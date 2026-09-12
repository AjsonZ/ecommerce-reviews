#!/usr/bin/env python3
"""Collect native-filtered ecommerce reviews and export them to Excel."""

from __future__ import annotations

import argparse
import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Sequence
from urllib.parse import parse_qs, urlparse


TAOBAO_RATE_TYPES = {
    "all": "-8",
    "positive": "1",
    "neutral": "0",
    "negative": "-1",
}
MAX_ALL_REVIEWS = 10_000
MAX_ALL_PAGES = 500

REVIEW_TYPE_ALIASES = {
    "all": "all",
    "全部": "all",
    "全部评价": "all",
    "positive": "positive",
    "good": "positive",
    "好评": "positive",
    "neutral": "neutral",
    "medium": "neutral",
    "中评": "neutral",
    "negative": "negative",
    "poor": "negative",
    "bad": "negative",
    "差评": "negative",
}


class EcommerceReviewsError(RuntimeError):
    """Base error for expected collection failures."""


class InputError(EcommerceReviewsError):
    """Raised when the product input or requested options are invalid."""


class OpenCLIError(EcommerceReviewsError):
    """Raised when OpenCLI or its browser bridge cannot complete a request."""


class PlatformResponseError(EcommerceReviewsError):
    """Raised when a platform returns a blocked or malformed response."""


class ExportError(EcommerceReviewsError):
    """Raised when the workbook exporter cannot complete."""


@dataclass(frozen=True)
class ProductRef:
    platform: str
    product_id: str
    canonical_url: str


@dataclass(frozen=True)
class ReviewRequest:
    product: ProductRef
    count: int
    review_type: str
    output_dir: Path
    all_reviews: bool = False


@dataclass(frozen=True)
class CollectionResult:
    rows: list[dict[str, Any]]
    pages_fetched: int = 0
    stop_reason: str = ""
    message: str = ""


@dataclass(frozen=True)
class Review:
    rank: int
    platform: str
    product_id: str
    review_id: str
    review_type: str
    user: str
    content: str
    date: str
    spec: str
    score: int | None
    source_order: int


SUPPORTED_PLATFORMS = {"taobao", "tmall"}
TMALL_HOSTS = {"detail.tmall.com", "chaoshi.detail.tmall.com"}
TAOBAO_HOSTS = {"item.taobao.com", "item.m.taobao.com"}


def _validate_numeric_id(value: str) -> str:
    cleaned = str(value or "").strip()
    if not re.fullmatch(r"\d{3,24}", cleaned):
        raise InputError("商品 ID 必须是 3–24 位数字。")
    return cleaned


def _canonical_for(platform: str, product_id: str) -> str:
    if platform == "tmall":
        return f"https://detail.tmall.com/item.htm?id={product_id}"
    if platform == "taobao":
        return f"https://item.taobao.com/item.htm?id={product_id}"
    raise InputError("平台仅支持 taobao 或 tmall。")


def parse_product(value: str, platform: str | None = None) -> ProductRef:
    raw = str(value or "").strip()
    platform_hint = str(platform or "").strip().lower() or None
    if platform_hint and platform_hint not in SUPPORTED_PLATFORMS:
        raise InputError("平台仅支持 taobao 或 tmall。")

    if re.fullmatch(r"\d+", raw):
        if not platform_hint:
            raise InputError("仅提供商品 ID 时必须同时指定平台。")
        product_id = _validate_numeric_id(raw)
        return ProductRef(platform_hint, product_id, _canonical_for(platform_hint, product_id))

    parsed = urlparse(raw)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise InputError("请输入完整的淘宝或天猫商品链接。")

    host = parsed.hostname.lower().rstrip(".")
    detected: str
    product_id: str | None = None
    if host in TMALL_HOSTS:
        detected = "tmall"
        product_id = parse_qs(parsed.query).get("id", [None])[0]
    elif host in TAOBAO_HOSTS:
        detected = "taobao"
        product_id = parse_qs(parsed.query).get("id", [None])[0]
    else:
        raise InputError("该链接不是受支持的淘宝或天猫商品域名。")

    if platform_hint and platform_hint != detected:
        raise InputError(f"指定平台 {platform_hint} 与链接平台 {detected} 不一致。")
    if not product_id:
        raise InputError("无法从商品链接中提取数字商品 ID。")
    product_id = _validate_numeric_id(product_id)
    return ProductRef(detected, product_id, _canonical_for(detected, product_id))


def validate_count(value: Any) -> int:
    if isinstance(value, bool):
        raise InputError(f"评论数量必须是 1–{MAX_ALL_REVIEWS} 的整数。")
    try:
        count = int(value)
    except (TypeError, ValueError) as error:
        raise InputError(f"评论数量必须是 1–{MAX_ALL_REVIEWS} 的整数。") from error
    if str(value).strip() != str(count) or not 1 <= count <= MAX_ALL_REVIEWS:
        raise InputError(f"评论数量必须是 1–{MAX_ALL_REVIEWS} 的整数。")
    return count


def normalize_review_type(value: str) -> str:
    normalized = str(value or "").strip().lower()
    result = REVIEW_TYPE_ALIASES.get(normalized)
    if not result:
        raise InputError("评价类型必须是全部、好评、中评或差评。")
    return result


def _json_config(data: dict[str, Any]) -> str:
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"))


def build_taobao_script(request: ReviewRequest) -> str:
    cfg = _json_config(
        {
            "auctionNumId": request.product.product_id,
            "rateType": TAOBAO_RATE_TYPES[request.review_type],
            "reviewType": request.review_type,
            "limit": request.count,
            "allReviews": request.all_reviews,
            "maxPages": MAX_ALL_PAGES,
        }
    )
    return f"""
(async () => {{
  const cfg={cfg};
  const normalize=(value)=>String(value ?? '').replace(/\\s+/g,' ').trim();
  const mtop=window.lib?.mtop;
  if (!mtop || typeof mtop.request !== 'function')
    return {{ok:false,error:'mtop_sdk_missing'}};

  const requestPage=async(page)=>{{
    const response=await mtop.request({{
      api:'mtop.taobao.rate.detaillist.get',
      v:'6.0',
      type:'get',
      dataType:'jsonp',
      ecode:'1',
      valueType:'string',
      timeout:20000,
      data:{{
        showTrueCount:false,
        auctionNumId:cfg.auctionNumId,
        pageNo:page,
        pageSize:20,
        orderType:'',
        searchImpr:'-8',
        expression:'',
        skuVids:'',
        rateSrc:'pc_rate_list',
        rateType:cfg.rateType,
        foldFlag:'0',
      }},
    }});
    const ret=String(response?.ret?.[0] || '');
    if (ret && !/^SUCCESS/i.test(ret)) throw new Error(ret);
    return response?.data || {{}};
  }};

  const output=[];
  const seen=new Set();
  let pagesFetched=0;
  try {{
    const maxPages=cfg.allReviews ? cfg.maxPages : Math.min(cfg.maxPages,Math.ceil(cfg.limit/20)+5);
    for (let page=1;page<=maxPages && output.length<cfg.limit;page++) {{
      const result=await requestPage(page);
      pagesFetched=page;
      const rows=Array.isArray(result.rateList) ? result.rateList : [];
      for (const item of rows) {{
        const content=normalize(item.feedback || '');
        if (!content) continue;
        const reviewId=normalize(item.id || item.feedbackId || item.rateId || '');
        const user=normalize(item.reduceUserNick || item.userNick || '');
        const date=normalize(item.feedbackDate || item.createTime || '').slice(0,30);
        const spec=normalize(item.skuValueStr || item.skuMap?.['商品规格'] || '');
        const nativeRate=normalize(item.rateType);
        const nativeType=nativeRate === '1' ? 'positive'
          : nativeRate === '0' ? 'neutral'
          : nativeRate === '-1' ? 'negative'
          : '';
        if (!nativeType) return {{ok:false,error:'unknown_native_rate_type'}};
        if (cfg.reviewType !== 'all' && nativeType !== cfg.reviewType)
          return {{ok:false,error:'filter_mismatch'}};
        const key=reviewId || [user,content,date,spec].join('\\u001f');
        if (seen.has(key)) continue;
        seen.add(key);
        output.push({{
          review_id:reviewId,
          review_type:nativeType,
          user,
          content,
          date,
          spec,
          score:null,
          source_order:output.length+1,
        }});
        if (output.length>=cfg.limit) break;
      }}
      if (rows.length === 0 || String(result.hasNext) !== 'true')
        return {{ok:true,rows:output.slice(0,cfg.limit),pages_fetched:pagesFetched,stop_reason:'source_exhausted'}};
    }}
    const stopReason=cfg.allReviews
      ? (output.length>=cfg.limit ? 'row_limit' : 'page_limit')
      : (output.length>=cfg.limit ? 'requested_count' : 'page_limit');
    return {{ok:true,rows:output.slice(0,cfg.limit),pages_fetched:pagesFetched,stop_reason:stopReason}};
  }} catch (error) {{
    if (cfg.allReviews && output.length>0)
      return {{ok:true,rows:output,pages_fetched:pagesFetched,stop_reason:'request_interrupted',message:normalize(error?.message || error)}};
    return {{ok:false,error:'request_failed',message:normalize(error?.message || error)}};
  }}
}})()
""".strip()


def _sanitize_error(message: str) -> str:
    text = re.sub(
        r"(?i)(cookie|token|authorization|password)\s*[:=]\s*[^\s;,]+",
        r"\1=[REDACTED]",
        str(message or ""),
    )
    if re.search(r"login|auth|cookie|登录|验证|captcha", text, re.I):
        return "浏览器登录状态无效或平台要求验证，请在浏览器中手动完成后重试。"
    return text.strip()[:500] or "OpenCLI 命令执行失败。"


def _decode_first_json(text: str) -> Any:
    decoder = json.JSONDecoder()
    source = str(text or "").lstrip("\ufeff\r\n ")
    for index, char in enumerate(source):
        if char not in "[{":
            continue
        try:
            value, _ = decoder.raw_decode(source[index:])
            return value
        except json.JSONDecodeError:
            continue
    raise OpenCLIError("OpenCLI 返回了无法解析的结果。")


class OpenCLIRunner:
    def __init__(
        self,
        executable: str | None = None,
        run_func: Callable[..., subprocess.CompletedProcess[str]] = subprocess.run,
    ) -> None:
        self.executable = executable or shutil.which("opencli") or shutil.which("opencli.cmd") or ""
        self.run_func = run_func
        if not self.executable:
            raise OpenCLIError("未找到 opencli，请先安装并连接 OpenCLI 浏览器桥。")
        self.command_prefix = self._resolve_command_prefix(self.executable)

    @staticmethod
    def _resolve_command_prefix(executable: str) -> list[str]:
        wrapper = Path(executable)
        if wrapper.suffix.lower() in {".cmd", ".ps1"} and wrapper.is_absolute():
            main_js = (
                wrapper.parent
                / "node_modules"
                / "@jackwener"
                / "opencli"
                / "dist"
                / "src"
                / "main.js"
            )
            node = _default_node()
            if main_js.exists() and node:
                return [node, str(main_js)]
        return [executable]

    def _run(self, args: list[str]) -> subprocess.CompletedProcess[str]:
        result = self.run_func(
            [*self.command_prefix, *args],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            shell=False,
        )
        if result.returncode != 0:
            raise OpenCLIError(_sanitize_error(result.stderr or result.stdout))
        return result

    def open(self, session: str, url: str) -> None:
        self._run(["browser", session, "open", url, "--window", "background"])

    def eval(self, session: str, javascript: str) -> dict[str, Any]:
        result = self._run(["browser", session, "eval", javascript])
        payload = _decode_first_json(result.stdout)
        if not isinstance(payload, dict):
            raise OpenCLIError("OpenCLI 浏览器返回的结果不是对象。")
        return payload

    def close(self, session: str) -> None:
        self._run(["browser", session, "close"])


def _browser_session_name(platform: str) -> tuple[str, bool]:
    """Return an optional user-selected session and whether it should be reused."""
    override = os.environ.get("ECOMMERCE_REVIEWS_BROWSER_SESSION", "").strip()
    if override:
        if not re.fullmatch(r"[A-Za-z0-9._:-]{1,80}", override):
            raise InputError("浏览器会话名包含不支持的字符。")
        return override, True
    return f"ecommerce-reviews-{os.getpid()}-{secrets.token_hex(4)}", False


def collect_review_result(
    request: ReviewRequest,
    runner: OpenCLIRunner | Any | None = None,
) -> CollectionResult:
    active_runner = runner or OpenCLIRunner()
    session, reused_session = _browser_session_name(request.product.platform)
    primary_error: BaseException | None = None
    try:
        active_runner.open(session, request.product.canonical_url)
        script = build_taobao_script(request)
        payload = active_runner.eval(session, script)
        if payload.get("ok") is not True:
            error = payload.get("error") or "malformed_response"
            if error == "mtop_sdk_missing":
                raise PlatformResponseError("商品页面未加载淘宝评价接口，可能需要登录或页面已改版。")
            if error == "filter_mismatch":
                raise PlatformResponseError("平台返回了与请求评价类型不一致的评论，已停止导出。")
            if error == "unknown_native_rate_type":
                raise PlatformResponseError("平台返回了无法识别的原生评价类型，已停止导出。")
            if error == "request_timeout":
                raise PlatformResponseError("平台评论接口请求超时，未生成 Excel。")
            if error == "request_failed":
                raise PlatformResponseError("平台评论接口请求失败，未生成 Excel。")
            raise PlatformResponseError(f"平台评论接口失败：{error}")
        rows = payload.get("rows")
        if not isinstance(rows, list):
            raise PlatformResponseError("平台评论响应缺少评论列表。")
        return CollectionResult(
            rows=[row for row in rows if isinstance(row, dict)],
            pages_fetched=int(payload.get("pages_fetched") or 0),
            stop_reason=_clean_field(payload.get("stop_reason")),
            message=_clean_field(payload.get("message")),
        )
    except BaseException as error:
        primary_error = error
        raise
    finally:
        if not reused_session:
            try:
                active_runner.close(session)
            except Exception:
                if primary_error is None:
                    raise


def collect_reviews(
    request: ReviewRequest,
    runner: OpenCLIRunner | Any | None = None,
) -> list[dict[str, Any]]:
    """Compatibility wrapper returning only normalized source rows."""
    return collect_review_result(request, runner).rows


def collection_note(request: ReviewRequest, result: CollectionResult, exported_count: int) -> str:
    if request.all_reviews:
        prefix = f"全部评论模式，上限 {MAX_ALL_REVIEWS} 条或 {MAX_ALL_PAGES} 页。"
        reasons = {
            "source_exhausted": "平台已无下一页，抓取正常结束。",
            "row_limit": f"已达到 {MAX_ALL_REVIEWS} 条安全上限。",
            "page_limit": f"已达到 {MAX_ALL_PAGES} 页安全上限。",
            "request_interrupted": "抓取因平台限流、验证或请求异常中断，已导出此前取得的有效评论。",
        }
        detail = reasons.get(result.stop_reason, "抓取已结束。")
        return f"{prefix}{detail} 共导出 {exported_count} 条。"
    if result.stop_reason == "page_limit":
        return f"已达到 {MAX_ALL_PAGES} 页安全上限，共导出 {exported_count} 条；少于请求的 {request.count} 条。"
    if exported_count < request.count:
        return f"平台仅返回 {exported_count} 条符合条件的有效评论。"
    return ""


def _clean_field(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize_reviews(
    request: ReviewRequest,
    rows: list[dict[str, Any]],
) -> list[Review]:
    reviews: list[Review] = []
    seen: set[str] = set()
    for row in rows:
        content = _clean_field(row.get("content"))
        if not content:
            continue
        review_id = _clean_field(row.get("review_id"))
        user = _clean_field(row.get("user"))
        date = _clean_field(row.get("date"))
        spec = _clean_field(row.get("spec"))
        native_type = normalize_review_type(str(row.get("review_type") or request.review_type))
        if request.review_type != "all" and native_type != request.review_type:
            raise InputError("平台返回了与请求评价类型不一致的评论，已停止导出。")
        key = review_id or "\x1f".join((user, content, date, spec))
        if key in seen:
            continue
        seen.add(key)
        raw_score = row.get("score")
        score = int(raw_score) if isinstance(raw_score, (int, float)) else None
        reviews.append(
            Review(
                rank=len(reviews) + 1,
                platform=request.product.platform,
                product_id=request.product.product_id,
                review_id=review_id,
                review_type=native_type,
                user=user,
                content=content,
                date=date,
                spec=spec,
                score=score,
                source_order=int(row.get("source_order") or len(reviews) + 1),
            )
        )
        if len(reviews) >= request.count:
            break
    return reviews


def _default_node() -> str:
    override = os.environ.get("ECOMMERCE_REVIEWS_NODE")
    if override:
        return override
    bundled = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe"
    if bundled.exists():
        return str(bundled)
    return shutil.which("node") or ""


def _default_node_modules() -> str:
    override = os.environ.get("ECOMMERCE_REVIEWS_NODE_MODULES")
    if override:
        return override
    bundled = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules"
    return str(bundled) if bundled.exists() else os.environ.get("NODE_PATH", "")


def export_workbook(
    request: ReviewRequest,
    reviews: list[Review],
    output_path: Path,
    note: str = "",
    run_func: Callable[..., subprocess.CompletedProcess[str]] = subprocess.run,
) -> Path:
    node = _default_node()
    node_modules = _default_node_modules()
    exporter = Path(__file__).with_name("export_reviews.mjs")
    if not node or not exporter.exists() or not node_modules:
        raise ExportError("Codex 的 Node.js 或 Artifact Tool 运行时不可用。")
    payload = {
        "metadata": {
            "platform": request.product.platform,
            "product_id": request.product.product_id,
            "canonical_url": request.product.canonical_url,
            "review_type": request.review_type,
            "requested_count": request.count,
            "all_reviews": request.all_reviews,
            "all_reviews_row_limit": MAX_ALL_REVIEWS,
            "all_reviews_page_limit": MAX_ALL_PAGES,
            "exported_count": len(reviews),
            "collected_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "note": note,
        },
        "reviews": [asdict(review) for review in reviews],
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = ""
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".json",
            delete=False,
        ) as handle:
            json.dump(payload, handle, ensure_ascii=False)
            temp_path = handle.name
        env = os.environ.copy()
        env["NODE_PATH"] = node_modules
        result = run_func(
            [node, str(exporter), temp_path, str(output_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            shell=False,
            env=env,
        )
        if result.returncode != 0:
            raise ExportError(_sanitize_error(result.stderr or result.stdout))
        if not output_path.exists():
            raise ExportError("导出器未生成 Excel 文件。")
        return output_path
    finally:
        if temp_path:
            Path(temp_path).unlink(missing_ok=True)


def _output_path(request: ReviewRequest, value: str | None) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = (
        f"{request.product.platform}_{request.product.product_id}_"
        f"{request.review_type}_{'all-available_' if request.all_reviews else ''}reviews_{timestamp}.xlsx"
    )
    if value:
        requested = Path(value).expanduser()
        return requested if requested.suffix.lower() == ".xlsx" else requested / filename
    return request.output_dir / filename


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="抓取淘宝或天猫原生评分评论并导出 Excel。"
    )
    parser.add_argument("input", help="商品链接；裸商品 ID 需同时使用 --platform")
    parser.add_argument("--platform", choices=sorted(SUPPORTED_PLATFORMS))
    amount = parser.add_mutually_exclusive_group(required=True)
    amount.add_argument("--count", help=f"抓取数量，1–{MAX_ALL_REVIEWS}")
    amount.add_argument(
        "--all-reviews",
        action="store_true",
        help=f"抓取接口可返回的全部评论，最多 {MAX_ALL_REVIEWS} 条或 {MAX_ALL_PAGES} 页",
    )
    parser.add_argument(
        "--review-type",
        required=True,
        help="全部/好评/中评/差评，或 all/positive/neutral/negative",
    )
    parser.add_argument("--output", help="输出 .xlsx 路径或目录")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
        product = parse_product(args.input, args.platform)
        count = MAX_ALL_REVIEWS if args.all_reviews else validate_count(args.count)
        review_type = normalize_review_type(args.review_type)
        request = ReviewRequest(
            product=product,
            count=count,
            review_type=review_type,
            output_dir=Path.cwd() / "outputs",
            all_reviews=args.all_reviews,
        )
        result = collect_review_result(request)
        reviews = normalize_reviews(request, result.rows)
        if not reviews:
            raise PlatformResponseError("平台未返回符合条件的有效评论，未生成 Excel。")
        note = collection_note(request, result, len(reviews))
        output_path = export_workbook(request, reviews, _output_path(request, args.output), note)
        print(
            json.dumps(
                {
                    "platform": product.platform,
                    "product_id": product.product_id,
                    "requested_count": "all" if args.all_reviews else count,
                    "all_reviews_row_limit": MAX_ALL_REVIEWS if args.all_reviews else None,
                    "all_reviews_page_limit": MAX_ALL_PAGES if args.all_reviews else None,
                    "pages_fetched": result.pages_fetched,
                    "stop_reason": result.stop_reason,
                    "exported_count": len(reviews),
                    "review_type": review_type,
                    "output_path": str(output_path.resolve()),
                    "note": note,
                },
                ensure_ascii=False,
            )
        )
        return 0
    except EcommerceReviewsError as error:
        print(str(error), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
