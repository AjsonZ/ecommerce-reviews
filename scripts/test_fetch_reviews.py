import json
import os
import shutil
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from fetch_reviews import (
    CollectionResult,
    InputError,
    MAX_ALL_PAGES,
    MAX_ALL_REVIEWS,
    OpenCLIError,
    OpenCLIRunner,
    ProductRef,
    ReviewRequest,
    TAOBAO_RATE_TYPES,
    build_taobao_script,
    collection_note,
    collect_reviews,
    export_workbook,
    main,
    normalize_review_type,
    normalize_reviews,
    parse_product,
    validate_count,
)


class ParsingTests(unittest.TestCase):
    def test_parses_supported_urls(self):
        tmall = parse_product(
            "https://detail.tmall.com/item.htm?id=1008809150955&skuId=6008539376934"
        )
        self.assertEqual(tmall.platform, "tmall")
        self.assertEqual(tmall.product_id, "1008809150955")
        self.assertEqual(
            tmall.canonical_url,
            "https://detail.tmall.com/item.htm?id=1008809150955",
        )

        taobao = parse_product("https://item.taobao.com/item.htm?id=827563850178")
        self.assertEqual(taobao.platform, "taobao")
        self.assertEqual(taobao.product_id, "827563850178")

    def test_rejects_deceptive_or_ambiguous_inputs(self):
        with self.assertRaises(InputError):
            parse_product("https://detail.tmall.com.evil.example/item.htm?id=1008809150955")
        with self.assertRaises(InputError):
            parse_product("1008809150955")
        with self.assertRaises(InputError):
            parse_product("https://example.com/item.htm?id=1008809150955")
        with self.assertRaises(InputError):
            parse_product("https://item.jd.com/100291143898.html")
        with self.assertRaises(InputError):
            parse_product("100291143898", "jd")

    def test_accepts_bare_id_with_platform(self):
        product = parse_product("1008809150955", "taobao")
        self.assertEqual(product.platform, "taobao")
        self.assertEqual(
            product.canonical_url,
            "https://item.taobao.com/item.htm?id=1008809150955",
        )

    def test_validates_count_and_review_type(self):
        self.assertEqual(validate_count(1), 1)
        self.assertEqual(validate_count(MAX_ALL_REVIEWS), MAX_ALL_REVIEWS)
        for invalid in (0, MAX_ALL_REVIEWS + 1, "ten"):
            with self.assertRaises(InputError):
                validate_count(invalid)
        self.assertEqual(normalize_review_type("好评"), "positive")
        self.assertEqual(normalize_review_type("中评"), "neutral")
        self.assertEqual(normalize_review_type("差评"), "negative")
        self.assertEqual(normalize_review_type("全部"), "all")


class FilterMappingTests(unittest.TestCase):
    def test_native_filter_mappings(self):
        self.assertEqual(
            TAOBAO_RATE_TYPES,
            {"all": "-8", "positive": "1", "neutral": "0", "negative": "-1"},
        )

    def test_scripts_embed_only_validated_request_values(self):
        request = ReviewRequest(
            product=ProductRef("tmall", "1008809150955", "https://detail.tmall.com/item.htm?id=1008809150955"),
            count=17,
            review_type="negative",
            output_dir=Path("outputs"),
        )
        taobao_script = build_taobao_script(request)
        self.assertIn('"auctionNumId":"1008809150955"', taobao_script)
        self.assertIn('"rateType":"-1"', taobao_script)
        self.assertIn('"limit":17', taobao_script)
        self.assertIn("mtop.taobao.rate.detaillist.get", taobao_script)
        self.assertIn("filter_mismatch", taobao_script)
        self.assertIn("score:null", taobao_script)

        all_request = ReviewRequest(
            product=request.product,
            count=MAX_ALL_REVIEWS,
            review_type="all",
            output_dir=Path("outputs"),
            all_reviews=True,
        )
        all_script = build_taobao_script(all_request)
        self.assertIn('"allReviews":true', all_script)
        self.assertIn(f'"maxPages":{MAX_ALL_PAGES}', all_script)
        self.assertIn("request_interrupted", all_script)

    def test_generated_javascript_compiles(self):
        node = os.environ.get("ECOMMERCE_REVIEWS_NODE") or shutil.which("node")
        if not node:
            self.skipTest("Node.js is unavailable")
        requests = [
            ReviewRequest(
                ProductRef("tmall", "1008809150955", "https://detail.tmall.com/item.htm?id=1008809150955"),
                10,
                "all",
                Path("outputs"),
            ),
            ReviewRequest(
                ProductRef("taobao", "827563850178", "https://item.taobao.com/item.htm?id=827563850178"),
                MAX_ALL_REVIEWS,
                "all",
                Path("outputs"),
                all_reviews=True,
            ),
        ]
        compiler = "new Function(require('fs').readFileSync(0,'utf8'));"
        for request in requests:
            result = subprocess.run(
                [node, "-e", compiler],
                input=build_taobao_script(request),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
                shell=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)


class RunnerTests(unittest.TestCase):
    def test_runner_uses_argument_arrays_and_parses_json(self):
        calls = []

        def fake_run(args, **kwargs):
            calls.append((args, kwargs))
            return subprocess.CompletedProcess(args, 0, stdout='{"ok":true}\n', stderr="")

        runner = OpenCLIRunner(executable="opencli.cmd", run_func=fake_run)
        runner.open("session-a", "https://detail.tmall.com/item.htm?id=123")
        result = runner.eval("session-a", "({ok:true})")
        runner.close("session-a")

        self.assertEqual(result, {"ok": True})
        self.assertEqual(calls[0][0][:4], ["opencli.cmd", "browser", "session-a", "open"])
        self.assertFalse(calls[0][1]["shell"])
        self.assertEqual(calls[-1][0], ["opencli.cmd", "browser", "session-a", "close"])

    def test_runner_raises_sanitized_error(self):
        def fake_run(args, **kwargs):
            return subprocess.CompletedProcess(
                args,
                1,
                stdout="",
                stderr="authentication required; cookie=secret-value",
            )

        runner = OpenCLIRunner(executable="opencli.cmd", run_func=fake_run)
        with self.assertRaises(OpenCLIError) as caught:
            runner.open("session-a", "https://detail.tmall.com/item.htm?id=123")
        self.assertNotIn("secret-value", str(caught.exception))

    def test_collection_always_closes_session(self):
        class FakeRunner:
            def __init__(self):
                self.closed = []

            def open(self, session, url):
                return None

            def eval(self, session, javascript):
                raise OpenCLIError("blocked")

            def close(self, session):
                self.closed.append(session)

        request = ReviewRequest(
            product=ProductRef("tmall", "1008809150955", "https://detail.tmall.com/item.htm?id=1008809150955"),
            count=10,
            review_type="all",
            output_dir=Path("outputs"),
        )
        runner = FakeRunner()
        with self.assertRaises(OpenCLIError):
            collect_reviews(request, runner)
        self.assertEqual(len(runner.closed), 1)

    def test_collection_can_reuse_named_session_without_closing_it(self):
        class FakeRunner:
            def __init__(self):
                self.opened = []
                self.closed = []

            def open(self, session, url):
                self.opened.append(session)

            def eval(self, session, javascript):
                return {"ok": True, "rows": [{"content": "内容", "review_type": "positive"}]}

            def close(self, session):
                self.closed.append(session)

        request = ReviewRequest(
            product=ProductRef("tmall", "1008809150955", "https://detail.tmall.com/item.htm?id=1008809150955"),
            count=1,
            review_type="positive",
            output_dir=Path("outputs"),
        )
        runner = FakeRunner()
        with patch.dict(os.environ, {"ECOMMERCE_REVIEWS_BROWSER_SESSION": "tmall-debug"}):
            rows = collect_reviews(request, runner)
        self.assertEqual(len(rows), 1)
        self.assertEqual(runner.opened, ["tmall-debug"])
        self.assertEqual(runner.closed, [])


class NormalizationTests(unittest.TestCase):
    def test_deduplicates_and_preserves_content(self):
        request = ReviewRequest(
            product=ProductRef("taobao", "123", "https://item.taobao.com/item.htm?id=123"),
            count=10,
            review_type="all",
            output_dir=Path("outputs"),
        )
        rows = [
            {"review_id": "a", "user": "u***1", "content": "完整评论，不应截断", "date": "2026-01-02", "spec": "黑色", "score": 5, "review_type": "positive", "source_order": 1},
            {"review_id": "a", "user": "u***1", "content": "重复", "date": "2026-01-02", "spec": "黑色", "score": 5, "review_type": "positive", "source_order": 2},
            {"review_id": "", "user": "🌈用户", "content": " 第二条 ", "date": "", "spec": "", "score": 3, "review_type": "neutral", "source_order": 3},
            {"review_id": "", "user": "🌈用户", "content": "第二条", "date": "", "spec": "", "score": 3, "review_type": "neutral", "source_order": 4},
            {"review_id": "b", "user": "empty", "content": "   ", "date": "", "spec": "", "score": 1, "review_type": "negative", "source_order": 5},
        ]
        reviews = normalize_reviews(request, rows)
        self.assertEqual(len(reviews), 2)
        self.assertEqual([review.rank for review in reviews], [1, 2])
        self.assertEqual(reviews[0].content, "完整评论，不应截断")
        self.assertEqual(reviews[1].user, "🌈用户")

    def test_rejects_category_mismatch(self):
        request = ReviewRequest(
            product=ProductRef("tmall", "123", "https://detail.tmall.com/item.htm?id=123"),
            count=10,
            review_type="negative",
            output_dir=Path("outputs"),
        )
        rows = [{"review_id": "a", "content": "好", "review_type": "positive"}]
        with self.assertRaises(InputError):
            normalize_reviews(request, rows)

    def test_all_reviews_notes_disclose_limits_and_stop_reason(self):
        request = ReviewRequest(
            ProductRef("tmall", "123", "https://detail.tmall.com/item.htm?id=123"),
            MAX_ALL_REVIEWS,
            "all",
            Path("outputs"),
            all_reviews=True,
        )
        exhausted = collection_note(
            request,
            CollectionResult([], pages_fetched=12, stop_reason="source_exhausted"),
            217,
        )
        self.assertIn("10000 条或 500 页", exhausted)
        self.assertIn("无下一页", exhausted)
        interrupted = collection_note(
            request,
            CollectionResult([], pages_fetched=3, stop_reason="request_interrupted"),
            40,
        )
        self.assertIn("中断", interrupted)


class WorkbookTests(unittest.TestCase):
    def test_artifact_exporter_creates_required_workbook_structure(self):
        request = ReviewRequest(
            product=ProductRef("tmall", "1008809150955", "https://detail.tmall.com/item.htm?id=1008809150955"),
            count=2,
            review_type="all",
            output_dir=Path("outputs"),
            all_reviews=True,
        )
        rows = [
            {"review_id": "r1", "user": "t***1", "content": "佩戴舒服", "date": "2026-09-01", "spec": "黑色", "score": 1, "review_type": "positive", "source_order": 1},
            {"review_id": "r2", "user": "t***2", "content": "续航很好", "date": "2026-09-02", "spec": "白色", "score": 1, "review_type": "positive", "source_order": 2},
        ]
        reviews = normalize_reviews(request, rows)
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "reviews.xlsx"
            export_workbook(request, reviews, output)
            self.assertTrue(output.exists())
            self.assertGreater(output.stat().st_size, 5000)
            with zipfile.ZipFile(output) as archive:
                workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
                detail_xml = archive.read("xl/worksheets/sheet2.xml").decode("utf-8")
                table_xml = archive.read("xl/tables/table1.xml").decode("utf-8")
                workbook_text = "\n".join(
                    archive.read(name).decode("utf-8", errors="ignore")
                    for name in archive.namelist()
                    if name.endswith(".xml")
                )
            self.assertIn('name="商品信息"', workbook_xml)
            self.assertIn('name="评论明细"', workbook_xml)
            self.assertIn('state="frozen"', detail_xml)
            self.assertIn('ySplit="1"', detail_xml)
            self.assertIn('autoFilter ref="A1:K3"', table_xml)
            self.assertNotIn("<f>", detail_xml)
            self.assertIn("全部（上限 10000 条或 500 页）", workbook_text)


class CliTests(unittest.TestCase):
    def test_main_emits_summary(self):
        raw = [{"review_id": "r1", "user": "u", "content": "内容", "review_type": "positive", "source_order": 1}]
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "out.xlsx"
            with patch("fetch_reviews.collect_review_result", return_value=CollectionResult(raw)), patch(
                "fetch_reviews.export_workbook", return_value=output
            ):
                with patch("builtins.print") as print_mock:
                    code = main([
                        "1008809150955",
                        "--platform",
                        "tmall",
                        "--count",
                        "1",
                        "--review-type",
                        "好评",
                        "--output",
                        str(output),
                    ])
            self.assertEqual(code, 0)
            summary = json.loads(print_mock.call_args.args[0])
            self.assertEqual(summary["platform"], "tmall")
            self.assertEqual(summary["review_type"], "positive")
            self.assertEqual(summary["exported_count"], 1)

    def test_main_rejects_empty_result_without_export(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "should-not-exist.xlsx"
            with patch("fetch_reviews.collect_review_result", return_value=CollectionResult([])), patch(
                "fetch_reviews.export_workbook"
            ) as export_mock:
                with patch("builtins.print") as print_mock:
                    code = main([
                        "1008809150955",
                        "--platform",
                        "tmall",
                        "--count",
                        "1",
                        "--review-type",
                        "全部",
                        "--output",
                        str(output),
                    ])
            self.assertEqual(code, 2)
            export_mock.assert_not_called()
            self.assertIn("未返回符合条件", print_mock.call_args.args[0])

    def test_main_accepts_all_reviews_without_count(self):
        raw = [{"review_id": "r1", "user": "u", "content": "内容", "review_type": "positive", "source_order": 1}]
        result = CollectionResult(raw, pages_fetched=1, stop_reason="source_exhausted")
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "all.xlsx"
            with patch("fetch_reviews.collect_review_result", return_value=result), patch(
                "fetch_reviews.export_workbook", return_value=output
            ):
                with patch("builtins.print") as print_mock:
                    code = main([
                        "1008809150955",
                        "--platform",
                        "tmall",
                        "--all-reviews",
                        "--review-type",
                        "全部",
                        "--output",
                        str(output),
                    ])
            self.assertEqual(code, 0)
            summary = json.loads(print_mock.call_args.args[0])
            self.assertEqual(summary["requested_count"], "all")
            self.assertEqual(summary["all_reviews_row_limit"], MAX_ALL_REVIEWS)
            self.assertEqual(summary["all_reviews_page_limit"], MAX_ALL_PAGES)


if __name__ == "__main__":
    unittest.main()
