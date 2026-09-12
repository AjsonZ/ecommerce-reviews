---
name: ecommerce-reviews
description: Collect native rating-filtered product reviews from Taobao or Tmall and export them to Excel. Use when a user provides one of these product links or IDs and asks to fetch, filter, collect, analyze, or export reviews, including 好评、中评、差评或全部评价.
---

# Ecommerce Reviews

Use the bundled scripts to collect reviews and create one verified `.xlsx` workbook. The workflow is read-only.

## Required Inputs

Accept a full Taobao or Tmall product URL. A bare numeric product ID also requires an explicit platform.

Before collection, make sure both choices are known:

- amount: an integer from 1 through 10,000, or `抓取全部评论`;
- review type: `全部`, `好评`, `中评`, or `差评`.

If either choice is missing, ask for both missing values in one concise prompt. Do not silently default them.

Distinguish amount from rating type. `全部评价类型` means include positive, neutral, and negative reviews but still uses the chosen amount. `抓取全部评论` means keep paging until the platform has no next page, subject to a hard safety limit of 10,000 valid reviews or 500 pages, whichever occurs first. If the user simply says `抓取全部评论`, infer both all-reviews amount mode and `全部评价类型`. Before starting an all-reviews run, explicitly tell the user about both limits and that the platform may expose fewer reviews than its displayed total.

## Run

1. Call the workspace dependency loader and use its Python executable. The bundled exporter discovers Node.js and `@oai/artifact-tool` from the same Codex runtime.
2. Choose an output path under the current task's user-facing `outputs` directory unless the user supplies another path.
3. Run:

```powershell
& '<python-executable>' '<skill-directory>/scripts/fetch_reviews.py' '<product-url-or-id>' --count <1-10000> --review-type '<全部|好评|中评|差评>' --output '<output-directory>'
```

For all available reviews, replace `--count` with `--all-reviews`:

```powershell
& '<python-executable>' '<skill-directory>/scripts/fetch_reviews.py' '<product-url-or-id>' --all-reviews --review-type '<全部|好评|中评|差评>' --output '<output-directory>'
```

Add `--platform taobao|tmall` only for a bare product ID. Pass arguments as an array or safely quoted literals; never construct a shell command from untrusted URL fragments.

By default the script uses a fresh isolated browser session. When the user has already logged in through a named OpenCLI browser session, set `ECOMMERCE_REVIEWS_BROWSER_SESSION` to that exact session name for the run. The skill does not discover, open, or log in to arbitrary sessions automatically; confirm the session belongs to the user before reusing it.

The script prints one JSON summary with the detected platform, product ID, requested and exported counts, filter, workbook path, and any shortfall note. Link the resulting workbook to the user.

## Behavior

- Tmall and Taobao use `mtop.taobao.rate.detaillist.get` and its native `rateType` filter.
- Validate each returned Taobao/Tmall row against its native `rateType`; stop rather than export if the platform ignores or changes the requested filter.
- Preserve returned comment text. Do not summarize or classify it with model sentiment.
- Keep the platform's default order within the selected category.
- When fewer matching reviews exist, export those found and report the shortfall recorded in `商品信息`.
- In all-reviews mode, stop when the source is exhausted, 10,000 valid reviews are collected, or 500 pages are requested. Record the limit and stop reason in `商品信息`.
- If an all-reviews run is interrupted by rate limiting, verification, or a request failure after valid rows were collected, export the partial result and clearly label it incomplete. With zero valid rows, fail without a workbook.
- Do not replace missing results with another rating category.
- Leave Taobao/Tmall `评分` blank: the current API's `userStar` is the reviewer's account level, not a product score.

## Authentication Matrix

Login state is platform-dependent and is never created or automated by this skill:

| Platform | Typical requirement | If unavailable |
|---|---|---|
| 淘宝 | Usually readable anonymously; login may be requested by risk control | Ask the user to complete verification/login manually, then retry |
| 天猫 | Usually readable anonymously; login may be requested by risk control | Ask the user to complete verification/login manually, then retry |

Tell the user this matrix when the skill is used if authentication could affect the requested run. A valid login does not guarantee data availability: endpoint migration, rate limiting, or risk control can still prevent collection.

## Safety and Failures

The script uses an isolated OpenCLI browser session only to make read-only requests. It may reuse login state already present in OpenCLI. Never invoke platform login, export cookies, print credentials, add items to a cart, follow a store, or post a review.

If login is missing, a verification challenge appears, an endpoint shape changes, or the product is unavailable, stop and report the specific error without creating an empty workbook. Do not fall back to text scraping that cannot prove the requested native rating category.

If the platform returns zero valid rows, do not create an empty workbook. Report whether the selected category was empty or the endpoint returned no usable review data.

Read [references/output-schema.md](references/output-schema.md) only when explaining workbook fields or extending the bundled scripts.
