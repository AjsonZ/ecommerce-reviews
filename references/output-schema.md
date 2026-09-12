# Output Schema

## Native Rating Filters

| User choice | Normalized value | Taobao/Tmall `rateType` |
|---|---|---:|
| 全部 | `all` | -8 |
| 好评 | `positive` | 1 |
| 中评 | `neutral` | 0 |
| 差评 | `negative` | -1 |

These values select platform-native categories. They are not sentiment scores calculated from comment text.

## Workbook

`商品信息` contains:

- platform and product ID;
- canonical product URL;
- requested review type and count;
- exported count;
- collection time;
- shortfall or status note.

No workbook is created when the platform returns zero valid comments or an unusable response. The caller receives an actionable error instead.

## Collection Amount

- Fixed mode accepts 1-10,000 reviews and is also limited to 500 requested pages.
- All-reviews mode continues until the platform reports no next page, with a hard ceiling of 10,000 valid reviews or 500 requested pages.
- The `请求数量` cell displays `全部（上限 10000 条或 500 页）` in all-reviews mode.
- The `说明` cell records whether collection ended naturally, reached a safety limit, or exported a partial result after interruption.
- “All” means all reviews currently exposed by the platform interface; it may be less than the total displayed on the product page.

## Authentication Notes

- Taobao/Tmall commonly expose review data without login, but risk control or verification may require the user's existing browser login state.
- Login success alone does not guarantee collection when a platform endpoint has migrated or returns an empty or malformed response.

`评论明细` contains:

| Column | Internal field | Meaning |
|---|---|---|
| 序号 | `rank` | One-based export order |
| 平台 | `platform` | Taobao or Tmall |
| 商品 ID | `product_id` | Item ID |
| 评论 ID | `review_id` | Native review ID when returned |
| 评价类型 | `review_type` | Native category |
| 用户 | `user` | Platform-masked nickname |
| 评论内容 | `content` | Full returned review text |
| 日期 | `date` | Platform-provided date/time |
| 商品规格 | `spec` | SKU attributes when returned |
| 评分 | `score` | Native product score when returned; blank for Taobao/Tmall because `userStar` is an account level |
| 平台排序 | `source_order` | Order within the selected platform filter |
