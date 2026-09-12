#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";


const REVIEW_TYPE_LABELS = {
  all: "全部",
  positive: "好评",
  neutral: "中评",
  negative: "差评",
};

const PLATFORM_LABELS = {
  taobao: "淘宝",
  tmall: "天猫",
};


async function loadArtifactTool() {
  const require = createRequire(import.meta.url);
  let resolved;
  try {
    resolved = require.resolve("@oai/artifact-tool");
  } catch (error) {
    throw new Error(
      "无法加载 @oai/artifact-tool；请通过工作区依赖加载器提供 NODE_PATH。",
      { cause: error },
    );
  }
  return import(pathToFileURL(resolved).href);
}


function assertPayload(payload) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new Error("导出输入必须是 JSON 对象。");
  }
  if (!payload.metadata || typeof payload.metadata !== "object") {
    throw new Error("导出输入缺少 metadata。");
  }
  if (!Array.isArray(payload.reviews)) {
    throw new Error("导出输入缺少 reviews 数组。");
  }
}


function setColumnWidth(sheet, column, lastRow, width) {
  sheet.getRange(`${column}1:${column}${Math.max(lastRow, 1)}`).format.columnWidth = width;
}


async function main() {
  const [inputPath, outputPath] = process.argv.slice(2);
  if (!inputPath || !outputPath) {
    throw new Error("用法：export_reviews.mjs INPUT.json OUTPUT.xlsx");
  }

  const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
  assertPayload(payload);
  const { SpreadsheetFile, Workbook } = await loadArtifactTool();
  const workbook = Workbook.create();
  const infoSheet = workbook.worksheets.add("商品信息");
  const reviewSheet = workbook.worksheets.add("评论明细");
  const metadata = payload.metadata;
  const reviews = payload.reviews;
  const fontFamily = "Arial";
  const darkGreen = "#245447";
  const lightGreen = "#E8F1ED";
  const borderColor = "#CFD8D3";
  const bodyText = "#202624";

  infoSheet.showGridLines = false;
  reviewSheet.showGridLines = false;

  infoSheet.getRange("A1:B1").values = [["电商评论导出", null]];
  infoSheet.getRange("A1").format.font = {
    name: fontFamily,
    size: 16,
    bold: true,
    color: bodyText,
  };
  infoSheet.getRange("A2:B2").values = [["按平台原生评价类型筛选", null]];
  infoSheet.getRange("A2").format.font = {
    name: fontFamily,
    size: 10,
    italic: true,
    color: "#5B6863",
  };
  infoSheet.getRange("A3:B3").format.borders = {
    bottom: { style: "thin", color: borderColor },
  };

  const infoRows = [
    ["平台", PLATFORM_LABELS[metadata.platform] ?? metadata.platform ?? ""],
    ["商品 ID", String(metadata.product_id ?? "")],
    ["商品链接", String(metadata.canonical_url ?? "")],
    ["评价类型", REVIEW_TYPE_LABELS[metadata.review_type] ?? metadata.review_type ?? ""],
    ["请求数量", metadata.all_reviews
      ? `全部（上限 ${Number(metadata.all_reviews_row_limit ?? 10000)} 条或 ${Number(metadata.all_reviews_page_limit ?? 500)} 页）`
      : Number(metadata.requested_count ?? 0)],
    ["导出数量", Number(metadata.exported_count ?? reviews.length)],
    ["抓取时间", String(metadata.collected_at ?? "")],
    ["说明", String(metadata.note ?? "")],
  ];
  infoSheet.getRange("B5").format.numberFormat = "@";
  infoSheet.getRange(`A4:B${3 + infoRows.length}`).values = infoRows;
  infoSheet.getRange(`A4:A${3 + infoRows.length}`).format = {
    fill: lightGreen,
    font: { name: fontFamily, bold: true, color: bodyText },
    verticalAlignment: "center",
  };
  infoSheet.getRange(`B4:B${3 + infoRows.length}`).format = {
    font: { name: fontFamily, color: bodyText },
    verticalAlignment: "center",
    wrapText: true,
  };
  infoSheet.getRange("B10").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
  infoSheet.getRange(`A4:B${3 + infoRows.length}`).format.borders = {
    bottom: { style: "thin", color: borderColor },
  };
  setColumnWidth(infoSheet, "A", 11, 16);
  setColumnWidth(infoSheet, "B", 11, 72);
  infoSheet.getRange("A1:B11").format.rowHeight = 22;
  infoSheet.getRange("B6").format.rowHeight = 36;

  const headers = [
    "序号",
    "平台",
    "商品 ID",
    "评论 ID",
    "评价类型",
    "用户",
    "评论内容",
    "日期",
    "商品规格",
    "评分",
    "平台排序",
  ];
  const detailRows = reviews.map((review) => [
    Number(review.rank ?? 0),
    PLATFORM_LABELS[review.platform] ?? review.platform ?? "",
    String(review.product_id ?? ""),
    String(review.review_id ?? ""),
    REVIEW_TYPE_LABELS[review.review_type] ?? review.review_type ?? "",
    String(review.user ?? ""),
    String(review.content ?? ""),
    String(review.date ?? ""),
    String(review.spec ?? ""),
    review.score === null || review.score === undefined ? null : Number(review.score),
    Number(review.source_order ?? review.rank ?? 0),
  ]);
  const matrix = [headers, ...detailRows];
  const lastRow = matrix.length;
  if (lastRow > 1) reviewSheet.getRange(`C2:D${lastRow}`).format.numberFormat = "@";
  reviewSheet.getRange(`A1:K${lastRow}`).values = matrix;
  reviewSheet.getRange("A1:K1").format = {
    fill: darkGreen,
    font: { name: fontFamily, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#FFFFFF" },
  };
  if (lastRow > 1) {
    reviewSheet.getRange(`A2:K${lastRow}`).format = {
      font: { name: fontFamily, size: 10, color: bodyText },
      verticalAlignment: "top",
    };
    reviewSheet.getRange(`G2:G${lastRow}`).format.wrapText = true;
    reviewSheet.getRange(`I2:I${lastRow}`).format.wrapText = true;
    reviewSheet.getRange(`A2:F${lastRow}`).format.borders = {
      bottom: { style: "thin", color: borderColor },
    };
    reviewSheet.getRange(`G2:K${lastRow}`).format.borders = {
      bottom: { style: "thin", color: borderColor },
    };
    reviewSheet.getRange(`A2:K${lastRow}`).format.rowHeight = 42;
    for (let rowIndex = 0; rowIndex < detailRows.length; rowIndex += 1) {
      const contentLength = String(detailRows[rowIndex][6] ?? "").length;
      const rowHeight = Math.max(42, Math.min(180, Math.ceil(contentLength / 48) * 18));
      reviewSheet.getRange(`A${rowIndex + 2}:K${rowIndex + 2}`).format.rowHeight = rowHeight;
    }
    const table = reviewSheet.tables.add(`A1:K${lastRow}`, true, "ReviewsTable");
    table.style = "TableStyleMedium4";
    table.showFilterButton = true;
  }
  reviewSheet.getRange("A1:K1").format.rowHeight = 30;
  reviewSheet.freezePanes.freezeRows(1);

  const widths = {
    A: 8,
    B: 10,
    C: 19,
    D: 22,
    E: 12,
    F: 18,
    G: 64,
    H: 20,
    I: 34,
    J: 9,
    K: 12,
  };
  for (const [column, width] of Object.entries(widths)) {
    setColumnWidth(reviewSheet, column, lastRow, width);
  }
  reviewSheet.getRange(`A1:A${lastRow}`).format.horizontalAlignment = "center";
  reviewSheet.getRange(`B1:B${lastRow}`).format.horizontalAlignment = "center";
  reviewSheet.getRange(`E1:E${lastRow}`).format.horizontalAlignment = "center";
  reviewSheet.getRange(`J1:K${lastRow}`).format.horizontalAlignment = "center";

  const infoInspection = await workbook.inspect({
    kind: "table",
    range: "商品信息!A1:B11",
    include: "values,formulas",
    tableMaxRows: 20,
    tableMaxCols: 4,
  });
  const reviewInspection = await workbook.inspect({
    kind: "table",
    range: `评论明细!A1:K${Math.min(lastRow, 12)}`,
    include: "values,formulas",
    tableMaxRows: 12,
    tableMaxCols: 11,
  });
  const formulaErrors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
    options: { useRegex: true, maxResults: 100 },
    summary: "final formula error scan",
  });
  if (String(formulaErrors.ndjson ?? "").match(/#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A|#NUM!|#NULL!|#SPILL!|#CALC!/)) {
    throw new Error("工作簿检查发现公式错误。");
  }

  const infoPreview = await workbook.render({
    sheetName: "商品信息",
    range: "A1:B11",
    scale: 1,
    format: "png",
  });
  const reviewPreview = await workbook.render({
    sheetName: "评论明细",
    range: `A1:K${Math.min(lastRow, 12)}`,
    scale: 1,
    format: "png",
  });
  const infoPreviewBytes = new Uint8Array(await infoPreview.arrayBuffer());
  const reviewPreviewBytes = new Uint8Array(await reviewPreview.arrayBuffer());
  if (infoPreviewBytes.byteLength < 500 || reviewPreviewBytes.byteLength < 500) {
    throw new Error("工作簿视觉预览为空或异常。")
  }

  await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  process.stdout.write(JSON.stringify({
    outputPath: path.resolve(outputPath),
    sheets: ["商品信息", "评论明细"],
    reviewRows: detailRows.length,
    infoInspectBytes: String(infoInspection.ndjson ?? "").length,
    reviewInspectBytes: String(reviewInspection.ndjson ?? "").length,
    renderBytes: [infoPreviewBytes.byteLength, reviewPreviewBytes.byteLength],
  }));
}


main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 1;
});
