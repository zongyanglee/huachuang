import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const outputDir = path.join(root, "outputs", "parity_group_distribution");
const inputPath = path.join(outputDir, "parity_group_distribution.json");
const outputPath = path.join(outputDir, "平价分组转债个数占比.xlsx");

const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
const labels = payload.labels;
const ratioLastCol = String.fromCharCode("A".charCodeAt(0) + labels.length + 1);
const countLastCol = String.fromCharCode("A".charCodeAt(0) + labels.length + 2);

const workbook = Workbook.create();
const ratioSheet = workbook.worksheets.add("占比");
const countSheet = workbook.worksheets.add("数量");
const noteSheet = workbook.worksheets.add("说明");

function setTitle(sheet, title, subtitle, lastCol) {
  const titleRange = sheet.getRangeByIndexes(0, 0, 1, lastCol);
  titleRange.merge();
  titleRange.values = [[title]];
  titleRange.format = {
    fill: "#1F4E78",
    font: { bold: true, color: "#FFFFFF", size: 14 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  titleRange.format.rowHeightPx = 30;

  const subtitleRange = sheet.getRangeByIndexes(1, 0, 1, lastCol);
  subtitleRange.merge();
  subtitleRange.values = [[subtitle]];
  subtitleRange.format = {
    fill: "#D9EAF7",
    font: { color: "#1F4E78" },
    horizontalAlignment: "left",
  };
}

function styleTable(sheet, rowCount, colCount, percentCols = []) {
  const header = sheet.getRangeByIndexes(2, 0, 1, colCount);
  header.format = {
    fill: "#5B9BD5",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  const body = sheet.getRangeByIndexes(3, 0, Math.max(rowCount - 3, 1), colCount);
  body.format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };
  header.format.borders = { preset: "all", style: "thin", color: "#FFFFFF" };
  sheet.freezePanes.freezeRows(3);
  sheet.getRangeByIndexes(0, 0, rowCount, colCount).format.font = { name: "Microsoft YaHei" };
  sheet.getRangeByIndexes(0, 0, rowCount, colCount).format.autofitColumns();
  sheet.getRangeByIndexes(0, 0, rowCount, 1).format.columnWidthPx = 95;
  for (const col of percentCols) {
    sheet.getRangeByIndexes(3, col, Math.max(rowCount - 3, 1), 1).format.numberFormat = "0.00%";
  }
}

const ratioHeaders = ["日期", ...labels, "合计占比"];
const ratioRows = payload.ratio_rows.map((row) => [
  row.date,
  ...labels.map((label) => row.values[label]),
  row.classified_ratio,
]);
setTitle(
  ratioSheet,
  "平价分组转债个数占比",
  "分组规则：补充 <=50 和 >200；中间区间左开右闭，每20一档；分母为当日有效平价样本数。",
  ratioHeaders.length,
);
ratioSheet.getRangeByIndexes(2, 0, 1, ratioHeaders.length).values = [ratioHeaders];
ratioSheet.getRangeByIndexes(3, 0, ratioRows.length, ratioHeaders.length).values = ratioRows;
styleTable(ratioSheet, ratioRows.length + 3, ratioHeaders.length, labels.map((_, i) => i + 1).concat([ratioHeaders.length - 1]));
ratioSheet.tables.add(`A3:${ratioLastCol}${ratioRows.length + 3}`, true, "ParityRatioTable");

const countHeaders = ["日期", ...labels, "合计数量", "有效平价样本数"];
const countRows = payload.count_rows.map((row) => [
  row.date,
  ...labels.map((label) => row.values[label]),
  row.classified_total,
  row.total_valid,
]);
setTitle(
  countSheet,
  "平价分组转债数量",
  "数量口径与占比页一致；所有有效平价样本均被纳入 <=50、区间分组或 >200。",
  countHeaders.length,
);
countSheet.getRangeByIndexes(2, 0, 1, countHeaders.length).values = [countHeaders];
countSheet.getRangeByIndexes(3, 0, countRows.length, countHeaders.length).values = countRows;
styleTable(countSheet, countRows.length + 3, countHeaders.length);
countSheet.tables.add(`A3:${countLastCol}${countRows.length + 3}`, true, "ParityCountTable");

const firstDate = payload.summary_rows[0]?.date ?? "";
const lastDate = payload.summary_rows.at(-1)?.date ?? "";
const latest = payload.summary_rows.at(-1) ?? {};
const noteRows = [
  ["项目", "内容"],
  ["数据源", "转债个券历史序列/年份/月度 parquet；筛选 __sheet_name = 平价"],
  ["日期范围", `${firstDate} 至 ${lastDate}`],
  ["交易日数量", payload.summary_rows.length],
  ["分组规则", "<=50、(50,70]、(70,90]、(90,110]、(110,130]、(130,150]、(150,170]、(170,190]、(190,200]、>200"],
  ["占比分母", "当日有效平价样本数，即平价非空且可转为数值的转债数量"],
  ["最新日期有效样本数", latest.total_valid ?? ""],
  ["最新日期分组合计数量", latest.classified_total ?? ""],
  ["最新日期分组合计占比", latest.classified_ratio ?? ""],
  ["生成时间", payload.generated_at],
];
noteSheet.getRange("A1:B1").values = [["平价分组转债个数占比 - 说明与校验", ""]];
noteSheet.getRange("A1:B1").merge();
noteSheet.getRange("A1:B1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF", size: 14 },
  horizontalAlignment: "center",
};
noteSheet.getRangeByIndexes(2, 0, noteRows.length, 2).values = noteRows;
noteSheet.getRange("A3:B3").format = {
  fill: "#5B9BD5",
  font: { bold: true, color: "#FFFFFF" },
};
noteSheet.getRangeByIndexes(2, 0, noteRows.length, 2).format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };
noteSheet.getRange("B11").format.numberFormat = "0.00%";
noteSheet.getRange("A:B").format.autofitColumns();
noteSheet.getRange("A:A").format.columnWidthPx = 170;
noteSheet.getRange("B:B").format.columnWidthPx = 760;
noteSheet.getRange("B:B").format.wrapText = true;

const ratioChartRange = ratioSheet.getRange(`A3:${ratioLastCol}${ratioRows.length + 3}`);
const chart = ratioSheet.charts.add("line", ratioChartRange);
chart.title = "平价分组占比走势";
chart.hasLegend = true;
chart.xAxis = { axisType: "textAxis", tickLabelInterval: 180 };
chart.yAxis = { numberFormatCode: "0%" };
chart.setPosition("L3", "T22");

await fs.mkdir(outputDir, { recursive: true });

await workbook.inspect({
  kind: "table",
  range: `占比!A1:${ratioLastCol}8`,
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: ratioHeaders.length,
});
await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
await workbook.render({ sheetName: "占比", range: "A1:V25", scale: 1, format: "png" });
await workbook.render({ sheetName: "数量", range: "A1:M15", scale: 1, format: "png" });
await workbook.render({ sheetName: "说明", range: "A1:B12", scale: 1, format: "png" });

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
