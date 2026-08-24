import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = process.cwd();
const dir = path.join(root, "outputs", "rating_annual_stats_since_2018");
const data = JSON.parse(await fs.readFile(path.join(dir, "stats.json"), "utf8"));
const outputPath = path.join(dir, "2018年以来转债各评级规模数量及收益率统计.xlsx");

const wb = Workbook.create();
const overallSheet = wb.worksheets.add("总体统计");
const annualSheet = wb.worksheets.add("年度统计");
const detailSheet = wb.worksheets.add("个券年度明细");
const noteSheet = wb.worksheets.add("口径说明");

const navy = "#17365D";
const blue = "#2F75B5";
const teal = "#2A7F79";
const border = "#D9E1F2";
const text = "#1F2937";

function title(sheet, range, value) {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[value]];
  sheet.getRange(range).format = {
    fill: navy,
    font: { name: "Microsoft YaHei", size: 16, bold: true, color: "#FFFFFF" },
    verticalAlignment: "center",
    rowHeight: 30,
  };
}

function header(range) {
  range.format = {
    fill: blue,
    font: { name: "Microsoft YaHei", size: 10, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeight: 30,
    borders: { preset: "all", style: "thin", color: border },
  };
}

function body(range) {
  range.format = {
    font: { name: "Microsoft YaHei", size: 9, color: text },
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#E7E6E6" },
  };
}

function writeStats(sheet, startRow, rows, tableName) {
  const headers = ["期间", "评级", "发行规模合计(亿元)", "发行规模占比", "个数", "个数占比", "收益率简单平均", "收益率中位数", "收益率有效样本数"];
  sheet.getRangeByIndexes(startRow - 1, 0, 1, headers.length).values = [headers];
  header(sheet.getRangeByIndexes(startRow - 1, 0, 1, headers.length));
  const values = rows.map((row) => headers.map((key) => row[key]));
  sheet.getRangeByIndexes(startRow, 0, values.length, headers.length).values = values;
  body(sheet.getRangeByIndexes(startRow, 0, values.length, headers.length));
  sheet.getRangeByIndexes(startRow, 2, values.length, 1).format.numberFormat = "0.00";
  sheet.getRangeByIndexes(startRow, 3, values.length, 1).format.numberFormat = "0.0%";
  sheet.getRangeByIndexes(startRow, 5, values.length, 4).format.numberFormat = "0.0%;[Red](0.0%);-";
  sheet.getRangeByIndexes(startRow, 4, values.length, 1).format.numberFormat = "0";
  sheet.getRangeByIndexes(startRow, 8, values.length, 1).format.numberFormat = "0";
  sheet.tables.add(`A${startRow}:I${startRow + values.length}`, true, tableName).style = "TableStyleMedium2";
  return { headers, values };
}

// Overall
overallSheet.showGridLines = false;
title(overallSheet, "A1:I1", `2018年以来转债各评级总体统计（截至${data.latest_date}）`);
overallSheet.getRange("A3:I3").merge();
overallSheet.getRange("A3:I3").values = [["每只转债只计一次；按2018年以来首次正常交易日的债项评级归类。收益率为区间首末正常交易收盘价收益，组内简单算术平均。"]];
overallSheet.getRange("A3:I3").format = {
  fill: "#F2F2F2",
  font: { name: "Microsoft YaHei", size: 10, italic: true, color: "#44546A" },
  wrapText: true,
  rowHeight: 28,
};
writeStats(overallSheet, 5, data.overall_stats, "OverallRatingStats");
overallSheet.getRange("A:A").format.columnWidth = 14;
overallSheet.getRange("B:B").format.columnWidth = 10;
overallSheet.getRange("C:C").format.columnWidth = 20;
overallSheet.getRange("D:I").format.columnWidth = 16;
overallSheet.freezePanes.freezeRows(5);

overallSheet.getRange("K20:L20").values = [["评级", "发行规模占比"]];
overallSheet.getRange(`K21:L${20 + data.overall_stats.length}`).values = data.overall_stats.map((row) => [row["评级"], row["发行规模占比"]]);
const overallChart = overallSheet.charts.add("bar", overallSheet.getRange(`K20:L${20 + data.overall_stats.length}`));
overallChart.title = "各评级发行规模占比";
overallChart.hasLegend = false;
overallChart.yAxis = { numberFormatCode: "0%" };
overallChart.setPosition("K3", "S18");

// Annual long-form table
annualSheet.showGridLines = false;
title(annualSheet, "A1:I1", `2018-${data.end_year}年各评级年度统计`);
annualSheet.getRange("A3:I3").merge();
annualSheet.getRange("A3:I3").values = [["年度样本为年内至少有一个正常交易日的转债；评级取该年首个正常交易日评级，上市首日不计入收益。2026年为截至最新数据日的年内统计。"]];
annualSheet.getRange("A3:I3").format = {
  fill: "#F2F2F2",
  font: { name: "Microsoft YaHei", size: 10, italic: true, color: "#44546A" },
  wrapText: true,
  rowHeight: 28,
};
writeStats(annualSheet, 5, data.annual_stats, "AnnualRatingStats");
annualSheet.getRange("A:B").format.columnWidth = 11;
annualSheet.getRange("C:C").format.columnWidth = 20;
annualSheet.getRange("D:I").format.columnWidth = 16;
annualSheet.freezePanes.freezeRows(5);

// Detail
detailSheet.showGridLines = false;
title(detailSheet, "A1:L1", "个券年度收益明细");
const detailHeaders = data.detail_columns;
detailSheet.getRangeByIndexes(2, 0, 1, detailHeaders.length).values = [detailHeaders];
header(detailSheet.getRangeByIndexes(2, 0, 1, detailHeaders.length));
const dateFields = new Set(["首个正常交易日", "末个正常交易日"]);
const detailRows = data.detail.map((row) => detailHeaders.map((key) => {
  const value = row[key];
  if (value == null) return null;
  if (dateFields.has(key)) return new Date(String(value).slice(0, 10));
  return value;
}));
detailSheet.getRangeByIndexes(3, 0, detailRows.length, detailHeaders.length).values = detailRows;
body(detailSheet.getRangeByIndexes(3, 0, detailRows.length, detailHeaders.length));
for (const field of dateFields) {
  const index = detailHeaders.indexOf(field);
  detailSheet.getRangeByIndexes(3, index, detailRows.length, 1).format.numberFormat = "yyyy-mm-dd";
}
detailSheet.getRangeByIndexes(3, detailHeaders.indexOf("发行规模"), detailRows.length, 1).format.numberFormat = "0.00";
detailSheet.getRangeByIndexes(3, detailHeaders.indexOf("首日收盘价"), detailRows.length, 2).format.numberFormat = "0.000";
detailSheet.getRangeByIndexes(3, detailHeaders.indexOf("收益率"), detailRows.length, 1).format.numberFormat = "0.0%;[Red](0.0%);-";
detailSheet.tables.add(`A3:L${detailRows.length + 3}`, true, "BondYearDetail").style = "TableStyleMedium2";
detailSheet.getRange("A:A").format.columnWidth = 9;
detailSheet.getRange("B:C").format.columnWidth = 14;
detailSheet.getRange("D:D").format.columnWidth = 9;
detailSheet.getRange("E:F").format.columnWidth = 14;
detailSheet.getRange("G:H").format.columnWidth = 15;
detailSheet.getRange("I:L").format.columnWidth = 14;
detailSheet.freezePanes.freezeRows(3);
detailSheet.freezePanes.freezeColumns(3);

// Notes
noteSheet.showGridLines = false;
title(noteSheet, "A1:D1", "统计口径说明");
noteSheet.getRange("A3:D3").values = [["项目", "口径", "单位", "备注"]];
header(noteSheet.getRange("A3:D3"));
const notes = [
  ["统计期间", `2018-01-01至${data.latest_date}`, "日期", `${data.end_year}年为不完整年度`],
  ["年度样本", "年内至少有一个状态为“交易”且收盘价有效的转债", "个券-年度", "状态为“新股上市”的上市首日不计入"],
  ["年度评级", "该转债当年首个正常交易日的债项评级", "评级", "同一转债每年只归入一个评级"],
  ["发行规模", "总表静态发行规模；组内求和", "亿元", "不是当年末余额"],
  ["发行规模占比", "组内发行规模合计/当年全部有效发行规模合计", "%", "总体统计中每只券只计一次"],
  ["个数占比", "组内个数/当年样本总个数", "%", "评级缺失归入未评级"],
  ["年度收益率", "当年末个正常交易日收盘价/当年首个正常交易日收盘价-1", "%", "至少两个正常交易日才计为有效收益样本"],
  ["总体收益率", "2018年以来末个正常交易日收盘价/首个正常交易日收盘价-1", "%", "未年化，不同个券持有区间不同"],
  ["简单算术平均", "组内有效个券收益率之和/有效收益样本数", "%", "可能受极端个券影响，附中位数用于核对"],
  ["数据来源", "工作区转债个券历史序列Parquet与总表", null, "计算结果可由个券年度明细复核"],
];
noteSheet.getRange(`A4:D${3 + notes.length}`).values = notes;
body(noteSheet.getRange(`A4:D${3 + notes.length}`));
noteSheet.getRange(`B4:B${3 + notes.length}`).format.wrapText = true;
noteSheet.getRange(`D4:D${3 + notes.length}`).format.wrapText = true;
noteSheet.getRange("A:A").format.columnWidth = 20;
noteSheet.getRange("B:B").format.columnWidth = 55;
noteSheet.getRange("C:C").format.columnWidth = 14;
noteSheet.getRange("D:D").format.columnWidth = 38;
noteSheet.freezePanes.freezeRows(3);

const previewOverall = await wb.render({ sheetName: "总体统计", range: "A1:S18", scale: 1.2, format: "png" });
await fs.writeFile(path.join(dir, "总体统计预览.png"), new Uint8Array(await previewOverall.arrayBuffer()));
const previewAnnual = await wb.render({ sheetName: "年度统计", range: "A1:I35", scale: 1.2, format: "png" });
await fs.writeFile(path.join(dir, "年度统计预览.png"), new Uint8Array(await previewAnnual.arrayBuffer()));

const check = await wb.inspect({ kind: "table", range: "总体统计!A5:I15", include: "values,formulas", tableMaxRows: 15, tableMaxCols: 10 });
console.log(check.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
