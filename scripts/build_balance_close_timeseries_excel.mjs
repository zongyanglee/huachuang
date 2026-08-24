import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = process.cwd();
const outputDir = path.join(root, "outputs", "balance_close_timeseries");
const summaryPath = path.join(outputDir, "summary.json");
const balanceCsvPath = path.join(outputDir, "余额_成交额大于零清洗.csv");
const closeCsvPath = path.join(outputDir, "收盘价_成交额大于零清洗.csv");
const xlsxPath = path.join(outputDir, "余额与收盘价时间序列_成交额大于零清洗.xlsx");

const summary = JSON.parse(await fs.readFile(summaryPath, "utf8"));
const balanceCsv = await fs.readFile(balanceCsvPath, "utf8");
const closeCsv = await fs.readFile(closeCsvPath, "utf8");

const workbook = await Workbook.fromCSV(balanceCsv, { sheetName: "余额" });
await workbook.fromCSV(closeCsv, { sheetName: "收盘价" });
const note = workbook.worksheets.add("说明");

for (const sheetName of ["余额", "收盘价"]) {
  const sheet = workbook.worksheets.getItem(sheetName);
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(1);
  sheet.getRangeByIndexes(0, 0, 1, summary.bond_count + 1).format = {
    fill: "#1F4E78",
    font: { color: "#FFFFFF", bold: true },
    horizontalAlignment: "center",
  };
  sheet.getRangeByIndexes(0, 0, summary.date_count + 1, 1).setNumberFormat("yyyy-mm-dd");
  sheet.getRange("A:A").format.columnWidthPx = 95;
  if (summary.bond_count > 0) {
    sheet.getRangeByIndexes(0, 1, 1, summary.bond_count).format.columnWidthPx = 78;
  }
}

note.showGridLines = false;
note.getRange("A1:B1").values = [["项目", "说明"]];
note.getRange("A1:B1").format = {
  fill: "#1F4E78",
  font: { color: "#FFFFFF", bold: true },
};
note.getRange("A2:B11").values = [
  ["数据来源", summary.source],
  ["源文件数量", summary.source_file_count],
  ["日期范围", `${summary.date_start} 至 ${summary.date_end}`],
  ["交易日数量", summary.date_count],
  ["转债代码数量", summary.bond_count],
  ["清洗规则", summary.cleaning_rule],
  ["成交额>0单元格数", summary.positive_amount_cells],
  ["候选单元格数", summary.total_candidate_cells],
  ["余额有效单元格数", summary.balance_non_null_cells],
  ["收盘价有效单元格数", summary.close_non_null_cells],
];
note.getRange("A1:B11").format.borders = {
  preset: "all",
  style: "thin",
  color: "#D9D9D9",
};
note.getRange("A2:A11").format = {
  fill: "#D9EAF7",
  font: { bold: true, color: "#1F4E78" },
};
note.getRange("A:A").format.columnWidthPx = 160;
note.getRange("B:B").format.columnWidthPx = 620;
note.getRange("B2:B11").format.wrapText = true;

const inspect = await workbook.inspect({
  kind: "table",
  range: "说明!A1:B11",
  include: "values",
  tableMaxRows: 12,
  tableMaxCols: 2,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
});
console.log(errors.ndjson);

for (const sheetName of ["余额", "收盘价", "说明"]) {
  const range = sheetName === "说明" ? "A1:B11" : "A1:J20";
  const preview = await workbook.render({
    sheetName,
    range,
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(outputDir, `${sheetName}_preview.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(xlsxPath);
console.log(xlsxPath);
