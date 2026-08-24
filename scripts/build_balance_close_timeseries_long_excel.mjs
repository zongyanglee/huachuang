import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = process.cwd();
const outputDir = path.join(root, "outputs", "balance_close_timeseries");
const summaryPath = path.join(outputDir, "summary.json");
const longCsvPath = path.join(outputDir, "余额与收盘价_成交额大于零清洗_长表.csv");
const xlsxPath = path.join(outputDir, "余额与收盘价时间序列_成交额大于零清洗.xlsx");

const summary = JSON.parse(await fs.readFile(summaryPath, "utf8"));
const longCsv = await fs.readFile(longCsvPath, "utf8");

const workbook = await Workbook.fromCSV(longCsv, { sheetName: "时间序列" });
const data = workbook.worksheets.getItem("时间序列");
const note = workbook.worksheets.add("说明");

data.showGridLines = false;
data.freezePanes.freezeRows(1);
data.getRange("A1:D1").format = {
  fill: "#1F4E78",
  font: { color: "#FFFFFF", bold: true },
  horizontalAlignment: "center",
};
data.getRange("A:A").format.columnWidthPx = 95;
data.getRange("B:B").format.columnWidthPx = 95;
data.getRange("C:D").format.columnWidthPx = 105;
data.getRangeByIndexes(0, 0, summary.long_row_count + 1, 1).setNumberFormat("yyyy-mm-dd");

note.showGridLines = false;
note.getRange("A1:B1").values = [["项目", "说明"]];
note.getRange("A1:B1").format = {
  fill: "#1F4E78",
  font: { color: "#FFFFFF", bold: true },
};
note.getRange("A2:B12").values = [
  ["数据来源", summary.source],
  ["源文件数量", summary.source_file_count],
  ["日期范围", `${summary.date_start} 至 ${summary.date_end}`],
  ["交易日数量", summary.date_count],
  ["转债代码数量", summary.bond_count],
  ["清洗规则", summary.cleaning_rule],
  ["Excel数据行数", summary.long_row_count],
  ["成交额>0单元格数", summary.positive_amount_cells],
  ["候选单元格数", summary.total_candidate_cells],
  ["宽表CSV", "同目录保留“余额_成交额大于零清洗.csv”和“收盘价_成交额大于零清洗.csv”"],
  ["导出日期", new Date().toISOString().slice(0, 10)],
];
note.getRange("A1:B12").format.borders = {
  preset: "all",
  style: "thin",
  color: "#D9D9D9",
};
note.getRange("A2:A12").format = {
  fill: "#D9EAF7",
  font: { bold: true, color: "#1F4E78" },
};
note.getRange("A:A").format.columnWidthPx = 160;
note.getRange("B:B").format.columnWidthPx = 620;
note.getRange("B2:B12").format.wrapText = true;

const inspect = await workbook.inspect({
  kind: "table",
  range: "时间序列!A1:D10",
  include: "values",
  tableMaxRows: 10,
  tableMaxCols: 4,
});
console.log(inspect.ndjson);

const noteInspect = await workbook.inspect({
  kind: "table",
  range: "说明!A1:B12",
  include: "values",
  tableMaxRows: 12,
  tableMaxCols: 2,
});
console.log(noteInspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
});
console.log(errors.ndjson);

for (const [sheetName, range] of [["时间序列", "A1:D20"], ["说明", "A1:B12"]]) {
  const preview = await workbook.render({
    sheetName,
    range,
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(outputDir, `${sheetName}_long_preview.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(xlsxPath);
console.log(xlsxPath);
