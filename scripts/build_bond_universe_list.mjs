import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = process.cwd();
const dir = path.join(root, "outputs", "bond_universe_since_2017");
const data = JSON.parse(await fs.readFile(path.join(dir, "universe.json"), "utf8"));
const outputPath = path.join(dir, "2017年以来全部存续转债列表.xlsx");

const wb = Workbook.create();
const sheet = wb.worksheets.add("转债列表");
sheet.showGridLines = false;

const headers = data.columns;
const dateColumns = new Set(["上市日期", "最后交易日", "发行日期", "转股期起始日", "赎回公告日", "最新数据日"]);
const rows = data.rows.map((row) => headers.map((header) => {
  const value = row[header];
  if (value == null) return null;
  if (dateColumns.has(header)) {
    const parsed = new Date(String(value).slice(0, 10));
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
  return value;
}));

sheet.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
sheet.getRangeByIndexes(1, 0, rows.length, headers.length).values = rows;

const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
headerRange.format = {
  fill: "#17365D",
  font: { name: "Microsoft YaHei", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  rowHeight: 32,
  borders: { preset: "all", style: "thin", color: "#D9E1F2" },
};
const bodyRange = sheet.getRangeByIndexes(1, 0, rows.length, headers.length);
bodyRange.format = {
  font: { name: "Microsoft YaHei", size: 9, color: "#1F2937" },
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#E7E6E6" },
};

for (const header of dateColumns) {
  const index = headers.indexOf(header);
  if (index >= 0) sheet.getRangeByIndexes(1, index, rows.length, 1).format.numberFormat = "yyyy-mm-dd";
}
for (const header of ["发行规模", "最新余额", "最新收盘价"]) {
  const index = headers.indexOf(header);
  if (index >= 0) sheet.getRangeByIndexes(1, index, rows.length, 1).format.numberFormat = "0.0000";
}

headers.forEach((header, index) => {
  let width = 12;
  if (["转债代码", "转债名称", "申万行业", "最新交易状态"].includes(header)) width = 14;
  if (dateColumns.has(header)) width = 13;
  if (["起始日期说明", "备注"].includes(header)) width = 30;
  sheet.getRangeByIndexes(0, index, rows.length + 1, 1).format.columnWidth = width;
});

sheet.tables.add(`A1:Y${rows.length + 1}`, true, "BondUniverseTable").style = "TableStyleMedium2";
sheet.freezePanes.freezeRows(1);
sheet.freezePanes.freezeColumns(2);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
