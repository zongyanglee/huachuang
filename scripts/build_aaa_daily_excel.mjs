import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs", "aaa_daily_excel");
const dataPath = path.join(outputDir, "aaa_daily_data.json");
const payload = JSON.parse(await fs.readFile(dataPath, "utf8"));

const headers = [
  "日期",
  "债项AAA个数",
  "债项AAA余额(亿元)",
  "全市场个数",
  "全市场余额(亿元)",
  "债项AAA数量占比",
  "债项AAA余额占比",
  "主体AAA个数",
  "主体AAA余额(亿元)",
  "主体AAA数量占比",
  "主体AAA余额占比",
];

const rows = payload.records.map((record) => headers.map((header) => record[header]));

const workbook = Workbook.create();
const daily = workbook.worksheets.add("日度数据");
const note = workbook.worksheets.add("说明");

daily.showGridLines = false;
daily.getRange("A1:K1").merge();
daily.getRange("A1").values = [["AAA评级转债日度数量及余额"]];
daily.getRange("A1").format = {
  fill: "#1F4E78",
  font: { color: "#FFFFFF", bold: true, size: 14 },
  horizontalAlignment: "center",
};
daily.getRange("A2:K2").merge();
daily.getRange("A2").values = [[`数据区间：${payload.date_start} 至 ${payload.date_end}；规模单位：亿元`]];
daily.getRange("A2").format = {
  fill: "#D9EAF7",
  font: { color: "#1F4E78", italic: true },
};

const data = [headers, ...rows];
daily.getRangeByIndexes(2, 0, data.length, headers.length).values = data;
const tableRange = `A3:K${data.length + 2}`;
const table = daily.tables.add(tableRange, true, "AAADailyTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;

daily.freezePanes.freezeRows(3);
daily.getRange("A3:K3").format = {
  fill: "#5B9BD5",
  font: { color: "#FFFFFF", bold: true },
  horizontalAlignment: "center",
};
daily.getRange(`A4:A${data.length + 2}`).setNumberFormat("yyyy-mm-dd");
daily.getRange(`B4:B${data.length + 2}`).setNumberFormat("0");
daily.getRange(`C4:C${data.length + 2}`).setNumberFormat("0.00");
daily.getRange(`D4:D${data.length + 2}`).setNumberFormat("0");
daily.getRange(`E4:E${data.length + 2}`).setNumberFormat("0.00");
daily.getRange(`F4:G${data.length + 2}`).setNumberFormat("0.0%");
daily.getRange(`H4:H${data.length + 2}`).setNumberFormat("0");
daily.getRange(`I4:I${data.length + 2}`).setNumberFormat("0.00");
daily.getRange(`J4:K${data.length + 2}`).setNumberFormat("0.0%");

daily.getRange("A1:A1").format.columnWidthPx = 95;
daily.getRange("B1:B1").format.columnWidthPx = 95;
daily.getRange("C1:C1").format.columnWidthPx = 130;
daily.getRange("D1:D1").format.columnWidthPx = 95;
daily.getRange("E1:E1").format.columnWidthPx = 130;
daily.getRange("F1:G1").format.columnWidthPx = 120;
daily.getRange("H1:H1").format.columnWidthPx = 95;
daily.getRange("I1:I1").format.columnWidthPx = 130;
daily.getRange("J1:K1").format.columnWidthPx = 120;
daily.getRange(`A3:K${data.length + 2}`).format.borders = {
  preset: "all",
  style: "thin",
  color: "#D9E2F3",
};

note.showGridLines = false;
note.getRange("A1:B1").values = [["项目", "说明"]];
note.getRange("A1:B1").format = {
  fill: "#1F4E78",
  font: { color: "#FFFFFF", bold: true },
};
note.getRange("A2:B7").values = [
  ["主口径", "债项评级 == AAA，且当日余额 > 0 的存续转债"],
  ["规模口径", "使用月度 parquet 中的“余额”指标求和，单位为亿元"],
  ["校验口径", "同时提供主体评级 AAA 的数量、余额和占比"],
  ["全市场口径", "当日余额 > 0 的全部转债"],
  ["数据来源", `转债个券历史序列月度 parquet，共 ${payload.source_file_count} 个文件`],
  ["导出日期", new Date().toISOString().slice(0, 10)],
];
note.getRange("A1:B1").format.columnWidthPx = 240;
note.getRange("A1:B7").format.borders = {
  preset: "all",
  style: "thin",
  color: "#D9D9D9",
};
note.getRange("A2:A7").format = {
  fill: "#D9EAF7",
  font: { bold: true, color: "#1F4E78" },
};

const preview = await workbook.render({
  sheetName: "日度数据",
  range: "A1:K20",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "aaa_daily_preview.png"),
  new Uint8Array(await preview.arrayBuffer()),
);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
});
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
const outPath = path.join(outputDir, "AAA评级转债日度数量及余额.xlsx");
await output.save(outPath);
console.log(outPath);
