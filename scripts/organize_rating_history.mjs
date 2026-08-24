import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "D:/JupyterFiles/huachuang";
const inputPath = path.join(root, "历史评级.xlsx");
const outputDir = path.join(root, "outputs", "rating_history_changes_20260709");
const outputPath = path.join(outputDir, "历史评级调整明细.xlsx");
const datasetPath = path.join(outputDir, "rating_history_dataset.json");

const sheets = [
  {
    name: "摘要",
    tableName: null,
    headers: ["项目", "数值"],
  },
  {
    name: "评级调整明细",
    tableName: "RatingChangeTable",
    headers: [
      "转债代码",
      "转债简称",
      "网上发行日期",
      "评级日期",
      "调整类型",
      "调整前评级",
      "调整后评级",
      "上次评级日期",
      "评级类型",
      "评级机构",
      "原始记录",
      "源表行号",
    ],
  },
  {
    name: "全部评级记录",
    tableName: "AllRatingEventsTable",
    headers: [
      "转债代码",
      "转债简称",
      "网上发行日期",
      "评级日期",
      "序号",
      "是否评级变化",
      "调整类型",
      "调整前评级",
      "调整后评级",
      "上次评级日期",
      "评级类型",
      "评级机构",
      "原始记录",
      "源表行号",
    ],
  },
  {
    name: "未解析或无缓存",
    tableName: "RatingIssueTable",
    headers: ["源表行号", "转债代码", "转债简称", "网上发行日期", "问题", "原始内容"],
  },
];

function excelSerialToDate(serial) {
  if (typeof serial !== "number" || !Number.isFinite(serial)) return "";
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + serial * 24 * 60 * 60 * 1000);
}

function dateStringToDate(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, day));
}

function colLetter(index) {
  let n = index + 1;
  let result = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function tableRange(rowCount, colCount) {
  return `A1:${colLetter(colCount - 1)}${Math.max(rowCount, 1)}`;
}

function changeRow(row) {
  return [
    row.bondCode,
    row.bondName,
    excelSerialToDate(row.issueDateSerial),
    dateStringToDate(row.ratingDate),
    row.changeType,
    row.previousRating,
    row.currentRating,
    dateStringToDate(row.previousRatingDate),
    row.ratingType,
    row.agency,
    row.rawRecord,
    row.sourceRow,
  ];
}

function allEventRow(row) {
  return [
    row.bondCode,
    row.bondName,
    excelSerialToDate(row.issueDateSerial),
    dateStringToDate(row.ratingDate),
    row.sequence,
    row.isRatingChange,
    row.changeType,
    row.previousRating,
    row.currentRating,
    dateStringToDate(row.previousRatingDate),
    row.ratingType,
    row.agency,
    row.rawRecord,
    row.sourceRow,
  ];
}

function issueRow(row) {
  return [
    row.sourceRow,
    row.bondCode,
    row.bondName,
    excelSerialToDate(row.issueDateSerial),
    row.issue,
    row.raw,
  ];
}

function applyTableStyle(sheet, rowCount, colCount, dateCols = []) {
  const range = sheet.getRange(tableRange(rowCount, colCount));
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.getRange(`A1:${colLetter(colCount - 1)}1`).format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF" },
  };
  range.format.borders = {
    insideHorizontal: { style: "thin", color: "#E5E7EB" },
    top: { style: "thin", color: "#BFC7D5" },
    bottom: { style: "thin", color: "#BFC7D5" },
  };
  range.format.autofitColumns();
  range.format.autofitRows();
  dateCols.forEach((col) => {
    sheet.getRange(`${col}2:${col}${rowCount}`).setNumberFormat("yyyy-mm-dd");
  });
}

function setColumnWidths(sheet, widths) {
  Object.entries(widths).forEach(([col, width]) => {
    sheet.getRange(`${col}1:${col}1`).format.columnWidth = width;
  });
}

async function main() {
  await fs.mkdir(outputDir, { recursive: true });

  const dataset = JSON.parse(await fs.readFile(datasetPath, "utf8"));
  const workbook = Workbook.create();

  for (const sheetSpec of sheets) {
    workbook.worksheets.add(sheetSpec.name);
  }

  const summarySheet = workbook.worksheets.getItem("摘要");
  const summaryRows = [
    ["项目", "数值"],
    ["源文件", "历史评级.xlsx"],
    ["源表数据行数", dataset.summary.sourceRows],
    ["有评级缓存的转债数量", dataset.summary.bondsWithCachedRating],
    ["全部评级事件数", dataset.summary.allEventRows],
    ["实际评级调整行数（不含首次评级）", dataset.summary.actualChangeRows],
    ["首次评级及调整行数", dataset.summary.changeRowsIncludingFirstRating],
    ["未解析或无缓存行数", dataset.summary.issueRows],
    ["整理口径", "主表仅保留前后评级不同的真实调整；全部评级记录表保留首次评级和评级不变的跟踪记录。"],
  ];
  summarySheet.getRange(`A1:B${summaryRows.length}`).values = summaryRows;
  applyTableStyle(summarySheet, summaryRows.length, 2);
  setColumnWidths(summarySheet, { A: 28, B: 110 });
  summarySheet.getRange("B9:B9").format.wrapText = true;

  const changeSheet = workbook.worksheets.getItem("评级调整明细");
  const actualChanges = dataset.changes.filter((row) => row.previousRating);
  const changeValues = [sheets[1].headers, ...actualChanges.map(changeRow)];
  changeSheet.getRange(tableRange(changeValues.length, sheets[1].headers.length)).values = changeValues;
  changeSheet.tables.add(tableRange(changeValues.length, sheets[1].headers.length), true, sheets[1].tableName);
  applyTableStyle(changeSheet, changeValues.length, sheets[1].headers.length, ["C", "D", "H"]);
  setColumnWidths(changeSheet, {
    A: 14,
    B: 18,
    C: 14,
    D: 14,
    E: 14,
    F: 12,
    G: 12,
    H: 14,
    I: 16,
    J: 38,
    K: 58,
    L: 10,
  });

  const allSheet = workbook.worksheets.getItem("全部评级记录");
  const allValues = [sheets[2].headers, ...dataset.allEvents.map(allEventRow)];
  allSheet.getRange(tableRange(allValues.length, sheets[2].headers.length)).values = allValues;
  allSheet.tables.add(tableRange(allValues.length, sheets[2].headers.length), true, sheets[2].tableName);
  applyTableStyle(allSheet, allValues.length, sheets[2].headers.length, ["C", "D", "J"]);
  setColumnWidths(allSheet, {
    A: 14,
    B: 18,
    C: 14,
    D: 14,
    E: 8,
    F: 12,
    G: 14,
    H: 12,
    I: 12,
    J: 14,
    K: 16,
    L: 38,
    M: 58,
    N: 10,
  });

  const issueSheet = workbook.worksheets.getItem("未解析或无缓存");
  const issueValues = [sheets[3].headers, ...dataset.issues.map(issueRow)];
  issueSheet.getRange(tableRange(issueValues.length, sheets[3].headers.length)).values = issueValues;
  issueSheet.tables.add(tableRange(issueValues.length, sheets[3].headers.length), true, sheets[3].tableName);
  applyTableStyle(issueSheet, issueValues.length, sheets[3].headers.length, ["D"]);
  setColumnWidths(issueSheet, { A: 10, B: 14, C: 18, D: 14, E: 24, F: 58 });

  const check = await workbook.inspect({
    kind: "table",
    maxChars: 8000,
    tableMaxRows: 8,
    tableMaxCols: 12,
    tableMaxCellChars: 100,
  });
  console.log(check.ndjson);

  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan",
  });
  console.log(errors.ndjson);

  for (const sheetName of ["摘要", "评级调整明细", "全部评级记录", "未解析或无缓存"]) {
    const preview = await workbook.render({
      sheetName,
      range: sheetName === "摘要" ? "A1:B9" : "A1:L30",
      scale: 1,
      format: "png",
    });
    await fs.writeFile(
      path.join(outputDir, `${sheetName}.png`),
      new Uint8Array(await preview.arrayBuffer()),
    );
  }

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
  console.log(`saved: ${outputPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
