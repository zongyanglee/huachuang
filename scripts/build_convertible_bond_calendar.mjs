import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

function parseArgs(argv) {
  const args = {};
  for (let i = 2; i < argv.length; i += 2) {
    const key = argv[i];
    const value = argv[i + 1];
    if (!key?.startsWith("--") || value === undefined) {
      throw new Error(`参数格式错误：${key ?? ""}`);
    }
    args[key.slice(2)] = value;
  }
  for (const required of ["input", "output", "preview-dir"]) {
    if (!args[required]) throw new Error(`缺少参数 --${required}`);
  }
  return args;
}

function excelColumnName(number) {
  let result = "";
  let value = number;
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function asDate(value) {
  return value ? new Date(`${value}T00:00:00`) : null;
}

function formatTitle(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF", size: 16 },
    verticalAlignment: "center",
    horizontalAlignment: "left",
  };
  range.format.rowHeight = 30;
}

function formatHeader(sheet, rangeAddress) {
  sheet.getRange(rangeAddress).format = {
    fill: "#4472C4",
    font: { bold: true, color: "#FFFFFF" },
    verticalAlignment: "center",
    horizontalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: "#A6A6A6" },
  };
}

function formatBody(sheet, rangeAddress) {
  sheet.getRange(rangeAddress).format = {
    verticalAlignment: "center",
    borders: {
      insideHorizontal: { style: "thin", color: "#D9E2F3" },
      bottom: { style: "thin", color: "#A6A6A6" },
    },
  };
}

function addCalendarSheet(workbook, payload) {
  const sheet = workbook.worksheets.add("未来5天日历");
  sheet.showGridLines = false;
  sheet.mergeCells("A1:G1");
  sheet.getRange("A1").values = [["未来5天转债发行及上市日历"]];
  formatTitle(sheet, "A1:G1");

  sheet.mergeCells("A2:G2");
  sheet.getRange("A2").values = [[
    `统计日次日起的 ${payload.window_days} 个自然日：${payload.window_start} 至 ${payload.window_end}`,
  ]];
  sheet.getRange("A2:G2").format = {
    fill: "#D9E2F3",
    font: { color: "#44546A", italic: true },
    verticalAlignment: "center",
  };

  const headers = ["事件日期", "星期", "事件类型", "转债代码", "转债简称", "当前状态", "数据来源"];
  const values = payload.calendar.map((row) => [
    asDate(row["事件日期"]),
    row["星期"],
    row["事件类型"],
    row["转债代码"],
    row["转债简称"],
    row["当前状态"],
    row["数据来源"],
  ]);
  sheet.getRange("A4:G4").values = [headers];
  formatHeader(sheet, "A4:G4");

  const endRow = Math.max(5, values.length + 4);
  if (values.length) {
    sheet.getRange(`A5:G${endRow}`).values = values;
    formatBody(sheet, `A5:G${endRow}`);
    sheet.getRange(`A5:A${endRow}`).format.numberFormat = 'mm"\u6708"dd"\u65e5"';
    values.forEach((row, index) => {
      const eventType = row[2];
      const fill = eventType === "上市交易"
        ? "#E2F0D9"
        : eventType === "网上发行"
          ? "#FCE4D6"
          : "#F2F2F2";
      sheet.getRange(`A${index + 5}:G${index + 5}`).format.fill = fill;
    });
  }

  sheet.getRange(`A1:G${endRow}`).format.font.name = "Microsoft YaHei";
  const widths = [15, 12, 14, 15, 14, 15, 22];
  widths.forEach((width, index) => {
    const column = excelColumnName(index + 1);
    sheet.getRange(`${column}1:${column}${endRow}`).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(4);
  return sheet;
}

function addOverview(workbook, payload) {
  const sheet = workbook.worksheets.add("概览");
  sheet.showGridLines = false;
  sheet.mergeCells("A1:F1");
  sheet.getRange("A1").values = [["转债发行及上市日历摘要"]];
  formatTitle(sheet, "A1:F1");

  sheet.mergeCells("A2:F2");
  sheet.getRange("A2").values = [[
    `统计日：${payload.as_of_date} ｜ 未来区间：${payload.window_start} 至 ${payload.window_end} ｜ 存量名单：${payload.universe_source}`,
  ]];
  sheet.getRange("A2:F2").format = {
    fill: "#D9E2F3",
    font: { color: "#44546A", italic: true },
    verticalAlignment: "center",
  };
  sheet.getRange("A2:F2").format.rowHeight = 22;

  sheet.mergeCells("A4:F6");
  sheet.getRange("A4").values = [[payload.paragraph]];
  sheet.getRange("A4:F6").format = {
    fill: "#FFF2CC",
    font: { bold: true, color: "#262626", size: 12 },
    wrapText: true,
    verticalAlignment: "center",
    horizontalAlignment: "left",
    borders: { preset: "outside", style: "medium", color: "#BF9000" },
  };

  sheet.getRange("A8:B8").values = [["指标", "数量（只）"]];
  formatHeader(sheet, "A8:B8");
  sheet.getRange("A9:A10").values = [["未来5天上市"], ["未来5天发行"]];
  const calendarEnd = Math.max(5, payload.calendar.length + 4);
  sheet.getRange("B9:B10").formulas = [
    [`=COUNTIF('未来5天日历'!C5:C${calendarEnd},"上市交易")`],
    [`=COUNTIF('未来5天日历'!C5:C${calendarEnd},"网上发行")`],
  ];
  sheet.getRange("B9:B10").format.font = { color: "#008000" };
  sheet.getRange("B9:B10").format.numberFormat = "#,##0";
  formatBody(sheet, "A9:B10");

  sheet.getRange("A12:F12").values = [[
    "说明：未来5天指统计日次日起的5个自然日；发行来自 p00600，上市来自已发行未上市转债的 THS_BD 上市日期。",
    null, null, null, null, null,
  ]];
  sheet.mergeCells("A12:F12");
  sheet.getRange("A12:F12").format = {
    font: { color: "#7F6000", italic: true },
    wrapText: true,
  };

  sheet.getRange("A1:F12").format.font.name = "Microsoft YaHei";
  sheet.getRange("A1:A12").format.columnWidth = 20;
  sheet.getRange("B1:B12").format.columnWidth = 14;
  sheet.getRange("C1:F12").format.columnWidth = 16;
  sheet.freezePanes.freezeRows(2);
  return sheet;
}

function addUnlistedSheet(workbook, rows) {
  const sheet = workbook.worksheets.add("未上市转债");
  sheet.showGridLines = false;
  const headers = [
    "转债代码", "转债简称", "正股代码", "正股简称",
    "发行方式", "交易状态", "转债余额（亿元）", "上市日期",
  ];
  const values = rows.map((row) => [
    row["转债代码"],
    row["转债简称"],
    row["正股代码"],
    row["正股简称"],
    row["发行方式"],
    row["交易状态"],
    row["转债余额"],
    asDate(row["上市日期"]),
  ]);
  sheet.getRange("A1:H1").values = [headers];
  formatHeader(sheet, "A1:H1");
  if (values.length) {
    sheet.getRange(`A2:H${values.length + 1}`).values = values;
    formatBody(sheet, `A2:H${values.length + 1}`);
    sheet.getRange(`G2:G${values.length + 1}`).format.numberFormat = "0.00";
    sheet.getRange(`H2:H${values.length + 1}`).format.numberFormat = "yyyy-mm-dd";
    values.forEach((_, index) => {
      if (index % 2 === 0) {
        sheet.getRange(`A${index + 2}:H${index + 2}`).format.fill = "#DDEBF7";
      }
    });
  }
  sheet.getRange(`A1:H${Math.max(2, values.length + 1)}`).format.font.name = "Microsoft YaHei";
  const widths = [15, 14, 15, 13, 22, 15, 18, 15];
  widths.forEach((width, index) => {
    const column = excelColumnName(index + 1);
    sheet.getRange(`${column}1:${column}${Math.max(2, values.length + 1)}`).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(1);
  return sheet;
}

function addIssueSheet(workbook, rows) {
  const sheet = workbook.worksheets.add("待网上发行");
  sheet.showGridLines = false;
  const headers = ["转债代码", "转债简称", "发行日期", "事件类型"];
  const values = rows.map((row) => [
    row["转债代码"],
    row["转债简称"],
    asDate(row["发行日期"]),
    "网上发行",
  ]);
  sheet.getRange("A1:D1").values = [headers];
  formatHeader(sheet, "A1:D1");
  if (values.length) {
    sheet.getRange(`A2:D${values.length + 1}`).values = values;
    formatBody(sheet, `A2:D${values.length + 1}`);
    sheet.getRange(`C2:C${values.length + 1}`).format.numberFormat = 'mm"月"dd"日"';
    values.forEach((_, index) => {
      if (index % 2 === 0) {
        sheet.getRange(`A${index + 2}:D${index + 2}`).format.fill = "#DDEBF7";
      }
    });
  }
  const issueEnd = Math.max(2, values.length + 1);
  sheet.getRange(`A1:D${issueEnd}`).format.font.name = "Microsoft YaHei";
  sheet.getRange(`A1:A${issueEnd}`).format.columnWidth = 15;
  sheet.getRange(`B1:B${issueEnd}`).format.columnWidth = 14;
  sheet.getRange(`C1:C${issueEnd}`).format.columnWidth = 15;
  sheet.getRange(`D1:D${issueEnd}`).format.columnWidth = 14;
  sheet.freezePanes.freezeRows(1);
  return sheet;
}

async function main() {
  const args = parseArgs(process.argv);
  const payload = JSON.parse(await fs.readFile(args.input, "utf8"));
  const workbook = Workbook.create();

  addCalendarSheet(workbook, payload);
  addOverview(workbook, payload);
  addUnlistedSheet(workbook, payload.unlisted);
  addIssueSheet(workbook, payload.upcoming_issues);

  const calendarEnd = Math.max(5, payload.calendar.length + 4);
  const overviewCheck = await workbook.inspect({
    kind: "table",
    range: `未来5天日历!A1:G${calendarEnd}`,
    include: "values,formulas",
    tableMaxRows: 20,
    tableMaxCols: 7,
    maxChars: 5000,
  });
  console.log(overviewCheck.ndjson);

  const errorScan = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "转债发行日历公式错误扫描",
  });
  console.log(errorScan.ndjson);

  await fs.mkdir(args["preview-dir"], { recursive: true });
  for (const [sheetName, range] of [
    ["未来5天日历", `A1:G${calendarEnd}`],
    ["概览", "A1:F12"],
    ["未上市转债", `A1:H${Math.max(2, payload.unlisted.length + 1)}`],
    ["待网上发行", `A1:D${Math.max(2, payload.upcoming_issues.length + 1)}`],
  ]) {
    const preview = await workbook.render({ sheetName, range, scale: 1.5, format: "png" });
    const previewPath = path.join(args["preview-dir"], `${sheetName}.png`);
    await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
  }

  await fs.mkdir(path.dirname(args.output), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(args.output);

  const savedWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(args.output));
  const savedCheck = await savedWorkbook.inspect({
    kind: "table",
    range: `未来5天日历!A1:G${calendarEnd}`,
    include: "values,formulas",
    tableMaxRows: 20,
    tableMaxCols: 7,
    maxChars: 2500,
  });
  console.log(savedCheck.ndjson);

  // artifact-tool 在部分环境会留下诊断 sidecar；验证后清理，避免污染正式输出目录。
  await fs.rm(`${args.output}.inspect.ndjson`, { force: true });
  console.log(`已生成：${args.output}`);
}

await main();
