import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

function parseArgs(argv) {
  const out = {};
  for (let i = 2; i < argv.length; i += 2) {
    if (!argv[i]?.startsWith("--") || argv[i + 1] === undefined) {
      throw new Error(`参数格式错误：${argv[i] ?? ""}`);
    }
    out[argv[i].slice(2)] = argv[i + 1];
  }
  for (const key of ["input", "xlsx", "png"]) {
    if (!out[key]) throw new Error(`缺少 --${key}`);
  }
  return out;
}

function isoDate(value) {
  return value ? String(value).slice(0, 10) : null;
}

function dateLabel(iso) {
  const [year, month, day] = iso.split("-").map(Number);
  const weekday = ["周日", "周一", "周二", "周三", "周四", "周五", "周六"][
    new Date(`${iso}T00:00:00`).getDay()
  ];
  return `${year}/${month}/${day}${weekday}`;
}

function nthTradingDate(tradingDates, issueDate, offset) {
  const index = tradingDates.indexOf(issueDate);
  if (index < 0) return null;
  return tradingDates[index + offset] ?? null;
}

function buildEvents(bond, tradingDates) {
  const events = new Map();
  const add = (date, label, colorKey) => {
    if (!date) return;
    const current = events.get(date) ?? [];
    current.push({ label, colorKey });
    events.set(date, current);
  };

  const issueDate = isoDate(bond["发行日期"]);
  if (issueDate) {
    add(nthTradingDate(tradingDates, issueDate, -1), "原股东股权登记日", "原股东股权登记日");
    add(issueDate, "原股东配售日，网上发行日", "原股东配售日，网上发行日");
    add(nthTradingDate(tradingDates, issueDate, 1), "中签率", "中签率");
    add(nthTradingDate(tradingDates, issueDate, 2), "中签结果", "中签结果");
  }
  add(isoDate(bond["上市日期"]), "上市", "上市");
  return events;
}

function columnName(index) {
  let result = "";
  let n = index;
  while (n > 0) {
    n -= 1;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

function applyBorders(range, color = "#000000") {
  range.format.borders = { preset: "all", style: "thin", color };
}

async function build(payload, xlsxPath, pngPath) {
  // 这里只生成一个独立的临时工作簿。主工作簿由 Excel 原生接口更新，
  // 避免第三方解析器重写其中的 Wind / THS 插件公式和 OOXML 扩展信息。
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("发行日历");
  sheet.showGridLines = false;

  const dateCount = payload.calendar_dates.length;
  const totalCols = 8 + dateCount;
  const lastCol = columnName(totalCols);
  const headerRow = 3;
  const firstDataRow = 4;
  const hasBonds = payload.bonds.length > 0;
  const lastDataRow = firstDataRow + Math.max(payload.bonds.length, 1) - 1;
  const footerRow = lastDataRow + 1;

  sheet.mergeCells("A1:A2");
  sheet.getRange("A1").values = [["华创证券\n研究"]];
  sheet.getRange("A1:A2").format = {
    fill: "#FFFFFF",
    font: { bold: true, color: "#1F4E79", size: 10 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };

  sheet.mergeCells(`B1:${lastCol}1`);
  sheet.getRange("B1").values = [[payload.title]];
  sheet.mergeCells(`B2:${lastCol}2`);
  sheet.getRange("B2").values = [[payload.subtitle]];
  sheet.getRange(`B1:${lastCol}2`).format = {
    fill: "#365F9C",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };

  const headers = [
    "网上申购代码",
    "转债代码",
    "简称",
    "所属行业",
    "发行规模\n（亿元）",
    "债项评级",
    "评级公司",
    "上市预测价\n（元）",
    ...payload.calendar_dates.map(dateLabel),
  ];
  sheet.getRange(`A${headerRow}:${lastCol}${headerRow}`).values = [headers];
  sheet.getRange(`A${headerRow}:${lastCol}${headerRow}`).format = {
    fill: "#FFFFFF",
    font: { bold: true, color: "#000000", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };

  if (hasBonds) {
    const baseRows = payload.bonds.map((bond) => {
      const predictedPrice = bond["上市价格预测V2.1"];
      const displayPrice =
        predictedPrice !== null &&
        predictedPrice !== undefined &&
        predictedPrice !== "" &&
        Number.isFinite(Number(predictedPrice))
          ? Number(predictedPrice)
          : "—";
      return [
        String(bond["网上申购代码"]),
        String(bond["转债代码"]),
        bond["简称"],
        bond["所属行业"],
        Number(bond["发行规模"]),
        bond["债项评级"],
        bond["评级公司"],
        displayPrice,
        ...payload.calendar_dates.map(() => null),
      ];
    });
    sheet.getRange(`A${firstDataRow}:${lastCol}${lastDataRow}`).values = baseRows;
    sheet.getRange(`A${firstDataRow}:${lastCol}${lastDataRow}`).format = {
      fill: "#FFFFFF",
      font: { color: "#000000", size: 11 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: true,
    };
    sheet.getRange(`E${firstDataRow}:E${lastDataRow}`).format.numberFormat = "0.00";
    sheet.getRange(`H${firstDataRow}:H${lastDataRow}`).format.numberFormat = "0.00";
  } else {
    sheet.mergeCells(`A${firstDataRow}:${lastCol}${firstDataRow}`);
    sheet.getRange(`A${firstDataRow}`).values = [["未来5个交易日暂无发行或上市安排"]];
    sheet.getRange(`A${firstDataRow}:${lastCol}${firstDataRow}`).format = {
      fill: "#FFFFFF",
      font: { color: "#666666", size: 11 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
    };
  }

  payload.bonds.forEach((bond, bondIndex) => {
    const row = firstDataRow + bondIndex;
    const events = buildEvents(bond, payload.trading_dates);
    payload.calendar_dates.forEach((date, dateIndex) => {
      const cellEvents = events.get(date) ?? [];
      if (!cellEvents.length) return;
      const cell = sheet.getCell(row - 1, 8 + dateIndex);
      cell.values = [[cellEvents.map((event) => event.label).join("，")]];
      const chosen = cellEvents.find((event) => event.colorKey === "上市") ?? cellEvents[0];
      cell.format = {
        fill: payload.colors[chosen.colorKey],
        font: { bold: chosen.colorKey === "上市", color: "#000000", size: 11 },
        horizontalAlignment: "center",
        verticalAlignment: "center",
        wrapText: true,
      };
    });
  });

  sheet.mergeCells(`A${footerRow}:${lastCol}${footerRow}`);
  const [year, month, day] = payload.updated_date.split("-").map(Number);
  sheet.getRange(`A${footerRow}`).values = [[`更新时间：${year}年${month}月${day}日`]];
  sheet.getRange(`A${footerRow}:${lastCol}${footerRow}`).format = {
    fill: "#FFFFFF",
    font: { color: "#000000", size: 10 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };

  applyBorders(sheet.getRange(`A1:${lastCol}${footerRow}`));
  const calendarRange = sheet.getRange(`A1:${lastCol}${footerRow}`);
  calendarRange.format.font.name = "KaiTi_GB2312";
  calendarRange.format.font.size = 11;
  sheet.getRange(`B1:${lastCol}1`).format.font = {
    name: "Microsoft YaHei",
    bold: true,
    color: "#FFFFFF",
    size: 14,
  };
  sheet.getRange(`B2:${lastCol}2`).format.font = {
    name: "Microsoft YaHei",
    bold: true,
    color: "#FFFFFF",
    size: 10,
  };
  sheet.getRange(`A${headerRow}:${lastCol}${headerRow}`).format.font.bold = true;

  for (let index = 0; index < totalCols; index += 1) {
    const column = columnName(index + 1);
    sheet.getRange(`${column}1:${column}${footerRow}`).format.columnWidth = 14;
  }
  sheet.getRange(`A1:${lastCol}${footerRow}`).format.rowHeight = 30;

  const check = await workbook.inspect({
    kind: "table",
    range: `发行日历!A1:${lastCol}${footerRow}`,
    include: "values,formulas",
    tableMaxRows: 15,
    tableMaxCols: 15,
    maxChars: 8000,
  });
  console.log(check.ndjson);
  const errors = await workbook.inspect({
    kind: "match",
    sheetId: "发行日历",
    range: `A1:${lastCol}${footerRow}`,
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "发行日历公式错误扫描",
  });
  console.log(errors.ndjson);

  await fs.mkdir(path.dirname(pngPath), { recursive: true });
  const preview = await workbook.render({
    sheetName: "发行日历",
    range: `A1:${lastCol}${footerRow}`,
    scale: 1.6,
    format: "png",
    headers: false,
  });
  await fs.writeFile(pngPath, new Uint8Array(await preview.arrayBuffer()));

  await fs.mkdir(path.dirname(xlsxPath), { recursive: true });
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(xlsxPath);

  const saved = await SpreadsheetFile.importXlsx(await FileBlob.load(xlsxPath));
  const savedCheck = await saved.inspect({
    kind: "table",
    range: `发行日历!A1:${lastCol}${footerRow}`,
    include: "values,formulas",
    tableMaxRows: 15,
    tableMaxCols: 15,
    maxChars: 5000,
  });
  console.log(savedCheck.ndjson);
  await fs.rm(`${xlsxPath}.inspect.ndjson`, { force: true });
}

const args = parseArgs(process.argv);
const payload = JSON.parse(await fs.readFile(args.input, "utf8"));
await build(payload, args.xlsx, args.png);
console.log(`已生成图片：${args.png}`);
