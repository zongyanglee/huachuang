import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";


function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 2) {
    const key = argv[index];
    const value = argv[index + 1];
    if (!key?.startsWith("--") || value === undefined) {
      throw new Error(`参数格式错误：${key ?? ""}`);
    }
    result[key.slice(2)] = value;
  }
  if (!result.input || !result.output) {
    throw new Error("必须提供 --input 和 --output");
  }
  return result;
}


function columnName(columnNumber) {
  let number = columnNumber;
  let name = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    number = Math.floor((number - 1) / 26);
  }
  return name;
}


function buildHeaders() {
  const headers = ["查询日期", "转债代码", "转债简称", "正股代码", "正股简称", "剩余期限（年）"];
  for (let rank = 1; rank <= 10; rank += 1) {
    headers.push(`NO${rank}大股东`, `NO${rank}持股比例（%）`, `NO${rank}持债情况`);
  }
  return headers;
}


function recordsToMatrix(records, headers) {
  return [headers, ...records.map((record) => headers.map((header) => record[header] ?? null))];
}


async function main() {
  const args = parseArgs(process.argv.slice(2));
  const payload = JSON.parse(await fs.readFile(args.input, "utf8"));
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("5-5.5年转债");
  const headers = buildHeaders();
  const records = payload.records || [];
  const matrix = recordsToMatrix(records, headers);
  const lastColumn = columnName(headers.length);
  const lastRow = Math.max(1, matrix.length);

  sheet.getRangeByIndexes(0, 0, matrix.length, headers.length).values = matrix;
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.freezePanes.freezeColumns(6);

  const used = sheet.getRange(`A1:${lastColumn}${lastRow}`);
  used.format.font = { name: "Microsoft YaHei", size: 9, color: "#1F2937" };
  used.format.verticalAlignment = "center";
  used.format.borders = {
    insideHorizontal: { style: "thin", color: "#E5E7EB" },
    bottom: { style: "thin", color: "#CBD5E1" },
  };

  const header = sheet.getRange(`A1:${lastColumn}1`);
  header.format = {
    fill: "#17365D",
    font: { name: "Microsoft YaHei", size: 9, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeight: 42,
    borders: { preset: "all", style: "thin", color: "#FFFFFF" },
  };

  sheet.tables.add(`A1:${lastColumn}${lastRow}`, true, "MaturityMajorHolderTable");
  sheet.getRange("A1:F1").format.fill = "#17365D";
  for (let rank = 1; rank <= 10; rank += 1) {
    const firstColumn = 7 + (rank - 1) * 3;
    const lastGroupColumn = firstColumn + 2;
    const groupRange = sheet.getRange(
      `${columnName(firstColumn)}1:${columnName(lastGroupColumn)}1`,
    );
    groupRange.format.fill = rank % 2 === 1 ? "#4472C4" : "#5B9BD5";
  }

  const baseWidths = [13, 15, 14, 15, 14, 15];
  baseWidths.forEach((width, index) => {
    sheet.getRange(`${columnName(index + 1)}:${columnName(index + 1)}`).format.columnWidth = width;
  });
  for (let rank = 1; rank <= 10; rank += 1) {
    const firstColumn = 7 + (rank - 1) * 3;
    sheet.getRange(`${columnName(firstColumn)}:${columnName(firstColumn)}`).format.columnWidth = 31;
    sheet.getRange(`${columnName(firstColumn + 1)}:${columnName(firstColumn + 1)}`).format.columnWidth = 18;
    sheet.getRange(`${columnName(firstColumn + 2)}:${columnName(firstColumn + 2)}`).format.columnWidth = 27;
  }

  if (records.length > 0) {
    sheet.getRange(`A2:${lastColumn}${lastRow}`).format.rowHeight = 34;
    sheet.getRange(`F2:F${lastRow}`).format.numberFormat = "0.000";
    sheet.getRange(`F2:F${lastRow}`).format.horizontalAlignment = "right";
    for (let rank = 1; rank <= 10; rank += 1) {
      const firstColumn = 7 + (rank - 1) * 3;
      const nameColumn = columnName(firstColumn);
      const stockRatioColumn = columnName(firstColumn + 1);
      const holdingColumn = columnName(firstColumn + 2);
      sheet.getRange(`${nameColumn}2:${nameColumn}${lastRow}`).format.wrapText = true;
      sheet.getRange(`${stockRatioColumn}2:${stockRatioColumn}${lastRow}`).format.numberFormat = "0.00";
      sheet.getRange(`${stockRatioColumn}2:${stockRatioColumn}${lastRow}`).format.horizontalAlignment = "right";
      sheet.getRange(`${holdingColumn}2:${holdingColumn}${lastRow}`).format.wrapText = true;
      const holdingRange = sheet.getRange(`${holdingColumn}2:${holdingColumn}${lastRow}`);
      holdingRange.conditionalFormats.add("beginsWith", {
        text: "转债第",
        format: { fill: "#E2F0D9", font: { color: "#006100", bold: true } },
      });
      holdingRange.conditionalFormats.add("containsText", {
        text: "未进入",
        format: { fill: "#FFF2CC", font: { color: "#9C6500" } },
      });
    }
  }

  await fs.mkdir(path.dirname(args.output), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(args.output);

  const inspection = await workbook.inspect({
    kind: "table",
    range: `5-5.5年转债!A1:${lastColumn}${Math.min(lastRow, 6)}`,
    include: "values,formulas",
    tableMaxRows: 6,
    tableMaxCols: headers.length,
    maxChars: 12000,
  });
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
    summary: "final formula error scan",
    maxChars: 3000,
  });
  console.log(inspection.ndjson);
  console.log(errors.ndjson);

  if (args["qa-dir"]) {
    await fs.mkdir(args["qa-dir"], { recursive: true });
    const previewRanges = [
      ["NO1-NO5", `A1:U${Math.min(lastRow, 12)}`],
      ["NO6-NO10", `V1:${lastColumn}${Math.min(lastRow, 12)}`],
    ];
    for (const [name, range] of previewRanges) {
      const preview = await workbook.render({ sheetName: sheet.name, range, scale: 1.35, format: "png" });
      await fs.writeFile(
        path.join(args["qa-dir"], `${name}.png`),
        new Uint8Array(await preview.arrayBuffer()),
      );
    }
  }
}


await main();

