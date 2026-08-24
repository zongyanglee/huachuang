# -*- coding: utf-8 -*-
"""华创固收转债日报：查询市场数据、读取转债 Parquet、绘图并生成 Excel 底稿。"""

from __future__ import annotations

import argparse
from configparser import ConfigParser
import json
import math
import os
import subprocess
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib import ticker as mticker
import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from iFinDPy import THS_DR, THS_DS, THS_DataStatistics, THS_GetErrorInfo, THS_iFinDLogin


WORKSPACE = Path(__file__).resolve().parents[2]
FONT_PATH = WORKSPACE / "assets/fonts/KaiTi_GB2312.ttf"
TITLE_FONT_PATH = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "simhei.ttf"
CHART_FONT_PATH = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "STKAITI.TTF"
REPORT_HEADER_PATH = WORKSPACE / "assets/images/条款表头.png"
CB_PARQUET_ROOT = WORKSPACE / "data/转债个券历史序列"
CB_MASTER_PARQUET = CB_PARQUET_ROOT / "_special" / "总表.parquet"
INDEX_PARQUET = CB_PARQUET_ROOT / "_special" / "指数.parquet"
CODEX_DEPENDENCIES = (
    Path.home()
    / ".cache"
    / "codex-runtimes"
    / "codex-primary-runtime"
    / "dependencies"
)
BUNDLED_NODE = CODEX_DEPENDENCIES / "node" / "bin" / "node.exe"
BUNDLED_NODE_MODULES = CODEX_DEPENDENCIES / "node" / "node_modules"
MARGIN_BALANCE_START_DATE = date(2024, 1, 1)

WORKBOOK_BUILDER_SOURCE = r'''import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";


const [, , payloadPath, outputPath, previewDirArg] = process.argv;
if (!payloadPath || !outputPath) {
  throw new Error("用法：node build_daily_market_workbook.mjs <payload.json> <output.xlsx>");
}

const payload = JSON.parse(await fs.readFile(payloadPath, "utf8"));
const workbook = Workbook.create();

const REPORT_RED = "#963634";
const REPORT_BLUE = "#0262BA";
const LIGHT_BLUE = "#DCE6F1";
const GRID = "#D9D9D9";
const NEGATIVE_RED = "#FF0000";
const BODY_FONT = "KaiTi_GB2312";
const TITLE_FONT = "SimHei";


function excelDate(value) {
  return new Date(`${value}T00:00:00`);
}


function styleSheet(sheet, titleRange, noteRange, headerRange, dataRange) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(3);

  titleRange.format = {
    fill: REPORT_RED,
    font: { name: TITLE_FONT, bold: true, color: "#FFFFFF", size: 16 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  titleRange.format.rowHeight = 30;

  noteRange.format = {
    fill: LIGHT_BLUE,
    font: { name: BODY_FONT, color: "#404040", size: 9 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    wrapText: true,
  };
  noteRange.format.rowHeight = 38;

  headerRange.format = {
    fill: REPORT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
  };
  headerRange.format.rowHeight = 24;

  dataRange.format = {
    font: { name: BODY_FONT, color: "#000000", size: 10 },
    borders: {
      insideHorizontal: { style: "thin", color: GRID },
      bottom: { style: "thin", color: GRID },
    },
    verticalAlignment: "center",
  };
  dataRange.format.rowHeight = 20;
}


function addMarketSheet() {
  const sheet = workbook.worksheets.add("两融余额");
  sheet.getRange("A1:B1").merge();
  sheet.getRange("A1").values = [["沪深两市融资融券余额"]];
  sheet.getRange("A2:B2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：同花顺 iFinD（p03438）｜查询区间：${payload.marketStartDate} 至 ${payload.runDate}｜最新数据：${payload.marketLatestDate}｜单位：亿元`,
  ]];
  sheet.getRange("A3:B3").values = [["交易日期", "沪深两市融资融券余额（亿元）"]];
  const rows = payload.market.map((row) => [excelDate(row.date), row.balance]);
  sheet.getRangeByIndexes(3, 0, rows.length, 2).values = rows;

  const lastRow = rows.length + 3;
  const dataRange = sheet.getRange(`A4:B${lastRow}`);
  styleSheet(
    sheet,
    sheet.getRange("A1:B1"),
    sheet.getRange("A2:B2"),
    sheet.getRange("A3:B3"),
    dataRange,
  );
  sheet.getRange(`A4:A${lastRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`B4:B${lastRow}`).format.numberFormat = "#,##0.00";
  sheet.getRange(`A4:A${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B4:B${lastRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`A1:A${lastRow}`).format.columnWidth = 16;
  sheet.getRange(`B1:B${lastRow}`).format.columnWidth = 34;
}


function addIndexSheet() {
  const sheet = workbook.worksheets.add("指数成交额");
  sheet.getRange("A1:E1").merge();
  sheet.getRange("A1").values = [["中证转债与沪深两市成交额"]];
  sheet.getRange("A2:E2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：同花顺 iFinD（ths_trans_amt_index）｜查询区间：${payload.indexStartDate} 至 ${payload.runDate}｜最新数据：${payload.indexLatestDate}｜单位：亿元`,
  ]];
  sheet.getRange("A3:E3").values = [[
    "交易日期",
    "中证转债指数成交额（亿元）",
    "上证指数成交额（亿元）",
    "深证成指成交额（亿元）",
    "沪深成交额合计（亿元）",
  ]];
  const rows = payload.index.map((row) => [
    excelDate(row.date),
    row.convertibleBond,
    row.shanghai,
    row.shenzhen,
    row.total,
  ]);
  sheet.getRangeByIndexes(3, 0, rows.length, 5).values = rows;

  const lastRow = rows.length + 3;
  const dataRange = sheet.getRange(`A4:E${lastRow}`);
  styleSheet(
    sheet,
    sheet.getRange("A1:E1"),
    sheet.getRange("A2:E2"),
    sheet.getRange("A3:E3"),
    dataRange,
  );
  sheet.getRange(`A4:A${lastRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`B4:E${lastRow}`).format.numberFormat = "#,##0.00";
  sheet.getRange(`A4:A${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B4:E${lastRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`A1:A${lastRow}`).format.columnWidth = 16;
  sheet.getRange(`B1:E${lastRow}`).format.columnWidth = 25;
}


function addReturnDistributionSheet() {
  const sheet = workbook.worksheets.add("涨跌分布");
  const summary = payload.returnSummary;
  const distributionRows = payload.returnDistribution.map((row) => [row.bucket, row.count]);
  const detailRows = payload.returnDetails.map((row) => [
    row.code,
    row.name,
    excelDate(row.previousDate),
    row.previousClose,
    excelDate(row.currentDate),
    row.currentClose,
    null,
    row.bucket,
    row.tradingStatus,
  ]);
  const lastDistributionRow = 7 + distributionRows.length;
  const lastDetailRow = 7 + detailRows.length;

  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(7);

  sheet.getRange("A1:L1").merge();
  sheet.getRange("A1").values = [[`可转债当日涨跌幅分布（${payload.runDate}）`]];
  sheet.getRange("A1:L1").format = {
    fill: REPORT_RED,
    font: { name: TITLE_FONT, bold: true, color: "#FFFFFF", size: 16 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  sheet.getRange("A1:L1").format.rowHeight = 30;

  sheet.getRange("A2:L2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：${payload.returnSource.currentParquet}｜口径：${payload.returnSource.sampleRule}｜公式：${payload.returnSource.returnFormula}`,
  ]];
  sheet.getRange("A2:L2").format = {
    fill: LIGHT_BLUE,
    font: { name: BODY_FONT, color: "#404040", size: 9 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    wrapText: true,
  };
  sheet.getRange("A2:L2").format.rowHeight = 38;

  sheet.getRange("A3:B3").values = [["上涨（只）", "下跌（只）"]];
  sheet.getRange("D3:E3").values = [["平盘（只）", "有效样本（只）"]];
  sheet.getRange("A4:B4").values = [[summary["上涨"], summary["下跌"]]];
  sheet.getRange("D4:E4").values = [[summary["平盘"], summary["有效样本"]]];
  sheet.getRange("A3").format.fill = REPORT_RED;
  sheet.getRange("B3").format.fill = REPORT_BLUE;
  sheet.getRange("D3").format.fill = "#A6A6A6";
  sheet.getRange("E3").format.fill = "#595959";
  for (const rangeAddress of ["A3:B3", "D3:E3"]) {
    sheet.getRange(rangeAddress).format.font = {
      name: BODY_FONT,
      bold: true,
      color: "#FFFFFF",
      size: 11,
    };
    sheet.getRange(rangeAddress).format.horizontalAlignment = "center";
    sheet.getRange(rangeAddress).format.verticalAlignment = "center";
    sheet.getRange(rangeAddress).format.rowHeight = 24;
  }
  for (const rangeAddress of ["A4:B4", "D4:E4"]) {
    sheet.getRange(rangeAddress).format = {
      fill: "#F2F2F2",
      font: { name: BODY_FONT, bold: true, color: "#000000", size: 14 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      numberFormat: "#,##0",
      borders: { preset: "outside", style: "thin", color: GRID },
    };
    sheet.getRange(rangeAddress).format.rowHeight = 28;
  }

  sheet.getRange("A6:B6").merge();
  sheet.getRange("A6").values = [["涨跌幅区间统计"]];
  sheet.getRange("D6:L6").merge();
  sheet.getRange("D6").values = [["个券明细"]];
  for (const rangeAddress of ["A6:B6", "D6:L6"]) {
    sheet.getRange(rangeAddress).format = {
      fill: REPORT_RED,
      font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
      horizontalAlignment: "left",
      verticalAlignment: "center",
    };
    sheet.getRange(rangeAddress).format.rowHeight = 23;
  }

  sheet.getRange("A7:B7").values = [["涨跌幅区间", "转债数量（只）"]];
  sheet.getRange("D7:L7").values = [[
    "转债代码",
    "转债简称",
    "前收盘日期",
    "前收盘价",
    "当日日期",
    "当日收盘价",
    "当日涨跌幅（%）",
    "涨跌幅区间",
    "交易状态",
  ]];
  for (const rangeAddress of ["A7:B7", "D7:L7"]) {
    sheet.getRange(rangeAddress).format = {
      fill: REPORT_BLUE,
      font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 10 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
    };
    sheet.getRange(rangeAddress).format.rowHeight = 24;
  }

  sheet.getRangeByIndexes(7, 0, distributionRows.length, 2).values = distributionRows;
  sheet.getRangeByIndexes(7, 3, detailRows.length, 9).values = detailRows;
  sheet.getRange("J8").formulasR1C1 = [["=IFERROR((RC[-1]/RC[-3]-1)*100,\"\")"]];
  sheet.getRange(`J8:J${lastDetailRow}`).fillDown();
  for (const rangeAddress of [`A8:B${lastDistributionRow}`, `D8:L${lastDetailRow}`]) {
    sheet.getRange(rangeAddress).format = {
      font: { name: BODY_FONT, color: "#000000", size: 10 },
      borders: {
        insideHorizontal: { style: "thin", color: GRID },
        bottom: { style: "thin", color: GRID },
      },
      verticalAlignment: "center",
    };
    sheet.getRange(rangeAddress).format.rowHeight = 20;
  }
  sheet.getRange(`A8:A${lastDistributionRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B8:B${lastDistributionRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`B8:B${lastDistributionRow}`).format.numberFormat = "#,##0";
  sheet.getRange(`D8:E${lastDetailRow}`).format.horizontalAlignment = "left";
  sheet.getRange(`F8:F${lastDetailRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`H8:H${lastDetailRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`F8:F${lastDetailRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`H8:H${lastDetailRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`G8:G${lastDetailRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`I8:J${lastDetailRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`G8:G${lastDetailRow}`).format.numberFormat = "0.000";
  sheet.getRange(`I8:I${lastDetailRow}`).format.numberFormat = "0.000";
  sheet.getRange(`J8:J${lastDetailRow}`).format.numberFormat = "0.00";
  sheet.getRange(`K8:L${lastDetailRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`J8:J${lastDetailRow}`).conditionalFormats.add("cellIs", {
    operator: "greaterThan",
    formula: 0,
    format: { font: { color: REPORT_RED } },
  });
  sheet.getRange(`J8:J${lastDetailRow}`).conditionalFormats.add("cellIs", {
    operator: "lessThan",
    formula: 0,
    format: { font: { color: REPORT_BLUE } },
  });

  sheet.getRange(`A1:A${lastDistributionRow}`).format.columnWidth = 15;
  sheet.getRange(`B1:B${lastDistributionRow}`).format.columnWidth = 14;
  sheet.getRange("C1:C7").format.columnWidth = 3;
  sheet.getRange(`D1:D${lastDetailRow}`).format.columnWidth = 16;
  sheet.getRange(`E1:E${lastDetailRow}`).format.columnWidth = 18;
  sheet.getRange(`F1:F${lastDetailRow}`).format.columnWidth = 14;
  sheet.getRange(`G1:G${lastDetailRow}`).format.columnWidth = 13;
  sheet.getRange(`H1:H${lastDetailRow}`).format.columnWidth = 14;
  sheet.getRange(`I1:I${lastDetailRow}`).format.columnWidth = 13;
  sheet.getRange(`J1:J${lastDetailRow}`).format.columnWidth = 18;
  sheet.getRange(`K1:K${lastDetailRow}`).format.columnWidth = 15;
  sheet.getRange(`L1:L${lastDetailRow}`).format.columnWidth = 12;
}


function addIndexPerformanceSheet() {
  const sheet = workbook.worksheets.add("指数表现");
  const mainRows = payload.indexPerformance.filter((row) => row.group === "主要指数");
  const styleRows = payload.indexPerformance.filter((row) => row.group === "风格指数");
  if (mainRows.length !== 9 || styleRows.length !== 9) {
    throw new Error(`指数表现分组数量异常：主要指数${mainRows.length}，风格指数${styleRows.length}`);
  }

  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(2);

  sheet.getRange("A1:G1").merge();
  sheet.getRange("I1:O1").merge();
  sheet.getRange("A1").values = [["主要指数"]];
  sheet.getRange("I1").values = [["风格指数"]];
  for (const rangeAddress of ["A1:G1", "I1:O1"]) {
    sheet.getRange(rangeAddress).format = {
      fill: "#D9E2F3",
      font: { name: TITLE_FONT, bold: true, color: "#000000", size: 15 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      borders: { preset: "all", style: "thin", color: "#000000" },
    };
    sheet.getRange(rangeAddress).format.rowHeight = 28;
  }

  const headers = [
    "代码",
    "指数名称",
    "收盘价",
    "日涨跌幅",
    "近一周",
    "近一月",
    `${payload.indexPerformanceYear}年初至今涨跌幅`,
  ];
  sheet.getRange("A2:G2").values = [headers];
  sheet.getRange("I2:O2").values = [headers];
  for (const rangeAddress of ["A2:G2", "I2:O2"]) {
    sheet.getRange(rangeAddress).format = {
      fill: "#203864",
      font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 10 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: true,
      borders: { preset: "all", style: "thin", color: "#000000" },
    };
    sheet.getRange(rangeAddress).format.rowHeight = 34;
  }

  sheet.getRange("A3:G11").values = mainRows.map((row) => [
    row.code, row.name, row.close, null, null, null, null,
  ]);
  sheet.getRange("I3:O11").values = styleRows.map((row) => [
    row.code, row.name, row.close, null, null, null, null,
  ]);
  sheet.getRange("P3:S11").values = mainRows.map((row) => [
    row.dailyBaseClose, row.weekBaseClose, row.monthBaseClose, row.yearBaseClose,
  ]);
  sheet.getRange("T3:W11").values = styleRows.map((row) => [
    row.dailyBaseClose, row.weekBaseClose, row.monthBaseClose, row.yearBaseClose,
  ]);

  for (let row = 3; row <= 11; row += 1) {
    sheet.getRange(`D${row}:G${row}`).formulas = [[
      `=IFERROR((C${row}/P${row}-1)*100,"")`,
      `=IFERROR((C${row}/Q${row}-1)*100,"")`,
      `=IFERROR((C${row}/R${row}-1)*100,"")`,
      `=IFERROR((C${row}/S${row}-1)*100,"")`,
    ]];
    sheet.getRange(`L${row}:O${row}`).formulas = [[
      `=IFERROR((K${row}/T${row}-1)*100,"")`,
      `=IFERROR((K${row}/U${row}-1)*100,"")`,
      `=IFERROR((K${row}/V${row}-1)*100,"")`,
      `=IFERROR((K${row}/W${row}-1)*100,"")`,
    ]];
  }

  for (const [rangeAddress, fill] of [["A3:G11", "#FCE4D6"], ["I3:O11", "#FFFFFF"]]) {
    sheet.getRange(rangeAddress).format = {
      fill,
      font: { name: BODY_FONT, color: "#000000", size: 10 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      borders: { preset: "all", style: "thin", color: "#000000" },
    };
    sheet.getRange(rangeAddress).format.rowHeight = 23;
  }
  sheet.getRange("C3:C11").format.numberFormat = "#,##0.00";
  sheet.getRange("D3:G11").format.numberFormat = "0.00";
  sheet.getRange("K3:K11").format.numberFormat = "#,##0.00";
  sheet.getRange("L3:O11").format.numberFormat = "0.00";
  for (const rangeAddress of ["D3:G11", "L3:O11"]) {
    sheet.getRange(rangeAddress).conditionalFormats.add("cellIs", {
      operator: "lessThan",
      formula: 0,
      format: { font: { color: NEGATIVE_RED } },
    });
  }

  sheet.getRange("P1:S1").merge();
  sheet.getRange("T1:W1").merge();
  sheet.getRange("P1").values = [["主要指数计算基准"]];
  sheet.getRange("T1").values = [["风格指数计算基准"]];
  sheet.getRange("P2:S2").values = [[
    `日基准 ${mainRows[0].dailyBaseDate}`,
    `周基准 ${mainRows[0].weekBaseDate}`,
    `月基准 ${mainRows[0].monthBaseDate}`,
    `年基准 ${mainRows[0].yearBaseDate}`,
  ]];
  sheet.getRange("T2:W2").values = [[
    `日基准 ${styleRows[0].dailyBaseDate}`,
    `周基准 ${styleRows[0].weekBaseDate}`,
    `月基准 ${styleRows[0].monthBaseDate}`,
    `年基准 ${styleRows[0].yearBaseDate}`,
  ]];
  for (const rangeAddress of ["P1:S1", "T1:W1"]) {
    sheet.getRange(rangeAddress).format = {
      fill: LIGHT_BLUE,
      font: { name: BODY_FONT, bold: true, color: "#404040", size: 10 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
    };
  }
  sheet.getRange("P2:W2").format = {
    fill: "#F2F2F2",
    font: { name: BODY_FONT, bold: true, color: "#404040", size: 9 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: GRID },
  };
  sheet.getRange("P3:W11").format = {
    font: { name: BODY_FONT, color: "#008000", size: 9 },
    horizontalAlignment: "right",
    verticalAlignment: "center",
    numberFormat: "#,##0.0000",
    borders: { preset: "all", style: "thin", color: GRID },
  };

  sheet.getRange("A12:O12").merge();
  sheet.getRange("A12").values = [[
    `数据来源：${payload.indexPerformanceSource.parquet}｜数据日期：${payload.runDate}｜口径：日/周/月分别相对前1/6/23个交易日收盘价，全年相对上一年最后交易日收盘价；收益率单元格由右侧基准收盘价公式计算。`,
  ]];
  sheet.getRange("A12:O12").format = {
    fill: LIGHT_BLUE,
    font: { name: BODY_FONT, color: "#404040", size: 9 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    wrapText: true,
  };
  sheet.getRange("A12:O12").format.rowHeight = 34;

  for (const column of ["A", "I"]) sheet.getRange(`${column}1:${column}12`).format.columnWidth = 16;
  for (const column of ["B", "J"]) sheet.getRange(`${column}1:${column}12`).format.columnWidth = 22;
  for (const column of ["C", "K"]) sheet.getRange(`${column}1:${column}12`).format.columnWidth = 13;
  for (const column of ["D", "E", "F", "L", "M", "N"]) sheet.getRange(`${column}1:${column}12`).format.columnWidth = 12;
  for (const column of ["G", "O"]) sheet.getRange(`${column}1:${column}12`).format.columnWidth = 17;
  sheet.getRange("H1:H12").format.columnWidth = 2;
  sheet.getRange("P1:W11").format.columnWidth = 18;
}


addMarketSheet();
addIndexSheet();
addReturnDistributionSheet();
addIndexPerformanceSheet();

const summary = await workbook.inspect({
  kind: "sheet,table",
  maxChars: 2500,
  tableMaxRows: 5,
  tableMaxCols: 5,
});
console.log(summary.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const previewDir = previewDirArg || path.dirname(payloadPath);
await fs.mkdir(previewDir, { recursive: true });
for (const sheetName of ["两融余额", "指数成交额", "涨跌分布", "指数表现"]) {
  const preview = await workbook.render({
    sheetName,
    range:
      sheetName === "两融余额"
        ? "A1:B18"
        : sheetName === "指数成交额"
          ? "A1:E18"
          : sheetName === "涨跌分布"
            ? "A1:L20"
            : "A1:O12",
    scale: 1.5,
    format: "png",
  });
  await fs.writeFile(
    path.join(previewDir, `${sheetName}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);'''

RED = "#E6121B"
BLUE = "#0262BA"
GRID_MAJOR = "#D9D9D9"
GRID_MINOR = "#ECECEC"
CM_PER_INCH = 2.54
CHART_WIDTH_CM = 8.81
CHART_HEIGHT_CM = 5.47
CHART_DPI = 254
CHART_PIXEL_WIDTH = 881
CHART_PIXEL_HEIGHT = 547
TABLE_PIXEL_HEIGHT = 480
CHART_FIGSIZE = (
    (CHART_PIXEL_WIDTH + 0.01) / CHART_DPI,
    (CHART_PIXEL_HEIGHT + 0.01) / CHART_DPI,
)
DOUBLE_CHART_PIXEL_WIDTH = CHART_PIXEL_WIDTH * 2
DOUBLE_CHART_FIGSIZE = (
    (DOUBLE_CHART_PIXEL_WIDTH + 0.01) / CHART_DPI,
    (CHART_PIXEL_HEIGHT + 0.01) / CHART_DPI,
)
TABLE_FIGSIZE = (
    (DOUBLE_CHART_PIXEL_WIDTH + 0.01) / CHART_DPI,
    (TABLE_PIXEL_HEIGHT + 0.01) / CHART_DPI,
)
SECTION_BAR_HEIGHT = 36
TITLE_FONT_SIZE = 8
AXIS_FONT_SIZE = 6
TICK_FONT_SIZE = 6
LEGEND_FONT_SIZE = 6
NOTE_FONT_SIZE = 6

DISTRIBUTION_LABELS = (
    ["<-5%"]
    + [f"{lower}%~{lower + 1}%" for lower in range(-5, 5)]
    + [">5%"]
)

MAIN_INDEX_SPECS = (
    ("000832.CSI", "转债指数", "中证转债"),
    ("889033.WI", "转债等权", "可转债等权"),
    ("889035.WI", "正股等权指数", "可转债正股等权"),
    ("884257.WI", "转债预案", "可转债预案"),
    ("000001.SH", "上证综指", "上证综指"),
    ("399001.SZ", "深证成指", "深证成指"),
    ("399006.SZ", "创业板指", "创业板指"),
    ("000016.SH", "上证50", "上证50"),
    ("000852.SH", "中证1000", "中证1000"),
)
STYLE_INDEX_SPECS = (
    ("801811.SI", "大盘指数", "大盘指数(申万)"),
    ("801812.SI", "中盘指数", "中盘指数(申万)"),
    ("801813.SI", "小盘指数", "小盘指数(申万)"),
    ("399372.SZ", "大盘成长", "大盘成长"),
    ("399373.SZ", "大盘价值", "大盘价值"),
    ("399374.SZ", "中盘成长", "中盘成长"),
    ("399375.SZ", "中盘价值", "中盘价值"),
    ("399376.SZ", "小盘成长", "小盘成长"),
    ("399377.SZ", "小盘价值", "小盘价值"),
)


def same_day_last_year(value: date) -> date:
    """返回上一年同月同日；2 月 29 日回退至 2 月 28 日。"""
    try:
        return value.replace(year=value.year - 1)
    except ValueError:
        return value.replace(year=value.year - 1, day=28)


IFIND_CREDENTIAL_FILE = WORKSPACE / "private/ifind账号.txt"


def load_ifind_credentials() -> tuple[str, str]:
    """从项目目录的 ifind账号.txt 读取统一登录账号。"""
    if not IFIND_CREDENTIAL_FILE.is_file():
        raise FileNotFoundError(f"未找到iFinD账号文件：{IFIND_CREDENTIAL_FILE}")
    config = ConfigParser(interpolation=None)
    config.read(IFIND_CREDENTIAL_FILE, encoding="utf-8")
    username = config.get("ifind", "username", fallback="").strip()
    password = config.get("ifind", "password", fallback="").strip()
    if not username or not password:
        raise RuntimeError("ifind账号.txt中的[ifind] username或password为空")
    return username, password


THS_LOGIN_OK_CODES = (0, -201)


def ths_login_errmsg(code: int) -> str:
    """获取 iFinD 登录/API 状态码说明。"""
    try:
        info = THS_GetErrorInfo(code)
        if isinstance(info, dict):
            return str(info.get("errmsg", info))
        return str(info)
    except Exception:
        return f"未知错误（状态码 {code}）"


def is_ths_login_ok(code: int) -> bool:
    """0、-201 视为登录可用（-201 通常为已登录）。"""
    return code in THS_LOGIN_OK_CODES


def print_ifind_usage() -> None:
    """显示iFinD各数据项的已用额度比例。"""
    try:
        result = THS_DataStatistics()
        tables = result.get("tables", {}) if isinstance(result, dict) else {}
        if not tables:
            detail = result.get("errmsg", "未返回额度数据") if isinstance(result, dict) else str(result)
            print(f"[警告] iFinD使用额度查询失败：{detail}")
            return
        print("iFinD使用额度：")
        for key, value in tables.items():
            ratio = value.get("ratio", "N/A") if isinstance(value, dict) else value
            print(f"{key} 已用：{ratio}")
    except Exception as exc:
        print(f"[警告] iFinD使用额度查询失败：{exc}")


def ths_login(
    ths_id: Optional[str] = None, ths_password: Optional[str] = None
) -> int:
    """登录 iFinD 并返回状态码。"""
    if not ths_id or not ths_password:
        file_id, file_password = load_ifind_credentials()
        ths_id = ths_id or file_id
        ths_password = ths_password or file_password
    code = THS_iFinDLogin(ths_id, ths_password)
    print(f"登录状态码: {code}")
    if not is_ths_login_ok(code):
        print(f"登录失败: {ths_login_errmsg(code)}")
    elif code == -201:
        print("登录成功！")
    else:
        print(ths_login_errmsg(code))
    if is_ths_login_ok(code):
        print_ifind_usage()
    return code


def extract_data(result, api_name: str) -> pd.DataFrame:
    errorcode = getattr(result, "errorcode", None)
    if errorcode not in (None, 0):
        errmsg = getattr(result, "errmsg", "")
        raise RuntimeError(f"{api_name} 调用失败（{errorcode}）：{errmsg}")
    data = getattr(result, "data", None)
    if data is None or data.empty:
        raise RuntimeError(f"{api_name} 未返回数据")
    return data.copy()


def fetch_market_statistics(start: date, end: date) -> pd.DataFrame:
    result = THS_DR(
        "p03438",
        (
            f"sdate={start:%Y%m%d};edate={end:%Y%m%d};"
            "sclx=沪深两市;pl=日"
        ),
        "p03438_f001:Y,p03438_f002:Y",
        "format:dataframe",
    )
    data = extract_data(result, "THS_DR/p03438")
    required = {"p03438_f001", "p03438_f002"}
    if not required.issubset(data.columns):
        raise RuntimeError(f"p03438 返回字段异常：{data.columns.tolist()}")
    output = data.rename(
        columns={"p03438_f001": "交易日期", "p03438_f002": "沪深两市融资融券余额_亿元"}
    )
    output["交易日期"] = pd.to_datetime(output["交易日期"], errors="coerce")
    output["沪深两市融资融券余额_亿元"] = pd.to_numeric(
        output["沪深两市融资融券余额_亿元"], errors="coerce"
    )
    return output.dropna().sort_values("交易日期").drop_duplicates("交易日期")


def fetch_index_turnover(start: date, end: date) -> pd.DataFrame:
    result = THS_DS(
        "000001.SH,399001.SZ,000832.CSI",
        "ths_trans_amt_index",
        "",
        "mode:thscode,block:history",
        f"{start:%Y-%m-%d}",
        f"{end:%Y-%m-%d}",
    )
    data = extract_data(result, "THS_DS/ths_trans_amt_index")
    required = {"time", "000001.SH", "399001.SZ", "000832.CSI"}
    if not required.issubset(data.columns):
        raise RuntimeError(f"指数成交额返回字段异常：{data.columns.tolist()}")
    output = data.rename(
        columns={
            "time": "交易日期",
            "000001.SH": "上证指数成交额_亿元",
            "399001.SZ": "深证成指成交额_亿元",
            "000832.CSI": "中证转债指数成交额_亿元",
        }
    )
    output["交易日期"] = pd.to_datetime(output["交易日期"], errors="coerce")
    for column in (
        "上证指数成交额_亿元",
        "深证成指成交额_亿元",
        "中证转债指数成交额_亿元",
    ):
        output[column] = pd.to_numeric(output[column], errors="coerce") / 100_000_000
    output["沪深成交额合计_亿元"] = (
        output["上证指数成交额_亿元"] + output["深证成指成交额_亿元"]
    )
    return output.dropna().sort_values("交易日期").drop_duplicates("交易日期")


def _return_bucket(value: float) -> str:
    if value < -5:
        return "<-5%"
    if value > 5:
        return ">5%"
    if math.isclose(value, 5.0, abs_tol=1e-12):
        return "4%~5%"
    lower = max(-5, min(4, math.floor(value)))
    return f"{lower}%~{lower + 1}%"


def _month_parquet_path(value: date) -> Path:
    return CB_PARQUET_ROOT / f"{value:%Y}" / f"{value:%Y%m}.parquet"


def latest_cb_trade_date() -> date:
    """返回月度转债 Parquet 中最新的实际交易日。"""
    parquet_files = sorted(
        (
            path
            for path in CB_PARQUET_ROOT.glob("[0-9][0-9][0-9][0-9]/*.parquet")
            if path.stem.isdigit() and len(path.stem) == 6
        ),
        key=lambda path: path.stem,
        reverse=True,
    )
    if not parquet_files:
        raise FileNotFoundError(f"未找到月度转债 Parquet：{CB_PARQUET_ROOT}")

    for parquet_path in parquet_files:
        frame = pd.read_parquet(parquet_path, columns=["交易日期", "交易状态"])
        trading_dates = pd.to_datetime(
            frame.loc[
                frame["交易状态"].astype(str).str.strip().eq("交易"), "交易日期"
            ],
            errors="coerce",
        ).dropna()
        if not trading_dates.empty:
            return trading_dates.max().date()
    raise RuntimeError("月度转债 Parquet 中未找到交易状态为“交易”的日期")


def fetch_index_performance(run_date: date) -> tuple[pd.DataFrame, dict[str, object]]:
    """从指数 Parquet 计算主要指数与风格指数的区间表现。"""
    if not INDEX_PARQUET.is_file():
        raise FileNotFoundError(f"未找到指数 Parquet：{INDEX_PARQUET}")

    data = pd.read_parquet(INDEX_PARQUET)
    required = {"指数名称", "交易日期", "指数值"}
    if not required.issubset(data.columns):
        raise RuntimeError(
            f"指数 Parquet 字段异常，缺少：{sorted(required - set(data.columns))}"
        )
    data = data.copy()
    data["交易日期"] = pd.to_datetime(data["交易日期"], errors="coerce")
    data["指数值"] = pd.to_numeric(data["指数值"], errors="coerce")
    data = data.dropna(subset=["指数名称", "交易日期", "指数值"])
    data = data.loc[data["交易日期"].le(pd.Timestamp(run_date))]

    rows: list[dict[str, object]] = []
    all_specs = (("主要指数", MAIN_INDEX_SPECS), ("风格指数", STYLE_INDEX_SPECS))
    for group_name, specs in all_specs:
        for code, parquet_name, display_name in specs:
            series = (
                data.loc[data["指数名称"].astype(str).eq(parquet_name),
                         ["交易日期", "指数值"]]
                .sort_values("交易日期")
                .drop_duplicates("交易日期", keep="last")
                .set_index("交易日期")["指数值"]
                .dropna()
            )
            if series.empty or series.index[-1].date() != run_date:
                latest = series.index[-1].date() if not series.empty else None
                raise RuntimeError(
                    f"指数“{parquet_name}”未更新至 {run_date:%Y-%m-%d}，"
                    f"当前最新日期：{latest}"
                )
            if len(series) < 24:
                raise RuntimeError(f"指数“{parquet_name}”历史记录不足24个交易日")

            year_history = series.loc[series.index < pd.Timestamp(run_date.year, 1, 1)]
            if year_history.empty:
                raise RuntimeError(f"指数“{parquet_name}”缺少上年末基准收盘价")

            current_date = series.index[-1]
            current_close = float(series.iloc[-1])
            daily_date, daily_base = series.index[-2], float(series.iloc[-2])
            week_date, week_base = series.index[-7], float(series.iloc[-7])
            month_date, month_base = series.index[-24], float(series.iloc[-24])
            year_date, year_base = year_history.index[-1], float(year_history.iloc[-1])
            rows.append(
                {
                    "组别": group_name,
                    "代码": code,
                    "指数名称": display_name,
                    "Parquet指数名称": parquet_name,
                    "数据日期": current_date,
                    "收盘价": current_close,
                    "日基准日期": daily_date,
                    "日基准收盘价": daily_base,
                    "日涨跌幅": (current_close / daily_base - 1) * 100,
                    "周基准日期": week_date,
                    "周基准收盘价": week_base,
                    "近一周涨跌幅": (current_close / week_base - 1) * 100,
                    "月基准日期": month_date,
                    "月基准收盘价": month_base,
                    "近一月涨跌幅": (current_close / month_base - 1) * 100,
                    "年基准日期": year_date,
                    "年基准收盘价": year_base,
                    "年初至今涨跌幅": (current_close / year_base - 1) * 100,
                }
            )

    result = pd.DataFrame(rows)
    source = {
        "parquet": str(INDEX_PARQUET.relative_to(WORKSPACE)),
        "sampleDate": f"{run_date:%Y-%m-%d}",
        "dailyRule": "相对前1个交易日收盘价",
        "weeklyRule": "相对前6个交易日收盘价",
        "monthlyRule": "相对前23个交易日收盘价",
        "yearRule": "相对上一年最后交易日收盘价",
    }
    return result, source


def _previous_month(value: date) -> date:
    return (pd.Timestamp(value.replace(day=1)) - pd.Timedelta(days=1)).date()


def fetch_cb_daily_returns(
    run_date: date,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, int], dict[str, object]]:
    """从月度 Parquet 读取交易中转债，并按收盘价计算当日涨跌幅。"""
    current_path = _month_parquet_path(run_date)
    if not current_path.is_file():
        raise FileNotFoundError(f"未找到运行日转债 Parquet：{current_path}")
    if not CB_MASTER_PARQUET.is_file():
        raise FileNotFoundError(f"未找到转债总表 Parquet：{CB_MASTER_PARQUET}")

    required = {"转债代码", "交易日期", "收盘价", "交易状态"}
    current_month = pd.read_parquet(current_path)
    if not required.issubset(current_month.columns):
        raise RuntimeError(
            f"转债 Parquet 字段异常，缺少：{sorted(required - set(current_month.columns))}"
        )
    current_month = current_month.copy()
    current_month["交易日期"] = pd.to_datetime(
        current_month["交易日期"], errors="coerce"
    )
    current_month["收盘价"] = pd.to_numeric(current_month["收盘价"], errors="coerce")
    current = current_month.loc[
        current_month["交易日期"].eq(pd.Timestamp(run_date))
        & current_month["交易状态"].astype(str).str.strip().eq("交易")
        & current_month["收盘价"].notna(),
        ["转债代码", "交易日期", "收盘价", "交易状态"],
    ].drop_duplicates("转债代码", keep="last")
    if current.empty:
        raise RuntimeError(f"{run_date:%Y-%m-%d} 的 Parquet 中没有交易中转债")

    codes = set(current["转债代码"].astype(str))
    history_frames = [current_month]
    source_paths = [current_path]
    cursor = _previous_month(run_date)
    matched_codes: set[str] = set()
    for _ in range(24):
        history = pd.concat(history_frames, ignore_index=True)
        history_dates = pd.to_datetime(history["交易日期"], errors="coerce")
        history_close = pd.to_numeric(history["收盘价"], errors="coerce")
        matched_codes = set(
            history.loc[
                history["转债代码"].astype(str).isin(codes)
                & history_dates.lt(pd.Timestamp(run_date))
                & history_close.notna(),
                "转债代码",
            ].astype(str)
        )
        if matched_codes == codes:
            break
        prior_path = _month_parquet_path(cursor)
        if prior_path.is_file():
            prior = pd.read_parquet(
                prior_path, columns=["转债代码", "交易日期", "收盘价", "交易状态"]
            )
            history_frames.append(prior)
            source_paths.append(prior_path)
        cursor = _previous_month(cursor)

    history = pd.concat(history_frames, ignore_index=True)
    history["交易日期"] = pd.to_datetime(history["交易日期"], errors="coerce")
    history["收盘价"] = pd.to_numeric(history["收盘价"], errors="coerce")
    previous = (
        history.loc[
            history["转债代码"].astype(str).isin(codes)
            & history["交易日期"].lt(pd.Timestamp(run_date))
            & history["收盘价"].notna(),
            ["转债代码", "交易日期", "收盘价"],
        ]
        .sort_values(["转债代码", "交易日期"])
        .drop_duplicates("转债代码", keep="last")
        .rename(columns={"交易日期": "前收盘日期", "收盘价": "前收盘价"})
    )

    master = pd.read_parquet(CB_MASTER_PARQUET, columns=["转债代码", "转债名称"])
    details = (
        current.rename(columns={"交易日期": "当日日期", "收盘价": "当日收盘价"})
        .merge(previous, on="转债代码", how="inner", validate="one_to_one")
        .merge(master.drop_duplicates("转债代码"), on="转债代码", how="left")
        .rename(columns={"转债名称": "转债简称"})
    )
    details["当日涨跌幅_百分比"] = (
        details["当日收盘价"].div(details["前收盘价"]).sub(1).mul(100)
    )
    details["涨跌幅区间"] = details["当日涨跌幅_百分比"].map(_return_bucket)
    details = details[
        [
            "转债代码",
            "转债简称",
            "前收盘日期",
            "前收盘价",
            "当日日期",
            "当日收盘价",
            "当日涨跌幅_百分比",
            "涨跌幅区间",
            "交易状态",
        ]
    ]
    if details.empty:
        raise RuntimeError(f"{run_date:%Y-%m-%d} 未取得可计算涨跌幅的交易中转债")
    details = details.sort_values(
        ["当日涨跌幅_百分比", "转债代码"], ascending=[False, True]
    ).reset_index(drop=True)

    returns = details["当日涨跌幅_百分比"]
    flat_mask = np.isclose(returns, 0.0, atol=1e-12)
    summary = {
        "上涨": int((returns > 0).sum()),
        "下跌": int((returns < 0).sum()),
        "平盘": int(flat_mask.sum()),
        "有效样本": int(len(details)),
    }
    counts = (
        details["涨跌幅区间"]
        .value_counts()
        .reindex(DISTRIBUTION_LABELS, fill_value=0)
        .astype(int)
    )
    distribution = pd.DataFrame(
        {"涨跌幅区间": list(DISTRIBUTION_LABELS), "转债数量": counts.to_list()}
    )
    source_info: dict[str, object] = {
        "currentParquet": str(current_path.relative_to(WORKSPACE)),
        "historyParquets": [str(path.relative_to(WORKSPACE)) for path in source_paths],
        "previousLatestDate": f"{details['前收盘日期'].max():%Y-%m-%d}",
        "sampleRule": "运行日交易状态=交易，且当日及此前最近收盘价均有效",
        "returnFormula": "(当日收盘价/此前最近收盘价-1)×100%",
    }
    return details, distribution, summary, source_info


def setup_font() -> fm.FontProperties:
    if not FONT_PATH.exists():
        raise FileNotFoundError(f"未找到报告字体：{FONT_PATH}")
    fm.fontManager.addfont(str(FONT_PATH))
    # 图表基准字号统一为 7；表格绘制时单独覆盖为 10。
    font = fm.FontProperties(fname=str(FONT_PATH), size=7)
    plt.rcParams["font.family"] = font.get_name()
    plt.rcParams["axes.unicode_minus"] = False
    return font


def get_title_font(size: float = TITLE_FONT_SIZE) -> fm.FontProperties:
    if not TITLE_FONT_PATH.exists():
        raise FileNotFoundError(f"未找到标题黑体字体：{TITLE_FONT_PATH}")
    fm.fontManager.addfont(str(TITLE_FONT_PATH))
    return fm.FontProperties(fname=str(TITLE_FONT_PATH), size=size)


def date_axis_format_for_span(date_values: pd.Series) -> str:
    """跨度严格超过一个公历年时省略日期中的日。"""
    minimum_date = pd.Timestamp(date_values.min())
    maximum_date = pd.Timestamp(date_values.max())
    if maximum_date > minimum_date + pd.DateOffset(years=1):
        return "%Y%m"
    return "%Y-%m-%d"


def style_axis(ax, font: fm.FontProperties) -> None:
    ax.tick_params(axis="both", colors="black", labelsize=TICK_FONT_SIZE, width=0.5)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(font)
        label.set_fontsize(TICK_FONT_SIZE)
    for spine in ax.spines.values():
        spine.set_color("black")
        spine.set_linewidth(0.5)
    ax.grid(True, which="major", axis="both", color=GRID_MAJOR, linewidth=0.4, alpha=0.75)
    ax.grid(True, which="minor", axis="x", color=GRID_MINOR, linewidth=0.3, alpha=0.6)
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=4))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    ax.xaxis.set_minor_locator(mdates.MonthLocator())


def add_chart_panel_title(fig: plt.Figure, title: str) -> None:
    """按日报版式增加浅灰标题栏和完整图表外框。"""
    fig.add_artist(
        plt.Rectangle(
            (0, 0),
            1,
            1,
            transform=fig.transFigure,
            facecolor="none",
            edgecolor="black",
            linewidth=0.65,
            zorder=20,
            clip_on=False,
        )
    )
    fig.add_artist(
        plt.Rectangle(
            (0, 0.93),
            1,
            0.07,
            transform=fig.transFigure,
            facecolor="#D9E2F3",
            edgecolor="black",
            linewidth=0.65,
            zorder=19,
            clip_on=False,
        )
    )
    fig.text(
        0.5,
        0.965,
        title,
        ha="center",
        va="center",
        fontproperties=get_title_font(TITLE_FONT_SIZE),
        fontsize=TITLE_FONT_SIZE,
        fontweight="bold",
        color=RED,
        zorder=21,
    )


def plot_market_statistics(
    data: pd.DataFrame, output_path: Path, font: fm.FontProperties
) -> None:
    """按已确认的华创“单轴”模板固化两融余额折线图。"""
    if not CHART_FONT_PATH.exists():
        raise FileNotFoundError(f"未找到华文楷体字体：{CHART_FONT_PATH}")
    fm.fontManager.addfont(str(CHART_FONT_PATH))
    chart_font = fm.FontProperties(fname=str(CHART_FONT_PATH), size=7)

    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    ax.plot(
        data["交易日期"],
        data["沪深两市融资融券余额_亿元"],
        color=RED,
        linewidth=1.0,
        marker=None,
        label="沪深两市融资融券余额（亿元）",
    )

    ax.set_xlim(data["交易日期"].min(), data["交易日期"].max())
    value_max = float(data["沪深两市融资融券余额_亿元"].max())
    y_step = 5000.0
    ax.set_ylim(0, max(y_step, math.ceil(value_max / y_step) * y_step))
    ax.yaxis.set_major_locator(mticker.MultipleLocator(y_step))
    ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.2f}"))
    ax.xaxis.set_major_locator(
        mdates.MonthLocator(bymonth=(1, 3, 5, 7, 9, 11), bymonthday=2)
    )
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(data["交易日期"]))
    )
    ax.tick_params(
        axis="both",
        which="major",
        colors="black",
        labelsize=7,
        width=0.6,
        length=3,
        top=False,
        right=False,
    )
    for label in ax.get_xticklabels():
        label.set_fontproperties(chart_font)
        label.set_fontsize(7)
        label.set_rotation(90)
        label.set_horizontalalignment("center")
    for label in ax.get_yticklabels():
        label.set_fontproperties(chart_font)
        label.set_fontsize(7)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color("black")
        ax.spines[side].set_linewidth(0.7)
    ax.grid(False)
    ax.set_xlabel("")
    ax.set_ylabel("")

    legend = ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.50),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        handlelength=3.2,
        handletextpad=0.8,
        borderaxespad=0,
        ncol=1,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.15, right=0.955, top=0.93, bottom=0.39)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_index_turnover(
    data: pd.DataFrame, output_path: Path, font: fm.FontProperties
) -> None:
    """按已确认的华创“双轴”模板固化成交额折线图。"""
    if not CHART_FONT_PATH.exists():
        raise FileNotFoundError(f"未找到华文楷体字体：{CHART_FONT_PATH}")
    fm.fontManager.addfont(str(CHART_FONT_PATH))
    chart_font = fm.FontProperties(fname=str(CHART_FONT_PATH), size=7)

    fig, ax_left = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    ax_right = ax_left.twinx()
    line_cb, = ax_left.plot(
        data["交易日期"],
        data["中证转债指数成交额_亿元"],
        color=RED,
        linewidth=1.0,
        marker=None,
        label="中证转债成交额（亿元）",
    )
    line_total, = ax_right.plot(
        data["交易日期"],
        data["沪深成交额合计_亿元"],
        color=BLUE,
        linewidth=1.0,
        marker=None,
        label="沪深两市成交额合计（亿元）",
    )
    latest = (
        data.dropna(subset=["中证转债指数成交额_亿元", "沪深成交额合计_亿元"])
        .sort_values("交易日期")
        .iloc[-1]
    )
    panel_title = (
        f"成交额:转债{float(latest['中证转债指数成交额_亿元']):.2f}亿，"
        f"A股{float(latest['沪深成交额合计_亿元']):.2f}亿"
    )
    add_chart_panel_title(fig, panel_title)

    left_step = 200.0
    right_step = 5000.0
    left_max = float(data["中证转债指数成交额_亿元"].max())
    right_max = float(data["沪深成交额合计_亿元"].max())
    ax_left.set_ylim(
        0, max(left_step, math.ceil(left_max * 1.05 / left_step) * left_step)
    )
    ax_right.set_ylim(
        0, max(right_step, math.ceil(right_max * 1.05 / right_step) * right_step)
    )
    ax_left.yaxis.set_major_locator(mticker.MultipleLocator(left_step))
    ax_right.yaxis.set_major_locator(mticker.MultipleLocator(right_step))
    ax_left.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.0f}"))
    ax_right.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.0f}"))

    ax_left.set_xlim(data["交易日期"].min(), data["交易日期"].max())
    tick_start = pd.Timestamp(data["交易日期"].min())
    tick_end = pd.Timestamp(data["交易日期"].max())
    tick_dates = []
    tick_date = tick_start
    while tick_date <= tick_end:
        tick_dates.append(tick_date)
        tick_date = tick_date + pd.DateOffset(months=2)
    if tick_dates[-1] != tick_end and (tick_end - tick_dates[-1]).days >= 20:
        tick_dates.append(tick_end)
    ax_left.set_xticks(tick_dates)
    ax_left.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(data["交易日期"]))
    )

    for axis in (ax_left, ax_right):
        axis.tick_params(
            axis="both",
            which="major",
            colors="black",
            labelsize=7,
            width=0.6,
            length=3,
            top=False,
        )
        for label in axis.get_yticklabels():
            label.set_fontproperties(chart_font)
            label.set_fontsize(7)
        axis.grid(False)
        axis.set_xlabel("")
        axis.set_ylabel("")
    for label in ax_left.get_xticklabels():
        label.set_fontproperties(chart_font)
        label.set_fontsize(7)
        label.set_rotation(90)
        label.set_horizontalalignment("center")

    ax_left.spines["top"].set_visible(False)
    ax_right.spines["top"].set_visible(False)
    ax_left.spines["right"].set_visible(False)
    ax_right.spines["left"].set_visible(False)
    for spine in (ax_left.spines["left"], ax_left.spines["bottom"], ax_right.spines["right"]):
        spine.set_color("black")
        spine.set_linewidth(0.7)

    legend = ax_left.legend(
        [line_cb, line_total],
        [line_cb.get_label(), line_total.get_label()],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.47),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=2,
        handlelength=2.7,
        handletextpad=0.6,
        columnspacing=1.0,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.115, right=0.885, top=0.88, bottom=0.36)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_cb_return_distribution(
    distribution: pd.DataFrame,
    summary: dict[str, int],
    run_date: date,
    output_path: Path,
    font: fm.FontProperties,
) -> None:
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    x = np.arange(len(distribution))
    colors = [BLUE] * 6 + [RED] * 6
    bars = ax.bar(
        x,
        distribution["转债数量"],
        width=0.82,
        color=colors,
        edgecolor="none",
        zorder=3,
    )
    valid_count = max(int(summary["有效样本"]), 1)
    up_pct = int(summary["上涨"]) / valid_count * 100
    down_pct = int(summary["下跌"]) / valid_count * 100
    add_chart_panel_title(
        fig,
        f"上涨转债占比{up_pct:.2f}%，下跌转债占比{down_pct:.2f}%",
    )
    ax.set_xticks(x)
    ax.set_xticklabels(
        distribution["涨跌幅区间"], rotation=45, ha="right", fontproperties=font
    )
    ax.tick_params(
        axis="both", colors="black", labelsize=TICK_FONT_SIZE, width=0.5
    )
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(font)
        label.set_fontsize(TICK_FONT_SIZE)
    for spine in ax.spines.values():
        spine.set_color("black")
        spine.set_linewidth(0.5)
    ax.grid(True, axis="y", color=GRID_MAJOR, linewidth=0.4, alpha=0.75, zorder=0)
    ax.set_axisbelow(True)
    ax.axvline(5.5, color="#A6A6A6", linewidth=0.5, linestyle="--", zorder=2)
    max_count = max(int(distribution["转债数量"].max()), 1)
    ax.set_ylim(0, max_count * 1.25)
    for bar, value in zip(bars, distribution["转债数量"]):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + max_count * 0.025,
            str(int(value)),
            ha="center",
            va="bottom",
            fontproperties=font,
            fontsize=6,
            color="black",
        )
    fig.tight_layout(rect=(0.015, 0.015, 0.985, 0.92), pad=0.5)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def _draw_index_performance_panel(
    ax,
    data: pd.DataFrame,
    year: int,
    font: fm.FontProperties,
    body_fill: str,
) -> None:
    ax.set_axis_off()
    headers = [
        "代码",
        "指数名称",
        "收盘价",
        "日涨跌幅",
        "近一周",
        "近一月",
        f"{year}年初至\n今涨跌幅",
    ]
    rows = [
        [
            str(row.代码),
            str(row.指数名称),
            f"{float(row.收盘价):.2f}",
            f"{float(row.日涨跌幅):.2f}",
            f"{float(row.近一周涨跌幅):.2f}",
            f"{float(row.近一月涨跌幅):.2f}",
            f"{float(row.年初至今涨跌幅):.2f}",
        ]
        for row in data.itertuples(index=False)
    ]
    table = ax.table(
        cellText=rows,
        colLabels=headers,
        cellLoc="center",
        colLoc="center",
        colWidths=[0.16, 0.22, 0.13, 0.11, 0.11, 0.11, 0.16],
        bbox=[0, 0, 1, 1],
    )
    table.auto_set_font_size(False)
    for (row_index, column_index), cell in table.get_celld().items():
        cell.set_edgecolor("black")
        cell.set_linewidth(0.4)
        cell.PAD = 0.015
        cell.get_text().set_fontproperties(font)
        if row_index == 0:
            cell.set_facecolor("#203864")
            cell.get_text().set_color("white")
            cell.get_text().set_fontsize(7.4)
        else:
            cell.set_facecolor(body_fill)
            cell.get_text().set_color("black")
            cell.get_text().set_fontsize(6.7)
            metric_column = {
                3: "日涨跌幅",
                4: "近一周涨跌幅",
                5: "近一月涨跌幅",
                6: "年初至今涨跌幅",
            }.get(column_index)
            if metric_column is not None:
                value = float(data.iloc[row_index - 1][metric_column])
                if value < 0:
                    cell.get_text().set_color(RED)


def plot_index_performance_table(
    data: pd.DataFrame,
    run_date: date,
    output_path: Path,
    font: fm.FontProperties,
) -> None:
    """生成主要指数、风格指数左右并列的报告表格图。"""
    main = data.loc[data["组别"].eq("主要指数")].reset_index(drop=True)
    style = data.loc[data["组别"].eq("风格指数")].reset_index(drop=True)
    if len(main) != 9 or len(style) != 9:
        raise RuntimeError(
            f"指数表现分组数量异常：主要指数 {len(main)} 个，风格指数 {len(style)} 个"
        )

    fig, axes = plt.subplots(
        1,
        2,
        figsize=TABLE_FIGSIZE,
        dpi=CHART_DPI,
        gridspec_kw={"wspace": 0},
    )
    _draw_index_performance_panel(
        axes[0], main, run_date.year, font, "#FCE4D6"
    )
    _draw_index_performance_panel(
        axes[1], style, run_date.year, font, "#FFFFFF"
    )
    fig.subplots_adjust(left=0, right=1, bottom=0, top=1, wspace=0)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def compose_index_market_overview(
    table_path: Path,
    turnover_path: Path,
    distribution_path: Path,
    output_path: Path,
    run_date: date,
) -> None:
    """将指数表、资金表现分隔栏及两张并列图合成为报告模块。"""
    if not REPORT_HEADER_PATH.is_file():
        raise FileNotFoundError(f"未找到日报表头图片：{REPORT_HEADER_PATH}")
    with Image.open(REPORT_HEADER_PATH) as source_header:
        source_header = source_header.convert("RGBA")
        header_draw = ImageDraw.Draw(source_header)
        header_text = (
            "【华创固收·周冠南团队】\n"
            f"可转债市场日度跟踪{run_date:%Y%m%d}"
        )
        header_font = ImageFont.truetype(str(FONT_PATH), 60)
        text_bbox = header_draw.textbbox((0, 0), header_text, font=header_font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        text_x = (source_header.width - text_width) // 2
        text_y = (source_header.height - text_height) // 2 - 40
        header_draw.text(
            (text_x, text_y),
            header_text,
            fill="white",
            font=header_font,
        )
        header_height = round(
            source_header.height * DOUBLE_CHART_PIXEL_WIDTH / source_header.width
        )
        resized_header = source_header.resize(
            (DOUBLE_CHART_PIXEL_WIDTH, header_height), Image.Resampling.LANCZOS
        )
        report_header = np.asarray(resized_header, dtype=np.float32) / 255.0

    table_image = plt.imread(table_path)
    turnover_image = plt.imread(turnover_path)
    distribution_image = plt.imread(distribution_path)
    expected_sizes = {
        "指数表": (TABLE_PIXEL_HEIGHT, DOUBLE_CHART_PIXEL_WIDTH),
        "成交额图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "涨跌幅直方图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
    }
    for label, image, expected in (
        ("指数表", table_image, expected_sizes["指数表"]),
        ("成交额图", turnover_image, expected_sizes["成交额图"]),
        ("涨跌幅直方图", distribution_image, expected_sizes["涨跌幅直方图"]),
    ):
        if image.shape[:2] != expected:
            raise RuntimeError(
                f"{label}尺寸异常：{image.shape[1]}×{image.shape[0]}，"
                f"预期{expected[1]}×{expected[0]}"
            )

    def to_rgba(image: np.ndarray) -> np.ndarray:
        if image.shape[2] == 4:
            return image
        alpha = np.ones((*image.shape[:2], 1), dtype=image.dtype)
        return np.concatenate([image, alpha], axis=2)

    section_figure = plt.figure(
        figsize=(
            (DOUBLE_CHART_PIXEL_WIDTH + 0.01) / CHART_DPI,
            (SECTION_BAR_HEIGHT + 0.01) / CHART_DPI,
        ),
        dpi=CHART_DPI,
        facecolor="#203864",
    )
    section_figure.text(
        0.5,
        0.5,
        "资金表现",
        ha="center",
        va="center",
        fontproperties=get_title_font(7),
        fontsize=7,
        fontweight="bold",
        color="white",
    )
    section_figure.canvas.draw()
    section_bar = (
        np.asarray(section_figure.canvas.buffer_rgba(), dtype=np.float32) / 255.0
    )
    plt.close(section_figure)
    if section_bar.shape[:2] != (SECTION_BAR_HEIGHT, DOUBLE_CHART_PIXEL_WIDTH):
        raise RuntimeError(
            f"资金表现分隔栏尺寸异常：{section_bar.shape[1]}×{section_bar.shape[0]}"
        )

    canvas = np.ones(
        (
            header_height + TABLE_PIXEL_HEIGHT + CHART_PIXEL_HEIGHT + SECTION_BAR_HEIGHT,
            DOUBLE_CHART_PIXEL_WIDTH,
            4,
        ),
        dtype=np.float32,
    )
    canvas[:header_height, :, :] = report_header
    table_start = header_height
    table_end = table_start + TABLE_PIXEL_HEIGHT
    canvas[table_start:table_end, :, :] = to_rgba(table_image)
    section_start = table_end
    chart_start = section_start + SECTION_BAR_HEIGHT
    canvas[section_start:chart_start, :, :] = section_bar
    canvas[chart_start:, :CHART_PIXEL_WIDTH, :] = to_rgba(turnover_image)
    canvas[chart_start:, CHART_PIXEL_WIDTH:, :] = to_rgba(
        distribution_image
    )
    plt.imsave(output_path, canvas, dpi=CHART_DPI)


def build_workbook(
    market: pd.DataFrame,
    index: pd.DataFrame,
    index_performance: pd.DataFrame,
    index_performance_source: dict[str, object],
    return_details: pd.DataFrame,
    return_distribution: pd.DataFrame,
    return_summary: dict[str, int],
    return_source: dict[str, object],
    run_date: date,
    index_start_date: date,
    market_start_date: date,
    output_path: Path,
) -> None:
    """通过 bundled artifact-tool 生成合并 Excel 底稿。"""
    for dependency in (BUNDLED_NODE, BUNDLED_NODE_MODULES):
        if not dependency.exists():
            raise FileNotFoundError(f"生成 Excel 所需依赖不存在：{dependency}")

    payload = {
        "runDate": f"{run_date:%Y-%m-%d}",
        "indexStartDate": f"{index_start_date:%Y-%m-%d}",
        "marketStartDate": f"{market_start_date:%Y-%m-%d}",
        "marketLatestDate": f"{market['交易日期'].max():%Y-%m-%d}",
        "indexLatestDate": f"{index['交易日期'].max():%Y-%m-%d}",
        "market": [
            {
                "date": f"{row.交易日期:%Y-%m-%d}",
                "balance": float(row.沪深两市融资融券余额_亿元),
            }
            for row in market.itertuples(index=False)
        ],
        "index": [
            {
                "date": f"{row.交易日期:%Y-%m-%d}",
                "shanghai": float(row.上证指数成交额_亿元),
                "shenzhen": float(row.深证成指成交额_亿元),
                "convertibleBond": float(row.中证转债指数成交额_亿元),
                "total": float(row.沪深成交额合计_亿元),
            }
            for row in index.itertuples(index=False)
        ],
        "indexPerformanceYear": run_date.year,
        "indexPerformanceSource": index_performance_source,
        "indexPerformance": [
            {
                "group": str(row.组别),
                "code": str(row.代码),
                "name": str(row.指数名称),
                "parquetName": str(row.Parquet指数名称),
                "date": f"{row.数据日期:%Y-%m-%d}",
                "close": float(row.收盘价),
                "dailyBaseDate": f"{row.日基准日期:%Y-%m-%d}",
                "dailyBaseClose": float(row.日基准收盘价),
                "weekBaseDate": f"{row.周基准日期:%Y-%m-%d}",
                "weekBaseClose": float(row.周基准收盘价),
                "monthBaseDate": f"{row.月基准日期:%Y-%m-%d}",
                "monthBaseClose": float(row.月基准收盘价),
                "yearBaseDate": f"{row.年基准日期:%Y-%m-%d}",
                "yearBaseClose": float(row.年基准收盘价),
            }
            for row in index_performance.itertuples(index=False)
        ],
        "returnSummary": return_summary,
        "returnSource": return_source,
        "returnDistribution": [
            {"bucket": str(row.涨跌幅区间), "count": int(row.转债数量)}
            for row in return_distribution.itertuples(index=False)
        ],
        "returnDetails": [
            {
                "code": str(row.转债代码),
                "name": str(row.转债简称),
                "previousDate": f"{row.前收盘日期:%Y-%m-%d}",
                "previousClose": float(row.前收盘价),
                "currentDate": f"{row.当日日期:%Y-%m-%d}",
                "currentClose": float(row.当日收盘价),
                "returnPct": float(row.当日涨跌幅_百分比),
                "bucket": str(row.涨跌幅区间),
                "tradingStatus": str(row.交易状态),
            }
            for row in return_details.itertuples(index=False)
        ],
    }

    with tempfile.TemporaryDirectory(prefix="daily_market_workbook_") as temp_dir_text:
        temp_dir = Path(temp_dir_text)
        payload_path = temp_dir / "payload.json"
        builder_path = temp_dir / "build_daily_market_workbook.mjs"
        node_modules_link = temp_dir / "node_modules"
        payload_path.write_text(
            json.dumps(payload, ensure_ascii=False), encoding="utf-8"
        )
        builder_path.write_text(WORKBOOK_BUILDER_SOURCE, encoding="utf-8")
        subprocess.run(
            [
                "cmd.exe",
                "/c",
                "mklink",
                "/J",
                str(node_modules_link),
                str(BUNDLED_NODE_MODULES),
            ],
            check=True,
            capture_output=True,
        )
        command = [
            str(BUNDLED_NODE),
            str(builder_path),
            str(payload_path),
            str(output_path),
        ]
        verification_dir = os.environ.get("DAILY_MARKET_VERIFY_DIR", "").strip()
        if verification_dir:
            command.append(str(Path(verification_dir).resolve()))
        result = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0:
            raise RuntimeError(
                "Excel 底稿生成失败：\n"
                + (result.stdout or "")
                + (result.stderr or "")
            )


def remove_obsolete_outputs(output_dir: Path, workbook_path: Path) -> None:
    """仅清理此前版本生成的已知 CSV/JSON 中间文件。"""
    for name in (
        "市场交易统计_p03438.csv",
        "指数成交金额_ths_trans_amt_index.csv",
        "运行信息.json",
        workbook_path.name + ".inspect.ndjson",
        "上证指数与深证成指成交额.png",
        "可转债上涨下跌占比.png",
    ):
        path = output_dir / name
        if path.is_file():
            path.unlink()


def run(run_date: date, output_dir: Path) -> dict[str, object]:
    start_date = same_day_last_year(run_date)
    market_start_date = MARGIN_BALANCE_START_DATE
    output_dir.mkdir(parents=True, exist_ok=True)

    login_code = ths_login()
    if not is_ths_login_ok(login_code):
        raise RuntimeError(
            f"iFinD 登录失败（{login_code}）：{ths_login_errmsg(login_code)}"
        )

    market = fetch_market_statistics(market_start_date, run_date)
    index = fetch_index_turnover(start_date, run_date)
    index_performance, index_performance_source = fetch_index_performance(run_date)
    return_details, return_distribution, return_summary, return_source = (
        fetch_cb_daily_returns(run_date)
    )

    market_png = output_dir / "沪深两市融资融券余额.png"
    index_png = output_dir / "中证转债与沪深两市成交额.png"
    return_png = output_dir / "可转债当日涨跌幅分布.png"
    index_performance_png = output_dir / "主要指数与风格指数表现.png"
    overview_png = output_dir / "主要指数与市场表现组合图.png"
    workbook_path = output_dir / f"转债日报市场数据底稿_{run_date:%Y%m%d}.xlsx"

    font = setup_font()
    plot_market_statistics(market, market_png, font)
    plot_index_turnover(index, index_png, font)
    plot_cb_return_distribution(
        return_distribution, return_summary, run_date, return_png, font
    )
    plot_index_performance_table(
        index_performance, run_date, index_performance_png, font
    )
    compose_index_market_overview(
        index_performance_png, index_png, return_png, overview_png, run_date
    )
    build_workbook(
        market,
        index,
        index_performance,
        index_performance_source,
        return_details,
        return_distribution,
        return_summary,
        return_source,
        run_date,
        start_date,
        market_start_date,
        workbook_path,
    )
    remove_obsolete_outputs(output_dir, workbook_path)

    metadata: dict[str, object] = {
        "运行日期": f"{run_date:%Y-%m-%d}",
        "指数成交额查询开始日期": f"{start_date:%Y-%m-%d}",
        "两融余额查询开始日期": f"{market_start_date:%Y-%m-%d}",
        "市场交易统计记录数": len(market),
        "市场交易统计最新日期": f"{market['交易日期'].max():%Y-%m-%d}",
        "指数成交金额记录数": len(index),
        "指数成交金额最新日期": f"{index['交易日期'].max():%Y-%m-%d}",
        "指数表现记录数": len(index_performance),
        "指数表现数据源": index_performance_source["parquet"],
        "指数表现最新日期": f"{index_performance['数据日期'].max():%Y-%m-%d}",
        "指数表现区间口径": "日/周/月分别相对前1/6/23个交易日，全年相对上一年最后交易日",
        "p03438_f002指标": "沪深两市融资融券余额（亿元）",
        "ths_trans_amt_index单位换算": "原始值（元）除以1亿，转为亿元",
        "成交额图左轴": "中证转债指数成交额（000832.CSI）",
        "成交额图右轴": "上证指数与深证成指成交额合计",
        "上涨转债数": return_summary["上涨"],
        "下跌转债数": return_summary["下跌"],
        "平盘转债数": return_summary["平盘"],
        "涨跌幅有效样本数": return_summary["有效样本"],
        "涨跌幅数据源": return_source["currentParquet"],
        "涨跌幅计算公式": return_source["returnFormula"],
        "报告字体": font.get_name(),
        "报告配色": {"红": RED, "蓝": BLUE},
        "Excel底稿": str(workbook_path),
    }
    return metadata


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--run-date",
        default=None,
        help="运行基准日，格式 YYYY-MM-DD；默认取转债 Parquet 最新交易日",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="输出目录；默认在 huachuang 目录下生成 YYYYMMDD_转债日报",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    run_date = (
        datetime.strptime(args.run_date, "%Y-%m-%d").date()
        if args.run_date
        else latest_cb_trade_date()
    )
    output_dir = args.output_dir or (
        WORKSPACE / "runs" / "daily" / f"{run_date:%Y%m%d}_转债日报"
    )
    metadata = run(run_date, output_dir.resolve())
    print(json.dumps(metadata, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
