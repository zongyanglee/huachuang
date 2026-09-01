# -*- coding: utf-8 -*-
"""华创固收转债日报：查询市场数据、读取转债 Parquet、绘图并生成 Excel 底稿。"""

from __future__ import annotations

import argparse
from copy import deepcopy
from configparser import ConfigParser
from contextlib import redirect_stdout
import hashlib
import io
import json
import math
import os
import posixpath
import re
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import Optional
from unicodedata import east_asian_width

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib import ticker as mticker
from matplotlib.patches import Patch
import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import pythoncom
import win32com.client

with redirect_stdout(io.StringIO()):
    from iFinDPy import THS_DR, THS_DS, THS_GetErrorInfo, THS_iFinDLogin


WORKSPACE = Path(__file__).resolve().parents[2]
DAILY_WORD_TEMPLATE_PATH = (
    WORKSPACE / "【华创固收】转债市场日度跟踪20260831.docx"
)
DAILY_WORD_TEMPLATE_SHA256 = (
    "922AA1FC6DA6C384E264A7077CBB7CE02FAE9C7501539D4E8B74CC76C4303FB3"
)
WORD_XML_NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
FONT_PATH = WORKSPACE / "assets/fonts/KaiTi_GB2312.ttf"
TITLE_FONT_PATH = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "simhei.ttf"
CHART_FONT_PATH = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "STKAITI.TTF"
REPORT_HEADER_PATH = WORKSPACE / "assets/images/条款表头.png"
REPORT_HEADER_FONT_SIZE = 72
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
MARGIN_BALANCE_MAX_DAILY_CHANGE_RATIO = 0.10
VALUATION_START_DATE = date(2019, 1, 1)
INDEX_TURNOVER_REQUEST_ATTEMPTS = 2
INDEX_TURNOVER_RETRY_DELAY_SECONDS = 3.0
INVERSE_CUBIC_VALUATION_NAME = "百元拟合溢价率"
MULTIFACTOR_VALUATION_NAME = "多因子修正百元拟合溢价率"
CB_INDEX_PARQUET_NAME = "转债指数"
VALUATION_REPAIR_WINDOW = 100
VALUATION_REPAIR_FORWARD_TRADING_DAYS = 70
VALUATION_REPAIR_INDEX_NAME = "转债估值修复指数"
VALUATION_REPAIR_FORWARD_RETURN_NAME = "中证转债指数后推70日涨跌幅"
VALUATION_REPAIR_PLOT_START_DATE = date(2024, 1, 1)
PARITY_GROUP_VALUATION_PREFIX = "多因子修正拟合溢价率_分平价_"
PARITY_GROUP_SPECS = tuple(
    (group, f"{PARITY_GROUP_VALUATION_PREFIX}{group}")
    for group in ("70-90", "90-110", "110-130", "130-150")
)
PARITY_INTERVAL_PREMIUM_COLUMNS = (
    "130以上",
    "120-130（含130）",
    "110-120（含120）",
    "100-110（含110）",
    "90-100（含100）",
    "80-90（含90）",
    "80以下（含80）",
)
PARITY_INTERVAL_PREMIUM_ASCENDING_COLUMNS = tuple(
    reversed(PARITY_INTERVAL_PREMIUM_COLUMNS)
)
EQUITY_BOND_GROUP_PREFIX = "多因子修正拟合溢价率_股债型_"
EQUITY_BOND_GROUP_SPECS = tuple(
    (group, f"{EQUITY_BOND_GROUP_PREFIX}{group}")
    for group in ("偏股型", "平衡型", "偏债型")
)
EQUITY_BOND_TYPE_GROUPS = ("偏股型", "平衡型", "偏债型")
RATING_GROUP_PREFIX = "多因子修正拟合溢价率_分评级_"
RATING_GROUP_SPECS = tuple(
    (group, f"{RATING_GROUP_PREFIX}{group}")
    for group in ("AAA/AA+", "AA/AA-", "A+/A")
)
MATURITY_GROUP_PREFIX = "多因子修正拟合溢价率_剩余期限_"
MATURITY_GROUP_SPECS = tuple(
    (group, f"{MATURITY_GROUP_PREFIX}{group}")
    for group in ("0-1", "1-2", "2-3", "3-4", "4-5", "5-6")
)
BALANCE_GROUP_PREFIX = "多因子修正拟合溢价率_分余额_"
BALANCE_GROUP_SPECS = (
    ("0-3", f"{BALANCE_GROUP_PREFIX}0-3亿元"),
    ("3-10", f"{BALANCE_GROUP_PREFIX}3-10亿元"),
    ("10-20", f"{BALANCE_GROUP_PREFIX}10-20亿元"),
    ("20-50", f"{BALANCE_GROUP_PREFIX}20-50亿元"),
    ("50+", f"{BALANCE_GROUP_PREFIX}50亿元以上"),
)
MARKET_CAP_GROUP_PREFIX = "多因子修正拟合溢价率_正股市值_"
MARKET_CAP_GROUP_SPECS = (
    ("0-50", f"{MARKET_CAP_GROUP_PREFIX}0-50亿元"),
    ("50-300", f"{MARKET_CAP_GROUP_PREFIX}50-300亿元"),
    ("300+", f"{MARKET_CAP_GROUP_PREFIX}300亿元以上"),
)
SECTOR_GROUP_SPECS = tuple(
    (group, f"多因子修正拟合溢价率_{group}")
    for group in ("科技", "金融", "制造", "消费", "周期")
)
SECTOR_ORDER = ("科技", "金融", "制造", "消费", "周期")
SECTOR_INDUSTRIES = {
    "科技": ("传媒", "电子", "国防军工", "计算机", "通信"),
    "金融": ("非银金融", "银行"),
    "制造": ("电力设备", "机械设备", "汽车", "轻工制造"),
    "消费": (
        "农林牧渔",
        "纺织服饰",
        "家用电器",
        "商贸零售",
        "社会服务",
        "食品饮料",
        "医药生物",
        "美容护理",
    ),
    "周期": (
        "基础化工",
        "钢铁",
        "公用事业",
        "环保",
        "建筑材料",
        "建筑装饰",
        "交通运输",
        "煤炭",
        "石油石化",
        "有色金属",
    ),
}
SECTOR_MEAN_METRICS = (
    ("收盘价", "各行业平均收盘价", ""),
    ("平价", "各行业平均平价", ""),
    ("转股溢价率", "各行业平均转股溢价率", "%"),
    ("纯债溢价率", "各行业平均纯债溢价率", "%"),
)
SECTOR_MEAN_PLOT_START_DATE = pd.Timestamp("2023-01-01")
INDUSTRY_INDEX_SPECS = (
    ("801960.SI", "石油石化"),
    ("801010.SI", "农林牧渔"),
    ("801030.SI", "基础化工"),
    ("801130.SI", "纺织服饰"),
    ("801200.SI", "商贸零售"),
    ("801040.SI", "钢铁"),
    ("801180.SI", "房地产"),
    ("801170.SI", "交通运输"),
    ("801210.SI", "社会服务"),
    ("801050.SI", "有色金属"),
    ("801710.SI", "建筑材料"),
    ("801120.SI", "食品饮料"),
    ("801140.SI", "轻工制造"),
    ("801720.SI", "建筑装饰"),
    ("801950.SI", "煤炭"),
    ("801760.SI", "传媒"),
    ("801980.SI", "美容护理"),
    ("801740.SI", "国防军工"),
    ("801880.SI", "汽车"),
    ("801750.SI", "计算机"),
    ("801160.SI", "公用事业"),
    ("801970.SI", "环保"),
    ("801790.SI", "非银金融"),
    ("801110.SI", "家用电器"),
    ("801780.SI", "银行"),
    ("801730.SI", "电力设备"),
    ("801150.SI", "医药生物"),
    ("801890.SI", "机械设备"),
    ("801770.SI", "通信"),
    ("801080.SI", "电子"),
)
INDUSTRY_HISTORY_ORDER = (
    "农林牧渔",
    "基础化工",
    "传媒",
    "电力设备",
    "电子",
    "房地产",
    "纺织服饰",
    "非银金融",
    "钢铁",
    "公用事业",
    "国防军工",
    "环保",
    "机械设备",
    "计算机",
    "家用电器",
    "建筑材料",
    "建筑装饰",
    "交通运输",
    "煤炭",
    "汽车",
    "轻工制造",
    "商贸零售",
    "社会服务",
    "石油石化",
    "食品饮料",
    "通信",
    "医药生物",
    "银行",
    "有色金属",
    "美容护理",
)
INDUSTRY_HISTORY_METRICS = ("收盘价", "转股溢价率", "平价", "纯债溢价率")
INDUSTRY_EXTREME_PRICE_THRESHOLD = 150.0
INDUSTRY_EXTREME_PREMIUM_THRESHOLD = 50.0
CLOSE_PRICE_DISTRIBUTION_LABELS = (
    "80以下（含80）",
    "80-90（含90）",
    "90-100（含100）",
    "100-110（含110）",
    "110-120（含120）",
    "120-130（含130）",
    "130-150（含150）",
    "150以上",
)
CLOSE_PRICE_DISTRIBUTION_BINS = (
    -np.inf,
    80.0,
    90.0,
    100.0,
    110.0,
    120.0,
    130.0,
    150.0,
    np.inf,
)
INTRADAY_VALUATION_SHEET = "百元平价拟合溢价率"
MAIN_MONEY_FLOW_WSET_FORMULA = (
    '=@wset("marketmoneyflows","startdate="&K5,'
    '"enddate="&K6,"frequency="&K7,"sector="&K8,'
    '"securitytype="&K10,"field=date,mainInflowMoney",'
    '"cols=2;rows=65")'
)
ETF_SHARE_START_SERIAL = 43831
ETF_SHARE_START_DATE = date(2020, 1, 1)
ETF_SHARE_SPECS = (
    ("博时可转债ETF", "511380.OF"),
    ("海富通可转债ETF", "511180.OF"),
)
ETF_SHARE_WSD_FORMULA = (
    '=@WSD(B4:C4,"unit_fundshare_total",B1,B2,'
    '"TradingCalendar=SSE","rptType=1","Version=1",'
    '"cols=2;rows=1614")'
)

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


function excelDateTime(value) {
  return new Date(`${value}Z`);
}


function newestFirst(rows) {
  return rows.slice().reverse();
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
  const rows = newestFirst(payload.market).map((row) => [excelDate(row.date), row.balance]);
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


function addMainMoneyFlowSheet() {
  const sheet = workbook.worksheets.add("主力净流入");
  sheet.getRange("A1:B1").merge();
  sheet.getRange("A1").values = [["沪深两市主力净流入"]];
  sheet.getRange("A2:B2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：${payload.mainMoneyFlowSource.source}｜查询区间：${payload.mainMoneyFlowSource.startDate} 至 ${payload.mainMoneyFlowSource.latestDate}｜${payload.mainMoneyFlowSource.unitRule}`,
  ]];
  sheet.getRange("A3:B3").values = [["交易日期", "主力净流入（亿元）"]];
  const rows = newestFirst(payload.mainMoneyFlow).map((row) => [excelDate(row.date), row.amount]);
  sheet.getRangeByIndexes(3, 0, rows.length, 2).values = rows;
  const lastRow = rows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("A1:B1"),
    sheet.getRange("A2:B2"),
    sheet.getRange("A3:B3"),
    sheet.getRange(`A4:B${lastRow}`),
  );
  sheet.getRange(`A4:A${lastRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`B4:B${lastRow}`).format.numberFormat = "#,##0.00";
  sheet.getRange(`A4:A${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B4:B${lastRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`A1:A${lastRow}`).format.columnWidth = 16;
  sheet.getRange(`B1:B${lastRow}`).format.columnWidth = 28;

  sheet.getRange("J4:K4").merge();
  sheet.getRange("J4").values = [["WSET查询参数"]];
  sheet.getRange("J4:K4").format = {
    fill: REPORT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
  };
  sheet.getRange("J5:J10").values = [
    ["开始日期"],
    ["截止日期"],
    ["频率"],
    ["市场"],
    [""],
    ["证券类型"],
  ];
  sheet.getRange("K5:K10").values = [
    [payload.mainMoneyFlowSource.startDate.replaceAll("-", "")],
    [payload.mainMoneyFlowSource.latestDate.replaceAll("-", "")],
    [payload.mainMoneyFlowSource.frequency],
    [payload.mainMoneyFlowSource.sector],
    [null],
    [payload.mainMoneyFlowSource.securityType],
  ];
  sheet.getRange("J5:K10").format = {
    fill: "#F2F2F2",
    font: { name: BODY_FONT, color: "#000000", size: 10 },
    borders: { preset: "all", style: "thin", color: GRID },
    verticalAlignment: "center",
  };
  sheet.getRange("J5:J10").format.horizontalAlignment = "left";
  sheet.getRange("K5:K10").format.horizontalAlignment = "right";
  sheet.getRange("K5:K6").format.numberFormat = "@";
  sheet.getRange("J12:K12").merge();
  sheet.getRange("J12").values = [["WSET公式（文本）"]];
  sheet.getRange("J12:K12").format = {
    fill: LIGHT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#404040", size: 9 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  sheet.getRange("J13:K15").merge();
  sheet.getRange("J13").values = [[`'${payload.mainMoneyFlowSource.formula}`]];
  sheet.getRange("J13:K15").format = {
    fill: "#F2F2F2",
    font: { name: BODY_FONT, color: "#000000", size: 9 },
    borders: { preset: "all", style: "thin", color: GRID },
    horizontalAlignment: "left",
    verticalAlignment: "top",
    wrapText: true,
  };
  sheet.getRange("J1:J15").format.columnWidth = 18;
  sheet.getRange("K1:K15").format.columnWidth = 34;
}


function addEtfShareSheet() {
  const sheet = workbook.worksheets.add("ETF份额与净申赎");
  sheet.getRange("A1:E1").merge();
  sheet.getRange("A1").values = [["可转债ETF份额与净申赎"]];
  sheet.getRange("A2:E2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：${payload.etfShareSource.source}｜查询区间：${payload.etfShareSource.startDate} 至 ${payload.etfShareSource.requestedEndDate}｜最新有效数据：${payload.etfShareSource.latestDate}｜${payload.etfShareSource.unitRule}｜${payload.etfShareSource.netSubscriptionRule}`,
  ]];
  sheet.getRange("A3:E3").values = [[
    "交易日期",
    "博时可转债ETF份额（亿份）",
    "博时净申赎（亿份）",
    "海富通可转债ETF份额（亿份）",
    "海富通净申赎（亿份）",
  ]];
  const rows = newestFirst(payload.etfShare).map((row) => [
    excelDate(row.date),
    row.boshiShare,
    null,
    row.haifutongShare,
    null,
  ]);
  sheet.getRangeByIndexes(3, 0, rows.length, 5).values = rows;
  const lastRow = rows.length + 3;
  const lastFormulaRow = lastRow - 1;
  if (lastFormulaRow >= 4) {
    sheet.getRange("C4").formulasR1C1 = [[
      '=IF(OR(RC[-1]="",R[1]C[-1]=""),"",RC[-1]-R[1]C[-1])',
    ]];
    sheet.getRange(`C4:C${lastFormulaRow}`).fillDown();
    sheet.getRange("E4").formulasR1C1 = [[
      '=IF(OR(RC[-1]="",R[1]C[-1]=""),"",RC[-1]-R[1]C[-1])',
    ]];
    sheet.getRange(`E4:E${lastFormulaRow}`).fillDown();
  }
  styleSheet(
    sheet,
    sheet.getRange("A1:E1"),
    sheet.getRange("A2:E2"),
    sheet.getRange("A3:E3"),
    sheet.getRange(`A4:E${lastRow}`),
  );
  sheet.getRange(`A4:A${lastRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`B4:E${lastRow}`).format.numberFormat = "0.00";
  sheet.getRange(`A4:A${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B4:E${lastRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`A1:A${lastRow}`).format.columnWidth = 16;
  sheet.getRange(`B1:E${lastRow}`).format.columnWidth = 26;

  sheet.getRange("G1:H1").merge();
  sheet.getRange("G1").values = [["最新指标摘要"]];
  sheet.getRange("G2:G5").values = [
    ["博时份额（亿份）"],
    ["博时净申赎（亿份）"],
    ["海富通份额（亿份）"],
    ["海富通净申赎（亿份）"],
  ];
  sheet.getRange("H2").formulas = [["=B4"]];
  sheet.getRange("H3").formulas = [["=C4"]];
  sheet.getRange("H4").formulas = [["=D4"]];
  sheet.getRange("H5").formulas = [["=E4"]];
  sheet.getRange("G1:H1").format = {
    fill: REPORT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  sheet.getRange("G2:H5").format = {
    fill: "#F2F2F2",
    font: { name: BODY_FONT, color: "#000000", size: 10 },
    borders: { preset: "all", style: "thin", color: GRID },
    verticalAlignment: "center",
  };
  sheet.getRange("H2:H5").format.numberFormat = "0.00";
  sheet.getRange("G1:G5").format.columnWidth = 24;
  sheet.getRange("H1:H5").format.columnWidth = 18;

  sheet.getRange("J1:K1").merge();
  sheet.getRange("J1").values = [["Wind WSD查询参数"]];
  sheet.getRange("J2:J7").values = [
    ["起始日期Excel序列"],
    ["请求截止日期"],
    ["博时代码"],
    ["海富通代码"],
    ["指标"],
    ["公式（文本）"],
  ];
  sheet.getRange("K2:K7").values = [
    [43831],
    [payload.etfShareSource.requestedEndDate],
    ["511380.OF"],
    ["511180.OF"],
    [payload.etfShareSource.field],
    [`'${payload.etfShareSource.formula}`],
  ];
  sheet.getRange("J1:K1").format = {
    fill: REPORT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  sheet.getRange("J2:K7").format = {
    fill: "#F2F2F2",
    font: { name: BODY_FONT, color: "#000000", size: 9 },
    borders: { preset: "all", style: "thin", color: GRID },
    verticalAlignment: "center",
    wrapText: true,
  };
  sheet.getRange("J1:J7").format.columnWidth = 22;
  sheet.getRange("K1:K7").format.columnWidth = 44;
}


function addSubnewBondSheet() {
  const sheet = workbook.worksheets.add("次新券表现");
  sheet.getRange("A1:E1").merge();
  sheet.getRange("A1").values = [["次新券相对上市表现与转股溢价率"]];
  sheet.getRange("A2:E2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：${payload.subnewBondSource.parquetRoot}、${payload.subnewBondSource.masterParquet}｜区间：${payload.subnewBondSource.startDate} 至 ${payload.subnewBondSource.latestDate}｜口径：${payload.subnewBondSource.sampleRule}`,
  ]];
  sheet.getRange("A3:E3").values = [[
    "交易日期",
    "相对上市涨跌幅均值（%）",
    "价格样本数",
    "平均转股溢价率（%）",
    "溢价率样本数",
  ]];
  const rows = newestFirst(payload.subnewBond).map((row) => [
    excelDate(row.date),
    row.listingReturnMean,
    row.listingReturnSampleCount,
    row.premiumMean,
    row.premiumSampleCount,
  ]);
  sheet.getRangeByIndexes(3, 0, rows.length, 5).values = rows;
  const lastRow = rows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("A1:E1"),
    sheet.getRange("A2:E2"),
    sheet.getRange("A3:E3"),
    sheet.getRange(`A4:E${lastRow}`),
  );
  sheet.getRange(`A4:A${lastRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`B4:B${lastRow}`).format.numberFormat = "0.00";
  sheet.getRange(`C4:C${lastRow}`).format.numberFormat = "#,##0";
  sheet.getRange(`D4:D${lastRow}`).format.numberFormat = "0.00";
  sheet.getRange(`E4:E${lastRow}`).format.numberFormat = "#,##0";
  sheet.getRange(`A4:A${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B4:E${lastRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`A1:A${lastRow}`).format.columnWidth = 16;
  sheet.getRange(`B1:B${lastRow}`).format.columnWidth = 27;
  sheet.getRange(`C1:C${lastRow}`).format.columnWidth = 14;
  sheet.getRange(`D1:D${lastRow}`).format.columnWidth = 25;
  sheet.getRange(`E1:E${lastRow}`).format.columnWidth = 14;

  sheet.getRange("G1:H1").merge();
  sheet.getRange("G1").values = [["最新指标摘要"]];
  sheet.getRange("G2:G7").values = [[
    "相对上市涨跌幅均值（%）",
  ], [
    "相对上市涨跌幅日变动（pct）",
  ], [
    "平均转股溢价率（%）",
  ], [
    "平均转股溢价率日变动（pct）",
  ], [
    "价格样本数",
  ], [
    "溢价率样本数",
  ]];
  sheet.getRange("H2").formulas = [["=B4"]];
  sheet.getRange("H3").formulas = [["=B4-B5"]];
  sheet.getRange("H4").formulas = [["=D4"]];
  sheet.getRange("H5").formulas = [["=D4-D5"]];
  sheet.getRange("H6").formulas = [["=C4"]];
  sheet.getRange("H7").formulas = [["=E4"]];
  sheet.getRange("G1:H1").format = {
    fill: REPORT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
  };
  sheet.getRange("G2:H7").format = {
    fill: "#F2F2F2",
    font: { name: BODY_FONT, color: "#000000", size: 10 },
    borders: { preset: "all", style: "thin", color: GRID },
    verticalAlignment: "center",
  };
  sheet.getRange("G2:G7").format.horizontalAlignment = "left";
  sheet.getRange("H2:H7").format.horizontalAlignment = "right";
  sheet.getRange("H2:H5").format.numberFormat = "0.00";
  sheet.getRange("H6:H7").format.numberFormat = "#,##0";
  sheet.getRange("G1:G7").format.columnWidth = 33;
  sheet.getRange("H1:H7").format.columnWidth = 18;
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
  const rows = newestFirst(payload.index).map((row) => [
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


function addValuationSheet() {
  const sheet = workbook.worksheets.add("转债估值");
  const dailyRows = newestFirst(payload.valuationDaily).map((row) => [
    excelDate(row.date),
    row.inverseCubic,
    row.multifactor,
  ]);
  const intradayRows = payload.valuationIntraday.map((row) => [
    excelDateTime(row.datetime),
    row.premium,
  ]);
  const priceParityRows = newestFirst(payload.priceParity).map((row) => [
    excelDate(row.date),
    row.weightedParity,
    row.medianClose,
    row.paritySampleCount,
    row.priceSampleCount,
    row.effectiveBalance,
  ]);
  const parityGroupRows = newestFirst(payload.parityGroupValuation).map((row) => [
    excelDate(row.date),
    row.group70_90,
    row.group90_110,
    row.group110_130,
    row.group130_150,
  ]);
  if (
    dailyRows.length < 2 ||
    intradayRows.length < 1 ||
    priceParityRows.length < 2 ||
    parityGroupRows.length < 2
  ) {
    throw new Error("转债估值底稿数据不足");
  }

  sheet.getRange("A1:C1").merge();
  sheet.getRange("A1").values = [["百元拟合溢价率历史序列"]];
  sheet.getRange("A2:C2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：${payload.valuationSource.parquet}｜区间：${payload.valuationSource.startDate} 至 ${payload.valuationSource.latestDate}｜单位：%`,
  ]];
  sheet.getRange("A3:C3").values = [[
    "交易日期",
    "百元拟合溢价率（反三次，%）",
    "多因子修正百元拟合溢价率（%）",
  ]];
  sheet.getRangeByIndexes(3, 0, dailyRows.length, 3).values = dailyRows;
  const lastDailyRow = dailyRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("A1:C1"),
    sheet.getRange("A2:C2"),
    sheet.getRange("A3:C3"),
    sheet.getRange(`A4:C${lastDailyRow}`),
  );
  sheet.getRange(`A4:A${lastDailyRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`B4:C${lastDailyRow}`).format.numberFormat = "0.00";
  sheet.getRange(`A4:A${lastDailyRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B4:C${lastDailyRow}`).format.horizontalAlignment = "right";

  sheet.getRange("E1:F1").merge();
  sheet.getRange("E1").values = [["盘中百元平价拟合溢价率"]];
  sheet.getRange("E2:F2").merge();
  sheet.getRange("E2").values = [[
    `数据来源：${payload.intradayValuationSource.workbook}｜工作表：${payload.intradayValuationSource.sheet}｜前一交易日参考：${payload.valuationSource.previousDate}`,
  ]];
  sheet.getRange("E3:F3").values = [["盘中时点", "拟合溢价率（%）"]];
  sheet.getRangeByIndexes(3, 4, intradayRows.length, 2).values = intradayRows;
  const lastIntradayRow = intradayRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("E1:F1"),
    sheet.getRange("E2:F2"),
    sheet.getRange("E3:F3"),
    sheet.getRange(`E4:F${lastIntradayRow}`),
  );
  sheet.getRange(`E4:E${lastIntradayRow}`).format.numberFormat = "yyyy-mm-dd hh:mm";
  sheet.getRange(`F4:F${lastIntradayRow}`).format.numberFormat = "0.00";
  sheet.getRange(`E4:E${lastIntradayRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`F4:F${lastIntradayRow}`).format.horizontalAlignment = "right";

  sheet.getRange("H1:I1").merge();
  sheet.getRange("H1").values = [["三次反比例最新摘要"]];
  sheet.getRange("H2:H5").values = [["最新值（%）"], ["环比变动（pct）"], ["2019年以来分位数"], ["前一交易日参考值（%）"]];
  sheet.getRange("I2").formulas = [["=B4"]];
  sheet.getRange("I3").formulas = [["=B4-B5"]];
  sheet.getRange("I4").formulas = [[`=COUNTIF(B4:B${lastDailyRow},"<="&I2)/COUNT(B4:B${lastDailyRow})`]];
  sheet.getRange("I5").formulas = [["=B5"]];
  sheet.getRange("H1:I1").format = {
    fill: REPORT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
  };
  sheet.getRange("H2:I5").format = {
    fill: "#F2F2F2",
    font: { name: BODY_FONT, color: "#000000", size: 10 },
    borders: { preset: "all", style: "thin", color: GRID },
    verticalAlignment: "center",
  };
  sheet.getRange("H2:H5").format.horizontalAlignment = "left";
  sheet.getRange("I2:I5").format.horizontalAlignment = "right";
  sheet.getRange("I2:I3").format.numberFormat = "0.00";
  sheet.getRange("I4").format.numberFormat = "0.00%";
  sheet.getRange("I5").format.numberFormat = "0.00";

  sheet.getRange("K1:P1").merge();
  sheet.getRange("K1").values = [["余额加权平价与收盘价中位数"]];
  sheet.getRange("K2:P2").merge();
  sheet.getRange("K2").values = [[
    `数据来源：${payload.priceParitySource.parquetRoot}｜区间：${payload.priceParitySource.startDate} 至 ${payload.priceParitySource.latestDate}｜口径：${payload.priceParitySource.sampleRule}`,
  ]];
  sheet.getRange("K3:P3").values = [[
    "交易日期",
    "余额加权平价",
    "收盘价中位数",
    "平价样本数",
    "价格样本数",
    "有效余额",
  ]];
  sheet.getRangeByIndexes(3, 10, priceParityRows.length, 6).values = priceParityRows;
  const lastPriceParityRow = priceParityRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("K1:P1"),
    sheet.getRange("K2:P2"),
    sheet.getRange("K3:P3"),
    sheet.getRange(`K4:P${lastPriceParityRow}`),
  );
  sheet.getRange(`K4:K${lastPriceParityRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`L4:M${lastPriceParityRow}`).format.numberFormat = "0.00";
  sheet.getRange(`N4:O${lastPriceParityRow}`).format.numberFormat = "#,##0";
  sheet.getRange(`P4:P${lastPriceParityRow}`).format.numberFormat = "#,##0.00";
  sheet.getRange(`K4:K${lastPriceParityRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`L4:P${lastPriceParityRow}`).format.horizontalAlignment = "right";

  sheet.getRange("R1:S1").merge();
  sheet.getRange("R1").values = [["价格与平价最新摘要"]];
  sheet.getRange("R2:R6").values = [["平均平价"], ["平均平价日变动"], ["价格中位数"], ["价格中位数日变动"], ["2019年以来分位数"]];
  sheet.getRange("S2").formulas = [["=L4"]];
  sheet.getRange("S3").formulas = [["=L4/L5-1"]];
  sheet.getRange("S4").formulas = [["=M4"]];
  sheet.getRange("S5").formulas = [["=M4/M5-1"]];
  sheet.getRange("S6").formulas = [[`=COUNTIF(M4:M${lastPriceParityRow},"<="&S4)/COUNT(M4:M${lastPriceParityRow})`]];
  sheet.getRange("R1:S1").format = {
    fill: REPORT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
  };
  sheet.getRange("R2:S6").format = {
    fill: "#F2F2F2",
    font: { name: BODY_FONT, color: "#000000", size: 10 },
    borders: { preset: "all", style: "thin", color: GRID },
    verticalAlignment: "center",
  };
  sheet.getRange("R2:R6").format.horizontalAlignment = "left";
  sheet.getRange("S2:S6").format.horizontalAlignment = "right";
  sheet.getRange("S2").format.numberFormat = "0.00";
  sheet.getRange("S3").format.numberFormat = "0.00%";
  sheet.getRange("S4").format.numberFormat = "0.00";
  sheet.getRange("S5:S6").format.numberFormat = "0.00%";

  sheet.getRange("U1:Y1").merge();
  sheet.getRange("U1").values = [["分平价多因子修正拟合溢价率"]];
  sheet.getRange("U2:Y2").merge();
  sheet.getRange("U2").values = [[
    `数据来源：${payload.parityGroupValuationSource.parquet}｜区间：${payload.parityGroupValuationSource.startDate} 至 ${payload.parityGroupValuationSource.runDate}｜口径：${payload.parityGroupValuationSource.readRule}`,
  ]];
  sheet.getRange("U3:Y3").values = [[
    "交易日期",
    "70-90（%）",
    "90-110（%）",
    "110-130（%）",
    "130-150（%）",
  ]];
  sheet.getRangeByIndexes(3, 20, parityGroupRows.length, 5).values = parityGroupRows;
  const lastParityGroupRow = parityGroupRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("U1:Y1"),
    sheet.getRange("U2:Y2"),
    sheet.getRange("U3:Y3"),
    sheet.getRange(`U4:Y${lastParityGroupRow}`),
  );
  sheet.getRange(`U4:U${lastParityGroupRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`V4:Y${lastParityGroupRow}`).format.numberFormat = "0.00";
  sheet.getRange(`U4:U${lastParityGroupRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`V4:Y${lastParityGroupRow}`).format.horizontalAlignment = "right";

  sheet.getRange(`A1:A${lastDailyRow}`).format.columnWidth = 16;
  sheet.getRange(`B1:C${lastDailyRow}`).format.columnWidth = 31;
  sheet.getRange(`D1:D${lastDailyRow}`).format.columnWidth = 3;
  sheet.getRange(`E1:E${lastIntradayRow}`).format.columnWidth = 22;
  sheet.getRange(`F1:F${lastIntradayRow}`).format.columnWidth = 22;
  sheet.getRange("G1:G10").format.columnWidth = 3;
  sheet.getRange("H1:H5").format.columnWidth = 25;
  sheet.getRange("I1:I5").format.columnWidth = 18;
  sheet.getRange("J1:J10").format.columnWidth = 3;
  sheet.getRange(`K1:K${lastPriceParityRow}`).format.columnWidth = 16;
  sheet.getRange(`L1:M${lastPriceParityRow}`).format.columnWidth = 21;
  sheet.getRange(`N1:O${lastPriceParityRow}`).format.columnWidth = 16;
  sheet.getRange(`P1:P${lastPriceParityRow}`).format.columnWidth = 18;
  sheet.getRange("Q1:Q10").format.columnWidth = 3;
  sheet.getRange("R1:R6").format.columnWidth = 24;
  sheet.getRange("S1:S6").format.columnWidth = 18;
  sheet.getRange("T1:T10").format.columnWidth = 3;
  sheet.getRange(`U1:U${lastParityGroupRow}`).format.columnWidth = 16;
  sheet.getRange(`V1:Y${lastParityGroupRow}`).format.columnWidth = 21;
}


function addClassificationValuationSheet() {
  const sheet = workbook.worksheets.add("分类拟合溢价率");
  const equityBondRows = newestFirst(payload.equityBondGroupValuation).map((row) => [
    excelDate(row.date), row.stock, row.balance, row.bond,
  ]);
  const ratingRows = newestFirst(payload.ratingGroupValuation).map((row) => [
    excelDate(row.date), row.top, row.middle, row.lower,
  ]);
  const maturityRows = newestFirst(payload.maturityGroupValuation).map((row) => [
    excelDate(row.date), row.group0_1, row.group1_2, row.group2_3,
    row.group3_4, row.group4_5, row.group5_6,
  ]);
  const balanceRows = newestFirst(payload.balanceGroupValuation).map((row) => [
    excelDate(row.date), row.group0_3, row.group3_10, row.group10_20,
    row.group20_50, row.group50_plus,
  ]);
  const marketCapRows = newestFirst(payload.marketCapGroupValuation).map((row) => [
    excelDate(row.date), row.group0_50, row.group50_300, row.group300_plus,
  ]);
  if (
    equityBondRows.length < 2 || ratingRows.length < 2 || maturityRows.length < 2 ||
    balanceRows.length < 2 || marketCapRows.length < 2
  ) {
    throw new Error("分类拟合溢价率底稿数据不足");
  }

  sheet.getRange("A1:D1").merge();
  sheet.getRange("A1").values = [["股债型多因子修正拟合溢价率"]];
  sheet.getRange("A2:D2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：${payload.equityBondGroupValuationSource.parquet}｜区间：${payload.equityBondGroupValuationSource.startDate} 至 ${payload.equityBondGroupValuationSource.runDate}｜口径：${payload.equityBondGroupValuationSource.readRule}`,
  ]];
  sheet.getRange("A3:D3").values = [["交易日期", "偏股型（%）", "平衡型（%）", "偏债型（%）"]];
  sheet.getRangeByIndexes(3, 0, equityBondRows.length, 4).values = equityBondRows;
  const lastEquityBondRow = equityBondRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("A1:D1"),
    sheet.getRange("A2:D2"),
    sheet.getRange("A3:D3"),
    sheet.getRange(`A4:D${lastEquityBondRow}`),
  );
  sheet.getRange(`A4:A${lastEquityBondRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`B4:D${lastEquityBondRow}`).format.numberFormat = "0.00";

  sheet.getRange("F1:I1").merge();
  sheet.getRange("F1").values = [["分评级多因子修正拟合溢价率"]];
  sheet.getRange("F2:I2").merge();
  sheet.getRange("F2").values = [[
    `数据来源：${payload.ratingGroupValuationSource.parquet}｜区间：${payload.ratingGroupValuationSource.startDate} 至 ${payload.ratingGroupValuationSource.runDate}｜口径：${payload.ratingGroupValuationSource.readRule}`,
  ]];
  sheet.getRange("F3:I3").values = [[
    "交易日期", "AAA/AA+（%）", "AA/AA-（%）", "A+/A（%）",
  ]];
  sheet.getRangeByIndexes(3, 5, ratingRows.length, 4).values = ratingRows;
  const lastRatingRow = ratingRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("F1:I1"),
    sheet.getRange("F2:I2"),
    sheet.getRange("F3:I3"),
    sheet.getRange(`F4:I${lastRatingRow}`),
  );
  sheet.getRange(`F4:F${lastRatingRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`G4:I${lastRatingRow}`).format.numberFormat = "0.00";

  sheet.getRange("T1:Z1").merge();
  sheet.getRange("T1").values = [["分剩余期限多因子修正拟合溢价率"]];
  sheet.getRange("T2:Z2").merge();
  sheet.getRange("T2").values = [[
    `数据来源：${payload.maturityGroupValuationSource.parquet}｜区间：${payload.maturityGroupValuationSource.startDate} 至 ${payload.maturityGroupValuationSource.runDate}｜口径：${payload.maturityGroupValuationSource.readRule}`,
  ]];
  sheet.getRange("T3:Z3").values = [[
    "交易日期", "0-1（%）", "1-2（%）", "2-3（%）", "3-4（%）", "4-5（%）", "5-6（%）",
  ]];
  sheet.getRangeByIndexes(3, 19, maturityRows.length, 7).values = maturityRows;
  const lastMaturityRow = maturityRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("T1:Z1"),
    sheet.getRange("T2:Z2"),
    sheet.getRange("T3:Z3"),
    sheet.getRange(`T4:Z${lastMaturityRow}`),
  );
  sheet.getRange(`T4:T${lastMaturityRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`U4:Z${lastMaturityRow}`).format.numberFormat = "0.00";

  sheet.getRange("AB1:AC1").merge();
  sheet.getRange("AB1").values = [["剩余期限标题摘要"]];
  sheet.getRange("AB2:AB5").values = [["5-6"], ["5-6日变动"], ["0-1"], ["0-1日变动"]];
  sheet.getRange("AC2").formulas = [["=Z4"]];
  sheet.getRange("AC3").formulas = [["=Z4-Z5"]];
  sheet.getRange("AC4").formulas = [["=U4"]];
  sheet.getRange("AC5").formulas = [["=U4-U5"]];
  sheet.getRange("AB1:AC1").format = {
    fill: REPORT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
  };
  sheet.getRange("AB2:AC5").format = {
    fill: "#F2F2F2", font: { name: BODY_FONT, color: "#000000", size: 10 },
    borders: { preset: "all", style: "thin", color: GRID },
  };
  sheet.getRange("AC2:AC5").format.numberFormat = "0.00";

  sheet.getRange("AE1:AJ1").merge();
  sheet.getRange("AE1").values = [["分余额多因子修正拟合溢价率"]];
  sheet.getRange("AE2:AJ2").merge();
  sheet.getRange("AE2").values = [[
    `数据来源：${payload.balanceGroupValuationSource.parquet}｜区间：${payload.balanceGroupValuationSource.startDate} 至 ${payload.balanceGroupValuationSource.runDate}｜口径：${payload.balanceGroupValuationSource.readRule}`,
  ]];
  sheet.getRange("AE3:AJ3").values = [[
    "交易日期", "0-3（%）", "3-10（%）", "10-20（%）", "20-50（%）", "50+（%）",
  ]];
  sheet.getRangeByIndexes(3, 30, balanceRows.length, 6).values = balanceRows;
  const lastBalanceRow = balanceRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("AE1:AJ1"),
    sheet.getRange("AE2:AJ2"),
    sheet.getRange("AE3:AJ3"),
    sheet.getRange(`AE4:AJ${lastBalanceRow}`),
  );
  sheet.getRange(`AE4:AE${lastBalanceRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`AF4:AJ${lastBalanceRow}`).format.numberFormat = "0.00";

  sheet.getRange("AL1:AO1").merge();
  sheet.getRange("AL1").values = [["分正股市值多因子修正拟合溢价率"]];
  sheet.getRange("AL2:AO2").merge();
  sheet.getRange("AL2").values = [[
    `数据来源：${payload.marketCapGroupValuationSource.parquet}｜区间：${payload.marketCapGroupValuationSource.startDate} 至 ${payload.marketCapGroupValuationSource.runDate}｜口径：${payload.marketCapGroupValuationSource.readRule}`,
  ]];
  sheet.getRange("AL3:AO3").values = [[
    "交易日期", "0-50（%）", "50-300（%）", "300+（%）",
  ]];
  sheet.getRangeByIndexes(3, 37, marketCapRows.length, 4).values = marketCapRows;
  const lastMarketCapRow = marketCapRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("AL1:AO1"),
    sheet.getRange("AL2:AO2"),
    sheet.getRange("AL3:AO3"),
    sheet.getRange(`AL4:AO${lastMarketCapRow}`),
  );
  sheet.getRange(`AL4:AL${lastMarketCapRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`AM4:AO${lastMarketCapRow}`).format.numberFormat = "0.00";

  sheet.getRange("AQ1:AR1").merge();
  sheet.getRange("AQ1").values = [["余额分类标题摘要"]];
  sheet.getRange("AQ2:AQ5").values = [["0-3"], ["0-3日变动"], ["50+"], ["50+日变动"]];
  sheet.getRange("AR2").formulas = [["=AF4"]];
  sheet.getRange("AR3").formulas = [["=AF4-AF5"]];
  sheet.getRange("AR4").formulas = [["=AJ4"]];
  sheet.getRange("AR5").formulas = [["=AJ4-AJ5"]];
  sheet.getRange("AT1:AU1").merge();
  sheet.getRange("AT1").values = [["市值分类标题摘要"]];
  sheet.getRange("AT2:AT5").values = [["0-50"], ["0-50日变动"], ["300+"], ["300+日变动"]];
  sheet.getRange("AU2").formulas = [["=AM4"]];
  sheet.getRange("AU3").formulas = [["=AM4-AM5"]];
  sheet.getRange("AU4").formulas = [["=AO4"]];
  sheet.getRange("AU5").formulas = [["=AO4-AO5"]];
  for (const range of ["AQ1:AR1", "AT1:AU1"]) {
    sheet.getRange(range).format = {
      fill: REPORT_BLUE,
      font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
    };
  }
  for (const range of ["AQ2:AR5", "AT2:AU5"]) {
    sheet.getRange(range).format = {
      fill: "#F2F2F2", font: { name: BODY_FONT, color: "#000000", size: 10 },
      borders: { preset: "all", style: "thin", color: GRID },
    };
  }
  sheet.getRange("AR2:AR5").format.numberFormat = "0.00";
  sheet.getRange("AU2:AU5").format.numberFormat = "0.00";

  sheet.getRange("N1:O1").merge();
  sheet.getRange("N1").values = [["图表标题摘要"]];
  sheet.getRange("N2:N6").values = [["偏股型"], ["偏股型日变动"], ["偏债型"], ["偏债型日变动"], ["AAA"]];
  sheet.getRange("O2").formulas = [["=B4"]];
  sheet.getRange("O3").formulas = [["=B4-B5"]];
  sheet.getRange("O4").formulas = [["=D4"]];
  sheet.getRange("O5").formulas = [["=D4-D5"]];
  sheet.getRange("O6").formulas = [["=G4"]];
  sheet.getRange("Q1:R1").merge();
  sheet.getRange("Q1").values = [["评级标题摘要"]];
  sheet.getRange("Q2:Q5").values = [["AAA/AA+日变动"], ["AA/AA-"], ["AA/AA-日变动"], ["数据来源"]];
  sheet.getRange("R2").formulas = [["=G4-G5"]];
  sheet.getRange("R3").formulas = [["=H4"]];
  sheet.getRange("R4").formulas = [["=H4-H5"]];
  sheet.getRange("R5").values = [[payload.ratingGroupValuationSource.parquet]];
  for (const range of ["N1:O1", "Q1:R1"]) {
    sheet.getRange(range).format = {
      fill: REPORT_BLUE,
      font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
    };
  }
  sheet.getRange("N2:O6").format = {
    fill: "#F2F2F2", font: { name: BODY_FONT, color: "#000000", size: 10 },
    borders: { preset: "all", style: "thin", color: GRID },
  };
  sheet.getRange("Q2:R5").format = {
    fill: "#F2F2F2", font: { name: BODY_FONT, color: "#000000", size: 10 },
    borders: { preset: "all", style: "thin", color: GRID },
  };
  sheet.getRange("O2:O6").format.numberFormat = "0.00";
  sheet.getRange("R2:R4").format.numberFormat = "0.00";
  sheet.getRange("R5").format.wrapText = true;

  sheet.getRange(`A1:A${lastEquityBondRow}`).format.columnWidth = 16;
  sheet.getRange(`B1:D${lastEquityBondRow}`).format.columnWidth = 18;
  sheet.getRange("E1:E10").format.columnWidth = 3;
  sheet.getRange(`F1:F${lastRatingRow}`).format.columnWidth = 16;
  sheet.getRange(`G1:I${lastRatingRow}`).format.columnWidth = 18;
  sheet.getRange("J1:J10").format.columnWidth = 3;
  sheet.getRange("N1:N6").format.columnWidth = 20;
  sheet.getRange("O1:O6").format.columnWidth = 18;
  sheet.getRange("P1:P10").format.columnWidth = 3;
  sheet.getRange("Q1:Q5").format.columnWidth = 20;
  sheet.getRange("R1:R5").format.columnWidth = 28;
  sheet.getRange("S1:S10").format.columnWidth = 3;
  sheet.getRange(`T1:T${lastMaturityRow}`).format.columnWidth = 16;
  sheet.getRange(`U1:Z${lastMaturityRow}`).format.columnWidth = 15;
  sheet.getRange("AA1:AA10").format.columnWidth = 3;
  sheet.getRange("AB1:AB5").format.columnWidth = 20;
  sheet.getRange("AC1:AC5").format.columnWidth = 18;
  sheet.getRange("AD1:AD10").format.columnWidth = 3;
  sheet.getRange(`AE1:AE${lastBalanceRow}`).format.columnWidth = 16;
  sheet.getRange(`AF1:AJ${lastBalanceRow}`).format.columnWidth = 15;
  sheet.getRange("AK1:AK10").format.columnWidth = 3;
  sheet.getRange(`AL1:AL${lastMarketCapRow}`).format.columnWidth = 16;
  sheet.getRange(`AM1:AO${lastMarketCapRow}`).format.columnWidth = 18;
  sheet.getRange("AP1:AP10").format.columnWidth = 3;
  sheet.getRange("AQ1:AQ5").format.columnWidth = 20;
  sheet.getRange("AR1:AR5").format.columnWidth = 18;
  sheet.getRange("AS1:AS10").format.columnWidth = 3;
  sheet.getRange("AT1:AT5").format.columnWidth = 20;
  sheet.getRange("AU1:AU5").format.columnWidth = 18;
}


function addEquityBondWeightedSheet() {
  const sheet = workbook.worksheets.add("股债性加权指标");
  const rows = newestFirst(payload.equityBondWeighted).map((row) => [
    excelDate(row.date),
    row.stockPremium, row.balancePremium, row.bondPremium,
    row.stockPrice, row.balancePrice, row.bondPrice,
  ]);
  if (rows.length < 2) {
    throw new Error("股债性分类余额加权底稿数据不足");
  }
  sheet.getRange("A1:G1").merge();
  sheet.getRange("A1").values = [["股债性分类余额加权转股溢价率与均价"]];
  sheet.getRange("A2:G2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：${payload.equityBondWeightedSource.parquetRoot}｜区间：${payload.equityBondWeightedSource.startDate} 至 ${payload.equityBondWeightedSource.runDate}｜口径：${payload.equityBondWeightedSource.sampleRule}`,
  ]];
  sheet.getRange("A3:G3").values = [[
    "交易日期", "偏股型转股溢价率（%）", "平衡型转股溢价率（%）", "偏债型转股溢价率（%）",
    "偏股型均价", "平衡型均价", "偏债型均价",
  ]];
  sheet.getRangeByIndexes(3, 0, rows.length, 7).values = rows;
  const lastRow = rows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("A1:G1"),
    sheet.getRange("A2:G2"),
    sheet.getRange("A3:G3"),
    sheet.getRange(`A4:G${lastRow}`),
  );
  sheet.getRange(`A4:A${lastRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`B4:G${lastRow}`).format.numberFormat = "0.00";
  sheet.getRange(`A1:A${lastRow}`).format.columnWidth = 16;
  sheet.getRange(`B1:D${lastRow}`).format.columnWidth = 20;
  sheet.getRange(`E1:G${lastRow}`).format.columnWidth = 16;
}


function addSectorAndPriceDistributionSheet() {
  const sheet = workbook.worksheets.add("板块与价格分布");
  const sectorRows = newestFirst(payload.sectorGroupValuation).map((row) => [
    excelDate(row.date), row.technology, row.finance, row.manufacturing,
    row.consumption, row.cyclical,
  ]);
  const distributionRows = newestFirst(payload.closePriceDistribution).map((row) => [
    excelDate(row.date), row.le80, row.p80_90, row.p90_100, row.p100_110,
    row.p110_120, row.p120_130, row.p130_150, row.gt150, row.sampleCount,
  ]);
  if (sectorRows.length < 2 || distributionRows.length < 2) {
    throw new Error("板块拟合溢价率或收盘价分布底稿数据不足");
  }

  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").values = [["分板块多因子修正拟合溢价率"]];
  sheet.getRange("A2:F2").merge();
  sheet.getRange("A2").values = [[
    `数据来源：${payload.sectorGroupValuationSource.parquet}｜区间：${payload.sectorGroupValuationSource.startDate} 至 ${payload.sectorGroupValuationSource.runDate}｜口径：${payload.sectorGroupValuationSource.readRule}`,
  ]];
  sheet.getRange("A3:F3").values = [[
    "交易日期", "科技（%）", "金融（%）", "制造（%）", "消费（%）", "周期（%）",
  ]];
  sheet.getRangeByIndexes(3, 0, sectorRows.length, 6).values = sectorRows;
  const lastSectorRow = sectorRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("A1:F1"),
    sheet.getRange("A2:F2"),
    sheet.getRange("A3:F3"),
    sheet.getRange(`A4:F${lastSectorRow}`),
  );
  sheet.getRange(`A4:A${lastSectorRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`B4:F${lastSectorRow}`).format.numberFormat = "0.00";

  sheet.getRange("H1:Q1").merge();
  sheet.getRange("H1").values = [["2019年以来收盘价八档分布"]];
  sheet.getRange("H2:Q2").merge();
  sheet.getRange("H2").values = [[
    `数据来源：${payload.closePriceDistributionSource.parquetRoot}｜区间：${payload.closePriceDistributionSource.startDate} 至 ${payload.closePriceDistributionSource.runDate}｜口径：${payload.closePriceDistributionSource.sampleRule}`,
  ]];
  sheet.getRange("H3:Q3").values = [[
    "交易日期", "≤80（%）", "80-90（%）", "90-100（%）", "100-110（%）",
    "110-120（%）", "120-130（%）", "130-150（%）", ">150（%）", "有效样本数",
  ]];
  sheet.getRangeByIndexes(3, 7, distributionRows.length, 10).values = distributionRows;
  const lastDistributionRow = distributionRows.length + 3;
  styleSheet(
    sheet,
    sheet.getRange("H1:Q1"),
    sheet.getRange("H2:Q2"),
    sheet.getRange("H3:Q3"),
    sheet.getRange(`H4:Q${lastDistributionRow}`),
  );
  sheet.getRange(`H4:H${lastDistributionRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`I4:P${lastDistributionRow}`).format.numberFormat = "0.00";
  sheet.getRange(`Q4:Q${lastDistributionRow}`).format.numberFormat = "#,##0";

  sheet.getRange("S1:T1").merge();
  sheet.getRange("S1").values = [["板块标题摘要"]];
  sheet.getRange("S2:S5").values = [["科技"], ["科技日变动"], ["周期"], ["周期日变动"]];
  sheet.getRange("T2").formulas = [["=B4"]];
  sheet.getRange("T3").formulas = [["=B4-B5"]];
  sheet.getRange("T4").formulas = [["=F4"]];
  sheet.getRange("T5").formulas = [["=F4-F5"]];
  sheet.getRange("V1:W1").merge();
  sheet.getRange("V1").values = [["收盘价分布标题摘要"]];
  sheet.getRange("V2:V5").values = [["破底"], ["破底日变动"], ["破面"], ["破面日变动"]];
  sheet.getRange("W2").formulas = [["=I4"]];
  sheet.getRange("W3").formulas = [["=I4-I5"]];
  sheet.getRange("W4").formulas = [["=SUM(I4:K4)"]];
  sheet.getRange("W5").formulas = [["=SUM(I4:K4)-SUM(I5:K5)"]];
  for (const range of ["S1:T1", "V1:W1"]) {
    sheet.getRange(range).format = {
      fill: REPORT_BLUE,
      font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 11 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      borders: { preset: "outside", style: "thin", color: REPORT_BLUE },
    };
  }
  for (const range of ["S2:T5", "V2:W5"]) {
    sheet.getRange(range).format = {
      fill: "#F2F2F2",
      font: { name: BODY_FONT, color: "#000000", size: 10 },
      borders: { preset: "all", style: "thin", color: GRID },
    };
  }
  sheet.getRange("T2:T5").format.numberFormat = "0.00";
  sheet.getRange("W2:W5").format.numberFormat = "0.00";

  sheet.getRange(`A1:A${lastSectorRow}`).format.columnWidth = 16;
  sheet.getRange(`B1:F${lastSectorRow}`).format.columnWidth = 17;
  sheet.getRange("G1:G10").format.columnWidth = 3;
  sheet.getRange(`H1:H${lastDistributionRow}`).format.columnWidth = 16;
  sheet.getRange(`I1:P${lastDistributionRow}`).format.columnWidth = 15;
  sheet.getRange(`Q1:Q${lastDistributionRow}`).format.columnWidth = 15;
  sheet.getRange("R1:R10").format.columnWidth = 3;
  sheet.getRange("S1:S5").format.columnWidth = 20;
  sheet.getRange("T1:T5").format.columnWidth = 18;
  sheet.getRange("U1:U10").format.columnWidth = 3;
  sheet.getRange("V1:V5").format.columnWidth = 22;
  sheet.getRange("W1:W5").format.columnWidth = 18;
}


function addSectorMeanMetricsSheet() {
  const sheet = workbook.worksheets.add("行业均值");
  const sectorKeys = ["technology", "finance", "manufacturing", "consumption", "cyclical"];
  const sectorLabels = ["科技", "金融", "制造", "消费", "周期"];
  const blocks = [
    {
      startColumn: 0, titleRange: "A1:F1", noteRange: "A2:F2", headerRange: "A3:F3",
      title: "各行业平均收盘价", key: "close", unit: "",
    },
    {
      startColumn: 7, titleRange: "H1:M1", noteRange: "H2:M2", headerRange: "H3:M3",
      title: "各行业平均平价", key: "parity", unit: "",
    },
    {
      startColumn: 14, titleRange: "O1:T1", noteRange: "O2:T2", headerRange: "O3:T3",
      title: "各行业平均转股溢价率", key: "conversionPremium", unit: "%",
    },
    {
      startColumn: 21, titleRange: "V1:AA1", noteRange: "V2:AA2", headerRange: "V3:AA3",
      title: "各行业平均纯债溢价率", key: "bondPremium", unit: "%",
    },
  ];
  const sourceNote =
    `数据来源：${payload.sectorMeanSource.parquetRoot}、${payload.sectorMeanSource.masterParquet}｜` +
    `区间：${payload.sectorMeanSource.startDate} 至 ${payload.sectorMeanSource.runDate}｜` +
    `板块：${payload.sectorMeanSource.sectorRule}｜口径：${payload.sectorMeanSource.sampleRule}`;
  for (const block of blocks) {
    const rows = newestFirst(payload.sectorMeanMetrics).map((row) => [
      excelDate(row.date),
      ...sectorKeys.map((sector) => row[`${block.key}_${sector}`]),
    ]);
    const lastRow = rows.length + 3;
    sheet.getRange(block.titleRange).merge();
    sheet.getRange(block.titleRange.split(":")[0]).values = [[block.title]];
    sheet.getRange(block.noteRange).merge();
    sheet.getRange(block.noteRange.split(":")[0]).values = [[sourceNote]];
    sheet.getRange(block.headerRange).values = [[
      "交易日期",
      ...sectorLabels.map((sector) => block.unit ? `${sector}（${block.unit}）` : sector),
    ]];
    sheet.getRangeByIndexes(3, block.startColumn, rows.length, 6).values = rows;
    styleSheet(
      sheet,
      sheet.getRange(block.titleRange),
      sheet.getRange(block.noteRange),
      sheet.getRange(block.headerRange),
      sheet.getRangeByIndexes(3, block.startColumn, rows.length, 6),
    );
    sheet.getRangeByIndexes(3, block.startColumn, rows.length, 1).format.numberFormat = "yyyy-mm-dd";
    sheet.getRangeByIndexes(3, block.startColumn + 1, rows.length, 5).format.numberFormat = "0.00";
    sheet.getRangeByIndexes(0, block.startColumn, lastRow, 1).format.columnWidth = 16;
    sheet.getRangeByIndexes(0, block.startColumn + 1, lastRow, 5).format.columnWidth = 16;
  }
  sheet.getRange("G1:G20").format.columnWidth = 3;
  sheet.getRange("N1:N20").format.columnWidth = 3;
  sheet.getRange("U1:U20").format.columnWidth = 3;
}


function addIndustryPerformanceSheet() {
  const sheet = workbook.worksheets.add("行业涨跌与估值");
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(2);
  sheet.getRange("A1:A2").merge();
  sheet.getRange("B1:E1").merge();
  sheet.getRange("F1:I1").merge();
  sheet.getRange("J1:O1").merge();
  sheet.getRange("A1").values = [["名称"]];
  sheet.getRange("B1").values = [["正股"]];
  sheet.getRange("F1").values = [["转债"]];
  sheet.getRange("J1").values = [["正股估值分位数"]];
  sheet.getRange("B2:O2").values = [[
    "日涨跌幅", "周涨跌幅", "月涨跌幅", "年初至今涨跌幅",
    "日涨跌幅", "周涨跌幅", "月涨跌幅", "年初至今涨跌幅",
    "PE（TTM）", "3年分位数", "10年分位数",
    "PB（LF）", "3年分位数", "10年分位数",
  ]];
  sheet.getRange("A1:O2").format = {
    fill: "#203864",
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 10 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#000000" },
  };
  sheet.getRange("A1:O1").format.rowHeight = 24;
  sheet.getRange("A2:O2").format.rowHeight = 32;

  const rows = payload.industryPerformance.map((row) => [
    row.name,
    row.stockDaily,
    row.stockWeekly,
    row.stockMonthly,
    row.stockYtd,
    row.bondDaily === null ? null : row.bondDaily / 100,
    row.bondWeekly === null ? null : row.bondWeekly / 100,
    row.bondMonthly === null ? null : row.bondMonthly / 100,
    row.bondYtd === null ? null : row.bondYtd / 100,
    row.pe,
    row.pe3y,
    row.pe10y,
    row.pb,
    row.pb3y,
    row.pb10y,
  ]);
  const lastRow = rows.length + 2;
  sheet.getRangeByIndexes(2, 0, rows.length, 15).values = rows;
  sheet.getRange(`A3:O${lastRow}`).format = {
    font: { name: BODY_FONT, color: "#000000", size: 10 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#000000" },
  };
  sheet.getRange(`A3:O${lastRow}`).format.rowHeight = 21;
  sheet.getRange(`B3:E${lastRow}`).format.numberFormat = "0.00";
  sheet.getRange(`F3:I${lastRow}`).format.numberFormat = "0.00%";
  sheet.getRange(`J3:O${lastRow}`).format.numberFormat = "0.00";
  sheet.getRange(`B3:E${lastRow}`).conditionalFormats.add("colorScale", {
    colors: ["#5B9BD5", "#FFFFFF", "#F8696B"],
    thresholds: ["min", 0, "max"],
  });
  sheet.getRange(`F3:I${lastRow}`).conditionalFormats.add("colorScale", {
    colors: ["#63BE7B", "#FFEB84", "#F8696B"],
    thresholds: ["min", 0, "max"],
  });
  for (const range of [`B3:E${lastRow}`, `F3:I${lastRow}`]) {
    sheet.getRange(range).conditionalFormats.add("cellIs", {
      operator: "lessThan",
      formula: 0,
      format: { font: { color: "#FF0000" } },
    });
  }
  for (const range of [`K3:K${lastRow}`, `N3:N${lastRow}`]) {
    sheet.getRange(range).conditionalFormats.add("dataBar", {
      color: "#5B9BD5",
      gradient: true,
    });
  }
  for (const range of [`L3:L${lastRow}`, `O3:O${lastRow}`]) {
    sheet.getRange(range).conditionalFormats.add("dataBar", {
      color: "#FF6666",
      gradient: true,
    });
  }
  sheet.getRange(`A1:A${lastRow}`).format.columnWidth = 14;
  sheet.getRange(`B1:I${lastRow}`).format.columnWidth = 12;
  sheet.getRange(`J1:J${lastRow}`).format.columnWidth = 11;
  sheet.getRange(`K1:L${lastRow}`).format.columnWidth = 13;
  sheet.getRange(`M1:M${lastRow}`).format.columnWidth = 11;
  sheet.getRange(`N1:O${lastRow}`).format.columnWidth = 13;

  sheet.getRange("Q1:AB1").merge();
  sheet.getRange("Q1").values = [["行业代码、Wind公式与转债样本数（审计）"]];
  sheet.getRange("Q2:AB2").values = [[
    "行业代码", "日涨跌幅公式", "周涨跌幅公式", "月涨跌幅公式", "年初至今公式",
    "PE公式", "PE三年公式", "PE十年公式", "PB公式", "PB三年公式", "PB十年公式",
    "转债当日样本数",
  ]];
  const auditRows = payload.industryPerformance.map((row) => [
    row.code,
    `公式：${row.stockDailyFormula}`,
    `公式：${row.stockWeeklyFormula}`,
    `公式：${row.stockMonthlyFormula}`,
    `公式：${row.stockYtdFormula}`,
    `公式：${row.peFormula}`,
    `公式：${row.pe3yFormula}`,
    `公式：${row.pe10yFormula}`,
    `公式：${row.pbFormula}`,
    `公式：${row.pb3yFormula}`,
    `公式：${row.pb10yFormula}`,
    row.bondSampleCount,
  ]);
  sheet.getRangeByIndexes(2, 16, auditRows.length, 12).values = auditRows;
  sheet.getRange(`Q1:AB2`).format = {
    fill: REPORT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#FFFFFF", size: 9 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: REPORT_BLUE },
  };
  sheet.getRange(`Q3:AB${lastRow}`).format = {
    font: { name: BODY_FONT, color: "#000000", size: 9 },
    borders: { preset: "all", style: "thin", color: GRID },
    wrapText: false,
  };
  sheet.getRange(`Q1:Q${lastRow}`).format.columnWidth = 13;
  sheet.getRange(`R1:AA${lastRow}`).format.columnWidth = 34;
  sheet.getRange(`AB1:AB${lastRow}`).format.columnWidth = 14;

  const noteRow = lastRow + 2;
  sheet.getRange(`A${noteRow}:O${noteRow + 1}`).merge();
  sheet.getRange(`A${noteRow}`).values = [[
    `数据来源：正股为${payload.industryPerformanceSource.equity.source}；` +
    `转债为${payload.industryPerformanceSource.convertibleBond.parquetRoot}。` +
    `排序：${payload.industryPerformanceSource.sortRule}。` +
    `转债口径：${payload.industryPerformanceSource.convertibleBond.sampleRule}；` +
    `${payload.industryPerformanceSource.convertibleBond.returnRule}`,
  ]];
  sheet.getRange(`A${noteRow}:O${noteRow + 1}`).format = {
    fill: LIGHT_BLUE,
    font: { name: BODY_FONT, color: "#404040", size: 9 },
    verticalAlignment: "center",
    wrapText: true,
  };
}


function addPureDataSheet(name, headers, rows, formats = {}, dateFormats = { 0: "yyyy-mm-dd" }) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  sheet.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
  sheet.getRangeByIndexes(0, 0, 1, headers.length).format = {
    font: { name: BODY_FONT, bold: true, color: "#000000", size: 10 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "bottom", style: "thin", color: "#000000" },
  };
  sheet.getRangeByIndexes(0, 0, 1, headers.length).format.rowHeight = 20;
  if (rows.length > 0) {
    sheet.getRangeByIndexes(1, 0, rows.length, headers.length).values = rows;
    sheet.getRangeByIndexes(1, 0, rows.length, headers.length).format = {
      font: { name: BODY_FONT, color: "#000000", size: 10 },
      verticalAlignment: "center",
    };
    for (const [column, format] of Object.entries(dateFormats)) {
      sheet.getRangeByIndexes(1, Number(column), rows.length, 1).format.numberFormat = format;
    }
    for (const [column, format] of Object.entries(formats)) {
      sheet.getRangeByIndexes(1, Number(column), rows.length, 1).format.numberFormat = format;
    }
  }
  for (let column = 0; column < headers.length; column += 1) {
    const width = column === 0 ? 16 : 18;
    sheet.getRangeByIndexes(0, column, Math.max(1, rows.length + 1), 1).format.columnWidth = width;
  }
}


function addIndexPerformanceReportSheet() {
  const sheet = workbook.worksheets.add("指数表现");
  const mainRows = payload.indexPerformance.filter((row) => row.group === "主要指数");
  const styleRows = payload.indexPerformance.filter((row) => row.group === "风格指数");
  if (payload.indexPerformance.length > 0 && (mainRows.length !== 9 || styleRows.length !== 9)) {
    throw new Error(`指数表现底稿分组数量异常：主要指数${mainRows.length}，风格指数${styleRows.length}`);
  }
  const valuesFor = (row) => row
    ? [
        row.name,
        row.close,
        (row.close / row.dailyBaseClose - 1) * 100,
        (row.close / row.weekBaseClose - 1) * 100,
        (row.close / row.monthBaseClose - 1) * 100,
        (row.close / row.yearBaseClose - 1) * 100,
      ]
    : [null, null, null, null, null, null];
  const rows = Array.from({ length: 9 }, (_, position) => [
    ...valuesFor(mainRows[position]),
    ...valuesFor(styleRows[position]),
  ]);

  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(2);
  sheet.getRange("A1:L1").merge();
  sheet.getRange("A1").values = [["图表 1  主要市场指数"]];
  sheet.getRange("A2:L2").values = [[
    "主要指数", "收盘价", "日涨跌幅", "近一周", "近一月", "年初至今",
    "主要指数", "收盘价", "日涨跌幅", "近一周", "近一月", "年初至今",
  ]];
  sheet.getRange("A3:L11").values = rows;
  sheet.getRange("A12:L12").merge();
  sheet.getRange("A12").values = [["资料来源：Wind，华创证券"]];

  sheet.getRange("A1:L1").format = {
    font: { name: TITLE_FONT, bold: true, color: "#000000", size: 12 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    borders: { preset: "bottom", style: "medium", color: "#000000" },
  };
  sheet.getRange("A1:L1").format.rowHeight = 25;
  sheet.getRange("A2:L2").format = {
    fill: LIGHT_BLUE,
    font: { name: BODY_FONT, bold: true, color: "#000000", size: 10 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: GRID },
  };
  sheet.getRange("A2:L2").format.rowHeight = 22;
  sheet.getRange("A3:L11").format = {
    font: { name: BODY_FONT, color: "#000000", size: 10 },
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: GRID },
  };
  sheet.getRange("A3:A11").format.horizontalAlignment = "left";
  sheet.getRange("G3:G11").format.horizontalAlignment = "left";
  sheet.getRange("B3:F11").format.horizontalAlignment = "right";
  sheet.getRange("H3:L11").format.horizontalAlignment = "right";
  sheet.getRange("A3:L11").format.rowHeight = 20;
  sheet.getRange("B3:F11").format.numberFormat = "0.00";
  sheet.getRange("H3:L11").format.numberFormat = "0.00";
  for (const range of ["C3:F11", "I3:L11"]) {
    sheet.getRange(range).conditionalFormats.add("cellIs", {
      operator: "lessThan",
      formula: 0,
      format: { font: { color: NEGATIVE_RED } },
    });
  }
  sheet.getRange("A12:L12").format = {
    font: { name: BODY_FONT, color: "#000000", size: 9 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    borders: { preset: "top", style: "thin", color: "#000000" },
  };
  sheet.getRange("A12:L12").format.rowHeight = 20;
  for (const column of ["A", "G"]) {
    sheet.getRange(`${column}1:${column}12`).format.columnWidth = 16;
  }
  for (const column of ["B", "C", "D", "E", "F", "H", "I", "J", "K", "L"]) {
    sheet.getRange(`${column}1:${column}12`).format.columnWidth = 12;
  }
}


function addRawChartSheets() {
  const newest = (rows, makeRow) => newestFirst(rows).map(makeRow);
  const add = (name, headers, rows, formats = {}, dateFormats = { 0: "yyyy-mm-dd" }) => addPureDataSheet(name, headers, rows, formats, dateFormats);

  addIndexPerformanceReportSheet();
  add("成交额", ["日期", "上证指数（亿元）", "深证成指（亿元）", "中证转债（亿元）", "沪深合计（亿元）"], newest(payload.index, (row) => [excelDate(row.date), row.shanghai, row.shenzhen, row.convertibleBond, row.total]), { 1: "#,##0.00", 2: "#,##0.00", 3: "#,##0.00", 4: "#,##0.00" });
  add("涨跌幅分布", ["涨跌幅区间", "转债数量"], payload.returnDistribution.map((row) => [row.bucket, row.count]), { 1: "#,##0" }, {});
  add("主力净流入", ["日期", "主力净流入（亿元）"], newest(payload.mainMoneyFlow, (row) => [excelDate(row.date), row.amount]), { 1: "#,##0.00" });
  add("两融余额", ["日期", "融资融券余额（亿元）"], newest(payload.market, (row) => [excelDate(row.date), row.balance]), { 1: "#,##0.00" });
  add("百元拟合溢价率", ["日期", "百元拟合溢价率（%）", "多因子修正百元拟合溢价率（%）"], newest(payload.valuationDaily, (row) => [excelDate(row.date), row.inverseCubic, row.multifactor]), { 1: "0.00", 2: "0.00" });
  add("盘中百元平价拟合溢价率", ["日期时间", "拟合溢价率（%）"], payload.valuationIntraday.map((row) => [excelDateTime(row.datetime), row.premium]), { 1: "0.00" }, { 0: "yyyy-mm-dd hh:mm" });
  add("转债估值修复指数", ["日期", "转债估值修复指数", "中证转债指数后推70日涨跌幅（%）"], newest(payload.valuationRepair, (row) => [excelDate(row.date), row.repair, row.forwardReturn]), { 1: "0.00", 2: "0.00" });
  add("余额加权平价与收盘价中位数", ["日期", "余额加权平价", "收盘价中位数"], newest(payload.priceParity, (row) => [excelDate(row.date), row.weightedParity, row.medianClose]), { 1: "0.00", 2: "0.00" });
  add("股债性分类转股溢价率", ["日期", "偏股型（%）", "平衡型（%）", "偏债型（%）"], newest(payload.equityBondWeighted, (row) => [excelDate(row.date), row.stockPremium, row.balancePremium, row.bondPremium]), { 1: "0.00", 2: "0.00", 3: "0.00" });
  add("分平价区间转股溢价率", ["日期", "130以上（%）", "120-130（含130）（%）", "110-120（含120）（%）", "100-110（含110）（%）", "90-100（含100）（%）", "80-90（含90）（%）", "80以下（含80）（%）"], newest(payload.parityIntervalPremium, (row) => [excelDate(row.date), row.above130, row.p120_130, row.p110_120, row.p100_110, row.p90_100, row.p80_90, row.le80]), { 1: "0.00", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00", 6: "0.00", 7: "0.00" });
  add("平价分类拟合溢价率", ["日期", "70-90（%）", "90-110（%）", "110-130（%）", "130-150（%）"], newest(payload.parityGroupValuation, (row) => [excelDate(row.date), row.group70_90, row.group90_110, row.group110_130, row.group130_150]), { 1: "0.00", 2: "0.00", 3: "0.00", 4: "0.00" });
  add("分剩余期限拟合溢价率", ["日期", "0-1（%）", "1-2（%）", "2-3（%）", "3-4（%）", "4-5（%）", "5-6（%）"], newest(payload.maturityGroupValuation, (row) => [excelDate(row.date), row.group0_1, row.group1_2, row.group2_3, row.group3_4, row.group4_5, row.group5_6]), { 1: "0.00", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00", 6: "0.00" });
  add("次新券平均转股溢价率", ["日期", "平均转股溢价率（%）"], newest(payload.subnewBond, (row) => [excelDate(row.date), row.premiumMean]), { 1: "0.00" });
  add("股债型拟合溢价率", ["日期", "偏股型（%）", "平衡型（%）", "偏债型（%）"], newest(payload.equityBondGroupValuation, (row) => [excelDate(row.date), row.stock, row.balance, row.bond]), { 1: "0.00", 2: "0.00", 3: "0.00" });
  add("分评级拟合溢价率", ["日期", "AAA/AA+（%）", "AA/AA-（%）", "A+/A（%）"], newest(payload.ratingGroupValuation, (row) => [excelDate(row.date), row.top, row.middle, row.lower]), { 1: "0.00", 2: "0.00", 3: "0.00" });
  add("分余额拟合溢价率", ["日期", "0-3（%）", "3-10（%）", "10-20（%）", "20-50（%）", "50+（%）"], newest(payload.balanceGroupValuation, (row) => [excelDate(row.date), row.group0_3, row.group3_10, row.group10_20, row.group20_50, row.group50_plus]), { 1: "0.00", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00" });
  add("分正股市值拟合溢价率", ["日期", "0-50（%）", "50-300（%）", "300+（%）"], newest(payload.marketCapGroupValuation, (row) => [excelDate(row.date), row.group0_50, row.group50_300, row.group300_plus]), { 1: "0.00", 2: "0.00", 3: "0.00" });
  add("分板块拟合溢价率", ["日期", "科技（%）", "金融（%）", "制造（%）", "消费（%）", "周期（%）"], newest(payload.sectorGroupValuation, (row) => [excelDate(row.date), row.technology, row.finance, row.manufacturing, row.consumption, row.cyclical]), { 1: "0.00", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00" });
  add("收盘价分布", ["日期", "80以下（%）", "80-90（%）", "90-100（%）", "100-110（%）", "110-120（%）", "120-130（%）", "130-150（%）", "150以上（%）"], newest(payload.closePriceDistribution, (row) => [excelDate(row.date), row.le80, row.p80_90, row.p90_100, row.p100_110, row.p110_120, row.p120_130, row.p130_150, row.gt150]), { 1: "0.00", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00", 6: "0.00", 7: "0.00", 8: "0.00" });
  add("博时ETF份额", ["日期", "ETF份额（亿份）"], newest(payload.etfShare, (row) => [excelDate(row.date), row.boshiShare]), { 1: "0.00" });
  add("海富通ETF份额", ["日期", "ETF份额（亿份）"], newest(payload.etfShare, (row) => [excelDate(row.date), row.haifutongShare]), { 1: "0.00" });
  const sectorMetricSpecs = [
    ["各行业平均收盘价", "close", "收盘价"],
    ["各行业平均平价", "parity", "平价"],
    ["各行业平均转股溢价率", "conversionPremium", "转股溢价率（%）"],
    ["各行业平均纯债溢价率", "bondPremium", "纯债溢价率（%）"],
  ];
  for (const [sheetName, key, label] of sectorMetricSpecs) {
    add(sheetName, ["日期", `科技${label}`, `金融${label}`, `制造${label}`, `消费${label}`, `周期${label}`], newest(payload.sectorMeanMetrics, (row) => [excelDate(row.date), row[`${key}_technology`], row[`${key}_finance`], row[`${key}_manufacturing`], row[`${key}_consumption`], row[`${key}_cyclical`]]), { 1: "0.00", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00" });
  }
  add("行业涨跌与估值", ["行业代码", "行业名称", "转债日涨跌幅", "转债周涨跌幅", "转债月涨跌幅", "转债年初至今涨跌幅", "正股日涨跌幅", "正股周涨跌幅", "正股月涨跌幅", "正股年初至今涨跌幅", "PE（TTM）", "PE三年分位数", "PE十年分位数", "PB（LF）", "PB三年分位数", "PB十年分位数"], payload.industryPerformance.map((row) => [row.code, row.name, row.bondDaily, row.bondWeekly, row.bondMonthly, row.bondYtd, row.stockDaily, row.stockWeekly, row.stockMonthly, row.stockYtd, row.pe, row.pe3y, row.pe10y, row.pb, row.pb3y, row.pb10y]), { 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00", 6: "0.00", 7: "0.00", 8: "0.00", 9: "0.00", 10: "0.00", 11: "0.00", 12: "0.00", 13: "0.00", 14: "0.00", 15: "0.00" }, {});
  add("股债性分类均价", ["日期", "偏股型", "平衡型", "偏债型"], newest(payload.equityBondWeighted, (row) => [excelDate(row.date), row.stockPrice, row.balancePrice, row.bondPrice]), { 1: "0.00", 2: "0.00", 3: "0.00" });
}


function addIndustryHistorySheets() {
  const specs = [
    ["行业收盘价历史", "close"],
    ["行业平价历史", "parity"],
    ["行业转股溢价率历史", "conversionPremium"],
    ["行业纯债溢价率历史", "bondPremium"],
  ];
  const industryOrder = payload.industryHistoryOrder || [];
  const dates = Array.from(
    new Set(payload.industryHistoryMetrics.map((row) => row.date))
  ).sort((left, right) => left.localeCompare(right));
  if (dates.length > 16383) {
    throw new Error(`行业历史序列日期列数超过Excel上限：${dates.length}`);
  }
  for (const [sheetName, metricKey] of specs) {
    const sheet = workbook.worksheets.add(sheetName);
    const lookup = new Map(
      payload.industryHistoryMetrics.map((row) => [
        `${row.industry}||${row.date}`,
        row[metricKey],
      ])
    );
    const rows = industryOrder.map((industry) => [
      industry,
      ...dates.map((date) => lookup.get(`${industry}||${date}`) ?? null),
    ]);
    const columnCount = dates.length + 1;
    sheet.showGridLines = false;
    sheet.freezePanes.freezeColumns(1);
    sheet.freezePanes.freezeRows(1);
    sheet.getRangeByIndexes(0, 0, 1, columnCount).values = [["行业", ...dates]];
    sheet.getRangeByIndexes(0, 0, 1, columnCount).format = {
      font: { name: BODY_FONT, bold: true, color: "#000000", size: 10 },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      borders: { preset: "bottom", style: "thin", color: "#000000" },
    };
    sheet.getRangeByIndexes(0, 0, 1, columnCount).format.rowHeight = 20;
    if (rows.length > 0) {
      sheet.getRangeByIndexes(1, 0, rows.length, columnCount).values = rows;
      sheet.getRangeByIndexes(1, 0, rows.length, columnCount).format = {
        font: { name: BODY_FONT, color: "#000000", size: 10 },
        verticalAlignment: "center",
      };
      if (dates.length > 0) {
        sheet.getRangeByIndexes(1, 1, rows.length, dates.length).format.numberFormat = "0.00";
      }
    }
    sheet.getRangeByIndexes(0, 0, rows.length + 1, 1).format.columnWidth = 16;
    if (dates.length > 0) {
      sheet.getRangeByIndexes(0, 1, rows.length + 1, dates.length).format.columnWidth = 12;
    }
  }
}


addRawChartSheets();
addIndustryHistorySheets();

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

if (previewDirArg) {
  await fs.mkdir(previewDirArg, { recursive: true });
  for (const sheetName of ["指数表现", "成交额", "涨跌幅分布", "主力净流入", "两融余额", "百元拟合溢价率", "盘中百元平价拟合溢价率", "转债估值修复指数", "余额加权平价与收盘价中位数", "股债性分类转股溢价率", "分平价区间转股溢价率", "平价分类拟合溢价率", "分剩余期限拟合溢价率", "次新券平均转股溢价率", "股债型拟合溢价率", "分评级拟合溢价率", "分余额拟合溢价率", "分正股市值拟合溢价率", "分板块拟合溢价率", "收盘价分布", "博时ETF份额", "海富通ETF份额", "各行业平均收盘价", "各行业平均平价", "各行业平均转股溢价率", "各行业平均纯债溢价率", "行业涨跌与估值", "股债性分类均价", "行业收盘价历史", "行业平价历史", "行业转股溢价率历史", "行业纯债溢价率历史"]) {
    const preview = await workbook.render({ sheetName, range: "A1:Z40", scale: 1.25, format: "png" });
    await fs.writeFile(path.join(previewDirArg, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
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
INDUSTRY_TABLE_PIXEL_HEIGHT = 1029
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
SINGLE_LINE_TITLE_BAND_HEIGHT = 0.07
DOUBLE_LINE_TITLE_BAND_HEIGHT = 0.12
SMALL_CHART_EXPORT_SPECS = (
    (1, "成交额", "中证转债与沪深两市成交额.png", SINGLE_LINE_TITLE_BAND_HEIGHT),
    (2, "涨跌幅分布", "可转债当日涨跌幅分布.png", SINGLE_LINE_TITLE_BAND_HEIGHT),
    (3, "主力净流入", "沪深两市主力净流入.png", SINGLE_LINE_TITLE_BAND_HEIGHT),
    (4, "两融余额", "沪深两市融资融券余额.png", SINGLE_LINE_TITLE_BAND_HEIGHT),
    (5, "百元拟合溢价率", "百元拟合溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (6, "盘中百元平价拟合溢价率", "盘中百元平价拟合溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (7, "转债估值修复指数", "转债估值修复指数.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (8, "余额加权平价与收盘价中位数", "余额加权平价与收盘价中位数.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (9, "股债性分类转股溢价率", "股债性分类转股溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (10, "平价分类拟合溢价率", "分平价多因子修正拟合溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (11, "分剩余期限拟合溢价率", "分剩余期限拟合溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (12, "次新券平均转股溢价率", "次新券平均转股溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (13, "股债型拟合溢价率", "股债型拟合溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (14, "分评级拟合溢价率", "分评级拟合溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (15, "分余额拟合溢价率", "分余额拟合溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (16, "分正股市值拟合溢价率", "分正股市值拟合溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (17, "分板块拟合溢价率", "分板块拟合溢价率.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (18, "收盘价分布", "收盘价分布面积图.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (19, "博时ETF份额与净申赎", "博时可转债ETF份额与净申赎.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (20, "海富通ETF份额与净申赎", "海富通可转债ETF份额与净申赎.png", DOUBLE_LINE_TITLE_BAND_HEIGHT),
    (21, "各行业平均收盘价", "各行业平均收盘价.png", SINGLE_LINE_TITLE_BAND_HEIGHT),
    (22, "各行业平均平价", "各行业平均平价.png", SINGLE_LINE_TITLE_BAND_HEIGHT),
    (23, "各行业平均转股溢价率", "各行业平均转股溢价率.png", SINGLE_LINE_TITLE_BAND_HEIGHT),
    (24, "各行业平均纯债溢价率", "各行业平均纯债溢价率.png", SINGLE_LINE_TITLE_BAND_HEIGHT),
)
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
WORD_INDEX_VALUE_COLUMNS = (
    "收盘价",
    "日涨跌幅",
    "近一周涨跌幅",
    "近一月涨跌幅",
    "年初至今涨跌幅",
)
WORD_INDEX_DISPLAY_NAMES = {
    "可转债等权": "转债等权",
    "可转债正股等权": "正股等权",
    "可转债预案": "转债预案",
    "大盘指数(申万)": "大盘指数",
    "中盘指数(申万)": "中盘指数",
    "小盘指数(申万)": "小盘指数",
}


def build_word_index_table_rows(
    index_performance: pd.DataFrame,
) -> list[list[str]]:
    """生成 Word 图表1的九行左右并列显示数据。"""
    required = {"组别", "指数名称", *WORD_INDEX_VALUE_COLUMNS}
    missing = required - set(index_performance.columns)
    if missing:
        raise RuntimeError(f"Word指数表缺少字段：{sorted(missing)}")
    rows_by_name = index_performance.set_index("指数名称", drop=False)
    main_names = [display for _, _, display in MAIN_INDEX_SPECS]
    style_names = [display for _, _, display in STYLE_INDEX_SPECS]
    missing_names = [
        name
        for name in (*main_names, *style_names)
        if name not in rows_by_name.index
    ]
    if missing_names:
        raise RuntimeError(f"Word指数表缺少指数：{missing_names}")

    result: list[list[str]] = []
    for main_name, style_name in zip(main_names, style_names):
        row: list[str] = []
        for name in (main_name, style_name):
            values = rows_by_name.loc[name]
            if isinstance(values, pd.DataFrame):
                raise RuntimeError(f"Word指数表存在重复指数：{name}")
            word_name = WORD_INDEX_DISPLAY_NAMES.get(name, name)
            row.extend(
                [
                    word_name,
                    *(
                        f"{float(values[column]):.2f}"
                        for column in WORD_INDEX_VALUE_COLUMNS
                    ),
                ]
            )
        result.append(row)
    return result


def build_industry_rotation_title(
    industry_performance: pd.DataFrame,
) -> str:
    """按正股行业指数日涨跌幅生成行业轮动标题。"""
    required = {"行业名称", "正股日涨跌幅"}
    missing = required - set(industry_performance.columns)
    if missing:
        raise RuntimeError(f"行业轮动标题缺少字段：{sorted(missing)}")
    ranked = industry_performance.dropna(subset=list(required)).copy()
    ranked["_原始顺序"] = range(len(ranked))
    ranked = ranked.sort_values(
        ["正股日涨跌幅", "_原始顺序"],
        ascending=[False, True],
        kind="stable",
    )
    if len(ranked) < 3:
        raise RuntimeError("行业轮动标题至少需要3个有效行业")
    leaders = "、".join(ranked.head(3)["行业名称"].astype(str))
    return f"行业轮动情况：{leaders}领涨"


def _sha256(path: Path) -> str:
    """返回文件的大写 SHA-256。"""
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _word_tag(local_name: str) -> str:
    return f"{{{WORD_XML_NAMESPACES['w']}}}{local_name}"


def _relationship_tag(local_name: str) -> str:
    return f"{{{WORD_XML_NAMESPACES['r']}}}{local_name}"


def _package_relationship_tag(local_name: str) -> str:
    return f"{{{WORD_XML_NAMESPACES['pr']}}}{local_name}"


def _top_level_word_tables(document_root: ET.Element) -> list[ET.Element]:
    body = document_root.find("w:body", WORD_XML_NAMESPACES)
    if body is None:
        raise RuntimeError("Word模板缺少主文档正文")
    return [child for child in body if child.tag == _word_tag("tbl")]


def _count_word_field_instructions(
    package: zipfile.ZipFile,
    pattern: re.Pattern[str],
) -> int:
    count = 0
    for name in package.namelist():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(package.read(name))
        except ET.ParseError:
            continue
        instruction = "".join(
            node.text or ""
            for node in root.findall(".//w:instrText", WORD_XML_NAMESPACES)
        )
        count += len(pattern.findall(instruction))
    return count


def inspect_daily_word_template(template_path: Path) -> dict[str, object]:
    """读取日报模板的冻结结构，用于写入前的严格校验。"""
    if not template_path.is_file():
        raise FileNotFoundError(f"Word日报模板不存在：{template_path}")
    with zipfile.ZipFile(template_path) as package:
        document_root = ET.fromstring(package.read("word/document.xml"))
        tables = _top_level_word_tables(document_root)
        relationship_ids: list[str] = []
        for table in tables[2:15]:
            for blip in table.findall(".//a:blip", WORD_XML_NAMESPACES):
                relationship_id = blip.get(_relationship_tag("embed"))
                if relationship_id:
                    relationship_ids.append(relationship_id)
        sequence_count = sum(
            "SEQ 图表" in "".join(
                node.text or ""
                for node in paragraph.findall(
                    ".//w:instrText", WORD_XML_NAMESPACES
                )
            )
            for paragraph in document_root.findall(
                ".//w:p", WORD_XML_NAMESPACES
            )
        )
        page_count = _count_word_field_instructions(
            package, re.compile(r"\bPAGE\b")
        )
        return {
            "sha256": _sha256(template_path),
            "topLevelTableCount": len(tables),
            "chartImageRelationshipIds": relationship_ids,
            "sequenceFieldCount": sequence_count,
            "pageFieldCount": page_count,
            "memberNames": tuple(package.namelist()),
        }


def validate_daily_word_template(template_path: Path) -> dict[str, object]:
    """验证模板仍与已审计版本一致，防止改版后错位写入。"""
    contract = inspect_daily_word_template(template_path)
    expected = {
        "sha256": DAILY_WORD_TEMPLATE_SHA256,
        "topLevelTableCount": 18,
        "chartImageRelationshipIds": [
            f"rId{number}" for number in range(14, 39)
        ],
        "sequenceFieldCount": 26,
        "pageFieldCount": 1,
    }
    mismatches = {
        key: {"expected": value, "actual": contract.get(key)}
        for key, value in expected.items()
        if contract.get(key) != value
    }
    if mismatches:
        raise RuntimeError(f"Word日报模板结构已变化，停止写入：{mismatches}")
    return contract


def set_plain_text_control_value(
    root: ET.Element,
    tag: str,
    value: str,
) -> int:
    """更新指定标签的纯文本内容控件，保留控件及其格式。"""
    count = 0
    tag_attribute = _word_tag("val")
    for control in root.findall(".//w:sdt", WORD_XML_NAMESPACES):
        tag_node = control.find("w:sdtPr/w:tag", WORD_XML_NAMESPACES)
        if tag_node is None or tag_node.get(tag_attribute) != tag:
            continue
        text_nodes = control.findall("w:sdtContent//w:t", WORD_XML_NAMESPACES)
        if not text_nodes:
            raise RuntimeError(f"Word内容控件没有文本节点：{tag}")
        text_nodes[0].text = value
        for node in text_nodes[1:]:
            node.text = ""
        count += 1
    return count


def _daily_commentary_summary_lines(commentary: str) -> list[str]:
    """提取首页点评正文，排除写给报告复制使用的图表标题清单。"""
    body = str(commentary).replace("\r\n", "\n").replace("\r", "\n")
    body = body.split("\n\n图表标题：", 1)[0]
    lines = [line.strip() for line in body.splitlines() if line.strip()]
    if len(lines) != 12:
        raise RuntimeError(
            "Word首页点评必须包含12个非空段落："
            f"实际{len(lines)}个"
        )
    return lines


def _replace_word_paragraph_text_preserving_runs(
    paragraph: ET.Element,
    value: str,
) -> None:
    """复用段落现有文字节点，保留字体、缩进和项目符号格式。"""
    text_nodes = paragraph.findall(".//w:t", WORD_XML_NAMESPACES)
    if not text_nodes:
        raise RuntimeError("Word首页点评目标段落没有文本节点")
    for node in text_nodes:
        node.text = ""

    heading, separator, remainder = str(value).partition("：")
    if separator and remainder and len(text_nodes) >= 2:
        text_nodes[0].text = f"{heading}{separator}"
        text_nodes[1].text = remainder
    else:
        text_nodes[0].text = str(value)


def replace_daily_word_summary(
    root: ET.Element,
    commentary: str,
) -> None:
    """把日报点评逐段写入首页摘要控件，并保留固定风险提示。"""
    tag_attribute = _word_tag("val")
    controls = []
    for control in root.findall(".//w:sdt", WORD_XML_NAMESPACES):
        tag_node = control.find("w:sdtPr/w:tag", WORD_XML_NAMESPACES)
        if (
            tag_node is not None
            and tag_node.get(tag_attribute) == "ReportSummary"
        ):
            controls.append(control)
    if len(controls) != 1:
        raise RuntimeError(
            "Word模板中的ReportSummary内容控件数量异常："
            f"{len(controls)}"
        )

    paragraphs = [
        paragraph
        for paragraph in controls[0].findall(
            "w:sdtContent//w:p", WORD_XML_NAMESPACES
        )
        if "".join(paragraph.itertext()).strip()
    ]
    risk_index = next(
        (
            index
            for index, paragraph in enumerate(paragraphs)
            if "".join(paragraph.itertext()).strip() == "风险提示："
        ),
        None,
    )
    if risk_index is None:
        raise RuntimeError("Word首页摘要缺少固定风险提示")
    dynamic_paragraphs = paragraphs[:risk_index]
    lines = _daily_commentary_summary_lines(commentary)
    if len(dynamic_paragraphs) != len(lines):
        raise RuntimeError(
            "Word首页点评段落槽位数量异常："
            f"{len(dynamic_paragraphs)} != {len(lines)}"
        )
    for paragraph, line in zip(dynamic_paragraphs, lines):
        _replace_word_paragraph_text_preserving_runs(paragraph, line)


def replace_text_after_seq_field(
    paragraph: ET.Element,
    value: str,
) -> None:
    """保留图表SEQ域与书签，仅替换域结束后的标题文本。"""
    children = list(paragraph)
    field_end_index: Optional[int] = None
    for index, child in enumerate(children):
        field_node = child.find("w:fldChar", WORD_XML_NAMESPACES)
        if (
            child.tag == _word_tag("r")
            and field_node is not None
            and field_node.get(_word_tag("fldCharType")) == "end"
        ):
            field_end_index = index
            break
    if field_end_index is None:
        raise RuntimeError("Word图表标题缺少SEQ域结束节点")

    post_field_runs = [
        child
        for child in children[field_end_index + 1 :]
        if child.tag == _word_tag("r")
    ]
    if not post_field_runs:
        raise RuntimeError("Word图表标题缺少可复用的标题格式节点")
    leading_space_run: Optional[ET.Element] = None
    for run in post_field_runs:
        run_text = "".join(run.itertext())
        if leading_space_run is None and run_text and not run_text.strip():
            leading_space_run = run
            continue
        if run_text.strip():
            style_run = run
            break
    else:
        style_run = post_field_runs[-1]

    for run in post_field_runs:
        if run is not leading_space_run:
            paragraph.remove(run)

    remaining_children = list(paragraph)
    insertion_index = len(remaining_children)
    if leading_space_run is not None:
        insertion_index = remaining_children.index(leading_space_run) + 1
    else:
        for index, child in enumerate(remaining_children):
            if index > field_end_index and child.tag != _word_tag("r"):
                insertion_index = index
                break

    new_run = ET.Element(_word_tag("r"), dict(style_run.attrib))
    run_properties = style_run.find("w:rPr", WORD_XML_NAMESPACES)
    if run_properties is not None:
        new_run.append(deepcopy(run_properties))
    lines = str(value).split("\n")
    for line_index, line in enumerate(lines):
        if line_index:
            new_run.append(ET.Element(_word_tag("br")))
        text_node = ET.SubElement(new_run, _word_tag("t"))
        text_node.text = line
    paragraph.insert(insertion_index, new_run)


def replace_sequence_title(cell: ET.Element, value: str) -> None:
    """把图表标题逐行写入模板既有段落，保持标题栏原始高度。"""
    paragraphs = cell.findall("w:p", WORD_XML_NAMESPACES)
    if not paragraphs:
        raise RuntimeError("Word图表标题单元格缺少段落")
    sequence_paragraph = _sequence_title_paragraph(cell)
    sequence_index = paragraphs.index(sequence_paragraph)
    title_paragraphs = paragraphs[sequence_index:]
    lines = str(value).split("\n")
    if len(lines) > len(title_paragraphs):
        raise RuntimeError(
            "Word图表标题行数超过模板槽位："
            f"{len(lines)} > {len(title_paragraphs)}"
        )
    replace_text_after_seq_field(sequence_paragraph, lines[0])
    for paragraph, line in zip(title_paragraphs[1:], lines[1:]):
        text_nodes = paragraph.findall(".//w:t", WORD_XML_NAMESPACES)
        if not text_nodes:
            raise RuntimeError("Word图表续行标题缺少文本节点")
        text_nodes[0].text = line
        for node in text_nodes[1:]:
            node.text = ""
    for paragraph in title_paragraphs[len(lines) :]:
        for node in paragraph.findall(".//w:t", WORD_XML_NAMESPACES):
            node.text = ""


def _register_word_xml_namespaces(xml_bytes: bytes) -> dict[str, str]:
    """注册原部件命名空间前缀，避免序列化时生成ns0等新前缀。"""
    namespaces: dict[str, str] = {}
    for _, (prefix, uri) in ET.iterparse(
        io.BytesIO(xml_bytes), events=("start-ns",)
    ):
        namespaces.setdefault(prefix, uri)
        if prefix == "xml":
            continue
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            continue
    return namespaces


def _serialize_word_xml(root: ET.Element, source_bytes: bytes) -> bytes:
    source_namespaces = _register_word_xml_namespaces(source_bytes)
    rendered = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    root_end = rendered.find(b">", rendered.find(b"<", 5))
    if root_end < 0:
        raise RuntimeError("Word XML根节点序列化异常")
    root_start = rendered[:root_end].decode("utf-8")
    declared_prefixes = {
        prefix or ""
        for prefix in re.findall(
            r"xmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?=\"[^\"]+\"",
            root_start,
        )
    }
    missing_declarations = []
    for prefix, uri in source_namespaces.items():
        if prefix in declared_prefixes or prefix == "xml":
            continue
        name = "xmlns" if not prefix else f"xmlns:{prefix}"
        missing_declarations.append(f' {name}="{uri}"')
    if not missing_declarations:
        return rendered
    addition = "".join(missing_declarations).encode("utf-8")
    return rendered[:root_end] + addition + rendered[root_end:]


def _direct_table_rows(table: ET.Element) -> list[ET.Element]:
    return [child for child in table if child.tag == _word_tag("tr")]


def _direct_row_cells(row: ET.Element) -> list[ET.Element]:
    return [child for child in row if child.tag == _word_tag("tc")]


def _set_word_cell_text(cell: ET.Element, value: str) -> None:
    text_nodes = cell.findall(".//w:t", WORD_XML_NAMESPACES)
    if not text_nodes:
        raise RuntimeError("Word目标单元格没有文本节点")
    text_nodes[0].text = value
    for node in text_nodes[1:]:
        node.text = ""


def _sequence_title_paragraph(cell: ET.Element) -> ET.Element:
    for paragraph in cell.findall(".//w:p", WORD_XML_NAMESPACES):
        instruction = "".join(
            node.text or ""
            for node in paragraph.findall(
                ".//w:instrText", WORD_XML_NAMESPACES
            )
        )
        if "SEQ 图表" in instruction:
            return paragraph
    raise RuntimeError("Word目标标题单元格缺少SEQ图表域")


def _patch_daily_word_document(
    document_bytes: bytes,
    run_date: date,
    index_rows: list[list[str]],
    chart_titles: list[str],
    industry_title: str,
    commentary: str,
) -> tuple[bytes, list[str]]:
    root = ET.fromstring(document_bytes)
    tables = _top_level_word_tables(root)
    if len(tables) != 18:
        raise RuntimeError(f"Word主文档表格数量异常：{len(tables)}")
    if len(index_rows) != 9 or any(len(row) != 12 for row in index_rows):
        raise RuntimeError("Word图表1必须提供9行、每行12列数据")
    if len(chart_titles) != 24:
        raise RuntimeError(f"Word图表标题数量异常：{len(chart_titles)}")

    if set_plain_text_control_value(
        root, "ReportTitle", f"转债市场日度跟踪{run_date:%Y%m%d}"
    ) != 1:
        raise RuntimeError("Word模板中的ReportTitle内容控件数量异常")
    replace_daily_word_summary(root, commentary)

    index_table_rows = _direct_table_rows(tables[1])
    if len(index_table_rows) != 12:
        raise RuntimeError(f"Word图表1行数异常：{len(index_table_rows)}")
    for row_element, values in zip(index_table_rows[2:11], index_rows):
        cells = _direct_row_cells(row_element)
        if len(cells) != 12:
            raise RuntimeError(f"Word图表1列数异常：{len(cells)}")
        for cell, value in zip(cells, values):
            _set_word_cell_text(cell, value)

    title_index = 0
    for table in tables[2:14]:
        rows = _direct_table_rows(table)
        if not rows:
            raise RuntimeError("Word双图表格缺少标题行")
        cells = _direct_row_cells(rows[0])
        if len(cells) < 3:
            raise RuntimeError(f"Word双图标题行列数异常：{len(cells)}")
        for cell_index in (0, 2):
            replace_sequence_title(cells[cell_index], chart_titles[title_index])
            title_index += 1
    if title_index != 24:
        raise RuntimeError(f"Word双图标题实际写入数量异常：{title_index}")

    industry_rows = _direct_table_rows(tables[14])
    industry_cells = _direct_row_cells(industry_rows[0])
    replace_sequence_title(industry_cells[0], industry_title)

    relationship_ids: list[str] = []
    for table in tables[2:15]:
        for blip in table.findall(".//a:blip", WORD_XML_NAMESPACES):
            relationship_id = blip.get(_relationship_tag("embed"))
            if relationship_id:
                relationship_ids.append(relationship_id)
    if len(relationship_ids) != 25:
        raise RuntimeError(f"Word日报图片槽位数量异常：{len(relationship_ids)}")
    return _serialize_word_xml(root, document_bytes), relationship_ids


def _patch_daily_word_header(header_bytes: bytes, run_date: date) -> bytes:
    root = ET.fromstring(header_bytes)
    updated = set_plain_text_control_value(
        root, "ReportDate", f"{run_date:%Y年%m月%d日}"
    )
    if updated != 2:
        raise RuntimeError(f"Word模板中的ReportDate内容控件数量异常：{updated}")
    return _serialize_word_xml(root, header_bytes)


def _resolve_word_media_members(
    relationship_bytes: bytes,
    relationship_ids: list[str],
) -> list[str]:
    root = ET.fromstring(relationship_bytes)
    by_id = {
        relationship.get("Id"): relationship.get("Target")
        for relationship in root.findall(
            _package_relationship_tag("Relationship")
        )
    }
    media_members: list[str] = []
    for relationship_id in relationship_ids:
        target = by_id.get(relationship_id)
        if not target:
            raise RuntimeError(f"Word图片关系不存在：{relationship_id}")
        member = posixpath.normpath(posixpath.join("word", target))
        if not member.startswith("word/media/"):
            raise RuntimeError(
                f"Word图片关系未指向媒体部件：{relationship_id} -> {target}"
            )
        media_members.append(member)
    if len(set(media_members)) != len(media_members):
        raise RuntimeError("Word日报图片槽位存在重复媒体目标")
    return media_members


def _write_patched_word_package(
    template_path: Path,
    output_path: Path,
    run_date: date,
    index_rows: list[list[str]],
    chart_titles: list[str],
    industry_title: str,
    image_paths: list[Path],
    commentary: str,
) -> None:
    with zipfile.ZipFile(template_path, "r") as source:
        document_bytes = source.read("word/document.xml")
        header_bytes = source.read("word/header3.xml")
        relationship_bytes = source.read("word/_rels/document.xml.rels")
        patched_document, relationship_ids = _patch_daily_word_document(
            document_bytes,
            run_date,
            index_rows,
            chart_titles,
            industry_title,
            commentary,
        )
        patched_header = _patch_daily_word_header(header_bytes, run_date)
        media_members = _resolve_word_media_members(
            relationship_bytes, relationship_ids
        )
        if len(image_paths) != len(media_members):
            raise RuntimeError(
                "Word图片输入数量与模板槽位不一致："
                f"{len(image_paths)} != {len(media_members)}"
            )
        replacements = {
            "word/document.xml": patched_document,
            "word/header3.xml": patched_header,
            **{
                member: image_path.read_bytes()
                for member, image_path in zip(media_members, image_paths)
            },
        }
        with zipfile.ZipFile(output_path, "w") as destination:
            for information in source.infolist():
                destination.writestr(
                    information,
                    replacements.get(information.filename, source.read(information)),
                )


def validate_generated_daily_word_report(
    output_path: Path,
    run_date: date,
    template_member_names: tuple[str, ...],
    index_rows: list[list[str]],
    chart_titles: list[str],
    industry_title: str,
    commentary: str,
) -> None:
    """验证生成文件的结构与全部动态槽位。"""
    with zipfile.ZipFile(output_path) as package:
        broken_member = package.testzip()
        if broken_member is not None:
            raise RuntimeError(f"Word输出压缩包损坏：{broken_member}")
        if tuple(package.namelist()) != template_member_names:
            raise RuntimeError("Word输出部件清单与模板不一致")
        document_root = ET.fromstring(package.read("word/document.xml"))
        header_root = ET.fromstring(package.read("word/header3.xml"))
        tables = _top_level_word_tables(document_root)
        if len(tables) != 18:
            raise RuntimeError(f"Word输出表格数量异常：{len(tables)}")
        relationship_ids = [
            blip.get(_relationship_tag("embed"))
            for table in tables[2:15]
            for blip in table.findall(".//a:blip", WORD_XML_NAMESPACES)
        ]
        if relationship_ids != [f"rId{number}" for number in range(14, 39)]:
            raise RuntimeError("Word输出图片关系顺序异常")
        sequence_count = sum(
            "SEQ 图表" in "".join(
                node.text or ""
                for node in paragraph.findall(
                    ".//w:instrText", WORD_XML_NAMESPACES
                )
            )
            for paragraph in document_root.findall(
                ".//w:p", WORD_XML_NAMESPACES
            )
        )
        if sequence_count != 26:
            raise RuntimeError(f"Word输出图表序号域数量异常：{sequence_count}")

        document_text = "".join(document_root.itertext())
        header_text = "".join(header_root.itertext())
        required_text = [
            f"转债市场日度跟踪{run_date:%Y%m%d}",
            industry_title,
            *(title.replace("\n", "") for title in chart_titles),
        ]
        missing_text = [text for text in required_text if text not in document_text]
        if missing_text:
            raise RuntimeError(f"Word输出缺少动态标题：{missing_text}")
        expected_date = f"{run_date:%Y年%m月%d日}"
        if header_text.count(expected_date) != 2:
            raise RuntimeError(f"Word输出日期内容控件异常：{expected_date}")

        actual_index_rows = [
            ["".join(cell.itertext()) for cell in _direct_row_cells(row)]
            for row in _direct_table_rows(tables[1])[2:11]
        ]
        if actual_index_rows != index_rows:
            raise RuntimeError("Word输出图表1数据与输入不一致")

        summary_controls = []
        for control in document_root.findall(
            ".//w:sdt", WORD_XML_NAMESPACES
        ):
            tag_node = control.find(
                "w:sdtPr/w:tag", WORD_XML_NAMESPACES
            )
            if (
                tag_node is not None
                and tag_node.get(_word_tag("val")) == "ReportSummary"
            ):
                summary_controls.append(control)
        if len(summary_controls) != 1:
            raise RuntimeError("Word输出ReportSummary内容控件数量异常")
        summary_paragraphs = [
            "".join(paragraph.itertext()).strip()
            for paragraph in summary_controls[0].findall(
                "w:sdtContent//w:p", WORD_XML_NAMESPACES
            )
            if "".join(paragraph.itertext()).strip()
        ]
        expected_summary = _daily_commentary_summary_lines(commentary)
        if summary_paragraphs[: len(expected_summary)] != expected_summary:
            raise RuntimeError("Word输出首页点评与输入不一致")
        if "图表标题：" in "".join(summary_paragraphs):
            raise RuntimeError("Word输出首页误写入图表标题清单")


def build_daily_word_report(
    run_date: date,
    output_dir: Path,
    index_performance: pd.DataFrame,
    chart_titles: list[str],
    industry_performance: pd.DataFrame,
    industry_chart_path: Path,
    commentary: str,
    template_path: Path = DAILY_WORD_TEMPLATE_PATH,
) -> Path:
    """按冻结模板生成当日 Word 报告，不调用 Word COM。"""
    if len(chart_titles) != 24:
        raise RuntimeError(f"Word图表标题数量异常：{len(chart_titles)}")
    contract = validate_daily_word_template(template_path)
    chart_paths = [
        output_dir / f"{sequence:02d}_{label}.png"
        for sequence, label, _, _ in SMALL_CHART_EXPORT_SPECS
    ]
    image_paths = [*chart_paths, industry_chart_path]
    missing_images = [str(path) for path in image_paths if not path.is_file()]
    if missing_images:
        raise FileNotFoundError(f"Word报告缺少图片：{missing_images}")

    index_rows = build_word_index_table_rows(index_performance)
    industry_title = build_industry_rotation_title(industry_performance)
    output_path = (
        output_dir / f"【华创固收】转债市场日度跟踪{run_date:%Y%m%d}.docx"
    )
    temporary_path = output_path.with_name(
        f".{output_path.stem}.{os.getpid()}.tmp.docx"
    )
    try:
        _write_patched_word_package(
            template_path,
            temporary_path,
            run_date,
            index_rows,
            chart_titles,
            industry_title,
            image_paths,
            commentary,
        )
        validate_generated_daily_word_report(
            temporary_path,
            run_date,
            contract["memberNames"],
            index_rows,
            chart_titles,
            industry_title,
            commentary,
        )
        os.replace(temporary_path, output_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
    return output_path


class ConsoleProgress:
    """在终端同一行显示日报生成进度，不引入额外依赖。"""

    def __init__(self, label: str = "转债日报", width: int = 28) -> None:
        if not sys.stderr.isatty() and hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        self.label = label
        self.width = width
        self.percent = 0
        self._line_width = 0
        self.update(0, "初始化")

    def update(self, percent: int, message: str) -> None:
        """更新进度和阶段说明；percent 自动限制在 0~100。"""
        self.percent = max(0, min(100, int(percent)))
        completed = round(self.width * self.percent / 100)
        bar = "#" * completed + "-" * (self.width - completed)
        line = f"\r{self.label} [{bar}] {self.percent:3d}% {message}"
        line_width = sum(
            2 if east_asian_width(character) in {"W", "F", "A"} else 1
            for character in line
        )
        padding = " " * max(0, self._line_width - line_width)
        sys.stderr.write(line + padding)
        sys.stderr.flush()
        self._line_width = line_width

    def finish(self, message: str = "完成") -> None:
        self.update(100, message)
        sys.stderr.write("\n")
        sys.stderr.flush()

    def fail(self) -> None:
        self.update(self.percent, "失败")
        sys.stderr.write("\n")
        sys.stderr.flush()


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


def ths_login(
    ths_id: Optional[str] = None, ths_password: Optional[str] = None
) -> int:
    """登录 iFinD 并返回状态码。"""
    if not ths_id or not ths_password:
        file_id, file_password = load_ifind_credentials()
        ths_id = ths_id or file_id
        ths_password = ths_password or file_password
    return THS_iFinDLogin(ths_id, ths_password)


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


def fetch_main_money_flow(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """通过 Wind Excel WSET 公式获取近三个月沪深两市主力净流入。"""
    start_date = (pd.Timestamp(run_date) - pd.DateOffset(months=3)).date()
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    latest_formula_value: object = None
    latest_region_address = ""
    evaluated_formula = MAIN_MONEY_FLOW_WSET_FORMULA
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        try:
            excel.COMAddIns("WDF.Addin").Connect = True
        except Exception:
            pass

        workbook = excel.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        sheet.Name = "主力净流入"
        for cell, value in (
            ("K5", f"{start_date:%Y%m%d}"),
            ("K6", f"{run_date:%Y%m%d}"),
        ):
            sheet.Range(cell).NumberFormat = "@"
            sheet.Range(cell).Value = value
        sheet.Range("K7").Value = "日"
        sheet.Range("K8").Value = "沪深两市"
        sheet.Range("K10").Value = "A股"
        try:
            sheet.Range("A1").Formula2 = MAIN_MONEY_FLOW_WSET_FORMULA
        except Exception:
            sheet.Range("A1").Formula = MAIN_MONEY_FLOW_WSET_FORMULA.replace("=@", "=")
        excel.CalculateFullRebuild()

        stable_signature: tuple[object, ...] | None = None
        stable_rounds = 0
        output = pd.DataFrame()
        started = time.monotonic()
        while time.monotonic() - started < 45:
            time.sleep(1)
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            latest_formula_value = sheet.Range("A1").Value
            evaluated_formula = str(sheet.Range("A1").Formula2)
            region = sheet.Range("A1").CurrentRegion
            latest_region_address = str(region.Address)
            values = region.Value
            if not isinstance(values, tuple):
                continue
            rows: list[dict[str, object]] = []
            for value_row in values:
                if not isinstance(value_row, tuple) or len(value_row) < 2:
                    continue
                raw_date, raw_amount = value_row[:2]
                try:
                    if all(hasattr(raw_date, part) for part in ("year", "month", "day")):
                        timestamp = pd.Timestamp(
                            date(
                                int(raw_date.year),
                                int(raw_date.month),
                                int(raw_date.day),
                            )
                        )
                    else:
                        timestamp = pd.Timestamp(raw_date).normalize()
                    amount = float(raw_amount)
                except (TypeError, ValueError):
                    continue
                if not math.isfinite(amount):
                    continue
                rows.append(
                    {
                        "交易日期": timestamp,
                        "主力净流入_万元": amount,
                        "主力净流入_亿元": amount / 10000.0,
                    }
                )
            if len(rows) < 2:
                continue
            candidate = (
                pd.DataFrame(rows)
                .drop_duplicates("交易日期", keep="last")
                .sort_values("交易日期")
                .reset_index(drop=True)
            )
            if candidate["交易日期"].iloc[-1].date() != run_date:
                continue
            signature = (
                len(candidate),
                candidate["交易日期"].iloc[0],
                candidate["交易日期"].iloc[-1],
                float(candidate["主力净流入_万元"].iloc[-1]),
            )
            if signature == stable_signature:
                stable_rounds += 1
            else:
                stable_signature = signature
                stable_rounds = 0
            output = candidate
            if stable_rounds >= 2:
                break

        if output.empty or stable_rounds < 2:
            raise RuntimeError(
                "Wind WSET 主力净流入数据刷新超时或未更新至运行日："
                f"公式值={latest_formula_value!r}，返回区域={latest_region_address}"
            )
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()

    latest_value = float(output["主力净流入_亿元"].iloc[-1])
    source: dict[str, object] = {
        "source": "Wind Excel WSET/marketmoneyflows",
        "formula": MAIN_MONEY_FLOW_WSET_FORMULA,
        "evaluatedFormula": evaluated_formula,
        "startDate": f"{start_date:%Y-%m-%d}",
        "latestDate": f"{run_date:%Y-%m-%d}",
        "frequency": "日",
        "sector": "沪深两市",
        "securityType": "A股",
        "field": "date,mainInflowMoney",
        "rawUnit": "万元",
        "unitRule": "mainInflowMoney原始值（万元）除以10000，转换为亿元",
        "points": len(output),
        "latestValue": latest_value,
        "latestDirection": "净流入" if latest_value >= 0 else "净流出",
    }
    return output, source


def _sse_trade_calendar(start_date: date, end_date: date) -> pd.DatetimeIndex:
    """从本地指数 Parquet 取得上交所交易日，用于对齐不返回日期列的 Wind WSD。"""
    if not INDEX_PARQUET.is_file():
        raise FileNotFoundError(f"未找到指数 Parquet：{INDEX_PARQUET}")
    calendar_data = pd.read_parquet(
        INDEX_PARQUET, columns=["指数名称", "交易日期", "指数值"]
    )
    dates = pd.to_datetime(
        calendar_data.loc[
            calendar_data["指数名称"].astype(str).eq("上证综指"), "交易日期"
        ],
        errors="coerce",
    ).dropna()
    dates = dates.loc[
        dates.ge(pd.Timestamp(start_date)) & dates.le(pd.Timestamp(end_date))
    ]
    dates = dates.drop_duplicates().sort_values()
    if dates.empty:
        raise RuntimeError(
            f"指数 Parquet 未提供 {start_date:%Y-%m-%d} 至 {end_date:%Y-%m-%d} 的上交所交易日"
        )
    return pd.DatetimeIndex(dates)


def _industry_period_context(run_date: date) -> dict[str, pd.Timestamp]:
    """生成行业涨跌幅及估值分位数所需的交易日锚点。"""
    calendar = _sse_trade_calendar(date(run_date.year - 10, 1, 1), run_date)
    run_timestamp = pd.Timestamp(run_date)
    if calendar[-1] != run_timestamp:
        raise RuntimeError(
            f"本地上交所交易日历未更新至 {run_date:%Y-%m-%d}，"
            f"当前最新日期：{calendar[-1].date()}"
        )

    def last_before(target: pd.Timestamp, label: str) -> pd.Timestamp:
        candidates = calendar[calendar < target]
        if candidates.empty:
            raise RuntimeError(f"无法确定{label}的前一交易日：{target.date()}")
        return pd.Timestamp(candidates[-1])

    def last_on_or_before(target: pd.Timestamp, label: str) -> pd.Timestamp:
        candidates = calendar[calendar <= target]
        if candidates.empty:
            raise RuntimeError(f"无法确定{label}：{target.date()}")
        return pd.Timestamp(candidates[-1])

    week_start = run_timestamp - pd.Timedelta(days=7)
    month_start = run_timestamp - pd.DateOffset(months=1)
    year_start = pd.Timestamp(run_date.year, 1, 1)
    three_year_target = run_timestamp - pd.DateOffset(years=3)
    ten_year_target = run_timestamp - pd.DateOffset(years=10)
    return {
        "current": run_timestamp,
        "dailyBase": last_before(run_timestamp, "日涨跌幅"),
        "weeklyBase": last_before(week_start, "周涨跌幅起始日"),
        "monthlyBase": last_before(month_start, "月涨跌幅起始日"),
        "yearBase": last_before(year_start, "年初至今涨跌幅起始日"),
        "threeYearStart": last_on_or_before(three_year_target, "三年前交易日"),
        "tenYearStart": last_on_or_before(ten_year_target, "十年前交易日"),
    }


def fetch_cb_industry_performance(
    run_date: date,
    period_context: dict[str, pd.Timestamp],
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从个券 Parquet 计算申万一级行业转债的四类涨跌幅。"""
    if not CB_MASTER_PARQUET.is_file():
        raise FileNotFoundError(f"未找到转债总表 Parquet：{CB_MASTER_PARQUET}")
    master = pd.read_parquet(
        CB_MASTER_PARQUET, columns=["转债代码", "申万行业"]
    ).copy()
    master["转债代码"] = master["转债代码"].astype(str).str.strip()
    code_to_industry = (
        master.drop_duplicates("转债代码", keep="last")
        .set_index("转债代码")["申万行业"]
        .astype(str)
        .str.strip()
    )

    anchor_keys = ("current", "dailyBase", "weeklyBase", "monthlyBase", "yearBase")
    anchor_dates = {pd.Timestamp(period_context[key]).normalize() for key in anchor_keys}
    parquet_paths = sorted(
        {
            _month_parquet_path(timestamp.date())
            for timestamp in anchor_dates
        }
    )
    missing_paths = [path for path in parquet_paths if not path.is_file()]
    if missing_paths:
        raise FileNotFoundError(
            "行业转债涨跌幅缺少月度 Parquet："
            + "、".join(str(path) for path in missing_paths)
        )

    frames: list[pd.DataFrame] = []
    required_columns = [
        "转债代码",
        "交易日期",
        "交易状态",
        "收盘价",
        "转股溢价率",
    ]
    for parquet_path in parquet_paths:
        frame = pd.read_parquet(parquet_path, columns=required_columns).copy()
        frame["交易日期"] = pd.to_datetime(
            frame["交易日期"], errors="coerce"
        ).dt.normalize()
        frame = frame.loc[frame["交易日期"].isin(anchor_dates)].copy()
        if not frame.empty:
            frames.append(frame)
    if not frames:
        raise RuntimeError("未从个券 Parquet 读取到行业转债涨跌幅锚点数据")

    data = pd.concat(frames, ignore_index=True)
    data["转债代码"] = data["转债代码"].astype(str).str.strip()
    data["行业名称"] = data["转债代码"].map(code_to_industry)
    data["收盘价"] = pd.to_numeric(data["收盘价"], errors="coerce")
    data["转股溢价率"] = pd.to_numeric(data["转股溢价率"], errors="coerce")
    data = data.loc[
        data["交易状态"].astype(str).str.strip().eq("交易")
        & data["行业名称"].notna()
        & data["收盘价"].gt(0)
    ].copy()
    extreme_mask = data["收盘价"].gt(INDUSTRY_EXTREME_PRICE_THRESHOLD) & data[
        "转股溢价率"
    ].gt(INDUSTRY_EXTREME_PREMIUM_THRESHOLD)
    data = data.loc[~extreme_mask]
    grouped = data.groupby(["交易日期", "行业名称"], observed=True)["收盘价"].agg(
        ["mean", "count"]
    )

    def level(industry: str, key: str) -> float | None:
        lookup = (pd.Timestamp(period_context[key]).normalize(), industry)
        if lookup not in grouped.index:
            return None
        value = grouped.loc[lookup, "mean"]
        return None if pd.isna(value) else float(value)

    def return_pct(current: float | None, base: float | None) -> float | None:
        if current is None or base in (None, 0):
            return None
        return (current / base - 1.0) * 100.0

    rows: list[dict[str, object]] = []
    for code, industry in INDUSTRY_INDEX_SPECS:
        current = level(industry, "current")
        current_lookup = (period_context["current"], industry)
        current_count = (
            int(grouped.loc[current_lookup, "count"])
            if current_lookup in grouped.index
            else 0
        )
        rows.append(
            {
                "行业代码": code,
                "行业名称": industry,
                "转债日涨跌幅": return_pct(current, level(industry, "dailyBase")),
                "转债周涨跌幅": return_pct(current, level(industry, "weeklyBase")),
                "转债月涨跌幅": return_pct(current, level(industry, "monthlyBase")),
                "转债年初至今涨跌幅": return_pct(
                    current, level(industry, "yearBase")
                ),
                "转债当日样本数": current_count,
            }
        )

    source: dict[str, object] = {
        "parquetRoot": str(CB_PARQUET_ROOT.relative_to(WORKSPACE)),
        "masterParquet": str(CB_MASTER_PARQUET.relative_to(WORKSPACE)),
        "sampleRule": (
            "交易状态=交易，收盘价>0；按总表申万一级行业取个券收盘价算术均值；"
            f"剔除收盘价>{INDUSTRY_EXTREME_PRICE_THRESHOLD:g}且转股溢价率>"
            f"{INDUSTRY_EXTREME_PREMIUM_THRESHOLD:g}%的极端样本"
        ),
        "returnRule": (
            "行业转债涨跌幅=当日行业平均收盘价/对应基准日行业平均收盘价-1；"
            "日、周、月、年初至今分别对应前一交易日、7个自然日前的前一交易日、"
            "上月同日前的前一交易日、上一年最后交易日"
        ),
        "anchorDates": {
            key: f"{pd.Timestamp(period_context[key]):%Y-%m-%d}"
            for key in anchor_keys
        },
    }
    return pd.DataFrame(rows), source


def fetch_industry_index_performance_and_valuation(
    run_date: date,
    period_context: dict[str, pd.Timestamp],
) -> tuple[pd.DataFrame, dict[str, object]]:
    """通过 Wind Excel 标量公式取得行业指数涨跌幅、PE/PB及分位数。"""
    run_text = f"{run_date:%Y%m%d}"
    three_year_text = f"{period_context['threeYearStart']:%Y%m%d}"
    ten_year_text = f"{period_context['tenYearStart']:%Y%m%d}"
    formula_specs = (
        ("正股日涨跌幅", "i_dq_pctchange", lambda code: f'=@i_dq_pctchange("{code}","{run_text}")'),
        ("正股周涨跌幅", "s_wq_pctchange", lambda code: f'=@s_wq_pctchange("{code}","{run_text}")'),
        ("正股月涨跌幅", "s_mq_pctchange", lambda code: f'=@s_mq_pctchange("{code}","{run_text}")'),
        ("正股年初至今涨跌幅", "s_yq_pctchange", lambda code: f'=@s_yq_pctchange("{code}","{run_text}")'),
        ("行业指数PE", "s_val_pe_ttm", lambda code: f'=@s_val_pe_ttm("{code}","{run_text}")'),
        (
            "PE三年分位数",
            "i_val_pe_percentile",
            lambda code: (
                f'=@i_val_pe_percentile("{code}","{run_text}",'
                f'"{three_year_text}","{run_text}")'
            ),
        ),
        (
            "PE十年分位数",
            "i_val_pe_percentile",
            lambda code: (
                f'=@i_val_pe_percentile("{code}","{run_text}",'
                f'"{ten_year_text}","{run_text}")'
            ),
        ),
        ("行业指数PB", "s_val_pb_lf", lambda code: f'=@s_val_pb_lf("{code}","{run_text}")'),
        (
            "PB三年分位数",
            "i_val_pb_percentile",
            lambda code: (
                f'=@i_val_pb_percentile("{code}","{run_text}",'
                f'"{three_year_text}","{run_text}")'
            ),
        ),
        (
            "PB十年分位数",
            "i_val_pb_percentile",
            lambda code: (
                f'=@i_val_pb_percentile("{code}","{run_text}",'
                f'"{ten_year_text}","{run_text}")'
            ),
        ),
    )
    formula_matrix = [
        [builder(code) for _, _, builder in formula_specs]
        for code, _ in INDUSTRY_INDEX_SPECS
    ]

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    raw_values: tuple[tuple[object, ...], ...] | None = None
    stable_signature: tuple[float, ...] | None = None
    stable_rounds = 0
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = False
        excel.EnableEvents = True
        try:
            excel.COMAddIns("WDF.Addin").Connect = True
        except Exception:
            pass

        workbook = excel.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        sheet.Name = "行业指数与估值"
        sheet.Range("A1:L1").Value = (
            (
                "行业代码",
                "行业名称",
                *[label for label, _, _ in formula_specs],
            ),
        )
        sheet.Range("A2:B31").Value = tuple(
            (code, industry) for code, industry in INDUSTRY_INDEX_SPECS
        )
        try:
            sheet.Range("C2:L31").Formula2 = tuple(
                tuple(row) for row in formula_matrix
            )
        except Exception:
            for row_index, formulas in enumerate(formula_matrix, start=2):
                for column_index, formula in enumerate(formulas, start=3):
                    try:
                        sheet.Cells(row_index, column_index).Formula2 = formula
                    except Exception:
                        sheet.Cells(row_index, column_index).Formula = formula.replace(
                            "=@", "="
                        )
        excel.CalculateFullRebuild()

        started = time.monotonic()
        while time.monotonic() - started < 120:
            time.sleep(1)
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            candidate = sheet.Range("C2:L31").Value
            if not isinstance(candidate, tuple) or len(candidate) != len(
                INDUSTRY_INDEX_SPECS
            ):
                continue
            numeric_values: list[float] = []
            valid = True
            for row in candidate:
                if not isinstance(row, tuple) or len(row) != len(formula_specs):
                    valid = False
                    break
                for value in row:
                    try:
                        number = float(value)
                    except (TypeError, ValueError):
                        valid = False
                        break
                    if not math.isfinite(number):
                        valid = False
                        break
                    numeric_values.append(number)
                if not valid:
                    break
            if not valid:
                continue
            signature = tuple(round(value, 10) for value in numeric_values)
            if signature == stable_signature:
                stable_rounds += 1
            else:
                stable_signature = signature
                stable_rounds = 0
            raw_values = candidate
            if stable_rounds >= 2:
                break
        if raw_values is None or stable_rounds < 2:
            preview = sheet.Range("A1:L5").Value
            raise RuntimeError(
                "Wind 行业指数涨跌幅及估值公式刷新超时，前5行状态："
                f"{preview!r}"
            )
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()

    rows: list[dict[str, object]] = []
    assert raw_values is not None
    for (code, industry), values, formulas in zip(
        INDUSTRY_INDEX_SPECS, raw_values, formula_matrix
    ):
        row: dict[str, object] = {"行业代码": code, "行业名称": industry}
        for (label, _, _), value, formula in zip(formula_specs, values, formulas):
            row[label] = float(value)
            row[f"{label}公式"] = formula
        rows.append(row)

    source: dict[str, object] = {
        "source": "Wind Excel 标量公式",
        "returnFunctions": [
            "i_dq_pctchange",
            "s_wq_pctchange",
            "s_mq_pctchange",
            "s_yq_pctchange",
        ],
        "valuationFunctions": [
            "s_val_pe_ttm",
            "i_val_pe_percentile",
            "s_val_pb_lf",
            "i_val_pb_percentile",
        ],
        "runDate": f"{run_date:%Y-%m-%d}",
        "threeYearStart": f"{period_context['threeYearStart']:%Y-%m-%d}",
        "tenYearStart": f"{period_context['tenYearStart']:%Y-%m-%d}",
        "returnUnit": "Wind函数直接返回百分数，例如2.05表示2.05%",
        "mapping": [
            {"code": code, "name": industry}
            for code, industry in INDUSTRY_INDEX_SPECS
        ],
    }
    return pd.DataFrame(rows), source


def fetch_industry_market_table(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """汇总行业正股指数、转债涨跌幅与正股估值。"""
    period_context = _industry_period_context(run_date)
    cb_data, cb_source = fetch_cb_industry_performance(run_date, period_context)
    equity_data, equity_source = fetch_industry_index_performance_and_valuation(
        run_date, period_context
    )
    result = equity_data.merge(
        cb_data,
        on=["行业代码", "行业名称"],
        how="left",
        validate="one_to_one",
    )
    result = result.sort_values(
        "正股日涨跌幅", ascending=False, na_position="last"
    ).reset_index(drop=True)
    source: dict[str, object] = {
        "runDate": f"{run_date:%Y-%m-%d}",
        "sortRule": "按正股行业指数日涨跌幅降序",
        "equity": equity_source,
        "convertibleBond": cb_source,
    }
    return result, source


def fetch_cb_etf_share_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """通过 Wind Excel WSD 获取两只可转债 ETF 份额，并计算单日净申赎。"""
    trade_calendar = _sse_trade_calendar(ETF_SHARE_START_DATE, run_date)
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    latest_formula_value: object = None
    evaluated_formula = ETF_SHARE_WSD_FORMULA
    returned_rows = 0
    raw_rows: list[tuple[object, object]] = []
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        try:
            excel.COMAddIns("WDF.Addin").Connect = True
        except Exception:
            pass

        workbook = excel.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        sheet.Name = "ETF份额"
        sheet.Range("B1").Value = ETF_SHARE_START_SERIAL
        sheet.Range("B2").Value = datetime(
            run_date.year, run_date.month, run_date.day
        )
        sheet.Range("B2").NumberFormat = "yyyy-mm-dd"
        sheet.Range("B4").Value = ETF_SHARE_SPECS[0][1]
        sheet.Range("C4").Value = ETF_SHARE_SPECS[1][1]
        try:
            sheet.Range("A6").Formula2 = ETF_SHARE_WSD_FORMULA
        except Exception:
            sheet.Range("A6").Formula = ETF_SHARE_WSD_FORMULA.replace("=@", "=")
        excel.CalculateFullRebuild()

        stable_signature: tuple[object, ...] | None = None
        stable_rounds = 0
        started = time.monotonic()
        while time.monotonic() - started < 90:
            time.sleep(1)
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            latest_formula_value = sheet.Range("A6").Value
            try:
                evaluated_formula = str(sheet.Range("A6").Formula2)
            except Exception:
                evaluated_formula = str(sheet.Range("A6").Formula)
            rows_match = re.search(r"rows=(\d+)", evaluated_formula, flags=re.I)
            if rows_match is None:
                continue
            candidate_rows = int(rows_match.group(1))
            if candidate_rows < 2:
                continue
            values = sheet.Range(f"A6:B{candidate_rows + 5}").Value
            if not isinstance(values, tuple) or len(values) != candidate_rows:
                continue
            candidate_raw = [
                (
                    row[0] if isinstance(row, tuple) and len(row) >= 1 else None,
                    row[1] if isinstance(row, tuple) and len(row) >= 2 else None,
                )
                for row in values
            ]
            numeric_tail = []
            for row in candidate_raw[-5:]:
                numeric_tail.extend(row)
            if not any(
                isinstance(value, (int, float)) and math.isfinite(float(value))
                for value in numeric_tail
            ):
                continue
            signature = (candidate_rows, *numeric_tail)
            if signature == stable_signature:
                stable_rounds += 1
            else:
                stable_signature = signature
                stable_rounds = 0
            returned_rows = candidate_rows
            raw_rows = candidate_raw
            if stable_rounds >= 2:
                break

        if not raw_rows or stable_rounds < 2:
            raise RuntimeError(
                "Wind WSD ETF份额数据刷新超时："
                f"公式值={latest_formula_value!r}，公式={evaluated_formula}"
            )
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()

    if returned_rows > len(trade_calendar):
        raise RuntimeError(
            "Wind WSD 返回行数超过本地上交所交易日数量："
            f"Wind={returned_rows}，交易日={len(trade_calendar)}"
        )
    aligned_dates = trade_calendar[:returned_rows]
    output = pd.DataFrame(
        {
            "交易日期": aligned_dates,
            "博时可转债ETF份额_万份": [row[0] for row in raw_rows],
            "海富通可转债ETF份额_万份": [row[1] for row in raw_rows],
        }
    )
    for name, _ in ETF_SHARE_SPECS:
        raw_column = f"{name}份额_万份"
        share_column = f"{name}份额_亿份"
        flow_column = f"{name}净申赎_亿份"
        output[raw_column] = pd.to_numeric(output[raw_column], errors="coerce")
        output[share_column] = output[raw_column] / 10000.0
        output[flow_column] = output[share_column].diff()
    output = output.dropna(
        subset=[f"{name}份额_亿份" for name, _ in ETF_SHARE_SPECS], how="all"
    ).reset_index(drop=True)
    if output.empty:
        raise RuntimeError("Wind WSD 未返回两只可转债 ETF 的有效份额数据")

    details: dict[str, dict[str, object]] = {}
    latest_effective_dates: list[pd.Timestamp] = []
    for name, code in ETF_SHARE_SPECS:
        share_column = f"{name}份额_亿份"
        flow_column = f"{name}净申赎_亿份"
        valid = output.dropna(subset=[share_column]).sort_values("交易日期")
        if valid.empty:
            raise RuntimeError(f"Wind WSD 未返回{name}（{code}）份额")
        latest = valid.iloc[-1]
        latest_effective_dates.append(pd.Timestamp(latest["交易日期"]))
        details[name] = {
            "code": code,
            "latestDate": f"{pd.Timestamp(latest['交易日期']):%Y-%m-%d}",
            "latestShareYi": float(latest[share_column]),
            "latestNetSubscriptionYi": (
                None if pd.isna(latest[flow_column]) else float(latest[flow_column])
            ),
        }
    latest_effective_date = max(latest_effective_dates)
    source: dict[str, object] = {
        "source": "Wind Excel WSD/unit_fundshare_total",
        "formula": ETF_SHARE_WSD_FORMULA,
        "evaluatedFormula": evaluated_formula,
        "startDate": f"{ETF_SHARE_START_DATE:%Y-%m-%d}",
        "requestedEndDate": f"{run_date:%Y-%m-%d}",
        "latestDate": f"{latest_effective_date:%Y-%m-%d}",
        "codes": [code for _, code in ETF_SHARE_SPECS],
        "field": "unit_fundshare_total",
        "tradingCalendar": "SSE",
        "rawUnit": "万份",
        "unitRule": "unit_fundshare_total原始值（万份）除以10000，转换为亿份",
        "netSubscriptionRule": "当日总份额减前一有效交易日总份额；正值为净申购，负值为净赎回",
        "returnedRows": returned_rows,
        "points": len(output),
        "funds": details,
    }
    return output, source


def fetch_index_turnover(start: date, end: date) -> pd.DataFrame:
    result = None
    for attempt in range(INDEX_TURNOVER_REQUEST_ATTEMPTS):
        result = THS_DS(
            "000001.SH,399001.SZ,000832.CSI",
            "ths_trans_amt_index",
            "",
            "mode:thscode,block:history",
            f"{start:%Y-%m-%d}",
            f"{end:%Y-%m-%d}",
        )
        if getattr(result, "errorcode", None) != -205:
            break
        if attempt + 1 < INDEX_TURNOVER_REQUEST_ATTEMPTS:
            time.sleep(INDEX_TURNOVER_RETRY_DELAY_SECONDS)
    assert result is not None
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


def fetch_daily_valuation_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从指数 Parquet 读取反三次与多因子修正百元拟合溢价率。"""
    if not INDEX_PARQUET.is_file():
        raise FileNotFoundError(f"未找到指数 Parquet：{INDEX_PARQUET}")

    data = pd.read_parquet(INDEX_PARQUET)
    required = {"指数名称", "交易日期", "指数值"}
    if not required.issubset(data.columns):
        raise RuntimeError(
            f"指数 Parquet 字段异常，缺少：{sorted(required - set(data.columns))}"
        )
    names = (INVERSE_CUBIC_VALUATION_NAME, MULTIFACTOR_VALUATION_NAME)
    data = data.loc[
        data["指数名称"].astype(str).isin(names),
        ["指数名称", "交易日期", "指数值"],
    ].copy()
    data["交易日期"] = pd.to_datetime(data["交易日期"], errors="coerce").dt.normalize()
    data["指数值"] = pd.to_numeric(data["指数值"], errors="coerce")
    data = data.dropna(subset=["指数名称", "交易日期", "指数值"])
    data = data.loc[
        data["交易日期"].between(
            pd.Timestamp(VALUATION_START_DATE), pd.Timestamp(run_date), inclusive="both"
        )
    ]
    data = data.drop_duplicates(["指数名称", "交易日期"], keep="last")
    wide = (
        data.pivot(index="交易日期", columns="指数名称", values="指数值")
        .sort_index()
        .rename_axis(columns=None)
        .reset_index()
    )
    if INVERSE_CUBIC_VALUATION_NAME not in wide.columns:
        raise RuntimeError("指数 Parquet 中缺少百元拟合溢价率历史序列")
    inverse = wide.dropna(subset=[INVERSE_CUBIC_VALUATION_NAME]).copy()
    if inverse.empty or inverse["交易日期"].iloc[-1].date() != run_date:
        latest = inverse["交易日期"].max() if not inverse.empty else pd.NaT
        raise RuntimeError(
            f"百元拟合溢价率未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
        )
    if MULTIFACTOR_VALUATION_NAME not in wide.columns:
        wide[MULTIFACTOR_VALUATION_NAME] = np.nan
    multifactor = wide.dropna(subset=[MULTIFACTOR_VALUATION_NAME])
    if multifactor.empty or multifactor["交易日期"].iloc[-1].date() != run_date:
        latest = multifactor["交易日期"].max() if not multifactor.empty else pd.NaT
        raise RuntimeError(
            f"多因子修正百元拟合溢价率未更新至 {run_date:%Y-%m-%d}，"
            f"当前最新日期：{latest}"
        )

    inverse_series = inverse.set_index("交易日期")[INVERSE_CUBIC_VALUATION_NAME]
    latest_value = float(inverse_series.iloc[-1])
    previous_value = float(inverse_series.iloc[-2])
    percentile = float(inverse_series.le(latest_value).mean() * 100)
    source = {
        "parquet": str(INDEX_PARQUET.relative_to(WORKSPACE)),
        "startDate": f"{VALUATION_START_DATE:%Y-%m-%d}",
        "latestDate": f"{inverse_series.index[-1]:%Y-%m-%d}",
        "previousDate": f"{inverse_series.index[-2]:%Y-%m-%d}",
        "latestValue": latest_value,
        "previousValue": previous_value,
        "dailyChangePctPoint": latest_value - previous_value,
        "percentileSince2019": percentile,
        "inverseMethod": "反三次；平价70—130、换手率<50%，剔除溢价率两端各3%",
        "multifactorMethod": "幂衰减基准曲线加六因子OLS修正；平价50—200、换手率<50%，剔除溢价率两端各3%",
    }
    return wide, source


def calculate_valuation_repair_index(data: pd.DataFrame) -> pd.DataFrame:
    """计算百元溢价率 100 日走廊位置及中证转债指数后推收益。"""
    required = {"交易日期", INVERSE_CUBIC_VALUATION_NAME, CB_INDEX_PARQUET_NAME}
    if not required.issubset(data.columns):
        raise RuntimeError(f"估值修复指数输入缺少字段：{sorted(required - set(data.columns))}")
    result = data.loc[:, list(required)].copy()
    result["交易日期"] = pd.to_datetime(result["交易日期"], errors="coerce").dt.normalize()
    for column in (INVERSE_CUBIC_VALUATION_NAME, CB_INDEX_PARQUET_NAME):
        result[column] = pd.to_numeric(result[column], errors="coerce")
    result = (
        result.dropna(subset=list(required))
        .sort_values("交易日期")
        .drop_duplicates("交易日期", keep="last")
        .reset_index(drop=True)
    )
    if len(result) < VALUATION_REPAIR_WINDOW:
        raise RuntimeError(
            f"估值修复指数历史不足{VALUATION_REPAIR_WINDOW}个交易日"
        )

    valuation = result[INVERSE_CUBIC_VALUATION_NAME]
    rolling_mean = valuation.rolling(VALUATION_REPAIR_WINDOW).mean()
    rolling_std = valuation.rolling(VALUATION_REPAIR_WINDOW).std(ddof=0)
    lower_corridor = rolling_mean - 2.0 * rolling_std
    upper_corridor = rolling_mean + 2.0 * rolling_std
    result[VALUATION_REPAIR_INDEX_NAME] = 100.0 * (
        upper_corridor - valuation
    ) / (upper_corridor - lower_corridor).replace(0, np.nan)

    cb_index = result[CB_INDEX_PARQUET_NAME].to_numpy(dtype=float)
    positions = np.arange(len(result))
    forward_positions = np.minimum(
        positions + VALUATION_REPAIR_FORWARD_TRADING_DAYS,
        len(result) - 1,
    )
    result[VALUATION_REPAIR_FORWARD_RETURN_NAME] = (
        cb_index[forward_positions] / cb_index - 1.0
    ) * 100.0
    return result


def fetch_valuation_repair_index(run_date: date) -> pd.DataFrame:
    """从指数 Parquet 读取百元拟合溢价率和中证转债指数计算修复指数。"""
    if not INDEX_PARQUET.is_file():
        raise FileNotFoundError(f"未找到指数 Parquet：{INDEX_PARQUET}")
    data = pd.read_parquet(
        INDEX_PARQUET, columns=["指数名称", "交易日期", "指数值"]
    )
    names = (INVERSE_CUBIC_VALUATION_NAME, CB_INDEX_PARQUET_NAME)
    data = data.loc[
        data["指数名称"].astype(str).isin(names),
        ["指数名称", "交易日期", "指数值"],
    ].copy()
    data["交易日期"] = pd.to_datetime(data["交易日期"], errors="coerce").dt.normalize()
    data["指数值"] = pd.to_numeric(data["指数值"], errors="coerce")
    data = data.dropna(subset=["指数名称", "交易日期", "指数值"])
    data = data.loc[data["交易日期"].le(pd.Timestamp(run_date))]
    wide = (
        data.drop_duplicates(["指数名称", "交易日期"], keep="last")
        .pivot(index="交易日期", columns="指数名称", values="指数值")
        .rename_axis(columns=None)
        .reset_index()
    )
    result = calculate_valuation_repair_index(wide)
    if result["交易日期"].iloc[-1].date() != run_date:
        latest = result["交易日期"].iloc[-1]
        raise RuntimeError(
            f"估值修复指数未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest:%Y-%m-%d}"
        )
    return result


def fetch_parity_group_valuation_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从指数 Parquet 直接读取分平价多因子修正拟合溢价率。"""
    if not INDEX_PARQUET.is_file():
        raise FileNotFoundError(f"未找到指数 Parquet：{INDEX_PARQUET}")

    data = pd.read_parquet(INDEX_PARQUET)
    required = {"指数名称", "交易日期", "指数值"}
    if not required.issubset(data.columns):
        raise RuntimeError(
            f"指数 Parquet 字段异常，缺少：{sorted(required - set(data.columns))}"
        )

    parquet_names = [parquet_name for _, parquet_name in PARITY_GROUP_SPECS]
    data = data.loc[
        data["指数名称"].astype(str).isin(parquet_names),
        ["指数名称", "交易日期", "指数值"],
    ].copy()
    data["交易日期"] = pd.to_datetime(data["交易日期"], errors="coerce").dt.normalize()
    data["指数值"] = pd.to_numeric(data["指数值"], errors="coerce")
    data = data.dropna(subset=["指数名称", "交易日期", "指数值"])
    data = data.loc[
        data["交易日期"].between(
            pd.Timestamp(VALUATION_START_DATE), pd.Timestamp(run_date), inclusive="both"
        )
    ]
    data = data.drop_duplicates(["指数名称", "交易日期"], keep="last")
    wide = (
        data.pivot(index="交易日期", columns="指数名称", values="指数值")
        .sort_index()
        .rename_axis(columns=None)
        .reset_index()
    )
    if wide.empty:
        raise RuntimeError("指数 Parquet 中缺少分平价多因子修正拟合溢价率历史序列")

    group_details: dict[str, dict[str, object]] = {}
    current_changes: dict[str, float] = {}
    for group, parquet_name in PARITY_GROUP_SPECS:
        if parquet_name not in wide.columns:
            raise RuntimeError(f"指数 Parquet 中缺少“{parquet_name}”")
        series = (
            wide.dropna(subset=[parquet_name])
            .set_index("交易日期")[parquet_name]
            .sort_index()
        )
        if len(series) < 2:
            raise RuntimeError(f"“{parquet_name}”有效历史记录不足2条")
        latest_date = pd.Timestamp(series.index[-1])
        previous_date = pd.Timestamp(series.index[-2])
        latest_value = float(series.iloc[-1])
        previous_value = float(series.iloc[-2])
        daily_change = latest_value - previous_value
        group_details[group] = {
            "parquetName": parquet_name,
            "latestDate": f"{latest_date:%Y-%m-%d}",
            "latestValue": latest_value,
            "previousDate": f"{previous_date:%Y-%m-%d}",
            "previousValue": previous_value,
            "dailyChangePctPoint": daily_change,
        }
        if latest_date.date() == run_date:
            current_changes[group] = daily_change

    if not current_changes:
        latest_by_group = {
            group: details["latestDate"] for group, details in group_details.items()
        }
        raise RuntimeError(
            f"分平价多因子修正拟合溢价率均未更新至 {run_date:%Y-%m-%d}："
            f"{latest_by_group}"
        )

    largest_change_group = max(current_changes, key=lambda group: abs(current_changes[group]))
    source: dict[str, object] = {
        "parquet": str(INDEX_PARQUET.relative_to(WORKSPACE)),
        "startDate": f"{VALUATION_START_DATE:%Y-%m-%d}",
        "runDate": f"{run_date:%Y-%m-%d}",
        "readRule": "直接读取指数 Parquet 中既有分平价多因子修正拟合溢价率，不重新计算",
        "groupDetails": group_details,
        "largestChangeGroup": largest_change_group,
        "largestChangePctPoint": current_changes[largest_change_group],
    }
    return wide, source


def aggregate_parity_interval_premium_series(data: pd.DataFrame) -> pd.DataFrame:
    """按日度数据更新的七档平价口径计算余额加权转股溢价率。"""
    required_columns = {"交易日期", "交易状态", "余额", "平价", "转股溢价率"}
    if not required_columns.issubset(data.columns):
        raise RuntimeError(
            "分平价区间转股溢价率缺少字段："
            f"{sorted(required_columns - set(data.columns))}"
        )

    base = data.loc[:, list(required_columns)].copy()
    base["交易日期"] = pd.to_datetime(
        base["交易日期"], errors="coerce"
    ).dt.normalize()
    for column in ("余额", "平价", "转股溢价率"):
        base[column] = pd.to_numeric(base[column], errors="coerce")
        base.loc[~np.isfinite(base[column]), column] = np.nan
    base = base.loc[
        base["交易日期"].notna()
        & base["交易状态"].astype(str).str.strip().eq("交易")
        & base[["余额", "平价", "转股溢价率"]].notna().all(axis=1)
    ].copy()
    if base.empty:
        raise RuntimeError("分平价区间转股溢价率没有有效交易样本")

    base["平价区间"] = pd.cut(
        base["平价"],
        bins=(-np.inf, 80.0, 90.0, 100.0, 110.0, 120.0, 130.0, np.inf),
        labels=PARITY_INTERVAL_PREMIUM_ASCENDING_COLUMNS,
        right=True,
        include_lowest=True,
    )
    base["加权值"] = base["转股溢价率"] * base["余额"]
    grouped = base.groupby(["交易日期", "平价区间"], observed=True).agg(
        加权值=("加权值", "sum"),
        有效余额=("余额", "sum"),
    )
    weighted = grouped["加权值"].div(grouped["有效余额"].replace(0, np.nan))
    result = (
        weighted.unstack("平价区间")
        .reindex(columns=PARITY_INTERVAL_PREMIUM_COLUMNS)
        .rename_axis(columns=None)
        .reset_index()
        .sort_values("交易日期")
        .reset_index(drop=True)
    )
    return result


def fetch_parity_interval_premium_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从月度个券 Parquet 计算2019年以来分平价区间转股溢价率。"""
    end_month = f"{run_date:%Y%m}"
    parquet_paths = sorted(
        path
        for path in CB_PARQUET_ROOT.glob("20*/20*.parquet")
        if "201901" <= path.stem <= end_month
    )
    if not parquet_paths:
        raise FileNotFoundError("未找到2019年以来的月度转债 Parquet")

    columns = ["交易日期", "交易状态", "余额", "平价", "转股溢价率"]
    frames: list[pd.DataFrame] = []
    for parquet_path in parquet_paths:
        frame = pd.read_parquet(parquet_path, columns=columns).copy()
        frame["交易日期"] = pd.to_datetime(
            frame["交易日期"], errors="coerce"
        ).dt.normalize()
        frame = frame.loc[
            frame["交易日期"].between(
                pd.Timestamp(VALUATION_START_DATE),
                pd.Timestamp(run_date),
                inclusive="both",
            )
        ]
        if not frame.empty:
            frames.append(frame)
    if not frames:
        raise RuntimeError("月度转债 Parquet 中没有分平价区间转股溢价率数据")

    result = aggregate_parity_interval_premium_series(
        pd.concat(frames, ignore_index=True)
    )
    latest_date = pd.Timestamp(result["交易日期"].iloc[-1]).date()
    if latest_date != run_date:
        raise RuntimeError(
            f"分平价区间转股溢价率未更新至 {run_date:%Y-%m-%d}，"
            f"当前最新日期：{latest_date:%Y-%m-%d}"
        )
    source: dict[str, object] = {
        "parquetRoot": str(CB_PARQUET_ROOT.relative_to(WORKSPACE)),
        "startDate": f"{result['交易日期'].min():%Y-%m-%d}",
        "runDate": f"{run_date:%Y-%m-%d}",
        "sampleRule": (
            "交易状态=交易；按平价划分≤80、80-90、90-100、100-110、"
            "110-120、120-130、>130七档；各档转股溢价率按余额加权"
        ),
    }
    return result, source


def fetch_equity_bond_group_valuation_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从指数 Parquet 读取偏股、平衡、偏债三类多因子拟合溢价率。"""
    if not INDEX_PARQUET.is_file():
        raise FileNotFoundError(f"未找到指数 Parquet：{INDEX_PARQUET}")
    names = [parquet_name for _, parquet_name in EQUITY_BOND_GROUP_SPECS]
    data = pd.read_parquet(
        INDEX_PARQUET, columns=["指数名称", "交易日期", "指数值"]
    )
    data = data.loc[
        data["指数名称"].astype(str).isin(names),
        ["指数名称", "交易日期", "指数值"],
    ].copy()
    data["交易日期"] = pd.to_datetime(data["交易日期"], errors="coerce").dt.normalize()
    data["指数值"] = pd.to_numeric(data["指数值"], errors="coerce")
    data = data.dropna(subset=["指数名称", "交易日期", "指数值"])
    data = data.loc[
        data["交易日期"].between(
            pd.Timestamp(VALUATION_START_DATE), pd.Timestamp(run_date), inclusive="both"
        )
    ].drop_duplicates(["指数名称", "交易日期"], keep="last")
    wide = (
        data.pivot(index="交易日期", columns="指数名称", values="指数值")
        .sort_index()
        .rename_axis(columns=None)
        .reset_index()
    )
    details: dict[str, dict[str, object]] = {}
    for group, parquet_name in EQUITY_BOND_GROUP_SPECS:
        if parquet_name not in wide.columns:
            raise RuntimeError(f"指数 Parquet 中缺少“{parquet_name}”")
        series = wide.dropna(subset=[parquet_name]).set_index("交易日期")[parquet_name]
        if len(series) < 2 or pd.Timestamp(series.index[-1]).date() != run_date:
            latest = series.index[-1] if len(series) else pd.NaT
            raise RuntimeError(
                f"“{parquet_name}”未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
            )
        details[group] = {
            "latestValue": float(series.iloc[-1]),
            "previousValue": float(series.iloc[-2]),
            "dailyChangePctPoint": float(series.iloc[-1] - series.iloc[-2]),
        }
    return wide, {
        "parquet": str(INDEX_PARQUET.relative_to(WORKSPACE)),
        "startDate": f"{VALUATION_START_DATE:%Y-%m-%d}",
        "runDate": f"{run_date:%Y-%m-%d}",
        "readRule": "直接读取指数 Parquet 中既有股债型多因子修正拟合溢价率，不重新计算",
        "groupDetails": details,
    }


def fetch_rating_group_valuation_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从指数 Parquet 读取已计算好的分评级多因子拟合溢价率。"""
    if not INDEX_PARQUET.is_file():
        raise FileNotFoundError(f"未找到指数 Parquet：{INDEX_PARQUET}")
    names = [parquet_name for _, parquet_name in RATING_GROUP_SPECS]
    data = pd.read_parquet(
        INDEX_PARQUET, columns=["指数名称", "交易日期", "指数值"]
    )
    data = data.loc[
        data["指数名称"].astype(str).isin(names),
        ["指数名称", "交易日期", "指数值"],
    ].copy()
    data["交易日期"] = pd.to_datetime(data["交易日期"], errors="coerce").dt.normalize()
    data["指数值"] = pd.to_numeric(data["指数值"], errors="coerce")
    data = data.dropna(subset=["指数名称", "交易日期", "指数值"])
    data = data.loc[
        data["交易日期"].between(
            pd.Timestamp(VALUATION_START_DATE), pd.Timestamp(run_date), inclusive="both"
        )
    ].drop_duplicates(["指数名称", "交易日期"], keep="last")
    wide = (
        data.pivot(index="交易日期", columns="指数名称", values="指数值")
        .sort_index()
        .rename_axis(columns=None)
        .reset_index()
    )
    details: dict[str, dict[str, object]] = {}
    for group, parquet_name in RATING_GROUP_SPECS:
        if parquet_name not in wide.columns:
            raise RuntimeError(f"指数 Parquet 中缺少“{parquet_name}”")
        series = wide.dropna(subset=[parquet_name]).set_index("交易日期")[parquet_name]
        if len(series) < 2 or pd.Timestamp(series.index[-1]).date() != run_date:
            latest = series.index[-1] if len(series) else pd.NaT
            raise RuntimeError(
                f"“{parquet_name}”未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
            )
        details[group] = {
            "latestValue": float(series.iloc[-1]),
            "previousValue": float(series.iloc[-2]),
            "dailyChangePctPoint": float(series.iloc[-1] - series.iloc[-2]),
        }
    return wide, {
        "parquet": str(INDEX_PARQUET.relative_to(WORKSPACE)),
        "startDate": f"{VALUATION_START_DATE:%Y-%m-%d}",
        "runDate": f"{run_date:%Y-%m-%d}",
        "readRule": "直接读取指数 Parquet 中既有分评级多因子修正拟合溢价率，不重新计算",
        "groupDetails": details,
    }


def fetch_maturity_group_valuation_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从指数 Parquet 读取已计算好的分剩余期限多因子拟合溢价率。"""
    if not INDEX_PARQUET.is_file():
        raise FileNotFoundError(f"未找到指数 Parquet：{INDEX_PARQUET}")
    names = [parquet_name for _, parquet_name in MATURITY_GROUP_SPECS]
    data = pd.read_parquet(
        INDEX_PARQUET, columns=["指数名称", "交易日期", "指数值"]
    )
    data = data.loc[
        data["指数名称"].astype(str).isin(names),
        ["指数名称", "交易日期", "指数值"],
    ].copy()
    data["交易日期"] = pd.to_datetime(data["交易日期"], errors="coerce").dt.normalize()
    data["指数值"] = pd.to_numeric(data["指数值"], errors="coerce")
    data = data.dropna(subset=["指数名称", "交易日期", "指数值"])
    data = data.loc[
        data["交易日期"].between(
            pd.Timestamp(VALUATION_START_DATE), pd.Timestamp(run_date), inclusive="both"
        )
    ].drop_duplicates(["指数名称", "交易日期"], keep="last")
    wide = (
        data.pivot(index="交易日期", columns="指数名称", values="指数值")
        .sort_index()
        .rename_axis(columns=None)
        .reset_index()
    )
    details: dict[str, dict[str, object]] = {}
    for group, parquet_name in MATURITY_GROUP_SPECS:
        if parquet_name not in wide.columns:
            raise RuntimeError(f"指数 Parquet 中缺少“{parquet_name}”")
        series = wide.dropna(subset=[parquet_name]).set_index("交易日期")[parquet_name]
        if len(series) < 2 or pd.Timestamp(series.index[-1]).date() != run_date:
            latest = series.index[-1] if len(series) else pd.NaT
            raise RuntimeError(
                f"“{parquet_name}”未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
            )
        details[group] = {
            "latestValue": float(series.iloc[-1]),
            "previousValue": float(series.iloc[-2]),
            "dailyChangePctPoint": float(series.iloc[-1] - series.iloc[-2]),
        }
    return wide, {
        "parquet": str(INDEX_PARQUET.relative_to(WORKSPACE)),
        "startDate": f"{VALUATION_START_DATE:%Y-%m-%d}",
        "runDate": f"{run_date:%Y-%m-%d}",
        "readRule": "直接读取指数 Parquet 中既有分剩余期限多因子修正拟合溢价率，不重新计算",
        "groupDetails": details,
    }


def fetch_precalculated_classification_valuation_series(
    run_date: date,
    specs: tuple[tuple[str, str], ...],
    classification_label: str,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从指数 Parquet 读取指定分类的已计算多因子拟合溢价率。"""
    if not INDEX_PARQUET.is_file():
        raise FileNotFoundError(f"未找到指数 Parquet：{INDEX_PARQUET}")
    names = [parquet_name for _, parquet_name in specs]
    data = pd.read_parquet(
        INDEX_PARQUET, columns=["指数名称", "交易日期", "指数值"]
    )
    data = data.loc[
        data["指数名称"].astype(str).isin(names),
        ["指数名称", "交易日期", "指数值"],
    ].copy()
    data["交易日期"] = pd.to_datetime(data["交易日期"], errors="coerce").dt.normalize()
    data["指数值"] = pd.to_numeric(data["指数值"], errors="coerce")
    data = data.dropna(subset=["指数名称", "交易日期", "指数值"])
    data = data.loc[
        data["交易日期"].between(
            pd.Timestamp(VALUATION_START_DATE), pd.Timestamp(run_date), inclusive="both"
        )
    ].drop_duplicates(["指数名称", "交易日期"], keep="last")
    wide = (
        data.pivot(index="交易日期", columns="指数名称", values="指数值")
        .sort_index()
        .rename_axis(columns=None)
        .reset_index()
    )
    details: dict[str, dict[str, object]] = {}
    for group, parquet_name in specs:
        if parquet_name not in wide.columns:
            raise RuntimeError(f"指数 Parquet 中缺少“{parquet_name}”")
        series = wide.dropna(subset=[parquet_name]).set_index("交易日期")[parquet_name]
        if len(series) < 2 or pd.Timestamp(series.index[-1]).date() != run_date:
            latest = series.index[-1] if len(series) else pd.NaT
            raise RuntimeError(
                f"“{parquet_name}”未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
            )
        details[group] = {
            "latestValue": float(series.iloc[-1]),
            "previousValue": float(series.iloc[-2]),
            "dailyChangePctPoint": float(series.iloc[-1] - series.iloc[-2]),
        }
    return wide, {
        "parquet": str(INDEX_PARQUET.relative_to(WORKSPACE)),
        "startDate": f"{VALUATION_START_DATE:%Y-%m-%d}",
        "runDate": f"{run_date:%Y-%m-%d}",
        "readRule": (
            f"直接读取指数 Parquet 中既有{classification_label}多因子修正拟合溢价率，"
            "不重新计算"
        ),
        "groupDetails": details,
    }


def fetch_balance_group_valuation_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """读取分余额多因子修正拟合溢价率。"""
    return fetch_precalculated_classification_valuation_series(
        run_date, BALANCE_GROUP_SPECS, "分余额"
    )


def fetch_market_cap_group_valuation_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """读取分正股市值多因子修正拟合溢价率。"""
    return fetch_precalculated_classification_valuation_series(
        run_date, MARKET_CAP_GROUP_SPECS, "分正股市值"
    )


def fetch_sector_group_valuation_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """读取科技、金融、制造、消费和周期板块拟合溢价率。"""
    return fetch_precalculated_classification_valuation_series(
        run_date, SECTOR_GROUP_SPECS, "分板块"
    )


def fetch_sector_mean_metrics(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从个券 Parquet 计算五大板块的四项日度截面算术均值。"""
    end_month = f"{run_date:%Y%m}"
    parquet_paths = sorted(
        path
        for path in CB_PARQUET_ROOT.glob("20*/20*.parquet")
        if "201901" <= path.stem <= end_month
    )
    if not parquet_paths:
        raise FileNotFoundError("未找到2019年以来的月度转债 Parquet")
    if not CB_MASTER_PARQUET.is_file():
        raise FileNotFoundError(f"未找到转债总表 Parquet：{CB_MASTER_PARQUET}")

    master = pd.read_parquet(
        CB_MASTER_PARQUET, columns=["转债代码", "申万行业"]
    ).copy()
    industry_to_sector = {
        industry: sector
        for sector, industries in SECTOR_INDUSTRIES.items()
        for industry in industries
    }
    master["转债代码"] = master["转债代码"].astype(str).str.strip()
    master["板块"] = master["申万行业"].astype(str).str.strip().map(industry_to_sector)
    code_to_sector = (
        master.dropna(subset=["板块"])
        .drop_duplicates("转债代码", keep="last")
        .set_index("转债代码")["板块"]
    )

    metric_names = [metric for metric, _, _ in SECTOR_MEAN_METRICS]
    required_columns = [
        "转债代码",
        "交易日期",
        "交易状态",
        *metric_names,
    ]
    monthly_results: list[pd.DataFrame] = []
    for parquet_path in parquet_paths:
        frame = pd.read_parquet(parquet_path, columns=required_columns).copy()
        frame["交易日期"] = pd.to_datetime(
            frame["交易日期"], errors="coerce"
        ).dt.normalize()
        frame["转债代码"] = frame["转债代码"].astype(str).str.strip()
        frame["板块"] = frame["转债代码"].map(code_to_sector)
        frame = frame.loc[
            frame["交易日期"].between(
                pd.Timestamp(VALUATION_START_DATE),
                pd.Timestamp(run_date),
                inclusive="both",
            )
            & frame["交易状态"].astype(str).str.strip().eq("交易")
            & frame["板块"].notna()
        ].copy()
        if frame.empty:
            continue
        for metric in metric_names:
            frame[metric] = pd.to_numeric(frame[metric], errors="coerce")
            frame.loc[~np.isfinite(frame[metric]), metric] = np.nan
        frame.loc[frame["收盘价"].le(0), "收盘价"] = np.nan
        frame.loc[frame["平价"].le(0), "平价"] = np.nan

        grouped = frame.groupby(["交易日期", "板块"], observed=True)[
            metric_names
        ].mean()
        metric_frames = []
        for metric in metric_names:
            wide = grouped[metric].unstack("板块").reindex(columns=SECTOR_ORDER)
            wide = wide.rename(columns={sector: f"{metric}_{sector}" for sector in SECTOR_ORDER})
            metric_frames.append(wide)
        monthly_results.append(pd.concat(metric_frames, axis=1).reset_index())

    if not monthly_results:
        raise RuntimeError("未计算出板块分类日度均值")
    result = (
        pd.concat(monthly_results, ignore_index=True)
        .drop_duplicates("交易日期", keep="last")
        .sort_values("交易日期")
        .reset_index(drop=True)
    )
    if result.empty or pd.Timestamp(result["交易日期"].iloc[-1]).date() != run_date:
        latest = result["交易日期"].max() if not result.empty else pd.NaT
        raise RuntimeError(
            f"板块分类日度均值未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
        )
    latest = result.iloc[-1]
    latest_values = {
        metric: {
            sector: float(latest[f"{metric}_{sector}"])
            for sector in SECTOR_ORDER
            if pd.notna(latest[f"{metric}_{sector}"])
        }
        for metric in metric_names
    }
    source: dict[str, object] = {
        "parquetRoot": str(CB_PARQUET_ROOT.relative_to(WORKSPACE)),
        "masterParquet": str(CB_MASTER_PARQUET.relative_to(WORKSPACE)),
        "startDate": f"{VALUATION_START_DATE:%Y-%m-%d}",
        "runDate": f"{run_date:%Y-%m-%d}",
        "sectorRule": "按总表申万一级行业映射为科技、金融、制造、消费、周期五大板块",
        "sampleRule": "交易状态=交易；收盘价和平价要求大于0；各指标按交易日和板块取个券截面算术均值",
        "points": len(result),
        "latestValues": latest_values,
    }
    return result, source


def aggregate_industry_history_metrics(
    panel: pd.DataFrame,
    master: pd.DataFrame,
    *,
    start_date: pd.Timestamp,
    run_date: pd.Timestamp,
) -> pd.DataFrame:
    """计算申万一级行业四项日度截面均值，并保留完整行业日期网格。"""
    required_panel = {
        "转债代码",
        "交易日期",
        "交易状态",
        *INDUSTRY_HISTORY_METRICS,
    }
    required_master = {"转债代码", "上市日期", "最后交易日", "申万行业"}
    if not required_panel.issubset(panel.columns):
        raise RuntimeError(
            "行业历史序列缺少个券字段："
            f"{sorted(required_panel - set(panel.columns))}"
        )
    if not required_master.issubset(master.columns):
        raise RuntimeError(
            "行业历史序列缺少总表字段："
            f"{sorted(required_master - set(master.columns))}"
        )

    history = panel.loc[:, list(required_panel)].copy()
    security_master = master.loc[:, list(required_master)].copy()
    history["转债代码"] = history["转债代码"].astype(str).str.strip()
    security_master["转债代码"] = security_master["转债代码"].astype(str).str.strip()
    history["交易日期"] = pd.to_datetime(
        history["交易日期"], errors="coerce"
    ).dt.normalize()
    for column in INDUSTRY_HISTORY_METRICS:
        history[column] = pd.to_numeric(history[column], errors="coerce")
        history.loc[~np.isfinite(history[column]), column] = np.nan
    for column in ("上市日期", "最后交易日"):
        security_master[column] = pd.to_datetime(
            security_master[column], errors="coerce"
        ).dt.normalize()
    security_master["申万行业"] = security_master["申万行业"].astype(str).str.strip()
    security_master = security_master.drop_duplicates("转债代码", keep="last")

    history = history.merge(
        security_master,
        on="转债代码",
        how="left",
        validate="many_to_one",
    )
    lifecycle_valid = (
        (history["上市日期"].isna() | history["交易日期"].ge(history["上市日期"]))
        & (
            history["最后交易日"].isna()
            | history["交易日期"].le(history["最后交易日"])
        )
    )
    extreme = history["收盘价"].gt(INDUSTRY_EXTREME_PRICE_THRESHOLD) & history[
        "转股溢价率"
    ].gt(INDUSTRY_EXTREME_PREMIUM_THRESHOLD)
    sample = history.loc[
        history["交易日期"].between(start_date, run_date, inclusive="both")
        & history["交易状态"].astype(str).str.strip().eq("交易")
        & lifecycle_valid
        & ~extreme
        & history["收盘价"].notna()
        & history["申万行业"].isin(INDUSTRY_HISTORY_ORDER),
        ["交易日期", "申万行业", *INDUSTRY_HISTORY_METRICS],
    ].copy()
    if sample.empty:
        raise RuntimeError("未计算出申万行业四项日度均值")

    trading_dates = pd.DatetimeIndex(sorted(sample["交易日期"].dropna().unique()))
    complete_index = pd.MultiIndex.from_product(
        [INDUSTRY_HISTORY_ORDER, trading_dates], names=["行业", "日期"]
    )
    result = (
        sample.groupby(["申万行业", "交易日期"], observed=True)[
            list(INDUSTRY_HISTORY_METRICS)
        ]
        .mean()
        .rename_axis(index=["行业", "日期"])
        .reindex(complete_index)
        .reset_index()
    )
    return result


def fetch_industry_history_metrics(run_date: date) -> pd.DataFrame:
    """读取2019年以来个券数据并计算30个申万行业的四项历史均值。"""
    end_month = f"{run_date:%Y%m}"
    parquet_paths = sorted(
        path
        for path in CB_PARQUET_ROOT.glob("20*/20*.parquet")
        if "201901" <= path.stem <= end_month
    )
    if not parquet_paths:
        raise FileNotFoundError("未找到2019年以来的月度转债 Parquet")
    if not CB_MASTER_PARQUET.is_file():
        raise FileNotFoundError(f"未找到转债总表 Parquet：{CB_MASTER_PARQUET}")

    panel = pd.concat(
        [
            pd.read_parquet(
                parquet_path,
                columns=[
                    "转债代码",
                    "交易日期",
                    "交易状态",
                    *INDUSTRY_HISTORY_METRICS,
                ],
            )
            for parquet_path in parquet_paths
        ],
        ignore_index=True,
    )
    master = pd.read_parquet(
        CB_MASTER_PARQUET,
        columns=["转债代码", "上市日期", "最后交易日", "申万行业"],
    )
    result = aggregate_industry_history_metrics(
        panel,
        master,
        start_date=pd.Timestamp(VALUATION_START_DATE),
        run_date=pd.Timestamp(run_date),
    )
    latest = result.loc[result[list(INDUSTRY_HISTORY_METRICS)].notna().any(axis=1), "日期"].max()
    if pd.isna(latest) or pd.Timestamp(latest).date() != run_date:
        raise RuntimeError(
            f"行业四项历史均值未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
        )
    return result


def aggregate_equity_bond_weighted_series(
    data: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """按平价底价溢价率分三类，计算余额加权转股溢价率与收盘价。"""
    required_columns = {
        "交易日期",
        "交易状态",
        "余额",
        "平价底价溢价率",
        "转股溢价率",
        "收盘价",
    }
    if not required_columns.issubset(data.columns):
        raise RuntimeError(
            "股债性分类余额加权数据缺少字段："
            f"{sorted(required_columns - set(data.columns))}"
        )

    base = data.loc[:, list(required_columns)].copy()
    base["交易日期"] = pd.to_datetime(base["交易日期"], errors="coerce").dt.normalize()
    for column in ("余额", "平价底价溢价率", "转股溢价率", "收盘价"):
        base[column] = pd.to_numeric(base[column], errors="coerce")
        base.loc[~np.isfinite(base[column]), column] = np.nan
    base = base.loc[
        base["交易日期"].notna()
        & base["交易状态"].astype(str).str.strip().eq("交易")
        & base["余额"].gt(0)
        & base["平价底价溢价率"].notna()
    ].copy()
    base["股债性分类"] = np.select(
        (
            base["平价底价溢价率"].gt(20),
            base["平价底价溢价率"].gt(-20)
            & base["平价底价溢价率"].lt(20),
            base["平价底价溢价率"].lt(-20),
        ),
        EQUITY_BOND_TYPE_GROUPS,
        default="",
    )
    base = base.loc[base["股债性分类"].isin(EQUITY_BOND_TYPE_GROUPS)].copy()
    if base.empty:
        raise RuntimeError("股债性分类余额加权序列无有效交易样本")

    result = pd.DataFrame({"交易日期": sorted(base["交易日期"].unique())})
    for metric in ("转股溢价率", "收盘价"):
        weighted = base.dropna(subset=[metric]).copy()
        weighted["加权值"] = weighted[metric] * weighted["余额"]
        grouped = weighted.groupby(["交易日期", "股债性分类"], observed=True).agg(
            加权值=("加权值", "sum"),
            有效余额=("余额", "sum"),
        )
        average = (grouped["加权值"] / grouped["有效余额"]).unstack("股债性分类")
        average = average.reindex(columns=EQUITY_BOND_TYPE_GROUPS)
        average = average.rename(
            columns={group: f"{group}_{metric}" for group in EQUITY_BOND_TYPE_GROUPS}
        ).reset_index()
        result = result.merge(average, on="交易日期", how="left")

    source: dict[str, object] = {
        "sampleRule": (
            "交易状态=交易、余额>0；按平价底价溢价率划分偏股型>20、"
            "平衡型(-20,20)、偏债型<-20；各组按余额加权"
        )
    }
    for metric, source_key in (("转股溢价率", "premium"), ("收盘价", "price")):
        metric_details: dict[str, dict[str, object]] = {}
        for group in EQUITY_BOND_TYPE_GROUPS:
            column = f"{group}_{metric}"
            series = result.loc[:, ["交易日期", column]].dropna().sort_values("交易日期")
            if len(series) < 2:
                raise RuntimeError(f"{group}余额加权{metric}有效历史记录不足2条")
            latest = series.iloc[-1]
            previous = series.iloc[-2]
            latest_value = float(latest[column])
            previous_value = float(previous[column])
            metric_details[group] = {
                "latestDate": f"{latest['交易日期']:%Y-%m-%d}",
                "latestValue": latest_value,
                "previousDate": f"{previous['交易日期']:%Y-%m-%d}",
                "previousValue": previous_value,
                "dailyChangePctPoint": latest_value - previous_value,
                "dailyChangePct": (
                    (latest_value / previous_value - 1.0) * 100
                    if previous_value != 0
                    else np.nan
                ),
            }
        source[source_key] = metric_details
    return result.sort_values("交易日期").reset_index(drop=True), source


def fetch_equity_bond_weighted_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从月度个券 Parquet 读取股债性分类余额加权序列。"""
    end_month = f"{run_date:%Y%m}"
    parquet_paths = sorted(
        path
        for path in CB_PARQUET_ROOT.glob("20*/20*.parquet")
        if "201901" <= path.stem <= end_month
    )
    if not parquet_paths:
        raise FileNotFoundError("未找到2019年以来的月度转债 Parquet")

    columns = [
        "交易日期",
        "交易状态",
        "余额",
        "平价底价溢价率",
        "转股溢价率",
        "收盘价",
    ]
    frames: list[pd.DataFrame] = []
    for parquet_path in parquet_paths:
        frame = pd.read_parquet(parquet_path, columns=columns)
        frame["交易日期"] = pd.to_datetime(frame["交易日期"], errors="coerce").dt.normalize()
        frame = frame.loc[frame["交易日期"].le(pd.Timestamp(run_date))].copy()
        if not frame.empty:
            frames.append(frame)
    if not frames:
        raise RuntimeError("月度转债 Parquet 中没有股债性分类数据")

    result, source = aggregate_equity_bond_weighted_series(
        pd.concat(frames, ignore_index=True)
    )
    if pd.Timestamp(result["交易日期"].iloc[-1]).date() != run_date:
        latest = pd.Timestamp(result["交易日期"].iloc[-1]).date()
        raise RuntimeError(
            f"股债性分类余额加权序列未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
        )
    for metric_key in ("premium", "price"):
        for group, details in source[metric_key].items():
            if details["latestDate"] != f"{run_date:%Y-%m-%d}":
                raise RuntimeError(
                    f"{group}余额加权{metric_key}未更新至 {run_date:%Y-%m-%d}，"
                    f"当前最新日期：{details['latestDate']}"
                )
    source.update(
        {
            "parquetRoot": str(CB_PARQUET_ROOT.relative_to(WORKSPACE)),
            "startDate": f"{result['交易日期'].min():%Y-%m-%d}",
            "runDate": f"{run_date:%Y-%m-%d}",
        }
    )
    return result, source


def build_equity_bond_weighted_titles(source: dict[str, object]) -> dict[str, str]:
    """生成股债性余额加权转股溢价率与均价标题。"""
    premium = source["premium"]
    price = source["price"]
    return {
        "premium": (
            "股债性分类转股溢价率：\n"
            f"偏股型{float(premium['偏股型']['latestValue']):.2f}%，"
            f"{float(premium['偏股型']['dailyChangePctPoint']):+.2f}pct；"
            f"偏债型{float(premium['偏债型']['latestValue']):.2f}%，"
            f"{float(premium['偏债型']['dailyChangePctPoint']):+.2f}pct"
        ),
        "price": (
            "股债性分类均价：\n"
            f"偏股型{float(price['偏股型']['latestValue']):.2f}，"
            f"{float(price['偏股型']['dailyChangePct']):+.2f}%；"
            f"偏债型{float(price['偏债型']['latestValue']):.2f}，"
            f"{float(price['偏债型']['dailyChangePct']):+.2f}%"
        ),
    }


def aggregate_price_parity_series(data: pd.DataFrame) -> pd.DataFrame:
    """按交易日聚合余额加权平价、余额加权收盘价与收盘价中位数。"""
    required = {"交易日期", "交易状态", "余额", "平价", "收盘价"}
    if not required.issubset(data.columns):
        raise RuntimeError(
            f"价格与平价聚合输入缺少字段：{sorted(required - set(data.columns))}"
        )

    frame = data[list(required)].copy()
    frame["交易日期"] = pd.to_datetime(
        frame["交易日期"], errors="coerce"
    ).dt.normalize()
    frame = frame.loc[
        frame["交易日期"].notna()
        & frame["交易状态"].astype(str).str.strip().eq("交易")
    ].copy()
    for column in ("余额", "平价", "收盘价"):
        frame[column] = pd.to_numeric(frame[column], errors="coerce")
    if frame.empty:
        return pd.DataFrame(
            columns=[
                "交易日期",
                "余额加权平价",
                "余额加权收盘价",
                "平价样本数",
                "价格样本数",
                "有效余额",
                "收盘价中位数",
            ]
        )

    weighted = frame.loc[frame["余额"].gt(0)].copy()
    parity_weighted = weighted.dropna(subset=["平价"]).copy()
    parity_weighted["平价乘余额"] = (
        parity_weighted["平价"] * parity_weighted["余额"]
    )
    parity_daily = parity_weighted.groupby("交易日期", as_index=False).agg(
        平价乘余额=("平价乘余额", "sum"),
        有效余额=("余额", "sum"),
        平价样本数=("平价", "size"),
    )
    parity_daily["余额加权平价"] = parity_daily["平价乘余额"].div(
        parity_daily["有效余额"].replace(0, np.nan)
    )
    price_weighted = weighted.dropna(subset=["收盘价"]).copy()
    price_weighted["收盘价乘余额"] = (
        price_weighted["收盘价"] * price_weighted["余额"]
    )
    price_daily = price_weighted.groupby("交易日期", as_index=False).agg(
        收盘价乘余额=("收盘价乘余额", "sum"),
        价格有效余额=("余额", "sum"),
    )
    price_daily["余额加权收盘价"] = price_daily["收盘价乘余额"].div(
        price_daily["价格有效余额"].replace(0, np.nan)
    )
    median_daily = (
        frame.dropna(subset=["收盘价"])
        .groupby("交易日期", as_index=False)
        .agg(收盘价中位数=("收盘价", "median"), 价格样本数=("收盘价", "size"))
    )
    return (
        parity_daily[
            [
                "交易日期",
                "余额加权平价",
                "平价样本数",
                "有效余额",
            ]
        ]
        .merge(
            price_daily[["交易日期", "余额加权收盘价"]],
            on="交易日期",
            how="outer",
        )
        .merge(median_daily, on="交易日期", how="outer")
        .sort_values("交易日期")
        .reset_index(drop=True)
    )


def fetch_price_parity_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从月度个券 Parquet 聚合余额加权平价、收盘价与价格中位数。"""
    end_month = f"{run_date:%Y%m}"
    parquet_paths = sorted(
        path
        for path in CB_PARQUET_ROOT.glob("20*/20*.parquet")
        if "201901" <= path.stem <= end_month
    )
    if not parquet_paths:
        raise FileNotFoundError("未找到2019年以来的月度转债 Parquet")

    raw_frames: list[pd.DataFrame] = []
    required = {"交易日期", "交易状态", "余额", "平价", "收盘价"}
    for parquet_path in parquet_paths:
        frame = pd.read_parquet(parquet_path, columns=list(required))
        if not required.issubset(frame.columns):
            raise RuntimeError(
                f"月度转债 Parquet 字段异常：{parquet_path}，"
                f"缺少{sorted(required - set(frame.columns))}"
            )
        frame["交易日期"] = pd.to_datetime(
            frame["交易日期"], errors="coerce"
        ).dt.normalize()
        frame = frame.loc[
            frame["交易日期"].le(pd.Timestamp(run_date))
            & frame["交易状态"].astype(str).str.strip().eq("交易")
        ].copy()
        if frame.empty:
            continue
        raw_frames.append(frame)

    if not raw_frames:
        raise RuntimeError("月度转债 Parquet 中未聚合出有效价格与平价序列")
    result = aggregate_price_parity_series(
        pd.concat(raw_frames, ignore_index=True)
    )
    result = result.dropna(
        subset=["余额加权平价", "余额加权收盘价", "收盘价中位数"]
    ).reset_index(drop=True)
    if len(result) < 2 or result["交易日期"].iloc[-1].date() != run_date:
        latest = result["交易日期"].max() if not result.empty else pd.NaT
        raise RuntimeError(
            f"价格与平价序列未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
        )

    latest = result.iloc[-1]
    previous = result.iloc[-2]
    parity_change = (
        float(latest["余额加权平价"]) / float(previous["余额加权平价"]) - 1
    ) * 100
    median_change = (
        float(latest["收盘价中位数"]) / float(previous["收盘价中位数"]) - 1
    ) * 100
    weighted_close_change = (
        float(latest["余额加权收盘价"])
        / float(previous["余额加权收盘价"])
        - 1
    ) * 100
    median_percentile = float(
        result["收盘价中位数"].le(float(latest["收盘价中位数"])).mean() * 100
    )
    source = {
        "parquetRoot": str(CB_PARQUET_ROOT.relative_to(WORKSPACE)),
        "startDate": f"{result['交易日期'].min():%Y-%m-%d}",
        "latestDate": f"{latest['交易日期']:%Y-%m-%d}",
        "previousDate": f"{previous['交易日期']:%Y-%m-%d}",
        "latestParity": float(latest["余额加权平价"]),
        "parityDailyChangePct": parity_change,
        "latestWeightedClose": float(latest["余额加权收盘价"]),
        "weightedCloseDailyChangePct": weighted_close_change,
        "latestMedianPrice": float(latest["收盘价中位数"]),
        "medianPriceDailyChangePct": median_change,
        "medianPricePercentileSince2019": median_percentile,
        "sampleRule": (
            "交易状态=交易；平价与整体收盘价按有效余额加权，"
            "价格中位数取有效收盘价样本截面中位数"
        ),
    }
    return result, source


def fetch_close_price_distribution_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从2019年以来月度个券 Parquet 重算收盘价八档分布占比。"""
    end_month = f"{run_date:%Y%m}"
    parquet_paths = sorted(
        path
        for path in CB_PARQUET_ROOT.glob("20*/20*.parquet")
        if "201901" <= path.stem <= end_month
    )
    if not parquet_paths:
        raise FileNotFoundError("未找到2019年以来的月度转债 Parquet")

    history_frames: list[pd.DataFrame] = []
    required = {"转债代码", "交易日期", "交易状态", "收盘价"}
    for parquet_path in parquet_paths:
        frame = pd.read_parquet(parquet_path, columns=list(required))
        if not required.issubset(frame.columns):
            raise RuntimeError(
                f"月度转债 Parquet 字段异常：{parquet_path}，"
                f"缺少{sorted(required - set(frame.columns))}"
            )
        frame["交易日期"] = pd.to_datetime(
            frame["交易日期"], errors="coerce"
        ).dt.normalize()
        frame["收盘价"] = pd.to_numeric(frame["收盘价"], errors="coerce")
        frame = frame.loc[
            frame["交易日期"].between(
                pd.Timestamp(VALUATION_START_DATE),
                pd.Timestamp(run_date),
                inclusive="both",
            )
            & frame["交易状态"].astype(str).str.strip().eq("交易")
            & frame["收盘价"].notna()
            & frame["收盘价"].gt(0),
            ["转债代码", "交易日期", "收盘价"],
        ].drop_duplicates(["转债代码", "交易日期"], keep="last")
        if not frame.empty:
            history_frames.append(frame)

    if not history_frames:
        raise RuntimeError("2019年以来月度转债 Parquet 中没有有效收盘价")
    history = pd.concat(history_frames, ignore_index=True)
    history["价格区间"] = pd.cut(
        history["收盘价"],
        bins=CLOSE_PRICE_DISTRIBUTION_BINS,
        labels=CLOSE_PRICE_DISTRIBUTION_LABELS,
        right=True,
        include_lowest=True,
    )
    counts = pd.crosstab(history["交易日期"], history["价格区间"], dropna=False)
    counts = counts.reindex(columns=CLOSE_PRICE_DISTRIBUTION_LABELS, fill_value=0)
    sample_count = counts.sum(axis=1)
    shares = counts.div(sample_count.replace(0, np.nan), axis=0).mul(100.0)
    result = shares.reset_index().sort_values("交易日期").reset_index(drop=True)
    result["有效样本数"] = result["交易日期"].map(sample_count).astype(int)
    if len(result) < 2 or pd.Timestamp(result.iloc[-1]["交易日期"]).date() != run_date:
        latest = result["交易日期"].max() if not result.empty else pd.NaT
        raise RuntimeError(
            f"收盘价分布未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest}"
        )

    latest = result.iloc[-1]
    previous = result.iloc[-2]
    floor_latest = float(latest[CLOSE_PRICE_DISTRIBUTION_LABELS[0]])
    floor_previous = float(previous[CLOSE_PRICE_DISTRIBUTION_LABELS[0]])
    par_columns = list(CLOSE_PRICE_DISTRIBUTION_LABELS[:3])
    par_latest = float(latest[par_columns].sum())
    par_previous = float(previous[par_columns].sum())
    source: dict[str, object] = {
        "parquetRoot": str(CB_PARQUET_ROOT.relative_to(WORKSPACE)),
        "startDate": f"{result['交易日期'].min():%Y-%m-%d}",
        "runDate": f"{run_date:%Y-%m-%d}",
        "previousDate": f"{previous['交易日期']:%Y-%m-%d}",
        "latestBreakFloorPct": floor_latest,
        "breakFloorDailyChangePctPoint": floor_latest - floor_previous,
        "latestBreakParPct": par_latest,
        "breakParDailyChangePctPoint": par_latest - par_previous,
        "latestSampleCount": int(latest["有效样本数"]),
        "sampleRule": (
            "交易状态=交易且收盘价>0；破底=收盘价<=80；破面=收盘价<=100；"
            "八档采用右闭区间"
        ),
    }
    return result, source


def fetch_subnew_bond_series(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """从个券与总表 Parquet 计算次新券价格表现和平均转股溢价率。"""
    if not CB_MASTER_PARQUET.is_file():
        raise FileNotFoundError(f"未找到转债总表 Parquet：{CB_MASTER_PARQUET}")

    end_month = f"{run_date:%Y%m}"
    parquet_paths = sorted(
        path
        for path in CB_PARQUET_ROOT.glob("20*/20*.parquet")
        if path.stem <= end_month
    )
    if not parquet_paths:
        raise FileNotFoundError("未找到月度转债 Parquet")

    required = {"转债代码", "交易日期", "交易状态", "收盘价", "转股溢价率"}
    history_frames: list[pd.DataFrame] = []
    for parquet_path in parquet_paths:
        frame = pd.read_parquet(parquet_path, columns=list(required))
        if not required.issubset(frame.columns):
            raise RuntimeError(
                f"月度转债 Parquet 字段异常：{parquet_path}，"
                f"缺少{sorted(required - set(frame.columns))}"
            )
        history_frames.append(frame)

    history = pd.concat(history_frames, ignore_index=True)
    history["交易日期"] = pd.to_datetime(
        history["交易日期"], errors="coerce"
    ).dt.normalize()
    history["转债代码"] = history["转债代码"].astype(str).str.strip()
    for column in ("收盘价", "转股溢价率"):
        history[column] = pd.to_numeric(history[column], errors="coerce")
    history = history.loc[
        history["交易日期"].le(pd.Timestamp(run_date))
        & history["交易状态"].astype(str).str.strip().eq("交易")
    ].copy()

    master = pd.read_parquet(
        CB_MASTER_PARQUET,
        columns=["转债代码", "上市日期", "转股期起始日"],
    )
    master["转债代码"] = master["转债代码"].astype(str).str.strip()
    master["上市日期"] = pd.to_datetime(
        master["上市日期"], errors="coerce"
    ).dt.normalize()
    master["转股期起始日"] = pd.to_datetime(
        master["转股期起始日"], errors="coerce"
    ).dt.normalize()
    history = history.merge(master, on="转债代码", how="left", validate="many_to_one")

    listing_candidates = history.loc[
        history["上市日期"].notna()
        & history["交易日期"].ge(history["上市日期"])
        & history["收盘价"].notna()
        & history["收盘价"].gt(0),
        ["转债代码", "交易日期", "收盘价"],
    ].sort_values(["转债代码", "交易日期"])
    listing_price = (
        listing_candidates.drop_duplicates("转债代码", keep="first")
        .set_index("转债代码")["收盘价"]
    )
    history["上市首个有效收盘价"] = history["转债代码"].map(listing_price)

    sample = history.loc[
        history["交易日期"].ge(pd.Timestamp(VALUATION_START_DATE))
        & history["上市日期"].notna()
        & history["转股期起始日"].notna()
        & history["交易日期"].ge(history["上市日期"])
        & history["交易日期"].lt(history["转股期起始日"])
    ].copy()
    sample["相对上市涨跌幅均值"] = (
        sample["收盘价"] / sample["上市首个有效收盘价"] - 1
    ) * 100

    listing_return = (
        sample.dropna(subset=["相对上市涨跌幅均值"])
        .groupby("交易日期", as_index=False)
        .agg(
            次新券相对上市涨跌幅均值=("相对上市涨跌幅均值", "mean"),
            价格样本数=("相对上市涨跌幅均值", "size"),
        )
    )
    premium = (
        sample.dropna(subset=["转股溢价率"])
        .groupby("交易日期", as_index=False)
        .agg(
            次新券平均转股溢价率=("转股溢价率", "mean"),
            溢价率样本数=("转股溢价率", "size"),
        )
    )
    result = (
        listing_return.merge(premium, on="交易日期", how="outer")
        .sort_values("交易日期")
        .drop_duplicates("交易日期", keep="last")
        .dropna(
            subset=["次新券相对上市涨跌幅均值", "次新券平均转股溢价率"]
        )
        .reset_index(drop=True)
    )
    if len(result) < 2 or result["交易日期"].iloc[-1].date() != run_date:
        latest_date = result["交易日期"].max() if not result.empty else pd.NaT
        raise RuntimeError(
            f"次新券指标未更新至 {run_date:%Y-%m-%d}，当前最新日期：{latest_date}"
        )

    latest = result.iloc[-1]
    previous = result.iloc[-2]
    source: dict[str, object] = {
        "parquetRoot": str(CB_PARQUET_ROOT.relative_to(WORKSPACE)),
        "masterParquet": str(CB_MASTER_PARQUET.relative_to(WORKSPACE)),
        "startDate": f"{result['交易日期'].min():%Y-%m-%d}",
        "latestDate": f"{latest['交易日期']:%Y-%m-%d}",
        "previousDate": f"{previous['交易日期']:%Y-%m-%d}",
        "latestListingReturnMeanPct": float(
            latest["次新券相对上市涨跌幅均值"]
        ),
        "listingReturnDailyChangePctPoint": float(
            latest["次新券相对上市涨跌幅均值"]
            - previous["次新券相对上市涨跌幅均值"]
        ),
        "latestPremiumMeanPct": float(latest["次新券平均转股溢价率"]),
        "premiumDailyChangePctPoint": float(
            latest["次新券平均转股溢价率"]
            - previous["次新券平均转股溢价率"]
        ),
        "latestListingReturnSampleCount": int(latest["价格样本数"]),
        "latestPremiumSampleCount": int(latest["溢价率样本数"]),
        "sampleRule": (
            "交易状态=交易；上市日期<=交易日<转股期起始日；"
            "上市价格取上市日起首个有效收盘价；两项均为截面算术均值"
        ),
    }
    return result, source


def fetch_intraday_valuation(
    run_date: date,
) -> tuple[pd.DataFrame, dict[str, object]]:
    """读取日内估值更新生成的盘中百元平价拟合溢价率结果。"""
    daily_runs = WORKSPACE / "runs" / "daily"
    candidates = sorted(
        daily_runs.rglob("*百元平价溢价率拟合结果.xlsx"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    valid: list[tuple[int, float, Path, pd.DataFrame]] = []
    for path in candidates:
        try:
            frame = pd.read_excel(path, sheet_name=INTRADAY_VALUATION_SHEET)
        except Exception:
            continue
        if not {"日期", "转股溢价率"}.issubset(frame.columns):
            continue
        frame = frame[["日期", "转股溢价率"]].copy()
        frame["日期"] = pd.to_datetime(frame["日期"], errors="coerce")
        frame["转股溢价率"] = pd.to_numeric(frame["转股溢价率"], errors="coerce")
        frame = frame.dropna().loc[lambda x: x["日期"].dt.date.eq(run_date)]
        frame = frame.drop_duplicates("日期", keep="last").sort_values("日期")
        if len(frame) < 6 or frame["日期"].dt.time.nunique() < 6:
            continue
        score = 2 if "日内估值数据更新" in str(path.parent) else 0
        score += 1 if path.stem.startswith(f"{run_date:%m%d}") else 0
        valid.append((score, path.stat().st_mtime, path, frame))
    if not valid:
        raise FileNotFoundError(
            f"未找到 {run_date:%Y-%m-%d} 的日内百元平价拟合溢价率结果"
        )
    _, _, source_path, result = max(valid, key=lambda item: (item[0], item[1]))
    source = {
        "workbook": str(source_path.relative_to(WORKSPACE)),
        "sheet": INTRADAY_VALUATION_SHEET,
        "points": len(result),
        "startTime": f"{result['日期'].min():%Y-%m-%d %H:%M}",
        "endTime": f"{result['日期'].max():%Y-%m-%d %H:%M}",
    }
    return result.reset_index(drop=True), source


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
    """跨度严格超过三个公历月时省略日期中的日。"""
    minimum_date = pd.Timestamp(date_values.min())
    maximum_date = pd.Timestamp(date_values.max())
    if maximum_date > minimum_date + pd.DateOffset(months=3):
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


def add_chart_panel_title(
    fig: plt.Figure,
    title: str,
    band_height: float | None = None,
) -> None:
    """按日报版式增加浅灰标题栏和完整图表外框。"""
    required_height = (
        DOUBLE_LINE_TITLE_BAND_HEIGHT
        if "\n" in title
        else SINGLE_LINE_TITLE_BAND_HEIGHT
    )
    if band_height is None:
        band_height = required_height
    elif band_height < required_height:
        raise ValueError(
            f"标题栏高度 {band_height:.2f} 小于标题所需高度 {required_height:.2f}"
        )
    band_bottom = 1.0 - band_height
    fig.add_artist(
        plt.Rectangle(
            (0, band_bottom),
            1,
            band_height,
            transform=fig.transFigure,
            facecolor="#D9E2F3",
            edgecolor="none",
            linewidth=0,
            zorder=19,
            clip_on=False,
        )
    )
    fig.add_artist(
        plt.Line2D(
            (0.002, 0.998),
            (band_bottom, band_bottom),
            transform=fig.transFigure,
            color="#7F7F7F",
            linewidth=0.5,
            zorder=20,
            clip_on=False,
        )
    )
    fig.text(
        0.5,
        band_bottom + band_height / 2,
        title,
        ha="center",
        va="center",
        fontproperties=get_title_font(TITLE_FONT_SIZE),
        fontsize=TITLE_FONT_SIZE,
        fontweight="bold",
        color=RED,
        zorder=21,
    )


def prepare_margin_balance_plot(
    data: pd.DataFrame,
) -> tuple[pd.DataFrame, str]:
    """剔除相对上一个正常点跳变超过 10% 的未完整更新数据。"""
    balance_column = "沪深两市融资融券余额_亿元"
    valid = (
        data.dropna(subset=["交易日期", balance_column])
        .sort_values("交易日期")
        .copy()
    )
    accepted_positions: list[int] = []
    previous_balance: float | None = None
    for position, raw_balance in enumerate(valid[balance_column]):
        balance = float(raw_balance)
        if previous_balance is None:
            accepted_positions.append(position)
            previous_balance = balance
            continue
        if previous_balance == 0:
            relative_change = 0.0 if balance == 0 else math.inf
        else:
            relative_change = abs(balance / previous_balance - 1.0)
        if relative_change <= MARGIN_BALANCE_MAX_DAILY_CHANGE_RATIO + 1e-12:
            accepted_positions.append(position)
            previous_balance = balance

    plot_data = valid.iloc[accepted_positions].reset_index(drop=True)
    if plot_data.empty:
        raise RuntimeError("两融余额无有效绘图数据")
    return plot_data, "沪深两市融资融券余额"


def plot_market_statistics(
    data: pd.DataFrame, output_path: Path, font: fm.FontProperties
) -> None:
    """按已确认的华创“单轴”模板固化两融余额折线图。"""
    if not CHART_FONT_PATH.exists():
        raise FileNotFoundError(f"未找到华文楷体字体：{CHART_FONT_PATH}")
    fm.fontManager.addfont(str(CHART_FONT_PATH))
    chart_font = fm.FontProperties(fname=str(CHART_FONT_PATH), size=7)
    plot_data, panel_title = prepare_margin_balance_plot(data)

    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    ax.plot(
        plot_data["交易日期"],
        plot_data["沪深两市融资融券余额_亿元"],
        color=RED,
        linewidth=1.0,
        marker=None,
        label="沪深两市融资融券余额（亿元）",
    )
    add_chart_panel_title(
        fig,
        panel_title,
    )

    ax.set_xlim(plot_data["交易日期"].min(), plot_data["交易日期"].max())
    value_max = float(plot_data["沪深两市融资融券余额_亿元"].max())
    y_step = 5000.0
    ax.set_ylim(0, max(y_step, math.ceil(value_max / y_step) * y_step))
    ax.yaxis.set_major_locator(mticker.MultipleLocator(y_step))
    ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.2f}"))
    ax.xaxis.set_major_locator(
        mdates.MonthLocator(bymonth=(1, 3, 5, 7, 9, 11), bymonthday=2)
    )
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(plot_data["交易日期"]))
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
    fig.subplots_adjust(left=0.15, right=0.955, top=0.88, bottom=0.39)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_main_money_flow(
    data: pd.DataFrame,
    source: dict[str, object],
    output_path: Path,
) -> None:
    """按华创单轴柱状图样式绘制沪深两市主力净流入。"""
    if not CHART_FONT_PATH.exists():
        raise FileNotFoundError(f"未找到华文楷体字体：{CHART_FONT_PATH}")
    fm.fontManager.addfont(str(CHART_FONT_PATH))
    chart_font = fm.FontProperties(fname=str(CHART_FONT_PATH), size=7)

    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    values = data["主力净流入_亿元"]
    bar_colors = np.where(values >= 0, RED, BLUE)
    ax.bar(
        data["交易日期"],
        values,
        width=0.8,
        color=bar_colors,
        edgecolor="none",
        zorder=3,
    )
    ax.axhline(0, color="#7F7F7F", linewidth=0.7, linestyle="--", zorder=2)

    latest_value = float(source["latestValue"])
    direction = "净流入" if latest_value >= 0 else "净流出"
    add_chart_panel_title(
        fig,
        f"沪深两市主力{direction}{abs(latest_value):.2f}亿元",
    )

    data_min = float(values.min())
    data_max = float(values.max())
    value_span = max(data_max - data_min, 1.0)
    y_padding = value_span * 0.04
    y_lower = min(0.0, data_min - y_padding)
    y_upper = max(0.0, data_max + y_padding)
    ax.set_ylim(y_lower, y_upper)
    ax.yaxis.set_major_locator(
        mticker.MaxNLocator(
            nbins=7,
            steps=[1, 2, 2.5, 5, 10],
            min_n_ticks=5,
        )
    )
    ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.0f}"))
    ax.set_xlim(
        data["交易日期"].min() - pd.Timedelta(days=1),
        data["交易日期"].max() + pd.Timedelta(days=1),
    )
    ax.xaxis.set_major_locator(mdates.DayLocator(bymonthday=[1, 15]))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y%m%d"))
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
        handles=[
            Patch(facecolor=RED, edgecolor="none", label="主力净流入（亿元）"),
            Patch(facecolor=BLUE, edgecolor="none", label="主力净流出（亿元）"),
        ],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.50),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        handlelength=1.8,
        handletextpad=0.8,
        borderaxespad=0,
        ncol=2,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.15, right=0.955, top=0.88, bottom=0.39)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_cb_etf_share_flow(
    data: pd.DataFrame,
    source: dict[str, object],
    etf_name: str,
    output_path: Path,
) -> None:
    """绘制 ETF 总份额折线与净申赎红蓝柱状双轴图。"""
    chart_font = _valuation_chart_font()
    share_column = f"{etf_name}份额_亿份"
    flow_column = f"{etf_name}净申赎_亿份"
    chart_data = (
        data.loc[:, ["交易日期", share_column, flow_column]]
        .loc[lambda frame: frame["交易日期"].ge(pd.Timestamp("2024-01-01"))]
        .dropna(subset=[share_column])
        .sort_values("交易日期")
        .reset_index(drop=True)
    )
    if chart_data.empty:
        raise RuntimeError(f"{etf_name}份额序列为空，无法绘图")

    fig, ax_left = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    ax_right = ax_left.twinx()
    line_share, = ax_left.plot(
        chart_data["交易日期"],
        chart_data[share_column],
        color="#404040",
        linewidth=1.0,
        marker=None,
        label=f"{etf_name}份额（亿份）",
        zorder=4,
    )
    flow_values = chart_data[flow_column].fillna(0.0)
    flow_colors = np.where(flow_values >= 0, RED, BLUE)
    ax_right.bar(
        chart_data["交易日期"],
        flow_values,
        width=1.0,
        color=flow_colors,
        edgecolor="none",
        alpha=0.88,
        zorder=2,
    )
    ax_right.axhline(0, color="#7F7F7F", linewidth=0.65, linestyle="--", zorder=1)

    detail = source["funds"][etf_name]
    latest_share = float(detail["latestShareYi"])
    latest_flow_value = detail["latestNetSubscriptionYi"]
    latest_flow = 0.0 if latest_flow_value is None else float(latest_flow_value)
    latest_flow_text = (
        f"{latest_flow * 10000:+,.0f}万份"
        if abs(latest_flow) < 1.0
        else f"{latest_flow:+,.2f}亿份"
    )
    add_chart_panel_title(
        fig,
        f"{etf_name}：\n份额{latest_share:.2f}亿份，净申赎{latest_flow_text}",
        band_height=DOUBLE_LINE_TITLE_BAND_HEIGHT,
    )

    share_min = float(chart_data[share_column].min())
    share_max = float(chart_data[share_column].max())
    share_span = max(share_max - share_min, max(abs(share_max), 1.0) * 0.08)
    ax_left.set_ylim(
        max(0.0, share_min - share_span * 0.06),
        share_max + share_span * 0.08,
    )
    ax_left.yaxis.set_major_locator(
        mticker.MaxNLocator(nbins=6, steps=[1, 2, 2.5, 5, 10], min_n_ticks=4)
    )
    ax_left.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.1f}"))

    valid_flow = chart_data[flow_column].dropna()
    if valid_flow.empty:
        flow_lower, flow_upper = -1.0, 1.0
    else:
        flow_min = float(valid_flow.min())
        flow_max = float(valid_flow.max())
        flow_span = max(flow_max - flow_min, 0.1)
        flow_lower = min(0.0, flow_min - flow_span * 0.05)
        flow_upper = max(0.0, flow_max + flow_span * 0.05)
        if math.isclose(flow_lower, flow_upper):
            flow_lower, flow_upper = flow_lower - 0.5, flow_upper + 0.5
    ax_right.set_ylim(flow_lower, flow_upper)
    ax_right.yaxis.set_major_locator(
        mticker.MaxNLocator(nbins=6, steps=[1, 2, 2.5, 5, 10], min_n_ticks=4)
    )
    ax_right.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.1f}"))

    ax_left.set_xlim(
        chart_data["交易日期"].min() - pd.Timedelta(days=10),
        chart_data["交易日期"].max() + pd.Timedelta(days=10),
    )
    ax_left.xaxis.set_major_locator(mdates.YearLocator())
    ax_left.xaxis.set_major_formatter(mdates.DateFormatter("%Y%m"))
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
    for spine in (
        ax_left.spines["left"],
        ax_left.spines["bottom"],
        ax_right.spines["right"],
    ):
        spine.set_color("black")
        spine.set_linewidth(0.7)

    legend = ax_left.legend(
        handles=[
            line_share,
            Patch(facecolor=RED, edgecolor="none", label="净申购（亿份）"),
            Patch(facecolor=BLUE, edgecolor="none", label="净赎回（亿份）"),
        ],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.42),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=3,
        handlelength=2.3,
        handletextpad=0.5,
        columnspacing=0.8,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.12, right=0.88, top=0.82, bottom=0.34)
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


def _valuation_chart_font() -> fm.FontProperties:
    if not CHART_FONT_PATH.exists():
        raise FileNotFoundError(f"未找到华文楷体字体：{CHART_FONT_PATH}")
    fm.fontManager.addfont(str(CHART_FONT_PATH))
    return fm.FontProperties(fname=str(CHART_FONT_PATH), size=7)


def _style_valuation_axis(ax, chart_font: fm.FontProperties) -> None:
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
    for label in ax.get_xticklabels() + ax.get_yticklabels():
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


def plot_daily_valuation(
    data: pd.DataFrame,
    source: dict[str, object],
    output_path: Path,
) -> None:
    """绘制2019年以来反三次、多因子修正及反三次历史分位线。"""
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    inverse = data.dropna(subset=[INVERSE_CUBIC_VALUATION_NAME]).copy()
    multifactor = data.dropna(subset=[MULTIFACTOR_VALUATION_NAME]).copy()
    inverse_series = inverse.set_index("交易日期")[INVERSE_CUBIC_VALUATION_NAME]
    quantiles = inverse_series.quantile([0.25, 0.50, 0.75])

    line_inverse, = ax.plot(
        inverse["交易日期"],
        inverse[INVERSE_CUBIC_VALUATION_NAME],
        color=RED,
        linewidth=1.0,
        marker=None,
        label="百元拟合溢价率",
        zorder=5,
    )
    line_multifactor, = ax.plot(
        multifactor["交易日期"],
        multifactor[MULTIFACTOR_VALUATION_NAME],
        color=BLUE,
        linewidth=1.0,
        marker=None,
        label="多因子修正百元拟合溢价率",
        zorder=4,
    )
    quantile_lines = []
    for quantile, color, label in (
        (0.25, "#9DC3E6", "25%"),
        (0.50, "#A6A6A6", "50%"),
        (0.75, "#E6B9B8", "75%"),
    ):
        quantile_lines.append(
            ax.axhline(
                float(quantiles.loc[quantile]),
                color=color,
                linewidth=1.0,
                linestyle=(0, (5, 4)),
                label=label,
                zorder=2,
            )
        )

    panel_title = (
        f"百元拟合溢价率：{float(source['latestValue']):.2f}%，"
        f"{float(source['dailyChangePctPoint']):+.2f}pct\n"
        f"2019年以来{float(source['percentileSince2019']):.2f}%分位数"
    )
    add_chart_panel_title(fig, panel_title)
    ax.set_xlim(inverse["交易日期"].min(), inverse["交易日期"].max())
    all_values = pd.concat(
        [
            inverse[INVERSE_CUBIC_VALUATION_NAME],
            multifactor[MULTIFACTOR_VALUATION_NAME],
            pd.Series(quantiles.values),
        ],
        ignore_index=True,
    ).dropna()
    y_step = 5.0
    y_min = max(0.0, math.floor((float(all_values.min()) - 1.0) / y_step) * y_step)
    y_max = math.ceil((float(all_values.max()) + 1.0) / y_step) * y_step
    ax.set_ylim(y_min, max(y_min + y_step, y_max))
    ax.yaxis.set_major_locator(mticker.MultipleLocator(y_step))
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    tick_dates = [pd.Timestamp(inverse["交易日期"].min())]
    tick_dates.extend(
        pd.Timestamp(year=year, month=1, day=1)
        for year in range(tick_dates[0].year + 1, inverse["交易日期"].max().year + 1)
    )
    ax.set_xticks(tick_dates)
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(inverse["交易日期"]))
    )
    _style_valuation_axis(ax, chart_font)
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")

    legend = ax.legend(
        [line_inverse, line_multifactor, *quantile_lines],
        ["百元拟合溢价率", "多因子修正百元拟合溢价率", "25%", "50%", "75%"],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.36),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=3,
        handlelength=2.7,
        handletextpad=0.6,
        columnspacing=1.0,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.965, top=0.84, bottom=0.34)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_intraday_valuation(
    data: pd.DataFrame,
    previous_date: pd.Timestamp,
    previous_value: float,
    output_path: Path,
) -> None:
    """绘制盘中百元拟合溢价率，并以虚线标示前一交易日收盘估值。"""
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    x = np.arange(len(data))
    line_intraday, = ax.plot(
        x,
        data["转股溢价率"],
        color=RED,
        linewidth=1.0,
        marker=None,
        label="盘中拟合溢价率",
        zorder=4,
    )
    line_previous = ax.axhline(
        previous_value,
        color=BLUE,
        linewidth=1.0,
        linestyle=(0, (5, 4)),
        label=f"前一交易日（{previous_date:%Y%m%d}）",
        zorder=3,
    )
    # 与同排左侧两行标题共用该排较高的标题栏高度。
    add_chart_panel_title(
        fig,
        "盘中百元平价拟合溢价率",
        band_height=DOUBLE_LINE_TITLE_BAND_HEIGHT,
    )

    all_values = pd.concat(
        [data["转股溢价率"], pd.Series([previous_value])], ignore_index=True
    ).dropna()
    y_step = 0.5
    y_min = math.floor((float(all_values.min()) - 0.15) / y_step) * y_step
    y_max = math.ceil((float(all_values.max()) + 0.15) / y_step) * y_step
    ax.set_ylim(y_min, max(y_min + y_step, y_max))
    ax.yaxis.set_major_locator(mticker.MultipleLocator(y_step))
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    tick_positions = list(range(0, len(data), 10))
    if tick_positions[-1] != len(data) - 1:
        tick_positions.append(len(data) - 1)
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(
        [pd.Timestamp(data.iloc[position]["日期"]).strftime("%H:%M") for position in tick_positions]
    )
    ax.set_xlim(0, len(data) - 1)
    _style_valuation_axis(ax, chart_font)
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")

    legend = ax.legend(
        [line_intraday, line_previous],
        [line_intraday.get_label(), line_previous.get_label()],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.35),
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
    fig.subplots_adjust(left=0.13, right=0.965, top=0.84, bottom=0.34)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_price_parity_series(
    data: pd.DataFrame,
    source: dict[str, object],
    output_path: Path,
) -> None:
    """绘制余额加权平价与收盘价中位数历史序列。"""
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    parity_line, = ax.plot(
        data["交易日期"],
        data["余额加权平价"],
        color=RED,
        linewidth=1.0,
        marker=None,
        label="余额加权平价",
        zorder=4,
    )
    median_line, = ax.plot(
        data["交易日期"],
        data["收盘价中位数"],
        color=BLUE,
        linewidth=1.0,
        marker=None,
        label="收盘价中位数",
        zorder=4,
    )
    panel_title = (
        f"平均平价：{float(source['latestParity']):.2f}，"
        f"{float(source['parityDailyChangePct']):+.2f}%\n"
        f"价格中位数：{float(source['latestMedianPrice']):.2f}，"
        f"{float(source['medianPriceDailyChangePct']):+.2f}%；"
        f"2019年以来{float(source['medianPricePercentileSince2019']):.2f}%分位数"
    )
    add_chart_panel_title(fig, panel_title)

    values = pd.concat(
        [data["余额加权平价"], data["收盘价中位数"]], ignore_index=True
    ).dropna()
    y_step = 10.0
    y_min = math.floor((float(values.min()) - 2.0) / y_step) * y_step
    y_max = math.ceil((float(values.max()) + 2.0) / y_step) * y_step
    ax.set_ylim(y_min, max(y_min + y_step, y_max))
    ax.yaxis.set_major_locator(mticker.MultipleLocator(y_step))
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    ax.set_xlim(data["交易日期"].min(), data["交易日期"].max())
    tick_dates = [pd.Timestamp(data["交易日期"].min())]
    tick_dates.extend(
        pd.Timestamp(year=year, month=1, day=1)
        for year in range(tick_dates[0].year + 1, data["交易日期"].max().year + 1)
    )
    ax.set_xticks(tick_dates)
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(data["交易日期"]))
    )
    _style_valuation_axis(ax, chart_font)
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")

    legend = ax.legend(
        [parity_line, median_line],
        [parity_line.get_label(), median_line.get_label()],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.35),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=2,
        handlelength=2.7,
        handletextpad=0.6,
        columnspacing=1.5,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.965, top=0.84, bottom=0.32)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_subnew_bond_metric(
    data: pd.DataFrame,
    *,
    value_column: str,
    latest_value: float,
    daily_change: float,
    panel_title_label: str,
    legend_label: str,
    output_path: Path,
    include_zero: bool,
) -> None:
    """绘制次新券单指标历史折线图。"""
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    line, = ax.plot(
        data["交易日期"],
        data[value_column],
        color=RED,
        linewidth=1.0,
        marker=None,
        label=legend_label,
        zorder=4,
    )
    add_chart_panel_title(
        fig,
        f"{panel_title_label}：{latest_value:.2f}%，{daily_change:+.2f}pct",
        band_height=DOUBLE_LINE_TITLE_BAND_HEIGHT,
    )

    values = data[value_column].dropna()
    value_min = float(values.min())
    value_max = float(values.max())
    value_span = max(value_max - value_min, 1.0)
    padding = value_span * 0.04
    y_lower = value_min - padding
    y_upper = value_max + padding
    if include_zero:
        y_lower = min(0.0, y_lower)
        y_upper = max(0.0, y_upper)
        ax.axhline(0, color="#7F7F7F", linewidth=0.7, linestyle="--", zorder=2)
    ax.set_ylim(y_lower, y_upper)
    ax.yaxis.set_major_locator(
        mticker.MaxNLocator(
            nbins=7,
            steps=[1, 2, 2.5, 5, 10],
            min_n_ticks=5,
        )
    )
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    ax.set_xlim(data["交易日期"].min(), data["交易日期"].max())
    tick_dates = [pd.Timestamp(data["交易日期"].min())]
    tick_dates.extend(
        pd.Timestamp(year=year, month=1, day=1)
        for year in range(tick_dates[0].year + 1, data["交易日期"].max().year + 1)
    )
    ax.set_xticks(tick_dates)
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(data["交易日期"]))
    )
    _style_valuation_axis(ax, chart_font)
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")

    legend = ax.legend(
        [line],
        [line.get_label()],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.35),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=1,
        handlelength=2.7,
        handletextpad=0.6,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.965, top=0.84, bottom=0.32)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def _set_weighted_metric_axis_scale(ax, values: pd.Series) -> None:
    """为余额加权指标设置紧凑且不留过多空白的纵轴范围。"""
    valid = pd.to_numeric(values, errors="coerce").dropna()
    if valid.empty:
        raise RuntimeError("余额加权指标没有可绘制的有效数值")
    lower = float(valid.min())
    upper = float(valid.max())
    padding = max(0.5, (upper - lower) * 0.1)
    ax.set_ylim(lower - padding, upper + padding)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(nbins=6, min_n_ticks=4))
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))


def plot_equity_bond_weighted_premium(
    data: pd.DataFrame,
    source: dict[str, object],
    output_path: Path,
) -> None:
    """绘制股债性分类余额加权转股溢价率，偏债型使用右轴。"""
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    right_ax = ax.twinx()
    stock_line, = ax.plot(
        data["交易日期"],
        data["偏股型_转股溢价率"],
        color=RED,
        linewidth=1.0,
        label="偏股型",
        zorder=4,
    )
    balanced_line, = ax.plot(
        data["交易日期"],
        data["平衡型_转股溢价率"],
        color="#A6A6A6",
        linewidth=1.0,
        label="平衡型",
        zorder=4,
    )
    bond_line, = right_ax.plot(
        data["交易日期"],
        data["偏债型_转股溢价率"],
        color=BLUE,
        linewidth=1.0,
        label="偏债型",
        zorder=4,
    )
    add_chart_panel_title(
        fig,
        build_equity_bond_weighted_titles(source)["premium"],
    )

    _set_weighted_metric_axis_scale(
        ax,
        pd.concat(
            [data["偏股型_转股溢价率"], data["平衡型_转股溢价率"]]
        ),
    )
    _set_weighted_metric_axis_scale(right_ax, data["偏债型_转股溢价率"])
    ax.set_xlim(data["交易日期"].min(), data["交易日期"].max())
    tick_dates = [pd.Timestamp(data["交易日期"].min())]
    tick_dates.extend(
        pd.Timestamp(year=year, month=1, day=1)
        for year in range(tick_dates[0].year + 1, data["交易日期"].max().year + 1)
    )
    ax.set_xticks(tick_dates)
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(data["交易日期"]))
    )
    _style_valuation_axis(ax, chart_font)
    right_ax.tick_params(
        axis="y",
        which="major",
        colors="black",
        labelsize=7,
        width=0.6,
        length=3,
        right=True,
        labelright=True,
    )
    for label in right_ax.get_yticklabels():
        label.set_fontproperties(chart_font)
        label.set_fontsize(7)
    right_ax.spines["right"].set_color("black")
    right_ax.spines["right"].set_linewidth(0.7)
    right_ax.spines["top"].set_visible(False)
    right_ax.spines["left"].set_visible(False)
    right_ax.tick_params(axis="x", bottom=False, labelbottom=False)
    right_ax.set_ylabel("")
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")

    legend = ax.legend(
        [stock_line, balanced_line, bond_line],
        ["偏股型", "平衡型", "偏债型"],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.35),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=3,
        handlelength=2.5,
        handletextpad=0.5,
        columnspacing=1.2,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.88, top=0.84, bottom=0.32)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_equity_bond_weighted_price(
    data: pd.DataFrame,
    source: dict[str, object],
    output_path: Path,
) -> None:
    """绘制股债性分类余额加权收盘价。"""
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    lines = []
    for group, color in zip(EQUITY_BOND_TYPE_GROUPS, (RED, "#A6A6A6", BLUE)):
        line, = ax.plot(
            data["交易日期"],
            data[f"{group}_收盘价"],
            color=color,
            linewidth=1.0,
            label=group,
            zorder=4,
        )
        lines.append(line)
    add_chart_panel_title(fig, build_equity_bond_weighted_titles(source)["price"])
    _set_weighted_metric_axis_scale(
        ax,
        pd.concat([data[f"{group}_收盘价"] for group in EQUITY_BOND_TYPE_GROUPS]),
    )
    ax.set_xlim(data["交易日期"].min(), data["交易日期"].max())
    tick_dates = [pd.Timestamp(data["交易日期"].min())]
    tick_dates.extend(
        pd.Timestamp(year=year, month=1, day=1)
        for year in range(tick_dates[0].year + 1, data["交易日期"].max().year + 1)
    )
    ax.set_xticks(tick_dates)
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(data["交易日期"]))
    )
    _style_valuation_axis(ax, chart_font)
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")
    legend = ax.legend(
        lines,
        [line.get_label() for line in lines],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.35),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=3,
        handlelength=2.5,
        handletextpad=0.5,
        columnspacing=1.2,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.965, top=0.84, bottom=0.32)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_valuation_repair_index(data: pd.DataFrame, output_path: Path) -> None:
    """绘制估值修复指数及中证转债指数后推70日涨跌幅。"""
    plot_data = data.dropna(subset=["交易日期"]).sort_values("交易日期").copy()
    plot_data = plot_data.loc[
        plot_data["交易日期"].ge(pd.Timestamp(VALUATION_REPAIR_PLOT_START_DATE))
    ]
    if plot_data[VALUATION_REPAIR_INDEX_NAME].notna().sum() < 2:
        raise RuntimeError("估值修复指数有效历史不足")
    chart_font = _valuation_chart_font()
    fig, ax_left = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    ax_right = ax_left.twinx()
    repair_line, = ax_left.plot(
        plot_data["交易日期"],
        plot_data[VALUATION_REPAIR_INDEX_NAME],
        color=RED,
        linewidth=1.0,
        label=VALUATION_REPAIR_INDEX_NAME,
        zorder=4,
    )
    return_line, = ax_right.plot(
        plot_data["交易日期"],
        plot_data[VALUATION_REPAIR_FORWARD_RETURN_NAME],
        color=BLUE,
        linewidth=1.0,
        label="中证转债指数后推70日涨跌幅（%）",
        zorder=4,
    )
    repair_history = plot_data.loc[
        plot_data[VALUATION_REPAIR_INDEX_NAME].notna(),
        ["交易日期", VALUATION_REPAIR_INDEX_NAME],
    ].sort_values("交易日期")
    latest_repair = float(repair_history[VALUATION_REPAIR_INDEX_NAME].iloc[-1])
    previous_repair = float(repair_history[VALUATION_REPAIR_INDEX_NAME].iloc[-2])
    panel_title = (
        f"{VALUATION_REPAIR_INDEX_NAME}{latest_repair:.2f}%；"
        f"{latest_repair - previous_repair:+.2f}pct"
    )
    add_chart_panel_title(
        fig,
        panel_title,
        band_height=DOUBLE_LINE_TITLE_BAND_HEIGHT,
    )

    repair_values = repair_history[VALUATION_REPAIR_INDEX_NAME]
    repair_lower = float(repair_values.min())
    repair_upper = float(repair_values.max())
    repair_span = repair_upper - repair_lower
    if repair_span <= np.finfo(float).eps:
        repair_span = max(abs(repair_upper) * 0.1, 1.0)
    repair_padding = repair_span * 0.05
    ax_left.set_ylim(repair_lower - repair_padding, repair_upper + repair_padding)
    ax_left.yaxis.set_major_locator(
        mticker.MaxNLocator(nbins=6, steps=[1, 2, 2.5, 5, 10], min_n_ticks=4)
    )
    ax_left.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f"))

    return_values = plot_data[VALUATION_REPAIR_FORWARD_RETURN_NAME].dropna()
    return_lower = float(return_values.min())
    return_upper = float(return_values.max())
    return_span = return_upper - return_lower
    if return_span <= np.finfo(float).eps:
        return_span = max(abs(return_upper) * 0.1, 1.0)
    return_padding = return_span * 0.05
    ax_right.set_ylim(return_lower - return_padding, return_upper + return_padding)
    ax_right.yaxis.set_major_locator(
        mticker.MaxNLocator(nbins=7, steps=[1, 2, 2.5, 5, 10], min_n_ticks=4)
    )
    ax_right.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))

    ax_left.set_xlim(
        pd.Timestamp(VALUATION_REPAIR_PLOT_START_DATE),
        plot_data["交易日期"].max(),
    )
    ax_left.xaxis.set_major_locator(mdates.YearLocator())
    ax_left.xaxis.set_minor_locator(mdates.MonthLocator(interval=3))
    ax_left.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(plot_data["交易日期"]))
    )
    style_axis(ax_left, chart_font)
    ax_left.grid(False, which="both", axis="both")
    ax_right.grid(False, which="both", axis="both")
    ax_left.spines["top"].set_visible(False)
    ax_right.spines["top"].set_visible(False)
    ax_left.tick_params(axis="x", top=False, labeltop=False)
    ax_right.tick_params(axis="x", top=False, labeltop=False)
    ax_right.tick_params(axis="y", colors="black", labelsize=TICK_FONT_SIZE, width=0.5)
    ax_right.spines["right"].set_color("black")
    ax_right.spines["right"].set_linewidth(0.5)
    for label in ax_right.get_yticklabels():
        label.set_fontproperties(chart_font)
        label.set_fontsize(TICK_FONT_SIZE)
    legend = ax_left.legend(
        [repair_line, return_line],
        [repair_line.get_label(), return_line.get_label()],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.22),
        ncol=1,
        frameon=False,
        prop=chart_font,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.87, top=0.87, bottom=0.28)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def build_group_valuation_titles(
    sources: dict[str, dict[str, object]],
) -> dict[str, str]:
    """生成分组拟合溢价率图表标题，保留既有数据变动字段格式。"""
    titles: dict[str, str] = {}

    parity = sources.get("parity")
    if parity is not None:
        titles["parity"] = (
            f"平价分类拟合溢价率，{str(parity['largestChangeGroup'])}："
            f"{float(parity['largestChangePctPoint']):+.2f}pct"
        )

    title_specs = (
        ("equity", "股债性分类拟合溢价率", "偏股型", "偏债型", "："),
        ("rating", "评级分类拟合溢价率", "AAA/AA+", "AA/AA-", "："),
        ("balance", "余额分类拟合溢价率：", "0-3", "50+", "："),
        ("market_cap", "市值分类拟合溢价率：", "0-50", "300+", "："),
        ("sector", "各板块拟合溢价率：", "科技", "周期", "："),
    )
    for key, title, first_group, second_group, separator in title_specs:
        details = sources.get(key)
        if details is None:
            continue
        first = details[first_group]
        second = details[second_group]
        titles[key] = (
            f"{title}\n"
            f"{first_group}{separator}{float(first['latestValue']):.2f}%，"
            f"{float(first['dailyChangePctPoint']):+.2f}pct；"
            f"{second_group}{separator}{float(second['latestValue']):.2f}%，"
            f"{float(second['dailyChangePctPoint']):+.2f}pct"
        )

    return titles


def plot_parity_group_valuation(
    data: pd.DataFrame,
    source: dict[str, object],
    output_path: Path,
    panel_title: Optional[str] = None,
) -> None:
    """绘制直接取自 Parquet 的分平价多因子修正拟合溢价率。"""
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    lines = []
    labels = []
    colors = (RED, BLUE, "#ED7D31", "#7F7F7F")
    for (group, parquet_name), color in zip(PARITY_GROUP_SPECS, colors):
        series = data.dropna(subset=[parquet_name])
        line, = ax.plot(
            series["交易日期"],
            series[parquet_name],
            color=color,
            linewidth=1.0,
            marker=None,
            label=group,
            zorder=4,
        )
        lines.append(line)
        labels.append(line.get_label())

    if panel_title is None:
        panel_title = build_group_valuation_titles({"parity": source})["parity"]
    # 与同排右侧双行标题共用该排较高的标题栏高度。
    add_chart_panel_title(
        fig,
        panel_title,
        band_height=DOUBLE_LINE_TITLE_BAND_HEIGHT,
    )

    value_columns = [parquet_name for _, parquet_name in PARITY_GROUP_SPECS]
    values = data[value_columns].stack().dropna()
    y_step = 10.0
    y_min = math.floor((float(values.min()) - 2.0) / y_step) * y_step
    y_max = math.ceil((float(values.max()) + 2.0) / y_step) * y_step
    ax.set_ylim(y_min, max(y_min + y_step, y_max))
    ax.yaxis.set_major_locator(mticker.MultipleLocator(y_step))
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    ax.set_xlim(data["交易日期"].min(), data["交易日期"].max())
    tick_dates = [pd.Timestamp(data["交易日期"].min())]
    tick_dates.extend(
        pd.Timestamp(year=year, month=1, day=1)
        for year in range(tick_dates[0].year + 1, data["交易日期"].max().year + 1)
    )
    ax.set_xticks(tick_dates)
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(data["交易日期"]))
    )
    _style_valuation_axis(ax, chart_font)
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")

    legend = ax.legend(
        lines,
        labels,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.35),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=2,
        handlelength=2.7,
        handletextpad=0.6,
        columnspacing=1.4,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.965, top=0.84, bottom=0.34)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_classification_valuation(
    data: pd.DataFrame,
    specs: tuple[tuple[str, str], ...],
    panel_title: str,
    output_path: Path,
) -> None:
    """按日报统一版式绘制分类拟合溢价率日度序列。"""
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    palette = (RED, BLUE, "#A6A6A6", "#E6B9B8", "#B7DEE8", "#F79646")
    lines = []
    labels = []
    for (label, column), color in zip(specs, palette):
        series = data.dropna(subset=[column])
        line, = ax.plot(
            series["交易日期"],
            series[column],
            color=color,
            linewidth=1.0,
            marker=None,
            label=label,
            zorder=4,
        )
        lines.append(line)
        labels.append(label)
    add_chart_panel_title(
        fig,
        panel_title,
        band_height=DOUBLE_LINE_TITLE_BAND_HEIGHT,
    )

    value_columns = [column for _, column in specs]
    values = data[value_columns].stack().dropna()
    if values.empty:
        raise RuntimeError(f"{panel_title}没有可绘制的有效数据")
    value_range = float(values.max() - values.min())
    y_step = 10.0 if value_range >= 30 else 5.0
    y_min = math.floor((float(values.min()) - y_step * 0.25) / y_step) * y_step
    y_max = math.ceil((float(values.max()) + y_step * 0.25) / y_step) * y_step
    ax.set_ylim(y_min, max(y_min + y_step, y_max))
    ax.yaxis.set_major_locator(mticker.MultipleLocator(y_step))
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    ax.set_xlim(data["交易日期"].min(), data["交易日期"].max())
    tick_dates = [pd.Timestamp(data["交易日期"].min())]
    tick_dates.extend(
        pd.Timestamp(year=year, month=1, day=1)
        for year in range(tick_dates[0].year + 1, data["交易日期"].max().year + 1)
    )
    ax.set_xticks(tick_dates)
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(data["交易日期"]))
    )
    _style_valuation_axis(ax, chart_font)
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")
    legend = ax.legend(
        lines,
        labels,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.35),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=min(3, len(labels)),
        handlelength=2.5,
        handletextpad=0.5,
        columnspacing=1.2,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.965, top=0.84, bottom=0.32)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_sector_mean_metric(
    data: pd.DataFrame,
    metric: str,
    panel_title: str,
    unit_suffix: str,
    output_path: Path,
) -> None:
    """绘制自2023年起的科技、金融、制造、消费和周期日度均值序列。"""
    plot_data = data.loc[
        pd.to_datetime(data["交易日期"], errors="coerce").ge(
            SECTOR_MEAN_PLOT_START_DATE
        )
    ].copy()
    if plot_data.empty:
        raise RuntimeError(
            f"{panel_title}在{SECTOR_MEAN_PLOT_START_DATE:%Y-%m-%d}后没有可绘制的数据"
        )
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    palette = (RED, BLUE, "#A6A6A6", "#E6B9B8", "#B7DEE8")
    lines = []
    value_columns = []
    for sector, color in zip(SECTOR_ORDER, palette):
        column = f"{metric}_{sector}"
        value_columns.append(column)
        series = plot_data.dropna(subset=[column])
        line, = ax.plot(
            series["交易日期"],
            series[column],
            color=color,
            linewidth=1.0,
            marker=None,
            label=sector,
            zorder=4,
        )
        lines.append(line)
    add_chart_panel_title(
        fig,
        panel_title,
        band_height=SINGLE_LINE_TITLE_BAND_HEIGHT,
    )

    values = plot_data[value_columns].stack().dropna()
    if values.empty:
        raise RuntimeError(f"{panel_title}没有可绘制的有效数据")
    value_min = float(values.min())
    value_max = float(values.max())
    value_span = max(value_max - value_min, max(abs(value_max), 1.0) * 0.04)
    ax.set_ylim(value_min - value_span * 0.04, value_max + value_span * 0.05)
    ax.yaxis.set_major_locator(
        mticker.MaxNLocator(nbins=7, steps=[1, 2, 2.5, 5, 10], min_n_ticks=5)
    )
    if unit_suffix == "%":
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    else:
        ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.2f}"))
    ax.set_xlim(plot_data["交易日期"].min(), plot_data["交易日期"].max())
    tick_dates = [pd.Timestamp(plot_data["交易日期"].min())]
    tick_dates.extend(
        pd.Timestamp(year=year, month=1, day=1)
        for year in range(
            tick_dates[0].year + 1, plot_data["交易日期"].max().year + 1
        )
    )
    ax.set_xticks(tick_dates)
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(plot_data["交易日期"]))
    )
    _style_valuation_axis(ax, chart_font)
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")
    legend = ax.legend(
        lines,
        list(SECTOR_ORDER),
        loc="upper center",
        bbox_to_anchor=(0.5, -0.35),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=5,
        handlelength=2.0,
        handletextpad=0.4,
        columnspacing=0.8,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.965, top=0.88, bottom=0.32)
    fig.savefig(output_path, dpi=CHART_DPI, facecolor="white")
    plt.close(fig)


def plot_close_price_distribution_area(
    data: pd.DataFrame,
    source: dict[str, object],
    output_path: Path,
) -> None:
    """绘制2019年以来八档收盘价占比的100%堆积面积图。"""
    chart_font = _valuation_chart_font()
    fig, ax = plt.subplots(figsize=CHART_FIGSIZE, dpi=CHART_DPI)
    colors = (
        "#1F4E79",
        "#5B9BD5",
        "#9DC3E6",
        "#D9E2F3",
        "#D9D9D9",
        "#F4B183",
        "#ED7D31",
        "#C00000",
    )
    legend_labels = ("≤80", "80-90", "90-100", "100-110", "110-120", "120-130", "130-150", ">150")
    areas = ax.stackplot(
        data["交易日期"],
        *[data[label] for label in CLOSE_PRICE_DISTRIBUTION_LABELS],
        colors=colors,
        labels=legend_labels,
        edgecolor="white",
        linewidth=0.2,
        zorder=3,
    )
    panel_title = (
        "收盘价分布：\n"
        f"破底：{float(source['latestBreakFloorPct']):.2f}%，"
        f"{float(source['breakFloorDailyChangePctPoint']):+.2f}pct；"
        f"破面：{float(source['latestBreakParPct']):.2f}%，"
        f"{float(source['breakParDailyChangePctPoint']):+.2f}pct"
    )
    add_chart_panel_title(
        fig,
        panel_title,
        band_height=DOUBLE_LINE_TITLE_BAND_HEIGHT,
    )
    ax.set_xlim(data["交易日期"].min(), data["交易日期"].max())
    ax.set_ylim(0, 100)
    ax.yaxis.set_major_locator(mticker.MultipleLocator(20))
    ax.yaxis.set_major_formatter(mticker.PercentFormatter(xmax=100, decimals=0))
    tick_dates = [pd.Timestamp(data["交易日期"].min())]
    tick_dates.extend(
        pd.Timestamp(year=year, month=1, day=1)
        for year in range(tick_dates[0].year + 1, data["交易日期"].max().year + 1)
    )
    ax.set_xticks(tick_dates)
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter(date_axis_format_for_span(data["交易日期"]))
    )
    _style_valuation_axis(ax, chart_font)
    for label in ax.get_xticklabels():
        label.set_rotation(90)
        label.set_horizontalalignment("center")
    legend = ax.legend(
        areas,
        legend_labels,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.34),
        frameon=False,
        prop=chart_font,
        fontsize=7,
        ncol=4,
        handlelength=1.8,
        handletextpad=0.5,
        columnspacing=0.9,
        borderaxespad=0,
    )
    for legend_text in legend.get_texts():
        legend_text.set_fontproperties(chart_font)
        legend_text.set_fontsize(7)
    fig.subplots_adjust(left=0.13, right=0.965, top=0.84, bottom=0.34)
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


def _blend_rgb(
    start: tuple[int, int, int],
    end: tuple[int, int, int],
    ratio: float,
) -> tuple[int, int, int]:
    ratio = max(0.0, min(1.0, float(ratio)))
    return tuple(
        int(round(start_value + (end_value - start_value) * ratio))
        for start_value, end_value in zip(start, end)
    )


def _return_heat_color(
    value: float | None,
    scale: float,
    negative_color: tuple[int, int, int],
    zero_color: tuple[int, int, int],
    positive_color: tuple[int, int, int],
) -> tuple[int, int, int]:
    if value is None or pd.isna(value) or scale <= 0:
        return (255, 255, 255)
    ratio = min(abs(float(value)) / scale, 1.0)
    if value < 0:
        return _blend_rgb(zero_color, negative_color, ratio)
    return _blend_rgb(zero_color, positive_color, ratio)


def plot_industry_market_table(
    data: pd.DataFrame,
    output_path: Path,
) -> None:
    """绘制申万行业正股、转债涨跌幅与估值分位数大表。"""
    if len(data) != len(INDUSTRY_INDEX_SPECS):
        raise RuntimeError(
            f"行业大表行数异常：{len(data)}，预期 {len(INDUSTRY_INDEX_SPECS)}"
        )
    column_specs = (
        ("行业名称", "名称", "name"),
        ("正股日涨跌幅", "日涨跌幅", "stock"),
        ("正股周涨跌幅", "周涨跌幅", "stock"),
        ("正股月涨跌幅", "月涨跌幅", "stock"),
        ("正股年初至今涨跌幅", "年初至今\n涨跌幅", "stock"),
        ("转债日涨跌幅", "日涨跌幅", "bond"),
        ("转债周涨跌幅", "周涨跌幅", "bond"),
        ("转债月涨跌幅", "月涨跌幅", "bond"),
        ("转债年初至今涨跌幅", "年初至今\n涨跌幅", "bond"),
        ("行业指数PE", "PE（TTM）", "valuation"),
        ("PE三年分位数", "3年分位数", "percentileBlue"),
        ("PE十年分位数", "10年分位数", "percentileRed"),
        ("行业指数PB", "PB（LF）", "valuation"),
        ("PB三年分位数", "3年分位数", "percentileBlue"),
        ("PB十年分位数", "10年分位数", "percentileRed"),
    )
    column_widths = [138] + [98] * 8 + [100, 160, 160, 100, 160, 160]
    if sum(column_widths) != DOUBLE_CHART_PIXEL_WIDTH:
        raise RuntimeError("行业大表列宽合计与日报宽度不一致")
    group_header_height = 39
    header_height = 60
    row_height = 31
    expected_height = group_header_height + header_height + row_height * len(data)
    if expected_height != INDUSTRY_TABLE_PIXEL_HEIGHT:
        raise RuntimeError(
            f"行业大表高度常量异常：{expected_height} != {INDUSTRY_TABLE_PIXEL_HEIGHT}"
        )

    image = Image.new(
        "RGB",
        (DOUBLE_CHART_PIXEL_WIDTH, INDUSTRY_TABLE_PIXEL_HEIGHT),
        "white",
    )
    draw = ImageDraw.Draw(image)
    body_font = ImageFont.truetype(str(FONT_PATH), 21)
    header_font = ImageFont.truetype(str(FONT_PATH), 20)
    group_font = ImageFont.truetype(str(FONT_PATH), 24)
    header_fill = (32, 56, 100)
    header_text = (255, 255, 255)
    grid_color = (0, 0, 0)
    negative_text = (255, 0, 0)

    x_positions = [0]
    for width in column_widths:
        x_positions.append(x_positions[-1] + width)

    def draw_centered_text(
        box: tuple[int, int, int, int],
        text_value: str,
        font: ImageFont.FreeTypeFont,
        fill: tuple[int, int, int],
    ) -> None:
        left, top, right, bottom = box
        bbox = draw.multiline_textbbox(
            (0, 0), text_value, font=font, spacing=1, align="center"
        )
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        draw.multiline_text(
            (
                left + (right - left - text_width) / 2,
                top + (bottom - top - text_height) / 2 - bbox[1],
            ),
            text_value,
            font=font,
            fill=fill,
            spacing=1,
            align="center",
        )

    group_ranges = (
        (0, 1, "名称"),
        (1, 5, "正股"),
        (5, 9, "转债"),
        (9, 15, "正股估值分位数"),
    )
    for start_index, end_index, label in group_ranges:
        group_bottom = (
            group_header_height + header_height
            if label == "名称"
            else group_header_height
        )
        box = (
            x_positions[start_index],
            0,
            x_positions[end_index],
            group_bottom,
        )
        draw.rectangle(box, fill=header_fill, outline=grid_color, width=1)
        draw_centered_text(box, label, group_font, header_text)

    for column_index, (_, label, _) in enumerate(column_specs):
        if column_index == 0:
            continue
        box = (
            x_positions[column_index],
            group_header_height,
            x_positions[column_index + 1],
            group_header_height + header_height,
        )
        draw.rectangle(box, fill=header_fill, outline=grid_color, width=1)
        draw_centered_text(box, label, header_font, header_text)

    stock_columns = [spec[0] for spec in column_specs if spec[2] == "stock"]
    bond_columns = [spec[0] for spec in column_specs if spec[2] == "bond"]
    stock_scales = {
        column: max(
            1e-9,
            float(pd.to_numeric(data[column], errors="coerce").abs().max()),
        )
        for column in stock_columns
    }
    bond_scales = {
        column: max(
            1e-9,
            float(pd.to_numeric(data[column], errors="coerce").abs().max()),
        )
        for column in bond_columns
    }
    for row_position, (_, row) in enumerate(data.iterrows()):
        top = group_header_height + header_height + row_position * row_height
        bottom = top + row_height
        for column_index, (column, _, kind) in enumerate(column_specs):
            left = x_positions[column_index]
            right = x_positions[column_index + 1]
            raw_value = row[column]
            numeric_value = (
                None
                if kind == "name" or raw_value is None or pd.isna(raw_value)
                else float(raw_value)
            )
            if kind == "stock":
                cell_fill = _return_heat_color(
                    numeric_value,
                    stock_scales[column],
                    (91, 155, 213),
                    (255, 255, 255),
                    (248, 105, 107),
                )
            elif kind == "bond":
                cell_fill = _return_heat_color(
                    numeric_value,
                    bond_scales[column],
                    (99, 190, 123),
                    (255, 235, 132),
                    (248, 105, 107),
                )
            elif kind == "valuation":
                cell_fill = (242, 242, 242)
            else:
                cell_fill = (255, 255, 255)
            box = (left, top, right, bottom)
            draw.rectangle(box, fill=cell_fill, outline=grid_color, width=1)

            if kind in ("percentileBlue", "percentileRed") and numeric_value is not None:
                bar_ratio = max(0.0, min(100.0, numeric_value)) / 100.0
                bar_right = left + 2 + int((right - left - 4) * bar_ratio)
                bar_color = (91, 155, 213) if kind == "percentileBlue" else (255, 102, 102)
                draw.rectangle(
                    (left + 2, top + 3, bar_right, bottom - 3),
                    fill=bar_color,
                )

            if kind == "name":
                display_text = str(raw_value)
            elif numeric_value is None:
                display_text = ""
            elif kind == "bond":
                display_text = f"{numeric_value:.2f}%"
            else:
                display_text = f"{numeric_value:.2f}"
            text_color = (
                negative_text
                if numeric_value is not None
                and numeric_value < 0
                and kind in ("stock", "bond")
                else (0, 0, 0)
            )
            draw_centered_text(box, display_text, body_font, text_color)
            draw.rectangle(box, outline=grid_color, width=1)

    temporary_output = output_path.with_name(
        f".{output_path.stem}.{os.getpid()}.tmp{output_path.suffix}"
    )
    try:
        image.save(temporary_output, dpi=(CHART_DPI, CHART_DPI))
        os.replace(temporary_output, output_path)
    finally:
        if temporary_output.exists():
            temporary_output.unlink()


def render_report_header(run_date: date) -> Image.Image:
    """在条款表头底图上绘制加大加粗的日报标题。"""
    if not REPORT_HEADER_PATH.is_file():
        raise FileNotFoundError(f"未找到日报表头图片：{REPORT_HEADER_PATH}")
    with Image.open(REPORT_HEADER_PATH) as source:
        header = source.convert("RGBA")
    draw = ImageDraw.Draw(header)
    text = (
        "【华创固收·周冠南团队】\n"
        f"可转债市场日度跟踪{run_date:%Y%m%d}"
    )
    font = ImageFont.truetype(str(FONT_PATH), REPORT_HEADER_FONT_SIZE)
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    text_x = (header.width - text_width) // 2
    text_y = (header.height - text_height) // 2 - 40
    draw.text(
        (text_x, text_y),
        text,
        fill="white",
        font=font,
    )
    return header


def compose_index_market_overview(
    table_path: Path,
    turnover_path: Path,
    distribution_path: Path,
    main_money_flow_path: Path,
    margin_balance_path: Path,
    daily_valuation_path: Path,
    intraday_valuation_path: Path,
    parity_group_valuation_path: Path,
    price_parity_path: Path,
    equity_bond_weighted_premium_path: Path,
    valuation_repair_path: Path,
    maturity_group_path: Path,
    subnew_premium_path: Path,
    equity_bond_group_path: Path,
    rating_group_path: Path,
    balance_group_path: Path,
    market_cap_group_path: Path,
    sector_group_path: Path,
    close_price_distribution_path: Path,
    boshi_etf_path: Path,
    haifutong_etf_path: Path,
    sector_mean_close_path: Path,
    sector_mean_parity_path: Path,
    sector_mean_conversion_premium_path: Path,
    sector_mean_bond_premium_path: Path,
    industry_table_path: Path,
    output_path: Path,
    run_date: date,
) -> None:
    """将指数表、资金表现与转债估值模块合成为日报长图。"""
    source_header = render_report_header(run_date)
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
    main_money_flow_image = plt.imread(main_money_flow_path)
    margin_balance_image = plt.imread(margin_balance_path)
    daily_valuation_image = plt.imread(daily_valuation_path)
    intraday_valuation_image = plt.imread(intraday_valuation_path)
    parity_group_valuation_image = plt.imread(parity_group_valuation_path)
    price_parity_image = plt.imread(price_parity_path)
    equity_bond_weighted_premium_image = plt.imread(equity_bond_weighted_premium_path)
    valuation_repair_image = plt.imread(valuation_repair_path)
    maturity_group_image = plt.imread(maturity_group_path)
    subnew_premium_image = plt.imread(subnew_premium_path)
    equity_bond_group_image = plt.imread(equity_bond_group_path)
    rating_group_image = plt.imread(rating_group_path)
    balance_group_image = plt.imread(balance_group_path)
    market_cap_group_image = plt.imread(market_cap_group_path)
    sector_group_image = plt.imread(sector_group_path)
    close_price_distribution_image = plt.imread(close_price_distribution_path)
    boshi_etf_image = plt.imread(boshi_etf_path)
    haifutong_etf_image = plt.imread(haifutong_etf_path)
    sector_mean_close_image = plt.imread(sector_mean_close_path)
    sector_mean_parity_image = plt.imread(sector_mean_parity_path)
    sector_mean_conversion_premium_image = plt.imread(
        sector_mean_conversion_premium_path
    )
    sector_mean_bond_premium_image = plt.imread(sector_mean_bond_premium_path)
    industry_table_image = plt.imread(industry_table_path)
    expected_sizes = {
        "指数表": (TABLE_PIXEL_HEIGHT, DOUBLE_CHART_PIXEL_WIDTH),
        "成交额图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "涨跌幅直方图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "主力净流入图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "两融余额图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "百元拟合溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "盘中百元拟合溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "分平价多因子修正拟合溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "余额加权平价与收盘价中位数图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "股债性分类转股溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "转债估值修复指数图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "分剩余期限拟合溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "次新券平均转股溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "股债型拟合溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "分评级拟合溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "分余额拟合溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "分正股市值拟合溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "分板块拟合溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "收盘价分布面积图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "博时可转债ETF份额与净申赎图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "海富通可转债ETF份额与净申赎图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "各行业平均收盘价图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "各行业平均平价图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "各行业平均转股溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "各行业平均纯债溢价率图": (CHART_PIXEL_HEIGHT, CHART_PIXEL_WIDTH),
        "行业涨跌与估值大表": (
            INDUSTRY_TABLE_PIXEL_HEIGHT,
            DOUBLE_CHART_PIXEL_WIDTH,
        ),
    }
    for label, image, expected in (
        ("指数表", table_image, expected_sizes["指数表"]),
        ("成交额图", turnover_image, expected_sizes["成交额图"]),
        ("涨跌幅直方图", distribution_image, expected_sizes["涨跌幅直方图"]),
        ("主力净流入图", main_money_flow_image, expected_sizes["主力净流入图"]),
        ("两融余额图", margin_balance_image, expected_sizes["两融余额图"]),
        ("百元拟合溢价率图", daily_valuation_image, expected_sizes["百元拟合溢价率图"]),
        ("盘中百元拟合溢价率图", intraday_valuation_image, expected_sizes["盘中百元拟合溢价率图"]),
        (
            "分平价多因子修正拟合溢价率图",
            parity_group_valuation_image,
            expected_sizes["分平价多因子修正拟合溢价率图"],
        ),
        (
            "余额加权平价与收盘价中位数图",
            price_parity_image,
            expected_sizes["余额加权平价与收盘价中位数图"],
        ),
        (
            "股债性分类转股溢价率图",
            equity_bond_weighted_premium_image,
            expected_sizes["股债性分类转股溢价率图"],
        ),
        (
            "转债估值修复指数图",
            valuation_repair_image,
            expected_sizes["转债估值修复指数图"],
        ),
        (
            "分剩余期限拟合溢价率图",
            maturity_group_image,
            expected_sizes["分剩余期限拟合溢价率图"],
        ),
        (
            "次新券平均转股溢价率图",
            subnew_premium_image,
            expected_sizes["次新券平均转股溢价率图"],
        ),
        (
            "股债型拟合溢价率图",
            equity_bond_group_image,
            expected_sizes["股债型拟合溢价率图"],
        ),
        (
            "分评级拟合溢价率图",
            rating_group_image,
            expected_sizes["分评级拟合溢价率图"],
        ),
        (
            "分余额拟合溢价率图",
            balance_group_image,
            expected_sizes["分余额拟合溢价率图"],
        ),
        (
            "分正股市值拟合溢价率图",
            market_cap_group_image,
            expected_sizes["分正股市值拟合溢价率图"],
        ),
        (
            "分板块拟合溢价率图",
            sector_group_image,
            expected_sizes["分板块拟合溢价率图"],
        ),
        (
            "收盘价分布面积图",
            close_price_distribution_image,
            expected_sizes["收盘价分布面积图"],
        ),
        (
            "博时可转债ETF份额与净申赎图",
            boshi_etf_image,
            expected_sizes["博时可转债ETF份额与净申赎图"],
        ),
        (
            "海富通可转债ETF份额与净申赎图",
            haifutong_etf_image,
            expected_sizes["海富通可转债ETF份额与净申赎图"],
        ),
        (
            "各行业平均收盘价图",
            sector_mean_close_image,
            expected_sizes["各行业平均收盘价图"],
        ),
        (
            "各行业平均平价图",
            sector_mean_parity_image,
            expected_sizes["各行业平均平价图"],
        ),
        (
            "各行业平均转股溢价率图",
            sector_mean_conversion_premium_image,
            expected_sizes["各行业平均转股溢价率图"],
        ),
        (
            "各行业平均纯债溢价率图",
            sector_mean_bond_premium_image,
            expected_sizes["各行业平均纯债溢价率图"],
        ),
        (
            "行业涨跌与估值大表",
            industry_table_image,
            expected_sizes["行业涨跌与估值大表"],
        ),
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

    def render_section_bar(title: str) -> np.ndarray:
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
            title,
            ha="center",
            va="center",
            fontproperties=get_title_font(7),
            fontsize=7,
            fontweight="bold",
            color="white",
        )
        section_figure.canvas.draw()
        rendered = (
            np.asarray(section_figure.canvas.buffer_rgba(), dtype=np.float32)
            / 255.0
        )
        plt.close(section_figure)
        if rendered.shape[:2] != (
            SECTION_BAR_HEIGHT,
            DOUBLE_CHART_PIXEL_WIDTH,
        ):
            raise RuntimeError(
                f"{title}分隔栏尺寸异常：{rendered.shape[1]}×{rendered.shape[0]}"
            )
        return rendered

    section_bar = render_section_bar("资金表现")
    valuation_section_bar = render_section_bar("转债估值")
    etf_section_bar = render_section_bar("ETF表现")
    industry_section_bar = render_section_bar("行业表现")

    canvas = np.ones(
        (
            header_height
            + TABLE_PIXEL_HEIGHT
            + CHART_PIXEL_HEIGHT * 12
            + SECTION_BAR_HEIGHT * 4
            + INDUSTRY_TABLE_PIXEL_HEIGHT,
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
    first_chart_end = chart_start + CHART_PIXEL_HEIGHT
    second_chart_end = first_chart_end + CHART_PIXEL_HEIGHT
    canvas[chart_start:first_chart_end, :CHART_PIXEL_WIDTH, :] = to_rgba(
        turnover_image
    )
    canvas[chart_start:first_chart_end, CHART_PIXEL_WIDTH:, :] = to_rgba(
        distribution_image
    )
    canvas[first_chart_end:second_chart_end, :CHART_PIXEL_WIDTH, :] = to_rgba(
        main_money_flow_image
    )
    canvas[first_chart_end:second_chart_end, CHART_PIXEL_WIDTH:, :] = to_rgba(
        margin_balance_image
    )
    valuation_section_end = second_chart_end + SECTION_BAR_HEIGHT
    canvas[second_chart_end:valuation_section_end, :, :] = valuation_section_bar
    valuation_chart_end = valuation_section_end + CHART_PIXEL_HEIGHT
    canvas[
        valuation_section_end:valuation_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(daily_valuation_image)
    canvas[
        valuation_section_end:valuation_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(intraday_valuation_image)
    second_valuation_chart_end = valuation_chart_end + CHART_PIXEL_HEIGHT
    canvas[
        valuation_chart_end:second_valuation_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(valuation_repair_image)
    canvas[
        valuation_chart_end:second_valuation_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(price_parity_image)
    third_valuation_chart_end = second_valuation_chart_end + CHART_PIXEL_HEIGHT
    canvas[
        second_valuation_chart_end:third_valuation_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(equity_bond_weighted_premium_image)
    canvas[
        second_valuation_chart_end:third_valuation_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(parity_group_valuation_image)
    fourth_valuation_chart_end = third_valuation_chart_end + CHART_PIXEL_HEIGHT
    canvas[
        third_valuation_chart_end:fourth_valuation_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(maturity_group_image)
    canvas[
        third_valuation_chart_end:fourth_valuation_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(subnew_premium_image)
    fifth_valuation_chart_end = fourth_valuation_chart_end + CHART_PIXEL_HEIGHT
    canvas[
        fourth_valuation_chart_end:fifth_valuation_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(equity_bond_group_image)
    canvas[
        fourth_valuation_chart_end:fifth_valuation_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(rating_group_image)
    sixth_valuation_chart_end = fifth_valuation_chart_end + CHART_PIXEL_HEIGHT
    canvas[
        fifth_valuation_chart_end:sixth_valuation_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(balance_group_image)
    canvas[
        fifth_valuation_chart_end:sixth_valuation_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(market_cap_group_image)
    seventh_valuation_chart_end = sixth_valuation_chart_end + CHART_PIXEL_HEIGHT
    canvas[
        sixth_valuation_chart_end:seventh_valuation_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(sector_group_image)
    canvas[
        sixth_valuation_chart_end:seventh_valuation_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(close_price_distribution_image)
    etf_section_end = seventh_valuation_chart_end + SECTION_BAR_HEIGHT
    canvas[
        seventh_valuation_chart_end:etf_section_end,
        :,
        :,
    ] = etf_section_bar
    etf_chart_end = etf_section_end + CHART_PIXEL_HEIGHT
    canvas[
        etf_section_end:etf_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(boshi_etf_image)
    canvas[
        etf_section_end:etf_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(haifutong_etf_image)
    industry_section_end = etf_chart_end + SECTION_BAR_HEIGHT
    canvas[
        etf_chart_end:industry_section_end,
        :,
        :,
    ] = industry_section_bar
    first_industry_chart_end = industry_section_end + CHART_PIXEL_HEIGHT
    canvas[
        industry_section_end:first_industry_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(sector_mean_close_image)
    canvas[
        industry_section_end:first_industry_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(sector_mean_parity_image)
    second_industry_chart_end = first_industry_chart_end + CHART_PIXEL_HEIGHT
    canvas[
        first_industry_chart_end:second_industry_chart_end,
        :CHART_PIXEL_WIDTH,
        :,
    ] = to_rgba(sector_mean_conversion_premium_image)
    canvas[
        first_industry_chart_end:second_industry_chart_end,
        CHART_PIXEL_WIDTH:,
        :,
    ] = to_rgba(sector_mean_bond_premium_image)
    industry_table_end = second_industry_chart_end + INDUSTRY_TABLE_PIXEL_HEIGHT
    canvas[
        second_industry_chart_end:industry_table_end,
        :,
        :,
    ] = to_rgba(industry_table_image)
    temporary_output = output_path.with_name(
        f".{output_path.stem}.{os.getpid()}.tmp{output_path.suffix}"
    )
    try:
        plt.imsave(temporary_output, canvas, dpi=CHART_DPI)
        os.replace(temporary_output, output_path)
    finally:
        if temporary_output.exists():
            temporary_output.unlink()


def compose_fuguo_daily_report(
    daily_valuation_path: Path,
    intraday_valuation_path: Path,
    valuation_repair_path: Path,
    price_parity_path: Path,
    parity_group_valuation_path: Path,
    equity_bond_weighted_premium_path: Path,
    output_path: Path,
) -> None:
    """将转债估值分隔栏与六张指定图按三行双列合成为富国日报。"""
    chart_paths = (
        daily_valuation_path,
        intraday_valuation_path,
        valuation_repair_path,
        price_parity_path,
        parity_group_valuation_path,
        equity_bond_weighted_premium_path,
    )
    charts: list[Image.Image] = []
    try:
        for path in chart_paths:
            if not path.is_file():
                raise FileNotFoundError(f"未找到富国日报图表：{path}")
            with Image.open(path) as source:
                chart = source.convert("RGB")
            if chart.size != (CHART_PIXEL_WIDTH, CHART_PIXEL_HEIGHT):
                raise RuntimeError(
                    f"富国日报图表尺寸异常：{path.name}为"
                    f"{chart.width}×{chart.height}，预期"
                    f"{CHART_PIXEL_WIDTH}×{CHART_PIXEL_HEIGHT}"
                )
            charts.append(chart)

        output_height = SECTION_BAR_HEIGHT + CHART_PIXEL_HEIGHT * 3
        canvas = Image.new(
            "RGB", (DOUBLE_CHART_PIXEL_WIDTH, output_height), "white"
        )
        draw = ImageDraw.Draw(canvas)
        draw.rectangle(
            (0, 0, DOUBLE_CHART_PIXEL_WIDTH, SECTION_BAR_HEIGHT),
            fill="#203864",
        )
        section_font_path = (
            TITLE_FONT_PATH if TITLE_FONT_PATH.is_file() else FONT_PATH
        )
        section_font = ImageFont.truetype(str(section_font_path), 20)
        title = "转债估值"
        title_box = draw.textbbox((0, 0), title, font=section_font)
        title_width = title_box[2] - title_box[0]
        title_height = title_box[3] - title_box[1]
        draw.text(
            (
                (DOUBLE_CHART_PIXEL_WIDTH - title_width) / 2,
                (SECTION_BAR_HEIGHT - title_height) / 2 - title_box[1],
            ),
            title,
            font=section_font,
            fill="white",
        )
        for position, chart in enumerate(charts):
            row, column = divmod(position, 2)
            canvas.paste(
                chart,
                (
                    column * CHART_PIXEL_WIDTH,
                    SECTION_BAR_HEIGHT + row * CHART_PIXEL_HEIGHT,
                ),
            )

        temporary_output = output_path.with_name(
            f".{output_path.stem}.{os.getpid()}.tmp{output_path.suffix}"
        )
        try:
            canvas.save(temporary_output, dpi=(CHART_DPI, CHART_DPI))
            os.replace(temporary_output, output_path)
        finally:
            if temporary_output.exists():
                temporary_output.unlink()
    finally:
        for chart in charts:
            chart.close()


def _remove_first_raster_row_preserving_size(image: Image.Image) -> Image.Image:
    """移除标题分隔线所在首行，同时保持模板所需的像素尺寸。"""
    if image.height < 2:
        raise RuntimeError("小图正文高度不足2像素")
    cleaned = Image.new(image.mode, image.size)
    cleaned.paste(image.crop((0, 1, image.width, image.height)), (0, 0))
    cleaned.paste(
        image.crop((0, image.height - 1, image.width, image.height)),
        (0, image.height - 1),
    )
    return cleaned


def export_numbered_titleless_small_charts(
    output_dir: Path,
    chart_specs: tuple[tuple[int, str, str, float], ...] = SMALL_CHART_EXPORT_SPECS,
) -> None:
    """长图合成后导出不带标题栏、按阅读顺序编号的独立小图。"""
    for sequence, label, source_name, title_band_height in chart_specs:
        source_path = output_dir / source_name
        if not source_path.is_file():
            raise FileNotFoundError(f"未找到小图源文件：{source_path}")
        output_path = output_dir / f"{sequence:02d}_{label}.png"
        with Image.open(source_path) as source:
            crop_top = round(source.height * title_band_height)
            if crop_top <= 0 or crop_top >= source.height:
                raise RuntimeError(f"小图标题栏高度异常：{source_path.name}")
            titleless = source.crop((0, crop_top, source.width, source.height))
            if titleless.height < 2:
                raise RuntimeError(f"小图正文高度异常：{source_path.name}")
            titleless = _remove_first_raster_row_preserving_size(titleless)
            temporary_output = output_path.with_name(
                f".{output_path.stem}.{os.getpid()}.tmp.png"
            )
            try:
                titleless.save(temporary_output)
                os.replace(temporary_output, output_path)
            finally:
                if temporary_output.exists():
                    temporary_output.unlink()
        source_path.unlink()


def build_workbook(
    market: pd.DataFrame,
    main_money_flow: pd.DataFrame,
    main_money_flow_source: dict[str, object],
    etf_share: pd.DataFrame,
    etf_share_source: dict[str, object],
    index: pd.DataFrame,
    index_performance: pd.DataFrame,
    index_performance_source: dict[str, object],
    return_details: pd.DataFrame,
    return_distribution: pd.DataFrame,
    return_summary: dict[str, int],
    return_source: dict[str, object],
    daily_valuation: pd.DataFrame,
    valuation_source: dict[str, object],
    valuation_repair: pd.DataFrame,
    intraday_valuation: pd.DataFrame,
    intraday_valuation_source: dict[str, object],
    parity_interval_premium: pd.DataFrame,
    parity_group_valuation: pd.DataFrame,
    parity_group_valuation_source: dict[str, object],
    equity_bond_group_valuation: pd.DataFrame,
    equity_bond_group_valuation_source: dict[str, object],
    rating_group_valuation: pd.DataFrame,
    rating_group_valuation_source: dict[str, object],
    maturity_group_valuation: pd.DataFrame,
    maturity_group_valuation_source: dict[str, object],
    balance_group_valuation: pd.DataFrame,
    balance_group_valuation_source: dict[str, object],
    market_cap_group_valuation: pd.DataFrame,
    market_cap_group_valuation_source: dict[str, object],
    sector_group_valuation: pd.DataFrame,
    sector_group_valuation_source: dict[str, object],
    sector_mean_metrics: pd.DataFrame,
    sector_mean_source: dict[str, object],
    industry_history_metrics: pd.DataFrame,
    industry_performance: pd.DataFrame,
    industry_performance_source: dict[str, object],
    close_price_distribution: pd.DataFrame,
    close_price_distribution_source: dict[str, object],
    price_parity: pd.DataFrame,
    price_parity_source: dict[str, object],
    equity_bond_weighted: pd.DataFrame,
    equity_bond_weighted_source: dict[str, object],
    subnew_bond: pd.DataFrame,
    subnew_bond_source: dict[str, object],
    run_date: date,
    index_start_date: date,
    market_start_date: date,
    output_path: Path,
) -> None:
    """通过 bundled artifact-tool 生成合并 Excel 底稿。"""
    for dependency in (BUNDLED_NODE, BUNDLED_NODE_MODULES):
        if not dependency.exists():
            raise FileNotFoundError(f"生成 Excel 所需依赖不存在：{dependency}")

    sector_payload_keys = {
        "科技": "technology",
        "金融": "finance",
        "制造": "manufacturing",
        "消费": "consumption",
        "周期": "cyclical",
    }
    metric_payload_keys = {
        "收盘价": "close",
        "平价": "parity",
        "转股溢价率": "conversionPremium",
        "纯债溢价率": "bondPremium",
    }
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
        "mainMoneyFlowSource": main_money_flow_source,
        "mainMoneyFlow": [
            {
                "date": f"{row.交易日期:%Y-%m-%d}",
                "amount": float(row.主力净流入_亿元),
            }
            for row in main_money_flow.itertuples(index=False)
        ],
        "etfShareSource": etf_share_source,
        "etfShare": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "boshiShare": (
                    None
                    if pd.isna(row["博时可转债ETF份额_亿份"])
                    else float(row["博时可转债ETF份额_亿份"])
                ),
                "haifutongShare": (
                    None
                    if pd.isna(row["海富通可转债ETF份额_亿份"])
                    else float(row["海富通可转债ETF份额_亿份"])
                ),
            }
            for _, row in etf_share.iterrows()
        ],
        "subnewBondSource": subnew_bond_source,
        "subnewBond": [
            {
                "date": f"{row.交易日期:%Y-%m-%d}",
                "listingReturnMean": float(row.次新券相对上市涨跌幅均值),
                "listingReturnSampleCount": int(row.价格样本数),
                "premiumMean": float(row.次新券平均转股溢价率),
                "premiumSampleCount": int(row.溢价率样本数),
            }
            for row in subnew_bond.itertuples(index=False)
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
        "valuationSource": valuation_source,
        "intradayValuationSource": intraday_valuation_source,
        "valuationDaily": [
            {
                "date": f"{row.交易日期:%Y-%m-%d}",
                "inverseCubic": (
                    None
                    if pd.isna(getattr(row, INVERSE_CUBIC_VALUATION_NAME))
                    else float(getattr(row, INVERSE_CUBIC_VALUATION_NAME))
                ),
                "multifactor": (
                    None
                    if pd.isna(getattr(row, MULTIFACTOR_VALUATION_NAME))
                    else float(getattr(row, MULTIFACTOR_VALUATION_NAME))
                ),
            }
            for row in daily_valuation.itertuples(index=False)
        ],
        "valuationRepair": [
            {
                "date": f"{row.交易日期:%Y-%m-%d}",
                "repair": (
                    None
                    if pd.isna(getattr(row, VALUATION_REPAIR_INDEX_NAME))
                    else float(getattr(row, VALUATION_REPAIR_INDEX_NAME))
                ),
                "forwardReturn": float(
                    getattr(row, VALUATION_REPAIR_FORWARD_RETURN_NAME)
                ),
            }
            for row in valuation_repair.itertuples(index=False)
        ],
        "valuationIntraday": [
            {
                "datetime": pd.Timestamp(row.日期).strftime("%Y-%m-%dT%H:%M:%S"),
                "premium": float(row.转股溢价率),
            }
            for row in intraday_valuation.itertuples(index=False)
        ],
        "parityIntervalPremium": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "above130": (
                    None if pd.isna(row["130以上"]) else float(row["130以上"])
                ),
                "p120_130": (
                    None
                    if pd.isna(row["120-130（含130）"])
                    else float(row["120-130（含130）"])
                ),
                "p110_120": (
                    None
                    if pd.isna(row["110-120（含120）"])
                    else float(row["110-120（含120）"])
                ),
                "p100_110": (
                    None
                    if pd.isna(row["100-110（含110）"])
                    else float(row["100-110（含110）"])
                ),
                "p90_100": (
                    None
                    if pd.isna(row["90-100（含100）"])
                    else float(row["90-100（含100）"])
                ),
                "p80_90": (
                    None
                    if pd.isna(row["80-90（含90）"])
                    else float(row["80-90（含90）"])
                ),
                "le80": (
                    None
                    if pd.isna(row["80以下（含80）"])
                    else float(row["80以下（含80）"])
                ),
            }
            for _, row in parity_interval_premium.iterrows()
        ],
        "parityGroupValuationSource": parity_group_valuation_source,
        "parityGroupValuation": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "group70_90": (
                    None
                    if pd.isna(row[f"{PARITY_GROUP_VALUATION_PREFIX}70-90"])
                    else float(row[f"{PARITY_GROUP_VALUATION_PREFIX}70-90"])
                ),
                "group90_110": (
                    None
                    if pd.isna(row[f"{PARITY_GROUP_VALUATION_PREFIX}90-110"])
                    else float(row[f"{PARITY_GROUP_VALUATION_PREFIX}90-110"])
                ),
                "group110_130": (
                    None
                    if pd.isna(row[f"{PARITY_GROUP_VALUATION_PREFIX}110-130"])
                    else float(row[f"{PARITY_GROUP_VALUATION_PREFIX}110-130"])
                ),
                "group130_150": (
                    None
                    if pd.isna(row[f"{PARITY_GROUP_VALUATION_PREFIX}130-150"])
                    else float(row[f"{PARITY_GROUP_VALUATION_PREFIX}130-150"])
                ),
            }
            for _, row in parity_group_valuation.iterrows()
        ],
        "equityBondGroupValuationSource": equity_bond_group_valuation_source,
        "equityBondGroupValuation": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "stock": (
                    None
                    if pd.isna(row[f"{EQUITY_BOND_GROUP_PREFIX}偏股型"])
                    else float(row[f"{EQUITY_BOND_GROUP_PREFIX}偏股型"])
                ),
                "balance": (
                    None
                    if pd.isna(row[f"{EQUITY_BOND_GROUP_PREFIX}平衡型"])
                    else float(row[f"{EQUITY_BOND_GROUP_PREFIX}平衡型"])
                ),
                "bond": (
                    None
                    if pd.isna(row[f"{EQUITY_BOND_GROUP_PREFIX}偏债型"])
                    else float(row[f"{EQUITY_BOND_GROUP_PREFIX}偏债型"])
                ),
            }
            for _, row in equity_bond_group_valuation.iterrows()
        ],
        "ratingGroupValuationSource": rating_group_valuation_source,
        "ratingGroupValuation": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "top": (
                    None
                    if pd.isna(row[f"{RATING_GROUP_PREFIX}AAA/AA+"])
                    else float(row[f"{RATING_GROUP_PREFIX}AAA/AA+"])
                ),
                "middle": (
                    None
                    if pd.isna(row[f"{RATING_GROUP_PREFIX}AA/AA-"])
                    else float(row[f"{RATING_GROUP_PREFIX}AA/AA-"])
                ),
                "lower": (
                    None
                    if pd.isna(row[f"{RATING_GROUP_PREFIX}A+/A"])
                    else float(row[f"{RATING_GROUP_PREFIX}A+/A"])
                ),
            }
            for _, row in rating_group_valuation.iterrows()
        ],
        "maturityGroupValuationSource": maturity_group_valuation_source,
        "maturityGroupValuation": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                **{
                    f"group{group.replace('-', '_')}": (
                        None if pd.isna(row[parquet_name]) else float(row[parquet_name])
                    )
                    for group, parquet_name in MATURITY_GROUP_SPECS
                },
            }
            for _, row in maturity_group_valuation.iterrows()
        ],
        "balanceGroupValuationSource": balance_group_valuation_source,
        "balanceGroupValuation": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "group0_3": (
                    None
                    if pd.isna(row[f"{BALANCE_GROUP_PREFIX}0-3亿元"])
                    else float(row[f"{BALANCE_GROUP_PREFIX}0-3亿元"])
                ),
                "group3_10": (
                    None
                    if pd.isna(row[f"{BALANCE_GROUP_PREFIX}3-10亿元"])
                    else float(row[f"{BALANCE_GROUP_PREFIX}3-10亿元"])
                ),
                "group10_20": (
                    None
                    if pd.isna(row[f"{BALANCE_GROUP_PREFIX}10-20亿元"])
                    else float(row[f"{BALANCE_GROUP_PREFIX}10-20亿元"])
                ),
                "group20_50": (
                    None
                    if pd.isna(row[f"{BALANCE_GROUP_PREFIX}20-50亿元"])
                    else float(row[f"{BALANCE_GROUP_PREFIX}20-50亿元"])
                ),
                "group50_plus": (
                    None
                    if pd.isna(row[f"{BALANCE_GROUP_PREFIX}50亿元以上"])
                    else float(row[f"{BALANCE_GROUP_PREFIX}50亿元以上"])
                ),
            }
            for _, row in balance_group_valuation.iterrows()
        ],
        "marketCapGroupValuationSource": market_cap_group_valuation_source,
        "marketCapGroupValuation": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "group0_50": (
                    None
                    if pd.isna(row[f"{MARKET_CAP_GROUP_PREFIX}0-50亿元"])
                    else float(row[f"{MARKET_CAP_GROUP_PREFIX}0-50亿元"])
                ),
                "group50_300": (
                    None
                    if pd.isna(row[f"{MARKET_CAP_GROUP_PREFIX}50-300亿元"])
                    else float(row[f"{MARKET_CAP_GROUP_PREFIX}50-300亿元"])
                ),
                "group300_plus": (
                    None
                    if pd.isna(row[f"{MARKET_CAP_GROUP_PREFIX}300亿元以上"])
                    else float(row[f"{MARKET_CAP_GROUP_PREFIX}300亿元以上"])
                ),
            }
            for _, row in market_cap_group_valuation.iterrows()
        ],
        "sectorGroupValuationSource": sector_group_valuation_source,
        "sectorGroupValuation": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "technology": float(row["多因子修正拟合溢价率_科技"]),
                "finance": float(row["多因子修正拟合溢价率_金融"]),
                "manufacturing": float(row["多因子修正拟合溢价率_制造"]),
                "consumption": float(row["多因子修正拟合溢价率_消费"]),
                "cyclical": float(row["多因子修正拟合溢价率_周期"]),
            }
            for _, row in sector_group_valuation.dropna().iterrows()
        ],
        "sectorMeanSource": sector_mean_source,
        "sectorMeanMetrics": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                **{
                    f"{metric_payload_keys[metric]}_{sector_payload_keys[sector]}": (
                        None
                        if pd.isna(row[f"{metric}_{sector}"])
                        else float(row[f"{metric}_{sector}"])
                    )
                    for metric, _, _ in SECTOR_MEAN_METRICS
                    for sector in SECTOR_ORDER
                },
            }
            for _, row in sector_mean_metrics.iterrows()
        ],
        "industryHistoryOrder": list(INDUSTRY_HISTORY_ORDER),
        "industryHistoryMetrics": [
            {
                "industry": str(row["行业"]),
                "date": f"{row['日期']:%Y-%m-%d}",
                "close": None if pd.isna(row["收盘价"]) else float(row["收盘价"]),
                "conversionPremium": (
                    None
                    if pd.isna(row["转股溢价率"])
                    else float(row["转股溢价率"])
                ),
                "parity": None if pd.isna(row["平价"]) else float(row["平价"]),
                "bondPremium": (
                    None
                    if pd.isna(row["纯债溢价率"])
                    else float(row["纯债溢价率"])
                ),
            }
            for _, row in industry_history_metrics.iterrows()
        ],
        "industryPerformanceSource": industry_performance_source,
        "industryPerformance": [
            {
                "code": str(row["行业代码"]),
                "name": str(row["行业名称"]),
                "stockDaily": float(row["正股日涨跌幅"]),
                "stockWeekly": float(row["正股周涨跌幅"]),
                "stockMonthly": float(row["正股月涨跌幅"]),
                "stockYtd": float(row["正股年初至今涨跌幅"]),
                "bondDaily": (
                    None
                    if pd.isna(row["转债日涨跌幅"])
                    else float(row["转债日涨跌幅"])
                ),
                "bondWeekly": (
                    None
                    if pd.isna(row["转债周涨跌幅"])
                    else float(row["转债周涨跌幅"])
                ),
                "bondMonthly": (
                    None
                    if pd.isna(row["转债月涨跌幅"])
                    else float(row["转债月涨跌幅"])
                ),
                "bondYtd": (
                    None
                    if pd.isna(row["转债年初至今涨跌幅"])
                    else float(row["转债年初至今涨跌幅"])
                ),
                "pe": float(row["行业指数PE"]),
                "pe3y": float(row["PE三年分位数"]),
                "pe10y": float(row["PE十年分位数"]),
                "pb": float(row["行业指数PB"]),
                "pb3y": float(row["PB三年分位数"]),
                "pb10y": float(row["PB十年分位数"]),
                "stockDailyFormula": str(row["正股日涨跌幅公式"]),
                "stockWeeklyFormula": str(row["正股周涨跌幅公式"]),
                "stockMonthlyFormula": str(row["正股月涨跌幅公式"]),
                "stockYtdFormula": str(row["正股年初至今涨跌幅公式"]),
                "peFormula": str(row["行业指数PE公式"]),
                "pe3yFormula": str(row["PE三年分位数公式"]),
                "pe10yFormula": str(row["PE十年分位数公式"]),
                "pbFormula": str(row["行业指数PB公式"]),
                "pb3yFormula": str(row["PB三年分位数公式"]),
                "pb10yFormula": str(row["PB十年分位数公式"]),
                "bondSampleCount": int(row["转债当日样本数"]),
            }
            for _, row in industry_performance.iterrows()
        ],
        "closePriceDistributionSource": close_price_distribution_source,
        "closePriceDistribution": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "le80": float(row["80以下（含80）"]),
                "p80_90": float(row["80-90（含90）"]),
                "p90_100": float(row["90-100（含100）"]),
                "p100_110": float(row["100-110（含110）"]),
                "p110_120": float(row["110-120（含120）"]),
                "p120_130": float(row["120-130（含130）"]),
                "p130_150": float(row["130-150（含150）"]),
                "gt150": float(row["150以上"]),
                "sampleCount": int(row["有效样本数"]),
            }
            for _, row in close_price_distribution.iterrows()
        ],
        "priceParitySource": price_parity_source,
        "priceParity": [
            {
                "date": f"{row.交易日期:%Y-%m-%d}",
                "weightedParity": float(row.余额加权平价),
                "medianClose": float(row.收盘价中位数),
                "paritySampleCount": int(row.平价样本数),
                "priceSampleCount": int(row.价格样本数),
                "effectiveBalance": float(row.有效余额),
            }
            for row in price_parity.itertuples(index=False)
        ],
        "equityBondWeightedSource": equity_bond_weighted_source,
        "equityBondWeighted": [
            {
                "date": f"{row['交易日期']:%Y-%m-%d}",
                "stockPremium": (
                    None
                    if pd.isna(row["偏股型_转股溢价率"])
                    else float(row["偏股型_转股溢价率"])
                ),
                "balancePremium": (
                    None
                    if pd.isna(row["平衡型_转股溢价率"])
                    else float(row["平衡型_转股溢价率"])
                ),
                "bondPremium": (
                    None
                    if pd.isna(row["偏债型_转股溢价率"])
                    else float(row["偏债型_转股溢价率"])
                ),
                "stockPrice": (
                    None
                    if pd.isna(row["偏股型_收盘价"])
                    else float(row["偏股型_收盘价"])
                ),
                "balancePrice": (
                    None
                    if pd.isna(row["平衡型_收盘价"])
                    else float(row["平衡型_收盘价"])
                ),
                "bondPrice": (
                    None
                    if pd.isna(row["偏债型_收盘价"])
                    else float(row["偏债型_收盘价"])
                ),
            }
            for _, row in equity_bond_weighted.iterrows()
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
    enforce_industry_history_workbook_layout(output_path)


def enforce_industry_history_workbook_layout(workbook_path: Path) -> None:
    """确保四张行业历史表冻结首行首列。"""
    sheet_names = {
        "行业收盘价历史",
        "行业平价历史",
        "行业转股溢价率历史",
        "行业纯债溢价率历史",
    }
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_rel_ns = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", main_ns)
    ET.register_namespace("r", office_rel_ns)

    workbook_path = Path(workbook_path)
    with zipfile.ZipFile(workbook_path, "r") as source:
        workbook_root = ET.fromstring(source.read("xl/workbook.xml"))
        relationships_root = ET.fromstring(
            source.read("xl/_rels/workbook.xml.rels")
        )
        relationships = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships_root.findall(
                f"{{{package_rel_ns}}}Relationship"
            )
        }
        targets: set[str] = set()
        for sheet in workbook_root.findall(
            f"{{{main_ns}}}sheets/{{{main_ns}}}sheet"
        ):
            if sheet.attrib.get("name") not in sheet_names:
                continue
            relationship_id = sheet.attrib.get(f"{{{office_rel_ns}}}id")
            target = relationships.get(relationship_id or "")
            if not target:
                continue
            member = target.replace("\\", "/").lstrip("/")
            if not member.startswith("xl/"):
                member = f"xl/{member}"
            targets.add(member)

        if len(targets) != len(sheet_names):
            raise RuntimeError(
                "行业历史底稿缺少工作表："
                f"期望{len(sheet_names)}张，实际定位{len(targets)}张"
            )

        patched: dict[str, bytes] = {}
        for member in targets:
            root = ET.fromstring(source.read(member))
            sheet_views = root.find(f"{{{main_ns}}}sheetViews")
            if sheet_views is None:
                sheet_views = ET.Element(f"{{{main_ns}}}sheetViews")
                root.insert(0, sheet_views)
            sheet_view = sheet_views.find(f"{{{main_ns}}}sheetView")
            if sheet_view is None:
                sheet_view = ET.SubElement(
                    sheet_views,
                    f"{{{main_ns}}}sheetView",
                    {"workbookViewId": "0"},
                )
            for child in list(sheet_view):
                if child.tag in {
                    f"{{{main_ns}}}pane",
                    f"{{{main_ns}}}selection",
                }:
                    sheet_view.remove(child)
            pane = ET.Element(
                f"{{{main_ns}}}pane",
                {
                    "xSplit": "1",
                    "ySplit": "1",
                    "topLeftCell": "B2",
                    "activePane": "bottomRight",
                    "state": "frozen",
                },
            )
            sheet_view.insert(0, pane)
            for pane_name, active_cell in (
                ("topRight", "B1"),
                ("bottomLeft", "A2"),
                ("bottomRight", "B2"),
            ):
                ET.SubElement(
                    sheet_view,
                    f"{{{main_ns}}}selection",
                    {
                        "pane": pane_name,
                        "activeCell": active_cell,
                        "sqref": active_cell,
                    },
                )
            patched[member] = ET.tostring(
                root, encoding="utf-8", xml_declaration=True
            )

        temp_handle = tempfile.NamedTemporaryFile(
            prefix=f"{workbook_path.stem}_",
            suffix=".xlsx",
            dir=workbook_path.parent,
            delete=False,
        )
        temp_path = Path(temp_handle.name)
        temp_handle.close()
        try:
            with zipfile.ZipFile(temp_path, "w") as destination:
                for info in source.infolist():
                    destination.writestr(info, patched.get(info.filename, source.read(info)))
            source.close()
            os.replace(temp_path, workbook_path)
        except Exception:
            temp_path.unlink(missing_ok=True)
            raise


def _commentary_direction(
    value: float,
    positive: str = "上涨",
    negative: str = "下降",
    zero: str = "持平",
) -> tuple[str, float]:
    numeric = float(value)
    if numeric > 1e-12:
        return positive, abs(numeric)
    if numeric < -1e-12:
        return negative, abs(numeric)
    return zero, 0.0


def _commentary_index_clause(label: str, value: float) -> str:
    direction, magnitude = _commentary_direction(value)
    if direction == "持平":
        return f"{label}环比持平"
    return f"{label}环比{direction}{magnitude:.2f}%"


def _commentary_subscription_text(value_yi: float) -> str:
    value = float(value_yi)
    if abs(value) < 1.0:
        converted = value * 10000.0
        text = f"{converted:+.3f}".rstrip("0").rstrip(".")
        return f"{text}万份"
    text = f"{value:+.3f}".rstrip("0").rstrip(".")
    return f"{text}亿份"


def _commentary_ranked_industries(
    data: pd.DataFrame,
    value_column: str,
) -> tuple[int, int, str, str]:
    frame = data[["行业名称", value_column]].copy()
    frame[value_column] = pd.to_numeric(frame[value_column], errors="coerce")
    frame = frame.dropna(subset=["行业名称", value_column])
    gainers = frame.loc[frame[value_column].gt(0)].sort_values(
        value_column, ascending=False
    )
    decliners = frame.loc[frame[value_column].lt(0)].sort_values(
        value_column, ascending=True
    )

    def render(rows: pd.DataFrame) -> str:
        return "、".join(
            f"{row['行业名称']}（{float(row[value_column]):+.2f}%）"
            for _, row in rows.head(3).iterrows()
        )

    return len(gainers), len(decliners), render(gainers), render(decliners)


def _commentary_sector_change_line(
    sector_mean_metrics: pd.DataFrame,
    metric: str,
    *,
    percentage_point: bool,
) -> str:
    data = sector_mean_metrics.copy()
    data["交易日期"] = pd.to_datetime(data["交易日期"], errors="coerce")
    data = data.dropna(subset=["交易日期"]).sort_values("交易日期")
    if len(data) < 2:
        raise RuntimeError(f"{metric}板块点评至少需要两个交易日")
    latest = data.iloc[-1]
    previous = data.iloc[-2]
    labels = {
        "周期": "大周期",
        "制造": "制造",
        "科技": "科技",
        "消费": "大消费",
        "金融": "大金融",
    }
    clauses: list[str] = []
    for sector in ("周期", "制造", "科技", "消费", "金融"):
        column = f"{metric}_{sector}"
        if column not in data.columns:
            raise RuntimeError(f"板块点评数据缺少字段：{column}")
        latest_value = float(latest[column])
        previous_value = float(previous[column])
        if percentage_point:
            change = latest_value - previous_value
            suffix = "pct"
        else:
            if previous_value == 0:
                raise RuntimeError(f"板块点评前值为0，无法计算环比：{column}")
            change = (latest_value / previous_value - 1.0) * 100.0
            suffix = "%"
        clauses.append(f"{labels[sector]}环比{change:+.2f}{suffix}")
    return "、".join(clauses) + "。"


def build_fuguo_daily_text(
    run_date: date,
    *,
    index_performance: pd.DataFrame,
    turnover: pd.DataFrame,
    price_parity_source: dict[str, object],
    valuation_source: dict[str, object],
) -> str:
    """生成富国日报使用的精简版转债市场日度跟踪文字。"""
    index_rows = index_performance.loc[
        index_performance["指数名称"].astype(str).eq("中证转债"),
        "日涨跌幅",
    ]
    if index_rows.empty or pd.isna(index_rows.iloc[-1]):
        raise RuntimeError("富国日报点评缺少中证转债指数日涨跌幅")
    index_change = float(index_rows.iloc[-1])
    index_direction, index_magnitude = _commentary_direction(
        index_change, positive="上涨", negative="下跌"
    )

    turnover_data = turnover.copy()
    turnover_data["交易日期"] = pd.to_datetime(
        turnover_data["交易日期"], errors="coerce"
    )
    turnover_data["中证转债指数成交额_亿元"] = pd.to_numeric(
        turnover_data["中证转债指数成交额_亿元"], errors="coerce"
    )
    turnover_data = (
        turnover_data.dropna(
            subset=["交易日期", "中证转债指数成交额_亿元"]
        )
        .sort_values("交易日期")
        .drop_duplicates("交易日期", keep="last")
    )
    if len(turnover_data) < 2:
        raise RuntimeError("富国日报点评至少需要两个交易日的转债成交额")
    latest_turnover = float(
        turnover_data.iloc[-1]["中证转债指数成交额_亿元"]
    )
    previous_turnover = float(
        turnover_data.iloc[-2]["中证转债指数成交额_亿元"]
    )
    if previous_turnover == 0:
        raise RuntimeError("富国日报点评前一交易日转债成交额为0")
    turnover_change = (latest_turnover / previous_turnover - 1.0) * 100.0
    turnover_direction, turnover_magnitude = _commentary_direction(
        turnover_change, positive="增长", negative="下降"
    )
    sentiment = (
        "升温" if turnover_change > 0 else "降温" if turnover_change < 0 else "平稳"
    )

    price_previous_date = pd.Timestamp(price_parity_source["previousDate"])
    price_previous_text = (
        f"{price_previous_date.month}月{price_previous_date.day}日"
    )
    median_change = float(price_parity_source["medianPriceDailyChangePct"])
    median_direction, median_magnitude = _commentary_direction(
        median_change, positive="上升", negative="下降"
    )

    valuation_previous_date = pd.Timestamp(valuation_source["previousDate"])
    valuation_previous_text = (
        f"{valuation_previous_date.month}月{valuation_previous_date.day}日"
    )
    valuation_change = float(valuation_source["dailyChangePctPoint"])
    valuation_direction, valuation_magnitude = _commentary_direction(
        valuation_change, positive="上升", negative="下降"
    )
    return (
        f"转债市场日度跟踪{run_date:%Y%m%d}\n"
        f"中证转债指数{index_direction}{index_magnitude:.2f}%。"
        f"转债市场成交情绪{sentiment}，可转债市场成交额为"
        f"{latest_turnover:.2f}亿元，环比{turnover_direction}"
        f"{turnover_magnitude:.2f}%。"
        f"价格中位数为{float(price_parity_source['latestMedianPrice']):.2f}元，"
        f"环比{price_previous_text}{median_direction}{median_magnitude:.2f}%。"
        f"百元平价拟合转股溢价率为{float(valuation_source['latestValue']):.2f}%，"
        f"环比{valuation_previous_text}{valuation_direction}"
        f"{valuation_magnitude:.2f}pct。"
        f"溢价率在2019年以来"
        f"{float(valuation_source['percentileSince2019']):.2f}%分位数。"
    )


def build_long_chart_titles(
    *,
    turnover: pd.DataFrame,
    return_summary: dict[str, int],
    main_money_flow_source: dict[str, object],
    margin_balance: pd.DataFrame,
    valuation_source: dict[str, object],
    valuation_repair: pd.DataFrame,
    price_parity_source: dict[str, object],
    equity_bond_weighted_source: dict[str, object],
    group_valuation_titles: dict[str, str],
    maturity_title: str,
    subnew_bond_source: dict[str, object],
    close_price_distribution_source: dict[str, object],
    etf_share_source: dict[str, object],
) -> list[str]:
    """按长图由上至下、同排由左至右生成全部图表标题。"""
    turnover_data = turnover.dropna(
        subset=[
            "交易日期",
            "中证转债指数成交额_亿元",
            "沪深成交额合计_亿元",
        ]
    ).sort_values("交易日期")
    if turnover_data.empty:
        raise RuntimeError("长图标题缺少成交额数据")
    latest_turnover = turnover_data.iloc[-1]
    turnover_title = (
        f"成交额:转债{float(latest_turnover['中证转债指数成交额_亿元']):.2f}亿，"
        f"A股{float(latest_turnover['沪深成交额合计_亿元']):.2f}亿"
    )

    valid_count = max(int(return_summary["有效样本"]), 1)
    distribution_title = (
        f"上涨转债占比{int(return_summary['上涨']) / valid_count * 100:.2f}%，"
        f"下跌转债占比{int(return_summary['下跌']) / valid_count * 100:.2f}%"
    )
    main_flow = float(main_money_flow_source["latestValue"])
    main_flow_title = (
        f"沪深两市主力{'净流入' if main_flow >= 0 else '净流出'}"
        f"{abs(main_flow):.2f}亿元"
    )
    _, margin_title = prepare_margin_balance_plot(margin_balance)

    daily_valuation_title = (
        f"百元拟合溢价率：{float(valuation_source['latestValue']):.2f}%，"
        f"{float(valuation_source['dailyChangePctPoint']):+.2f}pct\n"
        f"2019年以来{float(valuation_source['percentileSince2019']):.2f}%分位数"
    )
    repair_history = valuation_repair.dropna(
        subset=["交易日期", VALUATION_REPAIR_INDEX_NAME]
    ).sort_values("交易日期")
    if len(repair_history) < 2:
        raise RuntimeError("长图标题缺少估值修复指数前后两个有效交易日")
    latest_repair = float(repair_history.iloc[-1][VALUATION_REPAIR_INDEX_NAME])
    previous_repair = float(repair_history.iloc[-2][VALUATION_REPAIR_INDEX_NAME])
    repair_title = (
        f"{VALUATION_REPAIR_INDEX_NAME}{latest_repair:.2f}%；"
        f"{latest_repair - previous_repair:+.2f}pct"
    )
    price_parity_title = (
        f"平均平价：{float(price_parity_source['latestParity']):.2f}，"
        f"{float(price_parity_source['parityDailyChangePct']):+.2f}%\n"
        f"价格中位数：{float(price_parity_source['latestMedianPrice']):.2f}，"
        f"{float(price_parity_source['medianPriceDailyChangePct']):+.2f}%；"
        f"2019年以来"
        f"{float(price_parity_source['medianPricePercentileSince2019']):.2f}%分位数"
    )
    weighted_premium_title = build_equity_bond_weighted_titles(
        equity_bond_weighted_source
    )["premium"]
    subnew_title = (
        "次新券平均转股溢价率："
        f"{float(subnew_bond_source['latestPremiumMeanPct']):.2f}%，"
        f"{float(subnew_bond_source['premiumDailyChangePctPoint']):+.2f}pct"
    )
    close_distribution_title = (
        "收盘价分布：\n"
        f"破底：{float(close_price_distribution_source['latestBreakFloorPct']):.2f}%，"
        f"{float(close_price_distribution_source['breakFloorDailyChangePctPoint']):+.2f}pct；"
        f"破面：{float(close_price_distribution_source['latestBreakParPct']):.2f}%，"
        f"{float(close_price_distribution_source['breakParDailyChangePctPoint']):+.2f}pct"
    )

    etf_titles: list[str] = []
    for etf_name in ("博时可转债ETF", "海富通可转债ETF"):
        detail = etf_share_source["funds"][etf_name]
        latest_flow_value = detail["latestNetSubscriptionYi"]
        latest_flow = 0.0 if latest_flow_value is None else float(latest_flow_value)
        latest_flow_text = (
            f"{latest_flow * 10000:+,.0f}万份"
            if abs(latest_flow) < 1.0
            else f"{latest_flow:+,.2f}亿份"
        )
        etf_titles.append(
            f"{etf_name}：\n份额{float(detail['latestShareYi']):.2f}亿份，"
            f"净申赎{latest_flow_text}"
        )

    titles = [
        turnover_title,
        distribution_title,
        main_flow_title,
        margin_title,
        daily_valuation_title,
        "盘中百元平价拟合溢价率",
        repair_title,
        price_parity_title,
        weighted_premium_title,
        group_valuation_titles["parity"],
        maturity_title,
        subnew_title,
        group_valuation_titles["equity"],
        group_valuation_titles["rating"],
        group_valuation_titles["balance"],
        group_valuation_titles["market_cap"],
        group_valuation_titles["sector"],
        close_distribution_title,
        *etf_titles,
        *[panel_title for _, panel_title, _ in SECTOR_MEAN_METRICS],
    ]
    if len(titles) != 24:
        raise RuntimeError(f"长图图表标题数量异常：{len(titles)}")
    return titles


def append_long_chart_titles(commentary: str, titles: list[str]) -> str:
    """在日报点评末尾追加可直接复制的长图标题清单。"""
    cleaned_titles = [str(title).strip() for title in titles if str(title).strip()]
    return (
        commentary.rstrip()
        + "\n\n图表标题：\n"
        + "\n\n".join(cleaned_titles)
    )


def build_daily_commentary(
    run_date: date,
    *,
    index_performance: pd.DataFrame,
    turnover: pd.DataFrame,
    main_money_flow_source: dict[str, object],
    etf_share_source: dict[str, object],
    price_parity_source: dict[str, object],
    equity_bond_weighted_source: dict[str, object],
    close_price_distribution: pd.DataFrame,
    close_price_distribution_source: dict[str, object],
    valuation_source: dict[str, object],
    valuation_repair: pd.DataFrame,
    subnew_bond_source: dict[str, object],
    equity_bond_group_valuation_source: dict[str, object],
    parity_group_valuation_source: dict[str, object],
    industry_performance: pd.DataFrame,
    sector_mean_metrics: pd.DataFrame,
) -> str:
    """按固定日报话术，用本次运行已经取得的数据生成点评文字。"""
    turnover_data = turnover.copy()
    turnover_data["交易日期"] = pd.to_datetime(
        turnover_data["交易日期"], errors="coerce"
    )
    turnover_data = turnover_data.dropna(subset=["交易日期"]).sort_values(
        "交易日期"
    )
    if len(turnover_data) < 2:
        raise RuntimeError("成交额点评至少需要两个交易日")
    turnover_latest = turnover_data.iloc[-1]
    turnover_previous = turnover_data.iloc[-2]
    cb_turnover = float(turnover_latest["中证转债指数成交额_亿元"])
    stock_turnover = float(turnover_latest["沪深成交额合计_亿元"])
    cb_turnover_change = (
        cb_turnover / float(turnover_previous["中证转债指数成交额_亿元"]) - 1.0
    ) * 100.0
    stock_turnover_change = (
        stock_turnover / float(turnover_previous["沪深成交额合计_亿元"]) - 1.0
    ) * 100.0

    index_map = {
        str(row["指数名称"]): float(row["日涨跌幅"])
        for _, row in index_performance.iterrows()
        if pd.notna(row.get("指数名称")) and pd.notna(row.get("日涨跌幅"))
    }
    required_indices = {
        "中证转债",
        "上证综指",
        "深证成指",
        "创业板指",
        "上证50",
        "中证1000",
        "大盘成长",
        "大盘价值",
        "中盘成长",
        "中盘价值",
        "小盘成长",
        "小盘价值",
    }
    if not required_indices.issubset(index_map):
        raise RuntimeError(
            f"日报点评缺少指数：{sorted(required_indices - set(index_map))}"
        )

    volume_word = "放量" if cb_turnover_change > 0 else "缩量"
    if abs(cb_turnover_change) <= 1e-12:
        volume_word = "量平"
    index_word = _commentary_direction(index_map["中证转债"])[0]
    valuation_word = _commentary_direction(
        float(valuation_source["dailyChangePctPoint"]),
        positive="估值扩张",
        negative="估值压缩",
        zero="估值持平",
    )[0]
    overview = f"市场概况：转债{volume_word}{index_word}，{valuation_word}"

    main_specs = (
        ("中证转债指数", "中证转债"),
        ("上证综指", "上证综指"),
        ("深证成指", "深证成指"),
        ("创业板指", "创业板指"),
        ("上证50指数", "上证50"),
        ("中证1000指数", "中证1000"),
    )
    index_text = "指数表现：" + "、".join(
        _commentary_index_clause(label, index_map[key])
        for label, key in main_specs
    ) + "。"

    style_names = (
        "大盘成长",
        "大盘价值",
        "中盘成长",
        "中盘价值",
        "小盘成长",
        "小盘价值",
    )
    dominant_style = max(style_names, key=lambda name: index_map[name])
    style_text = (
        f"市场风格：{dominant_style}相对占优。"
        + "、".join(
            _commentary_index_clause(name, index_map[name]) for name in style_names
        )
        + "。"
    )

    cb_turnover_direction, cb_turnover_magnitude = _commentary_direction(
        cb_turnover_change, positive="增加", negative="减少"
    )
    stock_turnover_direction, stock_turnover_magnitude = _commentary_direction(
        stock_turnover_change, positive="增加", negative="减少"
    )
    sentiment = (
        "增强" if cb_turnover_change > 0 else "减弱" if cb_turnover_change < 0 else "平稳"
    )
    fund_sentences = []
    for fund_name in ("博时可转债ETF", "海富通可转债ETF"):
        details = etf_share_source["funds"][fund_name]
        fund_sentences.append(
            f"{fund_name}份额{float(details['latestShareYi']):.2f}亿份，"
            f"净申赎{_commentary_subscription_text(float(details['latestNetSubscriptionYi']))}"
        )
    money_flow = float(main_money_flow_source["latestValue"])
    money_flow_direction = "净流入" if money_flow >= 0 else "净流出"
    funds_text = (
        f"资金表现：转债市场成交情绪{sentiment}。"
        f"可转债市场成交额为{cb_turnover:.2f}亿元，"
        f"环比{cb_turnover_direction}{cb_turnover_magnitude:.2f}%；"
        f"沪深两市总成交额为{stock_turnover:.2f}亿元，"
        f"环比{stock_turnover_direction}{stock_turnover_magnitude:.2f}%；"
        f"沪深两市主力{money_flow_direction}{abs(money_flow):.2f}亿元；"
        + "，".join(fund_sentences)
        + "。"
    )

    previous_date = pd.Timestamp(price_parity_source["previousDate"])
    previous_date_text = f"{previous_date.month}月{previous_date.day}日"
    weighted_close_change = float(
        price_parity_source["weightedCloseDailyChangePct"]
    )
    price_center = "上升" if weighted_close_change > 0 else "下降" if weighted_close_change < 0 else "持平"
    distribution = close_price_distribution.copy()
    distribution["交易日期"] = pd.to_datetime(
        distribution["交易日期"], errors="coerce"
    )
    distribution = distribution.dropna(subset=["交易日期"]).sort_values("交易日期")
    if len(distribution) < 2 or "150以上" not in distribution.columns:
        raise RuntimeError("收盘价分布点评缺少150元以上档或前一交易日")
    high_price_latest = float(distribution.iloc[-1]["150以上"])
    high_price_change = high_price_latest - float(distribution.iloc[-2]["150以上"])
    high_price_word = "上升" if high_price_change > 0 else "下降" if high_price_change < 0 else "持平"
    weighted_close_direction, weighted_close_magnitude = _commentary_direction(
        weighted_close_change, positive="上升", negative="下降"
    )
    median_direction, median_magnitude = _commentary_direction(
        float(price_parity_source["medianPriceDailyChangePct"]),
        positive="上升",
        negative="下降",
    )
    category_price_clauses = []
    category_prices = equity_bond_weighted_source["price"]
    for position, category in enumerate(("偏股型", "偏债型", "平衡型")):
        details = category_prices[category]
        direction, magnitude = _commentary_direction(
            float(details["dailyChangePct"]), positive="上升", negative="下降"
        )
        prefix = "其中" if position == 0 else ""
        category_price_clauses.append(
            f"{prefix}{category}转债的收盘价为{float(details['latestValue']):.2f}元，"
            f"环比{direction}{magnitude:.2f}%"
        )
    price_text = (
        f"转债价格：转债中枢{price_center}，高价券占比{high_price_word}。"
        f"转债整体收盘价加权平均值为{float(price_parity_source['latestWeightedClose']):.2f}元，"
        f"环比{previous_date_text}{weighted_close_direction}{weighted_close_magnitude:.2f}%。"
        + "；".join(category_price_clauses)
        + "。"
        f"从转债收盘价分布情况看，150元以上高价券个数占比{high_price_latest:.2f}%，"
        f"较{previous_date_text}环比{high_price_change:+.2f}pct；"
        f"价格中位数为{float(price_parity_source['latestMedianPrice']):.2f}元，"
        f"环比{previous_date_text}{median_direction}{median_magnitude:.2f}%；"
        f"破底占比为{float(close_price_distribution_source['latestBreakFloorPct']):.2f}%，"
        f"环比{float(close_price_distribution_source['breakFloorDailyChangePctPoint']):+.2f}pct。"
    )

    repair = valuation_repair[["交易日期", VALUATION_REPAIR_INDEX_NAME]].copy()
    repair["交易日期"] = pd.to_datetime(repair["交易日期"], errors="coerce")
    repair[VALUATION_REPAIR_INDEX_NAME] = pd.to_numeric(
        repair[VALUATION_REPAIR_INDEX_NAME], errors="coerce"
    )
    repair = repair.dropna().sort_values("交易日期")
    if len(repair) < 2:
        raise RuntimeError("转债估值修复指数点评至少需要两个交易日")
    repair_latest = float(repair.iloc[-1][VALUATION_REPAIR_INDEX_NAME])
    repair_change = repair_latest - float(repair.iloc[-2][VALUATION_REPAIR_INDEX_NAME])
    valuation_change = float(valuation_source["dailyChangePctPoint"])
    valuation_summary = "估值扩张" if valuation_change > 0 else "估值压缩" if valuation_change < 0 else "估值持平"
    fitted_clauses = []
    fitted_details = equity_bond_group_valuation_source["groupDetails"]
    for category in ("偏股型", "偏债型", "平衡型"):
        details = fitted_details[category]
        direction, magnitude = _commentary_direction(
            float(details["dailyChangePctPoint"])
        )
        fitted_clauses.append(
            f"{category}拟合溢价率为{float(details['latestValue']):.2f}%，"
            f"环比{direction}{magnitude:.2f}pct"
        )
    parity_clauses = []
    parity_details = parity_group_valuation_source["groupDetails"]
    for group in ("110-130", "90-110", "70-90"):
        details = parity_details[group]
        direction, magnitude = _commentary_direction(
            float(details["dailyChangePctPoint"])
        )
        parity_clauses.append(
            f"平价{group}转债拟合溢价率为{float(details['latestValue']):.2f}%，"
            f"环比{direction}{magnitude:.2f}pct"
        )
    valuation_text = (
        f"转债估值：{valuation_summary}。"
        f"百元平价拟合转股溢价率为{float(valuation_source['latestValue']):.2f}%，"
        f"环比{previous_date_text}{_commentary_direction(valuation_change)[0]}"
        f"{abs(valuation_change):.2f}pct；"
        f"转债估值修复指数{repair_latest:.2f}%；{repair_change:+.2f}pct；"
        f"次新券平均转股溢价率为{float(subnew_bond_source['latestPremiumMeanPct']):.2f}%，"
        f"{float(subnew_bond_source['premiumDailyChangePctPoint']):+.2f}pct。"
        f"整体加权平价为{float(price_parity_source['latestParity']):.2f}元，"
        f"环比{previous_date_text}{_commentary_direction(float(price_parity_source['parityDailyChangePct']), positive='上升', negative='下降')[0]}"
        f"{abs(float(price_parity_source['parityDailyChangePct'])):.2f}%。"
        f"分类别来看，{'；'.join(fitted_clauses)}。"
        + "；".join(parity_clauses)
        + "。"
    )

    stock_up, stock_down, stock_gainers, stock_decliners = (
        _commentary_ranked_industries(industry_performance, "正股日涨跌幅")
    )
    bond_up, bond_down, bond_gainers, bond_decliners = (
        _commentary_ranked_industries(industry_performance, "转债日涨跌幅")
    )
    stock_breadth = "上涨占比过半" if stock_up > stock_down else "上涨占比未过半"
    industry_date_text = f"{run_date.month}月{run_date.day}日"
    industry_text = (
        f"行业表现：{industry_date_text}正股行业指数{stock_breadth}，"
        f"共计{stock_up}个行业上涨。A股市场中，涨幅前三位行业为{stock_gainers}；"
        f"跌幅前三位行业为{stock_decliners}。\n"
        f"转债市场共计{bond_down}个行业下跌，跌幅前三位行业为{bond_decliners}；"
        f"涨幅前三位行业为{bond_gainers}。"
    )

    sector_lines = (
        "(1) 收盘价："
        + _commentary_sector_change_line(
            sector_mean_metrics, "收盘价", percentage_point=False
        ),
        "(2) 转股溢价率："
        + _commentary_sector_change_line(
            sector_mean_metrics, "转股溢价率", percentage_point=True
        ),
        "(3) 转换价值："
        + _commentary_sector_change_line(
            sector_mean_metrics, "平价", percentage_point=False
        ),
        "(4) 纯债溢价率："
        + _commentary_sector_change_line(
            sector_mean_metrics, "纯债溢价率", percentage_point=True
        ),
    )
    return "\n".join(
        (
            overview,
            index_text,
            style_text,
            funds_text,
            price_text,
            valuation_text,
            industry_text,
            *sector_lines,
        )
    )


def write_daily_commentary(text: str, output_path: Path) -> None:
    """以 UTF-8 原子写入日报点评文本。"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_output = output_path.with_name(
        f".{output_path.stem}.{os.getpid()}.tmp{output_path.suffix}"
    )
    try:
        temporary_output.write_text(text.rstrip() + "\n", encoding="utf-8")
        os.replace(temporary_output, output_path)
    finally:
        if temporary_output.exists():
            temporary_output.unlink()


def remove_obsolete_outputs(output_dir: Path, workbook_path: Path) -> None:
    """仅清理此前版本生成的已知 CSV/JSON 中间文件。"""
    for name in (
        "市场交易统计_p03438.csv",
        "指数成交金额_ths_trans_amt_index.csv",
        "运行信息.json",
        workbook_path.name + ".inspect.ndjson",
        "上证指数与深证成指成交额.png",
        "可转债上涨下跌占比.png",
        "次新券相对上市涨跌幅均值.png",
        "股债性分类均价.png",
        "07_平价分类拟合溢价率.png",
        "10_转债估值修复指数.png",
    ):
        path = output_dir / name
        if path.is_file():
            path.unlink()


def run(
    run_date: date,
    output_dir: Path,
    progress: Optional[ConsoleProgress] = None,
) -> dict[str, object]:
    def report_progress(percent: int, message: str) -> None:
        if progress is not None:
            progress.update(percent, message)

    start_date = same_day_last_year(run_date)
    market_start_date = MARGIN_BALANCE_START_DATE
    output_dir.mkdir(parents=True, exist_ok=True)

    report_progress(2, "登录 iFinD")
    login_code = ths_login()
    if not is_ths_login_ok(login_code):
        raise RuntimeError(
            f"iFinD 登录失败（{login_code}）：{ths_login_errmsg(login_code)}"
        )

    report_progress(6, "读取资金数据")
    market = fetch_market_statistics(market_start_date, run_date)
    main_money_flow, main_money_flow_source = fetch_main_money_flow(run_date)
    etf_share, etf_share_source = fetch_cb_etf_share_series(run_date)
    report_progress(16, "读取市场表现")
    index = fetch_index_turnover(start_date, run_date)
    index_performance, index_performance_source = fetch_index_performance(run_date)
    return_details, return_distribution, return_summary, return_source = (
        fetch_cb_daily_returns(run_date)
    )
    report_progress(27, "读取估值数据")
    daily_valuation, valuation_source = fetch_daily_valuation_series(run_date)
    valuation_repair = fetch_valuation_repair_index(run_date)
    intraday_valuation, intraday_valuation_source = fetch_intraday_valuation(run_date)
    parity_interval_premium, parity_interval_premium_source = (
        fetch_parity_interval_premium_series(run_date)
    )
    parity_group_valuation, parity_group_valuation_source = (
        fetch_parity_group_valuation_series(run_date)
    )
    equity_bond_group_valuation, equity_bond_group_valuation_source = (
        fetch_equity_bond_group_valuation_series(run_date)
    )
    rating_group_valuation, rating_group_valuation_source = (
        fetch_rating_group_valuation_series(run_date)
    )
    maturity_group_valuation, maturity_group_valuation_source = (
        fetch_maturity_group_valuation_series(run_date)
    )
    balance_group_valuation, balance_group_valuation_source = (
        fetch_balance_group_valuation_series(run_date)
    )
    market_cap_group_valuation, market_cap_group_valuation_source = (
        fetch_market_cap_group_valuation_series(run_date)
    )
    sector_group_valuation, sector_group_valuation_source = (
        fetch_sector_group_valuation_series(run_date)
    )
    report_progress(43, "读取行业与分布数据")
    sector_mean_metrics, sector_mean_source = fetch_sector_mean_metrics(run_date)
    industry_history_metrics = fetch_industry_history_metrics(run_date)
    industry_performance, industry_performance_source = (
        fetch_industry_market_table(run_date)
    )
    close_price_distribution, close_price_distribution_source = (
        fetch_close_price_distribution_series(run_date)
    )
    price_parity, price_parity_source = fetch_price_parity_series(run_date)
    equity_bond_weighted, equity_bond_weighted_source = (
        fetch_equity_bond_weighted_series(run_date)
    )
    subnew_bond, subnew_bond_source = fetch_subnew_bond_series(run_date)
    report_progress(55, "数据读取完成")

    market_png = output_dir / "沪深两市融资融券余额.png"
    main_money_flow_png = output_dir / "沪深两市主力净流入.png"
    boshi_etf_png = output_dir / "博时可转债ETF份额与净申赎.png"
    haifutong_etf_png = output_dir / "海富通可转债ETF份额与净申赎.png"
    index_png = output_dir / "中证转债与沪深两市成交额.png"
    return_png = output_dir / "可转债当日涨跌幅分布.png"
    index_performance_png = output_dir / "主要指数与风格指数表现.png"
    daily_valuation_png = output_dir / "百元拟合溢价率.png"
    intraday_valuation_png = output_dir / "盘中百元平价拟合溢价率.png"
    parity_group_valuation_png = output_dir / "分平价多因子修正拟合溢价率.png"
    price_parity_png = output_dir / "余额加权平价与收盘价中位数.png"
    equity_bond_weighted_premium_png = output_dir / "股债性分类转股溢价率.png"
    valuation_repair_png = output_dir / "转债估值修复指数.png"
    maturity_group_png = output_dir / "分剩余期限拟合溢价率.png"
    subnew_premium_png = output_dir / "次新券平均转股溢价率.png"
    equity_bond_group_png = output_dir / "股债型拟合溢价率.png"
    rating_group_png = output_dir / "分评级拟合溢价率.png"
    balance_group_png = output_dir / "分余额拟合溢价率.png"
    market_cap_group_png = output_dir / "分正股市值拟合溢价率.png"
    sector_group_png = output_dir / "分板块拟合溢价率.png"
    close_price_distribution_png = output_dir / "收盘价分布面积图.png"
    sector_mean_close_png = output_dir / "各行业平均收盘价.png"
    sector_mean_parity_png = output_dir / "各行业平均平价.png"
    sector_mean_conversion_premium_png = output_dir / "各行业平均转股溢价率.png"
    sector_mean_bond_premium_png = output_dir / "各行业平均纯债溢价率.png"
    industry_performance_png = output_dir / "各行业转债正股涨跌幅及估值.png"
    overview_png = output_dir / "主要指数与市场表现组合图.png"
    fuguo_daily_png = output_dir / "富国日报.png"
    workbook_path = output_dir / f"转债日报市场数据底稿_{run_date:%Y%m%d}.xlsx"
    commentary_path = output_dir / f"转债日报点评_{run_date:%Y%m%d}.txt"
    fuguo_daily_text_path = output_dir / "富国日报.txt"

    report_progress(57, "绘制市场图表")
    font = setup_font()
    plot_market_statistics(market, market_png, font)
    plot_main_money_flow(main_money_flow, main_money_flow_source, main_money_flow_png)
    plot_cb_etf_share_flow(
        etf_share, etf_share_source, "博时可转债ETF", boshi_etf_png
    )
    plot_cb_etf_share_flow(
        etf_share, etf_share_source, "海富通可转债ETF", haifutong_etf_png
    )
    plot_index_turnover(index, index_png, font)
    plot_cb_return_distribution(
        return_distribution, return_summary, run_date, return_png, font
    )
    plot_index_performance_table(
        index_performance, run_date, index_performance_png, font
    )
    report_progress(65, "绘制估值图表")
    plot_daily_valuation(daily_valuation, valuation_source, daily_valuation_png)
    previous_date = pd.Timestamp(valuation_source["previousDate"])
    plot_intraday_valuation(
        intraday_valuation,
        previous_date,
        float(valuation_source["previousValue"]),
        intraday_valuation_png,
    )
    group_valuation_titles = build_group_valuation_titles(
        {
            "parity": parity_group_valuation_source,
            "equity": equity_bond_group_valuation_source["groupDetails"],
            "rating": rating_group_valuation_source["groupDetails"],
            "balance": balance_group_valuation_source["groupDetails"],
            "market_cap": market_cap_group_valuation_source["groupDetails"],
            "sector": sector_group_valuation_source["groupDetails"],
        }
    )
    plot_parity_group_valuation(
        parity_group_valuation,
        parity_group_valuation_source,
        parity_group_valuation_png,
        panel_title=group_valuation_titles["parity"],
    )
    plot_price_parity_series(price_parity, price_parity_source, price_parity_png)
    plot_equity_bond_weighted_premium(
        equity_bond_weighted,
        equity_bond_weighted_source,
        equity_bond_weighted_premium_png,
    )
    plot_valuation_repair_index(valuation_repair, valuation_repair_png)
    maturity_details = maturity_group_valuation_source["groupDetails"]
    maturity_title = (
        "分剩余期限拟合溢价率：\n"
        f"5-6：{float(maturity_details['5-6']['latestValue']):.2f}%，"
        f"{float(maturity_details['5-6']['dailyChangePctPoint']):+.2f}pct；"
        f"0-1：{float(maturity_details['0-1']['latestValue']):.2f}%，"
        f"{float(maturity_details['0-1']['dailyChangePctPoint']):+.2f}pct"
    )
    plot_classification_valuation(
        maturity_group_valuation,
        MATURITY_GROUP_SPECS,
        maturity_title,
        maturity_group_png,
    )
    plot_subnew_bond_metric(
        subnew_bond,
        value_column="次新券平均转股溢价率",
        latest_value=float(subnew_bond_source["latestPremiumMeanPct"]),
        daily_change=float(subnew_bond_source["premiumDailyChangePctPoint"]),
        panel_title_label="次新券平均转股溢价率",
        legend_label="次新券平均转股溢价率（%）",
        output_path=subnew_premium_png,
        include_zero=False,
    )
    plot_classification_valuation(
        equity_bond_group_valuation,
        EQUITY_BOND_GROUP_SPECS,
        group_valuation_titles["equity"],
        equity_bond_group_png,
    )
    plot_classification_valuation(
        rating_group_valuation,
        RATING_GROUP_SPECS,
        group_valuation_titles["rating"],
        rating_group_png,
    )
    plot_classification_valuation(
        balance_group_valuation,
        BALANCE_GROUP_SPECS,
        group_valuation_titles["balance"],
        balance_group_png,
    )
    plot_classification_valuation(
        market_cap_group_valuation,
        MARKET_CAP_GROUP_SPECS,
        group_valuation_titles["market_cap"],
        market_cap_group_png,
    )
    plot_classification_valuation(
        sector_group_valuation,
        SECTOR_GROUP_SPECS,
        group_valuation_titles["sector"],
        sector_group_png,
    )
    plot_close_price_distribution_area(
        close_price_distribution,
        close_price_distribution_source,
        close_price_distribution_png,
    )
    report_progress(76, "绘制行业图表")
    for metric, panel_title, unit_suffix in SECTOR_MEAN_METRICS:
        output_path = {
            "收盘价": sector_mean_close_png,
            "平价": sector_mean_parity_png,
            "转股溢价率": sector_mean_conversion_premium_png,
            "纯债溢价率": sector_mean_bond_premium_png,
        }[metric]
        plot_sector_mean_metric(
            sector_mean_metrics,
            metric,
            panel_title,
            unit_suffix,
            output_path,
        )
    plot_industry_market_table(industry_performance, industry_performance_png)
    report_progress(84, "合成长图")
    compose_index_market_overview(
        index_performance_png,
        index_png,
        return_png,
        main_money_flow_png,
        market_png,
        daily_valuation_png,
        intraday_valuation_png,
        parity_group_valuation_png,
        price_parity_png,
        equity_bond_weighted_premium_png,
        valuation_repair_png,
        maturity_group_png,
        subnew_premium_png,
        equity_bond_group_png,
        rating_group_png,
        balance_group_png,
        market_cap_group_png,
        sector_group_png,
        close_price_distribution_png,
        boshi_etf_png,
        haifutong_etf_png,
        sector_mean_close_png,
        sector_mean_parity_png,
        sector_mean_conversion_premium_png,
        sector_mean_bond_premium_png,
        industry_performance_png,
        overview_png,
        run_date,
    )
    compose_fuguo_daily_report(
        daily_valuation_png,
        intraday_valuation_png,
        valuation_repair_png,
        price_parity_png,
        parity_group_valuation_png,
        equity_bond_weighted_premium_png,
        fuguo_daily_png,
    )
    report_progress(88, "整理独立小图")
    export_numbered_titleless_small_charts(output_dir)
    report_progress(90, "生成 Excel 底稿")
    build_workbook(
        market,
        main_money_flow,
        main_money_flow_source,
        etf_share,
        etf_share_source,
        index,
        index_performance,
        index_performance_source,
        return_details,
        return_distribution,
        return_summary,
        return_source,
        daily_valuation,
        valuation_source,
        valuation_repair,
        intraday_valuation,
        intraday_valuation_source,
        parity_interval_premium,
        parity_group_valuation,
        parity_group_valuation_source,
        equity_bond_group_valuation,
        equity_bond_group_valuation_source,
        rating_group_valuation,
        rating_group_valuation_source,
        maturity_group_valuation,
        maturity_group_valuation_source,
        balance_group_valuation,
        balance_group_valuation_source,
        market_cap_group_valuation,
        market_cap_group_valuation_source,
        sector_group_valuation,
        sector_group_valuation_source,
        sector_mean_metrics,
        sector_mean_source,
        industry_history_metrics,
        industry_performance,
        industry_performance_source,
        close_price_distribution,
        close_price_distribution_source,
        price_parity,
        price_parity_source,
        equity_bond_weighted,
        equity_bond_weighted_source,
        subnew_bond,
        subnew_bond_source,
        run_date,
        start_date,
        market_start_date,
        workbook_path,
    )
    report_progress(96, "生成点评文本")
    commentary = build_daily_commentary(
        run_date,
        index_performance=index_performance,
        turnover=index,
        main_money_flow_source=main_money_flow_source,
        etf_share_source=etf_share_source,
        price_parity_source=price_parity_source,
        equity_bond_weighted_source=equity_bond_weighted_source,
        close_price_distribution=close_price_distribution,
        close_price_distribution_source=close_price_distribution_source,
        valuation_source=valuation_source,
        valuation_repair=valuation_repair,
        subnew_bond_source=subnew_bond_source,
        equity_bond_group_valuation_source=equity_bond_group_valuation_source,
        parity_group_valuation_source=parity_group_valuation_source,
        industry_performance=industry_performance,
        sector_mean_metrics=sector_mean_metrics,
    )
    long_chart_titles = build_long_chart_titles(
        turnover=index,
        return_summary=return_summary,
        main_money_flow_source=main_money_flow_source,
        margin_balance=market,
        valuation_source=valuation_source,
        valuation_repair=valuation_repair,
        price_parity_source=price_parity_source,
        equity_bond_weighted_source=equity_bond_weighted_source,
        group_valuation_titles=group_valuation_titles,
        maturity_title=maturity_title,
        subnew_bond_source=subnew_bond_source,
        close_price_distribution_source=close_price_distribution_source,
        etf_share_source=etf_share_source,
    )
    commentary = append_long_chart_titles(commentary, long_chart_titles)
    write_daily_commentary(commentary, commentary_path)
    fuguo_daily_text = build_fuguo_daily_text(
        run_date,
        index_performance=index_performance,
        turnover=index,
        price_parity_source=price_parity_source,
        valuation_source=valuation_source,
    )
    write_daily_commentary(fuguo_daily_text, fuguo_daily_text_path)
    report_progress(97, "填充 Word 报告")
    word_report_path = build_daily_word_report(
        run_date,
        output_dir,
        index_performance,
        long_chart_titles,
        industry_performance,
        industry_performance_png,
        commentary,
    )
    report_progress(98, "整理输出文件")
    remove_obsolete_outputs(output_dir, workbook_path)

    metadata: dict[str, object] = {
        "运行日期": f"{run_date:%Y-%m-%d}",
        "指数成交额查询开始日期": f"{start_date:%Y-%m-%d}",
        "两融余额查询开始日期": f"{market_start_date:%Y-%m-%d}",
        "市场交易统计记录数": len(market),
        "市场交易统计最新日期": f"{market['交易日期'].max():%Y-%m-%d}",
        "主力净流入数据源": main_money_flow_source["source"],
        "主力净流入公式": main_money_flow_source["formula"],
        "主力净流入记录数": len(main_money_flow),
        "主力净流入最新日期": main_money_flow_source["latestDate"],
        "主力净流入最新值_亿元": main_money_flow_source["latestValue"],
        "主力净流入方向": main_money_flow_source["latestDirection"],
        "ETF份额数据源": etf_share_source["source"],
        "ETF份额公式": etf_share_source["formula"],
        "ETF份额请求截止日期": etf_share_source["requestedEndDate"],
        "ETF份额最新有效日期": etf_share_source["latestDate"],
        "博时可转债ETF最新份额_亿份": etf_share_source["funds"][
            "博时可转债ETF"
        ]["latestShareYi"],
        "博时可转债ETF最新净申赎_亿份": etf_share_source["funds"][
            "博时可转债ETF"
        ]["latestNetSubscriptionYi"],
        "海富通可转债ETF最新份额_亿份": etf_share_source["funds"][
            "海富通可转债ETF"
        ]["latestShareYi"],
        "海富通可转债ETF最新净申赎_亿份": etf_share_source["funds"][
            "海富通可转债ETF"
        ]["latestNetSubscriptionYi"],
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
        "百元拟合溢价率最新值": valuation_source["latestValue"],
        "百元拟合溢价率日变动_pct": valuation_source["dailyChangePctPoint"],
        "百元拟合溢价率2019年以来分位数": valuation_source[
            "percentileSince2019"
        ],
        "日度估值数据源": valuation_source["parquet"],
        "盘中估值数据源": intraday_valuation_source["workbook"],
        "盘中估值时点数": intraday_valuation_source["points"],
        "盘中估值参考线": (
            f"{valuation_source['previousDate']} "
            f"{float(valuation_source['previousValue']):.2f}%"
        ),
        "平价分类溢价率数据源": parity_group_valuation_source["parquet"],
        "平价分类溢价率读取规则": parity_group_valuation_source["readRule"],
        "分平价区间转股溢价率数据源": parity_interval_premium_source[
            "parquetRoot"
        ],
        "分平价区间转股溢价率口径": parity_interval_premium_source[
            "sampleRule"
        ],
        "分平价区间转股溢价率记录数": len(parity_interval_premium),
        "平价分类溢价率当日变动绝对值最大组别": parity_group_valuation_source[
            "largestChangeGroup"
        ],
        "平价分类溢价率最大组别当日变动_pct": parity_group_valuation_source[
            "largestChangePctPoint"
        ],
        "股债型拟合溢价率数据源": equity_bond_group_valuation_source["parquet"],
        "偏股型拟合溢价率": equity_bond_group_valuation_source["groupDetails"][
            "偏股型"
        ]["latestValue"],
        "偏股型拟合溢价率日变动_pct": equity_bond_group_valuation_source[
            "groupDetails"
        ]["偏股型"]["dailyChangePctPoint"],
        "偏债型拟合溢价率": equity_bond_group_valuation_source["groupDetails"][
            "偏债型"
        ]["latestValue"],
        "偏债型拟合溢价率日变动_pct": equity_bond_group_valuation_source[
            "groupDetails"
        ]["偏债型"]["dailyChangePctPoint"],
        "分评级拟合溢价率数据源": rating_group_valuation_source["parquet"],
        "AAA/AA+拟合溢价率": rating_group_valuation_source["groupDetails"][
            "AAA/AA+"
        ]["latestValue"],
        "AAA/AA+拟合溢价率日变动_pct": rating_group_valuation_source[
            "groupDetails"
        ]["AAA/AA+"]["dailyChangePctPoint"],
        "AA/AA-拟合溢价率": rating_group_valuation_source["groupDetails"][
            "AA/AA-"
        ]["latestValue"],
        "AA/AA-拟合溢价率日变动_pct": rating_group_valuation_source[
            "groupDetails"
        ]["AA/AA-"]["dailyChangePctPoint"],
        "分剩余期限拟合溢价率数据源": maturity_group_valuation_source["parquet"],
        "5-6年拟合溢价率": maturity_group_valuation_source["groupDetails"][
            "5-6"
        ]["latestValue"],
        "5-6年拟合溢价率日变动_pct": maturity_group_valuation_source[
            "groupDetails"
        ]["5-6"]["dailyChangePctPoint"],
        "0-1年拟合溢价率": maturity_group_valuation_source["groupDetails"][
            "0-1"
        ]["latestValue"],
        "0-1年拟合溢价率日变动_pct": maturity_group_valuation_source[
            "groupDetails"
        ]["0-1"]["dailyChangePctPoint"],
        "分余额拟合溢价率数据源": balance_group_valuation_source["parquet"],
        "0-3亿元余额拟合溢价率": balance_group_valuation_source["groupDetails"][
            "0-3"
        ]["latestValue"],
        "0-3亿元余额拟合溢价率日变动_pct": balance_group_valuation_source[
            "groupDetails"
        ]["0-3"]["dailyChangePctPoint"],
        "50亿元以上余额拟合溢价率": balance_group_valuation_source[
            "groupDetails"
        ]["50+"]["latestValue"],
        "50亿元以上余额拟合溢价率日变动_pct": balance_group_valuation_source[
            "groupDetails"
        ]["50+"]["dailyChangePctPoint"],
        "分正股市值拟合溢价率数据源": market_cap_group_valuation_source[
            "parquet"
        ],
        "0-50亿元市值拟合溢价率": market_cap_group_valuation_source[
            "groupDetails"
        ]["0-50"]["latestValue"],
        "0-50亿元市值拟合溢价率日变动_pct": market_cap_group_valuation_source[
            "groupDetails"
        ]["0-50"]["dailyChangePctPoint"],
        "300亿元以上市值拟合溢价率": market_cap_group_valuation_source[
            "groupDetails"
        ]["300+"]["latestValue"],
        "300亿元以上市值拟合溢价率日变动_pct": market_cap_group_valuation_source[
            "groupDetails"
        ]["300+"]["dailyChangePctPoint"],
        "分板块拟合溢价率数据源": sector_group_valuation_source["parquet"],
        "科技板块拟合溢价率": sector_group_valuation_source["groupDetails"][
            "科技"
        ]["latestValue"],
        "科技板块拟合溢价率日变动_pct": sector_group_valuation_source[
            "groupDetails"
        ]["科技"]["dailyChangePctPoint"],
        "周期板块拟合溢价率": sector_group_valuation_source["groupDetails"][
            "周期"
        ]["latestValue"],
        "周期板块拟合溢价率日变动_pct": sector_group_valuation_source[
            "groupDetails"
        ]["周期"]["dailyChangePctPoint"],
        "行业均值数据源": sector_mean_source["parquetRoot"],
        "行业板块映射规则": sector_mean_source["sectorRule"],
        "行业均值样本口径": sector_mean_source["sampleRule"],
        "行业均值记录数": len(sector_mean_metrics),
        "行业涨跌与估值记录数": len(industry_performance),
        "行业正股涨跌幅数据源": industry_performance_source["equity"]["source"],
        "行业正股涨跌幅函数": industry_performance_source["equity"][
            "returnFunctions"
        ],
        "行业正股估值函数": industry_performance_source["equity"][
            "valuationFunctions"
        ],
        "行业转债涨跌幅数据源": industry_performance_source["convertibleBond"][
            "parquetRoot"
        ],
        "行业转债涨跌幅口径": industry_performance_source["convertibleBond"][
            "returnRule"
        ],
        "行业涨跌与估值排序规则": industry_performance_source["sortRule"],
        "收盘价分布数据源": close_price_distribution_source["parquetRoot"],
        "收盘价分布样本口径": close_price_distribution_source["sampleRule"],
        "破底占比": close_price_distribution_source["latestBreakFloorPct"],
        "破底占比日变动_pct": close_price_distribution_source[
            "breakFloorDailyChangePctPoint"
        ],
        "破面占比": close_price_distribution_source["latestBreakParPct"],
        "破面占比日变动_pct": close_price_distribution_source[
            "breakParDailyChangePctPoint"
        ],
        "平均平价": price_parity_source["latestParity"],
        "平均平价日变动_pct": price_parity_source["parityDailyChangePct"],
        "整体余额加权收盘价": price_parity_source["latestWeightedClose"],
        "整体余额加权收盘价日变动_pct": price_parity_source[
            "weightedCloseDailyChangePct"
        ],
        "价格中位数": price_parity_source["latestMedianPrice"],
        "价格中位数日变动_pct": price_parity_source[
            "medianPriceDailyChangePct"
        ],
        "价格中位数2019年以来分位数": price_parity_source[
            "medianPricePercentileSince2019"
        ],
        "价格与平价数据源": price_parity_source["parquetRoot"],
        "股债性分类余额加权数据源": equity_bond_weighted_source["parquetRoot"],
        "股债性分类余额加权口径": equity_bond_weighted_source["sampleRule"],
        "偏股型余额加权转股溢价率": equity_bond_weighted_source["premium"][
            "偏股型"
        ]["latestValue"],
        "偏股型余额加权转股溢价率日变动_pct": equity_bond_weighted_source[
            "premium"
        ]["偏股型"]["dailyChangePctPoint"],
        "偏债型余额加权转股溢价率": equity_bond_weighted_source["premium"][
            "偏债型"
        ]["latestValue"],
        "偏债型余额加权转股溢价率日变动_pct": equity_bond_weighted_source[
            "premium"
        ]["偏债型"]["dailyChangePctPoint"],
        "偏股型余额加权均价": equity_bond_weighted_source["price"]["偏股型"][
            "latestValue"
        ],
        "偏股型余额加权均价日变动_pct": equity_bond_weighted_source["price"][
            "偏股型"
        ]["dailyChangePct"],
        "偏债型余额加权均价": equity_bond_weighted_source["price"]["偏债型"][
            "latestValue"
        ],
        "偏债型余额加权均价日变动_pct": equity_bond_weighted_source["price"][
            "偏债型"
        ]["dailyChangePct"],
        "次新券数据源": subnew_bond_source["parquetRoot"],
        "次新券样本口径": subnew_bond_source["sampleRule"],
        "次新券相对上市涨跌幅均值": subnew_bond_source[
            "latestListingReturnMeanPct"
        ],
        "次新券相对上市涨跌幅日变动_pct": subnew_bond_source[
            "listingReturnDailyChangePctPoint"
        ],
        "次新券平均转股溢价率": subnew_bond_source[
            "latestPremiumMeanPct"
        ],
        "次新券平均转股溢价率日变动_pct": subnew_bond_source[
            "premiumDailyChangePctPoint"
        ],
        "报告字体": font.get_name(),
        "报告配色": {"红": RED, "蓝": BLUE},
        "Excel底稿": str(workbook_path),
        "日报点评": str(commentary_path),
        "Word报告": str(word_report_path),
        "富国日报图片": str(fuguo_daily_png),
        "富国日报点评": str(fuguo_daily_text_path),
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
    progress = ConsoleProgress()
    try:
        run_date = (
            datetime.strptime(args.run_date, "%Y-%m-%d").date()
            if args.run_date
            else latest_cb_trade_date()
        )
        output_dir = args.output_dir or (
            WORKSPACE / "runs" / "daily" / f"{run_date:%Y%m%d}_转债日报"
        )
        run(run_date, output_dir.resolve(), progress=progress)
    except Exception:
        progress.fail()
        raise
    progress.finish(f"完成：{output_dir.name}")


if __name__ == "__main__":
    main()
