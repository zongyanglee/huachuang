import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = process.cwd();
const inputDir = path.join(root, "outputs", "rating_nav_202406_202408");
const outputPath = path.join(inputDir, "2024年6-8月转债评级净值与评级错位分析_剔除天创华锋.xlsx");
const report = JSON.parse(await fs.readFile(path.join(inputDir, "report.json"), "utf8"));

const wb = Workbook.create();
const summarySheet = wb.worksheets.add("结论摘要");
const navSheet = wb.worksheets.add("评级净值曲线");
const sampleSheet = wb.worksheets.add("典型样本");
const detailSheet = wb.worksheets.add("个券明细");
const sourceSheet = wb.worksheets.add("口径与来源");

const navy = "#17365D";
const blue = "#2F75B5";
const teal = "#2A7F79";
const lightBlue = "#DDEBF7";
const lightTeal = "#DDEFEA";
const lightRed = "#FCE4D6";
const gray = "#F2F2F2";
const border = "#D9E1F2";
const text = "#1F2937";
const red = "#C00000";
const green = "#008000";

function applyBase(sheet) {
  sheet.showGridLines = false;
}

function styleTitle(sheet, range, title) {
  sheet.getRange(range).merge();
  sheet.getRange(range).values = [[title]];
  sheet.getRange(range).format = {
    fill: navy,
    font: { name: "Microsoft YaHei", size: 16, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    rowHeight: 30,
  };
}

function styleHeader(range, fill = blue) {
  range.format = {
    fill,
    font: { name: "Microsoft YaHei", size: 10, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: border },
    rowHeight: 28,
  };
}

function styleTableBody(range) {
  range.format = {
    font: { name: "Microsoft YaHei", size: 10, color: text },
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: border },
  };
}

function valueByRating(rating, field) {
  return report.summary.find((row) => row["评级"] === rating)?.[field];
}

// Conclusion sheet
styleTitle(summarySheet, "A1:I1", "2024年6-8月转债评级净值与评级错位分析（剔除天创、华锋）");
summarySheet.getRange("A3:I3").merge();
summarySheet.getRange("A3:I3").values = [[
  `区间：${report.start}至${report.end}，共${report.trading_days}个交易日；组合按每日正常交易个券等权，剔除上市首日涨跌幅及天创、华锋转债。`
]];
summarySheet.getRange("A3:I3").format = {
  fill: gray,
  font: { name: "Microsoft YaHei", size: 10, italic: true, color: "#44546A" },
  verticalAlignment: "center",
  rowHeight: 24,
};

summarySheet.getRange("A5:I5").merge();
summarySheet.getRange("A5:I5").values = [["核心判断"]];
summarySheet.getRange("A5:I5").format = {
  fill: teal,
  font: { name: "Microsoft YaHei", size: 11, bold: true, color: "#FFFFFF" },
  rowHeight: 24,
};
const conclusions = [
  ["1", `评级与抗跌性并非单调对应：AA+组合下跌${Math.abs(valueByRating("AA+", "区间收益") * 100).toFixed(1)}%，弱于AAA的${Math.abs(valueByRating("AAA", "区间收益") * 100).toFixed(1)}%；AA、AA-、A+和A也下跌约7.6%-9.3%。`],
  ["2", "真正明显的是个券错位：部分AAA/AA+个券跌幅超过15%，同时若干AA及以下小规模券仅小幅回撤。评级不能替代对价格、行业、股性、条款和资产负债表的分析。"],
  ["3", "小规模发行人可能因主体体量、融资渠道或评级方法限制难获高评级；若负债率较低、偿债压力可控，仍可能具备较好的债底和抗跌性。"],
  ["4", `限定：这不是“低评级整体更安全”。剔除天创、华锋后A-收益为${(valueByRating("A-", "区间收益") * 100).toFixed(1)}%，但平均仅约${valueByRating("A-", "平均持仓数").toFixed(1)}只，仍存在明显小样本和尾部风险。`],
];
summarySheet.getRange(`A6:I${5 + conclusions.length}`).values = conclusions.map(([n, s]) => [n, s, null, null, null, null, null, null, null]);
for (let i = 0; i < conclusions.length; i += 1) {
  const row = 6 + i;
  summarySheet.getRange(`B${row}:I${row}`).merge();
  summarySheet.getRange(`A${row}:I${row}`).format = {
    fill: i === 3 ? lightRed : "#FFFFFF",
    font: { name: "Microsoft YaHei", size: 10, color: i === 3 ? "#9C0006" : text },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: border },
    rowHeight: 38,
  };
  summarySheet.getRange(`A${row}`).format.font = { name: "Microsoft YaHei", size: 11, bold: true, color: navy };
  summarySheet.getRange(`A${row}`).format.horizontalAlignment = "center";
}

const summaryHeaders = ["评级", "期末净值", "区间收益", "最大回撤", "日波动率", "平均持仓数", "期初持仓数", "期末持仓数"];
summarySheet.getRange("A12:H12").values = [summaryHeaders];
styleHeader(summarySheet.getRange("A12:H12"));
const summaryRows = report.summary.map((row) => summaryHeaders.map((key) => row[key]));
summarySheet.getRange(`A13:H${12 + summaryRows.length}`).values = summaryRows;
styleTableBody(summarySheet.getRange(`A13:H${12 + summaryRows.length}`));
summarySheet.getRange(`B13:E${12 + summaryRows.length}`).format.numberFormat = "0.0%;[Red](0.0%);-";
summarySheet.getRange(`F13:H${12 + summaryRows.length}`).format.numberFormat = "0.0";
summarySheet.getRange(`C13:D${12 + summaryRows.length}`).conditionalFormats.add("colorScale", {
  colors: ["#F8696B", "#FFEB84", "#63BE7B"],
  thresholds: ["min", "50%", "max"],
});

summarySheet.getRange("A22:I22").merge();
summarySheet.getRange("A22:I22").values = [["高评级大跌与低评级抗跌样本对照"]];
summarySheet.getRange("A22:I22").format = {
  fill: teal,
  font: { name: "Microsoft YaHei", size: 11, bold: true, color: "#FFFFFF" },
  rowHeight: 24,
};
const compareHeaders = ["类型", "转债名称", "评级", "区间收益", "最大回撤", "发行规模(亿元)", "期初余额(亿元)", "2023资产负债率", "说明"];
summarySheet.getRange("A23:I23").values = [compareHeaders];
styleHeader(summarySheet.getRange("A23:I23"));
const highCodes = ["127089.SZ", "113066.SH", "127049.SZ", "110081.SH", "118034.SH"];
const highRows = report.detail.filter((row) => highCodes.includes(row["代码"])).sort((a, b) => a["区间收益"] - b["区间收益"]);
const defensiveRows = report.defensive;
const compareRows = [
  ...highRows.map((row) => ["高评级大跌", row["转债名称"], row["期初评级"], row["区间收益"], row["最大回撤"], row["发行规模"], row["期初余额"], null, "评级高，但行业/正股与估值下行仍可造成明显损失"]),
  ...defensiveRows.map((row) => ["低评级抗跌", row["转债名称"], row["期初评级"], row["区间收益"], row["最大回撤"], row["发行规模"], row["期初余额"], row["2023资产负债率"], "规模较小、评级不高，但财务杠杆较低且区间回撤较浅"]),
];
summarySheet.getRange(`A24:I${23 + compareRows.length}`).values = compareRows;
styleTableBody(summarySheet.getRange(`A24:I${23 + compareRows.length}`));
summarySheet.getRange(`D24:E${23 + compareRows.length}`).format.numberFormat = "0.0%;[Red](0.0%);-";
summarySheet.getRange(`F24:G${23 + compareRows.length}`).format.numberFormat = "0.00";
summarySheet.getRange(`H24:H${23 + compareRows.length}`).format.numberFormat = "0.0%";
summarySheet.getRange(`A24:A${23 + highRows.length}`).format.fill = lightRed;
summarySheet.getRange(`A${24 + highRows.length}:A${23 + compareRows.length}`).format.fill = lightTeal;
summarySheet.getRange(`I24:I${23 + compareRows.length}`).format.wrapText = true;

summarySheet.getRange("A:A").format.columnWidth = 13;
summarySheet.getRange("B:B").format.columnWidth = 15;
summarySheet.getRange("C:C").format.columnWidth = 10;
summarySheet.getRange("D:E").format.columnWidth = 12;
summarySheet.getRange("F:H").format.columnWidth = 14;
summarySheet.getRange("I:I").format.columnWidth = 38;
summarySheet.freezePanes.freezeRows(3);

// NAV sheet
styleTitle(navSheet, "A1:H1", "评级分类等权净值曲线");
navSheet.getRange("A3:H3").values = [["日期", "AAA", "AA+", "AA", "AA-", "A+", "A", "A-"]];
styleHeader(navSheet.getRange("A3:H3"));
const navRows = report.nav.map((row) => [
  String(row["日期"]).slice(0, 10),
  row["AAA"], row["AA+"], row["AA"], row["AA-"], row["A+"], row["A"], row["A-"],
]);
navSheet.getRange(`A4:H${3 + navRows.length}`).values = navRows;
styleTableBody(navSheet.getRange(`A4:H${3 + navRows.length}`));
navSheet.getRange(`B4:H${3 + navRows.length}`).format.numberFormat = "0.0000";
navSheet.getRange("A:A").format.columnWidth = 13;
navSheet.getRange("B:H").format.columnWidth = 11;
navSheet.freezePanes.freezeRows(3);

const sampledNavRows = report.nav
  .filter((_, index) => index % 4 === 0 || index === report.nav.length - 1)
  .map((row) => [
    String(row["日期"]).slice(5, 10),
    row["AAA"] - 1, row["AA+"] - 1, row["AA"] - 1, row["AA-"] - 1,
    row["A+"] - 1, row["A"] - 1, row["A-"] - 1,
  ]);
navSheet.getRange("J26:Q26").values = [["日期", "AAA", "AA+", "AA", "AA-", "A+", "A", "A-"]];
navSheet.getRange(`J27:Q${26 + sampledNavRows.length}`).values = sampledNavRows;
const navChart = navSheet.charts.add("line", navSheet.getRange(`J26:Q${26 + sampledNavRows.length}`));
navChart.title = "评级等权组合累计收益（剔除天创、华锋）";
navChart.hasLegend = true;
navChart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 9 } };
navChart.yAxis = { numberFormatCode: "0%" };
navChart.setPosition("J3", "T23");
const seriesColors = ["#1F4E78", "#5B9BD5", "#70AD47", "#A5A5A5", "#ED7D31", "#C00000", "#7030A0"];
navChart.series.items.forEach((series, index) => {
  series.line = { color: seriesColors[index], width: 2 };
});

// Samples sheet
styleTitle(sampleSheet, "A1:L1", "典型评级错位样本");
sampleSheet.getRange("A3:L3").values = [[
  "代码", "转债名称", "期初评级", "期初收盘价", "期末收盘价", "区间收益", "最大回撤",
  "期初余额(亿元)", "发行规模(亿元)", "期初正股市值(亿元)", "2023资产负债率", "样本类型",
]];
styleHeader(sampleSheet.getRange("A3:L3"));
const sampleRows = [
  ...highRows.map((row) => [row["代码"], row["转债名称"], row["期初评级"], row["期初收盘价"], row["期末收盘价"], row["区间收益"], row["最大回撤"], row["期初余额"], row["发行规模"], row["期初正股市值"], null, "高评级大跌"]),
  ...defensiveRows.map((row) => [row["代码"], row["转债名称"], row["期初评级"], row["期初收盘价"], row["期末收盘价"], row["区间收益"], row["最大回撤"], row["期初余额"], row["发行规模"], row["期初正股市值"], row["2023资产负债率"], "低评级抗跌"]),
];
sampleSheet.getRange(`A4:L${3 + sampleRows.length}`).values = sampleRows;
styleTableBody(sampleSheet.getRange(`A4:L${3 + sampleRows.length}`));
sampleSheet.getRange(`D4:E${3 + sampleRows.length}`).format.numberFormat = "0.000";
sampleSheet.getRange(`F4:G${3 + sampleRows.length}`).format.numberFormat = "0.0%;[Red](0.0%);-";
sampleSheet.getRange(`H4:J${3 + sampleRows.length}`).format.numberFormat = "0.00";
sampleSheet.getRange(`K4:K${3 + sampleRows.length}`).format.numberFormat = "0.0%";
sampleSheet.getRange(`A4:L${3 + highRows.length}`).format.fill = "#FFF2F2";
sampleSheet.getRange(`A${4 + highRows.length}:L${3 + sampleRows.length}`).format.fill = "#F0F8F5";
sampleSheet.getRange("A:A").format.columnWidth = 13;
sampleSheet.getRange("B:B").format.columnWidth = 14;
sampleSheet.getRange("C:C").format.columnWidth = 9;
sampleSheet.getRange("D:K").format.columnWidth = 14;
sampleSheet.getRange("L:L").format.columnWidth = 14;
sampleSheet.freezePanes.freezeRows(3);

// Detail sheet
styleTitle(detailSheet, "A1:L1", "全量个券区间表现");
const detailHeaders = ["代码", "转债名称", "期初评级", "首个有效日", "末个有效日", "期初收盘价", "期末收盘价", "区间收益", "最大回撤", "期初余额", "发行规模", "期初正股市值"];
detailSheet.getRange("A3:L3").values = [detailHeaders];
styleHeader(detailSheet.getRange("A3:L3"));
const detailRows = report.detail.map((row) => detailHeaders.map((key) => row[key]));
detailSheet.getRange(`A4:L${3 + detailRows.length}`).values = detailRows;
styleTableBody(detailSheet.getRange(`A4:L${3 + detailRows.length}`));
detailSheet.getRange(`F4:G${3 + detailRows.length}`).format.numberFormat = "0.000";
detailSheet.getRange(`H4:I${3 + detailRows.length}`).format.numberFormat = "0.0%;[Red](0.0%);-";
detailSheet.getRange(`J4:L${3 + detailRows.length}`).format.numberFormat = "0.00";
detailSheet.tables.add(`A3:L${3 + detailRows.length}`, true, "BondDetailTable").style = "TableStyleMedium2";
detailSheet.getRange("A:A").format.columnWidth = 13;
detailSheet.getRange("B:B").format.columnWidth = 14;
detailSheet.getRange("C:C").format.columnWidth = 9;
detailSheet.getRange("D:E").format.columnWidth = 13;
detailSheet.getRange("F:L").format.columnWidth = 14;
detailSheet.freezePanes.freezeRows(3);

// Sources and methodology
styleTitle(sourceSheet, "A1:F1", "口径、限制与数据来源");
sourceSheet.getRange("A3:F3").values = [["项目", "口径/数值", "单位", "期间/时点", "来源", "备注"]];
styleHeader(sourceSheet.getRange("A3:F3"));
const eastmoneyUrl = (code) =>
  `https://datacenter.eastmoney.com/securities/api/data/v1/get?reportName=RPT_F10_FINANCE_MAINFINADATA&columns=ALL&filter=(SECUCODE%3D%22${code}%22)&pageNumber=1&pageSize=30&sortTypes=-1&sortColumns=REPORT_DATE`;
const sourceRows = [
  ["评级组合净值", "每日按债项评级分组，仅对状态为“交易”且收盘价、涨跌幅有效的个券等权", "净值", `${report.start}至${report.end}`, "工作区历史Parquet", "剔除上市首日涨跌幅及天创、华锋转债；新券自下一正常交易日纳入；每日动态分组；期初净值=1"],
  ["个券区间收益", "期末有效收盘价/期初有效收盘价-1", "%", `${report.start}至${report.end}`, "工作区历史Parquet", "要求首尾一周内均有有效交易数据"],
  ["个券最大回撤", "区间有效收盘价/此前最高收盘价-1的最小值", "%", `${report.start}至${report.end}`, "工作区历史Parquet", "用于衡量路径中的最深跌幅"],
  ["长久物流资产负债率", 0.4318989174, "%", "2023年报", eastmoneyUrl("603569.SH"), "字段ZCFZL"],
  ["奇精机械资产负债率", 0.460445982208, "%", "2023年报", eastmoneyUrl("603677.SH"), "字段ZCFZL"],
  ["好莱客资产负债率", 0.319446369445, "%", "2023年报", eastmoneyUrl("603898.SH"), "字段ZCFZL"],
  ["结论限制", "评级组表现受行业、价格、平价、转股溢价率、流动性和异常交易影响", null, "分析区间", "本分析", "不能将个券反例外推为低评级整体更安全"],
];
sourceSheet.getRange(`A4:F${3 + sourceRows.length}`).values = sourceRows;
styleTableBody(sourceSheet.getRange(`A4:F${3 + sourceRows.length}`));
sourceSheet.getRange(`B7:B9`).format.numberFormat = "0.0%";
sourceSheet.getRange(`B4:B${3 + sourceRows.length}`).format.wrapText = true;
sourceSheet.getRange(`E4:E${3 + sourceRows.length}`).format = {
  font: { name: "Microsoft YaHei", size: 9, color: green },
  wrapText: true,
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: border },
};
sourceSheet.getRange("A:A").format.columnWidth = 23;
sourceSheet.getRange("B:B").format.columnWidth = 48;
sourceSheet.getRange("C:C").format.columnWidth = 10;
sourceSheet.getRange("D:D").format.columnWidth = 21;
sourceSheet.getRange("E:E").format.columnWidth = 55;
sourceSheet.getRange("F:F").format.columnWidth = 38;
sourceSheet.freezePanes.freezeRows(3);

for (const sheet of [summarySheet, navSheet, sampleSheet, detailSheet, sourceSheet]) {
  applyBase(sheet);
}

const summaryPreview = await wb.render({
  sheetName: "结论摘要",
  range: `A1:I${23 + compareRows.length}`,
  scale: 1.3,
  format: "png",
});
await fs.writeFile(path.join(inputDir, "结论摘要预览.png"), new Uint8Array(await summaryPreview.arrayBuffer()));

const navPreview = await wb.render({
  sheetName: "评级净值曲线",
  range: "A1:T23",
  scale: 1.2,
  format: "png",
});
await fs.writeFile(path.join(inputDir, "评级净值曲线预览.png"), new Uint8Array(await navPreview.arrayBuffer()));

const inspect = await wb.inspect({
  kind: "table",
  range: "结论摘要!A12:I31",
  include: "values,formulas",
  tableMaxRows: 25,
  tableMaxCols: 10,
});
console.log(inspect.ndjson);
const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
