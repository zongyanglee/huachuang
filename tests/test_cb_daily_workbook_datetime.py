from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_workbook_datetime", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class DailyWorkbookDatetimeTests(unittest.TestCase):
    @staticmethod
    def _build_workbook(payload: dict[str, object], directory: str) -> tuple[subprocess.CompletedProcess[str], Path]:
        temp_dir = Path(directory)
        payload_path = temp_dir / "payload.json"
        builder_path = temp_dir / "builder.mjs"
        output_path = temp_dir / "output.xlsx"
        payload_path.write_text(json.dumps(payload), encoding="utf-8")
        builder_path.write_text(MODULE.WORKBOOK_BUILDER_SOURCE, encoding="utf-8")
        subprocess.run(
            [
                "cmd.exe",
                "/c",
                "mklink",
                "/J",
                str(temp_dir / "node_modules"),
                str(MODULE.BUNDLED_NODE_MODULES),
            ],
            check=True,
            capture_output=True,
        )
        result = subprocess.run(
            [str(MODULE.BUNDLED_NODE), str(builder_path), str(payload_path), str(output_path)],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode == 0:
            MODULE.enforce_industry_history_workbook_layout(output_path)
        return result, output_path

    @staticmethod
    def _empty_payload() -> dict[str, object]:
        series_keys = [
            "market",
            "mainMoneyFlow",
            "etfShare",
            "index",
            "returnDistribution",
            "indexPerformance",
            "valuationDaily",
            "valuationRepair",
            "valuationIntraday",
            "parityIntervalPremium",
            "parityGroupValuation",
            "priceParity",
            "equityBondWeighted",
            "maturityGroupValuation",
            "subnewBond",
            "equityBondGroupValuation",
            "ratingGroupValuation",
            "balanceGroupValuation",
            "marketCapGroupValuation",
            "sectorGroupValuation",
            "closePriceDistribution",
            "sectorMeanMetrics",
            "industryHistoryMetrics",
            "industryPerformance",
        ]
        return {key: [] for key in series_keys}

    def test_intraday_iso_datetime_exports_without_invalid_date(self) -> None:
        """盘中估值的 ISO 日期时间不能在底稿检查阶段变为 Invalid Date。"""
        payload = self._empty_payload()
        payload["valuationIntraday"] = [
            {"datetime": "2026-08-28T14:57:00", "premium": 23.45}
        ]

        with tempfile.TemporaryDirectory() as directory:
            result, output_path = self._build_workbook(payload, directory)

            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertTrue(output_path.is_file())

    def test_sheet_order_follows_long_report_from_top_left_to_bottom_right(self) -> None:
        """底稿按长图阅读顺序先上后下、同排先左后右。"""
        expected_sheet_names = [
            "指数表现", "成交额", "涨跌幅分布", "主力净流入", "两融余额",
            "百元拟合溢价率", "盘中百元平价拟合溢价率", "转债估值修复指数",
            "余额加权平价与收盘价中位数", "股债性分类转股溢价率",
            "分平价区间转股溢价率", "平价分类拟合溢价率",
            "分剩余期限拟合溢价率", "次新券平均转股溢价率", "股债型拟合溢价率",
            "分评级拟合溢价率", "分余额拟合溢价率", "分正股市值拟合溢价率",
            "分板块拟合溢价率", "收盘价分布", "博时ETF份额", "海富通ETF份额",
            "各行业平均收盘价", "各行业平均平价", "各行业平均转股溢价率",
            "各行业平均纯债溢价率", "行业涨跌与估值", "股债性分类均价",
            "行业收盘价历史", "行业平价历史", "行业转股溢价率历史",
            "行业纯债溢价率历史",
        ]

        with tempfile.TemporaryDirectory() as directory:
            result, output_path = self._build_workbook(self._empty_payload(), directory)

            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            with zipfile.ZipFile(output_path) as workbook:
                root = ET.fromstring(workbook.read("xl/workbook.xml"))
            namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            actual_sheet_names = [
                sheet.attrib["name"] for sheet in root.findall("main:sheets/main:sheet", namespace)
            ]

        self.assertEqual(actual_sheet_names, expected_sheet_names)

    def test_industry_history_sheets_are_last_and_use_industry_by_date_layout(self) -> None:
        """行业历史序列追加在末尾，日期升序、两位小数并冻结首行首列。"""
        payload = self._empty_payload()
        payload["industryHistoryOrder"] = list(MODULE.INDUSTRY_HISTORY_ORDER)
        payload["industryHistoryMetrics"] = [
            {
                "industry": "国防军工",
                "date": "2026-08-28",
                "close": 126.5,
                "conversionPremium": 31.25,
                "parity": 101.75,
                "bondPremium": 42.0,
            },
            {
                "industry": "国防军工",
                "date": "2019-01-02",
                "close": 125.0,
                "conversionPremium": 30.5,
                "parity": 101.0,
                "bondPremium": 41.5,
            }
        ]

        with tempfile.TemporaryDirectory() as directory:
            result, output_path = self._build_workbook(payload, directory)

            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            with zipfile.ZipFile(output_path) as workbook:
                workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
                relationships_root = ET.fromstring(
                    workbook.read("xl/_rels/workbook.xml.rels")
                )
                namespace = {
                    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
                }
                relationships = {
                    rel.attrib["Id"]: rel.attrib["Target"]
                    for rel in relationships_root.findall("pkg:Relationship", namespace)
                }
                sheets = workbook_root.findall("main:sheets/main:sheet", namespace)
                last_four = sheets[-4:]
                actual_names = [sheet.attrib["name"] for sheet in last_four]
                close_sheet_target = relationships[last_four[0].attrib[f"{{{namespace['rel']}}}id"]]
                close_sheet_member = close_sheet_target.lstrip("/")
                if not close_sheet_member.startswith("xl/"):
                    close_sheet_member = f"xl/{close_sheet_member}"
                close_sheet_xml = ET.fromstring(
                    workbook.read(close_sheet_member)
                )
                styles_root = ET.fromstring(workbook.read("xl/styles.xml"))
                shared_strings = []
                if "xl/sharedStrings.xml" in workbook.namelist():
                    shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
                    shared_strings = [
                        "".join(node.text or "" for node in item.findall(".//main:t", namespace))
                        for item in shared_root.findall("main:si", namespace)
                    ]

            cells: dict[str, str] = {}
            for cell in close_sheet_xml.findall(".//main:c", namespace):
                address = cell.attrib["r"]
                if cell.attrib.get("t") == "inlineStr":
                    cells[address] = "".join(
                        node.text or "" for node in cell.findall(".//main:t", namespace)
                    )
                    continue
                value_node = cell.find("main:v", namespace)
                if value_node is None:
                    continue
                value = value_node.text or ""
                cells[address] = shared_strings[int(value)] if cell.attrib.get("t") == "s" else value

        self.assertEqual(
            actual_names,
            ["行业收盘价历史", "行业平价历史", "行业转股溢价率历史", "行业纯债溢价率历史"],
        )
        self.assertEqual(cells["A1"], "行业")
        self.assertEqual(cells["B1"], "2019-01-02")
        self.assertEqual(cells["C1"], "2026-08-28")
        self.assertEqual(cells["A2"], "农林牧渔")
        defense_row = 2 + list(MODULE.INDUSTRY_HISTORY_ORDER).index("国防军工")
        self.assertEqual(cells[f"A{defense_row}"], "国防军工")
        self.assertEqual(float(cells[f"B{defense_row}"]), 125.0)
        self.assertEqual(float(cells[f"C{defense_row}"]), 126.5)

        pane = close_sheet_xml.find(".//main:sheetView/main:pane", namespace)
        self.assertIsNotNone(pane)
        self.assertEqual(pane.attrib.get("state"), "frozen")
        self.assertEqual(pane.attrib.get("xSplit"), "1")
        self.assertEqual(pane.attrib.get("ySplit"), "1")

        numeric_cell = close_sheet_xml.find(f".//main:c[@r='B{defense_row}']", namespace)
        self.assertIsNotNone(numeric_cell)
        style_index = int(numeric_cell.attrib["s"])
        cell_xfs = styles_root.find("main:cellXfs", namespace)
        self.assertIsNotNone(cell_xfs)
        num_fmt_id = int(list(cell_xfs)[style_index].attrib["numFmtId"])
        custom_formats = {
            int(node.attrib["numFmtId"]): node.attrib["formatCode"]
            for node in styles_root.findall("main:numFmts/main:numFmt", namespace)
        }
        format_code = custom_formats.get(num_fmt_id, {2: "0.00"}.get(num_fmt_id))
        self.assertEqual(format_code, "0.00")

    def test_first_sheet_uses_side_by_side_market_index_layout(self) -> None:
        """首个Sheet应按图表1格式左右并列主要指数与风格指数。"""
        payload = self._empty_payload()

        def index_row(group: str, name: str, close: float) -> dict[str, object]:
            return {
                "group": group,
                "code": name,
                "name": name,
                "date": "2026-08-28",
                "close": close,
                "dailyBaseClose": 100.0,
                "weekBaseClose": 100.0,
                "monthBaseClose": 100.0,
                "yearBaseClose": 100.0,
            }

        payload["indexPerformance"] = [
            index_row("主要指数", name, 101.0 + position)
            for position, name in enumerate(
                [
                    "中证转债", "转债等权", "正股等权", "转债预案", "上证综指",
                    "深证成指", "创业板指", "上证50", "中证1000",
                ]
            )
        ] + [
            index_row("风格指数", name, 201.0 + position)
            for position, name in enumerate(
                [
                    "大盘指数", "中盘指数", "小盘指数", "大盘成长", "大盘价值",
                    "中盘成长", "中盘价值", "小盘成长", "小盘价值",
                ]
            )
        ]

        with tempfile.TemporaryDirectory() as directory:
            result, output_path = self._build_workbook(payload, directory)

            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            with zipfile.ZipFile(output_path) as workbook:
                sheet_xml = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
                shared_strings = []
                if "xl/sharedStrings.xml" in workbook.namelist():
                    shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
                    namespace = {
                        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    }
                    shared_strings = [
                        "".join(node.text or "" for node in item.findall(".//main:t", namespace))
                        for item in shared_root.findall("main:si", namespace)
                    ]

            namespace = {
                "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            }
            cells: dict[str, str] = {}
            for cell in sheet_xml.findall(".//main:c", namespace):
                address = cell.attrib["r"]
                cell_type = cell.attrib.get("t")
                if cell_type == "inlineStr":
                    cells[address] = "".join(
                        node.text or ""
                        for node in cell.findall(".//main:t", namespace)
                    )
                    continue
                value_node = cell.find("main:v", namespace)
                if value_node is None:
                    continue
                value = value_node.text or ""
                cells[address] = (
                    shared_strings[int(value)] if cell_type == "s" else value
                )

        self.assertEqual(cells["A1"], "图表 1  主要市场指数")
        self.assertEqual(cells["A2"], "主要指数")
        self.assertEqual(cells["G2"], "主要指数")
        self.assertEqual(cells["A3"], "中证转债")
        self.assertEqual(cells["G3"], "大盘指数")
        self.assertEqual(cells["A12"], "资料来源：Wind，华创证券")


if __name__ == "__main__":
    unittest.main()
