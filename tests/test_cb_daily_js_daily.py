from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

import pandas as pd


SCRIPT = Path(__file__).resolve().parents[1] / "src" / "daily" / "【日报】转债日报.py"
SPEC = importlib.util.spec_from_file_location("cb_daily_js_daily", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


EXPECTED_COLUMNS = {
    "平价分类转股溢价率": [
        "日期",
        "平价80以下",
        "平价80-95",
        "平价95-110",
        "平价110-125",
        "平价125以上",
    ],
    "收盘价分位数统计": ["日期", "0.05", "0.25", "0.5", "0.75", "0.8", "0.9", "均值"],
    "百元溢价率": [
        "日期",
        "百元平价拟合溢价率",
        "百元平价拟合溢价率（1-5.5年）",
        "次新券转股溢价率均值",
    ],
    "JS全样本": ["日期", "平价算术平均", "纯债价值算术平均", "纯债溢价率算术平均", "隐含波动率算术平均"],
    "JS平底分类余额加权转股溢价率": ["日期", "偏股", "偏债", "平衡"],
    "JS偏债型余额YTM": ["日期", "YTM余额加权", "YTM中位数", "AAA评级YTM中位数"],
    "板块转股溢价率": ["日期", "周期", "制造", "科技", "消费", "金融"],
    "板块价格": ["日期", "周期", "制造", "科技", "消费", "金融"],
    "板块平价": ["日期", "周期", "制造", "科技", "消费", "金融"],
    "板块纯债溢价率": ["日期", "周期", "制造", "科技", "消费", "金融"],
}


def make_daily_rows(trade_date: str) -> pd.DataFrame:
    rows = [
        # 高价且高溢价的旧“剔妖”样本必须保留。
        (trade_date, "交易", 1.0, 80.0, 90.0, 10.0, 5.0, 100.0, 30.0, -5.0, 2.0, 200.0, "AAA", 10.0),
        # 剩余期限不足一年：仅在平底分类和偏债型 YTM 中剔除。
        (trade_date, "交易", 3.0, 90.0, 100.0, 20.0, 10.0, 20.0, 30.0, -2.0, 0.5, 100.0, "AAA", 10.0),
        (trade_date, "交易", 2.0, 100.0, 110.0, 30.0, 15.0, 30.0, 0.0, -3.0, 2.0, 120.0, "AA+", 10.0),
        (trade_date, "交易", 4.0, 120.0, 120.0, 40.0, 20.0, 40.0, -30.0, 4.0, 2.0, 130.0, "AAA", 10.0),
        (trade_date, "交易", 5.0, 130.0, 130.0, 50.0, 25.0, 50.0, -30.0, 6.0, 0.5, 140.0, "AAA", 10.0),
        (trade_date, "未上市", 100.0, 999.0, 999.0, 999.0, 999.0, 999.0, -30.0, 999.0, 2.0, 999.0, "AAA", 10.0),
    ]
    return pd.DataFrame(
        rows,
        columns=[
            "交易日期",
            "交易状态",
            "余额",
            "平价",
            "纯债价值",
            "纯债溢价率",
            "隐含波动率",
            "转股溢价率",
            "平价底价溢价率",
            "YTM",
            "剩余期限",
            "收盘价",
            "债项评级",
            "换手率",
        ],
    )


def make_inverse_cubic_rows() -> pd.DataFrame:
    rows: list[tuple[object, ...]] = []
    for plain in range(70, 131, 5):
        rows.append(
            (
                "2026-08-28",
                "交易",
                float(plain),
                10000.0 / plain - 50.0,
                10.0,
                2.0,
            )
        )
    for plain in range(75, 116, 5):
        rows.append(
            (
                "2026-08-29",
                "交易",
                float(plain),
                6000.0 / plain - 30.0,
                10.0,
                2.0,
            )
        )
    for plain in (70, 72, 120, 125, 128, 130):
        rows.append(
            (
                "2026-08-29",
                "交易",
                float(plain),
                10000.0 / plain - 50.0,
                10.0,
                6.0,
            )
        )
    rows.extend(
        [
            ("2026-08-29", "交易", 100.0, 999.0, 50.1, 2.0),
            ("2026-08-29", "未上市", 100.0, 999.0, 10.0, 2.0),
        ]
    )
    return pd.DataFrame(
        rows,
        columns=["交易日期", "交易状态", "平价", "转股溢价率", "换手率", "剩余期限"],
    )


def make_sector_mean_rows() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "交易日期": "2026-08-28",
                "转股溢价率_周期": 10.1,
                "转股溢价率_制造": 10.2,
                "转股溢价率_科技": 10.3,
                "转股溢价率_消费": 10.4,
                "转股溢价率_金融": 10.5,
                "收盘价_周期": 20.1,
                "收盘价_制造": 20.2,
                "收盘价_科技": 20.3,
                "收盘价_消费": 20.4,
                "收盘价_金融": 20.5,
                "平价_周期": 30.1,
                "平价_制造": 30.2,
                "平价_科技": 30.3,
                "平价_消费": 30.4,
                "平价_金融": 30.5,
                "纯债溢价率_周期": 40.1,
                "纯债溢价率_制造": 40.2,
                "纯债溢价率_科技": 40.3,
                "纯债溢价率_消费": 40.4,
                "纯债溢价率_金融": 40.5,
            },
            {
                "交易日期": "2026-08-27",
                "转股溢价率_周期": 1.1,
                "转股溢价率_制造": 1.2,
                "转股溢价率_科技": 1.3,
                "转股溢价率_消费": 1.4,
                "转股溢价率_金融": 1.5,
                "收盘价_周期": 2.1,
                "收盘价_制造": 2.2,
                "收盘价_科技": 2.3,
                "收盘价_消费": 2.4,
                "收盘价_金融": 2.5,
                "平价_周期": 3.1,
                "平价_制造": 3.2,
                "平价_科技": 3.3,
                "平价_消费": 3.4,
                "平价_金融": 3.5,
                "纯债溢价率_周期": 4.1,
                "纯债溢价率_制造": 4.2,
                "纯债溢价率_科技": 4.3,
                "纯债溢价率_消费": 4.4,
                "纯债溢价率_金融": 4.5,
            },
        ]
    )


class JsDailyAggregationTests(unittest.TestCase):
    def test_aaa_ytm_median_uses_rating_and_one_year_minimum(self) -> None:
        result = MODULE.aggregate_js_daily_metrics(
            make_daily_rows("2026-08-28"),
            sector_mean_metrics=make_sector_mean_rows(),
        )

        ytm = result["JS偏债型余额YTM"]
        self.assertIn("AAA评级YTM中位数", ytm.columns)
        self.assertAlmostEqual(float(ytm.iloc[0]["AAA评级YTM中位数"]), -0.5)

    def test_metrics_match_legacy_definitions_without_demon_filter(self) -> None:
        aggregate = getattr(MODULE, "aggregate_js_daily_metrics", None)
        self.assertIsNotNone(aggregate, "日报模块尚未提供旧 JS 指标聚合函数")

        result = aggregate(
            make_daily_rows("2026-08-28"),
            sector_mean_metrics=make_sector_mean_rows(),
        )

        self.assertEqual(list(result), list(EXPECTED_COLUMNS))
        for sheet_name, columns in EXPECTED_COLUMNS.items():
            self.assertEqual(result[sheet_name].columns.tolist(), columns)

        full = result["JS全样本"].iloc[0]
        self.assertAlmostEqual(float(full["平价算术平均"]), 104.0)
        self.assertAlmostEqual(float(full["纯债价值算术平均"]), 110.0)
        self.assertAlmostEqual(float(full["纯债溢价率算术平均"]), 30.0)
        self.assertAlmostEqual(float(full["隐含波动率算术平均"]), 15.0)

        parity = result["平价分类转股溢价率"].iloc[0]
        self.assertAlmostEqual(float(parity["平价80以下"]), 100.0)
        self.assertAlmostEqual(float(parity["平价80-95"]), 20.0)
        self.assertAlmostEqual(float(parity["平价95-110"]), 30.0)
        self.assertAlmostEqual(float(parity["平价110-125"]), 40.0)
        self.assertAlmostEqual(float(parity["平价125以上"]), 50.0)

        floor = result["JS平底分类余额加权转股溢价率"].iloc[0]
        self.assertAlmostEqual(float(floor["偏股"]), 100.0)
        self.assertAlmostEqual(float(floor["偏债"]), 40.0)
        self.assertAlmostEqual(float(floor["平衡"]), 30.0)

        ytm = result["JS偏债型余额YTM"].iloc[0]
        self.assertAlmostEqual(float(ytm["YTM余额加权"]), 4.0)
        self.assertAlmostEqual(float(ytm["YTM中位数"]), 4.0)
        self.assertAlmostEqual(float(ytm["AAA评级YTM中位数"]), -0.5)

        close = result["收盘价分位数统计"].iloc[0]
        expected_close = {
            "0.05": 104.0,
            "0.25": 120.0,
            "0.5": 130.0,
            "0.75": 140.0,
            "0.8": 152.0,
            "0.9": 176.0,
            "均值": 138.0,
        }
        for column, expected in expected_close.items():
            self.assertAlmostEqual(float(close[column]), expected)

    def test_every_sheet_is_sorted_by_date_ascending(self) -> None:
        data = pd.concat(
            [make_daily_rows("2026-08-28"), make_daily_rows("2026-08-27")],
            ignore_index=True,
        )

        result = MODULE.aggregate_js_daily_metrics(
            data, sector_mean_metrics=make_sector_mean_rows()
        )

        for frame in result.values():
            self.assertEqual(
                pd.to_datetime(frame["日期"]).dt.strftime("%Y-%m-%d").tolist(),
                ["2026-08-27", "2026-08-28"],
            )

    def test_sector_mean_metrics_are_reused_as_requested_wide_sheets(self) -> None:
        aggregate = getattr(MODULE, "aggregate_js_sector_mean_metrics", None)
        self.assertIsNotNone(aggregate, "日报模块尚未提供JS板块宽表转换函数")
        result = aggregate(make_sector_mean_rows())

        self.assertEqual(
            list(result)[-4:],
            ["板块转股溢价率", "板块价格", "板块平价", "板块纯债溢价率"],
        )
        expected_latest = {
            "板块转股溢价率": [10.1, 10.2, 10.3, 10.4, 10.5],
            "板块价格": [20.1, 20.2, 20.3, 20.4, 20.5],
            "板块平价": [30.1, 30.2, 30.3, 30.4, 30.5],
            "板块纯债溢价率": [40.1, 40.2, 40.3, 40.4, 40.5],
        }
        for sheet_name, values in expected_latest.items():
            frame = result[sheet_name]
            self.assertEqual(frame.columns.tolist(), EXPECTED_COLUMNS[sheet_name])
            self.assertEqual(
                pd.to_datetime(frame["日期"]).dt.strftime("%Y-%m-%d").tolist(),
                ["2026-08-27", "2026-08-28"],
            )
            self.assertEqual(frame.iloc[-1, 1:].tolist(), values)

    def test_sector_mean_wide_sheets_are_appended_to_js_daily_output(self) -> None:
        try:
            result = MODULE.aggregate_js_daily_metrics(
                make_daily_rows("2026-08-28"),
                sector_mean_metrics=make_sector_mean_rows(),
            )
        except TypeError as exc:
            self.fail(f"JS日报聚合尚未接入板块宽表：{exc}")

        self.assertEqual(list(result), list(EXPECTED_COLUMNS))

    def test_inverse_cubic_and_subnew_series_use_confirmed_sample_rules(self) -> None:
        aggregate = getattr(MODULE, "aggregate_js_hundred_premium_series", None)
        self.assertIsNotNone(aggregate, "日报模块尚未提供 JS 百元溢价率聚合函数")
        subnew = pd.DataFrame(
            {
                "交易日期": pd.to_datetime(["2026-08-28", "2026-08-29"]),
                "次新券平均转股溢价率": [11.11, 12.34],
            }
        )

        result = aggregate(make_inverse_cubic_rows(), subnew)

        self.assertEqual(result.columns.tolist(), EXPECTED_COLUMNS["百元溢价率"])
        self.assertEqual(
            pd.to_datetime(result["日期"]).dt.strftime("%Y-%m-%d").tolist(),
            ["2026-08-28", "2026-08-29"],
        )
        first = result.iloc[0]
        second = result.iloc[1]
        self.assertAlmostEqual(float(first["百元平价拟合溢价率"]), 50.0, places=5)
        self.assertAlmostEqual(
            float(first["百元平价拟合溢价率（1-5.5年）"]), 50.0, places=5
        )
        self.assertAlmostEqual(
            float(second["百元平价拟合溢价率（1-5.5年）"]), 30.0, places=5
        )
        self.assertAlmostEqual(float(second["次新券转股溢价率均值"]), 12.34)

    def test_full_inverse_cubic_fit_does_not_require_remaining_maturity(self) -> None:
        data = make_inverse_cubic_rows().loc[
            lambda frame: frame["交易日期"].eq("2026-08-28")
        ].copy()
        data["剩余期限"] = float("nan")

        result = MODULE.aggregate_js_hundred_premium_series(data)

        self.assertAlmostEqual(
            float(result.iloc[0]["百元平价拟合溢价率"]), 50.0, places=5
        )
        self.assertTrue(
            pd.isna(result.iloc[0]["百元平价拟合溢价率（1-5.5年）"])
        )

    def test_workbook_contains_only_requested_result_sheets(self) -> None:
        metrics = MODULE.aggregate_js_daily_metrics(
            pd.concat(
                [make_daily_rows("2026-08-27"), make_daily_rows("2026-08-28")],
                ignore_index=True,
            ),
            sector_mean_metrics=make_sector_mean_rows(),
        )
        main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        namespace = {"main": main_ns, "pkg": package_rel_ns}

        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "JS日报_20260828.xlsx"
            MODULE.build_js_daily_workbook(metrics, output_path)
            self.assertTrue(output_path.is_file())
            with zipfile.ZipFile(output_path) as workbook:
                workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
                relationship_root = ET.fromstring(
                    workbook.read("xl/_rels/workbook.xml.rels")
                )
                relationships = {
                    relation.attrib["Id"]: relation.attrib["Target"]
                    for relation in relationship_root.findall("pkg:Relationship", namespace)
                }
                shared_strings: list[str] = []
                if "xl/sharedStrings.xml" in workbook.namelist():
                    shared_root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
                    shared_strings = [
                        "".join(node.text or "" for node in item.findall(".//main:t", namespace))
                        for item in shared_root.findall("main:si", namespace)
                    ]

                sheets = workbook_root.findall("main:sheets/main:sheet", namespace)
                actual_names = [sheet.attrib["name"] for sheet in sheets]
                first_rows: dict[str, list[str]] = {}
                all_text: list[str] = []
                last_rows: dict[str, int] = {}
                third_column_widths: dict[str, float] = {}
                for sheet in sheets:
                    relationship_id = sheet.attrib[f"{{{office_rel_ns}}}id"]
                    member = relationships[relationship_id].replace("\\", "/").lstrip("/")
                    if not member.startswith("xl/"):
                        member = f"xl/{member}"
                    root = ET.fromstring(workbook.read(member))
                    for column in root.findall("main:cols/main:col", namespace):
                        if int(column.attrib["min"]) <= 3 <= int(column.attrib["max"]):
                            third_column_widths[sheet.attrib["name"]] = float(
                                column.attrib["width"]
                            )
                    row_nodes = root.findall(".//main:sheetData/main:row", namespace)
                    last_rows[sheet.attrib["name"]] = max(
                        int(row.attrib["r"]) for row in row_nodes
                    )
                    header_values: list[str] = []
                    for cell in root.findall(".//main:row[@r='1']/main:c", namespace):
                        if cell.attrib.get("t") == "inlineStr":
                            value = "".join(
                                node.text or "" for node in cell.findall(".//main:t", namespace)
                            )
                        else:
                            node = cell.find("main:v", namespace)
                            raw = "" if node is None else (node.text or "")
                            value = shared_strings[int(raw)] if cell.attrib.get("t") == "s" else raw
                        header_values.append(value)
                        all_text.append(value)
                    first_rows[sheet.attrib["name"]] = header_values

        self.assertEqual(actual_names, list(EXPECTED_COLUMNS))
        for sheet_name, columns in EXPECTED_COLUMNS.items():
            self.assertEqual(first_rows[sheet_name], columns)
            self.assertEqual(last_rows[sheet_name], 3)
        joined = "\n".join(all_text)
        self.assertNotIn("数据来源", joined)
        self.assertNotIn("说明", joined)
        self.assertNotIn("图表", joined)
        self.assertGreaterEqual(third_column_widths["百元溢价率"], 28.0)

    def test_workbook_freezes_first_row_and_first_column_for_every_sheet(self) -> None:
        metrics = MODULE.aggregate_js_daily_metrics(
            pd.concat(
                [make_daily_rows("2026-08-27"), make_daily_rows("2026-08-28")],
                ignore_index=True,
            ),
            sector_mean_metrics=make_sector_mean_rows(),
        )
        main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        namespace = {"main": main_ns, "pkg": package_rel_ns}

        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "JS日报_20260828.xlsx"
            MODULE.build_js_daily_workbook(metrics, output_path)
            with zipfile.ZipFile(output_path) as workbook:
                workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
                relationship_root = ET.fromstring(
                    workbook.read("xl/_rels/workbook.xml.rels")
                )
                relationships = {
                    relation.attrib["Id"]: relation.attrib["Target"]
                    for relation in relationship_root.findall("pkg:Relationship", namespace)
                }
                actual_panes: dict[str, dict[str, str]] = {}
                for sheet in workbook_root.findall(
                    "main:sheets/main:sheet", namespace
                ):
                    relationship_id = sheet.attrib[f"{{{office_rel_ns}}}id"]
                    member = relationships[relationship_id].replace("\\", "/").lstrip("/")
                    if not member.startswith("xl/"):
                        member = f"xl/{member}"
                    root = ET.fromstring(workbook.read(member))
                    pane = root.find("main:sheetViews/main:sheetView/main:pane", namespace)
                    actual_panes[sheet.attrib["name"]] = (
                        {} if pane is None else dict(pane.attrib)
                    )

        expected_pane = {
            "xSplit": "1",
            "ySplit": "1",
            "topLeftCell": "B2",
            "state": "frozen",
        }
        self.assertEqual(set(actual_panes), set(EXPECTED_COLUMNS))
        for sheet_name, pane in actual_panes.items():
            self.assertEqual(
                {key: pane.get(key) for key in expected_pane},
                expected_pane,
                sheet_name,
            )


if __name__ == "__main__":
    unittest.main()
