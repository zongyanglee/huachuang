from __future__ import annotations

from datetime import date
import importlib.util
from pathlib import Path
import sys
import tempfile
import textwrap
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile


SCRIPT = Path(__file__).resolve().parents[1] / "src" / "daily" / "发行日历.py"
SPEC = importlib.util.spec_from_file_location("convertible_bond_calendar", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


def workbook_cell_value(path: Path, reference: str) -> str:
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as archive:
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        cell = sheet.find(f".//x:c[@r='{reference}']", namespace)
        if cell is None:
            raise AssertionError(f"工作簿中未找到单元格 {reference}")
        cell_type = cell.get("t")
        if cell_type == "inlineStr":
            return "".join(node.text or "" for node in cell.findall(".//x:t", namespace))
        value = cell.findtext("x:v", default="", namespaces=namespace)
        if cell_type != "s":
            return value
        shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        item = shared.findall("x:si", namespace)[int(value)]
        return "".join(node.text or "" for node in item.findall(".//x:t", namespace))


class ListingPredictionTests(unittest.TestCase):
    def test_non_strict_t1_latest_prediction_is_returned_with_flag(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_text:
            fake_model = Path(temp_dir_text) / "fake_v21_model.py"
            fake_model.write_text(
                textwrap.dedent(
                    """
                    import argparse
                    import json
                    from pathlib import Path
                    import pandas as pd

                    parser = argparse.ArgumentParser()
                    parser.add_argument("--project-dir", required=True)
                    args = parser.parse_args()
                    root = Path(args.project_dir)
                    (root / "预测").mkdir(parents=True, exist_ok=True)
                    (root / "版本登记").mkdir(parents=True, exist_ok=True)
                    pd.DataFrame([{
                        "转债代码": "123282.SZ",
                        "转债名称": "震裕转02",
                        "上市日期": "2026-09-02",
                        "预测信息日": "2026-08-28",
                        "最终预测价": 156.18453597933495,
                        "模型版本": "v2.1_极低筹码非线性修正",
                    }]).to_csv(root / "预测" / "v21_current_forecasts.csv", index=False, encoding="utf-8-sig")
                    (root / "版本登记" / "v21_current_summary.json").write_text(
                        json.dumps({
                            "model_version": "cb_listing_v2.1_scarcity_overlay",
                            "data_as_of": "2026-08-28",
                        }, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    """
                ),
                encoding="utf-8",
            )

            with patch.object(MODULE, "V21_PRICING_SCRIPT", fake_model):
                predictions, _ = MODULE.fetch_v21_listing_prices(
                    date(2026, 8, 30),
                    [("123282.SZ", date(2026, 9, 2))],
                    [
                        date(2026, 8, 28),
                        date(2026, 8, 31),
                        date(2026, 9, 1),
                        date(2026, 9, 2),
                    ],
                )

        prediction = predictions[("123282.SZ", date(2026, 9, 2))]
        self.assertAlmostEqual(prediction["price"], 156.18453597933495)
        self.assertFalse(prediction["strict_t1"])
        self.assertEqual(prediction["data_as_of"], "2026-08-28")

    def test_non_strict_t1_prediction_renders_with_asterisk_only(self) -> None:
        payload = {
            "title": MODULE.TITLE,
            "subtitle": MODULE.SUBTITLE,
            "updated_date": "2026-08-30",
            "calendar_dates": [
                "2026-08-31",
                "2026-09-01",
                "2026-09-02",
                "2026-09-03",
                "2026-09-04",
            ],
            "trading_dates": [
                "2026-08-28",
                "2026-08-31",
                "2026-09-01",
                "2026-09-02",
                "2026-09-03",
                "2026-09-04",
            ],
            "colors": MODULE.COLORS,
            "bonds": [
                {
                    "网上申购代码": "370953",
                    "转债代码": "123282.SZ",
                    "简称": "震裕转02",
                    "所属行业": "电力设备",
                    "发行规模": 18.8,
                    "债项评级": "AA-",
                    "评级公司": "上海新世纪",
                    "上市价格预测V2.1": 156.18453597933495,
                    "上市价格预测非严格T1": True,
                    "上市价格预测状态": "已预测：非严格T-1",
                    "发行日期": "2026-08-11",
                    "上市日期": "2026-09-02",
                }
            ],
        }

        with tempfile.TemporaryDirectory() as temp_dir_text:
            outputs = MODULE.render_payload(payload, Path(temp_dir_text))
            self.assertEqual(workbook_cell_value(outputs["xlsx"], "H4"), "156.18*")
            footer = workbook_cell_value(outputs["xlsx"], "A5")

        self.assertEqual(footer, "更新时间：2026年8月30日")


if __name__ == "__main__":
    unittest.main()
