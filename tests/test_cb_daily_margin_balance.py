from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest

import pandas as pd


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_margin_balance", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class MarginBalancePlotDataTests(unittest.TestCase):
    def test_incomplete_jump_is_hidden_without_discarding_next_normal_point(self) -> None:
        source = pd.DataFrame(
            {
                "交易日期": pd.to_datetime(
                    ["2026-08-26", "2026-08-27", "2026-08-28", "2026-08-31"]
                ),
                "沪深两市融资融券余额_亿元": [
                    26_448.93,
                    26_567.87,
                    13_613.85,
                    26_580.00,
                ],
            }
        )
        prepare = getattr(MODULE, "prepare_margin_balance_plot", None)
        self.assertIsNotNone(prepare, "尚未实现两融余额绘图数据校验")

        plot_data, panel_title = prepare(source)

        self.assertEqual(
            plot_data["交易日期"].dt.strftime("%Y-%m-%d").tolist(),
            ["2026-08-26", "2026-08-27", "2026-08-31"],
        )
        self.assertEqual(panel_title, "沪深两市融资融券余额")
        self.assertEqual(len(source), 4, "绘图校验不应修改原始底稿数据")


class PriceParityPlotRegressionTests(unittest.TestCase):
    def test_price_parity_plot_uses_its_own_date_series(self) -> None:
        data = pd.DataFrame(
            {
                "交易日期": pd.to_datetime(["2025-01-02", "2026-08-28"]),
                "余额加权平价": [98.0, 102.0],
                "收盘价中位数": [112.0, 116.0],
            }
        )
        source = {
            "latestParity": 102.0,
            "parityDailyChangePct": 0.2,
            "latestMedianPrice": 116.0,
            "medianPriceDailyChangePct": 0.3,
            "medianPricePercentileSince2019": 65.0,
        }

        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "price_parity.png"
            try:
                MODULE.plot_price_parity_series(data, source, output_path)
            except NameError as exc:
                self.fail(f"价格与平价图错误引用其他图表变量：{exc}")
            self.assertTrue(output_path.is_file())
            self.assertGreater(output_path.stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
