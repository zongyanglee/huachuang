from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import unittest


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_valuation_titles", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class DailyValuationTitleTests(unittest.TestCase):
    def test_group_valuation_titles_use_fitted_premium_labels(self) -> None:
        sources = {
            "parity": {
                "largestChangeGroup": "90-110",
                "largestChangePctPoint": 0.25,
            },
            "equity": {
                "偏股型": {"latestValue": 22.11, "dailyChangePctPoint": 0.31},
                "偏债型": {"latestValue": 18.22, "dailyChangePctPoint": -0.42},
            },
            "rating": {
                "AAA/AA+": {"latestValue": 20.10, "dailyChangePctPoint": 0.21},
                "AA/AA-": {"latestValue": 25.20, "dailyChangePctPoint": -0.32},
            },
            "balance": {
                "0-3": {"latestValue": 19.10, "dailyChangePctPoint": 0.11},
                "50+": {"latestValue": 23.20, "dailyChangePctPoint": -0.22},
            },
            "market_cap": {
                "0-50": {"latestValue": 24.10, "dailyChangePctPoint": 0.41},
                "300+": {"latestValue": 18.20, "dailyChangePctPoint": -0.52},
            },
            "sector": {
                "科技": {"latestValue": 28.10, "dailyChangePctPoint": 0.61},
                "周期": {"latestValue": 16.20, "dailyChangePctPoint": -0.72},
            },
        }

        titles = MODULE.build_group_valuation_titles(sources)

        self.assertEqual(
            titles["parity"],
            "平价分类拟合溢价率，90-110：+0.25pct",
        )
        self.assertEqual(
            titles["equity"].splitlines()[0],
            "股债性分类拟合溢价率",
        )
        self.assertEqual(
            titles["rating"].splitlines()[0],
            "评级分类拟合溢价率",
        )
        self.assertEqual(
            titles["balance"].splitlines()[0],
            "余额分类拟合溢价率：",
        )
        self.assertEqual(
            titles["market_cap"].splitlines()[0],
            "市值分类拟合溢价率：",
        )
        self.assertEqual(
            titles["sector"].splitlines()[0],
            "各板块拟合溢价率：",
        )
        self.assertIn("偏股型：22.11%，+0.31pct", titles["equity"])
        self.assertIn("科技：28.10%，+0.61pct", titles["sector"])


if __name__ == "__main__":
    unittest.main()
