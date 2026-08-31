from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_equity_bond_weighted", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class EquityBondWeightedSeriesTests(unittest.TestCase):
    @staticmethod
    def _fixture() -> pd.DataFrame:
        return pd.DataFrame(
            {
                "交易日期": pd.to_datetime(
                    [
                        "2026-08-27",
                        "2026-08-27",
                        "2026-08-27",
                        "2026-08-27",
                        "2026-08-27",
                        "2026-08-28",
                        "2026-08-28",
                        "2026-08-28",
                        "2026-08-28",
                        "2026-08-28",
                    ]
                ),
                "交易状态": ["交易"] * 10,
                "余额": [100, 50, 100, 50, 150, 100, 50, 100, 50, 150],
                "平价底价溢价率": [30, 25, 0, -30, -25, 30, 25, 0, -30, -25],
                "转股溢价率": [10, 20, 30, 40, 60, 14, 20, 28, 50, 60],
                "收盘价": [100, 120, 110, 90, 120, 105, 125, 112, 95, 125],
            }
        )

    def test_balance_weighted_series_and_titles_keep_group_metrics(self) -> None:
        result, source = MODULE.aggregate_equity_bond_weighted_series(self._fixture())

        self.assertEqual(result["交易日期"].dt.strftime("%Y-%m-%d").tolist(), ["2026-08-27", "2026-08-28"])
        self.assertAlmostEqual(result.loc[0, "偏股型_转股溢价率"], 13.3333333333)
        self.assertAlmostEqual(result.loc[1, "偏股型_转股溢价率"], 16.0)
        self.assertAlmostEqual(result.loc[1, "偏债型_转股溢价率"], 57.5)
        self.assertAlmostEqual(result.loc[0, "偏债型_收盘价"], 112.5)
        self.assertAlmostEqual(result.loc[1, "偏股型_收盘价"], 111.6666666667)
        self.assertAlmostEqual(
            source["premium"]["偏股型"]["dailyChangePctPoint"], 2.6666666667
        )
        self.assertAlmostEqual(
            source["price"]["偏债型"]["dailyChangePct"], 4.4444444444
        )

        titles = MODULE.build_equity_bond_weighted_titles(source)
        self.assertEqual(
            titles["premium"],
            "股债性分类转股溢价率：\n偏股型16.00%，+2.67pct；偏债型57.50%，+2.50pct",
        )
        self.assertEqual(
            titles["price"],
            "股债性分类均价：\n偏股型111.67，+4.69%；偏债型117.50，+4.44%",
        )

    def test_premium_chart_places_bond_type_on_right_axis(self) -> None:
        data, source = MODULE.aggregate_equity_bond_weighted_series(self._fixture())
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "equity_bond_premium.png"
            with patch.object(MODULE.plt, "close"):
                MODULE.plot_equity_bond_weighted_premium(data, source, output_path)
                figure = MODULE.plt.gcf()
            try:
                self.assertEqual(len(figure.axes), 2)
                self.assertEqual(len(figure.axes[0].lines), 2)
                self.assertEqual(len(figure.axes[1].lines), 1)
                np.testing.assert_allclose(
                    figure.axes[1].lines[0].get_ydata(),
                    data["偏债型_转股溢价率"].to_numpy(),
                )
            finally:
                MODULE.plt.close(figure)


if __name__ == "__main__":
    unittest.main()
