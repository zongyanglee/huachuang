from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_valuation_legend", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class DailyValuationLegendTests(unittest.TestCase):
    def test_primary_curve_legend_omits_method_parenthetical(self) -> None:
        dates = pd.to_datetime(
            ["2019-01-02", "2021-01-04", "2024-01-02", "2026-08-28"]
        )
        data = pd.DataFrame(
            {
                "交易日期": dates,
                MODULE.INVERSE_CUBIC_VALUATION_NAME: [18.0, 22.0, 25.0, 27.0],
                MODULE.MULTIFACTOR_VALUATION_NAME: [17.0, 21.0, 24.0, 26.0],
            }
        )
        source = {
            "latestValue": 27.0,
            "dailyChangePctPoint": 2.0,
            "percentileSince2019": 100.0,
        }

        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "daily_valuation.png"
            with patch.object(MODULE.plt, "close"):
                MODULE.plot_daily_valuation(data, source, output_path)
                figure = MODULE.plt.gcf()
            try:
                legend = figure.axes[0].get_legend()
                labels = [text.get_text() for text in legend.get_texts()]
                self.assertEqual(labels[0], "百元拟合溢价率")
                self.assertNotIn("三次反比例", labels[0])
            finally:
                MODULE.plt.close(figure)


if __name__ == "__main__":
    unittest.main()
