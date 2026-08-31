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
SPEC = importlib.util.spec_from_file_location("cb_daily_sector_plot_window", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class SectorPlotWindowTests(unittest.TestCase):
    def test_sector_mean_plot_starts_from_2023_without_truncating_source_data(self) -> None:
        dates = pd.to_datetime(["2022-12-30", "2023-01-03", "2023-01-04"])
        data = pd.DataFrame({"交易日期": dates})
        for position, sector in enumerate(MODULE.SECTOR_ORDER, start=1):
            data[f"收盘价_{sector}"] = [100 + position, 101 + position, 102 + position]

        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "sector_mean.png"
            with patch.object(MODULE.plt, "close"):
                MODULE.plot_sector_mean_metric(
                    data,
                    "收盘价",
                    "各行业平均收盘价",
                    "",
                    output_path,
                )
                figure = MODULE.plt.gcf()
            try:
                first_line_dates = pd.to_datetime(figure.axes[0].lines[0].get_xdata())
                self.assertEqual(first_line_dates.min(), pd.Timestamp("2023-01-03"))
                self.assertEqual(data["交易日期"].min(), pd.Timestamp("2022-12-30"))
            finally:
                MODULE.plt.close(figure)


if __name__ == "__main__":
    unittest.main()
