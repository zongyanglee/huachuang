from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd
from PIL import Image
from matplotlib import colors as mcolors


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_valuation_repair", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class DailyValuationRepairTests(unittest.TestCase):
    def test_fetch_uses_parquet_name_for_cb_index(self) -> None:
        """Parquet 中的“转债指数”应被映射为修复指数右轴的中证转债指数。"""
        dates = pd.bdate_range("2026-01-01", periods=101)
        values = np.arange(101, dtype=float)
        source = pd.concat(
            [
                pd.DataFrame(
                    {
                        "指数名称": MODULE.INVERSE_CUBIC_VALUATION_NAME,
                        "交易日期": dates,
                        "指数值": values,
                    }
                ),
                pd.DataFrame(
                    {
                        "指数名称": "转债指数",
                        "交易日期": dates,
                        "指数值": 100.0 + values,
                    }
                ),
            ],
            ignore_index=True,
        )
        with patch.object(MODULE.pd, "read_parquet", return_value=source):
            result = MODULE.fetch_valuation_repair_index(dates[-1].date())

        self.assertEqual(result["交易日期"].iloc[-1].date(), dates[-1].date())

    def test_repair_index_uses_100_day_corridor_and_capped_forward_return(self) -> None:
        """修复指数按 100 日走廊映射，后推收益在样本末端使用最新交易日。"""
        dates = pd.bdate_range("2026-01-01", periods=171)
        values = np.arange(171, dtype=float)
        data = pd.DataFrame(
            {
                "交易日期": dates,
                MODULE.INVERSE_CUBIC_VALUATION_NAME: values,
                MODULE.CB_INDEX_PARQUET_NAME: 100.0 + values,
            }
        )

        result = MODULE.calculate_valuation_repair_index(data)

        first_valid = result.iloc[99]
        expected_std = float(np.std(np.arange(100, dtype=float), ddof=0))
        upper_corridor = 49.5 + 2 * expected_std
        lower_corridor = 49.5 - 2 * expected_std
        expected_repair = 100 * (upper_corridor - 99.0) / (
            upper_corridor - lower_corridor
        )
        self.assertAlmostEqual(first_valid["转债估值修复指数"], expected_repair, places=8)
        self.assertAlmostEqual(result.iloc[0]["中证转债指数后推70日涨跌幅"], 70.0, places=8)
        self.assertAlmostEqual(result.iloc[150]["中证转债指数后推70日涨跌幅"], 8.0, places=8)
        self.assertEqual(result.iloc[-1]["中证转债指数后推70日涨跌幅"], 0.0)

    def test_repair_index_plot_has_standard_chart_dimensions(self) -> None:
        """修复指数以日报标准双轴小图尺寸输出。"""
        dates = pd.bdate_range("2026-01-01", periods=171)
        values = np.arange(171, dtype=float)
        data = MODULE.calculate_valuation_repair_index(
            pd.DataFrame(
                {
                    "交易日期": dates,
                    MODULE.INVERSE_CUBIC_VALUATION_NAME: values,
                    MODULE.CB_INDEX_PARQUET_NAME: 100.0 + values,
                }
            )
        )
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "repair.png"
            MODULE.plot_valuation_repair_index(data, output_path)
            with Image.open(output_path) as image:
                self.assertEqual(
                    image.size,
                    (MODULE.CHART_PIXEL_WIDTH, MODULE.CHART_PIXEL_HEIGHT),
                )

    def test_repair_index_plot_starts_from_2024(self) -> None:
        """绘图只展示 2024 年以来，计算结果仍可包含更早历史。"""
        dates = pd.bdate_range("2023-01-02", periods=700)
        values = np.linspace(20.0, 45.0, len(dates)) + np.sin(
            np.arange(len(dates)) / 12.0
        )
        data = MODULE.calculate_valuation_repair_index(
            pd.DataFrame(
                {
                    "交易日期": dates,
                    MODULE.INVERSE_CUBIC_VALUATION_NAME: values,
                    MODULE.CB_INDEX_PARQUET_NAME: 400.0 + np.arange(len(dates)) * 0.1,
                }
            )
        )
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "repair.png"
            with patch.object(MODULE.plt, "close"):
                MODULE.plot_valuation_repair_index(data, output_path)
                figure = MODULE.plt.gcf()
            try:
                left_limit = MODULE.mdates.num2date(figure.axes[0].get_xlim()[0])
                self.assertEqual(left_limit.date(), pd.Timestamp("2024-01-01").date())
            finally:
                MODULE.plt.close(figure)

    def test_repair_index_plot_uses_tight_data_driven_y_limits(self) -> None:
        """双轴上下限应贴合有效数据，避免固定对称区间产生大面积空白。"""
        dates = pd.bdate_range("2024-01-02", periods=20)
        data = pd.DataFrame(
            {
                "交易日期": dates,
                MODULE.VALUATION_REPAIR_INDEX_NAME: np.linspace(10.0, 30.0, len(dates)),
                MODULE.VALUATION_REPAIR_FORWARD_RETURN_NAME: np.linspace(
                    -2.0, 4.0, len(dates)
                ),
            }
        )
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "repair.png"
            with patch.object(MODULE.plt, "close"):
                MODULE.plot_valuation_repair_index(data, output_path)
                figure = MODULE.plt.gcf()
            try:
                left_span = np.diff(figure.axes[0].get_ylim())[0]
                right_span = np.diff(figure.axes[1].get_ylim())[0]
                self.assertLess(left_span, 30.0)
                self.assertLess(right_span, 9.0)
            finally:
                MODULE.plt.close(figure)

    def test_repair_index_plot_has_no_grid_or_top_axis_line(self) -> None:
        """估值修复指数图不显示灰色网格线和顶部坐标轴线。"""
        dates = pd.bdate_range("2024-01-02", periods=20)
        data = pd.DataFrame(
            {
                "交易日期": dates,
                MODULE.VALUATION_REPAIR_INDEX_NAME: np.linspace(10.0, 30.0, len(dates)),
                MODULE.VALUATION_REPAIR_FORWARD_RETURN_NAME: np.linspace(
                    -2.0, 4.0, len(dates)
                ),
            }
        )
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "repair.png"
            with patch.object(MODULE.plt, "close"):
                MODULE.plot_valuation_repair_index(data, output_path)
                figure = MODULE.plt.gcf()
            try:
                for axis in figure.axes:
                    grid_lines = axis.get_xgridlines() + axis.get_ygridlines()
                    self.assertFalse(any(line.get_visible() for line in grid_lines))
                    self.assertFalse(axis.spines["top"].get_visible())
            finally:
                MODULE.plt.close(figure)

    def test_repair_index_plot_uses_double_line_title_band_height(self) -> None:
        """换位后应与同排右图共用双行标题栏高度。"""
        dates = pd.bdate_range("2024-01-02", periods=20)
        data = pd.DataFrame(
            {
                "交易日期": dates,
                MODULE.VALUATION_REPAIR_INDEX_NAME: np.linspace(10.0, 30.0, len(dates)),
                MODULE.VALUATION_REPAIR_FORWARD_RETURN_NAME: np.linspace(
                    -2.0, 4.0, len(dates)
                ),
            }
        )
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "repair.png"
            with patch.object(MODULE.plt, "close"):
                MODULE.plot_valuation_repair_index(data, output_path)
                figure = MODULE.plt.gcf()
            try:
                title_bands = [
                    artist
                    for artist in figure.artists
                    if isinstance(artist, MODULE.plt.Rectangle)
                    and artist.get_facecolor()[:3]
                    == mcolors.to_rgba("#D9E2F3")[:3]
                ]
                self.assertEqual(len(title_bands), 1)
                self.assertAlmostEqual(
                    title_bands[0].get_height(),
                    MODULE.DOUBLE_LINE_TITLE_BAND_HEIGHT,
                )
            finally:
                MODULE.plt.close(figure)

    def test_repair_index_title_shows_latest_value_and_daily_change(self) -> None:
        """标题应显示最新估值修复指数及相邻有效交易日变动。"""
        dates = pd.bdate_range("2024-01-02", periods=3)
        data = pd.DataFrame(
            {
                "交易日期": dates,
                MODULE.VALUATION_REPAIR_INDEX_NAME: [10.0, 20.0, 21.0],
                MODULE.VALUATION_REPAIR_FORWARD_RETURN_NAME: [1.0, 2.0, 3.0],
            }
        )
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "repair.png"
            with patch.object(MODULE.plt, "close"):
                MODULE.plot_valuation_repair_index(data, output_path)
                figure = MODULE.plt.gcf()
            try:
                self.assertIn(
                    "转债估值修复指数21.00%；+1.00pct",
                    [text.get_text() for text in figure.texts],
                )
            finally:
                MODULE.plt.close(figure)


if __name__ == "__main__":
    unittest.main()
