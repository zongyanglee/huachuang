from __future__ import annotations

import importlib.util
from datetime import date
from pathlib import Path
import sys
import tempfile
import unittest

import numpy as np
import pandas as pd
from PIL import Image


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_fuguo_report", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class FuguoDailyReportTests(unittest.TestCase):
    def test_text_uses_requested_compact_format(self) -> None:
        builder = getattr(MODULE, "build_fuguo_daily_text", None)
        self.assertIsNotNone(builder)
        index_performance = pd.DataFrame(
            [{"指数名称": "中证转债", "日涨跌幅": 0.68}]
        )
        turnover = pd.DataFrame(
            {
                "交易日期": pd.to_datetime(["2026-08-26", "2026-08-27"]),
                "中证转债指数成交额_亿元": [427.18, 500.56],
            }
        )

        text = builder(
            date(2026, 8, 27),
            index_performance=index_performance,
            turnover=turnover,
            price_parity_source={
                "previousDate": "2026-08-26",
                "latestMedianPrice": 134.92,
                "medianPriceDailyChangePct": 0.05,
            },
            valuation_source={
                "previousDate": "2026-08-26",
                "latestValue": 42.33,
                "dailyChangePctPoint": 1.36,
                "percentileSince2019": 99.50,
            },
        )

        self.assertEqual(
            text,
            "转债市场日度跟踪20260827\n"
            "中证转债指数上涨0.68%。转债市场成交情绪升温，"
            "可转债市场成交额为500.56亿元，环比增长17.18%。"
            "价格中位数为134.92元，环比8月26日上升0.05%。"
            "百元平价拟合转股溢价率为42.33%，环比8月26日上升1.36pct。"
            "溢价率在2019年以来99.50%分位数。",
        )

    def test_image_contains_section_bar_and_six_charts_in_three_rows(self) -> None:
        composer = getattr(MODULE, "compose_fuguo_daily_report", None)
        self.assertIsNotNone(composer)
        with tempfile.TemporaryDirectory() as temp_dir_text:
            temp_dir = Path(temp_dir_text)
            sources = []
            colors = [
                (255, 0, 0),
                (0, 0, 255),
                (0, 255, 0),
                (255, 255, 0),
                (255, 0, 255),
                (0, 255, 255),
            ]
            for position, color in enumerate(colors, start=1):
                path = temp_dir / f"source_{position}.png"
                Image.new(
                    "RGB",
                    (MODULE.CHART_PIXEL_WIDTH, MODULE.CHART_PIXEL_HEIGHT),
                    color,
                ).save(path)
                sources.append(path)
            output = temp_dir / "富国日报.png"

            composer(*sources, output)

            with Image.open(output) as image:
                self.assertEqual(
                    image.size,
                    (
                        MODULE.DOUBLE_CHART_PIXEL_WIDTH,
                        MODULE.SECTION_BAR_HEIGHT
                        + MODULE.CHART_PIXEL_HEIGHT * 3,
                    ),
                )
                pixels = np.asarray(image.convert("RGB"))
            row_starts = [
                MODULE.SECTION_BAR_HEIGHT,
                MODULE.SECTION_BAR_HEIGHT + MODULE.CHART_PIXEL_HEIGHT,
                MODULE.SECTION_BAR_HEIGHT + MODULE.CHART_PIXEL_HEIGHT * 2,
            ]
            sampled = []
            for row_start in row_starts:
                sampled.append(tuple(pixels[row_start + 10, 10]))
                sampled.append(
                    tuple(pixels[row_start + 10, MODULE.CHART_PIXEL_WIDTH + 10])
                )
            self.assertEqual(sampled, colors)


if __name__ == "__main__":
    unittest.main()
