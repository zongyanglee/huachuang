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


def _minimal_fuguo_text(ytm_latest: float, ytm_change: float) -> str:
    return MODULE.build_fuguo_daily_text(
        date(2026, 8, 31),
        index_performance=pd.DataFrame(
            [{"指数名称": "中证转债", "日涨跌幅": -0.19}]
        ),
        turnover=pd.DataFrame(
            {
                "交易日期": pd.to_datetime(["2026-08-28", "2026-08-31"]),
                "中证转债指数成交额_亿元": [520.20, 500.56],
            }
        ),
        price_parity_source={
            "previousDate": "2026-08-28",
            "latestMedianPrice": 134.49,
            "medianPriceDailyChangePct": -0.11,
        },
        valuation_source={
            "previousDate": "2026-08-28",
            "latestValue": 41.63,
            "dailyChangePctPoint": -0.31,
            "percentileSince2019": 99.09,
        },
        ytm_source={
            "latestValue": ytm_latest,
            "dailyChangePctPoint": ytm_change,
        },
    )


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
            ytm_source={
                "latestValue": -5.90,
                "dailyChangePctPoint": -0.05,
            },
        )

        self.assertEqual(
            text,
            "转债市场日度跟踪20260827\n"
            "中证转债指数上涨0.68%。转债市场成交情绪升温，"
            "可转债市场成交额为500.56亿元，环比增长17.18%。"
            "价格中位数为134.92元，环比8月26日上升0.05%。"
            "百元平价拟合转股溢价率为42.33%，环比8月26日上升1.36pct。"
            "溢价率在2019年以来99.50%分位数。"
            "YTM中位数为-5.90%，较前一日下降0.05pct。",
        )

    def test_ytm_source_uses_valid_trading_bonds_and_previous_trade_date(
        self,
    ) -> None:
        """YTM中位数应排除非交易及无效样本，并相对前一交易日计算。"""
        with tempfile.TemporaryDirectory() as temp_dir_text:
            parquet_root = Path(temp_dir_text)
            month_dir = parquet_root / "2026"
            month_dir.mkdir()
            pd.DataFrame(
                {
                    "交易日期": pd.to_datetime(
                        [
                            "2026-08-26",
                            "2026-08-26",
                            "2026-08-26",
                            "2026-08-27",
                            "2026-08-27",
                            "2026-08-27",
                            "2026-08-27",
                        ]
                    ),
                    "交易状态": [
                        "交易",
                        "交易",
                        "未上市",
                        "交易",
                        "交易",
                        "交易",
                        "暂停上市",
                    ],
                    "YTM": [-5.70, -6.00, -99.00, -5.80, -6.00, np.nan, -99.00],
                }
            ).to_parquet(month_dir / "202608.parquet", index=False)

            source = MODULE.fetch_ytm_median_source(
                date(2026, 8, 27), parquet_root=parquet_root
            )

            self.assertEqual(source["latestDate"], "2026-08-27")
            self.assertEqual(source["previousDate"], "2026-08-26")
            self.assertAlmostEqual(source["latestValue"], -5.90)
            self.assertAlmostEqual(source["dailyChangePctPoint"], -0.05)
            self.assertEqual(source["latestSampleCount"], 2)
            self.assertEqual(source["previousSampleCount"], 2)

    def test_text_calls_rising_ytm_up(self) -> None:
        """YTM中位数较前一日上升时应直接表述为上升。"""
        text = _minimal_fuguo_text(-6.23, 0.07)

        self.assertTrue(
            text.endswith("YTM中位数为-6.23%，较前一日上升0.07pct。")
        )

    def test_text_calls_unchanged_ytm_flat(self) -> None:
        """YTM中位数与前一日相同时应表述为持平。"""
        text = _minimal_fuguo_text(-5.90, 0.0)

        self.assertTrue(
            text.endswith("YTM中位数为-5.90%，较前一日持平0.00pct。")
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
