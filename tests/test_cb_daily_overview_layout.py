from __future__ import annotations

import importlib.util
from datetime import date
from pathlib import Path
import sys
import tempfile
import unittest

import numpy as np
from PIL import Image


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_overview_layout", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class DailyOverviewLayoutTests(unittest.TestCase):
    def test_repair_and_parity_charts_are_swapped_in_long_report(self) -> None:
        """估值修复指数应在第二排左侧，平价分类拟合溢价率应在第三排右侧。"""
        repair_color = np.array([241, 29, 203], dtype=np.int16)
        parity_color = np.array([17, 231, 91], dtype=np.int16)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            table_path = root / "table.png"
            Image.new(
                "RGB",
                (MODULE.DOUBLE_CHART_PIXEL_WIDTH, MODULE.TABLE_PIXEL_HEIGHT),
                "white",
            ).save(table_path)

            chart_paths = []
            for index in range(24):
                path = root / f"chart_{index}.png"
                color = "white"
                if index == 6:
                    color = tuple(parity_color.tolist())
                elif index == 9:
                    color = tuple(repair_color.tolist())
                Image.new(
                    "RGB",
                    (MODULE.CHART_PIXEL_WIDTH, MODULE.CHART_PIXEL_HEIGHT),
                    color,
                ).save(path)
                chart_paths.append(path)

            industry_path = root / "industry.png"
            Image.new(
                "RGB",
                (
                    MODULE.DOUBLE_CHART_PIXEL_WIDTH,
                    MODULE.INDUSTRY_TABLE_PIXEL_HEIGHT,
                ),
                "white",
            ).save(industry_path)
            output_path = root / "overview.png"

            MODULE.compose_index_market_overview(
                table_path,
                *chart_paths,
                industry_path,
                output_path,
                date(2026, 8, 28),
            )

            rendered = np.asarray(Image.open(output_path).convert("RGB"), dtype=np.int16)
            parity_positions = np.argwhere(
                np.max(np.abs(rendered - parity_color), axis=2) <= 1
            )
            repair_positions = np.argwhere(
                np.max(np.abs(rendered - repair_color), axis=2) <= 1
            )

        self.assertGreater(len(parity_positions), 0)
        self.assertGreater(len(repair_positions), 0)
        self.assertLess(float(np.median(repair_positions[:, 0])), float(np.median(parity_positions[:, 0])))
        self.assertLess(float(np.median(repair_positions[:, 1])), MODULE.CHART_PIXEL_WIDTH)
        self.assertGreaterEqual(float(np.median(parity_positions[:, 1])), MODULE.CHART_PIXEL_WIDTH)


if __name__ == "__main__":
    unittest.main()
