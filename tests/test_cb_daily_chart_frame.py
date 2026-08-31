from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import unittest

import numpy as np


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_chart_frame", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class DailyChartFrameTests(unittest.TestCase):
    def test_panel_title_does_not_draw_black_outer_canvas_frame(self) -> None:
        """统一标题栏不得在小图画布四周绘制黑色轮廓框。"""
        figure = MODULE.plt.figure(
            figsize=MODULE.CHART_FIGSIZE,
            dpi=MODULE.CHART_DPI,
            facecolor="white",
        )
        try:
            MODULE.add_chart_panel_title(figure, "测试标题")
            figure.canvas.draw()
            pixels = np.asarray(figure.canvas.buffer_rgba())[:, :, :3]
            perimeter = np.concatenate(
                [
                    pixels[0, :, :],
                    pixels[-1, :, :],
                    pixels[:, 0, :],
                    pixels[:, -1, :],
                ],
                axis=0,
            )
            black_pixels = np.all(perimeter < 50, axis=1)
            self.assertFalse(bool(black_pixels.any()))
        finally:
            MODULE.plt.close(figure)


if __name__ == "__main__":
    unittest.main()
