from __future__ import annotations

from datetime import date
import importlib.util
from pathlib import Path
import sys
import unittest

import numpy as np
from PIL import Image, ImageDraw, ImageFont


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_report_header", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


def render_header_at_size(run_date: date, font_size: int) -> Image.Image:
    with Image.open(MODULE.REPORT_HEADER_PATH) as source:
        image = source.convert("RGBA")
    text = (
        "【华创固收·周冠南团队】\n"
        f"可转债市场日度跟踪{run_date:%Y%m%d}"
    )
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(str(MODULE.FONT_PATH), font_size)
    bbox = draw.textbbox((0, 0), text, font=font)
    x = (image.width - (bbox[2] - bbox[0])) // 2
    y = (image.height - (bbox[3] - bbox[1])) // 2 - 40
    draw.text((x, y), text, fill="white", font=font)
    return image


def changed_pixel_count(rendered: Image.Image, source: Image.Image) -> int:
    rendered_array = np.asarray(rendered.convert("RGBA"))
    source_array = np.asarray(source.convert("RGBA"))
    return int(np.any(rendered_array != source_array, axis=2).sum())


class ReportHeaderTypographyTests(unittest.TestCase):
    def test_header_text_is_larger_without_bold_stroke(self) -> None:
        render = getattr(MODULE, "render_report_header", None)
        self.assertIsNotNone(render, "尚未实现加大加粗后的日报表头绘制")

        run_date = date(2026, 8, 28)
        with Image.open(MODULE.REPORT_HEADER_PATH) as source_image:
            source = source_image.convert("RGBA")
        previous = render_header_at_size(run_date, 60)
        expected = render_header_at_size(run_date, 72)
        updated = render(run_date)

        self.assertEqual(updated.size, source.size)
        self.assertTrue(
            np.array_equal(np.asarray(updated), np.asarray(expected)),
            "日报表头应使用72号常规楷体，不增加描边加粗",
        )
        self.assertGreater(
            changed_pixel_count(updated, source),
            changed_pixel_count(previous, source) * 1.3,
        )


if __name__ == "__main__":
    unittest.main()
