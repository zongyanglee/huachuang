from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest

from PIL import Image


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_small_chart_exports", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class DailySmallChartExportTests(unittest.TestCase):
    def test_titleless_export_crops_title_band_and_prefixes_sequence(self) -> None:
        """独立小图删除标题栏，并使用两位数序号作为文件名前缀。"""
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            source_path = output_dir / "成交额源图.png"
            image = Image.new("RGB", (100, 100), "#00FF00")
            for y in range(12):
                for x in range(100):
                    image.putpixel((x, y), (255, 0, 0))
            image.save(source_path)

            MODULE.export_numbered_titleless_small_charts(
                output_dir,
                chart_specs=((1, "成交额", source_path.name, 0.12),),
            )

            exported = output_dir / "01_成交额.png"
            self.assertTrue(exported.is_file())
            self.assertFalse(source_path.exists())
            with Image.open(exported) as result:
                self.assertEqual(result.size, (100, 88))
                self.assertEqual(result.getpixel((50, 0))[:3], (0, 255, 0))

    def test_titleless_export_removes_title_separator_at_crop_boundary(self) -> None:
        """标题栏底部分隔线不得成为独立小图顶端的细线。"""
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            source_path = output_dir / "含标题分隔线源图.png"
            image = Image.new("RGB", (100, 100), "white")
            for x in range(100):
                image.putpixel((x, 12), (127, 127, 127))
                image.putpixel((x, 13), (0, 255, 0))
            image.save(source_path)

            MODULE.export_numbered_titleless_small_charts(
                output_dir,
                chart_specs=((1, "成交额", source_path.name, 0.12),),
            )

            with Image.open(output_dir / "01_成交额.png") as result:
                self.assertEqual(result.size, (100, 88))
                self.assertEqual(result.getpixel((50, 0))[:3], (0, 255, 0))

    def test_default_numbering_follows_swapped_long_report_order(self) -> None:
        """估值修复指数和平价分类换位后，独立小图编号应同步。"""
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            repair_source = output_dir / "转债估值修复指数.png"
            parity_source = output_dir / "分平价多因子修正拟合溢价率.png"
            Image.new("RGB", (100, 100), "white").save(repair_source)
            Image.new("RGB", (100, 100), "white").save(parity_source)

            MODULE.export_numbered_titleless_small_charts(
                output_dir,
                chart_specs=(
                    MODULE.SMALL_CHART_EXPORT_SPECS[6],
                    MODULE.SMALL_CHART_EXPORT_SPECS[9],
                ),
            )

            self.assertTrue((output_dir / "07_转债估值修复指数.png").is_file())
            self.assertTrue((output_dir / "10_平价分类拟合溢价率.png").is_file())

    def test_cleanup_removes_pre_swap_numbered_outputs(self) -> None:
        """同日重跑时不应保留换位前的旧编号小图。"""
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            old_parity = output_dir / "07_平价分类拟合溢价率.png"
            old_repair = output_dir / "10_转债估值修复指数.png"
            current_repair = output_dir / "07_转债估值修复指数.png"
            for path in (old_parity, old_repair, current_repair):
                path.write_bytes(b"generated chart")

            MODULE.remove_obsolete_outputs(output_dir, output_dir / "底稿.xlsx")

            self.assertFalse(old_parity.exists())
            self.assertFalse(old_repair.exists())
            self.assertTrue(current_repair.exists())


if __name__ == "__main__":
    unittest.main()
