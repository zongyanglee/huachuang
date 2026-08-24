"""
convert_media.py - 将 docx 媒体转换为 PPT 兼容的 PNG

功能概述：
    1. 遍历输入目录的所有媒体文件
    2. PNG/JPEG：直接复制
    3. SVG：使用 cairosvg 转换为 PNG（降级到 rsvg-convert / inkscape）
    4. EMF：使用 libreoffice 转换为 PNG
    5. 过滤：删除过小或尺寸异常的装饰性图

输入：<input_dir>（如 extracted/media）
输出：<output_dir>（如 extracted/media_png）
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

# ---- 1. 常量配置区 ----

# 支持的格式
SUPPORTED_INPUTS = {".png", ".jpeg", ".jpg", ".svg", ".emf", ".wmf"}
PASSTHROUGH_FORMATS = {".png", ".jpeg", ".jpg"}

# 文件大小阈值（字节）
MIN_FILE_SIZE = 5 * 1024  # 5KB

# 图片最小尺寸（像素）
MIN_WIDTH_PX = 200
MIN_HEIGHT_PX = 80


# ---- 2. 转换函数 ----

def convert_svg_to_png(svg_path: Path, png_path: Path) -> bool:
    """将 SVG 转为 PNG。返回是否成功。"""
    # 优先 cairosvg
    try:
        import cairosvg
        cairosvg.svg2png(url=str(svg_path), write_to=str(png_path), output_width=1600)
        return True
    except ImportError:
        pass
    except Exception:
        pass

    # 降级到 rsvg-convert
    try:
        result = subprocess.run(
            ["rsvg-convert", "-w", "1600", "-o", str(png_path), str(svg_path)],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode == 0:
            return True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    # 降级到 inkscape
    try:
        result = subprocess.run(
            ["inkscape", str(svg_path), "--export-type=png",
             "--export-filename", str(png_path), "-w", "1600"],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode == 0:
            return True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    return False


def convert_emf_to_png(emf_path: Path, png_path: Path) -> bool:
    """将 EMF 转为 PNG（用 libreoffice）。返回是否成功。"""
    try:
        # 用 libreoffice 转 EMF 为 PNG
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "png",
             "--outdir", str(png_path.parent), str(emf_path)],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode == 0:
            # libreoffice 会用原文件名+.png 保存
            expected = png_path.parent / (emf_path.stem + ".png")
            if expected.exists() and expected != png_path:
                shutil.move(str(expected), str(png_path))
            return png_path.exists()
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return False


def get_image_dimensions(img_path: Path) -> tuple[int, int]:
    """获取图片的像素尺寸"""
    try:
        from PIL import Image
        with Image.open(img_path) as im:
            return im.size
    except Exception:
        return (0, 0)


# ---- 3. 主流程 ----

def main() -> int:
    parser = argparse.ArgumentParser(description="将 docx 媒体转换为 PNG")
    parser.add_argument("--input", required=True, help="输入媒体目录")
    parser.add_argument("--output", required=True, help="输出 PNG 目录")
    parser.add_argument("--min-size-kb", type=int, default=5, help="最小文件大小（KB）")
    args = parser.parse_args()

    input_dir = Path(args.input)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not input_dir.exists():
        print(f"[ERROR] 输入目录不存在: {input_dir}", file=sys.stderr)
        return 1

    min_bytes = args.min_size_kb * 1024

    stats = {
        "total": 0,
        "copied": 0,
        "converted_svg": 0,
        "converted_emf": 0,
        "filtered_small": 0,
        "filtered_tiny": 0,
        "failed": 0,
    }

    for src in sorted(input_dir.iterdir()):
        if not src.is_file():
            continue
        ext = src.suffix.lower()
        if ext not in SUPPORTED_INPUTS:
            continue
        stats["total"] += 1

        # 文件大小过滤
        if src.stat().st_size < min_bytes:
            print(f"  [SKIP small] {src.name} ({src.stat().st_size} bytes)")
            stats["filtered_small"] += 1
            continue

        dst = output_dir / (src.stem + ".png")

        if ext in PASSTHROUGH_FORMATS:
            shutil.copy(src, dst)
            stats["copied"] += 1
        elif ext == ".svg":
            if convert_svg_to_png(src, dst):
                stats["converted_svg"] += 1
            else:
                print(f"  [WARN] SVG 转换失败: {src.name}")
                stats["failed"] += 1
                continue
        elif ext in {".emf", ".wmf"}:
            if convert_emf_to_png(src, dst):
                stats["converted_emf"] += 1
            else:
                print(f"  [WARN] EMF 转换失败: {src.name}")
                stats["failed"] += 1
                continue

        # 尺寸过滤
        w, h = get_image_dimensions(dst)
        if w < MIN_WIDTH_PX or h < MIN_HEIGHT_PX:
            print(f"  [SKIP tiny] {dst.name} ({w}x{h})")
            dst.unlink()
            stats["filtered_tiny"] += 1
            continue

        print(f"  [OK] {src.name} -> {dst.name} ({w}x{h})")

    print(f"\n[OK] 完成：")
    print(f"  总数: {stats['total']}")
    print(f"  复制: {stats['copied']}")
    print(f"  SVG 转换: {stats['converted_svg']}")
    print(f"  EMF 转换: {stats['converted_emf']}")
    print(f"  过滤（小文件）: {stats['filtered_small']}")
    print(f"  过滤（小尺寸）: {stats['filtered_tiny']}")
    print(f"  失败: {stats['failed']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
