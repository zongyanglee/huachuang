from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from matplotlib.ticker import FuncFormatter, MultipleLocator


SCRIPT_DIR = Path(__file__).resolve().parents[2]
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from update_renbao_figure1 import find_parquet_root, load_index_data


RED = "#E6121B"
BLUE = "#0262BA"


def expanded_limits(values, base_min: float, base_max: float, step: float) -> tuple[float, float]:
    minimum = float(values.min())
    maximum = float(values.max())
    low = base_min if minimum >= base_min else math.floor(minimum / step) * step - step
    high = base_max if maximum <= base_max else math.ceil(maximum / step) * step + step
    return low, high


def main() -> None:
    parser = argparse.ArgumentParser(description="由 parquet 生成周报图1高分辨率双轴折线图")
    parser.add_argument("workspace", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    data = load_index_data(find_parquet_root(args.workspace.resolve()))
    args.output.parent.mkdir(parents=True, exist_ok=True)

    chinese_font = FontProperties(fname=r"C:\Windows\Fonts\simsun.ttc", size=13)
    latin_font = FontProperties(fname=r"C:\Windows\Fonts\times.ttf", size=12.5)

    fig, left = plt.subplots(figsize=(8, 4.75), dpi=250)
    right = left.twinx()
    fig.patch.set_facecolor("white")
    left.set_facecolor("white")

    line_cb, = left.plot(
        data.index, data["转债指数"], color=RED, linewidth=2.2, label="中证转债"
    )
    line_wind, = right.plot(
        data.index, data["万得全A"], color=BLUE, linewidth=2.2, label="万得全A"
    )

    left.set_xlim(data.index.min(), data.index.max())
    left.set_ylim(*expanded_limits(data["转债指数"], 200, 600, 50))
    right.set_ylim(*expanded_limits(data["万得全A"], 3000, 7500, 500))
    left.yaxis.set_major_locator(MultipleLocator(50))
    right.yaxis.set_major_locator(MultipleLocator(500))
    right.yaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:,.0f}"))

    years = range(data.index.min().year, data.index.max().year + 1)
    ticks = [data.index.min().replace(year=year) for year in years]
    left.set_xticks(ticks)
    left.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))

    left.grid(False)
    right.grid(False)
    left.spines["top"].set_visible(False)
    right.spines["top"].set_visible(False)
    left.spines["right"].set_visible(False)
    right.spines["left"].set_visible(False)
    for axis in (left, right):
        axis.tick_params(axis="both", which="major", labelsize=12.5, length=4, width=0.8)
        for label in axis.get_xticklabels() + axis.get_yticklabels():
            label.set_fontproperties(latin_font)
        for side in ("bottom", "left", "right"):
            axis.spines[side].set_linewidth(0.7)
            axis.spines[side].set_color("black")

    legend = left.legend(
        [line_cb, line_wind],
        ["中证转债", "万得全A"],
        loc="upper center",
        bbox_to_anchor=(0.5, -0.17),
        ncol=2,
        frameon=False,
        handlelength=3.2,
        columnspacing=2.2,
        prop=chinese_font,
    )
    for handle in legend.legend_handles:
        handle.set_linewidth(2.5)

    fig.subplots_adjust(left=0.105, right=0.895, top=0.96, bottom=0.25)
    fig.savefig(args.output, dpi=250, facecolor="white")
    plt.close(fig)
    print(
        f"已生成 {args.output}：{data.index.min():%Y-%m-%d}至"
        f"{data.index.max():%Y-%m-%d}，{len(data)}点"
    )


if __name__ == "__main__":
    main()
