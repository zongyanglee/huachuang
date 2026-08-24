from __future__ import annotations

import argparse
import math
import re
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.font_manager import FontProperties
from matplotlib.ticker import MultipleLocator


RED = "#E6121B"
INDEX_MAP = [
    ("沪深300", "沪深300"),
    ("中证500", "中证500"),
    ("中证1000", "中证1000"),
    ("中证2000", "中证2000"),
    ("中证转债", "转债指数"),
]


def monthly_files(root: Path) -> list[Path]:
    return sorted(
        path
        for path in root.glob("[0-9][0-9][0-9][0-9]/*.parquet")
        if re.fullmatch(r"\d{6}", path.stem)
    )


def load_recent_indices(root: Path) -> pd.DataFrame:
    paths = monthly_files(root)[-2:]
    records = []
    parquet_names = [parquet_name for _, parquet_name in INDEX_MAP]
    for path in paths:
        frame = pd.read_parquet(path)
        indices = frame.loc[
            (frame["__sheet_name"] == "指数") & frame["__row_id"].isin(parquet_names)
        ].set_index("__row_id")
        missing = set(parquet_names) - set(indices.index)
        if missing:
            raise RuntimeError(f"{path.name}缺少指数：{sorted(missing)}")
        for column in frame.columns[2:]:
            date = pd.Timestamp(column).normalize()
            for name in parquet_names:
                value = pd.to_numeric(indices.at[name, column], errors="coerce")
                if pd.isna(value):
                    raise RuntimeError(f"{path.name}的{name}在{column}为空")
                records.append((date, name, float(value)))
    return (
        pd.DataFrame(records, columns=["date", "index", "value"])
        .drop_duplicates(["date", "index"], keep="last")
        .pivot(index="date", columns="index", values="value")
        .sort_index()
    )


def calculate_weekly_returns(data: pd.DataFrame) -> tuple[pd.Timestamp, pd.Timestamp, pd.Series]:
    end_date = data.index.max()
    week_start = end_date - pd.Timedelta(days=end_date.weekday())
    prior_dates = data.index[data.index < week_start]
    if len(prior_dates) == 0:
        raise RuntimeError("没有找到上周开始前的基准交易日")
    base_date = prior_dates.max()
    parquet_order = [parquet_name for _, parquet_name in INDEX_MAP]
    returns = (data.loc[end_date, parquet_order] / data.loc[base_date, parquet_order] - 1) * 100
    returns.index = [display_name for display_name, _ in INDEX_MAP]
    return base_date, end_date, returns


def make_chart(returns: pd.Series, output: Path) -> None:
    chinese_font = FontProperties(fname=r"C:\Windows\Fonts\simsun.ttc", size=13)
    latin_font = FontProperties(fname=r"C:\Windows\Fonts\times.ttf", size=12.5)

    fig, axis = plt.subplots(figsize=(8, 4.75), dpi=250)
    fig.patch.set_facecolor("white")
    axis.set_facecolor("white")

    x = np.arange(len(returns))
    bars = axis.bar(x, returns.to_numpy(), width=0.52, color=RED, edgecolor="none")
    axis.axhline(0, color="black", linewidth=0.8)
    axis.set_xticks(x, returns.index.tolist())
    for label in axis.get_xticklabels():
        label.set_fontproperties(chinese_font)
    for label in axis.get_yticklabels():
        label.set_fontproperties(latin_font)

    minimum = min(0.0, float(returns.min()))
    maximum = max(0.0, float(returns.max()))
    span = max(maximum - minimum, 1.0)
    lower = 0.0 if minimum >= 0 else math.floor((minimum - span * 0.12) / 2) * 2
    upper = 0.0 if maximum <= 0 else max(2.0, math.ceil((maximum + span * 0.18) / 2) * 2)
    axis.set_ylim(lower, upper)
    axis.yaxis.set_major_locator(MultipleLocator(2))
    axis.tick_params(axis="both", which="major", labelsize=12.5, length=4, width=0.8)

    offset = span * 0.035
    for bar, value in zip(bars, returns.to_numpy()):
        if value >= 0:
            y = value + offset
            vertical = "bottom"
        else:
            y = value - offset
            vertical = "top"
        axis.text(
            bar.get_x() + bar.get_width() / 2,
            y,
            f"{value:.2f}",
            ha="center",
            va=vertical,
            fontproperties=latin_font,
            color="black",
        )

    axis.grid(False)
    axis.spines["top"].set_visible(False)
    axis.spines["right"].set_visible(False)
    axis.spines["left"].set_linewidth(0.7)
    axis.spines["bottom"].set_linewidth(0.7)
    axis.margins(x=0.08)
    fig.subplots_adjust(left=0.105, right=0.97, top=0.95, bottom=0.17)
    output.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output, dpi=250, facecolor="white")
    plt.close(fig)


def main() -> None:
    parser = argparse.ArgumentParser(description="生成人保转债周报图2：主要指数上周涨跌幅")
    parser.add_argument("parquet_root", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    data = load_recent_indices(args.parquet_root.resolve())
    base_date, end_date, returns = calculate_weekly_returns(data)
    make_chart(returns, args.output.resolve())
    print(f"计算区间：{base_date:%Y-%m-%d}至{end_date:%Y-%m-%d}")
    for name, value in returns.items():
        print(f"{name}: {value:.2f}%")
    print(f"已生成：{args.output.resolve()}")


if __name__ == "__main__":
    main()
