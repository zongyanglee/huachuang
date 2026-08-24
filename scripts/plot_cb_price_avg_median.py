from __future__ import annotations

import json
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
STATS_JSON = ROOT / "outputs" / "cb_market_stats_2018" / "stats.json"
FONT_PATH = ROOT / "assets/fonts/KaiTi_GB2312.ttf"

AVG_PRICE_COLOR = "#E6121B"
MEDIAN_PRICE_COLOR = "#0262BA"


def main() -> None:
    if not STATS_JSON.exists():
        raise FileNotFoundError(f"未找到统计数据文件：{STATS_JSON}")
    if not FONT_PATH.exists():
        raise FileNotFoundError(f"未找到字体文件：{FONT_PATH}")

    payload = json.loads(STATS_JSON.read_text(encoding="utf-8"))
    df = pd.DataFrame(payload["rows"])
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date")

    latest_date = pd.to_datetime(payload["audit"]["last_date"])
    date_tag = latest_date.strftime("%Y%m%d")
    output_dir = ROOT / "runs" / "weekly" / f"鹏华周报{date_tag}"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "转债平均价格与价格中位数.png"

    font_prop = fm.FontProperties(fname=str(FONT_PATH))
    fm.fontManager.addfont(str(FONT_PATH))
    plt.rcParams["font.family"] = font_prop.get_name()
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(14, 7), dpi=180)

    line_avg, = ax.plot(
        df["date"],
        df["avg_price"],
        color=AVG_PRICE_COLOR,
        linewidth=2.2,
        label="平均价格",
    )
    line_median, = ax.plot(
        df["date"],
        df["median_price"],
        color=MEDIAN_PRICE_COLOR,
        linewidth=2.2,
        label="价格中位数",
    )

    ax.set_ylabel("价格", fontproperties=font_prop, fontsize=16, color="black")
    ax.tick_params(axis="both", colors="black")
    for spine in ax.spines.values():
        spine.set_color("black")

    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_minor_locator(mdates.MonthLocator(bymonth=(1, 7)))

    ax.grid(True, which="major", axis="both", color="#D9D9D9", linewidth=0.8, alpha=0.75)
    ax.grid(True, which="minor", axis="x", color="#ECECEC", linewidth=0.5, alpha=0.6)
    ax.set_xlim(df["date"].min(), df["date"].max())

    legend = ax.legend(
        [line_avg, line_median],
        [line_avg.get_label(), line_median.get_label()],
        loc="upper left",
        frameon=False,
        prop=font_prop,
        fontsize=15,
    )
    for text in legend.get_texts():
        text.set_fontproperties(font_prop)
        text.set_fontsize(15)

    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(font_prop)
        label.set_fontsize(14)

    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    print(json.dumps({
        "latest_date": latest_date.strftime("%Y-%m-%d"),
        "output_dir": str(output_dir),
        "output_path": str(output_path),
        "rows": int(len(df)),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
