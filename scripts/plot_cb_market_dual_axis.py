from __future__ import annotations

import json
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.ticker import PercentFormatter


ROOT = Path(__file__).resolve().parents[1]
STATS_JSON = ROOT / "outputs" / "cb_market_stats_2018" / "stats.json"
FONT_PATH = ROOT / "assets/fonts/KaiTi_GB2312.ttf"

PRICE_COLOR = "#E6121B"
PREMIUM_COLOR = "#0262BA"


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
    output_path = output_dir / "转债平均价格与平均转股溢价率.png"

    font_prop = fm.FontProperties(fname=str(FONT_PATH))
    fm.fontManager.addfont(str(FONT_PATH))
    font_name = font_prop.get_name()
    plt.rcParams["font.family"] = font_name
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax_price = plt.subplots(figsize=(14, 7), dpi=180)
    ax_premium = ax_price.twinx()

    line_price, = ax_price.plot(
        df["date"],
        df["avg_price"],
        color=PRICE_COLOR,
        linewidth=2.2,
        label="平均价格",
    )
    line_premium, = ax_premium.plot(
        df["date"],
        df["avg_premium"],
        color=PREMIUM_COLOR,
        linewidth=2.2,
        label="平均转股溢价率",
    )

    ax_price.set_ylabel("平均价格", fontproperties=font_prop, fontsize=16, color="black")
    ax_premium.set_ylabel("平均转股溢价率", fontproperties=font_prop, fontsize=16, color="black")

    ax_price.tick_params(axis="both", colors="black")
    ax_premium.tick_params(axis="both", colors="black")
    for spine in ax_price.spines.values():
        spine.set_color("black")
    for spine in ax_premium.spines.values():
        spine.set_color("black")
    ax_premium.yaxis.set_major_formatter(PercentFormatter(xmax=1, decimals=0))

    ax_price.xaxis.set_major_locator(mdates.YearLocator())
    ax_price.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax_price.xaxis.set_minor_locator(mdates.MonthLocator(bymonth=(1, 7)))

    ax_price.grid(True, which="major", axis="both", color="#D9D9D9", linewidth=0.8, alpha=0.75)
    ax_price.grid(True, which="minor", axis="x", color="#ECECEC", linewidth=0.5, alpha=0.6)
    ax_price.set_xlim(df["date"].min(), df["date"].max())

    lines = [line_price, line_premium]
    labels = [line.get_label() for line in lines]
    legend = ax_price.legend(
        lines,
        labels,
        loc="upper left",
        frameon=False,
        prop=font_prop,
        fontsize=15,
    )
    for text in legend.get_texts():
        text.set_fontproperties(font_prop)
        text.set_fontsize(15)

    for axis in (ax_price, ax_premium):
        for label in axis.get_xticklabels() + axis.get_yticklabels():
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
