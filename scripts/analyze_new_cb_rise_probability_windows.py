from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib import font_manager
import numpy as np
import pandas as pd


WORKSPACE = Path(__file__).resolve().parents[1]
INPUT_DIR = WORKSPACE / "outputs" / "新券上市事件研究_溢价压缩与涨幅"
OUTPUT_DIR = WORKSPACE / "outputs" / "新券上市后上涨概率窗口"

FORWARD_DAYS = 20
FIRST_START_T = 6
LAST_START_T = 100
OBVIOUS_RISE = 0.05

COHORTS = (
    ("since_2017", "2017年以来", 901),
    ("since_2024_10", "2024年10月以来", 51),
)

TIME_BANDS = (
    (6, 20),
    (21, 40),
    (41, 60),
    (61, 80),
    (81, 100),
)


def configure_chinese_font() -> None:
    candidates = (
        "Microsoft YaHei",
        "Microsoft YaHei UI",
        "SimHei",
        "Noto Sans CJK SC",
        "Source Han Sans SC",
    )
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in installed:
            plt.rcParams["font.sans-serif"] = [name]
            break
    plt.rcParams["axes.unicode_minus"] = False


def load_panel(slug: str) -> pd.DataFrame:
    path = INPUT_DIR / f"{slug}_个券事件面板.parquet"
    panel = pd.read_parquet(path).copy()
    panel["交易日"] = pd.to_numeric(panel["交易日"], errors="coerce").astype("Int64")
    panel["价格较上市日涨跌幅"] = pd.to_numeric(
        panel["价格较上市日涨跌幅"], errors="coerce"
    )
    panel["价格指数"] = 1.0 + panel["价格较上市日涨跌幅"]
    panel["T"] = panel["交易日"] + 1
    return panel.dropna(subset=["转债代码", "T", "价格指数"])


def compute_forward_observations(panel: pd.DataFrame) -> pd.DataFrame:
    wide = panel.pivot(index="转债代码", columns="T", values="价格指数").sort_index(axis=1)
    rows: list[pd.DataFrame] = []
    for start_t in range(FIRST_START_T, LAST_START_T + 1):
        needed = list(range(start_t, start_t + FORWARD_DAYS + 1))
        available = [t for t in needed if t in wide.columns]
        if len(available) != len(needed):
            continue
        block = wide[needed].dropna()
        base = block[start_t]
        endpoint_return = block[start_t + FORWARD_DAYS] / base - 1.0
        max_return = block.loc[:, start_t + 1 : start_t + FORWARD_DAYS].max(axis=1) / base - 1.0
        rows.append(
            pd.DataFrame(
                {
                    "转债代码": block.index,
                    "起始交易日": start_t,
                    "未来20日终点收益率": endpoint_return.to_numpy(),
                    "未来20日内最大涨幅": max_return.to_numpy(),
                    "20日后上涨": endpoint_return.to_numpy() > 0,
                    "未来20日内明显上涨": max_return.to_numpy() >= OBVIOUS_RISE,
                }
            )
        )
    return pd.concat(rows, ignore_index=True)


def aggregate_daily(observations: pd.DataFrame) -> pd.DataFrame:
    daily = (
        observations.groupby("起始交易日", as_index=False)
        .agg(
            样本数=("转债代码", "size"),
            二十日后上涨概率=("20日后上涨", "mean"),
            未来二十日内涨超5概率=("未来20日内明显上涨", "mean"),
            二十日收益率中位数=("未来20日终点收益率", "median"),
            未来二十日最大涨幅中位数=("未来20日内最大涨幅", "median"),
        )
        .sort_values("起始交易日")
    )
    daily["二十日后上涨概率_5日均线"] = daily["二十日后上涨概率"].rolling(5, center=True, min_periods=1).mean()
    daily["未来二十日内涨超5概率_5日均线"] = (
        daily["未来二十日内涨超5概率"].rolling(5, center=True, min_periods=1).mean()
    )
    return daily


def aggregate_bands(observations: pd.DataFrame, cohort_label: str) -> pd.DataFrame:
    parts = []
    for lo, hi in TIME_BANDS:
        block = observations[observations["起始交易日"].between(lo, hi)]
        parts.append(
            {
                "样本": cohort_label,
                "起始时段": f"T+{lo}—T+{hi}",
                "区间起点": lo,
                "区间终点": hi,
                "债券-起点观测数": len(block),
                "20日后上涨概率": block["20日后上涨"].mean(),
                "未来20日内涨超5%概率": block["未来20日内明显上涨"].mean(),
                "20日收益率中位数": block["未来20日终点收益率"].median(),
                "未来20日最大涨幅中位数": block["未来20日内最大涨幅"].median(),
            }
        )
    return pd.DataFrame(parts)


def best_contiguous_window(daily: pd.DataFrame, metric: str, width: int = 20) -> tuple[int, int, float]:
    values = daily.set_index("起始交易日")[metric]
    rolling = values.rolling(width, min_periods=width).mean()
    end_t = int(rolling.idxmax())
    return end_t - width + 1, end_t, float(rolling.loc[end_t])


def make_chart(daily_by_slug: dict[str, pd.DataFrame], output_path: Path) -> None:
    configure_chinese_font()
    fig, axes = plt.subplots(2, 1, figsize=(13.2, 8.6), sharex=True, sharey=True)
    fig.patch.set_facecolor("white")

    line_colors = ("#155EEF", "#D92D20")
    for ax, (slug, label, expected_n) in zip(axes, COHORTS):
        data = daily_by_slug[slug]
        x = data["起始交易日"]
        ax.plot(
            x,
            data["二十日后上涨概率_5日均线"] * 100,
            color=line_colors[0],
            linewidth=2.4,
            label="20个交易日后价格上涨",
        )
        ax.plot(
            x,
            data["未来二十日内涨超5概率_5日均线"] * 100,
            color=line_colors[1],
            linewidth=2.4,
            label="未来20日内最高涨幅≥5%",
        )

        lo, hi, best = best_contiguous_window(data, "未来二十日内涨超5概率", width=20)
        ax.axvspan(lo, hi, color="#F79009", alpha=0.11, linewidth=0)
        ax.annotate(
            f"明显上涨概率最高的连续20日起点区间\nT+{lo}—T+{hi}（均值 {best:.1%}）",
            xy=((lo + hi) / 2, best * 100),
            xytext=((lo + hi) / 2, min(88, best * 100 + 14)),
            ha="center",
            va="bottom",
            fontsize=10.5,
            arrowprops=dict(arrowstyle="-|>", color="#B54708", lw=1.1),
            bbox=dict(boxstyle="round,pad=0.35", fc="#FFFAEB", ec="#FEDF89", alpha=0.96),
        )
        ax.set_title(f"{label}（{expected_n}只完整样本）", loc="left", fontsize=13, fontweight="bold")
        ax.set_ylabel("概率（%）")
        ax.set_ylim(20, 90)
        ax.set_yticks(np.arange(20, 91, 10))
        ax.grid(axis="y", color="#D0D5DD", linewidth=0.7, alpha=0.7)
        ax.spines[["top", "right", "left"]].set_visible(False)
        ax.tick_params(axis="y", length=0)
        ax.legend(loc="upper right", frameon=False, ncol=2, fontsize=10.5)

    axes[-1].set_xlabel("从上市后的第几个交易日开始观察（T+N）")
    axes[-1].set_xticks([6, 20, 40, 60, 80, 100])
    fig.suptitle("新券上市后，未来20个交易日的上涨概率", x=0.06, ha="left", fontsize=17, fontweight="bold")
    fig.text(
        0.06,
        0.925,
        "已剔除上市前5个交易日；曲线为每日概率的5日居中均线，橙色区为“涨超5%概率”最高的连续20日起点区间。",
        ha="left",
        fontsize=10.5,
        color="#475467",
    )
    fig.tight_layout(rect=(0.04, 0.055, 0.98, 0.9))
    fig.savefig(output_path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    all_bands = []
    summary_rows = []
    daily_by_slug: dict[str, pd.DataFrame] = {}

    for slug, label, expected_n in COHORTS:
        panel = load_panel(slug)
        observations = compute_forward_observations(panel)
        daily = aggregate_daily(observations)
        daily.insert(0, "样本", label)
        daily_by_slug[slug] = daily
        daily.to_csv(OUTPUT_DIR / f"{slug}_未来20日上涨概率.csv", index=False, encoding="utf-8-sig")

        bands = aggregate_bands(observations, label)
        all_bands.append(bands)

        pos_lo, pos_hi, pos_value = best_contiguous_window(daily, "二十日后上涨概率", width=20)
        jump_lo, jump_hi, jump_value = best_contiguous_window(daily, "未来二十日内涨超5概率", width=20)
        summary_rows.append(
            {
                "样本": label,
                "完整样本数": panel["转债代码"].nunique(),
                "预期完整样本数": expected_n,
                "20日后上涨概率最高的连续20日起点区间": f"T+{pos_lo}—T+{pos_hi}",
                "该区间日均上涨概率": pos_value,
                "未来20日内涨超5%概率最高的连续20日起点区间": f"T+{jump_lo}—T+{jump_hi}",
                "该区间日均涨超5%概率": jump_value,
            }
        )

    pd.concat(all_bands, ignore_index=True).to_csv(
        OUTPUT_DIR / "分时段汇总.csv", index=False, encoding="utf-8-sig"
    )
    pd.DataFrame(summary_rows).to_csv(OUTPUT_DIR / "结论汇总.csv", index=False, encoding="utf-8-sig")
    make_chart(daily_by_slug, OUTPUT_DIR / "新券上市后未来20日上涨概率_两组对比.png")

    print(pd.DataFrame(summary_rows).to_string(index=False))
    print()
    print(pd.concat(all_bands, ignore_index=True).to_string(index=False))


if __name__ == "__main__":
    main()
