from __future__ import annotations

import json
from datetime import timedelta
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
GROUPS = ["70以下", "70-90", "90-110", "110-130", "130-150", "150元以上"]
CSV_GROUPS = ["70以下", "70-90", "90-110", "110-130", "130-150", "150以上"]


def _fmt_pct(value: float, digits: int = 2) -> str:
    return f"{value * 100:.{digits}f}%"


def _fmt_pct_change(value: float) -> str:
    sign = "+" if value >= 0 else ""
    return f"{sign}{value * 100:.1f}"


def _previous_friday(date: pd.Timestamp) -> pd.Timestamp:
    days_since_friday = (date.weekday() - 4) % 7
    if days_since_friday == 0:
        days_since_friday = 7
    return date - timedelta(days=int(days_since_friday))


def _choose_base_date(dates: pd.Series, latest_date: pd.Timestamp) -> pd.Timestamp:
    target = _previous_friday(latest_date)
    candidates = dates[dates <= target]
    if candidates.empty:
        raise ValueError("没有找到可用的上周五或节前最近交易日")
    return candidates.max()


def main() -> None:
    latest_dir_candidates = sorted(ROOT.glob("鹏华周报20*"))
    if not latest_dir_candidates:
        raise FileNotFoundError("未找到鹏华周报输出目录")
    output_dir = latest_dir_candidates[-1]
    avg_csv = output_dir / "平价分组平均转股溢价率_2015以来.csv"
    pct_csv = output_dir / "平价分组平均转股溢价率分位数.csv"
    if not avg_csv.exists():
        avg_csv = output_dir / "平价分组平均转股溢价率.csv"
    if not avg_csv.exists() or not pct_csv.exists():
        raise FileNotFoundError("未找到平价分组均值或分位数 CSV，请先运行分组计算脚本")

    avg = pd.read_csv(avg_csv)
    pct = pd.read_csv(pct_csv)
    avg["date"] = pd.to_datetime(avg["date"])
    pct["date"] = pd.to_datetime(pct["date"])
    avg = avg.sort_values("date")
    pct = pct.sort_values("date")

    latest_date = avg["date"].max()
    base_date = _choose_base_date(avg["date"], latest_date)
    latest_row = avg.loc[avg["date"].eq(latest_date)].iloc[0]
    base_row = avg.loc[avg["date"].eq(base_date)].iloc[0]
    pct_row = pct.loc[pct["date"].eq(latest_date)].iloc[0]

    latest_values = [_fmt_pct(float(latest_row[col])) for col in CSV_GROUPS]
    changes = [_fmt_pct_change(float(latest_row[col]) - float(base_row[col])) for col in CSV_GROUPS]
    percentiles = [_fmt_pct(float(pct_row[col]), 1) for col in CSV_GROUPS]

    period = f"{base_date:%Y%m%d}-{latest_date:%Y%m%d}"
    group_text = "、".join(GROUPS)
    text = (
        f"{period}，最新平价{group_text}的平均转股溢价率分别为"
        f"{'、'.join(latest_values)}。"
        f"本周不同平价区间的转债溢价率走势不一，环比分别"
        f"{'、'.join(changes)} pct。"
        f"最新平价{group_text}的平均转股溢价率所处2015年以来历史分位数分别"
        f"{'、'.join(percentiles)}。"
    )

    output_path = output_dir / "平价分组平均转股溢价率文字描述.txt"
    output_path.write_text(text, encoding="utf-8")

    print(json.dumps({
        "latest_date": latest_date.strftime("%Y-%m-%d"),
        "base_date": base_date.strftime("%Y-%m-%d"),
        "output_path": str(output_path),
        "text": text,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
