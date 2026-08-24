from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[1]
INPUT_XLSX = BASE_DIR / "抢券配售.xlsx"
HISTORY_DIR = BASE_DIR / "data/转债个券历史序列"
OUTPUT_DIR = BASE_DIR / "outputs" / "qiangquan_placement_exit_20260709"
PRICE_SHEET = "正股收盘价"


def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def price_on_date(series: pd.Series, target: pd.Timestamp) -> tuple[pd.Timestamp | None, float | None]:
    available = series.dropna()
    if available.empty or pd.isna(target):
        return None, None
    target = target.normalize()
    if target not in available.index:
        return None, None
    return target, float(available.loc[target])


def first_price_on_or_after(series: pd.Series, target: pd.Timestamp) -> tuple[pd.Timestamp | None, float | None]:
    available = series.dropna()
    if available.empty or pd.isna(target):
        return None, None
    hit = available[available.index >= target.normalize()]
    if hit.empty:
        return None, None
    return hit.index[0], float(hit.iloc[0])


def first_price_after(series: pd.Series, target: pd.Timestamp) -> tuple[pd.Timestamp | None, float | None]:
    available = series.dropna()
    if available.empty or pd.isna(target):
        return None, None
    hit = available[available.index > target.normalize()]
    if hit.empty:
        return None, None
    return hit.index[0], float(hit.iloc[0])


def first_recovery_on_or_before(
    series: pd.Series,
    start_after: pd.Timestamp,
    deadline: pd.Timestamp,
    cost_price: float,
) -> tuple[pd.Timestamp | None, float | None]:
    available = series.dropna()
    if available.empty:
        return None, None
    window = available[
        (available.index > start_after.normalize())
        & (available.index <= deadline.normalize())
        & (available >= cost_price)
    ]
    if window.empty:
        return None, None
    return window.index[0], float(window.iloc[0])


def load_stock_close_for_codes(codes: list[str]) -> pd.DataFrame:
    code_set = set(codes)
    blocks: list[pd.DataFrame] = []
    parquet_files = sorted(HISTORY_DIR.glob("[0-9][0-9][0-9][0-9]/*.parquet"))
    if not parquet_files:
        raise FileNotFoundError(f"No monthly parquet files found under {HISTORY_DIR}")

    for parquet_file in parquet_files:
        try:
            block = pd.read_parquet(
                parquet_file,
                filters=[("__sheet_name", "==", PRICE_SHEET)],
            )
        except Exception:
            raw = pd.read_parquet(parquet_file)
            block = raw[raw["__sheet_name"] == PRICE_SHEET]

        if block.empty:
            continue

        block = block[block["__row_id"].map(normalize_code).isin(code_set)].copy()
        if block.empty:
            continue

        block["__row_id"] = block["__row_id"].map(normalize_code)
        date_cols = [c for c in block.columns if c not in {"__sheet_name", "__row_id"}]
        if not date_cols:
            continue

        wide = block.set_index("__row_id")[date_cols]
        wide.columns = pd.to_datetime(wide.columns, errors="coerce")
        wide = wide.loc[:, wide.columns.notna()]
        blocks.append(wide)

    if not blocks:
        return pd.DataFrame(index=codes)

    prices = pd.concat(blocks, axis=1)
    prices = prices.loc[:, ~prices.columns.duplicated(keep="last")]
    prices = prices.sort_index(axis=1)
    prices = prices.apply(pd.to_numeric, errors="coerce")
    return prices.reindex(codes)


def build_results() -> tuple[pd.DataFrame, dict]:
    source = pd.read_excel(INPUT_XLSX, dtype={0: str})
    source.columns = [str(c).strip() for c in source.columns]
    code_col, name_col, approval_col, reg_col = source.columns[:4]

    source[code_col] = source[code_col].map(normalize_code)
    source[approval_col] = pd.to_datetime(source[approval_col], errors="coerce")
    source[reg_col] = pd.to_datetime(source[reg_col], errors="coerce")

    codes = source[code_col].tolist()
    prices = load_stock_close_for_codes(codes)

    rows: list[dict] = []
    for _, row in source.iterrows():
        code = row[code_col]
        name = row[name_col]
        approval_date = row[approval_col]
        reg_date = row[reg_col]
        series = prices.loc[code] if code in prices.index else pd.Series(dtype="float64")
        series = series.dropna()

        cost_date, cost_price = first_price_after(series, approval_date)
        cost_rule = "审批日后第一个交易日" if cost_price is not None else "缺少成本价"
        observe_date, observe_price = first_price_after(series, reg_date)
        deadline = reg_date + pd.DateOffset(months=1) if pd.notna(reg_date) else pd.NaT

        sell_date = None
        sell_price = None
        sell_reason = ""
        status = ""

        if cost_price is None:
            status = "缺少成本价"
            sell_reason = "审批日及之后无正股价格"
        elif pd.isna(reg_date):
            status = "缺少股权登记日"
            sell_reason = "无法确定卖出观察起点"
        else:
            recovery_date, recovery_price = first_recovery_on_or_before(series, reg_date, deadline, cost_price)
            if recovery_date is not None:
                sell_date = recovery_date
                sell_price = recovery_price
                sell_reason = "登记日后回到成本价"
                status = "已卖出"
            else:
                forced_date, forced_price = first_price_on_or_after(series, deadline)
                if forced_date is not None:
                    sell_date = forced_date
                    sell_price = forced_price
                    sell_reason = "满一个月出清"
                    status = "已卖出"
                else:
                    status = "未卖出"
                    sell_reason = "未回到成本价且数据尚未覆盖一个月出清日"

        latest_date = series.index.max() if not series.empty else None
        latest_price = float(series.loc[latest_date]) if latest_date is not None else None
        cost_date_diff = (cost_date - approval_date.normalize()).days if cost_date is not None and pd.notna(approval_date) else None
        cost_missing_note = "" if cost_price is not None else "parquet中该转债在审批日之后无正股收盘价"

        rows.append(
            {
                "转债代码": code,
                "转债名称": name,
                "发审委审批通过时间": approval_date.date().isoformat() if pd.notna(approval_date) else "",
                "股权登记日": reg_date.date().isoformat() if pd.notna(reg_date) else "",
                "成本价日期": cost_date.date().isoformat() if cost_date is not None else "",
                "成本价取价规则": cost_rule,
                "审批日至成本价日期差值_天": cost_date_diff,
                "成本价_正股收盘价": cost_price,
                "成本价缺失说明": cost_missing_note,
                "登记日后首个观察日": observe_date.date().isoformat() if observe_date is not None else "",
                "登记日后首个观察价": observe_price,
                "一个月出清参考日": deadline.date().isoformat() if pd.notna(deadline) else "",
                "卖出日期": sell_date.date().isoformat() if sell_date is not None else "",
                "卖出正股价": sell_price,
                "卖出原因": sell_reason,
                "状态": status,
                "从买入日到卖出日_天": (sell_date - cost_date).days if sell_date is not None and cost_date is not None else None,
                "股权登记日后持有_天": (sell_date - reg_date).days if sell_date is not None and pd.notna(reg_date) else None,
                "卖出较成本涨跌幅": (sell_price / cost_price - 1) if sell_price is not None and cost_price else None,
                "最新可用价格日": latest_date.date().isoformat() if latest_date is not None else "",
                "最新正股收盘价": latest_price,
                "最新较成本涨跌幅": (latest_price / cost_price - 1) if latest_price is not None and cost_price else None,
            }
        )

    result = pd.DataFrame(rows)
    summary = {
        "input_file": str(INPUT_XLSX),
        "history_dir": str(HISTORY_DIR),
        "price_sheet": PRICE_SHEET,
        "row_count": int(len(result)),
        "status_counts": result["状态"].value_counts(dropna=False).to_dict(),
        "sell_reason_counts": result["卖出原因"].value_counts(dropna=False).to_dict(),
        "cost_date_after_count": int((result["成本价取价规则"] == "审批日后第一个交易日").sum()),
        "cost_date_missing_count": int(result["成本价_正股收盘价"].isna().sum()),
        "cost_rule_counts": result["成本价取价规则"].value_counts(dropna=False).to_dict(),
        "min_approval_date": str(source[approval_col].min().date()),
        "max_approval_date": str(source[approval_col].max().date()),
        "min_registration_date": str(source[reg_col].min().date()),
        "max_registration_date": str(source[reg_col].max().date()),
        "max_price_date": str(prices.columns.max().date()) if len(prices.columns) else "",
        "assumptions": [
            "成本价取发审委审批通过日之后第一个有正股收盘价的交易日价格，不使用审批日当日价格。",
            "明细中保留成本价日期、成本价取价规则、审批日至成本价日期差值_天，用于识别买入价格对应的实际交易日。",
            "卖出观察从股权登记日之后第一个有正股价格的交易日开始。",
            "登记日后一个月内若正股收盘价大于等于成本价，则按首次回本日卖出。",
            "若一个月内未回本，则在股权登记日+1个自然月后第一个有正股价格的交易日出清。",
        ],
    }
    return result, summary


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    result, summary = build_results()
    result.to_csv(OUTPUT_DIR / "exit_results.csv", index=False, encoding="utf-8-sig")
    result.to_json(OUTPUT_DIR / "exit_results.json", orient="records", force_ascii=False, indent=2)
    (OUTPUT_DIR / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(result.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
