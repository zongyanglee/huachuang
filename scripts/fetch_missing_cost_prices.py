from __future__ import annotations

import json
import time
from datetime import timedelta
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[1]
RESULT_CSV = BASE_DIR / "outputs" / "qiangquan_placement_exit_20260709" / "exit_results.csv"
OUTPUT_CSV = BASE_DIR / "outputs" / "qiangquan_placement_exit_20260709" / "network_cost_price_overrides.csv"


BOND_TO_STOCK = {
    "128142.SZ": ("002946.SZ", "新乳业"),
    "128139.SZ": ("002965.SZ", "祥鑫科技"),
    "110077.SH": ("600461.SH", "洪城环境"),
    "113041.SH": ("601899.SH", "紫金矿业"),
    "113594.SH": ("603516.SH", "淳中科技"),
    "123059.SZ": ("300231.SZ", "银信科技"),
    "127019.SZ": ("000688.SZ", "国城矿业"),
    "128103.SZ": ("002360.SZ", "同德化工"),
    "113563.SH": ("603368.SH", "柳药集团"),
}


def eastmoney_secid(stock_code: str) -> str:
    code, market = stock_code.split(".")
    prefix = "1" if market == "SH" else "0"
    return f"{prefix}.{code}"


def fetch_daily_kline(stock_code: str, start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    params = {
        "secid": eastmoney_secid(stock_code),
        "fields1": "f1,f2,f3,f4,f5,f6",
        "fields2": "f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61",
        "klt": "101",
        "fqt": "0",
        "beg": start.strftime("%Y%m%d"),
        "end": end.strftime("%Y%m%d"),
    }
    url = "http://push2his.eastmoney.com/api/qt/stock/kline/get?" + urlencode(params)
    request = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(request, timeout=20) as response:
        payload = json.loads(response.read().decode("utf-8"))
    data = payload.get("data") or {}
    rows = []
    for item in data.get("klines") or []:
        parts = item.split(",")
        rows.append(
            {
                "date": pd.Timestamp(parts[0]),
                "open": float(parts[1]),
                "close": float(parts[2]),
                "high": float(parts[3]),
                "low": float(parts[4]),
                "volume": float(parts[5]),
                "amount": float(parts[6]),
            }
        )
    frame = pd.DataFrame(rows)
    if not frame.empty:
        frame = frame.sort_values("date").reset_index(drop=True)
    return frame


def main() -> None:
    results = pd.read_csv(RESULT_CSV)
    missing = results[results["状态"].eq("缺少成本价")].copy()
    output_rows = []

    for _, row in missing.iterrows():
        bond_code = row["转债代码"]
        stock_code, stock_name = BOND_TO_STOCK[bond_code]
        approval_date = pd.Timestamp(row["发审委审批通过时间"])
        start = approval_date - timedelta(days=10)
        end = approval_date + timedelta(days=10)
        kline = fetch_daily_kline(stock_code, start, end)
        time.sleep(0.25)

        prev_rows = kline[kline["date"] < approval_date]
        exact_rows = kline[kline["date"] == approval_date]
        next_rows = kline[kline["date"] > approval_date]

        exact = exact_rows.iloc[0] if not exact_rows.empty else None
        prev_row = prev_rows.iloc[-1] if not prev_rows.empty else None
        next_row = next_rows.iloc[0] if not next_rows.empty else None

        chosen = exact if exact is not None else next_row
        basis = "审批日当日收盘价" if exact is not None else "审批日为非交易日，采用下一交易日收盘价"

        output_rows.append(
            {
                "转债代码": bond_code,
                "转债名称": row["转债名称"],
                "正股代码": stock_code,
                "正股名称": stock_name,
                "发审委审批通过时间": approval_date.date().isoformat(),
                "网络成本价日期": chosen["date"].date().isoformat() if chosen is not None else "",
                "网络成本价": float(chosen["close"]) if chosen is not None else None,
                "网络成本价口径": basis if chosen is not None else "未获取到可用行情",
                "前一交易日": prev_row["date"].date().isoformat() if prev_row is not None else "",
                "前一交易日收盘价": float(prev_row["close"]) if prev_row is not None else None,
                "后一交易日": next_row["date"].date().isoformat() if next_row is not None else "",
                "后一交易日收盘价": float(next_row["close"]) if next_row is not None else None,
                "数据来源": "东方财富历史行情API(push2his.eastmoney.com)，不复权日K",
            }
        )

    output = pd.DataFrame(output_rows)
    output.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(output.to_string(index=False))
    print(f"saved: {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
