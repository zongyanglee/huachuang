#!/usr/bin/env python
# coding: utf-8
"""独立计算日度百元拟合溢价率，仅输出日期、拟合值和拟合函数。"""

from __future__ import annotations

import argparse
from configparser import ConfigParser
from datetime import date, datetime
from pathlib import Path

from iFinDPy import (
    THS_BD,
    THS_DR,
    THS_DS,
    THS_Date_Offset,
    THS_GetErrorInfo,
    THS_iFinDLogin,
)
import numpy as np
import pandas as pd
from scipy.optimize import curve_fit


WORKSPACE = Path(__file__).resolve().parents[2]
IFIND_CREDENTIAL_FILE = WORKSPACE / "private" / "ifind账号.txt"
LOGIN_OK_CODES = (0, -201)
EXCLUDED_CB_CODE = "128085.SZ"

FIT_PARITY_MIN = 70.0
FIT_PARITY_MAX = 130.0
FIT_TURNOVER_MAX = 50.0
PREMIUM_LOW_QUANTILE = 0.03
PREMIUM_HIGH_QUANTILE = 0.97


def load_credentials() -> tuple[str, str]:
    if not IFIND_CREDENTIAL_FILE.is_file():
        raise FileNotFoundError(f"未找到iFinD账号文件：{IFIND_CREDENTIAL_FILE}")
    config = ConfigParser(interpolation=None)
    config.read(IFIND_CREDENTIAL_FILE, encoding="utf-8")
    username = config.get("ifind", "username", fallback="").strip()
    password = config.get("ifind", "password", fallback="").strip()
    if not username or not password:
        raise RuntimeError("ifind账号.txt中的[ifind] username或password为空")
    return username, password


def error_message(error_code: int | None) -> str:
    if error_code is None:
        return "未知错误"
    try:
        info = THS_GetErrorInfo(error_code)
        if isinstance(info, dict):
            return str(info.get("errmsg", info))
        return str(info)
    except Exception:
        return f"错误码 {error_code}"


def api_data(result, api_name: str):
    if result is None:
        raise RuntimeError(f"{api_name}返回None")
    error_code = getattr(result, "errorcode", None)
    if error_code not in (None, *LOGIN_OK_CODES):
        raise RuntimeError(f"{api_name}失败：{error_message(error_code)}")
    data = getattr(result, "data", None)
    if data is None:
        raise RuntimeError(
            f"{api_name}未返回数据（errorcode={error_code}）：{error_message(error_code)}"
        )
    return data


def login() -> None:
    username, password = load_credentials()
    error_code = THS_iFinDLogin(username, password)
    if error_code not in LOGIN_OK_CODES:
        raise RuntimeError(f"iFinD登录失败：{error_message(error_code)}")
    print("iFinD登录成功")


def resolve_trade_dates(
    start_date: str | None,
    end_date: str | None,
    days_backwards: int,
) -> tuple[date, date]:
    reference_date = end_date or datetime.now().strftime("%Y-%m-%d")
    end_text = api_data(
        THS_Date_Offset(
            "212001",
            "dateType:0,period:D,offset:0,dateFormat:0,output:singledate",
            reference_date,
        ),
        "THS_Date_Offset(结束日)",
    )
    end_trade_date = datetime.strptime(str(end_text), "%Y-%m-%d").date()

    if start_date:
        start_trade_date = datetime.strptime(start_date, "%Y-%m-%d").date()
    else:
        start_text = api_data(
            THS_Date_Offset(
                "212001",
                f"dateType:0,period:D,offset:{days_backwards},dateFormat:0,output:singledate",
                end_trade_date.strftime("%Y-%m-%d"),
            ),
            "THS_Date_Offset(起始日)",
        )
        start_trade_date = datetime.strptime(str(start_text), "%Y-%m-%d").date()

    if start_trade_date > end_trade_date:
        raise ValueError("起始日期不能晚于结束日期")
    return start_trade_date, end_trade_date


def fetch_cb_codes(trade_date: date) -> set[str]:
    yyyymmdd = trade_date.strftime("%Y%m%d")
    data = api_data(
        THS_DR(
            "p00570",
            f"jyzt=未到期;sfdb=全部;jysc=全部;edate={yyyymmdd}",
            "jydm:Y,jydm_mc:Y,p00570_f001:Y,p00570_f019:Y",
            "format:dataframe",
        ),
        f"THS_DR(转债列表,{yyyymmdd})",
    )
    if not isinstance(data, pd.DataFrame) or "jydm" not in data.columns:
        raise RuntimeError(f"{yyyymmdd}转债列表返回格式异常")
    return set(data["jydm"].dropna().astype(str))


def filter_cb_codes(codes: set[str], end_date: date) -> tuple[list[str], pd.Series]:
    codes.discard(EXCLUDED_CB_CODE)
    code_text = ",".join(sorted(codes))
    if not code_text:
        raise RuntimeError("转债列表为空")

    data = api_data(
        THS_BD(
            code_text,
            "ths_convertible_debt_short_name_cbond;ths_stock_code_cbond;"
            "ths_stock_short_name_cbond;ths_issue_method_cbond;ths_trading_status_bond;"
            "ths_bond_balance_cbond;ths_listed_date_cbond",
            f";;;;;{end_date:%Y-%m-%d};",
        ),
        "THS_BD(转债基础信息)",
    )
    if not isinstance(data, pd.DataFrame) or "thscode" not in data.columns:
        raise RuntimeError("转债基础信息返回格式异常")

    basic = data.set_index("thscode").copy()
    basic.columns = [
        "转债简称",
        "正股代码",
        "正股简称",
        "发行方式",
        "交易状态",
        "转债余额",
        "上市日期",
    ]
    basic.index = basic.index.astype(str)
    basic = basic[~basic.index.str.contains("NQ", na=False)]
    basic = basic[~basic["发行方式"].astype(str).str.contains("定向", na=False)]
    basic = basic[~basic["交易状态"].astype(str).str.contains("终止上市", na=False)]
    listed_date = pd.to_datetime(basic["上市日期"].astype(str), errors="coerce")
    return basic.index.tolist(), listed_date


def fetch_daily_series(
    codes: list[str],
    indicator: str,
    start_date: date,
    end_date: date,
    parameter: str = "",
) -> pd.DataFrame:
    result = THS_DS(
        ",".join(codes),
        indicator,
        parameter,
        "Fill:Blank,mode:thscode",
        str(start_date),
        str(end_date),
    )
    data = api_data(result, f"THS_DS({indicator})")
    if not isinstance(data, pd.DataFrame) or "time" not in data.columns:
        raise RuntimeError(f"{indicator}返回格式异常")
    frame = data.set_index("time").T
    frame.index = frame.index.astype(str)
    frame.columns = pd.to_datetime(frame.columns)
    return frame.apply(pd.to_numeric, errors="coerce").sort_index(axis=1)


def inverse_cubic(x, a, b, c, d):
    return a / np.power(x, 3) + b / np.power(x, 2) + c / x + d


def fit_one_date(
    trade_date: pd.Timestamp,
    parity: pd.Series,
    premium: pd.Series,
    turnover: pd.Series,
    listed_date: pd.Series,
) -> tuple[float, str]:
    sample = pd.concat(
        [parity, premium, turnover],
        axis=1,
        keys=["平价", "转股溢价率", "换手率"],
    )
    sample = sample.apply(pd.to_numeric, errors="coerce")

    listed = listed_date.reindex(sample.index)
    valid_listing = listed.isna() | (trade_date.normalize() >= listed.dt.normalize())
    sample = sample[
        valid_listing
        & sample["平价"].between(FIT_PARITY_MIN, FIT_PARITY_MAX, inclusive="both")
        & sample["换手率"].notna()
        & sample["换手率"].le(FIT_TURNOVER_MAX)
    ].copy()

    low = sample["转股溢价率"].quantile(PREMIUM_LOW_QUANTILE)
    high = sample["转股溢价率"].quantile(PREMIUM_HIGH_QUANTILE)
    sample = sample[
        sample["转股溢价率"].gt(low) & sample["转股溢价率"].lt(high)
    ]
    sample = sample[["平价", "转股溢价率"]].replace(0, np.nan).dropna()
    if len(sample) < 5:
        raise ValueError(f"有效拟合样本不足：{len(sample)}")

    x = sample["平价"].to_numpy(dtype=float)
    y = sample["转股溢价率"].to_numpy(dtype=float)
    parameters, _ = curve_fit(inverse_cubic, x, y)
    a, b, c, d = (float(value) for value in parameters)
    fitted_premium = float(inverse_cubic(100.0, a, b, c, d))
    formula = (
        f"转股溢价率 = {a:.2f}/平价^3 + {b:.2f}/平价^2 "
        f"+ {c:.2f}/平价 + {d:.2f}"
    )
    return fitted_premium, formula


def calculate(start_date: date, end_date: date) -> pd.DataFrame:
    codes = fetch_cb_codes(start_date) | fetch_cb_codes(end_date)
    filtered_codes, listed_date = filter_cb_codes(codes, end_date)

    parity = fetch_daily_series(
        filtered_codes, "ths_transfer_value_cbond", start_date, end_date
    )
    premium = fetch_daily_series(
        filtered_codes,
        "ths_conversion_premium_rate_cbond",
        start_date,
        end_date,
    )
    turnover = fetch_daily_series(
        filtered_codes, "ths_turnover_ratio_cbond", start_date, end_date
    )

    common_codes = parity.index.intersection(premium.index).intersection(turnover.index)
    common_dates = parity.columns.intersection(premium.columns).intersection(
        turnover.columns
    ).sort_values()
    parity = parity.reindex(index=common_codes, columns=common_dates)
    premium = premium.reindex(index=common_codes, columns=common_dates)
    turnover = turnover.reindex(index=common_codes, columns=common_dates)

    if len(common_dates) == 0:
        raise RuntimeError("日度行情没有共同交易日期")

    first_turnover = turnover.iloc[:, 0]
    last_turnover = turnover.iloc[:, -1]
    inactive = (
        first_turnover.isna() & last_turnover.isna()
    ) | (first_turnover.eq(0) & last_turnover.eq(0))
    active_codes = common_codes[~inactive]

    rows = []
    for trade_date in common_dates:
        try:
            fitted_premium, formula = fit_one_date(
                trade_date,
                parity.loc[active_codes, trade_date],
                premium.loc[active_codes, trade_date],
                turnover.loc[active_codes, trade_date],
                listed_date,
            )
        except Exception as exc:
            print(f"[警告] {trade_date:%Y-%m-%d}拟合失败：{exc}")
            fitted_premium, formula = np.nan, ""
        rows.append(
            {
                "日期": trade_date.strftime("%Y-%m-%d"),
                "百元拟合溢价率": fitted_premium,
                "拟合函数": formula,
            }
        )
    return pd.DataFrame(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--start-date", help="起始日期，格式YYYY-MM-DD")
    parser.add_argument("--end-date", help="结束日期，格式YYYY-MM-DD；默认最近交易日")
    parser.add_argument(
        "--days-backwards",
        type=int,
        default=-1,
        help="未指定起始日期时，从结束交易日起回溯的交易日偏移，默认-5",
    )
    parser.add_argument("--output", type=Path, help="输出xlsx路径")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    login()
    start_date, end_date = resolve_trade_dates(
        args.start_date, args.end_date, args.days_backwards
    )
    print(f"计算区间：{start_date}至{end_date}")
    result = calculate(start_date, end_date)

    output_path = args.output
    if output_path is None:
        mmdd = end_date.strftime("%m%d")
        output_path = (
            WORKSPACE
            / "runs"
            / "daily"
            / f"{mmdd}数据更新"
            / f"{mmdd}百元拟合溢价率.xlsx"
        )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(output_path, sheet_name="百元拟合溢价率", index=False)
    print(result.to_string(index=False))
    print(f"结果已保存：{output_path}")


if __name__ == "__main__":
    main()
