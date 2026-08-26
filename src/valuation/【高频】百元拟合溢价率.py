#!/usr/bin/env python
# coding: utf-8

from configparser import ConfigParser
from datetime import datetime
from pathlib import Path
import time

from iFinDPy import (
    THS_BD,
    THS_DR,
    THS_DS,
    THS_GetErrorInfo,
    THS_HF,
    THS_DataStatistics,
    THS_Date_Offset,
    THS_iFinDLogin,
)
import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from PIL import Image
from scipy.integrate import quad
from scipy.optimize import curve_fit
from tqdm import tqdm


DAYS_TODAY = 0
DAYS_BACKWARDS = 0
INTERVAL = 3
DROP = 0

RED = "#E6121B"
BLUE = "#0262BA"
GRAY = "#A6A6A6"
PINK = "#E6B9B8"
SKY = "#B7DEE8"
ORANGE = "#F79646"
WORKSPACE_ROOT = Path(__file__).resolve().parents[2]
RUNS_DAILY_ROOT = WORKSPACE_ROOT / "runs" / "daily"
IFIND_CREDENTIAL_FILE = WORKSPACE_ROOT / "private/ifind账号.txt"


def load_ifind_credentials() -> tuple[str, str]:
    """从项目目录的 ifind账号.txt 读取统一登录账号。"""
    if not IFIND_CREDENTIAL_FILE.is_file():
        raise FileNotFoundError(f"未找到iFinD账号文件：{IFIND_CREDENTIAL_FILE}")
    config = ConfigParser(interpolation=None)
    config.read(IFIND_CREDENTIAL_FILE, encoding="utf-8")
    username = config.get("ifind", "username", fallback="").strip()
    password = config.get("ifind", "password", fallback="").strip()
    if not username or not password:
        raise RuntimeError("ifind账号.txt中的[ifind] username或password为空")
    return username, password


def ths_login_demo():
    username, password = load_ifind_credentials()
    ths_login = THS_iFinDLogin(username, password)
    if ths_login != 0:
        error_info = THS_GetErrorInfo(ths_login)
        if isinstance(error_info, dict):
            error_type = error_info.get("errmsg", "未知错误")
        else:
            error_type = str(error_info)
        raise SystemExit(
            f"iFinD 登录失败，程序已终止（错误码：{ths_login}，错误类别：{error_type}）"
        )
    print_data_statistics()
    print("登录成功")


def make_output_dir(mmdd_today: str) -> Path:
    root = RUNS_DAILY_ROOT / f"{mmdd_today}数据更新"
    folder = root / "日内估值数据更新"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def print_data_statistics():
    try:
        result = THS_DataStatistics()
        tables = result.get("tables", {}) if isinstance(result, dict) else {}
        if not tables:
            detail = result.get("errmsg", "未返回额度数据") if isinstance(result, dict) else str(result)
            print(f"[警告] iFinD使用额度查询失败：{detail}")
            return
        print("iFinD使用额度：")
        for key, value in tables.items():
            ratio = value.get("ratio", "N/A") if isinstance(value, dict) else value
            print(f"{key} 已用：{ratio}")
    except Exception as exc:
        print(f"[警告] iFinD使用额度查询失败：{exc}")


def get_trade_dates():
    last_date = datetime.now().strftime("%Y-%m-%d")
    last_date = THS_Date_Offset(
        "212001",
        f"dateType:0,period:D,offset:{DAYS_TODAY},dateFormat:0,output:singledate",
        last_date,
    ).data

    days = THS_Date_Offset(
        "212001",
        f"dateType:0,period:D,offset:{DAYS_BACKWARDS},dateFormat:0,output:sequencedate",
        last_date,
    ).data

    days = [datetime.strptime(date, "%Y-%m-%d").date() for date in days.split(",")]
    start_date = days[0]
    last_date = days[-1]

    current_time = datetime.now().strftime("%H:%M:%S")
    demo = THS_HF(
        "000001.SH",
        "close",
        f"Fill:Original,Interval:{INTERVAL}",
        f"{start_date} 09:00:00",
        f"{last_date} {current_time}",
    ).data
    times = [datetime.strptime(t, "%Y-%m-%d %H:%M") for t in demo["time"].tolist()]
    current_time = times[-1]
    print("当前时间是", current_time)
    return start_date, last_date, current_time


def fetch_cb_basic_trade(start_date, last_date):
    start_cb_list = THS_DR(
        "p05004",
        f"jyzt=2;sfdb=1;jysc=1;edate={start_date.strftime('%Y%m%d')};gnfl=0",
        "jydm:Y,jydm_mc:Y",
        "format:dataframe",
    ).data
    last_cb_list = THS_DR(
        "p05004",
        f"jyzt=2;sfdb=1;jysc=1;edate={last_date.strftime('%Y%m%d')};gnfl=0",
        "jydm:Y,jydm_mc:Y",
        "format:dataframe",
    ).data
    cb_list = ", ".join(set(list(start_cb_list["jydm"]) + list(last_cb_list["jydm"])))

    cb_basic_trade = THS_BD(
        cb_list,
        "ths_convertible_debt_short_name_cbond;"
        "ths_stock_code_cbond;"
        "ths_stock_short_name_cbond;"
        "ths_issue_method_bond;"
        "ths_trading_status_bond;"
        "ths_last_td_date_convertible_cbond;"
        "ths_listed_date_bond;"
        "ths_object_the_sw_bond;"
        "ths_bond_latest_credict_rating_bond",
        f";;;;;;;100,{last_date.strftime('%Y-%m-%d')};100,100",
    ).data
    cb_basic_trade = cb_basic_trade.set_index("thscode").rename_axis("转债代码")
    cb_basic_trade.columns = [
        "转债简称",
        "正股代码",
        "正股简称",
        "发行方式",
        "交易状态",
        "最后交易日",
        "上市日期",
        "申万行业",
        "债项评级",
    ]

    index_to_drop = []
    future_date = last_date + pd.Timedelta(days=3650)
    cb_basic_trade["最后交易日"] = pd.to_datetime(cb_basic_trade["最后交易日"], errors="coerce")
    cb_basic_trade["最后交易日"] = cb_basic_trade["最后交易日"].fillna(pd.Timestamp(future_date))
    cb_basic_trade = cb_basic_trade[
        ~cb_basic_trade["发行方式"].str.contains("定向")
        & ~cb_basic_trade.index.str.contains("NQ")
        & (cb_basic_trade["交易状态"] == "正常上市")
        & ~cb_basic_trade.index.isin(index_to_drop)
        & (cb_basic_trade["最后交易日"] >= pd.Timestamp(last_date))
    ]
    return cb_basic_trade


def hf_pivot(code_list, indicator, start_date, end_time, reindex=None):
    data = THS_HF(
        code_list,
        indicator,
        f"Fill:Original,Interval:{INTERVAL}",
        f"{start_date} 09:00:00",
        f"{end_time}",
    ).data
    data = data.pivot_table(index=["time"], columns="thscode", values=indicator).T
    time_columns = pd.to_datetime(data.columns, errors="coerce")
    if time_columns.notna().all():
        # iFinD 偶尔会按不同证券返回错开 1 分钟的 3 分钟线，例如
        # 09:32/09:35 与 09:33/09:36。统一到标准刻度后合并非空值，
        # 避免同一截面被拆成两组，导致拟合时间点数量翻倍。
        data.columns = time_columns.round(f"{INTERVAL}min")
        data = data.T.groupby(level=0, sort=True).first().T
    if reindex is not None:
        data = data.reindex(reindex)
    return data


def calculate_intraday_turnover(cumulative_volume, balance):
    balance = pd.to_numeric(balance, errors="coerce").reindex(cumulative_volume.index)
    return cumulative_volume.div(balance, axis=0).mul(1 / 1000)


def fetch_market_data(cb_basic_trade, last_date, current_time):
    code_list = ",".join(cb_basic_trade.index.astype(str))
    stock_code_list = ",".join(cb_basic_trade["正股代码"].astype(str))
    end_time = f"{current_time}"

    outstandingbalance = THS_DS(
        code_list,
        "ths_bond_balance_cbond",
        "",
        "Fill:Blank,mode:thscode",
        str(last_date),
        str(last_date),
    ).data
    outstandingbalance = outstandingbalance.set_index("time").rename_axis("thscode").T
    outstandingbalance.rename(columns={last_date.strftime("%Y-%m-%d"): "余额"}, inplace=True)

    close = hf_pivot(code_list, "close", last_date, end_time, cb_basic_trade.index)

    conversion_price = THS_BD(code_list, "ths_conversion_price_cbond", str(last_date)).data
    conversion_price = conversion_price.set_index("thscode")
    conversion_price.index.name = "转债代码"
    conversion_price.rename(columns={"ths_conversion_price_cbond": "转股价"}, inplace=True)

    stock_close = hf_pivot(stock_code_list, "close", last_date, end_time, cb_basic_trade["正股代码"])
    stock_close = stock_close.reset_index()
    stock_close.insert(0, "转债代码", pd.Series(cb_basic_trade.index, name="转债代码"))
    stock_close = stock_close.set_index("转债代码").drop("正股代码", axis=1, errors="ignore")

    convalue = stock_close.div(conversion_price["转股价"], axis=0).mul(100)
    convpremiumratio = (close / convalue - 1) * 100

    volume = hf_pivot(code_list, "volume", last_date, end_time, cb_basic_trade.index)
    volume = volume.fillna(0).cumsum(axis=1)
    amount = hf_pivot(code_list, "amount", last_date, end_time, cb_basic_trade.index)
    turn = calculate_intraday_turnover(volume, outstandingbalance["余额"])

    stock_market_value = THS_BD(
        code_list,
        "ths_convertible_debt_stock_market_value_cbond",
        str(last_date),
    ).data
    stock_market_value = stock_market_value.set_index("thscode")
    stock_market_value.index.name = "转债代码"
    stock_market_value.rename(columns={"ths_convertible_debt_stock_market_value_cbond": "正股市值"}, inplace=True)
    stock_market_value = stock_market_value / 100000000

    remain_duration = THS_BD(code_list, "ths_remain_duration_y_bond", str(last_date)).data
    remain_duration = remain_duration.set_index("thscode")
    remain_duration.index.name = "转债代码"
    remain_duration.rename(columns={"ths_remain_duration_y_bond": "剩余期限"}, inplace=True)

    hs300 = THS_HF(
        "000300.SH",
        "close",
        f"Fill:Original,Interval:{INTERVAL}",
        f"{last_date} 09:00:00",
        end_time,
    ).data
    del hs300["thscode"]
    hs300 = hs300.set_index("time")
    hs300.index.name = "日期"
    hs300 = hs300.T

    return {
        "基础信息": cb_basic_trade,
        "转债余额": outstandingbalance,
        "转债价格": close,
        "转股价": conversion_price,
        "正股价": stock_close,
        "平价": convalue,
        "转股溢价率": convpremiumratio,
        "成交量": volume,
        "成交额": amount,
        "换手率": turn,
        "正股市值": stock_market_value,
        "剩余期限": remain_duration,
        "万得全A": hs300,
    }


def write_intraday_excel(new_data, folder_name, mmdd_today):
    output_path = folder_name / f"{mmdd_today}日内数据更新.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in new_data.items():
            df.to_excel(writer, sheet_name=sheet_name)


def inverse_cubic(x, a, b, c, d):
    return a / np.power(x, 3) + b / np.power(x, 2) + c / x + d


def fit_inverse_cubic_value(df_data, target_x):
    df_data = df_data.copy()
    df_data.replace(0, np.nan, inplace=True)
    df_data.dropna(inplace=True)
    x = df_data["平价"]
    y = df_data["转股溢价率"]
    popt, _ = curve_fit(inverse_cubic, x, y)
    return inverse_cubic(target_x, *popt), popt


def inverse_cubic_condition_number(df_data):
    df_data = df_data.copy()
    df_data.replace(0, np.nan, inplace=True)
    df_data.dropna(subset=["平价", "转股溢价率"], inplace=True)
    x = df_data["平价"].to_numpy(dtype=float)
    if len(x) < 4:
        return np.inf

    design = np.column_stack([
        1 / np.power(x, 3),
        1 / np.power(x, 2),
        1 / x,
        np.ones_like(x),
    ])
    return float(np.linalg.cond(design))


def trim_premium_outliers(df_data):
    df_data = df_data.dropna(subset=["平价", "转股溢价率", "换手率"]).copy()
    low = df_data["转股溢价率"].quantile(0.03)
    up = df_data["转股溢价率"].quantile(0.97)
    return df_data[
        (df_data["转股溢价率"] > low)
        & (df_data["转股溢价率"] < up)
    ].dropna(axis=0)


def fit_100_premium(new_data, result_path):
    df_plain = new_data["平价"].copy()
    df_premium = new_data["转股溢价率"].copy()
    df_turn = new_data["换手率"].copy()
    date_range = df_plain.columns

    result_columns = ["日期", "a", "b", "c", "d", "转股溢价率", "拟合公式", "积分"]
    rows = []

    for date in tqdm(date_range, total=len(date_range), desc="Processing"):
        df_date = pd.DataFrame({
            "平价": df_plain[date],
            "转股溢价率": df_premium[date],
            "换手率": df_turn[date],
        })
        df_date.replace("", np.nan, inplace=True)
        df_date.dropna(axis=0, how="all", inplace=True)
        df_date = df_date.astype(float)
        sample_mask = (
            (df_date["平价"] < 130)
            & (df_date["平价"] > 70)
            & (df_date["换手率"] < 50)
        )
        df_date = df_date[sample_mask]
        df_date = trim_premium_outliers(df_date)

        premium_100, popt = fit_inverse_cubic_value(df_date, 100)
        a, b, c, d = popt
        formula = f"转股溢价率 = {a:.2f}/平价^3 + {b:.2f}/平价^2 + {c:.2f}/平价 + {d:.2f}"
        integral_value, _ = quad(lambda x: inverse_cubic(x, a, b, c, d), 70, 130)
        rows.append([date, a, b, c, d, premium_100, formula, integral_value / 60])

    df_result = pd.DataFrame(rows, columns=result_columns).set_index("日期")
    with pd.ExcelWriter(result_path, engine="openpyxl") as writer:
        df_result.to_excel(writer, sheet_name="百元平价拟合溢价率", index=True)
    return df_result


def build_rolling_frame(new_data, date_range, i, extra_name=None, extra_values=None):
    window = date_range[i - 4:i + 1]
    df_data = pd.DataFrame({
        "平价": new_data["平价"][window].values.flatten(),
        "转股溢价率": new_data["转股溢价率"][window].values.flatten(),
        "换手率": new_data["换手率"][window].values.flatten(),
        "转债代码": new_data["平价"].index.repeat(len(window)),
    })
    if extra_name is not None and extra_values is not None:
        df_data[extra_name] = extra_values.repeat(len(window)).values

    df_data.replace("", np.nan, inplace=True)
    df_data.dropna(axis=0, how="all", inplace=True)
    numeric_cols = ["平价", "转股溢价率", "换手率"]
    if extra_name in {"剩余期限", "正股市值"}:
        numeric_cols.append(extra_name)
    df_data[numeric_cols] = df_data[numeric_cols].astype(float)
    return df_data


def fit_category_premiums(
    new_data,
    labels,
    selectors,
    target_x_values,
    extra_name=None,
    extra_values=None,
    base_selector=None,
    trim_before_selector=False,
    min_unique_bonds=None,
    max_condition_number=None,
):
    date_range = new_data["平价"].columns
    rows = []

    for i in tqdm(range(4, len(date_range)), desc="Processing"):
        df_data = build_rolling_frame(new_data, date_range, i, extra_name, extra_values)
        df_data = df_data[df_data["换手率"] < 50].copy()
        premium_values = []

        for label, selector, target_x in zip(labels, selectors, target_x_values):
            try:
                if trim_before_selector:
                    df_subset = df_data[base_selector(df_data)].copy()
                    df_subset = trim_premium_outliers(df_subset)
                    df_subset = df_subset[selector(df_subset)]
                else:
                    df_subset = df_data[selector(df_data)].copy()
                    if base_selector is not None:
                        df_subset = df_subset[base_selector(df_subset)].copy()
                    df_subset = trim_premium_outliers(df_subset)

                if min_unique_bonds is not None:
                    unique_bonds = df_subset["转债代码"].nunique()
                    if unique_bonds < min_unique_bonds:
                        raise ValueError(
                            f"独立转债数量不足: {unique_bonds} < {min_unique_bonds}"
                        )

                if max_condition_number is not None:
                    condition_number = inverse_cubic_condition_number(df_subset)
                    if (
                        not np.isfinite(condition_number)
                        or condition_number > max_condition_number
                    ):
                        raise ValueError(
                            f"拟合矩阵条件数过高: {condition_number:.3e}"
                        )

                premium, _ = fit_inverse_cubic_value(df_subset, target_x)
                premium_values.append(premium)
            except Exception as exc:
                tqdm.write(f"{date_range[i]} {label}拟合已跳过：{exc}")
                premium_values.append(np.nan)

        rows.append([pd.to_datetime(date_range[i])] + premium_values)

    return pd.DataFrame(rows, columns=["日期"] + list(labels))


def append_result_sheet(result_path, df_result, sheet_name):
    with pd.ExcelWriter(result_path, engine="openpyxl", mode="a") as writer:
        df_result.to_excel(writer, sheet_name=sheet_name, index=False)


def plot_main_premium(df_result, hs300, folder_name, mmdd_today):
    plot_times = pd.to_datetime(df_result.index)
    x = plot_times.strftime("%H:%M")[DROP:]
    y = df_result[["转股溢价率"]][DROP:]
    y2 = hs300.T.copy()
    y2.index = pd.to_datetime(y2.index)
    y2 = y2.reindex(plot_times).interpolate(method="time").ffill().bfill().iloc[DROP:]

    plt.rcParams["font.sans-serif"] = ["KaiTi_GB2312"]
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots()
    ax.plot(x, y, marker="", linestyle="-", color=BLUE, label="百元拟合转股溢价率")

    ax2 = ax.twinx()
    ax2.plot(x, y2, marker="", linestyle="-", color="#DCE6F1", label="沪深300指数（右）")

    plt.xticks(x[::5])
    x_min, x_max, y_min, y_max = ax.axis()
    fig.autofmt_xdate()

    ax.text(
        0.5 * (x_min + x_max),
        y_max - 0.02,
        "【华创固收】转债市场盘中百元平价拟合溢价率",
        ha="center",
        va="top",
        color="black",
    )

    min_y_value = y["转股溢价率"].min()
    ax.axhline(y=min_y_value, linestyle="--", color="red", label="")
    ax.text(x[DROP], min_y_value, f"{min_y_value:.2f}%", ha="left", va="bottom", color="red", fontsize=9)
    ax.text(x_max, min_y_value, "by 李宗阳", ha="right", va="top", color="gray", fontsize=5, alpha=0)
    ax.text(x_min, min_y_value, "盘中估值测算与盘后可能存在差异，仅供参考", ha="left", va="top", color="gray", fontsize=5, alpha=0)
    ax.text(0.5 * (x_min + x_max), 0.5 * (y_min + y_max), "华创固收", ha="center", va="center", color="gray", fontsize=50, alpha=0)

    handles1, labels1 = ax.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    legend = ax.legend(handles1 + handles2, labels1 + labels2, loc="upper right", fontsize=7)
    legend.set_bbox_to_anchor((1, 0.90))

    fig.savefig(folder_name / f"{mmdd_today}【华创固收】盘中百元平价拟合溢价.jpg", format="jpg", dpi=300, bbox_inches="tight")
    plt.close(fig)


def plot_category_change(df_result, labels, title, output_path, plot_labels=None):
    plot_labels = list(plot_labels or labels)
    plot_df = df_result.copy()
    for label in labels:
        plot_df[label] = plot_df[label] - plot_df[label].iloc[0]

    colors = [RED, BLUE, GRAY, PINK, SKY, ORANGE]
    plt.rcParams["font.sans-serif"] = ["KaiTi_GB2312"]
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots()
    x = np.arange(len(plot_df))
    for label, color in zip(plot_labels, colors):
        ax.plot(x, plot_df[label], color=color, label=label)

    ax.set_title(title)
    ax.legend()

    tick_interval = 5
    tick_locations = range(0, len(plot_df), tick_interval)
    tick_labels = [pd.to_datetime(plot_df["日期"].iloc[i]).strftime("%H:%M") for i in tick_locations]
    ax.set_xticks(tick_locations)
    ax.set_xticklabels(tick_labels)

    plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
    ax.axhline(y=0, linestyle="--", color="grey", label="")
    plt.tight_layout()
    fig.savefig(output_path, format="jpg", dpi=300, bbox_inches="tight")
    plt.close(fig)


def run_category_fit(new_data, result_path, folder_name, mmdd_today, config):
    df_result = fit_category_premiums(
        new_data,
        labels=config["labels"],
        selectors=config["selectors"],
        target_x_values=config["targets"],
        extra_name=config.get("extra_name"),
        extra_values=config.get("extra_values"),
        base_selector=config.get("base_selector"),
        trim_before_selector=config.get("trim_before_selector", False),
        min_unique_bonds=config.get("min_unique_bonds"),
        max_condition_number=config.get("max_condition_number"),
    )
    append_result_sheet(result_path, df_result, config["sheet_name"])
    plot_category_change(
        df_result,
        config["labels"],
        config["title"],
        folder_name / f"{mmdd_today}{config['image_name']}",
        plot_labels=config.get("plot_labels"),
    )
    return df_result


def get_category_configs(new_data):
    industry = new_data["基础信息"]["申万行业"]
    rating = new_data["基础信息"]["债项评级"]
    remain_duration = new_data["剩余期限"]["剩余期限"]
    stock_value = new_data["正股市值"]["正股市值"]

    base_0_200 = lambda df: (df["平价"] < 200) & (df["平价"] > 0)
    return [
        {
            "labels": ["70-90", "90-110", "110-130", "130-150", "150+"],
            "selectors": [
                lambda df: (df["平价"] < 90) & (df["平价"] > 70),
                lambda df: (df["平价"] < 110) & (df["平价"] > 90),
                lambda df: (df["平价"] < 130) & (df["平价"] > 110),
                lambda df: (df["平价"] < 150) & (df["平价"] > 130),
                lambda df: (df["平价"] < 170) & (df["平价"] > 150),
            ],
            "targets": [80, 100, 120, 140, 160],
            "sheet_name": "分平价百元平价拟合溢价率",
            "title": "平价分类拟合溢价率（PCT）",
            "image_name": "【华创固收】平价分类拟合溢价率.jpg",
            "plot_labels": ["70-90", "90-110", "110-130"],
        },
        {
            "labels": ["科技", "金融", "制造", "消费", "周期"],
            "selectors": [
                lambda df: df["行业"].isin(["传媒", "电子", "国防军工", "计算机", "通信"]),
                lambda df: df["行业"].isin(["非银金融", "银行"]),
                lambda df: df["行业"].isin(["电力设备", "机械设备", "汽车", "轻工制造"]),
                lambda df: df["行业"].isin(["农林牧渔", "纺织服饰", "家用电器", "商贸零售", "社会服务", "食品饮料", "医药生物", "美容护理"]),
                lambda df: df["行业"].isin(["基础化工", "钢铁", "公用事业", "环保", "建筑材料", "建筑装饰", "交通运输", "煤炭", "石油石化", "有色金属"]),
            ],
            "targets": [100, 100, 100, 100, 100],
            "extra_name": "行业",
            "extra_values": industry,
            "base_selector": base_0_200,
            "trim_before_selector": False,
            "min_unique_bonds": 4,
            "max_condition_number": 1e12,
            "sheet_name": "分板块百元平价拟合溢价率",
            "title": "板块分类拟合溢价率（PCT）",
            "image_name": "【华创固收】板块分类拟合溢价率.jpg",
        },
        {
            "labels": ["AAA/AA+", "AA/AA-", "A/A-"],
            "selectors": [
                lambda df: df["债项评级"].isin(["AAA", "AA+"]),
                lambda df: df["债项评级"].isin(["AA", "AA-"]),
                lambda df: df["债项评级"].isin(["A+", "A"]),
            ],
            "targets": [100, 100, 100],
            "extra_name": "债项评级",
            "extra_values": rating,
            "base_selector": base_0_200,
            "trim_before_selector": True,
            "sheet_name": "分评级百元平价拟合溢价率",
            "title": "评级分类拟合溢价率（PCT）",
            "image_name": "【华创固收】评级分类拟合溢价率.jpg",
        },
        {
            "labels": ["0-1", "1-2", "2-3", "3-4", "4-5", "5-6"],
            "selectors": [
                lambda df: base_0_200(df) & (df["剩余期限"] < 1) & (df["剩余期限"] > 0),
                lambda df: base_0_200(df) & (df["剩余期限"] < 2) & (df["剩余期限"] > 1),
                lambda df: base_0_200(df) & (df["剩余期限"] < 3) & (df["剩余期限"] > 2),
                lambda df: base_0_200(df) & (df["剩余期限"] < 4) & (df["剩余期限"] > 3),
                lambda df: base_0_200(df) & (df["剩余期限"] < 5) & (df["剩余期限"] > 4),
                lambda df: base_0_200(df) & (df["剩余期限"] < 6) & (df["剩余期限"] > 5),
            ],
            "targets": [100, 100, 100, 100, 100, 100],
            "extra_name": "剩余期限",
            "extra_values": remain_duration,
            "sheet_name": "剩余期限百元平价拟合溢价率",
            "title": "剩余期限分类拟合溢价率（PCT）",
            "image_name": "【华创固收】剩余期限分类拟合溢价率.jpg",
        },
        {
            "labels": ["0-50", "50-300", "300+"],
            "selectors": [
                lambda df: base_0_200(df) & (df["正股市值"] < 50) & (df["正股市值"] > 0),
                lambda df: base_0_200(df) & (df["正股市值"] < 300) & (df["正股市值"] > 50),
                lambda df: base_0_200(df) & (df["正股市值"] < np.inf) & (df["正股市值"] > 300),
            ],
            "targets": [100, 100, 100],
            "extra_name": "正股市值",
            "extra_values": stock_value,
            "sheet_name": "分市值百元平价拟合溢价率",
            "title": "正股市值分类拟合溢价率（PCT）",
            "image_name": "【华创固收】正股市值分类拟合溢价率.jpg",
        },
    ]


def combine_images(folder_name, mmdd_today):
    image_paths = [
        folder_name / f"{mmdd_today}【华创固收】盘中百元平价拟合溢价.jpg",
        folder_name / f"{mmdd_today}【华创固收】平价分类拟合溢价率.jpg",
        folder_name / f"{mmdd_today}【华创固收】板块分类拟合溢价率.jpg",
        folder_name / f"{mmdd_today}【华创固收】评级分类拟合溢价率.jpg",
        folder_name / f"{mmdd_today}【华创固收】剩余期限分类拟合溢价率.jpg",
        folder_name / f"{mmdd_today}【华创固收】正股市值分类拟合溢价率.jpg",
    ]
    output_path = folder_name / f"{mmdd_today}【华创固收】分类拟合溢价率.jpg"

    images = [Image.open(path) for path in image_paths]
    target_width, target_height = images[1].size
    images[0] = images[0].resize((target_width, target_height))

    new_image = Image.new("RGB", (2 * target_width, 3 * target_height))
    new_image.paste(images[0], (0, 0))
    new_image.paste(images[1], (target_width, 0))
    new_image.paste(images[2], (0, target_height))
    new_image.paste(images[3], (target_width, target_height))
    new_image.paste(images[4], (0, 2 * target_height))
    new_image.paste(images[5], (target_width, 2 * target_height))
    new_image.save(output_path, dpi=(300, 300))

    for image in images:
        image.close()


def print_runtime(start_time):
    end_time = time.time()
    total_time = end_time - start_time
    if total_time > 60:
        minutes = int(total_time // 60)
        seconds = int(total_time % 60)
        print(f"程序总运行时长：{int(minutes)} 分 {int(seconds)} 秒")
    else:
        print(f"程序总运行时长：{int(total_time)} 秒")
    print(time.strftime("%H:%M:%S", time.localtime(end_time)), "基础分时估值运行完成")
    print("基础分时估值文件已全部生成")


def main():
    ths_login_demo()
    print(DAYS_TODAY, DAYS_BACKWARDS, INTERVAL, DROP)

    start_time = time.time()
    mmdd_today = time.strftime("%m%d", time.localtime())
    folder_name = make_output_dir(mmdd_today)

    start_date, last_date, current_time = get_trade_dates()
    cb_basic_trade = fetch_cb_basic_trade(start_date, last_date)
    new_data = fetch_market_data(cb_basic_trade, last_date, current_time)

    write_intraday_excel(new_data, folder_name, mmdd_today)

    result_path = folder_name / f"{mmdd_today}百元平价溢价率拟合结果.xlsx"
    premium_result = fit_100_premium(new_data, result_path)
    plot_main_premium(premium_result, new_data["万得全A"], folder_name, mmdd_today)

    for config in get_category_configs(new_data):
        run_category_fit(new_data, result_path, folder_name, mmdd_today, config)

    combine_images(folder_name, mmdd_today)
    print_runtime(start_time)


if __name__ == "__main__":
    main()
