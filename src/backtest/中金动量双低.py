"""中金动量双低策略回测。

研究用途，不涉及实盘交易。脚本使用当前标准 parquet 底稿：
``data/转债个券历史序列``，不修改通用回测合集。

策略口径：
1. 每 20 个交易日调仓，调仓日使用前一交易日截面。
2. 基础筛选：余额>=2、收盘价>=70、剩余期限>=0.5、正股市值>=30。
3. 剔除定价依据日已经公告赎回的转债；交易状态仅允许“交易/新股上市”。
4. 双低指标=收盘价+转股溢价率，仅保留双低指标>100，不设上限。
5. 动量条件：EXPMA5>EXPMA10>EXPMA20，且正股收盘价>EXPMA5。
6. 主体评级限制：AAA、AA+、AA、AA-、A+；债项评级不限制。
7. 按双低指标从低到高排序取指定容量，等权持有，缺失日收益按 0 处理。

默认执行用户指定的 9 组参数敏感性回测；传入 ``--single`` 时仅运行原始容量 50 配置。
"""

from __future__ import annotations

import argparse
import importlib.util
import sys
from dataclasses import dataclass
from pathlib import Path
from types import ModuleType
from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.font_manager import FontProperties


ROOT = Path(__file__).resolve().parents[2]
BASE_SCRIPT = ROOT / "src" / "backtest" / "【回测】回测合集.py"
DEFAULT_STRATEGY_NAME = "中金动量双低"


@dataclass
class StrategyConfig:
    strategy_name: str = DEFAULT_STRATEGY_NAME
    parquet_root: str = str(ROOT / "data" / "转债个券历史序列")
    backtest_start_date: str = "2015-01-01"
    rebalance_every_n_days: int = 20
    max_holdings: int = 50
    min_balance: float = 2.0
    min_price: float = 70.0
    min_remaining_years: float = 0.5
    min_stock_market_value: float = 30.0
    min_double_low: float = 100.0
    allowed_subject_ratings: tuple[str, ...] = ("AAA", "AA+", "AA", "AA-", "A+")
    benchmark_names: tuple[str, ...] = ("转债指数", "万得全A")
    excluded_parity_range: Optional[tuple[float, float]] = None


def load_backtest_module() -> ModuleType:
    spec = importlib.util.spec_from_file_location("integrated_cb_backtest_for_momentum_double_low", BASE_SCRIPT)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载回测框架: {BASE_SCRIPT}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def cross_section(mats: dict[str, pd.DataFrame], dt: pd.Timestamp) -> pd.DataFrame:
    return pd.DataFrame({name: frame[dt] for name, frame in mats.items()})


def equal_weights(holdings: list[str]) -> dict[str, float]:
    if not holdings:
        return {}
    weight = 1.0 / len(holdings)
    return {str(code): weight for code in holdings}


def portfolio_return_and_drift(
    weights: dict[str, float],
    day_returns: pd.Series,
) -> tuple[float, dict[str, float]]:
    if not weights:
        return 0.0, {}
    codes = list(weights)
    returns = pd.to_numeric(day_returns.reindex(codes), errors="coerce").fillna(0.0).astype("float64")
    old_weights = pd.Series(weights, dtype="float64").reindex(codes).fillna(0.0)
    gross_return = float((old_weights * returns).sum())
    closing_values = old_weights * (1.0 + returns)
    total = float(closing_values.sum())
    if not np.isfinite(total) or total <= 0:
        return gross_return, dict(weights)
    drifted = closing_values / total
    return gross_return, {str(code): float(weight) for code, weight in drifted.items()}


def max_drawdown(nav: pd.Series) -> float:
    clean = pd.to_numeric(nav, errors="coerce").dropna()
    if clean.empty:
        return np.nan
    return float((clean / clean.cummax() - 1.0).min())


def series_metrics(nav: pd.Series) -> dict[str, float]:
    clean = pd.to_numeric(nav, errors="coerce").dropna()
    if len(clean) < 2:
        return {"区间收益率": np.nan, "年化收益率": np.nan, "最大回撤": np.nan, "年化波动率": np.nan}
    daily = clean.pct_change().dropna()
    elapsed_days = max((clean.index[-1] - clean.index[0]).days, 1)
    return {
        "区间收益率": float(clean.iloc[-1] / clean.iloc[0] - 1.0),
        "年化收益率": float((clean.iloc[-1] / clean.iloc[0]) ** (365.25 / elapsed_days) - 1.0),
        "最大回撤": max_drawdown(clean),
        "年化波动率": float(daily.std(ddof=1) * np.sqrt(252)) if len(daily) > 1 else np.nan,
    }


def font(size: int) -> FontProperties:
    for path in [
        Path(r"C:\Windows\Fonts\NotoSansSC-VF.ttf"),
        Path(r"C:\Windows\Fonts\msyh.ttc"),
        Path(r"C:\Windows\Fonts\simhei.ttf"),
    ]:
        if path.exists():
            return FontProperties(fname=str(path), size=size)
    return FontProperties(size=size)


def run_backtest(
    config: StrategyConfig,
    *,
    base: Optional[ModuleType] = None,
    data: Optional[dict[str, pd.DataFrame]] = None,
) -> dict[str, pd.DataFrame]:
    base = base or load_backtest_module()
    data = data or base.load_original_data(config.parquet_root, force_refresh=False)
    required = [
        "收盘价",
        "涨跌幅",
        "转股溢价率",
        "余额",
        "剩余期限",
        "正股市值",
        "正股收盘价",
        "EXPMA5",
        "EXPMA10",
        "EXPMA20",
        "主体评级",
    ]
    if config.excluded_parity_range is not None:
        required.append("平价")
    missing = [name for name in required if name not in data]
    if missing:
        raise KeyError(f"数据源缺少必要 sheet: {missing}")

    mats = {name: base._normalize_wide(data[name]) for name in required}
    common_codes: Optional[pd.Index] = None
    common_dates: Optional[pd.DatetimeIndex] = None
    for frame in mats.values():
        common_codes = frame.index if common_codes is None else common_codes.intersection(frame.index)
        frame_dates = base._date_columns(frame)
        common_dates = frame_dates if common_dates is None else common_dates.intersection(frame_dates)
    assert common_codes is not None and common_dates is not None
    dates = common_dates[common_dates >= pd.Timestamp(config.backtest_start_date)]
    if len(dates) < 2:
        raise ValueError("回测区间公共交易日不足。")

    mats = {name: frame.reindex(index=common_codes, columns=dates) for name, frame in mats.items()}
    numeric = {
        name: frame.apply(pd.to_numeric, errors="coerce").astype("float64")
        for name, frame in mats.items()
        if name != "主体评级"
    }
    ratings = mats["主体评级"].astype("string")

    status_df = data.get("交易状态")
    if status_df is not None:
        status_df = base._normalize_wide(status_df).reindex(index=common_codes, columns=dates)
    total_dates = base._clean_total_dates(data.get("总表"), common_codes)

    benchmark_navs: dict[str, pd.Series] = {}
    benchmark_df = data.get("指数")
    if benchmark_df is not None and not benchmark_df.empty:
        benchmark_df = base._normalize_wide(benchmark_df)
        for name in config.benchmark_names:
            if name not in benchmark_df.index:
                print(f"[benchmark] 未找到 {name}，跳过")
                continue
            levels = pd.to_numeric(benchmark_df.loc[name].reindex(dates), errors="coerce")
            valid = levels.dropna()
            if len(valid) >= 2 and float(valid.iloc[0]) != 0:
                benchmark_navs[name] = levels / float(valid.iloc[0])

    rebalance_positions = set(range(1, len(dates), config.rebalance_every_n_days))
    holdings: list[str] = []
    weights: dict[str, float] = {}
    nav = 1.0
    nav_rows: list[dict[str, object]] = []
    rebalance_rows: list[dict[str, object]] = []
    holdings_rows: list[dict[str, object]] = []

    for pos, dt in enumerate(dates):
        if pos in rebalance_positions:
            pricing_dt = dates[pos - 1]
            cross = cross_section(numeric, pricing_dt)
            rating = ratings[pricing_dt].reindex(cross.index)
            close = pd.to_numeric(cross["收盘价"], errors="coerce")
            premium = pd.to_numeric(cross["转股溢价率"], errors="coerce")
            stock_close = pd.to_numeric(cross["正股收盘价"], errors="coerce")
            expma5 = pd.to_numeric(cross["EXPMA5"], errors="coerce")
            expma10 = pd.to_numeric(cross["EXPMA10"], errors="coerce")
            expma20 = pd.to_numeric(cross["EXPMA20"], errors="coerce")
            double_low = close + premium

            active = close.ge(config.min_price)
            active &= pd.to_numeric(cross["余额"], errors="coerce").ge(config.min_balance)
            active &= pd.to_numeric(cross["剩余期限"], errors="coerce").ge(config.min_remaining_years)
            active &= pd.to_numeric(cross["正股市值"], errors="coerce").ge(config.min_stock_market_value)
            active &= double_low.gt(config.min_double_low)
            active &= stock_close.gt(expma5) & expma5.gt(expma10) & expma10.gt(expma20)
            active &= rating.isin(config.allowed_subject_ratings)
            if config.excluded_parity_range is not None:
                lower, upper = config.excluded_parity_range
                parity = pd.to_numeric(cross["平价"], errors="coerce")
                active &= ~parity.between(lower, upper, inclusive="both")

            listing = total_dates["上市日期"]
            last_trade = total_dates["最后交易日"]
            redeem_announce = total_dates["赎回公告日"]
            if listing.notna().any():
                active &= listing.isna() | listing.le(pricing_dt)
            if last_trade.notna().any():
                active &= last_trade.isna() | last_trade.ge(pricing_dt)
            if redeem_announce.notna().any():
                active &= redeem_announce.isna() | redeem_announce.gt(pricing_dt)
            if status_df is not None and pricing_dt in status_df.columns:
                active &= status_df[pricing_dt].astype("string").isin(["交易", "新股上市"])

            candidates = double_low[active].dropna()
            selected = candidates.nsmallest(config.max_holdings)
            holdings = [str(code) for code in selected.index]
            weights = equal_weights(holdings)
            for rank, code in enumerate(holdings, start=1):
                rebalance_rows.append(
                    {
                        "调仓日": dt,
                        "定价依据日": pricing_dt,
                        "排名": rank,
                        "转债代码": code,
                        "双低指标": float(double_low.loc[code]),
                        "收盘价": float(close.loc[code]),
                        "转股溢价率": float(premium.loc[code]),
                        "余额": float(pd.to_numeric(cross.loc[code, "余额"], errors="coerce")),
                        "剩余期限": float(pd.to_numeric(cross.loc[code, "剩余期限"], errors="coerce")),
                        "正股市值": float(pd.to_numeric(cross.loc[code, "正股市值"], errors="coerce")),
                        "平价": float(pd.to_numeric(cross.loc[code, "平价"], errors="coerce"))
                        if "平价" in cross.columns
                        else np.nan,
                        "主体评级": str(rating.loc[code]),
                        "正股收盘价": float(stock_close.loc[code]),
                        "EXPMA5": float(expma5.loc[code]),
                        "EXPMA10": float(expma10.loc[code]),
                        "EXPMA20": float(expma20.loc[code]),
                        "候选数": int(len(candidates)),
                    }
                )

        day_returns = numeric["涨跌幅"].loc[holdings, dt] / 100.0 if holdings else pd.Series(dtype="float64")
        daily_return, weights = portfolio_return_and_drift(weights, day_returns)
        nav *= 1.0 + daily_return
        nav_rows.append({"日期": dt, "组合日收益率": daily_return, "净值": nav, "持仓数量": len(holdings)})
        holdings_rows.append({"日期": dt, **{f"持仓{i}": code for i, code in enumerate(holdings, start=1)}})

    nav_df = pd.DataFrame(nav_rows).set_index("日期")
    for name, bench in benchmark_navs.items():
        nav_df[f"基准净值_{name}"] = bench.reindex(nav_df.index)
    rebalance_df = pd.DataFrame(rebalance_rows)
    holdings_df = pd.DataFrame(holdings_rows)

    metric_rows = []
    metric_cols = ["净值", *[f"基准净值_{name}" for name in config.benchmark_names if f"基准净值_{name}" in nav_df.columns]]
    for col in metric_cols:
        metric = series_metrics(nav_df[col])
        name = config.strategy_name if col == "净值" else col.replace("基准净值_", "")
        metric_rows.append({"组合": name, **metric, "末日净值": float(nav_df[col].dropna().iloc[-1])})
    metrics_df = pd.DataFrame(metric_rows)

    rule_df = pd.DataFrame(
        [
            ["策略名称", config.strategy_name],
            ["调仓频率", f"每{config.rebalance_every_n_days}个交易日"],
            ["定价依据", "调仓日前一交易日截面"],
            ["持仓数量", f"双低指标最低前{config.max_holdings}只"],
            ["双低指标", "收盘价 + 转股溢价率，要求 >100，无上限"],
            ["动量条件", "EXPMA5>EXPMA10>EXPMA20 且 正股收盘价>EXPMA5"],
            [
                "基础筛选",
                f"余额>={config.min_balance:g}、收盘价>={config.min_price:g}、剩余期限>={config.min_remaining_years:g}、正股市值>={config.min_stock_market_value:g}",
            ],
            [
                "平价剔除",
                "无" if config.excluded_parity_range is None else f"闭区间[{config.excluded_parity_range[0]:g}, {config.excluded_parity_range[1]:g}]",
            ],
            ["评级限制", "主体评级 AAA/AA+/AA/AA-/A+，债项评级不限制"],
            ["事件筛选", "已公告赎回后不买入；交易状态需为交易/新股上市"],
            ["数据源", config.parquet_root],
            ["样本起点", str(nav_df.index.min().date())],
            ["样本终点", str(nav_df.index.max().date())],
        ],
        columns=["项目", "内容"],
    )

    return {
        "策略规则": rule_df,
        "净值曲线": nav_df,
        "绩效摘要": metrics_df,
        "调仓组合": rebalance_df,
        "日度持仓": holdings_df,
    }


def save_outputs(results: dict[str, pd.DataFrame], config: StrategyConfig) -> tuple[Path, Path]:
    nav_df = results["净值曲线"]
    end_date = pd.Timestamp(nav_df.index.max()).strftime("%Y%m%d")
    out_dir = ROOT / "runs" / "research" / f"策略回测{end_date}"
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = out_dir / f"{config.strategy_name}.xlsx"
    png_path = out_dir / f"{config.strategy_name}.png"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for sheet, frame in results.items():
            frame.to_excel(writer, sheet_name=sheet[:31], index=frame.index.name is not None)

    plt.rcParams["axes.unicode_minus"] = False
    f = font(11)
    tf = font(14)
    fig, ax = plt.subplots(figsize=(13, 7), dpi=180)
    ax.plot(nav_df.index, nav_df["净值"], label=config.strategy_name, linewidth=2.2, color="#9467bd")
    for col, color in [("基准净值_转债指数", "#2ca02c"), ("基准净值_万得全A", "#d62728")]:
        if col in nav_df.columns:
            ax.plot(nav_df.index, nav_df[col], label=col.replace("基准净值_", ""), linewidth=1.8, linestyle="--", color=color)
    ax.set_title(f"{config.strategy_name}净值曲线", fontproperties=tf)
    ax.set_xlabel("日期", fontproperties=f)
    ax.set_ylabel("累计净值", fontproperties=f)
    ax.grid(True, alpha=0.25)
    ax.legend(prop=f, loc="upper left")
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(f)
    fig.tight_layout()
    fig.savefig(png_path)
    plt.close(fig)
    return xlsx_path, png_path


def nine_variant_configs() -> list[StrategyConfig]:
    """返回用户指定的九个独立配置；其中两组 0.5 年门槛配置按要求保留。"""
    base = StrategyConfig()
    return [
        StrategyConfig(strategy_name="中金动量双低_容量30", max_holdings=30),
        StrategyConfig(strategy_name="中金动量双低_容量40", max_holdings=40),
        StrategyConfig(strategy_name="中金动量双低_容量50", max_holdings=50),
        StrategyConfig(strategy_name="中金动量双低_容量30_剔除剩余期限0.5年内", max_holdings=30, min_remaining_years=0.5),
        StrategyConfig(strategy_name="中金动量双低_容量30_剔除剩余期限1年内", max_holdings=30, min_remaining_years=1.0),
        StrategyConfig(strategy_name="中金动量双低_容量30_剔除剩余期限1.5年内", max_holdings=30, min_remaining_years=1.5),
        StrategyConfig(strategy_name="中金动量双低_容量30_剔除剩余期限2年内", max_holdings=30, min_remaining_years=2.0),
        StrategyConfig(strategy_name="中金动量双低_容量30_剔除平价120至130", max_holdings=30, excluded_parity_range=(120.0, 130.0)),
        StrategyConfig(strategy_name="中金动量双低_容量30_剔除平价130至140", max_holdings=30, excluded_parity_range=(130.0, 140.0)),
    ]


def save_variant_summary(variant_results: list[tuple[StrategyConfig, dict[str, pd.DataFrame]]]) -> tuple[Path, Path]:
    nav_df = variant_results[0][1]["净值曲线"]
    end_date = pd.Timestamp(nav_df.index.max()).strftime("%Y%m%d")
    out_dir = ROOT / "runs" / "research" / f"策略回测{end_date}"
    summary_rows = []
    combined_nav = pd.DataFrame(index=nav_df.index)
    for config, results in variant_results:
        metric = results["绩效摘要"].query("组合 == @config.strategy_name").iloc[0].to_dict()
        summary_rows.append(metric)
        combined_nav[config.strategy_name] = results["净值曲线"]["净值"]
    summary_path = out_dir / "中金动量双低_9组参数敏感性汇总.xlsx"
    with pd.ExcelWriter(summary_path, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="绩效摘要", index=False)
        combined_nav.to_excel(writer, sheet_name="净值对比")
        pd.DataFrame(
            [
                ["说明", "本文件为同一策略的参数敏感性结果，属于探索性检验，不构成样本外验证或未来表现承诺。"],
                ["平价区间", "120–130、130–140 均按闭区间剔除。"],
                ["重复配置", "容量30 与 剩余期限剔除0.5年内 的参数完全一致；按用户要求分别输出。"],
            ],
            columns=["项目", "内容"],
        ).to_excel(writer, sheet_name="口径说明", index=False)

    plt.rcParams["axes.unicode_minus"] = False
    fig, ax = plt.subplots(figsize=(14, 7.5), dpi=180)
    for config, _ in variant_results:
        ax.plot(combined_nav.index, combined_nav[config.strategy_name], label=config.strategy_name, linewidth=1.35)
    ax.set_title("中金动量双低：9组参数敏感性净值对比", fontproperties=font(14))
    ax.set_xlabel("日期", fontproperties=font(11))
    ax.set_ylabel("累计净值", fontproperties=font(11))
    ax.grid(True, alpha=0.25)
    ax.legend(prop=font(8), ncol=2, loc="upper left")
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(font(10))
    fig.tight_layout()
    chart_path = out_dir / "中金动量双低_9组参数敏感性净值对比.png"
    fig.savefig(chart_path)
    plt.close(fig)
    return summary_path, chart_path


def main() -> None:
    parser = argparse.ArgumentParser(description="中金动量双低策略回测")
    parser.add_argument("--single", action="store_true", help="仅运行原始容量50配置")
    args = parser.parse_args()

    configs = [StrategyConfig()] if args.single else nine_variant_configs()
    base = load_backtest_module()
    data = base.load_original_data(configs[0].parquet_root, force_refresh=False)
    variant_results: list[tuple[StrategyConfig, dict[str, pd.DataFrame]]] = []
    for config in configs:
        results = run_backtest(config, base=base, data=data)
        xlsx_path, png_path = save_outputs(results, config)
        nav_df = results["净值曲线"]
        latest = nav_df.iloc[-1]
        print(f"策略名称: {config.strategy_name}")
        print(f"样本区间: {nav_df.index.min().date()} - {nav_df.index.max().date()}")
        print(f"末日净值: {float(latest['净值']):.6f}")
        print(f"末日持仓数: {int(latest['持仓数量'])}")
        print(f"结果文件: {xlsx_path}")
        print(f"净值图: {png_path}")
        variant_results.append((config, results))

    if not args.single:
        summary_path, chart_path = save_variant_summary(variant_results)
        print(f"9组汇总: {summary_path}")
        print(f"9组对比图: {chart_path}")


if __name__ == "__main__":
    main()
