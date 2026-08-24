# -*- coding: utf-8 -*-
"""Use parquet history and the daily-update calculation functions to build CSV tables.

This script deliberately produces calculation intermediates only.  The final XLSX is
authored and formatted by the artifact-tool builder.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
import sys
import time
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
PARQUET_ROOT = ROOT / "data/转债个券历史序列"
DAILY_SCRIPT = ROOT / "【更新】日报数据更新.py"
DRAFT_SCRIPT = ROOT / "底稿更新.py"

REQUIRED_METRICS = (
    "余额",
    "收盘价",
    "债项评级",
    "平价",
    "转股溢价率",
    "纯债溢价率",
    "纯债价值",
    "隐含波动率",
    "YTM",
    "平价底价溢价率",
    "换手率",
    "剩余期限",
    "正股市值",
)


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def fix_mojibake(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    try:
        repaired = value.encode("latin1").decode("gbk")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return value
    # Only accept the repair when it actually introduces CJK characters.
    if any("\u4e00" <= ch <= "\u9fff" for ch in repaired):
        return repaired
    return value


def repair_total(total: pd.DataFrame) -> pd.DataFrame:
    out = total.copy()
    out.columns = [fix_mojibake(c) for c in out.columns]
    for col in out.select_dtypes(include=["object", "string"]).columns:
        out[col] = out[col].map(fix_mojibake)
    out.index = out.index.astype(str)
    return out


def normalize_date_columns(original: dict[str, pd.DataFrame]) -> list[pd.Timestamp]:
    common: set[pd.Timestamp] | None = None
    for name in REQUIRED_METRICS:
        if name not in original:
            raise KeyError(f"Parquet 缺少日报计算所需 sheet: {name}")
        cols = {pd.Timestamp(c).normalize() for c in original[name].columns}
        common = cols if common is None else common & cols
    if not common:
        raise RuntimeError("所需指标没有共同交易日")
    return sorted(common)


def latest_rating(rating: pd.DataFrame, dates: list[pd.Timestamp]) -> pd.Series:
    cols = [c for c in dates if c in rating.columns]
    if not cols:
        return pd.Series(index=rating.index, dtype=object)
    return rating[cols].ffill(axis=1).iloc[:, -1]


def derive_listing_close(
    close: pd.DataFrame,
    listing_dates: pd.Series,
    dates: list[pd.Timestamp],
) -> pd.Series:
    values = pd.Series(np.nan, index=close.index, dtype=float)
    positions = {d: i for i, d in enumerate(dates)}
    arr = close[dates].apply(pd.to_numeric, errors="coerce").to_numpy(
        dtype=float,
        na_value=np.nan,
    )
    for row_idx, code in enumerate(close.index):
        listed = pd.to_datetime(listing_dates.get(code), errors="coerce")
        if pd.isna(listed):
            continue
        listed = pd.Timestamp(listed).normalize()
        start = positions.get(listed)
        if start is None:
            start = int(np.searchsorted(np.array(dates, dtype="datetime64[ns]"), np.datetime64(listed)))
        if start >= len(dates):
            continue
        row = arr[row_idx, start:]
        valid = np.flatnonzero(~np.isnan(row))
        if valid.size:
            values.at[code] = float(row[valid[0]])
    return values


def build_static_inputs(
    original: dict[str, pd.DataFrame],
    dates: list[pd.Timestamp],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.Series]:
    total_src = repair_total(original["总表"])
    codes = original["余额"].index.astype(str)
    total_src = total_src.reindex(codes)

    name_col = "转债名称"
    listing_col = "上市日期"
    last_trade_col = "最后交易日"
    conv_start_col = "转股期起始日"
    industry_col = "申万行业"
    issue_col = "发行规模"
    required_total = [name_col, listing_col, last_trade_col, conv_start_col, industry_col, issue_col]
    missing = [c for c in required_total if c not in total_src.columns]
    if missing:
        raise KeyError(f"总表缺少字段: {missing}; 实际字段={list(total_src.columns)}")

    rating_last = latest_rating(original["债项评级"].reindex(codes), dates)
    meta = pd.DataFrame(
        {
            "代码": codes,
            "名称": total_src[name_col].values,
            "上市日期": pd.to_datetime(total_src[listing_col], errors="coerce").dt.strftime("%Y-%m-%d").values,
            "评级": rating_last.reindex(codes).values,
        }
    )
    listing_close = derive_listing_close(
        original["收盘价"].reindex(codes),
        pd.to_datetime(total_src[listing_col], errors="coerce"),
        dates,
    )
    total = pd.DataFrame(
        {
            "代码": codes,
            "名称": total_src[name_col].values,
            "上市日期": pd.to_datetime(total_src[listing_col], errors="coerce").dt.strftime("%Y-%m-%d").values,
            "最后交易日": pd.to_datetime(total_src[last_trade_col], errors="coerce").dt.strftime("%Y-%m-%d").values,
            "转股期起始日": pd.to_datetime(total_src[conv_start_col], errors="coerce").dt.strftime("%Y-%m-%d").values,
            "所属申万行业(2021）1级": total_src[industry_col].values,
            "摘牌日": pd.to_datetime(total_src[last_trade_col], errors="coerce").dt.strftime("%Y-%m-%d").values,
            "上市首日价格": listing_close.reindex(codes).values,
        }
    )
    issue = meta.copy()
    issue["发行规模"] = pd.to_numeric(total_src[issue_col], errors="coerce").values
    conv_start = pd.to_datetime(total_src[conv_start_col], errors="coerce")
    conv_start.index = codes
    return meta, total, issue, conv_start


def chunk_data(
    original: dict[str, pd.DataFrame],
    dates: list[pd.Timestamp],
    meta: pd.DataFrame,
    total: pd.DataFrame,
    issue: pd.DataFrame,
    conv_start: pd.Series,
) -> dict[str, pd.DataFrame]:
    codes = meta["代码"].astype(str)
    date_names = [d.strftime("%Y-%m-%d") for d in dates]
    out: dict[str, pd.DataFrame] = {}
    for name in REQUIRED_METRICS:
        src = original[name].reindex(codes)
        block = src.reindex(columns=dates).copy()
        block.columns = date_names
        out[name] = pd.concat([meta.reset_index(drop=True), block.reset_index(drop=True)], axis=1)
    out["总表"] = total
    out["发行规模"] = issue

    close_m = out["收盘价"].set_index("代码")[date_names].apply(pd.to_numeric, errors="coerce")
    prem_m = out["转股溢价率"].set_index("代码")[date_names].apply(pd.to_numeric, errors="coerce")
    first_close = total.set_index("代码")["上市首日价格"].pipe(pd.to_numeric, errors="coerce")
    is_subnew = conv_start > dates[-1]
    eligible = is_subnew & first_close.notna() & (first_close != 0)
    ipo_vals = close_m.div(first_close, axis=0).sub(1).mul(100)
    ipo_vals = ipo_vals.where(eligible.reindex(close_m.index), np.nan)
    sub_prem = prem_m.where(is_subnew.reindex(prem_m.index), np.nan)
    for sheet_name, values in (
        ("次新券相对上市首日涨跌幅", ipo_vals),
        ("次新券转股溢价率", sub_prem),
    ):
        out[sheet_name] = pd.concat(
            [meta.reset_index(drop=True), values.reindex(codes).reset_index(drop=True)],
            axis=1,
        )
    return out


def append_date_tables(store: dict[str, list[pd.DataFrame]], tables: dict[str, pd.DataFrame]) -> None:
    for name, frame in tables.items():
        copy = frame.copy()
        copy.index = pd.to_datetime(copy.index, errors="coerce")
        store.setdefault(name, []).append(copy)


def append_fit_tables(store: dict[str, list[pd.DataFrame]], daily, fit_data) -> None:
    specs = (
        ("百元平价拟合溢价率", dict(with_greeks=True)),
        ("平衡型", dict(floor_mode="balance")),
        ("偏债型", dict(floor_mode="bond")),
        ("偏股型", dict(floor_mode="stock")),
        ("百元平价拟合溢价率（1-5.5年）", dict(remain_range=(1, 5.5))),
        ("百元平价拟合溢价率（0-2年）", dict(remain_range=(0, 2))),
        ("百元平价拟合溢价率（2-4年）", dict(remain_range=(2, 4))),
        ("百元平价拟合溢价率（4-6年）", dict(remain_range=(4, 6))),
    )
    for name, kwargs in specs:
        store.setdefault(name, []).append(daily._run_fit_table(fit_data, **kwargs))


def combine_industry(parts: dict[str, list[pd.DataFrame]]) -> dict[str, pd.DataFrame]:
    out: dict[str, pd.DataFrame] = {}
    for name, frames in parts.items():
        if name == "剔除转债":
            max_rows = max((len(f) for f in frames), default=0)
            normalized = [f.reset_index(drop=True).reindex(range(max_rows)) for f in frames]
            out[name] = pd.concat(normalized, axis=1)
        else:
            out[name] = pd.concat(frames, axis=1)
    return out


def clean_for_csv(frame: pd.DataFrame, include_index: bool, index_label: str | None) -> pd.DataFrame:
    out = frame.copy()
    if include_index:
        out.index.name = index_label or out.index.name
        out = out.reset_index()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%Y-%m-%d")
    return out


def safe_sheet_name(prefix: str, seq: int, name: str, used: set[str]) -> str:
    stem = f"{prefix}{seq:02d}_{name}"
    candidate = stem[:31]
    suffix = 1
    while candidate in used:
        tail = f"_{suffix}"
        candidate = (stem[: 31 - len(tail)] + tail)
        suffix += 1
    used.add(candidate)
    return candidate


def table_profile(frame: pd.DataFrame) -> dict[str, Any]:
    numeric = frame.select_dtypes(include=[np.number])
    return {
        "rows": int(frame.shape[0]),
        "cols": int(frame.shape[1]),
        "numeric_non_null": int(numeric.notna().sum().sum()) if not numeric.empty else 0,
        "numeric_null": int(numeric.isna().sum().sum()) if not numeric.empty else 0,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--chunk-size", type=int, default=63)
    args = parser.parse_args()
    output_dir = Path(args.output_dir).resolve()
    csv_dir = output_dir / "csv"
    csv_dir.mkdir(parents=True, exist_ok=True)

    t0 = time.time()
    draft = load_module("parquet_draft_source", DRAFT_SCRIPT)
    daily = load_module("daily_update_calculation", DAILY_SCRIPT)
    parquet_loader_path = str(PARQUET_ROOT.relative_to(ROOT))
    source = draft.load_original_data(
        source_type="parquet",
        parquet_root=parquet_loader_path,
        force_refresh=False,
    )
    dates = normalize_date_columns(source)
    meta, total, issue, conv_start = build_static_inputs(source, dates)

    stats_a_parts: dict[str, list[pd.DataFrame]] = {}
    stats_b_parts: dict[str, list[pd.DataFrame]] = {}
    fit_parts: dict[str, list[pd.DataFrame]] = {}
    js_parts: dict[str, list[pd.DataFrame]] = {}
    industry_parts: dict[str, list[pd.DataFrame]] = {}

    chunks = math.ceil(len(dates) / args.chunk_size)
    for chunk_no, start in enumerate(range(0, len(dates), args.chunk_size), 1):
        chunk_dates = dates[start : start + args.chunk_size]
        print(
            f"[compute] {chunk_no}/{chunks} "
            f"{chunk_dates[0].date()}..{chunk_dates[-1].date()}",
            flush=True,
        )
        data = chunk_data(source, chunk_dates, meta, total, issue, conv_start)
        stats_a = daily.compute_statistics(data, apply_track_b=False)
        stats_b = daily.compute_statistics(data, apply_track_b=True)
        append_date_tables(stats_a_parts, stats_a)
        append_date_tables(stats_b_parts, stats_b)

        fit_data = daily.prepare_fit_data(data)
        append_fit_tables(fit_parts, daily, fit_data)

        js = daily.run_js_update(data, stats_b, str(output_dir), chunk_dates[-1].strftime("%m%d"))
        append_date_tables(js_parts, js)

        industry = daily.compute_industry_and_exclusion(data)
        for name, frame in industry.items():
            industry_parts.setdefault(name, []).append(frame.copy())

    stats_a_all = {k: pd.concat(v).sort_index() for k, v in stats_a_parts.items()}
    stats_b_all = {k: pd.concat(v).sort_index() for k, v in stats_b_parts.items()}
    fit_all = {k: pd.concat(v, ignore_index=True) for k, v in fit_parts.items()}
    js_all = {k: pd.concat(v).sort_index() for k, v in js_parts.items()}
    industry_all = combine_industry(industry_parts)

    source_inventory_rows = []
    for name, frame in source.items():
        date_like = [pd.Timestamp(c) for c in frame.columns if isinstance(c, (pd.Timestamp, np.datetime64))]
        source_inventory_rows.append(
            {
                "源数据表": name,
                "行数": len(frame),
                "列数": len(frame.columns),
                "起始日期": min(date_like).strftime("%Y-%m-%d") if date_like else "",
                "截止日期": max(date_like).strftime("%Y-%m-%d") if date_like else "",
                "最新日非空数": int(frame[date_like[-1]].notna().sum()) if date_like else "",
            }
        )
    source_inventory = pd.DataFrame(source_inventory_rows)

    missing_latest_rows = []
    for name in REQUIRED_METRICS:
        frame = source[name]
        latest = dates[-1]
        missing_latest_rows.append(
            {
                "指标": name,
                "最新交易日": latest.strftime("%Y-%m-%d"),
                "总券数": len(frame),
                "非空数": int(frame[latest].notna().sum()),
                "缺失数": int(frame[latest].isna().sum()),
                "缺失率": float(frame[latest].isna().mean()),
            }
        )
    missing_latest = pd.DataFrame(missing_latest_rows)

    fit_failure_rows = []
    for name, frame in fit_all.items():
        failed = int((pd.to_numeric(frame["a"], errors="coerce").fillna(0) == 0).sum())
        fit_failure_rows.append(
            {
                "拟合表": name,
                "交易日数": len(frame),
                "拟合失败/无样本日": failed,
                "成功率": (len(frame) - failed) / len(frame) if len(frame) else np.nan,
            }
        )
    fit_checks = pd.DataFrame(fit_failure_rows)

    used_names: set[str] = set()
    manifest: list[dict[str, Any]] = []

    def write_group(
        group: str,
        prefix: str,
        tables: dict[str, pd.DataFrame],
        *,
        include_index: bool,
        index_label: str | None,
    ) -> None:
        for seq, (logical_name, frame) in enumerate(tables.items(), 1):
            excel_name = safe_sheet_name(prefix, seq, logical_name, used_names)
            out = clean_for_csv(frame, include_index, index_label)
            csv_name = f"{len(manifest)+1:03d}_{excel_name}.csv"
            out.to_csv(csv_dir / csv_name, index=False, encoding="utf-8")
            manifest.append(
                {
                    "group": group,
                    "logical_name": logical_name,
                    "sheet_name": excel_name,
                    "csv": f"csv/{csv_name}",
                    "profile": table_profile(out),
                }
            )

    write_group("A轨统计", "A", stats_a_all, include_index=True, index_label="日期")
    write_group("B轨剔妖统计", "B", stats_b_all, include_index=True, index_label="日期")
    write_group("百元拟合", "F", fit_all, include_index=False, index_label=None)
    write_group("JS", "J", js_all, include_index=True, index_label="日期")
    write_group("行业", "I", industry_all, include_index=True, index_label="行业")
    write_group(
        "源数据质量",
        "Q",
        {
            "源数据清单": source_inventory,
            "最新日缺失": missing_latest,
            "拟合核对": fit_checks,
        },
        include_index=False,
        index_label=None,
    )

    metadata = {
        "title": "Parquet全历史日报计算底稿",
        "parquet_root": str(PARQUET_ROOT),
        "source_fingerprint": draft._parquet_dir_fingerprint(parquet_loader_path),
        "source_cache": "matched parquet fingerprint disk cache",
        "daily_script": str(DAILY_SCRIPT),
        "source_start_date": dates[0].strftime("%Y-%m-%d"),
        "source_end_date": dates[-1].strftime("%Y-%m-%d"),
        "trade_date_count": len(dates),
        "bond_count": len(meta),
        "calculation_groups": {
            "A轨统计": len(stats_a_all),
            "B轨剔妖统计": len(stats_b_all),
            "百元拟合": len(fit_all),
            "JS": len(js_all),
            "行业": len(industry_all),
        },
        "rules": [
            "A轨：上市日前、最后交易日后置空，并按收盘价缺失同步全表。",
            "B轨：在A轨基础上剔除转股溢价率>50且收盘价>150的样本。",
            "拟合：平价70–130、换手率≤50，并按转股溢价率3%/97%分位去极值。",
            "次新券上市首日价格由Parquet收盘价中上市日起首个有效值还原。",
            "数值为Python按日报脚本函数全历史回算结果，不含外部Excel链接。",
        ],
        "elapsed_seconds": round(time.time() - t0, 1),
        "manifest": manifest,
    }
    (output_dir / "manifest.json").write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(
        f"[done] {len(manifest)} tables, {len(dates)} dates, "
        f"elapsed={time.time()-t0:.1f}s, manifest={output_dir / 'manifest.json'}",
        flush=True,
    )


if __name__ == "__main__":
    main()
