"""转债底稿更新主脚本。

功能概览：
1) 登录 iFinD，获取交易日历
2) 读取底稿并补齐宽表日期列
3) 按配置规则更新各 sheet（含总表字段、公式拉取、派生计算）
4) 保存更新后的标准 Parquet
"""

# ================== 依赖导入区 ==================
# --- 标准库 ---
import ctypes
import hashlib
import os
import pickle
import time as pytime
from configparser import ConfigParser
from datetime import date, datetime
from numbers import Number
from pathlib import Path

# --- 第三方库 ---
import numpy as np
import pandas as pd
import xlwings as xw
from iFinDPy import *
from tqdm import tqdm

import sys

_COMMON_MODULE_DIR = Path(__file__).resolve().parents[1] / "common"
if str(_COMMON_MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(_COMMON_MODULE_DIR))

from 转债Parquet标准读写模块 import (
    MONTHLY_METRICS,
    export_original_data_to_parquet as _export_standard_parquet,
    read_original_data_from_parquet as _read_standard_parquet,
)


# ========== 缓存机制 ==========
_PARQUET_MEM_CACHE = {}
TQDM_NCOLS = 92
EXCEL_FORMULA_BATCH_SIZE = 50_000
# 日常更新默认关闭；需要集中修复全部历史局部缺失时手动改为 True。
RUN_HISTORICAL_MISSING_REPAIR = False

# iFinD 的 cb_anal_lasttradingday 对下列异常退市券返回 0，通用公式会继而
# 错误回退到到期日。总表公式更新后必须用已核实的实际最后交易日覆盖。
MANUAL_LAST_TRADE_DATE_OVERRIDES = {
    "110072.SH": "2024-08-28",  # 广汇转债
    "128100.SZ": "2023-05-22",  # 搜特转债
    "128085.SZ": "2024-01-18",  # 鸿达转债
    "123015.SZ": "2023-07-28",  # 蓝盾转退
    "128012.SZ": "2021-04-19",  # 辉丰转债
}


def _parquet_dir_fingerprint(input_root):
    """
    生成 parquet 目录唯一指纹：
    - 遍历 input_root 下所有 .parquet 文件，收集其绝对路径、最后修改时间、文件大小。
    - 拼接所有文件信息后做 md5 哈希，作为该目录内容的唯一标识。
    - 用于缓存机制，目录内容变化时自动失效。
    参数：
        input_root (str): parquet 根目录路径。
    返回：
        str: 该目录内容的 md5 指纹。
    """
    file_info = []
    for root, _, files in os.walk(input_root):
        for f in sorted(files):
            if not f.endswith('.parquet'):
                continue
            fp = os.path.join(root, f)
            st = os.stat(fp)
            # 记录文件绝对路径、修改时间戳、文件大小
            file_info.append(f"{fp}|{st.st_mtime_ns}|{st.st_size}")
    raw = '|'.join(file_info)
    return hashlib.md5(raw.encode('utf-8')).hexdigest()

def load_parquet_with_cache(input_root="data/转债个券历史序列", force_refresh=False, cache_dir="tmp/cache/parquet"):
    """
    读取 parquet，尽量命中缓存以避免全量重读。
    策略：
      1. 计算 input_root 目录指纹作为 key。
      2. 命中内存缓存 → 直接返回。
      3. 命中磁盘 .pkl → 回填内存并返回。
      4. 都未命中 → 全量读取 parquet 返回（不在此处写缓存）。
         缓存的生成统一由 build_parquet_cache 在 parquet 导出后完成，
         避免读取路径顺手落盘导致时机混乱。
    """
    fp = _parquet_dir_fingerprint(input_root)
    mem_key = (input_root, fp)
    cache_file = os.path.join(cache_dir, f"parquet_{fp}.pkl")
    if not force_refresh:
        cached = _PARQUET_MEM_CACHE.get(mem_key)
        if cached is not None:
            print(f"[cache] parquet memory hit: {input_root}")
            return cached
        if os.path.exists(cache_file):
            print(f"[cache] 命中 Parquet 磁盘缓存，正在载入：{cache_file}", flush=True)
            with open(cache_file, "rb") as f:
                cached = pickle.load(f)
            _PARQUET_MEM_CACHE[mem_key] = cached
            print(f"[cache] Parquet 磁盘缓存载入完成：{cache_file}", flush=True)
            return cached
    print(f"[cache] parquet 缓存未命中，直接读取 parquet 目录：{input_root}")
    data = read_original_data_from_parquet(input_root)
    _PARQUET_MEM_CACHE[mem_key] = data
    return data


def build_parquet_cache(input_root="data/转债个券历史序列", cache_dir="tmp/cache/parquet", data=None):
    """
    基于当前 parquet 目录内容生成/刷新磁盘缓存（.cache_parquet/parquet_<fp>.pkl）。
    一般在 export_original_data_to_parquet 完成后调用，确保下次进程启动时
    可直接命中 parquet 缓存，不再临时现读。

    参数：
        input_root: parquet 根目录，用于计算指纹。
        cache_dir: 磁盘缓存目录。
        data: 可选。若传入，则直接 pickle 该对象，跳过从 parquet 再读一遍；
              典型用法是把 export 之前的内存数据直接复用，避免回读带来的
              dtype/列类型漂移与额外 IO。未传入时则退回"从目录重读"的行为。

    每次导出 parquet 后 mtime/size 都会变化，导致新的 fp，因此本函数
    写入后会清理同目录下其它 parquet_*.pkl，仅保留本次生成的一份。
    返回：
        str: 生成的缓存文件路径。
    """
    os.makedirs(cache_dir, exist_ok=True)
    fp = _parquet_dir_fingerprint(input_root)
    cache_file = os.path.join(cache_dir, f"parquet_{fp}.pkl")
    if data is None:
        data = read_original_data_from_parquet(input_root)
        source_tag = "reread"
    else:
        source_tag = "inmem"
    with open(cache_file, "wb") as f:
        pickle.dump(data, f, protocol=pickle.HIGHEST_PROTOCOL)
    _PARQUET_MEM_CACHE[(input_root, fp)] = data
    print(f"[cache] parquet rebuilt ({source_tag}): {cache_file}")

    # 清理同目录下其它 parquet_*.pkl，避免历史 pkl 堆积
    keep_name = os.path.basename(cache_file)
    removed = 0
    for name in os.listdir(cache_dir):
        if not (name.startswith("parquet_") and name.endswith(".pkl")):
            continue
        if name == keep_name:
            continue
        try:
            os.remove(os.path.join(cache_dir, name))
            removed += 1
        except OSError as e:
            print(f"[cache] 清理旧 parquet pkl 失败 {name}: {e}")
    if removed:
        print(f"[cache] 已清理 {removed} 个旧 parquet pkl，仅保留: {keep_name}")
    return cache_file


def apply_manual_last_trade_date_overrides(original_data, total_sheet="总表"):
    """用已核实日期覆盖异常退市券的最后交易日。"""
    total_df = original_data.get(total_sheet)
    if total_df is None or total_df.empty:
        raise ValueError(f"最后交易日人工覆盖失败：缺少或为空的{total_sheet}。")
    if "最后交易日" not in total_df.columns:
        raise ValueError(f"最后交易日人工覆盖失败：{total_sheet}缺少最后交易日字段。")

    index_as_text = total_df.index.astype(str)
    updated_total = total_df.copy()
    applied = 0
    missing_codes = []
    for code, date_text in MANUAL_LAST_TRADE_DATE_OVERRIDES.items():
        matches = index_as_text == code
        match_count = int(matches.sum())
        if match_count == 0:
            missing_codes.append(code)
            continue
        if match_count > 1:
            raise ValueError(f"最后交易日人工覆盖失败：{total_sheet}中代码 {code} 重复。")
        updated_total.loc[matches, "最后交易日"] = pd.Timestamp(date_text)
        applied += 1

    updated = dict(original_data)
    updated[total_sheet] = updated_total
    print(
        f"[总表更新] 最后交易日人工覆盖完成｜成功 {applied}/"
        f"{len(MANUAL_LAST_TRADE_DATE_OVERRIDES)}｜缺失代码 {len(missing_codes)}",
        flush=True,
    )
    if missing_codes:
        print(f"[总表更新] 最后交易日人工覆盖未找到：{missing_codes}", flush=True)
    return updated


# ========== 零值清理 ==========
# 所有 sheet 统一执行零值清理，不再设置保留零值的业务白名单。
# 文本结果 sheet 仍用于控制 Excel 回填的数据类型，但其中精确表示零的字符串也会清理。
_TEXT_RESULT_SHEETS = {"交易状态", "正股交易状态", "主体评级", "债项评级"}

# 下列状态型结果在条款尚未生效、尚无公告或尚无可计算标的时，整日全空属于正常现象。
# 它们不参与历史全空列自动补抓，也不触发全空数据告警。
STRUCTURAL_EMPTY_SHEETS = {
    "赎回累计天数",
    "下修累计天数",
}

# 新券补历史默认从发行日后开始；下列正股相关指标需要拉取历史全区间。
FULL_HISTORY_NEW_BOND_SHEETS = {
    "正股收盘价",
    "正股交易状态",
    "正股近1日均价",
    "正股近20日均价",
    "正股20日波动率",
    "正股市值",
    "每股净资产",
    "EXPMA5",
    "EXPMA10",
    "EXPMA20",
}

# 下列指标不受转债上市日、最后交易日约束，应保留转债生命周期以外的历史数据。
BOND_LIFECYCLE_EXEMPT_SHEETS = {
    "转股价",  # 完整保留人工补充序列，不做上市日或最后交易日清洗。
    "正股收盘价",
    "正股交易状态",
    "正股近1日均价",
    "正股近20日均价",
    "正股20日波动率",
    "正股市值",
    "每股净资产",
    "EXPMA5",
    "EXPMA10",
    "EXPMA20",
}

# 两项条款累计天数由各自的本地化脚本独立计算并回写，
# 底稿更新不对它们取数、重算或执行生命周期清洗。
UNCHANGED_BACKTEST_SHEETS = {
    "赎回累计天数",
    "下修累计天数",
}


def get_existing_max_data_date(original_data, skip_sheets=None):
    """返回扩列前宽表中实际已有数据的最大日期。

    只把至少存在一个非空值的日期列视为“已有数据列”，避免历史遗留的全空占位列
    抬高截止日。该日期用于分隔“新券历史补档”和“全市场新增日期更新”。
    """
    if skip_sheets is None:
        skip_sheets = {"总表"}

    max_dates = []
    for sheet_name, df in original_data.items():
        if df is None or df.empty or sheet_name in skip_sheets:
            continue

        parsed_cols = pd.to_datetime(pd.Index(df.columns), errors="coerce")
        dated_cols = [
            (col, pd.Timestamp(ts))
            for col, ts in zip(df.columns, parsed_cols)
            if pd.notna(ts)
        ]
        for col, ts in reversed(sorted(dated_cols, key=lambda x: x[1])):
            if df[col].notna().any():
                max_dates.append(ts)
                break

    if not max_dates:
        raise ValueError("扩列前数据中没有可识别且包含非空值的日期列，无法确定历史截止日。")
    return max(max_dates)


def build_formula_fetch_tasks(
    df,
    *,
    fetch_scope="auto",
    history_cutoff=None,
    new_bond_codes=None,
    new_bond_issue_dates=None,
    fetch_full_history=False,
    include_historical_missing=False,
    incremental_by_row_latest=False,
    eligible_start_dates=None,
    eligible_end_dates=None,
    require_eligible_start=False,
):
    """生成公式取数任务，元素为 ``(代码字符串, 原索引, 列名, 日期)``。

    fetch_scope:
    - ``new_bond_history``：仅新券、仅历史截止日及以前、仅空单元格；
    - ``historical_missing``：仅历史截止日及以前的局部缺失；
      ``incremental_by_row_latest=True`` 时只补齐每行首个与末个有效值之间的内部缺口；
    - ``incremental_dates``：仅历史截止日及之后、全部新老券的空单元格；
    - ``auto``：兼容旧调用，合并整列为空的新日期与新券历史缺口。
    """
    valid_scopes = {"auto", "new_bond_history", "historical_missing", "incremental_dates"}
    if fetch_scope not in valid_scopes:
        raise ValueError(f"fetch_scope 无效: {fetch_scope}，仅支持 {sorted(valid_scopes)}")

    cutoff_ts = pd.Timestamp(history_cutoff) if history_cutoff is not None else None
    if fetch_scope != "auto" and cutoff_ts is None:
        raise ValueError(f"fetch_scope={fetch_scope} 时必须提供 history_cutoff")

    new_code_set = {
        str(c) for c in (new_bond_codes or [])
        if str(c) not in ("", "nan", "None")
    }
    issue_map = {}
    for code, raw in (new_bond_issue_dates or {}).items():
        ts = pd.to_datetime(raw, errors="coerce")
        if pd.notna(ts):
            issue_map[str(code)] = pd.Timestamp(ts)

    def normalize_date_map(raw_map):
        normalized = {}
        for code, raw in (raw_map or {}).items():
            ts = pd.to_datetime(raw, errors="coerce")
            if pd.notna(ts):
                normalized[str(code)] = pd.Timestamp(ts).normalize()
        return normalized

    eligible_start_map = normalize_date_map(eligible_start_dates)
    eligible_end_map = normalize_date_map(eligible_end_dates)

    index_pairs = [(str(idx), idx) for idx in df.index]
    new_index_pairs = [(code, idx) for code, idx in index_pairs if code in new_code_set]
    parsed_cols = pd.to_datetime(pd.Index(df.columns), errors="coerce")
    dated_cols = [
        (col, pd.Timestamp(ts))
        for col, ts in zip(df.columns, parsed_cols)
        if pd.notna(ts)
    ]

    row_earliest_dates = {}
    row_latest_dates = {}
    if fetch_scope == "historical_missing" and incremental_by_row_latest:
        for code, idx in index_pairs:
            populated_dates = [
                ts for col, ts in dated_cols if pd.notna(df.at[idx, col])
            ]
            row_earliest_dates[(code, idx)] = (
                min(populated_dates) if populated_dates else None
            )
            row_latest_dates[(code, idx)] = (
                max(populated_dates) if populated_dates else None
            )

    tasks = []
    seen = set()

    def add_missing(code_pairs, col, ts, apply_issue_cutoff):
        for code, idx in code_pairs:
            eligible_start = eligible_start_map.get(code)
            if require_eligible_start and eligible_start is None:
                continue
            if eligible_start is not None and ts < eligible_start:
                continue
            eligible_end = eligible_end_map.get(code)
            if eligible_end is not None and ts > eligible_end:
                continue
            if apply_issue_cutoff and not fetch_full_history:
                issue_ts = issue_map.get(code)
                if issue_ts is not None and ts < issue_ts:
                    continue
            if pd.isna(df.at[idx, col]):
                key = (code, col)
                if key not in seen:
                    tasks.append((code, idx, col, ts))
                    seen.add(key)

    for col, ts in sorted(dated_cols, key=lambda x: x[1]):
        if fetch_scope == "new_bond_history":
            if ts <= cutoff_ts:
                add_missing(new_index_pairs, col, ts, apply_issue_cutoff=True)
            continue

        if fetch_scope == "historical_missing":
            if ts > cutoff_ts:
                continue
            if incremental_by_row_latest:
                historical_pairs = [
                    (code, idx)
                    for code, idx in index_pairs
                    if (
                        row_latest_dates[(code, idx)] is None
                        or row_earliest_dates[(code, idx)] <= ts <= row_latest_dates[(code, idx)]
                    )
                ]
                add_missing(historical_pairs, col, ts, apply_issue_cutoff=False)
            elif include_historical_missing:
                add_missing(index_pairs, col, ts, apply_issue_cutoff=False)
            continue

        if fetch_scope == "incremental_dates":
            if ts >= cutoff_ts:
                add_missing(index_pairs, col, ts, apply_issue_cutoff=False)
            continue

        # auto：保留旧接口语义，但最终仍统一成一个批量任务集合。
        if df[col].isna().all():
            add_missing(index_pairs, col, ts, apply_issue_cutoff=False)
        if new_index_pairs:
            add_missing(new_index_pairs, col, ts, apply_issue_cutoff=True)

    return tasks


def _value_is_numeric_zero(v):
    """判定数值零及其字符串形式；布尔值、空值和非数值文本不视为零。"""
    if v is None or isinstance(v, bool):
        return False
    if isinstance(v, Number):
        try:
            fv = float(v)
        except (TypeError, ValueError):
            return False
        if not np.isfinite(fv):
            return False
        return fv == 0.0
    if isinstance(v, str):
        text = v.strip().replace(",", "")
        if not text:
            return False
        if text.endswith("%"):
            text = text[:-1].strip()
        try:
            fv = float(text)
        except (TypeError, ValueError):
            return False
        return np.isfinite(fv) and fv == 0.0
    return False


def sanitize_zero_values(
    original_data,
    verbose=True,
    progress=None,
    row_codes=None,
    date_after=None,
    date_on_or_after=None,
    date_on_or_before=None,
    include_non_date_columns=True,
    scope_label=None,
):
    """把指定作用域内的零值统一替换为缺失值。

    - 数值型 0/0.0 以及字符串 "0"/"0.0"/"0%" 均清理
    - 布尔 False、日期值、非零数字字符串及普通文本保持原样
    - row_codes 可把范围限制到本阶段新增代码
    - date_after/date_on_or_after/date_on_or_before 可把范围限制到本阶段日期列
    - include_non_date_columns 控制有日期边界时是否同时处理静态字段
    - 原 DataFrame 不会被就地修改；仅在检测到 0 时复制一份再替换
    """
    if progress is None:
        progress = bool(verbose)

    row_code_set = (
        {str(code) for code in row_codes}
        if row_codes is not None
        else None
    )
    date_after_ts = pd.to_datetime(date_after, errors="coerce")
    date_on_or_after_ts = pd.to_datetime(date_on_or_after, errors="coerce")
    date_before_ts = pd.to_datetime(date_on_or_before, errors="coerce")
    if pd.notna(date_after_ts) and pd.notna(date_on_or_after_ts):
        raise ValueError("date_after 与 date_on_or_after 不能同时指定")
    has_date_filter = (
        pd.notna(date_after_ts)
        or pd.notna(date_on_or_after_ts)
        or pd.notna(date_before_ts)
    )
    display_scope = scope_label or ("局部" if row_code_set is not None or has_date_filter else "全表")

    sanitized = {}
    total_cleared = 0
    touched_sheets = 0
    sheet_items = list(original_data.items())
    if progress:
        print(
            f"[sanitize] 开始{display_scope}零值清理，共 {len(sheet_items)} 个 sheet。",
            flush=True,
        )
        iterator = tqdm(
            sheet_items,
            desc=f"{display_scope}零值清理",
            unit="sheet",
            total=len(sheet_items),
            ncols=TQDM_NCOLS,
            dynamic_ncols=False,
            mininterval=0.2,
        )
    else:
        iterator = sheet_items

    for sheet_name, df in iterator:
        if progress:
            iterator.set_postfix_str(f"{sheet_name}: 扫描中", refresh=True)
        if df is None or df.empty:
            sanitized[sheet_name] = df
            if progress:
                iterator.set_postfix_str(f"{sheet_name}: 空表/跳过", refresh=False)
            continue

        if row_code_set is None:
            target_index = df.index
        else:
            row_mask = df.index.astype(str).isin(row_code_set)
            target_index = df.index[row_mask]

        if has_date_filter:
            parsed_cols = pd.to_datetime(
                pd.Index(df.columns),
                errors="coerce",
                format="mixed",
            )
            col_mask = np.asarray(pd.notna(parsed_cols), dtype=bool)
            if pd.notna(date_after_ts):
                col_mask &= np.asarray(parsed_cols > pd.Timestamp(date_after_ts), dtype=bool)
            if pd.notna(date_on_or_after_ts):
                col_mask &= np.asarray(
                    parsed_cols >= pd.Timestamp(date_on_or_after_ts), dtype=bool
                )
            if pd.notna(date_before_ts):
                col_mask &= np.asarray(parsed_cols <= pd.Timestamp(date_before_ts), dtype=bool)
            if include_non_date_columns:
                col_mask |= np.asarray(pd.isna(parsed_cols), dtype=bool)
            target_cols = df.columns[col_mask]
        else:
            target_cols = df.columns

        if len(target_index) == 0 or len(target_cols) == 0:
            sanitized[sheet_name] = df
            if progress:
                iterator.set_postfix_str(f"{sheet_name}: 无本阶段目标", refresh=False)
            continue

        target_df = df.loc[target_index, target_cols]
        new_df = None
        sheet_cleared = 0
        numeric_cols = [
            col
            for col in target_cols
            if (
                pd.api.types.is_numeric_dtype(target_df[col].dtype)
                and not pd.api.types.is_bool_dtype(target_df[col].dtype)
            )
        ]
        if numeric_cols:
            numeric_part = target_df[numeric_cols]
            numeric_zero_mask = numeric_part.notna() & numeric_part.eq(0)
            numeric_cleared = int(numeric_zero_mask.to_numpy().sum())
            if numeric_cleared:
                new_df = df.copy()
                new_df.loc[target_index, numeric_cols] = numeric_part.mask(numeric_zero_mask)
                sheet_cleared += numeric_cleared

        numeric_col_set = set(numeric_cols)
        other_cols = [col for col in target_cols if col not in numeric_col_set]
        for col in other_cols:
            series = target_df[col]
            if pd.api.types.is_bool_dtype(series.dtype) or pd.api.types.is_datetime64_any_dtype(series.dtype):
                continue
            zero_mask = series.map(_value_is_numeric_zero)
            cleared = int(zero_mask.sum())
            if cleared == 0:
                continue
            if new_df is None:
                new_df = df.copy()
            new_df.loc[target_index, col] = series.mask(zero_mask)
            sheet_cleared += cleared

        if sheet_cleared == 0:
            sanitized[sheet_name] = df
            if progress:
                iterator.set_postfix_str(f"{sheet_name}: 0 个", refresh=False)
            continue

        sanitized[sheet_name] = new_df
        total_cleared += sheet_cleared
        touched_sheets += 1
        if progress:
            iterator.set_postfix_str(
                f"{sheet_name}: 清理 {sheet_cleared} 个",
                refresh=False,
            )

    if verbose:
        print(
            f"[sanitize] {display_scope}零值清理完成：扫描 {len(sheet_items)} 个 sheet，"
            f"涉及 {touched_sheets} 个 sheet，共移除 {total_cleared} 个零值。",
            flush=True,
        )
    return sanitized


def sanitize_bond_lifecycle_values(
    original_data,
    *,
    total_sheet="总表",
    verbose=True,
):
    """按转债生命周期清理个券日度指标。

    - 除生命周期豁免指标及独立维护的条款累计天数外，
      上市日期之前、最后交易日之后统一置空；
    - 仅处理标准 Parquet Schema 已登记的个券日度指标，不处理总表和指数。
    """
    total_df = original_data.get(total_sheet)
    if total_df is None or total_df.empty:
        raise ValueError(f"生命周期清理失败：缺少或为空的{total_sheet}。")

    required_columns = {"上市日期", "最后交易日"}
    missing_columns = sorted(required_columns - set(total_df.columns))
    if missing_columns:
        raise ValueError(f"生命周期清理失败：{total_sheet}缺少字段 {missing_columns}")

    total = total_df.copy()
    total.index = total.index.astype(str)
    updated = dict(original_data)
    total_cleared = 0
    touched_sheets = 0

    for sheet_name in MONTHLY_METRICS:
        if sheet_name in BOND_LIFECYCLE_EXEMPT_SHEETS or sheet_name in UNCHANGED_BACKTEST_SHEETS:
            continue
        df = original_data.get(sheet_name)
        if df is None or df.empty:
            continue

        parsed_columns = pd.to_datetime(pd.Index(df.columns), errors="coerce", format="mixed")
        date_mask = np.asarray(pd.notna(parsed_columns), dtype=bool)
        if not date_mask.any():
            continue

        date_columns = df.columns[date_mask]
        dates = pd.DatetimeIndex(parsed_columns[date_mask]).normalize().to_numpy(dtype="datetime64[ns]")
        index_as_text = pd.Index(df.index.astype(str))
        common_code_mask = index_as_text.isin(total.index)
        if not common_code_mask.any():
            continue

        target_index = df.index[common_code_mask]
        target_codes = index_as_text[common_code_mask]
        metadata = total.reindex(target_codes)
        listing_dates = pd.to_datetime(metadata["上市日期"], errors="coerce").to_numpy(dtype="datetime64[ns]")
        last_trade_dates = pd.to_datetime(metadata["最后交易日"], errors="coerce").to_numpy(dtype="datetime64[ns]")

        lifecycle_mask = np.zeros((len(target_index), len(date_columns)), dtype=bool)
        valid_listing = ~pd.isna(listing_dates)
        valid_last_trade = ~pd.isna(last_trade_dates)
        lifecycle_mask |= valid_listing[:, None] & (dates[None, :] < listing_dates[:, None])
        lifecycle_mask |= valid_last_trade[:, None] & (dates[None, :] > last_trade_dates[:, None])

        target_values = df.loc[target_index, date_columns]
        clear_mask = lifecycle_mask & target_values.notna().to_numpy()
        cleared = int(clear_mask.sum())
        if cleared == 0:
            continue

        cleaned_values = target_values.mask(
            pd.DataFrame(lifecycle_mask, index=target_index, columns=date_columns)
        )
        cleaned_df = df.copy()
        cleaned_df.loc[target_index, date_columns] = cleaned_values
        updated[sheet_name] = cleaned_df
        total_cleared += cleared
        touched_sheets += 1
        if verbose:
            print(f"[lifecycle] {sheet_name}: 清理 {cleared} 个生命周期外非空值。")

    if verbose:
        print(
            f"[lifecycle] 个券日度指标清理完成：涉及 {touched_sheets} 个指标，"
            f"共清理 {total_cleared} 个非空值；生命周期豁免指标 {len(BOND_LIFECYCLE_EXEMPT_SHEETS)} 项"
            f"及独立维护的条款累计天数不处理。"
        )
    return updated


def find_all_empty_date_items(
    original_data,
    date_on_or_after=None,
    skip_sheets=None,
):
    """查找指定日期起，每个日期型 sheet 中整列全空的“项目 × 日期”组合。"""
    start_ts = pd.to_datetime(date_on_or_after, errors="coerce")
    skip_set = set(STRUCTURAL_EMPTY_SHEETS if skip_sheets is None else skip_sheets)
    anomalies = []
    checked_items = 0
    checked_dates = set()

    for sheet_name, df in original_data.items():
        if df is None or df.empty or sheet_name in skip_set:
            continue

        parsed_cols = pd.to_datetime(
            pd.Index(df.columns),
            errors="coerce",
            format="mixed",
        )
        dated_cols = [
            (col, pd.Timestamp(dt))
            for col, dt in zip(df.columns, parsed_cols)
            if pd.notna(dt) and (pd.isna(start_ts) or pd.Timestamp(dt) >= pd.Timestamp(start_ts))
        ]
        if not dated_cols:
            continue

        checked_items += 1
        target_cols = [col for col, _ in dated_cols]
        target_df = df[target_cols]
        all_empty = target_df.isna().all(axis=0)

        # object/string 列中的空白字符串也视作空值。
        for col in target_cols:
            if bool(all_empty.loc[col]):
                continue
            series = target_df[col]
            if pd.api.types.is_object_dtype(series.dtype) or pd.api.types.is_string_dtype(series.dtype):
                effective_empty = series.isna() | series.astype("string").str.strip().eq("")
                all_empty.loc[col] = bool(effective_empty.all())

        date_by_col = {col: dt for col, dt in dated_cols}
        checked_dates.update(date_by_col.values())
        for col in target_cols:
            if bool(all_empty.loc[col]):
                anomalies.append(
                    {
                        "sheet_name": str(sheet_name),
                        "date": date_by_col[col],
                    }
                )

    anomalies.sort(key=lambda x: (x["date"], x["sheet_name"]))
    return anomalies, checked_items, pd.DatetimeIndex(sorted(checked_dates))


def notify_all_empty_date_items(
    original_data,
    *,
    date_on_or_after=None,
    skip_sheets=None,
    title="底稿更新后全空数据提示",
):
    """校验并汇总弹窗提示整列全空的日期型数据项。"""
    anomalies, checked_items, checked_dates = find_all_empty_date_items(
        original_data,
        date_on_or_after=date_on_or_after,
        skip_sheets=skip_sheets,
    )

    if len(checked_dates) == 0:
        print("[validation] 没有符合本次校验范围的日期列，跳过全空数据检查。")
        return anomalies

    date_range = (
        checked_dates[0].strftime("%Y-%m-%d")
        if len(checked_dates) == 1
        else f"{checked_dates[0].strftime('%Y-%m-%d')}~{checked_dates[-1].strftime('%Y-%m-%d')}"
    )
    if not anomalies:
        print(
            f"[validation] 全空数据校验通过：{checked_items} 个项目，"
            f"日期范围 {date_range}。"
        )
        return anomalies

    grouped = {}
    for item in anomalies:
        date_text = item["date"].strftime("%Y-%m-%d")
        grouped.setdefault(date_text, []).append(item["sheet_name"])

    detail_lines = [
        f"{date_text}：{'、'.join(sheet_names)}"
        for date_text, sheet_names in grouped.items()
    ]
    console_message = (
        f"[validation] 警告：发现 {len(anomalies)} 处“某日某项数据全空”，"
        f"涉及 {len(grouped)} 个日期：\n" + "\n".join(detail_lines)
    )
    print(console_message, flush=True)

    # Windows MessageBox 文本过长时截断显示，完整明细仍保留在控制台。
    popup_lines = detail_lines[:40]
    omitted = len(detail_lines) - len(popup_lines)
    popup_message = (
        f"发现 {len(anomalies)} 处“某日某项数据全空”，"
        f"涉及 {len(grouped)} 个日期。\n\n"
        + "\n".join(popup_lines)
    )
    if omitted > 0:
        popup_message += f"\n\n另有 {omitted} 个日期，请查看控制台完整明细。"

    try:
        # 使用阻塞式 MessageBox，仅在弹出瞬间切到前台，不保持始终置顶。
        # MessageBoxW 不设置超时，程序会等待用户点击“确定”后再继续落盘。
        message_box_flags = (
            0x00000000  # MB_OK
            | 0x00000030  # MB_ICONWARNING
            | 0x00010000  # MB_SETFOREGROUND
        )
        ctypes.windll.user32.MessageBoxW(
            0,
            popup_message,
            title,
            message_box_flags,
        )
    except Exception as exc:
        print(f"[validation] 弹窗显示失败，已保留控制台提示：{exc}")

    return anomalies


def load_original_data(parquet_root="data/转债个券历史序列", force_refresh=False):
    """仅从标准 Parquet 底稿读取数据；读取失败时直接终止，禁止回退旧 Excel。"""
    if not os.path.isdir(parquet_root):
        raise FileNotFoundError(f"未找到 Parquet 数据目录: {parquet_root}")
    print(f"[source] 使用 Parquet 数据源: {parquet_root}")
    return load_parquet_with_cache(parquet_root, force_refresh=force_refresh)


IFIND_CREDENTIAL_FILE = Path(__file__).resolve().parents[2] / "private/ifind账号.txt"


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


def print_ifind_usage() -> None:
    """显示iFinD各数据项的已用额度比例。"""
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


def thslogindemo():
    """
    登录 iFinD 并打印登录状态码与中文提示。
    """
    username, password = load_ifind_credentials()
    thsLogin = THS_iFinDLogin(username, password)
    print(thsLogin)
    if thsLogin not in (0, -201):
        print('登录失败')
    else:
        print('登录成功')
        print_ifind_usage()


def _is_ifind_fetching_value(value):
    """统一识别 iFinD 公式仍在抓取的中英文状态文本。"""
    if not isinstance(value, str):
        return False
    text = value.strip()
    return (
        text in {"抓取中", "抓取中...", "抓取中…", "抓取中……"}
        or "fetching" in text.lower()
    )


def _is_update_empty_value(value):
    """统一识别公式返回的空值。"""
    if value is None or value is pd.NA:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _is_update_error_value(value):
    """统一识别 Excel/iFinD 返回的错误状态文本。"""
    if not isinstance(value, str):
        return False
    text = value.strip().lower()
    return (
        text.startswith("#")
        or "error" in text
        or "错误" in text
        or "失败" in text
        or "invalid" in text
    )


UPDATE_STAGE_LABELS = {
    "new_bond_history": "新券历史补档",
    "historical_missing": "历史缺失修补",
    "incremental_dates": "新增数据",
}

HISTORICAL_BACKFILL_SCOPES = frozenset({"new_bond_history", "historical_missing"})


def _record_historical_backfill(
    audit,
    *,
    stage,
    item_name,
    written_count,
    written_dates=None,
):
    """记录历史补档阶段实际写入的有效值和受影响月份。"""
    if audit is None or stage not in HISTORICAL_BACKFILL_SCOPES:
        return
    count = int(written_count or 0)
    if count <= 0:
        return

    stage_entry = audit.setdefault(
        stage,
        {"written_cells": 0, "items": {}, "months": set()},
    )
    stage_entry["written_cells"] += count
    stage_entry["items"][str(item_name)] = (
        int(stage_entry["items"].get(str(item_name), 0)) + count
    )
    for value in written_dates or ():
        ts = pd.to_datetime(value, errors="coerce")
        if pd.notna(ts):
            stage_entry["months"].add(pd.Timestamp(ts).strftime("%Y%m"))


def notify_historical_backfill(audit):
    """历史补档实际发生时，输出终端警告并弹窗提醒网站同步范围。"""
    active = {
        stage: entry
        for stage, entry in (audit or {}).items()
        if int(entry.get("written_cells", 0)) > 0
    }
    if not active:
        return False

    all_months = sorted({
        month
        for entry in active.values()
        for month in entry.get("months", set())
    })
    total_cells = sum(int(entry.get("written_cells", 0)) for entry in active.values())
    all_items = {
        item
        for entry in active.values()
        for item in entry.get("items", {})
    }

    if not all_months:
        month_summary = "月份范围未识别"
    elif len(all_months) == 1:
        month_summary = all_months[0]
    else:
        month_summary = f"{all_months[0]}~{all_months[-1]}"

    detail_lines = []
    for stage in ("new_bond_history", "historical_missing"):
        entry = active.get(stage)
        if not entry:
            continue
        stage_months = sorted(entry.get("months", set()))
        if not stage_months:
            stage_month_text = "月份未识别"
        elif len(stage_months) == 1:
            stage_month_text = stage_months[0]
        else:
            stage_month_text = f"{stage_months[0]}~{stage_months[-1]}"
        detail_lines.append(
            f"{_update_stage_label(stage)}：写入 {int(entry['written_cells']):,} 个单元格，"
            f"涉及 {len(entry.get('items', {}))} 个指标、{len(stage_months)} 个月份"
            f"（{stage_month_text}）"
        )

    message_lines = [
        "检测到本次运行发生历史补档。",
        "",
        *detail_lines,
        "",
        f"合计：{total_cells:,} 个单元格、{len(all_items)} 个指标、"
        f"{len(all_months)} 个月份（{month_summary}）。",
        "网站同步时不能只上传当前月；请同步受影响历史月份，或上传全量个券 Parquet。",
    ]
    message = "\n".join(message_lines)
    print("\n[历史补档提醒] " + "\n".join(message_lines), flush=True)
    try:
        # 阻塞等待点击“确定”；仅在弹出瞬间切到前台，不保持始终置顶。
        message_box_flags = 0x00000030 | 0x00010000  # MB_ICONWARNING | MB_SETFOREGROUND
        ctypes.windll.user32.MessageBoxW(
            0,
            message,
            "历史补档提醒",
            message_box_flags,
        )
    except Exception as exc:
        print(f"[历史补档提醒] 弹窗显示失败：{exc}", flush=True)
    return True


def _update_stage_label(stage=None, item_name=None):
    """返回用户界面使用的更新阶段名，隐藏内部函数/实现标签。"""
    if item_name and str(item_name).startswith("总表"):
        return "总表更新"
    return UPDATE_STAGE_LABELS.get(stage, "新增数据")


def _print_update_result(
    item_name,
    values,
    *,
    result_date=None,
    target_indices=None,
    stage=None,
):
    """每项更新后打印一次结果，重点展示本阶段最新一日（或静态字段）的抓取情况。"""
    result_prefix = f"[{_update_stage_label(stage, item_name)}][result]"
    if isinstance(values, pd.DataFrame):
        parsed_cols = pd.to_datetime(pd.Index(values.columns), errors="coerce")
        date_map = {
            pd.Timestamp(dt): col
            for col, dt in zip(values.columns, parsed_cols)
            if pd.notna(dt)
        }
        if not date_map:
            print(f"{result_prefix} {item_name}｜无日期列，无法展示最新一日结果")
            return
        target_date = pd.to_datetime(result_date, errors="coerce")
        if pd.isna(target_date):
            target_date = max(date_map)
        target_date = pd.Timestamp(target_date)
        if target_date not in date_map:
            print(f"{result_prefix} {item_name}｜{target_date.strftime('%Y-%m-%d')} 无对应日期列")
            return
        result_series = values[date_map[target_date]]
        date_label = target_date.strftime("%Y-%m-%d")
    else:
        result_series = pd.Series(values)
        date_label = "静态字段"

    if target_indices is not None:
        wanted = pd.Index(list(dict.fromkeys(target_indices)))
        selected = [idx for idx in wanted if idx in result_series.index]
        result_series = result_series.reindex(selected)

    empty_mask = result_series.map(_is_update_empty_value)
    fetching_mask = result_series.map(_is_ifind_fetching_value) & ~empty_mask
    error_mask = result_series.map(_is_update_error_value) & ~empty_mask & ~fetching_mask
    success_mask = ~empty_mask & ~fetching_mask & ~error_mask
    print(
        f"{result_prefix} {item_name}｜{date_label}"
        f"｜目标 {len(result_series)}｜成功 {int(success_mask.sum())}"
        f"｜空值 {int(empty_mask.sum())}｜抓取中 {int(fetching_mask.sum())}"
        f"｜错误 {int(error_mask.sum())}｜"
    )


def fill_sheets_with_ifind(
    original_data,
    sheet_configs,
    wait_step_seconds=1.2,
    max_wait_seconds=10,
    trade_dates=None,
    new_bond_codes=None,
    new_bond_issue_dates=None,
    fetch_scope="auto",
    history_cutoff=None,
    skip_total_table=False,
    formula_batch_size=EXCEL_FORMULA_BATCH_SIZE,
    historical_backfill_audit=None,
):
    """
    使用统一流程更新多个宽表 sheet：
    - new_bond_history：先为新券批量补齐扩列前历史截止日及以前的数据
    - historical_missing：补齐历史截止日及以前、有效区间内的局部缺失
    - incremental_dates：扩列后，批量拉取历史截止日及之后的缺失值，覆盖新老全部券
    - 每个指标把本阶段“代码 × 日期”任务按固定批次铺进 Excel，不再逐日调用
    - new_bond_issue_dates: {code: pd.Timestamp}，用于裁剪抓取起始日（发行日）；
      若参数未提供，函数会在总表填好"发行日期"列后自动从总表刷新
    - historical_backfill_audit: 可选审计字典；仅记录历史补档阶段实际写入的
      非空、非零有效值及受影响月份，供落盘后提示网站同步范围
    - 在隐藏 Excel 中批量写入 @thsiFinD 公式
    - 以插件是否仍返回“抓取中”状态作为主要等待依据，再回填最终值

    sheet_configs 示例:
    [
        {"sheet_name": "收盘价", "formula_expr": '=@thsiFinD("ths_bond_close_cbond",A{r},B{r},101)'},
        {"sheet_name": "平价底价溢价率", "formula_expr": '=@thsiFinD("ths_conversion_parity_price_premium_cbond",A{r},B{r})'},
        {
            "sheet_name": "剩余期限",
            "rule_type": "derived_term_from_total",
            "total_start_col": "发行日期",
            "total_end_col": "到期日期",
        },
    ]
    """
    if not sheet_configs:
        return original_data
    valid_fetch_scopes = {"auto", "new_bond_history", "historical_missing", "incremental_dates"}
    if fetch_scope not in valid_fetch_scopes:
        raise ValueError(f"fetch_scope 无效: {fetch_scope}，仅支持 {sorted(valid_fetch_scopes)}")
    formula_batch_size = int(formula_batch_size)
    if formula_batch_size <= 0:
        raise ValueError("formula_batch_size 必须为正整数")
    # Excel 工作表共 1,048,576 行，第 1 行留作标题。
    formula_batch_size = min(formula_batch_size, 1_048_575)
    history_cutoff_ts = pd.Timestamp(history_cutoff) if history_cutoff is not None else None
    if fetch_scope != "auto" and history_cutoff_ts is None:
        raise ValueError(f"fetch_scope={fetch_scope} 时必须提供 history_cutoff")

    only_total_call = all(
        cfg.get("rule_type", "formula") == "total_table_formula"
        for cfg in sheet_configs
    )
    call_stage_label = "总表更新" if only_total_call else _update_stage_label(fetch_scope)
    call_stage_prefix = f"[{call_stage_label}]"

    updated = dict(original_data)
    # 新增券：在已有日期列触发补档（仅针对这些 code 补 NaN 单元，避免误抓老券缺失）
    new_bond_codes_set = {str(c) for c in (new_bond_codes or []) if str(c) not in ("", "nan")}
    # 新增券发行日期映射：用于把"2015-今天"的日期列裁剪到"发行日-今天"
    # 来源优先级：参数 new_bond_issue_dates（板块成分） < 总表"发行日期"列刷新（cb_list_dtonl 公式后）
    merged_issue_ts = {}
    if new_bond_issue_dates:
        for k, v in new_bond_issue_dates.items():
            try:
                ts = pd.Timestamp(v)
                if pd.notna(ts):
                    merged_issue_ts[str(k)] = ts
            except Exception:
                continue
    listing_logged_ref = {"done": False}  # 一次性日志开关

    def _refresh_issue_from_total():
        """总表处理完后，从总表"发行日期"列把缺失的新券发行日补进 merged_issue_ts。"""
        if not new_bond_codes_set:
            return
        total_df = updated.get("总表")
        if total_df is None or "发行日期" not in total_df.columns:
            return
        added = 0
        for code in new_bond_codes_set:
            if code in merged_issue_ts:
                continue
            if code not in total_df.index:
                continue
            try:
                ts = pd.to_datetime(total_df.at[code, "发行日期"], errors="coerce")
            except Exception:
                ts = pd.NaT
            if pd.notna(ts):
                merged_issue_ts[code] = pd.Timestamp(ts)
                added += 1
        if not listing_logged_ref["done"]:
            hit = sum(1 for c in new_bond_codes_set if c in merged_issue_ts)
            print(
                f"{call_stage_prefix} 新券发行日期可用 {hit}/{len(new_bond_codes_set)} 只"
                f"（总表刷新新增 {added} 只）"
            )
            listing_logged_ref["done"] = True

    def _build_dt_to_col(df):
        parsed_cols = pd.to_datetime(pd.Index(df.columns), errors="coerce")
        dt_map = {}
        for col, dt in zip(df.columns, parsed_cols):
            if pd.notna(dt):
                dt_map[pd.Timestamp(dt)] = col
        return dt_map

    def _collect_target_dates(target_df):
        """按当前阶段挑选待计算日期，并标记其作用范围。"""
        dt_map = _build_dt_to_col(target_df)
        index_str = target_df.index.astype(str)
        new_mask = index_str.isin(new_bond_codes_set)
        selected = []
        reasons = {}
        for dt, col in sorted(dt_map.items()):
            col_series = target_df[col]
            if fetch_scope == "new_bond_history":
                if dt <= history_cutoff_ts and new_mask.any() and col_series[new_mask].isna().any():
                    selected.append(dt)
                    reasons[dt] = "new_bond"
                continue
            if fetch_scope == "historical_missing":
                if (
                    dt <= history_cutoff_ts
                    and sheet_name not in STRUCTURAL_EMPTY_SHEETS
                    and col_series.isna().any()
                ):
                    selected.append(dt)
                    reasons[dt] = "historical_missing"
                continue
            if fetch_scope == "incremental_dates":
                if dt >= history_cutoff_ts and col_series.isna().any():
                    selected.append(dt)
                    reasons[dt] = "new_date"
                continue

            if col_series.isna().all():
                selected.append(dt)
                reasons[dt] = "new_date"
            elif new_mask.any() and col_series[new_mask].isna().any():
                selected.append(dt)
                reasons[dt] = "new_bond"
        return selected, reasons

    trade_calendar = None
    if trade_dates is not None:
        trade_calendar = pd.DatetimeIndex(pd.to_datetime(pd.Series(list(trade_dates)), errors="coerce").dropna().unique()).sort_values()

    def _get_offset_trade_date(end_dt, days_ago):
        """按交易日历回溯 days_ago 个交易日。"""
        if trade_calendar is None or len(trade_calendar) == 0:
            return None
        end_ts = pd.Timestamp(end_dt)
        loc = int(trade_calendar.searchsorted(end_ts))
        if loc >= len(trade_calendar) or trade_calendar[loc] != end_ts:
            return None
        idx = int(loc - int(days_ago))
        if idx < 0:
            return None
        return trade_calendar[idx]

    def _normalize_formula_literal(v):
        """把值转换为可放进Excel公式的字符串字面量内容（不含外层引号）。"""
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return ""
        if isinstance(v, (pd.Timestamp, datetime, date, np.datetime64)):
            ts = pd.Timestamp(v)
            if pd.isna(ts):
                return ""
            return ts.strftime("%Y-%m-%d")
        return str(v)

    def _build_lifecycle_fetch_bounds(current_sheet_name):
        """生成公式抓取边界；抓取边界与最终数据清洗边界相互独立。"""
        if current_sheet_name not in MONTHLY_METRICS or current_sheet_name in UNCHANGED_BACKTEST_SHEETS:
            return None, None, False

        total_df = updated.get("总表")
        if total_df is None or total_df.empty:
            raise ValueError(f"{current_sheet_name} 历史缺失补抓失败：缺少总表生命周期信息。")

        # 转股价已有数据不做生命周期清洗，但发行前为空属于结构性空档，禁止补抓；
        # 发行日之后不设结束边界，避免改变“转股价不做生命周期处理”的既有规则。
        if current_sheet_name == "转股价":
            if "发行日期" not in total_df.columns:
                raise ValueError("转股价历史缺失补抓失败：总表缺少字段 发行日期")
            bounds = total_df.loc[:, ["发行日期"]].copy()
            bounds.index = bounds.index.astype(str)
            bounds = bounds.loc[~bounds.index.duplicated(keep="last")]
            issue_dates = pd.to_datetime(bounds["发行日期"], errors="coerce").to_dict()
            return issue_dates, None, True

        if current_sheet_name in BOND_LIFECYCLE_EXEMPT_SHEETS:
            return None, None, False

        required_columns = {"上市日期", "最后交易日"}
        missing_columns = sorted(required_columns - set(total_df.columns))
        if missing_columns:
            raise ValueError(
                f"{current_sheet_name} 历史缺失补抓失败：总表缺少字段 {missing_columns}"
            )

        bounds = total_df.loc[:, ["上市日期", "最后交易日"]].copy()
        bounds.index = bounds.index.astype(str)
        bounds = bounds.loc[~bounds.index.duplicated(keep="last")]
        start_dates = pd.to_datetime(bounds["上市日期"], errors="coerce").to_dict()
        end_dates = pd.to_datetime(bounds["最后交易日"], errors="coerce").to_dict()
        # 未取得上市日期的券视为尚未上市，不提交任何个券日度抓取任务。
        return start_dates, end_dates, True

    # 当总表已在独立阶段更新完成时，后续历史/增量阶段直接复用其中的发行日期。
    _refresh_issue_from_total()

    # 与模块级常量保持一致，避免两处维护
    text_result_sheets = _TEXT_RESULT_SHEETS

    app = None
    wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.add()
        sht = wb.sheets[0]

        for cfg in sheet_configs:
            sheet_name = cfg["sheet_name"]
            rule_type = cfg.get("rule_type", "formula")
            formula_expr = cfg.get("formula_expr")
            stage_label = "总表更新" if rule_type == "total_table_formula" else call_stage_label
            stage_prefix = f"[{stage_label}]"

            if sheet_name not in updated:
                print(f"{stage_prefix} 跳过：{sheet_name} 不存在")
                continue

            df = updated[sheet_name]
            if df is None or df.empty:
                print(f"{stage_prefix} 跳过：{sheet_name} 为空")
                continue

            if rule_type == "derived_ratio":
                numerator_sheet = cfg.get("numerator_sheet")
                denominator_sheet = cfg.get("denominator_sheet")
                scale = float(cfg.get("scale", 100.0))

                if numerator_sheet not in updated or denominator_sheet not in updated:
                    print(f"{stage_prefix} 跳过：{sheet_name} 缺少依赖表 {numerator_sheet}/{denominator_sheet}")
                    continue

                target_df = df.copy()
                numerator_df = updated[numerator_sheet]
                denominator_df = updated[denominator_sheet]

                if numerator_df is None or denominator_df is None:
                    print(f"{stage_prefix} 跳过：{sheet_name} 依赖表为空")
                    continue

                tgt_dt_map = _build_dt_to_col(target_df)
                num_dt_map = _build_dt_to_col(numerator_df)
                den_dt_map = _build_dt_to_col(denominator_df)

                new_dt_list, new_dt_reason = _collect_target_dates(target_df)

                common_dates = [dt for dt in sorted(new_dt_list) if dt in num_dt_map and dt in den_dt_map]
                if not common_dates:
                    print(f"{stage_prefix} {sheet_name} 无可计算新增日期，跳过。")
                    continue

                common_index = target_df.index.intersection(numerator_df.index).intersection(denominator_df.index)
                if len(common_index) == 0:
                    print(f"{stage_prefix} {sheet_name} 无代码交集，跳过。")
                    continue

                # 派生结果一定是数值；若目标列被 parquet 回读成 string dtype，
                # 直接赋 float 会报 "Invalid value for dtype 'str'"，预先转 numeric。
                for dt in common_dates:
                    tgt_col = tgt_dt_map[dt]
                    if not pd.api.types.is_numeric_dtype(target_df[tgt_col]):
                        target_df[tgt_col] = pd.to_numeric(target_df[tgt_col], errors="coerce")

                min_common_date = min(common_dates).strftime("%Y-%m-%d")
                max_common_date = max(common_dates).strftime("%Y-%m-%d")
                derived_preview_target_count = 0
                for preview_dt in common_dates:
                    preview_col = tgt_dt_map[preview_dt]
                    if new_dt_reason.get(preview_dt) == "new_bond":
                        preview_mask = common_index.astype(str).isin(new_bond_codes_set)
                        preview_index = common_index[preview_mask]
                    else:
                        preview_index = common_index
                    derived_preview_target_count += int(
                        target_df.loc[preview_index, preview_col].isna().sum()
                    )
                print(
                    f"{stage_prefix} {sheet_name}｜{min_common_date}~{max_common_date}"
                    f"｜阶段任务 {derived_preview_target_count}｜开始更新"
                )
                filled_derived_count = 0
                derived_written_dates = set()
                derived_target_count = 0
                derived_empty_count = 0
                derived_error_count = 0
                derived_zero_count = 0
                for dt in common_dates:
                    tgt_col = tgt_dt_map[dt]
                    num_col = num_dt_map[dt]
                    den_col = den_dt_map[dt]
                    scope_index = pd.Index([])
                    try:
                        # 仅对"新券补档"的日期列缩窄到新券子集，避免重算老券
                        if new_dt_reason.get(dt) == "new_bond":
                            scope_mask = common_index.astype(str).isin(new_bond_codes_set)
                            candidate_index = common_index[scope_mask]
                            missing_mask = target_df.loc[candidate_index, tgt_col].isna()
                            scope_index = candidate_index[missing_mask.to_numpy()]
                        else:
                            missing_mask = target_df.loc[common_index, tgt_col].isna()
                            scope_index = common_index[missing_mask.to_numpy()]

                        if len(scope_index) == 0:
                            continue
                        derived_target_count += len(scope_index)

                        numerator_s = pd.to_numeric(numerator_df.loc[scope_index, num_col], errors="coerce")
                        denominator_s = pd.to_numeric(denominator_df.loc[scope_index, den_col], errors="coerce")
                        with np.errstate(divide="ignore", invalid="ignore"):
                            derived_s = (numerator_s / denominator_s - 1.0) * scale
                        derived_s = derived_s.replace([np.inf, -np.inf], np.nan)
                        zero_mask = derived_s.eq(0) & derived_s.notna()
                        empty_mask = derived_s.isna()
                        fill_mask = ~empty_mask & ~zero_mask
                        one_empty_count = int(empty_mask.sum())
                        one_zero_count = int(zero_mask.sum())
                        one_filled_count = int(fill_mask.sum())
                        if fill_mask.any():
                            fill_index = scope_index[fill_mask.values]
                            target_df.loc[fill_index, tgt_col] = derived_s.loc[fill_index]
                            derived_written_dates.add(pd.Timestamp(dt))
                        derived_empty_count += one_empty_count
                        derived_zero_count += one_zero_count
                        filled_derived_count += one_filled_count
                    except Exception as e:
                        derived_error_count += len(scope_index)
                        print(f"{stage_prefix} {sheet_name} {dt.strftime('%Y-%m-%d')} 派生计算报错: {e}")

                print(
                    f"{stage_prefix} {sheet_name} 更新完成｜阶段任务 {derived_target_count}"
                    f"｜成功写入 {filled_derived_count}｜空值 {derived_empty_count}"
                    f"｜抓取中 0｜错误 {derived_error_count}｜零值丢弃 {derived_zero_count}"
                    "｜批次失败 0批/0项"
                )

                updated[sheet_name] = target_df
                _record_historical_backfill(
                    historical_backfill_audit,
                    stage=fetch_scope,
                    item_name=sheet_name,
                    written_count=filled_derived_count,
                    written_dates=derived_written_dates,
                )
                latest_result_date = max(common_dates)
                if new_dt_reason.get(latest_result_date) == "new_bond":
                    latest_result_indices = [
                        idx for idx in common_index
                        if str(idx) in new_bond_codes_set
                    ]
                else:
                    latest_result_indices = list(common_index)
                _print_update_result(
                    sheet_name,
                    target_df,
                    result_date=latest_result_date,
                    target_indices=latest_result_indices,
                    stage=fetch_scope,
                )
                continue

            if rule_type == "derived_term_from_total":
                total_sheet_name = cfg.get("total_sheet_name", "总表")
                if total_sheet_name not in updated or updated[total_sheet_name] is None:
                    print(f"{stage_prefix} 跳过：{sheet_name} 缺少依赖表 {total_sheet_name}")
                    continue

                total_df = updated[total_sheet_name]
                target_df = df.copy()

                start_col = cfg.get("total_start_col")
                end_col = cfg.get("total_end_col")
                missing_total_cols = [
                    col for col in (start_col, end_col)
                    if not col or col not in total_df.columns
                ]
                if missing_total_cols:
                    print(
                        f"{stage_prefix} 跳过：{sheet_name} 的总表缺少字段 "
                        f"{missing_total_cols}"
                    )
                    continue

                tgt_dt_map = _build_dt_to_col(target_df)
                new_dates, new_dt_reason = _collect_target_dates(target_df)
                if not new_dates:
                    print(f"{stage_prefix} {sheet_name} 无新增日期，跳过。")
                    continue

                common_index = target_df.index.intersection(total_df.index)
                if len(common_index) == 0:
                    print(f"{stage_prefix} {sheet_name} 与{total_sheet_name}无代码交集，跳过。")
                    continue

                start_dates_all = pd.to_datetime(total_df.loc[common_index, start_col], errors="coerce")
                end_dates_all = pd.to_datetime(total_df.loc[common_index, end_col], errors="coerce")

                # 剩余期限等派生结果是数值，避免被 parquet 漂移成 string 的列拒写 float
                for dt in new_dates:
                    tgt_col = tgt_dt_map[dt]
                    if not pd.api.types.is_numeric_dtype(target_df[tgt_col]):
                        target_df[tgt_col] = pd.to_numeric(target_df[tgt_col], errors="coerce")

                min_new_date = min(new_dates).strftime("%Y-%m-%d")
                max_new_date = max(new_dates).strftime("%Y-%m-%d")
                term_preview_target_count = 0
                for preview_dt in sorted(new_dates):
                    preview_col = tgt_dt_map[preview_dt]
                    if new_dt_reason.get(preview_dt) == "new_bond":
                        preview_mask = common_index.astype(str).isin(new_bond_codes_set)
                        preview_candidates = common_index[preview_mask]
                    else:
                        preview_candidates = common_index
                    preview_missing = target_df.loc[preview_candidates, preview_col].isna()
                    preview_index = preview_candidates[preview_missing.to_numpy()]
                    preview_starts = start_dates_all.loc[preview_index]
                    preview_ends = end_dates_all.loc[preview_index]
                    preview_date = pd.Timestamp(preview_dt)
                    preview_valid = (
                        preview_starts.notna()
                        & preview_ends.notna()
                        & (preview_date >= preview_starts)
                        & (preview_date <= preview_ends)
                    )
                    term_preview_target_count += int(preview_valid.sum())
                print(
                    f"{stage_prefix} {sheet_name}｜{min_new_date}~{max_new_date}"
                    f"｜阶段任务 {term_preview_target_count}｜开始更新"
                )
                filled_term_count = 0
                term_written_dates = set()
                term_target_count = 0
                term_empty_count = 0
                term_error_count = 0
                term_zero_count = 0
                for dt in sorted(new_dates):
                    tgt_col = tgt_dt_map[dt]
                    valid_idx = pd.Index([])
                    try:
                        # 对“新券补档”缩窄到新券子集，避免误覆盖老券
                        if new_dt_reason.get(dt) == "new_bond":
                            scope_mask = common_index.astype(str).isin(new_bond_codes_set)
                            candidate_index = common_index[scope_mask]
                            missing_mask = target_df.loc[candidate_index, tgt_col].isna()
                            scope_index = candidate_index[missing_mask.to_numpy()]
                        else:
                            missing_mask = target_df.loc[common_index, tgt_col].isna()
                            scope_index = common_index[missing_mask.to_numpy()]

                        if len(scope_index) == 0:
                            continue

                        start_dates = start_dates_all.loc[scope_index]
                        end_dates = end_dates_all.loc[scope_index]

                        date_ts = pd.Timestamp(dt)
                        valid_mask = start_dates.notna() & end_dates.notna() & (date_ts >= start_dates) & (date_ts <= end_dates)
                        if valid_mask.any():
                            valid_idx = scope_index[valid_mask.values]
                            term_target_count += len(valid_idx)
                            calc_values = (end_dates.loc[valid_idx] - date_ts).dt.days / 365.0
                            zero_mask = calc_values.eq(0) & calc_values.notna()
                            empty_mask = calc_values.isna()
                            fill_mask = ~empty_mask & ~zero_mask
                            one_empty_count = int(empty_mask.sum())
                            one_zero_count = int(zero_mask.sum())
                            one_filled_count = int(fill_mask.sum())
                            if fill_mask.any():
                                fill_idx = valid_idx[fill_mask.values]
                                target_df.loc[fill_idx, tgt_col] = calc_values.loc[fill_idx]
                                term_written_dates.add(pd.Timestamp(dt))
                            term_empty_count += one_empty_count
                            term_zero_count += one_zero_count
                            filled_term_count += one_filled_count
                    except Exception as e:
                        term_error_count += len(valid_idx)
                        print(f"{stage_prefix} {sheet_name} {dt.strftime('%Y-%m-%d')} 计算报错: {e}")

                print(
                    f"{stage_prefix} {sheet_name} 更新完成｜阶段任务 {term_target_count}"
                    f"｜成功写入 {filled_term_count}｜空值 {term_empty_count}"
                    f"｜抓取中 0｜错误 {term_error_count}｜零值丢弃 {term_zero_count}"
                    "｜批次失败 0批/0项"
                )

                updated[sheet_name] = target_df
                _record_historical_backfill(
                    historical_backfill_audit,
                    stage=fetch_scope,
                    item_name=sheet_name,
                    written_count=filled_term_count,
                    written_dates=term_written_dates,
                )
                latest_result_date = max(new_dates)
                if new_dt_reason.get(latest_result_date) == "new_bond":
                    latest_result_indices = [
                        idx for idx in common_index
                        if str(idx) in new_bond_codes_set
                    ]
                else:
                    latest_result_indices = list(common_index)
                _print_update_result(
                    sheet_name,
                    target_df,
                    result_date=latest_result_date,
                    target_indices=latest_result_indices,
                    stage=fetch_scope,
                )
                continue

            if rule_type == "total_table_formula":
                if skip_total_table:
                    print(f"{stage_prefix} 跳过：{sheet_name} 已在新券历史更新中完成")
                    continue
                total_df = df.copy()
                fields = cfg.get("fields", [])
                update_mode = cfg.get("update_mode", "only_blank")
                if not fields:
                    print(f"{stage_prefix} 跳过：{sheet_name} 未配置 fields")
                    continue

                for field_cfg in fields:
                    target_col = field_cfg.get("target_col")
                    formula_expr_field = field_cfg.get("formula_expr")
                    if not target_col or not formula_expr_field or "{r}" not in formula_expr_field:
                        print(f"{stage_prefix} {sheet_name} 跳过无效字段配置: {field_cfg}")
                        continue

                    if target_col not in total_df.columns:
                        # 针对文本字段（如转债名称、申万行业等）初始化为object类型，避免写入str时报错
                        if target_col in {"转债名称", "申万行业", "申万二级行业", "申万三级行业", "下修条款全文"}:
                            total_df[target_col] = pd.Series([None] * len(total_df), index=total_df.index, dtype=object)
                        else:
                            total_df[target_col] = np.nan


                    # 对“最后交易日”“赎回公告日”始终全量更新，其余字段按 only_blank
                    if target_col in {"最后交易日", "赎回公告日"}:
                        fill_mask = pd.Series(True, index=total_df.index)
                    elif target_col == "转债名称":
                        name_text = total_df[target_col].astype("string").str.strip()
                        fill_mask = total_df[target_col].isna() | name_text.str.lower().isin(
                            {"", "——", "--", "-", "nan", "nat", "none", "<na>"}
                        )
                    else:
                        fill_mask = total_df[target_col].isna()

                    codes_to_fetch = [str(c) for c in total_df.index[fill_mask].tolist()]
                    if not codes_to_fetch:
                        print(f"{stage_prefix} {sheet_name}.{target_col} 无需更新")
                        continue

                    filled_field_count = 0
                    field_empty_count = 0
                    field_fetching_count = 0
                    field_error_count = 0
                    field_zero_count = 0
                    field_failed_batch_count = 0
                    field_failed_task_count = 0
                    print(
                        f"{stage_prefix} {sheet_name}.{target_col}｜静态字段"
                        f"｜阶段任务 {len(codes_to_fetch)}｜批次 1｜开始更新"
                    )
                    try:
                        sht.range((1, 1), (max(len(codes_to_fetch) + 2, 10), 2)).clear_contents()
                        sht.range((1, 1)).value = "转债代码"
                        sht.range((1, 2)).value = "数值"

                        sht.range((2, 1)).options(transpose=True).value = codes_to_fetch
                        formula_range = sht.range((2, 2), (len(codes_to_fetch) + 1, 2))
                        formulas = [[formula_expr_field.format(r=r)] for r in range(2, len(codes_to_fetch) + 2)]
                        formula_range.formula = formulas

                        waited = 0.0
                        values = []
                        while True:
                            app.calculate()
                            raw_values = formula_range.value
                            if len(codes_to_fetch) == 1:
                                values = [raw_values]
                            else:
                                values = [row[0] if isinstance(row, list) else row for row in raw_values]

                            fetching_count = sum(1 for v in values if _is_ifind_fetching_value(v))
                            empty_count = sum(1 for v in values if _is_update_empty_value(v))
                            total_items = len(values)
                            # 全空视同仍在抓取（Excel 公式算术错误会被 xlwings 读成 None），
                            # 避免 fetching_count=0 但实际没数据时提前退出。
                            if fetching_count == 0 and empty_count < total_items:
                                break
                            if waited >= max_wait_seconds:
                                break

                            pytime.sleep(wait_step_seconds)
                            waited += wait_step_seconds

                        for task_pos, code in enumerate(codes_to_fetch):
                            if task_pos >= len(values):
                                field_error_count += 1
                                continue
                            val = values[task_pos]
                            if _is_ifind_fetching_value(val):
                                field_fetching_count += 1
                                continue
                            if _is_update_empty_value(val):
                                field_empty_count += 1
                                continue
                            if _is_update_error_value(val):
                                field_error_count += 1
                                continue
                            # 对总表的数值字段（如发行规模），插件返回 0 通常代表"查不到"，
                            # 直接跳过避免把 0 当真实数据落盘。
                            if _value_is_numeric_zero(val):
                                field_zero_count += 1
                                continue
                            if code in total_df.index:
                                try:
                                    total_df.at[code, target_col] = val
                                except Exception:
                                    field_error_count += 1
                                    continue
                                filled_field_count += 1
                    except Exception as e:
                        filled_field_count = 0
                        field_empty_count = 0
                        field_fetching_count = 0
                        field_error_count = 0
                        field_zero_count = 0
                        field_failed_batch_count = 1
                        field_failed_task_count = len(codes_to_fetch)
                        print(f"{stage_prefix} {sheet_name}.{target_col} 报错: {e}")

                    print(
                        f"{stage_prefix} {sheet_name}.{target_col} 更新完成"
                        f"｜阶段任务 {len(codes_to_fetch)}｜成功写入 {filled_field_count}"
                        f"｜空值 {field_empty_count}｜抓取中 {field_fetching_count}"
                        f"｜错误 {field_error_count}｜零值丢弃 {field_zero_count}"
                        f"｜批次失败 {field_failed_batch_count}批/{field_failed_task_count}项"
                    )

                    _print_update_result(
                        f"{sheet_name}.{target_col}",
                        total_df[target_col],
                        target_indices=codes_to_fetch,
                        stage=fetch_scope,
                    )

                updated[sheet_name] = total_df
                # 总表"发行日期"列已通过 cb_list_dtonl 公式填好，刷新新券发行日映射
                _refresh_issue_from_total()
                continue

            row_formula_map = {
                str(k): v for k, v in (cfg.get("row_formula_map") or {}).items()
            }

            # 默认规则：Excel 公式抓取。逐行公式映射与单一公式二选一；
            # 有映射时不设置默认回退函数，避免本地派生序列被误抓成其他指标。
            if row_formula_map:
                invalid_formula_rows = sorted(
                    code for code, expr in row_formula_map.items()
                    if not expr or "{r}" not in expr
                )
                if invalid_formula_rows:
                    print(
                        f"{stage_prefix} 跳过：{sheet_name} 的逐行公式配置无效："
                        f"{invalid_formula_rows[:10]}"
                    )
                    continue
            elif not formula_expr or "{r}" not in formula_expr:
                print(f"{stage_prefix} 跳过：{sheet_name} 的 formula_expr 无效，需包含 {{r}} 行占位符")
                continue

            total_lookup = None
            if "total_lookup_col" in cfg:
                total_sheet_name = cfg.get("total_sheet_name", "总表")
                if total_sheet_name not in updated or updated[total_sheet_name] is None:
                    print(f"{stage_prefix} 跳过：{sheet_name} 需要 {total_sheet_name} 但不存在")
                    continue
                total_df = updated[total_sheet_name]
                lookup_col = cfg.get("total_lookup_col")
                if not lookup_col or lookup_col not in total_df.columns:
                    print(f"{stage_prefix} 跳过：{sheet_name} 的总表缺少字段 {lookup_col}")
                    continue
                total_lookup = total_df[lookup_col]
            fetch_full_history = (
                bool(cfg.get("new_bond_fetch_full_history", False))
                or sheet_name in FULL_HISTORY_NEW_BOND_SHEETS
            )
            eligible_start_dates, eligible_end_dates, require_eligible_start = (
                _build_lifecycle_fetch_bounds(sheet_name)
            )
            tasks = build_formula_fetch_tasks(
                df,
                fetch_scope=fetch_scope,
                history_cutoff=history_cutoff_ts,
                new_bond_codes=new_bond_codes_set,
                new_bond_issue_dates=merged_issue_ts,
                fetch_full_history=fetch_full_history,
                include_historical_missing=(
                    fetch_scope == "historical_missing"
                    and sheet_name not in STRUCTURAL_EMPTY_SHEETS
                ),
                incremental_by_row_latest=(
                    fetch_scope == "historical_missing" and sheet_name == "指数"
                ),
                eligible_start_dates=eligible_start_dates,
                eligible_end_dates=eligible_end_dates,
                require_eligible_start=require_eligible_start,
            )

            # 配置了逐行公式映射时，只允许更新映射中明确登记的行。
            # “指数”中还包含多因子拟合等本地模型序列；这些序列由各自脚本写回，
            # 不得回退使用指数配置的默认公式（此前默认值是十年国债函数）。
            if row_formula_map:
                unmapped_task_codes = sorted({
                    code for code, *_ in tasks if code not in row_formula_map
                })
                tasks = [task for task in tasks if task[0] in row_formula_map]
                if unmapped_task_codes:
                    print(
                        f"{stage_prefix} {sheet_name} 跳过 {len(unmapped_task_codes)} 个"
                        "未配置函数的本地派生序列。"
                    )

            # 需要交易日回溯的公式，若某个历史日期无法找到起始/截止交易日，只跳过该任务。
            prepared_tasks = []
            skipped_offset = 0
            for code, idx, col, ts in tasks:
                start_ts = ts
                end_ts = ts
                if "date_start_trade_days_ago" in cfg:
                    start_ts = _get_offset_trade_date(ts, cfg.get("date_start_trade_days_ago", 0))
                    if start_ts is None:
                        skipped_offset += 1
                        continue
                if "date_end_trade_days_ago" in cfg:
                    end_ts = _get_offset_trade_date(ts, cfg.get("date_end_trade_days_ago", 0))
                    if end_ts is None:
                        skipped_offset += 1
                        continue
                prepared_tasks.append(
                    (code, idx, col, ts, pd.Timestamp(start_ts), pd.Timestamp(end_ts))
                )
            tasks = prepared_tasks

            if not tasks:
                suffix = f"（另有 {skipped_offset} 个任务无法回溯所需交易日）" if skipped_offset else ""
                print(f"{stage_prefix} {sheet_name} 无待抓取差异，跳过。{suffix}")
                continue

            filled = df.copy()
            task_cols = list(dict.fromkeys(col for _, _, col, _, _, _ in tasks))
            if sheet_name in text_result_sheets:
                for col in task_cols:
                    if filled[col].dtype != object:
                        filled[col] = filled[col].astype(object)
            else:
                for col in task_cols:
                    if not pd.api.types.is_numeric_dtype(filled[col]):
                        filled[col] = pd.to_numeric(filled[col], errors="coerce")

            n = len(tasks)
            min_task_date = min(ts for _, _, _, ts, _, _ in tasks).strftime("%Y-%m-%d")
            max_task_date = max(ts for _, _, _, ts, _, _ in tasks).strftime("%Y-%m-%d")
            batch_total = (n + formula_batch_size - 1) // formula_batch_size
            print(
                f"{stage_prefix} {sheet_name}｜{min_task_date}~{max_task_date}"
                f"｜阶段任务 {n}｜批次 {batch_total}｜开始更新"
            )

            filled_cnt = 0
            filled_dates = set()
            stage_empty_count = 0
            stage_fetching_count = 0
            stage_error_count = 0
            stage_zero_count = 0
            batch_error_count = 0
            batch_failed_task_count = 0
            previous_used_end_row = 10
            for batch_no, batch_start in enumerate(range(0, n, formula_batch_size), start=1):
                batch_tasks = tasks[batch_start: batch_start + formula_batch_size]
                batch_n = len(batch_tasks)
                try:
                    clear_end_row = max(previous_used_end_row, batch_n + 1, 10)
                    sht.range((1, 1), (clear_end_row, 3)).clear_contents()
                    previous_used_end_row = batch_n + 1
                    sht.range((1, 1)).value = "转债代码"
                    sht.range((1, 2)).value = "日期"
                    sht.range((1, 3)).value = "数值"

                    codes = [code for code, _, _, _, _, _ in batch_tasks]
                    date_strs = [ts.strftime("%Y-%m-%d") for _, _, _, ts, _, _ in batch_tasks]
                    sht.range((2, 1)).options(transpose=True).value = codes
                    sht.range((2, 2)).options(transpose=True).value = date_strs

                    formulas = []
                    for r, (code, _, _, ts, start_ts, end_ts) in enumerate(batch_tasks, start=2):
                        total_val = ""
                        if total_lookup is not None and code in total_lookup.index:
                            total_val = _normalize_formula_literal(total_lookup.loc[code])
                        code_formula_expr = (
                            row_formula_map[code] if row_formula_map else formula_expr
                        )
                        formulas.append([code_formula_expr.format(
                            r=r,
                            start_date=start_ts.strftime("%Y-%m-%d"),
                            end_date=end_ts.strftime("%Y-%m-%d"),
                            total_col_value=total_val,
                        )])

                    formula_range = sht.range((2, 3), (batch_n + 1, 3))
                    formula_range.formula = formulas

                    waited = 0.0
                    values = []
                    while True:
                        app.calculate()
                        raw_values = formula_range.value
                        if batch_n == 1:
                            values = [raw_values]
                        else:
                            values = [row[0] if isinstance(row, list) else row for row in raw_values]

                        fetching_count = sum(1 for v in values if _is_ifind_fetching_value(v))
                        empty_count = sum(1 for v in values if _is_update_empty_value(v))
                        if fetching_count == 0 and empty_count < batch_n:
                            break
                        if waited >= max_wait_seconds:
                            break

                        pytime.sleep(wait_step_seconds)
                        waited += wait_step_seconds

                except Exception as e:
                    batch_error_count += 1
                    batch_failed_task_count += batch_n
                    print(
                        f"{stage_prefix} {sheet_name} 更新报错｜批次 {batch_no}/{batch_total}: {e}"
                    )
                    continue

                for task_pos, (_, idx, col, task_date, _, _) in enumerate(batch_tasks):
                    if task_pos >= len(values):
                        stage_error_count += 1
                        continue
                    val = values[task_pos]
                    if _is_ifind_fetching_value(val):
                        stage_fetching_count += 1
                        continue
                    if _is_update_empty_value(val):
                        stage_empty_count += 1
                        continue
                    if _is_update_error_value(val):
                        stage_error_count += 1
                        continue
                    if _value_is_numeric_zero(val):
                        stage_zero_count += 1
                        continue
                    try:
                        filled.at[idx, col] = val
                    except Exception:
                        stage_error_count += 1
                        continue
                    filled_cnt += 1
                    filled_dates.add(pd.Timestamp(task_date))

            print(
                f"{stage_prefix} {sheet_name} 更新完成｜阶段任务 {n}｜成功写入 {filled_cnt}"
                f"｜空值 {stage_empty_count}｜抓取中 {stage_fetching_count}"
                f"｜错误 {stage_error_count}｜零值丢弃 {stage_zero_count}"
                f"｜批次失败 {batch_error_count}批/{batch_failed_task_count}项"
            )

            updated[sheet_name] = filled
            _record_historical_backfill(
                historical_backfill_audit,
                stage=fetch_scope,
                item_name=sheet_name,
                written_count=filled_cnt,
                written_dates=filled_dates,
            )
            latest_result_date = max(ts for _, _, _, ts, _, _ in tasks)
            latest_result_indices = [
                idx for _, idx, _, ts, _, _ in tasks
                if pd.Timestamp(ts) == pd.Timestamp(latest_result_date)
            ]
            _print_update_result(
                sheet_name,
                filled,
                result_date=latest_result_date,
                target_indices=latest_result_indices,
                stage=fetch_scope,
            )
    finally:
        if wb is not None:
            wb.close()
        if app is not None:
            app.quit()

    return updated


def parse_api_dates_from_query_result(date_query_res):
    """兼容 iFinD Date_Query 的不同返回形态，统一解析为 date 列表。"""
    if isinstance(date_query_res, str):
        tokens = [x.strip() for x in date_query_res.split(",") if x and x.strip()]
        return [datetime.strptime(x, "%Y-%m-%d").date() for x in tokens]

    if isinstance(date_query_res, pd.DataFrame):
        arr = date_query_res.to_numpy().ravel().tolist()
    elif isinstance(date_query_res, dict):
        arr = []
        for v in date_query_res.values():
            if isinstance(v, (list, tuple, np.ndarray, pd.Series, pd.Index)):
                arr.extend(list(v))
            else:
                arr.append(v)
    elif isinstance(date_query_res, (list, tuple, np.ndarray, pd.Series, pd.Index)):
        arr = list(date_query_res)
    else:
        arr = [date_query_res]

    parsed = pd.to_datetime(pd.Series(arr), errors="coerce").dropna()
    return [x.date() for x in parsed.tolist()]


def build_trade_date_index(date_query_res, start_date="2015-01-01"):
    """将 iFinD 返回的交易日结果解析为 DatetimeIndex。"""
    start_ts = pd.Timestamp(start_date)

    if isinstance(date_query_res, pd.DataFrame):
        raw_values = date_query_res.to_numpy().ravel().tolist()
    elif isinstance(date_query_res, dict):
        raw_values = []
        for value in date_query_res.values():
            if isinstance(value, (list, tuple, np.ndarray, pd.Series, pd.Index)):
                raw_values.extend(list(value))
            else:
                raw_values.append(value)
    elif isinstance(date_query_res, (list, tuple, np.ndarray, pd.Series, pd.Index)):
        raw_values = list(date_query_res)
    else:
        raw_values = [date_query_res]

    trade_dates = pd.to_datetime(pd.Series(raw_values), errors="coerce")
    trade_dates = trade_dates.dropna()
    trade_dates = trade_dates[trade_dates >= start_ts]
    trade_dates = pd.DatetimeIndex(sorted(pd.unique(trade_dates)))

    if trade_dates.empty:
        raise ValueError("Date_Query_res 未解析出 2015-01-01 及之后的有效交易日。")

    return trade_dates


def clip_trade_dates_before_data_ready(api_dates, cutoff_hour=17, now=None):
    """若运行日恰好是日历最新交易日且尚未到 cutoff_hour，则剔除当天，更新至上一交易日。

    盘后许多截面数据要到傍晚才齐；在最新交易日过早更新容易缺数，因此默认 17 点前
    不把“当天”纳入更新截止日。非交易日运行、或已过 cutoff，则原样返回。
    """
    if not api_dates:
        return api_dates

    now = datetime.now() if now is None else now
    today = now.date() if isinstance(now, datetime) else date.today()
    latest = max(api_dates)
    if latest != today or now.hour >= cutoff_hour:
        return api_dates

    clipped = [d for d in api_dates if d < latest]
    if not clipped:
        return api_dates

    prev = max(clipped)
    print(
        f"[trade_dates] 当前 {now.strftime('%Y-%m-%d %H:%M')}，未到 {cutoff_hour}:00，"
        f"最新交易日 {latest} 数据可能未齐，更新截止日回退至上一交易日 {prev}。"
    )
    return clipped


def extend_data_columns_to_latest(original_data, trade_dates, skip_sheets=None):
    """静默将宽表 sheet 的日期列补齐到最新交易日，总表等静态表跳过。"""
    if skip_sheets is None:
        skip_sheets = {"总表"}

    updated_data = {}
    for sheet_name, df in original_data.items():
        if df is None or sheet_name in skip_sheets:
            updated_data[sheet_name] = df
            continue

        parsed_cols = pd.to_datetime(pd.Index(df.columns), errors="coerce")
        date_col_mask = np.asarray(pd.notna(parsed_cols), dtype=bool)

        if not date_col_mask.any():
            updated_data[sheet_name] = df
            continue

        non_date_cols = df.columns[~date_col_mask].tolist()
        existing_date_cols = pd.DatetimeIndex(parsed_cols[date_col_mask])
        existing_date_cols = existing_date_cols[existing_date_cols >= trade_dates.min()]
        target_date_cols = existing_date_cols.union(trade_dates).sort_values()

        date_part = df.loc[:, date_col_mask].copy()
        date_part.columns = pd.DatetimeIndex(parsed_cols[date_col_mask])
        date_part = date_part.loc[:, ~date_part.columns.duplicated(keep="first")]
        date_part = date_part.reindex(columns=target_date_cols)

        if non_date_cols:
            updated_df = pd.concat([df[non_date_cols], date_part], axis=1)
        else:
            updated_df = date_part

        updated_data[sheet_name] = updated_df

    return updated_data


def export_original_data_to_parquet(original_data, output_root="data/转债个券历史序列"):
    return _export_standard_parquet(original_data, output_root)


def read_original_data_from_parquet(input_root="data/转债个券历史序列"):
    return _read_standard_parquet(input_root)


def fetch_latest_board_members():
    """从 iFinD 获取最新可转债板块成分。

    过滤规则与历史脚本保持一致：剔除定向发行、NQ 代码、终止上市的券。

    返回：
        pd.DataFrame: index=转债代码，columns=['转债简称', '发行方式', '交易状态']
    """
    empty_df = pd.DataFrame(columns=['转债简称', '发行方式', '交易状态'])
    current_date = pytime.strftime("%Y%m%d", pytime.localtime())

    try:
        cb_list_res = THS_DR(
            'p00570',
            f'jyzt=未到期;sfdb=全部;jysc=全部;edate={current_date}',
            'jydm:Y,jydm_mc:Y,p00570_f001:Y,p00570_f019:Y',
            'format:dataframe',
        ).data
    except Exception as e:
        print(f"[board] THS_DR p00570 调用失败: {e}")
        return empty_df

    if cb_list_res is None or (hasattr(cb_list_res, 'empty') and cb_list_res.empty):
        print("[board] THS_DR p00570 返回为空。")
        return empty_df

    cb_list_res = cb_list_res.set_index('jydm')
    codes_str = ','.join(cb_list_res.index.astype(str))
    if not codes_str:
        return empty_df

    try:
        cb_basic_res = THS_BD(
            codes_str,
            'ths_convertible_debt_short_name_cbond;ths_issue_method_cbond;ths_trading_status_bond;ths_online_issue_date_cbond',
            ';;;',
        ).data
    except Exception as e:
        print(f"[board] THS_BD 调用失败: {e}")
        return empty_df

    if cb_basic_res is None or (hasattr(cb_basic_res, 'empty') and cb_basic_res.empty):
        print("[board] THS_BD 返回为空。")
        return empty_df

    cb_basic_res = cb_basic_res.set_index('thscode').rename_axis('转债代码')
    cb_basic_res.columns = ['转债简称', '发行方式', '交易状态','网上发行日期']

    # 过滤：定向发行 / NQ 代码 / 终止上市（与历史逻辑保持一致）
    issue_method = cb_basic_res['发行方式'].fillna('').astype(str)
    cb_basic_res = cb_basic_res[~issue_method.str.contains('定向', na=False)]
    cb_basic_res = cb_basic_res[~cb_basic_res.index.astype(str).str.contains('NQ', na=False)]
    trade_status = cb_basic_res['交易状态'].fillna('').astype(str)
    cb_basic_res = cb_basic_res[~trade_status.str.contains('终止上市', na=False)]

    print(f"[board] 当前存续 {len(cb_basic_res)} 只转债。")
    return cb_basic_res


def extract_issue_dates_from_missing(missing_df, date_col="网上发行日期"):
    """从板块成分子集（通常是 missing_df）提取"网上发行日期"映射。

    fetch_latest_board_members 的 THS_BD 参数已包含 ths_online_issue_date_cbond
    字段，返回 DataFrame 带"网上发行日期"列。此函数直接从该列构造
    {code: pd.Timestamp} 映射，供 fill_sheets_with_、ifind 按发行日裁剪使用。
    """
    if missing_df is None or missing_df.empty:
        return {}
    if date_col not in missing_df.columns:
        print(f"[board] 提取发行日期失败：板块数据无列 '{date_col}'，实际列: {list(missing_df.columns)}")
        return {}
    mapping = {}
    for code, raw in missing_df[date_col].items():
        ts = pd.to_datetime(raw, errors='coerce')
        if pd.notna(ts):
            mapping[str(code)] = pd.Timestamp(ts)
    print(f"[board] 从板块成分提取发行日期：命中 {len(mapping)}/{len(missing_df)} 只")
    return mapping


def detect_missing_bonds(original_data, board_members, baseline_sheet="总表"):
    """识别板块成分里存在、但底稿基准 sheet 中缺失的转债。

    参数：
        original_data: 底稿字典
        board_members: fetch_latest_board_members 的返回
        baseline_sheet: 用作"底稿已收录个券"的基准 sheet，默认总表

    返回：
        pd.DataFrame: 与 board_members 同结构，仅包含底稿缺失的行。
    """
    empty_df = board_members.iloc[0:0].copy() if isinstance(board_members, pd.DataFrame) else pd.DataFrame()
    if board_members is None or board_members.empty:
        return empty_df
    if baseline_sheet not in original_data or original_data[baseline_sheet] is None:
        print(f"[missing] 基准 sheet '{baseline_sheet}' 不存在，跳过缺失个券检测。")
        return empty_df

    existing_codes = set(original_data[baseline_sheet].index.astype(str).tolist())
    board_codes = board_members.index.astype(str)
    missing_mask = ~board_codes.isin(existing_codes)
    missing_df = board_members[missing_mask].copy()
    return missing_df


def notify_missing_bonds(missing_df, title="检测到底稿缺失的转债"):
    """以 Windows 弹窗提示缺失个券（转债代码、转债简称、交易状态）。"""
    if missing_df is None or missing_df.empty:
        return
    try:
        header = f"共检测到 {len(missing_df)} 只需补充的转债：\n\n"
        header += f"{'转债代码':<14}{'转债简称':<16}交易状态\n"
        header += "-" * 48 + "\n"
        body_lines = []
        for code, row in missing_df.iterrows():
            short_name = str(row.get('转债简称', '') or '')
            status = str(row.get('交易状态', '') or '')
            body_lines.append(f"{str(code):<14}{short_name:<16}{status}")
        msg = header + "\n".join(body_lines)
        try:
            ctypes.windll.user32.MessageBoxW(0, msg, title, 0x40)
        except Exception:
            # 非 Windows 或无图形界面时退化为终端打印
            print(f"\n===== {title} =====")
            print(msg)
            print("=" * (len(title) + 12))
    except Exception as e:
        print(f"[missing] 弹窗显示失败: {e}")


def append_missing_bonds_to_sheets(original_data, missing_df, skip_sheets=None, allowed_sheets=None):
    """把缺失转债追加到各 sheet 的 index。

    - 宽表：所有日期列初始化为 NaN
    - 总表：所有字段列初始化为 NaN（后续由 total_table_formula 规则回填）
    - skip_sheets 中的 sheet 不受影响（如“指数”这种 index 不是转债代码的）
    - allowed_sheets 若提供，仅追加其中存在的 sheet（推荐传入 sheet_configs 的名单，
      避免误改底稿中 index 非转债代码的其他 sheet）

    返回：
        dict[str, pd.DataFrame]: 更新后的副本
    """
    if missing_df is None or missing_df.empty:
        return original_data
    if skip_sheets is None:
        skip_sheets = {"指数"}

    new_codes = [str(c) for c in missing_df.index.tolist() if str(c) not in ("", "nan", "None")]
    if not new_codes:
        return original_data

    allowed_set = set(allowed_sheets) if allowed_sheets is not None else None

    updated = dict(original_data)
    for sheet_name, df in original_data.items():
        if df is None or sheet_name in skip_sheets:
            continue
        if allowed_set is not None and sheet_name not in allowed_set:
            continue
        existing_idx = set(df.index.astype(str).tolist())
        to_add = [c for c in new_codes if c not in existing_idx]
        if not to_add:
            continue
        empty_rows = pd.DataFrame(index=pd.Index(to_add, name=df.index.name), columns=df.columns)
        new_df = pd.concat([df, empty_rows], axis=0)
        updated[sheet_name] = new_df
        print(f"[missing] {sheet_name}: 追加 {len(to_add)} 只新券 index")
    return updated


INDEX_ROW_FORMULA_MAP = {
    "十年国债": '=@b_calc_curve_chinabond("1232",B{r},"10.0y")',
    "转债指数": '=@i_dq_close("000832.CSI",B{r})',
    "万得全A": '=@i_dq_close("881001.WI",B{r})',
    "沪深300": '=@i_dq_close("000300.SH",B{r})',
    "正股等权指数": '=@i_dq_close("889035.WI",B{r})',
    "中证500": '=@i_dq_close("000905.SH",B{r})',
    "中证1000": '=@i_dq_close("000852.SH",B{r})',
    "中证2000": '=@i_dq_close("932000.CSI",B{r})',
    "中证800": '=@i_dq_close("000906.SH",B{r})',
    "中债新综合财富总指数": '=@i_dq_close("CBA00101.CS",B{r})',
    "普通股票型基金": '=@i_dq_close("885000.WI",B{r})',
    "偏股混合型基金": '=@i_dq_close("885001.WI",B{r})',
    "平衡混合型基金": '=@i_dq_close("885002.WI",B{r})',
    "偏债混合型基金": '=@i_dq_close("885003.WI",B{r})',
    "股票型基金": '=@i_dq_close("885004.WI",B{r})',
    "债券型基金": '=@i_dq_close("885005.WI",B{r})',
    "混合债券型一级基金": '=@i_dq_close("885006.WI",B{r})',
    "混合债券型二级基金": '=@i_dq_close("885007.WI",B{r})',
    "中长期纯债型基金": '=@i_dq_close("885008.WI",B{r})',
    "股票型基金总": '=@i_dq_close("885012.WI",B{r})',
    "混合型基金总": '=@i_dq_close("885013.WI",B{r})',
    "增强型基金": '=@i_dq_close("885044.WI",B{r})',
    "灵活配置型基金": '=@i_dq_close("885061.WI",B{r})',
    "短期纯债型基金": '=@i_dq_close("885062.WI",B{r})',
    "债券指数型基金": '=@i_dq_close("885063.WI",B{r})',
    "可转换债券型基金": '=@i_dq_close("885077.WI",B{r})',
    "纯债型基金总": '=@i_dq_close("885080.WI",B{r})',
    "转债等权": '=@i_dq_close("889033.WI",B{r})',
    "转债预案": '=@i_dq_close("884257.WI",B{r})',
    "上证综指": '=@i_dq_close("000001.SH",B{r})',
    "深证成指": '=@i_dq_close("399001.SZ",B{r})',
    "创业板指": '=@i_dq_close("399006.SZ",B{r})',
    "上证50": '=@i_dq_close("000016.SH",B{r})',
    "大盘指数": '=@i_dq_close("801811.SI",B{r})',
    "中盘指数": '=@i_dq_close("801812.SI",B{r})',
    "小盘指数": '=@i_dq_close("801813.SI",B{r})',
    "大盘成长": '=@i_dq_close("399372.SZ",B{r})',
    "大盘价值": '=@i_dq_close("399373.SZ",B{r})',
    "中盘成长": '=@i_dq_close("399374.SZ",B{r})',
    "中盘价值": '=@i_dq_close("399375.SZ",B{r})',
    "小盘成长": '=@i_dq_close("399376.SZ",B{r})',
    "小盘价值": '=@i_dq_close("399377.SZ",B{r})',
}


def ensure_index_sheet(original_data, trade_dates):
    """确保“指数”sheet存在且包含约定指数行。

    规则：
    - 不做“十年国债 -> 指数”自动迁移（由人工一次性改底稿）
    - 指数行应覆盖 INDEX_ROW_FORMULA_MAP 中的全部配置项
    - 缺失行会自动追加，日期列使用 trade_dates（后续仍可由 extend 流程统一补齐）
    """
    required_rows = list(INDEX_ROW_FORMULA_MAP)
    updated = dict(original_data)

    idx_df = updated.get("指数")
    if idx_df is None:
        idx_df = pd.DataFrame(index=pd.Index(required_rows, name="指数名称"), columns=trade_dates)
        updated["指数"] = idx_df
        print(f"[index] 已创建“指数”sheet，初始化 {len(required_rows)} 行。")
        return updated

    idx_df = idx_df.copy()
    existing_rows = set(idx_df.index.astype(str).tolist())
    to_add = [r for r in required_rows if r not in existing_rows]
    if to_add:
        extra = pd.DataFrame(index=pd.Index(to_add, name=idx_df.index.name), columns=idx_df.columns)
        idx_df = pd.concat([idx_df, extra], axis=0)
        print(f"[index] “指数”sheet 已补充缺失指数行 {len(to_add)} 个: {to_add}")
    updated["指数"] = idx_df
    return updated


# ================== 主流程入口 ==================
def build_sheet_configs():
    """集中维护所有 sheet 更新规则，便于后续增量扩展。"""
    return [
        {
            "sheet_name": "总表",
            "rule_type": "total_table_formula",
            "update_mode": "only_blank",
            "fields": [
                {"target_col": "转债名称", "formula_expr": '=@b_info_name(A{r})'},
                {"target_col": "上市日期", "formula_expr": '=@b_info_listeddate(A{r})'},
                {"target_col": "最后交易日", "formula_expr": '=@IF(@cb_anal_lasttradingday(A{r})=0,@b_info_maturitydate(A{r}),@cb_anal_lasttradingday(A{r}))'},
                {"target_col": "最后转股日", "formula_expr": '=@thsiFinD("ths_conversion_ed_cbond",A{r})'},
                {"target_col": "摘牌日期", "formula_expr": '=@thsiFinD("ths_delist_date_bond",A{r})'},
                {"target_col": "到期日期", "formula_expr": '=@thsiFinD("ths_maturity_date_cbond",A{r})'},
                {"target_col": "到期赎回价", "formula_expr": '=@thsiFinD("ths_maturity_redemp_price_cbond",A{r})'},
                {"target_col": "发行日期", "formula_expr": '=@cb_list_dtonl(A{r})'},
                {"target_col": "发行规模", "formula_expr": '=@cb_info_issueamount(A{r},100000000)'},
                {"target_col": "股票发行面值", "formula_expr": '=@thsiFinD("ths_stock_issue_coupon_value_stock",thsiFinD("ths_stock_code_cbond",A{r}))'},
                {"target_col": "申万行业", "formula_expr": '=@s_info_industry_sw_2021(A{r},"",1)'},
                {"target_col": "申万二级行业", "formula_expr": '=@s_info_industry_sw_2021(A{r},"",2)'},
                {"target_col": "申万三级行业", "formula_expr": '=@s_info_industry_sw_2021(A{r},"",3)'},
                {"target_col": "赎回公告日", "formula_expr": '=@IF(@cb_clause_calloption_indicativedatey(A{r})=0,"",@cb_clause_calloption_indicativedatey(A{r}))'},
                {"target_col": "转股期起始日", "formula_expr": '=@cb_clause_conversion_2_swapsharestartdate(A{r})'},
                {"target_col": "回售起始日期", "formula_expr": '=@cb_clause_putoption_conditionalputbackstartenddate(A{r})'},
                {"target_col": "赎回触发比例", "formula_expr": '=@cb_clause_calloption_triggerproportion(A{r})'},
                {"target_col": "赎回触发计算时间区间", "formula_expr": '=@cb_clause_calloption_redeemspan(A{r})'},
                {"target_col": "赎回触发计算最大时间区间", "formula_expr": '=@cb_clause_calloption_redeemmaxspan(A{r})'},
                {"target_col": "下修触发比例", "formula_expr": '=@cb_clause_reset_resettriggerratio(A{r})'},
                {"target_col": "下修条款全文", "formula_expr": '=@cb_clause_reset_item(A{r})'},
                {"target_col": "重设触发计算时间区间", "formula_expr": '=@cb_clause_reset_resettimespan(A{r})'},
                {"target_col": "重设触发计算最大时间区间", "formula_expr": '=@cb_clause_reset_resetmaxtimespan(A{r})'},
            ],
        },
        {"sheet_name": "余额", "formula_expr": '=@thsiFinD("ths_bond_balance_cbond",A{r},B{r})'},
        {"sheet_name": "收盘价", "formula_expr": '=@thsiFinD("ths_bond_close_cbond",A{r},B{r},101)'},
        {"sheet_name": "平价", "formula_expr": '=@thsiFinD("ths_transfer_value_cbond",A{r},B{r})'},
        {"sheet_name": "转股价", "formula_expr": '=@cb_anal_convprice(A{r},B{r})'},
        {"sheet_name": "平价底价溢价率", "formula_expr": '=@thsiFinD("ths_conversion_parity_price_premium_cbond",A{r},B{r})'},
        {"sheet_name": "转股溢价率", "formula_expr": '=@thsiFinD("ths_conversion_premium_rate_cbond",A{r},B{r})'},
        {"sheet_name": "纯债价值", "formula_expr": '=@cb_anal_straightbondvalue(A{r},B{r})'},
        {"sheet_name": "纯债溢价率", "formula_expr": '=@thsiFinD("ths_pure_bond_premium_rate_cbond",A{r},B{r})'},
        {"sheet_name": "YTM", "formula_expr": '=@cb_anal_ytm(A{r},B{r})'},
        {"sheet_name": "换手率", "formula_expr": '=@thsiFinD("ths_turnover_ratio_cbond",A{r},B{r})'},
        {"sheet_name": "成交量", "formula_expr": '=@thsiFinD("ths_bond_vol_yz_cbond",A{r},B{r},104)'},
        {"sheet_name": "成交额", "formula_expr": '=@thsiFinD("ths_bond_amt_cbond",A{r},B{r},101)/100000000'},
        {"sheet_name": "涨跌幅", "formula_expr": '=@thsiFinD("ths_bond_chg_ratio_cbond",A{r},B{r},101)'},
        {"sheet_name": "隐含波动率", "formula_expr": '=@cb_anal_impliedvol(A{r},B{r},3)'},
        {
            "sheet_name": "正股20日波动率",
            "formula_expr": '=@thsiFinD("ths_volatility_annual_stock",cb_info_underlyingcode(A{r}),"{start_date}","{end_date}",100,101)',
            "date_start_trade_days_ago": 20,
        },
        # {
        #     "sheet_name": "WR",
        #     "formula_expr": '=@wr(cb_info_underlyingcode(A{r}),B{r},14,"1",1)',
        # },
        # {"sheet_name": "正股换手率", "formula_expr": '=@thsiFinD("ths_stock_turnover_cbond",A{r},B{r})'},
        # {"sheet_name": "正股换手率", "formula_expr": '=@thsiFinD("ths_stock_turnover_cbond",A{r},B{r})'},
        # {"sheet_name": "PE", "formula_expr": '=@thsiFinD("ths_pe_stock",cb_info_underlyingcode(A{r}),B{r},106)'},
        # {"sheet_name": "PB", "formula_expr": '=@thsiFinD("ths_pb_stock",cb_info_underlyingcode(A{r}),B{r},108)'},
        {"sheet_name": "累计转股比例", "formula_expr": '=@thsiFinD("ths_accum_conversion_ratio_cbond",A{r},B{r})'},
        {"sheet_name": "转股稀释率", "formula_expr": '=@thsiFinD("ths_conversion_dlt_ratio_cbond",A{r},B{r})'},
        {"sheet_name": "正股市值", "formula_expr": '=@thsiFinD("ths_market_value_stock",cb_info_underlyingcode(A{r}),B{r})/100000000'},
        {"sheet_name": "每股净资产", "formula_expr": '=@thsiFinD("ths_net_asset_bps_latest_announ_stock",A{r},B{r})'},
        # {
        #     "sheet_name": "最近季度同比",
        #     "formula_expr": '=@IF(thsiFinD("ths_transfer_value_cbond",A{r},B{r})>0,thsiFinD("ths_sq_np_atsopc_yoy_bond",A{r},s_fa_latelyrd_bt(A{r},B{r})),"")',
        # },
        # {
        #     "sheet_name": "最近季度环比",
        #     "formula_expr": '=@IF(thsiFinD("ths_transfer_value_cbond",A{r},B{r})>0,thsiFinD("ths_sq_np_atsopc_lrr_bond",A{r},s_fa_latelyrd_bt(A{r},B{r})),"")',
        # },
        {"sheet_name": "EXPMA5", "formula_expr": '=@thsiFinD("ths_expma_stock",cb_info_underlyingcode(A{r}),B{r},5,100,100)'},
        {"sheet_name": "EXPMA10", "formula_expr": '=@thsiFinD("ths_expma_stock",cb_info_underlyingcode(A{r}),B{r},10,100,100)'},
        {"sheet_name": "EXPMA20", "formula_expr": '=@thsiFinD("ths_expma_stock",cb_info_underlyingcode(A{r}),B{r},20,100,100)'},
        {
            "sheet_name": "交易状态",
            "formula_expr": '=@IF(VALUE(B{r})=VALUE("{total_col_value}"),"新股上市",IF(@s_dq_tradestatus(A{r},B{r})="盘中停牌","交易",s_dq_tradestatus(A{r},B{r})))',
            "total_sheet_name": "总表",
            "total_lookup_col": "上市日期",
        },
        {"sheet_name": "主体评级", "formula_expr": '=@b_info_latestissurercreditrating2(A{r},B{r},"101","1")'},
        {"sheet_name": "债项评级", "formula_expr": '=@thsiFinD("ths_specified_date_bond_rating_bond",A{r},B{r},100)'},
        {"sheet_name": "正股收盘价", "formula_expr": '=@thsiFinD("ths_stock_close_cbond",A{r},B{r},100,"")'},
        {
            "sheet_name": "正股交易状态",
            "formula_expr": '=@thsiFinD("ths_trading_status_stock",thsiFinD("ths_stock_code_cbond",A{r}),B{r})',
        },
        {
            "sheet_name": "正股近1日均价",
            "formula_expr": '=@thsiFinD("ths_stock_avg_prx_nday_bond",A{r},-1,"{end_date}",100)',
            "date_end_trade_days_ago": 1,
        },
        {
            "sheet_name": "正股近20日均价",
            "formula_expr": '=@thsiFinD("ths_stock_avg_prx_nday_bond",A{r},-20,"{end_date}",100)',
            "date_end_trade_days_ago": 1,
        },
        {
            "sheet_name": "指数",
            "row_formula_map": INDEX_ROW_FORMULA_MAP,
        },
        {
            "sheet_name": "剩余期限",
            "rule_type": "derived_term_from_total",
            "total_start_col": "发行日期",
            "total_end_col": "到期日期",
        },
    ]


def main():
    """脚本主入口：执行登录、扩列、填充、保存。"""
    # 1) 登录 iFinD 并获取交易日历。
    thslogindemo()
    yyyymmdd_today = str(pytime.strftime("%Y-%m-%d", pytime.localtime()))
    date_query_res = THS_Date_Query('212001', 'mode:1,dateType:0,period:D,dateFormat:0', '2010-01-01', yyyymmdd_today).data
    api_dates = parse_api_dates_from_query_result(date_query_res)
    # 若当天就是最新交易日且未到 17 点，多数数据尚未齐备，更新截止日回退至上一交易日
    api_dates = clip_trade_dates_before_data_ready(api_dates, cutoff_hour=17)
    global_max_date_str = max(api_dates).strftime('%Y-%m-%d')
    print(f"✅ API 交易日历获取成功，共 {len(api_dates)} 天。更新截止交易日: {global_max_date_str}")

    # 2) 仅读取标准 Parquet 底稿；读取异常时直接终止，禁止回退旧 Excel。
    parquet_root = "data/转债个券历史序列"
    original_data = load_original_data(
        parquet_root=parquet_root,
        force_refresh=False,
    )
    # 历史底稿中存在大量 iFinD 插件遗留的 0 值（发行前、未交易、取数失败等场景），
    # 对多数指标而言它们并非真实 0，只是"无数据"。先统一清理一次，后续的扩列、
    # 补档与 parquet 落盘都能享受 null 压缩，显著降低存储与内存占用。
    original_data = sanitize_zero_values(original_data)
    trade_dates = build_trade_date_index(api_dates, start_date="2015-01-01")
    history_cutoff = get_existing_max_data_date(original_data)
    print(f"[update] 扩列前 Parquet 实际最大日期（历史截止日）: {history_cutoff.strftime('%Y-%m-%d')}")
    original_data = ensure_index_sheet(original_data, trade_dates)

    # 2.5) 先对齐最新板块成分并追加新券行；此时尚未扩展新增交易日列。
    sheet_configs = build_sheet_configs()
    allowed_sheets_for_append = {cfg["sheet_name"] for cfg in sheet_configs}

    board_members = fetch_latest_board_members()
    missing_df = detect_missing_bonds(original_data, board_members, baseline_sheet="总表")
    existing_codes = set(original_data["总表"].index.astype(str).tolist())
    print(
        f"[missing] 板块成分 {len(board_members)} 只 vs 底稿 {len(existing_codes)} 只，缺失 {len(missing_df)} 只"
        "。"
    )

    new_bond_codes = []
    new_bond_issue_dates = {}
    if not missing_df.empty:
        notify_missing_bonds(missing_df)
        original_data = append_missing_bonds_to_sheets(
            original_data,
            missing_df,
            skip_sheets={"指数"},
            allowed_sheets=allowed_sheets_for_append,
        )
        new_bond_codes = [str(c) for c in missing_df.index.tolist()]
        # 从板块成分提取"网上发行日期"，用于 fill 时按券裁剪历史日期范围，减少无效抓取
        # 缺失项 fill_sheets_with_ifind 会再从总表"发行日期"列兜底刷新
        new_bond_issue_dates = extract_issue_dates_from_missing(missing_df)

    total_sheet_configs = [cfg for cfg in sheet_configs if cfg.get("rule_type") == "total_table_formula"]
    daily_sheet_configs = [cfg for cfg in sheet_configs if cfg.get("rule_type") != "total_table_formula"]
    historical_backfill_audit = {}

    # 3.0) 总表静态字段独立更新一次，为后续历史修补提供发行日、上市日和最后交易日边界。
    original_data = fill_sheets_with_ifind(
        original_data,
        total_sheet_configs,
        trade_dates=api_dates,
        new_bond_codes=new_bond_codes,
        new_bond_issue_dates=new_bond_issue_dates,
        fetch_scope="new_bond_history",
        history_cutoff=history_cutoff,
    )
    # 通用公式对异常退市券可能错误回退到到期日；人工核实日期拥有最终优先级。
    original_data = apply_manual_last_trade_date_overrides(original_data)

    # 3.1) 第一阶段：只为新券批量补齐 history_cutoff 及以前的已有历史列。
    if new_bond_codes:
        original_data = fill_sheets_with_ifind(
            original_data,
            daily_sheet_configs,
            trade_dates=api_dates,
            new_bond_codes=new_bond_codes,
            new_bond_issue_dates=new_bond_issue_dates,
            fetch_scope="new_bond_history",
            history_cutoff=history_cutoff,
            historical_backfill_audit=historical_backfill_audit,
        )
        original_data = sanitize_zero_values(
            original_data,
            row_codes=new_bond_codes,
            date_on_or_before=history_cutoff,
            include_non_date_columns=True,
            scope_label="新券历史新增结果",
        )
        print(f"[update] 新券历史补档完成，截止 {history_cutoff.strftime('%Y-%m-%d')}。")

    # 3.2) 可选阶段：只修补 history_cutoff 及以前、各指标有效区间内的历史局部缺失。
    if RUN_HISTORICAL_MISSING_REPAIR:
        original_data = fill_sheets_with_ifind(
            original_data,
            daily_sheet_configs,
            trade_dates=api_dates,
            fetch_scope="historical_missing",
            history_cutoff=history_cutoff,
            skip_total_table=True,
            historical_backfill_audit=historical_backfill_audit,
        )
    else:
        print("[历史缺失修补] 开关已关闭，本次跳过。")

    # 3.3) 第三阶段：历史修补完成后再扩列，抓取 history_cutoff 及之后的缺失值。
    original_data = extend_data_columns_to_latest(
        original_data,
        trade_dates,
        skip_sheets={"总表"},
    )
    original_data = fill_sheets_with_ifind(
        original_data,
        daily_sheet_configs,
        trade_dates=api_dates,
        fetch_scope="incremental_dates",
        history_cutoff=history_cutoff,
        skip_total_table=True,
    )
    # 仅清理本阶段涉及的截止日及新增日期列，避免再次扫描、复制全部历史数据。
    original_data = sanitize_zero_values(
        original_data,
        date_on_or_after=history_cutoff,
        include_non_date_columns=False,
        scope_label="截止日及新增日期结果",
    )

    # 赎回累计天数由「【条款】P强赎进度跟踪.py」使用原始 Parquet
    # 行情、正股交易状态和公告 Excel 独立全历史重算并回写。
    # 本流程只更新底层原始数据，不再执行赎回累计天数测算。
    print("⏭️ 赎回累计天数由【条款】P强赎进度跟踪独立维护；底稿更新不再执行该项测算。")

    # 下修累计天数由“下修进度跟踪_本地化.py”独立全历史重算并写回。
    # 本流程只扩展底层交易日并保留既有结果，新增日期在下修脚本运行前保持为空。
    print("⏭️ 下修累计天数由下修进度跟踪独立维护；底稿更新不再执行该项测算。")

    # 落盘前统一清除转债生命周期外的个券日度数据。正股相关指标保留完整历史；
    # 赎回和下修累计天数都由各自的本地化脚本维护，不参与本次统一清洗。
    original_data = sanitize_bond_lifecycle_values(original_data)

    # 最终落盘前再执行一次全表零值清理，覆盖派生及跳过更新的历史 sheet。
    original_data = sanitize_zero_values(original_data)
    notify_all_empty_date_items(
        original_data,
        date_on_or_after=history_cutoff,
    )
    print("✅ iFinD 数据填充完成。")

    # 4) 保存更新后的 parquet（可选）。
    save_updated_parquet = True  # 设置为 True 则导出 parquet，False 则跳过导出
    if save_updated_parquet:
        export_original_data_to_parquet(original_data, output_root=parquet_root)
        # 紧随 parquet 导出同步刷新磁盘缓存；直接复用内存中的 original_data，
        # 避免回读整个目录造成的额外 IO 与类型漂移。
        build_parquet_cache(parquet_root, data=original_data)
        notify_historical_backfill(historical_backfill_audit)
    else:
        print("[save] 已跳过导出 parquet（save_updated_parquet=False）。")

if __name__ == "__main__":
    main()
