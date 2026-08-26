"""可转债历史数据的标准 Parquet Schema 与公共读写接口。

月度文件按“转债代码＋交易日期”保存一条观测；公共读写函数继续兼容日更程序
使用的 ``dict[sheet_name, DataFrame]`` 宽表接口。
"""

from __future__ import annotations

import json
import os
import re
import tempfile
import time
from pathlib import Path
from typing import Iterable, Mapping

import numpy as np
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq
from tqdm import tqdm


SCHEMA_VERSION = "2.15.0"
BOND_CODE = "转债代码"
TRADE_DATE = "交易日期"
INDEX_NAME = "指数名称"
INDEX_VALUE = "指数值"
MASTER_SHEET = "总表"
INDEX_SHEET = "指数"
TQDM_NCOLS = 92
WINDOWS_REPLACE_RETRY_DELAYS = (0.25, 0.5, 1.0, 2.0, 4.0, 4.0, 4.0)

MONTHLY_METRICS = [
    "余额",
    "收盘价",
    "平价",
    "转股价",
    "转股溢价率",
    "纯债价值",
    "纯债溢价率",
    "平价底价溢价率",
    "YTM",
    "换手率",
    "成交量",
    "成交额",
    "涨跌幅",
    "剩余期限",
    "隐含波动率",
    "主体评级",
    "债项评级",
    "正股收盘价",
    "正股交易状态",
    "正股近1日均价",
    "正股近20日均价",
    "正股20日波动率",
    "累计转股比例",
    "转股稀释率",
    "正股市值",
    "每股净资产",
    "EXPMA5",
    "EXPMA10",
    "EXPMA20",
    "交易状态",
    "赎回累计天数",
    "下修累计天数",
]

DEPRECATED_METRICS = {
    "不强赎承诺期倒计时",
    "不下修承诺期倒计时",
    "转债20天波动率",
    "近两年正股波动率",
}

STRING_METRICS = {"主体评级", "债项评级", "交易状态", "正股交易状态"}
NUMERIC_METRICS = [m for m in MONTHLY_METRICS if m not in STRING_METRICS]

MASTER_DATE_COLUMNS = [
    "上市日期",
    "最后交易日",
    "最后转股日",
    "摘牌日期",
    "到期日期",
    "发行日期",
    "赎回公告日",
    "转股期起始日",
    "回售起始日期",
]
MASTER_FLOAT_COLUMNS = [
    "发行规模",
    "到期赎回价",
    "股票发行面值",
    "赎回触发比例",
    "下修触发比例",
]
MASTER_INTEGER_COLUMNS = [
    "赎回触发计算时间区间",
    "赎回触发计算最大时间区间",
    "重设触发计算时间区间",
    "重设触发计算最大时间区间",
]

NULL_TEXT = {"", "--", "-", "N/A", "NA", "nan", "None", "null"}
DATE_COLUMN_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _normalise_missing(series: pd.Series) -> pd.Series:
    result = series.copy()
    if pd.api.types.is_object_dtype(result.dtype) or pd.api.types.is_string_dtype(result.dtype):
        result = result.map(lambda value: value.strip() if isinstance(value, str) else value)
        result = result.mask(result.isin(NULL_TEXT))
    return result


def _coerce_numeric_strict(series: pd.Series, context: str) -> pd.Series:
    source = _normalise_missing(series)
    result = pd.to_numeric(source, errors="coerce")
    invalid = source.notna() & result.isna()
    if invalid.any():
        examples = source.loc[invalid].astype(str).drop_duplicates().head(10).tolist()
        raise ValueError(f"{context} 含有无法解析为数值的内容: {examples}")
    return result.astype("float64")


def _coerce_integer_strict(series: pd.Series, context: str) -> pd.Series:
    numeric = _coerce_numeric_strict(series, context)
    non_integral = numeric.notna() & ~np.isclose(numeric % 1, 0.0, rtol=0.0, atol=1e-12)
    if non_integral.any():
        examples = numeric.loc[non_integral].drop_duplicates().head(10).tolist()
        raise ValueError(f"{context} 应为整数，发现: {examples}")
    return numeric.round().astype("Int16")


def _coerce_date_strict(series: pd.Series, context: str) -> pd.Series:
    source = _normalise_missing(series)
    try:
        parsed = pd.to_datetime(source, errors="coerce", format="mixed")
    except TypeError:
        parsed = pd.to_datetime(source, errors="coerce")
    invalid = source.notna() & parsed.isna()
    if invalid.any():
        examples = source.loc[invalid].astype(str).drop_duplicates().head(10).tolist()
        raise ValueError(f"{context} 含有无法解析为日期的内容: {examples}")
    return parsed.dt.normalize()


def _string_series(series: pd.Series) -> pd.Series:
    return _normalise_missing(series).astype("string")


def _schema_metadata(dataset_type: str, primary_key: Iterable[str]) -> dict[bytes, bytes]:
    return {
        b"schema_version": SCHEMA_VERSION.encode(),
        b"dataset_type": dataset_type.encode(),
        b"primary_key": ",".join(primary_key).encode("utf-8"),
    }


def monthly_schema() -> pa.Schema:
    fields = [pa.field(BOND_CODE, pa.string(), nullable=False), pa.field(TRADE_DATE, pa.date32(), nullable=False)]
    fields.extend(pa.field(metric, pa.string() if metric in STRING_METRICS else pa.float64()) for metric in MONTHLY_METRICS)
    return pa.schema(fields, metadata=_schema_metadata("convertible_bond_monthly_panel", [BOND_CODE, TRADE_DATE]))


def index_schema() -> pa.Schema:
    return pa.schema(
        [
            pa.field(INDEX_NAME, pa.string(), nullable=False),
            pa.field(TRADE_DATE, pa.date32(), nullable=False),
            pa.field(INDEX_VALUE, pa.float64()),
        ],
        metadata=_schema_metadata("market_index_history", [INDEX_NAME, TRADE_DATE]),
    )


def _master_schema(columns: Iterable[str]) -> pa.Schema:
    fields = []
    for column in columns:
        if column == BOND_CODE:
            field = pa.field(column, pa.string(), nullable=False)
        elif column in MASTER_DATE_COLUMNS:
            field = pa.field(column, pa.date32())
        elif column in MASTER_FLOAT_COLUMNS:
            field = pa.field(column, pa.float64())
        elif column in MASTER_INTEGER_COLUMNS:
            field = pa.field(column, pa.int16())
        else:
            field = pa.field(column, pa.string())
        fields.append(field)
    return pa.schema(fields, metadata=_schema_metadata("convertible_bond_master", [BOND_CODE]))


def _table_from_frame(frame: pd.DataFrame, schema: pa.Schema) -> pa.Table:
    arrays = [pa.array(frame[field.name], type=field.type, from_pandas=True, safe=True) for field in schema]
    return pa.Table.from_arrays(arrays, schema=schema)


def _replace_file_with_retry(temp_path: Path, path: Path) -> None:
    """Replace a file atomically, tolerating short-lived Windows file locks."""

    waited = 0.0
    for attempt, delay in enumerate((*WINDOWS_REPLACE_RETRY_DELAYS, None), start=1):
        try:
            os.replace(temp_path, path)
            return
        except OSError as exc:
            transient_lock = (
                isinstance(exc, PermissionError)
                or getattr(exc, "winerror", None) in {5, 32}
            )
            if not transient_lock or delay is None:
                if transient_lock:
                    raise PermissionError(
                        f"{path}: 文件持续被其他进程占用，等待 {waited:.2f} 秒、"
                        f"重试 {attempt - 1} 次后仍无法替换。请暂停 Git/LFS 同步、"
                        "关闭占用该文件的 Python/Excel 进程后重试。"
                    ) from exc
                raise
            time.sleep(delay)
            waited += delay


def write_typed_parquet(frame: pd.DataFrame, path: str | Path, schema: pa.Schema) -> None:
    """Write a validated Parquet file atomically beside its final path."""

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if list(frame.columns) != schema.names:
        raise ValueError(f"{path}: 字段不符合 Schema\n实际: {list(frame.columns)}\n预期: {schema.names}")
    table = _table_from_frame(frame, schema)
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        pq.write_table(table, temp_path, compression="zstd", use_dictionary=True, row_group_size=128_000)
        # Reading through an explicit file object avoids a lingering Windows
        # memory-mapped handle that would block the following atomic replace.
        with temp_path.open("rb") as handle:
            reread = pq.read_table(handle)
        if reread.schema != schema or reread.num_rows != len(frame):
            raise RuntimeError(f"{path}: Parquet 回读校验失败")
        del reread
        _replace_file_with_retry(temp_path, path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def _legacy_date_columns(frame: pd.DataFrame) -> list[str]:
    columns = [str(column) for column in frame.columns if DATE_COLUMN_RE.fullmatch(str(column))]
    for column in columns:
        pd.Timestamp(column)
    return columns


def _long_metric(group: pd.DataFrame, date_columns: list[str], metric: str) -> pd.Series:
    sub = group[["__row_id", *date_columns]].drop_duplicates("__row_id", keep="last")
    long = sub.melt(id_vars="__row_id", value_vars=date_columns, var_name=TRADE_DATE, value_name=metric)
    long[metric] = _normalise_missing(long[metric])
    long = long.dropna(subset=[metric])
    long[BOND_CODE] = long["__row_id"].astype("string")
    long[TRADE_DATE] = pd.to_datetime(long[TRADE_DATE]).dt.normalize()
    return long.set_index([BOND_CODE, TRADE_DATE])[metric]


def legacy_month_to_standard(frame: pd.DataFrame, source: str = "") -> tuple[pd.DataFrame, pd.DataFrame]:
    """Convert one legacy monthly frame to the v2 bond panel and index table."""

    required = {"__sheet_name", "__row_id"}
    if not required.issubset(frame.columns):
        raise ValueError(f"{source}: 不是旧版月度 Schema")
    date_columns = _legacy_date_columns(frame)
    if not date_columns:
        raise ValueError(f"{source}: 未找到交易日期列")

    groups = {str(name): group for name, group in frame.groupby("__sheet_name", sort=False)}
    unknown = sorted(set(groups) - set(MONTHLY_METRICS) - DEPRECATED_METRICS - {INDEX_SHEET})
    if unknown:
        raise ValueError(f"{source}: 发现未登记指标 {unknown}")

    parts = []
    for metric in MONTHLY_METRICS:
        group = groups.get(metric)
        if group is not None:
            parts.append(_long_metric(group, date_columns, metric))
    panel = pd.concat(parts, axis=1).reset_index() if parts else pd.DataFrame(columns=[BOND_CODE, TRADE_DATE])
    if panel.duplicated([BOND_CODE, TRADE_DATE]).any():
        raise ValueError(f"{source}: 转债代码＋交易日期主键重复")
    panel[BOND_CODE] = _string_series(panel[BOND_CODE])
    for metric in MONTHLY_METRICS:
        if metric not in panel:
            panel[metric] = pd.NA
        if metric in STRING_METRICS:
            panel[metric] = _string_series(panel[metric])
        else:
            panel[metric] = _coerce_numeric_strict(panel[metric], f"{source}/{metric}")
    panel = panel[[BOND_CODE, TRADE_DATE, *MONTHLY_METRICS]].sort_values([TRADE_DATE, BOND_CODE], kind="stable").reset_index(drop=True)

    index_group = groups.get(INDEX_SHEET)
    if index_group is None:
        indices = pd.DataFrame(columns=[INDEX_NAME, TRADE_DATE, INDEX_VALUE])
    else:
        long = index_group[["__row_id", *date_columns]].drop_duplicates("__row_id", keep="last").melt(
            id_vars="__row_id", value_vars=date_columns, var_name=TRADE_DATE, value_name=INDEX_VALUE
        )
        long[INDEX_VALUE] = _coerce_numeric_strict(long[INDEX_VALUE], f"{source}/{INDEX_SHEET}")
        indices = long.dropna(subset=[INDEX_VALUE]).rename(columns={"__row_id": INDEX_NAME})
        indices[INDEX_NAME] = _string_series(indices[INDEX_NAME])
        indices[TRADE_DATE] = pd.to_datetime(indices[TRADE_DATE]).dt.normalize()
        indices = indices[[INDEX_NAME, TRADE_DATE, INDEX_VALUE]].sort_values([TRADE_DATE, INDEX_NAME], kind="stable").reset_index(drop=True)
    if indices.duplicated([INDEX_NAME, TRADE_DATE]).any():
        raise ValueError(f"{source}: 指数名称＋交易日期主键重复")
    return panel, indices


def standardize_master(frame: pd.DataFrame, source: str = "总表") -> pd.DataFrame:
    result = frame.copy()
    if "__row_id" in result:
        result = result.rename(columns={"__row_id": BOND_CODE})
    if "__sheet_name" in result:
        result = result.drop(columns="__sheet_name")
    if BOND_CODE not in result:
        if result.index.name is not None or not isinstance(result.index, pd.RangeIndex):
            result.insert(0, BOND_CODE, result.index.astype(str))
        else:
            raise ValueError(f"{source}: 缺少 {BOND_CODE}")
    if result[BOND_CODE].isna().any() or result.duplicated(BOND_CODE).any():
        raise ValueError(f"{source}: {BOND_CODE} 存在空值或重复")
    result[BOND_CODE] = _string_series(result[BOND_CODE])
    for column in result.columns:
        if column == BOND_CODE:
            continue
        if column in MASTER_DATE_COLUMNS:
            result[column] = _coerce_date_strict(result[column], f"{source}/{column}")
        elif column in MASTER_FLOAT_COLUMNS:
            result[column] = _coerce_numeric_strict(result[column], f"{source}/{column}")
        elif column in MASTER_INTEGER_COLUMNS:
            result[column] = _coerce_integer_strict(result[column], f"{source}/{column}")
        else:
            result[column] = _string_series(result[column])
    return result.reset_index(drop=True)


def _original_metric_to_long(frame: pd.DataFrame, metric: str) -> pd.DataFrame:
    data = frame.copy()
    data.index = data.index.astype(str)
    parsed = pd.to_datetime(pd.Index(data.columns), errors="coerce")
    keep = ~parsed.isna()
    data = data.loc[:, keep]
    data.columns = parsed[keep]
    long = data.rename_axis(BOND_CODE).reset_index().melt(id_vars=BOND_CODE, var_name=TRADE_DATE, value_name=metric)
    long[TRADE_DATE] = pd.to_datetime(long[TRADE_DATE])
    long[metric] = _normalise_missing(long[metric])
    return long.dropna(subset=[metric])


def _build_month_from_original(original_data: Mapping[str, pd.DataFrame], year: int, month: int) -> pd.DataFrame:
    series = []
    for metric in MONTHLY_METRICS:
        frame = original_data.get(metric)
        if frame is None:
            continue
        long = _original_metric_to_long(frame, metric)
        long = long.loc[(long[TRADE_DATE].dt.year == year) & (long[TRADE_DATE].dt.month == month)]
        if not long.empty:
            series.append(long.set_index([BOND_CODE, TRADE_DATE])[metric])
    panel = pd.concat(series, axis=1).reset_index() if series else pd.DataFrame(columns=[BOND_CODE, TRADE_DATE])
    panel[BOND_CODE] = _string_series(panel[BOND_CODE])
    for metric in MONTHLY_METRICS:
        if metric not in panel:
            panel[metric] = pd.NA
        panel[metric] = _string_series(panel[metric]) if metric in STRING_METRICS else _coerce_numeric_strict(panel[metric], metric)
    return panel[[BOND_CODE, TRADE_DATE, *MONTHLY_METRICS]].sort_values([TRADE_DATE, BOND_CODE], kind="stable").reset_index(drop=True)


def _build_all_months_from_original(
    original_data: Mapping[str, pd.DataFrame],
    progress: tqdm | None = None,
    stop_check=None,
) -> dict[tuple[int, int], pd.DataFrame]:
    """Build every monthly panel while stacking each source metric only once."""

    parts_by_month: dict[tuple[int, int], list[pd.Series]] = {}
    for metric in MONTHLY_METRICS:
        if stop_check is not None:
            stop_check()
        if progress is not None:
            progress.set_postfix_str(f"整理指标：{metric}", refresh=True)
        try:
            frame = original_data.get(metric)
            if frame is None:
                continue
            long = _original_metric_to_long(frame, metric)
            if long.empty:
                continue
            for key, group in long.groupby([long[TRADE_DATE].dt.year, long[TRADE_DATE].dt.month], sort=True):
                month_key = (int(key[0]), int(key[1]))
                parts_by_month.setdefault(month_key, []).append(group.set_index([BOND_CODE, TRADE_DATE])[metric])
        finally:
            if progress is not None:
                progress.update(1)

    result = {}
    if progress is not None:
        progress.total += len(parts_by_month)
        progress.refresh()
    for key, parts in parts_by_month.items():
        if stop_check is not None:
            stop_check()
        if progress is not None:
            progress.set_postfix_str(f"重建月份：{key[0]}-{key[1]:02d}", refresh=True)
        panel = pd.concat(parts, axis=1).reset_index()
        if panel.duplicated([BOND_CODE, TRADE_DATE]).any():
            raise ValueError(f"{key}: 转债代码＋交易日期主键重复")
        panel[BOND_CODE] = _string_series(panel[BOND_CODE])
        for metric in MONTHLY_METRICS:
            if metric not in panel:
                panel[metric] = pd.NA
            panel[metric] = _string_series(panel[metric]) if metric in STRING_METRICS else _coerce_numeric_strict(panel[metric], metric)
        result[key] = panel[[BOND_CODE, TRADE_DATE, *MONTHLY_METRICS]].sort_values(
            [TRADE_DATE, BOND_CODE], kind="stable"
        ).reset_index(drop=True)
        if progress is not None:
            progress.update(1)
    return result


def _all_months(original_data: Mapping[str, pd.DataFrame]) -> list[tuple[int, int]]:
    months = set()
    for metric in [*MONTHLY_METRICS, INDEX_SHEET]:
        frame = original_data.get(metric)
        if frame is None:
            continue
        parsed = pd.to_datetime(pd.Index(frame.columns), errors="coerce")
        months.update((int(value.year), int(value.month)) for value in parsed if pd.notna(value))
    return sorted(months)


def _build_index_from_original(original_data: Mapping[str, pd.DataFrame]) -> pd.DataFrame:
    frame = original_data.get(INDEX_SHEET)
    if frame is None:
        return pd.DataFrame(columns=[INDEX_NAME, TRADE_DATE, INDEX_VALUE])
    long = _original_metric_to_long(frame, INDEX_VALUE).rename(columns={BOND_CODE: INDEX_NAME})
    long[INDEX_NAME] = _string_series(long[INDEX_NAME])
    long[INDEX_VALUE] = _coerce_numeric_strict(long[INDEX_VALUE], INDEX_SHEET)
    return long[[INDEX_NAME, TRADE_DATE, INDEX_VALUE]].sort_values([TRADE_DATE, INDEX_NAME], kind="stable").reset_index(drop=True)


def _write_manifest(output_root: Path, rows: list[dict[str, object]]) -> None:
    manifest = pd.DataFrame(rows, columns=["sheet_name", "route", "rows", "cols"])
    manifest["sheet_name"] = manifest["sheet_name"].astype("string")
    manifest["route"] = manifest["route"].astype("string")
    manifest["rows"] = manifest["rows"].astype("int64")
    manifest["cols"] = manifest["cols"].astype("int64")
    schema = pa.schema(
        [pa.field("sheet_name", pa.string()), pa.field("route", pa.string()), pa.field("rows", pa.int64()), pa.field("cols", pa.int64())],
        metadata=_schema_metadata("dataset_manifest", ["sheet_name"]),
    )
    write_typed_parquet(manifest, output_root / "_meta" / "sheet_manifest.parquet", schema)


def export_original_data_to_parquet(
    original_data: Mapping[str, pd.DataFrame],
    output_root: str | Path = "data/转债个券历史序列",
    stop_check=None,
) -> None:
    """Write the in-memory legacy workbook model using the standard v2 files."""

    output_root = Path(output_root).resolve()
    print("[parquet] 开始重建并写入标准 Parquet。", flush=True)
    progress = tqdm(
        total=len(MONTHLY_METRICS),
        desc="Rebuilding parquet",
        unit="step",
        ncols=TQDM_NCOLS,
        dynamic_ncols=False,
        mininterval=0.2,
    )
    try:
        monthly_panels = _build_all_months_from_original(
            original_data,
            progress=progress,
            stop_check=stop_check,
        )
        # 写盘开始后不再响应停止，避免只更新部分月份造成数据集不一致。
        if stop_check is not None:
            stop_check()
        # 月度面板数量在完成指标整理后才能确定；动态补入“写月文件＋3个收尾步骤”。
        progress.total += len(monthly_panels) + 3
        progress.refresh()

        manifest_rows = []
        for (year, month), panel in sorted(monthly_panels.items()):
            progress.set_postfix_str(f"写入月份：{year}-{month:02d}", refresh=True)
            path = output_root / str(year) / f"{year}{month:02d}.parquet"
            write_typed_parquet(panel, path, monthly_schema())
            progress.update(1)
        for metric in MONTHLY_METRICS:
            frame = original_data.get(metric)
            observations = int(frame.notna().sum().sum()) if frame is not None else 0
            manifest_rows.append({"sheet_name": metric, "route": f"monthly_column:{metric}", "rows": observations, "cols": 1})

        progress.set_postfix_str("写入指数", refresh=True)
        index_frame = _build_index_from_original(original_data)
        write_typed_parquet(index_frame, output_root / "_special" / "指数.parquet", index_schema())
        manifest_rows.append({"sheet_name": INDEX_SHEET, "route": "_special/指数.parquet", "rows": len(index_frame), "cols": 3})
        progress.update(1)

        progress.set_postfix_str("写入总表", refresh=True)
        master_source = original_data.get(MASTER_SHEET)
        if master_source is None:
            raise ValueError("original_data 缺少总表")
        master = standardize_master(master_source)
        write_typed_parquet(master, output_root / "_special" / "总表.parquet", _master_schema(master.columns))
        manifest_rows.append({"sheet_name": MASTER_SHEET, "route": "_special/总表.parquet", "rows": len(master), "cols": len(master.columns)})
        progress.update(1)

        progress.set_postfix_str("写入清单", refresh=True)
        _write_manifest(output_root, manifest_rows)
        progress.update(1)
    finally:
        progress.close()
    print(f"[parquet] 标准 Parquet 重建完成：{output_root}", flush=True)


def read_original_data_from_parquet(input_root: str | Path = "data/转债个券历史序列") -> dict[str, pd.DataFrame]:
    """Read v2 files and reconstruct the legacy workbook-shaped dictionary."""

    input_root = Path(input_root)
    files = sorted(path for year in input_root.iterdir() if year.is_dir() and year.name.isdigit() for path in year.glob("*.parquet"))
    parts: dict[str, list[pd.DataFrame]] = {metric: [] for metric in MONTHLY_METRICS}
    all_dates = []
    for path in files:
        frame = pd.read_parquet(path)
        required = {BOND_CODE, TRADE_DATE, *MONTHLY_METRICS}
        if not required.issubset(frame.columns):
            raise ValueError(f"{path} 不是标准 v2 月度 Schema")
        frame[TRADE_DATE] = pd.to_datetime(frame[TRADE_DATE])
        all_dates.extend(frame[TRADE_DATE].drop_duplicates().tolist())
        for metric in MONTHLY_METRICS:
            wide = frame.pivot(index=BOND_CODE, columns=TRADE_DATE, values=metric)
            parts[metric].append(wide)

    master = pd.read_parquet(input_root / "_special" / "总表.parquet")
    master[BOND_CODE] = master[BOND_CODE].astype(str)
    code_order = master[BOND_CODE].tolist()
    original: dict[str, pd.DataFrame] = {}
    date_order = pd.DatetimeIndex(sorted(pd.Index(all_dates).unique()))
    for metric, metric_parts in parts.items():
        merged = pd.concat(metric_parts, axis=1) if metric_parts else pd.DataFrame()
        merged = merged.reindex(index=code_order, columns=date_order)
        merged.index.name = None
        original[metric] = merged

    index_frame = pd.read_parquet(input_root / "_special" / "指数.parquet")
    index_frame[TRADE_DATE] = pd.to_datetime(index_frame[TRADE_DATE])
    index_wide = index_frame.pivot(index=INDEX_NAME, columns=TRADE_DATE, values=INDEX_VALUE).reindex(columns=date_order)
    index_wide.index.name = None
    original[INDEX_SHEET] = index_wide

    master = master.set_index(BOND_CODE)
    master.index.name = None
    for column in MASTER_DATE_COLUMNS:
        if column in master:
            master[column] = pd.to_datetime(master[column])
    original[MASTER_SHEET] = master
    return original


def read_metric_wide(input_root: str | Path, metric: str) -> pd.DataFrame:
    """Read one metric as date-by-code matrix for analytical scripts."""

    root = Path(input_root)
    if metric == INDEX_SHEET:
        frame = pd.read_parquet(root / "_special" / "指数.parquet")
        frame[TRADE_DATE] = pd.to_datetime(frame[TRADE_DATE])
        return frame.pivot(index=TRADE_DATE, columns=INDEX_NAME, values=INDEX_VALUE).sort_index()
    if metric not in MONTHLY_METRICS:
        raise KeyError(f"未知指标: {metric}")
    parts = []
    for path in sorted(p for year in root.iterdir() if year.is_dir() and year.name.isdigit() for p in year.glob("*.parquet")):
        frame = pd.read_parquet(path, columns=[BOND_CODE, TRADE_DATE, metric])
        parts.append(frame)
    long = pd.concat(parts, ignore_index=True)
    long[TRADE_DATE] = pd.to_datetime(long[TRADE_DATE])
    return long.pivot(index=TRADE_DATE, columns=BOND_CODE, values=metric).sort_index()


def replace_monthly_metric_from_wide(
    input_root: str | Path,
    metric: str,
    wide: pd.DataFrame,
    progress: tqdm | None = None,
) -> int:
    """用完整宽表原子替换月度 Parquet 中的单个指标，并刷新清单计数。

    ``wide`` 必须以转债代码为行、交易日期为列。未出现在宽表或值为空的
    代码日期组合会写为空值；其他月度指标保持不变。
    """
    if metric not in MONTHLY_METRICS:
        raise KeyError(f"未知指标: {metric}")

    root = Path(input_root)
    data = wide.copy()
    data.index = _string_series(pd.Series(data.index, index=data.index)).to_numpy()
    data.index.name = BOND_CODE
    parsed_dates = pd.to_datetime(pd.Index(data.columns), errors="coerce")
    if parsed_dates.isna().any():
        bad = [str(col) for col, parsed in zip(data.columns, parsed_dates) if pd.isna(parsed)]
        raise ValueError(f"{metric} 宽表包含无法识别的日期列: {bad[:10]}")
    data.columns = pd.DatetimeIndex(parsed_dates).normalize()
    data.columns.name = TRADE_DATE
    if data.index.has_duplicates:
        raise ValueError(f"{metric} 宽表转债代码重复")
    if data.columns.has_duplicates:
        raise ValueError(f"{metric} 宽表交易日期重复")

    if metric in STRING_METRICS:
        data = data.apply(_string_series)
    else:
        data = data.apply(lambda series: _coerce_numeric_strict(series, metric))
    lookup = data.stack(future_stack=True).dropna()
    lookup.index = lookup.index.set_names([BOND_CODE, TRADE_DATE])

    observations = 0
    month_paths = sorted(
        path
        for year_dir in root.iterdir()
        if year_dir.is_dir() and year_dir.name.isdigit()
        for path in year_dir.glob("*.parquet")
    )
    if not month_paths:
        raise FileNotFoundError(f"未找到月度 Parquet: {root}")

    if progress is not None:
        progress.total = len(month_paths)
        progress.refresh()

    schema = monthly_schema()
    for path in month_paths:
        if progress is not None:
            progress.set_postfix_str(f"写入：{path.stem}", refresh=True)
        panel = pd.read_parquet(path)
        missing_columns = [column for column in schema.names if column not in panel.columns]
        if missing_columns:
            raise ValueError(f"{path}: 缺少标准字段 {missing_columns}")
        panel[BOND_CODE] = _string_series(panel[BOND_CODE])
        panel[TRADE_DATE] = pd.to_datetime(panel[TRADE_DATE]).dt.normalize()
        keys = pd.MultiIndex.from_frame(panel[[BOND_CODE, TRADE_DATE]])
        replacement = lookup.reindex(keys)
        panel[metric] = replacement.to_numpy()
        panel[metric] = (
            _string_series(panel[metric])
            if metric in STRING_METRICS
            else _coerce_numeric_strict(panel[metric], f"{path}/{metric}")
        )
        observations += int(panel[metric].notna().sum())
        write_typed_parquet(panel[schema.names], path, schema)
        if progress is not None:
            progress.update(1)

    manifest_path = root / "_meta" / "sheet_manifest.parquet"
    if manifest_path.is_file():
        manifest = pd.read_parquet(manifest_path)
        metric_mask = manifest["sheet_name"].astype(str).eq(metric)
        if metric_mask.any():
            manifest.loc[metric_mask, "rows"] = observations
            manifest.loc[metric_mask, "cols"] = 1
            manifest.loc[metric_mask, "route"] = f"monthly_column:{metric}"
        else:
            manifest = pd.concat(
                [
                    manifest,
                    pd.DataFrame(
                        [{
                            "sheet_name": metric,
                            "route": f"monthly_column:{metric}",
                            "rows": observations,
                            "cols": 1,
                        }]
                    ),
                ],
                ignore_index=True,
            )
        _write_manifest(root, manifest.to_dict("records"))

    return observations


def schema_document() -> dict[str, object]:
    return {
        "schema_version": SCHEMA_VERSION,
        "monthly": {"primary_key": [BOND_CODE, TRADE_DATE], "fields": {field.name: str(field.type) for field in monthly_schema()}},
        "master": {"primary_key": [BOND_CODE], "date_fields": MASTER_DATE_COLUMNS, "float_fields": MASTER_FLOAT_COLUMNS, "integer_fields": MASTER_INTEGER_COLUMNS},
        "index": {"primary_key": [INDEX_NAME, TRADE_DATE], "fields": {field.name: str(field.type) for field in index_schema()}},
    }


def write_schema_document(output_root: str | Path) -> None:
    path = Path(output_root) / "_meta" / "schema_v2.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(schema_document(), ensure_ascii=False, indent=2), encoding="utf-8")
