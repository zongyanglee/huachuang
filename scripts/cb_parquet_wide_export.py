#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Sequence, Tuple, TypeVar

import pandas as pd
import pyarrow.parquet as pq

T = TypeVar("T")

META_COLS = ("__sheet_name", "__row_id")


def _progress(
    items: Sequence[T] | Iterable[T],
    *,
    desc: str,
    unit: str,
    enabled: bool = True,
) -> Iterator[T]:
    """终端进度条；已安装 tqdm 时用 tqdm，否则用简易百分比输出。"""
    if not enabled:
        yield from items
        return

    seq: Sequence[T]
    if isinstance(items, Sequence) and not isinstance(items, (str, bytes)):
        seq = items
    else:
        seq = list(items)

    total = len(seq)
    if total == 0:
        return

    try:
        from tqdm import tqdm

        yield from tqdm(seq, desc=desc, unit=unit, total=total, dynamic_ncols=True)
        return
    except ImportError:
        pass

    width = max(len(desc), 20)
    for i, item in enumerate(seq, start=1):
        pct = i * 100 // total
        bar = "#" * (pct // 2) + "-" * (50 - pct // 2)
        print(f"\r{desc:<{width}} [{bar}] {i}/{total} ({pct}%)", end="", flush=True)
        yield item
    print()


def _iter_parquet_files(root: Path, start_year: int, end_year: int) -> List[Path]:
    files: List[Path] = []
    for year in range(start_year, end_year + 1):
        year_dir = root / str(year)
        if not year_dir.exists():
            continue
        files.extend(sorted(year_dir.glob("*.parquet")))
    return files


def _parse_date_col(col: str) -> pd.Timestamp | None:
    # parquet columns look like: "2015-01-05 00:00:00"
    try:
        return pd.to_datetime(col)
    except Exception:
        return None


_BOND_COL_RE = re.compile(r"[^0-9A-Za-z_]+")


def _bond_to_colname(bond_code: str) -> str:
    # Postgres identifier: keep ASCII, start with letter.
    # Example: "110009.SH" -> "b110009_SH"
    cleaned = _BOND_COL_RE.sub("_", bond_code.strip())
    cleaned = cleaned.strip("_")
    if not cleaned:
        cleaned = "unknown"
    if cleaned[0].isdigit():
        cleaned = f"b{cleaned}"
    else:
        cleaned = f"b_{cleaned}"
    return cleaned


def _infer_sql_type(series: pd.Series) -> str:
    # Heuristic: if mostly numeric after coercion -> double precision else text
    coerced = pd.to_numeric(series, errors="coerce")
    non_null = series.notna().sum()
    if non_null == 0:
        return "text"
    numeric_ratio = coerced.notna().sum() / non_null
    return "double precision" if numeric_ratio >= 0.8 else "text"


@dataclass
class MetricWide:
    sheet_name: str
    frames: List[pd.DataFrame]

    def add_frame(self, frame: pd.DataFrame) -> None:
        self.frames.append(frame)

    def finalize(self) -> pd.DataFrame:
        if not self.frames:
            return pd.DataFrame()
        wide = pd.concat(self.frames, axis=0)
        wide = wide[~wide.index.duplicated(keep="last")]
        wide = wide.sort_index()
        wide.index.name = "date"
        return wide


def export_wide_tables(
    input_root: Path,
    output_root: Path,
    start_year: int,
    end_year: int,
    schema_name: str,
    metrics: List[str] | None,
    write_csv: bool,
    write_parquet: bool,
    show_progress: bool = True,
) -> None:
    output_root.mkdir(parents=True, exist_ok=True)
    out_data_dir = output_root / "data"
    out_ddl_dir = output_root / "ddl"
    out_meta_dir = output_root / "meta"
    out_data_dir.mkdir(parents=True, exist_ok=True)
    out_ddl_dir.mkdir(parents=True, exist_ok=True)
    out_meta_dir.mkdir(parents=True, exist_ok=True)

    files = _iter_parquet_files(input_root, start_year, end_year)
    if not files:
        raise SystemExit(f"No parquet files found under: {input_root} (years {start_year}-{end_year})")

    if show_progress:
        print(f"共 {len(files)} 个 parquet 分片（{start_year}-{end_year}）")

    metric_map: Dict[str, MetricWide] = {}
    bond_colname_map: Dict[str, str] = {}

    for file in _progress(files, desc="读取分片", unit="个", enabled=show_progress):
        table = pq.read_table(str(file))
        df = table.to_pandas()

        if not all(col in df.columns for col in META_COLS):
            raise SystemExit(f"Missing required columns {META_COLS} in {file}")

        date_cols = [c for c in df.columns if c not in META_COLS]
        parsed_dates: List[Tuple[str, pd.Timestamp]] = []
        for c in date_cols:
            ts = _parse_date_col(c)
            if ts is not None:
                parsed_dates.append((c, ts))
        if not parsed_dates:
            raise SystemExit(f"No date-like columns found in {file}")

        # Keep date columns ordered
        parsed_dates.sort(key=lambda x: x[1])
        date_cols_ordered = [c for (c, _) in parsed_dates]

        if metrics is not None:
            df = df[df["__sheet_name"].isin(metrics)]

        for sheet_name, g in df.groupby("__sheet_name", sort=False):
            metric_wide = metric_map.get(sheet_name)
            if metric_wide is None:
                metric_wide = MetricWide(sheet_name=sheet_name, frames=[])
                metric_map[sheet_name] = metric_wide

            sub = g[["__row_id", *date_cols_ordered]].copy()
            # Normalize row_id to string; later becomes columns.
            sub["__row_id"] = sub["__row_id"].astype(str)

            for bond in sub["__row_id"].unique().tolist():
                if bond not in bond_colname_map:
                    bond_colname_map[bond] = _bond_to_colname(bond)

            sub = sub.set_index("__row_id")
            # transpose: rows=date, cols=bonds
            wide_part = sub.T
            wide_part.index = pd.to_datetime(wide_part.index)
            wide_part = wide_part.sort_index()

            # Convert values: keep original strings for now; we'll coerce per-metric later.
            metric_wide.add_frame(wide_part)

    # Persist mapping
    (out_meta_dir / "bond_column_mapping.json").write_text(
        json.dumps(bond_colname_map, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # Export per-metric
    registry: Dict[str, Dict[str, str]] = {}
    metric_items = sorted(metric_map.items(), key=lambda x: x[0])
    for sheet_name, metric_wide in _progress(
        metric_items, desc="写出宽表", unit="项", enabled=show_progress
    ):
        wide = metric_wide.finalize()
        if wide.empty:
            continue

        # Rename bond columns to safe SQL identifiers
        wide = wide.rename(columns=bond_colname_map)

        # Coerce columns to numeric where possible; we infer SQL types from the resulting columns
        sql_types: Dict[str, str] = {}
        coerced_wide = wide.copy()
        col_iter = _progress(
            list(coerced_wide.columns),
            desc=f"  {sheet_name}",
            unit="列",
            enabled=show_progress and len(coerced_wide.columns) > 100,
        )
        for col in col_iter:
            coerced = pd.to_numeric(coerced_wide[col], errors="coerce")
            # Keep numeric if it meaningfully converts; else keep original strings
            if coerced.notna().sum() >= max(1, int(coerced_wide[col].notna().sum() * 0.8)):
                coerced_wide[col] = coerced.astype("float64")
            sql_types[col] = _infer_sql_type(coerced_wide[col])

        safe_file_stem = re.sub(r"[\\\\/:*?\"<>|\\s]+", "_", sheet_name).strip("_") or "metric"
        registry[sheet_name] = {
            "table_schema": schema_name,
            "table_name": sheet_name,  # keep original; we'll quote it in DDL
            "data_stem": safe_file_stem,
        }

        if write_parquet:
            pq_path = out_data_dir / f"{safe_file_stem}.parquet"
            coerced_wide.reset_index().to_parquet(pq_path, index=False)

        if write_csv:
            csv_path = out_data_dir / f"{safe_file_stem}.csv"
            coerced_wide.reset_index().to_csv(csv_path, index=False, encoding="utf-8-sig")

        # DDL（public 等已有 schema 不建库级 schema）
        ddl_lines: List[str] = []
        if schema_name.lower() != "public":
            ddl_lines.append(f"CREATE SCHEMA IF NOT EXISTS {schema_name};")
        ddl_lines.append(f'CREATE TABLE IF NOT EXISTS {schema_name}."{sheet_name}" (')
        ddl_lines.append("  date date PRIMARY KEY,")
        cols = list(coerced_wide.columns)
        for i, col in enumerate(cols):
            comma = "," if i < len(cols) - 1 else ""
            ddl_lines.append(f'  "{col}" {sql_types[col]}{comma}')
        ddl_lines.append(");")
        (out_ddl_dir / f"{safe_file_stem}.sql").write_text("\n".join(ddl_lines) + "\n", encoding="utf-8")

    (out_meta_dir / "metric_registry.json").write_text(
        json.dumps(registry, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge monthly parquet shards (2015-2026) into per-metric wide tables: rows=date, cols=bonds."
    )
    parser.add_argument(
        "--input-root",
        type=Path,
        default=Path(r"D:\JupyterFiles\huachuang\转债个券历史序列"),
        help="Root directory containing year folders like 2015/201501.parquet",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=Path(r"D:\JupyterFiles\huachuang\cb_wide_export"),
        help="Output directory for data/ ddl/ meta/",
    )
    parser.add_argument("--start-year", type=int, default=2015)
    parser.add_argument("--end-year", type=int, default=2026)
    parser.add_argument("--schema", type=str, default="public")
    parser.add_argument(
        "--metrics",
        type=str,
        default="",
        help="Comma-separated sheet names to export (default: all). Example: 收盘价,转股溢价率",
    )
    parser.add_argument("--no-csv", action="store_true", help="Do not write CSV outputs")
    parser.add_argument("--no-parquet", action="store_true", help="Do not write Parquet outputs")
    parser.add_argument("--no-progress", action="store_true", help="Disable progress bars")

    args = parser.parse_args()
    metrics = [m.strip() for m in args.metrics.split(",") if m.strip()] or None

    export_wide_tables(
        input_root=args.input_root,
        output_root=args.output_root,
        start_year=args.start_year,
        end_year=args.end_year,
        schema_name=args.schema,
        metrics=metrics,
        write_csv=not args.no_csv,
        write_parquet=not args.no_parquet,
        show_progress=not args.no_progress,
    )


if __name__ == "__main__":
    main()
