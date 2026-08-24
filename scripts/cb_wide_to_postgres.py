#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Dict, List, Tuple


def _require_psycopg() -> "module":
    try:
        import psycopg  # type: ignore
    except Exception as e:  # pragma: no cover
        raise SystemExit(
            "Missing dependency: psycopg.\n"
            "Install (user scope): py -3.12 -m pip install psycopg[binary]\n"
            f"Original error: {e}"
        )
    return psycopg


def _read_registry(export_root: Path) -> Dict[str, Dict[str, str]]:
    p = export_root / "meta" / "metric_registry.json"
    if not p.exists():
        raise SystemExit(f"Not found: {p} (run cb_parquet_wide_export.py first)")
    return json.loads(p.read_text(encoding="utf-8"))


def _read_sql(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _quote_ident(ident: str) -> str:
    return '"' + ident.replace('"', '""') + '"'


def _table_ref(schema: str, table: str) -> str:
    return f"{_quote_ident(schema)}.{_quote_ident(table)}"


def _csv_header(csv_path: Path) -> List[str]:
    # Read first line only; files are utf-8-sig
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        header = f.readline().strip("\n\r")
    if not header:
        raise SystemExit(f"Empty CSV header: {csv_path}")
    return header.split(",")


def _run_ddl(cur, ddl_sql: str) -> None:
    # 跳过 CREATE SCHEMA（public 等已有 schema；避免无建库权限的账号报错）
    lines = [
        ln
        for ln in ddl_sql.splitlines()
        if not ln.strip().upper().startswith("CREATE SCHEMA")
    ]
    ddl_sql = "\n".join(lines).strip()
    if not ddl_sql:
        raise SystemExit("DDL is empty after removing CREATE SCHEMA statements.")
    cur.execute(ddl_sql)


def _copy_csv_to_staging(cur, staging_ref: str, csv_path: Path) -> None:
    cols = _csv_header(csv_path)
    cols_quoted = ", ".join(_quote_ident(c) for c in cols)
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        copy_sql = f"COPY {staging_ref} ({cols_quoted}) FROM STDIN WITH (FORMAT csv, HEADER true)"
        with cur.copy(copy_sql) as cp:  # psycopg3
            while True:
                chunk = f.read(1024 * 1024)
                if not chunk:
                    break
                cp.write(chunk)


def _upsert_from_staging(cur, target_ref: str, staging_ref: str, cols: List[str]) -> None:
    if "date" not in cols:
        raise SystemExit("CSV must include a 'date' column as first field.")
    non_pk = [c for c in cols if c != "date"]
    cols_quoted = ", ".join(_quote_ident(c) for c in cols)
    set_clause = ", ".join(f"{_quote_ident(c)} = EXCLUDED.{_quote_ident(c)}" for c in non_pk)
    cur.execute(
        f"INSERT INTO {target_ref} ({cols_quoted}) "
        f"SELECT {cols_quoted} FROM {staging_ref} "
        f"ON CONFLICT ({_quote_ident('date')}) DO UPDATE SET {set_clause}"
    )


def _insert_ignore_from_staging(cur, target_ref: str, staging_ref: str, cols: List[str]) -> None:
    cols_quoted = ", ".join(_quote_ident(c) for c in cols)
    cur.execute(
        f"INSERT INTO {target_ref} ({cols_quoted}) "
        f"SELECT {cols_quoted} FROM {staging_ref} "
        f"ON CONFLICT ({_quote_ident('date')}) DO NOTHING"
    )


def load_metrics_to_postgres(
    *,
    export_root: Path,
    dsn: str,
    schema: str,
    metrics: List[str] | None,
    mode: str,
) -> None:
    psycopg = _require_psycopg()

    registry = _read_registry(export_root)
    ddl_dir = export_root / "ddl"
    data_dir = export_root / "data"

    selected: List[Tuple[str, Dict[str, str]]] = []
    for metric, meta in registry.items():
        if metrics is None or metric in metrics:
            selected.append((metric, meta))
    if not selected:
        raise SystemExit("No metrics selected (check --metrics or export output).")

    with psycopg.connect(dsn, autocommit=False) as conn:
        with conn.cursor() as cur:
            for metric, meta in selected:
                table_name = meta["table_name"]  # original sheet name (may be Chinese)
                data_stem = meta["data_stem"]

                ddl_path = ddl_dir / f"{data_stem}.sql"
                csv_path = data_dir / f"{data_stem}.csv"
                if not ddl_path.exists():
                    raise SystemExit(f"Missing DDL file: {ddl_path}")
                if not csv_path.exists():
                    raise SystemExit(f"Missing CSV file: {csv_path} (re-run export without --no-csv)")

                ddl_sql = _read_sql(ddl_path)
                _run_ddl(cur, ddl_sql)

                target_ref = _table_ref(schema, table_name)
                staging_ref = _table_ref(schema, f"__stg_{data_stem}")

                # staging table mirrors target
                cur.execute(f"DROP TABLE IF EXISTS {staging_ref};")
                cur.execute(f"CREATE UNLOGGED TABLE {staging_ref} (LIKE {target_ref} INCLUDING DEFAULTS);")

                _copy_csv_to_staging(cur, staging_ref, csv_path)
                cols = _csv_header(csv_path)

                if mode == "replace":
                    cur.execute(f"TRUNCATE TABLE {target_ref};")
                    cur.execute(f"INSERT INTO {target_ref} SELECT * FROM {staging_ref};")
                elif mode == "append":
                    _insert_ignore_from_staging(cur, target_ref, staging_ref, cols)
                elif mode == "upsert":
                    _upsert_from_staging(cur, target_ref, staging_ref, cols)
                else:
                    raise SystemExit(f"Unknown mode: {mode}")

                cur.execute(f"DROP TABLE IF EXISTS {staging_ref};")
                conn.commit()


def main() -> None:
    p = argparse.ArgumentParser(description="Load exported wide-table CSVs into PostgreSQL directly (no DBeaver needed).")
    p.add_argument(
        "--export-root",
        type=Path,
        default=Path(r"D:\JupyterFiles\huachuang\cb_wide_export"),
        help="Export root generated by cb_parquet_wide_export.py",
    )
    p.add_argument(
        "--dsn",
        type=str,
        default="",
        help="PostgreSQL DSN, e.g. postgresql://user:pass@host:5432/dbname (recommend using env var PG_DSN)",
    )
    p.add_argument("--schema", type=str, default="public", help="Target schema (must match export --schema)")
    p.add_argument(
        "--metrics",
        type=str,
        default="",
        help="Comma-separated metric(sheet) names to load (default: all in registry). Example: 收盘价,转股溢价率",
    )
    p.add_argument(
        "--mode",
        type=str,
        default="upsert",
        choices=["replace", "append", "upsert"],
        help="replace=truncate+reload, append=only new dates, upsert=insert/update by date",
    )
    args = p.parse_args()

    dsn = args.dsn.strip()
    if not dsn:
        import os

        dsn = os.environ.get("PG_DSN", "").strip()
    if not dsn:
        raise SystemExit("Missing --dsn and env var PG_DSN. Refuse to proceed without explicit connection string.")

    metrics = [m.strip() for m in args.metrics.split(",") if m.strip()] or None

    load_metrics_to_postgres(
        export_root=args.export_root,
        dsn=dsn,
        schema=args.schema,
        metrics=metrics,
        mode=args.mode,
    )


if __name__ == "__main__":
    main()

