#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
一键：把 2015-2026 parquet 分片合并为“宽表”（行=日期，列=转债），并直接写入云端 PostgreSQL。

用法（不依赖 PowerShell）：
  py -3.12 D:\\JupyterFiles\\huachuang\\scripts\\upload_cb_wide_onefile.py

首次需要安装依赖：
  py -3.12 -m pip install pyarrow pandas psycopg[binary] tqdm

安全：不把密码写在文件里；入库前会弹出密码输入框（明文显示，便于核对）。
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

# 直接运行本文件时，Python 只会把 scripts/ 放进 sys.path，需补上项目根目录
_PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))


@dataclass(frozen=True)
class Settings:
    # 1) 本地 parquet 根目录
    input_root: Path = Path(r"D:\JupyterFiles\huachuang\转债个券历史序列")
    start_year: int = 2015
    end_year: int = 2026

    # 2) 导出中间文件目录（会生成 data/ ddl/ meta/）
    export_root: Path = Path(r"D:\JupyterFiles\huachuang\cb_wide_export")

    # 3) PostgreSQL 目标库
    pg_host: str = "pgm-2zecgh9960s16bvsuo.pg.rds.aliyuncs.com"
    pg_port: int = 5432
    pg_db: str = "zgn_db"
    pg_user: str = "hcgszgn"

    # 4) 目标 schema（public 为默认 schema，不执行 CREATE SCHEMA）
    target_schema: str = "public"

    # 5) 导入哪些指标；None=全部（非常多，建议先填几个测试）
    metrics: Optional[List[str]] = field(default_factory=lambda: ["收盘价"])

    # 6) 导入策略：upsert / append / replace
    mode: str = "upsert"


def prompt_pg_password(*, user: str, host: str, port: int, db: str) -> str:
    """弹出 GUI 密码框（明文显示）。无图形界面时回退到终端 input。"""
    target = f"{user}@{host}:{port}/{db}"
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        pwd = input(f"PostgreSQL 密码 ({target}): ")
        if not pwd:
            raise SystemExit("未输入密码，已取消。")
        return pwd

    result: list[str | None] = [None]

    root = tk.Tk()
    root.title("PostgreSQL 密码")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    frame = ttk.Frame(root, padding=12)
    frame.grid(row=0, column=0)

    ttk.Label(frame, text="请输入数据库密码（可见）：").grid(row=0, column=0, sticky="w")
    ttk.Label(frame, text=target).grid(row=1, column=0, sticky="w", pady=(0, 8))

    pwd_var = tk.StringVar()
    entry = ttk.Entry(frame, textvariable=pwd_var, width=42, show="")
    entry.grid(row=2, column=0, sticky="ew")
    entry.focus_set()

    btn_row = ttk.Frame(frame)
    btn_row.grid(row=3, column=0, pady=(10, 0), sticky="e")

    def on_ok() -> None:
        result[0] = pwd_var.get()
        root.destroy()

    def on_cancel() -> None:
        root.destroy()

    ttk.Button(btn_row, text="确定", command=on_ok).grid(row=0, column=0, padx=(0, 6))
    ttk.Button(btn_row, text="取消", command=on_cancel).grid(row=0, column=1)
    entry.bind("<Return>", lambda _event: on_ok())
    root.protocol("WM_DELETE_WINDOW", on_cancel)

    root.update_idletasks()
    w, h = root.winfo_reqwidth(), root.winfo_reqheight()
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 3
    root.geometry(f"+{x}+{y}")

    root.mainloop()

    if not result[0]:
        raise SystemExit("未输入密码，已取消。")
    return result[0]


def main() -> None:
    s = Settings()

    # 先导出宽表（CSV + DDL）
    from scripts.cb_parquet_wide_export import export_wide_tables

    export_wide_tables(
        input_root=s.input_root,
        output_root=s.export_root,
        start_year=s.start_year,
        end_year=s.end_year,
        schema_name=s.target_schema,
        metrics=s.metrics,
        write_csv=True,
        write_parquet=False,
    )

    # 再写入 PostgreSQL（会执行 DDL + COPY + upsert/append/replace）
    from scripts.cb_wide_to_postgres import load_metrics_to_postgres

    password = prompt_pg_password(
        user=s.pg_user, host=s.pg_host, port=s.pg_port, db=s.pg_db
    )
    dsn = f"postgresql://{s.pg_user}:{password}@{s.pg_host}:{s.pg_port}/{s.pg_db}"

    load_metrics_to_postgres(
        export_root=s.export_root,
        dsn=dsn,
        schema=s.target_schema,
        metrics=s.metrics,
        mode=s.mode,
    )

    print("Done.")


if __name__ == "__main__":
    main()

