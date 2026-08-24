# -*- coding: utf-8 -*-
"""Compare notebook vs script daily outputs (ignore naming / extra sheets)."""
from __future__ import annotations

import hashlib
import json
import os
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]

# 宽表底稿 sheet（统计 xlsx 中 notebook 可能多写，py 不写）
RAW_WIDE_SHEETS = frozenset({
    "YTM", "余额", "债项评级", "剩余期限", "发行规模", "平价", "平价底价溢价率",
    "总表", "换手率", "收盘价", "正股市值", "纯债价值", "纯债溢价率", "转股溢价率",
    "隐含波动率", "平底分类转股溢价率", "次新券相对上市首日涨跌幅", "次新券转股溢价率",
})

SHEET_ALIASES = {
    "JS偏债型YTM分位数统计": "收盘价分位数统计",
    "收盘价分位数统计": "收盘价分位数统计",
}

FILE_PATTERNS = [
    (re.compile(r"^\d{4}数据更新\.xlsx$"), "raw_data"),
    (re.compile(r"^\d{4}数据更新（清理后）统计\.xlsx$"), "stats_clean"),
    (re.compile(r"^\d{4}数据更新（清理后剔妖）统计\.xlsx$"), "stats_yao"),
    (re.compile(r"^\d{4}百元平价溢价率拟合结果\.xlsx$"), "fit"),
    (re.compile(r"^\d{4}JS更新结果\.xlsx$"), "js"),
    (re.compile(r"^\d{4}剔除妖债及行业均值\.xlsx$"), "industry"),
    (re.compile(r"^\d{4}可转债列表\.xlsx$"), "cb_list"),
    (re.compile(r"^\d{4}转债周报\.txt$"), "weekly_txt"),
]


def _canonical_file(rel: str) -> str | None:
    name = Path(rel).name
    for pat, key in FILE_PATTERNS:
        if pat.match(name):
            return key
    sub = Path(rel).parts
    parts_set = set(sub)
    if parts_set & {"日内估值数据更新",} or (len(sub) >= 2 and "日内估值" in sub[0]):
        jpg = Path(rel).name
        m = re.search(r"【华创固收】(.+)\.jpe?g$", jpg, re.I)
        if not m:
            m = re.search(r"\d{4}-\d{2}-\d{2}【华创固收】(.+)\.jpe?g$", jpg, re.I)
        if m:
            return f"img:{m.group(1)}"
        if "日内数据更新" in jpg:
            return "img:intra_xlsx"
        if "百元平价溢价率拟合结果" in jpg:
            return "img:fit_xlsx"
    return None


def list_canonical(root: Path) -> dict[str, Path]:
    out: dict[str, Path] = {}
    for dp, _, fns in os.walk(root):
        for f in fns:
            if f.startswith("~$"):
                continue
            p = Path(dp) / f
            rel = p.relative_to(root).as_posix()
            key = _canonical_file(rel)
            if key:
                out[key] = p
    return out


def _norm_sheet(name: str) -> str:
    return SHEET_ALIASES.get(name, name)


def load_indexed(p: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(p, sheet_name=sheet, index_col=0)
    df.index = pd.Index(df.index).map(lambda x: str(x)[:10] if pd.notna(x) else "")
    return df


def compare_excel(pa: Path, pb: Path, *, rtol=1e-9, atol=1e-4) -> list[dict]:
    xa, xb = pd.ExcelFile(pa), pd.ExcelFile(pb)
    sa = {_norm_sheet(s): s for s in xa.sheet_names}
    sb = {_norm_sheet(s): s for s in xb.sheet_names}
    only_a = sorted(set(sa) - set(sb) - RAW_WIDE_SHEETS)
    only_b = sorted(set(sb) - set(sa) - RAW_WIDE_SHEETS)
    diffs: list[dict] = []
    if only_a:
        diffs.append({"type": "only_a_sheets", "sheets": only_a})
    if only_b:
        diffs.append({"type": "only_b_sheets", "sheets": only_b})
    for norm in sorted(set(sa) & set(sb)):
        if norm in RAW_WIDE_SHEETS:
            continue
        da = load_indexed(pa, sa[norm])
        db = load_indexed(pb, sb[norm])
        idx = sorted(set(da.index) & set(db.index))
        if not idx:
            diffs.append({"type": "no_common_dates", "sheet": norm, "a": list(da.index), "b": list(db.index)})
            continue
        da, db = da.loc[idx], db.loc[idx]
        if da.shape != db.shape:
            diffs.append({"type": "shape", "sheet": norm, "a": da.shape, "b": db.shape})
            continue
        d = da.apply(pd.to_numeric, errors="coerce") - db.apply(pd.to_numeric, errors="coerce")
        ad = d.abs()
        n = int((ad > atol).sum().sum())
        if n:
            diffs.append({
                "type": "values",
                "sheet": norm,
                "n_diff": n,
                "max_abs": float(np.nanmax(ad.values)),
                "dates": idx,
            })
    return diffs


def compare_txt(pa: Path, pb: Path) -> list[dict]:
    t1 = pa.read_text(encoding="utf-8", errors="replace").splitlines()
    t2 = pb.read_text(encoding="utf-8", errors="replace").splitlines()
    if t1 == t2:
        return []
    n = min(len(t1), len(t2))
    first = next((i for i in range(n) if t1[i] != t2[i]), None)
    return [{"lines_a": len(t1), "lines_b": len(t2), "first_diff_line": first}]


def file_hash(p: Path) -> str:
    h = hashlib.md5()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def compare_dirs(dir_a: Path, dir_b: Path) -> dict:
    fa, fb = list_canonical(dir_a), list_canonical(dir_b)
    only_a = sorted(set(fa) - set(fb))
    only_b = sorted(set(fb) - set(fa))
    common = sorted(set(fa) & set(fb))
    report: dict = {"only_a": only_a, "only_b": only_b, "pairs": {}}
    for key in common:
        pa, pb = fa[key], fb[key]
        entry: dict = {"path_a": str(pa), "path_b": str(pb)}
        ext = pa.suffix.lower()
        if ext in (".xlsx", ".xls"):
            entry["excel_diffs"] = compare_excel(pa, pb)
            entry["match"] = len(entry["excel_diffs"]) == 0
        elif ext == ".txt":
            entry["txt_diffs"] = compare_txt(pa, pb)
            entry["match"] = len(entry["txt_diffs"]) == 0
        elif ext in (".jpg", ".jpeg", ".png"):
            entry["match"] = file_hash(pa) == file_hash(pb)
        else:
            entry["match"] = file_hash(pa) == file_hash(pb)
        report["pairs"][key] = entry
    mism = [k for k, v in report["pairs"].items() if not v.get("match")]
    report["mismatch_keys"] = mism
    report["mismatch_count"] = len(mism) + len(only_a) + len(only_b)
    return report


def main() -> int:
    dir_a = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "JUP结果"
    dir_b = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "PY结果"
    report = compare_dirs(dir_a, dir_b)
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))
    return 1 if report["mismatch_count"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
