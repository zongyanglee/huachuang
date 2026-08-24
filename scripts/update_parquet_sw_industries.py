from __future__ import annotations

import argparse
import json
import os
import shutil
import time
from pathlib import Path

import pandas as pd
import xlwings as xw


FIELDS = [
    ("申万二级行业", 2),
    ("申万三级行业", 3),
]
FETCHING_MARKERS = ("数据获取中", "正在获取", "loading", "fetching", "requesting")
ERROR_MARKERS = ("#NAME?", "#VALUE!", "#N/A", "#REF!", "#NUM!", "#DIV/0!")


def is_fetching(value) -> bool:
    if value is None:
        return False
    text = str(value).strip().lower()
    return any(marker.lower() in text for marker in FETCHING_MARKERS)


def normalize_result(value):
    if value is None:
        return pd.NA
    text = str(value).strip()
    if not text or text.lower() in {"none", "nan", "nat", "<na>"}:
        return pd.NA
    if any(marker in text.upper() for marker in ERROR_MARKERS):
        return pd.NA
    return text


def read_matrix(range_value, rows: int, cols: int) -> list[list]:
    if rows == 1 and cols == 1:
        return [[range_value]]
    if rows == 1:
        return [list(range_value)]
    if cols == 1:
        return [[value] for value in range_value]
    return [list(row) for row in range_value]


def fetch_industries(
    codes: list[str], max_wait_seconds: int = 180, poll_seconds: int = 3
) -> dict[str, list]:
    app = None
    workbook = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        workbook = app.books.add()
        sheet = workbook.sheets[0]
        sheet.name = "申万行业查询"
        sheet.range("A1:C1").value = [["转债代码", "申万二级行业", "申万三级行业"]]
        sheet.range("A:A").number_format = "@"
        sheet.range((2, 1), (len(codes) + 1, 1)).options(transpose=True).value = codes

        formulas = []
        for row in range(2, len(codes) + 2):
            formulas.append(
                [
                    f'=@s_info_industry_sw_2021(A{row},"",2)',
                    f'=@s_info_industry_sw_2021(A{row},"",3)',
                ]
            )
        formula_range = sheet.range((2, 2), (len(codes) + 1, 3))
        formula_range.formula = formulas

        started = time.monotonic()
        latest_matrix = []
        while True:
            app.calculate()
            latest_matrix = read_matrix(formula_range.value, len(codes), 2)
            flattened = [value for row in latest_matrix for value in row]
            fetching = sum(is_fetching(value) for value in flattened)
            nonempty = sum(pd.notna(normalize_result(value)) for value in flattened)
            elapsed = int(time.monotonic() - started)
            print(
                f"[excel] elapsed={elapsed}s fetching={fetching} "
                f"nonempty={nonempty}/{len(flattened)}",
                flush=True,
            )
            if fetching == 0 and nonempty > 0:
                break
            if elapsed >= max_wait_seconds:
                raise TimeoutError(
                    f"Excel取数超时：fetching={fetching}, nonempty={nonempty}/{len(flattened)}"
                )
            time.sleep(poll_seconds)

        results = {field: [] for field, _ in FIELDS}
        for row in latest_matrix:
            for column, (field, _) in enumerate(FIELDS):
                results[field].append(normalize_result(row[column]))
        return results
    finally:
        if workbook is not None:
            workbook.close()
        if app is not None:
            app.quit()


def audit_results(codes: list[str], results: dict[str, list]) -> dict:
    audit = {"codes": len(codes), "fields": {}}
    for field, _ in FIELDS:
        values = pd.Series(results[field], dtype="string")
        missing_mask = values.isna() | values.str.strip().eq("")
        audit["fields"][field] = {
            "nonempty": int((~missing_mask).sum()),
            "missing": int(missing_mask.sum()),
            "unique": int(values.dropna().nunique()),
            "missing_codes": [codes[index] for index in values.index[missing_mask]][:30],
        }
    return audit


def update_total_parquet(root: Path, results: dict[str, list], backup_dir: Path) -> None:
    total_path = root / "_special" / "总表.parquet"
    manifest_path = root / "_meta" / "sheet_manifest.parquet"
    total = pd.read_parquet(total_path)
    if len(total) != len(results[FIELDS[0][0]]):
        raise RuntimeError("Excel结果行数与总表不一致")

    for field, _ in reversed(FIELDS):
        if field in total.columns:
            total = total.drop(columns=[field])
    industry_position = list(total.columns).index("申万行业") + 1
    for offset, (field, _) in enumerate(FIELDS):
        total.insert(
            industry_position + offset,
            field,
            pd.Series(results[field], index=total.index, dtype="string"),
        )

    backup_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(total_path, backup_dir / "总表_更新前.parquet")
    shutil.copy2(manifest_path, backup_dir / "sheet_manifest_更新前.parquet")

    temp_total = total_path.with_name(total_path.name + ".sw_tmp")
    try:
        total.to_parquet(temp_total, index=False, engine="pyarrow", compression="snappy")
        reread = pd.read_parquet(temp_total)
        if len(reread) != len(total) or list(reread.columns) != list(total.columns):
            raise RuntimeError("总表临时文件结构校验失败")
        for field, _ in FIELDS:
            if int(reread[field].notna().sum()) != int(total[field].notna().sum()):
                raise RuntimeError(f"总表临时文件{field}非空数校验失败")
        os.replace(temp_total, total_path)
    finally:
        if temp_total.exists():
            temp_total.unlink()

    manifest = pd.read_parquet(manifest_path)
    mask = manifest["sheet_name"] == "总表"
    if int(mask.sum()) != 1:
        raise RuntimeError("元数据中无法唯一定位总表")
    manifest.loc[mask, "cols"] = len(total.columns) - 2
    temp_manifest = manifest_path.with_name(manifest_path.name + ".sw_tmp")
    try:
        manifest.to_parquet(
            temp_manifest, index=False, engine="pyarrow", compression="snappy"
        )
        reread_manifest = pd.read_parquet(temp_manifest)
        recorded = int(
            reread_manifest.loc[reread_manifest["sheet_name"] == "总表", "cols"].iloc[0]
        )
        if recorded != len(total.columns) - 2:
            raise RuntimeError("总表元数据列数校验失败")
        os.replace(temp_manifest, manifest_path)
    finally:
        if temp_manifest.exists():
            temp_manifest.unlink()


def main() -> None:
    parser = argparse.ArgumentParser(description="用iFinD Excel函数补充申万二级、三级行业")
    parser.add_argument("parquet_root", type=Path)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--max-wait-seconds", type=int, default=180)
    parser.add_argument("--backup-dir", type=Path, default=None)
    args = parser.parse_args()

    root = args.parquet_root.resolve()
    total_path = root / "_special" / "总表.parquet"
    total = pd.read_parquet(total_path)
    codes = total["__row_id"].astype(str).tolist()
    if args.limit is not None:
        codes = codes[: args.limit]

    results = fetch_industries(codes, max_wait_seconds=args.max_wait_seconds)
    audit = audit_results(codes, results)
    print(json.dumps(audit, ensure_ascii=False, indent=2), flush=True)
    for index in range(min(10, len(codes))):
        print(
            codes[index],
            results["申万二级行业"][index],
            results["申万三级行业"][index],
            flush=True,
        )

    if not args.apply:
        print("仅测试Excel函数，未修改parquet。", flush=True)
        return
    if args.limit is not None:
        raise RuntimeError("--limit 与 --apply 不可同时使用")

    minimum_coverage = int(len(codes) * 0.98)
    for field, _ in FIELDS:
        nonempty = audit["fields"][field]["nonempty"]
        if nonempty < minimum_coverage:
            raise RuntimeError(f"{field}覆盖不足：{nonempty}/{len(codes)}")

    backup_dir = (
        args.backup_dir.resolve()
        if args.backup_dir
        else root.parent / "tmp" / "sw_industry_update_20260810"
    )
    update_total_parquet(root, results, backup_dir)
    print("总表parquet与元数据已更新。", flush=True)


if __name__ == "__main__":
    main()
