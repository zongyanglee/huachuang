import argparse
import os
from pathlib import Path

import pandas as pd
import pyarrow.parquet as pq


def _to_number(series: pd.Series) -> pd.Series:
    # parquet里是string；可能含逗号/空格/NA
    s = series.astype("string")
    s = s.str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce")


def load_industry_map(root: Path) -> pd.DataFrame:
    info_path = root / "data/转债个券历史序列" / "_special" / "总表.parquet"
    if not info_path.exists():
        raise FileNotFoundError(f"未找到总表文件: {info_path}")

    info = pq.read_table(str(info_path)).to_pandas()
    if "__row_id" not in info.columns or "申万行业" not in info.columns:
        raise ValueError("总表.parquet缺少必要列：__row_id 或 申万行业")

    m = info[["__row_id", "申万行业"]].copy()
    m["__row_id"] = m["__row_id"].astype("string")
    m["申万行业"] = m["申万行业"].astype("string")
    m = m.dropna(subset=["__row_id"]).drop_duplicates(subset=["__row_id"])
    return m


def iter_monthly_parquets(root: Path) -> list[Path]:
    base = root / "data/转债个券历史序列"
    if not base.exists():
        raise FileNotFoundError(f"未找到目录: {base}")

    files: list[Path] = []
    for p in base.rglob("*.parquet"):
        # 跳过元数据和总表（总表单独读取）
        if p.name == "sheet_manifest.parquet":
            continue
        if p.name == "总表.parquet":
            continue
        files.append(p)

    # 排序保证可复现（按路径字符串即可）
    files.sort(key=lambda x: str(x))
    return files


def aggregate_one_file(path: Path, industry_map: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    t = pq.read_table(str(path))
    df = t.to_pandas()

    if "__sheet_name" not in df.columns or "__row_id" not in df.columns:
        raise ValueError(f"{path} 缺少 __sheet_name 或 __row_id")

    df = df[df["__sheet_name"] == sheet_name].copy()
    if df.empty:
        return pd.DataFrame(columns=["date", "申万行业", "count", "balance_sum"])

    date_cols = [c for c in df.columns if c not in ("__sheet_name", "__row_id")]
    if not date_cols:
        return pd.DataFrame(columns=["date", "申万行业", "count", "balance_sum"])

    # wide -> long
    long_df = df.melt(
        id_vars="__row_id",
        value_vars=date_cols,
        var_name="date",
        value_name="balance",
    )

    long_df["date"] = pd.to_datetime(long_df["date"], errors="coerce").dt.date
    long_df["balance"] = _to_number(long_df["balance"])
    long_df = long_df.dropna(subset=["date", "__row_id"])

    long_df = long_df.merge(industry_map, on="__row_id", how="left")
    long_df["申万行业"] = long_df["申万行业"].fillna("未知")

    # 只统计有余额的转债；余额<=0的按缺失处理
    long_df = long_df[long_df["balance"].notna() & (long_df["balance"] > 0)]
    if long_df.empty:
        return pd.DataFrame(columns=["date", "申万行业", "count", "balance_sum"])

    out = (
        long_df.groupby(["date", "申万行业"], as_index=False)
        .agg(count=("__row_id", "nunique"), balance_sum=("balance", "sum"))
        .copy()
    )
    return out


def weighted_term_one_file(path: Path, sheet_balance: str, sheet_term: str) -> pd.DataFrame:
    t = pq.read_table(str(path))
    df = t.to_pandas()

    if "__sheet_name" not in df.columns or "__row_id" not in df.columns:
        raise ValueError(f"{path} 缺少 __sheet_name 或 __row_id")

    date_cols = [c for c in df.columns if c not in ("__sheet_name", "__row_id")]
    if not date_cols:
        return pd.DataFrame(columns=["date", "w_sum", "bal_sum"])

    bal = df[df["__sheet_name"] == sheet_balance].copy()
    term = df[df["__sheet_name"] == sheet_term].copy()
    if bal.empty or term.empty:
        return pd.DataFrame(columns=["date", "w_sum", "bal_sum"])

    bal_long = bal.melt(
        id_vars="__row_id",
        value_vars=date_cols,
        var_name="date",
        value_name="balance",
    )
    term_long = term.melt(
        id_vars="__row_id",
        value_vars=date_cols,
        var_name="date",
        value_name="remaining_term",
    )

    bal_long["date"] = pd.to_datetime(bal_long["date"], errors="coerce").dt.date
    term_long["date"] = pd.to_datetime(term_long["date"], errors="coerce").dt.date
    bal_long["balance"] = _to_number(bal_long["balance"])
    term_long["remaining_term"] = _to_number(term_long["remaining_term"])

    merged = bal_long.merge(term_long, on=["__row_id", "date"], how="inner")
    merged = merged.dropna(subset=["date", "__row_id"])
    merged = merged[merged["balance"].notna() & (merged["balance"] > 0) & merged["remaining_term"].notna()]
    if merged.empty:
        return pd.DataFrame(columns=["date", "w_sum", "bal_sum"])

    merged["w"] = merged["balance"] * merged["remaining_term"]
    out = (
        merged.groupby("date", as_index=False)
        .agg(w_sum=("w", "sum"), bal_sum=("balance", "sum"))
        .copy()
    )
    return out[["date", "w_sum", "bal_sum"]]


def main() -> None:
    parser = argparse.ArgumentParser(description="按日统计转债：申万行业 个数与总余额（从parquet读取）")
    parser.add_argument(
        "--root",
        default=str(Path.cwd()),
        help="工作目录（默认当前目录）",
    )
    parser.add_argument(
        "--sheet",
        default="余额",
        help="需要聚合的sheet_name（默认：余额）",
    )
    parser.add_argument(
        "--term-sheet",
        default="剩余期限",
        help="剩余期限sheet_name（默认：剩余期限）",
    )
    parser.add_argument(
        "--out",
        default="转债_申万行业_每日数量与余额.xlsx",
        help="输出Excel文件名（默认：转债_申万行业_每日数量与余额.xlsx）",
    )
    args = parser.parse_args()

    root = Path(args.root).resolve()
    industry_map = load_industry_map(root)
    files = iter_monthly_parquets(root)

    parts: list[pd.DataFrame] = []
    term_parts: list[pd.DataFrame] = []
    for i, fp in enumerate(files, 1):
        try:
            agg = aggregate_one_file(fp, industry_map, args.sheet)
            if not agg.empty:
                agg["source_file"] = fp.name
                parts.append(agg)

            term_agg = weighted_term_one_file(fp, sheet_balance=args.sheet, sheet_term=args.term_sheet)
            if not term_agg.empty:
                term_agg["source_file"] = fp.name
                term_parts.append(term_agg)
        except Exception as e:
            raise RuntimeError(f"处理失败: {fp}\n{e}") from e

        if i % 20 == 0:
            print(f"已处理 {i}/{len(files)}: {fp}")

    if not parts:
        raise RuntimeError("未聚合出任何数据：请确认parquet中存在sheet=余额且余额为数值。")

    result = pd.concat(parts, ignore_index=True)

    # 不同月文件不应有重复日期；若存在则合并一次
    result = (
        result.groupby(["date", "申万行业"], as_index=False)
        .agg(count=("count", "sum"), balance_sum=("balance_sum", "sum"))
        .sort_values(["date", "申万行业"], kind="stable")
    )

    daily_total = (
        result.groupby("date", as_index=False)
        .agg(count=("count", "sum"), balance_sum=("balance_sum", "sum"))
        .sort_values(["date"], kind="stable")
    )

    if term_parts:
        term_all = pd.concat(term_parts, ignore_index=True)
        daily_wavg_term = (
            term_all.groupby("date", as_index=False)
            .agg(w_sum=("w_sum", "sum"), bal_sum=("bal_sum", "sum"))
            .sort_values(["date"], kind="stable")
        )
        daily_wavg_term["wavg_remaining_term"] = daily_wavg_term["w_sum"] / daily_wavg_term["bal_sum"]
        daily_wavg_term = daily_wavg_term[["date", "wavg_remaining_term"]]
    else:
        daily_wavg_term = pd.DataFrame(columns=["date", "wavg_remaining_term"])

    # 透视：每行行业，每列日期（两张表：数量/余额）
    count_wide = (
        result.pivot_table(index="申万行业", columns="date", values="count", aggfunc="sum")
        .fillna(0)
        .astype("int64")
        .sort_index()
    )
    balance_wide = (
        result.pivot_table(index="申万行业", columns="date", values="balance_sum", aggfunc="sum")
        .fillna(0.0)
        .sort_index()
    )
    # 列名转成YYYY-MM-DD字符串，便于Excel查看/筛选
    count_wide.columns = [pd.to_datetime(c).date().isoformat() for c in count_wide.columns]
    balance_wide.columns = [pd.to_datetime(c).date().isoformat() for c in balance_wide.columns]

    out_path = root / args.out
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="daily_industry")
        daily_total.to_excel(writer, index=False, sheet_name="daily_total")
        count_wide.to_excel(writer, sheet_name="industry_x_date_count")
        balance_wide.to_excel(writer, sheet_name="industry_x_date_balance")
        daily_wavg_term.to_excel(writer, index=False, sheet_name="daily_wavg_remaining_term")

    print(f"已写出: {out_path}")


if __name__ == "__main__":
    # 避免某些环境下pandas的线程警告
    os.environ.setdefault("OMP_NUM_THREADS", "1")
    main()
