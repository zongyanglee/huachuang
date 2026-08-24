from pathlib import Path
import json

import pandas as pd


BALANCE = "\u4f59\u989d"
DEBT_RATING = "\u503a\u9879\u8bc4\u7ea7"
SUBJECT_RATING = "\u4e3b\u4f53\u8bc4\u7ea7"


def main() -> None:
    files = sorted(
        p
        for p in Path(".").rglob("*.parquet")
        if p.parent.name.isdigit() and p.stem.isdigit()
    )

    records = []
    for path in files:
        df = pd.read_parquet(path)
        date_cols = [c for c in df.columns if c not in ["__sheet_name", "__row_id"]]

        balance = (
            df[df["__sheet_name"].eq(BALANCE)]
            .set_index("__row_id")[date_cols]
            .apply(pd.to_numeric, errors="coerce")
        )
        debt_rating = df[df["__sheet_name"].eq(DEBT_RATING)].set_index("__row_id")[
            date_cols
        ]
        subject_rating = df[df["__sheet_name"].eq(SUBJECT_RATING)].set_index(
            "__row_id"
        )[date_cols]

        common = balance.index.intersection(debt_rating.index).intersection(
            subject_rating.index
        )
        balance = balance.loc[common]
        debt_rating = debt_rating.loc[common]
        subject_rating = subject_rating.loc[common]

        for col in date_cols:
            balances = balance[col]
            active = (balances.notna() & (balances > 0)).fillna(False)
            debt_aaa = active & (
                debt_rating[col].astype("string").str.strip().eq("AAA").fillna(False)
            )
            subject_aaa = active & (
                subject_rating[col]
                .astype("string")
                .str.strip()
                .eq("AAA")
                .fillna(False)
            )

            total_count = int(active.sum())
            total_balance = float(balances[active].sum())
            debt_count = int(debt_aaa.sum())
            debt_balance = float(balances[debt_aaa].sum())
            subject_count = int(subject_aaa.sum())
            subject_balance = float(balances[subject_aaa].sum())

            records.append(
                {
                    "\u65e5\u671f": pd.to_datetime(col).strftime("%Y-%m-%d"),
                    "\u503a\u9879AAA\u4e2a\u6570": debt_count,
                    "\u503a\u9879AAA\u4f59\u989d(\u4ebf\u5143)": round(
                        debt_balance, 4
                    ),
                    "\u5168\u5e02\u573a\u4e2a\u6570": total_count,
                    "\u5168\u5e02\u573a\u4f59\u989d(\u4ebf\u5143)": round(
                        total_balance, 4
                    ),
                    "\u503a\u9879AAA\u6570\u91cf\u5360\u6bd4": (
                        debt_count / total_count if total_count else None
                    ),
                    "\u503a\u9879AAA\u4f59\u989d\u5360\u6bd4": (
                        debt_balance / total_balance if total_balance else None
                    ),
                    "\u4e3b\u4f53AAA\u4e2a\u6570": subject_count,
                    "\u4e3b\u4f53AAA\u4f59\u989d(\u4ebf\u5143)": round(
                        subject_balance, 4
                    ),
                    "\u4e3b\u4f53AAA\u6570\u91cf\u5360\u6bd4": (
                        subject_count / total_count if total_count else None
                    ),
                    "\u4e3b\u4f53AAA\u4f59\u989d\u5360\u6bd4": (
                        subject_balance / total_balance if total_balance else None
                    ),
                }
            )

    out_dir = Path("outputs") / "aaa_daily_excel"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "aaa_daily_data.json"
    out_path.write_text(
        json.dumps(
            {
                "source_file_count": len(files),
                "row_count": len(records),
                "date_start": records[0]["\u65e5\u671f"] if records else None,
                "date_end": records[-1]["\u65e5\u671f"] if records else None,
                "records": records,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(out_path)
    print(f"rows={len(records)} files={len(files)}")


if __name__ == "__main__":
    main()
