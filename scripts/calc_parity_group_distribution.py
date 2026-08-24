import json
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
OUTPUT_DIR = ROOT / "outputs" / "parity_group_distribution"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

PARITY_SHEET_NAME = "\u5e73\u4ef7"
EDGES = [50, 70, 90, 110, 130, 150, 170, 190, 200]
LABELS = ["<=50"] + [f"({EDGES[i]},{EDGES[i + 1]}]" for i in range(len(EDGES) - 1)] + [">200"]


def dedupe_sort(rows):
    by_date = {row["date"]: row for row in rows}
    return [by_date[date] for date in sorted(by_date)]


def main():
    files = sorted(
        path
        for path in ROOT.glob("**/*.parquet")
        if path.parent.name.isdigit() and path.stem.isdigit()
    )

    ratio_rows = []
    count_rows = []
    summary_rows = []

    for path in files:
        df = pd.read_parquet(path)
        if "__sheet_name" not in df.columns:
            continue

        parity = df[df["__sheet_name"].astype(str).eq(PARITY_SHEET_NAME)].copy()
        if parity.empty:
            continue

        date_cols = [col for col in parity.columns if col not in ["__sheet_name", "__row_id"]]
        for col in date_cols:
            values = pd.to_numeric(parity[col], errors="coerce").dropna()
            total = int(values.shape[0])
            counts = [0] * len(LABELS)
            if total:
                counts[0] = int((values <= 50).sum())
                for idx in range(len(EDGES) - 1):
                    counts[idx + 1] = int(((values > EDGES[idx]) & (values <= EDGES[idx + 1])).sum())
                counts[-1] = int((values > 200).sum())

            ratio_values = [count / total if total else None for count in counts]
            date = pd.to_datetime(str(col)).strftime("%Y-%m-%d")
            classified = int(sum(counts))

            ratio_rows.append(
                {
                    "date": date,
                    "values": dict(zip(LABELS, ratio_values)),
                    "classified_ratio": classified / total if total else None,
                }
            )
            count_rows.append(
                {
                    "date": date,
                    "values": dict(zip(LABELS, counts)),
                    "classified_total": classified,
                    "total_valid": total,
                }
            )
            summary_rows.append(
                {
                    "date": date,
                    "total_valid": total,
                    "classified_total": classified,
                    "classified_ratio": classified / total if total else None,
                }
            )

    payload = {
        "source_files": [str(path.relative_to(ROOT)) for path in files],
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        "bin_rule": "include <=50 and >200; middle bins are left-open-right-closed; denominator is valid parity sample count per date.",
        "edges": EDGES,
        "labels": LABELS,
        "ratio_rows": dedupe_sort(ratio_rows),
        "count_rows": dedupe_sort(count_rows),
        "summary_rows": dedupe_sort(summary_rows),
    }

    json_path = OUTPUT_DIR / "parity_group_distribution.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json_path)
    print("files", len(files), "dates", len(payload["ratio_rows"]))
    if payload["summary_rows"]:
        print(payload["ratio_rows"][0]["date"], payload["ratio_rows"][-1]["date"])
        print(payload["summary_rows"][-1])


if __name__ == "__main__":
    main()
