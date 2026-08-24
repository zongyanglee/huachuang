import html
import json
import pathlib
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RATING_ORDER = {
    "AAA": 1,
    "AA+": 2,
    "AA": 3,
    "AA-": 4,
    "A+": 5,
    "A": 6,
    "A-": 7,
    "BBB+": 8,
    "BBB": 9,
    "BBB-": 10,
    "BB+": 11,
    "BB": 12,
    "BB-": 13,
    "B+": 14,
    "B": 15,
    "B-": 16,
    "CCC": 17,
    "CC": 18,
    "C": 19,
    "D": 20,
}
RATING_PREFIXES = sorted(RATING_ORDER, key=len, reverse=True)


def load_shared_strings(archive):
    try:
        shared_xml = archive.read("xl/sharedStrings.xml").decode("utf-8", "ignore")
    except KeyError:
        return []
    root = ET.fromstring(shared_xml)
    shared = []
    for si in root.findall("main:si", NS):
        texts = [node.text or "" for node in si.findall(".//main:t", NS)]
        shared.append("".join(texts))
    return shared


def cell_text(cell, shared):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        texts = [node.text or "" for node in cell.findall(".//main:t", NS)]
        return "".join(texts)

    value = cell.find("main:v", NS)
    if value is None:
        return ""

    raw = value.text or ""
    if cell_type == "s":
        return shared[int(raw)]
    return html.unescape(raw)


def maybe_number(value):
    if value == "":
        return ""
    try:
        number = float(value)
    except ValueError:
        return value
    if number.is_integer():
        return int(number)
    return number


def clean_rating_text(value):
    return (
        str(value or "")
        .replace("_x000D_", "\n")
        .replace("_x000A_", "\n")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .strip()
    )


def parse_rating_line(line):
    line = line.strip()
    match = re.match(r"^(.+?)\((.+)\)$", line)
    if not match:
        return None

    rating = match.group(1).strip()
    detail = match.group(2).strip()
    date_match = re.match(r"^(.*)-(\d{8})$", detail)
    if not date_match:
        return None

    type_and_agency = date_match.group(1)
    raw_date = date_match.group(2)
    split_at = type_and_agency.find("-")
    if split_at >= 0:
        rating_type = type_and_agency[:split_at].strip()
        agency = type_and_agency[split_at + 1 :].strip()
    else:
        rating_type = type_and_agency.strip()
        agency = ""

    return {
        "rating": rating,
        "ratingType": rating_type,
        "agency": agency,
        "date": f"{raw_date[:4]}-{raw_date[4:6]}-{raw_date[6:]}",
        "dateKey": raw_date,
        "raw": line,
    }


def direction(previous_rating, current_rating):
    previous_rank = rating_rank(previous_rating)
    current_rank = rating_rank(current_rating)
    if previous_rank is None or current_rank is None:
        return "评级变更"
    if current_rank < previous_rank:
        return "评级上调"
    if current_rank > previous_rank:
        return "评级下调"
    if str(previous_rating) != str(current_rating):
        return "评级变更"
    return "评级不变"


def rating_rank(rating):
    for prefix in RATING_PREFIXES:
        if str(rating).startswith(prefix):
            return RATING_ORDER[prefix]
    return None


def read_rows(workbook_path):
    with zipfile.ZipFile(workbook_path) as archive:
        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8", "ignore")
        shared = load_shared_strings(archive)

    root = ET.fromstring(sheet_xml)
    rows = []
    for row in root.findall(".//main:sheetData/main:row", NS):
        record = {}
        for cell in row.findall("main:c", NS):
            address = cell.attrib["r"]
            column = re.sub(r"\d+", "", address)
            if column in {"A", "B", "C", "D"}:
                record[column] = cell_text(cell, shared)
        if record:
            rows.append(record)
    return rows


def build_dataset(input_path):
    rows = read_rows(input_path)
    data_rows = rows[1:]
    all_events = []
    change_events = []
    issues = []
    processed_rows = 0

    for source_index, source_row in enumerate(data_rows, start=2):
        code = str(source_row.get("A", "")).strip()
        name = str(source_row.get("B", "")).strip()
        issue_date = maybe_number(str(source_row.get("C", "")).strip())
        raw_history = clean_rating_text(source_row.get("D", ""))
        if not code and not name and not raw_history:
            continue

        processed_rows += 1

        if not raw_history or raw_history == "#NAME?":
            issues.append(
                {
                    "sourceRow": source_index,
                    "bondCode": code,
                    "bondName": name,
                    "issueDateSerial": issue_date,
                    "issue": "无历史评级缓存",
                    "raw": raw_history,
                }
            )
            continue

        parsed = []
        seen = set()
        for line_index, line in enumerate([x.strip() for x in raw_history.split("\n") if x.strip()], start=1):
            event = parse_rating_line(line)
            if event is None:
                issues.append(
                    {
                        "sourceRow": source_index,
                        "bondCode": code,
                        "bondName": name,
                        "issueDateSerial": issue_date,
                        "issue": f"无法解析第{line_index}条评级记录",
                        "raw": line,
                    }
                )
                continue
            key = (event["dateKey"], event["rating"], event["ratingType"], event["agency"])
            if key in seen:
                continue
            seen.add(key)
            parsed.append(event)

        parsed.sort(key=lambda item: item["dateKey"])
        previous = None
        for sequence, event in enumerate(parsed, start=1):
            is_first = previous is None
            event_direction = "首次评级" if is_first else direction(previous["rating"], event["rating"])
            is_change = is_first or event["rating"] != previous["rating"]
            row = {
                "bondCode": code,
                "bondName": name,
                "issueDateSerial": issue_date,
                "ratingDate": event["date"],
                "sequence": sequence,
                "changeType": event_direction,
                "previousRating": "" if is_first else previous["rating"],
                "currentRating": event["rating"],
                "previousRatingDate": "" if is_first else previous["date"],
                "ratingType": event["ratingType"],
                "agency": event["agency"],
                "isRatingChange": "是" if is_change else "否",
                "rawRecord": event["raw"],
                "sourceRow": source_index,
            }
            all_events.append(row)
            if is_change:
                change_events.append(row)
            previous = event

    summary = {
        "sourceRows": processed_rows,
        "bondsWithCachedRating": len({row["bondCode"] for row in all_events}),
        "allEventRows": len(all_events),
        "changeRowsIncludingFirstRating": len(change_events),
        "actualChangeRows": len([row for row in change_events if row["previousRating"]]),
        "issueRows": len(issues),
    }
    return {"summary": summary, "changes": change_events, "allEvents": all_events, "issues": issues}


def main():
    if len(sys.argv) != 3:
        raise SystemExit("usage: extract_rating_history_cache.py <input.xlsx> <output.json>")

    input_path = pathlib.Path(sys.argv[1])
    output_path = pathlib.Path(sys.argv[2])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    dataset = build_dataset(input_path)
    output_path.write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(dataset["summary"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
