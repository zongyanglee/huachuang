import html
import pathlib
import re
import zipfile
import xml.etree.ElementTree as ET

path = pathlib.Path(r"D:\JupyterFiles\huachuang\历史评级.xlsx")
with zipfile.ZipFile(path) as archive:
    xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8", "ignore")
    shared_xml = archive.read("xl/sharedStrings.xml").decode("utf-8", "ignore")

print(xml[:1000])
print("D2 around:")
match = re.search(r'<c r="D2".*?</c>', xml)
print(match.group(0)[:2000] if match else "not found")

ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
root = ET.fromstring(xml)
shared_root = ET.fromstring(shared_xml)
shared = []
for si in shared_root.findall("main:si", ns):
    texts = [node.text or "" for node in si.findall(".//main:t", ns)]
    shared.append("".join(texts))

def cell_text(cell):
    value = cell.find("main:v", ns)
    if value is None:
        return ""
    raw = value.text or ""
    if cell.attrib.get("t") == "s":
        return shared[int(raw)]
    return html.unescape(raw)

rows = []
for row in root.findall(".//main:sheetData/main:row", ns):
    record = {}
    for cell in row.findall("main:c", ns):
        address = cell.attrib["r"]
        col = re.sub(r"\d+", "", address)
        if col in {"A", "B", "C", "D"}:
            record[col] = cell_text(cell)
    if record:
        rows.append(record)

data = rows[1:]
rating_rows = [(i + 2, row) for i, row in enumerate(data) if row.get("D")]
print("rows with cached rating:", len(rating_rows))
for idx, row in sorted(rating_rows, key=lambda item: len(item[1].get("D", "")), reverse=True)[:20]:
    print("ROW", idx, row.get("A"), row.get("B"), "len", len(row.get("D", "")))
    print(row.get("D", "")[:1000])
