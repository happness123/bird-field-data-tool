"""Generate review-ready XLSX and DOCX line-transect tables from KML points.

Uses only Python standard library. It never guesses species or counts from free text.
"""

import argparse
from datetime import datetime, timedelta, timezone
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED


S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
X = "http://www.w3.org/2001/XMLSchema-instance"
ET.register_namespace("", S)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def text_at(node, path):
    x = node.find(path)
    return "" if x is None else "".join(x.itertext()).strip()


def local_time(raw):
    if not raw:
        return ""
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00")).astimezone(
            timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return ""


def read_kml(path):
    root = ET.parse(path).getroot()
    match = re.search(r"([ML]\d+|海上)", path.stem, flags=re.I)
    transect = match.group(1).upper() if match else path.stem
    rows = []
    for i, mark in enumerate(root.findall(".//{*}Placemark"), 1):
        if mark.find("{*}Point") is None:
            continue
        name = text_at(mark, "{*}name")
        if name in {"起点", "终点"}:
            continue
        when = text_at(mark, "{*}TimeStamp/{*}when")
        description = text_at(mark, "{*}description")
        description_time = re.search(r"时间[：:]\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})", description)
        time_local = local_time(when) if when else (description_time.group(1)[:16] if description_time else "")
        coords = text_at(mark, "{*}Point/{*}coordinates")
        rows.append({"样线": transect, "日期时间": time_local,
                     "时间原文": when or (description_time.group(1) if description_time else ""),
                     "物种": "", "数量": "", "行为": "", "小生境": "", "照片编号": "",
                     "备注": "", "原始标记": name, "描述原文": description,
                     "坐标原文": coords, "来源文件": path.name, "标记序号": str(i),
                     "源文件SHA256": sha(path), "复核状态": "待人工核对"})
    return rows


def xml_bytes(root):
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def make_xlsx(path, rows, headers=None):
    headers = headers or ["样线", "日期时间", "物种", "数量", "行为", "小生境", "照片编号", "备注",
               "原始标记", "描述原文", "坐标原文", "来源文件", "标记序号", "时间原文",
               "源文件SHA256", "复核状态"]
    ws = ET.Element(f"{{{S}}}worksheet")
    data = ET.SubElement(ws, f"{{{S}}}sheetData")
    for number, values in enumerate([headers] + [[r[h] for h in headers] for r in rows], 1):
        row = ET.SubElement(data, f"{{{S}}}row", {"r": str(number)})
        for col, value in enumerate(values):
            letters = ""
            n = col + 1
            while n:
                n, rem = divmod(n - 1, 26)
                letters = chr(65 + rem) + letters
            c = ET.SubElement(row, f"{{{S}}}c", {"r": f"{letters}{number}", "t": "inlineStr"})
            ET.SubElement(ET.SubElement(c, f"{{{S}}}is"), f"{{{S}}}t").text = str(value)
    workbook = ET.Element(f"{{{S}}}workbook", {"xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"})
    sheets = ET.SubElement(workbook, f"{{{S}}}sheets")
    ET.SubElement(sheets, f"{{{S}}}sheet", {"name": "样线记录_待核对", "sheetId": "1",
                                         "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": "rId1"})
    with ZipFile(path, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", xml_bytes(workbook))
        z.writestr("xl/worksheets/sheet1.xml", xml_bytes(ws))


def paragraph(parent, content):
    p = ET.SubElement(parent, f"{{{W}}}p")
    run = ET.SubElement(p, f"{{{W}}}r")
    ET.SubElement(run, f"{{{W}}}t").text = str(content)


def make_docx(path, rows):
    root = ET.Element(f"{{{W}}}document")
    body = ET.SubElement(root, f"{{{W}}}body")
    grouped = {}
    for row in rows:
        grouped.setdefault(row["来源文件"], []).append(row)
    headers = ["NO", "物种", "时间", "数量", "行为", "小生境", "照片编号", "备注", "原始标记", "标记序号"]
    for filename, group in grouped.items():
        paragraph(body, "鸟类调查样线数据表（待人工核对）")
        paragraph(body, f"样线：{group[0]['样线']}  来源：{filename}")
        paragraph(body, "KML 自由文本尚未鉴定或拆分；空白字段须由调查者核对。")
        table = ET.SubElement(body, f"{{{W}}}tbl")
        for values in [headers] + [[str(i), r["物种"], r["日期时间"], r["数量"], r["行为"],
                                     r["小生境"], r["照片编号"], r["备注"], r["原始标记"],
                                     r["标记序号"]] for i, r in enumerate(group, 1)]:
            tr = ET.SubElement(table, f"{{{W}}}tr")
            for value in values:
                cell = ET.SubElement(tr, f"{{{W}}}tc")
                paragraph(cell, value)
    ET.SubElement(body, f"{{{W}}}sectPr")
    with ZipFile(path, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", xml_bytes(root))


def main():
    parser = argparse.ArgumentParser(description="从 KML 生成 Excel 与 Word 样线表（待核对）")
    parser.add_argument("input", type=Path, help="KML 文件或目录")
    parser.add_argument("output", type=Path, help="必须不存在的新输出目录")
    args = parser.parse_args()
    if args.output.exists():
        parser.error("输出目录已存在，拒绝覆盖")
    files = [args.input] if args.input.is_file() else sorted(args.input.glob("*.kml"))
    if not files:
        parser.error("未找到 KML")
    rows = [row for file in files for row in read_kml(file)]
    args.output.mkdir(parents=True)
    make_xlsx(args.output / "鸟类样线汇总_待核对.xlsx", rows)
    make_docx(args.output / "鸟类样线数据表_待核对.docx", rows)
    (args.output / "处理记录.json").write_text(json.dumps({"inputs": [{"path": str(f.resolve()), "sha256": sha(f)} for f in files],
        "point_rows": len(rows), "warning": "输出待核对；物种、数量等未由自由文本自动判断"}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"files": len(files), "rows": len(rows), "output": str(args.output)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
