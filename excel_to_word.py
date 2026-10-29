"""Turn a reviewed line-transect XLSX into Word survey tables using stdlib only."""

import argparse
from collections import OrderedDict
from datetime import datetime, timedelta
from pathlib import Path
import hashlib
import json
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from kml_to_excel import normalize_path, HERE


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("w", W)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_xlsx(path):
    with ZipFile(path) as z:
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ["".join(si.itertext()) for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(f"{{{S}}}si")]
        root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    records = []
    for row in root.findall(f".//{{{S}}}sheetData/{{{S}}}row"):
        values = {}
        for cell in row.findall(f"{{{S}}}c"):
            column = re.match(r"[A-Z]+", cell.attrib.get("r", ""))
            if not column:
                continue
            kind = cell.attrib.get("t", "")
            v = cell.find(f"{{{S}}}v")
            if kind == "inlineStr":
                inline = cell.find(f"{{{S}}}is")
                value = "" if inline is None else "".join(inline.itertext())
            elif kind == "s" and v is not None:
                value = strings[int(v.text)]
            else:
                value = "" if v is None or v.text is None else v.text
            values[column.group()] = value.strip()
        records.append(values)
    if not records:
        raise ValueError("Excel 没有表头")
    header = records[0]
    names = {col: title for col, title in header.items() if title}
    result = []
    for row in records[1:]:
        record = {title: row.get(col, "") for col, title in names.items()}
        if any(record.values()):
            result.append(record)
    return set(names.values()), result


def format_time(value):
    if re.fullmatch(r"\d{1,2}:\d{2}", value):
        h, m = map(int, value.split(":"))
        return f"{h:02d}:{m:02d}" if h < 24 and m < 60 else value
    try:
        number = float(value)
        if 0 <= number < 1:
            minutes = round(number * 24 * 60) % (24 * 60)
            return f"{minutes // 60:02d}:{minutes % 60:02d}"
    except ValueError:
        pass
    return value


def paragraph(parent, value, bold=False):
    p = ET.SubElement(parent, f"{{{W}}}p")
    run = ET.SubElement(p, f"{{{W}}}r")
    if bold:
        ET.SubElement(ET.SubElement(run, f"{{{W}}}rPr"), f"{{{W}}}b")
    t = ET.SubElement(run, f"{{{W}}}t")
    t.text = str(value)
    return p


def remarks(row):
    note = row.get("备注", "")
    if "待" in row.get("复核状态", ""):
        raw = row.get("原始标记", "")
        note = f"{note}；待复核原始标记：{raw}" if note else f"待复核原始标记：{raw}"
    return note


def write_docx(path, groups):
    root = ET.Element(f"{{{W}}}document")
    body = ET.SubElement(root, f"{{{W}}}body")
    headers = ["NO", "物种", "时间", "数量", "行为", "小生境", "照片编号", "备注"]
    for index, ((date, transect), rows) in enumerate(groups.items()):
        if index:
            p = ET.SubElement(body, f"{{{W}}}p")
            pp = ET.SubElement(p, f"{{{W}}}pPr")
            ET.SubElement(pp, f"{{{W}}}pageBreakBefore")
        paragraph(body, "鸟类调查记录表", bold=True)
        paragraph(body, f"调查日期：{date or '待填写'}    天气：待填写    调查人：待填写")
        paragraph(body, f"样带编号：{transect or '待填写'}    起点：待填写    终点：待填写")
        paragraph(body, "调查开始时间：待填写    结束时间：待填写    样带长度：待填写")
        table = ET.SubElement(body, f"{{{W}}}tbl")
        props = ET.SubElement(table, f"{{{W}}}tblPr")
        borders = ET.SubElement(props, f"{{{W}}}tblBorders")
        for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
            ET.SubElement(borders, f"{{{W}}}{edge}", {f"{{{W}}}val": "single", f"{{{W}}}sz": "4"})
        for values in [headers] + [[str(i), r.get("物种", ""), format_time(r.get("时间", "")),
                                     r.get("数量", ""), r.get("行为", ""), r.get("小生境", ""),
                                     r.get("照片编号", ""), remarks(r)]
                                    for i, r in enumerate(rows, 1)]:
            tr = ET.SubElement(table, f"{{{W}}}tr")
            for value in values:
                cell = ET.SubElement(tr, f"{{{W}}}tc")
                paragraph(cell, value)
        paragraph(body, "原始标记、经纬度和来源见输入 Excel；请在核对后使用本表。")
    ET.SubElement(body, f"{{{W}}}sectPr")
    xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with ZipFile(path, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", xml)


def main():
    parser = argparse.ArgumentParser(description="将已核对的 Excel 按日期和样线转为 Word 样线数据表")
    parser.add_argument("excel", help="Excel 文件路径，可用 Windows 或 WSL 路径")
    parser.add_argument("--output-dir", help="可选：把 Word 放入指定的已有输出目录")
    args = parser.parse_args()
    source = normalize_path(args.excel)
    if not source.is_file() or source.suffix.lower() != ".xlsx":
        parser.error("请提供现有的 .xlsx 文件")
    headers, records = read_xlsx(source)
    required = {"样线", "日期", "时间", "物种", "数量", "行为"}
    missing = required - headers
    if missing:
        parser.error("Excel 缺少必需列：" + "、".join(sorted(missing)))
    groups = OrderedDict()
    for row in records:
        groups.setdefault((row.get("日期", ""), row.get("样线", "")), []).append(row)
    output_dir = normalize_path(args.output_dir) if args.output_dir else HERE / "outputs" / (
        re.sub(r"[^\w\u4e00-\u9fff-]", "_", source.stem)[:60]
        + "_Word_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
    output_dir.mkdir(parents=True, exist_ok=bool(args.output_dir))
    target = output_dir / "鸟类样线数据表_待核对.docx"
    if target.exists():
        parser.error("Word 输出文件已存在，拒绝覆盖")
    write_docx(target, groups)
    manifest = output_dir / ("Word处理记录.json" if args.output_dir else "处理记录.json")
    manifest.write_text(json.dumps({"input": str(source.resolve()),
        "input_sha256": sha(source), "output": str(target), "groups": len(groups),
        "rows": len(records), "rule": "按日期和样线分组；未知元数据留待填写；原始 Excel 不修改"},
        ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Word 文件：{target}")
    print(f"共 {len(groups)} 张样线表、{len(records)} 条 Excel 记录。")


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, ET.ParseError) as exc:
        sys.exit(f"生成失败：{exc}")
