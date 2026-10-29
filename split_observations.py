"""Conservative KML note splitting to reviewable Excel rows."""

import argparse
from datetime import datetime, timedelta, timezone
import json
from pathlib import Path
import re
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from generate_forms import read_kml, make_xlsx, sha


BEHAVIORS = ["盘旋飞过", "水中停留", "鸣叫", "飞过", "停留", "盘旋", "休憩", "捕食", "繁殖", "求偶", "筑巢"]
CN = {"一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
NUM = r"(?:\d+|[一二两三四五六七八九十])"


def reference_species(path):
    """Read species strings only; corrections in the old workbook are not applied."""
    with ZipFile(path) as z:
        strings = ["".join(x.itertext()) for x in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("{*}si")]
        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    names = set()
    for row in sheet.findall(".//{*}sheetData/{*}row")[1:]:
        for cell in row.findall("{*}c"):
            ref = cell.attrib.get("r", "")
            if not ref.startswith(("B", "C")) or cell.attrib.get("t") != "s":
                continue
            v = cell.find("{*}v")
            if v is not None:
                names.add(strings[int(v.text)].strip())
    return sorted((n for n in names if n), key=len, reverse=True)


def parse_num(token):
    if not token:
        return ""
    token = token.replace("只", "").replace("个", "")
    return str(CN[token]) if token in CN else token if token.isdigit() else ""


def split_note(note, names):
    """Return species candidates; keep uncertain text rather than inventing species."""
    structured = list(re.finditer(r"(?<!\S)([\u4e00-\u9fff]+)\s+(\d+)\s*(?:只|个)?(?=\s|$)", note))
    if structured:
        out = []
        for i, hit in enumerate(structured):
            segment = note[hit.end():structured[i + 1].start() if i + 1 < len(structured) else len(note)].strip()
            behavior = next((b for b in BEHAVIORS if b in segment), "")
            habitat = segment.replace(behavior, "", 1).strip() if behavior else segment
            out.append({"物种": hit.group(1), "数量": hit.group(2), "行为": behavior,
                        "小生境": habitat, "复核状态": "待人工核对" if behavior else "行为待人工核对",
                        "片段": note[hit.start():structured[i + 1].start() if i + 1 < len(structured) else len(note)].strip()})
        return out
    pattern = re.compile("|".join(re.escape(n) for n in names)) if names else None
    hits = list(pattern.finditer(note)) if pattern else []
    if not hits:
        return [{"物种": "", "数量": "", "行为": "", "小生境": "", "复核状态": "未识别物种，待人工拆分", "片段": note}]
    out = []
    for i, hit in enumerate(hits):
        prev_end = hits[i - 1].end() if i else 0
        next_start = hits[i + 1].start() if i + 1 < len(hits) else len(note)
        before = note[prev_end:hit.start()]
        after = note[hit.end():next_start]
        prefix = re.search(rf"({NUM})\s*[只个]?\s*$", before)
        suffix = re.match(rf"\s*({NUM})\s*[只个]?", after)
        trailing = re.search(rf"({NUM})\s*[只个]?\s*$", after) if not suffix else None
        count = parse_num(suffix.group(1) if suffix else prefix.group(1) if prefix else trailing.group(1) if trailing else "")
        behaviors = [b for b in BEHAVIORS if b in after and not any(b != longer and b in longer and longer in after for longer in BEHAVIORS)]
        behavior = behaviors[0] if len(behaviors) == 1 else ""
        status = "待人工核对"
        if not count or not behavior or len(behaviors) > 1:
            status = "数量或行为不明确，待人工核对"
        if i and before.strip() and not prefix and not suffix:
            status = "片段归属不明确，待人工核对"
        out.append({"物种": hit.group(), "数量": count, "行为": behavior, "小生境": "",
                    "复核状态": status, "片段": before + hit.group() + after})
    return out


HEADERS = ["样线", "日期", "时间", "物种", "数量", "行为", "小生境", "经度", "纬度", "复核状态",
           "原始标记", "拆分片段", "来源文件", "标记序号", "源文件SHA256", "时间原文"]


def main():
    p = argparse.ArgumentParser(description="KML 鸟类标记拆分为待复核 Excel")
    p.add_argument("input", type=Path, help="KML 文件或目录")
    p.add_argument("output", type=Path, help="必须不存在的新输出目录")
    p.add_argument("--reference", type=Path, help="可选：已有 Excel 的物种词表；紧凑写法需要")
    a = p.parse_args()
    if a.output.exists():
        p.error("输出目录已存在，拒绝覆盖")
    files = [a.input] if a.input.is_file() else sorted(a.input.glob("*.kml"))
    if not files:
        p.error("未找到 KML")
    if a.reference and not a.reference.is_file():
        p.error("物种参考表不存在")
    names = reference_species(a.reference) if a.reference else []
    rows = []
    for file in files:
        for raw in read_kml(file):
            note = raw["原始标记"]
            if not note or note in {"开始", "出发开始"}:
                continue
            coord = raw["坐标原文"].split(",")
            date_time = raw["日期时间"].split(" ")
            for part in split_note(note, names):
                rows.append({"样线": raw["样线"], "日期": date_time[0] if len(date_time) == 2 else "",
                    "时间": date_time[1] if len(date_time) == 2 else "", "物种": part["物种"],
                    "数量": part["数量"], "行为": part["行为"], "小生境": part["小生境"],
                    "经度": coord[0].strip() if coord else "", "纬度": coord[1].strip() if len(coord) > 1 else "",
                    "复核状态": part["复核状态"], "原始标记": note, "拆分片段": part["片段"],
                    "来源文件": raw["来源文件"], "标记序号": raw["标记序号"],
                    "源文件SHA256": raw["源文件SHA256"], "时间原文": raw["时间原文"]})
    a.output.mkdir(parents=True)
    make_xlsx(a.output / "鸟类观测拆分_待复核.xlsx", rows, HEADERS)
    summary = {"inputs": [{"path": str(f.resolve()), "sha256": sha(f)} for f in files],
               "reference": {"path": str(a.reference.resolve()), "sha256": sha(a.reference)} if a.reference else None,
               "species_lexicon_size": len(names), "rows": len(rows),
               "unidentified": sum(not r["物种"] for r in rows),
               "review_required": sum("待人工" in r["复核状态"] for r in rows),
               "rule": "词表来自旧 Excel；只匹配原文，未自动采用旧表的物种调整；全部结果待人工核对"}
    (a.output / "处理记录.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"rows": len(rows), "unidentified": summary["unidentified"], "output": str(a.output)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
