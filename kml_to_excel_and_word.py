"""Single-command KML to reviewable Excel and Word output."""

import argparse
from datetime import datetime
from pathlib import Path
import re
import subprocess
import sys

from kml_to_excel import HERE, normalize_path


def main():
    parser = argparse.ArgumentParser(description="一个 KML 同时生成 Excel 和 Word 样线表")
    parser.add_argument("kml", help="输入 KML 文件路径")
    parser.add_argument("--reference", help="可选：旧 Excel 物种词表，供紧凑写法识别")
    args = parser.parse_args()
    source = normalize_path(args.kml)
    if not source.is_file() or source.suffix.lower() != ".kml":
        parser.error("KML 文件不存在，或扩展名不是 .kml")
    reference = normalize_path(args.reference) if args.reference else None
    if reference and not reference.is_file():
        parser.error("物种参考表不存在")
    safe_stem = re.sub(r"[^\w\u4e00-\u9fff-]", "_", source.stem)[:60]
    output = HERE / "outputs" / f"{safe_stem}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
    output.parent.mkdir(exist_ok=True)
    first = [sys.executable, str(HERE / "split_observations.py"), str(source), str(output)]
    if reference:
        first.extend(["--reference", str(reference)])
    if subprocess.run(first, check=False).returncode:
        return 1
    excel = output / "鸟类观测拆分_待复核.xlsx"
    second = [sys.executable, str(HERE / "excel_to_word.py"), str(excel), "--output-dir", str(output)]
    if subprocess.run(second, check=False).returncode:
        print(f"Excel 已生成，Word 失败；Excel 路径：{excel}", file=sys.stderr)
        return 1
    print(f"Excel：{excel}")
    print(f"Word：{output / '鸟类样线数据表_待核对.docx'}")
    print("两份文件均待人工核对；调查天气、人员及起止时间等缺失信息需补充。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
