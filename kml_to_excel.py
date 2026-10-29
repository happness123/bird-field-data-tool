"""One-argument launcher for the reviewable KML-to-Excel converter."""

import argparse
from datetime import datetime
import os
from pathlib import Path
import re
import subprocess
import sys


HERE = Path(__file__).resolve().parent
DEFAULT_REFERENCE = HERE / "species_reference.xlsx"


def normalize_path(raw):
    if os.name == "nt":
        return Path(raw).expanduser()
    match = re.fullmatch(r"([A-Za-z]):[\\/](.*)", raw)
    if match:
        return Path("/mnt") / match.group(1).lower() / match.group(2).replace("\\", "/")
    return Path(raw).expanduser()


def main():
    parser = argparse.ArgumentParser(description="给一个 KML 文件，生成待复核 Excel")
    parser.add_argument("kml", help="KML 路径；可用 F:\\... 或 /mnt/f/... 格式")
    parser.add_argument("--reference", help="可选：已有物种表 Excel；紧凑写法建议提供")
    args = parser.parse_args()
    source = normalize_path(args.kml)
    reference = normalize_path(args.reference) if args.reference else DEFAULT_REFERENCE
    if not source.is_file() or source.suffix.lower() != ".kml":
        parser.error("KML 文件不存在，或扩展名不是 .kml")
    if args.reference and not reference.is_file():
        parser.error(f"物种参考表不存在：{reference}")
    safe_stem = re.sub(r"[^\w\u4e00-\u9fff-]", "_", source.stem)[:60]
    output = HERE / "outputs" / f"{safe_stem}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
    output.parent.mkdir(exist_ok=True)
    command = [sys.executable, str(HERE / "split_observations.py"), str(source), str(output)]
    if reference.is_file():
        command.extend(["--reference", str(reference)])
    completed = subprocess.run(command, check=False)
    if completed.returncode:
        return completed.returncode
    print(f"Excel 文件：{output / '鸟类观测拆分_待复核.xlsx'}")
    print("请打开复核状态、原始标记和来源列，确认物种、数量及行为。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
