"""Regenerate the mutation-differential seed corpus for change 0588.

Five synthetic MCE documents are retained beside this script; the four large
seeds are real fixture parts and are extracted here instead of being copied, so
the packet stays small.

usage: make-seeds.py <test-data-dir> <seed-dir>
"""

import glob
import os
import sys
import zipfile

ROOT, OUT = sys.argv[1], sys.argv[2]
MC = b"http://schemas.openxmlformats.org/markup-compatibility/2006"

FIXED = [
    ("poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx",
     "xl/worksheets/sheet1.xml", "sheet1.xml"),
]
# The remaining three are the first fixture of each family whose named member
# carries the MCE namespace and is between 500 and 60,000 bytes.
SEARCH = [
    ("*.docx", "word/document.xml"),
    ("*.pptx", "ppt/slides/slide1.xml"),
    ("*.xlsx", "xl/workbook.xml"),
]

os.makedirs(OUT, exist_ok=True)
for relative, member, name in FIXED:
    with zipfile.ZipFile(os.path.join(ROOT, relative)) as archive:
        open(os.path.join(OUT, name), "wb").write(archive.read(member))
        print("seed", name)
for pattern, member in SEARCH:
    for path in sorted(glob.glob(os.path.join(ROOT, "**", pattern), recursive=True)):
        try:
            with zipfile.ZipFile(path) as archive:
                data = archive.read(member)
        except Exception:  # noqa: BLE001 - the corpus contains deliberate bad zips
            continue
        if MC in data and 500 < len(data) < 60000:
            name = os.path.basename(path) + "-" + member.replace("/", "_")
            open(os.path.join(OUT, name), "wb").write(data)
            print("seed", name, "from", path)
            break
