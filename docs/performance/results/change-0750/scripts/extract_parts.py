"""Change 0750: rebuild the audit probe's inputs from the repository.

Usage: extract_parts.py <repo-root> <census.jsonl or census.jsonl.gz> <out-dir>

Writes the five named parts the probe times and `corpus-accepted.bin`: every
`test-data` XML member of the census whose source audit both legs accept, in
census order, each as a little-endian u32 length followed by its bytes. The
SHA-256 of each output is listed in `probe/parts.txt`.
"""

import gzip
import json
import os
import struct
import sys
import zipfile

PARTS = {
    "ws-structured-1.5MB.xml": ("test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx",
                                "xl/worksheets/sheet3.xml"),
    "ws-patriarch-3.4MB.xml": ("test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx",
                               "xl/worksheets/sheet1.xml"),
    "docx-drawing-288KB.xml": ("test-data/ooxml/docx/drawing.docx", "word/document.xml"),
    "docx-table-alignment-40KB.xml": ("test-data/ooxml/docx/table-alignment.docx",
                                      "word/document.xml"),
    "pptx-slide11-121KB.xml": ("test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx",
                               "ppt/slides/slide11.xml"),
}


def rows(census):
    opener = gzip.open if census.endswith(".gz") else open
    with opener(census, "rt") as handle:
        for line in handle:
            yield json.loads(line)


def main(repo, census, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    for name, (package, member) in PARTS.items():
        with zipfile.ZipFile(os.path.join(repo, package)) as archive:
            data = archive.read(member)
        with open(os.path.join(out_dir, name), "wb") as handle:
            handle.write(data)
    payload, count, archives = bytearray(), 0, {}
    for row in rows(census):
        if not row["file"].startswith("test-data/"):
            continue
        if row["before_source"] != "OK" or row["after_source"] != "OK":
            continue
        path = os.path.join(repo, row["file"])
        if row["member"]:
            archive = archives.get(path)
            if archive is None:
                archive = archives[path] = zipfile.ZipFile(path)
            data = archive.read(row["member"])
        else:
            with open(path, "rb") as handle:
                data = handle.read()
        payload += struct.pack("<I", len(data)) + data
        count += 1
    with open(os.path.join(out_dir, "corpus-accepted.bin"), "wb") as handle:
        handle.write(payload)
    print(f"{len(PARTS)} parts; corpus-accepted.bin: {count} members, {len(payload) - 4 * count} bytes")


if __name__ == "__main__":
    main(*sys.argv[1:4])
