#!/usr/bin/env python3
"""Report whether each DOCX-family fixture's `word/document.xml` is marked.

Usage: bom-presence.py <test-data-root>
"""

import os
import sys
import zipfile

BOM = b"\xef\xbb\xbf"
EXTENSIONS = ("docx", "docm", "dotx", "dotm")


def main():
    root = sys.argv[1]
    rows = []
    for directory, _, names in os.walk(root):
        for name in sorted(names):
            if name.lower().rsplit(".", 1)[-1] not in EXTENSIONS:
                continue
            path = os.path.join(directory, name)
            relative = os.path.relpath(path, root)
            try:
                document = zipfile.ZipFile(path).read("word/document.xml")
            except Exception as error:  # noqa: BLE001 - a census, not a parser
                rows.append((relative, "unreadable", str(error)[:60]))
                continue
            marked = "BOM" if document.startswith(BOM) else "no-bom"
            rows.append((relative, marked, f"{len(document)}bytes"))
    rows.sort()
    for row in rows:
        print(" ".join(row))


if __name__ == "__main__":
    main()
