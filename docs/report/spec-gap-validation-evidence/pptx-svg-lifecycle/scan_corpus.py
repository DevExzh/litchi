#!/usr/bin/env python3
"""Bounded local discovery of svgBlip tokens in PPTX slide members."""

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import zipfile


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("root", type=Path)
    args = parser.parse_args()
    reports = []
    for directory in ("test-data/ooxml", "3rdparty"):
        paths = subprocess.check_output(
            ["rg", "--files", directory], cwd=args.root, text=True
        ).splitlines()
        packages = 0
        skipped = []
        members = []
        hits = []
        oversized = []
        for shown in sorted(paths):
            path = args.root / shown
            if path.suffix not in {".pptx", ".pptm", ".potx", ".ppsx"}:
                continue
            packages += 1
            try:
                with zipfile.ZipFile(path) as archive:
                    for item in archive.infolist():
                        if "ppt/slides/" not in item.filename or not item.filename.endswith(".xml"):
                            continue
                        if item.file_size > 2 * 1024 * 1024:
                            oversized.append([shown, item.filename, item.file_size])
                            continue
                        data = archive.read(item)
                        row = [shown, item.filename, len(data), hashlib.sha256(data).hexdigest()]
                        members.append(row)
                        if b"svgBlip" in data:
                            hits.append(row)
            except (zipfile.BadZipFile, OSError, UnicodeError, RuntimeError) as error:
                skipped.append([shown, type(error).__name__])
        reports.append({
            "directory": directory, "packages": packages,
            "slide_members": len(members), "skipped": skipped,
            "oversized": oversized, "hits": hits,
            "inspected_members": members,
        })
    print(json.dumps({"scope": "Exact token discovery only, slide members <= 2 MiB", "reports": reports}, indent=2))


if __name__ == "__main__":
    main()
