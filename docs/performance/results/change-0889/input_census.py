#!/usr/bin/env python3
"""Independently enumerate and CRC-check the qualification ZIP population."""
import argparse
import hashlib
import json
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]
PACKAGES = (
    "test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx",
    "test-data/ooxml/docx/drawing.docx",
    "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx",
)


def census():
    packages = []
    for name in PACKAGES:
        path = ROOT / name
        data = path.read_bytes()
        members = []
        with zipfile.ZipFile(path) as archive:
            names = [info.filename for info in archive.infolist()]
            if len(names) != len(set(names)):
                raise ValueError(f"duplicate member names in {name}")
            for info in archive.infolist():
                if info.filename.lower().endswith((".xml", ".rels", ".vml")):
                    member = archive.read(info)
                    members.append({
                        "name": info.filename,
                        "bytes": len(member),
                        "sha256": hashlib.sha256(member).hexdigest(),
                    })
        packages.append({
            "path": name,
            "bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest(),
            "members": sorted(members, key=lambda member: member["name"]),
        })
    return {
        "schema": "litchi.mce-independent-input-census-0889.v1",
        "reader": "Python zipfile; reads and CRC-validates each selected member",
        "extensions": [".xml", ".rels", ".vml"],
        "packages": packages,
        "package_count": len(packages),
        "member_count": sum(len(package["members"]) for package in packages),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    actual = census()
    retained = Path(__file__).with_name("independent-input-census.json")
    if actual != json.loads(retained.read_text()):
        raise ValueError("qualification input population changed")
    if args.output:
        with args.output.open("x") as output:
            json.dump(actual, output, indent=2)
            output.write("\n")
    print(f"PASS: {actual['package_count']} packages, {actual['member_count']} members")


if __name__ == "__main__":
    main()
