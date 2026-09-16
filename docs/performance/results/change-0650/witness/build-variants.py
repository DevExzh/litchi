#!/usr/bin/env python3
"""Rebuild change 0650's four constructed admission variants.

Each variant republishes a tracked fixture with `word/document.xml` replaced by
the same bytes with or without a leading UTF-8 byte order mark, so the mark is
the only difference between the `0` and `1` member of each pair:

    a0-nomc-nobom.docx   ooxml/docx/documentProperties.docx, unchanged
    a1-nomc-bom.docx     the same, with a byte order mark (no mc: markup)
    b0-mc-nobom.docx     ooxml/docx/Hyperlink.docx, unchanged
    b1-mc-bom.docx       the same, with a byte order mark (mc:Ignorable present)

Usage: build-variants.py <test-data-root> <output-directory>
"""

import os
import sys
import zipfile

BOM = b"\xef\xbb\xbf"


def load(path):
    archive = zipfile.ZipFile(path)
    names = [item.filename for item in archive.infolist()]
    return names, {name: archive.read(name) for name in names}


def build(names, data, out, name, document):
    path = os.path.join(out, name)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for member in names:
            payload = document if member == "word/document.xml" else data[member]
            archive.writestr(member, payload)
    return path


def main():
    root, out = sys.argv[1], sys.argv[2]
    os.makedirs(out, exist_ok=True)
    for prefix, fixture in (
        ("a", "ooxml/docx/documentProperties.docx"),
        ("b", "ooxml/docx/Hyperlink.docx"),
    ):
        names, data = load(os.path.join(root, fixture))
        document = data["word/document.xml"]
        marked = b"Ignorable" in document
        suffix = "mc" if marked else "nomc"
        build(names, data, out, f"{prefix}0-{suffix}-nobom.docx", document)
        build(names, data, out, f"{prefix}1-{suffix}-bom.docx", BOM + document)


if __name__ == "__main__":
    main()
