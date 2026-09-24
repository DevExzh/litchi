#!/usr/bin/env python3
"""Create deterministic accepted-diagnostic and refused-graph XLSX cases."""

from __future__ import annotations

import pathlib
import sys
import zipfile


FIXED_DATE = (1980, 1, 1, 0, 0, 0)


def clone_info(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    clone = zipfile.ZipInfo(info.filename, FIXED_DATE)
    clone.compress_type = zipfile.ZIP_DEFLATED
    clone.create_system = 0
    clone.external_attr = 0
    return clone


def rewrite_members(base: pathlib.Path, output: pathlib.Path, edits: dict[str, str]) -> None:
    with zipfile.ZipFile(base, "r") as source:
        members = {info.filename: source.read(info.filename) for info in source.infolist()}
    for name, (old, new) in edits.items():
        value = members[name].decode("utf-8")
        if old not in value:
            raise RuntimeError(f"missing {old!r} in {name}")
        members[name] = value.replace(old, new, 1).encode()
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w") as destination:
        for name, payload in members.items():
            destination.writestr(clone_info(zipfile.ZipInfo(name)), payload)


def main() -> None:
    if len(sys.argv) != 4:
        raise SystemExit("usage: generate_cases.py BASE OUTPUT DIAGNOSTIC|REFUSAL")
    base = pathlib.Path(sys.argv[1])
    output = pathlib.Path(sys.argv[2])
    mode = sys.argv[3]
    if mode == "DIAGNOSTIC":
        rewrite_members(
            base,
            output,
            {
                "xl/drawings/vmlDrawing1.vml": (
                    "<x:Checked>1</x:Checked>",
                    "<x:Checked>0</x:Checked>",
                )
            },
        )
    elif mode == "REFUSAL":
        drawing = zipfile.ZipFile(base, "r").read("xl/drawings/drawing1.xml").decode("utf-8")
        shape_start = drawing.index("<xdr:sp ")
        shape_end = drawing.index("</xdr:sp>", shape_start) + len("</xdr:sp>")
        shape = drawing[shape_start:shape_end]
        wrapped = '<f:wrapper xmlns:f="urn:test:foreign">' + shape + "</f:wrapper>"
        rewrite_members(base, output, {"xl/drawings/drawing1.xml": (shape, wrapped)})
    else:
        raise SystemExit(f"unknown mode {mode!r}")


if __name__ == "__main__":
    main()
