#!/usr/bin/env python3
"""Reproduce the local LibreOffice value-inspection observation receipt.

The conversion uses a fresh temporary LibreOffice profile and a temporary
output directory.  The retained ODS is a raw capture; the stable
``content.xml`` hash and the typed formula sequence are the portable checks.
"""

from __future__ import annotations

import hashlib
import json
import math
import os
from pathlib import Path
import re
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "inspection-native.fods"
RESULTS = ROOT / "native-results.json"
PROVENANCE = ROOT / "provenance.json"

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
CALCEXT = "{urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0}"
FORMULA = TABLE + "formula"
VALUE_TYPE = OFFICE + "value-type"
VALUE = OFFICE + "value"
BOOLEAN_VALUE = OFFICE + "boolean-value"


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def element_text(element: ET.Element) -> str:
    return "".join(element.itertext())


def typed_native(cell: ET.Element) -> dict[str, object]:
    kind = cell.attrib.get(VALUE_TYPE)
    raw = element_text(cell)
    if cell.attrib.get(CALCEXT + "value-type") == "error":
        return {"type": "error", "value": raw}
    if kind == "float":
        value = float(cell.attrib.get(VALUE, raw))
        if value.is_integer() and abs(value) < 2**53:
            value = int(value)
        return {"type": "number", "value": value}
    if kind == "boolean":
        return {
            "type": "logical",
            "value": cell.attrib.get(BOOLEAN_VALUE, raw).lower() == "true",
        }
    if kind == "string":
        return {"type": "text", "value": raw}
    raise RuntimeError(f"unexpected native result type: {kind!r}")


def extract(output: Path) -> list[dict[str, object]]:
    with zipfile.ZipFile(output) as archive:
        content = archive.read("content.xml")
    root = ET.fromstring(content)
    observed: list[dict[str, object]] = []
    for row in root.iter(TABLE + "table-row"):
        cells = list(row.findall(TABLE + "table-cell"))
        formula_cells = [cell for cell in cells if FORMULA in cell.attrib]
        if not formula_cells:
            continue
        # The first formula row is the source row.  It has no case label and
        # is intentionally excluded from the receipt's formula sequence.
        case = element_text(cells[0]) if cells else ""
        if not case:
            continue
        formula_cell = formula_cells[0]
        observed.append(
            {
                "case": case,
                "formula": formula_cell.attrib[FORMULA],
                "native": typed_native(formula_cell),
            }
        )
    return observed


def compare_native(
    expected: list[dict[str, object]], observed: list[dict[str, object]]
) -> None:
    projection = [
        {
            "case": item["case"],
            "formula": item["formula"],
            "native_formula": item.get("native_formula", item["formula"]),
            "native": item["native"],
        }
        for item in expected
    ]
    if len(observed) != len(projection):
        raise RuntimeError(f"formula row count changed: {len(observed)} != {len(projection)}")
    for expected_item, observed_item in zip(projection, observed, strict=True):
        if expected_item["case"] != observed_item["case"]:
            raise RuntimeError(
                f"formula case sequence changed: {expected_item['case']!r} != "
                f"{observed_item['case']!r}"
            )
        expected_formula = expected_item.get("native_formula", expected_item["formula"])
        if expected_formula != observed_item["formula"]:
            raise RuntimeError(
                f"formula spelling changed for {expected_item['case']}: "
                f"{expected_formula!r} != {observed_item['formula']!r}"
            )
        if expected_item["native"] != observed_item["native"]:
            raise RuntimeError(
                "typed native observations changed:\n"
                + json.dumps(
                    {"expected": expected_item, "observed": observed_item},
                    ensure_ascii=False,
                    indent=2,
                )
            )


def main() -> None:
    provenance = json.loads(PROVENANCE.read_text(encoding="utf-8"))
    expected = json.loads(RESULTS.read_text(encoding="utf-8"))
    input_digest = sha256(INPUT.read_bytes())
    if input_digest != provenance["fixture"]["input_sha256"]:
        raise RuntimeError(
            f"input SHA-256 changed: {input_digest} != "
            f"{provenance['fixture']['input_sha256']}"
        )

    executable = os.environ.get("LIBREOFFICE", "/usr/bin/libreoffice")
    environment = os.environ.copy()
    for name in ("LANG", "LC_ALL", "LC_CTYPE", "LC_NUMERIC"):
        environment[name] = "C.UTF-8"
    with tempfile.TemporaryDirectory(prefix="litchi-inspection-native-") as temporary_name:
        temporary = Path(temporary_name)
        profile = temporary / "profile"
        output_dir = temporary / "output"
        profile.mkdir()
        output_dir.mkdir()
        command = [
            executable,
            "--headless",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to",
            "ods",
            "--outdir",
            str(output_dir),
            str(INPUT),
        ]
        completed = subprocess.run(
            command, env=environment, text=True, capture_output=True
        )
        if completed.returncode:
            raise RuntimeError(
                f"LibreOffice conversion failed ({completed.returncode}):\n"
                f"{completed.stdout}{completed.stderr}"
            )
        output = output_dir / "inspection-native.ods"
        if not output.is_file():
            raise RuntimeError("LibreOffice did not create inspection-native.ods")

        with zipfile.ZipFile(output) as archive:
            content = archive.read("content.xml")
        content_digest = sha256(content)
        expected_content_digest = provenance["fixture"][
            "recalculated_content_xml_sha256"
        ]
        if content_digest != expected_content_digest:
            raise RuntimeError(
                f"content.xml SHA-256 changed: {content_digest} != "
                f"{expected_content_digest}"
            )
        observed = extract(output)
        compare_native(expected, observed)

    comparisons: dict[str, int] = {}
    for item in expected:
        comparison = item["comparison"]
        comparisons[comparison] = comparisons.get(comparison, 0) + 1
    print(
        json.dumps(
            {
                "converter": executable,
                "fresh_profile": True,
                "temporary_tree_cleaned": True,
                "rows": len(observed),
                "comparison_counts": comparisons,
                "content_xml_sha256": content_digest,
            },
            ensure_ascii=False,
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
