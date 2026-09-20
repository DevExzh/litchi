#!/usr/bin/env python3
"""Reproduce the retained local LibreOffice reference-metadata receipt.

The conversion always uses a fresh headless profile in a temporary directory.
The retained ODS is compared through its stable ``content.xml`` and typed
formula sequence; ZIP container metadata is not used as a semantic oracle.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "reference-metadata-native.fods"
OUTPUT = ROOT / "recalculated.ods"
RESULTS = ROOT / "native-results.json"
PROVENANCE = ROOT / "provenance.json"

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
STYLE = "{urn:oasis:names:tc:opendocument:xmlns:style:1.0}"
CALCEXT = "{urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0}"
FORMULA = TABLE + "formula"
VALUE_TYPE = OFFICE + "value-type"
VALUE = OFFICE + "value"
BOOLEAN_VALUE = OFFICE + "boolean-value"
TABLE_STYLE_NAME = TABLE + "style-name"
TABLE_DISPLAY = TABLE + "display"
STYLE_NAME = STYLE + "name"


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def text_value(element: ET.Element) -> str:
    return "".join(element.itertext())


def typed_native(cell: ET.Element) -> dict[str, object]:
    raw = text_value(cell)
    if cell.attrib.get(CALCEXT + "value-type") == "error":
        return {"type": "error", "value": raw}
    kind = cell.attrib.get(VALUE_TYPE)
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
    raise RuntimeError(f"unexpected native formula result type: {kind!r}")


def formula_rows(content: bytes, *, typed: bool) -> list[dict[str, object]]:
    root = ET.fromstring(content)
    rows: list[dict[str, object]] = []
    for row in root.iter(TABLE + "table-row"):
        cells = list(row.findall(TABLE + "table-cell"))
        formula_cells = [cell for cell in cells if FORMULA in cell.attrib]
        if not formula_cells:
            continue
        formula_cell = formula_cells[0]
        item: dict[str, object] = {
            "case": text_value(cells[0]) if cells else "",
            "formula": formula_cell.attrib[FORMULA],
        }
        if typed:
            item["native"] = typed_native(formula_cell)
        rows.append(item)
    return rows


def hidden_table_is_hidden(content: bytes) -> bool:
    """Require a real table style with ``table:display=false`` for Hidden."""

    root = ET.fromstring(content)
    styles: dict[str, bool] = {}
    for style in root.iter(STYLE + "style"):
        name = style.attrib.get(STYLE_NAME)
        if not name:
            continue
        properties = style.find(STYLE + "table-properties")
        styles[name] = (
            properties is not None and properties.attrib.get(TABLE_DISPLAY) == "false"
        )
    for table in root.iter(TABLE + "table"):
        if table.attrib.get(TABLE + "name") == "Hidden":
            style_name = table.attrib.get(TABLE_STYLE_NAME)
            return bool(style_name and styles.get(style_name))
    return False


def main() -> None:
    provenance = json.loads(PROVENANCE.read_text(encoding="utf-8"))
    results = json.loads(RESULTS.read_text(encoding="utf-8"))
    expected = results["observations"] if isinstance(results, dict) else results
    if not isinstance(expected, list) or not expected:
        raise RuntimeError("native results observations are empty")
    input_bytes = INPUT.read_bytes()
    if sha256(input_bytes) != provenance["fixture"]["input_sha256"]:
        raise RuntimeError("native input SHA-256 changed")
    if not hidden_table_is_hidden(input_bytes):
        raise RuntimeError("native input Hidden table is not display=false")
    expected_input_rows = formula_rows(input_bytes, typed=False)
    expected_projection = [
        {
            "case": row["case"],
            "formula": row["formula"],
        }
        for row in expected
    ]
    if expected_input_rows != expected_projection:
        raise RuntimeError("retained native results do not match the input formula sequence")

    executable = os.environ.get("LIBREOFFICE", "/usr/bin/libreoffice")
    environment = dict(os.environ)
    for name in ("LANG", "LC_ALL", "LC_CTYPE", "LC_NUMERIC"):
        environment[name] = "C.UTF-8"
    with tempfile.TemporaryDirectory(prefix="litchi-reference-metadata-native-") as name:
        temporary = Path(name)
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
            command,
            env=environment,
            text=True,
            capture_output=True,
        )
        if completed.returncode:
            raise RuntimeError(
                f"LibreOffice conversion failed ({completed.returncode}):\n"
                f"{completed.stdout}{completed.stderr}"
            )
        generated = output_dir / "reference-metadata-native.ods"
        if not generated.is_file():
            raise RuntimeError("LibreOffice did not create reference-metadata-native.ods")
        generated_bytes = generated.read_bytes()
        with zipfile.ZipFile(generated) as archive:
            content = archive.read("content.xml")
        if not hidden_table_is_hidden(content):
            raise RuntimeError("recalculated Hidden table is not display=false")
        content_hash = sha256(content)
        expected_content_hash = provenance["fixture"]["recalculated_content_xml_sha256"]
        if content_hash != expected_content_hash:
            raise RuntimeError(
                f"content.xml SHA-256 changed: {content_hash} != {expected_content_hash}"
            )
        observed = formula_rows(content, typed=True)
        expected_native = [
            {
                "case": row["case"],
                "formula": row.get("native_formula", row["formula"]),
                "native": row["native"],
            }
            for row in expected
        ]
        if observed != expected_native:
            raise RuntimeError(
                "typed native observations changed:\n"
                + json.dumps(
                    {"expected": expected_native, "observed": observed},
                    ensure_ascii=False,
                    indent=2,
                )
            )

    # Keep the retained capture immutable.  The fresh conversion above is the
    # reproducible check; its ZIP bytes are intentionally not copied over the
    # retained output because ZIP member timestamps are host/container data.
    comparisons: dict[str, int] = {}
    for row in expected:
        comparison = row.get("comparison", "observation")
        comparisons[comparison] = comparisons.get(comparison, 0) + 1
    print(
        json.dumps(
            {
                "status": "verified",
                "converter": executable,
                "fresh_profile": True,
                "temporary_tree_cleaned": True,
                "rows": len(observed),
                "comparison_counts": comparisons,
                "content_xml_sha256": content_hash,
                "hidden_sheet_verified": True,
            },
            ensure_ascii=False,
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
