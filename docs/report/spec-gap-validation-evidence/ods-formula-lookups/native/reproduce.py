#!/usr/bin/env python3
"""Reproduce the retained local LibreOffice lookup-function receipt.

Every conversion uses a fresh headless profile.  The retained ODS ZIP is
checked through its stable ``content.xml`` and typed formula rows, while the
retained ZIP hash remains provenance only because container metadata can vary
between conversions.
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
INPUT = ROOT / "lookup-functions-native.fods"
OUTPUT = ROOT / "recalculated.ods"
CONTENT = ROOT / "content.xml"
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


def element_text(element: ET.Element) -> str:
    return "".join(element.itertext())


def typed_native(cell: ET.Element) -> dict[str, object]:
    raw = element_text(cell)
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
    for table in root.iter(TABLE + "table"):
        for row in table.findall(TABLE + "table-row"):
            cells = list(row.findall(TABLE + "table-cell"))
            formula_cells = [cell for cell in cells if FORMULA in cell.attrib]
            if not formula_cells:
                continue
            cell = formula_cells[0]
            item: dict[str, object] = {
                "case": element_text(cells[0]) if cells else "",
                "formula": cell.attrib[FORMULA],
            }
            if typed:
                item["native"] = typed_native(cell)
            rows.append(item)
    return rows


def hidden_table_is_hidden(content: bytes) -> bool:
    root = ET.fromstring(content)
    styles: dict[str, bool] = {}
    for style in root.iter(STYLE + "style"):
        name = style.attrib.get(STYLE_NAME)
        if not name:
            continue
        properties = style.find(STYLE + "table-properties")
        styles[name] = bool(
            properties is not None
            and properties.attrib.get(TABLE_DISPLAY) == "false"
        )
    for table in root.iter(TABLE + "table"):
        if table.attrib.get(TABLE + "name") == "Hidden":
            style_name = table.attrib.get(TABLE_STYLE_NAME)
            return bool(style_name and styles.get(style_name))
    return False


def main() -> None:
    provenance = json.loads(PROVENANCE.read_text(encoding="utf-8"))
    results_document = json.loads(RESULTS.read_text(encoding="utf-8"))
    expected = results_document["observations"]
    if not isinstance(expected, list) or not expected:
        raise RuntimeError("native results observations are empty")

    input_bytes = INPUT.read_bytes()
    fixture = provenance["fixture"]
    if sha256(input_bytes) != fixture["input_sha256"]:
        raise RuntimeError("native input SHA-256 changed")
    if not hidden_table_is_hidden(input_bytes):
        raise RuntimeError("native input Hidden table is not display=false")
    source_rows = formula_rows(input_bytes, typed=False)
    projection = [{"case": row["case"], "formula": row["formula"]} for row in expected]
    if source_rows != projection:
        raise RuntimeError("retained native results do not match input formula rows")

    executable = os.environ.get("LIBREOFFICE", "/usr/bin/libreoffice")
    environment = dict(os.environ)
    for name in ("LANG", "LC_ALL", "LC_CTYPE", "LC_NUMERIC"):
        environment[name] = "C.UTF-8"
    with tempfile.TemporaryDirectory(prefix="litchi-lookups-native-") as name:
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
        generated = output_dir / "lookup-functions-native.ods"
        if not generated.is_file():
            raise RuntimeError("LibreOffice did not create lookup-functions-native.ods")
        with zipfile.ZipFile(generated) as archive:
            generated_content = archive.read("content.xml")
        if not hidden_table_is_hidden(generated_content):
            raise RuntimeError("recalculated Hidden table is not display=false")
        content_hash = sha256(generated_content)
        if content_hash != fixture["recalculated_content_xml_sha256"]:
            raise RuntimeError(
                f"content.xml SHA-256 changed: {content_hash} != "
                f"{fixture['recalculated_content_xml_sha256']}"
            )
        observed = formula_rows(generated_content, typed=True)
        expected_native = [
            {
                "case": row["case"],
                "formula": row["formula"],
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

    if sha256(CONTENT.read_bytes()) != fixture["recalculated_content_xml_sha256"]:
        raise RuntimeError("retained content.xml does not match the retained ODS")
    parity = sum(row.get("comparison") == "parity" for row in expected)
    divergences = sum(row.get("comparison") == "native-divergence" for row in expected)
    print(
        json.dumps(
            {
                "status": "verified",
                "converter": executable,
                "fresh_profile": True,
                "temporary_tree_cleaned": True,
                "rows": len(observed),
                "parity": parity,
                "divergences": divergences,
                "content_xml_sha256": content_hash,
            },
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
