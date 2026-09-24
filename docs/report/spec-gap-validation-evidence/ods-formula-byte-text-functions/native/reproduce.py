#!/usr/bin/env python3
"""Reproduce the local LibreOffice byte-function observation receipt.

The converter runs with a fresh temporary user profile.  Only the generated
ODS is inspected; the temporary profile and output directory are removed by
TemporaryDirectory when the check finishes.  The retained ODS ZIP hash is a
capture receipt, while the content.xml hash and typed formula results are the
portable checks because ZIP metadata may vary between conversions.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "byte-functions-native.fods"
RESULTS = ROOT / "native-results.json"
PROVENANCE = ROOT / "provenance.json"

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
TEXT = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
FORMULA = TABLE + "formula"
VALUE_TYPE = OFFICE + "value-type"
VALUE = OFFICE + "value"
STRING_VALUE = OFFICE + "string-value"


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def text_value(cell: ET.Element) -> str:
    return "".join(cell.itertext())


def typed_native(cell: ET.Element) -> dict[str, object]:
    kind = cell.attrib.get(VALUE_TYPE)
    if kind == "float":
        raw = cell.attrib[VALUE]
        if re.fullmatch(r"[-+]?\d+", raw):
            value: object = int(raw)
        else:
            value = float(raw)
        return {"type": "number", "value": value}
    if kind == "string":
        return {"type": "text", "value": cell.attrib.get(STRING_VALUE, text_value(cell))}
    if kind == "error":
        return {"type": "error", "value": cell.attrib.get(VALUE, text_value(cell))}
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
        if len(cells) != 3 or len(formula_cells) != 1:
            raise RuntimeError("fixture row shape changed")
        formula_cell = formula_cells[0]
        observed.append(
            {
                "case": text_value(cells[0]),
                "formula": formula_cell.attrib[FORMULA],
                "native": typed_native(formula_cell),
            }
        )
    return observed


def main() -> None:
    provenance = json.loads(PROVENANCE.read_text(encoding="utf-8"))
    expected = json.loads(RESULTS.read_text(encoding="utf-8"))
    input_bytes = INPUT.read_bytes()
    input_digest = sha256(input_bytes)
    expected_input_digest = provenance["fixture"]["input_sha256"]
    if input_digest != expected_input_digest:
        raise RuntimeError(f"input SHA-256 changed: {input_digest} != {expected_input_digest}")

    executable = os.environ.get("LIBREOFFICE", "/usr/bin/libreoffice")
    with tempfile.TemporaryDirectory(prefix="litchi-byte-native-") as temporary_name:
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
        completed = subprocess.run(command, text=True, capture_output=True)
        if completed.returncode:
            raise RuntimeError(
                f"LibreOffice conversion failed ({completed.returncode}):\n"
                f"{completed.stdout}{completed.stderr}"
            )
        output = output_dir / "byte-functions-native.ods"
        if not output.is_file():
            raise RuntimeError("LibreOffice did not create byte-functions-native.ods")

        with zipfile.ZipFile(output) as archive:
            content = archive.read("content.xml")
        content_digest = sha256(content)
        expected_content_digest = provenance["fixture"]["recalculated_content_xml_sha256"]
        if content_digest != expected_content_digest:
            raise RuntimeError(
                f"content.xml SHA-256 changed: {content_digest} != {expected_content_digest}"
            )

        observed = extract(output)
        expected_projection = [
            {"case": item["case"], "formula": item["formula"], "native": item["native"]}
            for item in expected
        ]
        if observed != expected_projection:
            raise RuntimeError(
                "typed native observations changed:\n"
                + json.dumps({"expected": expected_projection, "observed": observed}, ensure_ascii=False, indent=2)
            )

    print(
        json.dumps(
            {
                "converter": executable,
                "fresh_profile": True,
                "temporary_tree_cleaned": True,
                "rows": len(observed),
                "ascii_rows_matching": sum(item["matches_profile"] for item in expected[:7]),
                "nonascii_divergences": sum(
                    not item["matches_profile"] for item in expected[7:]
                ),
                "content_xml_sha256": content_digest,
            },
            ensure_ascii=False,
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
