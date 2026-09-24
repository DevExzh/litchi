#!/usr/bin/env python3
"""Reproduce the retained LibreOffice financial reducer capture.

The fixture expands the compact 2,000-zero NPV descriptor to a bounded
one-column array literal and maps the MIRR text/empty/logical vector to a
five-cell reference so the empty cell remains representable in LibreOffice.
Those source mappings are retained in ``native-results.json``. Every other
formula keeps the corpus array or reference geometry verbatim.
"""

from __future__ import annotations

import hashlib
import json
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

sys.dont_write_bytecode = True


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "financial-reducers-native.fods"
OUTPUT = ROOT / "recalculated.ods"
CONTENT = ROOT / "content.xml"
RESULTS = ROOT / "native-results.json"
PROVENANCE = ROOT / "provenance.json"
CONTRACT = ROOT.parent / "contract.md"
ORACLE = ROOT / "oracle-reducer-vectors.json"

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
CALCEXT = "{urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0}"
FORMULA = TABLE + "formula"


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def text_value(cell: ET.Element) -> str:
    return "".join(cell.itertext())


def typed_native(cell: ET.Element) -> dict[str, object]:
    attrs = cell.attrib
    kind = attrs.get(CALCEXT + "value-type") or attrs.get(OFFICE + "value-type")
    if kind in {"float", "percentage", "currency"}:
        return {"kind": "number", "value": float(attrs[OFFICE + "value"])}
    if kind == "error":
        return {"kind": "error", "raw": text_value(cell)}
    if kind == "boolean":
        return {"kind": "logical", "value": attrs.get(OFFICE + "boolean-value", "").lower() == "true"}
    if kind == "string":
        return {"kind": "text", "value": attrs.get(OFFICE + "string-value", text_value(cell))}
    return {"kind": kind or "unknown", "raw": text_value(cell)}


def formula_rows(content: bytes, *, typed: bool) -> list[dict[str, object]]:
    root = ET.fromstring(content)
    rows: list[dict[str, object]] = []
    for row in root.iter(TABLE + "table-row"):
        cells = list(row.findall(TABLE + "table-cell"))
        formula_cells = [cell for cell in cells if FORMULA in cell.attrib]
        if not formula_cells:
            continue
        if len(cells) != 5 or len(formula_cells) != 1:
            raise RuntimeError("financial reducer fixture row shape changed")
        formula_cell = formula_cells[0]
        item: dict[str, object] = {
            "id": text_value(cells[0]),
            "function": text_value(cells[1]),
            "formula": formula_cell.attrib[FORMULA],
            "source_formula": text_value(cells[3]),
        }
        if typed:
            item["native"] = typed_native(formula_cell)
        rows.append(item)
    return rows


def formula_function(formula: str) -> str:
    match = re.match(r"of:=([A-Za-z][A-Za-z0-9_.]*)\(", formula)
    if match is None:
        raise RuntimeError(f"unexpected formula syntax: {formula!r}")
    return match.group(1).upper()


def assert_rows(expected: list[dict[str, object]], observed: list[dict[str, object]]) -> None:
    if len(expected) != len(observed):
        raise RuntimeError(f"formula row count changed: {len(expected)} != {len(observed)}")
    for wanted, actual in zip(expected, observed):
        if wanted["id"] != actual["id"] or wanted["function"] != actual["function"]:
            raise RuntimeError(f"formula row identity changed for {wanted['id']!r}")
        if wanted["source_formula"] != actual["source_formula"]:
            raise RuntimeError(f"source formula changed for {wanted['id']!r}")
        if formula_function(str(actual["formula"])) != str(wanted["function"]).upper():
            raise RuntimeError(f"formula function changed for {wanted['id']!r}: {actual['formula']!r}")
        if wanted.get("native") != actual.get("native"):
            raise RuntimeError(
                f"typed native observation changed for {wanted['id']!r}: "
                f"{wanted.get('native')!r} != {actual.get('native')!r}"
            )


def main() -> None:
    provenance = json.loads(PROVENANCE.read_text(encoding="utf-8"))
    retained = json.loads(RESULTS.read_text(encoding="utf-8"))
    corpus = json.loads(ORACLE.read_text(encoding="utf-8"))

    if sha256(CONTRACT.read_bytes()) != provenance["contract_sha256"]:
        raise RuntimeError("financial contract SHA-256 changed")
    if sha256(ORACLE.read_bytes()) != provenance["oracle_sha256"]:
        raise RuntimeError("financial reducer corpus SHA-256 changed")
    if retained.get("contract_sha256") != provenance["contract_sha256"]:
        raise RuntimeError("native result and provenance contract identities differ")
    if retained.get("oracle_sha256") != provenance["oracle_sha256"]:
        raise RuntimeError("native result and provenance corpus identities differ")
    if retained.get("rows") != len(corpus.get("vectors", [])):
        raise RuntimeError("native result row count is not bound to the corpus")
    if sha256(INPUT.read_bytes()) != provenance["fixture"]["input_sha256"]:
        raise RuntimeError("financial reducer native input SHA-256 changed")
    if sha256(OUTPUT.read_bytes()) != provenance["fixture"]["recalculated_output_sha256"]:
        raise RuntimeError("retained financial reducer ODS SHA-256 changed")
    if sha256(CONTENT.read_bytes()) != provenance["fixture"]["content_xml_sha256"]:
        raise RuntimeError("retained financial reducer content.xml SHA-256 changed")
    if sha256(RESULTS.read_bytes()) != provenance["native_results_sha256"]:
        raise RuntimeError("retained financial reducer result SHA-256 changed")
    capture_log = ROOT / provenance["capture_log"]
    if sha256(capture_log.read_bytes()) != provenance["capture_log_sha256"]:
        raise RuntimeError("retained financial reducer capture log SHA-256 changed")

    retained_rows = retained.get("observations")
    if not isinstance(retained_rows, list) or len(retained_rows) != len(corpus["vectors"]):
        raise RuntimeError("native reducer observations are incomplete")
    input_rows = formula_rows(INPUT.read_bytes(), typed=False)
    if len(input_rows) != len(retained_rows):
        raise RuntimeError("retained reducer results do not match FODS input row count")
    for wanted, actual in zip(retained_rows, input_rows):
        if wanted["id"] != actual["id"] or wanted["function"] != actual["function"]:
            raise RuntimeError(f"retained result does not match FODS input at {wanted['id']!r}")
        if wanted["source_formula"] != actual["source_formula"]:
            raise RuntimeError(f"retained source mapping changed at {wanted['id']!r}")
        if formula_function(str(actual["formula"])) != str(wanted["function"]).upper():
            raise RuntimeError(f"input formula changed for {wanted['id']!r}")

    retained_typed = [
        {
            "id": item["id"],
            "function": item["function"],
            "formula": item["formula"],
            "source_formula": item["source_formula"],
            "native": item["native"],
        }
        for item in retained_rows
    ]
    assert_rows(retained_typed, formula_rows(CONTENT.read_bytes(), typed=True))

    executable = os.environ.get("LIBREOFFICE", "/usr/bin/libreoffice")
    environment = dict(os.environ)
    for name in ("LANG", "LC_ALL", "LC_CTYPE", "LC_NUMERIC"):
        environment[name] = "C.UTF-8"
    with tempfile.TemporaryDirectory(prefix="litchi-financial-reducers-native-") as temporary_name:
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
        completed = subprocess.run(command, env=environment, text=True, capture_output=True)
        if completed.returncode:
            raise RuntimeError(
                f"LibreOffice conversion failed ({completed.returncode}):\n"
                f"{completed.stdout}{completed.stderr}"
            )
        generated = output_dir / "financial-reducers-native.ods"
        if not generated.is_file():
            raise RuntimeError("LibreOffice did not create financial-reducers-native.ods")
        with zipfile.ZipFile(generated) as archive:
            fresh_content = archive.read("content.xml")
        fresh_rows = formula_rows(fresh_content, typed=True)
        assert_rows(retained_typed, fresh_rows)
        fresh_content_sha256 = sha256(fresh_content)

    print(
        json.dumps(
            {
                "status": "verified",
                "converter": executable,
                "fresh_profile": True,
                "temporary_profile_cleaned": True,
                "rows": len(retained_rows),
                "comparison_counts": retained.get("comparison_counts", {}),
                "source_transformations": retained.get("source_transformations", []),
                "fresh_content_xml_sha256": fresh_content_sha256,
                "native_capture_is_compatibility_only": True,
            },
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
