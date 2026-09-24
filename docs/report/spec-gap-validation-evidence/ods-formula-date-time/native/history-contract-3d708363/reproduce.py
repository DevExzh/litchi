#!/usr/bin/env python3
"""Reproduce the retained LibreOffice date/time compatibility capture.

The command uses a fresh headless profile and a temporary output directory.
The retained FODS input, ODS output, content.xml, and typed observations are
checked before conversion.  Deterministic rows are compared byte-for-byte at
the typed observation level; NOW, TODAY, and no-Year EASTERSUNDAY are checked
only as host observations because LibreOffice uses its own calculation clock.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "date-time-native.fods"
OUTPUT = ROOT / "recalculated.ods"
CONTENT = ROOT / "content.xml"
RESULTS = ROOT / "native-results.json"
PROVENANCE = ROOT / "provenance.json"

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
CALCEXT = "{urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0}"
FORMULA = TABLE + "formula"

VOLATILE = {
    "now.explicit_timestamp",
    "now.missing_timestamp",
    "today.explicit_timestamp",
    "today.missing_timestamp",
    "eastersunday.timestamp_after_current_easter",
    "eastersunday.missing_timestamp",
}


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def text_value(element: ET.Element) -> str:
    return "".join(element.itertext())


def typed_native(cell: ET.Element) -> dict[str, object]:
    kind = cell.attrib.get(CALCEXT + "value-type") or cell.attrib.get(OFFICE + "value-type")
    if kind == "error":
        return {
            "kind": "error",
            "raw": text_value(cell),
            "error_value": cell.attrib.get(OFFICE + "string-value", ""),
        }
    if kind == "float":
        return {"kind": "number", "value": float(cell.attrib[OFFICE + "value"])}
    if kind == "date":
        return {"kind": "date", "date_value": cell.attrib.get(OFFICE + "date-value", "")}
    if kind == "time":
        return {"kind": "time", "time_value": cell.attrib.get(OFFICE + "time-value", "")}
    if kind == "boolean":
        return {
            "kind": "logical",
            "value": cell.attrib.get(OFFICE + "boolean-value", "").lower() == "true",
        }
    if kind == "string":
        return {"kind": "text", "value": cell.attrib.get(OFFICE + "string-value", text_value(cell))}
    return {"kind": kind or "unknown", "raw": text_value(cell)}


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
            "id": text_value(cells[0]) if cells else "",
            "formula": formula_cell.attrib[FORMULA],
        }
        if typed:
            item["native"] = typed_native(formula_cell)
        rows.append(item)
    return rows


def compare_rows(expected: list[dict[str, object]], observed: list[dict[str, object]]) -> int:
    if len(expected) != len(observed):
        raise RuntimeError(f"formula row count changed: {len(expected)} != {len(observed)}")
    deterministic = 0
    for wanted, actual in zip(expected, observed):
        if wanted["id"] != actual["id"]:
            raise RuntimeError(f"formula row identity changed: {wanted['id']!r} != {actual['id']!r}")
        if not str(actual["formula"]).startswith(f"of:={wanted['id'].split('.', 1)[0].upper()}("):
            raise RuntimeError(f"formula function changed for {wanted['id']}: {actual['formula']!r}")
        if wanted["id"] in VOLATILE:
            continue
        if wanted.get("native") != actual.get("native"):
            raise RuntimeError(
                f"deterministic native observation changed for {wanted['id']}: "
                f"{wanted.get('native')!r} != {actual.get('native')!r}"
            )
        deterministic += 1
    return deterministic


def main() -> None:
    provenance = json.loads(PROVENANCE.read_text(encoding="utf-8"))
    results = json.loads(RESULTS.read_text(encoding="utf-8"))
    if provenance.get("contract_sha256") != results.get("contract_sha256"):
        raise RuntimeError("native provenance and result contract identities differ")
    if provenance.get("oracle_sha256") != results.get("oracle_sha256"):
        raise RuntimeError("native provenance and result oracle identities differ")
    if sha256(INPUT.read_bytes()) != provenance["fixture"]["input_sha256"]:
        raise RuntimeError("native input SHA-256 changed")
    if sha256(OUTPUT.read_bytes()) != provenance["fixture"]["recalculated_output_sha256"]:
        raise RuntimeError("retained recalculated ODS SHA-256 changed")
    if sha256(CONTENT.read_bytes()) != provenance["fixture"]["content_xml_sha256"]:
        raise RuntimeError("retained content.xml SHA-256 changed")
    if sha256(RESULTS.read_bytes()) != provenance["native_results_sha256"]:
        raise RuntimeError("retained native result SHA-256 changed")

    retained = results.get("observations")
    if not isinstance(retained, list) or not retained:
        raise RuntimeError("native observations are empty")
    retained_rows = [
        {"id": row["id"], "formula": "of:" + row["formula"], "native": row["native"]}
        for row in retained
    ]
    input_rows = formula_rows(INPUT.read_bytes(), typed=False)
    if len(input_rows) != len(retained_rows):
        raise RuntimeError("retained result count does not match FODS input")
    for wanted, actual in zip(retained_rows, input_rows):
        if wanted["id"] != actual["id"] or wanted["formula"] != actual["formula"]:
            raise RuntimeError(f"retained result does not match FODS input at {wanted['id']!r}")
    compare_rows(retained_rows, formula_rows(CONTENT.read_bytes(), typed=True))

    executable = os.environ.get("LIBREOFFICE", "/usr/bin/libreoffice")
    environment = dict(os.environ)
    for name in ("LANG", "LC_ALL", "LC_CTYPE", "LC_NUMERIC"):
        environment[name] = "C.UTF-8"
    with tempfile.TemporaryDirectory(prefix="litchi-date-time-native-") as temporary_name:
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
        generated = output_dir / "date-time-native.ods"
        if not generated.is_file():
            raise RuntimeError("LibreOffice did not create date-time-native.ods")
        with zipfile.ZipFile(generated) as archive:
            fresh_content = archive.read("content.xml")
        fresh_rows = formula_rows(fresh_content, typed=True)
        deterministic = compare_rows(retained_rows, fresh_rows)
        fresh_content_sha256 = sha256(fresh_content)

    counts: dict[str, int] = {}
    for row in retained:
        comparison = row["comparison"]
        counts[comparison] = counts.get(comparison, 0) + 1
    print(
        json.dumps(
            {
                "status": "verified",
                "converter": executable,
                "fresh_profile": True,
                "temporary_profile_cleaned": True,
                "rows": len(retained),
                "deterministic_rows_checked": deterministic,
                "comparison_counts": counts,
                "fresh_content_xml_sha256": fresh_content_sha256,
                "volatile_rows_host_observations": len(VOLATILE),
            },
            ensure_ascii=False,
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
