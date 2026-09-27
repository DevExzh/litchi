#!/usr/bin/env python3
"""Replay the bounded negative controls used by the 0778 oracle preflight."""

from __future__ import annotations

import importlib.util
import json
import re
import sys
from dataclasses import replace
from pathlib import Path


EVIDENCE = Path(__file__).resolve().parent
PACKET = EVIDENCE.parent
MANIFEST = PACKET / "artifacts-0" / "manifest.json"
ORACLE_PATH = PACKET / "oracle.py"

spec = importlib.util.spec_from_file_location("ordinary_save_oracle", ORACLE_PATH)
if spec is None or spec.loader is None:
    raise RuntimeError(f"cannot load oracle: {ORACLE_PATH}")
oracle = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = oracle
spec.loader.exec_module(oracle)


manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
root = PACKET


def load_case(case_id: str):
    case = next(row for row in manifest["cases"] if row["case_id"] == case_id)
    source = oracle.ZipArchive.read(
        root / "artifacts-0" / case["source_archive"]["path"], f"{case_id}.source"
    )
    output_path = next(
        row["output"]["path"] for row in case["policy_outputs"] if row["policy"] == "default"
    )
    output = oracle.ZipArchive.read(root / "artifacts-0" / output_path, f"{case_id}.output")
    return case, source, output


def output_with_member(output, member: str, data: bytes):
    members = dict(output.members)
    members[member] = data
    return replace(output, members=members)


def expect_rejected(label: str, callback) -> None:
    try:
        callback()
    except oracle.OracleError as exc:
        print(f"{label}: REJECTED: {exc}")
    else:
        raise RuntimeError(f"{label}: mutation was accepted")


conditional, conditional_source, conditional_output = load_case("real-002-xlsx")
conditional_closure = oracle.fixed_mutation_closure(
    "xlsx", conditional_source, conditional["edit_target"], "negative.conditional"
)
content_types = conditional_output.members["[Content_Types].xml"]
printer_type = b"ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\""


def conditional_output_with_content_types(data: bytes):
    return output_with_member(conditional_output, "[Content_Types].xml", data)


expect_rejected(
    "printer-override-content-type",
    lambda: oracle.check_xlsx_owned_parts(
        conditional_source,
        conditional_output_with_content_types(
            content_types.replace(printer_type, b'ContentType="application/x-evil"', 1)
        ),
        "negative.printer-override-content-type",
    ),
)
printer_pattern = rb'<Override\s+PartName="/xl/printerSettings/printerSettings1\.bin"\s+ContentType="[^"]+"/>'
expect_rejected(
    "missing-retained-printer-override",
    lambda: oracle.check_xlsx_owned_parts(
        conditional_source,
        conditional_output_with_content_types(re.sub(printer_pattern, b"", content_types, count=1)),
        "negative.missing-retained-printer-override",
    ),
)
expect_rejected(
    "extra-content-type-declaration",
    lambda: oracle.check_xlsx_owned_parts(
        conditional_source,
        conditional_output_with_content_types(
            content_types.replace(
                b"</Types>", b'<Override PartName="/xl/other.bin" ContentType="application/x-evil"/></Types>', 1
            )
        ),
        "negative.extra-content-type-declaration",
    ),
)
expect_rejected(
    "unrelated-default-content-type",
    lambda: oracle.check_xlsx_owned_parts(
        conditional_source,
        conditional_output_with_content_types(
            content_types.replace(
                b'ContentType="application/vnd.openxmlformats-package.relationships+xml"',
                b'ContentType="application/x-evil"',
                1,
            )
        ),
        "negative.unrelated-default-content-type",
    ),
)

generated, generated_source, generated_output = load_case("generated-xlsx-medium")
generated_closure = oracle.fixed_mutation_closure(
    "xlsx", generated_source, generated["edit_target"], "negative.generated"
)
workbook = generated_output.members["xl/workbook.xml"]
workbook_rels = generated_output.members["xl/_rels/workbook.xml.rels"]
generated_content_types = generated_output.members["[Content_Types].xml"]


expect_rejected(
    "workbook-attribute",
    lambda: oracle.check_xlsx_owned_parts(
        generated_source,
        output_with_member(generated_output, "xl/workbook.xml", workbook.replace(b"<sheets>", b'<bookViews activeTab="1"/><sheets>', 1)),
        "negative.workbook-attribute",
    ),
)
expect_rejected(
    "unrelated-workbook-relationship",
    lambda: oracle.check_preservation(
        generated_source,
        output_with_member(
            generated_output,
            "xl/_rels/workbook.xml.rels",
            re.sub(rb"<Relationship\b[^>]*/>", b"", workbook_rels, count=1),
        ),
        generated_closure,
        "negative.unrelated-workbook-relationship",
    ),
)
expect_rejected(
    "unrelated-content-type",
    lambda: oracle.check_xlsx_owned_parts(
        generated_source,
        output_with_member(
            generated_output,
            "[Content_Types].xml",
            re.sub(rb"<Default\b[^>]*/>", b"", generated_content_types, count=1),
        ),
        "negative.unrelated-content-type",
    ),
)
expect_rejected(
    "calcPr-flag",
    lambda: oracle.check_xlsx_owned_parts(
        generated_source,
        output_with_member(
            generated_output,
            "xl/workbook.xml",
            workbook.replace(b'fullCalcOnLoad="true"', b'fullCalcOnLoad="false"', 1),
        ),
        "negative.calcPr-flag",
    ),
)
print("baseline printer-settings normalization: ACCEPTED")
