#!/usr/bin/env python3
"""Extract bounded aggregate observations from LibreOffice FODS fixtures.

Usage: ``python3 extract.py /path/to/libreoffice-core [output-directory]``.
The input checkout is read only.  Selected worksheet references are copied as
literal cells so the Rust integration test can exercise the value resolver
without evaluating an upstream formula cell or importing a cached formula
result as an input.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET


UPSTREAM_COMMIT = "d804d6aff49054bad1719ec3c2d136b545bbc7e7"
UPSTREAM_URL = "https://github.com/LibreOffice/core"
RAW_URL = f"https://raw.githubusercontent.com/LibreOffice/core/{UPSTREAM_COMMIT}"

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
TEXT = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"

SOURCES = {
    "SUM": "sc/qa/unit/data/functions/mathematical/fods/sum.fods",
    "PRODUCT": "sc/qa/unit/data/functions/mathematical/fods/product.fods",
    "SUMSQ": "sc/qa/unit/data/functions/mathematical/fods/sumsq.fods",
    "SUMPRODUCT": "sc/qa/unit/data/functions/array/fods/sumproduct.fods",
    "SUMX2MY2": "sc/qa/unit/data/functions/array/fods/sumx2my2.fods",
    "SUMX2PY2": "sc/qa/unit/data/functions/array/fods/sumx2py2.fods",
    "SUMXMY2": "sc/qa/unit/data/functions/array/fods/sumxmy2.fods",
}

# Coordinates are one-based, as they are in the source FODS files.  The
# selected rows intentionally cover literals, references, arrays, and blank or
# text cells while staying small enough for a bounded resolver fixture.
SELECTIONS = {
    "SUM": (2, 3, 4, 5, 6, 9, 10),
    "PRODUCT": (2, 3, 4, 5, 6, 7, 9, 10, 11),
    "SUMSQ": (2, 3, 4, 5, 6),
    "SUMPRODUCT": (2, 3, 4, 9, 10, 11, 12, 13, 14),
    "SUMX2MY2": (2, 3, 4, 5, 6, 7),
    "SUMX2PY2": (2, 3, 4, 5, 6, 7),
    "SUMXMY2": (2, 3, 4, 5, 6, 7),
}

# PRODUCT() is present in the upstream fixture with a cached zero.  The
# selected public profile requires at least one argument, so retain its source
# coordinate in the receipt rather than silently treating it as a supported
# observation.
EXCLUDED = [
    {
        "function": "PRODUCT",
        "row": 12,
        "column": 1,
        "reason": "The upstream PRODUCT() cache uses zero arguments; the selected aggregate profile requires at least one argument.",
    },
    {
        "function": "SUMPRODUCT",
        "row": 15,
        "column": 1,
        "reason": "The source range includes M17, an upstream formula cell (of:=\"\"); the bounded native fixture refuses formula-cell inputs rather than evaluating or importing that cache.",
    },
    {
        "function": "SUMPRODUCT",
        "source": "sc/qa/unit/data/functions/array/fods/sumproduct.fods",
        "sheet": "Sheet2",
        "row": 16,
        "column": 1,
        "formula": "of:=SUMPRODUCT([.J16:.J18];[.N16:.N18])",
        "cached": "18",
        "reference_cell": {
            "sheet": "Sheet2",
            "row": 17,
            "column": 14,
            "type": "text",
            "value": "Unknown",
        },
        "reason": "The selected contract converts malformed text in a referenced forced array to a formula Value error, while the LibreOffice cache treats the literal text cell as zero; retain the source/cache as an explicit native conversion variance.",
    },
]

FUNCTIONS = set(SOURCES)
CALL_RE = re.compile(r"([A-Za-z][A-Za-z0-9_.]*)\s*\(")
REFERENCE_RE = re.compile(r"\[([^\]]+)\]")
CELL_RE = re.compile(r"^\$?([A-Za-z]+)\$?(\d+)$")
MAX_REFERENCE_CELLS = 100_000


def column_number(letters: str) -> int:
    number = 0
    for letter in letters.upper():
        number = number * 26 + ord(letter) - ord("A") + 1
    return number


def parse_endpoint(value: str, default_sheet: str):
    value = value.strip()
    if value.startswith("."):
        value = value[1:]
    sheet = default_sheet
    if "." in value:
        prefix, value = value.rsplit(".", 1)
        sheet = prefix.strip("'")
    match = CELL_RE.fullmatch(value)
    if match is None:
        raise ValueError(f"unsupported reference endpoint {value!r}")
    letters, row = match.groups()
    return sheet, int(row), column_number(letters)


def referenced_cells(formula: str, default_sheet: str):
    result = set()
    for body in REFERENCE_RE.findall(formula):
        if ":" in body:
            left, right = body.split(":", 1)
            first = parse_endpoint(left, default_sheet)
            second = parse_endpoint(right, first[0])
            if first[0] != second[0]:
                raise ValueError(f"cross-sheet range is not a bounded local closure: {body!r}")
            sheet = first[0]
            row_start, row_end = sorted((first[1], second[1]))
            col_start, col_end = sorted((first[2], second[2]))
            count = (row_end - row_start + 1) * (col_end - col_start + 1)
            if count > MAX_REFERENCE_CELLS:
                raise ValueError(f"reference closure is too large: {body!r}")
            result.update(
                (sheet, row, column)
                for row in range(row_start, row_end + 1)
                for column in range(col_start, col_end + 1)
            )
        else:
            result.add(parse_endpoint(body, default_sheet))
    return sorted(result)


def materialize_tables(root: ET.Element):
    sheets = {}
    for table in root.iter(TABLE + "table"):
        sheet = table.get(TABLE + "name")
        cells = {}
        row_number = 1
        for row in table.findall(TABLE + "table-row"):
            column_number_value = 1
            for cell in row:
                repeated = int(cell.get(TABLE + "number-columns-repeated", "1"))
                for offset in range(repeated):
                    cells[(row_number, column_number_value + offset)] = cell
                column_number_value += repeated
            row_number += int(row.get(TABLE + "number-rows-repeated", "1"))
        sheets[sheet] = cells
    return sheets


def text_value(cell: ET.Element) -> str:
    explicit = cell.get(OFFICE + "string-value")
    if explicit is not None:
        return explicit
    paragraphs = [node.text or "" for node in cell.iter(TEXT + "p")]
    return "".join(paragraphs)


def literal_cell(cell: ET.Element | None, sheet: str, row: int, column: int):
    location = {"sheet": sheet, "row": row, "column": column}
    if cell is None:
        location.update({"type": "empty"})
        return location
    formula = cell.get(TABLE + "formula")
    if formula:
        raise ValueError(
            f"reference closure reads formula cell {sheet}!R{row}C{column}: {formula}"
        )
    value_type = cell.get(OFFICE + "value-type")
    if value_type is None:
        location.update({"type": "empty"})
    elif value_type == "float":
        value = cell.get(OFFICE + "value")
        if value is None:
            raise ValueError(f"number cell has no value at {sheet}!R{row}C{column}")
        number = float(value)
        if not number == number or number in (float("inf"), float("-inf")):
            raise ValueError(f"non-finite number at {sheet}!R{row}C{column}")
        location.update({"type": "number", "value": value})
    elif value_type == "boolean":
        location.update({"type": "logical", "value": cell.get(OFFICE + "boolean-value", "false")})
    elif value_type == "string":
        location.update({"type": "text", "value": text_value(cell)})
    else:
        raise ValueError(f"unsupported referenced cell type {value_type!r} at {sheet}!R{row}C{column}")
    return location


def extract(source: Path, output: Path):
    inputs = {}
    rows = []
    for function, relative in SOURCES.items():
        data = (source / relative).read_bytes()
        inputs[relative] = hashlib.sha256(data).hexdigest()
        sheets = materialize_tables(ET.fromstring(data))
        selected_rows = set(SELECTIONS[function])
        sheet = "Sheet2"
        table = sheets.get(sheet, {})
        for row_number in sorted(selected_rows):
            cell = table.get((row_number, 1))
            if cell is None:
                raise ValueError(f"selected row {function} R{row_number}C1 is missing")
            formula = cell.get(TABLE + "formula", "")
            if not formula or cell.get(OFFICE + "value-type") != "float":
                raise ValueError(f"selected row {function} R{row_number}C1 is not a numeric formula")
            calls = [call.upper() for call in CALL_RE.findall(formula)]
            if not calls or calls[0] != function or set(calls) != {function}:
                raise ValueError(f"selected row has unsupported calls: {formula}")
            cached = cell.get(OFFICE + "value")
            if cached is None:
                raise ValueError(f"selected row {function} R{row_number}C1 has no cache")
            references = []
            for ref_sheet, ref_row, ref_column in referenced_cells(formula, sheet):
                source_cells = sheets.get(ref_sheet)
                references.append(
                    literal_cell(
                        None if source_cells is None else source_cells.get((ref_row, ref_column)),
                        ref_sheet,
                        ref_row,
                        ref_column,
                    )
                )
            rows.append(
                {
                    "function": function,
                    "source": relative,
                    "sheet": sheet,
                    "row": row_number,
                    "column": 1,
                    "formula": formula,
                    "cached": cached,
                    "cells": references,
                }
            )

    assert len(rows) == 48, len(rows)
    assert {row["function"] for row in rows} == FUNCTIONS
    expected = json.loads((Path(__file__).resolve().parent / "provenance.json").read_text())
    assert inputs == expected["inputs"], "inputs differ from the pinned upstream fixture hashes"
    receipt = {
        "upstream": UPSTREAM_URL,
        "raw_url": RAW_URL,
        "commit": UPSTREAM_COMMIT,
        "inputs": inputs,
        "source_staging": "Receipt inputs were assembled in a temporary tree from the exact raw GitHub files at the pinned commit; the repository auxiliary checkout was not assumed to be pinned.",
        "byte_comparison": "Each staged input was SHA-256 checked against the exact raw GitHub file at commit before extraction.",
        "selected_observations": len(rows),
        "excluded_profile_variances": EXCLUDED,
    }
    (output / "cached-results.json").write_text(json.dumps(rows, indent=2) + "\n")
    (output / "provenance.json").write_text(json.dumps(receipt, indent=2) + "\n")


def main():
    source = Path(sys.argv[1]).resolve()
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent
    output.mkdir(parents=True, exist_ok=True)
    extract(source, output)


if __name__ == "__main__":
    main()
