#!/usr/bin/env python3
"""Extract bounded LibreOffice paired-statistics observations.

The extractor consumes only the pinned FODS fixtures and retains the cached
formula value together with the literal cells reachable from that formula.
Formula cells in a retained source closure are rejected, so this receipt
cannot turn an upstream cached intermediate into an evaluator input.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET

from steyx_deviation import write_proof


UPSTREAM = "https://github.com/LibreOffice/core"
UPSTREAM_COMMIT = "d804d6aff49054bad1719ec3c2d136b545bbc7e7"
RAW_URL = f"https://raw.githubusercontent.com/LibreOffice/core/{UPSTREAM_COMMIT}"

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
TEXT = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"

SOURCES = {
    "CORREL": "sc/qa/unit/data/functions/statistical/fods/correl.fods",
    "COVAR": "sc/qa/unit/data/functions/statistical/fods/covar.fods",
    "PEARSON": "sc/qa/unit/data/functions/statistical/fods/pearson.fods",
    "RSQ": "sc/qa/unit/data/functions/statistical/fods/rsq.fods",
    "SLOPE": "sc/qa/unit/data/functions/statistical/fods/slope.fods",
    "INTERCEPT": "sc/qa/unit/data/functions/statistical/fods/intercept.fods",
    "STEYX": "sc/qa/unit/data/functions/statistical/fods/steyx.fods",
    "FORECAST": "sc/qa/unit/data/functions/statistical/fods/forecast.fods",
}

# Coordinates are one-based, matching the LibreOffice source files. Every
# selected formula has a finite numeric cache and a closure containing only
# literal number, logical, text, or empty cells. Array constants are retained
# where the formula has no source closure; they exercise the native parser and
# reducer without importing an upstream formula result.
SELECTIONS = {
    "CORREL": [(2, 1), (3, 1), (5, 1)],
    "COVAR": [(2, 1), (3, 1), (4, 1)],
    "PEARSON": [(2, 1), (4, 1), (6, 1), (7, 1)],
    "RSQ": [(2, 1), (3, 1), (4, 1), (5, 1), (6, 1), (7, 1)],
    "SLOPE": [(2, 1), (4, 1), (5, 1), (6, 1), (7, 1), (9, 1)],
    "INTERCEPT": [(2, 1), (3, 1), (4, 1), (5, 1), (7, 1), (9, 1), (10, 1)],
    "STEYX": [(2, 1), (3, 1), (4, 1), (5, 1), (6, 1), (7, 1)],
    "FORECAST": [(2, 1), (3, 1), (43, 1), (44, 1), (45, 1), (46, 1), (47, 1)],
}

# Keep unsupported upstream cases explicit instead of promoting their cached
# result or materializing an unbounded/formula-dependent closure.
EXCLUDED_PROFILE_VARIANCES = [
    {
        "function": "CORREL",
        "source": SOURCES["CORREL"],
        "sheet": "Sheet2",
        "row": 4,
        "column": 1,
        "formula": "of:=CORREL([.Q2:.Q101];[.R2:.R101])",
        "cached": "-0.99844495050475",
        "kind": "formula_closure",
        "reason": "The large paired closure contains formula-generated regression data; the cached result is excluded rather than importing formulas.",
    },
    {
        "function": "COVAR",
        "source": SOURCES["COVAR"],
        "sheet": "Sheet2",
        "row": 8,
        "column": 1,
        "formula": "of:=COVAR([.I7:.I12];[.J7:.J12])",
        "cached": "2.47222264607747",
        "kind": "formula_closure",
        "reason": "The paired range contains formula-generated values; the cached result is excluded rather than importing formulas.",
    },
    {
        "function": "PEARSON",
        "source": SOURCES["PEARSON"],
        "sheet": "Sheet2",
        "row": 3,
        "column": 1,
        "formula": "of:=PEARSON([.J12:.J17];[.I12:.I17])",
        "cached": "1",
        "kind": "formula_closure",
        "reason": "The paired range contains formula-generated values; the cached result is excluded rather than importing formulas.",
    },
    {
        "function": "PEARSON",
        "source": SOURCES["PEARSON"],
        "sheet": "Sheet2",
        "row": 5,
        "column": 1,
        "formula": "of:=PEARSON([.O1:.O6];[.P1:.P6])",
        "cached": "0.713771885812213",
        "kind": "formula_closure",
        "reason": "One paired range contains formula-generated values; the cached result is excluded rather than importing formulas.",
    },
    {
        "function": "INTERCEPT",
        "source": SOURCES["INTERCEPT"],
        "sheet": "Sheet2",
        "row": 6,
        "column": 1,
        "formula": "of:=INTERCEPT(sl_j;sl_i)",
        "cached": "6.75000000513889",
        "kind": "named_range_unmaterialized",
        "reason": "The named ranges have no bounded literal closure in the FODS table.",
    },
    {
        "function": "INTERCEPT",
        "source": SOURCES["INTERCEPT"],
        "sheet": "Sheet2",
        "row": 8,
        "column": 1,
        "formula": "of:=INTERCEPT([.R2:.R101];[.Q2:.Q101])",
        "cached": "1800130712.69176",
        "kind": "formula_closure",
        "reason": "The large paired closure contains formula-generated regression data; the cached result is excluded rather than importing formulas.",
    },
    {
        "function": "FORECAST",
        "source": SOURCES["FORECAST"],
        "sheet": "Sheet2",
        "row": 4,
        "column": 1,
        "formula": "of:=FORECAST([.S1:.S39];[.T1:.T39];[.S1:.S39])",
        "cached": "112.723528625461",
        "kind": "reference_argument",
        "reason": "The first forecast value is a reference argument rather than a scalar literal; it remains outside this scalar-probe receipt.",
    },
    {
        "function": "FORECAST",
        "source": SOURCES["FORECAST"],
        "sheet": "Sheet2",
        "row": 4,
        "column": 10,
        "formula": "of:=FORECAST([.J1:.J2];[.K1:.K2];[.J1:.J2])",
        "cached": None,
        "kind": "invalid_forecast",
        "reason": "LibreOffice reports an error for this two-point forecast probe.",
    },
]

HOST_NUMERIC_DEVIATIONS = [
    {
        "function": "STEYX",
        "source": SOURCES["STEYX"],
        "sheet": "Sheet2",
        "row": 3,
        "column": 1,
        "formula": "of:=STEYX([.I2:.I101];[.J2:.J101])",
        "cached": "1.60215046095659",
        "kind": "pinned_host_numeric_deviation",
        "reference": "1.6021504609570352",
        "reference_hex": "0x1.9a26885914f04p+0",
        "ulp_difference": 2005,
        "reason": "Exact Fraction arithmetic over the 99 decoded binary64 pairs followed by a 300-digit Decimal square root yields the Rust reference value; the retained LibreOffice cache is 2005 ULP lower. The independent proof is steyx_deviation.py and steyx-deviation-proof.json.",
    },
]

FUNCTIONS = set(SOURCES)
CALL_RE = re.compile(r"([A-Za-z][A-Za-z0-9_.]*)\s*\(")
REFERENCE_RE = re.compile(r"\[([^\]]+)\]")
CELL_RE = re.compile(r"^\$?([A-Za-z]+)\$?(\d+)$")
MAX_REFERENCE_CELLS = 100_000
MAX_INPUT_BYTES = 2_000_000


def column_number(letters: str) -> int:
    result = 0
    for letter in letters.upper():
        result = result * 26 + ord(letter) - ord("A") + 1
    return result


def parse_endpoint(value: str, default_sheet: str):
    value = value.strip()
    if value.startswith("."):
        value = value[1:]
    sheet = default_sheet
    if "." in value:
        prefix, value = value.rsplit(".", 1)
        sheet = prefix.strip("'").lstrip("$")
    match = CELL_RE.fullmatch(value)
    if match is None:
        raise ValueError(f"unsupported reference endpoint {value!r}")
    letters, row = match.groups()
    return sheet, int(row), column_number(letters)


def referenced_cells(formula: str, default_sheet: str):
    result = set()
    for body in REFERENCE_RE.findall(formula):
        if "~" in body:
            raise ValueError(f"multi-area reference is intentionally outside this receipt: {body!r}")
        if ":" in body:
            left, right = body.split(":", 1)
            first = parse_endpoint(left, default_sheet)
            second = parse_endpoint(right, first[0])
            if first[0] != second[0]:
                raise ValueError(f"cross-sheet range is not a bounded local closure: {body!r}")
            row_start, row_end = sorted((first[1], second[1]))
            col_start, col_end = sorted((first[2], second[2]))
            count = (row_end - row_start + 1) * (col_end - col_start + 1)
            if count > MAX_REFERENCE_CELLS:
                raise ValueError(f"reference closure is too large: {body!r}")
            result.update(
                (first[0], row, column)
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
    paragraphs = list(cell.iter(TEXT + "p"))
    nested = [node for paragraph in paragraphs for node in paragraph.iter() if node is not paragraph]
    unsupported = [node.tag for node in nested if node.tag != TEXT + "a"]
    if unsupported:
        raise ValueError(
            "unsupported nested text markup in retained string cell: "
            + ", ".join(sorted(set(unsupported)))
        )
    for anchor in (node for node in nested if node.tag == TEXT + "a"):
        if any(child is not anchor and child.tag != TEXT + "a" for child in anchor.iter()):
            raise ValueError("unsupported nested markup below retained text hyperlink")
    explicit = cell.get(OFFICE + "string-value")
    if explicit is not None:
        return explicit
    return "".join("".join(paragraph.itertext()) for paragraph in paragraphs)


def literal_cell(cell: ET.Element | None, sheet: str, row: int, column: int):
    result = {"sheet": sheet, "row": row, "column": column}
    if cell is None:
        result["type"] = "empty"
        return result
    formula = cell.get(TABLE + "formula")
    if formula:
        raise ValueError(f"reference closure reads formula cell {sheet}!R{row}C{column}: {formula}")
    value_type = cell.get(OFFICE + "value-type")
    if value_type is None:
        result["type"] = "empty"
    elif value_type == "float":
        value = cell.get(OFFICE + "value")
        if value is None:
            raise ValueError(f"number cell has no value at {sheet}!R{row}C{column}")
        number = float(value)
        if not number == number or number in (float("inf"), float("-inf")):
            raise ValueError(f"non-finite number at {sheet}!R{row}C{column}")
        result.update({"type": "number", "value": value})
    elif value_type == "boolean":
        result.update({"type": "logical", "value": cell.get(OFFICE + "boolean-value", "false")})
    elif value_type == "string":
        nested = [node.tag for paragraph in cell.iter(TEXT + "p") for node in paragraph.iter() if node is not paragraph]
        result.update({"type": "text", "value": text_value(cell)})
        if nested:
            result["nested_markup"] = sorted(set(nested))
    else:
        raise ValueError(f"unsupported referenced cell type {value_type!r} at {sheet}!R{row}C{column}")
    return result


def extract(source: Path, output: Path):
    inputs = {}
    parsed = {}
    for relative in SOURCES.values():
        data = (source / relative).read_bytes()
        if len(data) > MAX_INPUT_BYTES:
            raise ValueError(f"bounded fixture exceeds 2 MiB: {relative}")
        inputs[relative] = hashlib.sha256(data).hexdigest()
        parsed[relative] = materialize_tables(ET.fromstring(data))

    rows = []
    for function, selections in SELECTIONS.items():
        relative = SOURCES[function]
        sheets = parsed[relative]
        sheet = "Sheet2"
        table = sheets.get(sheet, {})
        for row_number, column_number_value in selections:
            cell = table.get((row_number, column_number_value))
            if cell is None:
                raise ValueError(f"selected row {function} R{row_number}C{column_number_value} is missing")
            formula = cell.get(TABLE + "formula", "")
            if not formula or cell.get(OFFICE + "value-type") not in ("float", "currency"):
                raise ValueError(f"selected row {function} R{row_number}C{column_number_value} is not a numeric formula")
            calls = [call.upper() for call in CALL_RE.findall(formula)]
            if not calls or calls[0] != function or set(calls) != {function}:
                raise ValueError(f"selected row has unsupported calls: {formula}")
            cached = cell.get(OFFICE + "value")
            if cached is None:
                raise ValueError(f"selected row {function} R{row_number}C{column_number_value} has no cache")
            references = []
            for ref_sheet, ref_row, ref_column in referenced_cells(formula, sheet):
                source_cells = sheets.get(ref_sheet)
                references.append(literal_cell(None if source_cells is None else source_cells.get((ref_row, ref_column)), ref_sheet, ref_row, ref_column))
            rows.append(
                {
                    "function": function,
                    "source": relative,
                    "sheet": sheet,
                    "row": row_number,
                    "column": column_number_value,
                    "formula": formula,
                    "cached": cached,
                    "valid_modes": ["scalar", "matrix"],
                    "cells": references,
                }
            )

    expected_count = sum(len(selection) for selection in SELECTIONS.values())
    assert len(rows) == expected_count, (len(rows), expected_count)
    assert {row["function"] for row in rows} == FUNCTIONS
    receipt = {
        "upstream": UPSTREAM,
        "raw_url": RAW_URL,
        "commit": UPSTREAM_COMMIT,
        "inputs": dict(sorted(inputs.items())),
        "source_staging": "Receipt inputs were assembled in a temporary tree from exact raw GitHub files at the pinned commit; the repository auxiliary checkout was not assumed to be pinned.",
        "byte_comparison": "Each staged input was SHA-256 checked against the exact raw GitHub file at commit before extraction.",
        "selected_observations": len(rows),
        "selected_functions": sorted(FUNCTIONS),
        "selected_modes": ["scalar", "matrix"],
        "host_numeric_deviations": HOST_NUMERIC_DEVIATIONS,
        "normative_source": {
            "archive": "3rdparty/specs/OpenDocument-v1.4-os.zip",
            "archive_sha256": "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4",
            "member": "part4-formula/OpenDocument-v1.4-os-part4-formula.html",
            "member_sha256": "ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1",
            "sections": ["6.18.38 INTERCEPT", "6.18.41 LINEST"],
            "intercept_text": "INTERCEPT returns the intercept (a) calculated as described in 6.18.41 for the function call LINEST(Data_Y, Data_X, FALSE()).",
            "linest_const_text": "Const: If set to FALSE, the model constant a is equal to 0.",
        },
        "excluded_profile_variances": EXCLUDED_PROFILE_VARIANCES,
        "coverage_notes": {
            "source_format": "All selected sources are upstream FODS fixtures; no conversion tool or recalculation step is used.",
            "closure": "Each selected reference closure contains only literal number, logical, text, or empty cells. Formula cells, dates, named ranges, and nested formula dependencies remain explicit exclusions.",
            "cache": "The retained cached value is used only as an upstream observation. The Rust test reconstructs the literal closure and evaluates the formula in both scalar and matrix modes. Forty-one rows use the ordinary 1e-13 relative comparison; the retained STEYX I2:I101/J2:J101 row is a declared host numeric deviation with an independent exact reference proof.",
            "profile": "Host-specific date/error coercions and unsupported named-range or nested-expression behavior are not silently promoted to the evaluator contract.",
            "intercept_normative_ambiguity": "ODF 1.4 §6.18.38 literally describes INTERCEPT as LINEST(Data_Y, Data_X, FALSE()), while §6.18.41 defines Const=FALSE as a zero intercept. The selected native caches show ordinary nonzero intercepts (for example 6.75000000513889 and -9). This receipt records LibreOffice observations only; the normative contract must choose and document the resolution.",
        },
    }
    root = Path(__file__).resolve().parent
    expected_path = root / "provenance.json"
    if expected_path.is_file():
        expected = json.loads(expected_path.read_text())
        assert inputs == expected["inputs"], "inputs differ from pinned upstream fixture hashes"
    output.mkdir(parents=True, exist_ok=True)
    (output / "cached-results.json").write_text(json.dumps(rows, indent=2) + "\n")
    (output / "provenance.json").write_text(json.dumps(receipt, indent=2) + "\n")
    write_proof(output / "cached-results.json", output / "steyx-deviation-proof.json")


def main():
    if len(sys.argv) < 2:
        raise SystemExit("usage: extract.py /path/to/staged-source [output-directory]")
    source = Path(sys.argv[1]).resolve()
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent
    extract(source, output)


if __name__ == "__main__":
    main()
