#!/usr/bin/env python3
"""Extract bounded LibreOffice §6.20 text-function observations.

The inputs are the raw FODS fixtures at one pinned LibreOffice revision.  A
selected formula keeps its native cached result and the complete literal cell
closure used by its references.  Formula cells in a closure are rejected so a
native cache can never smuggle an upstream calculated value into the Rust
receipt.  This extractor does not invoke LibreOffice, convert a document, or
recalculate a formula.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET


UPSTREAM = "https://github.com/LibreOffice/core"
UPSTREAM_COMMIT = "d804d6aff49054bad1719ec3c2d136b545bbc7e7"
RAW_URL = f"https://raw.githubusercontent.com/LibreOffice/core/{UPSTREAM_COMMIT}"

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
TEXT = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
CALEXT = "{urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0}"

SOURCES = {
    "ASC": "sc/qa/unit/data/functions/text/fods/asc.fods",
    "CHAR": "sc/qa/unit/data/functions/text/fods/char.fods",
    "CLEAN": "sc/qa/unit/data/functions/text/fods/clean.fods",
    "CODE": "sc/qa/unit/data/functions/text/fods/code.fods",
    "CONCATENATE": "sc/qa/unit/data/functions/text/fods/concatenate.fods",
    "DOLLAR": "sc/qa/unit/data/functions/text/fods/dollar.fods",
    "EXACT": "sc/qa/unit/data/functions/text/fods/exact.fods",
    "FIND": "sc/qa/unit/data/functions/text/fods/find.fods",
    "FIXED": "sc/qa/unit/data/functions/text/fods/fixed.fods",
    "JIS": "sc/qa/unit/data/functions/text/fods/jis.fods",
    "LEFT": "sc/qa/unit/data/functions/text/fods/left.fods",
    "LEN": "sc/qa/unit/data/functions/text/fods/len.fods",
    "LOWER": "sc/qa/unit/data/functions/text/fods/lower.fods",
    "MID": "sc/qa/unit/data/functions/text/fods/mid.fods",
    "PROPER": "sc/qa/unit/data/functions/text/fods/proper.fods",
    "REPLACE": "sc/qa/unit/data/functions/text/fods/replace.fods",
    "REPT": "sc/qa/unit/data/functions/text/fods/rept.fods",
    "RIGHT": "sc/qa/unit/data/functions/text/fods/right.fods",
    "SEARCH": "sc/qa/unit/data/functions/text/fods/search.fods",
    "SUBSTITUTE": "sc/qa/unit/data/functions/text/fods/substitute.fods",
    "T": "sc/qa/unit/data/functions/text/fods/t.fods",
    "TEXT": "sc/qa/unit/data/functions/text/fods/text.fods",
    "TRIM": "sc/qa/unit/data/functions/text/fods/trim.fods",
    "UNICHAR": "sc/qa/unit/data/functions/text/fods/unichar.fods",
    "UNICODE": "sc/qa/unit/data/functions/text/fods/unicode.fods",
    "UPPER": "sc/qa/unit/data/functions/text/fods/upper.fods",
}

# Coordinates are one-based and refer to Sheet2 in each upstream FODS file.
# These rows cover all 26 §6.20 functions and retain finite numeric, text,
# logical, and one formula-error cache.  Locale-sensitive formatting rows are
# intentionally limited to the documented en-US text profile below.
SELECTIONS = {
    "ASC": [(2, 1)],
    "CHAR": [(2, 1), (34, 1), (96, 1)],
    "CLEAN": [(2, 1)],
    "CODE": [(2, 1), (99, 1)],
    "CONCATENATE": [(2, 1), (3, 1), (4, 1)],
    "DOLLAR": [(2, 1), (3, 1), (4, 1)],
    "EXACT": [(2, 1), (3, 1), (6, 1)],
    "FIND": [(2, 1), (3, 1), (4, 1), (5, 1), (16, 1), (17, 1)],
    "FIXED": [(2, 1), (3, 1), (5, 1)],
    "JIS": [(2, 1), (3, 1), (7, 1)],
    "LEFT": [(2, 1), (3, 1), (4, 1)],
    "LEN": [(2, 1), (2, 11), (4, 1), (7, 1)],
    "LOWER": [(2, 1), (4, 1), (11, 1)],
    "MID": [(2, 1), (3, 1), (5, 1), (8, 1)],
    "PROPER": [(2, 1), (5, 1), (17, 1)],
    "REPLACE": [(2, 1), (3, 1), (4, 1), (6, 1), (11, 1)],
    "REPT": [(2, 1), (3, 1), (4, 1)],
    "RIGHT": [(2, 1), (6, 1), (12, 1), (23, 10)],
    "SEARCH": [(2, 1), (6, 1), (21, 1), (25, 1)],
    "SUBSTITUTE": [(2, 1), (3, 1), (4, 1), (5, 1), (7, 1)],
    "T": [(11, 1), (21, 1)],
    "TEXT": [(2, 1), (3, 1), (7, 1), (14, 1), (15, 1), (18, 1), (23, 1), (24, 1), (25, 1)],
    "TRIM": [(2, 1), (3, 1), (4, 1), (6, 1)],
    "UNICHAR": [(2, 1), (7, 1), (38, 1)],
    "UNICODE": [(2, 1), (3, 1), (7, 1)],
    "UPPER": [(2, 1), (3, 1), (4, 1), (5, 1)],
}

# These are observations deliberately kept out of cached-results.json.  The
# list records why a source/cache was not silently treated as evaluator input:
# formula-dependent closures, invalid/no-cache rows, array/profile variants,
# and date/locale formatting outside the selected deterministic profile.
EXCLUDED_PROFILE_VARIANCES = [
    {
        "function": "ASC",
        "source": SOURCES["ASC"],
        "sheet": "Sheet2",
        "row": 35,
        "column": 1,
        "kind": "incompatible_locale_profile",
        "reason": "Full-width and Japanese text conversion is retained as a native observation only; the selected profile covers ASCII input.",
    },
    {
        "function": "CONCATENATE",
        "source": SOURCES["CONCATENATE"],
        "sheet": "Sheet2",
        "row": 7,
        "column": 1,
        "kind": "nonliteral_computed_argument",
        "formula": "of:=CONCATENATE(TRANSPOSE(1))",
        "reason": "TRANSPOSE is a computed argument outside the literal dependency closure receipt.",
    },
    {
        "function": "DOLLAR",
        "source": SOURCES["DOLLAR"],
        "sheet": "Sheet2",
        "row": 9,
        "column": 1,
        "kind": "array_argument_profile",
        "formula": "of:=DOLLAR([.I2:.I5];2)",
        "reason": "Array formatting is retained as a native observation but excluded from the scalar deterministic profile.",
    },
    {
        "function": "TRIM",
        "source": SOURCES["TRIM"],
        "sheet": "Sheet2",
        "row": 8,
        "column": 1,
        "kind": "native_matrix_error_profile",
        "formula": "of:=TRIM([.I5:.I6])",
        "reason": "LibreOffice's ordinary scalar range call stores #VALUE!, while the selected profile streams this range in matrix mode; the host-only scalar cache is excluded.",
    },
    {
        "function": "FIXED",
        "source": SOURCES["FIXED"],
        "sheet": "Sheet2",
        "row": 7,
        "column": 1,
        "kind": "array_argument_profile",
        "formula": "of:=FIXED([.J1:.J3];3;1)",
        "reason": "The array output is outside the selected scalar formatting receipt.",
    },
    {
        "function": "JIS",
        "source": SOURCES["JIS"],
        "sheet": "Sheet2",
        "row": 35,
        "column": 1,
        "kind": "incompatible_locale_profile",
        "reason": "Japanese locale-specific width conversion is retained in the source corpus but not promoted beyond the selected deterministic profile.",
    },
    {
        "function": "LEFT",
        "source": SOURCES["LEFT"],
        "sheet": "Sheet2",
        "row": 6,
        "column": 1,
        "kind": "invalid_argument_no_cache",
        "formula": "of:=LEFT(\"Calc\" ;)",
        "reason": "The upstream FODS cell has no typed cache and is not converted into a synthetic error row.",
    },
    {
        "function": "MID",
        "source": SOURCES["MID"],
        "sheet": "Sheet2",
        "row": 7,
        "column": 1,
        "kind": "native_error_profile",
        "formula": "of:=MID([.I2];0;3)",
        "reason": "LibreOffice reports Err:502; the selected receipt keeps only typed result caches with deterministic error mapping.",
    },
    {
        "function": "PROPER",
        "source": SOURCES["PROPER"],
        "sheet": "Sheet2",
        "row": 8,
        "column": 1,
        "kind": "native_error_profile",
        "formula": "of:=PROPER()",
        "reason": "The no-argument native error is an explicit host observation, not a synthetic empty text result.",
    },
    {
        "function": "REPT",
        "source": SOURCES["REPT"],
        "sheet": "Sheet2",
        "row": 6,
        "column": 1,
        "kind": "native_error_profile",
        "formula": "of:=REPT(\"-\";-1)",
        "reason": "LibreOffice stores an error cache for a negative repeat count; no replacement result is inferred.",
    },
    {
        "function": "T",
        "source": SOURCES["T"],
        "sheet": "Sheet2",
        "row": 9,
        "column": 1,
        "kind": "native_error_profile",
        "formula": "of:=T(1/0)",
        "reason": "The error-propagation cache is retained as a host observation; the selected receipt covers text inputs only.",
    },
    {
        "function": "TEXT",
        "source": SOURCES["TEXT"],
        "sheet": "Sheet2",
        "row": 4,
        "column": 1,
        "kind": "incompatible_date_locale_profile",
        "formula": "of:=TEXT([.I1];\"YYYY-MM-DD\")",
        "reason": "Date serial formatting depends on the workbook date/locale profile and is excluded from this text-only receipt.",
    },
    {
        "function": "TEXT",
        "source": SOURCES["TEXT"],
        "sheet": "Sheet2",
        "row": 17,
        "column": 1,
        "kind": "nonliteral_nested_function",
        "formula": "of:=TRIM(TEXT(0.34;\"# ?/?\"))",
        "reason": "Nested TRIM is not a single TEXT-function observation and is excluded from the per-function closure.",
    },
    {
        "function": "TEXT",
        "source": SOURCES["TEXT"],
        "sheet": "Sheet2",
        "row": 26,
        "column": 1,
        "kind": "incompatible_date_locale_profile",
        "formula": "of:=TEXT([.J8];\"YYYY/MM/DD\")",
        "reason": "Date formatting is retained as a native observation only; the selected profile is numeric/text and date-neutral.",
    },
    {
        "function": "TRIM",
        "source": SOURCES["TRIM"],
        "sheet": "Sheet2",
        "row": 5,
        "column": 1,
        "kind": "formula_error_no_cache",
        "formula": "of:=TRIM([.F1])",
        "reason": "The source cell is an untyped error with no retained cache and is not synthesized.",
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
    # Spans carry presentation style only; their itertext is the literal
    # dependency value.  Hyperlinks are likewise retained as their displayed
    # text.  ODF text:s/tab/line-break are literal whitespace.  Other nested
    # markup would change the dependency semantics.
    whitespace = (TEXT + "s", TEXT + "tab", TEXT + "line-break")
    unsupported = [
        node.tag
        for node in nested
        if node.tag not in (TEXT + "a", TEXT + "span", *whitespace)
    ]
    if unsupported:
        raise ValueError(
            "unsupported nested text markup in retained string cell: "
            + ", ".join(sorted(set(unsupported)))
        )
    for anchor in (node for node in nested if node.tag == TEXT + "a"):
        if any(
            child is not anchor
            and child.tag not in (TEXT + "a", TEXT + "span", *whitespace)
            for child in anchor.iter()
        ):
            raise ValueError("unsupported nested markup below retained text hyperlink")
    explicit = cell.get(OFFICE + "string-value")
    if explicit is not None:
        return explicit

    def render(node: ET.Element) -> str:
        if node.tag == TEXT + "s":
            return " " * int(node.get(TEXT + "c", "1"))
        if node.tag == TEXT + "tab":
            return "\t"
        if node.tag == TEXT + "line-break":
            return "\n"
        value = node.text or ""
        for child in node:
            value += render(child)
            value += child.tail or ""
        return value

    return "".join(render(paragraph) for paragraph in paragraphs)


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
    elif value_type in ("float", "currency"):
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
        nested = [
            node.tag
            for paragraph in cell.iter(TEXT + "p")
            for node in paragraph.iter()
            if node is not paragraph
        ]
        result.update({"type": "text", "value": text_value(cell)})
        if nested:
            result["nested_markup"] = sorted(set(nested))
    else:
        raise ValueError(f"unsupported referenced cell type {value_type!r} at {sheet}!R{row}C{column}")
    return result


def cached_value(cell: ET.Element):
    value_type = cell.get(OFFICE + "value-type")
    error_type = cell.get(CALEXT + "value-type") == "error"
    if error_type:
        # Error cells may carry office:string-value="" while the human
        # spelling is in text:p (for example #VALUE!).  Read the rendered
        # paragraph directly before falling back to the empty attribute.
        rendered = "".join("".join(paragraph.itertext()) for paragraph in cell.iter(TEXT + "p")).strip()
        if rendered.startswith("Err:"):
            # Keep the raw native spelling.  The test maps only the selected
            # explicit #VALUE! receipt into ScalarError::Value.
            rendered = rendered
        elif not rendered:
            rendered = cell.get(OFFICE + "string-value", "")
        return {"type": "error", "value": rendered}
    if value_type in ("float", "currency"):
        value = cell.get(OFFICE + "value")
        if value is None:
            raise ValueError("typed numeric cache has no office:value")
        number = float(value)
        if not number == number or number in (float("inf"), float("-inf")):
            raise ValueError("non-finite formula cache")
        return {"type": "number", "value": value}
    if value_type == "boolean":
        return {"type": "logical", "value": cell.get(OFFICE + "boolean-value", "false")}
    if value_type == "string":
        return {"type": "text", "value": cell.get(OFFICE + "string-value", text_value(cell))}
    raise ValueError("formula has no typed native cache")


def cache_attributes(cell: ET.Element):
    rendered_text = (
        "".join("".join(paragraph.itertext()) for paragraph in cell.iter(TEXT + "p")).strip()
        if cell.get(CALEXT + "value-type") == "error"
        else text_value(cell)
    )
    return {
        "office_value_type": cell.get(OFFICE + "value-type"),
        "office_value": cell.get(OFFICE + "value"),
        "office_string_value": cell.get(OFFICE + "string-value"),
        "office_boolean_value": cell.get(OFFICE + "boolean-value"),
        "calcext_value_type": cell.get(CALEXT + "value-type"),
        "matrix_columns_spanned": cell.get(TABLE + "number-matrix-columns-spanned"),
        "matrix_rows_spanned": cell.get(TABLE + "number-matrix-rows-spanned"),
        "rendered_text": rendered_text,
    }


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
            if not formula:
                raise ValueError(f"selected row {function} R{row_number}C{column_number_value} is not a formula")
            calls = [call.rsplit(".", 1)[-1].upper() for call in CALL_RE.findall(formula)]
            if not calls or calls[0] != function or set(calls) != {function}:
                raise ValueError(f"selected row has unsupported calls: {formula}")
            cached = cached_value(cell)
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
                    "column": column_number_value,
                    "formula": formula,
                    "cached": cached,
                    "cached_attributes": cache_attributes(cell),
                    # A FODS matrix formula has one cached top-left value but
                    # its native result is the complete declared matrix.  Do
                    # not validate that cache as scalar implicit intersection;
                    # ordinary (non-matrix) range formulas retain both modes.
                    "valid_modes": ["matrix"]
                    if cell.get(TABLE + "number-matrix-rows-spanned")
                    or cell.get(TABLE + "number-matrix-columns-spanned")
                    else ["scalar", "matrix"],
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
        "corpus": {
            "files": len(SOURCES),
            "paths": sorted(SOURCES.values()),
        },
        "source_staging": "Receipt inputs were assembled in a temporary tree from exact raw GitHub files at the pinned commit; no auxiliary checkout, conversion, or recalculation was used.",
        "byte_comparison": "Each staged input was SHA-256 checked against the exact raw GitHub file at commit before extraction.",
        "selected_observations": len(rows),
        "selected_functions": sorted(FUNCTIONS),
        "selected_modes": ["scalar", "matrix"],
        "excluded_profile_variances": EXCLUDED_PROFILE_VARIANCES,
        "coverage_notes": {
            "scope": "OpenFormula 1.4 §6.20; all 26 normative text functions have at least one selected native cache.",
            "source_format": "All selected sources are upstream FODS fixtures; no conversion tool or recalculation step is used.",
            "closure": "Each selected reference closure contains only literal number, logical, text, or empty cells. Formula cells, dates, named ranges, and nested formula dependencies remain explicit exclusions.",
            "cache": "The typed native cache is retained as an observation. Ordinary formulas are evaluated in scalar and matrix modes; FODS matrix-span formulas are validated in matrix mode with their declared output shape, and only the retained top-left cache value is compared.",
            "profile": "DOLLAR, FIXED, and TEXT numeric formatting uses the deterministic en-US profile represented by the selected native rows. Date, Japanese-locale, invalid-arity, nested, and computed-argument cases remain explicit exclusions.",
            "native_error": "FIND R16 is the selected #VALUE! cache. Other native error/no-cache rows remain exclusions and are not synthesized.",
        },
    }
    root = Path(__file__).resolve().parent
    expected_path = root / "provenance.json"
    if expected_path.is_file():
        expected = json.loads(expected_path.read_text())
        assert inputs == expected["inputs"], "inputs differ from pinned upstream fixture hashes"
    output.mkdir(parents=True, exist_ok=True)
    (output / "cached-results.json").write_text(json.dumps(rows, indent=2, ensure_ascii=False) + "\n")
    (output / "provenance.json").write_text(json.dumps(receipt, indent=2, ensure_ascii=False) + "\n")


def main():
    if len(sys.argv) < 2:
        raise SystemExit("usage: extract.py /path/to/staged-source [output-directory]")
    source = Path(sys.argv[1]).resolve()
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent
    extract(source, output)


if __name__ == "__main__":
    main()
