#!/usr/bin/env python3
"""Extract bounded discrete-math observations from LibreOffice FODS fixtures.

Usage: ``python3 extract.py /path/to/libreoffice-core [output-directory]``.
The input checkout is read only.  Formula cells are never imported as resolver
inputs: every selected reference closure is retained as literal worksheet
cells, and a formula-cell dependency is an extraction error.
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
    "COMBIN": "sc/qa/unit/data/functions/mathematical/fods/combin.fods",
    "COMBINA": "sc/qa/unit/data/functions/mathematical/fods/combina.fods",
    "FACT": "sc/qa/unit/data/functions/mathematical/fods/fact.fods",
    "FACTDOUBLE": "sc/qa/unit/data/functions/addin/fods/factdouble.fods",
    "GCD": "sc/qa/unit/data/functions/mathematical/fods/gcd.fods",
    "LCM": "sc/qa/unit/data/functions/mathematical/fods/lcm.fods",
    "MULTINOMIAL": "sc/qa/unit/data/functions/mathematical/fods/multinomial.fods",
    "EVEN": "sc/qa/unit/data/functions/mathematical/fods/even.fods",
    "ODD": "sc/qa/unit/data/functions/mathematical/fods/odd.fods",
    "DELTA": "sc/qa/unit/data/functions/addin/fods/delta.fods",
    "GESTEP": "sc/qa/unit/data/functions/addin/fods/gestep.fods",
}

# Coordinates are one-based, as they are in the source FODS files.  The
# selected rows retain direct literals and small literal reference closures.
# Inputs that exercise a disputed host coercion or an unsupported source
# dependency remain in EXCLUDED below rather than becoming invented results.
SELECTIONS = {
    "COMBIN": (2, 3, 4, 5, 6, 11, 13),
    "COMBINA": (2, 3, 5, 6, 11, 13),
    "FACT": (2, 3, 4, 5, 6),
    "FACTDOUBLE": (2, 3, 4, 7, 11, 15, 25),
    "GCD": (2, 3, 4, 7, 8, 9, 10, 11, 21, 39),
    "LCM": (2, 3, 4, 7, 10, 21, 39),
    "MULTINOMIAL": (2, 3, 4, 5, 6, 8, 13, 14, 15, 17, 18, 19, 20),
    "EVEN": (2, 3, 4, 5, 6),
    "ODD": (2, 3, 4, 5, 6),
    "DELTA": (2, 3, 4, 8, 12, 13),
    "GESTEP": (2, 3, 4, 8, 12, 13),
}

# These are source observations deliberately outside the selected bounded
# contract.  Their formula/cache/coordinate details make each omission
# auditable without treating a host result as a normative expected value.
EXCLUDED = [
    {
        "function": "COMBINA",
        "source": SOURCES["COMBINA"],
        "sheet": "Sheet2",
        "row": 4,
        "column": 1,
        "formula": "of:=COMBINA(0;0)",
        "cached": "0",
        "kind": "contract_divergence",
        "reason": "LibreOffice caches COMBINA(0;0) as 0; the selected evaluator contract defines the zero/zero case as 1, so this host variance is excluded from numeric corroboration.",
    },
    {
        "function": "COMBIN",
        "source": SOURCES["COMBIN"],
        "sheet": "Sheet2",
        "row": 12,
        "column": 1,
        "formula": "of:=COMBIN(12;)",
        "cached": "1",
        "kind": "contract_divergence",
        "reason": "The fixture omits number_chosen; omitted-argument defaulting is outside the selected two-argument contract.",
    },
    {
        "function": "COMBIN",
        "source": SOURCES["COMBIN"],
        "sheet": "Sheet2",
        "row": 9,
        "column": 1,
        "formula": "of:=COMBIN(1;2)",
        "kind": "domain_exclusion",
        "reason": "The fixture has number_chosen greater than number; the selected binomial domain requires number_chosen <= number, so its formula error is excluded from numeric cache corroboration.",
    },
    {
        "function": "COMBINA",
        "source": SOURCES["COMBINA"],
        "sheet": "Sheet2",
        "row": 9,
        "column": 1,
        "formula": "of:=COMBINA(1;2)",
        "kind": "domain_exclusion",
        "reason": "The fixture has number_chosen greater than number; the selected strict COMBINA domain requires number_chosen <= number, so its formula error is excluded from numeric cache corroboration.",
    },
    {
        "function": "COMBINA",
        "source": SOURCES["COMBINA"],
        "sheet": "Sheet2",
        "row": 12,
        "column": 1,
        "formula": "of:=COMBINA(12;)",
        "cached": "1",
        "kind": "contract_divergence",
        "reason": "The fixture omits number_chosen; omitted-argument defaulting is outside the selected two-argument contract.",
    },
    {
        "function": "MULTINOMIAL",
        "source": SOURCES["MULTINOMIAL"],
        "sheet": "Sheet2",
        "row": 7,
        "column": 1,
        "formula": "of:=MULTINOMIAL([.G7];[.H7];[.I7])",
        "cached": "6.18970023101455E+026",
        "reference_cell": {"sheet": "Sheet2", "row": 7, "column": 7, "formula": "of:=2^30"},
        "kind": "formula_cell_input",
        "reason": "The first referenced operand is an upstream formula cell (G7 of:=2^30); the bounded native resolver refuses formula-cell inputs rather than evaluating or importing its cache.",
    },
    {
        "function": "LCM",
        "source": SOURCES["LCM"],
        "sheet": "Sheet2",
        "row": 8,
        "column": 1,
        "formula": "of:=LCM(1.2;2.4)",
        "cached": "2",
        "kind": "contract_divergence",
        "reason": "The native cache accepts fractional LCM operands, while the resolved LCM contract rejects non-integer inputs; the host result is retained only as an explicit variance.",
    },
    {
        "function": "LCM",
        "source": SOURCES["LCM"],
        "sheet": "Sheet2",
        "row": 9,
        "column": 1,
        "formula": "of:=LCM(1.2;3.6)",
        "cached": "3",
        "kind": "contract_divergence",
        "reason": "The native cache accepts fractional LCM operands, while the resolved LCM contract rejects non-integer inputs; the host result is retained only as an explicit variance.",
    },
    {
        "function": "LCM",
        "source": SOURCES["LCM"],
        "sheet": "Sheet2",
        "row": 11,
        "column": 1,
        "formula": "of:=LCM([.L1];[.M1])",
        "cached": "2",
        "reference_cells": [
            {"sheet": "Sheet2", "row": 1, "column": 12, "type": "number", "value": "1.2"},
            {"sheet": "Sheet2", "row": 1, "column": 13, "type": "number", "value": "2.4"},
        ],
        "kind": "contract_divergence",
        "reason": "The native cache accepts fractional LCM operands from literal references, while the resolved LCM contract rejects non-integer inputs; the host result is retained only as an explicit variance.",
    },
    {
        "function": "GCD",
        "source": SOURCES["GCD"],
        "sheet": "Sheet2",
        "row": 37,
        "column": 1,
        "formula": "of:=GCD(6;{2;4})",
        "cached": "2",
        "kind": "unsupported_dependency",
        "reason": "The forced-array operand is outside this bounded scalar/reference native fixture; no array cache is promoted without the corresponding discrete sequence contract.",
    },
    {
        "function": "LCM",
        "source": SOURCES["LCM"],
        "sheet": "Sheet2",
        "row": 37,
        "column": 1,
        "formula": "of:=LCM(6;{2;4})",
        "cached": "12",
        "kind": "unsupported_dependency",
        "reason": "The forced-array operand is outside this bounded scalar/reference native fixture; no array cache is promoted without the corresponding discrete sequence contract.",
    },
    {
        "function": "GESTEP",
        "source": SOURCES["GESTEP"],
        "sheet": "Sheet2",
        "row": 9,
        "column": 1,
        "formula": "of:=GESTEP([.G7];[.H7])",
        "kind": "unsupported_dependency",
        "reason": "The referenced input includes the fixture's literal text cell G7 and the native formula is an error; ordinary Number conversion yields a formula error in the resolved GESTEP contract, so this nonnumeric case remains an explicit exclusion.",
    },
    {
        "function": "GESTEP",
        "source": SOURCES["GESTEP"],
        "sheet": "Sheet2",
        "row": 10,
        "column": 1,
        "formula": "of:=GESTEP([.H7];[.G7])",
        "kind": "unsupported_dependency",
        "reason": "The referenced input includes the fixture's literal text cell G7 and the native formula is an error; ordinary Number conversion yields a formula error in the resolved GESTEP contract, so this nonnumeric case remains an explicit exclusion.",
    },
    {
        "function": "GESTEP",
        "source": SOURCES["GESTEP"],
        "sheet": "Sheet2",
        "row": 11,
        "column": 1,
        "formula": "of:=GESTEP(\"dog\";0)",
        "kind": "contract_divergence",
        "reason": "The fixture supplies text as the number operand and records a formula error; the resolved GESTEP contract uses ordinary Number conversion, and no logical-valued GESTEP case exists in the pinned source.",
    },
    {
        "function": "DELTA",
        "source": SOURCES["DELTA"],
        "sheet": "Sheet2",
        "row": 11,
        "column": 1,
        "formula": "of:=DELTA(\"dog\";0)",
        "kind": "contract_divergence",
        "reason": "The fixture supplies text as the number operand and records a formula error; ordinary Number conversion yields a formula error in the resolved DELTA contract, so this nonnumeric case remains an explicit exclusion.",
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
        if len(data) > 2_000_000:
            raise ValueError(f"bounded fixture exceeds 2 MiB: {relative}")
        inputs[relative] = hashlib.sha256(data).hexdigest()
        sheets = materialize_tables(ET.fromstring(data))
        sheet = "Sheet2"
        table = sheets.get(sheet, {})
        for row_number in SELECTIONS[function]:
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

    expected_count = sum(len(selection) for selection in SELECTIONS.values())
    assert len(rows) == expected_count, len(rows)
    assert {row["function"] for row in rows} == FUNCTIONS
    expected = json.loads((Path(__file__).resolve().parent / "provenance.json").read_text())
    assert inputs == expected["inputs"], "inputs differ from pinned upstream fixture hashes"
    receipt = {
        "upstream": UPSTREAM_URL,
        "raw_url": RAW_URL,
        "commit": UPSTREAM_COMMIT,
        "inputs": inputs,
        "source_staging": "Receipt inputs were assembled in a temporary tree from the exact raw GitHub files at the pinned commit; the repository auxiliary checkout was not assumed to be pinned.",
        "byte_comparison": "Each staged input was SHA-256 checked against the exact raw GitHub file at commit before extraction.",
        "selected_observations": len(rows),
        "selected_functions": sorted(FUNCTIONS),
        "excluded_profile_variances": EXCLUDED,
        "coverage_notes": {
            "gestep_text": "Pinned GESTEP rows 9-11 contain text/error inputs and remain explicit exclusions; ordinary Number conversion yields a formula error, and no logical-valued GESTEP input is present in the source fixture.",
            "multinomial_fractional": "Pinned MULTINOMIAL row 19 uses fractional operands (3.4 and 2.3) and is retained; its cached 10 agrees with the resolved raw-sum-before-floor contract, although this pair does not distinguish that rule from per-argument flooring.",
        },
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
