#!/usr/bin/env python3
"""Extract bounded conditional-aggregate observations from LibreOffice.

The FODS inputs are read directly.  The pinned SUMIFS input is an older BIFF
workbook rather than a FODS fixture; it is converted in a temporary directory
with the pinned receipt's LibreOffice converter so its cached formula and
literal source closure can be checked without changing a checkout.  Formula
cells in a retained source closure are refused.

Usage::

    python3 extract.py /path/to/staged-source [output-directory]
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET


UPSTREAM_COMMIT = "d804d6aff49054bad1719ec3c2d136b545bbc7e7"
UPSTREAM_URL = "https://github.com/LibreOffice/core"
RAW_URL = f"https://raw.githubusercontent.com/LibreOffice/core/{UPSTREAM_COMMIT}"

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
TEXT = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"

# The wildcard/regex fixtures are included in the byte receipt so the native
# profile exclusions remain auditable, but only the ordinary fixtures below
# contribute selected observations.
SOURCES = {
    "SUMIF": "sc/qa/unit/data/functions/mathematical/fods/sumif.fods",
    "SUMIFS": "sc/qa/unit/data/xls/opencl/math/sumifs.xls",
    "COUNTIF": "sc/qa/unit/data/functions/statistical/fods/countif.fods",
    "COUNTIFS": "sc/qa/unit/data/functions/statistical/fods/countifs.fods",
    "AVERAGEIF": "sc/qa/unit/data/functions/statistical/fods/averageif.fods",
    "AVERAGEIFS": "sc/qa/unit/data/functions/statistical/fods/averageifs.fods",
}

SUPPLEMENTAL_INPUTS = {
    "SUMIF wildcard profile": "sc/qa/unit/data/functions/mathematical/fods/sumif_wildcards.fods",
    "AVERAGEIF wildcard profile": "sc/qa/unit/data/functions/statistical/fods/averageif_wildcards.fods",
}

# Coordinates are one-based as they are in the source files.  SUMIFS is the
# one selected cell in the BIFF workbook; its source ranges contain 10,000
# bounded literal rows and are retained as a deliberately larger closure.
SELECTIONS = {
    "SUMIF": [(row, 1) for row in (
        3,
        4,
        5,
        6,
        8,
        9,
        10,
        11,
        12,
        13,
        14,
        15,
        16,
        17,
        19,
        20,
        21,
        22,
        23,
        24,
        25,
        26,
        27,
        29,
    )],
    "SUMIFS": [(3, 6)],
    "COUNTIF": [(row, 1) for row in (
        2,
        3,
        4,
        5,
        6,
        12,
        13,
        14,
        32,
    )],
    "COUNTIFS": [(row, 1) for row in (
        2,
        3,
        4,
        5,
        6,
        12,
        13,
        14,
        32,
        33,
        34,
        35,
        38,
    )],
    "AVERAGEIF": [(row, 1) for row in (2, 5, 8, 13, 14, 16)],
    "AVERAGEIFS": [(row, 1) for row in (2, 3, 6, 8, 9, 10, 11, 12, 13, 14, 15)],
}

# These rows stay in the provenance receipt but are not turned into selected
# expected values.  Host regular-expression/wildcard behavior and formula
# dependencies are profile choices; a native cache cannot silently decide
# those choices for the resolver contract.
EXCLUDED = [
    {
        "function": "SUMIF",
        "source": SOURCES["SUMIF"],
        "sheet": "Sheet2",
        "row": 2,
        "column": 1,
        "formula": 'of:=SUMIF([.I1:.I5];"unpaid";[.H1:.H5])',
        "cached": "6",
        "kind": "case_profile_variance",
        "reason": "LibreOffice's cache matches the case-variant literal Unpaid; the selected conditional profile keeps case policy explicit, so this native case-insensitive result is not promoted as normative.",
    },
    {
        "function": "SUMIF",
        "source": SOURCES["SUMIF"],
        "sheet": "Sheet2",
        "row": 18,
        "column": 1,
        "formula": "of:=SUMIF([.H2:.H8])",
        "kind": "arity_exclusion",
        "reason": "The fixture omits the criterion argument; the selected SUMIF contract requires a criterion.",
    },
    {
        "function": "SUMIF",
        "source": SOURCES["SUMIF"],
        "sheet": "Sheet2",
        "row": 31,
        "column": 1,
        "formula": 'of:=SUMIF({-10|10|20|30};">0")',
        "cached": "60",
        "kind": "reference_profile_variance",
        "reason": "The native fixture supplies an inline array as the range; the selected ODF profile admits worksheet references and reference lists, so this host result is retained only as an explicit out-of-profile source observation.",
    },
    {
        "function": "SUMIF",
        "source": SOURCES["SUMIF"],
        "sheet": "Sheet2",
        "rows": [28, 30],
        "kind": "empty_relational_criterion_variance",
        "reason": "LibreOffice interprets bare > and >= criteria as a host numeric comparison; the selected Criterion profile gives an empty right side no matches for relational operators, so these native cache results are excluded.",
    },
    {
        "function": "SUMIF",
        "source": SUPPLEMENTAL_INPUTS["SUMIF wildcard profile"],
        "sheet": "Sheet2",
        "rows": list(range(2, 26)),
        "kind": "host_regex_and_formula_dependency",
        "reason": "The wildcard-profile rows exercise LibreOffice regular-expression matching and their source range includes upstream formula cells; the selected profile leaves host regex/wildcard behavior explicit and the bounded resolver refuses formula-cell inputs.",
    },
    {
        "function": "AVERAGEIF",
        "source": SUPPLEMENTAL_INPUTS["AVERAGEIF wildcard profile"],
        "sheet": "Sheet2",
        "rows": [14, 15],
        "kind": "host_wildcard_variance",
        "reason": "Rows 14 and 15 use host wildcard criteria (*West and *(New Office)); wildcard enablement and whole-cell policy are host properties, so their native caches remain explicit exclusions.",
    },
    {
        "function": "AVERAGEIFS",
        "source": SOURCES["AVERAGEIFS"],
        "sheet": "Sheet2",
        "row": 5,
        "column": 1,
        "formula": 'of:=AVERAGEIFS([.L2:.L6];[.J2:.J6];"pen.*";[.K2:.K6];"<"&MAX([.K2:.K6]))',
        "cached": "65",
        "kind": "host_regex_variance",
        "reason": "The criterion pen.* relies on LibreOffice regular-expression matching; the selected profile does not infer regex enablement from a host cache.",
    },
    {
        "function": "COUNTIF",
        "source": SOURCES["COUNTIF"],
        "sheet": "Sheet2",
        "rows": list(range(34, 43)),
        "kind": "host_regex_and_case_profile",
        "reason": "These rows intentionally exercise regex, inline case controls, and case variants; they remain native profile evidence rather than selected numeric expectations.",
    },
    {
        "function": "COUNTIFS",
        "source": SOURCES["COUNTIFS"],
        "sheet": "Sheet2",
        "row": 8,
        "column": 1,
        "formula": 'of:=COUNTIFS([.L1:.L8]; ".+")',
        "cached": "4",
        "kind": "host_regex_variance",
        "reason": "The criterion .+ relies on LibreOffice regular-expression matching; the selected profile does not infer regex enablement from a host cache.",
    },
    {
        "function": "COUNTIF",
        "source": SOURCES["COUNTIF"],
        "sheet": "Sheet2",
        "row": 7,
        "column": 1,
        "formula": 'of:=COUNTIF([.L1:.L5]; ">=P")',
        "cached": "2",
        "kind": "case_profile_variance",
        "reason": "The native cache applies its host text collation to the mixed-case comparison; the selected profile uses case-sensitive Unicode scalar ordering and does not promote this host ordering result.",
    },
    {
        "function": "COUNTIF",
        "source": SOURCES["COUNTIF"],
        "sheet": "Sheet2",
        "row": 17,
        "column": 1,
        "formula": "of:=COUNTIF([.F24:.F28];[.F24:.F28])",
        "cached": "2",
        "kind": "criterion_reference_shape_exclusion",
        "reason": "The criterion argument is a multi-cell reference; the selected criterion contract admits only a single-cell reference.",
    },
    {
        "function": "COUNTIFS",
        "source": SOURCES["COUNTIFS"],
        "sheet": "Sheet2",
        "row": 39,
        "column": 1,
        "formula": "of:=COUNTIFS([.$I$38:.$I$39];[.F39])",
        "cached": "1",
        "formula_cell_attributes": {
            "office_value": "1",
            "office_value_type": "float",
            "display": "0",
        },
        "expected_literal": {
            "sheet": "Sheet2",
            "row": 39,
            "column": 2,
            "value_type": "float",
            "office_value": "0",
            "display": "0",
        },
        "criterion_cell": {
            "sheet": "Sheet2",
            "row": 39,
            "column": 6,
            "value_type": "string",
            "display": "2",
        },
        "candidate_cells": [
            {
                "sheet": "Sheet2",
                "row": 38,
                "column": 9,
                "value_type": "string",
                "display": "3",
            },
            {
                "sheet": "Sheet2",
                "row": 39,
                "column": 9,
                "value_type": "string",
                "display": "1",
            },
        ],
        "kind": "stale_cached_result_variance",
        "reason": "The pinned FODS stores A39 with office:value=1 but displays 0 in its text:p, while the expected literal B39 is float 0/display 0. F39 is Text 2 and I38/I39 are Text 3 and 1. The attribute, displayed result, and expected literal disagree, so the native cache is internally stale/inconsistent and is excluded rather than promoted; the selected evaluator result 0 is not treated as a production failure.",
    },
    {
        "function": "COUNTIF",
        "source": SOURCES["COUNTIF"],
        "sheet": "Sheet2",
        "rows": [33, 62],
        "kind": "empty_numeric_inequality_variance",
        "reason": "LibreOffice counts Empty cells as unequal for the numeric-looking <>7 criterion; the selected profile keeps Empty separate from Number criteria and therefore does not promote these host-coercion results.",
    },
    {
        "function": "COUNTIFS",
        "source": SOURCES["COUNTIFS"],
        "sheet": "Sheet2",
        "row": 7,
        "column": 1,
        "formula": 'of:=COUNTIFS([.L1:.L5]; ">=P")',
        "cached": "2",
        "kind": "case_profile_variance",
        "reason": "The native cache applies its host text collation to the mixed-case comparison; the selected profile uses case-sensitive Unicode scalar ordering and does not promote this host ordering result.",
    },
    {
        "function": "COUNTIFS",
        "source": SOURCES["COUNTIFS"],
        "sheet": "Sheet2",
        "row": 17,
        "column": 1,
        "formula": "of:=COUNTIFS([.F24:.F28];[.F24:.F28])",
        "cached": "2",
        "kind": "criterion_reference_shape_exclusion",
        "reason": "The criterion argument is a multi-cell reference; the selected criterion contract admits only a single-cell reference.",
    },
    {
        "function": "COUNTIF",
        "source": SOURCES["COUNTIF"],
        "sheet": "Sheet2",
        "row": 43,
        "column": 1,
        "formula": "of:=COUNTIF([.BM5:.BM9];0)",
        "cached": "0",
        "kind": "unmaterialized_source_variance",
        "reason": "The selected fixture closure lies beyond the bounded source cells materialized by the FODS table; no inferred empty-source cache is promoted.",
    },
]

FUNCTIONS = set(SOURCES)
CALL_RE = re.compile(r"([A-Za-z][A-Za-z0-9_.]*)\s*\(")
REFERENCE_RE = re.compile(r"\[([^\]]+)\]")
CELL_RE = re.compile(r"^\$?([A-Za-z]+)\$?(\d+)$")
MAX_REFERENCE_CELLS = 100_000
MAX_INPUT_BYTES = 2_000_000
CONVERTER_VERSION = "LibreOffice 26.2.5.2 620(Build:2)"
CONVERTER_BINARY_SHA256 = "51ac4b1ad2e310024b9d55f49f615bc134bb7304a87808c40f5d5c847a45b05d"


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
        sheet = prefix.strip("'").lstrip("$")
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
    paragraphs = list(cell.iter(TEXT + "p"))
    nested = [
        node.tag
        for paragraph in paragraphs
        for node in paragraph.iter()
        if node is not paragraph
    ]
    if nested:
        raise ValueError(
            "unsupported nested text markup in retained string cell: "
            + ", ".join(sorted(set(nested)))
        )
    explicit = cell.get(OFFICE + "string-value")
    if explicit is not None:
        return explicit
    return "".join(node.text or "" for node in paragraphs)


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
        raise ValueError(
            f"unsupported referenced cell type {value_type!r} at {sheet}!R{row}C{column}"
        )
    return location


def converter_receipt() -> dict[str, str]:
    command = shutil.which("libreoffice")
    if command is None:
        raise RuntimeError("SUMIFS native extraction requires libreoffice")
    version = subprocess.run(
        [command, "--version"], check=True, capture_output=True, text=True
    ).stdout.strip()
    binary = Path(command).resolve().parent / "soffice.bin"
    if not binary.is_file():
        raise RuntimeError(f"LibreOffice binary is missing: {binary}")
    digest = hashlib.sha256(binary.read_bytes()).hexdigest()
    receipt = {"command": command, "version": version, "soffice_binary_sha256": digest}
    if version != CONVERTER_VERSION or digest != CONVERTER_BINARY_SHA256:
        raise RuntimeError(
            "SUMIFS conversion tool differs from the pinned receipt: "
            f"{receipt!r} != version={CONVERTER_VERSION!r}, "
            f"sha256={CONVERTER_BINARY_SHA256!r}"
        )
    return receipt


def staged_fods(source: Path, relative: str, temporary: Path):
    path = source / relative
    if path.suffix.lower() != ".xls":
        data = path.read_bytes()
        if len(data) > MAX_INPUT_BYTES:
            raise ValueError(f"bounded fixture exceeds 2 MiB: {relative}")
        return data, None
    conversion_root = temporary / "converted"
    conversion_root.mkdir(parents=True, exist_ok=True)
    staged = conversion_root / Path(relative).name
    staged.write_bytes(path.read_bytes())
    command = shutil.which("libreoffice")
    if command is None:
        raise RuntimeError("SUMIFS native extraction requires libreoffice")
    subprocess.run(
        [command, "--headless", "--convert-to", "fods", "--outdir", str(conversion_root), str(staged)],
        check=True,
        capture_output=True,
        text=True,
    )
    converted = conversion_root / f"{staged.stem}.fods"
    if not converted.is_file():
        raise RuntimeError(f"LibreOffice did not emit {converted}")
    data = converted.read_bytes()
    if len(data) > 30_000_000:
        raise ValueError(f"converted fixture exceeds 30 MiB: {relative}")
    return data, converter_receipt()


def extract(source: Path, output: Path):
    inputs = {}
    rows = []
    converter = None
    all_inputs = {**SOURCES, **SUPPLEMENTAL_INPUTS}
    with tempfile.TemporaryDirectory(prefix="litchi-ods-conditional-materialize-") as temporary_name:
        temporary = Path(temporary_name)
        parsed = {}
        for relative in all_inputs.values():
            data = (source / relative).read_bytes()
            if len(data) > MAX_INPUT_BYTES:
                raise ValueError(f"bounded fixture exceeds 2 MiB: {relative}")
            inputs[relative] = hashlib.sha256(data).hexdigest()
        for function, relative in SOURCES.items():
            data, conversion = staged_fods(source, relative, temporary)
            if conversion is not None:
                converter = conversion
            parsed[function] = materialize_tables(ET.fromstring(data))

        for function, selections in SELECTIONS.items():
            sheets = parsed[function]
            sheet = "Main" if function == "SUMIFS" else "Sheet2"
            table = sheets.get(sheet, {})
            for row_number, column_number_value in selections:
                cell = table.get((row_number, column_number_value))
                if cell is None:
                    raise ValueError(
                        f"selected row {function} R{row_number}C{column_number_value} is missing"
                    )
                formula = cell.get(TABLE + "formula", "")
                if not formula or cell.get(OFFICE + "value-type") != "float":
                    raise ValueError(
                        f"selected row {function} R{row_number}C{column_number_value} is not a numeric formula"
                    )
                calls = [call.upper() for call in CALL_RE.findall(formula)]
                if not calls or calls[0] != function or set(calls) != {function}:
                    raise ValueError(f"selected row has unsupported calls: {formula}")
                cached = cell.get(OFFICE + "value")
                if cached is None:
                    raise ValueError(f"selected row {function} R{row_number}C{column_number_value} has no cache")
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
                        "source": SOURCES[function],
                        "sheet": sheet,
                        "row": row_number,
                        "column": column_number_value,
                        "formula": formula,
                        "cached": cached,
                        "cells": references,
                    }
                )

    expected_count = sum(len(selection) for selection in SELECTIONS.values())
    assert len(rows) == expected_count, (len(rows), expected_count)
    assert {row["function"] for row in rows} == FUNCTIONS
    receipt = {
        "upstream": UPSTREAM_URL,
        "raw_url": RAW_URL,
        "commit": UPSTREAM_COMMIT,
        "inputs": dict(sorted(inputs.items())),
        "source_staging": "Receipt inputs were assembled in a temporary tree from exact raw GitHub files at the pinned commit; the repository auxiliary checkout was not assumed to be pinned.",
        "byte_comparison": "Each staged input was SHA-256 checked against the exact raw GitHub file at commit before extraction.",
        "selected_observations": len(rows),
        "selected_functions": sorted(FUNCTIONS),
        "excluded_profile_variances": EXCLUDED,
        "converter": converter,
        "sumifs_materialization": "The pinned SUMIFS source is BIFF8/XLS.  A temporary, SHA-checked copy is converted to FODS by the exact LibreOffice binary recorded above; the retained closure contains only literal Data/Main cells and the formula value emitted by that conversion.  This receipt does not claim byte-preservation of the original BIFF8 cached value or absence of conversion recalculation.  The checkout is never modified.",
        "coverage_notes": {
            "regex_and_wildcards": "Regex and wildcard rows remain explicit host-profile exclusions; selected rows use literal, numeric, empty, and relational criteria whose behavior is independent of those host switches.",
            "sumifs_source": "No SUMIFS FODS fixture exists at the pinned upstream revision.  The selected SUMIFS observation therefore comes from the pinned opencl/math/sumifs.xls source through the bounded conversion receipt.",
            "text_materialization": "Retained string closures were audited at the pinned source bytes and contain direct text:p content only.  text_value rejects nested span, space, tab, or line-break markup rather than silently dropping it.",
        },
    }
    root = Path(__file__).resolve().parent
    expected_path = root / "provenance.json"
    if expected_path.is_file():
        expected = json.loads(expected_path.read_text())
        assert inputs == expected["inputs"], "inputs differ from pinned upstream fixture hashes"
        assert converter == expected["converter"], "converter differs from pinned receipt"
    output.mkdir(parents=True, exist_ok=True)
    (output / "cached-results.json").write_text(json.dumps(rows, indent=2) + "\n")
    (output / "provenance.json").write_text(json.dumps(receipt, indent=2) + "\n")


def main():
    if len(sys.argv) < 2:
        raise SystemExit("usage: extract.py /path/to/staged-source [output-directory]")
    source = Path(sys.argv[1]).resolve()
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent
    output.mkdir(parents=True, exist_ok=True)
    extract(source, output)


if __name__ == "__main__":
    main()
