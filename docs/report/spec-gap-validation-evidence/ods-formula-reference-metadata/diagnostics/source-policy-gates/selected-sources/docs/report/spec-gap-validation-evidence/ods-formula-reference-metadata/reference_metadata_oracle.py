#!/usr/bin/env python3
"""Independent reference-metadata oracle.

The model in this file deliberately owns its small reference geometry types and
does not import, execute, or mirror the Rust evaluator.  It describes local
references, ordered reference lists, arrays, and the four-sheet workbook used
by the retained Rust/native receipts.  Cell contents are intentionally absent:
all eight functions are metadata operations and must be able to answer their
admitted cases without reading a cell.

The ODF entries leave a few conversion details to an evaluator.  This profile
chooses the following bounded behavior for the slice:

* AREAS counts logical reference records, including records in an ordered
  ReferenceList; physical planes of a 3-D record do not increase the count.
* COLUMN and ROW admit exactly one logical Reference record.  Their
  omitted-argument cases use the current position.  Matrix cases publish the
  complete horizontal COLUMN vector or vertical ROW vector required by the
  function entry.  Scalar publication of an explicit multi-cell reference
  selects the first element of that generated axis; it does not intersect the
  input reference at the caller's position.
* COLUMNS and ROWS accept one direct Reference or rectangular Array.  A
  ReferenceList, including a list with one retained record, is a pseudotype
  refusal and is not silently flattened into an Array.
* SHEET and SHEETS use local sheet metadata only.  A reference is not
  dereferenced; a 3-D reference starts at its first sheet.  Text SHEET names
  a local sheet.  Source-qualified references are classified by ISREF and are
  refused by SHEET/SHEETS without fetching the external workbook.
* ISREF observes the runtime kind, including ReferenceList, and never projects
  or reads its argument.

Every observation records its formula, mode, caller position, and expected
typed value.  The retained native fixture is a separate observation of local
LibreOffice behavior; it is not used to construct these expected values.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import json
from pathlib import Path
from typing import Final


HERE = Path(__file__).resolve().parent
GOLDENS = HERE / "reference-metadata-goldens.json"
ARCHIVE_SHA256: Final = (
    "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4"
)
PART4_SHA256: Final = (
    "ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1"
)
BASELINE: Final = "049c09cdde3978593149079c4257df047a3fa419"
FUNCTIONS: Final = (
    "AREAS",
    "COLUMN",
    "COLUMNS",
    "ISREF",
    "ROW",
    "ROWS",
    "SHEET",
    "SHEETS",
)
SHEETS: Final = ("Main", "Data", "Archive", "Hidden")
SHEET_INDEX: Final = {name: index for index, name in enumerate(SHEETS)}


@dataclass(frozen=True)
class Ref:
    """One logical Reference record.

    ``sheet_start``/``sheet_end`` are None for a current-sheet reference.  The
    endpoints are inclusive and use zero-based coordinates internally.  A
    record spanning several sheets remains one record for AREAS.
    """

    text: str
    sheet_start: int | None
    sheet_end: int | None
    row_start: int
    row_end: int
    column_start: int
    column_end: int

    @property
    def is_source(self) -> bool:
        return "#" in self.text and self.text.startswith("[")

    def sheets(self, current_sheet: str) -> tuple[int, ...]:
        start = SHEET_INDEX[current_sheet] if self.sheet_start is None else self.sheet_start
        end = start if self.sheet_end is None else self.sheet_end
        if end < start:
            raise ValueError("reference sheet endpoints are reversed")
        return tuple(range(start, end + 1))

    def rows(self) -> int:
        return self.row_end - self.row_start + 1

    def columns(self) -> int:
        return self.column_end - self.column_start + 1


@dataclass(frozen=True)
class RefList:
    records: tuple[Ref, ...]


@dataclass(frozen=True)
class Array:
    rows: int
    columns: int
    values: tuple["Scalar", ...]


@dataclass(frozen=True)
class Scalar:
    kind: str
    value: object


Arg = Ref | RefList | Array | Scalar


def number(value: int | float) -> dict[str, object]:
    return {"type": "number", "value": value}


def logical(value: bool) -> dict[str, object]:
    return {"type": "logical", "value": value}


def text(value: str) -> dict[str, object]:
    return {"type": "text", "value": value}


def error(value: str = "#VALUE!") -> dict[str, object]:
    return {"type": "error", "value": value}


def array_value(rows: int, columns: int, values: list[object]) -> dict[str, object]:
    if len(values) != rows * columns:
        raise ValueError("array values do not match shape")
    return {"type": "array", "rows": rows, "columns": columns, "values": values}


def ref(
    spelling: str,
    *,
    sheet_start: str | None,
    sheet_end: str | None = None,
    row_start: int,
    row_end: int | None = None,
    column_start: int,
    column_end: int | None = None,
) -> Ref:
    return Ref(
        spelling,
        None if sheet_start is None else SHEET_INDEX[sheet_start],
        None if sheet_end is None else SHEET_INDEX[sheet_end],
        row_start,
        row_start if row_end is None else row_end,
        column_start,
        column_start if column_end is None else column_end,
    )


def current_cell(
    spelling: str = "[.A1]",
    *,
    row: int = 0,
    column: int = 0,
) -> Ref:
    return ref(
        spelling,
        sheet_start=None,
        row_start=row,
        column_start=column,
    )


def parse_ref(spelling: str) -> Ref:
    """Parse the deliberately small, independently authored fixture syntax."""

    if not (spelling.startswith("[") and spelling.endswith("]")):
        raise ValueError(f"not a bracket reference: {spelling}")
    body = spelling[1:-1]
    if "#" in body:
        # Source-qualified references are retained as a kind marker.  The
        # local metadata profile can classify ISREF and reject SHEET/SHEETS
        # without resolving the external workbook.  Keep only the local
        # coordinate tail for the descriptor shape; other geometry functions
        # are intentionally absent from the source corpus because they have a
        # typed external-capability refusal.
        body = body.rsplit("#", 1)[1]
    endpoints = body.split(":")
    if len(endpoints) > 2:
        raise ValueError(f"ambiguous reference: {spelling}")

    def endpoint(value: str) -> tuple[str | None, int, int]:
        if value.startswith("."):
            value = value[1:]
            sheet = None
        else:
            sheet, value = value.split(".", 1)
        column = ""
        row = ""
        for character in value:
            if character.isalpha():
                column += character
            elif character.isdigit():
                row += character
            else:
                raise ValueError(f"unsupported endpoint: {spelling}")
        if not column or not row or not column.isupper():
            raise ValueError(f"unsupported endpoint: {spelling}")
        column_number = 0
        for character in column:
            column_number = column_number * 26 + ord(character) - ord("A") + 1
        return sheet, int(row) - 1, column_number - 1

    left_sheet, left_row, left_column = endpoint(endpoints[0])
    if len(endpoints) == 1:
        return ref(
            spelling,
            sheet_start=left_sheet,
            row_start=left_row,
            column_start=left_column,
        )
    right = endpoints[1]
    right_sheet, right_row, right_column = endpoint(right)
    if right_sheet is not None and left_sheet is None:
        raise ValueError("fixture does not use a right-only sheet locator")
    if right_sheet is None:
        right_sheet = left_sheet
    if left_sheet is not None and right_sheet is not None:
        sheet_start = left_sheet
        sheet_end = right_sheet
    else:
        sheet_start = None
        sheet_end = None
    return ref(
        spelling,
        sheet_start=sheet_start,
        sheet_end=sheet_end,
        row_start=min(left_row, right_row),
        row_end=max(left_row, right_row),
        column_start=min(left_column, right_column),
        column_end=max(left_column, right_column),
    )


def array_from_formula(spelling: str) -> Array:
    if spelling == "{1;2|3;4}":
        return Array(
            2,
            2,
            (
                Scalar("number", 1),
                Scalar("number", 2),
                Scalar("number", 3),
                Scalar("number", 4),
            ),
        )
    if spelling == "{1;2;3}":
        return Array(
            1,
            3,
            (Scalar("number", 1), Scalar("number", 2), Scalar("number", 3)),
        )
    if spelling == '{"Main";"Data"}':
        return Array(
            1,
            2,
            (Scalar("text", "Main"), Scalar("text", "Data")),
        )
    raise ValueError(f"unknown oracle array {spelling}")


def arg(spelling: str) -> Arg:
    if spelling.startswith(("IFERROR(", "IFNA(")) and spelling.endswith(")"):
        parts = spelling[spelling.find("(") + 1 : -1].split(";")
        if len(parts) == 2:
            # The source descriptor is retained through either handler.  The
            # typed arithmetic variants are exercised by the semantic suite;
            # this value-only oracle records the descriptor-preserving path.
            return arg(parts[0])
    if spelling.startswith("IF(") and spelling.endswith(")"):
        parts = spelling[3:-1].split(";")
        if len(parts) == 3 and parts[0] in ("TRUE()", "FALSE()"):
            return arg(parts[1] if parts[0] == "TRUE()" else parts[2])
    if spelling == "(([.A1]~[.B2])![.A1])":
        # A list/intersection retains list kind even when only one record
        # survives.  This is the executable one-record ReferenceList shape
        # used to exercise the §4.9/§5.9 refusal rule.
        return RefList((parse_ref("[.A1]"),))
    if "~" in spelling:
        pieces = spelling.split("~")
        return RefList(tuple(parse_ref(piece) for piece in pieces))
    if spelling.startswith("["):
        return parse_ref(spelling)
    if spelling.startswith("{"):
        return array_from_formula(spelling)
    if spelling == "#N/A":
        return Scalar("error", spelling)
    if spelling == "#VALUE!":
        return Scalar("error", spelling)
    if spelling == "TRUE":
        return Scalar("logical", True)
    if spelling == "FALSE":
        return Scalar("logical", False)
    if spelling == "TRUE()":
        return Scalar("logical", True)
    if spelling == "FALSE()":
        return Scalar("logical", False)
    if spelling.startswith('"') and spelling.endswith('"'):
        return Scalar("text", json.loads(spelling))
    try:
        return Scalar("number", float(spelling) if "." in spelling else int(spelling))
    except ValueError as exc:
        raise ValueError(f"unknown scalar {spelling}") from exc


def reference_list_or_error(value: Arg) -> RefList | dict[str, object]:
    if isinstance(value, RefList):
        return value
    if isinstance(value, Ref):
        return RefList((value,))
    return error()


def one_reference_or_error(value: Arg) -> Ref | dict[str, object]:
    if isinstance(value, Ref):
        return value
    return error()


def metadata(function: str, arguments: list[Arg], *, mode: str, position: dict[str, object]) -> dict[str, object]:
    current_sheet = str(position["sheet"])
    current_row = int(position["row"]) - 1
    current_column = int(position["column"]) - 1
    if function == "AREAS":
        if len(arguments) != 1:
            return error()
        value = reference_list_or_error(arguments[0])
        return number(len(value.records)) if isinstance(value, RefList) else value
    if function == "ISREF":
        if len(arguments) != 1:
            return error()
        return logical(isinstance(arguments[0], (Ref, RefList)))
    if function in ("COLUMN", "ROW"):
        if len(arguments) > 1:
            return error()
        value: Arg = (
            current_cell(
                f"[{current_sheet}.{chr(ord('A') + current_column)}{current_row + 1}]",
                row=current_row,
                column=current_column,
            )
            if not arguments
            else arguments[0]
        )
        reference = one_reference_or_error(value)
        if not isinstance(reference, Ref):
            return reference
        if mode == "matrix":
            if function == "COLUMN":
                values = list(range(reference.column_start + 1, reference.column_end + 2))
                if len(values) == 1:
                    return number(values[0])
                return array_value(1, len(values), values)
            values = list(range(reference.row_start + 1, reference.row_end + 2))
            if len(values) == 1:
                return number(values[0])
            return array_value(len(values), 1, values)
        # Scalar publication projects the generated axis at its first element.
        # The complete axis is retained by the matrix operation above; this
        # boundary projection must not intersect the input reference at the
        # caller's position.
        return number(
            reference.column_start + 1
            if function == "COLUMN"
            else reference.row_start + 1
        )
    if function in ("COLUMNS", "ROWS"):
        if len(arguments) != 1:
            return error()
        value = arguments[0]
        if isinstance(value, Array):
            return number(value.columns if function == "COLUMNS" else value.rows)
        if not isinstance(value, Ref):
            return error()
        reference = value
        if isinstance(reference, Ref):
            return number(reference.columns() if function == "COLUMNS" else reference.rows())
        return error()
    if function == "SHEET":
        if len(arguments) > 1:
            return error()
        if not arguments:
            return number(SHEET_INDEX[current_sheet] + 1)
        value = arguments[0]
        if isinstance(value, Array):
            if not value.values:
                return error()
            if mode == "matrix":
                return array_value(
                    value.rows,
                    value.columns,
                    [sheet_scalar(item) for item in value.values],
                )
            return sheet_scalar(value.values[0])
        if isinstance(value, Scalar):
            return sheet_scalar(value)
        if not isinstance(value, Ref):
            return error()
        reference = value
        if isinstance(reference, Ref):
            if reference.is_source:
                return error()
            first = reference.sheets(current_sheet)[0]
            return number(first + 1)
        return error()
    if function == "SHEETS":
        if len(arguments) > 1:
            return error()
        if not arguments:
            return number(len(SHEETS))
        value = arguments[0]
        if not isinstance(value, Ref):
            return error()
        reference = value
        if isinstance(reference, Ref):
            if reference.is_source:
                return error()
            return number(len(reference.sheets(current_sheet)))
        return error()
    raise ValueError(function)


def sheet_scalar(value: Scalar) -> dict[str, object]:
    if value.kind == "error":
        return error(str(value.value))
    if value.kind == "logical":
        name = "TRUE" if value.value else "FALSE"
    elif value.kind == "number":
        name = str(value.value)
    elif value.kind == "text":
        name = str(value.value)
    else:
        return error()
    return number(SHEET_INDEX[name] + 1) if name in SHEET_INDEX else error("#REF!")


def formula_call(name: str, arguments: list[str]) -> str:
    return "=" + name + "(" + ";".join(arguments) + ")"


def position(sheet: str = "Main", row: int = 1, column: int = 1) -> dict[str, object]:
    return {"sheet": sheet, "row": row, "column": column}


def add(
    rows: list[dict[str, object]],
    case: str,
    function: str,
    arguments: list[str],
    *,
    mode: str = "scalar",
    at: dict[str, object] | None = None,
    expected: dict[str, object] | None = None,
) -> None:
    caller = position() if at is None else at
    typed_arguments = [arg(item) for item in arguments]
    wanted = metadata(function, typed_arguments, mode=mode, position=caller)
    if expected is not None and wanted != expected:
        raise AssertionError(f"{case}: supplied expected differs from model: {wanted} != {expected}")
    rows.append(
        {
            "case": case,
            "function": function,
            "arguments": arguments,
            "formula": formula_call(function, arguments),
            "mode": mode,
            "position": caller,
            "expected": wanted,
            "expected_reads": 0,
        }
    )


def document() -> dict[str, object]:
    rows: list[dict[str, object]] = []

    # AREAS counts logical records and keeps duplicate list members.
    add(rows, "areas.single_cell", "AREAS", ["[.A1]"])
    add(rows, "areas.two_dimensional", "AREAS", ["[.B2:.D4]"])
    add(rows, "areas.three_dimensional_record", "AREAS", ["[Data.A1:Archive.C3]"])
    add(rows, "areas.ordered_list", "AREAS", ["[.A1]~[.B2:.C3]"])
    add(rows, "areas.duplicate_list_member", "AREAS", ["[.A1]~[.A1]"])
    add(rows, "areas.scalar_refusal", "AREAS", ["42"])
    add(rows, "areas.array_refusal", "AREAS", ["{1;2|3;4}"])

    # COLUMN and ROW have an omitted/current-cell path, single-cell and
    # multi-cell scalar paths, and complete vectors in matrix mode.  The
    # scalar cases use the first element of the generated axis even when the
    # caller position is elsewhere.
    add(rows, "column.omitted_current", "COLUMN", [], at=position("Main", 5, 6))
    add(rows, "column.single_cell", "COLUMN", ["[Data.C4]"], at=position("Main", 2, 2))
    add(rows, "column.vector_scalar", "COLUMN", ["[.B2:.D2]"], at=position("Main", 9, 9))
    add(rows, "column.two_dimensional_scalar", "COLUMN", ["[.B2:.D4]"], at=position("Main", 9, 9))
    add(rows, "column.vector_matrix", "COLUMN", ["[Data.B2:Archive.D4]"], mode="matrix")
    add(rows, "column.list_refusal", "COLUMN", ["[.A1]~[.B2]"])
    add(rows, "column.one_record_list_refusal", "COLUMN", ["(([.A1]~[.B2])![.A1])"])
    add(rows, "column.array_refusal", "COLUMN", ["{1;2;3}"])

    add(rows, "row.omitted_current", "ROW", [], at=position("Data", 4, 6))
    add(rows, "row.single_cell", "ROW", ["[Archive.C4]"], at=position("Main", 2, 2))
    add(rows, "row.vector_scalar", "ROW", ["[.B2:.B4]"], at=position("Main", 9, 9))
    add(rows, "row.two_dimensional_scalar", "ROW", ["[.B2:.D4]"], at=position("Main", 9, 9))
    add(rows, "row.vector_matrix", "ROW", ["[Data.B2:Archive.D4]"], mode="matrix")
    add(rows, "row.list_refusal", "ROW", ["[.A1]~[.B2]"])
    add(rows, "row.one_record_list_refusal", "ROW", ["(([.A1]~[.B2])![.A1])"])
    add(rows, "row.array_refusal", "ROW", ["{1;2;3}"])

    # COLUMNS/ROWS are geometry-only and accept rectangular arrays as well as
    # one reference.  A 3-D record reports the rectangle's dimensions once.
    add(rows, "columns.single_cell", "COLUMNS", ["[.A1]"])
    add(rows, "columns_two_dimensional", "COLUMNS", ["[.B2:.D4]"])
    add(rows, "columns_three_dimensional", "COLUMNS", ["[Data.B2:Archive.D4]"])
    add(rows, "columns_array", "COLUMNS", ["{1;2|3;4}"])
    add(rows, "columns_list_refusal", "COLUMNS", ["[.A1]~[.B2]"])
    add(rows, "columns.one_record_list_refusal", "COLUMNS", ["(([.A1]~[.B2])![.A1])"])
    add(rows, "columns_scalar_refusal", "COLUMNS", ["42"])
    add(rows, "rows.single_cell", "ROWS", ["[.A1]"])
    add(rows, "rows_two_dimensional", "ROWS", ["[.B2:.D4]"])
    add(rows, "rows_three_dimensional", "ROWS", ["[Data.B2:Archive.D4]"])
    add(rows, "rows_array", "ROWS", ["{1;2|3;4}"])
    add(rows, "rows_list_refusal", "ROWS", ["[.A1]~[.B2]"])
    add(rows, "rows.one_record_list_refusal", "ROWS", ["(([.A1]~[.B2])![.A1])"])
    add(rows, "rows_scalar_refusal", "ROWS", ["42"])

    # ISREF observes every ReferenceList.  AREAS also counts a list; the other
    # six reference parameters require a direct Reference and reject every
    # ReferenceList, including one-record lists.
    add(rows, "isref_single_cell", "ISREF", ["[.A1]"])
    add(rows, "isref_three_dimensional", "ISREF", ["[Data.A1:Archive.B2]"])
    add(rows, "isref_ordered_list", "ISREF", ["[.A1]~[.B2]"])
    add(rows, "isref_source_reference", "ISREF", ["['file:///book.ods'#.A1]"])
    add(
        rows,
        "isref_selected_source",
        "ISREF",
        ["IF(TRUE();['file:///book.ods'#.A1];[.A1])"],
    )
    add(
        rows,
        "isref_selected_local_after_source",
        "ISREF",
        ["IF(FALSE();['file:///book.ods'#.A1];[.A1])"],
    )
    add(
        rows,
        "isref_iferror_source",
        "ISREF",
        ["IFERROR(['file:///book.ods'#.A1];\"Main\")"],
    )
    add(
        rows,
        "isref_ifna_source",
        "ISREF",
        ["IFNA(['file:///book.ods'#.A1];\"Main\")"],
    )
    add(rows, "isref_array", "ISREF", ["{1;2|3;4}"])
    add(rows, "isref_number", "ISREF", ["42"])
    add(rows, "isref_formula_error_value", "ISREF", ["#N/A"])

    # SHEET uses workbook order, with local text lookup and first-sheet 3-D
    # reference semantics.  SHEETS counts the cuboid span or all workbook
    # sheets when omitted; hidden sheets are still present in the model.
    add(rows, "sheet_omitted_current", "SHEET", [], at=position("Archive", 1, 1))
    add(rows, "sheet_local_text", "SHEET", ['"Data"'])
    add(rows, "sheet_unknown_text", "SHEET", ['"Missing"'])
    add(rows, "sheet_local_reference", "SHEET", ["[.B2]"])
    add(rows, "sheet_explicit_reference", "SHEET", ["[Archive.B2]"])
    add(rows, "sheet_three_dimensional_first", "SHEET", ["[Data.A1:Archive.B2]"])
    add(rows, "sheet_list_refusal", "SHEET", ["[.A1]~[.B2]"])
    add(rows, "sheet.one_record_list_refusal", "SHEET", ["(([.A1]~[.B2])![.A1])"])
    add(rows, "sheet_source_refusal", "SHEET", ["['file:///book.ods'#.A1]"])
    add(
        rows,
        "sheet_selected_source_refusal",
        "SHEET",
        ["IF(TRUE();['file:///book.ods'#.A1];[Data.A1])"],
    )
    add(
        rows,
        "sheet_selected_local_after_source",
        "SHEET",
        ["IF(FALSE();['file:///book.ods'#.A1];[Data.A1])"],
    )
    add(
        rows,
        "sheet_iferror_source_refusal",
        "SHEET",
        ["IFERROR(['file:///book.ods'#.A1];\"Main\")"],
    )
    add(
        rows,
        "sheet_ifna_source_refusal",
        "SHEET",
        ["IFNA(['file:///book.ods'#.A1];\"Main\")"],
    )
    add(rows, "sheet_number_conversion", "SHEET", ["1"])
    add(rows, "sheet_logical_conversion", "SHEET", ["TRUE()"])
    add(rows, "sheet_array_scalar_projection", "SHEET", ['{"Main";"Data"}'])
    add(rows, "sheet_array_matrix_conversion", "SHEET", ['{"Main";"Data"}'], mode="matrix")
    add(rows, "sheet_array_number_refusal", "SHEET", ["{1;2;3}"])
    add(rows, "sheets_omitted_document", "SHEETS", [])
    add(rows, "sheets_local_reference", "SHEETS", ["[.B2]"])
    add(rows, "sheets_three_dimensional", "SHEETS", ["[Data.A1:Archive.B2]"])
    add(rows, "sheets_explicit_same_sheet", "SHEETS", ["[Archive.B2:Archive.D4]"])
    add(rows, "sheets_list_refusal", "SHEETS", ["[.A1]~[.B2]"])
    add(rows, "sheets.one_record_list_refusal", "SHEETS", ["(([.A1]~[.B2])![.A1])"])
    add(rows, "sheets_source_refusal", "SHEETS", ["['file:///book.ods'#.A1]"])
    add(
        rows,
        "sheets_selected_source_refusal",
        "SHEETS",
        ["IF(TRUE();['file:///book.ods'#.A1];[Data.A1])"],
    )
    add(
        rows,
        "sheets_selected_local_after_source",
        "SHEETS",
        ["IF(FALSE();['file:///book.ods'#.A1];[Data.A1])"],
    )
    add(
        rows,
        "sheets_iferror_source_refusal",
        "SHEETS",
        ["IFERROR(['file:///book.ods'#.A1];\"Main\")"],
    )
    add(
        rows,
        "sheets_ifna_source_refusal",
        "SHEETS",
        ["IFNA(['file:///book.ods'#.A1];\"Main\")"],
    )
    add(rows, "sheets_array_refusal", "SHEETS", ["{1;2|3;4}"])

    observed = {row["function"] for row in rows}
    if observed != set(FUNCTIONS):
        raise RuntimeError(f"function coverage changed: {sorted(observed)}")
    contract = HERE / "contract.md"
    contract_sha = hashlib.sha256(contract.read_bytes()).hexdigest() if contract.is_file() else None
    return {
        "schema": "ods-formula-reference-metadata-oracle-v1",
        "status": "independent normative/profile observations; production support pending",
        "baseline_commit": BASELINE,
        "normative_archive_sha256": ARCHIVE_SHA256,
        "normative_part4_sha256": PART4_SHA256,
        "contract_sha256": contract_sha,
        "functions": list(FUNCTIONS),
        "workbook": {
            "sheets": list(SHEETS),
            "sheet_numbers": {name: index + 1 for name, index in SHEET_INDEX.items()},
            "hidden_sheets_included": True,
            "extent": {"rows": 16, "columns": 12},
        },
        "profile": {
            "areas": "logical reference records; a 3-D record remains one area; list duplicates retained",
            "column_row": "one Reference record; omitted scalar uses current position; matrix result is full vector; explicit scalar projection selects the first axis element",
            "dimensions": "Reference or rectangular Array; ReferenceList is a type refusal",
            "sheet": "local sheet metadata; first sheet of a 3-D Reference; Text names a local sheet",
            "sheets": "local workbook count or inclusive 3-D reference span; hidden sheets included",
            "source_references": "ISREF classifies source descriptors; SHEET/SHEETS reject them as #VALUE!; no source is fetched",
            "cell_reads": "zero for every admitted metadata observation",
        },
        "observations": rows,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    data = document()
    payload = (json.dumps(data, ensure_ascii=False, indent=2) + "\n").encode()
    if args.write:
        GOLDENS.write_bytes(payload)
    if args.check:
        if not GOLDENS.is_file() or GOLDENS.read_bytes() != payload:
            raise SystemExit("reference metadata oracle bytes differ")
    print(
        json.dumps(
            {
                "functions": len({row["function"] for row in data["observations"]}),
                "observations": len(data["observations"]),
                "oracle_sha256": hashlib.sha256(payload).hexdigest(),
                "verified": args.check,
            },
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
