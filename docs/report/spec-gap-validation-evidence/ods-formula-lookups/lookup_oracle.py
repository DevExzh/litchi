#!/usr/bin/env python3
"""Small independent oracle for the ODF lookup/reference batch.

This module intentionally owns its value, array, address, and reference types.
It does not import the Rust evaluator and it does not call LibreOffice.  The
fixture is deliberately bounded: it exercises scalar results, inline arrays,
reference descriptors, Empty-versus-omitted arguments, and formula-error
search precedence against a finite resolver profile.

The script has two phases.  ``--self-check`` runs the model and checks all
invariant observations without writing a claim-bearing corpus.  ``--write``
and ``--check`` require a local ``contract.md`` and bind the JSON bytes to its
SHA-256.  This makes it impossible to accidentally publish a pre-contract
golden file.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import json
from pathlib import Path
import re
from typing import Final, Iterable


HERE = Path(__file__).resolve().parent
GOLDENS = HERE / "lookup-goldens.json"
CONTRACT = HERE / "contract.md"

ARCHIVE_SHA256: Final = (
    "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4"
)
PART4_SHA256: Final = (
    "ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1"
)
BASELINE: Final = "635fd2e1348b621426b50909cbd5765c91837306"
FUNCTIONS: Final = (
    "ADDRESS",
    "CHOOSE",
    "HLOOKUP",
    "INDEX",
    "INDIRECT",
    "LOOKUP",
    "MATCH",
    "OFFSET",
    "VLOOKUP",
)

# The lookup batch has its own pinned full C+F comparison profile.  The
# invariant corpus deliberately uses a small table of mappings rather than
# treating the Python runtime's Unicode version as normative.  Characters
# outside this table are rejected by ``lookup_fold`` until the contract binds
# a complete Unicode-17 CaseFolding extract.  No Unicode normalization is
# applied: in particular, ``İ`` and ``i`` + COMBINING DOT remain distinct.
PINNED_CASEFOLD_PROFILE: Final = (
    "Unicode17 stable fixture subset: ASCII, U+00DF/U+1E9E and "
    "U+03A3/U+03C2/U+03C3/U+0130/U+0307; no normalization"
)
PINNED_CASEFOLD: Final = {
    "ß": "ss",
    "ẞ": "ss",
    "Σ": "σ",
    "σ": "σ",
    "ς": "σ",
    # Preserve dotted capital I and the combining dot as distinct scalars;
    # this is intentionally not NFKC/NFKD or locale-aware lowercasing.
    "İ": "İ",
    "\u0307": "\u0307",
}


@dataclass(frozen=True)
class Scalar:
    kind: str
    value: object


@dataclass(frozen=True)
class Array:
    rows: int
    columns: int
    values: tuple[Scalar, ...]

    def __post_init__(self) -> None:
        if self.rows <= 0 or self.columns <= 0:
            raise ValueError("array dimensions must be positive")
        if len(self.values) != self.rows * self.columns:
            raise ValueError("array values do not match shape")

    def row(self, index: int) -> tuple[Scalar, ...]:
        start = index * self.columns
        return self.values[start : start + self.columns]

    def column(self, index: int) -> tuple[Scalar, ...]:
        return tuple(self.values[index + row * self.columns] for row in range(self.rows))


@dataclass(frozen=True)
class Reference:
    sheet: str
    row: int
    column: int
    rows: int = 1
    columns: int = 1
    sheet_end: str | None = None
    # ``direct`` means the retained descriptor still has a lexical owner;
    # ``derived`` is the descriptor produced by INDEX/OFFSET/INDIRECT.  The
    # distinction is part of the value contract and is checked by the Rust
    # consumer through ReferenceView::reference().
    owner: str = "direct"

    def shifted(self, row_delta: int, column_delta: int) -> "Reference":
        return Reference(
            self.sheet,
            self.row + row_delta,
            self.column + column_delta,
            self.rows,
            self.columns,
            self.sheet_end,
            "derived",
        )


@dataclass(frozen=True)
class ReferenceList:
    records: tuple[Reference, ...]


def number(value: int | float) -> Scalar:
    return Scalar("number", value)


def text(value: str) -> Scalar:
    return Scalar("text", value)


def logical(value: bool) -> Scalar:
    return Scalar("logical", value)


def empty() -> Scalar:
    return Scalar("empty", None)


def error(value: str) -> Scalar:
    return Scalar("error", value)


def array(rows: int, columns: int, values: Iterable[Scalar]) -> Array:
    return Array(rows, columns, tuple(values))


def array_expected(value: Array) -> dict[str, object]:
    return {
        "type": "array",
        "rows": value.rows,
        "columns": value.columns,
        "values": [scalar_expected(item) for item in value.values],
    }


def scalar_expected(value: Scalar) -> dict[str, object]:
    return {"type": value.kind, "value": value.value}


def reference_expected(value: Reference) -> dict[str, object]:
    # Coordinates in the retained JSON are one-based, matching formula
    # notation.  ``rows``/``columns`` describe the retained descriptor and do
    # not imply a cell read.
    area: dict[str, object] = {
        "sheet": value.sheet,
        "row": value.row,
        "column": value.column,
        "rows": value.rows,
        "columns": value.columns,
    }
    if value.sheet_end is not None:
        area["sheet_end"] = value.sheet_end
    return {
        "type": "reference",
        "areas": [area],
        "owner": value.owner,
    }


def reference_list_expected(value: ReferenceList) -> dict[str, object]:
    return {
        "type": "reference_list",
        "records": [
            {
                **reference_expected(record)["areas"][0],
                "owner": record.owner,
            }
            for record in value.records
        ],
    }


def ref(sheet: str, row: int, column: int, rows: int = 1, columns: int = 1,
        sheet_end: str | None = None, owner: str = "direct") -> Reference:
    return Reference(sheet, row, column, rows, columns, sheet_end, owner)


def ref_list(*records: Reference) -> ReferenceList:
    return ReferenceList(tuple(records))


def col_name(column: int) -> str:
    """Convert a positive one-based column number to A1 letters."""

    if column < 1:
        raise ValueError("A1 columns are one-based")
    letters: list[str] = []
    value = column
    while value:
        value, remainder = divmod(value - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def address(row: int, column: int, absolute: int, a1: bool,
            sheet: str | None = None) -> str:
    """Model ADDRESS for valid positive row/column inputs.

    The R1C1 branch is intentionally independent of a caller position.  A
    relative component retains the supplied row/column in brackets; it does
    not subtract an origin.  INDIRECT has a separate relative-coordinate
    parser below, because its R1C1 text is interpreted against the caller.
    """

    if row < 1 or column < 1 or absolute not in (1, 2, 3, 4):
        raise ValueError("invalid ADDRESS coordinate or absolute mode")
    prefix = ""
    if sheet:
        quoted = any(not (character.isascii() and
                          (character.isalnum() or character == "_"))
                     for character in sheet)
        escaped = sheet.replace("'", "''")
        prefix = (f"'{escaped}'" if quoted else sheet) + ("." if a1 else "!")
    if a1:
        column_text = col_name(column)
        row_text = str(row)
        if absolute in (1, 3):
            column_text = "$" + column_text
        if absolute in (1, 2):
            row_text = "$" + row_text
        return prefix + column_text + row_text

    row_text = str(row) if absolute in (1, 2) else f"[{row}]"
    column_text = str(column) if absolute in (1, 3) else f"[{column}]"
    return prefix + f"R{row_text}C{column_text}"


def choose(index: int, values: tuple[Scalar, ...]) -> Scalar:
    if index < 1 or index > len(values):
        raise ValueError("CHOOSE index outside supplied scalar options")
    return values[index - 1]


def index(array_value: Array, row: int, column: int) -> Scalar:
    if not (1 <= row <= array_value.rows and 1 <= column <= array_value.columns):
        raise ValueError("INDEX coordinate outside inline array")
    return array_value.values[(row - 1) * array_value.columns + column - 1]


def comparison_key(value: Scalar) -> tuple[int, object]:
    """Independent lookup ordering for the admitted scalar profile.

    The invariant corpus uses numbers, logicals, and a pinned Unicode text
    subset.  Lookup rows do not combine unlike types.  Mixed-type barriers and
    Empty/Error conversion are intentionally deferred to the contract review.
    """

    if value.kind == "empty":
        return 0, 0.0
    ranks = {"number": 0, "text": 1, "logical": 2}
    if value.kind not in ranks:
        raise ValueError(f"lookup comparison does not admit {value.kind}")
    comparable = lookup_fold(value.value) if value.kind == "text" else value.value
    return ranks[value.kind], comparable


def equal_lookup(left: Scalar, right: Scalar) -> bool:
    if left.kind == "empty" or right.kind == "empty":
        return comparison_key(left) == comparison_key(right)
    if left.kind != right.kind:
        return False
    if left.kind == "text":
        return lookup_fold(left.value) == lookup_fold(right.value)
    return left.value == right.value


def lookup_fold(value: str) -> str:
    """Fold only the independently pinned lookup fixture alphabet.

    ASCII lower-casing and the listed stable mappings cover every non-error
    text observation below.  Raising for another non-ASCII scalar prevents a
    newly assigned character in the host Python Unicode table from silently
    becoming a normative oracle result.
    """

    result: list[str] = []
    for character in value:
        if "A" <= character <= "Z":
            result.append(character.lower())
        elif character in PINNED_CASEFOLD:
            result.append(PINNED_CASEFOLD[character])
        elif ord(character) < 128:
            result.append(character)
        else:
            raise ValueError(
                f"text {value!r} uses an unpinned non-ASCII scalar U+{ord(character):04X}"
            )
    return "".join(result)


def match(lookup: Scalar, vector: tuple[Scalar, ...], match_type: int) -> int:
    if match_type == 0:
        for position, candidate in enumerate(vector, start=1):
            if equal_lookup(lookup, candidate):
                return position
        raise LookupError("#N/A")

    if match_type == 1:
        best: int | None = None
        lookup_rank = comparison_key(lookup)[0]
        for position, candidate in enumerate(vector, start=1):
            candidate_rank = comparison_key(candidate)[0]
            if candidate_rank > lookup_rank:
                break
            # A Text lookup may pass through lower-ranked numeric cells, but
            # it must not publish one of them as a fallback result.  The same
            # rule applies to Logical over numeric/text cells.
            if candidate_rank < lookup_rank:
                continue
            if comparison_key(candidate) <= comparison_key(lookup):
                best = position
            else:
                break
        if best is None:
            raise LookupError("#N/A")
        return best

    if match_type == -1:
        best: int | None = None
        lookup_rank = comparison_key(lookup)[0]
        for position, candidate in enumerate(vector, start=1):
            candidate_rank = comparison_key(candidate)[0]
            if candidate_rank > lookup_rank:
                # Higher-ranked cells precede this type in a valid descending
                # vector.  They cannot be a fallback, but the target type may
                # still occur later.
                continue
            if candidate_rank < lookup_rank:
                break
            if comparison_key(candidate) >= comparison_key(lookup):
                best = position
            else:
                break
        if best is None:
            raise LookupError("#N/A")
        return best
    raise ValueError("MATCH match_type must be -1, 0, or 1")


def table_lookup(
    lookup: Scalar,
    table: Array,
    result_index: int,
    horizontal: bool,
    approximate: bool,
) -> Scalar:
    line = table.row(0) if horizontal else table.column(0)
    if approximate:
        # The fixture uses ascending numeric keys.  Retaining the last
        # duplicate is the observable distinction from an exact first match.
        position = match(lookup, line, 1)
    else:
        position = match(lookup, line, 0)
    if horizontal:
        if result_index < 1 or result_index > table.rows:
            raise ValueError("HLOOKUP result row outside table")
        return table.row(result_index - 1)[position - 1]
    if result_index < 1 or result_index > table.columns:
        raise ValueError("VLOOKUP result column outside table")
    return table.column(result_index - 1)[position - 1]


def lookup(lookup_value: Scalar, lookup_vector: Array, result_vector: Array) -> Scalar:
    if lookup_vector.rows != 1 or result_vector.rows != 1:
        raise ValueError("invariant LOOKUP case uses horizontal vectors")
    position = match(lookup_value, lookup_vector.row(0), 1)
    return result_vector.row(0)[position - 1]


def indirect_r1c1(spelling: str, *, sheet: str, row: int, column: int) -> Reference:
    """Resolve the bounded R1C1 subset relative to a one-based caller."""

    pattern = re.fullmatch(r"R(\[(?P<row_delta>[+-]?\d+)\]|(?P<row_abs>\d+))"
                           r"C(\[(?P<column_delta>[+-]?\d+)\]|(?P<column_abs>\d+))",
                           spelling)
    if pattern is None:
        raise ValueError(f"unsupported R1C1 spelling {spelling!r}")
    target_row = (
        row + int(pattern.group("row_delta"))
        if pattern.group("row_delta") is not None
        else int(pattern.group("row_abs"))
    )
    target_column = (
        column + int(pattern.group("column_delta"))
        if pattern.group("column_delta") is not None
        else int(pattern.group("column_abs"))
    )
    if target_row < 1 or target_column < 1:
        raise ValueError("INDIRECT relative coordinate leaves the sheet")
    return Reference(sheet, target_row, target_column, owner="derived")


def indirect_a1(spelling: str, *, sheet: str) -> Reference:
    match_value = re.fullmatch(r"(?P<column>[A-Z]+)(?P<row>\d+)", spelling)
    if match_value is None:
        raise ValueError(f"unsupported A1 spelling {spelling!r}")
    column = 0
    for character in match_value.group("column"):
        column = column * 26 + ord(character) - ord("A") + 1
    return Reference(
        sheet, int(match_value.group("row")), column, owner="derived"
    )


def case(case_name: str, function: str, formula: str, expected: object,
         *, row: int = 5, column: int = 3, mode: str = "scalar",
         expected_reads: int | None = 0,
         expected_reads_min: int | None = None,
         expected_reads_max: int | None = None) -> dict[str, object]:
    if function not in FUNCTIONS:
        raise ValueError(function)
    result: dict[str, object] = {
        "case": case_name,
        "function": function,
        "formula": formula,
        "mode": mode,
        "position": {"sheet": "Main", "row": row, "column": column},
        "expected": expected,
        "expected_reads": expected_reads,
    }
    if expected_reads_min is not None:
        result["expected_reads_min"] = expected_reads_min
    if expected_reads_max is not None:
        result["expected_reads_max"] = expected_reads_max
    return result


def observations() -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []

    for absolute, suffix in ((1, "default"), (2, "row_absolute"),
                             (3, "column_absolute"), (4, "relative")):
        rows.append(case(
            f"address.a1.{suffix}",
            "ADDRESS",
            f"=ADDRESS(7;28;{absolute})",
            scalar_expected(Scalar("text", address(7, 28, absolute, True))),
        ))
        rows.append(case(
            f"address.r1c1.{suffix}",
            "ADDRESS",
            f"=ADDRESS(7;28;{absolute};FALSE())",
            scalar_expected(Scalar("text", address(7, 28, absolute, False))),
        ))

    options = (number(10), text("alpha"), number(30))
    for choice in (1, 2, 3):
        selected = choose(choice, options)
        rows.append(case(
            f"choose.scalar_{choice}",
            "CHOOSE",
            f'=CHOOSE({choice};10;"alpha";30)',
            scalar_expected(selected),
        ))

    h_table = array(2, 4, (number(1), number(2), number(2), number(3),
                           number(10), number(20), number(21), number(30)))
    rows.append(case(
        "hlookup.exact_first_duplicate",
        "HLOOKUP",
        "=HLOOKUP(2;{1;2;2;3|10;20;21;30};2;FALSE())",
        scalar_expected(table_lookup(number(2), h_table, 2, True, False)),
    ))
    rows.append(case(
        "hlookup.approximate_last_duplicate",
        "HLOOKUP",
        "=HLOOKUP(2.5;{1;2;2;3|10;20;21;30};2;TRUE())",
        scalar_expected(table_lookup(number(2.5), h_table, 2, True, True)),
    ))

    indexed = array(2, 2, (number(10), number(20), number(30), number(40)))
    simple_indexed = array(2, 2, (number(1), number(2), number(3), number(4)))
    rows.append(case(
        "index.array_explicit_row_column",
        "INDEX",
        "=INDEX({10;20|30;40};2;1)",
        scalar_expected(index(indexed, 2, 1)),
    ))

    rows.append(case(
        "indirect.a1_current_sheet",
        "INDIRECT",
        '=INDIRECT("B2")',
        reference_expected(indirect_a1("B2", sheet="Main")),
    ))
    rows.append(case(
        "indirect.r1c1_relative_to_caller",
        "INDIRECT",
        '=INDIRECT("R[1]C[2]";FALSE())',
        reference_expected(indirect_r1c1("R[1]C[2]", sheet="Main", row=5, column=3)),
    ))
    rows.append(case(
        "indirect.r1c1_absolute",
        "INDIRECT",
        '=INDIRECT("R6C5";FALSE())',
        reference_expected(indirect_r1c1("R6C5", sheet="Main", row=5, column=3)),
    ))

    lookup_keys = array(1, 3, (number(1), number(2), number(3)))
    lookup_values = array(1, 3, (number(10), number(20), number(30)))
    rows.append(case(
        "lookup.approximate_largest_less_equal",
        "LOOKUP",
        "=LOOKUP(2.5;{1;2;3};{10;20;30})",
        scalar_expected(lookup(number(2.5), lookup_keys, lookup_values)),
    ))

    match_vector = (number(10), number(20), number(20), number(30))
    rows.append(case(
        "match.exact_first_duplicate",
        "MATCH",
        "=MATCH(20;{10;20;20;30};0)",
        scalar_expected(number(match( number(20), match_vector, 0))),
    ))
    rows.append(case(
        "match.approximate_ascending",
        "MATCH",
        "=MATCH(25;{10;20;30};1)",
        scalar_expected(number(match(number(25), (number(10), number(20), number(30)), 1))),
    ))
    rows.append(case(
        "match.exact_text_casefold",
        "MATCH",
        '=MATCH("ALPHA";{"alpha";"Beta"};0)',
        scalar_expected(number(match(text("ALPHA"), (text("alpha"), text("Beta")), 0))),
    ))
    rows.append(case(
        "match.exact_text_sharp_s_expansion",
        "MATCH",
        '=MATCH("STRASSE";{"Straße";"other"};0)',
        scalar_expected(number(match(text("STRASSE"), (text("Straße"), text("other")), 0))),
    ))
    rows.append(case(
        "match.exact_text_sigma_final_sigma",
        "MATCH",
        '=MATCH("Σ";{"ς";"other"};0)',
        scalar_expected(number(match(text("Σ"), (text("ς"), text("other")), 0))),
    ))
    # Dotted I is kept as a deliberate no-normalization observation.  The
    # expected #N/A is a lookup miss, not an Empty/Error coercion case.
    try:
        match(text("i"), (text("İ"),), 0)
    except LookupError:
        dotted_i = Scalar("error", "#N/A")
    else:  # pragma: no cover - protects the pinned no-normalization invariant
        raise AssertionError("dotted I unexpectedly matched ASCII i")
    rows.append(case(
        "match.exact_text_dotted_i_no_normalization",
        "MATCH",
        '=MATCH("i";{"İ"};0)',
        scalar_expected(dotted_i),
    ))

    base = Reference("Main", 2, 2)
    shifted = base.shifted(1, 2)
    rows.append(case(
        "offset.checked_positive_shift",
        "OFFSET",
        "=OFFSET([.B2];1;2)",
        reference_expected(shifted),
    ))

    v_table = array(4, 2, (number(1), number(10), number(2), number(20),
                           number(2), number(21), number(3), number(30)))
    rows.append(case(
        "vlookup.exact_first_duplicate",
        "VLOOKUP",
        "=VLOOKUP(2;{1;10|2;20|2;21|3;30};2;FALSE())",
        scalar_expected(table_lookup(number(2), v_table, 2, False, False)),
    ))
    rows.append(case(
        "vlookup.approximate_last_duplicate",
        "VLOOKUP",
        "=VLOOKUP(2.5;{1;10|2;20|2;21|3;30};2;TRUE())",
        scalar_expected(table_lookup(number(2.5), v_table, 2, False, True)),
    ))

    # ADDRESS conversion, quoting, and matrix publication.
    rows.extend([
        case(
            "address.quoted_sheet_a1",
            "ADDRESS",
            '=ADDRESS(7;28;1;TRUE();"Data Sheet")',
            scalar_expected(Scalar("text", address(7, 28, 1, True, "Data Sheet"))),
        ),
        case(
            "address.apostrophe_escape_a1",
            "ADDRESS",
            '=ADDRESS(1;1;1;TRUE();"O\'Brien")',
            scalar_expected(Scalar("text", address(1, 1, 1, True, "O'Brien"))),
        ),
        case(
            "address.quoted_sheet_r1c1",
            "ADDRESS",
            '=ADDRESS(7;28;4;FALSE();"Data Sheet")',
            scalar_expected(Scalar("text", address(7, 28, 4, False, "Data Sheet"))),
        ),
        case(
            "address.empty_sheet_is_unqualified",
            "ADDRESS",
            '=ADDRESS(1;1;1;TRUE();"")',
            scalar_expected(Scalar("text", address(1, 1, 1, True, ""))),
        ),
        case(
            "address.finite_numeric_text_integer_conversion",
            "ADDRESS",
            '=ADDRESS("7";28)',
            scalar_expected(Scalar("text", address(7, 28, 1, True))),
        ),
        case(
            "address.logical_integer_conversion",
            "ADDRESS",
            "=ADDRESS(TRUE();1)",
            scalar_expected(Scalar("text", address(1, 1, 1, True))),
        ),
        case(
            "address.logical_text_style_refusal",
            "ADDRESS",
            '=ADDRESS(1;1;1;"1")',
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "address.missing_optional_slots_use_defaults",
            "ADDRESS",
            "=ADDRESS(1;1;;)",
            scalar_expected(Scalar("text", address(1, 1, 1, True))),
        ),
        case(
            "address.malformed_numeric_text_is_value_error",
            "ADDRESS",
            '=ADDRESS("x";1)',
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "address.matrix_row_parameter",
            "ADDRESS",
            "=ADDRESS({1|2};1)",
            array_expected(array(2, 1, (text("$A$1"), text("$A$2")))),
            mode="matrix",
        ),
        case(
            "address.invalid_row_is_value_error",
            "ADDRESS",
            "=ADDRESS(0;1)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "address.invalid_absolute_mode_is_value_error",
            "ADDRESS",
            "=ADDRESS(1;1;5)",
            scalar_expected(error("#VALUE!")),
        ),
    ])

    # CHOOSE keeps selected references/lists and performs matrix index
    # selection.  The out-of-range and selected-error rows are formula values;
    # unselected references are represented by the zero-read receipt.
    rows.extend([
        case(
            "choose.fractional_index_truncates",
            "CHOOSE",
            "=CHOOSE(2.9;10;20;30)",
            scalar_expected(number(20)),
        ),
        case(
            "choose.invalid_zero_index",
            "CHOOSE",
            "=CHOOSE(0;10;20)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "choose.selected_reference_identity",
            "CHOOSE",
            "=CHOOSE(2;42;[.A1])",
            reference_expected(ref("Main", 1, 1)),
        ),
        case(
            "choose.selected_reference_matrix_mode_preserves_descriptor",
            "CHOOSE",
            "=CHOOSE(2;42;[.A1])",
            reference_expected(ref("Main", 1, 1)),
            mode="matrix",
        ),
        case(
            "choose.selected_reference_list_identity",
            "CHOOSE",
            "=CHOOSE(2;42;[.A1]~[.B1])",
            reference_list_expected(ref_list(ref("Main", 1, 1), ref("Main", 1, 2))),
        ),
        case(
            "choose.matrix_index_selection",
            "CHOOSE",
            "=CHOOSE({1;3};10;20;30)",
            array_expected(array(1, 2, (number(10), number(30)))),
            mode="matrix",
        ),
        case(
            "choose.selected_formula_error_identity",
            "CHOOSE",
            "=CHOOSE(2;10;NA())",
            scalar_expected(error("#N/A")),
        ),
        case(
            "choose.unselected_missing_sheet_is_lazy",
            "CHOOSE",
            "=CHOOSE(1;42;[Missing.A1])",
            scalar_expected(number(42)),
        ),
    ])

    # INDEX array selectors, slices, logical list records and scalar/type
    # refusals.  The list and 3-D rows are descriptor-only and therefore read
    # no cells.
    rows.extend([
        case(
            "index.array_one_argument",
            "INDEX",
            "=INDEX({1;2|3;4})",
            array_expected(simple_indexed),
            mode="matrix",
        ),
        case(
            "index.array_zero_row_column_slice",
            "INDEX",
            "=INDEX({1;2|3;4};;2)",
            array_expected(array(2, 1, (number(2), number(4)))),
            mode="matrix",
        ),
        case(
            "index.array_zero_column_row_slice",
            "INDEX",
            "=INDEX({1;2|3;4};2;)",
            array_expected(array(1, 2, (number(3), number(4)))),
            mode="matrix",
        ),
        case(
            "index.array_both_zero_returns_complete_array",
            "INDEX",
            "=INDEX({1;2|3;4};;)",
            array_expected(simple_indexed),
            mode="matrix",
        ),
        case(
            "index.negative_row_is_value_error",
            "INDEX",
            "=INDEX({1;2|3;4};-1;1)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "index.row_beyond_shape_is_reference_error",
            "INDEX",
            "=INDEX({1;2|3;4};3;1)",
            scalar_expected(error("#REF!")),
        ),
        case(
            "index.array_area_number_must_be_one",
            "INDEX",
            "=INDEX({1;2|3;4};1;1;2)",
            scalar_expected(error("#REF!")),
        ),
        case(
            "index.scalar_data_is_value_error",
            "INDEX",
            "=INDEX(2;1;1)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "index.reference_slice_clears_owner",
            "INDEX",
            "=INDEX([.A1:.C4];2;3)",
            reference_expected(ref("Main", 2, 3, owner="derived")),
        ),
        case(
            "index.reference_list_second_record",
            "INDEX",
            "=INDEX([.A1:.B2]~[.C3:.D4];1;1;2)",
            reference_expected(ref("Main", 3, 3, owner="derived")),
        ),
        case(
            "index.reference_list_second_record_full_area",
            "INDEX",
            "=INDEX([.A1:.B2]~[.C3:.D4];0;0;2)",
            reference_expected(ref("Main", 3, 3, 2, 2, owner="derived")),
        ),
        case(
            "index.reference_list_duplicate_record_identity",
            "INDEX",
            "=INDEX([.A1]~[.A1];1;1;2)",
            reference_expected(ref("Main", 1, 1, owner="derived")),
        ),
        case(
            "index.three_dimensional_record_preserved",
            "INDEX",
            "=INDEX([Main.A1:Data.B2];1;1)",
            reference_expected(ref("Main", 1, 1, 1, 1, "Data", "derived")),
        ),
        case(
            "index.three_dimensional_full_area_preserved",
            "INDEX",
            "=INDEX([Main.A1:Data.B2];0;0)",
            reference_expected(ref("Main", 1, 1, 2, 2, "Data", "derived")),
        ),
    ])

    # OFFSET and INDIRECT are descriptor constructors.  The matrix/reference
    # rows are retained as geometry, while a real consumer is covered by the
    # later bounded read rows.
    rows.extend([
        case(
            "offset.missing_dimensions_use_original_shape",
            "OFFSET",
            "=OFFSET([.A1:.B2];0;0;;)",
            reference_expected(ref("Main", 1, 1, 2, 2, owner="derived")),
        ),
        case(
            "offset.explicit_positive_dimensions",
            "OFFSET",
            "=OFFSET([.B2];0;0;2;3)",
            reference_expected(ref("Main", 2, 2, 2, 3, owner="derived")),
        ),
        case(
            "offset.negative_shift",
            "OFFSET",
            "=OFFSET([.B2:.C3];-1;-1)",
            reference_expected(ref("Main", 1, 1, 2, 2, owner="derived")),
        ),
        case(
            "offset.zero_height_is_value_error",
            "OFFSET",
            "=OFFSET([.B2];0;0;0;1)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "offset.out_of_sheet_is_reference_error",
            "OFFSET",
            "=OFFSET([.A1];-1;0)",
            scalar_expected(error("#REF!")),
        ),
        case(
            "offset.reference_list_is_value_error_before_reads",
            "OFFSET",
            "=OFFSET([.A1]~[.B1];0;0)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "offset.three_dimensional_planes_preserved",
            "OFFSET",
            "=OFFSET([Main.A1:Data.B2];0;0)",
            reference_expected(ref("Main", 1, 1, 2, 2, "Data", "derived")),
        ),
        case(
            "offset.actual_empty_height_is_not_omitted",
            "OFFSET",
            "=OFFSET([.A1:.B2];0;0;[.J1];)",
            scalar_expected(error("#VALUE!")),
            expected_reads=1,
        ),
        case(
            "indirect.a1_range_descriptor",
            "INDIRECT",
            '=INDIRECT("B2:C3")',
            reference_expected(ref("Main", 2, 2, 2, 2, owner="derived")),
        ),
        case(
            "indirect.dot_sheet_separator",
            "INDIRECT",
            '=INDIRECT("Data.C3")',
            reference_expected(ref("Data", 3, 3, owner="derived")),
        ),
        case(
            "indirect.exclamation_sheet_separator",
            "INDIRECT",
            '=INDIRECT("Data!C3")',
            reference_expected(ref("Data", 3, 3, owner="derived")),
        ),
        case(
            "indirect.three_dimensional_descriptor",
            "INDIRECT",
            '=INDIRECT("Main.A1:Data.B2")',
            reference_expected(ref("Main", 1, 1, 2, 2, "Data", "derived")),
        ),
        case(
            "indirect.whole_column_descriptor",
            "INDIRECT",
            '=INDIRECT("A:A")',
            reference_expected(ref("Main", 1, 1, 16, 1, owner="derived")),
        ),
        case(
            "indirect.whole_row_descriptor",
            "INDIRECT",
            '=INDIRECT("1:1")',
            reference_expected(ref("Main", 1, 1, 1, 24, owner="derived")),
        ),
        case(
            "indirect.r1c1_absolute_descriptor",
            "INDIRECT",
            '=INDIRECT("R6C5";FALSE())',
            reference_expected(ref("Main", 6, 5, owner="derived")),
        ),
        case(
            "indirect.r1c1_relative_descriptor",
            "INDIRECT",
            '=INDIRECT("R[1]C[2]";FALSE())',
            reference_expected(ref("Main", 6, 5, owner="derived")),
        ),
        case(
            "indirect.r1c1_omitted_current_column",
            "INDIRECT",
            '=INDIRECT("R[-1]C";FALSE())',
            reference_expected(ref("Main", 4, 3, owner="derived")),
        ),
        case(
            "indirect.array_text_uses_first_element_in_scalar_mode",
            "INDIRECT",
            '=INDIRECT({"B2";"C3"})',
            reference_expected(ref("Main", 2, 2, owner="derived")),
        ),
        case(
            "indirect.malformed_text_is_reference_error",
            "INDIRECT",
            '=INDIRECT("not-a-reference")',
            scalar_expected(error("#REF!")),
        ),
    ])

    # Search rows use both inline arrays (read-free matrix key/index cases)
    # and the fixed four-sheet resolver profile described in the contract.
    rows.extend([
        case(
            "match.descending_last_duplicate",
            "MATCH",
            "=MATCH(2;{4;2;2;1};-1)",
            scalar_expected(number(3)),
        ),
        case(
            "match.invalid_match_type",
            "MATCH",
            "=MATCH(2;{1;2;3};2)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "match.finite_numeric_text_match_type",
            "MATCH",
            '=MATCH(2;{1;2;3};"0")',
            scalar_expected(number(2)),
        ),
        case(
            "match.scalar_data_refusal",
            "MATCH",
            "=MATCH(2;2;0)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "match.reference_list_data_refusal",
            "MATCH",
            "=MATCH(2;[.A1]~[.A2];0)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "match.two_dimensional_data_refusal",
            "MATCH",
            "=MATCH(2;{1;2|3;4};0)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "match.number_text_barrier",
            "MATCH",
            '=MATCH(2;{"1";"2"};0)',
            scalar_expected(error("#N/A")),
        ),
        case(
            "match.ascending_text_key_does_not_fallback_to_numbers",
            "MATCH",
            '=MATCH("2";{1;2;3};1)',
            scalar_expected(error("#N/A")),
        ),
        case(
            "match.descending_number_key_does_not_fallback_to_text",
            "MATCH",
            '=MATCH(2;{"3";"2";"1"};-1)',
            scalar_expected(error("#N/A")),
        ),
        case(
            "match.ascending_mixed_type_midpoint",
            "MATCH",
            '=MATCH("b";{1|"a"|"c"};1)',
            scalar_expected(number(2)),
        ),
        case(
            "match.descending_mixed_type_midpoint",
            "MATCH",
            '=MATCH(2;{"z"|2|1};-1)',
            scalar_expected(number(2)),
        ),
        case(
            "match.exact_mixed_midpoint",
            "MATCH",
            "=MATCH(1;[.J1:.J4];0)",
            scalar_expected(number(2)),
            expected_reads=2,
        ),
        case(
            "match.logical_ascending_order",
            "MATCH",
            "=MATCH(TRUE();{FALSE();TRUE()};1)",
            scalar_expected(number(2)),
        ),
        case(
            "match.matrix_key_iteration",
            "MATCH",
            "=MATCH({1;4};{1;2;2;4};0)",
            array_expected(array(1, 2, (number(1), number(4)))),
            mode="matrix",
        ),
        case(
            "match.empty_search_cell_normalizes_to_zero",
            "MATCH",
            "=MATCH(0;[.J1:.J4];0)",
            scalar_expected(number(1)),
            expected_reads=1,
        ),
        case(
            "match.empty_search_key_normalizes_to_zero",
            "MATCH",
            "=MATCH([.J1];[.A1:.A4];0)",
            scalar_expected(error("#N/A")),
            expected_reads=5,
        ),
        case(
            "match.visited_formula_error_propagates",
            "MATCH",
            "=MATCH(2;[.L1:.L2];0)",
            scalar_expected(error("#N/A")),
            expected_reads=2,
        ),
        case(
            "match.formula_error_scans_suffix_before_returning_error",
            "MATCH",
            "=MATCH(2;[.L1:.L3];0)",
            scalar_expected(error("#N/A")),
            expected_reads=3,
        ),
        case(
            "match.approximate_reference_reads_are_bounded",
            "MATCH",
            "=MATCH(3;[.A1:.A4];1)",
            scalar_expected(number(3)),
            expected_reads=None,
            expected_reads_min=2,
            expected_reads_max=4,
        ),
        case(
            "hlookup.scalar_data_refusal",
            "HLOOKUP",
            "=HLOOKUP(2;2;1;FALSE())",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "hlookup.selector_bounds_before_reads",
            "HLOOKUP",
            "=HLOOKUP(2;{1;2|10;20};3;FALSE())",
            scalar_expected(error("#REF!")),
        ),
        case(
            "hlookup.matrix_key_iteration",
            "HLOOKUP",
            "=HLOOKUP({1;4};{1;2;2;4|10;20;21;40};2;FALSE())",
            array_expected(array(1, 2, (number(10), number(40)))),
            mode="matrix",
        ),
        case(
            "hlookup.matrix_index_iteration",
            "HLOOKUP",
            "=HLOOKUP(2;{1;2;2;4|10;20;21;40};{1;2};FALSE())",
            array_expected(array(1, 2, (number(2), number(20)))),
            mode="matrix",
        ),
        case(
            "hlookup.reference_exact_first_duplicate",
            "HLOOKUP",
            "=HLOOKUP(2;[.E1:.H3];2;FALSE())",
            scalar_expected(text("two-first")),
            expected_reads=3,
        ),
        case(
            "hlookup.reference_approximate_duplicate_last",
            "HLOOKUP",
            "=HLOOKUP(3;[.E1:.H3];2;TRUE())",
            scalar_expected(text("two-last")),
            expected_reads=None,
            expected_reads_min=3,
            expected_reads_max=5,
        ),
        case(
            "hlookup.reference_list_data_refusal",
            "HLOOKUP",
            "=HLOOKUP(2;[.E1:.H3]~[.A1:.C4];2;FALSE())",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "vlookup.scalar_data_refusal",
            "VLOOKUP",
            "=VLOOKUP(2;2;1;FALSE())",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "vlookup.selector_bounds_before_reads",
            "VLOOKUP",
            "=VLOOKUP(2;{1;10|2;20};3;FALSE())",
            scalar_expected(error("#REF!")),
        ),
        case(
            "vlookup.reference_exact_first_duplicate",
            "VLOOKUP",
            "=VLOOKUP(2;[.A1:.C4];2;FALSE())",
            scalar_expected(text("two-first")),
            expected_reads=3,
        ),
        case(
            "vlookup.reference_approximate_duplicate_last",
            "VLOOKUP",
            "=VLOOKUP(3;[.A1:.C4];2;TRUE())",
            scalar_expected(text("two-last")),
            expected_reads=None,
            expected_reads_min=3,
            expected_reads_max=5,
        ),
        case(
            "vlookup.logical_text_range_lookup_refusal",
            "VLOOKUP",
            '=VLOOKUP(2;{1;10|2;20};2;"1")',
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "vlookup.referenced_empty_range_lookup_is_false",
            "VLOOKUP",
            "=VLOOKUP(2;{1;10|2;20};2;[.J1])",
            scalar_expected(number(20)),
            expected_reads=1,
        ),
        case(
            "vlookup.matrix_key_iteration",
            "VLOOKUP",
            "=VLOOKUP({1;4};{1;10|2;20|2;21|4;40};2;FALSE())",
            array_expected(array(1, 2, (number(10), number(40)))),
            mode="matrix",
        ),
        case(
            "vlookup.matrix_index_iteration",
            "VLOOKUP",
            "=VLOOKUP(2;{1;10|2;20|2;21|4;40};{1;2};FALSE())",
            array_expected(array(1, 2, (number(2), number(20)))),
            mode="matrix",
        ),
        case(
            "vlookup.reference_result_empty_is_preserved",
            "VLOOKUP",
            "=VLOOKUP(0;[.J1:.K4];2;FALSE())",
            scalar_expected(empty()),
            expected_reads=2,
        ),
        case(
            "vlookup.reference_result_error_is_preserved",
            "VLOOKUP",
            "=VLOOKUP(2;[.M1:.N2];2;FALSE())",
            scalar_expected(error("#N/A")),
            expected_reads=2,
        ),
    ])

    # LOOKUP orientation, vector checks, and the contract's lazy short-result
    # extension policy.  Inline short Arrays fail at the selected index;
    # short References extend only when the match actually needs it.
    rows.extend([
        case(
            "lookup.scalar_data_refusal",
            "LOOKUP",
            "=LOOKUP(2;2)",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "lookup.tall_orientation",
            "LOOKUP",
            "=LOOKUP(3;[.A1:.B4])",
            scalar_expected(text("two-last")),
            expected_reads=None,
            expected_reads_min=3,
            expected_reads_max=5,
        ),
        case(
            "lookup.wide_orientation",
            "LOOKUP",
            "=LOOKUP(3;[.E1:.H2])",
            scalar_expected(text("two-last")),
            expected_reads=None,
            expected_reads_min=3,
            expected_reads_max=5,
        ),
        case(
            "lookup.short_array_result_is_na_only_when_selected",
            "LOOKUP",
            "=LOOKUP(4;{1;2;2;4};{10;20})",
            scalar_expected(error("#N/A")),
        ),
        case(
            "lookup.result_must_be_vector",
            "LOOKUP",
            "=LOOKUP(2;{1;2};{10;20|30;40})",
            scalar_expected(error("#VALUE!")),
        ),
        case(
            "lookup.short_reference_result_extends_for_needed_match",
            "LOOKUP",
            "=LOOKUP(4;[.A1:.A4];[.B1:.B2])",
            scalar_expected(text("four")),
            expected_reads=None,
            expected_reads_min=3,
            expected_reads_max=6,
        ),
        case(
            "lookup.short_reference_result_not_extended_for_in_range_match",
            "LOOKUP",
            "=LOOKUP(1;[.A1:.A4];[.B1:.B2])",
            scalar_expected(text("one")),
            expected_reads=None,
            expected_reads_min=2,
            expected_reads_max=5,
        ),
        case(
            "lookup.long_reference_result_keeps_selected_duplicate",
            "LOOKUP",
            "=LOOKUP(2;[.A1:.A4];[.B1:.B4])",
            scalar_expected(text("two-last")),
            expected_reads=None,
            expected_reads_min=3,
            expected_reads_max=6,
        ),
        case(
            "lookup.reference_extension_sheet_bound_is_na",
            "LOOKUP",
            "=LOOKUP(4;[.A1:.A4];[.B16])",
            scalar_expected(error("#N/A")),
            expected_reads=None,
            expected_reads_min=3,
            expected_reads_max=5,
        ),
        case(
            "lookup.empty_key_and_empty_result",
            "LOOKUP",
            "=LOOKUP(0;[.J1:.J4];[.K1:.K4])",
            scalar_expected(empty()),
            expected_reads=None,
            expected_reads_min=2,
            expected_reads_max=5,
        ),
        case(
            "lookup.matrix_key_iteration",
            "LOOKUP",
            "=LOOKUP({1;4};{1;2;2;4};{10;20;21;40})",
            array_expected(array(1, 2, (number(10), number(40)))),
            mode="matrix",
        ),
    ])
    return rows


def contract_sha256() -> str | None:
    if not CONTRACT.is_file():
        return None
    return hashlib.sha256(CONTRACT.read_bytes()).hexdigest()


def document(*, require_contract: bool) -> dict[str, object]:
    digest = contract_sha256()
    if require_contract and digest is None:
        raise SystemExit(
            f"refusing to publish lookup goldens: missing contract {CONTRACT}"
        )
    return {
        "schema": "ods-formula-lookups-oracle-v0",
        "status": (
            "contract-bound independent observations"
            if digest is not None
            else "pending contract; self-check only"
        ),
        "baseline_commit": BASELINE,
        "normative_archive_sha256": ARCHIVE_SHA256,
        "normative_part4_sha256": PART4_SHA256,
        "contract_sha256": digest,
        "functions": list(FUNCTIONS),
        "workbook": {
            "sheets": ["Main", "Data", "Hidden", "Archive"],
            "extent": {"rows": 16, "columns": 24},
            "empty_search_range": "Main.J1:J4",
            "formula_error_search_range": "Main.L1:L3",
            "formula_error_result_range": "Main.M1:N2",
        },
        "profile": {
            "address_r1c1": (
                "relative components retain supplied row/column in brackets; "
                "ADDRESS does not subtract an origin"
            ),
            "indirect_r1c1": "relative brackets are offsets from caller position",
            "lookup_text": PINNED_CASEFOLD_PROFILE,
            "duplicate_ties": "exact first match; ascending approximate last duplicate",
            "empty_profile": "Empty search keys/cells normalize to Number 0; Empty results remain Empty",
            "read_profile": "inline arrays and descriptor-only cases require zero resolver reads",
            "formula_error_scan": "retain first visited error and scan suffix before publication",
            "deferred": "typed provider/resource failure receipts remain in focused resource tests",
        },
        "observations": observations(),
    }


def canonical_bytes(data: dict[str, object]) -> bytes:
    return (json.dumps(data, ensure_ascii=False, indent=2) + "\n").encode()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--self-check", action="store_true",
                        help="run invariant model checks without writing goldens")
    parser.add_argument("--write", action="store_true",
                        help="write contract-bound lookup-goldens.json")
    parser.add_argument("--check", action="store_true",
                        help="compare retained goldens with regenerated contract-bound bytes")
    args = parser.parse_args()
    if not (args.self_check or args.write or args.check):
        parser.error("choose --self-check, --write, or --check")
    data = document(require_contract=args.write or args.check)
    payload = canonical_bytes(data)
    if args.write:
        GOLDENS.write_bytes(payload)
    if args.check:
        if not GOLDENS.is_file():
            raise SystemExit(f"missing retained goldens: {GOLDENS}")
        if GOLDENS.read_bytes() != payload:
            raise SystemExit("lookup oracle bytes differ from retained goldens")
    print(json.dumps({
        "contract_bound": data["contract_sha256"] is not None,
        "functions": len(data["functions"]),
        "observations": len(data["observations"]),
        "goldens": str(GOLDENS) if args.write or args.check else None,
        "verified": bool(args.check),
    }, sort_keys=True))


if __name__ == "__main__":
    main()
