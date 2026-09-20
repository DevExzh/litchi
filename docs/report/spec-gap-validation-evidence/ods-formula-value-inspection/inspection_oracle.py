#!/usr/bin/env python3
"""Independent draft oracle for the sixteen value-inspection functions.

This module intentionally does not import the evaluator or copy its conversion
tables.  It is a small, deterministic Python profile derived from the local
ODF 1.4 entries.  The repository semantic contract is still pending, so
implementation-defined and locale-dependent choices are labelled as a draft
profile in the generated goldens.
"""

from __future__ import annotations

import argparse
from datetime import date
from decimal import Decimal
import hashlib
import json
import math
from pathlib import Path
import re


HERE = Path(__file__).resolve().parent
GOLDENS = HERE / "inspection-goldens.json"
ARCHIVE_SHA256 = (
    "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4"
)
CONTRACT_SHA256 = (
    "f227e861c55e3360d60d41a42ee7f101ba3c29bbc00d25d56cc01d80c3cd7923"
)
FUNCTIONS = (
    "ERROR.TYPE",
    "ISBLANK",
    "ISERR",
    "ISERROR",
    "ISEVEN",
    "ISLOGICAL",
    "ISNA",
    "ISNONTEXT",
    "ISNUMBER",
    "ISODD",
    "ISTEXT",
    "N",
    "NA",
    "NUMBERVALUE",
    "TYPE",
    "VALUE",
)


def value(kind: str, item: object) -> dict[str, object]:
    return {"type": kind, "value": item}


def error(code: str = "#VALUE!") -> dict[str, object]:
    return value("error", code)


EMPTY = {"type": "empty"}
NUMBER = lambda item: value("number", item)
TEXT = lambda item: value("text", item)
LOGICAL = lambda item: value("logical", item)
ERROR_NA = error("#N/A")
ERROR_DIV0 = error("#DIV/0!")
ERROR_VALUE = error("#VALUE!")


def is_error(item: dict[str, object]) -> bool:
    return item.get("type") == "error"


def inspect_function(name: str, item: dict[str, object]) -> dict[str, object]:
    kind = item.get("type")
    if name == "ISBLANK":
        return LOGICAL(kind == "empty")
    if name == "ISERR":
        return LOGICAL(is_error(item) and item.get("value") != "#N/A")
    if name == "ISERROR":
        return LOGICAL(is_error(item))
    if name == "ISLOGICAL":
        return LOGICAL(kind == "logical")
    if name == "ISNA":
        return LOGICAL(item == ERROR_NA)
    if name == "ISNONTEXT":
        return LOGICAL(kind != "text")
    if name == "ISNUMBER":
        return LOGICAL(kind == "number")
    if name == "ISTEXT":
        return LOGICAL(kind == "text")
    if name == "ERROR.TYPE":
        if not is_error(item):
            return error()
        return NUMBER(
            {
                "#NULL!": 1,
                "#DIV/0!": 2,
                "#VALUE!": 3,
                "#REF!": 4,
                "#NAME?": 5,
                "#NUM!": 6,
                "#N/A": 7,
            }.get(str(item["value"]), 3)
        )
    raise ValueError(name)


def type_code(item: dict[str, object]) -> dict[str, object]:
    if item.get("type") == "empty":
        return NUMBER(1)
    return NUMBER(
        {
            "number": 1,
            "text": 2,
            "logical": 4,
            "error": 16,
            "array": 64,
        }[str(item["type"])]
    )


def n_function(item: dict[str, object]) -> dict[str, object]:
    kind = item.get("type")
    if kind == "number":
        return item
    if kind == "logical":
        return NUMBER(1 if item["value"] else 0)
    if kind in ("text", "empty"):
        # Explicit draft choice for the ODF implementation-defined Text and
        # Empty cases.  The native receipt keeps these rows unresolved.
        return NUMBER(0)
    if kind == "error":
        return item
    raise ValueError(item)


def parity(name: str, item: dict[str, object]) -> dict[str, object]:
    kind = item.get("type")
    if is_error(item):
        return item
    if kind == "logical":
        number = 1.0 if item["value"] else 0.0
    elif kind == "empty":
        number = 0.0
    elif kind == "text":
        parsed = numbervalue(str(item["value"]))
        if parsed["type"] == "error":
            return parsed
        number = float(parsed["value"])
    elif kind == "number":
        number = float(item["value"])
    else:
        return error()
    if not math.isfinite(number):
        return error("#NUM!")
    truncated = math.trunc(number)
    even = abs(truncated) % 2 == 0
    return LOGICAL(even if name == "ISEVEN" else not even)


def numbervalue(text: str, decimal: str = ".", group: str = ",") -> dict[str, object]:
    """Evaluate the explicit-separator NUMBERVALUE draft profile."""

    if len(decimal) != 1 or decimal in group:
        return error()
    transformed = text
    decimal_index = transformed.find(decimal)
    if decimal_index < 0:
        transformed = transformed.replace(group, "")
    else:
        transformed = (
            transformed[:decimal_index].replace(group, "")
            + transformed[decimal_index:]
        )
    transformed = "".join(character for character in transformed if not character.isspace())
    transformed = transformed.replace(decimal, ".", 1)
    if transformed.startswith("."):
        transformed = "0" + transformed
    percent_count = 0
    while transformed.endswith("%"):
        percent_count += 1
        transformed = transformed[:-1]
    if transformed in {"INF", "+INF", "-INF", "NaN"}:
        return error("#NUM!")
    if not re.fullmatch(
        r"[+-]?(?:(?:[0-9]+(?:\.[0-9]*)?)|(?:\.[0-9]+))(?:[eE][+-]?[0-9]+)?",
        transformed,
    ):
        return error()
    try:
        result = float(transformed) / (100**percent_count)
    except ValueError:
        return error()
    return NUMBER(result) if math.isfinite(result) else error("#NUM!")


SERIAL_EPOCH = date(1899, 12, 30)
MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


def serial(year: int, month: int, day: int, fraction: float = 0.0) -> dict[str, object]:
    if (year, month, day) < (1899, 12, 30) or (year, month, day) > (9999, 12, 31):
        return error("#NUM!")
    try:
        result = float((date(year, month, day) - SERIAL_EPOCH).days) + fraction
    except ValueError:
        return error()
    return NUMBER(result)


def time_fraction(text: str) -> float | None:
    match = re.fullmatch(r"(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?", text)
    if match is None:
        return None
    hour = int(match.group(1))
    minute = int(match.group(2))
    try:
        second = Decimal(match.group(3) or "0")
    except ArithmeticError:
        return None
    if hour > 23 or minute > 59 or second >= Decimal(60):
        return None
    return float((Decimal(hour * 3600 + minute * 60) + second) / Decimal(86400))


def date_parts(text: str) -> tuple[int, int, int] | None:
    match = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", text)
    if match:
        return tuple(map(int, match.groups()))  # type: ignore[return-value]
    match = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", text)
    if match:
        month, day, year = map(int, match.groups())
        return year, month, day
    match = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2})", text)
    if match:
        month, day, short_year = map(int, match.groups())
        year = 1900 + short_year if short_year >= 30 else 2000 + short_year
        return year, month, day
    match = re.fullmatch(r"([A-Za-z]+) (\d{1,2}), (\d{4})", text)
    if match:
        month, day, year = match.groups()
        return int(year), MONTHS.get(month.lower(), 0), int(day)
    match = re.fullmatch(r"(\d{1,2}) ([A-Za-z]+) (\d{4})", text)
    if match:
        day, month, year = match.groups()
        return int(year), MONTHS.get(month.lower(), 0), int(day)
    return None


def value_function(text: str) -> dict[str, object]:
    """Evaluate the deterministic English/ISO VALUE draft profile."""

    text = text.strip()
    fraction = re.fullmatch(r"([+-]?)(\d+) (\d+)/([1-9]\d?)", text)
    if fraction:
        sign, whole, numerator, denominator = fraction.groups()
        result = int(whole) + int(numerator) / int(denominator)
        return NUMBER(-result if sign == "-" else result)

    if "T" in text or " " in text:
        match = re.fullmatch(r"(.+?)[ T](\d{1,2}:\d{1,2}(?::\d{1,2}(?:\.\d+)?)?)", text)
        if match:
            parts = date_parts(match.group(1))
            fraction = time_fraction(match.group(2))
            if parts is not None and fraction is not None:
                return serial(*parts, fraction)

    fraction = time_fraction(text)
    if fraction is not None:
        return NUMBER(fraction)
    parts = date_parts(text)
    if parts is not None:
        return serial(*parts)

    numeric = text
    percent = numeric.endswith("%")
    if percent:
        numeric = numeric[:-1]
    match = re.fullmatch(
        r"[+-]?(?:\$?\d+(?:,\d{3})*(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?",
        numeric,
    )
    if match is None:
        return error()
    numeric = numeric.replace(",", "").replace("$", "")
    try:
        result = float(numeric)
    except ValueError:
        return error()
    if percent:
        result /= 100
    return NUMBER(result) if math.isfinite(result) else error("#NUM!")


def formula_call(name: str, args: list[str]) -> str:
    rendered = ["TRUE()" if item == "TRUE" else "FALSE()" if item == "FALSE" else item for item in args]
    return "=" + name + "(" + ";".join(rendered) + ")"


def document() -> dict[str, object]:
    contract = HERE / "contract.md"
    if contract.is_file():
        digest = hashlib.sha256(contract.read_bytes()).hexdigest()
        if digest != CONTRACT_SHA256:
            raise RuntimeError(f"contract hash changed: {digest} != {CONTRACT_SHA256}")
    observations: list[dict[str, object]] = []

    def add(case: str, name: str, args: list[str], expected: dict[str, object]) -> None:
        observations.append(
            {
                "case": case,
                "function": name,
                "args": args,
                "formula": formula_call(name, args),
                "expected": expected,
            }
        )

    # Raw identity and error classification.
    add("blank.empty", "ISBLANK", ["<Empty>"], inspect_function("ISBLANK", EMPTY))
    add("blank.empty_text", "ISBLANK", ['""'], inspect_function("ISBLANK", TEXT("")))
    add("err.na", "ISERR", ["#N/A"], inspect_function("ISERR", ERROR_NA))
    add("err.div0", "ISERR", ["#DIV/0!"], inspect_function("ISERR", ERROR_DIV0))
    add("error.any", "ISERROR", ["#N/A"], inspect_function("ISERROR", ERROR_NA))
    add("error.type.na", "ERROR.TYPE", ["#N/A"], inspect_function("ERROR.TYPE", ERROR_NA))
    add("error.type.number", "ERROR.TYPE", ["42"], inspect_function("ERROR.TYPE", NUMBER(42)))
    add("logical.true", "ISLOGICAL", ["TRUE"], inspect_function("ISLOGICAL", LOGICAL(True)))
    add("logical.number", "ISLOGICAL", ["42"], inspect_function("ISLOGICAL", NUMBER(42)))
    add("na.only", "ISNA", ["#N/A"], inspect_function("ISNA", ERROR_NA))
    add("na.other_error", "ISNA", ["#DIV/0!"], inspect_function("ISNA", ERROR_DIV0))
    add("nont text", "ISNONTEXT", ["<Empty>"], inspect_function("ISNONTEXT", EMPTY))
    add("nont text_value", "ISNONTEXT", ['"x"'], inspect_function("ISNONTEXT", TEXT("x")))
    add("number.number", "ISNUMBER", ["42"], inspect_function("ISNUMBER", NUMBER(42)))
    add("number.error", "ISNUMBER", ["#N/A"], inspect_function("ISNUMBER", ERROR_NA))
    add("text.text", "ISTEXT", ['"x"'], inspect_function("ISTEXT", TEXT("x")))
    add("text.empty", "ISTEXT", ["<Empty>"], inspect_function("ISTEXT", EMPTY))

    # N and strict NA arity.
    add("n.number", "N", ["42"], n_function(NUMBER(42)))
    add("n.logical", "N", ["TRUE"], n_function(LOGICAL(True)))
    add("n.text_draft", "N", ['"x"'], n_function(TEXT("x")))
    add("n.empty_draft", "N", ["<Empty>"], n_function(EMPTY))
    add("n.error", "N", ["#N/A"], n_function(ERROR_NA))
    add("na.zero_arguments", "NA", [], ERROR_NA)

    # TYPE codes are normative for the four scalar categories and Array.
    add("type.number", "TYPE", ["42"], type_code(NUMBER(42)))
    add("type.text", "TYPE", ['"x"'], type_code(TEXT("x")))
    add("type.logical", "TYPE", ["TRUE"], type_code(LOGICAL(True)))
    add("type.error", "TYPE", ["#N/A"], type_code(ERROR_NA))
    add("type.empty", "TYPE", ["<Empty>"], type_code(EMPTY))
    add("type.array", "TYPE", ["{1;2}"], type_code(value("array", [1, 2])))

    # Numeric parity: truncation toward zero, signs, fractional boundaries and
    # exact binary64 integers around 2^53.
    for name, number, expected in (
        ("ISEVEN", 0, True),
        ("ISEVEN", 1, False),
        ("ISEVEN", -1, False),
        ("ISEVEN", 2.5, True),
        ("ISEVEN", -2.5, True),
        ("ISEVEN", 0.999, True),
        ("ISEVEN", -0.999, True),
        ("ISEVEN", 9007199254740992, True),
        ("ISEVEN", 9007199254740991, False),
        ("ISODD", 0, False),
        ("ISODD", 1, True),
        ("ISODD", -1, True),
        ("ISODD", 2.5, False),
        ("ISODD", -2.5, False),
        ("ISODD", 0.999, False),
        ("ISODD", -0.999, False),
        ("ISODD", 9007199254740992, False),
        ("ISODD", 9007199254740991, True),
        ):
        add(
            f"{name.lower()}.{number}",
            name,
            [str(number)],
            parity(name, NUMBER(number)),
        )

    add("iseven.logical", "ISEVEN", ["TRUE"], parity("ISEVEN", LOGICAL(True)))
    add("isodd.logical", "ISODD", ["TRUE"], parity("ISODD", LOGICAL(True)))
    add("iseven.empty", "ISEVEN", ["<Empty>"], parity("ISEVEN", EMPTY))
    add("isodd.empty", "ISODD", ["<Empty>"], parity("ISODD", EMPTY))
    add("iseven.text", "ISEVEN", ['"2"'], parity("ISEVEN", TEXT("2")))
    add("isodd.text", "ISODD", ['"2"'], parity("ISODD", TEXT("2")))
    add("iseven.malformed_text", "ISEVEN", ['"x"'], parity("ISEVEN", TEXT("x")))

    # NUMBERVALUE's ordered transform with explicit separators.  The omitted
    # separator defaults are deliberately outside the fixed oracle cases.
    for case, text, decimal, group in (
        ("us", "1,234.56", ".", ","),
        ("eu", "1.234,56", ",", "."),
        ("percent", " .5 % ", ".", ","),
        ("space_group", "1 234,50", ",", " "),
        ("leading_decimal", ".5", ".", ","),
        ("exponent", "1E3", ".", ","),
        ("repeat_percent", "12%%", ".", ","),
        ("group_after_decimal", "1.2,3", ".", ","),
        ("equal_separators", "1,234", ",", ","),
        ("nonfinite", "INF", ".", ","),
        ("malformed", "abc", ".", ","),
    ):
        args = [json.dumps(text), json.dumps(decimal), json.dumps(group)]
        add(f"numbervalue.{case}", "NUMBERVALUE", args, numbervalue(text, decimal, group))

    # Missing optional AST slots use the fixed profile defaults.  Empty Text
    # remains an explicit separator and follows its own validation path.
    optional_text = "1,234.5"
    for case, args, expected in (
        ("omitted", [json.dumps(optional_text)], numbervalue(optional_text)),
        ("omitted_slots", [json.dumps(optional_text), "", ""], numbervalue(optional_text)),
        (
            "omitted_group",
            [json.dumps(optional_text), json.dumps("."), ""],
            numbervalue(optional_text, ".", ","),
        ),
        (
            "omitted_decimal",
            [json.dumps(optional_text), "", json.dumps(",")],
            numbervalue(optional_text, ".", ","),
        ),
        (
            "explicit_empty_group",
            [json.dumps(optional_text), json.dumps("."), json.dumps("")],
            numbervalue(optional_text, ".", ""),
        ),
    ):
        add(f"numbervalue.{case}", "NUMBERVALUE", args, expected)

    # VALUE required grammar and deterministic serial epoch.
    value_inputs = (
        ("integer", "123"),
        ("leading_decimal", ".5"),
        ("signed_exponent", "-1.25E2"),
        ("signed_negative_exponent_lower", "-1e-3"),
        ("signed_negative_exponent_upper", "-1E-3"),
        ("grouped_signed_negative_exponent", "-1,234e-3"),
        ("currency_signed_negative_exponent", "-$1,234.5e-3"),
        ("percent", "12%"),
        ("currency", "$1,234.50"),
        ("unbounded_first_group", "1234,567"),
        ("fraction", "1 1/2"),
        ("negative_fraction", "-1 1/2"),
        ("fraction_99", "1 1/99"),
        ("fraction_100_rejected", "1 1/100"),
        ("time_hm", "2:00"),
        ("time_hms", "02:03:04"),
        ("time_fractional", "12:34:56.5"),
        ("time_fractional_two_digits", "12:34:56.12"),
        ("time_fractional_three_digits", "00:00:00.001"),
        ("time_fractional_near_day_end", "23:59:59.999999999999999999"),
        ("iso_date", "2024-02-29"),
        ("short_us_date", "5/21/06"),
        ("dash_us_date", "5-21-2006"),
        ("short_month_day_date", "29 Oct 2006"),
        ("date_1900_02_28", "1900-02-28"),
        ("date_1900_03_01", "1900-03-01"),
        ("iso_datetime", "2024-02-29 12:34:56"),
        ("iso_datetime_fractional", "2024-02-29 12:34:56.12"),
        ("iso_t_datetime", "2024-02-29T12:34:56"),
        ("us_date", "5/21/2006"),
        ("abbreviated_month_date", "Oct 29, 2006"),
        ("full_month_date", "29 October 2006"),
        ("malformed", "bogus"),
    )
    for case, text in value_inputs:
        add(f"value.{case}", "VALUE", [json.dumps(text)], value_function(text))
    add("value.empty", "VALUE", ["<Empty>"], NUMBER(0))
    add("value.empty_text", "VALUE", ['""'], error())
    add("value.logical", "VALUE", ["TRUE"], error())
    add("value.number", "VALUE", ["42"], error())

    observed_functions = {item["function"] for item in observations}
    if observed_functions != set(FUNCTIONS):
        raise RuntimeError(
            f"function coverage changed: {sorted(observed_functions)} != {list(FUNCTIONS)}"
        )
    return {
        "status": "semantic contract landed; production support pending",
        "normative_archive_sha256": ARCHIVE_SHA256,
        "contract_sha256": CONTRACT_SHA256,
        "profile": {
            "date_epoch": "1899-12-30; proleptic Gregorian; no 1900 leap-day insertion",
            "numeric_text": "invariant decimal/exponent plus the required en_US currency/grouping forms",
            "fraction": "mixed fraction with a space and denominator 1 through 99",
            "numbervalue": "explicit separators; whitespace removal; repeated trailing percent scaling",
            "logical_type_code": 4,
            "n_text_empty": "explicit zero choice from the value-inspection contract",
        },
        "functions": list(FUNCTIONS),
        "observations": observations,
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
        if GOLDENS.read_bytes() != payload:
            raise SystemExit("inspection oracle bytes differ")
    print(
        json.dumps(
            {
                "functions": len({item["function"] for item in data["observations"]}),
                "observations": len(data["observations"]),
                "oracle_sha256": hashlib.sha256(payload).hexdigest(),
                "verified": args.check,
            },
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
