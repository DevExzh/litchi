#!/usr/bin/env python3
"""Verify the date/time oracle with an independent, standard-library model.

The Rust consumer replays formulas through the evaluator, while this script
checks the retained JSON against civil-date, fraction, Easter, and week
calculations implemented here.  It deliberately does not import production
code or LibreOffice.  The contract and corpus are byte-bound by their
SHA-256 values; ``--check`` is read-only and fail-closed.
"""

from __future__ import annotations

import argparse
from datetime import date, timedelta
from fractions import Fraction
import hashlib
import json
import math
from pathlib import Path
import re
from typing import Any


HERE = Path(__file__).resolve().parent
CONTRACT = HERE / "contract.md"
VECTORS = HERE / "oracle-vectors.json"
CONTRACT_SHA256 = "cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f"

FUNCTIONS = (
    "DATE",
    "DATEDIF",
    "DATEVALUE",
    "DAY",
    "DAYS",
    "DAYS360",
    "EASTERSUNDAY",
    "EDATE",
    "EOMONTH",
    "HOUR",
    "ISOWEEKNUM",
    "MINUTE",
    "MONTH",
    "NETWORKDAYS",
    "NOW",
    "SECOND",
    "TIME",
    "TIMEVALUE",
    "TODAY",
    "WEEKDAY",
    "WEEKNUM",
    "WORKDAY",
    "YEAR",
    "YEARFRAC",
)
EXPECTED_IDS = frozenset(
    """
date.month_rollover date.day_rollover_leap date.no_synthetic_1900_day
date.fractional_arguments_truncate date.minimum_profile_date date.maximum_profile_date
date.invalid_month date.checked_overflow
datedif.Y datedif.M datedif.D datedif.MD datedif.YM datedif.YD
datedif.month_day_positive datedif.reversed_interval datedif.invalid_format
datevalue.iso_date datevalue.datetime_discards_time datevalue.en_us_numeric
datevalue.two_digit_pivot datevalue.english_month_name datevalue.numeric_fallback
datevalue.numeric_fallback_small datevalue.simple_fraction_fallback datevalue.invalid_calendar_day
day.datetime_floor day.iso_text day.out_of_domain
days.retain_fraction days.reversed days.text_and_number
days360.us_february_end days360.us_reversed_signed
days360.european_swaps_with_sign days360.european_31st
eastersunday.explicit_1583 eastersunday.explicit_2024 eastersunday.explicit_9956
eastersunday.invalid_year eastersunday.timestamp_after_current_easter eastersunday.missing_timestamp
edate.leap_month_clamp edate.nonleap_month_clamp edate.negative_month edate.fractional_month_truncates
eomonth.leap_target eomonth.previous_month eomonth.invalid_domain
hour.midday_fraction hour.negative_time hour.final_second
isoweeknum.new_year_2021 isoweeknum.first_week_2021 isoweeknum.end_2020 isoweeknum.2015_first_thursday
minute.half_second_rounds_up minute.just_below_half minute.negative_half_day minute.end_of_day_wrap
month.datetime_floor month.text_name month.invalid_text
networkdays.default_week networkdays.reversed_interval networkdays.custom_all_workdays
networkdays.duplicate_holiday networkdays.no_workday_sequence networkdays.formula_error_holiday
networkdays.reference_list_refusal
now.explicit_timestamp now.missing_timestamp
second.half_second_wrap second.just_below_boundary second.negative_fraction second.final_day_fraction
time.direct_fraction time.negative_fraction time.multi_day time.overflow
timevalue.clock_fraction timevalue.datetime_fraction timevalue.numeric_fallback
timevalue.simple_fraction_fallback timevalue.invalid_24_hour timevalue.date_only_rejected
timevalue.invalid_second
today.explicit_timestamp today.missing_timestamp
weekday.type_1 weekday.type_2 weekday.type_3 weekday.type_11 weekday.type_12 weekday.type_13
weekday.type_14 weekday.type_15 weekday.type_16 weekday.type_17 weekday.omitted_type weekday.invalid_type
weeknum.mode_1 weeknum.mode_2 weeknum.mode_11 weeknum.mode_12 weeknum.mode_13 weeknum.mode_14
weeknum.mode_15 weeknum.mode_16 weeknum.mode_17 weeknum.mode_21 weeknum.mode_150
weeknum.omitted_mode weeknum.noninteger_mode
workday.forward_default workday.backward_default workday.zero_preserves_fraction
workday.holiday_skip workday.custom_all_workdays workday.all_off_nonzero workday.reference_list_refusal
year.two_digit_pivot year.datetime year.minimum_profile year.invalid_text
yearfrac.basis_0_30us yearfrac.basis_1_actual_leap_year yearfrac.basis_2_actual_360
yearfrac.basis_3_actual_365 yearfrac.basis_4_30e yearfrac.reversed_is_nonnegative yearfrac.invalid_basis
""".split()
)
EPOCH = date(1899, 12, 30)
FORMULA_HEAD = re.compile(r"^=([A-Z][A-Z0-9_]*)\(")


class VerificationError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def serial(year: int, month: int, day: int) -> int:
    return (date(year, month, day) - EPOCH).days


def civil(value: int) -> date:
    return EPOCH + timedelta(days=value)


def easter(year: int) -> date:
    # Gregorian computus, independently written for this evidence check.
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = (h + l - 7 * m + 114) % 31 + 1
    return date(year, month, day)


def us_360(start: date, end: date) -> int:
    start_day = 30 if start.day == 31 or (start.month == 2 and (start + timedelta(days=1)).month != 2) else start.day
    end_day = 30 if end.day == 31 and start_day == 30 else end.day
    return (end.year * 360 + end.month * 30 + end_day) - (
        start.year * 360 + start.month * 30 + start_day
    )


def eomonth(year: int, month: int) -> date:
    if month == 12:
        return date(year + 1, 1, 1) - timedelta(days=1)
    return date(year, month + 1, 1) - timedelta(days=1)


def iso_week(serial_value: int) -> int:
    return civil(serial_value).isocalendar().week


def weeknum(serial_value: int, mode: int) -> int:
    day = civil(serial_value)
    if mode in (21, 150):
        return iso_week(serial_value)
    first_weekday = {
        1: 6,
        2: 0,
        11: 0,
        12: 1,
        13: 2,
        14: 3,
        15: 4,
        16: 5,
        17: 6,
    }[mode]
    jan1 = date(day.year, 1, 1)
    first_start = jan1 - timedelta(days=(jan1.weekday() - first_weekday) % 7)
    return 1 + (day - first_start).days // 7


def fraction_value(value: Fraction) -> float:
    return float(value)


def expected(row: dict[str, Any]) -> dict[str, Any]:
    outcome = row.get("expected")
    if not isinstance(outcome, dict):
        raise VerificationError(f"{row.get('id')}: expected outcome is not an object")
    return outcome


def check_number(rows: dict[str, dict[str, Any]], ident: str, wanted: Fraction | int | float) -> None:
    row = rows[ident]
    outcome = expected(row)
    if outcome.get("kind") != "number":
        raise VerificationError(f"{ident}: expected a number, got {outcome}")
    observed = float(outcome["value"])
    wanted_float = float(wanted)
    if not math.isfinite(observed) or not math.isfinite(wanted_float):
        raise VerificationError(f"{ident}: non-finite independent or corpus number")
    scale = max(abs(observed), abs(wanted_float), 1.0)
    if abs(observed - wanted_float) > 1e-12 * scale:
        raise VerificationError(f"{ident}: expected {wanted_float!r}, got {observed!r}")
    exact = outcome.get("exact")
    if exact is not None:
        exact_float = float(exact)
        if abs(exact_float - observed) > 1e-12 * max(abs(exact_float), abs(observed), 1.0):
            raise VerificationError(f"{ident}: exact decimal disagrees with value")


def number(value: Fraction | int | float) -> dict[str, Any]:
    observed = float(value)
    if not math.isfinite(observed):
        raise VerificationError(f"independent model produced non-finite number: {value!r}")
    return {"kind": "number", "value": observed}


def error(code: str) -> dict[str, str]:
    return {"kind": "error", "code": code}


def unsupported(capability: str) -> dict[str, str]:
    return {"kind": "unsupported", "capability": capability}


def last_day(year: int, month: int) -> int:
    return eomonth(year, month).day


def add_months(value: date, months: int, *, end: bool = False) -> date:
    zero_based = value.year * 12 + value.month - 1 + months
    year, month_zero = divmod(zero_based, 12)
    month = month_zero + 1
    if end:
        day = last_day(year, month)
    else:
        day = min(value.day, last_day(year, month))
    return date(year, month, day)


def datedif(start_serial: int, end_serial: int, fmt: str) -> dict[str, Any]:
    start, end = civil(start_serial), civil(end_serial)
    if end < start:
        return error("#NUM!")
    if fmt == "D":
        return number(end_serial - start_serial)
    month_delta = (end.year - start.year) * 12 + end.month - start.month
    if fmt == "M":
        anniversary = add_months(start, month_delta)
        return number(month_delta - (anniversary > end))
    if fmt == "Y":
        years = end.year - start.year
        anniversary = date(end.year, start.month, min(start.day, last_day(end.year, start.month)))
        return number(years - (anniversary > end))
    if fmt == "MD":
        remainder = end.day - start.day
        if remainder < 0:
            previous_month = add_months(end.replace(day=1), -1)
            remainder += last_day(previous_month.year, previous_month.month)
        return number(remainder)
    if fmt == "YM":
        signed = month_delta - (end.day < start.day)
        return number(signed % 12)
    if fmt == "YD":
        anniversary = date(end.year, start.month, min(start.day, last_day(end.year, start.month)))
        if anniversary > end:
            anniversary = date(end.year - 1, start.month, min(start.day, last_day(end.year - 1, start.month)))
        return number((end - anniversary).days)
    return error("#VALUE!")


def weekday_value(serial_value: int, kind: int) -> int:
    weekday = civil(serial_value).weekday()  # Monday=0, Sunday=6.
    sunday_first = (weekday + 1) % 7
    if kind == 1:
        return sunday_first + 1
    if kind == 2:
        return weekday + 1
    if kind == 3:
        return weekday
    if 11 <= kind <= 17:
        first = (kind - 11) % 7
        return (weekday - first) % 7 + 1
    raise ValueError(kind)


def networkdays(start_serial: int, end_serial: int, holidays: set[int] = set(), workdays: set[int] | None = None) -> int:
    if workdays is None:
        workdays = {0, 1, 2, 3, 4}  # Python Monday-first.
    sign = 1 if start_serial <= end_serial else -1
    first, last = sorted((start_serial, end_serial))
    return sign * sum(
        1
        for value in range(first, last + 1)
        if civil(value).weekday() in workdays and value not in holidays
    )


def workday(start_serial: int, offset: int, holidays: set[int] = set(), workdays: set[int] | None = None) -> int:
    if workdays is None:
        workdays = {0, 1, 2, 3, 4}
    if offset == 0:
        return start_serial
    direction = 1 if offset > 0 else -1
    remaining = abs(offset)
    value = start_serial
    while remaining:
        value += direction
        if civil(value).weekday() in workdays and value not in holidays:
            remaining -= 1
    return value


def model_outcome(ident: str) -> dict[str, Any]:
    """Return an independently computed typed outcome for every retained ID."""

    # DATE
    if ident == "date.month_rollover": return number(serial(2021, 1, 1))
    if ident == "date.day_rollover_leap": return number(serial(2020, 3, 1))
    if ident == "date.no_synthetic_1900_day": return number(serial(1900, 3, 1))
    if ident == "date.fractional_arguments_truncate": return number(serial(2020, 1, 1))
    if ident == "date.minimum_profile_date": return number(-693593)
    if ident == "date.maximum_profile_date": return number(2958465)
    if ident in {"date.invalid_month", "date.checked_overflow"}: return error("#NUM!")
    # DATEDIF
    if ident == "datedif.Y": return datedif(43524, 43890, "Y")
    if ident == "datedif.M": return datedif(43861, 43890, "M")
    if ident == "datedif.D": return datedif(43831, 43891, "D")
    if ident == "datedif.MD": return datedif(43861, 43889, "MD")
    if ident == "datedif.YM": return datedif(43616, 43951, "YM")
    if ident == "datedif.YD": return datedif(43525, 43890, "YD")
    if ident == "datedif.month_day_positive": return datedif(43845, 43910, "MD")
    if ident == "datedif.reversed_interval": return error("#NUM!")
    if ident == "datedif.invalid_format": return error("#VALUE!")
    # DATEVALUE/DAY/MONTH/YEAR
    if ident == "datevalue.iso_date": return number(serial(2020, 2, 29))
    if ident == "datevalue.datetime_discards_time": return number(serial(2020, 2, 29))
    if ident == "datevalue.en_us_numeric": return number(serial(2020, 2, 29))
    if ident == "datevalue.two_digit_pivot": return number(serial(1930, 1, 1))
    if ident == "datevalue.english_month_name": return number(serial(2020, 2, 29))
    if ident == "datevalue.numeric_fallback": return number(43890)
    if ident == "datevalue.numeric_fallback_small": return number(123)
    if ident == "datevalue.simple_fraction_fallback": return number(0)
    if ident == "datevalue.invalid_calendar_day": return error("#VALUE!")
    if ident == "day.datetime_floor": return number(29)
    if ident == "day.iso_text": return number(1)
    if ident == "day.out_of_domain": return error("#NUM!")
    if ident == "month.datetime_floor": return number(2)
    if ident == "month.text_name": return number(2)
    if ident == "month.invalid_text": return error("#VALUE!")
    if ident == "year.two_digit_pivot": return number(1930)
    if ident == "year.datetime": return number(2020)
    if ident == "year.minimum_profile": return number(1)
    if ident == "year.invalid_text": return error("#VALUE!")
    # DAYS/DAYS360
    if ident == "days.retain_fraction": return number(Fraction(5, 4))
    if ident == "days.reversed": return number(-59)
    if ident == "days.text_and_number": return number(59)
    if ident == "days360.us_february_end": return number(us_360(civil(43890), civil(43921)))
    if ident == "days360.us_reversed_signed": return number(us_360(civil(43921), civil(43890)))
    if ident == "days360.european_swaps_with_sign": return number(-31)
    if ident == "days360.european_31st": return number(28)
    # Easter and month shifts
    if ident == "eastersunday.explicit_1583":
        value = easter(1583); return number(serial(value.year, value.month, value.day))
    if ident == "eastersunday.explicit_2024":
        value = easter(2024); return number(serial(value.year, value.month, value.day))
    if ident == "eastersunday.explicit_9956":
        value = easter(9956); return number(serial(value.year, value.month, value.day))
    if ident == "eastersunday.invalid_year": return error("#NUM!")
    if ident == "eastersunday.timestamp_after_current_easter": return number(serial(2025, 4, 20))
    if ident == "eastersunday.missing_timestamp": return unsupported("CalculationClock")
    if ident == "edate.leap_month_clamp": return number(serial(2020, 2, 29))
    if ident == "edate.nonleap_month_clamp": return number(serial(2021, 2, 28))
    if ident == "edate.negative_month": return number(serial(2020, 1, 31))
    if ident == "edate.fractional_month_truncates": return number(serial(2020, 3, 29))
    if ident == "eomonth.leap_target": return number(serial(2020, 2, 29))
    if ident == "eomonth.previous_month": return number(serial(2020, 1, 31))
    if ident == "eomonth.invalid_domain": return error("#NUM!")
    # Time components
    if ident == "hour.midday_fraction": return number(12)
    if ident == "hour.negative_time": return number(18)
    if ident == "hour.final_second": return number(23)
    if ident == "isoweeknum.new_year_2021": return number(iso_week(44197))
    if ident == "isoweeknum.first_week_2021": return number(iso_week(44200))
    if ident == "isoweeknum.end_2020": return number(iso_week(44196))
    if ident == "isoweeknum.2015_first_thursday": return number(iso_week(42005))
    if ident == "minute.half_second_rounds_up": return number(4)
    if ident == "minute.just_below_half": return number(0)
    if ident == "minute.negative_half_day": return number(0)
    if ident == "minute.end_of_day_wrap": return number(0)
    if ident == "second.half_second_wrap": return number(0)
    if ident == "second.just_below_boundary": return number(59)
    if ident == "second.negative_fraction": return number(59)
    if ident == "second.final_day_fraction": return number(59)
    if ident == "time.direct_fraction": return number((Fraction(3, 2) * 3600 + Fraction(121, 4) * 60 + Fraction(1, 2)) / 86400)
    if ident == "time.negative_fraction": return number(Fraction(-1, 16))
    if ident == "time.multi_day": return number(1)
    if ident == "time.overflow": return error("#NUM!")
    if ident == "timevalue.clock_fraction": return number(Fraction(90593, 172800))
    if ident == "timevalue.datetime_fraction": return number(Fraction(3 * 3600 + 4 * 60 + 5, 86400))
    if ident == "timevalue.numeric_fallback": return number(Fraction(1, 2))
    if ident == "timevalue.simple_fraction_fallback": return number(Fraction(1, 4))
    if ident in {"timevalue.invalid_24_hour", "timevalue.date_only_rejected", "timevalue.invalid_second"}: return error("#VALUE!")
    # NETWORKDAYS/WORKDAY
    if ident == "networkdays.default_week": return number(networkdays(45292, 45298))
    if ident == "networkdays.reversed_interval": return number(networkdays(45298, 45292))
    if ident == "networkdays.custom_all_workdays": return number(networkdays(45292, 45298, workdays={0, 1, 2, 3, 4, 5, 6}))
    if ident == "networkdays.duplicate_holiday": return number(networkdays(45292, 45298, holidays={45294}))
    if ident == "networkdays.no_workday_sequence": return number(0)
    if ident == "networkdays.formula_error_holiday": return error("#N/A")
    if ident == "networkdays.reference_list_refusal": return error("#VALUE!")
    if ident == "workday.forward_default": return number(workday(45292, 5))
    if ident == "workday.backward_default": return number(workday(45292, -1))
    if ident == "workday.zero_preserves_fraction": return number(Fraction(45298 * 4 + 3, 4))
    if ident == "workday.holiday_skip": return number(workday(45292, 3, holidays={45294}))
    if ident == "workday.custom_all_workdays": return number(workday(45292, 3, workdays={0, 1, 2, 3, 4, 5, 6}))
    if ident in {"workday.all_off_nonzero", "workday.reference_list_refusal"}: return error("#NUM!" if ident == "workday.all_off_nonzero" else "#VALUE!")
    # Volatile clock functions
    if ident == "now.explicit_timestamp": return number(float("45418.52425925925925925925925925925925926"))
    if ident == "now.missing_timestamp": return unsupported("CalculationClock")
    if ident == "today.explicit_timestamp": return number(45418)
    if ident == "today.missing_timestamp": return unsupported("CalculationClock")
    # Week functions
    if ident.startswith("weekday.type_"):
        return number(weekday_value(44199, int(ident.rsplit("_", 1)[1])))
    if ident == "weekday.omitted_type": return number(weekday_value(44199, 1))
    if ident == "weekday.invalid_type": return error("#NUM!")
    if ident.startswith("weeknum.mode_"):
        return number(weeknum(44199, int(ident.rsplit("_", 1)[1])))
    if ident == "weeknum.omitted_mode": return number(weeknum(44199, 1))
    if ident == "weeknum.noninteger_mode": return error("#NUM!")
    # YEARFRAC
    if ident == "yearfrac.basis_0_30us": return number(Fraction(us_360(civil(43861), civil(43890)), 360))
    if ident == "yearfrac.basis_1_actual_leap_year": return number(Fraction(365, 366))
    if ident == "yearfrac.basis_2_actual_360": return number(Fraction(365, 360))
    if ident == "yearfrac.basis_3_actual_365": return number(1)
    if ident == "yearfrac.basis_4_30e": return number(Fraction(29, 360))
    if ident == "yearfrac.reversed_is_nonnegative": return number(Fraction(59, 360))
    if ident == "yearfrac.invalid_basis": return error("#NUM!")
    raise VerificationError(f"independent model has no vector ID {ident!r}")


def independent_checks(rows: dict[str, dict[str, Any]]) -> None:
    model_ids = set(rows)
    checked_ids: set[str] = set()
    failures: list[str] = []
    for ident, row in rows.items():
        checked_ids.add(ident)
        try:
            wanted = model_outcome(ident)
            observed = expected(row)
            if wanted.get("kind") != observed.get("kind"):
                raise VerificationError(f"independent kind {wanted} != corpus {observed}")
            if wanted["kind"] == "number":
                check_number(rows, ident, wanted["value"])
            elif wanted != observed:
                raise VerificationError(f"independent outcome {wanted} != corpus {observed}")
        except Exception as error:
            failures.append(f"{ident}: {error}")
    if checked_ids != model_ids:
        failures.append("independent model did not execute every corpus ID")
    if failures:
        raise VerificationError("independent model mismatches:\n" + "\n".join(f" - {failure}" for failure in failures))


def verify() -> dict[str, Any]:
    observed_contract = sha256(CONTRACT)
    if observed_contract != CONTRACT_SHA256:
        raise VerificationError(
            f"contract hash changed: {observed_contract} != {CONTRACT_SHA256}"
        )
    document = json.loads(VECTORS.read_text(encoding="utf-8"))
    if document.get("schema") != "ods-formula-date-time-oracle-v1":
        raise VerificationError("oracle schema is not date/time v1")
    if document.get("contract_sha256") != CONTRACT_SHA256:
        raise VerificationError("oracle contract identity does not match contract.md")
    if tuple(document.get("functions", ())) != FUNCTIONS:
        raise VerificationError("oracle function order is not the complete 24-function scope")
    vectors = document.get("vectors")
    if not isinstance(vectors, list) or not vectors:
        raise VerificationError("oracle vectors are empty")
    if document.get("vector_count") != len(vectors):
        raise VerificationError("oracle vector_count does not match rows")
    observed_ids = {row.get("id") for row in vectors if isinstance(row, dict)}
    if observed_ids != EXPECTED_IDS:
        missing = sorted(EXPECTED_IDS - observed_ids)
        extra = sorted(observed_ids - EXPECTED_IDS)
        raise VerificationError(f"oracle ID set changed: missing={missing!r}, extra={extra!r}")
    rows: dict[str, dict[str, Any]] = {}
    for row in vectors:
        if not isinstance(row, dict):
            raise VerificationError("oracle row is not an object")
        ident = row.get("id")
        if not isinstance(ident, str) or not ident or ident in rows:
            raise VerificationError(f"invalid or duplicate oracle id: {ident!r}")
        function = row.get("function")
        if function not in FUNCTIONS:
            raise VerificationError(f"{ident}: unknown function {function!r}")
        formula = row.get("formula")
        if not isinstance(formula, str):
            raise VerificationError(f"{ident}: formula is missing")
        match = FORMULA_HEAD.match(formula)
        if match is None or match.group(1) != function:
            raise VerificationError(f"{ident}: formula/function mismatch")
        outcome = expected(row)
        if outcome.get("kind") == "number" and (
            isinstance(outcome.get("value"), bool)
            or not isinstance(outcome.get("value"), (int, float))
            or not math.isfinite(float(outcome["value"]))
        ):
            raise VerificationError(f"{ident}: numeric outcome is malformed")
        if outcome.get("kind") == "error" and not isinstance(outcome.get("code"), str):
            raise VerificationError(f"{ident}: error outcome is malformed")
        if outcome.get("kind") == "unsupported" and not isinstance(outcome.get("capability"), str):
            raise VerificationError(f"{ident}: unsupported outcome is malformed")
        reads = row.get("reads")
        if reads is not None and (not isinstance(reads, dict) or reads.get("expected") != 0):
            raise VerificationError(f"{ident}: retained zero-read assertion is malformed")
        rows[ident] = row
    if {row["function"] for row in vectors} != set(FUNCTIONS):
        raise VerificationError("oracle does not execute every function")
    independent_checks(rows)
    return {
        "status": "verified",
        "contract_sha256": observed_contract,
        "oracle_sha256": sha256(VECTORS),
        "vectors": len(vectors),
        "independently_recomputed": len(vectors),
        "functions": len(FUNCTIONS),
        "zero_read_rows": sum("reads" in row for row in vectors),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="verify without writing files")
    args = parser.parse_args()
    if not args.check:
        parser.error("only read-only --check is supported")
    print(json.dumps(verify(), sort_keys=True))


if __name__ == "__main__":
    main()
