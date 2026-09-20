#!/usr/bin/env python3
"""Generate and verify the independent IRR/RATE/XIRR root corpus.

The root functions have no closed form in ODF. This oracle therefore keeps
two responsibilities separate:

* exact, analytically constructed vectors validate the contract equations with
  high-precision Decimal residuals; and
* bounded-search rows explicitly record when the contract does not guarantee
  which root (or even a successful bracket) a finite guess will select.

The module imports no Rust evaluator and no spreadsheet host. The write mode
regenerates the adjacent JSON corpus from this source and the repository-local
contract/source hashes. The check mode is read-only, rejects stale source pins,
duplicate IDs, altered cardinality, and any changed vector, then revalidates
all analytic residuals.
"""

from __future__ import annotations

import argparse
from decimal import Context, Decimal, InvalidOperation, ROUND_HALF_EVEN, localcontext
import hashlib
import json
from pathlib import Path
from typing import Any, Iterable
import zipfile


HERE = Path(__file__).resolve().parent
CONTRACT = HERE / "contract.md"
DATE_TIME_CONTRACT = HERE.parent / "ods-formula-date-time" / "contract.md"
VECTORS = HERE / "oracle-root-vectors.json"
NORMATIVE_ARCHIVE = HERE.parents[3] / "3rdparty/specs/OpenDocument-v1.4-os.zip"
NORMATIVE_MEMBER = "part4-formula/OpenDocument-v1.4-os-part4-formula.html"

NORMATIVE_ARCHIVE_SHA256 = (
    "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4"
)
NORMATIVE_PART4_SHA256 = (
    "ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1"
)

# Decimal precision is deliberately independent of binary64 solver arithmetic.
PRECISION = 180
CONTEXT = Context(
    prec=PRECISION,
    rounding=ROUND_HALF_EVEN,
    Emax=999999,
    Emin=-999999,
)

FUNCTIONS = ("IRR", "RATE", "XIRR")
EXPECTED_COUNTS = {"IRR": 10, "RATE": 10, "XIRR": 11}
EXPECTED_VECTOR_COUNT = sum(EXPECTED_COUNTS.values())
NUM = "#NUM!"
VALUE = "#VALUE!"
DATE_MIN = Decimal("-693593")
DATE_MAX_EXCLUSIVE = Decimal("2958466")
ROOT_ABS_TOLERANCE = Decimal("1e-120")
ROOT_REL_TOLERANCE = Decimal("1e-120")


class OracleError(Exception):
    """An error in the independent mathematical model."""

    def __init__(self, value: str):
        super().__init__(value)
        self.value = value


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def normative_member_digest() -> str:
    with zipfile.ZipFile(NORMATIVE_ARCHIVE) as archive:
        return hashlib.sha256(archive.read(NORMATIVE_MEMBER)).hexdigest()


def source_hashes() -> dict[str, str]:
    archive_hash = digest(NORMATIVE_ARCHIVE)
    member_hash = normative_member_digest()
    if archive_hash != NORMATIVE_ARCHIVE_SHA256:
        raise RuntimeError(
            f"normative archive hash changed: {archive_hash} != "
            f"{NORMATIVE_ARCHIVE_SHA256}"
        )
    if member_hash != NORMATIVE_PART4_SHA256:
        raise RuntimeError(
            f"normative Part 4 hash changed: {member_hash} != "
            f"{NORMATIVE_PART4_SHA256}"
        )
    return {
        "financial_contract_sha256": digest(CONTRACT),
        "date_time_contract_sha256": digest(DATE_TIME_CONTRACT),
        "normative_archive_sha256": archive_hash,
        "normative_part4_member_sha256": member_hash,
    }


def d(value: str | int | Decimal) -> Decimal:
    """Parse an exact oracle operand, without importing implementation state."""

    try:
        result = Decimal(str(value))
    except (InvalidOperation, ValueError) as error:
        raise OracleError(NUM) from error
    if not result.is_finite():
        raise OracleError(NUM)
    return result


def binary64_decimal(value: str | int | Decimal) -> Decimal:
    """Model a finite formula Number after its binary64 input conversion."""

    try:
        result = Decimal.from_float(float(str(value)))
    except (OverflowError, ValueError) as error:
        raise OracleError(NUM) from error
    if not result.is_finite():
        raise OracleError(NUM)
    return result


def finite(value: Decimal) -> Decimal:
    if not value.is_finite():
        raise OracleError(NUM)
    return value


def real_power(base: Decimal, exponent: Decimal) -> Decimal:
    """High-precision real power with the profile's signed integer branch."""

    with localcontext(CONTEXT):
        if exponent == exponent.to_integral_value():
            try:
                return finite(base ** int(exponent))
            except (ArithmeticError, OverflowError) as error:
                raise OracleError(NUM) from error
        if base <= 0:
            raise OracleError(NUM)
        try:
            return finite((CONTEXT.ln(base) * exponent).exp(context=CONTEXT))
        except (ArithmeticError, OverflowError) as error:
            raise OracleError(NUM) from error


def scaled(values: Iterable[Decimal]) -> Decimal:
    return sum((abs(value) for value in values), Decimal(0))


def irr_residual(values: list[Decimal], rate: Decimal) -> tuple[Decimal, Decimal]:
    """Return residual and absolute term scale using contract i=1 indexing."""

    with localcontext(CONTEXT):
        base = 1 + rate
        if base == 0:
            raise OracleError(NUM)
        terms = [
            value * real_power(base, -Decimal(index))
            for index, value in enumerate(values, start=1)
        ]
        return finite(sum(terms, Decimal(0))), scaled(terms)


def rate_residual(
    nper: Decimal,
    payment: Decimal,
    present_value: Decimal,
    future_value: Decimal,
    pay_type: Decimal,
    rate: Decimal,
) -> tuple[Decimal, Decimal]:
    """Evaluate the contract balance equation at a supplied candidate rate."""

    with localcontext(CONTEXT):
        if rate == -1:
            if nper != nper.to_integral_value() or nper <= 0:
                raise OracleError(NUM)
            growth = Decimal(0)
            annuity_factor = Decimal(1)
            due_factor = 1 - pay_type
        else:
            growth = real_power(1 + rate, nper)
            if rate == 0:
                annuity_factor = nper
            else:
                annuity_factor = (growth - 1) / rate
            due_factor = 1 + rate * pay_type
        terms = [
            present_value * growth,
            payment * annuity_factor * due_factor,
            future_value,
        ]
        return finite(sum(terms, Decimal(0))), scaled(terms)


def xirr_residual(
    values: list[Decimal], dates: list[Decimal], rate: Decimal
) -> tuple[Decimal, Decimal]:
    with localcontext(CONTEXT):
        if len(values) != len(dates):
            raise OracleError(VALUE)
        base = 1 + rate
        if base <= 0:
            raise OracleError(NUM)
        date0 = dates[0]
        terms = [
            value * real_power(base, -(date - date0) / Decimal(365))
            for value, date in zip(values, dates)
        ]
        return finite(sum(terms, Decimal(0))), scaled(terms)


def check_residual(function: str, row: dict[str, Any], root: str) -> None:
    rate = d(root)
    args = row["model_args"]
    if function == "IRR":
        residual, scale = irr_residual(
            [binary64_decimal(value) for value in args["values"]], rate
        )
    elif function == "RATE":
        residual, scale = rate_residual(
            *(binary64_decimal(args[name]) for name in (
                "nper",
                "payment",
                "present_value",
                "future_value",
                "pay_type",
            )),
            rate,
        )
    elif function == "XIRR":
        values = [binary64_decimal(value) for value in args["values"]]
        dates = [binary64_decimal(value) for value in args["dates"]]
        residual, scale = xirr_residual(values, dates, rate)
    else:
        raise AssertionError(function)
    bound = ROOT_ABS_TOLERANCE + ROOT_REL_TOLERANCE * max(Decimal(1), scale)
    if abs(residual) > bound:
        raise AssertionError(
            f"{row['id']} residual {residual} exceeds {bound} at root {root}"
        )


def number_expected(value: str) -> dict[str, Any]:
    return {
        "kind": "number",
        "value": value,
        "subtype": "Number",
        "tolerance": {
            "absolute": "1e-12",
            "relative": "1e-12",
        },
        "residual_tolerance": {
            "absolute": str(ROOT_ABS_TOLERANCE),
            "relative": str(ROOT_REL_TOLERANCE),
        },
    }


def error_expected(value: str) -> dict[str, Any]:
    return {"kind": "error", "value": value}


def bounded_expected(roots: list[str], reason: str) -> dict[str, Any]:
    return {
        "kind": "bounded-search",
        "status": "selection-or-success-not-guaranteed",
        "candidate_roots": roots,
        "reason": reason,
        "tolerance": {
            "absolute": "1e-12",
            "relative": "1e-12",
        },
        "residual_tolerance": {
            "absolute": str(ROOT_ABS_TOLERANCE),
            "relative": str(ROOT_REL_TOLERANCE),
        },
    }


def case(
    identifier: str,
    function: str,
    formula: str,
    args: list[str],
    model_args: dict[str, Any],
    expected: dict[str, Any],
    tags: list[str],
    basis: str,
) -> dict[str, Any]:
    return {
        "id": identifier,
        "function": function,
        "formula": formula,
        "args": args,
        "model_args": model_args,
        "expected": expected,
        "tags": tags,
        "oracle_basis": basis,
    }


def power_of_two_scale(exponent: int) -> Decimal:
    with localcontext(Context(prec=PRECISION + 40)):
        if exponent >= 0:
            return Decimal(2) ** exponent
        return Decimal(1) / (Decimal(2) ** (-exponent))


def scaled_cashflows(exponent: int) -> tuple[str, str]:
    scale = power_of_two_scale(exponent)
    return format(-10 * scale, "f"), format(11 * scale, "f")


def scaled_rate_args(exponent: int) -> tuple[str, str]:
    scale = power_of_two_scale(exponent)
    return format(100 * scale, "f"), format(-110 * scale, "f")


def root_cases() -> list[dict[str, Any]]:
    low_irr, high_irr = scaled_cashflows(-500), scaled_cashflows(500)
    low_rate, high_rate = scaled_rate_args(-500), scaled_rate_args(500)
    low_xirr, high_xirr = low_irr, high_irr

    rows = [
        case(
            "irr.exact_positive_rational",
            "IRR",
            "=IRR({-100|110})",
            ["-100", "110"],
            {"values": ["-100", "110"]},
            number_expected("0.1"),
            ["positive-branch", "exact-rational", "baseline"],
            "The periodic residual is zero at 1/10 for [-100, 110].",
        ),
        case(
            "irr.exact_negative_integral_branch",
            "IRR",
            "=IRR({-100|100|200};-2)",
            ["-100", "100", "200", "-2"],
            {"values": ["-100", "100", "200"]},
            number_expected("-2"),
            ["negative-branch", "exact-integer-root", "sign-variation"],
            "The negative-base residual is zero at rate -2; guess -2 selects that branch.",
        ),
        case(
            "irr.multiple_root_exact_low",
            "IRR",
            "=IRR({-100|230|-132};0.1)",
            ["-100", "230", "-132", "0.1"],
            {"values": ["-100", "230", "-132"]},
            number_expected("0.1"),
            ["multiple-root", "exact-rational", "guess-selection"],
            "The two exact roots are 1/10 and 1/5; an exactly-zero low-root guess returns 1/10.",
        ),
        case(
            "irr.multiple_root_exact_high",
            "IRR",
            "=IRR({-100|230|-132};0.2)",
            ["-100", "230", "-132", "0.2"],
            {"values": ["-100", "230", "-132"]},
            number_expected("0.2"),
            ["multiple-root", "exact-rational", "guess-selection"],
            "The same polynomial has an exact 1/5 root; an exactly-zero high-root guess returns 1/5.",
        ),
        case(
            "irr.multiple_root_mid_guess_bounded",
            "IRR",
            "=IRR({-100|230|-132};0.15)",
            ["-100", "230", "-132", "0.15"],
            {"values": ["-100", "230", "-132"]},
            bounded_expected(
                ["0.1", "0.2"],
                "ODF does not choose among multiple roots; the bounded profile may select the first bracket or reach its cap for this mid-interval guess.",
            ),
            ["multiple-root", "bounded-search", "non-asserting-selection"],
            "Both candidate roots have zero high-precision residual; no single outcome is claimed here.",
        ),
        case(
            "irr.tangent_root_no_sign_bracket",
            "IRR",
            "=IRR({100|-220|121};0.2)",
            ["100", "-220", "121", "0.2"],
            {"values": ["100", "-220", "121"]},
            error_expected(NUM),
            ["tangent-root", "no-sign-bracket"],
            "The residual has a repeated root at 1/10 and no sign change; the profile maps a missing bracket to #NUM!.",
        ),
        case(
            "irr.sign_variation_missing",
            "IRR",
            "=IRR({100|110})",
            ["100", "110"],
            {"values": ["100", "110"]},
            error_expected(NUM),
            ["invalid-sign", "domain-error"],
            "IRR requires both a positive and a negative admitted cash flow.",
        ),
        case(
            "irr.guess_exact_minus_one",
            "IRR",
            "=IRR({-100|110};-1)",
            ["-100", "110", "-1"],
            {"values": ["-100", "110"]},
            error_expected(NUM),
            ["invalid-guess", "branch-boundary"],
            "IRR guess exactly -1 has no valid transformed branch and maps to #NUM!.",
        ),
        case(
            "irr.scale_low_power_two",
            "IRR",
            f"=IRR({{{low_irr[0]}|{low_irr[1]}}})",
            list(low_irr),
            {"values": list(low_irr)},
            number_expected("0.1"),
            ["scale-extreme", "positive-branch", "exact-rational"],
            "Both cash flows are scaled by the exactly representable binary power 2^-500.",
        ),
        case(
            "irr.scale_high_power_two",
            "IRR",
            f"=IRR({{{high_irr[0]}|{high_irr[1]}}})",
            list(high_irr),
            {"values": list(high_irr)},
            number_expected("0.1"),
            ["scale-extreme", "positive-branch", "exact-rational"],
            "Both cash flows are scaled by the exactly representable binary power 2^500.",
        ),
        case(
            "rate.exact_positive_rational",
            "RATE",
            "=RATE(1;-110;100)",
            ["1", "-110", "100"],
            {
                "nper": "1",
                "payment": "-110",
                "present_value": "100",
                "future_value": "0",
                "pay_type": "0",
            },
            number_expected("0.1"),
            ["positive-branch", "exact-rational", "baseline"],
            "The balance equation is zero at 1/10 with Nper 1 and default Fv/PayType.",
        ),
        case(
            "rate.exact_negative_integral_branch",
            "RATE",
            "=RATE(1;0;100;100;0;-2)",
            ["1", "0", "100", "100", "0", "-2"],
            {
                "nper": "1",
                "payment": "0",
                "present_value": "100",
                "future_value": "100",
                "pay_type": "0",
            },
            number_expected("-2"),
            ["negative-branch", "exact-integer-root"],
            "The integral-period balance equation is zero at rate -2; guess -2 selects the eligible negative branch.",
        ),
        case(
            "rate.exact_minus_one_boundary",
            "RATE",
            "=RATE(2;-50;100;50;0;-1)",
            ["2", "-50", "100", "50", "0", "-1"],
            {
                "nper": "2",
                "payment": "-50",
                "present_value": "100",
                "future_value": "50",
                "pay_type": "0",
            },
            number_expected("-1"),
            ["rate-minus-one", "boundary", "exact"],
            "The documented guarded boundary residual is exactly zero at rate -1.",
        ),
        case(
            "rate.exact_minus_one_due_boundary",
            "RATE",
            "=RATE(1;100;100;0;1;-1)",
            ["1", "100", "100", "0", "1", "-1"],
            {
                "nper": "1",
                "payment": "100",
                "present_value": "100",
                "future_value": "0",
                "pay_type": "1",
            },
            number_expected("-1"),
            ["rate-minus-one", "boundary", "annuity-due"],
            "At rate -1 the due factor is 1 - PayType = 0, so this guarded residual is zero.",
        ),
        case(
            "rate.nonintegral_minus_one_num",
            "RATE",
            "=RATE(1.5;-50;100;50;0;-1)",
            ["1.5", "-50", "100", "50", "0", "-1"],
            {
                "nper": "1.5",
                "payment": "-50",
                "present_value": "100",
                "future_value": "50",
                "pay_type": "0",
            },
            error_expected(NUM),
            ["rate-minus-one", "nonintegral-nper", "domain-error"],
            "A rate -1 boundary is numeric only for positive integral Nper.",
        ),
        case(
            "rate.invalid_pay_type_num",
            "RATE",
            "=RATE(1;-110;100;0;0.5;0.1)",
            ["1", "-110", "100", "0", "0.5", "0.1"],
            {
                "nper": "1",
                "payment": "-110",
                "present_value": "100",
                "future_value": "0",
                "pay_type": "0.5",
            },
            error_expected(NUM),
            ["invalid-control", "domain-error"],
            "PayType is a Number control and only exact 0 or 1 is admitted.",
        ),
        case(
            "rate.nonfinite_guess_num",
            "RATE",
            "=RATE(1;-110;100;0;0;NaN)",
            ["1", "-110", "100", "0", "0", "NaN"],
            {
                "nper": "1",
                "payment": "-110",
                "present_value": "100",
                "future_value": "0",
                "pay_type": "0",
            },
            error_expected(NUM),
            ["invalid-guess", "nonfinite"],
            "A supplied non-finite guess is rejected before root search.",
        ),
        case(
            "rate.no_positive_branch_root",
            "RATE",
            "=RATE(1;100;100;100;0;0.1)",
            ["1", "100", "100", "100", "0", "0.1"],
            {
                "nper": "1",
                "payment": "100",
                "present_value": "100",
                "future_value": "100",
                "pay_type": "0",
            },
            error_expected(NUM),
            ["no-root-on-selected-branch", "bounded-search"],
            "The only algebraic root is below -1, while a positive guess selects the positive-base branch.",
        ),
        case(
            "rate.scale_low_power_two",
            "RATE",
            f"=RATE(1;{low_rate[1]};{low_rate[0]})",
            ["1", low_rate[1], low_rate[0]],
            {
                "nper": "1",
                "payment": low_rate[1],
                "present_value": low_rate[0],
                "future_value": "0",
                "pay_type": "0",
            },
            number_expected("0.1"),
            ["scale-extreme", "positive-branch", "exact-rational"],
            "PV and Payment are scaled by the exactly representable binary power 2^-500.",
        ),
        case(
            "rate.scale_high_power_two",
            "RATE",
            f"=RATE(1;{high_rate[1]};{high_rate[0]})",
            ["1", high_rate[1], high_rate[0]],
            {
                "nper": "1",
                "payment": high_rate[1],
                "present_value": high_rate[0],
                "future_value": "0",
                "pay_type": "0",
            },
            number_expected("0.1"),
            ["scale-extreme", "positive-branch", "exact-rational"],
            "PV and Payment are scaled by the exactly representable binary power 2^500.",
        ),
        case(
            "xirr.exact_positive_rational",
            "XIRR",
            "=XIRR({-100|110};{43831|44196})",
            ["-100", "110", "43831", "44196"],
            {"values": ["-100", "110"], "dates": ["43831", "44196"]},
            number_expected("0.1"),
            ["positive-branch", "exact-rational", "baseline"],
            "The date offset is exactly 365 days, so the residual is zero at 1/10.",
        ),
        case(
            "xirr.exact_two_year_rational",
            "XIRR",
            "=XIRR({-100|121};{43831|44561})",
            ["-100", "121", "43831", "44561"],
            {"values": ["-100", "121"], "dates": ["43831", "44561"]},
            number_expected("0.1"),
            ["positive-branch", "exact-rational", "date-weighted"],
            "The date offset is exactly 730 days and 121 = 100*(11/10)^2.",
        ),
        case(
            "xirr.fractional_date_zero_rate",
            "XIRR",
            "=XIRR({-100|100};{43831|44011})",
            ["-100", "100", "43831", "44011"],
            {"values": ["-100", "100"], "dates": ["43831", "44011"]},
            number_expected("0"),
            ["positive-branch", "fractional-date-exponent", "exact"],
            "At rate zero the base is one, so the 180/365 fractional exponent is exact.",
        ),
        case(
            "xirr.unsorted_dates_valid",
            "XIRR",
            "=XIRR({-110|100};{44196|43831})",
            ["-110", "100", "44196", "43831"],
            {"values": ["-110", "100"], "dates": ["44196", "43831"]},
            number_expected("0.1"),
            ["positive-branch", "unsorted-dates", "exact-rational"],
            "XIRR retains source order; it has no XNPV-style date-order constraint.",
        ),
        case(
            "xirr.scale_low_power_two",
            "XIRR",
            f"=XIRR({{{low_xirr[0]}|{low_xirr[1]}}};{{43831|44196}})",
            list(low_xirr) + ["43831", "44196"],
            {"values": list(low_xirr), "dates": ["43831", "44196"]},
            number_expected("0.1"),
            ["scale-extreme", "positive-branch", "exact-rational"],
            "Both cash flows are scaled by the exactly representable binary power 2^-500.",
        ),
        case(
            "xirr.scale_high_power_two",
            "XIRR",
            f"=XIRR({{{high_xirr[0]}|{high_xirr[1]}}};{{43831|44196}})",
            list(high_xirr) + ["43831", "44196"],
            {"values": list(high_xirr), "dates": ["43831", "44196"]},
            number_expected("0.1"),
            ["scale-extreme", "positive-branch", "exact-rational"],
            "Both cash flows are scaled by the exactly representable binary power 2^500.",
        ),
        case(
            "xirr.first_cashflow_positive_num",
            "XIRR",
            "=XIRR({100|-110};{43831|44196})",
            ["100", "-110", "43831", "44196"],
            {"values": ["100", "-110"], "dates": ["43831", "44196"]},
            error_expected(NUM),
            ["invalid-sign", "domain-error"],
            "The first admitted cash flow must be negative.",
        ),
        case(
            "xirr.date_below_profile_num",
            "XIRR",
            "=XIRR({-100|110};{43831|-693594})",
            ["-100", "110", "43831", "-693594"],
            {"values": ["-100", "110"], "dates": ["43831", "-693594"]},
            error_expected(NUM),
            ["invalid-date", "domain-error"],
            "The second date is below the date/time profile minimum serial.",
        ),
        case(
            "xirr.unequal_flattened_count_value",
            "XIRR",
            "=XIRR({-100|110|0};{43831|44196})",
            ["-100", "110", "0", "43831", "44196"],
            {"values": ["-100", "110", "0"], "dates": ["43831", "44196"]},
            error_expected(VALUE),
            ["invalid-shape", "value-error"],
            "XIRR equal-count failure is a shape #VALUE! result.",
        ),
        case(
            "xirr.guess_exact_minus_one_num",
            "XIRR",
            "=XIRR({-100|110};{43831|44196};-1)",
            ["-100", "110", "43831", "44196", "-1"],
            {"values": ["-100", "110"], "dates": ["43831", "44196"]},
            error_expected(NUM),
            ["invalid-guess", "branch-boundary"],
            "XIRR always uses the positive-base branch, so guess -1 is invalid.",
        ),
        case(
            "xirr.guess_below_minus_one_num",
            "XIRR",
            "=XIRR({-100|110};{43831|44196};-2)",
            ["-100", "110", "43831", "44196", "-2"],
            {"values": ["-100", "110"], "dates": ["43831", "44196"]},
            error_expected(NUM),
            ["invalid-guess", "unsupported-negative-branch"],
            "Fractional date exponents do not admit the negative-base branch.",
        ),
    ]
    assert len(rows) == EXPECTED_VECTOR_COUNT, len(rows)
    return rows


def validate_root_cases(rows: list[dict[str, Any]]) -> None:
    for row in rows:
        expected = row["expected"]
        kind = expected["kind"]
        if kind == "number":
            check_residual(row["function"], row, expected["value"])
        elif kind == "bounded-search":
            candidates = expected["candidate_roots"]
            if not candidates:
                raise AssertionError(f"{row['id']} has no bounded-search candidates")
            for candidate in candidates:
                check_residual(row["function"], row, candidate)
        elif kind == "error":
            if expected["value"] not in {NUM, VALUE}:
                raise AssertionError(f"{row['id']} has unknown formula error")
        else:
            raise AssertionError(f"{row['id']} has unknown expected kind {kind}")


def validate_error_semantics(rows: list[dict[str, Any]]) -> None:
    """Check that negative rows still exercise their stated contract failure."""

    by_id = {row["id"]: row for row in rows}
    values = by_id["irr.sign_variation_missing"]["model_args"]["values"]
    assert all(d(value) > 0 for value in values), "IRR sign row lost its invalid input"
    assert by_id["irr.guess_exact_minus_one"]["args"][-1] == "-1"
    tangent = by_id["irr.tangent_root_no_sign_bracket"]
    check_residual("IRR", tangent, "0.1")
    assert d(tangent["args"][-1]) != d("0.1")

    rate_nonintegral = by_id["rate.nonintegral_minus_one_num"]["model_args"]
    assert rate_nonintegral["nper"] == "1.5"
    assert by_id["rate.nonintegral_minus_one_num"]["args"][-1] == "-1"
    assert d(by_id["rate.invalid_pay_type_num"]["model_args"]["pay_type"]) not in {0, 1}
    assert by_id["rate.nonfinite_guess_num"]["args"][-1] == "NaN"

    xirr_sign = by_id["xirr.first_cashflow_positive_num"]["model_args"]["values"]
    assert d(xirr_sign[0]) > 0 and any(d(value) < 0 for value in xirr_sign[1:])
    xirr_date = by_id["xirr.date_below_profile_num"]["model_args"]["dates"]
    assert d(xirr_date[1]) < DATE_MIN
    xirr_shape = by_id["xirr.unequal_flattened_count_value"]["model_args"]
    assert len(xirr_shape["values"]) != len(xirr_shape["dates"])
    assert by_id["xirr.guess_exact_minus_one_num"]["args"][-1] == "-1"
    assert d(by_id["xirr.guess_below_minus_one_num"]["args"][-1]) < -1


def build_document() -> dict[str, Any]:
    rows = root_cases()
    hashes = source_hashes()
    document = {
        "schema": "ods-formula-financial-root-oracle-v1",
        "status": "draft independent high-precision root corpus; no Rust replay claimed",
        "contract_sha256": hashes["financial_contract_sha256"],
        "source_hashes": hashes,
        "normative_source": {
            "sections": ["6.12.24 IRR", "6.12.42 RATE", "6.12.51 XIRR"],
            "archive_sha256": hashes["normative_archive_sha256"],
            "part4_member_sha256": hashes["normative_part4_member_sha256"],
            "date_sequence_profile_sha256": hashes["date_time_contract_sha256"],
        },
        "coverage": {
            "functions": list(FUNCTIONS),
            "claim": (
                "Independent equation and residual corpus for IRR, RATE, and "
                "XIRR only; it does not certify Rust implementation, native "
                "compatibility, resource behavior, or performance."
            ),
        },
        "precision": {
            "decimal_digits": PRECISION,
            "rounding": "ROUND_HALF_EVEN",
            "input_model": (
                "finite formula Number operands are converted through Python "
                "binary64 float() and Decimal.from_float(); expected analytic "
                "roots remain exact Decimal rationals where constructed"
            ),
            "residual_validation": (
                "absolute and relative residual bound is 1e-120 times "
                "max(1, sum(abs(term)))"
            ),
            "output_tolerance": {
                "absolute": "1e-12",
                "relative": "1e-12",
            },
        },
        "profile": {
            "default_guess": "0.1",
            "positive_branch": "rate > -1 with u = ln1p(rate)",
            "negative_branch": (
                "IRR and positive-integral-Nper RATE only, when guess < -1; "
                "v = ln(-(1 + rate))"
            ),
            "xirr_branch": "positive-base only because date exponents can be fractional",
            "rate_minus_one": (
                "positive integral Nper uses growth 0, annuity factor 1, "
                "and due factor 1 - PayType"
            ),
            "multiple_roots": (
                "the selected bounded bracket is deterministic, but ODF does "
                "not promise a globally smallest root"
            ),
            "bounded_search_uncertainty": (
                "rows marked bounded-search validate candidate roots without "
                "claiming a unique finite-guess selection or success"
            ),
            "date_domain": "Date serials must satisfy -693593 <= date < 2958466",
        },
        "vector_count": len(rows),
        "function_counts": {function: EXPECTED_COUNTS[function] for function in FUNCTIONS},
        "vectors": rows,
    }
    validate_document(document)
    validate_root_cases(rows)
    validate_error_semantics(rows)
    return document


def validate_document(document: dict[str, Any]) -> None:
    if document.get("schema") != "ods-formula-financial-root-oracle-v1":
        raise AssertionError("unexpected root oracle schema")
    if document.get("contract_sha256") != digest(CONTRACT):
        raise AssertionError("stale financial contract hash")
    hashes = source_hashes()
    if document.get("source_hashes") != hashes:
        raise AssertionError("source hash set is stale or malformed")
    normative = document.get("normative_source")
    if not isinstance(normative, dict):
        raise AssertionError("normative_source must be an object")
    if normative.get("archive_sha256") != hashes["normative_archive_sha256"]:
        raise AssertionError("stale normative archive hash")
    if normative.get("part4_member_sha256") != hashes["normative_part4_member_sha256"]:
        raise AssertionError("stale normative Part 4 hash")
    if normative.get("date_sequence_profile_sha256") != hashes["date_time_contract_sha256"]:
        raise AssertionError("stale DateSequence profile hash")
    if document.get("coverage", {}).get("functions") != list(FUNCTIONS):
        raise AssertionError("root coverage function set/order changed")
    rows = document.get("vectors")
    if not isinstance(rows, list):
        raise AssertionError("vectors must be a list")
    if document.get("vector_count") != len(rows) or len(rows) != EXPECTED_VECTOR_COUNT:
        raise AssertionError("root vector cardinality changed")
    if document.get("function_counts") != EXPECTED_COUNTS:
        raise AssertionError("root function cardinality changed")
    ids = [row.get("id") for row in rows]
    if any(not isinstance(identifier, str) or not identifier for identifier in ids):
        raise AssertionError("root vector has missing ID")
    if len(set(ids)) != len(ids):
        raise AssertionError("duplicate root vector ID")
    counts = {function: 0 for function in FUNCTIONS}
    for row in rows:
        function = row.get("function")
        if function not in counts:
            raise AssertionError(f"unsupported root function {function}")
        counts[function] += 1
        if row.get("expected", {}).get("kind") == "bounded-search":
            expected = row["expected"]
            if expected.get("status") != "selection-or-success-not-guaranteed":
                raise AssertionError(f"{row['id']} weak bounded-search status")
            if not expected.get("candidate_roots"):
                raise AssertionError(f"{row['id']} missing bounded candidates")
    if counts != EXPECTED_COUNTS:
        raise AssertionError(f"root function counts changed: {counts}")
    if document.get("status") != "draft independent high-precision root corpus; no Rust replay claimed":
        raise AssertionError("root corpus status must remain non-accepting")


def write_document(document: dict[str, Any]) -> None:
    VECTORS.write_text(
        json.dumps(document, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def check_document() -> dict[str, Any]:
    expected = build_document()
    try:
        actual = json.loads(VECTORS.read_text(encoding="utf-8"))
    except FileNotFoundError as error:
        raise SystemExit(f"missing {VECTORS}; run --write first") from error
    except json.JSONDecodeError as error:
        raise SystemExit(f"invalid JSON in {VECTORS}: {error}") from error
    validate_document(actual)
    if actual != expected:
        raise SystemExit("root vector JSON differs from independently regenerated corpus")
    return {
        "verified": True,
        "schema": expected["schema"],
        "contract_sha256": expected["contract_sha256"],
        "vector_count": expected["vector_count"],
        "function_counts": expected["function_counts"],
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--write", action="store_true", help="write the deterministic corpus")
    modes.add_argument("--check", action="store_true", help="verify the corpus without writing")
    args = parser.parse_args()
    if args.write:
        write_document(build_document())
        print(json.dumps({"written": str(VECTORS), "vector_count": EXPECTED_VECTOR_COUNT}))
    else:
        print(json.dumps(check_document(), sort_keys=True))


if __name__ == "__main__":
    main()

