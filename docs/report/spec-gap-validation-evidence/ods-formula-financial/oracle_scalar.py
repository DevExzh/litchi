#!/usr/bin/env python3
"""Generate and verify the independent scalar financial oracle corpus.

This module is deliberately separate from the Rust evaluator.  It evaluates
the eleven scalar, non-iterative kernels with :class:`decimal.Decimal` at high
precision and writes the checked JSON corpus beside this file.  The remaining
nine contract entries are sequence/reducer or iterative functions and are
listed as pending; this file makes no claim about their implementation.

Run ``python3 oracle_scalar.py --write`` after a contract change, or
``python3 oracle_scalar.py --check`` to recompute every vector and compare it
with the committed corpus.  The current contract hash is recorded in the
corpus so a later contract freeze requires an explicit regeneration.
"""

from __future__ import annotations

import argparse
from decimal import (
    Context,
    Decimal,
    DivisionByZero,
    InvalidOperation,
    ROUND_HALF_EVEN,
    ROUND_DOWN,
    localcontext,
)
import hashlib
import json
import math
from pathlib import Path
from typing import Any, Callable, Iterable
import zipfile


HERE = Path(__file__).resolve().parent
CONTRACT = HERE / "contract.md"
VECTORS = HERE / "oracle-scalar-vectors.json"
NORMATIVE_ARCHIVE = HERE.parents[3] / "3rdparty/specs/OpenDocument-v1.4-os.zip"
NORMATIVE_ARCHIVE_SHA256 = "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4"
NORMATIVE_PART4_SHA256 = "ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1"

# The precision is intentionally well above the evaluator's binary64 result.
# Corpus consumers compare the Decimal target with a stated binary64 tolerance.
PRECISION = 100
CONTEXT = Context(prec=PRECISION, rounding=ROUND_HALF_EVEN, Emax=999999, Emin=-999999)
NORMATIVE_PART4_MEMBER = "part4-formula/OpenDocument-v1.4-os-part4-formula.html"

COVERED_FUNCTIONS = (
    "EFFECT",
    "FV",
    "IPMT",
    "ISPMT",
    "NOMINAL",
    "NPER",
    "PDURATION",
    "PMT",
    "PPMT",
    "PV",
    "RRI",
)
PENDING_FUNCTIONS = (
    "CUMIPMT",
    "CUMPRINC",
    "FVSCHEDULE",
    "IRR",
    "MIRR",
    "NPV",
    "RATE",
    "XIRR",
    "XNPV",
)


class OracleError(Exception):
    """Formula-level error emitted by the independent scalar model."""

    def __init__(self, value: str):
        super().__init__(value)
        self.value = value


NUM = "#NUM!"
DIV0 = "#DIV/0!"


def dec(value: str | int | Decimal) -> Decimal:
    """Model the evaluator's binary64 literal before high-precision arithmetic."""

    try:
        result = Decimal.from_float(float(str(value)))
    except (InvalidOperation, ValueError, OverflowError) as error:
        raise OracleError(NUM) from error
    if not result.is_finite():
        raise OracleError(NUM)
    return result


def finite(value: Decimal) -> Decimal:
    if not value.is_finite():
        raise OracleError(NUM)
    return value


def trunc_integer(value: Decimal) -> Decimal:
    """Repository Integer profile: truncate toward zero."""

    return value.to_integral_value(rounding=ROUND_DOWN)


def integer_slot(value: Decimal, *, positive: bool = False) -> Decimal:
    value = trunc_integer(value)
    if positive and value <= 0:
        raise OracleError(NUM)
    return value


def control(value: Decimal) -> Decimal:
    """Number-declared Type/PayType: only exact numeric 0 or 1 is admitted."""

    if value == 0 or value == 1:
        return value
    raise OracleError(NUM)


def decimal_power(base: Decimal, exponent: Decimal) -> Decimal:
    """Real-valued power with an explicit signed-integer branch.

    Decimal's ``ln``/``exp`` methods are used only for positive bases.  A
    negative base is admitted when the exponent is an exact integer, which is
    required by the finite integer-power FV and negative-ratio RRI profiles.
    """

    with localcontext(CONTEXT):
        if exponent == exponent.to_integral_value():
            integer_exponent = int(exponent)
            if base == 0 and integer_exponent < 0:
                raise OracleError(DIV0)
            try:
                return finite(base**integer_exponent)
            except (InvalidOperation, DivisionByZero, OverflowError) as error:
                raise OracleError(NUM) from error
        if base <= 0:
            raise OracleError(NUM)
        try:
            return finite((CONTEXT.ln(base) * exponent).exp(context=CONTEXT))
        except (InvalidOperation, DivisionByZero, OverflowError) as error:
            raise OracleError(NUM) from error


def growth(rate: Decimal, periods: Decimal) -> Decimal:
    return decimal_power(1 + rate, periods)


def annuity(rate: Decimal, periods: Decimal, pay_type: Decimal) -> Decimal:
    """Discount-free annuity factor, including the beginning-payment factor."""

    if rate == 0:
        return finite(periods)
    return finite((growth(rate, periods) - 1) * (1 + rate * pay_type) / rate)


def effect(rate: Decimal, payments: Decimal) -> Decimal:
    payments = integer_slot(payments, positive=True)
    if rate < 0:
        raise OracleError(NUM)
    return finite(decimal_power(1 + rate / payments, payments) - 1)


def nominal(effective_rate: Decimal, payments: Decimal) -> Decimal:
    payments = integer_slot(payments, positive=True)
    if effective_rate <= 0:
        raise OracleError(NUM)
    return finite(payments * (decimal_power(1 + effective_rate, 1 / payments) - 1))


def fv(
    rate: Decimal,
    periods: Decimal,
    payment: Decimal,
    present_value: Decimal = Decimal(0),
    pay_type: Decimal = Decimal(0),
) -> Decimal:
    pay_type = control(pay_type)
    factor = growth(rate, periods)
    return finite(-(present_value * factor + payment * annuity(rate, periods, pay_type)))


def pv(
    rate: Decimal,
    periods: Decimal,
    payment: Decimal,
    future_value: Decimal = Decimal(0),
    pay_type: Decimal = Decimal(0),
) -> Decimal:
    pay_type = control(pay_type)
    factor = growth(rate, periods)
    if factor == 0:
        raise OracleError(DIV0)
    return finite(-(future_value + payment * annuity(rate, periods, pay_type)) / factor)


def pmt(
    rate: Decimal,
    periods: Decimal,
    present_value: Decimal,
    future_value: Decimal = Decimal(0),
    pay_type: Decimal = Decimal(0),
) -> Decimal:
    periods = integer_slot(periods, positive=True)
    pay_type = control(pay_type)
    factor = annuity(rate, periods, pay_type)
    if factor == 0:
        raise OracleError(DIV0)
    return finite(-(present_value * growth(rate, periods) + future_value) / factor)


def nper(
    rate: Decimal,
    payment: Decimal,
    present_value: Decimal,
    future_value: Decimal = Decimal(0),
    pay_type: Decimal = Decimal(0),
) -> Decimal:
    pay_type = control(pay_type)
    if rate == 0:
        if payment == 0:
            raise OracleError(DIV0)
        return finite(-(present_value + future_value) / payment)
    payment_factor = (1 + rate * pay_type) * payment
    numerator = payment_factor - future_value * rate
    denominator = payment_factor + present_value * rate
    if denominator == 0:
        raise OracleError(DIV0)
    ratio = numerator / denominator
    if ratio <= 0:
        raise OracleError(NUM)
    base = 1 + rate
    if base <= 0:
        raise OracleError(NUM)
    try:
        return finite(CONTEXT.ln(ratio) / CONTEXT.ln(base))
    except (InvalidOperation, DivisionByZero, OverflowError) as error:
        raise OracleError(NUM) from error


def payment_unchecked(
    rate: Decimal,
    periods: Decimal,
    present_value: Decimal,
    future_value: Decimal,
    pay_type: Decimal,
) -> Decimal:
    if rate == 0:
        if periods == 0:
            raise OracleError(DIV0)
        return finite(-(present_value + future_value) / periods)
    factor = annuity(rate, periods, pay_type)
    if factor == 0:
        raise OracleError(DIV0)
    return finite(-(present_value * growth(rate, periods) + future_value) / factor)


def ipmt(
    rate: Decimal,
    period: Decimal,
    periods: Decimal,
    present_value: Decimal,
    future_value: Decimal = Decimal(0),
    pay_type: Decimal = Decimal(0),
) -> Decimal:
    pay_type = control(pay_type)
    if periods == 0:
        raise OracleError(DIV0)
    if rate == 0:
        return Decimal(0)
    # Repository profile: no interest is charged on the first beginning-
    # payment period; subsequent due-period interest is divided by (1+r).
    if pay_type == 1 and period == 1:
        return Decimal(0)
    payment = payment_unchecked(rate, periods, present_value, future_value, pay_type)
    prior_periods = period - 1
    prior_balance = finite(
        present_value * growth(rate, prior_periods)
        + payment * annuity(rate, prior_periods, pay_type)
    )
    due_divisor = 1 + rate if pay_type == 1 else Decimal(1)
    if due_divisor == 0:
        raise OracleError(DIV0)
    return finite(-prior_balance * rate / due_divisor)


def ispmt(
    rate: Decimal,
    period: Decimal,
    periods: Decimal,
    present_value: Decimal,
) -> Decimal:
    if periods == 0:
        raise OracleError(DIV0)
    # Conventional ODF profile selected for the equation-less §6.12.25 entry.
    return finite(rate * present_value * (period / periods - 1))


def ppmt(
    rate: Decimal,
    period: Decimal,
    periods: Decimal,
    present_value: Decimal,
    future_value: Decimal = Decimal(0),
    pay_type: Decimal = Decimal(0),
) -> Decimal:
    period = integer_slot(period)
    periods = integer_slot(periods)
    pay_type = control(pay_type)
    if rate <= 0 or present_value <= 0 or period <= 0 or period >= periods:
        raise OracleError(NUM)
    payment = payment_unchecked(rate, periods, present_value, future_value, pay_type)
    return finite(payment - ipmt(rate, period, periods, present_value, future_value, pay_type))


def pduration(rate: Decimal, current: Decimal, specified: Decimal) -> Decimal:
    if rate <= 0 or current <= 0 or specified <= 0:
        raise OracleError(NUM)
    if current == specified:
        return Decimal(0)
    try:
        return finite((CONTEXT.ln(specified) - CONTEXT.ln(current)) / CONTEXT.ln(1 + rate))
    except (InvalidOperation, DivisionByZero, OverflowError) as error:
        raise OracleError(NUM) from error


def rri(periods: Decimal, present: Decimal, future: Decimal) -> Decimal:
    if periods <= 0 or present == 0:
        raise OracleError(NUM)
    if present == future:
        return Decimal(0)
    if future == 0:
        return Decimal(-1)
    ratio = future / present
    if ratio > 0:
        try:
            return finite((CONTEXT.ln(ratio) / periods).exp(context=CONTEXT) - 1)
        except (InvalidOperation, DivisionByZero, OverflowError) as error:
            raise OracleError(NUM) from error
    # The accepted profile classifies this boundary in binary64, matching the
    # evaluator's reciprocal operation.  Decimal arithmetic remains the
    # independent high-precision model for the admitted magnitude.
    try:
        binary_period = float(periods)
        binary_exponent = 1.0 / binary_period
    except (OverflowError, ZeroDivisionError) as error:
        raise OracleError(NUM) from error
    if (
        not math.isfinite(binary_exponent)
        or binary_exponent <= 0.0
        or binary_exponent != math.trunc(binary_exponent)
    ):
        raise OracleError(NUM)
    exponent = int(binary_exponent)
    try:
        magnitude = (CONTEXT.ln(-ratio) / periods).exp(context=CONTEXT)
    except (InvalidOperation, DivisionByZero, OverflowError) as error:
        raise OracleError(NUM) from error
    if exponent % 2:
        magnitude = -magnitude
    return finite(magnitude - 1)


FUNCTIONS: dict[str, Callable[..., Decimal]] = {
    "EFFECT": effect,
    "FV": fv,
    "IPMT": ipmt,
    "ISPMT": ispmt,
    "NOMINAL": nominal,
    "NPER": nper,
    "PDURATION": pduration,
    "PMT": pmt,
    "PPMT": ppmt,
    "PV": pv,
    "RRI": rri,
}


def evaluate(function: str, args: Iterable[str]) -> Decimal | OracleError:
    try:
        with localcontext(CONTEXT):
            values = [dec(value) for value in args]
            result = FUNCTIONS[function](*values)
            return finite(result)
    except OracleError as error:
        return error
    except (InvalidOperation, DivisionByZero, OverflowError, TypeError) as error:
        raise AssertionError(f"unmapped Decimal failure for {function}: {error}") from error


def canonical(value: Decimal) -> str:
    """Serialize a high-precision Decimal without losing significant digits."""

    with localcontext(CONTEXT):
        value = +value
    if value == 0:
        return "0"
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def vector(
    ident: str,
    function: str,
    formula: str,
    args: list[str],
    tags: list[str],
    basis: str,
) -> dict[str, Any]:
    outcome = evaluate(function, args)
    if isinstance(outcome, OracleError):
        expected: dict[str, Any] = {"kind": "error", "value": outcome.value}
    else:
        expected = {
            "kind": "number",
            "value": canonical(outcome),
            "subtype": "Number",
            "tolerance": {"absolute": "0", "relative": "1e-12"},
        }
    return {
        "id": ident,
        "function": function,
        "formula": formula,
        "args": args,
        "expected": expected,
        "tags": tags,
        "oracle_basis": basis,
    }


def vector_specs() -> list[dict[str, Any]]:
    """Hand-authored inputs; all expected outcomes are generated above."""

    rows: list[tuple[str, str, str, list[str], list[str], str]] = [
        (
            "effect.basic",
            "EFFECT",
            "=EFFECT(0.1;2)",
            ["0.1", "2"],
            ["baseline", "closed-form"],
            "(1 + Rate / Payments)^Payments - 1",
        ),
        (
            "effect.zero_rate",
            "EFFECT",
            "=EFFECT(0;12)",
            ["0", "12"],
            ["zero-rate", "exact"],
            "explicit zero-rate compounding branch",
        ),
        (
            "effect.integer_truncates",
            "EFFECT",
            "=EFFECT(0.1;2.9)",
            ["0.1", "2.9"],
            ["integer-profile", "truncation"],
            "declared Integer Payments truncates toward zero before validation",
        ),
        (
            "effect.negative_rate_num",
            "EFFECT",
            "=EFFECT(-0.1;2)",
            ["-0.1", "2"],
            ["domain-error"],
            "source Rate >= 0 constraint",
        ),
        (
            "effect.zero_payments_num",
            "EFFECT",
            "=EFFECT(0.1;0)",
            ["0.1", "0"],
            ["domain-error"],
            "source Payments > 0 constraint",
        ),
        (
            "fv.basic",
            "FV",
            "=FV(0.1;1;-110)",
            ["0.1", "1", "-110"],
            ["baseline", "default-arguments"],
            "balance equation with Pv=0 and PayType=0 defaults",
        ),
        (
            "fv.zero_rate",
            "FV",
            "=FV(0;10;-10)",
            ["0", "10", "-10"],
            ["zero-rate", "exact"],
            "explicit zero-rate annuity branch",
        ),
        (
            "fv.beginning_payment",
            "FV",
            "=FV(0.1;1;-100;0;1)",
            ["0.1", "1", "-100", "0", "1"],
            ["pay-type", "annuity-due"],
            "balance equation with due factor 1 + Rate * PayType",
        ),
        (
            "fv.negative_base_integer_power",
            "FV",
            "=FV(-2;2;0;100;0)",
            ["-2", "2", "0", "100", "0"],
            ["regression", "negative-base", "integer-power"],
            "finite signed integer power (1 + Rate)^Nper",
        ),
        (
            "fv.fractional_nper",
            "FV",
            "=FV(0.1;0.5;-10)",
            ["0.1", "0.5", "-10"],
            ["fractional-number"],
            "Number Nper remains fractional; positive-base real power",
        ),
        (
            "fv.invalid_pay_type_num",
            "FV",
            "=FV(0.1;2;0;0;0.5)",
            ["0.1", "2", "0", "0", "0.5"],
            ["control-error"],
            "Number PayType admits only exact 0 or 1",
        ),
        (
            "ipmt.ordinary_period_one",
            "IPMT",
            "=IPMT(0.1;1;2;100)",
            ["0.1", "1", "2", "100"],
            ["baseline", "ordinary-annuity"],
            "selected amortization profile, Type=0 default",
        ),
        (
            "ipmt.ordinary_period_two",
            "IPMT",
            "=IPMT(0.1;2;2;100)",
            ["0.1", "2", "2", "100"],
            ["ordinary-annuity", "period-boundary"],
            "selected amortization profile after the first payment",
        ),
        (
            "ipmt.due_first_period_zero",
            "IPMT",
            "=IPMT(0.1;1;2;100;0;1)",
            ["0.1", "1", "2", "100", "0", "1"],
            ["pay-type", "annuity-due", "exact"],
            "beginning payment occurs before first-period interest",
        ),
        (
            "ipmt.due_after_period_one",
            "IPMT",
            "=IPMT(0.1;2;2;100;0;1)",
            ["0.1", "2", "2", "100", "0", "1"],
            ["regression", "pay-type", "annuity-due"],
            "due-period interest uses the selected post-payment divisor 1 + Rate",
        ),
        (
            "ipmt.zero_rate",
            "IPMT",
            "=IPMT(0;1;10;100)",
            ["0", "1", "10", "100"],
            ["zero-rate", "exact"],
            "interest component is zero at Rate=0",
        ),
        (
            "ipmt.zero_nper_divzero",
            "IPMT",
            "=IPMT(0.1;1;0;100)",
            ["0.1", "1", "0", "100"],
            ["domain-error"],
            "selected zero-period divisor mapping",
        ),
        (
            "ipmt.invalid_type_num",
            "IPMT",
            "=IPMT(0.1;1;2;100;0;0.5)",
            ["0.1", "1", "2", "100", "0", "0.5"],
            ["control-error"],
            "Number Type admits only exact 0 or 1",
        ),
        (
            "ispmt.period_one",
            "ISPMT",
            "=ISPMT(0.1;1;2;100)",
            ["0.1", "1", "2", "100"],
            ["baseline", "profile-regression"],
            "conventional profile Rate * Pv * (Period / Nper - 1)",
        ),
        (
            "ispmt.period_two",
            "ISPMT",
            "=ISPMT(0.1;2;2;100)",
            ["0.1", "2", "2", "100"],
            ["profile-regression", "exact"],
            "conventional profile at Period=Nper",
        ),
        (
            "ispmt.period_zero",
            "ISPMT",
            "=ISPMT(0.1;0;2;100)",
            ["0.1", "0", "2", "100"],
            ["profile-boundary"],
            "conventional profile permits Number Period=0",
        ),
        (
            "ispmt.zero_rate",
            "ISPMT",
            "=ISPMT(0;1;2;100)",
            ["0", "1", "2", "100"],
            ["zero-rate", "exact"],
            "conventional profile linear zero-rate result",
        ),
        (
            "ispmt.zero_nper_divzero",
            "ISPMT",
            "=ISPMT(0.1;1;0;100)",
            ["0.1", "1", "0", "100"],
            ["domain-error"],
            "zero Nper denominator",
        ),
        (
            "nominal.basic",
            "NOMINAL",
            "=NOMINAL(0.1025;2)",
            ["0.1025", "2"],
            ["baseline", "closed-form"],
            "inverse of EFFECT",
        ),
        (
            "nominal.integer_truncates",
            "NOMINAL",
            "=NOMINAL(0.1025;2.9)",
            ["0.1025", "2.9"],
            ["integer-profile", "truncation"],
            "declared Integer CompoundingPeriods truncates toward zero",
        ),
        (
            "nominal.one_period",
            "NOMINAL",
            "=NOMINAL(0.1;1)",
            ["0.1", "1"],
            ["exact", "closed-form"],
            "one-period inverse compounding",
        ),
        (
            "nominal_zero_rate_num",
            "NOMINAL",
            "=NOMINAL(0;2)",
            ["0", "2"],
            ["domain-error"],
            "source EffectiveRate > 0 constraint",
        ),
        (
            "nper.basic",
            "NPER",
            "=NPER(0.1;-110;100)",
            ["0.1", "-110", "100"],
            ["baseline", "closed-form"],
            "logarithmic balance-equation solution",
        ),
        (
            "nper.zero_rate",
            "NPER",
            "=NPER(0;-10;100)",
            ["0", "-10", "100"],
            ["zero-rate", "exact"],
            "explicit zero-rate balance branch",
        ),
        (
            "nper.tiny_rate",
            "NPER",
            "=NPER(1e-20;-10;100;0;0)",
            ["1e-20", "-10", "100", "0", "0"],
            ["regression", "tiny-rate", "stability"],
            "unrounded logarithmic balance ratio at a tiny nonzero Rate",
        ),
        (
            "nper.negative_rate",
            "NPER",
            "=NPER(-0.1;-10;100)",
            ["-0.1", "-10", "100"],
            ["negative-rate-profile"],
            "source/profile admits a finite negative Rate when the real log branch exists",
        ),
        (
            "nper.beginning_payment",
            "NPER",
            "=NPER(0.1;-110;100;0;1)",
            ["0.1", "-110", "100", "0", "1"],
            ["pay-type", "annuity-due"],
            "balance equation with due factor",
        ),
        (
            "nper.zero_denominator_divzero",
            "NPER",
            "=NPER(0.1;-10;100)",
            ["0.1", "-10", "100"],
            ["domain-error"],
            "balance-equation denominator is zero",
        ),
        (
            "nper.invalid_pay_type_num",
            "NPER",
            "=NPER(0.1;-110;100;0;1.9)",
            ["0.1", "-110", "100", "0", "1.9"],
            ["control-error"],
            "Number PayType is not truncated; only exact 0 or 1 is admitted",
        ),
        (
            "pduration.basic",
            "PDURATION",
            "=PDURATION(0.1;100;121)",
            ["0.1", "100", "121"],
            ["baseline", "closed-form"],
            "(ln(Specified) - ln(Current)) / ln(1 + Rate)",
        ),
        (
            "pduration.equal_values",
            "PDURATION",
            "=PDURATION(0.1;100;100)",
            ["0.1", "100", "100"],
            ["exact", "boundary"],
            "equal values return zero before logarithm",
        ),
        (
            "pduration.decrease",
            "PDURATION",
            "=PDURATION(0.1;121;100)",
            ["0.1", "121", "100"],
            ["negative-result"],
            "signed logarithmic duration",
        ),
        (
            "pduration.invalid_rate_num",
            "PDURATION",
            "=PDURATION(0;100;121)",
            ["0", "100", "121"],
            ["domain-error"],
            "source Rate > 0 constraint",
        ),
        (
            "pduration.invalid_current_num",
            "PDURATION",
            "=PDURATION(0.1;0;121)",
            ["0.1", "0", "121"],
            ["domain-error"],
            "source CurrentValue > 0 constraint",
        ),
        (
            "pmt.basic",
            "PMT",
            "=PMT(0.1;1;100)",
            ["0.1", "1", "100"],
            ["baseline", "default-arguments"],
            "balance equation with Fv=0 and PayType=0 defaults",
        ),
        (
            "pmt.zero_rate",
            "PMT",
            "=PMT(0;10;100)",
            ["0", "10", "100"],
            ["zero-rate", "exact"],
            "explicit zero-rate payment branch",
        ),
        (
            "pmt.beginning_payment",
            "PMT",
            "=PMT(0.1;1;100;0;1)",
            ["0.1", "1", "100", "0", "1"],
            ["pay-type", "annuity-due"],
            "balance equation with due factor",
        ),
        (
            "pmt.integer_nper_truncates",
            "PMT",
            "=PMT(0.1;2.9;100)",
            ["0.1", "2.9", "100"],
            ["integer-profile", "truncation"],
            "declared Integer Nper truncates toward zero",
        ),
        (
            "pmt.invalid_nper_num",
            "PMT",
            "=PMT(0.1;0;100)",
            ["0.1", "0", "100"],
            ["domain-error"],
            "source Nper > 0 constraint",
        ),
        (
            "pmt.invalid_pay_type_num",
            "PMT",
            "=PMT(0.1;1;100;0;0.5)",
            ["0.1", "1", "100", "0", "0.5"],
            ["control-error"],
            "Number PayType admits only exact 0 or 1",
        ),
        (
            "ppmt.basic",
            "PPMT",
            "=PPMT(0.1;1;2;100)",
            ["0.1", "1", "2", "100"],
            ["baseline", "closed-form"],
            "Payment - IPMT under the selected amortization profile",
        ),
        (
            "ppmt.beginning_payment",
            "PPMT",
            "=PPMT(0.1;1;2;100;0;1)",
            ["0.1", "1", "2", "100", "0", "1"],
            ["pay-type", "annuity-due"],
            "Type=1 first-period interest is zero",
        ),
        (
            "ppmt.period_equal_nper_num",
            "PPMT",
            "=PPMT(0.1;2;2;100)",
            ["0.1", "2", "2", "100"],
            ["regression", "domain-error", "period-boundary"],
            "explicit PPMT precondition Period < Nper",
        ),
        (
            "ppmt.period_integer_truncates",
            "PPMT",
            "=PPMT(0.1;1.9;2.9;100)",
            ["0.1", "1.9", "2.9", "100"],
            ["integer-profile", "truncation"],
            "declared Integer Period and Nper truncate toward zero",
        ),
        (
            "ppmt.zero_rate_num",
            "PPMT",
            "=PPMT(0;1;2;100)",
            ["0", "1", "2", "100"],
            ["domain-error"],
            "source Rate > 0 constraint",
        ),
        (
            "ppmt.invalid_type_num",
            "PPMT",
            "=PPMT(0.1;1;2;100;0;0.5)",
            ["0.1", "1", "2", "100", "0", "0.5"],
            ["control-error"],
            "Number Type is not truncated; exact 0/1 only",
        ),
        (
            "pv.basic",
            "PV",
            "=PV(0.1;1;-110)",
            ["0.1", "1", "-110"],
            ["baseline", "default-arguments"],
            "balance equation with Fv=0 and PayType=0 defaults",
        ),
        (
            "pv.zero_rate",
            "PV",
            "=PV(0;10;-10)",
            ["0", "10", "-10"],
            ["zero-rate", "exact"],
            "explicit zero-rate present-value branch",
        ),
        (
            "pv.beginning_payment",
            "PV",
            "=PV(0.1;1;-100;0;1)",
            ["0.1", "1", "-100", "0", "1"],
            ["pay-type", "annuity-due"],
            "balance equation with due factor",
        ),
        (
            "pv.rate_minus_one_divzero",
            "PV",
            "=PV(-1;1;0;1)",
            ["-1", "1", "0", "1"],
            ["regression", "domain-error"],
            "zero growth factor is a direct PV divisor",
        ),
        (
            "pv.invalid_pay_type_num",
            "PV",
            "=PV(0.1;1;-100;0;1.9)",
            ["0.1", "1", "-100", "0", "1.9"],
            ["control-error"],
            "Number PayType is not truncated; exact 0/1 only",
        ),
        (
            "rri.basic",
            "RRI",
            "=RRI(2;100;121)",
            ["2", "100", "121"],
            ["baseline", "closed-form"],
            "(Fv / Pv)^(1 / Nper) - 1",
        ),
        (
            "rri.one_period",
            "RRI",
            "=RRI(1;100;110)",
            ["1", "100", "110"],
            ["exact", "closed-form"],
            "one-period ratio",
        ),
        (
            "rri.negative_ratio_integer_exponent",
            "RRI",
            "=RRI(1;100;-100)",
            ["1", "100", "-100"],
            ["regression", "negative-base", "integer-power"],
            "negative ratio with exact integer reciprocal exponent",
        ),
        (
            "rri.negative_ratio_even_power",
            "RRI",
            "=RRI(0.5;100;-100)",
            ["0.5", "100", "-100"],
            ["negative-base", "integer-power"],
            "reciprocal exponent 2 gives a positive real root",
        ),
        (
            "rri.negative_ratio_odd_root_refused",
            "RRI",
            "=RRI(3;100;-800)",
            ["3", "100", "-800"],
            ["negative-base", "domain-error", "profile-regression"],
            "the selected profile admits a negative ratio only when the binary64 reciprocal Nper is an exact integer",
        ),
        (
            "rri.zero_future_minus_one",
            "RRI",
            "=RRI(2;100;0)",
            ["2", "100", "0"],
            ["boundary", "exact"],
            "zero future value gives ratio limit -1",
        ),
        (
            "rri.negative_ratio_fractional_num",
            "RRI",
            "=RRI(2;100;-100)",
            ["2", "100", "-100"],
            ["domain-error", "negative-base"],
            "negative ratio with noninteger reciprocal exponent has no real result",
        ),
        (
            "rri.zero_nper_num",
            "RRI",
            "=RRI(0;100;121)",
            ["0", "100", "121"],
            ["domain-error"],
            "source Nper > 0 constraint",
        ),
        (
            "rri.zero_present_num",
            "RRI",
            "=RRI(2;0;121)",
            ["2", "0", "121"],
            ["domain-error"],
            "zero present value has no finite ratio",
        ),
    ]
    return [vector(*row) for row in rows]


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def validate_normative_source() -> None:
    """Fail closed if the repository-local ODF source identity changed."""

    if not NORMATIVE_ARCHIVE.is_file():
        raise SystemExit(f"missing normative archive: {NORMATIVE_ARCHIVE}")
    observed_archive = sha256(NORMATIVE_ARCHIVE)
    if observed_archive != NORMATIVE_ARCHIVE_SHA256:
        raise SystemExit(
            f"normative archive hash changed: {observed_archive} != {NORMATIVE_ARCHIVE_SHA256}"
        )
    try:
        with zipfile.ZipFile(NORMATIVE_ARCHIVE) as archive:
            member = archive.read(NORMATIVE_PART4_MEMBER)
    except (OSError, KeyError, zipfile.BadZipFile) as error:
        raise SystemExit("normative Part 4 member is unreadable") from error
    observed_member = sha256_bytes(member)
    if observed_member != NORMATIVE_PART4_SHA256:
        raise SystemExit(
            f"normative Part 4 hash changed: {observed_member} != {NORMATIVE_PART4_SHA256}"
        )


def document() -> dict[str, Any]:
    validate_normative_source()
    rows = vector_specs()
    return {
        "schema": "ods-formula-financial-scalar-oracle-v1",
        "status": "draft independent Decimal corpus; scalar slice only",
        "contract_sha256": sha256(CONTRACT),
        "normative_source": {
            "archive_sha256": NORMATIVE_ARCHIVE_SHA256,
            "part4_member_sha256": NORMATIVE_PART4_SHA256,
            "sections": "6.12.19-6.12.20, 6.12.23, 6.12.25, 6.12.28-6.12.29, 6.12.35-6.12.37, 6.12.41, 6.12.44 plus repository profiles",
        },
        "coverage": {
            "covered_functions": list(COVERED_FUNCTIONS),
            "pending_functions": list(PENDING_FUNCTIONS),
            "claim": "This corpus covers eleven scalar non-iterative kernels only; it makes no all-twenty implementation or validation claim.",
        },
        "precision": {
            "decimal_digits": PRECISION,
            "rounding": "ROUND_HALF_EVEN during independent arithmetic",
            "input_model": "formula literals are converted with IEEE-754 binary64 float() and Decimal.from_float before Decimal arithmetic",
            "comparison": "per-vector absolute tolerance is zero; finite evaluator outputs use a 1e-12 relative tolerance without a unit-scale floor",
        },
        "profile": {
            "integer_conversion": "declared Integer arguments truncate toward zero",
            "number_controls": "Type and PayType declared Number arguments require exact converted 0 or 1",
            "signed_integer_power": "finite negative-base integer powers are admitted where the equation has an integral exponent",
            "rri_negative_ratio": "negative ratios are admitted only when the binary64 reciprocal of Nper is an exact positive integer; otherwise the result is #NUM!",
            "ispmt": "Rate * Pv * (Period / Nper - 1)",
            "ipmt_due": "Type=1 period one has zero interest; later due-period interest uses the selected 1 + Rate divisor",
            "ppmt_boundary": "Period >= Nper is #NUM!",
        },
        "open_profile_questions": [],
        "vector_count": len(rows),
        "vectors": rows,
    }


def check_document(actual: dict[str, Any], expected: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    for key in (
        "schema",
        "contract_sha256",
        "normative_source",
        "coverage",
        "precision",
        "profile",
        "open_profile_questions",
    ):
        if actual.get(key) != expected.get(key):
            errors.append(f"document field differs: {key}")
    actual_raw = actual.get("vectors")
    expected_raw = expected.get("vectors")
    if not isinstance(actual_raw, list) or not isinstance(expected_raw, list):
        return errors + ["vectors must be arrays"]
    actual_ids = [row.get("id") if isinstance(row, dict) else None for row in actual_raw]
    expected_ids = [row.get("id") if isinstance(row, dict) else None for row in expected_raw]
    if len(actual_ids) != len(set(actual_ids)):
        errors.append("committed corpus contains duplicate vector identifiers")
    if len(expected_ids) != len(set(expected_ids)):
        errors.append("generated corpus contains duplicate vector identifiers")
    if len(actual_raw) != len(expected_raw):
        errors.append("vector cardinality differs")
    actual_rows = {row.get("id"): row for row in actual_raw if isinstance(row, dict)}
    expected_rows = {row.get("id"): row for row in expected_raw if isinstance(row, dict)}
    if set(actual_rows) != set(expected_rows):
        errors.append("vector identifiers differ")
    for ident, row in expected_rows.items():
        if actual_rows.get(ident) != row:
            errors.append(f"vector differs: {ident}")
    if actual.get("vector_count") != len(expected_raw):
        errors.append("vector_count differs")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true", help="write the generated JSON corpus")
    parser.add_argument("--check", action="store_true", help="compare the committed corpus with regenerated output")
    args = parser.parse_args()
    if args.write and args.check:
        parser.error("choose --write or --check")
    generated = document()
    if args.write:
        VECTORS.write_text(json.dumps(generated, indent=2) + "\n", encoding="utf-8")
        return 0
    if args.check or not VECTORS.is_file():
        if not VECTORS.is_file():
            print(f"missing corpus: {VECTORS}")
            return 1
        actual = json.loads(VECTORS.read_text(encoding="utf-8"))
        errors = check_document(actual, generated)
        if errors:
            for error in errors:
                print(error)
            return 1
        print(f"verified {generated['vector_count']} financial scalar vectors")
        return 0
    print(f"generated {generated['vector_count']} vectors; use --write to update {VECTORS}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
