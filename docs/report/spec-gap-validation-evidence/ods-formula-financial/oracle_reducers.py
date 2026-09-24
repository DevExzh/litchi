#!/usr/bin/env python3
"""Generate the independent Decimal oracle for the six reducer functions.

The model is intentionally separate from the Rust evaluator and from the
scalar oracle.  It covers ``CUMIPMT``, ``CUMPRINC``, ``FVSCHEDULE``, ``MIRR``,
``NPV``, and ``XNPV`` with 100-digit Decimal arithmetic after converting
formula literals to binary64 inputs.  ``IRR``, ``RATE``, and ``XIRR`` remain
pending because their bounded root-selection profile is a separate contract
surface.

Use ``--write`` to regenerate the adjacent JSON corpus after a contract
change, or ``--check`` to verify the committed corpus and source bindings.
"""

from __future__ import annotations

import argparse
from decimal import (
    Context,
    Decimal,
    DivisionByZero,
    InvalidOperation,
    ROUND_DOWN,
    ROUND_HALF_EVEN,
    localcontext,
)
import hashlib
import json
import math
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator
import zipfile
from fractions import Fraction


HERE = Path(__file__).resolve().parent
CONTRACT = HERE / "contract.md"
VECTORS = HERE / "oracle-reducer-vectors.json"
NORMATIVE_ARCHIVE = HERE.parents[3] / "3rdparty/specs/OpenDocument-v1.4-os.zip"
NORMATIVE_ARCHIVE_SHA256 = "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4"
NORMATIVE_PART4_MEMBER = "part4-formula/OpenDocument-v1.4-os-part4-formula.html"
NORMATIVE_PART4_SHA256 = "ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1"

PRECISION = 100
CONTEXT = Context(prec=PRECISION, rounding=ROUND_HALF_EVEN, Emax=999999, Emin=-999999)

COVERED_FUNCTIONS = (
    "CUMIPMT",
    "CUMPRINC",
    "FVSCHEDULE",
    "MIRR",
    "NPV",
    "XNPV",
)
PENDING_FUNCTIONS = ("IRR", "RATE", "XIRR")

NUM = "#NUM!"
VALUE = "#VALUE!"
DIV0 = "#DIV/0!"
F64_EPSILON = Decimal(2) ** Decimal(-52)
FORWARD_ERROR_MULTIPLIER = Decimal(16)


class OracleError(Exception):
    """A generated or retained formula-level error."""

    def __init__(self, value: str, *, formula: bool = False):
        super().__init__(value)
        self.value = value
        self.formula = formula


def decimal_literal(value: str | int | Decimal) -> Decimal:
    """Convert a formula literal as the evaluator does before Decimal math."""

    try:
        result = Decimal.from_float(float(str(value)))
    except (InvalidOperation, OverflowError, ValueError) as error:
        raise OracleError(NUM) from error
    if not result.is_finite():
        raise OracleError(NUM)
    return result


def finite(value: Decimal) -> Decimal:
    if not value.is_finite():
        raise OracleError(NUM)
    return value


def trunc_integer(value: Decimal) -> Decimal:
    return value.to_integral_value(rounding=ROUND_DOWN)


def integer_slot(value: Decimal, *, positive: bool = False) -> Decimal:
    value = trunc_integer(value)
    if positive and value <= 0:
        raise OracleError(NUM)
    return value


def integer_control(value: Decimal) -> Decimal:
    value = integer_slot(value)
    if value == 0 or value == 1:
        return value
    raise OracleError(NUM)


def number_power(base: Decimal, exponent: Decimal) -> Decimal:
    """Evaluate a real power, preserving finite signed integer powers."""

    with localcontext(CONTEXT):
        if base == 0:
            if exponent == 0:
                return Decimal(1)
            if exponent < 0:
                raise OracleError(DIV0)
            return Decimal(0)
        if exponent == exponent.to_integral_value():
            integer_exponent = int(exponent)
            try:
                return finite(base**integer_exponent)
            except (DivisionByZero, InvalidOperation, OverflowError) as error:
                raise OracleError(NUM) from error
        if base <= 0:
            raise OracleError(NUM)
        try:
            return finite((CONTEXT.ln(base) * exponent).exp(context=CONTEXT))
        except (DivisionByZero, InvalidOperation, OverflowError) as error:
            raise OracleError(NUM) from error


def growth(rate: Decimal, periods: Decimal) -> Decimal:
    return number_power(1 + rate, periods)


def annuity_factor(rate: Decimal, periods: Decimal, pay_type: Decimal) -> Decimal:
    if rate == 0:
        return finite(periods)
    return finite((growth(rate, periods) - 1) * (1 + rate * pay_type) / rate)


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
    factor = annuity_factor(rate, periods, pay_type)
    if factor == 0:
        raise OracleError(DIV0)
    return finite(-(present_value * growth(rate, periods) + future_value) / factor)


def ipmt_profile(
    rate: Decimal,
    period: Decimal,
    periods: Decimal,
    present_value: Decimal,
    future_value: Decimal,
    pay_type: Decimal,
) -> Decimal:
    """The selected repository IPMT profile used by both cumulative reducers."""

    if periods == 0:
        raise OracleError(DIV0)
    if rate == 0:
        return Decimal(0)
    if pay_type == 1 and period == 1:
        return Decimal(0)
    payment = payment_unchecked(rate, periods, present_value, future_value, pay_type)
    prior_periods = period - 1
    balance = finite(
        present_value * growth(rate, prior_periods)
        + payment * annuity_factor(rate, prior_periods, pay_type)
    )
    divisor = 1 + rate if pay_type == 1 else Decimal(1)
    if divisor == 0:
        raise OracleError(DIV0)
    return finite(-balance * rate / divisor)


def ppmt_profile(
    rate: Decimal,
    period: Decimal,
    periods: Decimal,
    present_value: Decimal,
    future_value: Decimal,
    pay_type: Decimal,
) -> Decimal:
    if rate <= 0 or present_value <= 0 or period <= 0 or period >= periods:
        raise OracleError(NUM)
    payment = payment_unchecked(rate, periods, present_value, future_value, pay_type)
    return finite(
        payment - ipmt_profile(rate, period, periods, present_value, future_value, pay_type)
    )


def iter_sequence(sequence: Any) -> Iterator[Any]:
    """Yield a compact sequence descriptor in source order.

    Lists represent ordinary row-major elements.  A mapping may use
    ``elements`` plus optional shape metadata, or ``repeat``/``tail`` for a
    compact long prefix such as the NPV underflow-rescue vector.
    """

    if isinstance(sequence, list):
        yield from sequence
        return
    if not isinstance(sequence, dict):
        raise OracleError(VALUE)
    if "elements" in sequence:
        elements = sequence["elements"]
        if not isinstance(elements, list):
            raise OracleError(VALUE)
        yield from elements
        return
    repeat = sequence.get("repeat")
    if repeat is not None:
        if not isinstance(repeat, dict):
            raise OracleError(VALUE)
        count = repeat.get("count")
        if not isinstance(count, int) or count < 0:
            raise OracleError(VALUE)
        for _ in range(count):
            yield repeat.get("value")
    tail = sequence.get("tail", [])
    if not isinstance(tail, list):
        raise OracleError(VALUE)
    yield from tail


def kind_value(item: Any) -> tuple[str, Any]:
    if isinstance(item, dict):
        kind = item.get("kind")
        if kind == "number":
            return "number", item.get("value")
        if kind == "logical":
            return "logical", bool(item.get("value"))
        if kind == "text":
            return "text", item.get("value", "")
        if kind == "empty":
            return "empty", None
        if kind == "error":
            value = item.get("value")
            if not isinstance(value, str) or not value.startswith("#"):
                raise OracleError(VALUE)
            return "error", value
        raise OracleError(VALUE)
    if isinstance(item, (str, int, float)):
        return "number", item
    if item is None:
        return "empty", None
    raise OracleError(VALUE)


def first_error(current: OracleError | None, candidate: OracleError) -> OracleError:
    return current if current is not None else candidate


def cumipmt(args: dict[str, Any]) -> Decimal:
    rate = decimal_literal(args["rate"])
    periods = decimal_literal(args["periods"])
    value = decimal_literal(args["value"])
    start = integer_slot(decimal_literal(args["start"]))
    end = integer_slot(decimal_literal(args["end"]))
    pay_type = integer_control(decimal_literal(args["type"]))
    if rate <= 0 or value <= 0 or start < 1 or start > end or end > periods:
        raise OracleError(NUM)
    total = Decimal(0)
    period = int(start)
    while period <= int(end):
        total += ipmt_profile(rate, Decimal(period), periods, value, Decimal(0), pay_type)
        total = finite(total)
        period += 1
    return total


def cumprinc(args: dict[str, Any]) -> Decimal:
    rate = decimal_literal(args["rate"])
    periods = decimal_literal(args["periods"])
    value = decimal_literal(args["value"])
    start = integer_slot(decimal_literal(args["start"]))
    end = integer_slot(decimal_literal(args["end"]))
    pay_type = integer_control(decimal_literal(args["type"]))
    # CUMPRINC validates Type first, then permits a reversed empty range
    # without importing CUMIPMT's outer positivity constraints.
    if start > end:
        return Decimal(0)
    total = Decimal(0)
    period = int(start)
    while period <= int(end):
        total += ppmt_profile(rate, Decimal(period), periods, value, Decimal(0), pay_type)
        total = finite(total)
        period += 1
    return total


def fvschedule(args: dict[str, Any]) -> Decimal:
    principal = decimal_literal(args["principal"])
    product = Decimal(1)
    formula_error: OracleError | None = None
    generated_error: OracleError | None = None
    try:
        items = iter_sequence(args["schedule"])
        for item in items:
            kind, raw = kind_value(item)
            if kind in ("empty", "text", "logical"):
                continue
            if kind == "error":
                formula_error = first_error(formula_error, OracleError(raw, formula=True))
                continue
            try:
                rate = decimal_literal(raw)
                product = finite(product * (1 + rate))
            except OracleError as error:
                generated_error = first_error(generated_error, error)
    except OracleError as error:
        generated_error = first_error(generated_error, error)
    if formula_error is not None:
        raise formula_error
    if generated_error is not None:
        raise generated_error
    return finite(principal * product)


def npv_sum(rate: Decimal, values: Iterable[Decimal]) -> Decimal:
    if rate == 0:
        # Preserve cancellation across binary64 magnitudes.  A 100-digit
        # Decimal accumulator would erase a unit term between +/-1e300.
        exact_total = Fraction(0)
        for value in values:
            exact_total += Fraction(value)
        return finite(Decimal(exact_total.numerator) / Decimal(exact_total.denominator))
    total = Decimal(0)
    for index, value in enumerate(values, 1):
        denominator = number_power(1 + rate, Decimal(index))
        if denominator == 0:
            raise OracleError(DIV0)
        total = finite(total + value / denominator)
    return total


def npv(args: dict[str, Any]) -> Decimal:
    rate = decimal_literal(args["rate"])
    sequences = args.get("sequences")
    if not isinstance(sequences, list) or not sequences:
        raise OracleError(VALUE)
    total = Decimal(0)
    exact_total: Fraction | None = Fraction(0) if rate == 0 else None
    index = 0
    formula_error: OracleError | None = None
    generated_error: OracleError | None = None
    for sequence in sequences:
        try:
            elements = iter_sequence(sequence)
            for item in elements:
                kind, raw = kind_value(item)
                # NumberSequenceList omission rules do not assign an exponent
                # to Empty/Text/distinguished Logical cells.
                if kind in ("empty", "text", "logical"):
                    continue
                index += 1
                if kind == "error":
                    formula_error = first_error(formula_error, OracleError(raw, formula=True))
                    continue
                try:
                    value = decimal_literal(raw)
                    if exact_total is not None:
                        exact_total += Fraction(value)
                    else:
                        denominator = number_power(1 + rate, Decimal(index))
                        if denominator == 0:
                            generated_error = first_error(generated_error, OracleError(DIV0))
                        else:
                            total = finite(total + value / denominator)
                except OracleError as error:
                    generated_error = first_error(generated_error, error)
        except OracleError as error:
            generated_error = first_error(generated_error, error)
    if formula_error is not None:
        raise formula_error
    if generated_error is not None:
        raise generated_error
    if exact_total is not None:
        return finite(Decimal(exact_total.numerator) / Decimal(exact_total.denominator))
    return finite(total)


def mirr(args: dict[str, Any]) -> Decimal:
    investment = decimal_literal(args["investment"])
    reinvest = decimal_literal(args["reinvest_rate"])
    retained: list[Decimal] = []
    formula_error: OracleError | None = None
    generated_error: OracleError | None = None
    try:
        elements = iter_sequence(args["values"])
        for item in elements:
            kind, raw = kind_value(item)
            if kind in ("empty", "text"):
                continue
            if kind == "error":
                formula_error = first_error(formula_error, OracleError(raw, formula=True))
                continue
            if kind == "logical":
                retained.append(Decimal(1 if raw else 0))
                continue
            try:
                retained.append(decimal_literal(raw))
            except OracleError as error:
                generated_error = first_error(generated_error, error)
    except OracleError as error:
        generated_error = first_error(generated_error, error)
    if formula_error is not None:
        raise formula_error
    if generated_error is not None:
        raise generated_error
    if not retained or not any(value > 0 for value in retained) or not any(
        value < 0 for value in retained
    ):
        raise OracleError(NUM)
    count = len(retained)
    positive_mask = [value if value > 0 else Decimal(0) for value in retained]
    negative_mask = [value if value < 0 else Decimal(0) for value in retained]
    positive_npv = npv_sum(reinvest, positive_mask)
    negative_npv = npv_sum(investment, negative_mask)
    denominator = negative_npv * (1 + investment)
    if denominator == 0:
        raise OracleError(DIV0)
    ratio = finite(
        -positive_npv * number_power(1 + reinvest, Decimal(count)) / denominator
    )
    # MIRR's MathML expression ends in POWER.  Use the shared real POWER
    # profile here: a negative ratio remains valid for an exact integer
    # exponent, while a non-integer negative-base exponent is #NUM.  Zero
    # with a positive exponent is the ordinary zero POWER result.
    exponent = Decimal(1) / Decimal(count - 1)
    return finite(number_power(ratio, exponent) - 1)


def xnpv(args: dict[str, Any]) -> Decimal:
    rate = decimal_literal(args["rate"])
    # The complete Values/Dates scans still conceptually happen in the
    # evaluator.  Keep the rate-domain failure separate so a later formula
    # Error remains the first retained formula value after those scans.
    rate_error = OracleError(NUM) if rate <= -1 else None
    formula_error: OracleError | None = None
    generated_error: OracleError | None = None
    value_items: list[tuple[str, Any]] = []
    date_items: list[tuple[str, Any]] = []
    try:
        value_items = [kind_value(item) for item in iter_sequence(args["values"])]
    except OracleError as error:
        generated_error = first_error(generated_error, error)
    for kind, raw in value_items:
        if kind == "error":
            formula_error = first_error(formula_error, OracleError(raw, formula=True))
        elif kind != "number":
            generated_error = first_error(generated_error, OracleError(VALUE))
    try:
        date_items = [kind_value(item) for item in iter_sequence(args["dates"])]
    except OracleError as error:
        generated_error = first_error(generated_error, error)
    for kind, raw in date_items:
        if kind == "error":
            formula_error = first_error(formula_error, OracleError(raw, formula=True))
        elif kind != "number":
            generated_error = first_error(generated_error, OracleError(VALUE))
    if len(value_items) != len(date_items):
        generated_error = first_error(generated_error, OracleError(VALUE))
    if formula_error is not None:
        raise formula_error
    if rate_error is not None:
        raise rate_error
    if generated_error is not None:
        raise generated_error
    values = [decimal_literal(raw) for _, raw in value_items]
    dates = [decimal_literal(raw) for _, raw in date_items]
    if not values:
        raise OracleError(VALUE)
    if values[0] >= 0 or not any(value > 0 for value in values):
        raise OracleError(NUM)
    first_date = dates[0]
    if any(day < first_date for day in dates):
        raise OracleError(NUM)
    total = Decimal(0)
    exact_total: Fraction | None = Fraction(0) if rate == 0 else None
    for value, day in zip(values, dates):
        if exact_total is not None:
            exact_total += Fraction(value)
            continue
        exponent = (day - first_date) / Decimal(365)
        denominator = number_power(1 + rate, exponent)
        if denominator == 0:
            raise OracleError(DIV0)
        total = finite(total + value / denominator)
    if exact_total is not None:
        return finite(Decimal(exact_total.numerator) / Decimal(exact_total.denominator))
    return total


FUNCTIONS: dict[str, Callable[[dict[str, Any]], Decimal]] = {
    "CUMIPMT": cumipmt,
    "CUMPRINC": cumprinc,
    "FVSCHEDULE": fvschedule,
    "MIRR": mirr,
    "NPV": npv,
    "XNPV": xnpv,
}


def evaluate(function: str, inputs: dict[str, Any]) -> Decimal | OracleError:
    try:
        with localcontext(CONTEXT):
            return finite(FUNCTIONS[function](inputs))
    except OracleError as error:
        return error
    except (DivisionByZero, InvalidOperation, OverflowError, TypeError) as error:
        raise AssertionError(f"unmapped Decimal failure for {function}: {error}") from error


def canonical(value: Decimal) -> str:
    with localcontext(CONTEXT):
        value = +value
    if value == 0:
        return "0"
    text = format(value, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def conditioned_forward_error(ident: str) -> dict[str, str] | None:
    """Return the narrow cancellation bound for the three near-zero rows.

    This is an acceptance bound for finite binary64 forward evaluation.  It
    uses the same absolute-term scale as the contract's residual analysis,
    but does not claim that the solver residual rule governs these reducers.
    All other vectors retain zero absolute tolerance and the common relative
    tolerance.
    """

    if ident in {"npv.basic", "npv.split_argument_order"}:
        rate = decimal_literal("0.1")
        terms = (
            decimal_literal("-100") / number_power(1 + rate, Decimal(1)),
            decimal_literal("110") / number_power(1 + rate, Decimal(2)),
        )
    elif ident == "xnpv.basic":
        rate = decimal_literal("0.1")
        terms = (
            decimal_literal("-100"),
            decimal_literal("110") / number_power(1 + rate, Decimal(1)),
        )
    else:
        return None
    term_scale = sum((abs(term) for term in terms), Decimal(0))
    absolute = FORWARD_ERROR_MULTIPLIER * F64_EPSILON * term_scale
    return {
        "term_scale": canonical(term_scale),
        "absolute": canonical(absolute),
    }


def vector(
    ident: str,
    function: str,
    formula: str,
    inputs: dict[str, Any],
    tags: list[str],
    basis: str,
) -> dict[str, Any]:
    outcome = evaluate(function, inputs)
    conditioned_bound = conditioned_forward_error(ident)
    if isinstance(outcome, OracleError):
        expected: dict[str, Any] = {
            "kind": "error",
            "value": outcome.value,
            "error_class": "formula" if outcome.formula else "generated",
        }
    else:
        expected = {
            "kind": "number",
            "value": canonical(outcome),
            "subtype": "Number",
            "tolerance": {
                "absolute": conditioned_bound["absolute"] if conditioned_bound else "0",
                "relative": "1e-12",
            },
        }
    row = {
        "id": ident,
        "function": function,
        "formula": formula,
        "inputs": inputs,
        "expected": expected,
        "tags": tags,
        "oracle_basis": basis,
    }
    if conditioned_bound is not None:
        row["tolerance_basis"] = {
            "model": "16 * f64_epsilon * sum(abs(discounted_terms))",
            "f64_epsilon": canonical(F64_EPSILON),
            "term_scale": conditioned_bound["term_scale"],
            "absolute_bound": conditioned_bound["absolute"],
            "scope": "finite reducer forward evaluation only; solver residual tolerances do not govern this row",
        }
    return row


def vector_specs() -> list[dict[str, Any]]:
    rows: list[tuple[str, str, str, dict[str, Any], list[str], str]] = [
        (
            "cumipmt.basic",
            "CUMIPMT",
            "=CUMIPMT(0.1;2;100;1;2;0)",
            {"rate": "0.1", "periods": "2", "value": "100", "start": "1", "end": "2", "type": "0"},
            ["baseline", "ordered-periods"],
            "sum of the selected IPMT profile for Start through End",
        ),
        (
            "cumipmt.beginning_payment",
            "CUMIPMT",
            "=CUMIPMT(0.1;2;100;1;2;1)",
            {"rate": "0.1", "periods": "2", "value": "100", "start": "1", "end": "2", "type": "1"},
            ["pay-type", "ordered-periods"],
            "Type=1 uses zero first-period interest and the due divisor later",
        ),
        (
            "cumipmt.integer_type_truncates",
            "CUMIPMT",
            "=CUMIPMT(0.1;2;100;1;1;1.9)",
            {"rate": "0.1", "periods": "2", "value": "100", "start": "1", "end": "1", "type": "1.9"},
            ["integer-profile", "truncation"],
            "declared Integer Type truncates toward zero before exact 0/1 validation",
        ),
        (
            "cumipmt.invalid_outer_domain",
            "CUMIPMT",
            "=CUMIPMT(0;2;100;1;1;0)",
            {"rate": "0", "periods": "2", "value": "100", "start": "1", "end": "1", "type": "0"},
            ["domain-error"],
            "source Rate > 0 outer constraint",
        ),
        (
            "cumprinc.basic",
            "CUMPRINC",
            "=CUMPRINC(0.1;2;100;1;1;0)",
            {"rate": "0.1", "periods": "2", "value": "100", "start": "1", "end": "1", "type": "0"},
            ["baseline", "delegated-ppmt"],
            "sum of the selected PPMT profile",
        ),
        (
            "cumprinc.period_equal_nper_num",
            "CUMPRINC",
            "=CUMPRINC(0.1;2;100;2;2;0)",
            {"rate": "0.1", "periods": "2", "value": "100", "start": "2", "end": "2", "type": "0"},
            ["domain-error", "period-boundary"],
            "included PPMT term applies Period < Nper",
        ),
        (
            "cumprinc.reversed_empty",
            "CUMPRINC",
            "=CUMPRINC(-1;0;-100;2;1;0)",
            {"rate": "-1", "periods": "0", "value": "-100", "start": "2", "end": "1", "type": "0"},
            ["reversed-empty", "profile-boundary"],
            "Type passes, then reversed integer interval has an empty zero sum",
        ),
        (
            "cumprinc.invalid_type_on_empty",
            "CUMPRINC",
            "=CUMPRINC(0.1;2;100;2;1;2)",
            {"rate": "0.1", "periods": "2", "value": "100", "start": "2", "end": "1", "type": "2"},
            ["reversed-empty", "control-error"],
            "Type is validated even when the integer interval is empty",
        ),
        (
            "fvschedule.basic",
            "FVSCHEDULE",
            "=FVSCHEDULE(100;{0.1|0.2})",
            {"principal": "100", "schedule": ["0.1", "0.2"]},
            ["baseline", "ordered-sequence"],
            "Principal times product of one plus schedule elements in source order",
        ),
        (
            "fvschedule.negative_factor",
            "FVSCHEDULE",
            "=FVSCHEDULE(100;{-2})",
            {"principal": "100", "schedule": ["-2"]},
            ["negative-factor"],
            "source has no positivity constraint on schedule entries",
        ),
        (
            "fvschedule.zero_factor_continues",
            "FVSCHEDULE",
            "=FVSCHEDULE(100;{-1|0.5})",
            {"principal": "100", "schedule": ["-1", "0.5"]},
            ["zero-factor", "complete-scan"],
            "zero product remains zero while later factors remain admitted",
        ),
        (
            "fvschedule.formula_error_retained",
            "FVSCHEDULE",
            "=FVSCHEDULE(100;{0.1|NA()|0.2})",
            {
                "principal": "100",
                "schedule": ["0.1", {"kind": "error", "value": "#N/A"}, "0.2"],
            },
            ["formula-error", "complete-scan"],
            "formula Error is retained while the admitted sequence continues",
        ),
        (
            "npv.basic",
            "NPV",
            "=NPV(0.1;{-100|110})",
            {"rate": "0.1", "sequences": [["-100", "110"]]},
            ["baseline", "one-based"],
            "sum(Value[i] / (1 + Rate)^i), with i beginning at one",
        ),
        (
            "npv.split_argument_order",
            "NPV",
            "=NPV(0.1;{-100};{110})",
            {"rate": "0.1", "sequences": [["-100"], ["110"]]},
            ["argument-order", "exact"],
            "separate NumberSequenceList arguments continue one-based exponents",
        ),
        (
            "npv.one_based_with_middle_zero",
            "NPV",
            "=NPV(0.1;{-100|0|110})",
            {"rate": "0.1", "sequences": [["-100", "0", "110"]]},
            ["one-based", "zero-term"],
            "the zero occupies its source position and the 110 term is exponent three",
        ),
        (
            "npv.rate_zero_cancellation",
            "NPV",
            "=NPV(0;{-1e300|1|1e300})",
            {"rate": "0", "sequences": [["-1e300", "1", "1e300"]]},
            ["cancellation", "exact-rate-zero", "large-magnitude"],
            "rate-zero terms sum in source order but exact rational accumulation preserves the unit cancellation residue",
        ),
        (
            "npv.negative_base_integer_path",
            "NPV",
            "=NPV(-2;{100|20})",
            {"rate": "-2", "sequences": [["100", "20"]]},
            ["negative-base", "integer-power"],
            "checked signed integer powers preserve a finite real result below Rate=-1",
        ),
        (
            "npv.scaled_underflow_rescue",
            "NPV",
            "=NPV(0.5;{0 repeated 2000 times;1e300})",
            {
                "rate": "0.5",
                "sequences": [{"repeat": {"value": "0", "count": 2000}, "tail": ["1e300"]}],
            },
            ["regression", "underflow-rescue", "compact-sequence"],
            "the exponent-2001 tail term remains observable with scaled high-precision arithmetic",
        ),
        (
            "npv.rate_minus_one_divzero",
            "NPV",
            "=NPV(-1;{100})",
            {"rate": "-1", "sequences": [["100"]]},
            ["domain-error", "zero-denominator"],
            "Rate=-1 gives a zero first discount denominator",
        ),
        (
            "npv.rate_minus_one_formula_error",
            "NPV",
            "=NPV(-1;{100|NA()})",
            {"rate": "-1", "sequences": [["100", {"kind": "error", "value": "#N/A"}]]},
            ["domain-error", "formula-error", "precedence"],
            "the later retained formula Error supersedes the earlier generated zero-denominator error",
        ),
        (
            "npv.formula_error_retained",
            "NPV",
            "=NPV(0.1;{100|NA()|110})",
            {"rate": "0.1", "sequences": [["100", {"kind": "error", "value": "#N/A"}, "110"]]},
            ["formula-error", "complete-scan"],
            "formula Error wins after the complete sequence admission scan",
        ),
        (
            "mirr.basic",
            "MIRR",
            "=MIRR({-100|110};0.1;0.1)",
            {"values": ["-100", "110"], "investment": "0.1", "reinvest_rate": "0.1"},
            ["baseline", "mathml-equation"],
            "shared-position MathML masks give the corrected result 0.1",
        ),
        (
            "mirr.shared_masks",
            "MIRR",
            "=MIRR({-100|110|20};0.1;0.1)",
            {"values": ["-100", "110", "20"], "investment": "0.1", "reinvest_rate": "0.1"},
            ["shared-mask", "position-sensitive"],
            "positive and negative masks retain the same admitted positions and n",
        ),
        (
            "mirr.negative_ratio_integer_power",
            "MIRR",
            "=MIRR({110|-100};0;-2)",
            {"values": ["110", "-100"], "investment": "0", "reinvest_rate": "-2"},
            ["negative-base", "integer-power", "shared-power-profile"],
            "the final ratio is negative but the retained count is two, so POWER uses its exact integer exponent",
        ),
        (
            "mirr.zero_ratio_power",
            "MIRR",
            "=MIRR({110|110|-100};0;-2)",
            {"values": ["110", "110", "-100"], "investment": "0", "reinvest_rate": "-2"},
            ["zero-base", "power-profile", "shared-mask"],
            "the signed positive mask cancels to zero and POWER(0;1/(n-1)) remains zero",
        ),
        (
            "mirr.negative_ratio_noninteger_num",
            "MIRR",
            "=MIRR({-100|110|0};-2;-2)",
            {"values": ["-100", "110", "0"], "investment": "-2", "reinvest_rate": "-2"},
            ["negative-base", "noninteger-power", "domain-error"],
            "a negative final ratio with exponent 1/2 is outside the shared real POWER profile",
        ),
        (
            "mirr.text_empty_logical_compact",
            "MIRR",
            "=MIRR({-100|\"ignored\"|EMPTY|TRUE()|110};0.1;0.1)",
            {
                "values": [
                    "-100",
                    {"kind": "text", "value": "110"},
                    {"kind": "empty"},
                    {"kind": "logical", "value": True},
                    "110",
                ],
                "investment": "0.1",
                "reinvest_rate": "0.1",
            },
            ["text-empty", "logical", "shared-mask"],
            "Text and Empty are removed; Logical TRUE remains as a retained one",
        ),
        (
            "mirr.no_positive_num",
            "MIRR",
            "=MIRR({-100|0};0.1;0.1)",
            {"values": ["-100", "0"], "investment": "0.1", "reinvest_rate": "0.1"},
            ["domain-error", "sign"],
            "at least one retained positive value is required",
        ),
        (
            "mirr.no_negative_num",
            "MIRR",
            "=MIRR({0|110};0.1;0.1)",
            {"values": ["0", "110"], "investment": "0.1", "reinvest_rate": "0.1"},
            ["domain-error", "sign"],
            "at least one retained negative value is required",
        ),
        (
            "mirr.formula_error_retained",
            "MIRR",
            "=MIRR({-100|NA()|110};0.1;0.1)",
            {
                "values": ["-100", {"kind": "error", "value": "#N/A"}, "110"],
                "investment": "0.1",
                "reinvest_rate": "0.1",
            },
            ["formula-error", "complete-scan"],
            "formula Error is retained while the Array scan continues",
        ),
        (
            "xnpv.basic",
            "XNPV",
            "=XNPV(0.1;{-100|110};{43831|44196})",
            {"rate": "0.1", "values": ["-100", "110"], "dates": ["43831", "44196"]},
            ["baseline", "date-weighted"],
            "sum(Value / (1 + Rate)^((Date - first_date) / 365))",
        ),
        (
            "xnpv.nonannual_interval",
            "XNPV",
            "=XNPV(0.1;{-100|110};{43831|43861})",
            {"rate": "0.1", "values": ["-100", "110"], "dates": ["43831", "43861"]},
            ["date-weighted", "fractional-exponent"],
            "30-day date difference uses the days/365 exponent",
        ),
        (
            "xnpv.rate_zero_cancellation",
            "XNPV",
            "=XNPV(0;{-1e300|1|1e300};{43831|43831|43831})",
            {
                "rate": "0",
                "values": ["-1e300", "1", "1e300"],
                "dates": ["43831", "43831", "43831"],
            },
            ["cancellation", "exact-rate-zero", "large-magnitude"],
            "equal dates make every discount factor one; exact rational accumulation preserves the unit cancellation residue",
        ),
        (
            "xnpv.equal_count_different_geometry",
            "XNPV",
            "=XNPV(0.1;{-100|110|20|30};{43831;43832;43833;43834})",
            {
                "rate": "0.1",
                "values": {"elements": ["-100", "110", "20", "30"], "rows": 2, "columns": 2},
                "dates": {"elements": ["43831", "43832", "43833", "43834"], "rows": 1, "columns": 4},
            },
            ["shape", "equal-count", "row-major"],
            "equal flattened element counts are pairable despite differing rectangular geometry",
        ),
        (
            "xnpv.rate_minus_one_num",
            "XNPV",
            "=XNPV(-1;{-100|110};{43831|44196})",
            {"rate": "-1", "values": ["-100", "110"], "dates": ["43831", "44196"]},
            ["domain-error"],
            "XNPV requires Rate > -1",
        ),
        (
            "xnpv.rate_minus_one_formula_error",
            "XNPV",
            "=XNPV(-1;{NA()|110};{43831|44196})",
            {
                "rate": "-1",
                "values": [{"kind": "error", "value": "#N/A"}, "110"],
                "dates": ["43831", "44196"],
            },
            ["domain-error", "formula-error", "precedence"],
            "the retained Values formula Error supersedes the earlier generated Rate-domain error",
        ),
        (
            "xnpv.count_mismatch_value",
            "XNPV",
            "=XNPV(0.1;{-100|110};{43831})",
            {"rate": "0.1", "values": ["-100", "110"], "dates": ["43831"]},
            ["shape-error"],
            "Values and Dates must have equal element counts",
        ),
        (
            "xnpv.non_number_value",
            "XNPV",
            "=XNPV(0.1;{-100|\"110\"};{43831|44196})",
            {
                "rate": "0.1",
                "values": ["-100", {"kind": "text", "value": "110"}],
                "dates": ["43831", "44196"],
            },
            ["value-error", "element-type"],
            "XNPV does not convert Text elements to cash flows",
        ),
        (
            "xnpv_date_before_first_num",
            "XNPV",
            "=XNPV(0.1;{-100|110};{43831|43830})",
            {"rate": "0.1", "values": ["-100", "110"], "dates": ["43831", "43830"]},
            ["domain-error", "date-order"],
            "every date must be at least the first date",
        ),
        (
            "xnpv_sign_num",
            "XNPV",
            "=XNPV(0.1;{100|110};{43831|44196})",
            {"rate": "0.1", "values": ["100", "110"], "dates": ["43831", "44196"]},
            ["domain-error", "sign"],
            "first admitted cash flow must be negative and a positive flow must exist",
        ),
        (
            "xnpv.formula_error_retained",
            "XNPV",
            "=XNPV(0.1;{NA()|110};{43831|44196})",
            {
                "rate": "0.1",
                "values": [{"kind": "error", "value": "#N/A"}, "110"],
                "dates": ["43831", "44196"],
            },
            ["formula-error", "values-first"],
            "Values formula Error is retained before the complete Dates scan",
        ),
    ]
    return [vector(*row) for row in rows]


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def validate_normative_source() -> None:
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
        "schema": "ods-formula-financial-reducer-oracle-v1",
        "status": "draft independent Decimal corpus; six reducer/sequence functions only",
        "contract_sha256": sha256(CONTRACT),
        "normative_source": {
            "archive_sha256": NORMATIVE_ARCHIVE_SHA256,
            "part4_member_sha256": NORMATIVE_PART4_SHA256,
            "sections": "6.12.11-6.12.12, 6.12.21, 6.12.27, 6.12.30, 6.12.52 plus repository profiles",
        },
        "coverage": {
            "covered_functions": list(COVERED_FUNCTIONS),
            "pending_functions": list(PENDING_FUNCTIONS),
            "claim": "This corpus covers six reducer/sequence functions only; IRR, RATE, and XIRR remain pending and there is no all-twenty claim.",
        },
        "precision": {
            "decimal_digits": PRECISION,
            "rounding": "ROUND_HALF_EVEN during independent arithmetic",
            "input_model": "formula literals are converted with IEEE-754 binary64 float() and Decimal.from_float before Decimal arithmetic",
            "comparison": "all rows use a 1e-12 relative tolerance without a unit-scale floor; absolute tolerance is zero by default, with only the three named cancellation rows using their recorded conditioned forward-error bound",
            "cancellation": "rate-zero NPV/XNPV and MIRR mask sums use exact rational accumulation of the converted binary64 Decimal inputs before Decimal serialization",
            "conditioned_forward_error": "Only npv.basic, npv.split_argument_order, and xnpv.basic add 16 * f64_epsilon * sum(abs(discounted_terms)) as a row-specific absolute bound for finite rounded reducer evaluation; this does not alter the solver residual rule",
        },
        "profile": {
            "integer_conversion": "CUM Start/End/Type truncate toward zero; Type then requires exact 0 or 1",
            "sequence_order": "NPV and FVSCHEDULE preserve argument/element order; NPV exponents begin at one",
            "npv_negative_base": "integer exponents use checked signed powers below Rate=-1",
            "mirr": "Text and Empty are removed, Logical maps to 0/1, retained positions share positive/negative masks and n; the final ratio uses the shared real POWER profile",
            "xnpv": "Values and Dates pair by row-major slot index; equal counts permit differing rectangular geometry",
            "typed_errors": "formula Errors are retained after complete scans and precede generated errors",
        },
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
        print(f"verified {generated['vector_count']} financial reducer vectors")
        return 0
    print(f"generated {generated['vector_count']} vectors; use --write to update {VECTORS}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
