#!/usr/bin/env python3
"""Independent high-precision observations for descriptive statistics.

The reducer inputs are the represented IEEE-754 binary64 values retained in
the fixture JSON.  Additive and centered sums use exact ``Fraction`` values;
the geometric mean and the standardized moments use a 240-digit Decimal
context for logarithm/exponential/square-root operations.  No Rust reducer or
spreadsheet host is consulted when this file derives an expected result.

The generator writes a compact JSON corpus and ``--check`` verifies the
retained bytes.  Reference fixtures use the NumberSequence conversion profile
(only Number cells and formula Errors); direct scalar fixtures exercise the
scalar bridge; selected ReferenceList rows exercise admission and refusal.
"""

from __future__ import annotations

from decimal import Context, Decimal, localcontext
from fractions import Fraction
import json
import math
from pathlib import Path
import random
import struct
import sys


HERE = Path(__file__).resolve().parent
FUNCTIONS = (
    "AVEDEV",
    "DEVSQ",
    "GEOMEAN",
    "HARMEAN",
    "KURT",
    "SKEW",
    "SKEWP",
)

# Part 4 distinguishes NumberSequence from NumberSequenceList.  The local
# profile admits a ReferenceList only for the List signatures.
LIST_ADMITTED = {
    "AVEDEV": True,
    "DEVSQ": False,
    "GEOMEAN": True,
    "HARMEAN": True,
    "KURT": True,
    "SKEW": True,
    "SKEWP": False,
}

DECIMAL_PRECISION = 240
ORDINARY_MAX_ULPS = 0
SENSITIVE_MAX_ULPS = 8
TRANSCENDENTAL_MAX_ULPS = 8
SEED = 20260919

# These are the repository profile choices: a real odd root is retained for a
# negative GEOMEAN product, a negative even product is #NUM!, and HARMEAN's
# zero denominator is #DIV/0!.  The implementation deliberately does not
# import the native Excel positive-only restriction.
GEOMEAN_NEGATIVE_EVEN_ERROR = "Number"
HARMEAN_ZERO_ERROR = "DivisionByZero"

SENSITIVE_FIXTURES = {
    "large-offset",
    "nearmax-opposite",
    "nearmax-pair",
    "subnormal",
    "subnormal-skew-correction",
    "subnormal-skew-correction-mirror",
    "high-dynamic-signed-moments",
    "tiny-spread",
    "seeded-0",
    "seeded-1",
    "harmonic-limb-cancel",
    "harmonic-limb-reverse",
    "harmonic-limb-interleaved",
    "harmonic-limb-large-residual",
    "harmonic-limb-zero",
}

# Pin the final binary64 publication for the subnormal skew correction.  The
# signed mirror also catches a reducer that gets the magnitude right but loses
# the sign when applying the sample correction after subnormal rounding.
EXACT_SUBNORMAL_SKEW_FIXTURES = {
    "subnormal-skew-correction",
    "subnormal-skew-correction-mirror",
}


def bits(value: float, *, canonical_zero: bool = False) -> str:
    if canonical_zero and value == 0.0:
        value = 0.0
    return struct.pack(">d", float(value)).hex()


def number(value: float) -> dict[str, str]:
    return {"kind": "Number", "bits": bits(value)}


def text(value: str) -> dict[str, str]:
    return {"kind": "Text", "value": value}


def logical(value: bool) -> dict[str, object]:
    return {"kind": "Logical", "value": bool(value)}


def empty() -> dict[str, str]:
    return {"kind": "Empty"}


def error(value: str) -> dict[str, str]:
    return {"kind": "Error", "value": value}


def represented_number(cell: dict[str, str]) -> Fraction:
    value = struct.unpack(">d", bytes.fromhex(cell["bits"]))[0]
    return Fraction(value)


def add(datasets, name, cells, **metadata):
    datasets.append((name, cells, metadata))


def is_prime(value: int) -> bool:
    """Deterministic Miller-Rabin for the <2**64 fixture candidates."""

    if value < 2:
        return False
    for divisor in (2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37):
        if value % divisor == 0:
            return value == divisor
    odd = value - 1
    powers = 0
    while odd % 2 == 0:
        powers += 1
        odd //= 2
    for base in (2, 325, 9375, 28178, 450775, 9780504, 1795265022):
        base %= value
        if base == 0:
            continue
        witness = pow(base, odd, value)
        if witness in (1, value - 1):
            continue
        for _ in range(powers - 1):
            witness = (witness * witness) % value
            if witness == value - 1:
                break
        else:
            return False
    return True


def harmonic_limb_values() -> list[int]:
    """Return 100 distinct exactly represented odd reciprocal denominators."""

    candidate = (1 << 52) - 1
    values = []
    while len(values) < 100:
        if is_prime(candidate):
            values.append(candidate)
        candidate -= 2
    # The product is also the LCM because these candidates are distinct
    # primes.  Compute the LCM explicitly so the width check remains a direct
    # regression against a historical fixed-limb helper.
    lcm = math.lcm(*values)
    assert lcm.bit_length() > 4096, lcm.bit_length()
    assert all(float(value).is_integer() and float(value) == value for value in values)
    return values


def build_datasets():
    maximum = sys.float_info.max
    tiny = math.ulp(0.0)
    rng = random.Random(SEED)
    datasets = []

    add(datasets, "basic-odd", [number(x) for x in (1, 2, 3, 4, 5)])
    add(datasets, "even-middle", [number(x) for x in (1, 2, 4, 8)])
    add(datasets, "fractional", [number(x) for x in (0.1, 0.2, 0.3, 0.4)])
    add(datasets, "duplicates", [number(x) for x in (1, 1, 2, 2, 4)])
    add(
        datasets,
        "signed-zero",
        [number(x) for x in (-0.0, 0.0, -1.0, 1.0, 2.0)],
    )
    add(datasets, "negative", [number(x) for x in (-9, -3, -1, -7, -5)])
    add(datasets, "mixed-sign-even", [number(x) for x in (-4, -2, 1, 3)])
    add(datasets, "mixed-sign-odd", [number(x) for x in (-1, 2, 3)])
    add(datasets, "rational-cancel-0", [number(x) for x in (3.0, 6.0, -2.0)])
    add(datasets, "rational-cancel-1", [number(x) for x in (6.0, -2.0, 3.0)])
    add(datasets, "rational-cancel-2", [number(x) for x in (-2.0, 3.0, 6.0)])
    add(
        datasets,
        "extremes",
        [number(math.ulp(0.0)), number(-math.ulp(0.0)), number(1.0e100)],
    )
    add(
        datasets,
        "near-cancel-nonzero",
        [number(3.0), number(6.0), number(math.nextafter(-2.0, -math.inf))],
    )
    add(datasets, "nearmax-opposite", [number(x) for x in (maximum, -maximum, 0.0, 1.0)])
    add(
        datasets,
        "nearmax-pair",
        [number(maximum), number(math.nextafter(maximum, 0.0)), number(1.0), number(2.0)],
    )
    add(datasets, "subnormal", [number(tiny * x) for x in (1.0, 2.0, 3.0, 4.0)])
    add(
        datasets,
        "subnormal-skew-correction",
        [number(x) for x in (-1.0, 1.0, tiny)],
    )
    add(
        datasets,
        "subnormal-skew-correction-mirror",
        [number(x) for x in (1.0, -1.0, -tiny)],
    )
    add(
        datasets,
        "high-dynamic-signed-moments",
        [
            number(x)
            for x in (
                1.556634821042735e-10,
                6.130384255269732e18,
                6.259353243371187e-21,
                -2.23367199722874e-21,
                9.698988031985045e-15,
                -133087232.0,
                6556286976.0,
                -8.670476161507865e-20,
                -6.951982392533473e-13,
            )
        ],
    )
    add(
        datasets,
        "tiny-spread",
        [
            number(1.0),
            number(math.nextafter(1.0, math.inf)),
            number(math.nextafter(1.0, 0.0)),
            number(1.0),
        ],
    )
    large = 1.0e16
    add(
        datasets,
        "large-offset",
        [
            number(large),
            number(math.nextafter(large, math.inf)),
            number(math.nextafter(large, 0.0)),
            number(large),
        ],
    )

    order_values = [1.0e16, 1.0, -1.0e16, 2.0, 3.0]
    add(datasets, "order-0", [number(x) for x in order_values])
    add(datasets, "order-1", [number(x) for x in reversed(order_values)])
    add(datasets, "order-2", [number(x) for x in order_values[2:] + order_values[:2]])

    for index in range(2):
        values = []
        for _ in range(5 + index * 4):
            mantissa = (rng.getrandbits(53) + 1) / 2**53
            exponent = rng.randrange(-900, 901)
            values.append(math.ldexp(rng.choice((-1.0, 1.0)) * mantissa, exponent))
        add(datasets, f"seeded-{index}", [number(x) for x in values])

    add(datasets, "all-equal", [number(7.0)] * 5)
    add(datasets, "singleton", [number(7.0)])
    add(datasets, "pair", [number(1.0), number(3.0)])
    add(datasets, "triple", [number(1.0), number(2.0), number(4.0)])
    add(datasets, "zero-values", [number(0.0)] * 4)
    add(datasets, "harmonic-cancel", [number(x) for x in (-1.0, 1.0, 2.0, 4.0)])
    add(datasets, "geomean-negative-even", [number(x) for x in (-1.0, 2.0)])

    limb_values = [float(value) for value in harmonic_limb_values()]
    ordered_limb = limb_values + [-value for value in limb_values] + [1.0]
    add(datasets, "harmonic-limb-cancel", [number(x) for x in ordered_limb])
    add(
        datasets,
        "harmonic-limb-reverse",
        [number(x) for x in reversed(ordered_limb)],
    )
    add(
        datasets,
        "harmonic-limb-interleaved",
        [number(x) for value in limb_values for x in (value, -value)] + [number(1.0)],
    )
    add(
        datasets,
        "harmonic-limb-large-residual",
        [number(x) for x in ordered_limb[:-1]] + [number(1.0e100)],
    )
    add(
        datasets,
        "harmonic-limb-zero",
        [number(x) for x in ordered_limb[:-1]],
    )

    add(
        datasets,
        "typed-mixed",
        [number(5.0), text("ignored"), logical(True), empty(), number(-2.0), number(8.0)],
    )
    add(
        datasets,
        "typed-positive",
        [number(1.0), text("2"), logical(False), empty(), number(4.0)],
    )
    add(
        datasets,
        "typed-nonnumeric",
        [text("word"), logical(True), empty(), number(6.0)],
    )
    add(datasets, "error-first", [error("NotAvailable"), number(1.0), number(2.0)])
    add(datasets, "error-middle", [number(1.0), error("DivisionByZero"), number(2.0)])
    add(datasets, "error-last", [number(1.0), number(2.0), error("Number")])
    add(
        datasets,
        "multiple-errors",
        [number(1.0), error("DivisionByZero"), error("NotAvailable"), number(2.0)],
    )
    add(datasets, "empty", [empty()])

    add(datasets, "repeated-large", [number(0.0)] + [number(1.0e150)] * 4096)
    add(datasets, "repeated-unit", [number(0.0)] + [number(1.0)] * 4096)

    # These are direct scalar-conversion fixtures.  Their values are not used
    # as resolver cells by the scalar rows below.
    add(datasets, "direct-basic", [number(1.0), number(2.0), number(4.0), number(8.0)])
    add(datasets, "direct-mixed", [number(1.0), text("2"), logical(True), number(4.0)])
    add(datasets, "direct-negative", [number(-9.0), number(-3.0), number(-1.0)])
    add(datasets, "direct-singleton", [number(7.0)])
    # The formula syntax below renders this as a quoted empty Text literal.
    # A worksheet Empty is covered by the reference ``empty`` fixture; a
    # scalar empty slot is covered separately by ``missing-slot`` because the
    # contract treats an omitted slot as Missing rather than as a cell Empty.
    add(datasets, "direct-empty", [text("")])

    assert len(datasets) == 53, len(datasets)
    return datasets


def encode_fixture(name, cells):
    if name == "repeated-large":
        return {
            "encoding": "prefix_repeat_suffix",
            "prefix": [number(0.0)],
            "repeat": 4096,
            "repeated": number(1.0e150),
            "suffix": [],
        }
    if name == "repeated-unit":
        return {
            "encoding": "prefix_repeat_suffix",
            "prefix": [number(0.0)],
            "repeat": 4096,
            "repeated": number(1.0),
            "suffix": [],
        }
    return {"encoding": "cells", "cells": cells}


def unpack_cells(cells, origin):
    """Return exact admitted values and the first formula error."""

    values = []
    for cell in cells:
        kind = cell["kind"]
        if kind == "Error":
            return values, cell["value"]
        if kind == "Number":
            values.append(represented_number(cell))
        elif origin == "Scalar":
            if kind == "Text":
                try:
                    parsed = float(cell["value"])
                except ValueError:
                    return values, "Value"
                if not math.isfinite(parsed):
                    return values, "Value"
                values.append(Fraction(parsed))
            elif kind == "Logical":
                values.append(Fraction(int(cell["value"])))
            elif kind == "Empty":
                values.append(Fraction(0))
        # A referenced NumberSequence omits Text, Logical, and Empty cells.
    return values, None


def f64_fraction(value: Fraction):
    try:
        result = float(value)
    except OverflowError:
        return None
    return result if math.isfinite(result) else None


def decimal_fraction(value: Fraction, context: Context) -> Decimal:
    return context.divide(Decimal(value.numerator), Decimal(value.denominator))


def decimal_to_float(value: Decimal):
    try:
        result = float(value)
    except (OverflowError, ValueError):
        return None
    return result if math.isfinite(result) else None


def geometric_mean(values):
    if any(value == 0 for value in values):
        return Decimal(0)
    negative = sum(value < 0 for value in values) % 2 == 1
    if negative and len(values) % 2 == 0:
        return GEOMEAN_NEGATIVE_EVEN_ERROR
    context = Context(prec=DECIMAL_PRECISION)
    with localcontext(context):
        log_sum = sum(
            context.ln(decimal_fraction(abs(value), context)) for value in values
        )
        result = context.exp(log_sum / Decimal(len(values)))
        return -result if negative else result


def harmonic_mean(values):
    if any(value == 0 for value in values):
        return HARMEAN_ZERO_ERROR
    # Decide the zero-denominator domain case in exact rational arithmetic.
    # Decimal conversion is only for the finite publication path; checking the
    # represented reciprocals after conversion could erase a very small but
    # non-zero cancellation at the context precision.
    exact_denominator = sum((Fraction(1, 1) / value for value in values), Fraction(0))
    if exact_denominator == 0:
        return HARMEAN_ZERO_ERROR
    context = Context(prec=DECIMAL_PRECISION)
    with localcontext(context):
        denominator = decimal_fraction(exact_denominator, context)
        return Decimal(len(values)) / denominator


def centered(values):
    mean = sum(values, Fraction(0)) / len(values)
    deltas = [value - mean for value in values]
    sumsq = sum((delta * delta for delta in deltas), Fraction(0))
    return mean, deltas, sumsq


def standardized_moment(values, function):
    count = len(values)
    if function == "KURT" and count < 4:
        return {"error": "Value", "error_source": "cardinality"}
    if function in {"SKEW", "SKEWP"} and count < 3:
        return {"error": "Value", "error_source": "cardinality"}
    _, deltas, sumsq = centered(values)
    if sumsq == 0:
        return {"error": "Value", "error_source": "degenerate"}
    cubes = sum((delta**3 for delta in deltas), Fraction(0))
    if function in {"SKEW", "SKEWP"} and cubes == 0:
        # The central third moment is an exact Fraction.  Do not normalize
        # each residual independently in Decimal: a mathematically symmetric
        # sequence must remain an exact +0 after publication.
        return {"fraction": Fraction(0), "interpolated": False}
    fourths = sum((delta**4 for delta in deltas), Fraction(0))
    if function == "KURT":
        # The sample standard deviation's square root cancels after taking
        # the fourth power, so excess kurtosis is itself an exact rational.
        ratio_sum = fourths * (count - 1) ** 2 / (sumsq**2)
        result = (
            Fraction(count * (count + 1), (count - 1) * (count - 2) * (count - 3))
            * ratio_sum
            - Fraction(3 * (count - 1) ** 2, (count - 2) * (count - 3))
        )
        return {"fraction": result, "interpolated": False}
    context = Context(prec=DECIMAL_PRECISION)
    with localcontext(context):
        divisor = count - 1 if function == "SKEW" or function == "KURT" else count
        deviation = context.sqrt(decimal_fraction(sumsq / divisor, context))
        normalized_cubes = decimal_fraction(cubes, context) / (deviation**3)
        if function == "SKEW":
            result = Decimal(count) * normalized_cubes / Decimal(
                (count - 1) * (count - 2)
            )
        else:
            result = normalized_cubes / Decimal(count)
        return {"decimal": result, "decimal_precision": DECIMAL_PRECISION}


def exact_result(function, cells, origin, name):
    values, formula_error = unpack_cells(cells, origin)
    if formula_error is not None:
        return {"error": formula_error, "error_source": "formula"}
    if not values:
        error = (
            "DivisionByZero"
            if function in {"AVEDEV", "DEVSQ", "GEOMEAN", "HARMEAN"}
            else "Value"
        )
        return {"error": error, "error_source": "empty"}

    if function == "AVEDEV":
        mean = sum(values, Fraction(0)) / len(values)
        return {
            "fraction": sum((abs(value - mean) for value in values), Fraction(0)) / len(values),
            "interpolated": False,
        }
    if function == "DEVSQ":
        _, deltas, _ = centered(values)
        return {"fraction": sum((delta * delta for delta in deltas), Fraction(0))}
    if function == "GEOMEAN":
        result = geometric_mean(values)
        if isinstance(result, str):
            return {"error": result, "error_source": "domain"}
        return {"decimal": result, "decimal_precision": DECIMAL_PRECISION}
    if function == "HARMEAN":
        result = harmonic_mean(values)
        if isinstance(result, str):
            return {"error": result, "error_source": "domain"}
        return {"decimal": result, "decimal_precision": DECIMAL_PRECISION}
    if function in {"KURT", "SKEW", "SKEWP"}:
        return standardized_moment(values, function)
    raise AssertionError(function)


def scalar_token(cell):
    kind = cell["kind"]
    if kind == "Number":
        return repr(struct.unpack(">d", bytes.fromhex(cell["bits"]))[0])
    if kind == "Text":
        return '"' + cell["value"].replace('"', '""') + '"'
    if kind == "Logical":
        return "TRUE()" if cell["value"] else "FALSE()"
    if kind == "Empty":
        return '""'
    if kind == "Error":
        return "#" + cell["value"].upper().replace("NOTAVAILABLE", "N/A")
    raise AssertionError(kind)


def formula_for(function, cells, origin):
    if origin == "Reference":
        return f"={function}([.A1:.A{len(cells)}])"
    if origin == "ReferenceList":
        return f"={function}(([.A1:.A{len(cells)}]~[.A1:.A{len(cells)}]))"
    if origin == "Scalar":
        return f"={function}({';'.join(scalar_token(cell) for cell in cells)})"
    raise AssertionError(origin)


def row_for(function, name, cells, origin):
    result = {
        "function": function,
        "fixture": name,
        "origin": origin,
        "formula": formula_for(function, cells, origin),
    }
    if origin == "Scalar":
        # Direct scalar arguments are already evaluated values and must not
        # touch the resolver.
        result["expected_reads"] = 0
    if origin == "ReferenceList" and not LIST_ADMITTED[function]:
        result.update({"error": "Value", "error_source": "shape", "expected_reads": 0})
        return result
    observed = cells if origin != "ReferenceList" else cells + cells
    exact = exact_result(function, observed, "Reference" if origin != "Scalar" else "Scalar", name)
    result.update(exact)
    if "fraction" in result:
        fraction = result.pop("fraction")
        value = f64_fraction(fraction)
        if value is None:
            result.update({"error": "Number", "error_source": "binary64_overflow"})
            return result
        # Exact zero is canonical +0.  A nonzero Fraction that rounds to zero
        # retains its IEEE sign so the contract's signed-underflow rule stays
        # observable in the retained bits.
        result["expected_bits"] = bits(value, canonical_zero=fraction == 0)
        result["exact_numerator"] = str(fraction.numerator)
        result["exact_denominator"] = str(fraction.denominator)
        result.pop("interpolated", None)
        result["comparison"] = (
            "exact_bits"
            if value == 0.0
            or (function in {"SKEW", "SKEWP"} and name in EXACT_SUBNORMAL_SKEW_FIXTURES)
            or name not in SENSITIVE_FIXTURES
            else "ulps"
        )
        result["max_ulps"] = (
            ORDINARY_MAX_ULPS if result["comparison"] == "exact_bits" else SENSITIVE_MAX_ULPS
        )
    elif "decimal" in result:
        decimal = result.pop("decimal")
        value = decimal_to_float(decimal)
        if value is None:
            result.update({"error": "Number", "error_source": "binary64_overflow"})
            return result
        result["expected_bits"] = bits(value, canonical_zero=decimal == 0)
        result["decimal_reference"] = format(decimal, "f")
        result["comparison"] = (
            "exact_bits"
            if value == 0.0
            or (function in {"SKEW", "SKEWP"} and name in EXACT_SUBNORMAL_SKEW_FIXTURES)
            else "ulps"
        )
        result["max_ulps"] = (
            ORDINARY_MAX_ULPS if result["comparison"] == "exact_bits" else TRANSCENDENTAL_MAX_ULPS
        )
    return result


def document():
    datasets = build_datasets()
    cells_by_name = {name: cells for name, cells, _ in datasets}
    fixtures = {name: encode_fixture(name, cells) for name, cells, _ in datasets}
    fixtures["missing-slot"] = {"encoding": "cells", "cells": []}
    reference_names = [name for name, _, _ in datasets if not name.startswith("direct-")]
    scalar_names = [
        "direct-basic",
        "direct-mixed",
        "direct-negative",
        "direct-singleton",
        "direct-empty",
    ]
    rows = []
    for function in FUNCTIONS:
        for name in reference_names:
            rows.append(row_for(function, name, cells_by_name[name], "Reference"))

    list_names = (
        "basic-odd",
        "typed-positive",
        "negative",
        "nearmax-opposite",
        "repeated-unit",
    )
    for function in FUNCTIONS:
        for name in list_names:
            rows.append(row_for(function, name, cells_by_name[name], "ReferenceList"))

    for function in FUNCTIONS:
        for name in scalar_names:
            rows.append(row_for(function, name, cells_by_name[name], "Scalar"))
        rows.append(row_for(function, "missing-slot", [], "Scalar"))
        rows[-1]["formula"] = f"={function}(;)"
        rows[-1]["error"] = "Value"
        rows[-1]["error_source"] = "missing_parameter"

    return {
        "oracle": "Exact Fraction for represented binary64 sums; 240-digit Decimal ln/exp/sqrt for transcendental moments",
        "production_rounding_claim": False,
        "seed": SEED,
        "harmonic_limb_regression": {
            "distinct_odd_values": 100,
            "lcm_bit_length": math.lcm(*harmonic_limb_values()).bit_length(),
            "expected_harmean": 201,
            "large_residual_expected_harmean": "201e100",
            "zero_residual_error": HARMEAN_ZERO_ERROR,
        },
        "functions": list(FUNCTIONS),
        "list_admission": {
            function: "accept" if accepted else "reject"
            for function, accepted in LIST_ADMITTED.items()
        },
        "reference_read_policy": "admitted references/lists must be read; exact replay counts are covered by the resource target, while shape refusals and scalar rows retain exact zero-read assertions",
        "error_profile": {
            "empty_average_derived": "DivisionByZero",
            "cardinality_and_degenerate_moment": "Value",
            "geomean_negative_even_product": GEOMEAN_NEGATIVE_EVEN_ERROR,
            "harmean_zero_or_zero_denominator": HARMEAN_ZERO_ERROR,
            "non_finite_result": "Number",
            "explicit_missing_parameter": "Value",
        },
        "comparison_policy": {
            "ordinary_exact_bits": ORDINARY_MAX_ULPS,
            "sensitive_fraction_max_ulps": SENSITIVE_MAX_ULPS,
            "transcendental_max_ulps": TRANSCENDENTAL_MAX_ULPS,
            "zero": "canonical +0 for exact mathematical zero; a nonzero result rounded to zero retains its sign bit",
            "decimal_precision": DECIMAL_PRECISION,
            "rationale": "Exact additive fractions are bit-exact for ordinary fixtures. A fixed eight-ULP bound is reserved before execution for large-offset/subnormal and Decimal transcendental paths, covering one binary64 publication and compensated/log-exp/sqrt rounding without adapting to evaluator output.",
        },
        "fixtures": fixtures,
        "observations": rows,
    }


def main():
    retained = document()
    encoded = json.dumps(retained, indent=2) + "\n"
    path = HERE / "numeric-goldens.json"
    if "--check" in sys.argv:
        if path.read_text() != encoded:
            raise SystemExit("retained descriptive-statistics oracle differs")
        print(
            json.dumps(
                {
                    "observations": len(retained["observations"]),
                    "fixtures": len(retained["fixtures"]),
                    "verified": True,
                }
            )
        )
    else:
        path.write_text(encoded)
        print(
            f"Wrote {len(retained['observations'])} observations over "
            f"{len(retained['fixtures'])} fixtures"
        )


if __name__ == "__main__":
    main()
