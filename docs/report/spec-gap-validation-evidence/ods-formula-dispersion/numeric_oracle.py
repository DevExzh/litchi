#!/usr/bin/env python3
"""Independent exact references for the eight dispersion reducers.

The input values are first converted to their represented binary64 values and
then to :class:`fractions.Fraction`.  Variance goldens therefore retain every
operand bit exactly.  Standard-deviation goldens take the exact rational
variance through a 600-digit :class:`decimal.Decimal` square root before the
reference binary64 conversion.  No production evaluator, native application,
or floating-point running variance is consulted here.

The production kernel is a fixed-size scaled compensated first-offset state,
so this file deliberately does not claim correct rounding for production.
Ordinary rows use an eight-ULP comparison, matching the existing fixed-width
numeric aggregate product allowance.  Adversarial cancellation rows use
the existing database numeric profile's explicit 2e-12 relative bound and
require a finite result with the expected sign; an expected zero must still be
the exact positive zero.  The majority-equal order probes are retained under
that bound to expose order-dependent cancellation instead of widening the
tolerance until it disappears.
"""

from decimal import Decimal, localcontext
from fractions import Fraction
import json
import math
from pathlib import Path
import random
import struct
import sys


HERE = Path(__file__).resolve().parent
FUNCTIONS = (
    "VAR",
    "VARA",
    "VARP",
    "VARPA",
    "STDEV",
    "STDEVA",
    "STDEVP",
    "STDEVPA",
)
SAMPLE_FUNCTIONS = {"VAR", "VARA", "STDEV", "STDEVA"}
ADVERSARIAL_MARKERS = (
    "adjacent",
    "opposite",
    "tiny",
    "underflow",
    "majority",
    "order-",
)
ORDINARY_MAX_ULPS = 8
ADVERSARIAL_MAX_RELATIVE_ERROR = 2e-12
DECIMAL_SQRT_PRECISION = 600


def bits(value):
    return struct.pack(">d", float(value)).hex()


def number(value):
    return {"kind": "Number", "bits": bits(value)}


def text(value):
    return {"kind": "Text", "value": value}


def logical(value):
    return {"kind": "Logical", "value": bool(value)}


def empty():
    return {"kind": "Empty"}


def error(value):
    return {"kind": "Error", "value": value}


def represented_number(cell):
    return Fraction(struct.unpack(">d", bytes.fromhex(cell["bits"]))[0])


def fixture_numbers(cells, function):
    """Return (admitted exact values, first formula error), in source order."""
    values = []
    for cell in cells:
        kind = cell["kind"]
        if kind == "Error":
            return values, cell["value"]
        if kind == "Number":
            values.append(represented_number(cell))
        elif function.endswith("A"):
            if kind == "Text":
                values.append(Fraction(0))
            elif kind == "Logical":
                values.append(Fraction(int(cell["value"])))
            # Any-family Empty cells are omitted.
    return values, None


def decimal_sqrt_fraction(value):
    with localcontext() as context:
        context.prec = DECIMAL_SQRT_PRECISION
        root = (Decimal(value.numerator) / Decimal(value.denominator)).sqrt()
        # Keep a high-precision decimal audit value in the JSON while making
        # the binary64 conversion from the same high-precision object.
        audit = format(root, ".160g")
        rounded = float(root)
    return rounded, audit


def exact_result(function, cells):
    values, formula_error = fixture_numbers(cells, function)
    if formula_error is not None:
        return {"error": formula_error, "error_source": "formula"}
    minimum = 2 if function in SAMPLE_FUNCTIONS else 1
    if len(values) < minimum:
        return {"error": "Value", "error_source": "minimum_count"}
    total = sum(values, Fraction())
    mean = total / len(values)
    variance = sum(((value - mean) ** 2 for value in values), Fraction())
    variance /= len(values) - 1 if function in SAMPLE_FUNCTIONS else len(values)
    result = {
        "exact_numerator": str(variance.numerator),
        "exact_denominator": str(variance.denominator),
    }
    if function.startswith("VAR"):
        try:
            rounded = float(variance)
        except OverflowError:
            return {**result, "error": "Number", "error_source": "binary64_overflow"}
        if not math.isfinite(rounded):
            return {**result, "error": "Number", "error_source": "binary64_overflow"}
        result["reference"] = "exact_fraction_to_binary64"
    else:
        rounded, audit = decimal_sqrt_fraction(variance)
        if not math.isfinite(rounded):
            return {**result, "error": "Number", "error_source": "binary64_overflow"}
        result["reference"] = f"decimal_sqrt_{DECIMAL_SQRT_PRECISION}_digits_to_binary64"
        result["decimal_sqrt"] = audit
    result["expected_bits"] = bits(rounded)
    return result


def add_dataset(datasets, name, cells):
    datasets.append((name, cells))


def build_datasets():
    maximum = sys.float_info.max
    tiny = math.ulp(0.0)
    datasets = []

    # Ordinary and finite-limit binary64 vectors.
    add_dataset(datasets, "small-ascending", [number(x) for x in (1.0, 2.0, 3.0, 4.0)])
    add_dataset(datasets, "fractional", [number(x) for x in (0.1, 0.2, 0.3, 0.4)])
    add_dataset(datasets, "signed-zero", [number(0.0), number(-0.0), number(0.0)])
    add_dataset(datasets, "negative", [number(x) for x in (-9.0, -3.0, -1.0, -7.0)])
    add_dataset(datasets, "powers", [number(2.0**-40), number(1.0), number(2.0**40)])
    add_dataset(datasets, "large-moderate", [number(x) for x in (1e100, -1e100, 3.0, -7.0)])
    large = 1e150
    add_dataset(
        datasets,
        "adjacent-large",
        [number(large), number(math.nextafter(large, math.inf)), number(large)],
    )
    small = 1e-150
    add_dataset(
        datasets,
        "adjacent-small",
        [number(small), number(math.nextafter(small, math.inf)), number(small)],
    )
    add_dataset(datasets, "opposite-extremes", [number(maximum), number(-maximum)])
    add_dataset(datasets, "opposite-with-zero", [number(maximum), number(-maximum), number(0.0)])
    add_dataset(
        datasets,
        "tiny-subnormals",
        [number(tiny), number(2.0 * tiny), number(3.0 * tiny), number(0.0)],
    )
    add_dataset(
        datasets,
        "underflow-delta",
        [number(tiny), number(0.0), number(-tiny), number(2.0 * tiny)],
    )

    # Typed reference admission.  NumberSequence functions keep only Number
    # cells; A-family functions additionally map Text to zero and Logical to
    # zero or one, while every family omits Empty cells.
    add_dataset(
        datasets,
        "typed-mixed",
        [number(5.0), text("ignored"), logical(True), empty(), number(-2.0)],
    )
    add_dataset(
        datasets,
        "typed-empty-text",
        [number(5.0), text(""), logical(False), empty(), number(9.0)],
    )
    add_dataset(
        datasets,
        "typed-text-nonnumeric",
        [text("word"), text("42"), number(6.0), logical(False)],
    )
    add_dataset(
        datasets,
        "typed-logicals",
        [logical(False), logical(True), number(-4.0), number(8.0)],
    )
    add_dataset(datasets, "typed-empty", [empty(), empty(), number(4.0), number(8.0)])
    add_dataset(
        datasets,
        "typed-large",
        [number(maximum), text("x"), logical(False), empty(), number(-maximum)],
    )
    add_dataset(
        datasets,
        "typed-subnormal",
        [number(tiny), text("x"), logical(True), empty(), number(2.0 * tiny)],
    )
    add_dataset(
        datasets,
        "typed-decimal-text",
        [text("0.5"), text("-2"), logical(True), number(4.0)],
    )

    # Formula errors are retained in conceptual source order by every one of
    # these reducers.  The first error is the expected published error.
    add_dataset(
        datasets,
        "error-first",
        [error("NotAvailable"), number(1.0), number(2.0)],
    )
    add_dataset(
        datasets,
        "error-middle",
        [number(1.0), error("DivisionByZero"), number(2.0)],
    )
    add_dataset(
        datasets,
        "error-last",
        [number(1.0), number(2.0), error("Number")],
    )
    add_dataset(
        datasets,
        "multiple-errors",
        [number(1.0), error("DivisionByZero"), error("NotAvailable"), number(2.0)],
    )
    add_dataset(
        datasets,
        "error-empty",
        [empty(), error("NotAvailable"), empty()],
    )
    add_dataset(
        datasets,
        "error-typed",
        [text("x"), logical(True), error("Number"), number(8.0)],
    )

    # Empty and singleton constraints are intentionally represented as real
    # references, rather than omitted arguments.
    add_dataset(datasets, "empty-one", [empty()])
    add_dataset(datasets, "singleton-number", [number(7.0)])
    add_dataset(datasets, "singleton-negative", [number(-7.0)])
    add_dataset(datasets, "singleton-text", [text("x")])

    # The same represented values in different source orders make order
    # sensitivity visible without relying on a host calculation engine.
    order_values = [1e16, 1.0, -1e16, 2.0, 3.0]
    for index, values in enumerate(
        (
            order_values,
            list(reversed(order_values)),
            order_values[1:] + order_values[:1],
            order_values[2:] + order_values[:2],
            [-x for x in order_values],
            list(reversed([-x for x in order_values])),
            [1e100, 1.0, -1e100, 2.0],
            list(reversed([1e100, 1.0, -1e100, 2.0])),
        )
    ):
        add_dataset(datasets, f"order-{index}", [number(x) for x in values])

    rng = random.Random(20260919)
    for index in range(8):
        values = []
        for _ in range(2 + index * 3):
            mantissa = rng.getrandbits(53) / 2**53
            exponent = rng.randrange(-900, 901)
            values.append(math.ldexp(rng.choice((-1.0, 1.0)) * mantissa, exponent))
        add_dataset(datasets, f"random-numeric-{index}", [number(x) for x in values])

    for index in range(6):
        cells = []
        for offset in range(2 + index):
            if offset % 4 == 0:
                cells.append(text("text" if offset else ""))
            elif offset % 4 == 1:
                cells.append(logical(bool((offset + index) % 2)))
            elif offset % 4 == 2:
                cells.append(number(float(index - offset)))
            else:
                cells.append(empty())
        add_dataset(datasets, f"random-typed-{index}", cells)

    # Majority-equal first-outlier probes.  The 10,000-element pair follows
    # the review request directly; the 100,000-element unit pair makes the
    # n*epsilon cancellation visible under the retained 2e-12 bound.
    add_dataset(
        datasets,
        "majority-large-forward-10k",
        [number(0.0)] + [number(1e150)] * 10_000,
    )
    add_dataset(
        datasets,
        "majority-large-reverse-10k",
        [number(1e150)] * 10_000 + [number(0.0)],
    )
    add_dataset(
        datasets,
        "majority-unit-forward-100k",
        [number(0.0)] + [number(1.0)] * 100_000,
    )
    add_dataset(
        datasets,
        "majority-unit-reverse-100k",
        [number(1.0)] * 100_000 + [number(0.0)],
    )

    assert len(datasets) == 56, len(datasets)
    return datasets


def row_for(function, fixture_name, cells, reference_kind):
    if reference_kind == "Reference":
        formula = f"={function}([.A1:.A{len(cells)}])"
        expected_reads = len(cells)
    else:
        formula = f"={function}([.A1:.A{len(cells)}]~[.A1:.A{len(cells)}])"
        expected_reads = 2 * len(cells)

    # VAR, VARP, and STDEVP use NumberSequence, whereas STDEV uses
    # NumberSequenceList.  The four A variants use Any and admit a list.
    list_rejected = reference_kind == "ReferenceList" and function in {
        "VAR",
        "VARP",
        "STDEVP",
    }
    result = {"function": function, "fixture": fixture_name,
              "reference_kind": reference_kind, "formula": formula}
    if list_rejected:
        result.update({"error": "Value", "error_source": "shape", "expected_reads": 0})
        return result

    result["expected_reads"] = expected_reads
    observed_cells = cells + cells if reference_kind == "ReferenceList" else cells
    result.update(exact_result(function, observed_cells))
    if "error" not in result:
        adversarial = any(marker in fixture_name for marker in ADVERSARIAL_MARKERS)
        result["comparison"] = "relative" if adversarial else "ulps"
        if adversarial:
            result["max_relative_error"] = ADVERSARIAL_MAX_RELATIVE_ERROR
        else:
            result["max_ulps"] = ORDINARY_MAX_ULPS
    return result


def encode_fixture(name, cells):
    """Keep long equal runs analytic in the retained JSON evidence."""
    if name == "majority-large-forward-10k":
        return {
            "encoding": "prefix_repeat_suffix",
            "prefix": [number(0.0)],
            "repeat": 10_000,
            "repeated": number(1e150),
            "suffix": [],
        }
    if name == "majority-large-reverse-10k":
        return {
            "encoding": "prefix_repeat_suffix",
            "prefix": [],
            "repeat": 10_000,
            "repeated": number(1e150),
            "suffix": [number(0.0)],
        }
    if name == "majority-unit-forward-100k":
        return {
            "encoding": "prefix_repeat_suffix",
            "prefix": [number(0.0)],
            "repeat": 100_000,
            "repeated": number(1.0),
            "suffix": [],
        }
    if name == "majority-unit-reverse-100k":
        return {
            "encoding": "prefix_repeat_suffix",
            "prefix": [],
            "repeat": 100_000,
            "repeated": number(1.0),
            "suffix": [number(0.0)],
        }
    return {"encoding": "cells", "cells": cells}


def document():
    datasets = build_datasets()
    cells_by_name = {name: cells for name, cells in datasets}
    fixtures = {name: encode_fixture(name, cells) for name, cells in datasets}
    rows = []
    for function in FUNCTIONS:
        for name, cells in datasets:
            rows.append(row_for(function, name, cells, "Reference"))

    # Keep a compact but representative list-admission matrix.  Reusing the
    # fixture payload means the JSON does not duplicate large cell vectors.
    list_names = (
        "small-ascending",
        "typed-mixed",
        "adjacent-large",
        "opposite-extremes",
        "tiny-subnormals",
        "random-numeric-0",
        "singleton-number",
        "majority-large-reverse-10k",
    )
    for function in FUNCTIONS:
        for name in list_names:
            rows.append(row_for(function, name, cells_by_name[name], "ReferenceList"))

    assert len(rows) == 512
    return {
        "oracle": "Python Fraction from represented binary64 operands; Decimal sqrt at 600 digits",
        "production_rounding_claim": False,
        "seed": 20260919,
        "functions": list(FUNCTIONS),
        "sample_functions": sorted(SAMPLE_FUNCTIONS),
        "cardinality_error": "Value",
        "list_admission": {
            "VAR": "reject",
            "VARP": "reject",
            "STDEVP": "reject",
            "STDEV": "accept",
            "VARA": "accept",
            "VARPA": "accept",
            "STDEVA": "accept",
            "STDEVPA": "accept",
        },
        "comparison_policy": {
            "ordinary_max_ulps": ORDINARY_MAX_ULPS,
            "adversarial_max_relative_error": ADVERSARIAL_MAX_RELATIVE_ERROR,
            "adversarial_zero": "expected zero requires exact positive zero",
            "rationale": "The relative bound is the existing database numerical profile; majority-equal rows remain under it to expose n*epsilon cancellation.",
        },
        "fixtures": fixtures,
        "observations": rows,
    }


def main():
    encoded = json.dumps(document(), indent=2) + "\n"
    path = HERE / "numeric-goldens.json"
    if "--check" in sys.argv:
        if path.read_text() != encoded:
            raise SystemExit("retained dispersion oracle differs")
        print(json.dumps({"observations": 512, "fixtures": 56, "verified": True}))
    else:
        path.write_text(encoded)
        print("Wrote 512 observations over 56 typed fixtures")


if __name__ == "__main__":
    main()
