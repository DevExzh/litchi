#!/usr/bin/env python3
"""Independent exact references for the order-statistics batch.

The observations in this file are derived from the represented IEEE-754
binary64 operands, never from the Rust evaluator or a spreadsheet host.  The
ordering and rank decisions use :class:`fractions.Fraction`; interpolation is
performed as an exact rational convex combination and converted to binary64
once at publication.  This matters for the max/min interpolation probes,
where ``upper - lower`` can overflow even though the mathematical result is
finite.

The eight functions intentionally share one compact typed fixture corpus.  A
fixture is used once as a reference and a small selected subset is repeated as
an ordered ReferenceList.  The latter is where pseudotype/list refusal and
duplicate occurrence semantics are checked.  Long repeated fixtures use a
prefix/repeat/suffix encoding so the retained JSON does not contain a cell
vector merely to describe a bounded scan.

Formula Number parameters are parsed through binary64 before their exact
Fraction representation is used.  Thus rank and interpolation arithmetic is
exact for the operands the evaluator receives, including decimal literals such
as 0.1, rather than exact for their source spelling.

This module is an evidence generator.  ``--check`` regenerates the bytes and
fails if the retained ``numeric-goldens.json`` differs.
"""

from __future__ import annotations

from collections import Counter
from fractions import Fraction
import json
import math
from pathlib import Path
import random
import struct
import sys


HERE = Path(__file__).resolve().parent
FUNCTIONS = (
    "MEDIAN",
    "MODE",
    "LARGE",
    "SMALL",
    "PERCENTILE",
    "PERCENTRANK",
    "QUARTILE",
    "RANK",
)

# A NumberSequenceList is admitted by these functions.  MODE follows the
# ForceArray/NumberSequence profile and QUARTILE follows NumberSequence, so an
# explicit multi-area union is rejected before a resolver read.
LIST_ADMITTED = {
    "MEDIAN": True,
    "MODE": False,
    "LARGE": True,
    "SMALL": True,
    "PERCENTILE": True,
    "PERCENTRANK": True,
    "QUARTILE": False,
    "RANK": True,
}

ORDINARY_MAX_ULPS = 2
INTERPOLATION_MAX_ULPS = 4
DECIMAL_PLACES = 3
SEED = 20260919


def bits(value: float, *, canonical_zero: bool = False) -> str:
    """Return a big-endian binary64 bit spelling, preserving signed zero."""

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


def build_datasets():
    maximum = sys.float_info.max
    tiny = math.ulp(0.0)
    large = 1e150
    datasets = []

    # Stable selection, ties, signed zero, and ordinary interpolation.
    add(datasets, "small-ascending", [number(x) for x in (1, 2, 3, 4, 5)])
    add(datasets, "even-middle", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "duplicates", [number(x) for x in (1, 2, 2, 3, 4, 4)])
    add(datasets, "mode-tie", [number(x) for x in (1, 1, 2, 2, 3)])
    add(datasets, "no-mode", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "all-equal", [number(x) for x in (7, 7, 7, 7, 7)])
    add(datasets, "signed-zero", [number(-0.0), number(0.0), number(-0.0), number(0.0), number(1.0), number(-1.0)])
    add(datasets, "fractional", [number(x) for x in (0.1, 0.2, 0.3, 0.4)])
    add(datasets, "negative", [number(x) for x in (-9, -3, -1, -7, -5)])
    add(datasets, "powers", [number(2.0**-40), number(1.0), number(2.0**40), number(-2.0**20)])
    add(datasets, "large-moderate", [number(x) for x in (1e100, -1e100, 3.0, -7.0, 11.0)])
    add(datasets, "adjacent-large", [number(large), number(math.nextafter(large, math.inf)), number(large), number(math.nextafter(large, 0.0))])
    add(datasets, "adjacent-small", [number(1e-150), number(math.nextafter(1e-150, math.inf)), number(1e-150), number(math.nextafter(1e-150, 0.0))])
    add(datasets, "opposite-extremes", [number(maximum), number(-maximum), number(0.0), number(1.0)])
    add(datasets, "opposite-with-zero", [number(maximum), number(-maximum), number(0.0)])
    add(datasets, "tiny-subnormals", [number(tiny), number(2.0 * tiny), number(3.0 * tiny), number(0.0)])
    add(datasets, "underflow-delta", [number(tiny), number(0.0), number(-tiny), number(2.0 * tiny)])
    add(datasets, "interpolation-overflow", [number(maximum), number(-maximum)])
    add(datasets, "binary-boundary", [number(2.0**53 - 1), number(2.0**53), number(2.0**53 + 2), number(-(2.0**53))])
    add(datasets, "rank-ties", [number(x) for x in (1, 2, 2, 2, 5, 7)])
    # The duplicated lower bracket exercises PERCENTRANK's first-occurrence
    # rank when X lies strictly between two represented values.
    add(datasets, "rank-ascending", [number(x) for x in (1, 3, 3, 7, 9)])
    add(datasets, "rank-descending", [number(x) for x in (9, 7, 5, 3, 1)])
    add(datasets, "percentile-flat", [number(x) for x in (4, 4, 4, 4)])
    add(datasets, "percentile-extreme", [number(x) for x in (-1e308, -1.0, 1.0, 1e308)])

    # Reference cell typing is deliberate: sequence references omit Text,
    # Logical, and Empty cells, while formula Errors remain visible.
    add(datasets, "typed-mixed", [number(5.0), text("ignored"), logical(True), empty(), number(-2.0), number(8.0)])
    add(datasets, "typed-nonnumeric", [text("word"), text("42"), number(6.0), logical(False), empty()])
    add(datasets, "typed-logicals", [logical(False), logical(True), number(-4.0), number(8.0)])
    add(datasets, "negative-underflow-midpoint", [number(-tiny), number(0.0)])
    add(datasets, "typed-large", [number(maximum), text("x"), logical(False), empty(), number(-maximum)])
    add(datasets, "typed-subnormal", [number(tiny), text("x"), logical(True), empty(), number(2.0 * tiny)])
    add(datasets, "typed-decimal-text", [text("0.5"), text("-2"), logical(True), number(4.0)])

    # Formula errors remain values in the independent model.  The evaluator
    # must retain the first one while continuing any admitted scan.
    add(datasets, "error-first", [error("NotAvailable"), number(1.0), number(2.0), number(3.0)])
    add(datasets, "error-middle", [number(1.0), error("DivisionByZero"), number(2.0), number(3.0)])
    add(datasets, "error-last", [number(1.0), number(2.0), number(3.0), error("Number")])
    add(datasets, "multiple-errors", [number(1.0), error("DivisionByZero"), error("NotAvailable"), number(2.0)])
    add(datasets, "error-empty", [empty(), error("NotAvailable"), empty()])
    add(datasets, "empty-one", [empty()])
    add(datasets, "singleton-number", [number(7.0)])
    add(datasets, "singleton-negative", [number(-7.0)])
    add(datasets, "singleton-zero", [number(0.0)])

    # Parameter/domain refusals.  Each invalid case has enough ordinary data
    # that its error is attributable to the parameter rather than emptiness.
    add(datasets, "invalid-k-zero", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "invalid-k-fraction", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "invalid-k-large", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "invalid-percentile-low", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "invalid-percentile-high", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "invalid-rank-outside", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "invalid-significance", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "invalid-quartile", [number(x) for x in (1, 2, 3, 4)])
    add(datasets, "missing-parameter", [number(x) for x in (1, 2, 3, 4)])

    # The same represented values in different source orders exercise the
    # order-independent sort and deterministic tie selection.
    order_values = [1e16, 1.0, -1e16, 2.0, 3.0]
    for index, values in enumerate(
        (order_values, list(reversed(order_values)), order_values[2:] + order_values[:2])
    ):
        add(datasets, f"order-{index}", [number(x) for x in values])

    rng = random.Random(SEED)
    for index in range(2):
        values = []
        for _ in range(5 + index * 4):
            mantissa = (rng.getrandbits(53) + 1) / 2**53
            exponent = rng.randrange(-900, 901)
            values.append(math.ldexp(rng.choice((-1.0, 1.0)) * mantissa, exponent))
        add(datasets, f"random-numeric-{index}", [number(x) for x in values])

    # Compact repeated fixtures make the scan and exact tie behavior visible
    # without retaining thousands of duplicate JSON objects.
    add(datasets, "repeated-large", [number(0.0)] + [number(1e150)] * 4096)
    add(datasets, "repeated-unit", [number(0.0)] + [number(1.0)] * 4096)

    assert len(datasets) == 56, len(datasets)
    return datasets


def encode_fixture(name, cells):
    if name == "repeated-large":
        return {"encoding": "prefix_repeat_suffix", "prefix": [number(0.0)], "repeat": 4096,
                "repeated": number(1e150), "suffix": []}
    if name == "repeated-unit":
        return {"encoding": "prefix_repeat_suffix", "prefix": [number(0.0)], "repeat": 4096,
                "repeated": number(1.0), "suffix": []}
    return {"encoding": "cells", "cells": cells}


def unpack_cells(cells):
    """Return admitted exact numbers and the first formula Error, if any."""

    values = []
    for cell in cells:
        kind = cell["kind"]
        if kind == "Error":
            return values, cell["value"]
        if kind == "Number":
            values.append(represented_number(cell))
        # NumberSequence reference conversion deliberately omits the other
        # distinguished cell kinds.  Inline arrays are not used by this corpus.
    return values, None


def f64_fraction(value: Fraction):
    try:
        result = float(value)
    except OverflowError:
        return None
    return result if math.isfinite(result) else None


def decimal_round_fraction(value: Fraction, places: int) -> Fraction:
    """Round a non-negative rank to decimal places, ties away from zero."""

    scale = 10**places
    numerator = value.numerator * scale
    quotient, remainder = divmod(numerator, value.denominator)
    if remainder * 2 >= value.denominator:
        quotient += 1
    return Fraction(quotient, scale)


def parameter_fraction(value: str) -> Fraction:
    """Convert a formula Number literal through its represented binary64."""

    return Fraction(float(value))


def percentile(values, fraction):
    ordered = sorted(values)
    rank = Fraction(1) + fraction * (len(ordered) - 1)
    lower = rank.numerator // rank.denominator
    distance = rank - lower
    left = ordered[lower - 1]
    if distance == 0:
        return left, False
    right = ordered[lower]
    return left + distance * (right - left), True


def parameter(name, function, numeric_values):
    n = len(numeric_values)
    if function in {"LARGE", "SMALL"} and name == "invalid-k-zero":
        return {"k": "0"}
    if function in {"LARGE", "SMALL"} and name == "invalid-k-fraction":
        return {"k": "2.5"}
    if function in {"LARGE", "SMALL"} and name == "invalid-k-large":
        return {"k": str(n + 1)}
    if function == "PERCENTILE" and name == "invalid-percentile-low":
        return {"p": "-0.1"}
    if function == "PERCENTILE" and name == "invalid-percentile-high":
        return {"p": "1.1"}
    if function == "PERCENTRANK" and name == "invalid-rank-outside":
        return {"x": "100", "significance": "3"}
    if function == "RANK" and name == "invalid-rank-outside":
        return {"x": "100", "order": "1"}
    if function == "PERCENTRANK" and name == "invalid-significance":
        return {"x": "2", "significance": "0"}
    if function == "QUARTILE" and name == "invalid-quartile":
        return {"quartile": "5"}
    if name == "missing-parameter":
        return {"missing": True}
    if function in {"LARGE", "SMALL"}:
        return {"k": str(max(1, min(n, 2)))}
    if function == "PERCENTILE":
        choices = {"percentile-flat": "0.3", "percentile-extreme": "0.9",
                   "interpolation-overflow": "0.5", "negative-underflow-midpoint": "0.5",
                   "even-middle": "0.3"}
        return {"p": choices.get(name, "0.25")}
    if function == "PERCENTRANK":
        if name == "singleton-number":
            return {"x": "7", "significance": "3"}
        if name == "singleton-zero":
            return {"x": "0", "significance": "3"}
        if name == "rank-ties":
            return {"x": "2", "significance": "3"}
        if name == "rank-ascending":
            return {"x": "5", "significance": "2"}
        if name == "even-middle":
            return {"x": "3", "significance": "4"}
        if name in {"opposite-extremes", "interpolation-overflow"}:
            return {"x": "0", "significance": "3"}
        if name == "small-ascending":
            return {"x": "3", "significance": "3", "omit_significance": True}
        return {"x": "3", "significance": "3"}
    if function == "QUARTILE":
        choices = {"percentile-flat": "3", "percentile-extreme": "1",
                   "opposite-extremes": "2", "interpolation-overflow": "2",
                   "negative-underflow-midpoint": "2"}
        return {"quartile": choices.get(name, "1")}
    if function == "RANK":
        if name == "invalid-rank-outside":
            return {"x": "100", "order": "1"}
        if name in {"rank-ties", "duplicates", "mode-tie"}:
            return {"x": "2", "order": "1"}
        if name == "singleton-negative":
            return {"x": "-7", "order": "0"}
        if name == "singleton-zero":
            return {"x": "0", "order": "0"}
        if name in {"order-1", "rank-ascending"}:
            return {"x": "3", "order": "1"}
        if name == "small-ascending":
            return {"x": "3", "order": "0", "omit_order": True}
        return {"x": "3", "order": "0"}
    return {}


def formula_for(function, name, length, reference_kind, params):
    if params.get("missing"):
        reference = f"[.A1:.A{length}]"
        if function in {"MEDIAN", "MODE"}:
            return f"={function}({reference};)"
        if function in {"LARGE", "SMALL"}:
            return f"={function}({reference};)"
        if function == "PERCENTILE":
            return f"=PERCENTILE({reference};)"
        if function == "PERCENTRANK":
            return f"=PERCENTRANK({reference};3;)"
        if function == "QUARTILE":
            return f"=QUARTILE({reference};)"
        return f"=RANK(3;{reference};)"

    if reference_kind == "Reference":
        reference = f"[.A1:.A{length}]"
    else:
        reference = f"([.A1:.A{length}]~[.A1:.A{length}])"

    if function in {"MEDIAN", "MODE"}:
        return f"={function}({reference})"
    if function in {"LARGE", "SMALL"}:
        return f"={function}({reference};{params['k']})"
    if function == "PERCENTILE":
        return f"=PERCENTILE({reference};{params['p']})"
    if function == "PERCENTRANK":
        if params.get("omit_significance"):
            return f"=PERCENTRANK({reference};{params['x']})"
        significance = params.get("significance", "3")
        return f"=PERCENTRANK({reference};{params['x']};{significance})"
    if function == "QUARTILE":
        return f"=QUARTILE({reference};{params['quartile']})"
    if params.get("omit_order"):
        return f"=RANK({params['x']};{reference})"
    return f"=RANK({params['x']};{reference};{params['order']})"


def exact_result(function, cells, name):
    values, formula_error = unpack_cells(cells)
    if formula_error is not None:
        return {"error": formula_error, "error_source": "formula"}
    params = parameter(name, function, values)
    if params.get("missing"):
        return {"error": "Value", "error_source": "missing_parameter"}
    if function in {"LARGE", "SMALL"}:
        k = parameter_fraction(params["k"])
        if k.denominator != 1 or k < 1 or k > len(values):
            return {"error": "Value", "error_source": "domain"}
        result = sorted(values, reverse=function == "LARGE")[int(k) - 1]
        return {"fraction": result, "interpolated": False}
    if function == "MEDIAN":
        if not values:
            return {"error": "Value", "error_source": "empty"}
        ordered = sorted(values)
        middle = len(ordered) // 2
        if len(ordered) % 2:
            return {"fraction": ordered[middle], "interpolated": False}
        return {"fraction": (ordered[middle - 1] + ordered[middle]) / 2,
                "interpolated": True}
    if function == "MODE":
        if not values:
            return {"error": "Value", "error_source": "empty"}
        counts = Counter(values)
        maximum_count = max(counts.values())
        if maximum_count < 2:
            return {"error": "Value", "error_source": "no_mode"}
        return {"fraction": min(value for value, count in counts.items()
                                  if count == maximum_count), "interpolated": False}
    if function == "PERCENTILE":
        if not values:
            return {"error": "Value", "error_source": "empty"}
        p = parameter_fraction(params["p"])
        if p < 0 or p > 1:
            return {"error": "Value", "error_source": "domain"}
        result, interpolated = percentile(values, p)
        return {"fraction": result, "interpolated": interpolated}
    if function == "PERCENTRANK":
        if not values:
            return {"error": "Value", "error_source": "empty"}
        x = parameter_fraction(params["x"])
        significance = parameter_fraction(params["significance"])
        if significance.denominator != 1 or significance < 1:
            return {"error": "Value", "error_source": "domain"}
        ordered = sorted(values)
        if x < ordered[0] or x > ordered[-1]:
            return {"error": "Value", "error_source": "domain"}
        if len(ordered) == 1:
            if x != ordered[0]:
                return {"error": "Value", "error_source": "domain"}
            result = Fraction(1)
        else:
            try:
                exact_index = ordered.index(x)
            except ValueError:
                upper = next(index for index, value in enumerate(ordered) if value > x)
                lower = upper - 1
                left, right = ordered[lower], ordered[upper]
                # The lower bracket may be repeated.  Its rank is the first
                # occurrence (the number of values strictly below it), not
                # the final duplicate index immediately before ``upper``.
                lower_first = ordered.index(left)
                rank = Fraction(lower_first) + (x - left) / (right - left)
            else:
                rank = Fraction(exact_index)
            result = rank / (len(ordered) - 1)
        rounded = decimal_round_fraction(result, int(significance))
        return {"fraction": rounded, "interpolated": True, "rounded": True}
    if function == "QUARTILE":
        if not values:
            return {"error": "Value", "error_source": "empty"}
        quartile = parameter_fraction(params["quartile"])
        if quartile.denominator != 1 or quartile < 0 or quartile > 4:
            return {"error": "Value", "error_source": "domain"}
        result, interpolated = percentile(values, quartile / 4)
        return {"fraction": result, "interpolated": interpolated}
    if function == "RANK":
        x = parameter_fraction(params["x"])
        if x not in values:
            return {"error": "Value", "error_source": "missing_value"}
        order = parameter_fraction(params.get("order", "0"))
        if order == 0:
            rank = 1 + sum(value > x for value in values)
        else:
            rank = 1 + sum(value < x for value in values)
        return {"fraction": Fraction(rank), "interpolated": False}
    raise AssertionError(function)


def row_for(function, name, cells, reference_kind):
    values, _ = unpack_cells(cells)
    params = parameter(name, function, values)
    result = {"function": function, "fixture": name,
              "reference_kind": reference_kind,
              "formula": formula_for(function, name, len(cells), reference_kind, params),
              "expected_reads": len(cells) if reference_kind == "Reference" else 2 * len(cells)}
    if reference_kind == "ReferenceList" and not LIST_ADMITTED[function]:
        result.update({"error": "Value", "error_source": "shape", "expected_reads": 0})
        return result
    observed_cells = cells if reference_kind == "Reference" else cells + cells
    result.update(exact_result(function, observed_cells, name))
    if "fraction" in result:
        exact_fraction = result.pop("fraction")
        rounded = f64_fraction(exact_fraction)
        if rounded is None:
            result.update({"error": "Number", "error_source": "binary64_overflow"})
        else:
            result["expected_bits"] = bits(rounded)
            result["exact_numerator"] = str(exact_fraction.numerator)
            result["exact_denominator"] = str(exact_fraction.denominator)
            interpolated = result.pop("interpolated", False)
            # Both positive and negative underflow zero are sign-sensitive
            # computed results; selected zero is always the canonical +0.
            result["comparison"] = "exact_bits" if rounded == 0.0 else (
                "ulps" if interpolated else "exact_bits"
            )
            result["max_ulps"] = (
                INTERPOLATION_MAX_ULPS
                if result["comparison"] == "ulps"
                else ORDINARY_MAX_ULPS
            )
    return result


def document():
    datasets = build_datasets()
    cells_by_name = {name: cells for name, cells, _ in datasets}
    fixtures = {name: encode_fixture(name, cells) for name, cells, _ in datasets}
    rows = []
    for function in FUNCTIONS:
        for name, cells, _ in datasets:
            rows.append(row_for(function, name, cells, "Reference"))

    list_names = (
        "small-ascending",
        "duplicates",
        "typed-mixed",
        "opposite-extremes",
        "random-numeric-0",
        "singleton-number",
        "rank-ties",
        "repeated-large",
    )
    for function in FUNCTIONS:
        for name in list_names:
            rows.append(row_for(function, name, cells_by_name[name], "ReferenceList"))
    assert len(rows) == 512, len(rows)
    return {
        "oracle": "Python Fraction from represented binary64 operands; exact rational sorting/interpolation",
        "production_rounding_claim": False,
        "seed": SEED,
        "functions": list(FUNCTIONS),
        "list_admission": {name: ("accept" if accepted else "reject")
                           for name, accepted in LIST_ADMITTED.items()},
        "error_profile": {
            "empty_and_domain": "Value",
            "mode_without_repeated_value": "Value",
            "rank_value_absent": "Value",
            "explicit_missing_parameter": "Value",
        },
        "comparison_policy": {
            "selected_or_integer_max_ulps": ORDINARY_MAX_ULPS,
            "interpolation_or_decimal_rank_max_ulps": INTERPOLATION_MAX_ULPS,
            "zero": "selected zero is +0; computed underflow retains its IEEE sign bit",
            "rationale": "Selection and integer ranks are bit-invariant. Exact rational interpolation is converted once, while a bounded four-ULP allowance covers a production binary64 interpolation/decimal-rounding sequence without masking order-statistics errors.",
            "percent_rank_decimal_places": DECIMAL_PLACES,
        },
        "fixtures": fixtures,
        "observations": rows,
    }


def main():
    encoded = json.dumps(document(), indent=2) + "\n"
    path = HERE / "numeric-goldens.json"
    if "--check" in sys.argv:
        if path.read_text() != encoded:
            raise SystemExit("retained order-statistics oracle differs")
        print(json.dumps({"observations": 512, "fixtures": 56, "verified": True}))
    else:
        path.write_text(encoded)
        print("Wrote 512 observations over 56 fixtures")


if __name__ == "__main__":
    main()
