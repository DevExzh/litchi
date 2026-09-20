#!/usr/bin/env python3
"""Independent exact references for paired statistics and regression.

The retained operands are represented IEEE-754 binary64 values.  All
additive, centered, covariance, regression, and RSQ quantities are formed as
exact :class:`fractions.Fraction` values.  Correlation and STEYX use a
high-precision Decimal square root only for their final publication.  The
generator never imports or calls the Rust evaluator.

The corpus deliberately keeps ``INTERCEPT`` in the same coherent included-
constant profile as the contract: ``y_mean - slope*x_mean``.  Every paired
data argument is shape-aware.  An aligned pair is admitted only when both
members are finite Numbers; Empty, Text, and Logical members suppress that
position, while formula Errors are retained as formula errors.
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
    "CORREL",
    "COVAR",
    "PEARSON",
    "RSQ",
    "SLOPE",
    "INTERCEPT",
    "STEYX",
    "FORECAST",
)
CORRELATION_FUNCTIONS = {"CORREL", "PEARSON", "RSQ"}
ORDINARY_MAX_ULPS = 0
SENSITIVE_MAX_ULPS = 8
TRANSCENDENTAL_MAX_ULPS = 8
DECIMAL_PRECISION = 300
SEED = 20260919

ERROR_VALUE = "Value"
ERROR_DIVISION = "DivisionByZero"
ERROR_NUMBER = "Number"
ERROR_NA = "NotAvailable"

SENSITIVE_FIXTURES = {
    "adjacent-large",
    "adjacent-small",
    "cancellation",
    "extreme-finite",
    "subnormal",
    "random-0",
    "random-1",
    "random-2",
    "repeated-line",
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


def represented_number(cell: dict[str, str]) -> Fraction | None:
    value = struct.unpack(">d", bytes.fromhex(cell["bits"]))[0]
    if not math.isfinite(value):
        return None
    return Fraction(value)


def add_dataset(
    datasets,
    name: str,
    x,
    y,
    *,
    shape: tuple[int, int] | None = None,
    y_shape: tuple[int, int] | None = None,
    query: dict | None = None,
    reference_list: bool = False,
):
    if shape is None:
        shape = (len(x), 1)
    if y_shape is None:
        y_shape = shape
    datasets.append(
        {
            "name": name,
            "x": x,
            "y": y,
            "x_shape": shape,
            "y_shape": y_shape,
            "query": query if query is not None else number(2.0),
            "reference_list": reference_list,
        }
    )


def sys_float_max() -> float:
    return float.fromhex("0x1.fffffffffffffp+1023")


def build_datasets():
    tiny = math.ulp(0.0)
    maximum = sys_float_max()
    datasets = []
    add = lambda name, x, y, **kwargs: add_dataset(datasets, name, x, y, **kwargs)

    add("line-basic", [number(x) for x in (0.0, 1.0, 2.0, 3.0)], [number(y) for y in (1.0, 3.0, 5.0, 7.0)], query=number(4.0))
    add("line-negative", [number(x) for x in (-2.0, -1.0, 0.0, 1.0)], [number(y) for y in (5.0, 3.0, 1.0, -1.0)], query=number(2.0))
    add("line-fractional", [number(x) for x in (0.1, 0.2, 0.3, 0.4)], [number(x * 1.5 - 0.2) for x in (0.1, 0.2, 0.3, 0.4)], query=number(0.5))
    add("near-perfect", [number(x) for x in (1.0, 2.0, 3.0, 4.0, 5.0)], [number(x) for x in (2.0, 4.0, 6.0, 8.0, math.nextafter(10.0, math.inf))], query=number(6.0))
    add("constant-x", [number(2.0)] * 4, [number(x) for x in (1.0, 3.0, 5.0, 7.0)], query=number(4.0))
    add("constant-y", [number(x) for x in (1.0, 2.0, 3.0, 4.0)], [number(7.0)] * 4, query=number(5.0))
    add("one-pair", [empty(), number(1.0), text("ignored")], [number(2.0), number(2.0), number(3.0)], query=number(2.0))
    add("two-pair", [number(0.0), number(1.0)], [number(1.0), number(3.0)], query=number(2.0))
    add("empty-pairs", [empty(), text("ignored")], [text("ignored"), empty()], query=number(2.0))

    add("typed-asymmetric-x", [number(1.0), text("ignored"), number(3.0), empty(), number(5.0)], [number(2.0), number(4.0), number(6.0), number(8.0), number(10.0)], query=number(4.0))
    add("typed-asymmetric-y", [number(1.0), number(2.0), number(3.0), number(4.0), number(5.0)], [number(2.0), text("ignored"), number(6.0), empty(), number(10.0)], query=number(4.0))
    add("typed-logical", [number(1.0), number(2.0), number(3.0), number(4.0)], [number(2.0), logical(True), number(6.0), logical(False)], query=number(3.0))
    add("formula-error-x", [number(1.0), error("NotAvailable"), number(3.0)], [number(2.0), number(4.0), number(6.0)], query=number(2.0))
    add("formula-error-y", [number(1.0), number(2.0), number(3.0)], [number(2.0), error("DivisionByZero"), number(6.0)], query=number(2.0))

    add("adjacent-large", [number(x) for x in (1.0e16, math.nextafter(1.0e16, math.inf), math.nextafter(1.0e16, 0.0), 1.0e16)], [number(x) for x in (2.0e16, math.nextafter(2.0e16, math.inf), math.nextafter(2.0e16, 0.0), 2.0e16)], query=number(1.0e16))
    add("adjacent-small", [number(x) for x in (1.0e-150, math.nextafter(1.0e-150, math.inf), math.nextafter(1.0e-150, 0.0), 1.0e-150)], [number(x) for x in (2.0e-150, math.nextafter(2.0e-150, math.inf), math.nextafter(2.0e-150, 0.0), 2.0e-150)], query=number(1.0e-150))
    add("subnormal", [number(x) for x in (tiny, 2.0 * tiny, 3.0 * tiny, 4.0 * tiny)], [number(x) for x in (2.0 * tiny, 4.0 * tiny, 6.0 * tiny, 8.0 * tiny)], query=number(tiny))
    add("extreme-finite", [number(x) for x in (maximum, -maximum, 0.0, 1.0)], [number(x) for x in (maximum, -maximum, 1.0, 0.0)], query=number(0.0))
    add("cancellation", [number(x) for x in (1.0e100, 1.0, -1.0e100, 2.0)], [number(x) for x in (-1.0e100, 2.0, 1.0e100, 3.0)], query=number(3.0))
    add("mixed-sign", [number(x) for x in (-4.0, -2.0, 1.0, 3.0)], [number(x) for x in (8.0, 4.0, -2.0, -6.0)], query=number(1.0))
    add("underflow-slope", [number(x) for x in (0.0, 1.0, 2.0, 3.0)], [number(x) for x in (-tiny, 0.0, tiny, 2.0 * tiny)], query=number(4.0))

    add("shape-row-vs-column", [number(x) for x in (1.0, 2.0, 3.0)], [number(x) for x in (1.0, 2.0, 3.0)], shape=(1, 3), y_shape=(3, 1))
    add("shape-length-mismatch", [number(x) for x in (1.0, 2.0, 3.0)], [number(x) for x in (1.0, 2.0)], shape=(3, 1), y_shape=(2, 1))
    add("list-refusal", [number(x) for x in (1.0, 2.0, 3.0)], [number(x) for x in (1.0, 2.0, 3.0)], reference_list=True)

    repeated_x = [number(float(i % 4)) for i in range(256)]
    repeated_y = [number(2.0 * float(i % 4) + 1.0) for i in range(256)]
    add("repeated-line", repeated_x, repeated_y, query=number(5.0))

    rng = random.Random(SEED)
    for index in range(3):
        x_values = []
        y_values = []
        for _ in range(5 + index * 2):
            x_value = math.ldexp((rng.getrandbits(53) + 1) / 2**53, rng.randrange(-500, 501))
            y_value = math.ldexp((rng.getrandbits(53) + 1) / 2**53, rng.randrange(-500, 501))
            x_values.append(number(rng.choice((-1.0, 1.0)) * x_value))
            y_values.append(number(rng.choice((-1.0, 1.0)) * y_value))
        add(f"random-{index}", x_values, y_values, query=number(float(index + 1)))

    add("query-logical", [number(x) for x in (0.0, 1.0, 2.0)], [number(x) for x in (1.0, 3.0, 5.0)], query=logical(True))
    add("query-text", [number(x) for x in (0.0, 1.0, 2.0)], [number(x) for x in (1.0, 3.0, 5.0)], query=text("4"))
    add("query-bad", [number(x) for x in (0.0, 1.0, 2.0)], [number(x) for x in (1.0, 3.0, 5.0)], query=text("bad"))
    add("query-bad-data-error", [number(0.0), error("NotAvailable"), number(2.0)], [number(1.0), number(3.0), number(5.0)], query=text("bad"))
    add("query-nonfinite", [number(x) for x in (0.0, 1.0, 2.0)], [number(x) for x in (1.0, 3.0, 5.0)], query=number(float("inf")))
    add("query-empty", [number(x) for x in (0.0, 1.0, 2.0)], [number(x) for x in (1.0, 3.0, 5.0)], query=empty())
    add("scalar-singleton", [number(1.0)], [number(3.0)], query=number(2.0))
    add("scalar-text", [text("2")], [number(3.0)], query=number(2.0))
    add("scalar-logical", [logical(True)], [number(3.0)], query=number(2.0))
    add("scalar-error", [error("NotAvailable")], [number(3.0)], query=number(2.0))
    add("nonfinite-number", [number(1.0), number(float("inf")), number(3.0)], [number(2.0), number(4.0), number(6.0)], query=number(2.0))

    assert len(datasets) == 39, len(datasets)
    return datasets


def col_name(index: int) -> str:
    result = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def cell_ref(row: int, column: int) -> str:
    return f"{col_name(column)}{row + 1}"


def range_ref(rows: int, columns: int, base_column: int) -> str:
    return f"[.{cell_ref(0, base_column)}:.{cell_ref(rows - 1, base_column + columns - 1)}]"


def scalar_token(cell: dict) -> str:
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
        return {
            ERROR_VALUE: "#VALUE!",
            ERROR_NUMBER: "#NUM!",
            ERROR_DIVISION: "#DIV/0!",
            ERROR_NA: "#N/A",
        }[cell["value"]]
    raise AssertionError(kind)


def array_literal(cells: list[dict], shape: tuple[int, int]) -> str:
    rows, columns = shape
    values = []
    for row in range(rows):
        start = row * columns
        values.append("|".join(scalar_token(cell) for cell in cells[start : start + columns]))
    return "{" + ";".join(values) + "}"


def expression_for(function: str, dataset: dict, origin: str) -> str:
    x = dataset["x"]
    y = dataset["y"]
    x_shape = dataset["x_shape"]
    y_shape = dataset["y_shape"]
    if origin == "ReferenceList":
        x_expr = f"({range_ref(*x_shape, 0)}~{range_ref(*x_shape, 0)})"
        y_expr = f"({range_ref(*y_shape, 4)}~{range_ref(*y_shape, 4)})"
    elif origin == "Reference":
        x_expr = range_ref(*x_shape, 0)
        y_expr = range_ref(*y_shape, 4)
    elif origin == "Array":
        x_expr = array_literal(x, x_shape)
        y_expr = array_literal(y, y_shape)
    elif origin == "Scalar":
        x_expr = scalar_token(x[0])
        y_expr = scalar_token(y[0])
    else:
        raise AssertionError(origin)
    if function in ("RSQ", "SLOPE", "INTERCEPT", "STEYX", "FORECAST"):
        first_data, second_data = y_expr, x_expr
    else:
        first_data, second_data = x_expr, y_expr
    if function == "FORECAST":
        query_expr = (
            "[.H1]"
            if origin == "Reference" and dataset["name"] in ("query-empty", "query-nonfinite")
            else scalar_token(dataset["query"])
        )
        return f"=FORECAST({query_expr};{first_data};{second_data})"
    return f"={function}({first_data};{second_data})"


def pair_values(function: str, dataset: dict):
    if dataset["x_shape"] != dataset["y_shape"]:
        return [], None, ERROR_VALUE, "shape"
    formula_error = None
    generated_error = None
    values = []
    for left, right in zip(dataset["x"], dataset["y"]):
        source_members = (right, left) if function in ("RSQ", "SLOPE", "INTERCEPT", "STEYX", "FORECAST") else (left, right)
        for cell in source_members:
            if cell["kind"] == "Error" and formula_error is None:
                formula_error = cell["value"]
        if left["kind"] == "Error" or right["kind"] == "Error":
            continue
        if left["kind"] != "Number" or right["kind"] != "Number":
            continue
        x = represented_number(left)
        y = represented_number(right)
        if x is None or y is None:
            if generated_error is None:
                generated_error = ERROR_NUMBER
            continue
        values.append((x, y))
    return values, formula_error, generated_error, "pairs"


def query_fraction(cell: dict):
    kind = cell["kind"]
    if kind == "Number":
        value = represented_number(cell)
        return value, None if value is not None else ERROR_NUMBER
    if kind == "Logical":
        return Fraction(int(cell["value"])), None
    if kind == "Empty":
        return Fraction(0), None
    if kind == "Text":
        try:
            value = float(cell["value"])
        except ValueError:
            return None, ERROR_VALUE
        if not math.isfinite(value):
            return None, ERROR_NUMBER
        return Fraction(value), None
    if kind == "Error":
        return None, cell["value"]
    return None, ERROR_VALUE


def centered(values):
    count = len(values)
    x_mean = sum((x for x, _ in values), Fraction(0)) / count
    y_mean = sum((y for _, y in values), Fraction(0)) / count
    sxx = sum(((x - x_mean) ** 2 for x, _ in values), Fraction(0))
    syy = sum(((y - y_mean) ** 2 for _, y in values), Fraction(0))
    sxy = sum(((x - x_mean) * (y - y_mean) for x, y in values), Fraction(0))
    return x_mean, y_mean, sxx, syy, sxy


def decimal_sqrt_fraction(value: Fraction):
    context = Context(prec=DECIMAL_PRECISION)
    with localcontext(context):
        decimal = (Decimal(value.numerator) / Decimal(value.denominator)).sqrt()
        return decimal, format(decimal, "f")


def finite_fraction_result(value: Fraction):
    try:
        result = float(value)
    except OverflowError:
        return None
    return result if math.isfinite(result) else None


def result_from_fraction(value: Fraction, fixture: str):
    rounded = finite_fraction_result(value)
    if rounded is None:
        return {"error": ERROR_NUMBER, "error_source": "binary64_overflow"}
    comparison = "exact_bits" if value == 0 or fixture not in SENSITIVE_FIXTURES else "ulps"
    return {
        "expected_bits": bits(rounded, canonical_zero=value == 0),
        "exact_numerator": str(value.numerator),
        "exact_denominator": str(value.denominator),
        "comparison": comparison,
        "max_ulps": 0 if comparison == "exact_bits" else SENSITIVE_MAX_ULPS,
    }


def result_from_decimal(value: Decimal, fixture: str):
    rounded = float(value)
    if not math.isfinite(rounded):
        return {"error": ERROR_NUMBER, "error_source": "binary64_overflow"}
    comparison = "exact_bits" if value == 0 or abs(rounded) == 1.0 else "ulps"
    return {
        "expected_bits": bits(rounded, canonical_zero=value == 0),
        "decimal_reference": format(value, "f"),
        "decimal_precision": DECIMAL_PRECISION,
        "comparison": comparison,
        "max_ulps": 0 if comparison == "exact_bits" else TRANSCENDENTAL_MAX_ULPS,
    }


def exact_result(function: str, dataset: dict, origin: str):
    if origin == "Scalar":
        dataset = {
            **dataset,
            "x": dataset["x"][:1],
            "y": dataset["y"][:1],
            "x_shape": (1, 1),
            "y_shape": (1, 1),
        }
    values, formula_error, generated_error, status = pair_values(function, dataset)
    if status == "shape":
        if function == "RSQ":
            left_cells = dataset["x_shape"][0] * dataset["x_shape"][1]
            right_cells = dataset["y_shape"][0] * dataset["y_shape"][1]
            if left_cells != right_cells:
                return {"error": ERROR_NA, "error_source": "shape_cell_count"}
        return {"error": ERROR_VALUE, "error_source": "shape"}
    query = None
    query_error = None
    if function == "FORECAST":
        query, query_error = query_fraction(dataset["query"])
        query_formula_error = (
            query_error if dataset["query"]["kind"] == "Error" else None
        )
    else:
        query_formula_error = None
    if formula_error is not None:
        if query_formula_error is not None:
            return {
                "error": query_formula_error,
                "error_source": "query_formula",
            }
        return {"error": formula_error, "error_source": "formula"}
    if query_error is not None:
        return {
            "error": query_error,
            "error_source": "query_formula" if dataset["query"]["kind"] == "Error" else "query",
        }
    if generated_error is not None:
        return {"error": generated_error, "error_source": "generated"}
    if not values:
        return {"error": ERROR_NA if function == "RSQ" else ERROR_VALUE, "error_source": "empty_pairs"}
    x_mean, y_mean, sxx, syy, sxy = centered(values)
    count = len(values)
    if function in CORRELATION_FUNCTIONS:
        if sxx == 0 or syy == 0:
            return {"error": ERROR_DIVISION, "error_source": "zero_variance"}
        if function == "RSQ":
            return result_from_fraction(sxy * sxy / (sxx * syy), dataset["name"])
        decimal, audit = decimal_sqrt_fraction(sxx * syy)
        with localcontext(Context(prec=DECIMAL_PRECISION)):
            value = Decimal(sxy.numerator) / Decimal(sxy.denominator) / decimal
        result = result_from_decimal(value, dataset["name"])
        result["decimal_reference"] = audit
        return result
    if function == "COVAR":
        return result_from_fraction(sxy / count, dataset["name"])
    if function == "STEYX" and count < 3:
        return {"error": ERROR_VALUE, "error_source": "minimum_count"}
    if sxx == 0:
        return {"error": ERROR_DIVISION, "error_source": "zero_x_variance"}
    slope = sxy / sxx
    if function == "SLOPE":
        return result_from_fraction(slope, dataset["name"])
    if function == "INTERCEPT":
        return result_from_fraction(y_mean - slope * x_mean, dataset["name"])
    if function == "FORECAST":
        return result_from_fraction(y_mean + slope * (query - x_mean), dataset["name"])
    if function == "STEYX":
        residual = syy - sxy * sxy / sxx
        if residual < 0:
            return {"error": ERROR_NUMBER, "error_source": "negative_residual"}
        if residual == 0:
            return result_from_fraction(Fraction(0), dataset["name"])
        decimal, audit = decimal_sqrt_fraction(residual / (count - 2))
        result = result_from_decimal(decimal, dataset["name"])
        result["decimal_reference"] = audit
        return result
    raise AssertionError(function)


def encode_fixture(dataset: dict):
    return {
        "encoding": "paired_cells",
        "x": dataset["x"],
        "y": dataset["y"],
        "x_shape": list(dataset["x_shape"]),
        "y_shape": list(dataset["y_shape"]),
        "query": dataset["query"],
        "reference_list": dataset["reference_list"],
    }


def row_for(function: str, dataset: dict, origin: str):
    result = {
        "function": function,
        "fixture": dataset["name"],
        "origin": origin,
        "formula": expression_for(function, dataset, origin),
    }
    shape_valid = dataset["x_shape"] == dataset["y_shape"]
    if origin == "ReferenceList" or not shape_valid:
        result["expected_reads"] = 0
    elif origin == "Reference":
        result["expected_reads"] = len(dataset["x"]) + len(dataset["y"])
        if function == "FORECAST" and dataset["name"] in ("query-empty", "query-nonfinite"):
            result["expected_reads"] += 1
    else:
        result["expected_reads"] = 0
    if origin == "ReferenceList":
        result.update({"error": ERROR_VALUE, "error_source": "shape"})
        return result
    result.update(exact_result(function, dataset, origin))
    return result


def document():
    datasets = build_datasets()
    by_name = {dataset["name"]: dataset for dataset in datasets}
    fixtures = {dataset["name"]: encode_fixture(dataset) for dataset in datasets}
    rows = []
    for function in FUNCTIONS:
        for dataset in datasets:
            rows.append(row_for(function, dataset, "Reference"))

    array_names = (
        "line-basic",
        "line-fractional",
        "typed-asymmetric-x",
        "typed-asymmetric-y",
        "typed-logical",
        "formula-error-x",
        "formula-error-y",
        "adjacent-large",
        "subnormal",
        "constant-x",
        "constant-y",
        "shape-row-vs-column",
        "shape-length-mismatch",
    )
    for function in FUNCTIONS:
        for name in array_names:
            rows.append(row_for(function, by_name[name], "Array"))

    scalar_names = ("one-pair", "scalar-singleton", "scalar-text", "scalar-logical", "scalar-error")
    for function in FUNCTIONS:
        for name in scalar_names:
            rows.append(row_for(function, by_name[name], "Scalar"))

    list_dataset = by_name["list-refusal"]
    for function in FUNCTIONS:
        rows.append(row_for(function, list_dataset, "ReferenceList"))

    expected_rows = len(FUNCTIONS) * (len(datasets) + len(array_names) + len(scalar_names) + 1)
    assert len(rows) == expected_rows, (len(rows), expected_rows)
    return {
        "oracle": "Exact Fraction over represented binary64 pairs; Decimal square root at 300 digits",
        "production_rounding_claim": False,
        "seed": SEED,
        "functions": list(FUNCTIONS),
        "intercept_profile": "included_constant_y_mean_minus_slope_x_mean",
        "pair_admission": "finite Numbers on both aligned sides; Empty/Text/Logical skip the position; formula Errors propagate",
        "comparison_policy": {
            "ordinary_max_ulps": ORDINARY_MAX_ULPS,
            "sensitive_max_ulps": SENSITIVE_MAX_ULPS,
            "transcendental_max_ulps": TRANSCENDENTAL_MAX_ULPS,
            "exact_zero": "exact positive zero for algebraic nonnegative or zero results; signed nonzero underflow retains sign",
            "rationale": "Exact Fraction rows use exact bits for ordinary rational fixtures and a predeclared eight-ULP ceiling for cancellation/extreme fixtures. Decimal square-root rows use the same eight-ULP ceiling except exact zero and exact unit correlation.",
            "decimal_precision": DECIMAL_PRECISION,
        },
        "error_profile": {
            "no_pair": ERROR_VALUE,
            "rsq_no_pair": ERROR_NA,
            "steyx_minimum": ERROR_VALUE,
            "zero_variance": ERROR_DIVISION,
            "shape_or_list": ERROR_VALUE,
        },
        "fixtures": fixtures,
        "observations": rows,
    }


def main():
    encoded = json.dumps(document(), indent=2) + "\n"
    path = HERE / "numeric-goldens.json"
    if "--check" in sys.argv:
        if path.read_text() != encoded:
            raise SystemExit("retained paired-statistics oracle differs")
        document_value = json.loads(encoded)
        print(json.dumps({"observations": len(document_value["observations"]), "fixtures": len(document_value["fixtures"]), "verified": True}))
    else:
        path.write_text(encoded)
        document_value = json.loads(encoded)
        print(f"Wrote {len(document_value['observations'])} observations over {len(document_value['fixtures'])} paired fixtures")


if __name__ == "__main__":
    main()
