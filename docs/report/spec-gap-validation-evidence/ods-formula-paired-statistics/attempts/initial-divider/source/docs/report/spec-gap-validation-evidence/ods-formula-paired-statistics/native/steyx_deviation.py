#!/usr/bin/env python3
"""Prove the retained LibreOffice STEYX host-rounding deviation.

The helper reads only the retained literal closure.  It decodes each decimal
cell to binary64, forms the centered regression quantities as exact
``Fraction`` values, and applies a 300-digit ``Decimal`` square root only to
the final residual mean square.  It never imports or calls the Rust evaluator.
"""

from __future__ import annotations

from decimal import Context, Decimal, localcontext
from fractions import Fraction
import argparse
import hashlib
import json
import math
from pathlib import Path
import struct
import sys

sys.dont_write_bytecode = True


ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = ROOT / "cached-results.json"
DEFAULT_OUTPUT = ROOT / "steyx-deviation-proof.json"
DECIMAL_PRECISION = 300
TARGET_FUNCTION = "STEYX"
TARGET_ROW = 3
TARGET_COLUMN = 1
Y_COLUMN = 9
X_COLUMN = 10


def binary64(value: str) -> float:
    result = float(value)
    if not math.isfinite(result):
        raise ValueError(f"non-finite binary64 source value: {value!r}")
    return result


def bits(value: float) -> str:
    return struct.pack(">d", value).hex()


def bits_integer(value: float) -> int:
    return int.from_bytes(struct.pack(">d", value), "big")


def target_row(input_path: Path) -> dict:
    rows = json.loads(input_path.read_text())
    matches = [
        row
        for row in rows
        if row.get("function") == TARGET_FUNCTION
        and row.get("row") == TARGET_ROW
        and row.get("column") == TARGET_COLUMN
    ]
    if len(matches) != 1:
        raise ValueError(f"expected one retained {TARGET_FUNCTION} target row, found {len(matches)}")
    row = matches[0]
    if row["formula"] != "of:=STEYX([.I2:.I101];[.J2:.J101])":
        raise ValueError(f"unexpected STEYX target formula: {row['formula']!r}")
    return row


def calculate(input_path: Path) -> dict:
    row = target_row(input_path)
    cells = row["cells"]
    y_cells = {cell["row"]: cell for cell in cells if cell["column"] == Y_COLUMN}
    x_cells = {cell["row"]: cell for cell in cells if cell["column"] == X_COLUMN}
    if set(y_cells) != set(x_cells):
        raise ValueError("STEYX target columns do not retain aligned rows")

    pairs = []
    skipped_rows = []
    for row_number in sorted(y_cells):
        y_cell = y_cells[row_number]
        x_cell = x_cells[row_number]
        if y_cell["type"] == "number" and x_cell["type"] == "number":
            # Fraction(float(...)) records the exact represented binary64
            # operand, rather than treating the source decimal as exact.
            x_value = binary64(x_cell["value"])
            y_value = binary64(y_cell["value"])
            pairs.append((Fraction(x_value), Fraction(y_value)))
        elif y_cell["type"] == "empty" and x_cell["type"] == "empty":
            skipped_rows.append(row_number)
        else:
            raise ValueError(
                f"unexpected nonnumeric asymmetry at retained row {row_number}: "
                f"{y_cell['type']}/{x_cell['type']}"
            )

    pair_count = len(pairs)
    if pair_count != 99 or skipped_rows != [2]:
        raise ValueError(f"unexpected target admission: pairs={pair_count}, skipped={skipped_rows}")
    x_values = [x for x, _ in pairs]
    y_values = [y for _, y in pairs]
    x_mean = sum(x_values, Fraction()) / pair_count
    y_mean = sum(y_values, Fraction()) / pair_count
    sxx = sum((x - x_mean) ** 2 for x in x_values)
    sxy = sum((x - x_mean) * (y - y_mean) for x, y in pairs)
    slope = sxy / sxx
    intercept = y_mean - slope * x_mean
    residual_sum = sum(
        (y - (intercept + slope * x)) ** 2
        for x, y in pairs
    )
    residual_mean_square = residual_sum / (pair_count - 2)
    with localcontext(Context(prec=DECIMAL_PRECISION)):
        decimal_mean_square = (
            Decimal(residual_mean_square.numerator)
            / Decimal(residual_mean_square.denominator)
        )
        decimal_root = decimal_mean_square.sqrt()
        reference = float(decimal_root)

    native = binary64(row["cached"])
    reference_bits = bits_integer(reference)
    native_bits = bits_integer(native)
    ulp_difference = reference_bits - native_bits
    absolute_difference = reference - native
    if (
        reference.hex() != "0x1.9a26885914f04p+0"
        or native.hex() != "0x1.9a2688591472fp+0"
        or ulp_difference != 2005
    ):
        raise ValueError(
            "retained host deviation changed: "
            f"reference={reference.hex()}, native={native.hex()}, ulps={ulp_difference}"
        )

    return {
        "oracle": "Exact Fraction over decoded binary64 operands; Decimal square root at 300 digits",
        "input": input_path.name,
        "input_sha256": hashlib.sha256(input_path.read_bytes()).hexdigest(),
        "function": TARGET_FUNCTION,
        "formula": row["formula"],
        "source": row["source"],
        "sheet": row["sheet"],
        "row": TARGET_ROW,
        "column": TARGET_COLUMN,
        "y_column": Y_COLUMN,
        "x_column": X_COLUMN,
        "pair_count": pair_count,
        "skipped_rows": skipped_rows,
        "decimal_precision": DECIMAL_PRECISION,
        "residual_mean_square_fraction": {
            "numerator": str(residual_mean_square.numerator),
            "denominator": str(residual_mean_square.denominator),
        },
        "reference_decimal_300": format(decimal_root, "f"),
        "reference_float": repr(reference),
        "reference_hex": reference.hex(),
        "reference_bits": bits(reference),
        "native_cache": row["cached"],
        "native_hex": native.hex(),
        "native_bits": bits(native),
        "absolute_difference": repr(absolute_difference),
        "native_ulp": repr(math.ulp(native)),
        "ulp_difference": ulp_difference,
        "comparison": "The Rust native test compares this row to reference_bits and asserts the retained native cache is exactly 2005 binary64 ULP below it; the other 41 rows retain the ordinary 1e-13 relative policy.",
    }


def write_proof(input_path: Path, output_path: Path) -> None:
    proof = calculate(input_path)
    output_path.write_text(json.dumps(proof, indent=2) + "\n")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    proof = calculate(args.input)
    if args.check:
        retained = json.loads(args.output.read_text())
        if retained != proof:
            raise SystemExit("retained STEYX proof differs from independent calculation")
    else:
        args.output.write_text(json.dumps(proof, indent=2) + "\n")
    print(json.dumps({"function": TARGET_FUNCTION, "ulp_difference": proof["ulp_difference"], "checked": args.check}, sort_keys=True))


if __name__ == "__main__":
    main()
