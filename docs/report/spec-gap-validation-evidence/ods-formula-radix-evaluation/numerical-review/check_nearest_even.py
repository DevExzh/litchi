#!/usr/bin/env python3
"""Independently check the fixed-limb integer-to-f64 rounding algorithm.

This mirrors the arithmetic and bit decisions in radix.rs without importing,
building, or executing the Rust implementation.  Python's integer-to-float
conversion supplies the independent IEEE-754 nearest-even oracle for results
that remain finite; overflow is represented as ``None`` on both sides.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
import struct
from pathlib import Path


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[4]
RADIX = ROOT / "crates/litchi-ods/src/codec/formula/evaluation/radix.rs"
TEST = ROOT / "crates/litchi-ods/tests/ods_formula_radix_evaluation.rs"
CHECK_COUNT = 19_251


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def float_from_bits(bits: int) -> float:
    return struct.unpack(">d", bits.to_bytes(8, "big"))[0]


def fixed_limb_algorithm(value: int) -> float | None:
    if value == 0:
        return 0.0
    highest = value.bit_length() - 1
    if highest <= 52:
        return float(value)

    shift = highest - 52
    significand = (value >> shift) & ((1 << 53) - 1)
    round_bit = (value >> (shift - 1)) & 1
    sticky = bool(value & ((1 << (shift - 1)) - 1)) if shift > 1 else False
    if round_bit and (sticky or (significand & 1)):
        significand += 1

    exponent = highest
    if significand == 1 << 53:
        significand >>= 1
        exponent += 1
    if exponent > 1023:
        return None
    bits = ((exponent + 1023) << 52) | (significand & ((1 << 52) - 1))
    return float_from_bits(bits)


def python_oracle(value: int) -> float | None:
    try:
        result = float(value)
    except OverflowError:
        return None
    return result if math.isfinite(result) else None


def check(value: int) -> None:
    actual = fixed_limb_algorithm(value)
    expected = python_oracle(value)
    if actual is None or expected is None:
        if actual is not None or expected is not None:
            raise AssertionError(f"overflow mismatch for {value}")
    elif actual.hex() != expected.hex():
        raise AssertionError(
            f"rounding mismatch for {value}: {actual.hex()} != {expected.hex()}"
        )


def run() -> int:
    rng = random.Random(0x4F44534652414449)
    checks = 0
    for highest in range(1024):
        samples = 100 if highest < 100 else 10
        for _ in range(samples):
            value = rng.getrandbits(highest + 1) | (1 << highest)
            check(value)
            checks += 1

    max_finite_integer = ((1 << 53) - 1) << 971
    for delta in (
        -(1 << 972),
        -(1 << 971),
        -(1 << 970) - 1,
        -(1 << 970),
        -(1 << 970) + 1,
        0,
        (1 << 970) - 1,
        1 << 970,
        (1 << 970) + 1,
        (1 << 971) - 1,
        1 << 971,
    ):
        check(max_finite_integer + delta)
        checks += 1

    if checks != CHECK_COUNT:
        raise AssertionError(f"unexpected check count {checks}")
    return checks


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--receipt", type=Path)
    args = parser.parse_args()
    checks = run()
    receipt = {
        "status": 0,
        "claim": "fixed-limb nearest-even formula agrees with an independent Python IEEE-754 oracle",
        "algorithm_scope": "mirrored radix.rs bit decisions; Rust binary was not executed",
        "checks": checks,
        "boundary_cases": 11,
        "random_seed": "0x4F44534652414449",
        "radix_source_sha256": sha256(RADIX),
        "focused_test_sha256": sha256(TEST),
        "script_sha256": sha256(Path(__file__).resolve()),
        "command": "python3 check_nearest_even.py --receipt receipt.json",
    }
    rendered = json.dumps(receipt, indent=2) + "\n"
    if args.receipt:
        args.receipt.write_text(rendered, encoding="utf-8")
    print(rendered, end="")


if __name__ == "__main__":
    main()
