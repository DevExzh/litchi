#!/usr/bin/env python3
"""Validate one source-probe JSON receipt without running the probe."""

import argparse
import json
import math
import pathlib
import sys


def fail(message: str) -> int:
    print(f"0359qha validator: {message}", file=sys.stderr)
    return 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, type=pathlib.Path)
    parser.add_argument("--scenario", required=True)
    parser.add_argument("--samples", required=True, type=int)
    parser.add_argument("--output-sha", required=True)
    parser.add_argument("--min-operation-us", type=float, default=0.0)
    args = parser.parse_args()
    if not args.json.is_file():
        return fail(f"missing receipt: {args.json}")
    try:
        rows = json.loads(args.json.read_text())
    except (OSError, ValueError) as error:
        return fail(f"unreadable JSON: {error}")
    if not isinstance(rows, list) or len(rows) != 1:
        return fail("receipt must contain exactly one scenario row")
    row = rows[0]
    if not isinstance(row, dict):
        return fail("scenario row is not an object")
    if row.get("scenario") != args.scenario:
        return fail(f"scenario mismatch: {row.get('scenario')!r}")
    if row.get("samples") != args.samples:
        return fail(f"sample-count mismatch: {row.get('samples')!r} != {args.samples}")
    if row.get("output_sha256") != args.output_sha:
        return fail("output SHA-256 mismatch")
    for field in ("median_us", "min_us"):
        value = row.get(field)
        if not isinstance(value, (int, float)) or not math.isfinite(value):
            return fail(f"missing or invalid operation timing: {field}")
        if value < args.min_operation_us:
            return fail(
                f"subthreshold operation timing: {field}={value} < {args.min_operation_us}"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
