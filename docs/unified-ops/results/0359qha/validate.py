#!/usr/bin/env python3
"""Validate source-probe receipts without running the probe."""

import argparse
import json
import math
import pathlib
import sys
from typing import Any


def fail(message: str) -> int:
    print(f"0359qha validator: {message}", file=sys.stderr)
    return 1


def read_row(path: pathlib.Path) -> dict[str, Any] | None:
    try:
        rows = json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"unreadable JSON {path}: {error}")
        return None
    if not isinstance(rows, list) or len(rows) != 1:
        fail(f"receipt must contain exactly one scenario row: {path}")
        return None
    row = rows[0]
    if not isinstance(row, dict):
        fail(f"scenario row is not an object: {path}")
        return None
    return row


def validate_row(
    row: dict[str, Any],
    path: pathlib.Path,
    scenario: str,
    samples: int,
    output_sha: str,
) -> float | None:
    if row.get("scenario") != scenario:
        fail(f"scenario mismatch in {path}: {row.get('scenario')!r}")
        return None
    if row.get("samples") != samples:
        fail(f"sample-count mismatch in {path}: {row.get('samples')!r} != {samples}")
        return None
    if row.get("output_sha256") != output_sha:
        fail(f"output SHA-256 mismatch in {path}")
        return None
    values: list[float] = []
    for field in ("median_us", "min_us"):
        value = row.get(field)
        if not isinstance(value, (int, float)) or isinstance(value, bool) or not math.isfinite(value):
            fail(f"missing or invalid operation timing in {path}: {field}")
            return None
        if value < 0:
            fail(f"negative operation timing in {path}: {field}")
            return None
        values.append(float(value))
    if values[1] > values[0]:
        fail(f"minimum exceeds median in {path}")
        return None
    # The source probe sorts 31 samples and reports sample[15] as the median
    # and sample[0] as the minimum. Therefore at least 16 samples are >= the
    # median and the remaining 15 are >= the minimum. This is a conservative
    # certified lower bound on timed work, not the actual sample sum.
    return 16 * values[0] + 15 * values[1]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--json", type=pathlib.Path)
    source.add_argument("--directory", type=pathlib.Path)
    parser.add_argument("--scenario", required=True)
    parser.add_argument("--samples", required=True, type=int)
    parser.add_argument("--invocations", type=int)
    parser.add_argument("--output-sha", required=True)
    parser.add_argument("--min-aggregate-us", type=float, default=0.0)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.json is not None:
        if not args.json.is_file():
            return fail(f"missing receipt: {args.json}")
        row = read_row(args.json)
        if row is None:
            return 1
        if validate_row(row, args.json, args.scenario, args.samples, args.output_sha) is None:
            return 1
        return 0

    if args.invocations is None or args.invocations <= 0:
        return fail("--invocations must be positive for a directory")
    if not args.directory.is_dir():
        return fail(f"missing receipt directory: {args.directory}")
    expected_paths = [args.directory / f"result-{index}.json" for index in range(1, args.invocations + 1)]
    paths = list(args.directory.glob("result-*.json"))
    if {path.name for path in paths} != {path.name for path in expected_paths}:
        return fail(
            f"invocation receipts must be exactly result-1.json through "
            f"result-{args.invocations}.json"
        )
    paths = expected_paths
    if args.samples != 31:
        return fail("aggregate certification requires exactly 31 samples")
    certified_lower_bound_us = 0.0
    for path in paths:
        row = read_row(path)
        if row is None:
            return 1
        operation_lower_bound_us = validate_row(
            row, path, args.scenario, args.samples, args.output_sha
        )
        if operation_lower_bound_us is None:
            return 1
        certified_lower_bound_us += operation_lower_bound_us
    if not math.isfinite(certified_lower_bound_us) or certified_lower_bound_us < args.min_aggregate_us:
        return fail(
            f"subthreshold certified aggregate lower bound: {certified_lower_bound_us} < "
            f"{args.min_aggregate_us}"
        )
    print(
        f"0359qha validator: {args.scenario} invocations={args.invocations} "
        f"samples={args.samples} certified_lower_bound_us="
        f"{certified_lower_bound_us:.1f}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
