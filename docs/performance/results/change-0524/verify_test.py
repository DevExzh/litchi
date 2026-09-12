#!/usr/bin/env python3
"""Exercise fail-closed numerical checks with four in-memory mutations.

The retained reports are controls.  Every mutation is made on a deep copy,
and the script writes only its small result document; no capture artifact is
changed.  Use ``--stage`` to run the same checks against either matched stage.
"""

from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path

import analyze


def read(path: Path):
    return analyze.read(path)


def run(stage: str, output: Path) -> dict:
    analyze.configure(stage)
    folder = analyze.FOLDER
    cfb = read(folder / "native-r1-cfb.json")["results"][0]
    xls = next(
        row for row in read(folder / "native-r1-xls.json")["results"]
        if row["case"] == "xls_source_backed_open"
    )
    allocation = read(folder / "alloc-r1-cfb.json")["results"][0]

    # Establish that the unmodified controls are accepted before trying to
    # reject the semantic mutations.
    for row, count, allocator in (
        (cfb, 1000, False),
        (xls, 1000, False),
        (allocation, 30, True),
    ):
        analyze.validate_row(row, count, allocator)

    short = copy.deepcopy(cfb)
    short["elapsed_ns"]["samples"].pop()

    alignment = copy.deepcopy(cfb)
    alignment["operation_metrics"]["sample_indices"][0] = (
        alignment["operation_metrics"]["sample_indices"][1]
    )

    source_contract = copy.deepcopy(xls)
    source_contract["source"]["read_calls"][0] += 1

    balance = copy.deepcopy(allocation)
    balance["operation_metrics"]["allocation"]["allocated_bytes"]["values"][0] += 1

    mutations = [
        ("short_native_vector", short, 1000, False),
        ("operation_sample_alignment", alignment, 1000, False),
        ("source_read_contract", source_contract, 1000, False),
        ("allocation_live_byte_balance", balance, 30, True),
    ]
    checks = []
    for name, row, count, allocator in mutations:
        try:
            analyze.validate_row(row, count, allocator)
        except (AssertionError, KeyError, TypeError, ValueError) as error:
            checks.append({
                "case": name,
                "rejected": True,
                "reason": str(error) or "semantic invariant assertion",
            })
        else:
            raise AssertionError(f"corruption accepted: {name}")

    result = {
        "schema": "cfb_ole2_numeric_verifier_tests_v1",
        "status": "pass",
        "stage": stage,
        "valid_controls_pass": True,
        "checks": checks,
        "scope": "Four in-memory raw semantic mutations; no retained artifact modified",
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("baseline", "candidate"), default="baseline")
    parser.add_argument("output", nargs="?", type=Path)
    parser.add_argument("--output", dest="output_option", type=Path)
    args = parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error("provide output either positionally or with --output")
    output = args.output_option or args.output or analyze.HERE / "verifier-tests.json"
    try:
        result = run(args.stage, output)
    except (analyze.EvidenceError, AssertionError, KeyError, OSError,
            TypeError, ValueError) as error:
        print(f"verify_test.py: evidence check failed: {error}")
        return 2
    print(f"0524 {result['stage']} verifier mutations rejected: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
