#!/usr/bin/env python3
"""Run the bounded ODS formula reference parser profile and emit normalized CSV."""

from __future__ import annotations

import argparse
import csv
import json
import re
import shlex
import subprocess
from pathlib import Path


COMPARABLE_CASES = (
    ("parse-bracket-current-cell", 1_000),
    ("parse-bracket-sheet-cell", 1_000),
    ("parse-bracket-range", 1_000),
    ("parse-quoted-doubled-sheet", 1_000),
    ("parse-local-refs-256", 128),
    ("parse-sum-unbracketed", 1_000),
    ("parse-vlookup-unbracketed", 1_000),
    ("parse-malformed-zero-row", 1_000),
    ("parse-malformed-missing-separator", 1_000),
    ("parse-malformed-unclosed-bracket", 1_000),
    ("parse-malformed-missing-row", 1_000),
    ("parse-malformed-bad-quote", 1_000),
)

COVERAGE_CASES = (
    ("coverage-source-cell", 1_000),
    ("coverage-source-range", 1_000),
    ("coverage-empty-source", 1_000),
    ("coverage-unicode-escaped-source", 1_000),
    ("coverage-whole-columns", 1_000),
    ("coverage-whole-rows", 1_000),
    ("coverage-cross-sheet-range", 1_000),
    ("coverage-nested-inherited", 1_000),
    ("coverage-ref-error", 1_000),
    ("coverage-colon-sheet-1k", 1_000),
    ("coverage-colon-sheet-4k", 1_000),
    ("coverage-colon-sheet-16k", 128),
    ("coverage-source-iri-1k", 1_000),
    ("coverage-source-iri-4k", 1_000),
    ("coverage-source-iri-16k", 128),
    ("coverage-source-iri-over-16k", 128),
)

FIELDS = (
    "group",
    "workload",
    "case",
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
    "expected_success",
    "mean_ns",
    "p50_ns",
    "p95_ns",
    "p99_ns",
    "alloc_calls_p50",
    "alloc_calls_max",
    "dealloc_calls_p50",
    "dealloc_calls_max",
    "requested_bytes_p50",
    "requested_bytes_max",
    "released_bytes_p50",
    "released_bytes_max",
    "live_before_p50",
    "live_after_p50",
    "live_after_max",
    "peak_live_delta_p50",
    "peak_live_delta_max",
    "successes_p50",
    "successes_max",
    "checksum_p50",
    "checksum_max",
    "max_rss_kib",
    "status",
)


def parse_key_values(line: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for token in line.split()[1:]:
        key, value = token.split("=", 1)
        values[key] = value
    return values


def rss_kib(time_path: Path) -> str:
    text = time_path.read_text()
    match = re.search(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", text, re.M)
    return match.group(1) if match else ""


def run_one(
    *,
    binary: Path,
    out_dir: Path,
    group: str,
    case: str,
    repeat: int,
    warmups: int,
    iterations: int,
) -> dict[str, str]:
    stem = f"parse-{case}"
    stdout_path = out_dir / f"{stem}.stdout"
    stderr_path = out_dir / f"{stem}.stderr"
    time_path = out_dir / f"{stem}.time"
    status_path = out_dir / f"{stem}.status"
    command = [
        "taskset",
        "-c",
        "2",
        "/usr/bin/time",
        "-v",
        "-o",
        str(time_path),
        str(binary),
        "--workload",
        "parse",
        "--case",
        case,
        "--warmups",
        str(warmups),
        "--iterations",
        str(iterations),
        "--repeat",
        str(repeat),
    ]
    with (out_dir / "commands.txt").open("a") as stream:
        stream.write(shlex.join(command) + "\n")
    with stdout_path.open("w") as stdout, stderr_path.open("w") as stderr:
        completed = subprocess.run(command, stdout=stdout, stderr=stderr)
    status_path.write_text(f"{completed.returncode}\n")
    stdout_lines = stdout_path.read_text().splitlines()
    config = parse_key_values(next(line for line in stdout_lines if line.startswith("config ")))
    result = parse_key_values(next(line for line in stdout_lines if line.startswith("result ")))
    row = {
        "group": group,
        "workload": "parse",
        "case": case,
        "input_bytes": config["input_bytes"],
        "repeat": config["repeat"],
        "warmups": config["warmups"],
        "iterations": config["iterations"],
        "expected_success": config["expected_success"],
        **result,
        "max_rss_kib": rss_kib(time_path),
        "status": str(completed.returncode),
    }
    return {field: row.get(field, "") for field in FIELDS}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--group",
        choices=("comparable", "coverage"),
        default="comparable",
        help="comparable keeps the twelve baseline lanes; coverage is candidate-only",
    )
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=15)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    cases = COMPARABLE_CASES if args.group == "comparable" else COVERAGE_CASES
    (args.output / "group.json").write_text(
        json.dumps(
            {
                "group": args.group,
                "case_count": len(cases),
                "cases": [{"case": case, "repeat": repeat} for case, repeat in cases],
                "warmups": args.warmups,
                "iterations": args.iterations,
            },
            indent=2,
        )
        + "\n"
    )
    rows = [
        run_one(
            binary=args.binary,
            out_dir=args.output,
            group=args.group,
            case=case,
            repeat=repeat,
            warmups=args.warmups,
            iterations=args.iterations,
        )
        for case, repeat in cases
    ]
    with (args.output / "raw.csv").open("w", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


if __name__ == "__main__":
    main()
