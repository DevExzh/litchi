#!/usr/bin/env python3
"""Run the bounded formula lookup/parser profile and emit a normalized CSV."""

from __future__ import annotations

import argparse
import csv
import re
import subprocess
from pathlib import Path


LOOKUP_CASES = (
    ("lookup-sum-upper", 10_000),
    ("lookup-sum-lower", 10_000),
    ("lookup-sum-mixed", 10_000),
    ("lookup-vlookup-upper", 10_000),
    ("lookup-vlookup-lower", 10_000),
    ("lookup-vlookup-mixed", 10_000),
    ("lookup-invalid-short", 10_000),
    ("lookup-invalid-ascii-4k", 10_000),
    ("lookup-invalid-ascii-64k", 10_000),
    ("lookup-invalid-unicode-4k", 10_000),
    ("lookup-invalid-unicode-64k", 10_000),
)

PARSE_CASES = (
    ("parse-sum-upper", 1_000),
    ("parse-sum-lower", 1_000),
    ("parse-sum-mixed", 1_000),
    ("parse-vlookup-upper", 1_000),
    ("parse-vlookup-lower", 1_000),
    ("parse-vlookup-mixed", 1_000),
    ("parse-absolute", 1_000),
    ("parse-cell-heavy", 128),
    ("parse-invalid-ascii-4k", 256),
    ("parse-invalid-ascii-64k", 4),
    ("parse-invalid-unicode-4k", 256),
    ("parse-invalid-unicode-64k", 4),
)

FIELDS = (
    "workload",
    "case",
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
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
    workload: str,
    case: str,
    repeat: int,
    warmups: int,
    iterations: int,
) -> dict[str, str]:
    stem = f"{workload}-{case}"
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
        workload,
        "--case",
        case,
        "--warmups",
        str(warmups),
        "--iterations",
        str(iterations),
        "--repeat",
        str(repeat),
    ]
    (out_dir / "commands.txt").open("a").write(" ".join(command) + "\n")
    completed = subprocess.run(command, stdout=stdout_path.open("w"), stderr=stderr_path.open("w"))
    status_path.write_text(f"{completed.returncode}\n")
    stdout_lines = stdout_path.read_text().splitlines()
    config = parse_key_values(next(line for line in stdout_lines if line.startswith("config ")))
    result = parse_key_values(next(line for line in stdout_lines if line.startswith("result ")))
    row = {
        "workload": workload,
        "case": case,
        "input_bytes": config["input_bytes"],
        "repeat": config["repeat"],
        "warmups": config["warmups"],
        "iterations": config["iterations"],
        **result,
        "max_rss_kib": rss_kib(time_path),
        "status": str(completed.returncode),
    }
    return {field: row.get(field, "") for field in FIELDS}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=15)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    rows: list[dict[str, str]] = []
    for workload, cases in (("lookup", LOOKUP_CASES), ("parse", PARSE_CASES)):
        for case, repeat in cases:
            rows.append(
                run_one(
                    binary=args.binary,
                    out_dir=args.output,
                    workload=workload,
                    case=case,
                    repeat=repeat,
                    warmups=args.warmups,
                    iterations=args.iterations,
                )
            )
    with (args.output / "raw.csv").open("w", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


if __name__ == "__main__":
    main()
