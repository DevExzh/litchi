#!/usr/bin/env python3
"""Run candidate-only scalar evaluation measurements by phase.

The parser phase measures Expression::parse.  The evaluate phase parses once
outside the timed region and measures evaluate_scalar.  The parse-evaluate
phase parses and evaluates on every repetition.  These phase labels are kept
in every row so the candidate-only numbers are not presented as a tokenizer
comparison or as a baseline for the new evaluator.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import re
import shlex
import subprocess
from pathlib import Path

CASES = (
    ("eval-flat-64", 64),
    ("eval-flat-256", 32),
    ("eval-flat-1024", 8),
    ("eval-flat-4096", 2),
    ("eval-utf8-text-text-64", 64),
    ("eval-utf8-text-text-256", 32),
    ("eval-utf8-text-text-1024", 8),
    ("eval-utf8-text-text-4096", 2),
    ("eval-utf8-number-left-64", 64),
    ("eval-utf8-number-left-256", 32),
    ("eval-utf8-number-left-1024", 8),
    ("eval-utf8-number-left-4096", 2),
    ("eval-utf8-text-number-right-64", 64),
    ("eval-utf8-text-number-right-256", 32),
    ("eval-utf8-text-number-right-1024", 8),
    ("eval-utf8-text-number-right-4096", 2),
    ("eval-escaped-text-text-64", 64),
    ("eval-escaped-text-text-256", 32),
    ("eval-escaped-text-text-1024", 8),
    ("eval-escaped-text-text-4096", 2),
    ("eval-escaped-text-number-right-64", 64),
    ("eval-escaped-text-number-right-256", 32),
    ("eval-escaped-text-number-right-1024", 8),
    ("eval-escaped-text-number-right-4096", 2),
    ("eval-coerce-64", 64),
    ("eval-coerce-256", 32),
    ("eval-coerce-1024", 8),
    ("eval-coerce-4096", 2),
    ("eval-long-name-4096", 2),
    ("eval-array-4096", 2),
    ("eval-reference", 128),
    ("eval-exact-step", 128),
    ("eval-work-limit", 128),
    ("eval-text-limit", 128),
    ("eval-cancelled", 128),
)
PHASES = ("parse", "evaluate", "parse-evaluate")
FIELDS = (
    "group",
    "workload",
    "phase",
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
    "refusals_p50",
    "refusals_max",
    "checksum_p50",
    "checksum_max",
    "output_reserved_bytes_p50",
    "output_reserved_bytes_max",
    "failure",
    "max_rss_kib",
    "started_at",
    "finished_at",
    "status",
)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="milliseconds").replace(
        "+00:00", "Z"
    )


def parse_key_values(line: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for token in line.split()[1:]:
        key, value = token.split("=", 1)
        values[key] = value
    return values


def rss_kib(path: Path) -> str:
    if not path.exists():
        return ""
    match = re.search(
        r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$",
        path.read_text(),
        re.MULTILINE,
    )
    return match.group(1) if match else ""


def run_one(
    binary: Path,
    out_dir: Path,
    phase: str,
    case: str,
    repeat: int,
    warmups: int,
    iterations: int,
) -> tuple[dict[str, str], dict[str, str]]:
    stem = f"{phase}-{case}"
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
        "evaluation",
        "--phase",
        phase,
        "--case",
        case,
        "--warmups",
        str(warmups),
        "--iterations",
        str(iterations),
        "--repeat",
        str(repeat),
    ]
    started_at = utc_now()
    with (out_dir / "commands.txt").open("a", encoding="utf-8") as stream:
        stream.write(shlex.join(command) + "\n")
    with stdout_path.open("w", encoding="utf-8") as stdout, stderr_path.open(
        "w", encoding="utf-8"
    ) as stderr:
        completed = subprocess.run(command, stdout=stdout, stderr=stderr, check=False)
    finished_at = utc_now()
    status_path.write_text(f"{completed.returncode}\n", encoding="utf-8")

    lines = stdout_path.read_text(encoding="utf-8").splitlines()
    config_line = next((line for line in lines if line.startswith("config ")), None)
    result_line = next((line for line in lines if line.startswith("result ")), None)
    config = parse_key_values(config_line) if config_line else {}
    result = parse_key_values(result_line) if result_line else {}
    row = {
        "group": "evaluation",
        "workload": "evaluation",
        **config,
        **result,
        "max_rss_kib": rss_kib(time_path),
        "started_at": started_at,
        "finished_at": finished_at,
        "status": str(completed.returncode),
    }
    sequence = {
        "phase": phase,
        "case": case,
        "repeat": str(repeat),
        "started_at": started_at,
        "finished_at": finished_at,
        "status": str(completed.returncode),
        "stdout": stdout_path.name,
        "stderr": stderr_path.name,
        "time": time_path.name,
    }
    return ({field: row.get(field, "") for field in FIELDS}, sequence)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--phase", choices=("all", *PHASES), default="all")
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=15)
    args = parser.parse_args()
    if args.warmups < 0 or args.iterations <= 0:
        raise SystemExit("warmups must be nonnegative and iterations must be positive")

    binary = args.binary.resolve()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    phases = PHASES if args.phase == "all" else (args.phase,)
    (output / "commands.txt").write_text("", encoding="utf-8")
    (output / "group.json").write_text(
        json.dumps(
            {
                "group": "evaluation",
                "candidate_only": True,
                "phase_selection": args.phase,
                "phases": list(phases),
                "case_count": len(CASES),
                "cases": [{"case": case, "repeat": repeat} for case, repeat in CASES],
                "warmups": args.warmups,
                "iterations": args.iterations,
                "binary": str(binary),
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    rows: list[dict[str, str]] = []
    sequence: list[dict[str, str]] = []
    for phase in phases:
        for case, repeat in CASES:
            row, item = run_one(
                binary, output, phase, case, repeat, args.warmups, args.iterations
            )
            rows.append(row)
            sequence.append(item)
    with (output / "raw.csv").open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    (output / "sequence.json").write_text(
        json.dumps(
            {
                "group": "evaluation",
                "candidate_only": True,
                "started_at": sequence[0]["started_at"] if sequence else None,
                "finished_at": sequence[-1]["finished_at"] if sequence else None,
                "rows": sequence,
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
