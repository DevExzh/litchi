#!/usr/bin/env python3
"""Run bounded scalar/logical/bitwise evaluation lanes by phase.

The comparable group is identical across the baseline and candidate source
snapshots and covers representative scalar and logical workloads. The bitwise
group measures the new function family; the baseline can be run only as a
capability/refusal control because those functions are intentionally unsupported
there. Evaluation contexts are built by the binary outside each timed operation
and remain alive through allocator counter reads; see src/main.rs.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import json
import re
import shlex
import subprocess
from pathlib import Path

COMPARABLE_CASES = (
    ("control-flat-64", 64),
    ("control-flat-256", 32),
    ("control-flat-1024", 8),
    ("control-flat-4096", 2),
    ("control-coerce-64", 64),
    ("control-coerce-256", 32),
    ("control-coerce-1024", 8),
    ("control-coerce-4096", 2),
    ("control-utf8-left-64", 64),
    ("control-utf8-left-256", 32),
    ("control-utf8-left-1024", 8),
    ("control-utf8-left-4096", 2),
    ("control-escaped-64", 64),
    ("control-escaped-256", 32),
    ("control-escaped-1024", 8),
    ("control-escaped-4096", 2),
    ("control-true", 128),
    ("control-false", 128),
    ("failure-reference", 128),
    ("failure-array", 128),
    ("failure-name", 128),
    ("failure-work", 128),
    ("failure-memory", 128),
    ("failure-stack", 128),
    ("failure-cancelled", 128),
    ("logical-true", 128),
    ("logical-false", 128),
    ("logical-and-64", 64),
    ("logical-and-1024", 8),
    ("logical-or-256", 32),
    ("logical-xor-4096", 2),
    ("logical-not-256", 32),
    ("logical-if-1024", 8),
    ("logical-iferror-4096", 2),
    ("logical-ifna-4096", 2),
    ("lazy-if-true-heavy-1024", 8),
    ("lazy-if-false-reference", 128),
    ("text-if-escaped", 128),
    ("text-if-concat", 128),
)

BITWISE_CASES = (
    *[
        (
            f"bitwise-{name}-{size}",
            2 if size == 4096 else 8 if size == 1024 else 32 if size == 256 else 64,
        )
        for name in ("and", "or", "xor", "lshift", "rshift")
        for size in (64, 256, 1024, 4096)
    ],
    ("bitwise-coerce-text", 128),
    ("bitwise-coerce-fraction", 128),
    ("bitwise-error-negative", 128),
    ("bitwise-error-overflow", 128),
    ("bitwise-error-shift", 128),
    ("bitwise-shift-negative", 128),
    ("bitwise-error-numeric", 128),
    ("bitwise-error-arity", 128),
    ("bitwise-lazy", 128),
    ("bitwise-lazy-error", 128),
    ("bitwise-lazy-selected", 128),
)

PHASES = ("parse", "evaluate", "parse-evaluate")
FIELDS = (
    "group",
    "revision",
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
    revision: str,
    group: str,
    phase: str,
    case: str,
    repeat: int,
    warmups: int,
    iterations: int,
) -> tuple[dict[str, str], dict[str, str]]:
    stem = f"{revision}-{phase}-{case}"
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
        "bitwise-evaluation",
        "--revision",
        revision,
        "--group",
        group,
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
        "group": group,
        "revision": revision,
        "workload": "bitwise-evaluation",
        **config,
        **result,
        "max_rss_kib": rss_kib(time_path),
        "started_at": started_at,
        "finished_at": finished_at,
        "status": str(completed.returncode),
    }
    sequence = {
        "revision": revision,
        "group": group,
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
    parser.add_argument("--revision", choices=("baseline", "candidate"), required=True)
    parser.add_argument("--group", choices=("comparable", "bitwise", "all"), default="all")
    parser.add_argument("--phase", choices=("all", *PHASES), default="all")
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=15)
    args = parser.parse_args()
    if args.warmups < 0 or args.iterations <= 0:
        raise SystemExit("warmups must be nonnegative and iterations must be positive")

    if args.group == "comparable":
        cases = COMPARABLE_CASES
    elif args.group == "bitwise":
        cases = BITWISE_CASES
    else:
        cases = (*COMPARABLE_CASES, *BITWISE_CASES)
    binary = args.binary.resolve()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    phases = PHASES if args.phase == "all" else (args.phase,)
    (output / "commands.txt").write_text("", encoding="utf-8")
    (output / "group.json").write_text(
        json.dumps(
            {
                "group": args.group,
                "revision": args.revision,
                "candidate_only": args.group == "bitwise",
                "phase_selection": args.phase,
                "phases": list(phases),
                "case_count": len(cases),
                "cases": [{"case": case, "repeat": repeat} for case, repeat in cases],
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
        for case, repeat in cases:
            row, item = run_one(
                binary,
                output,
                args.revision,
                args.group,
                phase,
                case,
                repeat,
                args.warmups,
                args.iterations,
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
                "group": args.group,
                "revision": args.revision,
                "started_at": sequence[0]["started_at"] if sequence else None,
                "finished_at": sequence[-1]["finished_at"] if sequence else None,
                "rows": sequence,
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    (output / "runner-sha256.txt").write_text(
        f"{hashlib.sha256(Path(__file__).read_bytes()).hexdigest()}  run.py\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
