#!/usr/bin/env python3
"""Run the candidate-only strict OpenFormula expression profile."""
from __future__ import annotations
import argparse, csv, json, re, shlex, subprocess
from pathlib import Path

CASES = (
    ("expr-flat-64", 256), ("expr-flat-256", 128), ("expr-flat-1024", 32), ("expr-flat-4096", 8),
    ("expr-array-64", 256), ("expr-array-256", 128), ("expr-array-1024", 32), ("expr-array-4096", 8),
    ("expr-name-64", 256), ("expr-name-256", 128), ("expr-name-1024", 32), ("expr-name-4096", 8),
    ("expr-string-64", 256), ("expr-string-256", 128), ("expr-string-1024", 32), ("expr-string-4096", 8),
    ("expr-reference-64", 256), ("expr-reference-256", 128), ("expr-reference-1024", 32), ("expr-reference-4096", 8),
    ("expr-depth-limit-256", 8), ("expr-malformed-trailing-4096", 8),
    ("expr-malformed-unclosed-array-4096", 8), ("expr-malformed-unclosed-string-64k", 8),
)
FIELDS = (
    "group", "workload", "case", "input_bytes", "repeat", "warmups", "iterations", "expected_success",
    "mean_ns", "p50_ns", "p95_ns", "p99_ns", "alloc_calls_p50", "alloc_calls_max", "dealloc_calls_p50",
    "dealloc_calls_max", "requested_bytes_p50", "requested_bytes_max", "released_bytes_p50", "released_bytes_max",
    "live_before_p50", "live_after_p50", "live_after_max", "peak_live_delta_p50", "peak_live_delta_max",
    "successes_p50", "successes_max", "checksum_p50", "checksum_max", "max_rss_kib", "status",
)

def parse_key_values(line: str) -> dict[str, str]:
    values = {}
    for token in line.split()[1:]:
        key, value = token.split("=", 1)
        values[key] = value
    return values

def rss_kib(path: Path) -> str:
    match = re.search(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", path.read_text(), re.M)
    return match.group(1) if match else ""

def run_one(binary: Path, out_dir: Path, case: str, repeat: int, warmups: int, iterations: int) -> dict[str, str]:
    stem = f"parse-{case}"
    stdout_path, stderr_path = out_dir / f"{stem}.stdout", out_dir / f"{stem}.stderr"
    time_path, status_path = out_dir / f"{stem}.time", out_dir / f"{stem}.status"
    command = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(time_path), str(binary),
               "--workload", "expression", "--case", case, "--warmups", str(warmups),
               "--iterations", str(iterations), "--repeat", str(repeat)]
    with (out_dir / "commands.txt").open("a") as stream:
        stream.write(shlex.join(command) + "\n")
    with stdout_path.open("w") as stdout, stderr_path.open("w") as stderr:
        completed = subprocess.run(command, stdout=stdout, stderr=stderr)
    status_path.write_text(f"{completed.returncode}\n")
    lines = stdout_path.read_text().splitlines()
    config = parse_key_values(next(line for line in lines if line.startswith("config ")))
    result = parse_key_values(next(line for line in lines if line.startswith("result ")))
    row = {"group": "expression", "workload": "expression", **config, **result,
           "max_rss_kib": rss_kib(time_path), "status": str(completed.returncode)}
    return {field: row.get(field, "") for field in FIELDS}

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=15)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    (args.output / "group.json").write_text(json.dumps({"group": "expression", "case_count": len(CASES),
        "cases": [{"case": case, "repeat": repeat} for case, repeat in CASES],
        "warmups": args.warmups, "iterations": args.iterations}, indent=2) + "\n")
    rows = [run_one(args.binary, args.output, case, repeat, args.warmups, args.iterations) for case, repeat in CASES]
    with (args.output / "raw.csv").open("w", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\n")
        writer.writeheader(); writer.writerows(rows)

if __name__ == "__main__":
    main()
