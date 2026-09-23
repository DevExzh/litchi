#!/usr/bin/env python3
"""Run the change 0751 before/after matrix in ABBA order on one pinned core.

Adapted from change 0742's runner. Every process is a fresh harness child
pinned with ``taskset -c CORE``; with ``--perf-stat`` the child runs under
``perf stat -e cycles,instructions -x ,`` whose CSV lands beside the report
(``<report>.perf.csv``) and counts the whole child. For each case and round
the order is before, after, after, before. Raw JSON reports land in
``OUT/<lane>/<case>/r<round>-s<slot>-<arm>.json`` and every invocation (argv,
arm, binary SHA-256, start/end monotonic and wall times, exit code) is
appended to ``OUT/<lane>/receipts.jsonl``.

Usage:
  run_abba.py --lane native --before BIN --after BIN --core 4 --rounds 4 \\
      --samples 20 --warmup 3 --out DIR [--perf-stat] CASE [CASE ...]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
import time
from pathlib import Path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--lane", required=True)
    parser.add_argument("--before", required=True, type=Path)
    parser.add_argument("--after", required=True, type=Path)
    parser.add_argument("--core", required=True)
    parser.add_argument("--rounds", type=int, default=4)
    parser.add_argument("--samples", type=int, required=True)
    parser.add_argument("--warmup", type=int, required=True)
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--perf-stat", action="store_true")
    parser.add_argument("cases", nargs="+")
    args = parser.parse_args()

    binaries = {"before": args.before, "after": args.after}
    hashes = {arm: sha256_file(path) for arm, path in binaries.items()}
    lane_dir = args.out / args.lane
    lane_dir.mkdir(parents=True, exist_ok=True)
    receipts = (lane_dir / "receipts.jsonl").open("a", encoding="utf-8")
    order = ["before", "after", "after", "before"]
    for case in args.cases:
        case_dir = lane_dir / case
        case_dir.mkdir(parents=True, exist_ok=True)
        for round_index in range(args.rounds):
            for slot, arm in enumerate(order):
                report = case_dir / f"r{round_index}-s{slot}-{arm}.json"
                argv = ["taskset", "-c", str(args.core)]
                if args.perf_stat:
                    argv += [
                        "perf",
                        "stat",
                        "-e",
                        "cycles,instructions",
                        "-x",
                        ",",
                        "-o",
                        str(report.with_suffix(".perf.csv")),
                        "--",
                    ]
                argv += [
                    str(binaries[arm]),
                    "--case",
                    case,
                    "--samples",
                    str(args.samples),
                    "--warmup",
                    str(args.warmup),
                    "--json",
                    str(report),
                ]
                started_wall = time.time()
                started = time.monotonic_ns()
                result = subprocess.run(
                    argv,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.PIPE,
                    check=False,
                )
                finished = time.monotonic_ns()
                receipt = {
                    "lane": args.lane,
                    "case": case,
                    "round": round_index,
                    "slot": slot,
                    "arm": arm,
                    "binary": str(binaries[arm]),
                    "binary_sha256": hashes[arm],
                    "argv": argv,
                    "report": str(report.relative_to(args.out)),
                    "started_unix": started_wall,
                    "monotonic_ns": [started, finished],
                    "exit_code": result.returncode,
                    "stderr_tail": result.stderr.decode("utf-8", "replace")[-400:],
                }
                receipts.write(json.dumps(receipt, sort_keys=True) + "\n")
                receipts.flush()
                if result.returncode != 0:
                    print(json.dumps(receipt, indent=2), file=sys.stderr)
                    return result.returncode
                print(
                    f"{args.lane} {case} r{round_index} s{slot} {arm}: "
                    f"{(finished - started) / 1e9:.1f}s",
                    flush=True,
                )
    return 0


if __name__ == "__main__":
    sys.exit(main())
