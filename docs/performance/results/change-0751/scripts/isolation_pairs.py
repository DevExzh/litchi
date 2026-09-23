#!/usr/bin/env python3
"""Per-iteration cycles and instructions by `perf stat` isolation pairs (change 0751).

Each measurement runs one harness child pinned with ``taskset -c CORE`` under
``perf stat -e cycles,instructions -x ,`` twice, at ``--samples LOW`` and
``--samples HIGH`` (same warmup), and divides the difference of the two
whole-child counts by ``HIGH - LOW``. The corpus construction, the untimed
gates and process start-up cancel; what remains is one loop iteration: the
timed lifecycle plus the per-iteration checks the runner performs after its
clocks stop (reopen, semantic verification, output digest), which are the
same code in both arms except where this change touches a capture. Arms
alternate before, after, after, before for each repetition.

Usage:
  isolation_pairs.py --before BIN --after BIN --core 4 --low 2 --high 12 \\
      --warmup 1 --repeats 2 --out DIR CASE [CASE ...]
"""

from __future__ import annotations

import argparse
import json
import statistics
import subprocess
import sys
from pathlib import Path


def counters(path: Path) -> dict[str, float]:
    values: dict[str, float] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        fields = line.split(",")
        if len(fields) >= 3 and fields[0].strip():
            event = fields[2].split(":")[0]
            try:
                values[event] = float(fields[0])
            except ValueError:
                continue
    return values


def run(binary: Path, core: str, case: str, samples: int, warmup: int, out: Path) -> dict[str, float]:
    perf = out.with_suffix(".perf.csv")
    argv = [
        "taskset", "-c", core,
        "perf", "stat", "-e", "cycles,instructions", "-x", ",", "-o", str(perf), "--",
        str(binary), "--case", case, "--samples", str(samples), "--warmup", str(warmup),
        "--json", str(out),
    ]
    result = subprocess.run(argv, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, check=False)
    if result.returncode != 0:
        sys.stderr.write(result.stderr.decode("utf-8", "replace"))
        raise SystemExit(result.returncode)
    return counters(perf)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--before", required=True, type=Path)
    parser.add_argument("--after", required=True, type=Path)
    parser.add_argument("--core", required=True)
    parser.add_argument("--low", type=int, default=2)
    parser.add_argument("--high", type=int, default=12)
    parser.add_argument("--warmup", type=int, default=1)
    parser.add_argument("--repeats", type=int, default=2)
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("cases", nargs="+")
    args = parser.parse_args()
    binaries = {"before": args.before, "after": args.after}
    report: dict = {}
    for case in args.cases:
        case_dir = args.out / case
        case_dir.mkdir(parents=True, exist_ok=True)
        per_arm: dict[str, list[dict[str, float]]] = {"before": [], "after": []}
        for repeat in range(args.repeats):
            for slot, arm in enumerate(("before", "after", "after", "before")):
                low = run(binaries[arm], args.core, case, args.low, args.warmup,
                          case_dir / f"p{repeat}-s{slot}-{arm}-low.json")
                high = run(binaries[arm], args.core, case, args.high, args.warmup,
                           case_dir / f"p{repeat}-s{slot}-{arm}-high.json")
                span = args.high - args.low
                per_arm[arm].append({
                    event: (high[event] - low[event]) / span
                    for event in ("cycles", "instructions")
                    if event in high and event in low
                })
                print(f"{case} p{repeat} s{slot} {arm}: {per_arm[arm][-1]}", flush=True)
        report[case] = {
            arm: {
                "per_iteration": rows,
                "median_cycles": statistics.median(row["cycles"] for row in rows),
                "median_instructions": statistics.median(row["instructions"] for row in rows),
            }
            for arm, rows in per_arm.items()
        }
        for event in ("cycles", "instructions"):
            before = report[case]["before"][f"median_{event}"]
            after = report[case]["after"][f"median_{event}"]
            report[case][f"{event}_ratio_after_over_before"] = after / before
    (args.out / "isolation-pairs.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({case: {key: value for key, value in body.items() if key.endswith("ratio_after_over_before")} for case, body in report.items()}, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
