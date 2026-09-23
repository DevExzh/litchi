#!/usr/bin/env python3
"""Change 0745 per-iteration hardware-counter lane.

For every case, arm and heap layout (argv[0] extra bytes), two pinned processes
run the same timed owner with 20 and 120 measured samples (5 warmups) under
`perf stat -e instructions:u,page-faults,cycles`. The difference divided by
100 is the per-owner cost with process setup, oracle and reporting work
removed except for their per-sample parts: the probe's per-sample output
digest and copy are identical in every arm, so they cancel in comparisons.
User-mode instruction counts are layout-invariant evidence of work removed;
page faults show the allocator-state effects that move wall-clock timings.

Usage: counters.py OUT_DIR
"""

import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run import ARMS, CASES, CORE, binary  # noqa: E402

EVENTS = "instructions:u,page-faults,cycles"
LAYOUTS = (0, 2, 5)  # round indexes: 0, 16 and 40 extra argv[0] bytes
SIZES = (20, 120)
COUNTED = [name for name in CASES if not name.startswith("harness-")]


def command(name, arm, layout, samples):
    base = CASES[name](arm, layout, "/dev/null")
    for flag in ("--samples",):
        index = base.index(flag)
        base[index + 1] = str(samples)
    return base


def parse(path):
    counters = {}
    with open(path) as handle:
        for line in handle:
            parts = line.strip().split(",")
            if len(parts) > 3 and parts[0] and not line.startswith("#"):
                counters[parts[2]] = float(parts[0])
    return counters


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    rows = []
    for layout in LAYOUTS:
        for name in COUNTED:
            for arm in ARMS:
                measured = {}
                for samples in SIZES:
                    stem = f"{out_dir}/{name}-{arm}-l{layout}-n{samples}"
                    probe = command(name, arm, layout, samples)
                    full = ["taskset", "-c", CORE, "perf", "stat", "-x,", "-e", EVENTS, "-o", f"{stem}.perf", "--"] + probe
                    completed = subprocess.run(full, capture_output=True, text=True)
                    if completed.returncode != 0:
                        raise SystemExit(f"{full} failed: {completed.stderr[-2000:]}")
                    measured[samples] = parse(f"{stem}.perf")
                per_owner = {
                    event: (measured[SIZES[1]][event] - measured[SIZES[0]][event]) / (SIZES[1] - SIZES[0])
                    for event in measured[SIZES[1]]
                }
                rows.append({"case": name, "arm": arm, "argv0_extra_bytes": 8 * layout, "per_owner": per_owner})
    summary = {}
    for row in rows:
        entry = summary.setdefault(row["case"], {})
        for event, value in row["per_owner"].items():
            entry.setdefault(event, {}).setdefault(row["arm"], []).append(round(value, 1))
    json.dump({"events": EVENTS, "sizes": SIZES, "layouts_argv0_extra_bytes": [8 * x for x in LAYOUTS], "rows": rows, "summary": summary}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
