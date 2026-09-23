#!/usr/bin/env python3
"""Change 0749 per-owner hardware-counter lane.

For every case, arm and heap layout (argv[0] extra bytes 0, 16 and 40), two
pinned processes run the same timed owner with a small and a large sample
count under `perf stat -e instructions,cycles,instructions:u,cycles:u,
page-faults`. The difference divided by the sample-count difference is the
per-owner cost with process setup and reporting removed. Probe processes run
with `--oracle first`, so only the first owner's output is digested and the
per-owner difference holds the timed owner and its loop. Harness processes
keep their own per-sample oracles, identical in both arms. Instruction counts are layout-invariant evidence of work
removed; cycles and page faults show what the layout moves.

Harness selectors run one (case, shape) per process pair.

Usage: counters.py OUT_DIR
"""

import json
import os
import statistics
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run import ARMS, CASES, CORE, binary  # noqa: E402

EVENTS = "instructions,cycles,instructions:u,cycles:u,page-faults"
LAYOUTS = (0, 2, 5)

PROBE_SIZES = {
    "cfb-": (200, 2200),
    "container-": (50, 550),
    "doc-replace-nohf": (50, 550),
    "ppt-noop": (50, 550),
    "": (20, 220),
}

HARNESS = [
    (case, shape)
    for case in ("doc_semantic_one_edit_save", "ppt_semantic_one_edit_save", "xls_semantic_one_edit_save")
    for shape in ("tiny", "large")
]


def probe_sizes(name):
    for prefix, sizes in PROBE_SIZES.items():
        if name.startswith(prefix):
            return sizes
    raise KeyError(name)


def probe_command(name, arm, layout, samples):
    # Only the first owner's output is digested, so the per-owner difference
    # contains the timed owner and the loop, not a whole-output SHA-256.
    command = CASES[name](arm, layout, "/dev/null")
    command[command.index("--samples") + 1] = str(samples)
    return command + ["--oracle", "first"]


def harness_command(case, shape, arm, layout, samples, out):
    return [
        binary(arm, "litchi-perf-baseline", layout),
        "--case", case,
        "--writer-shape", shape,
        "--samples", str(samples),
        "--warmup", "5",
        "--json", out,
    ]


def parse(path):
    counters = {}
    with open(path) as handle:
        for line in handle:
            parts = line.strip().split(",")
            if len(parts) > 3 and parts[0] and not line.startswith("#"):
                counters[parts[2]] = float(parts[0])
    return counters


def measure(stem, command):
    full = ["taskset", "-c", CORE, "perf", "stat", "-x,", "-e", EVENTS, "-o", f"{stem}.perf", "--"] + command
    completed = subprocess.run(full, capture_output=True, text=True)
    if completed.returncode != 0:
        raise SystemExit(f"{full} failed: {completed.stderr[-2000:]}")
    return parse(f"{stem}.perf")


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    rows = []
    jobs = []
    for name in CASES:
        if name.startswith("harness-"):
            continue
        jobs.append((name, probe_sizes(name), lambda arm, layout, samples, stem, name=name: probe_command(name, arm, layout, samples)))
    for case, shape in HARNESS:
        label = f"{case}/{shape}"
        sizes = (20, 1020) if shape == "tiny" else (20, 220)
        jobs.append((label, sizes, lambda arm, layout, samples, stem, case=case, shape=shape: harness_command(case, shape, arm, layout, samples, f"{stem}.json")))
    for layout in LAYOUTS:
        for label, sizes, build in jobs:
            for arm in ARMS:
                measured = {}
                for samples in sizes:
                    stem = f"{out_dir}/{label.replace('/', '__')}-{arm}-l{layout}-n{samples}"
                    measured[samples] = measure(stem, build(arm, layout, samples, stem))
                per_owner = {
                    event: (measured[sizes[1]][event] - measured[sizes[0]][event]) / (sizes[1] - sizes[0])
                    for event in measured[sizes[1]]
                }
                rows.append({"case": label, "arm": arm, "argv0_extra_bytes": 8 * layout,
                             "sizes": sizes, "per_owner": per_owner})
    summary = {}
    for row in rows:
        entry = summary.setdefault(row["case"], {})
        for event, value in row["per_owner"].items():
            entry.setdefault(event, {}).setdefault(row["arm"], []).append(round(value, 1))
    for case, events in summary.items():
        for event, arms in events.items():
            if "base" in arms and "cand" in arms:
                base = statistics.median(arms["base"])
                cand = statistics.median(arms["cand"])
                arms["median_change_pct"] = round(100.0 * (cand / base - 1.0), 3) if base else None
    json.dump({"events": EVENTS, "layouts_argv0_extra_bytes": [8 * x for x in LAYOUTS],
               "rows": rows, "summary": summary}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
