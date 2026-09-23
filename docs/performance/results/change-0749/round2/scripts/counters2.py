#!/usr/bin/env python3
"""Change 0749 round two: per-owner hardware counters.

For every case, arm and heap layout (argv[0] extra bytes 0, 16 and 40), two
pinned processes run the same timed owner with a small and a large sample
count under `perf stat -e instructions,cycles,instructions:u,cycles:u,
page-faults`. The difference divided by the sample-count difference is the
per-owner cost with process setup and reporting removed. Probe processes run
with `--oracle first`, so only the first owner's output is digested. Harness
selectors run one (case, shape) per process pair, arms A and C.

A complete `perf stat` output already in OUT_DIR is reused, so an
interrupted lane resumes where it stopped.

Usage: counters2.py OUT_DIR [CASE_PREFIX ...] > counters.json
"""

import json
import os
import statistics
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run2 import CASES, CORE, binary  # noqa: E402

EVENTS = "instructions,cycles,instructions:u,cycles:u,page-faults"
LAYOUTS = (0, 2, 5)

PROBE_SIZES = {
    "reuse-v3-1000x2000": (5, 25),
    "reuse-v3-3000x2000": (2, 6),
    "reuse-v4-3000x2000": (2, 6),
    "reuse-v4-10000x4000": (2, 6),
}

HARNESS = [
    ("--writer-shape", case, shape)
    for case in ("doc_semantic_one_edit_save", "ppt_semantic_one_edit_save", "xls_semantic_one_edit_save")
    for shape in ("tiny", "large")
] + [
    ("--shape", case, shape)
    for case in ("cfb_open", "cfb_read_one")
    for shape in ("few-large", "many-small")
]


def harness_sizes(case, shape):
    if case == "cfb_read_one" and shape == "many-small":
        return (100, 5100)
    if shape in ("tiny", "many-small"):
        return (20, 1020)
    return (20, 220)


def probe_command(name, arm, layout, samples):
    command = CASES[name]["build"](arm, layout, "/dev/null")
    command[command.index("--samples") + 1] = str(samples)
    return command + ["--oracle", "first"]


def harness_command(flag, case, shape, arm, layout, samples, out):
    return [
        binary(arm, "litchi-perf-baseline", layout),
        "--case", case,
        flag, shape,
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
    # Resumable: a complete earlier output of the same command is reused.
    if os.path.exists(f"{stem}.perf"):
        counters = parse(f"{stem}.perf")
        if all(event in counters for event in EVENTS.split(",")):
            return counters
    full = ["taskset", "-c", CORE, "perf", "stat", "-x,", "-e", EVENTS, "-o", f"{stem}.perf", "--"] + command
    completed = subprocess.run(full, capture_output=True, text=True, timeout=600)
    if completed.returncode != 0:
        raise SystemExit(f"{full} failed: {completed.stderr[-2000:]}")
    return parse(f"{stem}.perf")


def main():
    out_dir = sys.argv[1]
    prefixes = tuple(sys.argv[2:])
    os.makedirs(out_dir, exist_ok=True)
    jobs = []
    for name, case in CASES.items():
        if case["arms"] != "probe" or (prefixes and not name.startswith(prefixes)):
            continue
        sizes = next(
            (value for prefix, value in PROBE_SIZES.items() if name.startswith(prefix)),
            (200, 2200),
        )
        jobs.append((name, ("A", "B", "C"), sizes,
                     lambda arm, layout, samples, stem, name=name: probe_command(name, arm, layout, samples)))
    for flag, case, shape in HARNESS:
        label = f"{case}/{shape}"
        if prefixes and not label.startswith(prefixes):
            continue
        jobs.append((label, ("A", "C"), harness_sizes(case, shape),
                     lambda arm, layout, samples, stem, flag=flag, case=case, shape=shape:
                     harness_command(flag, case, shape, arm, layout, samples, f"{stem}.json")))
    rows = []
    for layout in LAYOUTS:
        for label, arms, sizes, build in jobs:
            for arm in arms:
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
    for events in summary.values():
        for arms in events.values():
            base = statistics.median(arms["A"])
            for other in ("B", "C"):
                if other in arms and base:
                    arms[f"median_change_{other}_vs_A_pct"] = round(100.0 * (statistics.median(arms[other]) / base - 1.0), 3)
            if "B" in arms and "C" in arms and statistics.median(arms["B"]):
                arms["median_change_C_vs_B_pct"] = round(
                    100.0 * (statistics.median(arms["C"]) / statistics.median(arms["B"]) - 1.0), 3)
    json.dump({"events": EVENTS, "layouts_argv0_extra_bytes": [8 * x for x in LAYOUTS],
               "rows": rows, "summary": summary}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
