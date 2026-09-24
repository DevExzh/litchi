#!/usr/bin/env python3
"""Change 0769: per-owner hardware counters (record 0767's method).

For every case, arm and heap layout (argv[0] extra bytes 0, 16 and 40), two
pinned processes run the same timed owner with a small and a large sample
count under `perf stat -e instructions:u,cycles:u,instructions,cycles,
page-faults`. The difference divided by the sample-count difference is the
per-owner cost with process setup, corpus construction and reporting
removed. Probe processes run with `--oracle first`, so only the first
owner's output is digested. Harness selectors run one (case, shape) per
process pair.

A complete `perf stat` output already in OUT_DIR is reused, so an
interrupted lane resumes where it stopped.

Usage: counters.py OUT_DIR [CASE_PREFIX ...] > counters.json
"""

import json
import os
import statistics
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run import CASES, CORE, binary  # noqa: E402

EVENTS = "instructions:u,cycles:u,instructions,cycles,page-faults"
LAYOUTS = (0, 2, 5)
ARMS = ("base", "cand")


def probe_sizes(name):
    if name.endswith("-1000"):
        return (10, 60)
    if name.endswith("-3000"):
        return (4, 24)
    if name.endswith("-10000"):
        return (2, 12)
    if name == "write-reuse-v3-mini-3000-grow":
        return (2, 7)
    return (200, 2200)


HARNESS = [
    ("--shape", "cfb_open", "tiny", (100, 5100), ["--payload", "incompressible"]),
    ("--shape", "cfb_open", "many-small", (20, 1020), ["--payload", "incompressible"]),
    ("--shape", "cfb_open", "few-large", (20, 1020), ["--payload", "incompressible"]),
    ("--shape", "cfb_open", "wide-root", (10, 210), ["--payload", "incompressible"]),
    ("--shape", "cfb_list_streams", "tiny", (100, 10100), ["--payload", "incompressible"]),
    ("--shape", "cfb_list_streams", "many-small", (100, 5100), ["--payload", "incompressible"]),
    ("--shape", "cfb_list_streams", "few-large", (100, 10100), ["--payload", "incompressible"]),
    ("--shape", "cfb_list_streams", "wide-root", (20, 1020), ["--payload", "incompressible"]),
    ("--semantic-shape", "doc_semantic_open", "tiny", (20, 520), []),
    ("--semantic-shape", "doc_semantic_open", "large", (5, 105), []),
]


def probe_command(name, arm, layout, samples):
    command = CASES[name](arm, layout, "/dev/null")
    command[command.index("--samples") + 1] = str(samples)
    return command + ["--oracle", "first"]


def harness_command(flag, case, shape, extra, arm, layout, samples, out):
    return [
        binary(arm, "litchi-perf-baseline", layout),
        "--case", case,
        flag, shape,
    ] + extra + [
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
    completed = subprocess.run(full, capture_output=True, text=True, timeout=900)
    if completed.returncode != 0:
        raise SystemExit(f"{full} failed: {completed.stderr[-2000:]}")
    return parse(f"{stem}.perf")


def main():
    out_dir = sys.argv[1]
    prefixes = tuple(sys.argv[2:])
    os.makedirs(out_dir, exist_ok=True)
    jobs = []
    for name in CASES:
        if name.startswith("harness-") or (prefixes and not name.startswith(prefixes)):
            continue
        jobs.append((name, probe_sizes(name),
                     lambda arm, layout, samples, stem, name=name: probe_command(name, arm, layout, samples)))
    for flag, case, shape, sizes, extra in HARNESS:
        label = f"{case}/{shape}"
        if prefixes and not label.startswith(prefixes):
            continue
        jobs.append((label, sizes,
                     lambda arm, layout, samples, stem, flag=flag, case=case, shape=shape, extra=extra:
                     harness_command(flag, case, shape, extra, arm, layout, samples, f"{stem}.json")))
    rows = []
    for layout in LAYOUTS:
        for label, sizes, build in jobs:
            # Alternate the arm order with the layout.
            for arm in (ARMS if layout % 2 == 0 else ARMS[::-1]):
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
            base = statistics.median(arms["base"])
            if "cand" in arms and base:
                arms["median_change_pct"] = round(100.0 * (statistics.median(arms["cand"]) / base - 1.0), 3)
    json.dump({"events": EVENTS, "layouts_argv0_extra_bytes": [8 * x for x in LAYOUTS],
               "rows": rows, "summary": summary}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
