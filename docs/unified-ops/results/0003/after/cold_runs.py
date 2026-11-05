#!/usr/bin/env python3
"""S10 follow-up: first-call costs in a fresh process, before vs after.

Runs the cold-open probe (cold_open_probe.rs) built at both revisions as
fresh processes, pinned, alternating ABBA per round, and reports for each
variant the median first-round and second-round elapsed time and minor
faults, plus the leaf step medians (XLSX).

usage: cold_runs.py CPU ROUNDS OUT.json NAME=BINARY:FORMAT:MODE:PATH ...
Each round runs the variants in order, then the next round in reverse.
"""
import json
import statistics
import subprocess
import sys

cpu, rounds, out = sys.argv[1], int(sys.argv[2]), sys.argv[3]
variants = []
for arg in sys.argv[4:]:
    name, rest = arg.split("=", 1)
    binary, fmt, mode, path = rest.split(":", 3)
    variants.append((name, binary, fmt, mode, path))
records = {name: [] for name, *_ in variants}
for index in range(rounds):
    order = variants if index % 2 == 0 else list(reversed(variants))
    for name, binary, fmt, mode, path in order:
        proc = subprocess.run(["taskset", "-c", cpu, binary, fmt, mode, path],
                              capture_output=True, check=True)
        records[name].append(json.loads(proc.stdout))
summary = {}
for name, *_ in variants:
    runs = records[name]
    def med(round_index, key):
        return statistics.median(r["rounds"][round_index][key] for r in runs)
    entry = {
        "processes": len(runs),
        "first_elapsed_us": med(0, "elapsed_ns") / 1000,
        "second_elapsed_us": med(1, "elapsed_ns") / 1000,
        "third_elapsed_us": med(2, "elapsed_ns") / 1000,
        "first_minor_faults": med(0, "minor_faults"),
        "second_minor_faults": med(1, "minor_faults"),
        "first_elapsed_us_min_max": [min(r["rounds"][0]["elapsed_ns"] for r in runs) / 1000,
                                     max(r["rounds"][0]["elapsed_ns"] for r in runs) / 1000],
    }
    steps = runs[0]["rounds"][0]["steps"]
    if steps:
        entry["first_steps_us"] = {step: statistics.median(r["rounds"][0]["steps"][step] for r in runs) / 1000
                                   for step in steps}
        entry["second_steps_us"] = {step: statistics.median(r["rounds"][1]["steps"][step] for r in runs) / 1000
                                    for step in steps}
    summary[name] = entry
json.dump({"variants": [v[:4] for v in variants], "rounds": rounds, "summary": summary,
           "raw": records}, open(out, "w"), indent=1)
for name, entry in summary.items():
    print(name, json.dumps({k: v for k, v in entry.items()}))
