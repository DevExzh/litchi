#!/usr/bin/env python3
"""Per-operation instructions and cycles of each timed region, change 0743.

Each probe mode runs its region N times; `perf stat` counts user-mode
instructions and cycles for the whole process at two iteration counts, and
the per-operation value is (count(N2) - count(N1)) / (N2 - N1), so corpus
construction and after-loop verification cancel. The edit/save `cycle` modes
also build a fresh package per iteration, which the `setup` mode measures on
its own so it can be subtracted. Legs run in ABBA order, two blocks; each
reported value is the median of the leg's four differenced pairs.
"""

import argparse
import json
import os
import statistics
import subprocess

MODES = [
    ("fulltext-large", ["fulltext", "large"], 5, 25, None),
    ("fulltext-medium", ["fulltext", "medium"], 50, 550, None),
    ("open-large", ["open", "large"], 50, 550, None),
    ("open-medium", ["open", "medium"], 100, 1100, None),
    ("setup-large", ["setup", "large"], 10, 110, None),
    ("setup-medium", ["setup", "medium"], 50, 550, None),
    ("noop-large", ["cycle", "large"], 3, 13, "noop"),
    ("one-large", ["cycle", "large"], 3, 13, "one"),
    ("pct-large", ["cycle", "large"], 2, 7, "pct"),
    ("noop-medium", ["cycle", "medium"], 20, 220, "noop"),
    ("one-medium", ["cycle", "medium"], 20, 220, "one"),
]


def counters(binary, arguments, iterations, edits, core, out):
    command = ["perf", "stat", "-x,", "-e", "instructions:u,cycles:u", "-o", out,
               "taskset", "-c", core, binary, *arguments, str(iterations)]
    if edits:
        command.append(edits)
    subprocess.run(command, check=True, capture_output=True, text=True)
    values = {}
    with open(out, encoding="utf-8") as handle:
        for line in handle:
            fields = line.strip().split(",")
            if len(fields) > 2 and fields[0].replace(".", "").isdigit():
                values[fields[2]] = float(fields[0])
    return values


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--a", required=True)
    parser.add_argument("--b", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--core", default="8")
    args = parser.parse_args()
    os.makedirs(args.out, exist_ok=True)
    legs = {"a": args.a, "b": args.b}
    records = []
    for name, arguments, low, high, edits in MODES:
        for block in range(2):
            for position, leg in enumerate(["a", "b", "b", "a"]):
                pair = {}
                for iterations in (low, high):
                    out = os.path.join(args.out, f"{name}-b{block}-p{position}-{leg}-{iterations}.csv")
                    pair[iterations] = counters(legs[leg], arguments, iterations, edits, args.core, out)
                per_operation = {
                    event: (pair[high][event] - pair[low][event]) / (high - low)
                    for event in ("instructions:u", "cycles:u")
                }
                records.append({"mode": name, "block": block, "position": position,
                                "leg": leg, "per_operation": per_operation})
                print(name, block, position, leg, {k: round(v) for k, v in per_operation.items()}, flush=True)
    summary = {}
    for name, *_ in MODES:
        for leg in ("a", "b"):
            rows = [r["per_operation"] for r in records if r["mode"] == name and r["leg"] == leg]
            summary.setdefault(name, {})[leg] = {
                event: statistics.median(row[event] for row in rows)
                for event in ("instructions:u", "cycles:u")
            }
    with open(os.path.join(args.out, "probe-counters.json"), "w", encoding="utf-8") as handle:
        json.dump({"records": records, "summary": summary}, handle, indent=1)


if __name__ == "__main__":
    main()
