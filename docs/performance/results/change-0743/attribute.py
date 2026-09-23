#!/usr/bin/env python3
"""Per-commit attribution for change 0743.

The retained probe (probe/src/main.rs) builds the harness's own semantic PPTX
corpus and times one region per mode with the same public calls the harness
makes; verification runs after the clocks. One probe binary per commit is run
in rotating order, each process pinned to one core, for several rounds. The
reported value per (commit, mode) is the median of the per-process medians.
"""

import argparse
import json
import os
import re
import statistics
import subprocess

MODES = [
    ("fulltext", ["fulltext", "large", "15"]),
    ("fulltext-medium", ["fulltext", "medium", "300"]),
    ("noop", ["cycle", "large", "7", "noop"]),
    ("one", ["cycle", "large", "7", "one"]),
    ("pct", ["cycle", "large", "3", "pct"]),
    ("one-medium", ["cycle", "medium", "150", "one"]),
]


def run(binary, arguments, core):
    completed = subprocess.run(
        ["taskset", "-c", core, binary, *arguments],
        capture_output=True,
        text=True,
        check=True,
    )
    values = {}
    for line in completed.stdout.splitlines():
        match = re.match(r"^(\S+)\s+median_ns (\d+)$", line.strip())
        if match:
            values[match.group(1)] = int(match.group(2))
    return values


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--binaries", required=True, help="directory of <index>-<commit> probes")
    parser.add_argument("--out", required=True)
    parser.add_argument("--rounds", type=int, default=3)
    parser.add_argument("--core", default="8")
    args = parser.parse_args()
    binaries = sorted(os.listdir(args.binaries))
    records = []
    for round_index in range(args.rounds):
        order = binaries if round_index % 2 == 0 else list(reversed(binaries))
        for mode, arguments in MODES:
            for name in order:
                values = run(os.path.join(args.binaries, name), arguments, args.core)
                records.append(
                    {"round": round_index, "binary": name, "mode": mode, "values": values}
                )
                print(round_index, mode, name, values, flush=True)
    summary = {}
    for name in binaries:
        for mode, _ in MODES:
            rows = [r["values"] for r in records if r["binary"] == name and r["mode"] == mode]
            keys = sorted({key for row in rows for key in row})
            summary.setdefault(name, {})[mode] = {
                key: statistics.median(row[key] for row in rows if key in row) for key in keys
            }
    with open(os.path.join(args.out, "attribution.json"), "w", encoding="utf-8") as handle:
        json.dump({"records": records, "summary": summary}, handle, indent=1)


if __name__ == "__main__":
    main()
