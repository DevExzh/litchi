#!/usr/bin/env python3
"""Summarize the tiny semantic control's dedicated runs (change 0751).

The runs were made with
``taskset -c 4 perf stat -e cycles,instructions -x , -- BIN --case
pptx_semantic_one_edit_save --semantic-shape tiny --samples N --warmup 10``
at N = 100 and N = 1100, in two ABBA rounds (``r<round>-s<slot>-<arm>-n<N>``).
Per-iteration counts are the difference of the two whole-child counts divided
by 1000; the paired p50 ratio pairs (s0 before, s1 after) and (s3 before,
s2 after) of the N = 1100 reports. ``fe-<arm>-*.csv`` are single whole-child
runs at N = 1100 with front-end counters, in the order before, after, after,
before.

Usage: tiny_semantic.py COUNTERS_DIR [--json OUT]
"""

from __future__ import annotations

import argparse
import collections
import gzip
import json
import re
import statistics
import sys
from pathlib import Path


def text(path: Path) -> str:
    raw = path.read_bytes()
    return gzip.decompress(raw).decode("utf-8") if path.suffix == ".gz" else raw.decode("utf-8")


def counters(path: Path) -> dict[str, float]:
    values: dict[str, float] = {}
    for line in text(path).splitlines():
        fields = line.split(",")
        if len(fields) > 2 and fields[0].strip():
            try:
                values[fields[2]] = float(fields[0])
            except ValueError:
                continue
    return values


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("counters", type=Path)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    runs = {}
    for path in args.counters.glob("r*-n*.perf.csv*"):
        match = re.search(r"r(\d)-s(\d)-(before|after)-n(\d+)", path.name)
        runs[(int(match[1]), int(match[2]), match[3], int(match[4]))] = counters(path)
    per_iteration = collections.defaultdict(list)
    for (round_index, slot, arm, samples), values in sorted(runs.items()):
        if samples != 1100:
            continue
        low = runs[(round_index, slot, arm, 100)]
        per_iteration[arm].append(
            {event: (values[event] - low[event]) / 1000 for event in ("cycles", "instructions")}
        )
    p50 = {}
    for path in args.counters.glob("r*-n1100.json*"):
        match = re.search(r"r(\d)-s(\d)-(before|after)-n", path.name)
        (result,) = json.loads(text(path))["results"]
        p50[(int(match[1]), int(match[2]), match[3])] = result["elapsed_ns"]["p50"] / 1e3
    pairs = [p50[(r, 1, "after")] / p50[(r, 0, "before")] for r in (0, 1)]
    pairs += [p50[(r, 2, "after")] / p50[(r, 3, "before")] for r in (0, 1)]
    frontend = collections.defaultdict(lambda: collections.defaultdict(list))
    for path in args.counters.glob("fe-*.csv*"):
        arm = "before" if "-before-" in path.name else "after"
        for event, value in counters(path).items():
            frontend[arm][event].append(value)
    report = {
        "per_iteration": dict(per_iteration),
        "median_per_iteration": {
            arm: {
                event: statistics.median(row[event] for row in rows)
                for event in ("cycles", "instructions")
            }
            for arm, rows in per_iteration.items()
        },
        "p50_us": {f"r{r}-s{s}-{arm}": value for (r, s, arm), value in sorted(p50.items())},
        "paired_p50_ratios": pairs,
        "median_paired_p50_ratio": statistics.median(pairs),
        "frontend_whole_child_median": {
            arm: {event: statistics.median(values) for event, values in events.items()}
            for arm, events in frontend.items()
        },
    }
    medians = report["median_per_iteration"]
    report["per_iteration_ratios"] = {
        event: medians["after"][event] / medians["before"][event]
        for event in ("cycles", "instructions")
    }
    output = json.dumps(report, indent=2, sort_keys=True)
    if args.json:
        args.json.write_text(output + "\n", encoding="utf-8")
    print(output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
