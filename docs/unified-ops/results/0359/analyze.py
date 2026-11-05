#!/usr/bin/env python3
"""Summarize the record 0359 A-B-B-A probe runs: per scenario, the median of
the per-run medians on each leg and after/before."""
import json, pathlib, statistics, sys

base = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else pathlib.Path(__file__).parent)
legs = {"before": {}, "after": {}}
for path in sorted(base.glob("probe-*-*.json")):
    leg = path.name.split("-")[1]
    for row in json.loads(path.read_text()):
        legs[leg].setdefault(row["scenario"], []).append(row["median_us"])
summary = []
for scenario, after in legs["after"].items():
    before = legs["before"].get(scenario)
    if not before:
        continue
    b, a = statistics.median(before), statistics.median(after)
    summary.append({"scenario": scenario, "before_us": round(b, 1), "after_us": round(a, 1),
                    "ratio": round(a / b, 3), "before_runs": before, "after_runs": after})
print(json.dumps(summary, indent=1))
