#!/usr/bin/env python3
"""Record 0360b: summarize one A-B-B-A run (abba.sh output directory).

Ratio = median of the two after-leg medians / median of the two before-leg
medians, per scenario. Writes summary.json and prints a Markdown table.
"""
import json
import statistics
import sys
from pathlib import Path

out = Path(sys.argv[1])
legs = {"before": [1, 4], "after": [2, 3]}


def probe(leg, n):
    return {r["scenario"]: r["median_us"] for r in json.loads((out / f"probe-{leg}-{n}.json").read_text())}


def sheet(leg, n):
    data = json.loads((out / f"sheet-{leg}-{n}.txt").read_text())
    return {k: v["median_us"] for k, v in data["scenarios"].items() if k.endswith("/xlsb")}


def crud(leg, n):
    data = json.loads((out / f"crud-{leg}-{n}.json").read_text())
    return {f"crud/{c['case']}": c["statistics"]["p50_ns"] / 1000 for c in data["cases"]}


rows = {}
for name, reader in [("probe", probe), ("sheet", sheet), ("crud", crud)]:
    per_leg = {}
    for leg, runs in legs.items():
        values = [reader(leg, n) for n in runs]
        per_leg[leg] = {k: statistics.median(v[k] for v in values) for k in values[0]}
    for key in per_leg["before"]:
        before, after = per_leg["before"][key], per_leg["after"].get(key)
        if after is None:
            continue
        rows[key] = {"before_us": before, "after_us": after, "ratio": after / before}
(out / "summary.json").write_text(json.dumps(rows, indent=2) + "\n")
print("| Scenario | Before (µs) | After (µs) | After/before |")
print("|---|---:|---:|---:|")
for key, row in rows.items():
    print(f"| `{key}` | {row['before_us']:.1f} | {row['after_us']:.1f} | {row['ratio']:.3f} |")
