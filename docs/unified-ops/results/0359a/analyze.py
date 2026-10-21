#!/usr/bin/env python3
"""Summarize the record 0359a A-B-B-A runs: per probe and scenario, the
median of the per-run medians on each leg and after/before; the source
probe's after rows also against the base's plan-only XLS route; the DOC 1k
batch's unified/leaf ratio per leg; and whether each route's output SHA-256
is the same on both legs."""
import json, pathlib, statistics, sys

base = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else pathlib.Path(__file__).parent)
out = {}
for probe in ("source", "ops"):
    legs = {"before": {}, "after": {}}
    digests = {"before": {}, "after": {}}
    for path in sorted(base.glob(f"{probe}-*-*.json")):
        leg = path.name.split("-")[1]
        for row in json.loads(path.read_text()):
            legs[leg].setdefault(row["scenario"], []).append(row["median_us"])
            if "output_sha256" in row:
                digests[leg].setdefault(row["scenario"], set()).add(row["output_sha256"])
    rows = []
    for scenario, after in legs["after"].items():
        before = legs["before"].get(scenario)
        if not before:
            continue
        b, a = statistics.median(before), statistics.median(after)
        entry = {"scenario": scenario, "before_us": round(b, 1), "after_us": round(a, 1),
                 "ratio": round(a / b, 3), "before_runs": before, "after_runs": after}
        plan = legs["before"].get(scenario + "_plan")
        if plan:
            p = statistics.median(plan)
            entry["before_plan_us"] = round(p, 1)
            entry["ratio_vs_plan"] = round(a / p, 3)
        if scenario in digests["after"]:
            same = digests["after"][scenario] == digests["before"].get(scenario)
            entry["output_identical"] = same and len(digests["after"][scenario]) == 1
            entry["output_sha256"] = sorted(digests["after"][scenario])
        rows.append(entry)
    out[probe] = rows
batch = {}
for leg in ("before", "after"):
    vals = {}
    for path in sorted(base.glob(f"source-{leg}-*.json")):
        for row in json.loads(path.read_text()):
            vals.setdefault(row["scenario"], []).append(row["median_us"])
    if "doc/batch1k_ops" in vals and "doc/batch1k_leaf" in vals:
        o, l = statistics.median(vals["doc/batch1k_ops"]), statistics.median(vals["doc/batch1k_leaf"])
        batch[leg] = {"ops_us": round(o, 1), "leaf_us": round(l, 1), "ops_over_leaf": round(o / l, 3)}
out["doc_batch1k_unified_over_leaf"] = batch
rss = {}
log = base / "log.txt"
if log.exists():
    for line in log.read_text().splitlines():
        parts = line.split()
        if len(parts) >= 3 and parts[2].startswith("maxrss_kb="):
            rss.setdefault(parts[1], int(parts[2].split("=")[1]))
out["max_rss_kb"] = rss
print(json.dumps(out, indent=1))
