#!/usr/bin/env python3
"""Summarize rtf_allocator_check.py: median of per-child p50s per mode/leg/row.

usage: rtf_allocator_summary.py OUT_DIR > summary.json
"""
import gzip
import json
import os
import statistics
import sys
from collections import defaultdict

root = sys.argv[1]
children = [json.loads(line) for line in open(os.path.join(root, "children.jsonl"))]
rows = defaultdict(list)
for child in children:
    if child["exit_code"] != 0:
        continue
    stem = f"B{child['block']}-{child['mode']}-{child['position']}{child['leg']}-rtf.json"
    path = os.path.join(root, "raw", stem)
    opener = gzip.open if os.path.exists(path + ".gz") else open
    report = json.load(opener(path + (".gz" if opener is gzip.open else ""), "rt"))
    for result in report["results"]:
        samples = sorted(result["elapsed_ns"]["samples"])
        n = len(samples)
        p50 = (samples[(n - 1) // 2] + samples[n // 2]) / 2
        case = result["case"].replace("_unified_", "_").replace("_leaf_", "_")
        rows[(child["mode"], case, result["corpus"]["name"], child["leg"])].append(p50)
out = []
for (mode, case, corpus, leg), values in sorted(rows.items()):
    out.append({"mode": mode, "route": case, "corpus": corpus, "leg": leg,
                "children": len(values), "median_child_p50_us": statistics.median(values) / 1000,
                "child_p50_us": [round(v / 1000, 1) for v in values]})
json.dump({"children": len(children), "failed": sum(c["exit_code"] != 0 for c in children),
           "rows": out}, sys.stdout, indent=1)
