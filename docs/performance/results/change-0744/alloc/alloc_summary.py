#!/usr/bin/env python3
"""Summarize operation-region allocator metrics (A = before, B = after)."""
import json
import sys

RAW = sys.argv[1] if len(sys.argv) > 1 else "alloc"
OUT = sys.argv[2] if len(sys.argv) > 2 else "alloc-summary.json"
METRICS = ["allocation_calls", "deallocation_calls", "reallocation_calls", "allocated_bytes", "region_peak_live_bytes"]


def load(slot):
    data = json.load(open(f"{RAW}/alloc-{slot}.json"))
    rows = {}
    for result in data["results"]:
        allocation = (result.get("operation_metrics") or {}).get("allocation") or {}
        if allocation.get("status") != "measured":
            continue
        rows[(result["case"], result["corpus"]["shape"])] = {
            metric: allocation[metric]["values"] for metric in METRICS
        }
    return rows


slots = {slot: load(slot) for slot in ["A1", "B1", "B2", "A2"]}
summary = []
for key in sorted(slots["A1"]):
    entry = {"case": key[0], "shape": key[1]}
    for metric in METRICS:
        before = {value for slot in ("A1", "A2") for value in slots[slot][key][metric]}
        after = {value for slot in ("B1", "B2") for value in slots[slot][key][metric]}
        entry[metric] = {
            "before": sorted(before),
            "after": sorted(after),
            "deterministic": len(before) == 1 and len(after) == 1,
            "change": (min(after) / min(before) - 1) if min(before) else None,
        }
    summary.append(entry)
json.dump(summary, open(OUT, "w"), indent=1)
for entry in summary:
    for metric in METRICS:
        value = entry[metric]
        change = value["change"]
        print(
            f"{entry['case']:30s} {entry['shape']:11s} {metric:24s} "
            f"{value['before'][0]:>12,} -> {value['after'][0]:>12,} "
            f"{'' if change is None else f'{change*100:+.2f}%':>9s} det={value['deterministic']}"
        )
