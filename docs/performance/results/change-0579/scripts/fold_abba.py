#!/usr/bin/env python3
"""Fold an A/B/B/A capture into paired medians and a same-binary noise floor."""
import json, pathlib, statistics, sys

BASE = pathlib.Path(sys.argv[1])


def quantiles(path):
    data = json.loads(path.read_text())
    values = sorted(record["elapsed_ns"] for record in data["records"])
    counters = {(r["metrics"]["read_calls"], r["metrics"]["read_bytes"],
                 r["metrics"]["version_calls"]) for r in data["records"]}
    def q(p):
        return values[min(len(values) - 1, int(p * len(values)))]
    return {"p50": statistics.median(values), "p95": q(0.95), "p99": q(0.99),
            "n": len(values), "counters": sorted(counters), "binary": data.get("binary")}


rows = []
for stem in ("flagship", "cv", "54016"):
    for mode in ("owned-readat", "file-source"):
        legs = {tag: quantiles(BASE / f"{stem}-{mode}-{tag}.json") for tag in ("A1", "B1", "B2", "A2")}
        counters = {tuple(map(tuple, legs[tag]["counters"])) for tag in legs}
        rows.append({
            "fixture": stem, "mode": mode,
            "before_p50": (legs["A1"]["p50"] + legs["A2"]["p50"]) / 2,
            "after_p50": (legs["B1"]["p50"] + legs["B2"]["p50"]) / 2,
            "dir1": 100 * (legs["B1"]["p50"] - legs["A1"]["p50"]) / legs["A1"]["p50"],
            "dir2": 100 * (legs["B2"]["p50"] - legs["A2"]["p50"]) / legs["A2"]["p50"],
            "AA": 100 * (legs["A2"]["p50"] - legs["A1"]["p50"]) / legs["A1"]["p50"],
            "BB": 100 * (legs["B2"]["p50"] - legs["B1"]["p50"]) / legs["B1"]["p50"],
            "p99_dir1": 100 * (legs["B1"]["p99"] - legs["A1"]["p99"]) / legs["A1"]["p99"],
            "counters_identical": len(counters) == 1,
            "legs": {tag: {k: v for k, v in legs[tag].items() if k != "counters"} for tag in legs},
            "counter_values": sorted(counters)[0] if len(counters) == 1 else sorted(counters),
        })
print(json.dumps(rows, indent=1))
