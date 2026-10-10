#!/usr/bin/env python3
"""Summarize paired ABBA blocks; confidence bounds are conditional on these runs."""
import json
import random
import statistics
import sys
from pathlib import Path


def percentile(values, q):
    values = sorted(values)
    point = (len(values) - 1) * q
    lower = int(point)
    upper = min(lower + 1, len(values) - 1)
    return values[lower] + (values[upper] - values[lower]) * (point - lower)


def summarize(directory):
    rows = []
    hashes = set()
    for version in ("control", "candidate"):
        runs = [json.loads(p.read_text()) for p in sorted(directory.glob(f"{version}-*.json"))]
        if len(runs) < 2:
            raise ValueError(f"two independent invocations required: {version}")
        for run in runs:
            hashes.add(run["corpus_sha256"])
            assert run["byte_equivalence"] and run["order"] == "ABBA"
        for route in runs[0]["routes"]:
            name = route["unified_route"]
            for template in route["measurements"]:
                count = template["operations"]
                blocks = []
                for run in runs:
                    found = next(r for r in run["routes"] if r["unified_route"] == name)
                    item = next(m for m in found["measurements"] if m["operations"] == count)
                    assert len(item["leaf_ns"]) == len(item["unified_ns"]) == 2 * run["rounds"]
                    for i in range(run["rounds"]):
                        blocks.append((item["leaf_ns"][2*i:2*i+2], item["unified_ns"][2*i:2*i+2]))
                leaf = [v for a, _ in blocks for v in a]
                unified = [v for _, b in blocks for v in b]
                rng = random.Random(0)
                ratios = []
                for _ in range(2000):
                    sample = [rng.choice(blocks) for _ in blocks]
                    ratios.append(statistics.median(v for _, b in sample for v in b) /
                                  statistics.median(v for a, _ in sample for v in a))
                rows.append({
                    "version": version, "route": name, "operations": count,
                    "samples_per_leg": len(leaf), "native_commits": template["native_commits"],
                    "leaf_median_ns": statistics.median(leaf),
                    "unified_median_ns": statistics.median(unified),
                    "unified_p95_ns": percentile(unified, .95),
                    "unified_p99_ns": percentile(unified, .99),
                    "unified_over_leaf": statistics.median(unified) / statistics.median(leaf),
                    "ratio_block_bootstrap_95_percent": [percentile(ratios, .025), percentile(ratios, .975)],
                })
    assert len(hashes) == 1, "corpus drift"
    return {"schema": 1, "corpus_sha256": hashes.pop(),
            "confidence_scope": "paired ABBA-block bootstrap; does not capture other-host or workload uncertainty",
            "measurements": rows}


if __name__ == "__main__":
    print(json.dumps(summarize(Path(sys.argv[1])), indent=2))
