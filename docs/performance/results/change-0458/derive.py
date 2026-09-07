#!/usr/bin/env python3
"""Recompute phase attribution and instrumentation overhead from raw rows."""
import hashlib
import json
import math
from pathlib import Path
import random
import re
import statistics

import oracle

ROOT = Path(__file__).resolve().parent
PHASES = ["snapshot_open", "transaction", "add", "commit", "publication"]


def stats(values):
    values = sorted(values)
    return {"count": len(values), "min": values[0], "p50": statistics.median(values),
            "p95": values[math.ceil(len(values) * .95) - 1], "p99": values[math.ceil(len(values) * .99) - 1],
            "max": values[-1], "mean": statistics.mean(values)}


def allocation(rows):
    if not rows or rows[0] is None:
        return None
    return {"allocated_bytes": stats([r["allocated_bytes"] for r in rows]),
            "allocation_calls": stats([r["allocation_calls"] for r in rows]),
            "retained_live_delta_bytes": stats([r["live_bytes_after"] - r["live_bytes_before"] for r in rows]),
            "peak_above_phase_entry_bytes": stats([r["region_peak_live_bytes"] - r["live_bytes_before"] for r in rows])}


def bootstrap(before, after, seed):
    rng = random.Random(seed)
    values = []
    for _ in range(10000):
        left = statistics.median(rng.choices(before, k=len(before)))
        right = statistics.median(rng.choices(after, k=len(after)))
        values.append((right / left - 1) * 100)
    values.sort()
    return {"low": values[249], "high": values[9749], "resamples": 10000, "seed": seed,
            "method": "independent row bootstrap of median ratio; Python Random.choices; nearest-rank 2.5/97.5 percentiles; no multiple-comparison correction"}


def derive():
    protocol = json.loads((ROOT / "protocol.json").read_text())
    lanes = []
    reports = {}
    resources = {}
    for lane in protocol["order"]:
        path = ROOT / "runs" / lane["id"] / "report.json"
        oracle.validate(path, lane, protocol)
        value = json.loads(path.read_text())
        reports[(lane["repeat"], lane["instrumentation"], lane["shape"], lane["scope"])] = value
        resource_text = path.with_name("resource.log").read_text()
        rss = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", resource_text)
        assert rss, "missing GNU time maximum RSS"
        resources[(lane["repeat"], lane["instrumentation"], lane["shape"], lane["scope"])] = int(rss.group(1))
        rows = value["rows"]
        result = {"lane": lane, "report_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                  "lifecycle_ns": stats([r["lifecycle_ns"] for r in rows]),
                  "process_lifetime_maxrss_kib": int(rss.group(1))}
        if lane["scope"] == "lifecycle":
            result["allocation"] = allocation([r.get("lifecycle_allocation_metrics") for r in rows])
        else:
            result["phase_sum_ns"] = stats([r["phase_sum_ns"] for r in rows])
            result["boundary_gap_ns"] = stats([r["boundary_gap_ns"] for r in rows])
            result["phases"] = {}
            for i, name in enumerate(PHASES):
                phase = [r["phases"][i] for r in rows]
                result["phases"][name] = {"elapsed_ns": stats([r["elapsed_ns"] for r in phase]),
                    "paired_share_of_phase_sum_pct": stats([r["phases"][i]["elapsed_ns"] / r["phase_sum_ns"] * 100 for r in rows]),
                    "allocation": allocation([r.get("allocation_metrics") for r in phase])}
        lanes.append(result)
    overhead = []
    flags = []
    for repeat in ["R1", "R2"]:
        for instrumentation in ["normal", "allocator"]:
            for shape in ["tiny", "medium", "large"]:
                key = (repeat, instrumentation, shape)
                before = reports[(*key, "lifecycle")]["rows"]
                after = reports[(*key, "phases")]["rows"]
                b = stats([r["lifecycle_ns"] for r in before])
                a = stats([r["lifecycle_ns"] for r in after])
                delta = {q: (a[q] / b[q] - 1) * 100 for q in ["p50", "p95", "p99"]}
                entry = {"repeat": repeat, "instrumentation": instrumentation, "shape": shape,
                         "phase_envelope_vs_lifecycle_delta_pct": delta,
                         "p50_delta_95pct_interval": bootstrap([r["lifecycle_ns"] for r in before], [r["lifecycle_ns"] for r in after], 458 + len(overhead))}
                rss_delta = (resources[(*key, "phases")] / resources[(*key, "lifecycle")] - 1) * 100
                entry["process_lifetime_maxrss_delta_pct"] = rss_delta
                if instrumentation == "allocator":
                    direct = [r["lifecycle_allocation_metrics"]["allocated_bytes"] for r in before]
                    segmented = [sum(p["allocation_metrics"]["allocated_bytes"] for p in r["phases"]) for r in after]
                    entry["direct_allocated_bytes"] = stats(direct)
                    entry["sum_phase_allocated_bytes"] = stats(segmented)
                    entry["allocation_volume_matches_exactly"] = direct == segmented
                overhead.append(entry)
                for q, d in delta.items():
                    if abs(d) > 5:
                        flags.append({"repeat": repeat, "instrumentation": instrumentation, "shape": shape, "metric": q, "delta_pct": d})
                if abs(rss_delta) > 5:
                    flags.append({"repeat": repeat, "instrumentation": instrumentation, "shape": shape, "metric": "process_lifetime_maxrss_kib", "delta_pct": rss_delta})
    return {"schema": "litchi-0458-derived-v1", "change": 458, "reports": len(lanes), "samples": sum(row["lifecycle_ns"]["count"] for row in lanes),
            "quantiles": "median (midpoint for even sample counts); p95/p99 nearest rank; phase shares computed per paired row",
            "scope": "public API phase diagnostic and observed instrumentation overhead; no production speedup or internal commit-stage attribution",
            "lanes": lanes, "instrumentation_overhead": overhead, "absolute_5pct_review_flags": flags}


if __name__ == "__main__":
    print(json.dumps(derive(), indent=2, sort_keys=True))
