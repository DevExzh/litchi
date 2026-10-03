#!/usr/bin/env python3
"""Offline allocation-site attribution; no workload or timing claim."""
from pathlib import Path
from collections import Counter
import argparse, hashlib, json, re

P = Path(__file__).resolve().parent
MCE = "litchi_ooxml_common::mce::stream::"
RUNNER = "litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan::"
GUARD = "litchi_perf_baseline::producer_shape::force_full_cell_worksheet_tolerance_reports::"

def descriptor(path):
    return {"path": str(path.resolve()), "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}

def analyze(path):
    regions = Counter()
    sites = {}
    total = 0
    unresolved = 0
    for line in path.read_text().splitlines():
        stack, cost = line.rsplit(None, 1)
        assert cost.isdecimal() and int(cost) > 0
        cost = int(cost)
        frames = stack.removesuffix(";").split(";")
        assert all(frames)
        # The timed closure is mapped in producer_shape.rs:2522-2570.
        # Presence is a stack attribution, not a timer or native latency metric.
        closure = any(RUNNER + "_$u7b$$u7b$closure$u7d$$u7d$::" in f for f in frames)
        guard = any(GUARD in f for f in frames)
        assert not (closure and guard)
        region = ("scan_closure" if closure else "forced_worksheet_guard" if guard
                  else "runner_other" if any(RUNNER in f for f in frames) else "process_other")
        regions[region] += cost
        total += cost
        if any(re.search(r"(?:^0x[0-9a-f]+|\?\?|<unknown>)", f) for f in frames):
            unresolved += cost
        nearest = next((f for f in reversed(frames) if MCE in f), None)
        if nearest:
            nearest = re.sub(r"::h[0-9a-f]{16}(?= |$)", "", nearest)
            sites.setdefault(region, Counter())[nearest] += cost
    return {"input": descriptor(path), "total": total, "regions": dict(regions),
            "unresolved_frame_allocation_events": unresolved,
            "nearest_mce_symbol_by_region": {r: dict(c.most_common()) for r,c in sites.items()}}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--export-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = {"schema": "litchi.heaptrack-site-attribution-0895.v1", "helper": descriptor(Path(__file__)),
              "scope": "whole-process disjoint stack-presence regions; MCE symbols inclusive of descendant work; no temporal segmentation",
              "performance_claim": None, "profiles": {}}
    for i, role in enumerate(("baseline", "candidate", "candidate", "baseline")):
        key = f"{i:02d}-{role}"
        result["profiles"][key] = analyze(args.export_root / "runs" / key / "stacks.txt")
    with args.output.open("x") as out:
        json.dump(result, out, indent=2, sort_keys=True); out.write("\n")
    print("Attribution written:", args.output)
if __name__ == "__main__":
    main()
