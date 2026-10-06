#!/usr/bin/env python3
"""Pairwise supplement for the validated formal v4 reduction.

The v4 reduction is retained as a descriptive per-cell aggregate.  This
supplement applies the frozen baseline/candidate rule to every same-cell,
same-block pair: 36 native pairs and 12 allocator pairs.  It consumes only
the reducer's already validated ``individual_rows`` and starts no workload.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import math
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DEFAULT_REDUCTION = ROOT.parent / "reductions/xlsx-formal-v1.json"
DEFAULT_PLAN = ROOT / "xlsx-formal-authorized-v4.json"


def read(path: Path) -> object:
    return json.loads(path.read_text())


def midpoint(values: list[int | float]) -> int | float:
    ordered = sorted(values)
    if len(ordered) % 2:
        return ordered[len(ordered) // 2]
    return (ordered[len(ordered) // 2 - 1] + ordered[len(ordered) // 2]) / 2


def nearest_rank(values: list[int | float], percentile: float) -> int | float:
    ordered = sorted(values)
    return ordered[max(1, math.ceil(percentile * len(ordered))) - 1]


def spread(values: list[int | float]) -> float | None:
    low, high = min(values), max(values)
    if low == 0:
        return 0.0 if high == 0 else None
    return (high - low) / low


def ratio(candidate: int | float | None, baseline: int | float | None) -> float | None:
    if candidate is None or baseline is None:
        return None
    if baseline == 0:
        return 1.0 if candidate == 0 else None
    return candidate / baseline


def flagged(value: float | None) -> bool | None:
    return None if value is None else value > 1.05


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def artifact(path: Path) -> dict:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def summarize_child(row: dict, metric: str) -> dict:
    values = row["sample_values"][metric]
    return {"samples": values, "p50": midpoint(values), "p95": nearest_rank(values, 0.95), "p99": nearest_rank(values, 0.99), "spread": spread(values)}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reduction", type=Path, default=DEFAULT_REDUCTION)
    parser.add_argument("--plan", type=Path, default=DEFAULT_PLAN)
    parser.add_argument("--json-output", type=Path, default=DEFAULT_REDUCTION.with_name("xlsx-formal-v1-pairwise-v5.json"))
    parser.add_argument("--text-output", type=Path, default=DEFAULT_REDUCTION.with_name("xlsx-formal-v1-pairwise-v5.txt"))
    args = parser.parse_args()
    reduction = read(args.reduction)
    plan = read(args.plan)
    if reduction.get("schema") != "litchi.performance.0941.xlsx-reduction.v4" or reduction.get("children") != 96:
        raise SystemExit("pairwise supplement refused: validated v4 reduction required")
    rows = reduction["individual_rows"]
    by_key = {(row["lane"], row["block"], row["cell"], row["leg"]): row for row in rows}
    expected = []
    for lane in ("native", "allocator"):
        for cell in plan["cells"]:
            for block in range(plan["scope"][lane]["blocks"]):
                expected.append((lane, block, cell["id"], "baseline"))
                expected.append((lane, block, cell["id"], "candidate"))
    if sorted(by_key) != sorted(expected) or len(by_key) != 96:
        raise SystemExit("pairwise supplement refused: individual row matrix changed")
    metrics = {"native": ("elapsed_ns", "rss_kib"), "allocator": ("elapsed_ns", "rss_kib", "live_bytes", "region_peak_live_bytes")}
    pairs = []
    for lane in ("native", "allocator"):
        for cell in plan["cells"]:
            for block in range(plan["scope"][lane]["blocks"]):
                pair_metrics = {}
                for metric in metrics[lane]:
                    baseline = summarize_child(by_key[(lane, block, cell["id"], "baseline")], metric)
                    candidate = summarize_child(by_key[(lane, block, cell["id"], "candidate")], metric)
                    ratios = {dimension: ratio(candidate[dimension], baseline[dimension]) for dimension in ("p50", "p95", "p99", "spread")}
                    flags = {dimension: flagged(ratios[dimension]) for dimension in ratios}
                    gate = "informational" if lane == "allocator" and metric == "elapsed_ns" else ("adoption_and_review" if lane == "native" and metric == "elapsed_ns" else "review")
                    pair_metrics[metric] = {"gate": gate, "baseline": baseline, "candidate": candidate, "candidate_over_baseline": ratios, "over_5pct_flags": flags, "hard_native_p50_adoption_reject": gate == "adoption_and_review" and flags["p50"] is True}
                pairs.append({"lane": lane, "cell": cell["id"], "block": block, "metrics": pair_metrics})
    hard_failures = [pair for pair in pairs if pair["lane"] == "native" and pair["metrics"]["elapsed_ns"]["hard_native_p50_adoption_reject"]]
    native_flags = [
        (pair, metric, dimension)
        for pair in pairs if pair["lane"] == "native"
        for metric in ("elapsed_ns", "rss_kib")
        for dimension, value in pair["metrics"][metric]["over_5pct_flags"].items()
        if value is True
    ]
    allocator_memory_flags = [
        (pair, metric, dimension)
        for pair in pairs if pair["lane"] == "allocator"
        for metric in ("rss_kib", "live_bytes", "region_peak_live_bytes")
        for dimension, value in pair["metrics"][metric]["over_5pct_flags"].items()
        if value is True
    ]
    counts = {"native_pairs": 36, "allocator_pairs": 12, "native_hard_p50_adoption_rejects": len(hard_failures), "native_review_flags": len(native_flags), "allocator_memory_review_flags": len(allocator_memory_flags), "allocator_rss_review_flags": sum(1 for pair, metric, dimension in allocator_memory_flags if metric == "rss_kib"), "allocator_retained_live_review_flags": sum(1 for pair, metric, dimension in allocator_memory_flags if metric == "live_bytes"), "allocator_region_peak_review_flags": sum(1 for pair, metric, dimension in allocator_memory_flags if metric == "region_peak_live_bytes")}
    result = {"schema": "litchi.performance.0941.xlsx-pairwise-supplement.v5", "created_utc": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"), "reduction": artifact(args.reduction), "plan": artifact(args.plan), "counts": counts, "pairs": pairs, "plan_gates": plan["gates"], "status": "validated_pairwise_supplement"}
    args.json_output.parent.mkdir(parents=True, exist_ok=True)
    args.json_output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    lines = ["lane | cell | block | timing p50 ratio | RSS p50 ratio | live p50 ratio | region p50 ratio | hard native p50", "-"]
    for pair in pairs:
        timing = pair["metrics"]["elapsed_ns"]
        rss = pair["metrics"]["rss_kib"]
        live = pair["metrics"].get("live_bytes", {})
        region = pair["metrics"].get("region_peak_live_bytes", {})
        lines.append(" | ".join([pair["lane"], pair["cell"], str(pair["block"]), f"{timing['candidate_over_baseline']['p50']:.4f}", f"{rss['candidate_over_baseline']['p50']:.4f}", f"{live.get('candidate_over_baseline', {}).get('p50', 'n/a')}", f"{region.get('candidate_over_baseline', {}).get('p50', 'n/a')}", str(timing["hard_native_p50_adoption_reject"])]))
    lines.extend(["", json.dumps(counts, sort_keys=True)])
    args.text_output.write_text("\n".join(lines) + "\n")
    print(json.dumps({"status": result["status"], "counts": counts, "json": str(args.json_output), "text": str(args.text_output)}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
