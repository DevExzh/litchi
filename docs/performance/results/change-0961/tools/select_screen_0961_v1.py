#!/usr/bin/env python3
"""Apply the 0961 screen rule; this never authorizes adoption."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
from typing import Any

CASES = {"medium-one-percent", "dense-sparse-one-percent", "medium-one-edit"}
PRIMARY = ("edit_commit_ns.p50", "edit_commit_ns.p95", "total_ns.p50", "total_ns.p95")
SIGNAL = 0.95
SAFETY = 1.05


def descriptor(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    return {"path": str(path.resolve()), "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}


def select(report: dict[str, Any]) -> dict[str, Any]:
    if report.get("status") != "screened":
        raise ValueError("reduced report is not screened")
    stops: list[str] = []
    signals: list[dict[str, Any]] = []
    observed: dict[str, Any] = {}
    for lane, expected in (("native", 6), ("allocator", 3)):
        cases = report.get(lane, {}).get("cases", {})
        if set(cases) != CASES:
            raise ValueError(f"{lane} case set changed")
        observed[lane] = {}
        required = (*PRIMARY, "max_rss_kib") if lane == "native" else (
            "allocated_bytes", "allocation_calls", "region_peak_live_bytes",
            "core_combined_with_drain.allocated_bytes", "max_rss_kib")
        for case, value in cases.items():
            metrics = value.get("metrics", {})
            observed[lane][case] = {}
            for key in required:
                metric = metrics.get(key, {})
                ratio = metric.get("median")
                valid = (metric.get("pair_count") == expected and metric.get("undefined_count") == 0
                         and type(ratio) in (int, float) and math.isfinite(ratio) and ratio > 0)
                observed[lane][case][key] = metric
                label = f"{lane}/{case}/{key}"
                if not valid:
                    stops.append(label + ": invalid or undefined")
                    continue
                if ratio <= SIGNAL and (lane == "native" or key == "region_peak_live_bytes"):
                    signals.append({"metric": label, "median": ratio})
                if ratio > SAFETY and (lane == "native" and (key in PRIMARY or key == "max_rss_kib")
                                       or lane == "allocator" and key in {"region_peak_live_bytes", "max_rss_kib"}):
                    stops.append(label + ": safety median exceeds 1.05")
    if not signals:
        stops.append("No qualifying metric among native primary/native RSS/allocator region peak")
    return {
        "decision": "reject_screen" if stops else "require_full_confirmation_and_regression_review",
        "adoption_allowed": False,
        "statistical_effect_claim": False,
        "signal_threshold": SIGNAL,
        "safety_review_threshold": SAFETY,
        "stops": stops,
        "signals": signals,
        "required_metrics": observed,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--report", required=True, type=Path)
    ap.add_argument("--output", required=True, type=Path)
    args = ap.parse_args()
    result = select(json.loads(args.report.read_text(encoding="utf-8")))
    result.update(schema="litchi.performance.0961.root-screen-selection.v1",
                  report=descriptor(args.report), helper=descriptor(Path(__file__)))
    with args.output.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write("\n")
    print(result["decision"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
