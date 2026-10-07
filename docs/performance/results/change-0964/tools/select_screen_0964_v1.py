#!/usr/bin/env python3
"""Apply the frozen 0964 screen rule; this never authorizes adoption."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
from typing import Any

CASES = ("medium-one-percent", "dense-sparse-one-percent", "medium-one-edit")
PRIMARY = ("edit_commit_ns.p50", "edit_commit_ns.p95", "total_ns.p50", "total_ns.p95")
TAIL = ("ingress_ns.p99", "edit_commit_ns.p99", "publication_ns.p99", "commit_drop_ns.p99", "total_ns.p99")
NATIVE_REQUIRED = PRIMARY + TAIL + ("max_rss_kib",)
ALLOCATOR_REQUIRED = (
    "allocated_bytes",
    "allocation_calls",
    "region_peak_live_bytes",
    "core_combined_with_drain.allocated_bytes",
    "max_rss_kib",
)
SIGNAL = 0.95
SAFETY = 1.05


def descriptor(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    stat = path.stat()
    return {
        "path": str(path.resolve()),
        "bytes": stat.st_size,
        "mode": stat.st_mode & 0o777,
        "executable": bool(stat.st_mode & 0o111),
        "sha256": hashlib.sha256(data).hexdigest(),
    }


def valid_metric(metric: Any, pair_count: int) -> bool:
    if not isinstance(metric, dict):
        return False
    value = metric.get("median")
    return (
        metric.get("pair_count") == pair_count
        and metric.get("undefined_count") == 0
        and type(value) in (int, float)
        and math.isfinite(float(value))
        and float(value) > 0
    )


def select(report: dict[str, Any]) -> dict[str, Any]:
    if report.get("status") != "screened":
        raise ValueError("reduced report is not screened")

    stops: list[str] = []
    signals: list[dict[str, Any]] = []
    observed: dict[str, dict[str, dict[str, Any]]] = {}
    for lane, expected_pairs, required in (
        ("native", 6, NATIVE_REQUIRED),
        ("allocator", 3, ALLOCATOR_REQUIRED),
    ):
        lane_value = report.get(lane)
        cases = lane_value.get("cases", {}) if isinstance(lane_value, dict) else {}
        if not isinstance(cases, dict) or set(cases) != set(CASES):
            raise ValueError(f"{lane} case set changed")
        observed[lane] = {}
        for case in CASES:
            metrics = cases[case].get("metrics", {})
            if not isinstance(metrics, dict):
                raise ValueError(f"{lane}/{case}: metrics missing")
            observed[lane][case] = {}
            for key in required:
                metric = metrics.get(key)
                observed[lane][case][key] = metric
                label = f"{lane}/{case}/{key}"
                if not valid_metric(metric, expected_pairs):
                    stops.append(label + ": invalid or undefined")
                    continue
                ratio = float(metric["median"])
                hard_safety = (
                    lane == "native" and (key in PRIMARY or key in TAIL or key == "max_rss_kib")
                ) or (
                    lane == "allocator" and key in {"region_peak_live_bytes", "max_rss_kib"}
                )
                if hard_safety and ratio > SAFETY:
                    stops.append(label + ": safety median exceeds 1.05")
                if (
                    ratio <= SIGNAL
                    and (lane == "native" and (key in PRIMARY or key == "max_rss_kib")
                         or lane == "allocator" and key == "region_peak_live_bytes")
                ):
                    signals.append({"metric": label, "median": ratio, "threshold": SIGNAL})

    if not signals:
        stops.append("No qualifying metric among native primary/native RSS/allocator region peak")
    return {
        "decision": "reject_screen" if stops else "require_full_confirmation_and_regression_review",
        "adoption_allowed": False,
        "statistical_effect_claim": False,
        "signal_threshold": SIGNAL,
        "safety_review_threshold": SAFETY,
        "signal_comparator": "candidate_over_baseline <= threshold",
        "safety_comparator": "candidate_over_baseline > threshold",
        "stops": stops,
        "signals": signals,
        "required_metrics": observed,
        "requested_bytes_or_allocation_calls_signal": False,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    result = select(json.loads(args.report.read_text(encoding="utf-8")))
    result.update(
        {
            "schema": "litchi.performance.0964.root-screen-selection.v1",
            "report": descriptor(args.report),
            "helper": descriptor(Path(__file__)),
        }
    )
    if args.output.exists() or args.output.is_symlink():
        raise RuntimeError(f"refusing to overwrite selection: {args.output}")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(result["decision"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
