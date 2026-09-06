#!/usr/bin/env python3
"""Derive the reproducible, descriptive 0434 ABBA summary.

The command first runs the portable matrix verifier, then summarizes retained
normal latency and allocator vectors.  It emits no candidate-speedup claim;
the exact cross-phase identity gate is the prerequisite for any later
interpretation.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
from pathlib import Path
import random
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent


class SummaryError(ValueError):
    pass


def load_module(filename: str, name: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / filename)
    if spec is None or spec.loader is None:
        raise SummaryError(f"cannot load {filename}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


verify = load_module("verify.py", "change0434_verify_for_summary")


def finite(value: Any, label: str, *, signed: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SummaryError(f"{label}: expected a number")
    number = float(value)
    if not math.isfinite(number) or (not signed and number < 0):
        raise SummaryError(f"{label}: expected a finite number")
    return number


def stats(values: list[int | float], label: str, *, signed: bool = False) -> dict[str, Any]:
    if not values:
        raise SummaryError(f"{label}: empty vector")
    ordered = sorted(finite(value, f"{label}[{i}]", signed=signed) for i, value in enumerate(values))
    count = len(ordered)
    median = (ordered[(count - 1) // 2] + ordered[count // 2]) / 2.0
    p95 = ordered[min(((95 * count + 99) // 100) - 1, count - 1)]
    p99 = ordered[min(((99 * count + 99) // 100) - 1, count - 1)]
    result: dict[str, Any] = {
        "count": count,
        "min": ordered[0], "p50": median, "p95": p95, "p99": p99,
        "max": ordered[-1], "mean": sum(ordered) / count,
    }
    for key in ("min", "p50", "p95", "p99", "max", "mean"):
        if isinstance(result[key], float) and result[key].is_integer():
            result[key] = int(result[key])
    return result


def bootstrap_p50(values: list[int], label: str) -> dict[str, Any]:
    """A fixed-seed percentile interval for the sample median.

    This is a descriptive uncertainty interval over the 30 timed samples
    collected inside one fresh process for a report, not an independent-run
    confidence claim.
    """

    if not values:
        raise SummaryError(f"{label}: empty bootstrap input")
    rng = random.Random(434)
    medians: list[float] = []
    for _ in range(2000):
        sample = sorted(values[rng.randrange(len(values))] for _ in values)
        n = len(sample)
        medians.append((sample[(n - 1) // 2] + sample[n // 2]) / 2.0)
    medians.sort()
    low = medians[int(0.025 * len(medians))]
    high = medians[int(0.975 * len(medians)) - 1]
    result: dict[str, Any] = {
        "method": "bootstrap_median_percentile",
        "resamples": 2000,
        "seed": 434,
        "level": 0.95,
        "low": low,
        "high": high,
    }
    for key in ("low", "high"):
        if isinstance(result[key], float) and result[key].is_integer():
            result[key] = int(result[key])
    return result


def metric_values(metric: Any, label: str) -> list[int] | None:
    if not isinstance(metric, dict):
        raise SummaryError(f"{label}: malformed metric")
    if metric.get("status") != "measured":
        return None
    values = metric.get("values")
    if not isinstance(values, list):
        raise SummaryError(f"{label}: measured metric has no values")
    return [verify.u64(value, f"{label}.values[{i}]") for i, value in enumerate(values)]


def row_summary(row: dict[str, Any]) -> dict[str, Any]:
    verified = row["verified"]
    result = verified["result"]
    lane = row["lane"]
    elapsed = [verify.u64(value, f"{row['name']}.elapsed[{i}]") for i, value in enumerate(verified["elapsed"])]
    elapsed_summary = stats(elapsed, f"{row['name']}.elapsed")
    if lane["mode"] == "normal":
        elapsed_summary["p50_ci"] = bootstrap_p50(elapsed, row["name"])
    metrics = verified["metrics"]
    process = metrics.get("process")
    process_peak = metric_values(process.get("peak_rss_bytes"), f"{row['name']}.process.peak_rss_bytes") if isinstance(process, dict) else None
    reported_peak_scope = process.get("peak_rss_bytes", {}).get("scope", "process_lifetime_high_water_after_not_operation_peak") if isinstance(process, dict) else "process_lifetime_high_water_after_not_operation_peak"
    process_summary: dict[str, Any] = {
        "whole_process_max_rss": row["resource"],
        "reported_process_peak_rss": {
            "scope": reported_peak_scope,
            "statistics": stats(process_peak, f"{row['name']}.reported_process_peak_rss") if process_peak is not None else {"status": "unavailable"},
        },
    }
    allocation_metric = metrics.get("allocation")
    if lane["mode"] == "normal" or allocation_metric is None or allocation_metric.get("status") != "measured":
        allocation: dict[str, Any] = {"status": "unavailable", "scope": "operation_global_system_allocator", "reason": "normal binary or unavailable allocator vectors"}
    else:
        vectors = {
            key: metric_values(allocation_metric.get(key), f"{row['name']}.allocation.{key}")
            for key in ("allocation_calls", "allocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes")
        }
        if any(value is None for value in vectors.values()):
            raise SummaryError(f"{row['name']}: allocator vectors are incomplete")
        assert all(value is not None for value in vectors.values())
        n = len(vectors["allocated_bytes"])
        if any(len(value) != n for value in vectors.values()):
            raise SummaryError(f"{row['name']}: allocator vectors have different lengths")
        regional_entry = [vectors["region_peak_live_bytes"][i] - vectors["live_bytes_before"][i] for i in range(n)]
        regional_exit = [vectors["region_peak_live_bytes"][i] - vectors["live_bytes_after"][i] for i in range(n)]
        allocation = {
            "status": "measured", "scope": allocation_metric.get("scope"),
            "requested_bytes": stats(vectors["allocated_bytes"], f"{row['name']}.allocated_bytes"),
            "requested_count": stats(vectors["allocation_calls"], f"{row['name']}.allocation_calls"),
            "regional_peak_minus_entry": stats(regional_entry, f"{row['name']}.regional_peak_minus_entry"),
            "regional_peak_minus_exit": stats(regional_exit, f"{row['name']}.regional_peak_minus_exit"),
            "raw_live_before": stats(vectors["live_bytes_before"], f"{row['name']}.live_before"),
            "raw_live_after": stats(vectors["live_bytes_after"], f"{row['name']}.live_after"),
        }
    identity = row["identity"]
    result_row = {
        "phase": row["phase"], "role": row["role"], "mode": lane["mode"], "shape": lane["shape"], "repeat": lane["repeat"], "name": row["name"],
        "report": str(row["report_path"]),
        "corpus": {key: identity[key] for key in ("archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256")},
        "output_sha256": identity["output_sha256"], "semantic_sha256": identity["semantic_sha256"],
        "sink": {key: identity[key] for key in ("sink_accepted_bytes", "sink_write_calls")},
        "latency_claim_scope": "normal operation elapsed vectors; 30 samples inside one fresh process per report" if lane["mode"] == "normal" else "descriptive allocator-process elapsed vectors; no allocator latency claim",
        "elapsed_ns": elapsed_summary,
        "process_memory": process_summary,
        "allocation": allocation,
        "source_identity": {
            "build_revision": row["build"]["revision"],
            "ambient_revision": identity["environment_revision"],
            "ambient_worktree_dirty": identity["environment_dirty"],
            "binary_sha256": row["build"]["binaries"][lane["mode"]]["sha256"],
            "source_manifest_sha256": row["build"]["source_manifest"]["sha256"],
        },
    }
    if lane["mode"] == "normal":
        result_row["throughput"] = {
            "archive_bytes_per_second": stats(
                [identity["archive_bytes"] * 1_000_000_000 / value for value in elapsed],
                f"{row['name']}.archive_bytes_per_second",
            )
        }
    return result_row


def relative(old: float, new: float, label: str) -> float:
    if old == 0:
        raise SummaryError(f"{label}: cannot compare a zero baseline")
    return (new - old) / old * 100.0


def repeat_flags(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    indexed = {(row["role"], row["mode"], row["shape"], row["repeat"]): row for row in rows}
    flags: list[dict[str, Any]] = []
    for role in ("before-streaming", "after-streaming"):
        for mode in ("normal", "allocator"):
            for shape in ("tiny", "medium", "large"):
                first = indexed[(role, mode, shape, "R1")]
                second = indexed[(role, mode, shape, "R2")]
                if mode == "normal":
                    metric_specs = (
                        ("elapsed_ns.p50", lambda row: row["elapsed_ns"]["p50"], "operation_samples_within_one_fresh_process"),
                        ("elapsed_ns.p95", lambda row: row["elapsed_ns"]["p95"], "operation_samples_within_one_fresh_process"),
                        ("elapsed_ns.p99", lambda row: row["elapsed_ns"]["p99"], "operation_samples_within_one_fresh_process"),
                        ("elapsed_ns.mean", lambda row: row["elapsed_ns"]["mean"], "operation_samples_within_one_fresh_process"),
                    )
                else:
                    metric_specs = (
                        ("allocation.requested_bytes.p50", lambda row: row["allocation"]["requested_bytes"]["p50"], "operation_samples_within_one_fresh_process"),
                        ("allocation.requested_count.p50", lambda row: row["allocation"]["requested_count"]["p50"], "operation_samples_within_one_fresh_process"),
                        ("allocation.regional_peak_minus_entry.p50", lambda row: row["allocation"]["regional_peak_minus_entry"]["p50"], "operation_samples_within_one_fresh_process"),
                    )
                metric_specs += (("whole_process_rss_bytes", lambda row: row["process_memory"]["whole_process_max_rss"]["bytes"], "gnu_time_v_verbose_whole_fresh_process"),)
                for metric, select, scope in metric_specs:
                    old = float(select(first)); new = float(select(second))
                    delta = relative(old, new, f"repeat.{role}.{mode}.{shape}.{metric}")
                    flags.append({"role": role, "mode": mode, "shape": shape, "metric": metric, "scope": scope, "r1": old, "r2": new, "relative_percent": delta, "threshold_percent": 5.0, "flagged": abs(delta) > 5.0})
    return flags


def matched_comparisons(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Compare each after role to its matched before role for each repeat.

    These are retained descriptive deltas.  The normal rows include operation
    latency and archive throughput; allocator rows include requested count,
    requested bytes, regional peak, and whole-process RSS only.  No allocator
    elapsed metric is compared.
    """

    indexed = {(row["role"], row["mode"], row["shape"], row["repeat"]): row for row in rows}
    comparisons: list[dict[str, Any]] = []
    for mode in ("normal", "allocator"):
        for shape in ("tiny", "medium", "large"):
            for repeat in ("R1", "R2"):
                before = indexed[("before-streaming", mode, shape, repeat)]
                after = indexed[("after-streaming", mode, shape, repeat)]
                if mode == "normal":
                    metric_specs = (
                        ("latency_ns.p50", lambda row: row["elapsed_ns"]["p50"], True),
                        ("latency_ns.p95", lambda row: row["elapsed_ns"]["p95"], True),
                        ("latency_ns.p99", lambda row: row["elapsed_ns"]["p99"], True),
                        ("latency_ns.mean", lambda row: row["elapsed_ns"]["mean"], True),
                        ("throughput.archive_bytes_per_second.p50", lambda row: row["throughput"]["archive_bytes_per_second"]["p50"], False),
                        ("whole_process_rss_bytes", lambda row: row["process_memory"]["whole_process_max_rss"]["bytes"], True),
                    )
                else:
                    metric_specs = (
                        ("allocation.requested_count.p50", lambda row: row["allocation"]["requested_count"]["p50"], True),
                        ("allocation.requested_bytes.p50", lambda row: row["allocation"]["requested_bytes"]["p50"], True),
                        ("allocation.regional_peak_minus_entry.p50", lambda row: row["allocation"]["regional_peak_minus_entry"]["p50"], True),
                        ("whole_process_rss_bytes", lambda row: row["process_memory"]["whole_process_max_rss"]["bytes"], True),
                    )
                for metric, select, higher_is_regression in metric_specs:
                    old = float(select(before)); new = float(select(after))
                    delta = relative(old, new, f"comparison.{mode}.{shape}.{repeat}.{metric}")
                    regression = delta > 5.0 if higher_is_regression else delta < -5.0
                    scope = "gnu_time_v_verbose_whole_fresh_process" if metric == "whole_process_rss_bytes" else "operation_samples_within_one_fresh_process"
                    comparisons.append({
                        "baseline_role": "before-streaming", "candidate_role": "after-streaming",
                        "mode": mode, "shape": shape, "repeat": repeat, "metric": metric,
                        "before": old, "after": new, "relative_percent": delta,
                        "threshold_percent": 5.0, "higher_is_regression": higher_is_regression, "scope": scope,
                        "regression": regression,
                    })
    return comparisons


def derive() -> dict[str, Any]:
    verified = verify.verify_matrix()
    rows = [row_summary(row) for row in verified["rows"]]
    return {
        "schema_version": 1, "change": 434,
        "classification": "Matched current-revision ODS streaming evidence; no speedup or causal optimization claim.",
        "timing_scope": "Operation timer covers deterministic scalar-row authoring, publication validation, compression/finalization, and writes to the hashing discard sink; artifact creation/reopen oracle, procfs probes, sink construction, and digest extraction are outside the operation timer.",
        "comparison_scope": "Matched before/after deltas are descriptive. Normal rows compare operation latency, archive throughput, and whole-process GNU time RSS; allocator rows compare requested count/bytes, regional peak, and whole-process GNU time RSS only.",
        "claims": [],
        "protocol_sha256": verified["protocol_sha256"], "oracle": verified["oracle"],
        "builds": verified["builds"], "ambient": verified["ambient"],
        "matrix": {**verified["matrix"], "formal_reports": len(rows), "retained_samples": len(rows) * 30},
        "cross_phase_identity": verified["cross_phase_identity"],
        "rows": rows,
        "repeat_flags": repeat_flags(rows),
        "comparisons": matched_comparisons(rows),
        "extra_attempts": verified["extra_attempts"],
        "uncertainty": {
            "method": "deterministic bootstrap median percentile", "resamples": 2000, "seed": 434, "level": 0.95,
            "scope": "normal operation samples", "sample_count": 30,
            "sample_scope": "30 within-process timed samples from one fresh process per report; not 30 independent process replicates",
        },
        "limitations": [
            "The operation timer excludes artifact creation/reopen oracle, procfs probes, sink construction, and digest extraction; GNU time RSS is whole-process and includes setup, corpus generation, and the harness oracle.",
            "Normal p50 intervals are descriptive bootstrap intervals over the 30 within-process samples from each report, not independent-process confidence intervals.",
            "GNU time RSS is one whole-process maximum-resident-set observation per fresh report process; it is reviewed separately from the within-process operation vectors.",
            "Allocator vectors are descriptive; allocator elapsed time is not a latency claim.",
            "Runtime ambient checkout identity is retained separately from executable build revision/source manifest.",
        ],
    }


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "summary.json")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    try:
        derived = derive()
        if args.check:
            if not args.output.is_file() or canonical(json.loads(args.output.read_text(encoding="utf-8"))) != canonical(derived):
                raise SummaryError(f"{args.output}: retained summary differs from fresh derivation")
        else:
            if args.output.exists():
                raise SummaryError(f"{args.output}: already exists; use --check")
            args.output.write_text(json.dumps(derived, indent=2, sort_keys=True, allow_nan=False) + "\n", encoding="utf-8")
    except (OSError, KeyError, TypeError, ValueError, AssertionError, verify.VerificationError, SummaryError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    print("VALID")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
