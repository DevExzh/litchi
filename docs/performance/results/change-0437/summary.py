#!/usr/bin/env python3
"""Derive the descriptive 0437 ODP matrix summary from ``verify.py``.

The summary keeps the buffered control comparison and the buffered-to-streaming
API comparison separate.  It reports timed normal vectors, allocator vectors,
and whole-process RSS evidence with review flags.  It never labels a role
change a speedup or assigns a causal hotspot from this matrix alone.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
from pathlib import Path
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
ROLES = ("before-buffered", "after-buffered", "after-streaming")
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
THRESHOLD = 5.0


class SummaryError(ValueError):
    pass


def load_module(filename: str, name: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / filename)
    if spec is None or spec.loader is None:
        raise SummaryError(f"cannot load {filename}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


verify = load_module("verify.py", "change0437_verify_for_summary")


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise SummaryError(f"cannot canonicalize summary: {error}") from error


def number(value: Any, label: str, *, signed: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)) or (not signed and float(value) < 0):
        raise SummaryError(f"{label}: expected a finite number")
    return float(value)


def integral(value: float) -> int | float:
    return int(value) if value.is_integer() else value


def stats(values: list[int | float], label: str, *, signed: bool = False) -> dict[str, Any]:
    if not values:
        raise SummaryError(f"{label}: empty vector")
    ordered = sorted(number(value, f"{label}[{index}]", signed=signed) for index, value in enumerate(values))
    count = len(ordered)
    p50 = (ordered[(count - 1) // 2] + ordered[count // 2]) / 2.0
    p95 = ordered[min(((95 * count + 99) // 100) - 1, count - 1)]
    p99 = ordered[min(((99 * count + 99) // 100) - 1, count - 1)]
    return {
        "count": count,
        "min": integral(ordered[0]),
        "p50": integral(p50),
        "p95": integral(p95),
        "p99": integral(p99),
        "max": integral(ordered[-1]),
        "mean": integral(sum(ordered) / count),
    }


def metric_values(metric: Any, label: str) -> list[int] | None:
    if not isinstance(metric, dict):
        raise SummaryError(f"{label}: malformed metric")
    if metric.get("status") != "measured":
        return None
    values = metric.get("values")
    if not isinstance(values, list):
        raise SummaryError(f"{label}: measured metric has no values")
    return [verify.u64(value, f"{label}.values[{index}]") for index, value in enumerate(values)]


def relative(old: float, new: float, label: str) -> float:
    if old == 0:
        raise SummaryError(f"{label}: zero baseline")
    return (new - old) / old * 100.0


def elapsed_summary(row: dict[str, Any]) -> dict[str, Any]:
    values = [verify.u64(value, f"{row['name']}.elapsed[{index}]") for index, value in enumerate(row["verified"]["elapsed"])]
    result = stats(values, f"{row['name']}.elapsed_ns")
    report_elapsed = row["verified"]["result"]["elapsed_ns"]
    result["standard_deviation"] = number(report_elapsed["standard_deviation"], f"{row['name']}.elapsed_ns.standard_deviation")
    interval = report_elapsed.get("confidence_interval_95")
    if not isinstance(interval, dict):
        raise SummaryError(f"{row['name']}.elapsed_ns.confidence_interval_95: missing")
    result["confidence_interval_95"] = {
        "method": interval.get("method"),
        "lower": number(interval.get("lower"), f"{row['name']}.elapsed_ns.confidence_interval_95.lower"),
        "upper": number(interval.get("upper"), f"{row['name']}.elapsed_ns.confidence_interval_95.upper"),
    }
    return result


def allocator_summary(row: dict[str, Any]) -> dict[str, Any]:
    allocation = row["verified"]["metrics"].get("allocation")
    if row["lane"]["mode"] == "normal" or not isinstance(allocation, dict) or allocation.get("status") != "measured":
        return {"status": "unavailable", "scope": "operation_global_system_allocator", "latency_claim": "allocator elapsed values are not a latency claim"}
    names = ("allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes")
    vectors: dict[str, list[int]] = {}
    for name in names:
        values = metric_values(allocation.get(name), f"{row['name']}.allocation.{name}")
        if values is None:
            raise SummaryError(f"{row['name']}: allocator vector {name} is unavailable")
        vectors[name] = values
    balance = [before + allocated - deallocated - after for before, allocated, deallocated, after in zip(vectors["live_bytes_before"], vectors["allocated_bytes"], vectors["deallocated_bytes"], vectors["live_bytes_after"], strict=True)]
    if any(value != 0 for value in balance):
        raise SummaryError(f"{row['name']}: allocator live-byte balance is nonzero")
    live_delta = [after - before for before, after in zip(vectors["live_bytes_before"], vectors["live_bytes_after"], strict=True)]
    if any(value != 0 for value in live_delta):
        raise SummaryError(f"{row['name']}: fresh complete operation has nonzero live-byte delta")
    region_above_entry = [peak - before for peak, before in zip(vectors["region_peak_live_bytes"], vectors["live_bytes_before"], strict=True)]
    region_above_exit = [peak - after for peak, after in zip(vectors["region_peak_live_bytes"], vectors["live_bytes_after"], strict=True)]
    return {
        "status": "measured",
        "scope": allocation.get("scope"),
        "statistics": {name: stats(values, f"{row['name']}.allocation.{name}") for name, values in vectors.items()},
        "statistics_extra": {
            "region_peak_above_entry": stats(region_above_entry, f"{row['name']}.allocation.region_peak_above_entry", signed=True),
            "region_peak_above_exit": stats(region_above_exit, f"{row['name']}.allocation.region_peak_above_exit", signed=True),
            "live_balance": stats(balance, f"{row['name']}.allocation.live_balance", signed=True),
            "live_delta": stats(live_delta, f"{row['name']}.allocation.live_delta", signed=True),
        },
        "latency_claim": "allocator elapsed values retained for process context only; no allocator latency claim",
    }


def row_summary(row: dict[str, Any]) -> dict[str, Any]:
    verified = row["verified"]
    identity = row["identity"]
    elapsed_values = [verify.u64(value, f"{row['name']}.elapsed[{index}]") for index, value in enumerate(verified["elapsed"])]
    throughput = None
    if row["lane"]["mode"] == "normal":
        slides = identity["slide_count"]
        text_bytes = identity["canonical_text_bytes"]
        throughput = {
            "slides_per_second": stats([slides * 1_000_000_000 / value for value in elapsed_values], f"{row['name']}.throughput.slides_per_second"),
            "canonical_text_bytes_per_second": stats([text_bytes * 1_000_000_000 / value for value in elapsed_values], f"{row['name']}.throughput.canonical_text_bytes_per_second"),
            "scope": "timed operation samples; canonical presentation text bytes are the independent ODP semantic projection",
        }
    return {
        "phase": row["phase"], "role": row["role"], "mode": row["lane"]["mode"],
        "shape": row["lane"]["shape"], "repeat": row["lane"]["repeat"], "name": row["name"],
        "report": row["report_path"], "catalog": row["catalog_path"],
        "corpus": {key: identity.get(key) for key in ("archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256")},
        "output_sha256": identity.get("output_sha256"),
        "semantic_sha256": identity.get("semantic_sha256"),
        "source_identity": {
            "binary_sha256": row["receipt"]["binary"]["sha256"],
            "source_manifest_sha256": row["receipt"]["source_manifest"]["sha256"],
            "ambient_git_revision": row["environment"]["git_revision"],
            "ambient_worktree_dirty": row["environment"]["git_worktree_dirty"],
            "styles_xml_sha256": identity.get("styles_xml_sha256"),
            "meta_xml_sha256": identity.get("meta_xml_sha256"),
        },
        "elapsed_ns": elapsed_summary(row),
        "process_memory": {
            "gnu_time_peak_rss_bytes": row["resource"].get("peak_rss_bytes"),
            "scope": "whole executable invocation including setup/corpus/warmups/timed calls/output hashing; copied Python oracle validation runs afterward outside GNU time/RSS",
        },
        "allocation": allocator_summary(row),
        "throughput": throughput,
        "timing_scope": "30 timed samples inside one fresh process per report; corpus/setup and copied Python oracle work are outside the operation timer",
    }


def metric_value(row: dict[str, Any], metric: str) -> float | None:
    elapsed = row["elapsed_ns"]
    if metric in elapsed:
        return float(elapsed[metric])
    if metric.startswith("throughput."):
        section, key = metric.split(".", 1)
        value = row.get(section)
        if not isinstance(value, dict) or not isinstance(value.get(key), dict):
            return None
        return float(value[key]["p50"])
    rss = row["process_memory"].get("gnu_time_peak_rss_bytes")
    if metric == "gnu_time_peak_rss_bytes":
        return float(rss) if rss is not None else None
    allocation = row["allocation"]
    if allocation.get("status") != "measured":
        return None
    if metric.startswith("allocation."):
        key = metric.removeprefix("allocation.")
        if key in allocation["statistics"]:
            return float(allocation["statistics"][key]["p50"])
        if key in allocation["statistics_extra"]:
            return float(allocation["statistics_extra"][key]["p50"])
    return None


def comparisons(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    indexed = {(row["role"], row["mode"], row["shape"], row["repeat"]): row for row in rows}
    normal_metrics = ("p50", "p95", "p99", "mean", "throughput.slides_per_second", "throughput.canonical_text_bytes_per_second", "gnu_time_peak_rss_bytes")
    allocator_metrics = ("allocation.allocation_calls", "allocation.allocated_bytes", "allocation.region_peak_above_entry", "allocation.live_bytes_after", "gnu_time_peak_rss_bytes")
    pairs: list[dict[str, Any]] = []
    flags: list[dict[str, Any]] = []
    for before, after, comparison_scope in (("before-buffered", "after-buffered", "buffered_control"), ("after-buffered", "after-streaming", "buffered_to_streaming_api")):
        for mode in MODES:
            for shape in SHAPES:
                metrics = normal_metrics if mode == "normal" else allocator_metrics
                for repeat in ("R1", "R2"):
                    left = indexed[(before, mode, shape, repeat)]
                    right = indexed[(after, mode, shape, repeat)]
                    for metric in metrics:
                        old = metric_value(left, metric)
                        new = metric_value(right, metric)
                        if old is None or new is None:
                            continue
                        change = relative(old, new, f"{comparison_scope}.{mode}.{shape}.{repeat}.{metric}")
                        higher_is_regression = not metric.startswith("throughput.")
                        flagged = change > THRESHOLD if higher_is_regression else change < -THRESHOLD
                        row = {"scope": comparison_scope, "mode": mode, "shape": shape, "repeat": repeat, "metric": metric, "before_role": before, "after_role": after, "before": old, "after": new, "relative_percent": change, "threshold_percent": THRESHOLD, "higher_is_regression": higher_is_regression, "flagged": flagged, "interpretation": "descriptive review flag; no speedup or causal claim"}
                        pairs.append(row)
                        if row["flagged"]:
                            flags.append(row)
    for role in ROLES:
        for mode in MODES:
            for shape in SHAPES:
                for metric in ("p50", "p95", "p99", "throughput.slides_per_second", "throughput.canonical_text_bytes_per_second", "gnu_time_peak_rss_bytes") if mode == "normal" else allocator_metrics:
                    first = indexed[(role, mode, shape, "R1")]
                    second = indexed[(role, mode, shape, "R2")]
                    old = metric_value(first, metric)
                    new = metric_value(second, metric)
                    if old is None or new is None:
                        continue
                    change = relative(old, new, f"repeat.{role}.{mode}.{shape}.{metric}")
                    row = {"scope": "within_role_repeat_drift", "role": role, "mode": mode, "shape": shape, "metric": metric, "r1": old, "r2": new, "relative_percent": change, "threshold_percent": THRESHOLD, "higher_is_regression": None, "flagged": abs(change) > THRESHOLD, "interpretation": "repeat stability review uses absolute drift; role-pair cost direction does not apply"}
                    if row["flagged"]:
                        flags.append(row)
    return pairs, flags


def derive() -> dict[str, Any]:
    envelope = verify.verify_matrix()
    rows = [row_summary(row) for row in envelope["rows"]]
    pairs, flags = comparisons(rows)
    return {
        "schema": "litchi-0437-odp-summary-v1",
        "change": 437,
        "protocol_sha256": envelope["protocol_sha256"],
        "matrix": envelope["matrix"],
        "ambient": envelope["ambient"],
        "rows": rows,
        "comparisons": pairs,
        "review_flags": flags,
        "claims": [
            "descriptive current-revision ODP evidence only",
            "normal elapsed vectors describe the timed operation scope",
            "allocator vectors and RSS retain whole-process/operation scope separately",
            "no speedup, causality, fixed-memory, cache-miss, or 10x claim is authorized by this summary",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    try:
        value = derive()
        if args.check:
            path = ROOT / "summary.json"
            if not path.is_file() or canonical(json.loads(path.read_text(encoding="utf-8"))) != canonical(value):
                raise SummaryError("retained summary differs from independent derivation")
        if args.write:
            (ROOT / "summary.json").write_bytes(canonical(value) + b"\n")
        print(json.dumps(value, indent=2, sort_keys=True, allow_nan=False))
    except (OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
