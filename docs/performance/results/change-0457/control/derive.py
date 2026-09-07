#!/usr/bin/env python3
"""Derive a descriptive 0457 ODP control summary from retained evidence.

This script reads successful capture receipts, reports, and resource logs.  It
does not invoke a workload, profiler, oracle, or source checkout.  The output
keeps the exact retained vectors and derives within-lane p50/p95/p99 values;
it makes no before/after, causal, bounded-memory, or speedup claim.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 457
SAMPLES = 30
WARMUPS = 3
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
PHASE_ORDER = {
    "R1": (("normal", "tiny"), ("normal", "medium"), ("normal", "large"),
           ("allocator", "tiny"), ("allocator", "medium"), ("allocator", "large")),
    "R2": (("allocator", "large"), ("allocator", "medium"), ("allocator", "tiny"),
           ("normal", "large"), ("normal", "medium"), ("normal", "tiny")),
}
RSS_RE = re.compile(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$")


class DeriveError(ValueError):
    pass


def fail(label: str, message: str) -> None:
    raise DeriveError(f"{label}: {message}")


def load(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(label, "expected object")
    return value


def integer(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        fail(label, "expected non-negative integer")
    return value


def vector(value: Any, label: str) -> list[int]:
    row = obj(value, label)
    values = row.get("values")
    if not isinstance(values, list) or len(values) != SAMPLES:
        fail(label, f"expected exactly {SAMPLES} values")
    return [integer(item, f"{label}.values[{index}]") for index, item in enumerate(values)]


def permutation(value: Any, label: str) -> list[int]:
    if not isinstance(value, list) or len(value) != SAMPLES or sorted(value) != list(range(SAMPLES)):
        fail(label, "expected a permutation of sample indices")
    return [integer(item, f"{label}[{index}]") for index, item in enumerate(value)]


def finite(value: int | float, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        fail(label, "expected a finite number")
    return float(value)


def clean(value: float) -> int | float:
    return int(value) if value.is_integer() else value


def percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def statistics(values: list[int], label: str) -> dict[str, Any]:
    if not values:
        fail(label, "empty vector")
    clean_values = [finite(value, f"{label}[{index}]") for index, value in enumerate(values)]
    ordered = sorted(clean_values)
    median = (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) / 2.0
    return {
        "count": len(values),
        "min": clean(ordered[0]),
        "p50": clean(median),
        "p95": clean(percentile(clean_values, 0.95)),
        "p99": clean(percentile(clean_values, 0.99)),
        "max": clean(ordered[-1]),
        "mean": clean(sum(clean_values) / len(clean_values)),
    }


def chronological(values: list[int], sample_order: list[int]) -> list[int]:
    by_index = {sample_index: values[position] for position, sample_index in enumerate(sample_order)}
    return [by_index[index] for index in range(SAMPLES)]


def artifact(base: Path, record: Any, label: str) -> Path:
    row = obj(record, label)
    raw_path = row.get("path")
    if not isinstance(raw_path, str) or not raw_path or Path(raw_path).is_absolute() or ".." in Path(raw_path).parts:
        fail(label, "artifact path must be bundle-relative")
    path = (base / raw_path).resolve()
    if not path.is_file() or path.is_symlink() or not path.is_relative_to(base.resolve()):
        fail(label, "artifact is missing or escapes the control bundle")
    if sha(path) != row.get("sha256") or path.stat().st_size != row.get("bytes"):
        fail(label, "artifact identity differs from its receipt")
    return path


def resource_rss(path: Path, label: str) -> int:
    matches = [int(match.group(1)) * 1024 for line in path.read_text(encoding="utf-8", errors="replace").splitlines() if (match := RSS_RE.match(line))]
    if len(matches) != 1:
        fail(label, "expected one GNU-time maximum RSS line")
    return matches[0]


def lane(root: Path, protocol: dict[str, Any], protocol_sha: str, phase: str, mode: str, shape: str, attempt: str) -> dict[str, Any]:
    name = f"{phase}-{mode}-{shape}-{phase.lower()}"
    directory = root / "runs" / phase / attempt
    receipt_path = directory / f"{name}-receipt.json"
    receipt = obj(load(receipt_path, str(receipt_path)), str(receipt_path))
    expected_lane = {"mode": mode, "phase": phase, "repeat": phase, "shape": shape}
    if receipt.get("schema") != "litchi-0457-control-capture-receipt-v1" or receipt.get("change") != CHANGE:
        fail(str(receipt_path), "receipt schema/change differs")
    if receipt.get("status") != "pass" or receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
        fail(str(receipt_path), "capture and oracle must pass")
    if receipt.get("phase") != phase or receipt.get("attempt") != attempt or receipt.get("name") != name or receipt.get("lane") != expected_lane:
        fail(str(receipt_path), "receipt lane identity differs")
    if receipt.get("protocol_sha256") != protocol_sha:
        fail(str(receipt_path), "receipt protocol binding differs")
    if receipt.get("source_unchanged") is not True:
        fail(str(receipt_path), "source custody is not unchanged")
    artifacts = obj(receipt.get("artifacts"), f"{receipt_path}.artifacts")
    required = {"report", "catalog", "workload_log", "resource_log", "oracle_log"}
    if set(artifacts) != required:
        fail(str(receipt_path), "artifact set differs")
    report_path = artifact(root, artifacts["report"], f"{receipt_path}.report")
    resource_path = artifact(root, artifacts["resource_log"], f"{receipt_path}.resource_log")
    artifact(root, artifacts["catalog"], f"{receipt_path}.catalog")
    artifact(root, artifacts["workload_log"], f"{receipt_path}.workload_log")
    artifact(root, artifacts["oracle_log"], f"{receipt_path}.oracle_log")
    report = obj(load(report_path, str(report_path)), str(report_path))
    configuration = obj(report.get("configuration"), f"{report_path}.configuration")
    if configuration.get("samples_per_case") != SAMPLES or configuration.get("warmup_iterations_per_case") != WARMUPS or configuration.get("execution_workers") != [1]:
        fail(str(report_path), "sample or worker configuration differs")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1:
        fail(str(report_path), "expected one result")
    result = obj(results[0], f"{report_path}.results[0]")
    corpus = obj(result.get("corpus"), f"{report_path}.results[0].corpus")
    if result.get("case") != protocol["selector"] or corpus.get("shape") != shape:
        fail(str(report_path), "selector or shape differs")
    elapsed = obj(result.get("elapsed_ns"), f"{report_path}.elapsed_ns")
    elapsed_values = [integer(value, f"{report_path}.elapsed_ns.samples[{index}]") for index, value in enumerate(elapsed.get("samples", []))]
    if len(elapsed_values) != SAMPLES:
        fail(str(report_path), "elapsed vector length differs")
    sample_order = permutation(elapsed.get("sample_order"), f"{report_path}.elapsed_ns.sample_order")
    operation = obj(result.get("operation_metrics"), f"{report_path}.operation_metrics")
    if operation.get("sample_count") != SAMPLES or permutation(operation.get("sample_indices"), f"{report_path}.operation_metrics.sample_indices") != sample_order:
        fail(str(report_path), "operation sample alignment differs")
    process = obj(operation.get("process"), f"{report_path}.operation_metrics.process")
    peak_rss = vector(process.get("peak_rss_bytes"), f"{report_path}.process.peak_rss_bytes")
    rss_delta = vector(process.get("rss_delta_bytes"), f"{report_path}.process.rss_delta_bytes")
    sink = obj(operation.get("sink"), f"{report_path}.operation_metrics.sink")
    accepted_bytes = vector(sink.get("accepted_bytes"), f"{report_path}.sink.accepted_bytes")
    write_calls = vector(sink.get("write_calls"), f"{report_path}.sink.write_calls")
    largest_write = vector(sink.get("largest_write"), f"{report_path}.sink.largest_write")
    allocation: dict[str, Any]
    if mode == "allocator":
        raw_allocation = obj(operation.get("allocation"), f"{report_path}.operation_metrics.allocation")
        if raw_allocation.get("status") != "measured":
            fail(str(report_path), "allocator vectors are not measured")
        vectors = {field: vector(raw_allocation.get(field), f"{report_path}.allocation.{field}") for field in ALLOCATION_FIELDS}
        heap_growth = [after - before for before, after in zip(vectors["live_bytes_before"], vectors["live_bytes_after"])]
        region_peak_above_entry = [peak - before for peak, before in zip(vectors["region_peak_live_bytes"], vectors["live_bytes_before"])]
        allocation = {
            "status": "measured",
            "vectors": vectors,
            "heap_growth_bytes": heap_growth,
            "region_peak_above_entry": region_peak_above_entry,
        }
    else:
        if "allocation" in operation:
            fail(str(report_path), "normal report must not synthesize allocator vectors")
        allocation = {
            "status": "unavailable",
            "vectors": {field: None for field in ALLOCATION_FIELDS},
            "heap_growth_bytes": None,
            "region_peak_above_entry": None,
        }
    derived_allocation = {
        **allocation["vectors"],
        "heap_growth_bytes": allocation["heap_growth_bytes"],
        "region_peak_above_entry": allocation["region_peak_above_entry"],
    }
    vector_output = {
        "elapsed_ns": {"sample_order": sample_order, "elapsed_order": elapsed_values, "sample_index_order": chronological(elapsed_values, sample_order)},
        "peak_rss_bytes": {"sample_order": sample_order, "elapsed_order": peak_rss, "sample_index_order": chronological(peak_rss, sample_order)},
        "rss_delta_bytes": {"sample_order": sample_order, "elapsed_order": rss_delta, "sample_index_order": chronological(rss_delta, sample_order)},
        "sink_accepted_bytes": {"sample_order": sample_order, "elapsed_order": accepted_bytes, "sample_index_order": chronological(accepted_bytes, sample_order)},
        "sink_write_calls": {"sample_order": sample_order, "elapsed_order": write_calls, "sample_index_order": chronological(write_calls, sample_order)},
        "sink_largest_write": {"sample_order": sample_order, "elapsed_order": largest_write, "sample_index_order": chronological(largest_write, sample_order)},
    }
    statistics_output = {
        "elapsed_ns": statistics(elapsed_values, f"{report_path}.elapsed_ns"),
        "peak_rss_bytes": statistics(peak_rss, f"{report_path}.peak_rss_bytes"),
        "rss_delta_bytes": statistics(rss_delta, f"{report_path}.rss_delta_bytes"),
        "sink_accepted_bytes": statistics(accepted_bytes, f"{report_path}.sink_accepted_bytes"),
        "sink_write_calls": statistics(write_calls, f"{report_path}.sink_write_calls"),
        "sink_largest_write": statistics(largest_write, f"{report_path}.sink_largest_write"),
    }
    for field, values in derived_allocation.items():
        if values is None:
            vector_output[field] = None
            statistics_output[field] = None
        else:
            vector_output[field] = {"sample_order": sample_order, "elapsed_order": values, "sample_index_order": chronological(values, sample_order)}
            statistics_output[field] = statistics(values, f"{report_path}.{field}")
    return {
        "phase": phase,
        "mode": mode,
        "shape": shape,
        "repeat": phase,
        "name": name,
        "receipt": str(receipt_path.relative_to(root)),
        "report": str(report_path.relative_to(root)),
        "vectors": vector_output,
        "statistics": statistics_output,
        "gnu_time_max_rss_bytes": resource_rss(resource_path, f"{receipt_path}.resource_log"),
    }


def derive(root: Path, protocol_path: Path, attempt: str) -> dict[str, Any]:
    protocol = obj(load(protocol_path, str(protocol_path)), str(protocol_path))
    if protocol.get("change") != CHANGE or protocol.get("selector") != "odp_existing_append_lifecycle" or protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS:
        fail("protocol", "control dimensions differ")
    protocol_sha = sha(protocol_path)
    rows = [lane(root, protocol, protocol_sha, phase, mode, shape, attempt) for phase in ("R1", "R2") for mode, shape in PHASE_ORDER[phase]]
    if len(rows) != 12:
        fail("matrix", "expected twelve rows")
    summary_rows: dict[str, dict[str, dict[str, Any]]] = {mode: {} for mode in MODES}
    for mode in MODES:
        for shape in SHAPES:
            selected = [row for row in rows if row["mode"] == mode and row["shape"] == shape]
            if len(selected) != 2:
                fail(f"{mode}.{shape}", "expected one R1 and one R2 row")
            metric_names = (
                "elapsed_ns",
                "peak_rss_bytes",
                "rss_delta_bytes",
                "sink_accepted_bytes",
                "sink_write_calls",
                "sink_largest_write",
                *ALLOCATION_FIELDS,
                "heap_growth_bytes",
                "region_peak_above_entry",
            )
            summary_rows[mode][shape] = {"repeats": [row["repeat"] for row in selected]}
            summary_rows[mode][shape].update({metric: [row["statistics"][metric] for row in selected] for metric in metric_names})
    return {
        "schema": "litchi-0457-odp-control-summary-v1",
        "change": CHANGE,
        "protocol_sha256": protocol_sha,
        "attempt": attempt,
        "scope": "descriptive current-revision ODP existing-append control; no before/after, causal, bounded-memory, or speedup claim",
        "statistic_definition": {
            "p50": "average of the two middle values after sorting the 30-value vector",
            "p95": "nearest-rank value at rank ceil(0.95 * n)",
            "p99": "nearest-rank value at rank ceil(0.99 * n)",
            "source": "derive.py recomputes these values from retained vectors; report-embedded summary fields are not silently mixed into these statistics",
        },
        "matrix": {"reports": 12, "retained_samples": 360, "samples_per_report": SAMPLES, "warmups_per_report": WARMUPS, "cpu": 2, "workers": 1, "selector": protocol["selector"], "phases": ["R1", "R2"]},
        "by_mode_shape": summary_rows,
        "rows": rows,
        "claims": [
            "p50, p95, and p99 are descriptive statistics over each retained thirty-sample vector",
            "allocation_calls, deallocation_calls, reallocation_calls, failed_allocation_calls, allocated_bytes, deallocated_bytes, endpoint/peak live bytes, and region_peak_above_entry retain exact allocator vectors and statistics",
            "peak_rss_bytes is the in-process operation metric; gnu_time_max_rss_bytes is a whole-process invocation observation",
            "heap_growth_bytes is available only for measured allocator lanes as live_bytes_after minus live_bytes_before; region_peak_above_entry is region_peak_live_bytes minus live_bytes_before",
            "normal lanes do not synthesize allocator or heap-growth zeros",
            "no before/after, causal, bounded-memory, physical-I/O, scaling, or speedup claim is authorized",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol-r1.json")
    parser.add_argument("--attempt", default="formal-r1")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    root = args.root.resolve()
    protocol_path = args.protocol.resolve()
    try:
        value = derive(root, protocol_path, args.attempt)
        raw = json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode("utf-8") + b"\n"
        if args.write:
            output = (args.output or root / "summary.json").resolve()
            output.write_bytes(raw)
        print(raw.decode("utf-8"))
    except (OSError, KeyError, TypeError, ValueError, DeriveError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
