#!/usr/bin/env python3
"""Derive descriptive statistics from the verified 0457 candidate evidence.

The script invokes the portable candidate verifier first, then reads only the
authenticated reports, receipts, and GNU-time resource logs.  It retains exact
per-lane vectors and derives p50/p95/p99 values from each thirty-sample vector.
Normal lanes leave allocator fields unavailable.  No control comparison,
before/after claim, speedup claim, or missing-data substitution is performed.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import subprocess
import sys
from typing import Any

import verify as candidate_verify


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


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(label, "expected object")
    return value


def integer(value: Any, label: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(label, f"expected integer >= {minimum}")
    return value


def finite(value: int | float, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        fail(label, "expected finite number")
    return float(value)


def clean(value: float) -> int | float:
    return int(value) if value.is_integer() else value


def vector(value: Any, label: str, *, minimum: int = 0) -> list[int]:
    row = obj(value, label)
    if row.get("status") != "measured":
        fail(label, "metric vector is not measured")
    values = row.get("values")
    if not isinstance(values, list) or len(values) != SAMPLES:
        fail(label, f"expected exactly {SAMPLES} measured values")
    return [integer(item, f"{label}.values[{index}]", minimum=minimum) for index, item in enumerate(values)]


def plain_vector(value: Any, label: str, *, minimum: int = 0) -> list[int]:
    if not isinstance(value, list) or len(value) != SAMPLES:
        fail(label, f"expected exactly {SAMPLES} values")
    return [integer(item, f"{label}[{index}]", minimum=minimum) for index, item in enumerate(value)]


def permutation(value: Any, label: str) -> list[int]:
    if not isinstance(value, list) or len(value) != SAMPLES or sorted(value) != list(range(SAMPLES)):
        fail(label, "expected a permutation of sample indices")
    return [integer(item, f"{label}[{index}]") for index, item in enumerate(value)]


def chronological(values: list[int], sample_order: list[int]) -> list[int]:
    by_index = {sample_index: values[position] for position, sample_index in enumerate(sample_order)}
    return [by_index[index] for index in range(SAMPLES)]


def percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    rank = max(1, math.ceil(fraction * len(ordered)))
    return ordered[rank - 1]


def statistics(values: list[int], label: str) -> dict[str, Any]:
    if not values:
        fail(label, "empty vector")
    ordered = sorted(finite(value, f"{label}[{index}]") for index, value in enumerate(values))
    middle = (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) / 2.0
    return {
        "count": len(values),
        "min": clean(ordered[0]),
        "p50": clean(middle),
        "p95": clean(percentile(ordered, 0.95)),
        "p99": clean(percentile(ordered, 0.99)),
        "max": clean(ordered[-1]),
        "mean": clean(sum(ordered) / len(ordered)),
    }


def artifact(base: Path, record: Any, label: str) -> Path:
    row = obj(record, label)
    raw_path = row.get("path")
    if not isinstance(raw_path, str) or not raw_path or Path(raw_path).is_absolute() or ".." in Path(raw_path).parts:
        fail(label, "artifact path must be bundle-relative")
    path = (base / raw_path).resolve()
    if not path.is_file() or path.is_symlink() or not path.is_relative_to(base.resolve()):
        fail(label, "artifact is missing or escapes candidate bundle")
    expected_bytes = integer(row.get("bytes"), f"{label}.bytes")
    expected_sha = row.get("sha256")
    if path.stat().st_size != expected_bytes or not isinstance(expected_sha, str):
        fail(label, "artifact size/hash record is malformed")
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    if digest != expected_sha:
        fail(label, "artifact identity differs")
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
    if receipt.get("status") != "pass" or receipt.get("phase") != phase or receipt.get("attempt") != attempt or receipt.get("name") != name or receipt.get("lane") != expected_lane or receipt.get("protocol_sha256") != protocol_sha:
        fail(str(receipt_path), "successful receipt identity differs")
    artifacts = obj(receipt.get("artifacts"), f"{receipt_path}.artifacts")
    required = {"report", "catalog", "workload_log", "resource_log", "oracle_log"}
    if set(artifacts) != required:
        fail(str(receipt_path), "artifact set differs")
    report_path = artifact(root, artifacts["report"], f"{receipt_path}.report")
    resource_path = artifact(root, artifacts["resource_log"], f"{receipt_path}.resource_log")
    for key in required - {"report", "resource_log"}:
        artifact(root, artifacts[key], f"{receipt_path}.{key}")
    report = obj(load(report_path, str(report_path)), str(report_path))
    configuration = obj(report.get("configuration"), f"{report_path}.configuration")
    if configuration.get("samples_per_case") != SAMPLES or configuration.get("warmup_iterations_per_case") != WARMUPS or configuration.get("execution_workers") != [1]:
        fail(str(report_path), "sample, warmup, or worker dimensions differ")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1:
        fail(str(report_path), "expected one result")
    result = obj(results[0], f"{report_path}.results[0]")
    corpus = obj(result.get("corpus"), f"{report_path}.results[0].corpus")
    if result.get("case") != protocol["selector"] or corpus.get("shape") != shape:
        fail(str(report_path), "selector or shape differs")
    elapsed = obj(result.get("elapsed_ns"), f"{report_path}.elapsed_ns")
    elapsed_values = [integer(item, f"{report_path}.elapsed_ns.samples[{index}]", minimum=1) for index, item in enumerate(elapsed.get("samples", []))]
    if len(elapsed_values) != SAMPLES or elapsed_values != sorted(elapsed_values):
        fail(str(report_path), "elapsed vector is not sorted and complete")
    sample_order = permutation(elapsed.get("sample_order"), f"{report_path}.elapsed_ns.sample_order")
    operation = obj(result.get("operation_metrics"), f"{report_path}.operation_metrics")
    if operation.get("sample_count") != SAMPLES or permutation(operation.get("sample_indices"), f"{report_path}.operation_metrics.sample_indices") != sample_order:
        fail(str(report_path), "operation sample alignment differs")
    process = obj(operation.get("process"), f"{report_path}.operation_metrics.process")
    sink = obj(operation.get("sink"), f"{report_path}.operation_metrics.sink")
    source_metrics = obj(operation.get("source"), f"{report_path}.operation_metrics.source")
    metric_vectors: dict[str, list[int]] = {
        "elapsed_ns": elapsed_values,
        "peak_rss_bytes": vector(process.get("peak_rss_bytes"), f"{report_path}.process.peak_rss_bytes"),
        "rss_delta_bytes": vector(process.get("rss_delta_bytes"), f"{report_path}.process.rss_delta_bytes"),
        "sink_accepted_bytes": vector(sink.get("accepted_bytes"), f"{report_path}.sink.accepted_bytes"),
        "sink_write_calls": vector(sink.get("write_calls"), f"{report_path}.sink.write_calls"),
        "sink_largest_write": vector(sink.get("largest_write"), f"{report_path}.sink.largest_write"),
    }
    summary = obj(obj(result.get("source"), f"{report_path}.source").get("odp_source_tail_append"), f"{report_path}.source.odp_source_tail_append")
    for field in (
        "lifecycle_ns",
        "open_ns",
        "append_plan_ns",
        "publication_ns",
        "publication_report_bytes",
        "publication_report_source_version_id",
        "publication_report_source_version_revision",
        "runtime_proof_source_version_id",
        "runtime_proof_source_version_revision",
        "source_read_calls",
        "source_read_bytes",
    ):
        metric_vectors[field] = plain_vector(summary.get(field), f"{report_path}.source.odp_source_tail_append.{field}")
    bucket_metrics = obj(sink.get("write_size_buckets"), f"{report_path}.sink.write_size_buckets")
    for field in ("bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536"):
        metric_vectors[f"sink_bucket_{field}"] = vector(bucket_metrics.get(field), f"{report_path}.sink.write_size_buckets.{field}")
    allocation_vectors: dict[str, list[int] | None]
    if mode == "allocator":
        allocation = obj(operation.get("allocation"), f"{report_path}.operation_metrics.allocation")
        if allocation.get("status") != "measured":
            fail(str(report_path), "allocator metrics are not measured")
        allocation_vectors = {field: vector(allocation.get(field), f"{report_path}.allocation.{field}") for field in ALLOCATION_FIELDS}
        if any(value != 0 for value in allocation_vectors["failed_allocation_calls"] or []):
            fail(str(report_path), "allocator failure calls are nonzero")
        for index in range(SAMPLES):
            if allocation_vectors["live_bytes_before"][index] + allocation_vectors["allocated_bytes"][index] - allocation_vectors["deallocated_bytes"][index] != allocation_vectors["live_bytes_after"][index]:
                fail(str(report_path), "allocator live-byte balance differs")
            if allocation_vectors["region_peak_live_bytes"][index] < max(allocation_vectors["live_bytes_before"][index], allocation_vectors["live_bytes_after"][index]):
                fail(str(report_path), "allocator region peak is below an endpoint")
        allocation_vectors["heap_growth_bytes"] = [after - before for before, after in zip(allocation_vectors["live_bytes_before"], allocation_vectors["live_bytes_after"])]
        allocation_vectors["region_peak_above_entry"] = [peak - before for peak, before in zip(allocation_vectors["region_peak_live_bytes"], allocation_vectors["live_bytes_before"])]
    else:
        if operation.get("allocation") is not None:
            fail(str(report_path), "normal lane must not publish allocator vectors")
        allocation_vectors = {field: None for field in (*ALLOCATION_FIELDS, "heap_growth_bytes", "region_peak_above_entry")}
    all_vectors: dict[str, list[int] | None] = {**metric_vectors, **allocation_vectors}
    vector_output: dict[str, Any] = {}
    statistics_output: dict[str, Any] = {}
    for field, values in all_vectors.items():
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
        "source_archive_sha256": summary.get("source_archive_sha256"),
        "source_archive_bytes": summary.get("source_archive_bytes"),
        "fixture_proof_source_version_id": summary.get("proof_source_version_id"),
        "fixture_proof_source_version_revision": summary.get("proof_source_version_revision"),
        "output_archive_sha256": summary.get("output_archive_sha256"),
        "output_archive_bytes": summary.get("output_archive_bytes"),
        "vectors": vector_output,
        "statistics": statistics_output,
        "gnu_time_max_rss_bytes": resource_rss(resource_path, f"{receipt_path}.resource_log"),
    }


def derive(root: Path, protocol_path: Path, attempt: str, precleanup: bool, repo_root: Path | None) -> dict[str, Any]:
    protocol_path = protocol_path.resolve()
    verification = candidate_verify.verify(protocol_path, attempt, precleanup, repo_root)
    protocol = obj(load(protocol_path, str(protocol_path)), str(protocol_path))
    protocol_sha = verification["protocol_sha256"]
    rows = [lane(root, protocol, protocol_sha, phase, mode, shape, attempt) for phase in ("R1", "R2") for mode, shape in PHASE_ORDER[phase]]
    if len(rows) != 12:
        fail("matrix", "expected twelve verified rows")
    summary_rows: dict[str, dict[str, dict[str, Any]]] = {mode: {} for mode in MODES}
    metric_names = tuple(rows[0]["statistics"])
    for mode in MODES:
        for shape in SHAPES:
            selected = [row for row in rows if row["mode"] == mode and row["shape"] == shape]
            if len(selected) != 2:
                fail(f"{mode}.{shape}", "expected one R1 and one R2 row")
            summary_rows[mode][shape] = {
                "repeats": [row["repeat"] for row in selected],
                "output_archive_sha256": [row["output_archive_sha256"] for row in selected],
                "output_archive_bytes": [row["output_archive_bytes"] for row in selected],
            }
            for metric in metric_names:
                summary_rows[mode][shape][metric] = [row["statistics"][metric] for row in selected]
    return {
        "schema": "litchi-0457-odp-source-tail-candidate-summary-v1",
        "change": CHANGE,
        "protocol_sha256": protocol_sha,
        "output_binding_sha256": verification["output_binding_sha256"],
        "attempt": attempt,
        "scope": "descriptive current-revision source-backed ODP publication-plan candidate; specialized retained-result contract only",
        "statistic_definition": {
            "p50": "average of the two middle values after sorting the retained thirty-value vector",
            "p95": "nearest-rank value at rank ceil(0.95 * n)",
            "p99": "nearest-rank value at rank ceil(0.99 * n)",
            "source": "derive.py recomputes statistics from retained vectors; report-embedded summaries are not silently mixed",
        },
        "matrix": {"reports": 12, "retained_samples": 360, "samples_per_report": SAMPLES, "warmups_per_report": WARMUPS, "cpu": 2, "workers": 1, "selector": protocol["selector"], "phases": ["R1", "R2"]},
        "by_mode_shape": summary_rows,
        "rows": rows,
        "failed_attempts": verification["failed_attempts"],
        "comparison_contract": {
            "status": "withheld",
            "control_comparison": "separate contract required",
            "reason": "source-backed publication has a different retained-result contract from owned Snapshot/Commit/Patch; no ordinary speedup, regression, or CRUD comparison is derived here",
        },
        "claims": [
            "p50, p95, and p99 are descriptive statistics over each retained thirty-sample vector",
            "peak_rss_bytes is the in-process operation metric; gnu_time_max_rss_bytes is a whole-process invocation observation",
            "allocator_calls, byte counters, live endpoints, peaks, heap_growth_bytes, and region_peak_above_entry retain exact measured allocator vectors only for allocator lanes",
            "normal lanes leave allocator vectors unavailable and never synthesize zero values",
            "source read calls/returned bytes and sink accepted bytes/write counters are retained; requested-byte histograms and physical I/O attribution remain outside this contract",
            "no ordinary Commit/Patch speedup, regression, general CRUD equivalence, before/after, causal, bounded-memory, scaling, or cancellation claim is authorized",
        ],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol.json")
    parser.add_argument("--attempt", default="formal")
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--precleanup", action="store_true")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args(argv)
    try:
        value = derive(args.root.resolve(), args.protocol.resolve(), args.attempt, args.precleanup, args.repo_root)
        raw = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False).encode("utf-8") + b"\n"
        if args.write:
            output = (args.output or args.root.resolve() / "summary.json").resolve()
            if output.exists() or output.is_symlink():
                fail("output", "refusing to overwrite an existing derived summary")
            output.write_bytes(raw)
        print(raw.decode("utf-8"))
        return 0
    except (OSError, KeyError, TypeError, ValueError, DeriveError, candidate_verify.VerifyError, subprocess.SubprocessError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
