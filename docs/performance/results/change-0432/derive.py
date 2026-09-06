#!/usr/bin/env python3
"""Derive a descriptive, lossless summary from the 0432 capture matrix.

The summary is a view of twelve already verified reports.  It keeps the raw
elapsed, allocator, operation-process, and external ``/usr/bin/time`` samples
alongside compact p50/p95/p99/mean values.  Repeat differences are descriptive
flags at a five-percent threshold; this batch makes no optimization claim.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("streaming_report", ROOT / "verify-report.py")
verify_report = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verify_report)


CHANGE = 432
CASE = "xlsx_streaming_create"
SAMPLES = 30
REPEATS = ("R1", "R2")
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
RSS_RE = re.compile(r"^Maximum resident set size \(kbytes\):\s*(\d+)\s*$")


class DerivationError(ValueError):
    """Capture rows cannot be summarized without changing their meaning."""


def fail(path: str, message: str) -> None:
    raise DerivationError(f"{path}: {message}")


def obj(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "expected an object")
    return value


def array(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "expected an array")
    return value


def u64(value: Any, path: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0 or value > verify_report.U64_MAX:
        fail(path, "expected a u64")
    return value


def finite(value: Any, path: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        fail(path, "expected a finite number")
    try:
        result = float(value)
    except (OverflowError, ValueError):
        fail(path, "number is outside the finite range")
    if not math.isfinite(result) or result < 0:
        fail(path, "expected a finite non-negative number")
    return result


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode("utf-8")


def load(path: Path, label: str) -> Any:
    return verify_report.load_json(path, label)


def sha256_file(path: Path, label: str) -> str:
    if not path.is_file():
        fail(label, "file is missing")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as error:
        fail(label, f"cannot read file: {error}")
    return digest.hexdigest()


def stats(values: list[int | float], path: str) -> dict[str, Any]:
    if not values:
        fail(path, "cannot summarize an empty vector")
    ordered = sorted(finite(value, f"{path}.samples[{index}]") for index, value in enumerate(values))
    count = len(ordered)
    p50 = (ordered[(count - 1) // 2] + ordered[count // 2]) / 2.0
    p95 = ordered[min(((95 * count + 99) // 100) - 1, count - 1)]
    p99 = ordered[min(((99 * count + 99) // 100) - 1, count - 1)]
    mean = sum(ordered) / count
    return {
        "min": ordered[0],
        "p50": p50,
        "p95": p95,
        "p99": p99,
        "max": ordered[-1],
        "mean": mean,
        "samples": list(values),
    }


def signed_stats(values: list[int], path: str) -> dict[str, Any]:
    """Summarize an endpoint delta without converting a negative delta to zero."""

    if not values:
        fail(path, "cannot summarize an empty vector")
    ordered = []
    for index, value in enumerate(values):
        if isinstance(value, bool) or not isinstance(value, int) or value < -verify_report.U64_MAX or value > verify_report.U64_MAX:
            fail(f"{path}.samples[{index}]", "expected a signed checked endpoint delta")
        ordered.append(float(value))
    ordered.sort()
    count = len(ordered)
    result = {
        "min": ordered[0],
        "p50": (ordered[(count - 1) // 2] + ordered[count // 2]) / 2.0,
        "p95": ordered[min(((95 * count + 99) // 100) - 1, count - 1)],
        "p99": ordered[min(((99 * count + 99) // 100) - 1, count - 1)],
        "max": ordered[-1],
        "mean": sum(ordered) / count,
        "samples": list(values),
    }
    for key in ("min", "p50", "p95", "p99", "max"):
        if isinstance(result[key], float) and result[key].is_integer():
            result[key] = int(result[key])
    return result


def integer_stats(values: list[int], path: str) -> dict[str, Any]:
    result = stats(values, path)
    for key in ("min", "p50", "p95", "p99", "max"):
        value = result[key]
        if isinstance(value, float) and value.is_integer():
            result[key] = int(value)
    return result


def read_resource_rss(path: Path, label: str) -> dict[str, Any]:
    candidates = [path]
    if not path.is_file() and path.with_suffix(path.suffix + ".gz").is_file():
        candidates = [path.with_suffix(path.suffix + ".gz")]
    selected = candidates[0]
    if not selected.is_file():
        fail(label, "resource log is missing")
    try:
        if selected.name.endswith(".gz"):
            with gzip.open(selected, "rt", encoding="utf-8", errors="strict") as stream:
                lines = stream.read().splitlines()
        else:
            lines = selected.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(label, f"cannot read resource log: {error}")
    matches = []
    for line in lines:
        normalized = line.lstrip()
        match = RSS_RE.match(normalized)
        if match is not None:
            matches.append(int(match.group(1)))
    if len(matches) != 1:
        fail(label, "must contain exactly one GNU time maximum-resident-set-size line")
    kib = matches[0]
    if kib > verify_report.U64_MAX // 1024:
        fail(label, "RSS value overflows bytes")
    # Keep the logical artifact name stable when seal.py losslessly changes
    # the storage representation from .log to .log.gz.
    return {"status": "measured", "scope": "gnu_time_v_verbose_whole_fresh_process", "kib": kib, "bytes": kib * 1024, "path": path.name}


def receipt_for(root: Path, lane: dict[str, Any], name: str) -> tuple[Path, Path]:
    captures = root / "captures"
    report = captures / f"{name}.json"
    receipt = captures / f"{name}-receipt.json"
    if not report.is_file():
        fail(f"captures/{name}.json", "report is missing")
    if not receipt.is_file():
        fail(f"captures/{name}-receipt.json", "capture receipt is missing")
    value = obj(load(receipt, str(receipt)), f"captures/{name}-receipt.json")
    if value.get("change") != CHANGE or value.get("name") != name or value.get("lane") != lane:
        fail(f"captures/{name}-receipt.json", "receipt does not bind the frozen lane")
    if value.get("status") != "pass" or value.get("exit_code") != 0:
        fail(f"captures/{name}-receipt.json", "lane receipt is not a passing command receipt")
    artifacts = obj(value.get("artifacts"), f"captures/{name}-receipt.json.artifacts")
    report_key = f"captures/{name}.json"
    artifact = obj(artifacts.get(report_key), f"captures/{name}-receipt.json.artifacts.{report_key}")
    if artifact.get("sha256") != sha256_file(report, report_key) or artifact.get("bytes") != report.stat().st_size:
        fail(f"captures/{name}-receipt.json.artifacts.{report_key}", "report receipt hash/byte binding is stale")
    return report, receipt


def lane_name(lane: dict[str, Any]) -> str:
    return f"{lane['mode']}-{lane['shape']}-{lane['repeat'].lower()}"


def load_lanes(root: Path, protocol: dict[str, Any]) -> list[tuple[dict[str, Any], str, Path, dict[str, Any]]]:
    expected_order = array(protocol.get("order"), "protocol.order")
    if len(expected_order) != 12:
        fail("protocol.order", "must contain the twelve frozen lanes")
    for index, lane_value in enumerate(expected_order):
        lane = obj(lane_value, f"protocol.order[{index}]")
        if lane.get("mode") not in MODES or lane.get("shape") not in SHAPES or lane.get("repeat") not in REPEATS:
            fail(f"protocol.order[{index}]", "contains an unknown mode, shape, or repeat")
    index_path = root / "capture-index.json"
    if not index_path.is_file():
        fail("capture-index.json", "capture index is required before deriving the matrix")
    indexed = array(load(index_path, "capture-index.json"), "capture-index.json")
    if len(indexed) != len(expected_order):
        fail("capture-index.json", "must list all twelve receipts")
    rows = []
    for position, (lane_value, indexed_value) in enumerate(zip(expected_order, indexed)):
        lane = obj(lane_value, f"protocol.order[{position}]")
        indexed_name = indexed_value if isinstance(indexed_value, str) else ""
        expected = f"captures/{lane_name(lane)}-receipt.json"
        if indexed_name != expected:
            fail(f"capture-index.json[{position}]", f"must be {expected!r} in protocol order")
        name = lane_name(lane)
        report_path, _receipt_path = receipt_for(root, lane, name)
        verified = verify_report.validate_report(report_path, lane["mode"], lane["shape"])
        rows.append((lane, name, report_path, verified))
    return rows


def vector_from_metric(metric: dict[str, Any], path: str) -> list[int] | None:
    status = metric.get("status")
    if status != "measured":
        if "values" in metric:
            fail(path, "unavailable metric must not carry fabricated values")
        return None
    values = array(metric.get("values"), f"{path}.values")
    return [u64(value, f"{path}.values[{index}]") for index, value in enumerate(values)]


def build_row(root: Path, lane: dict[str, Any], name: str, report_path: Path, verified: dict[str, Any]) -> dict[str, Any]:
    result = verified["result"]
    corpus = verified["corpus"]
    elapsed = verified["elapsed"]
    metrics = verified["metrics"]
    samples = list(verified["elapsed"])
    sample_order = list(verified["sample_order"])
    elapsed_statistics = dict(verified["statistics"])
    elapsed_statistics["samples"] = samples
    elapsed_statistics["sample_order"] = sample_order
    row: dict[str, Any] = {
        "name": name,
        "repeat": lane["repeat"],
        "mode": lane["mode"],
        "shape": lane["shape"],
        "case": CASE,
        "rows": corpus["rows"],
        "cells": corpus["cells"],
        "corpus": {
            "archive_bytes": corpus["archive_bytes"],
            "archive_sha256": corpus["archive_hash"],
            "target_payload_bytes": corpus["target_payload_bytes"],
            "target_payload_sha256": result["corpus"]["target_payload_sha256"],
        },
        "output": {
            "archive_bytes": corpus["archive_bytes"],
            "archive_sha256": result["output_sha256"],
            "sink_accepted_bytes": verified["sink"]["accepted_bytes"],
        },
        "latency_ns": elapsed_statistics,
        "throughput": {},
        "process_rss": {},
        "allocation": {},
        "source_identity": {
            "revision": result_path_revision(verified),
            "binary_sha256": verified["report"]["binary_identity"]["binary_sha256"],
            "tool_instrumentation": verified["report"]["tool"]["instrumentation"],
        },
    }
    rows_per_second = [corpus["rows"] * 1_000_000_000.0 / value for value in samples]
    bytes_per_second = [corpus["archive_bytes"] * 1_000_000_000.0 / value for value in samples]
    row["throughput"] = {
        "rows_per_second": stats(rows_per_second, f"{name}.throughput.rows_per_second"),
        "archive_bytes_per_second": stats(bytes_per_second, f"{name}.throughput.archive_bytes_per_second"),
    }

    process = metrics["process"]
    process_rss = process["peak_rss_bytes"]
    operation_rss = vector_from_metric(process_rss, f"{name}.operation_metrics.process.peak_rss_bytes")
    resource = read_resource_rss(root / "captures" / f"{name}-resource.log", f"captures/{name}-resource.log")
    row["process_rss"] = {
        "process_lifetime_peak_observed_at_operation": ({"status": "measured", "scope": process_rss["scope"], "statistics": integer_stats(operation_rss, f"{name}.process_rss.process_lifetime_peak_observed_at_operation"), "samples": operation_rss} if operation_rss is not None else {"status": "unavailable", "scope": process_rss["scope"]}),
        "whole_process_max_rss": resource,
    }

    allocation = metrics["allocation"]
    if lane["mode"] == "normal" and allocation is None:
        row["allocation"] = {"status": "unavailable", "scope": verify_report.ALLOCATOR_SCOPE, "reason": "normal binary does not install the benchmark allocator"}
    elif allocation is None:
        fail(name, "allocator report omitted its measured allocation envelope")
    elif allocation["status"] != "measured":
        row["allocation"] = {"status": allocation["status"], "scope": allocation["scope"]}
    else:
        vectors = {}
        for field in verify_report.ALLOCATOR_FIELDS:
            values = vector_from_metric(allocation[field], f"{name}.allocation.{field}")
            assert values is not None
            vectors[field] = values
        entry_delta = [vectors["region_peak_live_bytes"][index] - vectors["live_bytes_before"][index] for index in range(SAMPLES)]
        exit_delta = [vectors["region_peak_live_bytes"][index] - vectors["live_bytes_after"][index] for index in range(SAMPLES)]
        row["allocation"] = {
            "status": "measured",
            "scope": allocation["scope"],
            "requested_bytes": {"source_field": "allocated_bytes", "statistics": integer_stats(vectors["allocated_bytes"], f"{name}.allocation.requested_bytes"), "samples": vectors["allocated_bytes"]},
            "requested_count": {"source_field": "allocation_calls", "statistics": integer_stats(vectors["allocation_calls"], f"{name}.allocation.requested_count"), "samples": vectors["allocation_calls"]},
            "regional_peak_minus_entry": {"source_fields": ["region_peak_live_bytes", "live_bytes_before"], "statistics": integer_stats(entry_delta, f"{name}.allocation.regional_peak_minus_entry"), "samples": entry_delta},
            "regional_peak_minus_live_exit": {"source_fields": ["region_peak_live_bytes", "live_bytes_after"], "statistics": integer_stats(exit_delta, f"{name}.allocation.regional_peak_minus_live_exit"), "samples": exit_delta},
            "live_exit_minus_entry": {"source_fields": ["live_bytes_after", "live_bytes_before"], "statistics": signed_stats([vectors["live_bytes_after"][index] - vectors["live_bytes_before"][index] for index in range(SAMPLES)], f"{name}.allocation.live_exit_minus_entry" )},
            "raw_vectors": vectors,
        }
    return row


def result_path_revision(verified: dict[str, Any]) -> str:
    return verified["report"]["environment"]["git_revision"]


def compare_repeats(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_key = {(row["mode"], row["shape"], row["repeat"]): row for row in rows}
    flags = []
    metrics = (
        ("latency_ns.p50", lambda row: row["latency_ns"]["p50"]),
        ("latency_ns.p95", lambda row: row["latency_ns"]["p95"]),
        ("latency_ns.p99", lambda row: row["latency_ns"]["p99"]),
        ("latency_ns.mean", lambda row: row["latency_ns"]["mean"]),
        ("throughput.rows_per_second.p50", lambda row: row["throughput"]["rows_per_second"]["p50"]),
        ("throughput.rows_per_second.p95", lambda row: row["throughput"]["rows_per_second"]["p95"]),
        ("throughput.rows_per_second.p99", lambda row: row["throughput"]["rows_per_second"]["p99"]),
        ("throughput.rows_per_second.mean", lambda row: row["throughput"]["rows_per_second"]["mean"]),
    )
    for mode in MODES:
        for shape in SHAPES:
            first = by_key[(mode, shape, "R1")]
            second = by_key[(mode, shape, "R2")]
            for metric_name, select in metrics:
                before = float(select(first))
                after = float(select(second))
                if before == 0:
                    fail(f"repeat_flags.{mode}.{shape}.{metric_name}", "R1 metric is zero")
                delta = (after - before) / before * 100.0
                flags.append({"mode": mode, "shape": shape, "metric": metric_name, "r1": before, "r2": after, "relative_percent": delta, "threshold_percent": 5.0, "flagged": abs(delta) > 5.0})
    return flags


def compare_identity(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    checks = []
    for shape in SHAPES:
        selected = [row for row in rows if row["shape"] == shape]
        if len(selected) != 4:
            fail(f"same_corpus_output.{shape}", "must contain both modes and both repeats")
        corpus_hashes = {row["corpus"]["archive_sha256"] for row in selected}
        output_hashes = {row["output"]["archive_sha256"] for row in selected}
        corpus_bytes = {row["corpus"]["archive_bytes"] for row in selected}
        output_bytes = {row["output"]["archive_bytes"] for row in selected}
        if len(corpus_hashes) != 1 or len(output_hashes) != 1 or len(corpus_bytes) != 1 or len(output_bytes) != 1 or corpus_hashes != output_hashes or corpus_bytes != output_bytes:
            fail(f"same_corpus_output.{shape}", "corpus/output archive identities differ across modes or repeats")
        checks.append({"shape": shape, "same_corpus_sha256": True, "same_output_sha256": True, "same_corpus_bytes": True, "same_output_bytes": True, "archive_sha256": next(iter(corpus_hashes)), "archive_bytes": next(iter(corpus_bytes))})
    return checks


def derive(root: Path) -> dict[str, Any]:
    protocol = obj(load(root / "protocol.json", "protocol.json"), "protocol.json")
    if protocol.get("change") != CHANGE or protocol.get("samples") != SAMPLES or protocol.get("warmups") != 3 or protocol.get("repeats") != 2 or protocol.get("workers") != 1 or protocol.get("cells_per_row") != 4:
        fail("protocol", "does not match the frozen 0432 matrix")
    rows_info = load_lanes(root, protocol)
    rows = [build_row(root, lane, name, path, verified) for lane, name, path, verified in rows_info]
    if len(rows) != 12 or {(row["mode"], row["shape"], row["repeat"]) for row in rows} != {(mode, shape, repeat) for mode in MODES for shape in SHAPES for repeat in REPEATS}:
        fail("rows", "matrix is incomplete or contains duplicate lanes")
    revisions = {row["source_identity"]["revision"] for row in rows}
    if len(revisions) != 1:
        fail("source_identity", "source revision differs across fresh-process lanes")
    summary = {
        "schema_version": 1,
        "change": CHANGE,
        "classification": protocol["classification"],
        "timing_scope": protocol["timing_scope"],
        "memory_scope": protocol["memory_scope"],
        "claims": [],
        "source_revision": next(iter(revisions)),
        "matrix": {"modes": list(MODES), "shapes": list(SHAPES), "repeats": list(REPEATS), "samples": SAMPLES, "warmups": protocol["warmups"], "cells_per_row": protocol["cells_per_row"]},
        "rows": rows,
        "same_corpus_output": compare_identity(rows),
        "repeat_flags": compare_repeats(rows),
    }
    return summary


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "summary.json")
    parser.add_argument("--check", action="store_true", help="recompute and compare the existing summary")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    args = parser.parse_args()
    root = args.root.resolve()
    try:
        summary = derive(root)
        if args.check:
            existing = load(args.output, str(args.output))
            if canonical(existing) != canonical(summary):
                fail(str(args.output), "does not match a fresh derivation")
        else:
            if args.output.exists():
                fail(str(args.output), "already exists; use --check to validate it explicitly")
            args.output.write_text(json.dumps(summary, indent=2, sort_keys=True, allow_nan=False) + "\n", encoding="utf-8")
    except (OSError, KeyError, TypeError, AssertionError, verify_report.VerificationError, DerivationError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    print("VALID")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
