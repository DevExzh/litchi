#!/usr/bin/env python3
"""Derive a lossless matched summary for the 0433 ODS role matrix."""

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
_spec = importlib.util.spec_from_file_location("ods_report", ROOT / "verify-report.py")
if _spec is None or _spec.loader is None:
    raise RuntimeError("cannot load the independent report verifier")
verify_report = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(verify_report)

CHANGE = 433
ROLES = ("before-buffered", "after-buffered", "after-streaming")
BUILD_DIR = {"before-buffered": "before", "after-buffered": "after", "after-streaming": "after"}
MODES = ("normal", "allocator")
REPEATS = ("r1", "r2")
RSS_RE = re.compile(r"^Maximum resident set size \(kbytes\):\s*(\d+)\s*$")


class DerivationError(ValueError):
    pass


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


def finite(value: Any, path: str, *, signed: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        fail(path, "expected a finite number")
    result = float(value)
    if not math.isfinite(result) or (not signed and result < 0):
        fail(path, "expected a finite number in the permitted range")
    return result


def load(path: Path, label: str) -> Any:
    return verify_report.load_json(path, label)


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode("utf-8")


def sha256_file(path: Path, label: str) -> str:
    if not path.is_file():
        fail(label, "file is missing")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def stats(values: list[int | float], path: str, *, signed: bool = False) -> dict[str, Any]:
    if not values:
        fail(path, "cannot summarize an empty vector")
    ordered = sorted(finite(value, f"{path}.samples[{index}]", signed=signed) for index, value in enumerate(values))
    count = len(ordered)
    midpoint = (ordered[(count - 1) // 2] + ordered[count // 2]) / 2.0
    p95 = ordered[min(((95 * count + 99) // 100) - 1, count - 1)]
    p99 = ordered[min(((99 * count + 99) // 100) - 1, count - 1)]
    result = {"min": ordered[0], "p50": midpoint, "p95": p95, "p99": p99, "max": ordered[-1], "mean": sum(ordered) / count, "samples": list(values)}
    for key in ("min", "p50", "p95", "p99", "max"):
        if isinstance(result[key], float) and result[key].is_integer():
            result[key] = int(result[key])
    return result


def metric_values(metric: Any, path: str) -> list[int] | None:
    value = obj(metric, path)
    status = value.get("status")
    if status != "measured":
        if "values" in value:
            fail(path, "non-measured metric must not carry values")
        return None
    values = array(value.get("values"), f"{path}.values")
    return [u64(item, f"{path}.values[{index}]") for index, item in enumerate(values)]


def read_rss(path: Path, label: str) -> dict[str, Any]:
    selected = path
    if not selected.is_file() and path.with_suffix(path.suffix + ".gz").is_file():
        selected = path.with_suffix(path.suffix + ".gz")
    if not selected.is_file():
        fail(label, "GNU time resource log is missing")
    try:
        if selected.suffix == ".gz":
            with gzip.open(selected, "rt", encoding="utf-8") as stream:
                lines = stream.read().splitlines()
        else:
            lines = selected.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(label, f"cannot read resource log: {error}")
    matches = [int(match.group(1)) for line in lines if (match := RSS_RE.match(line.lstrip()))]
    if len(matches) != 1:
        fail(label, "must contain exactly one GNU time maximum-resident-set-size line")
    if matches[0] > verify_report.U64_MAX // 1024:
        fail(label, "RSS value overflows bytes")
    return {"status": "measured", "scope": "gnu_time_v_verbose_whole_fresh_process", "bytes": matches[0] * 1024, "kib": matches[0], "path": path.name}


def build_root(role: str) -> Path:
    return ROOT / BUILD_DIR[role]


def lane_name(role: str, lane: dict[str, Any]) -> str:
    return f"{role}-{lane['mode']}-{lane['shape']}-{str(lane['repeat']).lower()}"


def receipt_for(role: str, lane: dict[str, Any], name: str) -> tuple[Path, dict[str, Any]]:
    root = build_root(role)
    receipt_path = root / "captures" / f"{name}-receipt.json"
    report_path = root / "captures" / f"{name}.json"
    if not receipt_path.is_file() or not report_path.is_file():
        fail(f"{role}.captures.{name}", "report or receipt is missing")
    receipt = obj(load(receipt_path, str(receipt_path)), str(receipt_path))
    if receipt.get("role") != role or receipt.get("name") != name or receipt.get("lane") != lane:
        fail(str(receipt_path), "receipt does not bind the frozen role/lane")
    if receipt.get("status") != "pass" or receipt.get("exit_code") != 0:
        fail(str(receipt_path), "capture receipt is not passing")
    artifacts = obj(receipt.get("artifacts"), str(receipt_path))
    artifact = obj(
        artifacts.get(f"{BUILD_DIR[role]}/captures/{name}.json"),
        f"{receipt_path}.artifacts.report",
    )
    if artifact.get("sha256") != sha256_file(report_path, str(report_path)) or artifact.get("bytes") != report_path.stat().st_size:
        fail(str(receipt_path), "report artifact binding is stale")
    return report_path, receipt


def load_role_rows(protocol: dict[str, Any], role: str) -> list[dict[str, Any]]:
    order = array(protocol.get("order"), "protocol.order")
    index_path = build_root(role) / f"capture-index-{role}.json"
    if not index_path.is_file():
        # Permit the first role's legacy index while a pre-freeze capture is
        # being inspected; formal captures use the role-specific name so the
        # two after roles can share one build/captures directory.
        index_path = build_root(role) / "capture-index.json"
    indexed = array(load(index_path, str(index_path)), str(index_path))
    if len(indexed) != len(order):
        fail(str(index_path), "does not list every frozen lane")
    rows = []
    for position, lane_value in enumerate(order):
        lane = obj(lane_value, f"protocol.order[{position}]")
        expected_name = f"captures/{lane_name(role, lane)}-receipt.json"
        if indexed[position] != expected_name:
            fail(f"{index_path}[{position}]", f"must be {expected_name!r}")
        report_path, receipt = receipt_for(role, lane, lane_name(role, lane))
        verified = verify_report.validate_report(report_path, lane["mode"], lane["shape"], role)
        build = obj(load(build_root(role) / "build.json", f"{role}/build.json"), f"{role}/build.json")
        binary = obj(build["binaries"][lane["mode"]], f"{role}.build.binaries.{lane['mode']}")
        if receipt.get("revision") != build.get("revision") or receipt.get("source_manifest") != build.get("source_manifest"):
            fail(str(report_path), "receipt source custody differs from the role build")
        if receipt.get("binary") != binary:
            fail(str(report_path), "receipt binary identity differs from the role build")
        if receipt.get("protocol_sha256") != build.get("protocol_sha256") or receipt.get("verifier_sha256") != build.get("verifier_sha256"):
            fail(str(report_path), "receipt protocol/verifier custody differs from the role build")
        if verified["report"]["environment"]["git_revision"] != build["revision"]:
            fail(str(report_path), "report revision differs from role build")
        if verified["report"]["binary_identity"]["path"] != binary["path"] or verified["report"]["binary_identity"]["binary_sha256"] != binary["sha256"] or verified["report"]["binary_identity"]["binary_bytes"] != binary["bytes"]:
            fail(str(report_path), "report binary identity differs from receipt/build")
        resource = read_rss(build_root(role) / "captures" / f"{lane_name(role, lane)}-resource.log", f"{role}.{lane_name(role, lane)}.resource")
        rows.append({"role": role, "lane": lane, "name": lane_name(role, lane), "report_path": report_path, "receipt": receipt, "verified": verified, "resource": resource, "build": build})
    return rows


def semantic_key(verified: dict[str, Any]) -> dict[str, Any]:
    metadata = verified.get("source")
    if metadata is not None:
        metadata = obj(metadata, "report.results[0].source")
        oracle = obj(metadata["ods_scalar_rows"], "report.results[0].source.ods_scalar_rows")
        return {
            "rows": u64(oracle["rows_per_sheet"], "ods_scalar_rows.rows_per_sheet"),
            "cells_per_row": u64(oracle["columns_per_sheet"], "ods_scalar_rows.columns_per_sheet"),
            "cells": u64(oracle["rows_per_sheet"], "ods_scalar_rows.rows_per_sheet")
            * u64(oracle["columns_per_sheet"], "ods_scalar_rows.columns_per_sheet"),
            "semantic_sha256": verify_report.digest(oracle["semantic_sha256"], "ods_scalar_rows.semantic_sha256"),
            "sheet_count": u64(oracle["sheet_count"], "ods_scalar_rows.sheet_count"),
            "scalar_columns": list(oracle["scalar_columns"]),
        }
    fail("semantic_oracle", "validated report has no ODS source oracle")
    raise AssertionError("unreachable")


def build_summary_row(row: dict[str, Any]) -> dict[str, Any]:
    verified = row["verified"]
    result = verified["result"]
    corpus = verified["corpus"]
    samples = list(verified["elapsed"])
    rows = corpus["rows"]
    elapsed = stats(samples, f"{row['name']}.latency_ns")
    elapsed["sample_order"] = list(verified["sample_order"])
    metrics = verified["metrics"]
    process = metrics["process"]
    operation_peak = metric_values(process["peak_rss_bytes"], f"{row['name']}.process.peak_rss_bytes")
    process_rss = {"operation_lifetime_peak_observed_at_operation": {"status": "measured", "scope": process["peak_rss_bytes"]["scope"], "statistics": stats(operation_peak, f"{row['name']}.operation_peak_rss")} if operation_peak is not None else {"status": "unavailable", "scope": process["peak_rss_bytes"]["scope"]}, "whole_process_max_rss": row["resource"]}
    allocation = metrics.get("allocation")
    allocation_summary: dict[str, Any]
    if allocation is None:
        allocation_summary = {"status": "unavailable", "scope": verify_report.ALLOCATOR_SCOPE, "reason": "normal binary does not publish allocator vectors"}
    elif allocation["status"] != "measured":
        allocation_summary = {"status": allocation["status"], "scope": allocation["scope"]}
    else:
        vectors = {field: metric_values(allocation[field], f"{row['name']}.allocation.{field}") for field in verify_report.ALLOCATOR_FIELDS}
        if any(value is None for value in vectors.values()):
            fail(row["name"], "allocator vector became unavailable after report validation")
        typed = {key: value for key, value in vectors.items() if value is not None}
        allocation_summary = {
            "status": "measured", "scope": allocation["scope"],
            "requested_bytes": {"source_field": "allocated_bytes", "statistics": stats(typed["allocated_bytes"], f"{row['name']}.requested_bytes")},
            "requested_count": {"source_field": "allocation_calls", "statistics": stats(typed["allocation_calls"], f"{row['name']}.requested_count")},
            "regional_peak_minus_entry": {"statistics": stats([typed["region_peak_live_bytes"][i] - typed["live_bytes_before"][i] for i in range(len(samples))], f"{row['name']}.region_minus_entry")},
            "regional_peak_minus_live_exit": {"statistics": stats([typed["region_peak_live_bytes"][i] - typed["live_bytes_after"][i] for i in range(len(samples))], f"{row['name']}.region_minus_exit")},
            "live_exit_minus_entry": {"statistics": stats([typed["live_bytes_after"][i] - typed["live_bytes_before"][i] for i in range(len(samples))], f"{row['name']}.live_exit_minus_entry", signed=True)},
            "raw_vectors": typed,
        }
    return {
        "role": row["role"], "selector": row["verified"]["selector"], "mode": row["lane"]["mode"], "shape": row["lane"]["shape"], "repeat": row["lane"]["repeat"], "name": row["name"],
        "rows": rows, "cells": corpus["cells"], "semantic_oracle": semantic_key(verified),
        "corpus": {"archive_bytes": corpus["archive_bytes"], "archive_sha256": corpus["archive_hash"], "target_payload_bytes": corpus["target_payload_bytes"], "target_payload_sha256": result["corpus"]["target_payload_sha256"]},
        "output": {"archive_sha256": result["output_sha256"], "archive_bytes": corpus["archive_bytes"]},
        "retention": {"status": "fixed" if "retained_authoring_window_bytes" in verified["sink"] else "nonfixed", "window_bytes": verified["sink"].get("retained_authoring_window_bytes")},
        "latency_ns": elapsed,
        "throughput": {"rows_per_second": stats([rows * 1_000_000_000.0 / value for value in samples], f"{row['name']}.rows_per_second"), "archive_bytes_per_second": stats([corpus["archive_bytes"] * 1_000_000_000.0 / value for value in samples], f"{row['name']}.bytes_per_second")},
        "process_memory": process_rss, "allocation": allocation_summary,
        "source_identity": {
            "revision": verified["report"]["environment"]["git_revision"],
            "binary_sha256": verified["report"]["binary_identity"]["binary_sha256"],
            "source_manifest_sha256": verify_report.digest(row["build"]["source_manifest"]["sha256"], f"{row['name']}.source_manifest.sha256"),
            "source_manifest_files": u64(row["build"]["source_manifest"]["files"], f"{row['name']}.source_manifest.files"),
        },
    }


def relative(before: float, after: float, path: str) -> float:
    if before == 0:
        fail(path, "cannot compare against a zero baseline")
    return (after - before) / before * 100.0


def repeat_flags(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    flags = []
    # Elapsed time and throughput are claims for the normal binary only.  The
    # allocator binary retains raw elapsed vectors in each row for audit and
    # descriptive context, but those timings are not repeat/regression flags.
    normal_metrics = (
        ("latency_ns.p50", lambda r: r["latency_ns"]["p50"]),
        ("latency_ns.p95", lambda r: r["latency_ns"]["p95"]),
        ("latency_ns.p99", lambda r: r["latency_ns"]["p99"]),
        ("latency_ns.mean", lambda r: r["latency_ns"]["mean"]),
        ("throughput.rows_per_second.p50", lambda r: r["throughput"]["rows_per_second"]["p50"]),
        ("throughput.rows_per_second.mean", lambda r: r["throughput"]["rows_per_second"]["mean"]),
    )
    allocator_metrics = (
        ("allocation.requested_bytes.p50", lambda r: r["allocation"]["requested_bytes"]["statistics"]["p50"]),
        ("allocation.requested_count.p50", lambda r: r["allocation"]["requested_count"]["statistics"]["p50"]),
        ("allocation.regional_peak_minus_entry.p50", lambda r: r["allocation"]["regional_peak_minus_entry"]["statistics"]["p50"]),
    )
    grouped = {(row["role"], row["mode"], row["shape"], str(row["repeat"]).lower()): row for row in rows}
    for role in ROLES:
        for mode in MODES:
            shapes = {row["shape"] for row in rows if row["role"] == role and row["mode"] == mode}
            for shape in sorted(shapes):
                first = grouped.get((role, mode, shape, "r1")); second = grouped.get((role, mode, shape, "r2"))
                if first is None or second is None:
                    fail(f"repeat_flags.{role}.{mode}.{shape}", "both repeats are required")
                metrics = normal_metrics if mode == "normal" else allocator_metrics
                for name, select in metrics:
                    r1 = float(select(first)); r2 = float(select(second))
                    flags.append({"role": role, "mode": mode, "shape": shape, "metric": name, "r1": r1, "r2": r2, "relative_percent": relative(r1, r2, f"repeat_flags.{role}.{mode}.{shape}.{name}"), "threshold_percent": 5.0, "flagged": abs(relative(r1, r2, f"repeat_flags.{role}.{mode}.{shape}.{name}")) > 5.0})
    return flags


def comparisons(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    indexed = {(row["role"], row["mode"], row["shape"], str(row["repeat"]).lower()): row for row in rows}
    output = []
    for after_role in ("after-buffered", "after-streaming"):
        for mode in MODES:
            shapes = {row["shape"] for row in rows if row["role"] == after_role and row["mode"] == mode}
            for shape in sorted(shapes):
                for repeat in REPEATS:
                    before = indexed.get(("before-buffered", mode, shape, repeat)); after = indexed.get((after_role, mode, shape, repeat))
                    if before is None or after is None:
                        fail(f"comparisons.{after_role}.{mode}.{shape}.{repeat}", "matched before/after lane is missing")
                    if mode == "normal":
                        metrics = (
                            ("latency_ns.p50", before["latency_ns"]["p50"], after["latency_ns"]["p50"], True),
                            ("latency_ns.p95", before["latency_ns"]["p95"], after["latency_ns"]["p95"], True),
                            ("latency_ns.p99", before["latency_ns"]["p99"], after["latency_ns"]["p99"], True),
                            ("latency_ns.mean", before["latency_ns"]["mean"], after["latency_ns"]["mean"], True),
                            ("throughput.rows_per_second.p50", before["throughput"]["rows_per_second"]["p50"], after["throughput"]["rows_per_second"]["p50"], False),
                        )
                    else:
                        metrics = (
                            ("allocation.requested_bytes.p50", before["allocation"].get("requested_bytes", {}).get("statistics", {}).get("p50"), after["allocation"].get("requested_bytes", {}).get("statistics", {}).get("p50"), True),
                            ("allocation.requested_count.p50", before["allocation"].get("requested_count", {}).get("statistics", {}).get("p50"), after["allocation"].get("requested_count", {}).get("statistics", {}).get("p50"), True),
                            ("allocation.regional_peak_minus_entry.p50", before["allocation"].get("regional_peak_minus_entry", {}).get("statistics", {}).get("p50"), after["allocation"].get("regional_peak_minus_entry", {}).get("statistics", {}).get("p50"), True),
                        )
                    for metric, old, new, higher_is_regression in metrics:
                        if old is None or new is None:
                            continue
                        delta = relative(float(old), float(new), f"comparisons.{after_role}.{mode}.{shape}.{repeat}.{metric}")
                        flagged = delta > 5.0 if higher_is_regression else delta < -5.0
                        output.append({"baseline_role": "before-buffered", "candidate_role": after_role, "mode": mode, "shape": shape, "repeat": repeat, "metric": metric, "before": old, "after": new, "relative_percent": delta, "threshold_percent": 5.0, "higher_is_regression": higher_is_regression, "regression": flagged})
    return output


def derive() -> dict[str, Any]:
    protocol = obj(load(ROOT / "protocol.json", "protocol.json"), "protocol.json")
    if u64(protocol.get("change"), "protocol.change") != CHANGE:
        fail("protocol.change", "does not identify 0433")
    rows = []
    for role in ROLES:
        rows.extend(build_summary_row(row) for row in load_role_rows(protocol, role))
    if not rows:
        fail("rows", "capture matrix is empty")
    semantic_by_shape: dict[str, dict[str, Any]] = {}
    output_by_role_shape: dict[tuple[str, str], tuple[str, str]] = {}
    for row in rows:
        key = semantic_key_from_summary(row)
        old = semantic_by_shape.setdefault(row["shape"], key)
        if old != key:
            fail(f"semantic_oracle.{row['shape']}", "semantic rows/cells differ across roles, modes, or repeats")
        output_key = (row["role"], row["shape"])
        output_identity = (row["corpus"]["archive_sha256"], row["output"]["archive_sha256"])
        previous_output = output_by_role_shape.setdefault(output_key, output_identity)
        if previous_output != output_identity:
            fail(f"output_identity.{row['role']}.{row['shape']}", "corpus/output identity differs across modes or repeats")
    revisions = {row["role"]: {row["source_identity"]["revision"] for row in rows if row["role"] == role} for role in ROLES}
    for role, values in revisions.items():
        if len(values) != 1:
            fail(f"source_identity.{role}", "source revision differs within a role")
    return {"schema_version": 1, "change": CHANGE, "classification": protocol.get("classification", "descriptive matched ODS row creation evidence"), "timing_scope": protocol.get("timing_scope", ""), "preservation_scope": protocol.get("preservation_scope", ""), "comparison_scope": protocol.get("comparison_scope", ""), "claims": [], "matrix": {"roles": list(ROLES), "modes": list(MODES), "samples": protocol["samples"], "warmups": protocol["warmups"], "cells_per_row": protocol["cells_per_row"]}, "source_revisions": {role: next(iter(values)) for role, values in revisions.items()}, "rows": rows, "repeat_flags": repeat_flags(rows), "comparisons": comparisons(rows), "semantic_oracle_by_shape": semantic_by_shape}


def semantic_key_from_summary(row: dict[str, Any]) -> dict[str, Any]:
    key = dict(row["semantic_oracle"])
    key.pop("target_payload_sha256", None)
    return key


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "summary.json")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    try:
        summary = derive()
        if args.check:
            existing = load(args.output, str(args.output))
            if canonical(existing) != canonical(summary):
                fail(str(args.output), "does not match fresh derivation")
        else:
            if args.output.exists():
                fail(str(args.output), "already exists; use --check to validate")
            args.output.write_text(json.dumps(summary, indent=2, sort_keys=True, allow_nan=False) + "\n", encoding="utf-8")
    except (OSError, KeyError, TypeError, AssertionError, verify_report.VerificationError, DerivationError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    print("VALID")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
