#!/usr/bin/env python3
"""Derive the bounded ODP retention decision from sealed evidence.

This postprocessor consumes the independently validated ``summary.json`` and
the six formal profile receipts.  It reports every comparison and repeat flag,
allocator peak/call/byte deltas, RSS scope, and retained profiler artifact
bindings.  A retention-enabler decision is granted only when the measured
streaming region peak has the declared shape-growth property and the large
shape clears the explicit reduction threshold.  Latency flags are retained as
disclosures and never decide that gate.

The result is descriptive evidence for a bounded publication enabler.  It
does not make a fixed-memory, RSS, scaling, speedup, causal, or 10x claim.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import math
from pathlib import Path
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
ROLES = ("before-buffered", "after-buffered", "after-streaming")
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
REPEATS = ("R1", "R2")
PROFILE_KINDS = ("stat", "record")
SIGNIFICANT_PEAK_REDUCTION_PERCENT = 20.0
MAX_JSON_BYTES = 512 * 1024 * 1024
MAX_LOGICAL_ARTIFACT_BYTES = 2 * 1024 * 1024 * 1024


class DecisionError(ValueError):
    pass


def fail(message: str) -> None:
    raise DecisionError(message)


def reject_constant(value: str) -> Any:
    raise DecisionError(f"non-finite JSON value {value!r}")


def duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise DecisionError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def load_json(path: Path, label: str) -> Any:
    if not path.is_file():
        fail(f"{label}: file is missing")
    if path.stat().st_size > MAX_JSON_BYTES:
        fail(f"{label}: JSON exceeds {MAX_JSON_BYTES} bytes")
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=duplicate_pairs,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, DecisionError) as error:
        fail(f"{label}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected an object")
    return value


def array(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        fail(f"{label}: expected an array")
    return value


def finite(value: Any, label: str, *, signed: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        fail(f"{label}: expected a number")
    result = float(value)
    if not math.isfinite(result) or (not signed and result < 0):
        fail(f"{label}: expected a finite {'signed' if signed else 'nonnegative'} number")
    return result


def integer(value: Any, label: str, *, signed: bool = False) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or (not signed and value < 0):
        fail(f"{label}: expected an integer")
    return value


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def relative(bundle: Path, path: Path, label: str) -> str:
    try:
        value = path.resolve().relative_to(bundle.resolve()).as_posix()
    except ValueError:
        fail(f"{label}: path escapes the evidence bundle")
    return value


def bundle_path(bundle: Path, value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value or Path(value).is_absolute() or ".." in Path(value).parts:
        fail(f"{label}: expected a safe bundle-relative path")
    path = (bundle / value).resolve()
    if not path.is_relative_to(bundle.resolve()):
        fail(f"{label}: path escapes the evidence bundle")
    return path


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(
            value,
            sort_keys=True,
            separators=(",", ":"),
            ensure_ascii=False,
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot canonicalize result: {error}")
    raise AssertionError("unreachable")


def pct_change(old: float, new: float, label: str) -> float:
    if old == 0:
        fail(f"{label}: zero baseline")
    return (new - old) / old * 100.0


def row_key(row: dict[str, Any], label: str) -> tuple[str, str, str, str]:
    role = row.get("role")
    mode = row.get("mode")
    shape = row.get("shape")
    repeat = row.get("repeat")
    if role not in ROLES or mode not in MODES or shape not in SHAPES or repeat not in REPEATS:
        fail(f"{label}: malformed role/mode/shape/repeat")
    return role, mode, shape, repeat


def require_stat(row: dict[str, Any], section: str, key: str, label: str) -> float:
    value = obj(row.get(section), f"{label}.{section}").get(key)
    stat = obj(value, f"{label}.{section}.{key}")
    if "p50" not in stat:
        fail(f"{label}.{section}.{key}: p50 is missing")
    return finite(stat["p50"], f"{label}.{section}.{key}.p50", signed=True)


def elapsed_p50(row: dict[str, Any], label: str) -> float:
    elapsed = obj(row.get("elapsed_ns"), f"{label}.elapsed_ns")
    return finite(elapsed.get("p50"), f"{label}.elapsed_ns.p50")


def rss(row: dict[str, Any], label: str) -> float | None:
    memory = obj(row.get("process_memory"), f"{label}.process_memory")
    value = memory.get("gnu_time_peak_rss_bytes")
    if value is None:
        return None
    return finite(value, f"{label}.process_memory.gnu_time_peak_rss_bytes")


def allocation_p50(row: dict[str, Any], key: str, label: str) -> float | None:
    statistics = allocation_statistics(row, key, label)
    return None if statistics is None else statistics["p50"]


def allocation_statistics(row: dict[str, Any], key: str, label: str) -> dict[str, float] | None:
    allocation = obj(row.get("allocation"), f"{label}.allocation")
    if allocation.get("status") != "measured":
        return None
    section = "statistics_extra" if key == "region_peak_above_entry" or key in {"live_delta", "live_balance"} else "statistics"
    stat = obj(obj(allocation.get(section), f"{label}.allocation.{section}").get(key), f"{label}.allocation.{section}.{key}")
    return {
        bound: finite(stat.get(bound), f"{label}.allocation.{section}.{key}.{bound}", signed=True)
        for bound in ("min", "p50", "max")
    }


def index_rows(summary: dict[str, Any]) -> dict[tuple[str, str, str, str], dict[str, Any]]:
    matrix = obj(summary.get("matrix"), "summary.matrix")
    if matrix.get("formal_reports") != 36 or matrix.get("retained_samples") != 1080:
        fail("summary.matrix: expected 36 reports and 1080 samples")
    rows = array(summary.get("rows"), "summary.rows")
    if len(rows) != 36:
        fail("summary.rows: expected 36 rows")
    indexed: dict[tuple[str, str, str, str], dict[str, Any]] = {}
    for index, value in enumerate(rows):
        row = obj(value, f"summary.rows[{index}]")
        key = row_key(row, f"summary.rows[{index}]")
        if key in indexed:
            fail(f"summary.rows[{index}]: duplicate row {key}")
        indexed[key] = row
    expected = {
        (role, mode, shape, repeat)
        for role in ROLES
        for mode in MODES
        for shape in SHAPES
        for repeat in REPEATS
    }
    if set(indexed) != expected:
        fail("summary.rows: matrix keys differ from the frozen 36-row matrix")
    for key, row in indexed.items():
        if row.get("timing_scope") is None or row.get("process_memory", {}).get("scope") is None:
            fail(f"summary.rows.{key}: timing/RSS scope is missing")
    return indexed


def comparison_index(summary: dict[str, Any]) -> dict[tuple[str, str, str, str, str], dict[str, Any]]:
    values = array(summary.get("comparisons"), "summary.comparisons")
    result: dict[tuple[str, str, str, str, str], dict[str, Any]] = {}
    for index, value in enumerate(values):
        row = obj(value, f"summary.comparisons[{index}]")
        scope = row.get("scope")
        mode = row.get("mode")
        shape = row.get("shape")
        repeat = row.get("repeat")
        metric = row.get("metric")
        if not isinstance(scope, str) or mode not in MODES or shape not in SHAPES or repeat not in REPEATS or not isinstance(metric, str):
            fail(f"summary.comparisons[{index}]: malformed comparison identity")
        key = (scope, mode, shape, repeat, metric)
        if key in result:
            fail(f"summary.comparisons[{index}]: duplicate comparison {key}")
        if not isinstance(row.get("flagged"), bool):
            fail(f"summary.comparisons[{index}].flagged: expected boolean")
        result[key] = row
    if not isinstance(summary.get("review_flags"), list):
        fail("summary.review_flags: expected an array")
    return result


def pair_row(indexed: dict[tuple[str, str, str, str], dict[str, Any]], before: str, after: str, mode: str, shape: str, repeat: str) -> tuple[dict[str, Any], dict[str, Any]]:
    return indexed[(before, mode, shape, repeat)], indexed[(after, mode, shape, repeat)]


def pair_delta(indexed: dict[tuple[str, str, str, str], dict[str, Any]], comparisons: dict[tuple[str, str, str, str, str], dict[str, Any]], before: str, after: str, mode: str, shape: str, repeat: str, metric: str, getter: Any) -> dict[str, Any]:
    left, right = pair_row(indexed, before, after, mode, shape, repeat)
    old = getter(left, f"{before}.{mode}.{shape}.{repeat}")
    new = getter(right, f"{after}.{mode}.{shape}.{repeat}")
    if old is None or new is None:
        return {
            "status": "unavailable",
            "before_role": before,
            "after_role": after,
            "mode": mode,
            "shape": shape,
            "repeat": repeat,
            "metric": metric,
        }
    scope = "buffered_control" if before == "before-buffered" else "buffered_to_streaming_api"
    comparison = comparisons.get((scope, mode, shape, repeat, metric))
    change = pct_change(old, new, f"{scope}.{mode}.{shape}.{repeat}.{metric}")
    result = {
        "status": "measured",
        "scope": scope,
        "before_role": before,
        "after_role": after,
        "mode": mode,
        "shape": shape,
        "repeat": repeat,
        "metric": metric,
        "before": old,
        "after": new,
        "delta": new - old,
        "relative_percent": change,
        "reduction_percent": -change,
    }
    if comparison is not None:
        result["summary_comparison"] = copy_comparison(comparison)
    return result


def copy_comparison(value: dict[str, Any]) -> dict[str, Any]:
    # JSON round-tripping is unnecessary and would obscure non-finite errors;
    # the summary loader already rejects them.
    return dict(value)


def buffered_repeat_deltas(indexed: dict[tuple[str, str, str, str], dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for role in ("before-buffered", "after-buffered"):
        for shape in SHAPES:
            first = elapsed_p50(indexed[(role, "normal", shape, "R1")], f"{role}.normal.{shape}.R1")
            second = elapsed_p50(indexed[(role, "normal", shape, "R2")], f"{role}.normal.{shape}.R2")
            result.append({
                "scope": "same_api_buffered_repeat",
                "role": role,
                "shape": shape,
                "r1_p50_ns": first,
                "r2_p50_ns": second,
                "delta_ns": second - first,
                "relative_percent": pct_change(first, second, f"repeat.{role}.{shape}.p50"),
            })
    return result


def streaming_peak_consistency(indexed: dict[tuple[str, str, str, str], dict[str, Any]]) -> dict[str, Any]:
    values: list[dict[str, Any]] = []
    for shape in SHAPES:
        for repeat in REPEATS:
            row = indexed[("after-streaming", "allocator", shape, repeat)]
            statistics = allocation_statistics(row, "region_peak_above_entry", f"after-streaming.allocator.{shape}.{repeat}")
            if statistics is None:
                return {
                    "status": "not_evidenced",
                    "scope": "summary.statistics_extra.region_peak_above_entry.min/p50/max",
                    "claim": "six-row equality was not asserted because at least one allocator peak is unavailable",
                    "rows": values,
                }
            values.append({"shape": shape, "repeat": repeat, "min_bytes": statistics["min"], "p50_bytes": statistics["p50"], "max_bytes": statistics["max"]})
    equal = len({(row["min_bytes"], row["p50_bytes"], row["max_bytes"]) for row in values}) == 1
    exact = equal and all(row["min_bytes"] == row["p50_bytes"] == row["max_bytes"] for row in values)
    return {
        "status": "pass" if exact else "failed",
        "scope": "summary.statistics_extra.region_peak_above_entry.min/p50/max",
        "asserted": exact,
        "claim": "all six after-streaming allocator summary peaks have identical min=p50=max values" if exact else "after-streaming allocator summary peaks are not identical sampled constants",
        "rows": values,
    }


def shape_growth(indexed: dict[tuple[str, str, str, str], dict[str, Any]], role: str, repeat: str) -> dict[str, Any]:
    values = {
        shape: allocation_p50(indexed[(role, "allocator", shape, repeat)], "region_peak_above_entry", f"{role}.allocator.{shape}.{repeat}")
        for shape in SHAPES
    }
    if any(value is None for value in values.values()):
        return {"status": "unavailable", "role": role, "repeat": repeat, "values": values}
    tiny = float(values["tiny"])
    medium = float(values["medium"])
    large = float(values["large"])
    return {
        "status": "measured",
        "role": role,
        "repeat": repeat,
        "values": values,
        "tiny_to_medium_delta_bytes": medium - tiny,
        "medium_to_large_delta_bytes": large - medium,
    }


def growth_gate(indexed: dict[tuple[str, str, str, str], dict[str, Any]]) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    passed = True
    for repeat in REPEATS:
        control = shape_growth(indexed, "after-buffered", repeat)
        candidate = shape_growth(indexed, "after-streaming", repeat)
        if control["status"] != "measured" or candidate["status"] != "measured":
            passed = False
            rows.append({"repeat": repeat, "status": "unavailable", "control": control, "candidate": candidate})
            continue
        control_a = control["tiny_to_medium_delta_bytes"]
        control_b = control["medium_to_large_delta_bytes"]
        candidate_a = candidate["tiny_to_medium_delta_bytes"]
        candidate_b = candidate["medium_to_large_delta_bytes"]
        row_pass = candidate_a <= control_a and candidate_b <= control_b
        passed = passed and row_pass
        rows.append({
            "repeat": repeat,
            "status": "pass" if row_pass else "failed",
            "control": control,
            "candidate": candidate,
            "candidate_growth_no_more_than_control": row_pass,
        })
    return {
        "status": "pass" if passed else "failed",
        "definition": "candidate streaming region-peak p50 growth between adjacent shapes is no greater than after-buffered control for each repeat",
        "rows": rows,
    }


def live_delta_check(indexed: dict[tuple[str, str, str, str], dict[str, Any]]) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    passed = True
    for role in ROLES:
        for shape in SHAPES:
            for repeat in REPEATS:
                row = indexed[(role, "allocator", shape, repeat)]
                statistics = allocation_statistics(row, "live_delta", f"{role}.allocator.{shape}.{repeat}")
                ok = statistics is not None and statistics["min"] == statistics["p50"] == statistics["max"] == 0
                passed = passed and ok
                rows.append({"role": role, "shape": shape, "repeat": repeat, "min_bytes": None if statistics is None else statistics["min"], "p50_bytes": None if statistics is None else statistics["p50"], "max_bytes": None if statistics is None else statistics["max"], "zero": ok})
    return {"status": "pass" if passed else "failed", "scope": "summary.statistics_extra.live_delta.min/p50/max", "rows": rows}


def profile_digest(receipt: dict[str, Any], key: str, receipt_path: Path) -> str:
    value = obj(receipt.get(key), f"{receipt_path}.{key}").get("sha256")
    if not isinstance(value, str) or len(value) != 64 or any(char not in "0123456789abcdefABCDEF" for char in value):
        fail(f"{receipt_path}.{key}.sha256: malformed")
    return value.lower()


def retained_profiles(bundle: Path) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for role in ROLES:
        for kind in PROFILE_KINDS:
            receipt_path = bundle / "profiles" / role / kind / "receipt.json"
            receipt = obj(load_json(receipt_path, relative(bundle, receipt_path, "profile receipt")), str(receipt_path))
            if receipt.get("schema") != "litchi-0437-profile-receipt-v1" or receipt.get("status") != "pass":
                fail(f"{receipt_path}: formal profile is not passing")
            if receipt.get("role") != role or receipt.get("kind") != kind or receipt.get("preparatory") is not False:
                fail(f"{receipt_path}: profile role/kind/preparatory binding differs")
            scope = receipt.get("scope")
            if not isinstance(scope, str) or "whole fresh process" not in scope:
                fail(f"{receipt_path}: whole-process profile scope is missing")
            artifacts = obj(receipt.get("artifacts"), f"{receipt_path}.artifacts")
            required = {"report", "catalog", "resource", "workload_log", "oracle_log"}
            required.add("perf_stat" if kind == "stat" else "perf_data")
            if kind == "record":
                required.update({"perf_script", "perf_report"})
            if set(artifacts) != required:
                fail(f"{receipt_path}: profile artifact set differs")
            artifact_rows: dict[str, Any] = {}
            for name in sorted(required):
                record = obj(artifacts[name], f"{receipt_path}.artifacts.{name}")
                logical = integer(record.get("bytes"), f"{receipt_path}.artifacts.{name}.bytes")
                expected = record.get("sha256")
                if not isinstance(expected, str) or len(expected) != 64 or any(char not in "0123456789abcdefABCDEF" for char in expected):
                    fail(f"{receipt_path}.artifacts.{name}.sha256: malformed")
                raw_path = bundle_path(bundle, record.get("path"), f"{receipt_path}.artifacts.{name}.path")
                selected = raw_path if raw_path.is_file() else Path(str(raw_path) + ".gz")
                if not selected.is_file():
                    fail(f"{receipt_path}.artifacts.{name}: raw/gzip artifact is missing")
                digest = hashlib.sha256()
                count = 0
                try:
                    stream = gzip.open(selected, "rb") if selected.name.endswith(".gz") else selected.open("rb")
                    with stream:
                        for block in iter(lambda: stream.read(1024 * 1024), b""):
                            count += len(block)
                            if count > MAX_LOGICAL_ARTIFACT_BYTES:
                                fail(f"{selected}: logical artifact exceeds bounded size")
                            digest.update(block)
                except (OSError, EOFError, gzip.BadGzipFile) as error:
                    fail(f"{selected}: cannot read retained artifact: {error}")
                actual = digest.hexdigest()
                if count != logical or actual != expected.lower():
                    fail(f"{selected}: artifact hash or logical size differs from receipt")
                artifact_rows[name] = {
                    "logical_path": record["path"],
                    "logical_bytes": count,
                    "logical_sha256": actual,
                }
            result.append({
                "role": role,
                "kind": kind,
                "receipt": relative(bundle, receipt_path, "profile receipt"),
                "receipt_sha256": sha256_file(receipt_path),
                "source_manifest_sha256": profile_digest(receipt, "source_manifest", receipt_path),
                "binary_sha256": profile_digest(receipt, "binary", receipt_path),
                "scope": scope,
                "scope_correction": "The receipt scope text is retained verbatim; the independent copied Python oracle runs after the profiled executable exits and is outside the workload, perf, and GNU-time process scope.",
                "instrumentation": {
                    "stat_events": receipt.get("stat_events"),
                    "record_event": receipt.get("record_event"),
                    "record_frequency_hz": receipt.get("record_frequency_hz"),
                    "call_graph": receipt.get("call_graph"),
                },
                "artifacts": artifact_rows,
                "interpretation": "whole-process instrumentation retained for scope context; no causal hotspot attribution",
            })
    return result


def derive(bundle: Path, summary_path: Path, threshold: float) -> dict[str, Any]:
    summary = obj(load_json(summary_path, relative(bundle, summary_path, "summary")), "summary")
    indexed = index_rows(summary)
    comparisons = comparison_index(summary)
    repeat_deltas = buffered_repeat_deltas(indexed)

    buffered_control_p50 = [
        pair_delta(indexed, comparisons, "before-buffered", "after-buffered", "normal", shape, repeat, "p50", elapsed_p50)
        for shape in SHAPES for repeat in REPEATS
    ]
    cross_api_p50 = [
        pair_delta(indexed, comparisons, "after-buffered", "after-streaming", "normal", shape, repeat, "p50", elapsed_p50)
        for shape in SHAPES for repeat in REPEATS
    ]
    cross_api_ratios = []
    for item in cross_api_p50:
        if item["status"] != "measured":
            cross_api_ratios.append(dict(item, ratio=None))
        else:
            cross_api_ratios.append(dict(item, ratio=item["after"] / item["before"], ratio_definition="after-streaming p50 / after-buffered p50"))

    reductions: list[dict[str, Any]] = []
    for before, after in (("before-buffered", "after-buffered"), ("after-buffered", "after-streaming")):
        for shape in SHAPES:
            for repeat in REPEATS:
                for metric, key in (("allocation_calls", "allocation.allocation_calls"), ("allocated_bytes", "allocation.allocated_bytes"), ("region_peak_above_entry", "allocation.region_peak_above_entry")):
                    reductions.append(pair_delta(indexed, comparisons, before, after, "allocator", shape, repeat, key, lambda row, label, metric_name=metric: allocation_p50(row, metric_name, label)))

    rss_deltas = []
    for before, after in (("before-buffered", "after-buffered"), ("after-buffered", "after-streaming")):
        for mode in MODES:
            for shape in SHAPES:
                for repeat in REPEATS:
                    rss_deltas.append(pair_delta(indexed, comparisons, before, after, mode, shape, repeat, "gnu_time_peak_rss_bytes", rss))

    live_delta = live_delta_check(indexed)
    peak_consistency = streaming_peak_consistency(indexed)
    growth = growth_gate(indexed)
    large_rows = [
        pair_delta(indexed, comparisons, "after-buffered", "after-streaming", "allocator", "large", repeat, "allocation.region_peak_above_entry", lambda row, label: allocation_p50(row, "region_peak_above_entry", label))
        for repeat in REPEATS
    ]
    large_reduction = {
        "threshold_percent": threshold,
        "rows": large_rows,
        "each_repeat_at_or_above_threshold": all(row.get("status") == "measured" and row.get("reduction_percent", -math.inf) >= threshold for row in large_rows),
    }
    gate = {
        "peak_growth_no_more_than_control": growth["status"] == "pass",
        "large_peak_reduction": large_reduction["each_repeat_at_or_above_threshold"],
        "streaming_peak_summary_p50_consistency": peak_consistency["status"] == "pass",
        "allocator_live_delta_zero": live_delta["status"] == "pass",
    }
    retention_status = "pass" if all(gate.values()) else "withheld"

    review_flags = array(summary.get("review_flags"), "summary.review_flags")
    return {
        "schema": "litchi-0437-odp-retention-decision-v1",
        "change": 437,
        "summary_path": relative(bundle, summary_path, "summary"),
        "summary_sha256": sha256_file(summary_path),
        "summary_protocol_sha256": summary.get("protocol_sha256"),
        "decision": {
            "status": retention_status,
            "scope": "measured bounded publication enabler only",
            "criteria": gate,
            "significant_peak_reduction_threshold_percent": threshold,
            "latency_flags_do_not_gate_retention": True,
            "no_claims": [
                "universal fixed memory",
                "fixed RSS",
                "scaling beyond the measured shapes",
                "speedup or throughput improvement",
                "causal hotspot attribution",
                "10x improvement",
            ],
        },
        "buffered_same_api_p50_deltas": buffered_control_p50,
        "buffered_same_api_repeat_p50_deltas": repeat_deltas,
        "cross_api_streaming_over_buffered_p50": cross_api_ratios,
        "allocator_reductions": reductions,
        "process_rss_deltas": rss_deltas,
        "peak_shape_growth": growth,
        "streaming_allocator_peak_consistency": peak_consistency,
        "allocator_live_delta": live_delta,
        "large_shape_peak_reduction": large_reduction,
        "all_comparisons": array(summary.get("comparisons"), "summary.comparisons"),
        "all_review_flags": review_flags,
        "review_flag_count": len(review_flags),
        "profiles": retained_profiles(bundle),
        "profile_scope": "retained profiler artifacts cover the whole profiled fresh executable process, including setup, corpus generation, warmups, samples, and output hashing; the independent copied Python oracle runs after that process and is outside the workload, perf, and GNU-time scope",
        "evidence_limits": [
            "allocator peak values are operation-region values as summarized by the report contract",
            "GNU time RSS is whole-process high-water scope and is reported separately",
            "streaming peak consistency is asserted from summary min=p50=max statistics",
            "latency regressions and repeat drift remain fully listed in all_review_flags",
        ],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bundle", type=Path, default=ROOT)
    parser.add_argument("--summary", type=Path, default=None)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument("--threshold-percent", type=float, default=SIGNIFICANT_PEAK_REDUCTION_PERCENT)
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args(argv)
    bundle = args.bundle.resolve()
    summary_path = (args.summary or (bundle / "summary.json")).resolve()
    output = (args.output or (bundle / "retention-decision.json")).resolve()
    try:
        if not bundle.is_dir() or not output.is_relative_to(bundle):
            fail("bundle/output paths must name an existing bundle and an output inside it")
        threshold = finite(args.threshold_percent, "--threshold-percent")
        if threshold <= 0 or threshold >= 100:
            fail("--threshold-percent must be between 0 and 100")
        value = derive(bundle, summary_path, threshold)
        if args.check:
            if not output.is_file() or canonical(load_json(output, "retention-decision.json")) != canonical(value):
                fail("retained retention-decision.json differs from independent derivation")
        if args.write:
            output.write_bytes(canonical(value) + b"\n")
        print(json.dumps(value, indent=2, sort_keys=True, allow_nan=False))
    except (OSError, TypeError, ValueError, DecisionError, KeyError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
