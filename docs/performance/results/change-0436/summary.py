#!/usr/bin/env python3
"""Derive the descriptive 0436 ODT two-role ABBA summary.

The outer verifier owns report, corpus, source, repeat, custody, and profile
validation.  This file only derives statistics from that verified envelope. It
keeps the before-streaming -> after-streaming same-API comparison descriptive. It emits review flags;
it does not turn one matrix into a speedup, causality, cache-miss, or 10x
claim.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import random
import re
import sys
from typing import Any, Callable


ROOT = Path(__file__).resolve().parent
ROLES = ("before-streaming", "after-streaming")
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
REPEATS = ("R1", "R2")
THRESHOLD = 5.0
BOOTSTRAP_RESAMPLES = 2_000
BOOTSTRAP_SEED = 436
PROFILE_BUILD_DIR = {"before-streaming": "before", "after-streaming": "after"}
PROFILE_KINDS = ("stat", "record")
PROFILE_EVENTS = ("cycles:u", "instructions:u", "branches:u", "branch-misses:u", "L1-dcache-load-misses:u")
STAT_NUMBER = re.compile(r"^[0-9]+$")
LOST_RE = re.compile(r"^#\s*Total Lost Samples:\s*([0-9,]+)\s*$")
EVENT_HEADER_RE = re.compile(r"^#\s*Samples:\s+(.+?)\s+of event '([^']+)'\s*$")
EVENT_COUNT_RE = re.compile(r"^#\s*Event count \(approx\.\):\s*([0-9,]+)\s*$")
REPORT_ROW_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)%\s+\S+\s+\S+\s+(.+?)\s*$")
ADDR2LINE_TEXT = "could not read first record"


class SummaryError(ValueError):
    pass


def load_module(filename: str, name: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / filename)
    if spec is None or spec.loader is None:
        raise SummaryError(f"cannot load {filename}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


verify = load_module("verify.py", "change0436_verify_for_summary")


def finite(value: Any, label: str, *, signed: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SummaryError(f"{label}: expected a number")
    number = float(value)
    if not math.isfinite(number) or (not signed and number < 0):
        raise SummaryError(f"{label}: expected a finite number")
    return number


def integral(value: float | int) -> int | float:
    return int(value) if isinstance(value, float) and value.is_integer() else value


def stats(values: list[int | float], label: str, *, signed: bool = False) -> dict[str, Any]:
    if not values:
        raise SummaryError(f"{label}: empty vector")
    ordered = sorted(finite(value, f"{label}[{i}]", signed=signed) for i, value in enumerate(values))
    count = len(ordered)
    median = (ordered[(count - 1) // 2] + ordered[count // 2]) / 2.0
    p95 = ordered[min(((95 * count + 99) // 100) - 1, count - 1)]
    p99 = ordered[min(((99 * count + 99) // 100) - 1, count - 1)]
    return {
        "count": count,
        "min": integral(ordered[0]),
        "p50": integral(median),
        "p95": integral(p95),
        "p99": integral(p99),
        "max": integral(ordered[-1]),
        "mean": integral(sum(ordered) / count),
    }


def bootstrap(values: list[int], label: str, statistic: Callable[[list[int]], float], name: str) -> dict[str, Any]:
    if not values:
        raise SummaryError(f"{label}: empty bootstrap input")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates: list[float] = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        sample = [values[rng.randrange(len(values))] for _ in values]
        estimates.append(float(statistic(sample)))
    estimates.sort()
    result: dict[str, Any] = {
        "method": f"bootstrap_{name}_percentile",
        "resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED,
        "level": 0.95,
        "low": integral(estimates[int(0.025 * len(estimates))]),
        "high": integral(estimates[int(0.975 * len(estimates)) - 1]),
    }
    return result


def bootstrap_median(values: list[int], label: str) -> dict[str, Any]:
    return bootstrap(values, label, lambda sample: stats(sample, label, signed=False)["p50"], "median")


def bootstrap_mean(values: list[int], label: str) -> dict[str, Any]:
    return bootstrap(values, label, lambda sample: sum(sample) / len(sample), "mean")


def metric_values(metric: Any, label: str) -> list[int] | None:
    if not isinstance(metric, dict):
        raise SummaryError(f"{label}: malformed metric")
    if metric.get("status") != "measured":
        return None
    values = metric.get("values")
    if not isinstance(values, list):
        raise SummaryError(f"{label}: measured metric has no values")
    return [verify.u64(value, f"{label}.values[{i}]") for i, value in enumerate(values)]


def raw_sha256(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def profile_receipt_path(role: str, kind: str) -> Path:
    return ROOT / "profiles" / PROFILE_BUILD_DIR[role] / kind / "receipt.json"


def profile_artifact(row: dict[str, Any], name: str) -> tuple[bytes, dict[str, Any]]:
    role = row["role"]
    kind = row["kind"]
    receipt = row["receipt"]
    artifacts = receipt.get("artifacts")
    if not isinstance(artifacts, dict) or name not in artifacts:
        raise SummaryError(f"profile.{role}.{kind}: artifact {name!r} is missing")
    descriptor = artifacts[name]
    receipt_path = profile_receipt_path(role, kind)
    selected = verify.artifact_path(receipt_path, descriptor, f"{receipt_path}.artifacts.{name}")
    return verify.artifact_bytes(selected, f"{receipt_path}.artifacts.{name}"), descriptor


def profile_artifact_binding(name: str, descriptor: dict[str, Any]) -> dict[str, Any]:
    path = descriptor.get("path")
    size = descriptor.get("bytes")
    digest = descriptor.get("sha256")
    if not isinstance(path, str) or not path or Path(path).is_absolute() or not isinstance(size, int) or size < 0 or not isinstance(digest, str):
        raise SummaryError(f"profile artifact {name}: malformed binding")
    return {
        "path": path[:-3] if path.endswith(".gz") else path,
        "bytes": size,
        "sha256": digest,
    }


def parse_stat_count(value: str) -> int | None:
    normalized = value.strip().replace(",", "")
    return int(normalized) if STAT_NUMBER.fullmatch(normalized) else None


def parse_perf_stat(raw: bytes, label: str) -> dict[str, Any]:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        raise SummaryError(f"{label}: perf stat is not UTF-8: {error}") from error
    events: dict[str, dict[str, Any]] = {}
    for line in text.splitlines():
        fields = line.split(",")
        if len(fields) < 3:
            continue
        event = fields[2].strip()
        if event not in PROFILE_EVENTS:
            continue
        if event in events:
            raise SummaryError(f"{label}: duplicate event {event}")
        count = parse_stat_count(fields[0])
        running = fields[4].strip() if len(fields) > 4 else ""
        if count is None:
            events[event] = {
                "status": "missing",
                "reason": fields[0].strip() or "perf did not publish a count",
                "running_percent": running or None,
            }
        else:
            events[event] = {"status": "available", "count": count, "running_percent": running or None}
    for event in PROFILE_EVENTS:
        events.setdefault(event, {"status": "missing", "reason": "event line was absent from perf stat output"})

    def available(event: str) -> int | None:
        item = events[event]
        return item.get("count") if item.get("status") == "available" else None

    cycles = available("cycles:u")
    instructions = available("instructions:u")
    branches = available("branches:u")
    misses = available("branch-misses:u")
    return {
        "scope": "perf_stat_whole_process_including_setup_warmups_timed_calls_hashing_oracle",
        "events": events,
        "ipc": {"status": "available", "value": instructions / cycles} if cycles and instructions is not None else {"status": "missing", "reason": "cycles:u and instructions:u were not both available"},
        "branch_miss_rate": {"status": "available", "percent": misses * 100.0 / branches} if branches and misses is not None else {"status": "missing", "reason": "branches:u and branch-misses:u were not both available"},
        "l1_load_count": {"status": "missing", "reason": "L1 load denominator was not collected"},
        "llc_cache_misses": {"status": "missing", "reason": "LLC event was unavailable/not part of the frozen event set"},
        "raw_sha256": raw_sha256(raw),
    }


def parse_perf_report(raw: bytes, label: str) -> dict[str, Any]:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        raise SummaryError(f"{label}: perf report is not UTF-8: {error}") from error
    lost: int | None = None
    event: str | None = None
    event_count: int | None = None
    rows: list[dict[str, Any]] = []
    warnings = 0
    for line in text.splitlines():
        if ADDR2LINE_TEXT in line:
            warnings += 1
        if match := LOST_RE.match(line):
            lost = int(match.group(1).replace(",", ""))
        if match := EVENT_HEADER_RE.match(line):
            event = match.group(2)
        if match := EVENT_COUNT_RE.match(line):
            event_count = int(match.group(1).replace(",", ""))
        if match := REPORT_ROW_RE.match(line):
            overhead = float(match.group(1))
            symbol = match.group(2).strip()
            if not symbol.startswith("|") and not symbol.startswith("--"):
                rows.append({"overhead_percent": overhead, "symbol": symbol})
    return {
        "scope": "perf_report_no_children_whole_process_samples",
        "event": {"status": "available", "name": event} if event else {"status": "missing", "reason": "perf report event header was absent"},
        "event_count": {"status": "available", "count": event_count} if event_count is not None else {"status": "missing", "reason": "perf report event count was absent"},
        "lost_samples": {"status": "available", "count": lost} if lost is not None else {"status": "missing", "reason": "Total Lost Samples header was absent"},
        "addr2line_warnings": {"status": "available", "count": warnings, "scope": "symbolization_diagnostic_not_lost_sample_count"},
        "top_self_rows": rows[:20],
        "raw_sha256": raw_sha256(raw),
    }


def parse_record_event(raw: bytes, label: str) -> dict[str, Any]:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        raise SummaryError(f"{label}: perf report is not UTF-8: {error}") from error
    event = next((match.group(2) for line in text.splitlines() if (match := EVENT_HEADER_RE.match(line))), None)
    events = {
        "cycles:u": {"status": "available_sampled", "scope": "perf_record_sampling_event"} if event == "cycles:u" else {"status": "missing", "reason": "cycles:u sample event header was absent"},
        "instructions:u": {"status": "missing", "reason": "record lane sampled cycles:u only"},
        "branches:u": {"status": "missing", "reason": "record lane sampled cycles:u only"},
        "branch-misses:u": {"status": "missing", "reason": "record lane sampled cycles:u only"},
        "L1-dcache-load-misses:u": {"status": "missing", "reason": "record lane sampled cycles:u only"},
    }
    return {
        "scope": "perf_record_whole_process_including_setup_warmups_timed_calls_hashing_oracle",
        "events": events,
        "ipc": {"status": "missing", "reason": "instructions and cycles counts are not a paired perf-record counter measurement"},
        "branch_miss_rate": {"status": "missing", "reason": "branch counters were not collected by this record lane"},
        "l1_load_count": {"status": "missing", "reason": "L1 load denominator was not collected"},
        "llc_cache_misses": {"status": "missing", "reason": "LLC event was unavailable/not part of the frozen event set"},
    }


def summarize_formal_profile(row: dict[str, Any]) -> dict[str, Any]:
    role = row["role"]
    kind = row["kind"]
    receipt = row["receipt"]
    artifacts = receipt.get("artifacts")
    if not isinstance(artifacts, dict):
        raise SummaryError(f"profile.{role}.{kind}: artifact inventory is missing")
    result: dict[str, Any] = {
        "role": role,
        "kind": kind,
        "attempt": receipt.get("attempt"),
        "selector": receipt.get("selector"),
        "shape": receipt.get("shape"),
        "scope": receipt.get("scope"),
        "source_custody": {
            "source_manifest": receipt.get("source_manifest"),
            "ambient_source_manifest": receipt.get("ambient_source_manifest"),
            "source_unchanged": receipt.get("source_unchanged"),
        },
        "protocol_sha256": receipt.get("protocol_sha256"),
        "oracle_verifier_sha256": receipt.get("oracle_verifier_sha256"),
        "artifacts": {name: profile_artifact_binding(name, descriptor) for name, descriptor in artifacts.items()},
        "whole_process_rss": {
            **row["resource"],
            "scope": "gnu_time_v_verbose_whole_fresh_process_including_setup_warmups_timed_calls_hashing_oracle",
            "path": row["resource"]["path"][:-3] if row["resource"]["path"].endswith(".gz") else row["resource"]["path"],
        },
        "claims": [],
    }
    if kind == "stat":
        raw, _ = profile_artifact(row, "perf_stat")
        result["pmu"] = parse_perf_stat(raw, f"{role}.stat.perf_stat")
    else:
        raw, _ = profile_artifact(row, "perf_report")
        result["pmu"] = parse_record_event(raw, f"{role}.record.perf_report")
        result["record"] = parse_perf_report(raw, f"{role}.record.perf_report")
        result["record"]["perf_data"] = profile_artifact_binding("perf_data", artifacts["perf_data"])
        result["record"]["perf_script"] = profile_artifact_binding("perf_script", artifacts["perf_script"])
    return result


def formal_profile_evidence(verified: dict[str, Any]) -> dict[str, Any]:
    rows = {(row["role"], row["kind"]): row for row in verified.get("profiles", [])}
    expected = {(role, kind) for role in ROLES for kind in PROFILE_KINDS}
    if set(rows) != expected:
        raise SummaryError(f"formal profile set differs: expected {sorted(expected)}, got {sorted(rows)}")
    profiles = [summarize_formal_profile(rows[(role, kind)]) for role in ROLES for kind in PROFILE_KINDS]
    return {
        "schema_version": 1,
        "change": 436,
        "classification": "Whole-process ODT profiles retained for descriptive PMU counters and caller self rows; no causal hotspot, speedup, cache-miss, or 10x claim.",
        "matrix": {"formal_profiles": len(profiles), "required_profiles": 4, "roles": list(ROLES), "kinds": list(PROFILE_KINDS), "shape": "large", "source_matrix_reports": verified["matrix"]["reports"]},
        "profiles": profiles,
        "preparatory_evidence": {"scope": "retained separately from formal profiles and the formal latency matrix"},
        "limitations": [
            "Every counter, IPC result, miss rate, sample count, and self row is whole-command evidence including setup, corpus generation, warmups, timed calls, output hashing, and the untimed oracle.",
            "L1 and LLC cache claims are unavailable from the frozen event set; a missing or zero field is not evidence of zero cache misses.",
            "perf-record samples are caller evidence, not a paired instructions/cycles counter measurement; IPC and branch miss rate are unavailable in record lanes.",
            "Total Lost Samples and addr2line warnings have separate scopes: lost samples concern perf data; addr2line warnings concern symbolization and are not lost-sample evidence.",
            "Self rows come from perf report --no-children and may contain unresolved or inlined symbols; they do not isolate the timed operation from setup or oracle work.",
        ],
        "claims": [],
    }


def normalized_resource(resource: dict[str, Any], label: str) -> dict[str, Any]:
    """Keep archive-relative paths stable when a .log is later gzipped."""

    path = resource.get("path")
    if not isinstance(path, str) or not path:
        raise SummaryError(f"{label}.path: missing resource path")
    canonical_path = path[:-3] if path.endswith(".gz") else path
    result = dict(resource)
    result["path"] = canonical_path
    # The compressed/uncompressed storage suffix is intentionally omitted from
    # the derived summary.  The outer verifier binds the retained artifact
    # bytes; this keeps --check stable after a portable .log -> .log.gz seal.
    result.pop("stored_path", None)
    result["scope"] = "gnu_time_v_verbose_whole_fresh_process"
    return result


def elapsed_summary(row: dict[str, Any], values: list[int]) -> dict[str, Any]:
    label = f"{row['name']}.elapsed_ns"
    result = stats(values, label)
    report_elapsed = row["verified"]["result"]["elapsed_ns"]
    interval = report_elapsed.get("confidence_interval_95")
    if not isinstance(interval, dict):
        raise SummaryError(f"{label}: report mean confidence interval is missing")
    result["mean_ci"] = {
        "method": interval.get("method"),
        "lower": finite(interval.get("lower"), f"{label}.mean_ci.lower"),
        "upper": finite(interval.get("upper"), f"{label}.mean_ci.upper"),
    }
    if row["lane"]["mode"] == "normal":
        result["p50_ci"] = bootstrap_median(values, label)
        result["bootstrap_mean_ci"] = bootstrap_mean(values, label)
    return result


def allocation_summary(row: dict[str, Any], allocation_metric: dict[str, Any] | None) -> dict[str, Any]:
    label = row["name"]
    if row["lane"]["mode"] == "normal" or allocation_metric is None or allocation_metric.get("status") != "measured":
        return {
            "status": "unavailable",
            "scope": "operation_global_system_allocator",
            "reason": "normal binary or unavailable allocator vectors",
        }
    names = (
        "allocation_calls", "allocated_bytes", "deallocated_bytes",
        "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
        "peak_live_bytes_after", "region_peak_live_bytes",
    )
    vectors: dict[str, list[int]] = {}
    for name in names:
        values = metric_values(allocation_metric.get(name), f"{label}.allocation.{name}")
        if values is None:
            raise SummaryError(f"{label}: allocator vector {name} is unavailable")
        vectors[name] = values
    lengths = {len(values) for values in vectors.values()}
    if len(lengths) != 1:
        raise SummaryError(f"{label}: allocator vectors have different lengths")
    balances = [
        before + allocated - deallocated - after
        for before, allocated, deallocated, after in zip(
            vectors["live_bytes_before"], vectors["allocated_bytes"],
            vectors["deallocated_bytes"], vectors["live_bytes_after"], strict=True
        )
    ]
    if any(balance != 0 for balance in balances):
        raise SummaryError(f"{label}: allocator live-byte balance is nonzero")
    live_deltas = [after - before for before, after in zip(vectors["live_bytes_before"], vectors["live_bytes_after"], strict=True)]
    region_above_entry = [
        peak - entry
        for peak, entry in zip(vectors["region_peak_live_bytes"], vectors["live_bytes_before"], strict=True)
    ]
    region_above_exit = [
        peak - exit_value
        for peak, exit_value in zip(vectors["region_peak_live_bytes"], vectors["live_bytes_after"], strict=True)
    ]
    sample_indices = row["verified"].get("sample_order")
    if not isinstance(sample_indices, list):
        raise SummaryError(f"{label}: verified sample order is missing")
    return {
        "status": "measured",
        "scope": allocation_metric.get("scope"),
        "statistics": {
            "allocation_calls": stats(vectors["allocation_calls"], f"{label}.allocation_calls"),
            "allocated_bytes": stats(vectors["allocated_bytes"], f"{label}.allocated_bytes"),
            "deallocated_bytes": stats(vectors["deallocated_bytes"], f"{label}.deallocated_bytes"),
            "region_peak_above_entry": stats(region_above_entry, f"{label}.region_peak_above_entry", signed=True),
            "region_peak_above_exit": stats(region_above_exit, f"{label}.region_peak_above_exit", signed=True),
            "live_bytes_before": stats(vectors["live_bytes_before"], f"{label}.live_bytes_before"),
            "live_bytes_after": stats(vectors["live_bytes_after"], f"{label}.live_bytes_after"),
            "live_delta": stats(live_deltas, f"{label}.live_delta", signed=True),
            "live_balance": stats(balances, f"{label}.live_balance", signed=True),
            "peak_live_bytes_before": stats(vectors["peak_live_bytes_before"], f"{label}.peak_live_bytes_before"),
            "peak_live_bytes_after": stats(vectors["peak_live_bytes_after"], f"{label}.peak_live_bytes_after"),
        },
        "raw_aligned_vectors": {
            "sample_indices": list(sample_indices),
            "allocation_calls": vectors["allocation_calls"],
            "allocated_bytes": vectors["allocated_bytes"],
            "deallocated_bytes": vectors["deallocated_bytes"],
            "live_bytes_before": vectors["live_bytes_before"],
            "live_bytes_after": vectors["live_bytes_after"],
            "region_peak_live_bytes": vectors["region_peak_live_bytes"],
            "region_peak_above_entry": region_above_entry,
            "live_balance": balances,
            "live_delta": live_deltas,
        },
        "latency_claim": "allocator elapsed values retained for process context only; no allocator latency claim",
    }


def process_summary(row: dict[str, Any], process: dict[str, Any] | None) -> dict[str, Any]:
    if not isinstance(process, dict):
        return {
            "reported_process_peak_rss": {"status": "unavailable", "scope": "process_lifetime_high_water_after_not_operation_peak"},
            "whole_process_rss": normalized_resource(row["resource"], f"{row['name']}.resource"),
        }
    peak = metric_values(process.get("peak_rss_bytes"), f"{row['name']}.process.peak_rss_bytes")
    peak_scope = process.get("peak_rss_bytes", {}).get("scope", "process_lifetime_high_water_after_not_operation_peak")
    return {
        "whole_process_rss": normalized_resource(row["resource"], f"{row['name']}.resource"),
        "reported_process_peak_rss": {
            "scope": peak_scope,
            "statistics": stats(peak, f"{row['name']}.reported_process_peak_rss") if peak is not None else {"status": "unavailable"},
        },
        "scope_note": "GNU time RSS is whole-process; reported peak_rss_bytes is the process lifetime high-water vector, not operation region peak.",
    }


def row_summary(row: dict[str, Any]) -> dict[str, Any]:
    verified = row["verified"]
    lane = row["lane"]
    values = [verify.u64(value, f"{row['name']}.elapsed[{i}]") for i, value in enumerate(verified["elapsed"])]
    identity = row["identity"]
    metrics = verified["metrics"]
    result = verified["result"]
    row_result: dict[str, Any] = {
        "phase": row["phase"],
        "role": row["role"],
        "mode": lane["mode"],
        "shape": lane["shape"],
        "repeat": lane["repeat"],
        "name": row["name"],
        "report": str(row["report_path"]),
        "corpus": {key: identity[key] for key in ("archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256")},
        "output_sha256": identity["output_sha256"],
        "semantic_sha256": identity["semantic_sha256"],
        "source_identity": {
            "build_revision": row["build"]["revision"],
            "ambient_revision": identity["environment_revision"],
            "ambient_worktree_dirty": identity["environment_dirty"],
            "binary_sha256": row["build"]["binaries"][lane["mode"]]["sha256"],
            "source_manifest_sha256": row["build"]["source_manifest"]["sha256"],
            "styles_xml_sha256": identity["styles_xml_sha256"],
            "meta_xml_sha256": identity["meta_xml_sha256"],
        },
        "sink": {key: identity[key] for key in ("sink_accepted_bytes", "sink_write_calls")},
        "elapsed_ns": elapsed_summary(row, values),
        "process_memory": process_summary(row, metrics.get("process")),
        "allocation": allocation_summary(row, metrics.get("allocation")),
        "timing_scope": "30 timed samples inside one fresh process per formal report; corpus/oracle and GNU time RSS scopes are separate",
        "latency_claim_scope": "normal operation elapsed vectors only" if lane["mode"] == "normal" else "allocator elapsed retained for process context only; no allocator latency claim",
    }
    if lane["mode"] == "normal":
        row_result["throughput"] = {
            "paragraphs_per_second": stats(
                [identity["paragraph_count"] * 1_000_000_000 / value for value in values],
                f"{row['name']}.paragraphs_per_second",
            ),
            "archive_bytes_per_second": stats(
                [identity["archive_bytes"] * 1_000_000_000 / value for value in values],
                f"{row['name']}.archive_bytes_per_second",
            )
        }
    # Keep the report's operation metric claim available without copying all
    # vectors into the summary; allocator vectors are retained explicitly
    # above because they are the allocation comparison substrate.
    row_result["operation_metric_alignment"] = result["operation_metrics"]["alignment"]
    return row_result


def relative(old: float, new: float, label: str) -> float:
    if old == 0:
        raise SummaryError(f"{label}: cannot compare a zero baseline")
    return (new - old) / old * 100.0


def repeat_specs(mode: str) -> tuple[tuple[str, Callable[[dict[str, Any]], float], bool], ...]:
    if mode == "normal":
        return (
            ("elapsed_ns.p50", lambda row: float(row["elapsed_ns"]["p50"]), True),
            ("elapsed_ns.p95", lambda row: float(row["elapsed_ns"]["p95"]), True),
            ("elapsed_ns.p99", lambda row: float(row["elapsed_ns"]["p99"]), True),
            ("elapsed_ns.mean", lambda row: float(row["elapsed_ns"]["mean"]), True),
            ("throughput.paragraphs_per_second.p50", lambda row: float(row["throughput"]["paragraphs_per_second"]["p50"]), False),
        )
    return (
        ("allocation.statistics.allocation_calls.p50", lambda row: float(row["allocation"]["statistics"]["allocation_calls"]["p50"]), True),
        ("allocation.statistics.allocated_bytes.p50", lambda row: float(row["allocation"]["statistics"]["allocated_bytes"]["p50"]), True),
        ("allocation.statistics.region_peak_above_entry.p50", lambda row: float(row["allocation"]["statistics"]["region_peak_above_entry"]["p50"]), True),
        ("allocation.statistics.live_bytes_after.p50", lambda row: float(row["allocation"]["statistics"]["live_bytes_after"]["p50"]), True),
    )


def rss_specs() -> tuple[tuple[str, Callable[[dict[str, Any]], float], bool, str], ...]:
    return (
        ("whole_process_rss_bytes", lambda row: float(row["process_memory"]["whole_process_rss"]["bytes"]), True, "gnu_time_v_verbose_whole_fresh_process"),
        ("reported_process_peak_rss.p50", lambda row: float(row["process_memory"]["reported_process_peak_rss"]["statistics"]["p50"]), True, "process_lifetime_high_water_after_not_operation_peak"),
    )


def repeat_flags(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    indexed = {(row["role"], row["mode"], row["shape"], row["repeat"]): row for row in rows}
    flags: list[dict[str, Any]] = []
    for role in ROLES:
        for mode in MODES:
            for shape in SHAPES:
                first = indexed[(role, mode, shape, "R1")]
                second = indexed[(role, mode, shape, "R2")]
                specs = list(repeat_specs(mode))
                for metric, select, higher_is_regression in specs:
                    old, new = select(first), select(second)
                    delta = relative(old, new, f"repeat.{role}.{mode}.{shape}.{metric}")
                    flags.append({"scope": "within_role_repeat_drift", "role": role, "mode": mode, "shape": shape, "metric": metric, "r1": old, "r2": new, "relative_percent": delta, "threshold_percent": THRESHOLD, "flagged": abs(delta) > THRESHOLD, "higher_is_regression": higher_is_regression})
                for metric, select, higher_is_regression, scope in rss_specs():
                    old, new = select(first), select(second)
                    delta = relative(old, new, f"repeat.{role}.{mode}.{shape}.{metric}")
                    flags.append({"scope": scope, "role": role, "mode": mode, "shape": shape, "metric": metric, "r1": old, "r2": new, "relative_percent": delta, "threshold_percent": THRESHOLD, "flagged": abs(delta) > THRESHOLD, "higher_is_regression": higher_is_regression})
    return flags


def comparison_specs(mode: str) -> tuple[tuple[str, Callable[[dict[str, Any]], float], bool, str], ...]:
    if mode == "normal":
        return (
            ("latency_ns.p50", lambda row: float(row["elapsed_ns"]["p50"]), True, "timed_operation_samples"),
            ("latency_ns.p95", lambda row: float(row["elapsed_ns"]["p95"]), True, "timed_operation_samples"),
            ("latency_ns.p99", lambda row: float(row["elapsed_ns"]["p99"]), True, "timed_operation_samples"),
            ("latency_ns.mean", lambda row: float(row["elapsed_ns"]["mean"]), True, "timed_operation_samples"),
            ("throughput.paragraphs_per_second.p50", lambda row: float(row["throughput"]["paragraphs_per_second"]["p50"]), False, "timed_operation_samples"),
        )
    return (
        ("allocation_calls.p50", lambda row: float(row["allocation"]["statistics"]["allocation_calls"]["p50"]), True, "allocator_operation_vectors"),
        ("allocated_bytes.p50", lambda row: float(row["allocation"]["statistics"]["allocated_bytes"]["p50"]), True, "allocator_operation_vectors"),
        ("region_peak_above_entry.p50", lambda row: float(row["allocation"]["statistics"]["region_peak_above_entry"]["p50"]), True, "allocator_operation_vectors"),
        ("live_bytes_after.p50", lambda row: float(row["allocation"]["statistics"]["live_bytes_after"]["p50"]), True, "allocator_operation_vectors"),
    )


def matched_comparisons(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    indexed = {(row["role"], row["mode"], row["shape"], row["repeat"]): row for row in rows}
    pairs = (
        ("streaming-control", "before-streaming", "after-streaming"),
    )
    comparisons: list[dict[str, Any]] = []
    for pair_name, baseline_role, candidate_role in pairs:
        for mode in MODES:
            for shape in SHAPES:
                for repeat in REPEATS:
                    baseline = indexed[(baseline_role, mode, shape, repeat)]
                    candidate = indexed[(candidate_role, mode, shape, repeat)]
                    specs = list(comparison_specs(mode))
                    for metric, select, higher_is_regression, scope in specs:
                        old, new = select(baseline), select(candidate)
                        delta = relative(old, new, f"comparison.{pair_name}.{mode}.{shape}.{repeat}.{metric}")
                        regression = delta > THRESHOLD if higher_is_regression else delta < -THRESHOLD
                        comparisons.append({"comparison": pair_name, "baseline_role": baseline_role, "candidate_role": candidate_role, "mode": mode, "shape": shape, "repeat": repeat, "metric": metric, "before": old, "after": new, "relative_percent": delta, "threshold_percent": THRESHOLD, "higher_is_regression": higher_is_regression, "scope": scope, "regression": regression})
                    for metric, select, higher_is_regression, scope in rss_specs():
                        old, new = select(baseline), select(candidate)
                        delta = relative(old, new, f"comparison.{pair_name}.{mode}.{shape}.{repeat}.{metric}")
                        comparisons.append({"comparison": pair_name, "baseline_role": baseline_role, "candidate_role": candidate_role, "mode": mode, "shape": shape, "repeat": repeat, "metric": metric, "before": old, "after": new, "relative_percent": delta, "threshold_percent": THRESHOLD, "higher_is_regression": higher_is_regression, "scope": scope, "regression": delta > THRESHOLD})
    return comparisons


def strict_load(path: Path) -> Any:
    def duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise SummaryError(f"{path}: duplicate JSON key {key!r}")
            result[key] = value
        return result

    def reject_constant(value: str) -> Any:
        raise SummaryError(f"{path}: non-finite JSON constant {value!r}")

    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=duplicate_pairs, parse_constant=reject_constant)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise SummaryError(f"{path}: invalid JSON: {error}") from error


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode("utf-8")


def derive(*, require_binaries: bool = False) -> dict[str, Any]:
    verified = verify.verify_matrix(require_binaries=require_binaries)
    rows = [row_summary(row) for row in verified["rows"]]
    comparisons = matched_comparisons(rows)
    repeats = repeat_flags(rows)
    formal_profiles = [dict(row, resource=normalized_resource(row["resource"], f"profile.{row['role']}.{row['kind']}.resource")) for row in verified.get("profiles", [])]
    return {
        "schema_version": 1,
        "change": 436,
        "classification": "Matched current-revision ODT streaming paragraph creation evidence; descriptive role comparisons only, with no speedup, causal, cache-miss, or 10x claim.",
        "timing_scope": "The operation timer covers fresh paragraph construction, the selected public ODT authoring/publication path, and HashingDiscardSink writes. Corpus setup, package reopen/semantic gates, sink finalization/digest extraction, and GNU time/procfs observations are separate scopes.",
        "comparison_scope": "before-streaming to after-streaming is a same-API control comparison. Normal rows include elapsed and paragraph throughput. Allocator rows compare aligned allocation vectors and process memory only; allocator elapsed values are descriptive context.",
        "claims": [],
        "threshold_policy": {"regression_percent": THRESHOLD, "flagged_metrics": "Every latency, throughput, RSS, and allocator delta beyond the threshold is retained for review; flags do not establish causality or a blanket pass."},
        "protocol_sha256": verified["protocol_sha256"],
        "oracle": verified["oracle"],
        "builds": verified["builds"],
        "ambient": verified["ambient"],
        "matrix": {**verified["matrix"], "formal_reports": len(rows), "retained_samples": sum(row["elapsed_ns"]["count"] for row in rows)},
        "cross_phase_identity": verified["cross_phase_identity"],
        "rows": rows,
        "repeat_flags": repeats,
        "comparisons": comparisons,
        "profiles": {
            "formal": {"count": len(formal_profiles), "required_count": 4, "rows": formal_profiles, "scope": "whole-command normal large-shape perf stat/record; setup, warmups, hashing, and oracle are included"},
            "raw_pmu_and_self_samples": formal_profile_evidence(verified),
            "preparatory": {"scope": "pre-production pilot/profile evidence retained separately; not part of the 24-report formal matrix"},
        },
        "extra_attempts": verified.get("extra_attempts", []),
        "uncertainty": {
            "method": "report Student-t mean interval plus deterministic bootstrap percentile intervals",
            "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
            "bootstrap_seed": BOOTSTRAP_SEED,
            "level": 0.95,
            "sample_scope": "30 timed samples within one fresh process per formal report; not 30 independent process replicates",
        },
        "limitations": [
            "Normal p50/p95/p99 and mean are descriptive within-report statistics; bootstrap intervals resample those within-process samples.",
            "Allocator elapsed vectors are retained for process context only and are excluded from latency comparisons.",
            "region_peak_above_entry is an operation allocator metric, not RSS and not total process memory; live-byte balance is checked independently.",
            "GNU time maximum RSS and reported process peak high-water vectors are process-wide observations and are reviewed separately from operation region peaks.",
            "Whole-command profiles include setup, corpus generation, warmups, timed calls, output hashing, and untimed validation; they do not establish a causal hotspot or optimization speedup.",
            "The outer verifier requires exact archive, content, style, metadata, semantic, topology, output, and sink identity across the before/after streaming roles; this is an evidence gate, not a performance claim.",
            "Preparatory pilots/profiles remain separate from this formal 24-report summary.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "summary.json")
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--require-binaries", action="store_true")
    args = parser.parse_args()
    try:
        derived = derive(require_binaries=args.require_binaries)
        if args.check:
            if not args.output.is_file() or canonical(strict_load(args.output)) != canonical(derived):
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
