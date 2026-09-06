#!/usr/bin/env python3
"""Derive the immutable 0445 OPC Part-addition baseline summary.

The driver reads only retained receipts, reports, GNU-time logs, and the two
profile outputs.  It does not invoke the workload, Git, an oracle, or a
profiler.  The provider oracle remains responsible for semantic correctness;
this module binds the report fields needed for a descriptive baseline and
fails closed when those fields are absent or inconsistent.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import io
import json
import math
from pathlib import Path
import random
import re
import statistics
import sys
from typing import Any


CHANGE = 445
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
PHASES = {
    "A1": (("normal", "tiny"), ("normal", "medium"), ("normal", "large"), ("allocator", "tiny"), ("allocator", "medium"), ("allocator", "large")),
    "B1": (("normal", "tiny"), ("normal", "medium"), ("normal", "large"), ("allocator", "tiny"), ("allocator", "medium"), ("allocator", "large")),
    "B2": (("allocator", "large"), ("allocator", "medium"), ("allocator", "tiny"), ("normal", "large"), ("normal", "medium"), ("normal", "tiny")),
    "A2": (("allocator", "large"), ("allocator", "medium"), ("allocator", "tiny"), ("normal", "large"), ("normal", "medium"), ("normal", "tiny")),
}
SOURCE_MODES = {"A1":"observed", "B1":"plain", "B2":"plain", "A2":"observed"}
SELECTORS = {"observed":"opc_part_add_lifecycle", "plain":"opc_part_add_plain_lifecycle"}

SAMPLES = 30
WARMUPS = 3
BOOTSTRAP_RESAMPLES = 2_000
BOOTSTRAP_SEED = 4_450_301
REPEAT_THRESHOLD_PERCENT = 5.0
MAX_ARTIFACT_BYTES = 512 * 1024 * 1024
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
RSS_RE = re.compile(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$")
PERF_COUNT_RE = re.compile(r"^\s*([0-9][0-9,]*)\s*$")


class DeriveError(ValueError):
    pass


def fail(label: str, message: str) -> None:
    raise DeriveError(f"{label}: {message}")


def duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            raise DeriveError(f"duplicate JSON key {key!r}")
        value[key] = item
    return value


def reject_constant(value: str) -> Any:
    raise DeriveError(f"non-finite JSON value {value!r}")


def parse_json(raw: bytes, label: str) -> Any:
    try:
        return json.loads(raw.decode("utf-8"), object_pairs_hook=duplicate_keys, parse_constant=reject_constant)
    except (UnicodeError, json.JSONDecodeError, DeriveError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def load_json(path: Path, label: str) -> Any:
    try:
        raw = path.read_bytes()
    except OSError as error:
        fail(label, f"cannot read: {error}")
    if len(raw) > MAX_ARTIFACT_BYTES:
        fail(label, "JSON exceeds bounded artifact size")
    return parse_json(raw, label)


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(label, "expected object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(label, "expected non-empty string")
    return value


def u64(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0 or value > (1 << 64) - 1:
        fail(label, "expected u64")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label)
    if HEX64.fullmatch(value) is None:
        fail(label, "expected SHA-256")
    return value.lower()


def binding_identity(value: Any, label: str, *, binary: bool) -> dict[str, Any]:
    row = obj(value, label)
    size = row.get("bytes")
    if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
        fail(label, "expected positive byte count")
    return {"bytes": size, "sha256": digest(row.get("sha256"), f"{label}.sha256"), **({} if not binary else {"binary": True})}


def source_binding(value: Any, label: str) -> dict[str, Any]:
    row = obj(value, label)
    files = row.get("files")
    if isinstance(files, bool) or not isinstance(files, int) or files <= 0:
        fail(label, "expected positive source file count")
    return {"files": files, "sha256": digest(row.get("sha256"), f"{label}.sha256")}


def timestamp(value: Any, label: str) -> dt.datetime:
    value = text(value, label)
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        fail(label, f"invalid timestamp: {error}")
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        fail(label, "timestamp must include timezone")
    return parsed


def sha_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def sha(path: Path) -> str:
    return sha_bytes(path.read_bytes())


def safe_path(root: Path, value: Any, label: str) -> Path:
    name = text(value, label)
    path = Path(name)
    if path.is_absolute() or ".." in path.parts:
        fail(label, "path must be bundle-relative")
    resolved = (root / path).resolve()
    if not resolved.is_relative_to(root.resolve()):
        fail(label, "path escapes evidence root")
    return resolved


def bounded_gzip(raw: bytes, label: str) -> bytes:
    try:
        with gzip.GzipFile(fileobj=io.BytesIO(raw)) as stream:
            chunks: list[bytes] = []
            total = 0
            while True:
                chunk = stream.read(min(1024 * 1024, MAX_ARTIFACT_BYTES - total + 1))
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_ARTIFACT_BYTES:
                    fail(label, "decompressed artifact exceeds bounded size")
                chunks.append(chunk)
            return b"".join(chunks)
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        fail(label, f"invalid gzip artifact: {error}")
    raise AssertionError("unreachable")


def artifact(root: Path, record: Any, label: str) -> tuple[bytes, str]:
    row = obj(record, label)
    raw_path = safe_path(root, row.get("path"), f"{label}.path")
    selected = raw_path
    if not selected.is_file() and Path(str(raw_path) + ".gz").is_file():
        selected = Path(str(raw_path) + ".gz")
    if not selected.is_file():
        fail(label, "retained raw or gzip artifact is missing")
    try:
        stored = selected.read_bytes()
    except OSError as error:
        fail(label, f"cannot read artifact: {error}")
    if len(stored) > MAX_ARTIFACT_BYTES:
        fail(label, "stored artifact exceeds bounded size")
    content = bounded_gzip(stored, label) if selected.suffix == ".gz" else stored
    if len(content) > MAX_ARTIFACT_BYTES:
        fail(label, "artifact exceeds bounded size")
    if row.get("bytes") != len(content) or row.get("sha256") != sha_bytes(content):
        fail(label, "artifact hash or logical size differs from receipt")
    # Summary identities name the receipt's logical artifact. Storage may be
    # raw or deterministically gzipped without changing the derived result.
    return content, str(raw_path.relative_to(root))


def percentile(ordered: list[float], fraction: float) -> float:
    if not ordered:
        raise DeriveError("cannot calculate percentile of an empty vector")
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def median(ordered: list[float]) -> float:
    count = len(ordered)
    return (ordered[(count - 1) // 2] + ordered[count // 2]) / 2.0


def finite_number(value: Any, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        fail(label, "expected finite number")
    return float(value)


def statistic(values: list[float], which: str) -> float:
    ordered = sorted(values)
    if which == "p50":
        return median(ordered)
    if which == "p95":
        return percentile(ordered, 0.95)
    if which == "p99":
        return percentile(ordered, 0.99)
    if which == "mean":
        return statistics.fmean(values)
    raise DeriveError(f"unknown statistic {which}")


def clean_number(value: float) -> int | float:
    return int(value) if value.is_integer() else value


def bootstrap(values: list[float], which: str, seed: int) -> dict[str, Any]:
    rng = random.Random(seed)
    n = len(values)
    draws: list[float] = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        sample = [values[rng.randrange(n)] for _ in range(n)]
        draws.append(statistic(sample, which))
    draws.sort()
    return {
        "lower": clean_number(percentile(draws, 0.025)),
        "upper": clean_number(percentile(draws, 0.975)),
        "method": "deterministic nonparametric bootstrap percentile interval",
        "resamples": BOOTSTRAP_RESAMPLES,
        "seed": seed,
    }


def stats(values: list[int | float], label: str, *, bootstrap_seed: int | None = None) -> dict[str, Any]:
    if not values:
        fail(label, "empty vector")
    clean = [finite_number(value, f"{label}[{index}]") for index, value in enumerate(values)]
    ordered = sorted(clean)
    result: dict[str, Any] = {
        "count": len(clean),
        "min": clean_number(ordered[0]),
        "p50": clean_number(median(ordered)),
        "p95": clean_number(percentile(ordered, 0.95)),
        "p99": clean_number(percentile(ordered, 0.99)),
        "max": clean_number(ordered[-1]),
        "mean": clean_number(statistics.fmean(clean)),
        "standard_deviation": clean_number(statistics.stdev(clean)) if len(clean) > 1 else 0,
    }
    if bootstrap_seed is not None:
        result["bootstrap_95"] = {
            name: bootstrap(clean, name, bootstrap_seed + offset)
            for offset, name in enumerate(("p50", "p95", "p99", "mean"))
        }
    return result


def report_result(report: dict[str, Any], label: str) -> dict[str, Any]:
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1:
        fail(label, "expected exactly one result")
    return obj(results[0], f"{label}.results[0]")


def required_vector(value: Any, label: str) -> list[int]:
    metric = obj(value, label)
    values = metric.get("values")
    if not isinstance(values, list):
        fail(label, "measured vector is missing")
    return [u64(item, f"{label}.values[{index}]") for index, item in enumerate(values)]


def parse_rss(raw: bytes, label: str) -> int:
    matches = []
    for line in raw.decode("utf-8", errors="replace").splitlines():
        found = RSS_RE.match(line)
        if found:
            matches.append(int(found.group(1)) * 1024)
    if len(matches) != 1:
        fail(label, "expected one GNU-time maximum RSS line")
    return matches[0]


def lane_name(phase: str, mode: str, shape: str) -> str:
    return f"{phase}-{mode}-{shape}-r{phase[-1]}"


def load_lane(root: Path, protocol: dict[str, Any], phase: str, mode: str, shape: str, attempt: str) -> dict[str, Any]:
    name = lane_name(phase, mode, shape)
    path = root / "runs" / phase / attempt / f"{name}-receipt.json"
    receipt = obj(load_json(path, str(path)), str(path))
    if receipt.get("schema") != "litchi-0445-capture-receipt-v1" or receipt.get("change") != CHANGE:
        fail(str(path), "receipt schema/change differs")
    if receipt.get("status") != "pass" or receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
        fail(str(path), "capture and oracle must pass")
    if receipt.get("phase") != phase or receipt.get("role") != SOURCE_MODES[phase] or receipt.get("name") != name:
        fail(str(path), "receipt lane identity differs")
    if receipt.get("selector") != SELECTORS[SOURCE_MODES[phase]] or receipt.get("source_field") != "opc_part_add":
        fail(str(path), "selector/source field differs")
    if receipt.get("source_unchanged") is not True or receipt.get("outside_bundle_status_unchanged") is not True:
        fail(str(path), "source/repository custody is not unchanged")
    protocol_sha = digest(receipt.get("protocol_sha256"), f"{path}.protocol_sha256")
    if protocol_sha != sha(root / "protocol.json"):
        fail(str(path), "protocol binding differs")
    started = timestamp(receipt.get("started_utc"), f"{path}.started_utc")
    finished = timestamp(receipt.get("finished_utc"), f"{path}.finished_utc")
    if finished < started:
        fail(str(path), "capture interval is reversed")
    artifacts = obj(receipt.get("artifacts"), f"{path}.artifacts")
    required = {"report", "catalog", "resource_log", "workload_log", "oracle_log"}
    if set(artifacts) != required:
        fail(str(path), f"artifact set differs: {set(artifacts)!r}")
    report_raw, report_path = artifact(root, artifacts["report"], f"{path}.report")
    catalog_raw, catalog_path = artifact(root, artifacts["catalog"], f"{path}.catalog")
    resource_raw, resource_path = artifact(root, artifacts["resource_log"], f"{path}.resource_log")
    _, workload_path = artifact(root, artifacts["workload_log"], f"{path}.workload_log")
    oracle_raw, oracle_path = artifact(root, artifacts["oracle_log"], f"{path}.oracle_log")
    if b"VALID" not in oracle_raw:
        fail(str(path), "oracle log lacks VALID")
    report = obj(parse_json(report_raw, report_path), report_path)
    result = report_result(report, report_path)
    if result.get("case") != SELECTORS[SOURCE_MODES[phase]]:
        fail(report_path, "report selector differs")
    corpus = obj(result.get("corpus"), f"{report_path}.results[0].corpus")
    if corpus.get("shape") != shape:
        fail(report_path, "corpus shape differs")
    elapsed = obj(result.get("elapsed_ns"), f"{report_path}.elapsed_ns")
    elapsed_values = [u64(value, f"{report_path}.elapsed_ns.samples[{index}]") for index, value in enumerate(elapsed.get("samples", []))]
    if len(elapsed_values) != SAMPLES:
        fail(report_path, f"expected {SAMPLES} elapsed samples")
    sample_order = elapsed.get("sample_order")
    if not isinstance(sample_order, list) or len(sample_order) != SAMPLES or sorted(sample_order) != list(range(SAMPLES)):
        fail(report_path, "elapsed sample_order is not a permutation")
    operation = obj(result.get("operation_metrics"), f"{report_path}.operation_metrics")
    source = obj(obj(result.get("source"), f"{report_path}.source").get("opc_part_add"), f"{report_path}.source.opc_part_add")
    source_archive_sha256 = digest(source.get("source_archive_sha256"), f"{report_path}.source.opc_part_add.source_archive_sha256")
    output_archive_sha256 = digest(source.get("output_archive_sha256"), f"{report_path}.source.opc_part_add.output_archive_sha256")
    identity = {
        "archive_bytes": u64(corpus.get("archive_bytes"), f"{report_path}.corpus.archive_bytes"),
        "archive_sha256": digest(corpus.get("archive_sha256"), f"{report_path}.corpus.archive_sha256"),
        "target_payload_bytes": u64(corpus.get("target_payload_bytes"), f"{report_path}.corpus.target_payload_bytes"),
        "target_payload_sha256": digest(corpus.get("target_payload_sha256"), f"{report_path}.corpus.target_payload_sha256"),
        "source_archive_sha256": source_archive_sha256,
        "output_archive_sha256": output_archive_sha256,
        "output_sha256": digest(result.get("output_sha256"), f"{report_path}.output_sha256"),
        "added_payload_sha256": digest(source.get("added_payload_sha256"), f"{report_path}.source.opc_part_add.added_payload_sha256"),
    }
    allocation = None
    if mode == "allocator":
        allocation_obj = obj(operation.get("allocation"), f"{report_path}.operation_metrics.allocation")
        if allocation_obj.get("status") != "measured":
            fail(report_path, "allocator report is not measured")
        names = ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes")
        vectors = {name: required_vector(allocation_obj.get(name), f"{report_path}.allocation.{name}") for name in names}
        if any(len(values) != SAMPLES for values in vectors.values()):
            fail(report_path, "allocator vector length differs from elapsed samples")
        operation_indices = allocation_obj.get("sample_indices", operation.get("sample_indices", sample_order))
        if not isinstance(operation_indices, list) or len(operation_indices) != SAMPLES or sorted(operation_indices) != list(range(SAMPLES)):
            fail(report_path, "allocator sample indices are not a permutation")
        if operation_indices != sample_order:
            fail(report_path, "allocator sample indices do not align with elapsed sample order")
        chronological = {
            name: [values[operation_indices.index(index)] for index in range(SAMPLES)]
            for name, values in vectors.items()
        }
        failed = chronological["failed_allocation_calls"]
        if any(value != 0 for value in failed):
            fail(report_path, "allocator failure calls are nonzero")
        balance = [
            before + allocated - deallocated - after
            for before, allocated, deallocated, after in zip(
                chronological["live_bytes_before"], chronological["allocated_bytes"],
                chronological["deallocated_bytes"], chronological["live_bytes_after"]
            )
        ]
        if any(value != 0 for value in balance):
            fail(report_path, "allocator live-byte balance is nonzero")
        if any(
            peak < before or peak < after
            for peak, before, after in zip(
                chronological["region_peak_live_bytes"], chronological["live_bytes_before"], chronological["live_bytes_after"]
            )
        ):
            fail(report_path, "allocator region peak is below an endpoint")
        allocation = {
            "scope": text(allocation_obj.get("scope"), f"{report_path}.allocation.scope"),
            "sample_order": list(operation_indices),
            "elapsed_order_vectors": vectors,
            "chronological_vectors": chronological,
            "statistics": {name: stats(values, f"{report_path}.allocation.{name}") for name, values in vectors.items()},
            "derived": {
                "live_delta": [after - before for before, after in zip(chronological["live_bytes_before"], chronological["live_bytes_after"])],
                "live_balance": balance,
                "region_peak_above_entry": [peak - before for peak, before in zip(chronological["region_peak_live_bytes"], chronological["live_bytes_before"])],
                "region_peak_above_exit": [peak - after for peak, after in zip(chronological["region_peak_live_bytes"], chronological["live_bytes_after"])],
            },
        }
        allocation["derived_statistics"] = {name: stats(values, f"{report_path}.allocation.{name}", bootstrap_seed=None) for name, values in allocation["derived"].items()}
    environment = obj(report.get("environment"), f"{report_path}.environment")
    configuration = obj(report.get("configuration"), f"{report_path}.configuration")
    if configuration.get("samples_per_case") != SAMPLES or configuration.get("warmup_iterations_per_case") != WARMUPS:
        fail(report_path, "report sample/warmup configuration differs or is missing")
    cases = configuration.get("cases")
    if not isinstance(cases, list) or SELECTORS[SOURCE_MODES[phase]] not in cases:
        fail(report_path, "report configuration cases does not bind the append selector")
    corpus_shapes = configuration.get("semantic_shapes")
    if not isinstance(corpus_shapes, list) or shape not in corpus_shapes:
        fail(report_path, "report configuration semantic_shapes does not bind this shape")
    return {
        "phase": phase,
        "mode": mode,
        "shape": shape,
        "repeat": "R" + phase[-1],
        "source_mode": SOURCE_MODES[phase],
        "name": name,
        "receipt": str(path.relative_to(root)),
        "report": report_path,
        "catalog": catalog_path,
        "resource": resource_path,
        "workload_log": workload_path,
        "oracle_log": oracle_path,
        "started": started.isoformat(),
        "finished": finished.isoformat(),
        "elapsed_ns": stats(elapsed_values, f"{report_path}.elapsed_ns", bootstrap_seed=BOOTSTRAP_SEED + len(mode) + len(shape)),
        "elapsed_samples": elapsed_values,
        "sample_order": list(sample_order),
        "process_memory": {"gnu_time_peak_rss_bytes": parse_rss(resource_raw, f"{path}.resource_log"), "scope": "GNU time whole executable invocation; copied oracle is afterward"},
        "allocation": allocation if allocation is not None else {"status": "unavailable", "scope": "normal binary does not expose allocator vectors"},
        "identity": identity,
        "environment": {key: environment.get(key) for key in ("git_revision", "git_worktree_dirty", "rustc_version", "cpu_affinity")},
        "binding": {"revision": receipt.get("revision"), "source_manifest": source_binding(receipt.get("source_manifest"), f"{path}.source_manifest"), "binary": binding_identity(receipt.get("binary"), f"{path}.binary", binary=True), "protocol_sha256": protocol_sha},
    }


def relative_change(old: float, new: float, label: str) -> float | None:
    if old == 0:
        return None
    return (new - old) / old * 100.0


def repeat_flags_one(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_key = {(row["mode"], row["shape"], row["repeat"]): row for row in rows}
    result: list[dict[str, Any]] = []
    for mode in MODES:
        for shape in SHAPES:
            first = by_key[(mode, shape, "R1")]
            second = by_key[(mode, shape, "R2")]
            metrics: list[tuple[str, float | None, float | None]] = [
                ("elapsed_ns.p50", first["elapsed_ns"]["p50"], second["elapsed_ns"]["p50"]),
                ("elapsed_ns.p95", first["elapsed_ns"]["p95"], second["elapsed_ns"]["p95"]),
                ("elapsed_ns.p99", first["elapsed_ns"]["p99"], second["elapsed_ns"]["p99"]),
                ("elapsed_ns.mean", first["elapsed_ns"]["mean"], second["elapsed_ns"]["mean"]),
                ("process_memory.gnu_time_peak_rss_bytes", first["process_memory"]["gnu_time_peak_rss_bytes"], second["process_memory"]["gnu_time_peak_rss_bytes"]),
            ]
            if mode == "allocator":
                for name in ("allocation_calls", "allocated_bytes", "deallocated_bytes", "region_peak_above_entry", "live_bytes_after"):
                    metrics.append((f"allocation.{name}.p50", first["allocation"]["statistics"].get(name, {}).get("p50") if name in first["allocation"]["statistics"] else first["allocation"]["derived_statistics"][name]["p50"], second["allocation"]["statistics"].get(name, {}).get("p50") if name in second["allocation"]["statistics"] else second["allocation"]["derived_statistics"][name]["p50"]))
            for metric, old, new in metrics:
                if old is None or new is None:
                    continue
                change = relative_change(float(old), float(new), metric)
                result.append({"scope": "within_role_repeat_drift", "mode": mode, "shape": shape, "metric": metric, "r1": old, "r2": new, "relative_percent": change, "threshold_percent": REPEAT_THRESHOLD_PERCENT, "flagged": change is not None and abs(change) > REPEAT_THRESHOLD_PERCENT, "interpretation": "descriptive repeat review; no before/after or causal claim"})
    return result


def repeat_flags(rows):
    result = []
    for source_mode in ('observed', 'plain'):
        subset = repeat_flags_one([r for r in rows if r['source_mode'] == source_mode])
        result.extend(dict(row, source_mode=source_mode) for row in subset)
    return result


def comparisons(rows):
    index = {(r['source_mode'],r['mode'],r['shape'],r['repeat']):r for r in rows}
    result=[]
    for mode in MODES:
        for shape in SHAPES:
            for repeat in ('R1','R2'):
                observed=index['observed',mode,shape,repeat];plain=index['plain',mode,shape,repeat]
                pairs=[('elapsed_ns.'+key,observed['elapsed_ns'][key],plain['elapsed_ns'][key]) for key in ('p50','p95','p99','mean')]
                pairs.append(('whole_process_peak_rss_bytes',observed['process_memory']['gnu_time_peak_rss_bytes'],plain['process_memory']['gnu_time_peak_rss_bytes']))
                if mode=='allocator':
                    for key in ('allocation_calls','allocated_bytes','live_bytes_before','live_bytes_after','region_peak_live_bytes'):
                        pairs.append(('allocation.'+key,observed['allocation']['statistics'][key]['p50'],plain['allocation']['statistics'][key]['p50']))
                    pairs.append(('allocation.region_peak_above_entry',observed['allocation']['derived_statistics']['region_peak_above_entry']['p50'],plain['allocation']['derived_statistics']['region_peak_above_entry']['p50']))
                for metric,old,new in pairs:
                    result.append({'mode':mode,'shape':shape,'repeat':repeat,'metric':metric,'observed':old,'plain':new,'relative_percent':relative_change(old,new,metric),'scope':'matched observer/plain diagnostic; not production before/after'})
    return result


def parse_perf_stat(raw: bytes, label: str) -> dict[str, Any]:
    events = ("cycles:u", "instructions:u", "branches:u", "branch-misses:u", "L1-dcache-load-misses:u")
    rows: dict[str, Any] = {}
    for line in raw.decode("utf-8", errors="replace").splitlines():
        parts = line.split(",")
        for event in events:
            if event not in parts:
                continue
            if event in rows:
                fail(label, "duplicate perf event")
            count = parts[0].strip()
            match = PERF_COUNT_RE.fullmatch(count)
            rows[event] = {"raw": line, "available": bool(match), "count": int(count.replace(",", "")) if match else None, "event": event}
    if set(rows) != set(events):
        fail(label, "perf stat must contain every requested event exactly once")
    return {"events": rows, "scope": "whole profiled executable including setup, corpus generation, warmups, samples, hashing, and checks", "cache_zero_policy": "zero or unavailable L1 values are retained but not treated as validated cache behavior"}


def parse_perf_report(raw: bytes, label: str) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    pattern = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)%\s+(.*)$")
    for line in raw.decode("utf-8", errors="replace").splitlines():
        match = pattern.match(line)
        if match:
            result.append({"overhead_percent": float(match.group(1)), "line": match.group(0).rstrip()})
    return result[:20]


def profile_summary(root: Path, kind: str, protocol_sha: str, source_mode: str) -> dict[str, Any]:
    candidates = sorted((root / "profiles" / source_mode / kind).glob("**/receipt.json"))
    failed_attempts = [str(path.relative_to(root)) for path in candidates if load_json(path, str(path)).get("status") == "failed"]
    candidates = [path for path in candidates if load_json(path, str(path)).get("status") == "pass"]
    if len(candidates) != 1:
        fail(f"profiles.after.{kind}", f"expected one passing receipt, found {len(candidates)}; retained failures: {failed_attempts}")
    path = candidates[0]
    receipt = obj(load_json(path, str(path)), str(path))
    if receipt.get("schema") != "litchi-0445-profile-receipt-v1" or receipt.get("status") != "pass" or receipt.get("kind") != kind or receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
        fail(str(path), "profile receipt is not passing")
    if receipt.get("protocol_sha256") != protocol_sha:
        fail(str(path), "profile protocol binding differs")
    artifacts = obj(receipt.get("artifacts"), f"{path}.artifacts")
    if kind == "stat":
        raw, output = artifact(root, artifacts.get("perf_stat"), f"{path}.perf_stat")
        return {"source_mode": source_mode, "kind": kind, "receipt": str(path.relative_to(root)), "binary": binding_identity(receipt.get("binary"), f"{path}.binary", binary=True), "source_manifest": source_binding(receipt.get("source_manifest"), f"{path}.source_manifest"), "counters": parse_perf_stat(raw, output), "artifact": output, "scope": receipt.get("scope")}
    raw, report_path = artifact(root, artifacts.get("perf_report"), f"{path}.perf_report")
    script_raw, script_path = artifact(root, artifacts.get("perf_script"), f"{path}.perf_script")
    data_raw, data_path = artifact(root, artifacts.get("perf_data"), f"{path}.perf_data")
    return {"source_mode": source_mode, "kind": kind, "receipt": str(path.relative_to(root)), "binary": binding_identity(receipt.get("binary"), f"{path}.binary", binary=True), "source_manifest": source_binding(receipt.get("source_manifest"), f"{path}.source_manifest"), "top_self": parse_perf_report(raw, report_path), "perf_data_bytes": len(data_raw), "perf_script_bytes": len(script_raw), "artifacts": {"report": report_path, "script": script_path, "data": data_path}, "scope": receipt.get("scope")}


def derive(root: Path, attempt: str) -> dict[str, Any]:
    protocol_path = root / "protocol.json"
    protocol = obj(load_json(protocol_path, "protocol.json"), "protocol.json")
    if protocol.get("change") != CHANGE or protocol.get("selector") not in {None, "opc_part_add_lifecycle"}:
        fail("protocol", "change or selector differs")
    if protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS or protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol", "samples/warmups/cpu/worker binding differs")
    if protocol.get("shapes") != {"tiny": 64, "medium": 1024, "large": 4096}:
        fail("protocol.shapes", "shape binding differs")
    protocol_sha = sha(protocol_path)
    rows: list[dict[str, Any]] = []
    for phase, lanes in PHASES.items():
        for mode, shape in lanes:
            rows.append(load_lane(root, protocol, phase, mode, shape, attempt))
    if len(rows) != 24:
        fail("matrix", "expected 24 rows")
    for phase in PHASES:
        phase_rows = [row for row in rows if row["phase"] == phase]
        for previous, current in zip(phase_rows, phase_rows[1:]):
            if timestamp(previous["finished"], "previous.finished") > timestamp(current["started"], "current.started"):
                fail(f"phase.{phase}", "lanes overlap")
    identity_fields = ("archive_bytes", "archive_sha256", "target_payload_bytes", "target_payload_sha256", "source_archive_sha256", "output_archive_sha256", "output_sha256", "added_payload_sha256")
    for shape in SHAPES:
        reference = next(row["identity"] for row in rows if row["mode"] == "normal" and row["shape"] == shape and row["repeat"] == "R1")
        for row in rows:
            if row["shape"] == shape and any(row["identity"][field] != reference[field] for field in identity_fields):
                fail(f"identity.{shape}", "corpus/output identity differs across mode or repeat")
    source_bindings = {(row["binding"]["revision"], json.dumps(row["binding"]["source_manifest"], sort_keys=True)) for row in rows}
    if len(source_bindings) != 1:
        fail("bindings", "revision/source identity differs across baseline rows")
    binary_bindings = {(row["mode"], json.dumps(row["binding"]["binary"], sort_keys=True)) for row in rows}
    if len(binary_bindings) != len(MODES):
        fail("bindings", "binary identity differs within a normal or allocator lane")
    for source_mode, selector in SELECTORS.items():
        role_spec = obj(protocol['roles'][source_mode], 'protocol role')
        if role_spec.get('selector') != selector or role_spec.get('source_field') != 'opc_part_add':
            fail('protocol role', 'selector/source binding differs')
    profiles = [profile_summary(root, kind, protocol_sha, source_mode) for source_mode in ('observed','plain') for kind in ('stat','record')]
    return {
        "schema": "litchi-0445-opc-part-add-baseline-summary-v1",
        "change": CHANGE,
        "protocol_sha256": protocol_sha,
        "scope": "matched current-build observed/plain OPC Part-addition baseline; no production speedup claim",
        "matrix": {"reports": len(rows), "retained_samples": len(rows) * SAMPLES, "samples": SAMPLES, "warmups": WARMUPS, "cpu": 2, "workers": 1, "phases": list(PHASES), "selectors": SELECTORS},
        "bindings": {"revision": next(iter(source_bindings))[0], "source_rows": len(source_bindings), "binary_lane_bindings": len(binary_bindings)},
        "rows": rows,
        "repeat_review": repeat_flags(rows),
        "comparisons": comparisons(rows),
        "profiles": profiles,
        "claims": [
            "descriptive current-revision timing, allocation, process-RSS, and hardware-counter observations for the named selector and declared scope",
            "normal p50/p95/p99 vectors include deterministic bootstrap intervals over 30 samples from each fresh process",
            "profile self rows are whole-process observations and do not isolate a causal Part-addition hotspot",
            "no before/after speedup, fixed-memory, cache-miss, scaling, or allocator-latency claim is authorized",
        ],
    }


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode("utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--attempt", default="formal")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    root = args.root.resolve()
    output = (args.output or root / "summary.json").resolve()
    try:
        value = derive(root, args.attempt)
        if args.check:
            if not output.is_file() or canonical(load_json(output, str(output))) != canonical(value):
                fail(str(output), "retained summary differs from derivation")
        if args.write:
            output.write_bytes(canonical(value) + b"\n")
    except (OSError, KeyError, TypeError, ValueError, DeriveError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    print(json.dumps({"status": "pass", "output": str(output), "reports": 24, "profiles": 4}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
