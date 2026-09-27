"""Offline replay and custody analysis for the 0788 memory packet.

The benchmark binaries are deliberately not imported or executed here.  This
module consumes only retained receipts and report artifacts, reconstructs the
corpus independently, and emits deterministic summaries.  In particular,
the diagnostic probe is kept separate from native timing and RSS populations.
The packet is an attribution experiment for the already rejected 0787
candidate; it never makes an adoption decision.
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import math
import random
import re
import statistics
import struct
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
ORIGIN_PATH = PACKET / "origin.json"
PLAN_PATH = PACKET / "plan.json"
REPORT_SCHEMA = "litchi.execution-baseline.v1"
# ``memory_analysis.py`` owns the phase-snapshot document and uses the
# ``...-memory-analysis.v1`` schema.  Keep this packet-level replay document
# distinct so the two deterministic artifacts cannot be confused.
ANALYSIS_SCHEMA = "litchi.cached-part-memory-replay.0788.v1"
BOOTSTRAP_SEED = 788078
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_CONFIDENCE = 0.95
BOOTSTRAP_INDEXES = [249, 9749]
EXPECTED_PRODUCTION_FILES = 9_196
EXPECTED_ARCHITECTURE_FILES = 35
LEGS = ("before", "after")
NATIVE_METRICS = ("p50", "p95", "p99", "rss", "minor_faults", "major_faults")
HEX = frozenset("0123456789abcdefABCDEF")
SOURCE_METRICS = {
    "small": {"requested_bytes": 73590,
               "histogram": [32, 0, 0, 32, 0, 0, 0, 0]},
    "large": {"requested_bytes": 146041,
               "histogram": [32, 0, 0, 0, 32, 0, 0, 0]},
    "mixed": {"requested_bytes": 143782,
               "histogram": [32, 0, 0, 1, 31, 0, 0, 0]},
}


class ReplayError(RuntimeError):
    """Raised for missing, stale, malformed, or contradictory evidence."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def is_revision(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 40 and all(c in HEX for c in value)


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def historical_rel(path: Path) -> str:
    """Return a relocation-stable path for the retained sibling 0787 packet."""
    sibling = (PACKET.parent / "change-0787").resolve()
    try:
        return "../change-0787/" + str(path.relative_to(sibling))
    except ValueError:
        return rel(path)


def origin() -> dict[str, Any]:
    value = read_json(ORIGIN_PATH)
    require(isinstance(value, dict), "origin.json is malformed")
    require(is_revision(value.get("base")), "origin base revision is invalid")
    require(isinstance(value.get("owned_worktree"), str)
            and value["owned_worktree"], "origin owned worktree is missing")
    return value


def _owned_root() -> Path:
    return Path(origin()["owned_worktree"]).resolve()


def _path_candidates(raw: Path) -> list[Path]:
    candidates: list[Path] = []
    if raw.is_absolute():
        try:
            candidates.append(PACKET / raw.resolve().relative_to(_owned_root()))
        except ValueError:
            pass
        text = raw.as_posix()
        for marker in ("/docs/performance/results/change-0788/",
                       "/docs/performance/results/change-0787/"):
            if marker in text:
                relative = text.split(marker, 1)[1]
                if marker.endswith("0788/"):
                    candidates.append(PACKET / relative)
                else:
                    candidates.append(PACKET.parent / "change-0787" / relative)
        candidates.append(raw)
    else:
        text = raw.as_posix()
        if text.startswith("docs/performance/results/change-0788/"):
            candidates.append(PACKET / text.split("change-0788/", 1)[1])
        elif text.startswith("docs/performance/results/change-0787/"):
            candidates.append(PACKET.parent / "change-0787" /
                             text.split("change-0787/", 1)[1])
        candidates.extend((PACKET / raw, ROOT / raw))
    unique: list[Path] = []
    for candidate in candidates:
        candidate = candidate.resolve(strict=False)
        if candidate not in unique:
            unique.append(candidate)
    return unique


def resolve_path(value: Any, *, packet_bound: bool = True) -> Path:
    require(isinstance(value, str) and value, f"invalid artifact path: {value!r}")
    candidates = _path_candidates(Path(value))
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            if packet_bound:
                try:
                    candidate.relative_to(PACKET.resolve())
                except ValueError:
                    continue
            return candidate
    fallback = candidates[0] if candidates else Path(value).resolve(strict=False)
    if packet_bound:
        try:
            fallback.relative_to(PACKET.resolve())
        except ValueError:
            fail(f"artifact path escaped packet: {value}")
    return fallback


def artifact(value: Any, label: str, *, packet_bound: bool = True,
             allow_missing: bool = False) -> Path | None:
    require(isinstance(value, dict), f"{label} artifact is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = resolve_path(raw, packet_bound=packet_bound)
    if not path.is_file():
        if allow_missing:
            return None
        fail(f"missing {label}: {raw}")
    require(not path.is_symlink(), f"{label} is a symlink: {raw}")
    require(path.stat().st_size == size, f"{label}.bytes changed")
    require(sha256(path) == digest, f"{label}.sha256 changed")
    return path


def artifact_path(value: Any, label: str, *, packet_bound: bool = True) -> Path:
    path = artifact(value, label, packet_bound=packet_bound)
    assert path is not None
    return path


def file_identity(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def _walk(value: Any, path: tuple[str, ...] = ()) -> Iterable[tuple[tuple[str, ...], Any]]:
    yield path, value
    if isinstance(value, dict):
        for key, child in value.items():
            yield from _walk(child, path + (str(key),))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from _walk(child, path + (str(index),))


def lookup(value: Any, names: Iterable[str]) -> Any:
    wanted = {name.lower().replace("_", "").replace("-", "") for name in names}
    for path, child in _walk(value):
        if path and path[-1].lower().replace("_", "").replace("-", "") in wanted:
            return child
    return None


def _field(value: dict[str, Any], name: str, *aliases: str) -> Any:
    if name in value:
        return value[name]
    for alias in aliases:
        if alias in value:
            return value[alias]
    nested = value.get("config")
    if isinstance(nested, dict):
        if name in nested:
            return nested[name]
        for alias in aliases:
            if alias in nested:
                return nested[alias]
    return lookup(value, (name, *aliases))


def _case_key(case: dict[str, Any]) -> tuple[Any, ...]:
    return (case["route"], case["shape"], case["state"],
            case["task_floor"], case["workers"])


def _case_copy(case: dict[str, Any]) -> dict[str, Any]:
    return {key: case[key] for key in ("route", "shape", "state", "task_floor", "workers")}


def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "litchi.cached-part-memory.0788.v1",
            "plan schema changed")
    require(plan.get("affinity") == list(range(32)), "capture affinity changed")
    require(plan.get("bootstrap") == {"ci": 0.95, "indexes": BOOTSTRAP_INDEXES,
                                       "resamples": BOOTSTRAP_RESAMPLES,
                                       "seed": BOOTSTRAP_SEED},
            "bootstrap contract changed")
    expected = plan.get("expected_counts")
    require(isinstance(expected, dict), "expected counts are missing")
    wanted = {"all_reports": 220, "all_samples": 5092, "heaptrack_reports": 16,
              "heaptrack_samples": 480, "memory_off_reports": 32,
              "memory_off_samples": 496, "memory_on_reports": 32,
              "memory_on_samples": 496, "native_reports": 120,
              "native_samples": 3600, "qualification_reports": 20,
              "qualification_samples": 20}
    require(expected == wanted, "frozen expected counts changed")
    for name in ("native", "qualification", "memory", "heaptrack"):
        require(isinstance(plan.get(name), dict), f"{name} plan is missing")
    require(isinstance(plan.get("purpose"), str)
            and "Diagnostic attribution only" in plan["purpose"],
            "diagnostic purpose changed")
    return plan


def expected_cases(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    value = plan[lane].get("cases")
    require(isinstance(value, list) and value, f"{lane} cases are missing")
    result: list[dict[str, Any]] = []
    for index, case in enumerate(value):
        require(isinstance(case, dict), f"{lane} case {index} is malformed")
        for key in ("route", "shape", "state", "task_floor", "workers"):
            require(key in case, f"{lane} case {index} lacks {key}")
        result.append(_case_copy(case))
    return result


def payload_oracle() -> dict[str, Any]:
    """Rebuild the 0786/0787 deterministic member payloads independently."""
    result: dict[str, Any] = {}
    for shape in ("small", "large", "mixed"):
        members: list[str] = []
        sequence = hashlib.sha256()
        logical = 0
        for index in range(32):
            size = 4096 if shape == "small" or (shape == "mixed" and index == 31) else 262144
            label = f"litchi-0786-member-{index:02}-".encode()
            offset = index * 11 if index * 11 <= 255 else 0
            period = bytes((label[k % len(label)] + (k % 97) * 3 + offset) % 256
                           for k in range(len(label) * 97))
            payload = (period * ((size + len(period) - 1) // len(period)))[:size]
            members.append(hashlib.sha256(payload).hexdigest())
            sequence.update(struct.pack("<Q", index))
            sequence.update(payload)
            logical += size
        result[shape] = {"members": members, "sequence": sequence.hexdigest(),
                         "logical_bytes": logical}
    return result


def _expected_sizes(shape: str) -> list[int]:
    require(shape in ("small", "large", "mixed"), f"unknown corpus shape: {shape}")
    return [4096 if shape == "small" or (shape == "mixed" and i == 31)
            else 262144 for i in range(32)]


def _verification_ok(value: Any) -> bool:
    if not isinstance(value, dict):
        return False
    positive = False
    for _, child in _walk(value):
        if isinstance(child, bool):
            if not child:
                return False
            positive = True
    return positive


def check_corpus(report: dict[str, Any], expected_shape: str,
                 oracle: dict[str, Any], label: str) -> None:
    corpus = report.get("corpus")
    require(isinstance(corpus, dict), f"{label} corpus is missing")
    require(corpus.get("shape") == expected_shape, f"{label} corpus shape changed")
    require(corpus.get("metadata_members") == ["[Content_Types].xml", "_rels/.rels"],
            f"{label} metadata member manifest changed")
    members = corpus.get("members")
    require(isinstance(members, list) and len(members) == 32,
            f"{label} member cardinality changed")
    expected_sizes = _expected_sizes(expected_shape)
    for index, (member, size, digest) in enumerate(zip(members, expected_sizes,
                                                       oracle[expected_shape]["members"])):
        require(isinstance(member, dict), f"{label} member {index} is malformed")
        require(member.get("index") == index and member.get("bytes") == size,
                f"{label} member {index} size/order changed")
        require(member.get("sha256") == digest, f"{label} member {index} payload changed")
        require(member.get("opc_name") == f"custom/member{index:02}.bin"
                and member.get("opc_uri") == f"/custom/member{index:02}.bin"
                and member.get("cfb_name") == f"Member{index:02}",
                f"{label} member {index} naming changed")
    require(corpus.get("selected_payload_member_count") == 32,
            f"{label} selected member count changed")
    require(is_sha(corpus.get("opc_sha256")) and is_sha(corpus.get("cfb_sha256")),
            f"{label} container identity is invalid")


def _source_metrics(sample: dict[str, Any]) -> dict[str, Any] | None:
    value = sample.get("source_metrics", sample.get("source"))
    if value is None:
        return None
    if isinstance(value, bool):
        return {"available": value}
    require(isinstance(value, dict), "source metrics are malformed")
    aliases = {
        "availability": ("availability",),
        "calls": ("logical_calls", "calls", "read_calls"),
        "requested_bytes": ("requested_bytes", "bytes_requested"),
        "returned_bytes": ("returned_bytes", "bytes_returned"),
        "short_reads": ("short_reads", "short_read_count"),
        "active": ("active_reads_after_operation", "active", "active_reads"),
        "max_active": ("max_simultaneous_reads", "max_active", "max_concurrent"),
        "request_histogram": ("request_size_histogram", "request_histogram", "histogram"),
    }
    result = {"available": False}
    availability = value.get("availability", "")
    result["availability"] = availability
    result["available"] = bool(value.get("available", False)) or (
        isinstance(availability, str) and "unavailable" not in availability
        and "not-applicable" not in availability and "feature" in availability)
    for name, names in aliases.items():
        for key in names:
            if key in value:
                result[name] = value[key]
                break
    return result


def _resource_contract(sample: dict[str, Any], expected: dict[str, Any],
                       label: str) -> dict[str, Any]:
    resources = sample.get("resources", sample.get("resource_snapshots"))
    require(isinstance(resources, dict), f"{label} resources are missing")
    limits = resources.get("limits")
    require(isinstance(limits, dict), f"{label} resource limits are missing")
    for key in ("workers", "io_concurrency", "cpu_tasks"):
        nonnegative_int(limits.get(key), f"{label} limit {key}")
    require(limits["workers"] == expected["workers"]
            and limits["io_concurrency"] == expected["workers"],
            f"{label} worker limits changed")
    require(limits["cpu_tasks"] == 1_000_000, f"{label} CPU task limit changed")
    snapshots: dict[str, dict[str, int]] = {}
    for marker in ("before_operation", "after_operation", "after_drop"):
        value = resources.get(marker)
        require(isinstance(value, dict), f"{label} {marker} is missing")
        snapshots[marker] = {}
        for key in ("workers", "io_concurrency", "cpu_tasks"):
            counter = value.get(key)
            nonnegative_int(counter, f"{label} {marker}.{key}")
            require(counter <= limits[key], f"{label} {marker}.{key} exceeds limit")
            snapshots[marker][key] = counter
    initial_cpu = 0 if expected["state"] == "fresh" else 32
    require(snapshots["before_operation"]["cpu_tasks"] == initial_cpu
            and snapshots["after_operation"]["cpu_tasks"] == initial_cpu + 32
            and snapshots["after_drop"]["cpu_tasks"] == initial_cpu + 32,
            f"{label} CPU task lifecycle changed")
    require(snapshots["after_drop"]["workers"] == 0
            and snapshots["after_drop"]["io_concurrency"] == 0,
            f"{label} permits were not released")
    require(resources.get("worker_and_io_released") is True
            and resources.get("cpu_tasks_within_limit") is True,
            f"{label} resource witness failed")
    return {"limits": {key: limits[key] for key in ("workers", "io_concurrency", "cpu_tasks")},
            "snapshots": snapshots, "worker_and_io_released": True,
            "cpu_tasks_within_limit": True}


def _check_source_metrics(sample: dict[str, Any], expected: dict[str, Any],
                          feature: bool, label: str, logical_bytes: int) -> dict[str, Any]:
    metrics = _source_metrics(sample)
    require(metrics is not None, f"{label} source metrics are missing")
    if not feature:
        require(metrics.get("available") is False
                and metrics.get("availability") == "unavailable-normal-build",
                f"{label} has unexpected source metrics")
        for key in ("calls", "requested_bytes", "returned_bytes", "short_reads",
                    "active", "max_active", "request_histogram"):
            require(metrics.get(key) is None, f"{label} carries disabled source metric {key}")
        return {"available": False, "availability": metrics.get("availability", "")}
    require(metrics.get("available") is True, f"{label} source metrics unavailable")
    fresh = expected["state"] == "fresh"
    required = ("calls", "requested_bytes", "returned_bytes", "short_reads",
                "active", "max_active", "request_histogram")
    for name in required:
        require(name in metrics and metrics[name] is not None, f"{label} {name} missing")
    for name in ("calls", "requested_bytes", "returned_bytes", "short_reads", "active", "max_active"):
        nonnegative_int(metrics[name], f"{label} source {name}")
    require(metrics["max_active"] <= expected["workers"], f"{label} max active exceeds workers")
    if not fresh:
        require(metrics["max_active"] == 0, f"{label} primed source concurrency changed")
    require(metrics["active"] == 0 and metrics["short_reads"] == 0,
            f"{label} source reads did not quiesce")
    histogram = metrics["request_histogram"]
    require(isinstance(histogram, list) and all(isinstance(x, int) and x >= 0 for x in histogram),
            f"{label} request histogram is malformed")
    require(sum(histogram) == metrics["calls"], f"{label} request histogram does not sum")
    calls = 64 if fresh else 0
    # The observer counts logical ReadAt requests.  Requested bytes are the
    # compressed source ranges, so they intentionally differ from the
    # verified decompressed payload byte count.
    bytes_expected = SOURCE_METRICS[expected["shape"]]["requested_bytes"] if fresh else 0
    require(metrics["calls"] == calls and metrics["requested_bytes"] == bytes_expected
            and metrics["returned_bytes"] == bytes_expected,
            f"{label} source accounting changed")
    require(metrics["request_histogram"] ==
            (SOURCE_METRICS[expected["shape"]]["histogram"] if fresh else [0] * 8),
            f"{label} source request histogram changed")
    return {key: metrics[key] for key in ("available", "availability", "calls",
                                          "requested_bytes", "returned_bytes", "short_reads",
                                          "active", "max_active", "request_histogram")}


def parse_report(path: Path, receipt: dict[str, Any], expected: dict[str, Any],
                 *, feature: bool, samples_expected: int, warmup_expected: int,
                 oracle: dict[str, Any], label: str) -> dict[str, Any]:
    report = read_json(path)
    require(isinstance(report, dict), f"{label} report is malformed")
    require(report.get("schema", report.get("schema_version")) == REPORT_SCHEMA,
            f"{label} report schema changed")
    for key in ("route", "shape", "state", "workers", "task_floor"):
        require(_field(report, key) == expected[key], f"{label} {key} differs from receipt")
    config = report.get("config")
    require(isinstance(config, dict), f"{label} config is missing")
    require(config.get("samples") == samples_expected and config.get("warmup") == warmup_expected
            and config.get("aggregate_parallel_bytes") == 65536
            and config.get("cpu_task_limit") == 1_000_000,
            f"{label} report configuration changed")
    check_corpus(report, expected["shape"], oracle, label)
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == samples_expected,
            f"{label} sample count changed")
    metrics = report.get("metrics")
    require(isinstance(metrics, dict)
            and metrics.get("source_metrics_feature") is feature
            and isinstance(metrics.get("cpu_clock"), str)
            and "ProcessCPUTime" in metrics["cpu_clock"],
            f"{label} metric availability changed")
    logical_bytes = oracle[expected["shape"]]["logical_bytes"]
    walls: list[float] = []
    cpus: list[float] = []
    resources: list[dict[str, Any]] = []
    source_rows: list[dict[str, Any]] = []
    output_sequence: set[str] = set()
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"{label} sample {index} is malformed")
        wall = sample.get("wall_ns")
        finite_number(wall, f"{label} sample {index} wall_ns")
        require(float(wall) > 0, f"{label} sample {index} wall_ns is not positive")
        walls.append(float(wall))
        cpu = sample.get("cpu_ns")
        if cpu is not None:
            finite_number(cpu, f"{label} sample {index} cpu_ns")
            require(float(cpu) >= 0, f"{label} sample {index} cpu_ns is negative")
            cpus.append(float(cpu))
        verification = sample.get("verification")
        require(isinstance(verification, dict)
                and verification.get("ordered") is True
                and verification.get("all_member_sha256_match") is True
                and verification.get("members") == 32
                and verification.get("logical_bytes") == logical_bytes
                and verification.get("sequence_sha256") == oracle[expected["shape"]]["sequence"],
                f"{label} sample {index} payload verification failed")
        resources.append(_resource_contract(sample, expected, f"{label} sample {index}"))
        source_rows.append(_check_source_metrics(sample, expected, feature,
                                                  f"{label} sample {index}", logical_bytes))
        output_sequence.add(verification["sequence_sha256"])
    require(len(output_sequence) == 1, f"{label} output sequence changed")
    require(not cpus or len(cpus) == len(samples), f"{label} CPU metric availability changed")
    return {
        "receipt": {key: receipt.get(key) for key in
                    ("lane", "block", "leg", "mode", "repeat", "protocol")
                    if key in receipt},
        "case": _case_copy(expected),
        "report": file_identity(path),
        "samples": len(samples),
        "wall_ns": walls,
        "cpu_ns": cpus if cpus else None,
        "p50_ns": nearest_rank(walls, 0.50),
        "p95_ns": nearest_rank(walls, 0.95),
        "p99_ns": nearest_rank(walls, 0.99),
        "mean_ns": statistics.fmean(walls),
        "corpus_fingerprint": corpus_fingerprint(report["corpus"]),
        "source_metrics": source_rows,
        "resource_contract": resources,
        "verification_ok": True,
    }


def corpus_fingerprint(corpus: dict[str, Any]) -> str:
    selected = {key: corpus[key] for key in
                ("shape", "metadata_members", "members", "selected_payload_member_count",
                 "opc_sha256", "cfb_sha256")}
    return hashlib.sha256(json.dumps(selected, sort_keys=True,
                                     separators=(",", ":")).encode()).hexdigest()


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    ordered = sorted(float(value) for value in values)
    require(ordered, "nearest-rank received no values")
    index = max(0, min(len(ordered) - 1, math.ceil(quantile * len(ordered)) - 1))
    return ordered[index]


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(values, "bootstrap received no values")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates = [statistics.median(rng.choice(values) for _ in values)
                 for _ in range(BOOTSTRAP_RESAMPLES)]
    estimates.sort()
    return {"estimate": statistics.median(values), "lower": estimates[BOOTSTRAP_INDEXES[0]],
            "upper": estimates[BOOTSTRAP_INDEXES[1]], "seed": BOOTSTRAP_SEED,
            "resamples": BOOTSTRAP_RESAMPLES, "confidence": BOOTSTRAP_CONFIDENCE,
            "statistic": "median", "endpoint_indexes": BOOTSTRAP_INDEXES}


def paired_metric(before: list[float], after: list[float], name: str) -> dict[str, Any]:
    require(len(before) == len(after) == 6, f"{name} block cardinality changed")
    require(all(math.isfinite(x) and x >= 0 for x in before + after),
            f"{name} contains non-finite values")
    deltas = [new - old for old, new in zip(before, after)]
    ratios: list[float | None] = [new / old if old > 0 else None
                                  for old, new in zip(before, after)]
    positive = [float(value) for value in ratios if value is not None]
    result: dict[str, Any] = {
        "name": name,
        "before_block_values": before,
        "after_block_values": after,
        "block_deltas": deltas,
        "block_ratios": ratios,
        "raw_distributions_retained": True,
        "bootstrap": None,
        "difference_bootstrap": bootstrap(deltas),
    }
    if len(positive) == 6:
        result["bootstrap"] = bootstrap(positive)
        result["estimate"] = result["bootstrap"]["estimate"]
        result["ci95_low"] = result["bootstrap"]["lower"]
        result["ci95_high"] = result["bootstrap"]["upper"]
    else:
        result["estimate"] = None
        result["ci95_low"] = None
        result["ci95_high"] = None
    result["difference_estimate"] = result["difference_bootstrap"]["estimate"]
    result["difference_ci95_low"] = result["difference_bootstrap"]["lower"]
    result["difference_ci95_high"] = result["difference_bootstrap"]["upper"]
    return result


def _normalise_token(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    for prefix in (str(_owned_root()), str(ROOT.resolve())):
        value = value.replace(prefix, str(ROOT.resolve()))
    return value


def normalise_command(value: Any) -> list[Any]:
    require(isinstance(value, list) and value, "command receipt is malformed")
    return [_normalise_token(item) for item in value]


def cleanup_value(cleanup: Any, receipt: dict[str, Any]) -> bool:
    wanted = (receipt.get("path"), receipt.get("bytes", receipt.get("size")),
              receipt.get("sha256", receipt.get("digest")))
    if not (isinstance(wanted[0], str) and isinstance(wanted[1], int) and is_sha(wanted[2])):
        return False
    if isinstance(cleanup, dict):
        actual = (cleanup.get("path"), cleanup.get("bytes", cleanup.get("size")),
                  cleanup.get("sha256", cleanup.get("digest")))
        return actual == wanted or any(cleanup_value(child, receipt)
                                      for child in cleanup.values())
    if isinstance(cleanup, list):
        return any(cleanup_value(child, receipt) for child in cleanup)
    return False


def load_cleanup() -> tuple[Any, bool]:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None, False
    value = read_json(path)
    require(isinstance(value, dict), "cleanup.json is malformed")
    return value, bool(value.get("verified") is True
                       or value.get("executables_verified_before_removal") is True)


def validate_binary(value: Any, label: str, cleanup: Any, cleanup_verified: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} binary identity is missing")
    path = artifact(value, label, packet_bound=False, allow_missing=True)
    if path is None:
        require(cleanup_verified and cleanup_value(cleanup, value),
                f"{label} is missing without exact cleanup witness")
    return {"path": value["path"], "bytes": value["bytes"],
            "sha256": value["sha256"], "custody_verified": True}


def _git_output(arguments: list[str], label: str, *, binary: bool = False) -> bytes | str:
    try:
        completed = subprocess.run(["git", *arguments], cwd=ROOT, check=True,
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {label}: {error}")
    return completed.stdout if binary else completed.stdout.decode()


def _git_blob_hashes(revision: str, names: list[str]) -> dict[str, str]:
    request = b"".join(f"{revision}:{name}\n".encode() for name in names)
    # ``_git_output`` cannot carry input; use a direct batch invocation.
    try:
        completed = subprocess.run(["git", "cat-file", "--batch"], cwd=ROOT,
                                   input=request, stdout=subprocess.PIPE, check=True)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read production Git blobs: {error}")
    data = completed.stdout
    offset = 0
    result: dict[str, str] = {}
    for name in names:
        end = data.find(b"\n", offset)
        require(end >= 0, f"Git blob header is truncated: {name}")
        header = data[offset:end].split()
        require(len(header) == 3 and header[1] == b"blob", f"Git blob is invalid: {name}")
        size = int(header[2])
        offset = end + 1
        require(offset + size <= len(data), f"Git blob is truncated: {name}")
        result[name] = hashlib.sha256(data[offset:offset + size]).hexdigest()
        offset += size
        require(data[offset:offset + 1] == b"\n", f"Git blob terminator is missing: {name}")
        offset += 1
    require(offset == len(data), "Git blob batch has trailing data")
    return result


def production_names(revision: str) -> list[str]:
    raw = _git_output(["ls-tree", "-r", "-z", "--name-only", revision, "--",
                       "crates", "Cargo.toml", "clippy.toml", ".cargo/config.toml",
                       "rust-toolchain.toml"], "production census", binary=True)
    assert isinstance(raw, bytes)
    names = [item.decode() for item in raw.split(b"\0") if item]
    require(len(names) == EXPECTED_PRODUCTION_FILES, "production file census changed")
    require(len(set(names)) == len(names), "production file census contains duplicates")
    return names


def _base_production_files(revision: str) -> dict[str, str]:
    return _git_blob_hashes(revision, production_names(revision))


def load_source(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is malformed")
    revision = value.get("revision")
    require(is_revision(revision), f"{label}.revision is invalid")
    files = value.get("files")
    require(isinstance(files, dict) and len(files) == EXPECTED_PRODUCTION_FILES,
            f"{label}.files cardinality changed")
    require(all(isinstance(name, str) and name and is_sha(digest)
                for name, digest in files.items()), f"{label}.files is malformed")
    return {"revision": revision, "files": dict(files)}


def load_tool(value: Any, label: str) -> dict[str, Any]:
    if isinstance(value, dict) and "files" in value:
        files = value["files"]
        revision = value.get("revision")
    else:
        files = value
        revision = origin()["base"]
    require(isinstance(files, dict) and files and
            all(isinstance(k, str) and is_sha(v) for k, v in files.items()),
            f"{label} is malformed")
    if revision is not None:
        require(is_revision(revision), f"{label}.revision is invalid")
    return {"revision": revision, "files": dict(files)}


def current_tool() -> dict[str, str]:
    tool = ROOT / "tools/perf-execution"
    require(tool.is_dir(), "standalone native tool is missing")
    return {str(path.relative_to(ROOT)): sha256(path)
            for path in sorted(tool.rglob("*"))
            if path.is_file() and "target" not in path.parts}


def probe_files() -> dict[str, str]:
    directory = PACKET / "probe-src"
    return {path.name: sha256(path) for path in sorted(directory.iterdir()) if path.is_file()}


def source_identity(path: Path, label: str) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict), f"{label} source manifest is malformed")
    production = load_source(value.get("production"), f"{label} production source")
    tool = load_tool(value.get("tool"), f"{label} tool source")
    probe = value.get("probe")
    require(isinstance(probe, dict) and probe and
            all(isinstance(k, str) and is_sha(v) for k, v in probe.items()),
            f"{label} probe source is malformed")
    require(probe == probe_files(), f"{label} probe source changed")
    return {"production": production, "tool": tool, "probe": dict(probe)}


def load_build(leg: str, plan: dict[str, Any], base_files: dict[str, str],
               cleanup: Any, cleanup_verified: bool) -> dict[str, Any]:
    path = PACKET / f"build-{leg}" / "build.json"
    if not path.is_file():
        path = PACKET / "builds" / f"{leg}.json"
    build = read_json(path)
    require(isinstance(build, dict), f"{leg} build is malformed")
    source_path = artifact_path(build.get("source"), f"{leg} build source")
    source = source_identity(source_path, f"{leg} build")
    require(source["production"]["revision"] == origin()["base"],
            f"{leg} build production revision changed")
    require(source["tool"]["files"] == current_tool(), f"{leg} native tool source changed")
    require(len(source["production"]["files"]) == EXPECTED_PRODUCTION_FILES,
            f"{leg} production source count changed")
    if leg == "before":
        require(source["production"]["files"] == base_files,
                "before production source differs from base Git blobs")
    frozen_path = artifact_path(build.get("frozen_inputs"), f"{leg} frozen inputs")
    frozen = read_json(frozen_path)
    require(isinstance(frozen, dict) and isinstance(frozen.get("files", frozen), dict),
            f"{leg} frozen input manifest is malformed")
    frozen_map = frozen.get("files", frozen)
    for name, digest in frozen_map.items():
        require(isinstance(name, str) and is_sha(digest), f"{leg} frozen input is malformed")
        current = PACKET / name
        if current.is_file() and sha256(current) == digest:
            continue
        # The first ten qualification-before children were captured with a
        # decoder that called GNU time's %R field ``elapsed_seconds``.  The
        # raw .time artifacts are authoritative; capture-correction.json binds
        # the archived pre-correction driver to the frozen build input and
        # records the exact byte-preserving correction.
        if name == "capture.py":
            correction_path = PACKET / "capture-correction.json"
            archived = PACKET / "capture-pre-correction.py"
            require(correction_path.is_file() and archived.is_file()
                    and sha256(archived) == digest,
                    f"{leg} frozen capture driver cannot be replayed")
            correction = read_json(correction_path)
            require(isinstance(correction, dict)
                    and correction.get("before", {}).get("sha256") == digest
                    and correction.get("after", {}).get("sha256") == sha256(current),
                    f"{leg} capture correction witness changed")
            continue
        require(current.is_file() and sha256(current) == digest,
                f"{leg} frozen input changed: {name}")
    rows = build.get("rows")
    if rows is None:
        rows = build.get("commands")
    require(isinstance(rows, list) and len(rows) == 2, f"{leg} build command set changed")
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"{leg} build command {index} failed")
        normalise_command(row.get("command"))
        artifact_path(row.get("log"), f"{leg} build command {index} log")
    env = build.get("environment")
    require(isinstance(env, dict) and env.get("CARGO_BUILD_JOBS") == "2"
            and env.get("CARGO_INCREMENTAL") == "0", f"{leg} build environment changed")
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"native", "memory"},
            f"{leg} binary set changed")
    checked = {kind: validate_binary(value, f"{leg} {kind} binary", cleanup, cleanup_verified)
               for kind, value in binaries.items()}
    return {"receipt": file_identity(path), "source": source,
            "frozen_inputs": frozen_map, "binaries": checked,
            "environment": {key: env.get(key) for key in
                             ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                              "RUSTFLAGS", "RUSTUP_TOOLCHAIN")},
            "commands": [list(row["command"]) for row in rows]}


def load_quality(after: dict[str, Any], base_files: dict[str, str]) -> dict[str, Any]:
    path = PACKET / "quality.json"
    quality = read_json(path)
    require(isinstance(quality, dict), "quality.json is malformed")
    source_path = artifact_path(quality.get("source"), "quality source")
    source = source_identity(source_path, "quality")
    # The fresh probe quality run is intentionally performed with the restored
    # baseline live.  Candidate production quality is bound independently to
    # the sealed 0787 quality receipt below; accepting either frozen production
    # map here prevents the probe gate from being misrepresented as candidate
    # production quality.
    require(source["tool"] == after["source"]["tool"]
            and source["probe"] == after["source"]["probe"]
            and source["production"]["files"] in (
                after["source"]["production"]["files"],
                base_files),
            "quality source differs from frozen build sources")
    rows = quality.get("rows")
    require(isinstance(rows, list) and len(rows) == 6, "quality gate cardinality changed")
    logs: list[str] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"quality gate {index} failed")
        normalise_command(row.get("command"))
        logs.append(rel(artifact_path(row.get("log"), f"quality gate {index} log")))
    return {"receipt": file_identity(path), "source": rel(source_path),
            "gates": len(rows), "logs": logs}


def load_architecture() -> dict[str, Any]:
    value = read_json(PACKET / "architecture-inputs.json")
    require(isinstance(value, dict) and len(value) == EXPECTED_ARCHITECTURE_FILES,
            "architecture input cardinality changed")
    base = origin()["base"]
    names = sorted(value)
    require(all(isinstance(name, str) and not name.startswith("/") and is_sha(digest)
                for name, digest in value.items()), "architecture input manifest malformed")
    require(_git_blob_hashes(base, names) == value, "architecture input origin blobs changed")
    return {"revision": base, "count": len(value), "files": value,
            "receipt": file_identity(PACKET / "architecture-inputs.json")}


def load_candidate_binding(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    binding = read_json(PACKET / "candidate-binding.json")
    require(isinstance(binding, dict)
            and binding.get("purpose", "").startswith("Attribution only"),
            "candidate binding purpose changed")
    require(binding.get("historical_commit") == "f50e22fc3b649c1cc3b964db71b667ebf4028a36",
            "historical candidate commit changed")
    for name in ("candidate_receipt", "patch"):
        item = binding.get(name)
        require(isinstance(item, dict) and is_sha(item.get("sha256")),
                f"candidate {name} binding is malformed")
        path = resolve_path(item.get("path"), packet_bound=False)
        try:
            path.relative_to((PACKET.parent / "change-0787").resolve())
        except ValueError:
            fail(f"candidate {name} archive escaped historical 0787 packet")
        require(path.is_file() and sha256(path) == item["sha256"],
                f"candidate {name} archive changed")
    files = binding.get("files")
    require(isinstance(files, list) and len(files) == 3, "candidate file set changed")
    before_map = before["source"]["production"]["files"]
    after_map = after["source"]["production"]["files"]
    for item in files:
        require(isinstance(item, dict), "candidate file binding is malformed")
        name = item.get("path")
        require(name in before_map and is_sha(item.get("before_sha256"))
                and is_sha(item.get("candidate_sha256")),
                f"candidate file binding is malformed: {name}")
        require(before_map[name] == item["before_sha256"], f"candidate before differs: {name}")
        require(after_map[name] == item["candidate_sha256"], f"candidate after differs: {name}")
    require(set(after_map) == set(before_map), "candidate production census changed")
    diffs = [name for name in before_map if before_map[name] != after_map[name]]
    require(diffs == [item["path"] for item in files], "candidate diff is not exact")
    historical_quality = binding.get("historical_quality")
    require(isinstance(historical_quality, dict)
            and is_sha(historical_quality.get("sha256")),
            "historical candidate quality binding is malformed")
    quality_path = resolve_path(historical_quality.get("path"), packet_bound=False)
    try:
        quality_path.relative_to((PACKET.parent / "change-0787").resolve())
    except ValueError:
        fail("historical candidate quality escaped the 0787 packet")
    require(quality_path.is_file() and sha256(quality_path) == historical_quality["sha256"],
            "historical candidate quality receipt changed")
    quality = read_json(quality_path)
    require(isinstance(quality, dict) and isinstance(quality.get("rows"), list)
            and len(quality["rows"]) == 6,
            "historical candidate quality gates changed")
    require(all(isinstance(row, dict) and row.get("exit_code") == 0
                for row in quality["rows"]), "historical candidate quality failed")
    quality_source = quality.get("source")
    require(isinstance(quality_source, dict) and is_sha(quality_source.get("sha256")),
            "historical candidate quality source binding is malformed")
    quality_source_path = resolve_path(quality_source.get("path"), packet_bound=False)
    require(quality_source_path.is_file() and sha256(quality_source_path) == quality_source["sha256"],
            "historical candidate quality source changed")
    quality_source_value = read_json(quality_source_path)
    if "production" in quality_source_value:
        historical_production = load_source(quality_source_value["production"],
                                            "historical candidate quality production")
    else:
        historical_production = load_source(quality_source_value,
                                            "historical candidate quality production")
    require(historical_production["files"] == after["source"]["production"]["files"],
            "historical candidate quality source is not the candidate")
    candidate_quality = {"receipt": {"path": historical_rel(quality_path), "bytes": quality_path.stat().st_size,
                                      "sha256": sha256(quality_path)},
                         "source": {"path": historical_rel(quality_source_path),
                                    "bytes": quality_source_path.stat().st_size,
                                    "sha256": sha256(quality_source_path)},
                         "gates": len(quality["rows"]), "candidate_source_checked": True}
    return {"schema": binding.get("schema"), "historical_commit": binding["historical_commit"],
            "purpose": binding["purpose"], "files": files,
            "candidate_receipt": {"path": binding["candidate_receipt"]["path"],
                                   "sha256": binding["candidate_receipt"]["sha256"]},
            "patch": {"path": binding["patch"]["path"], "sha256": binding["patch"]["sha256"]},
            "historical_quality": candidate_quality}


def expected_order(plan: dict[str, Any], lane: str, *, fixed_leg: str | None = None) -> list[dict[str, Any]]:
    spec = plan["qualification" if fixed_leg else lane]
    cases = expected_cases(plan, "qualification" if fixed_leg else lane)
    result: list[dict[str, Any]] = []
    orders = spec.get("orders", [])
    if fixed_leg:
        # Qualification is one fixed-leg sample per case and intentionally
        # has no alternating order list in the frozen plan.
        orders = [fixed_leg]
    for block, order in enumerate(orders):
        if fixed_leg:
            wanted_legs = [fixed_leg]
        else:
            wanted_legs = order
        for case in cases:
            for leg in wanted_legs:
                result.append({"block": block, "leg": leg, **case})
    return result


def _binary_identity(value: Any) -> tuple[Any, Any, Any]:
    require(isinstance(value, dict), "receipt binary identity is malformed")
    return value.get("path"), value.get("bytes", value.get("size")), value.get("sha256")


def load_native_lane(name: str, plan: dict[str, Any], builds: dict[str, Any],
                     oracle: dict[str, Any]) -> list[dict[str, Any]]:
    path = PACKET / name / "receipts.json"
    receipt_value = read_json(path)
    if isinstance(receipt_value, dict):
        receipts = receipt_value.get("receipts", receipt_value.get("rows"))
        require(isinstance(receipts, list), f"{name} receipt list is missing")
    else:
        receipts = receipt_value
    require(isinstance(receipts, list), f"{name} receipts are malformed")
    if name == "native":
        wanted = expected_order(plan, "native")
        expected_count, samples, warmup = 120, 30, 3
    else:
        leg = name.split("-", 1)[1]
        wanted = expected_order(plan, "qualification", fixed_leg=leg)
        expected_count, samples, warmup = 10, 1, 0
    require(len(receipts) == expected_count, f"{name} report cardinality changed")
    result: list[dict[str, Any]] = []
    for index, (receipt, expected) in enumerate(zip(receipts, wanted)):
        require(isinstance(receipt, dict), f"{name} receipt {index} is malformed")
        require(receipt.get("exit_code") == 0, f"{name} receipt {index} failed")
        for key, value in expected.items():
            require(receipt.get(key) == value, f"{name} receipt {index} {key} changed")
        leg = expected["leg"]
        binary = builds[leg]["binaries"]["native"]
        require(_binary_identity(receipt.get("binary")) == _binary_identity(binary),
                f"{name} receipt {index} binary changed")
        report_path = artifact_path(receipt.get("report"), f"{name} report {index}")
        artifact_path(receipt.get("log"), f"{name} log {index}")
        parsed = parse_report(report_path, receipt, expected,
                              feature=False, samples_expected=samples,
                              warmup_expected=warmup, oracle=oracle,
                              label=f"{name} report {index}")
        parsed["block"] = expected["block"]
        parsed["leg"] = leg
        time_value = receipt.get("time", receipt.get("rss"))
        time_path = artifact_path(time_value, f"{name} process time {index}")
        parsed["process"] = parse_process_time(
            time_path, f"{name} process time {index}",
            allow_legacy_minor=name == "qualification-before")
        result.append(parsed)
    return result


def parse_process_time(path: Path, label: str, *, allow_legacy_minor: bool = False) -> dict[str, Any]:
    text = path.read_text().strip()
    # The capture freezes `/usr/bin/time -f "%M %R %F"`: maximum RSS KiB,
    # minor page faults, and major page faults.  The ten qualification-before
    # reports retain an older JSON decoder's ``elapsed_seconds`` key for the
    # second field; their raw `.time` files remain authoritative and this
    # compatibility alias is explicitly bounded to that lane.
    if text.startswith("{"):
        value = json.loads(text)
        require(isinstance(value, dict), f"{label} process record is malformed")
        rss = next((value[key] for key in ("ru_maxrss_kib", "rss_kib", "max_rss_kib", "rss")
                    if key in value), None)
        minor = next((value[key] for key in ("minor_faults", "minor_page_faults", "minor")
                      if key in value), None)
        if minor is None and allow_legacy_minor:
            minor = value.get("elapsed_seconds")
        major = next((value[key] for key in ("major_faults", "major_page_faults", "major")
                      if key in value), None)
        nonnegative_int(rss, f"{label}.rss_kib")
        nonnegative_int(minor, f"{label}.minor_faults")
        nonnegative_int(major, f"{label}.major_faults")
        return {"rss_kib": rss, "minor_faults": minor, "major_faults": major}
    fields = text.split()
    require(len(fields) == 3 and re.fullmatch(r"[0-9]+", fields[0])
            and re.fullmatch(r"[0-9]+", fields[1])
            and re.fullmatch(r"[0-9]+", fields[2]),
            f"{label} must contain RSS, minor faults, and major faults")
    return {"rss_kib": int(fields[0]), "minor_faults": int(fields[1]),
            "major_faults": int(fields[2])}


def parse_rss_control(path: Path, label: str) -> int:
    """Read the diagnostic lane's retained ru_maxrss JSON control."""
    text = path.read_text().strip()
    if text.startswith("{"):
        value = json.loads(text)
        require(isinstance(value, dict), f"{label} is malformed")
        rss = next((value[key] for key in ("ru_maxrss_kib", "rss_kib", "max_rss_kib", "rss")
                    if key in value), None)
    else:
        require(re.fullmatch(r"[0-9]+", text) is not None, f"{label} is malformed")
        rss = int(text)
    nonnegative_int(rss, f"{label}.rss_kib")
    require(rss > 0, f"{label}.rss_kib is zero")
    return rss


def load_memory(plan: dict[str, Any], builds: dict[str, Any], oracle: dict[str, Any]) -> list[dict[str, Any]]:
    path = PACKET / "memory" / "receipts.json"
    receipt_value = read_json(path)
    if isinstance(receipt_value, dict):
        receipts = receipt_value.get("receipts", receipt_value.get("processes",
                                                                     receipt_value.get("rows")))
        require(isinstance(receipts, list), "memory receipts list is missing")
    else:
        receipts = receipt_value
    require(isinstance(receipts, list) and len(receipts) == 64,
            "memory report cardinality changed")
    cases = expected_cases(plan, "memory")
    protocols = {(item["samples"], item["warmup"]) for item in plan["memory"]["protocols"]}
    result: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()
    for index, receipt in enumerate(receipts):
        require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
                f"memory receipt {index} failed")
        nested_case = receipt.get("case")
        source_case = nested_case if isinstance(nested_case, dict) else receipt
        case = {key: source_case.get(key) for key in
                ("route", "shape", "state", "task_floor", "workers")}
        require(case in cases, f"memory receipt {index} has unknown case")
        mode = receipt.get("mode")
        if isinstance(mode, bool):
            mode = "on" if mode else "off"
        require(mode in ("off", "on"), f"memory receipt {index} mode changed")
        protocol = receipt.get("protocol")
        if isinstance(protocol, dict):
            protocol_key = (protocol.get("samples"), protocol.get("warmup"))
        else:
            protocol_key = (receipt.get("samples"), receipt.get("warmup"))
        require(protocol_key in protocols, f"memory receipt {index} protocol changed")
        repeat = receipt.get("repeat")
        require(repeat in (0, 1), f"memory receipt {index} repeat changed")
        leg = receipt.get("leg", receipt.get("build"))
        require(leg in LEGS, f"memory receipt {index} leg changed")
        key = (_case_key(case), protocol_key, repeat, leg, mode)
        require(key not in seen, f"duplicate memory coordinate: {key}")
        seen.add(key)
        binary = builds[leg]["binaries"]["memory"]
        require(_binary_identity(receipt.get("binary")) == _binary_identity(binary),
                f"memory receipt {index} binary changed")
        report_path = artifact_path(receipt.get("report"), f"memory report {index}")
        artifact_path(receipt.get("log"), f"memory log {index}")
        rss_receipt = receipt.get("rss", receipt.get("time", receipt.get("rss_time")))
        rss_path = artifact_path(rss_receipt, f"memory RSS control {index}")
        rss_kib = parse_rss_control(rss_path, f"memory RSS control {index}")
        parsed = parse_report(report_path, receipt, case, feature=True,
                              samples_expected=protocol_key[0], warmup_expected=protocol_key[1],
                              oracle=oracle, label=f"memory report {index}")
        snapshots = receipt.get("snapshots")
        if mode == "on":
            require(snapshots is not None, f"memory receipt {index} snapshots missing")
            snapshot_path = artifact_path(snapshots, f"memory snapshots {index}")
            parsed["snapshots"] = file_identity(snapshot_path)
        else:
            require(snapshots is None or isinstance(snapshots, dict),
                    f"memory receipt {index} snapshots malformed")
            if isinstance(snapshots, dict):
                snapshot_path = artifact(snapshots, f"memory snapshots {index}")
                parsed["snapshots"] = None if snapshot_path is None else file_identity(snapshot_path)
            else:
                parsed["snapshots"] = None
        parsed.update({"mode": mode, "repeat": repeat, "leg": leg,
                       "protocol": {"samples": protocol_key[0], "warmup": protocol_key[1]},
                       "case": case, "rss_kib": rss_kib})
        result.append(parsed)
    require(len(seen) == 64, "memory coordinate coverage changed")
    require(sum(row["samples"] for row in result if row["mode"] == "off") == 496
            and sum(row["samples"] for row in result if row["mode"] == "on") == 496,
            "memory sample totals changed")
    return result


def load_heaptrack(plan: dict[str, Any], builds: dict[str, Any]) -> list[dict[str, Any]]:
    path = PACKET / "heaptrack" / "receipts.json"
    receipt_value = read_json(path)
    if isinstance(receipt_value, dict):
        receipts = receipt_value.get("receipts", receipt_value.get("rows"))
        require(isinstance(receipts, list), "heaptrack receipt list is missing")
    else:
        receipts = receipt_value
    require(isinstance(receipts, list) and len(receipts) == 16,
            "heaptrack report cardinality changed")
    cases = expected_cases(plan, "heaptrack")
    expected_keys = {(_case_key(case), repeat, leg)
                     for case in cases for repeat in (0, 1) for leg in ("before", "after")}
    seen: set[tuple[Any, ...]] = set()
    result: list[dict[str, Any]] = []
    for index, receipt in enumerate(receipts):
        require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
                f"heaptrack receipt {index} failed")
        case = {key: receipt.get(key) for key in
                ("route", "shape", "state", "task_floor", "workers")}
        leg, repeat = receipt.get("leg"), receipt.get("repeat")
        key = (_case_key(case), repeat, leg)
        require(key in expected_keys and key not in seen, f"heaptrack coordinate changed: {key}")
        seen.add(key)
        require(_binary_identity(receipt.get("binary")) ==
                _binary_identity(builds[leg]["binaries"]["native"]),
                f"heaptrack receipt {index} binary changed")
        raw = receipt.get("raw", receipt.get("trace"))
        printed = receipt.get("print", receipt.get("summary"))
        stacks = receipt.get("stacks", receipt.get("flamegraph"))
        raw_path = artifact_path(raw, f"heaptrack raw {index}", packet_bound=False)
        printed_path = artifact_path(printed, f"heaptrack print {index}")
        stacks_path = artifact_path(stacks, f"heaptrack stacks {index}")
        command = normalise_command(receipt.get("command"))
        print_command = normalise_command(receipt.get("print_command"))
        require(receipt.get("merge_backtraces") is False
                or ("-m" in print_command and
                    print_command[print_command.index("-m") + 1] == "0"),
                f"heaptrack receipt {index} merged backtraces")
        result.append({"case": case, "leg": leg, "repeat": repeat,
                       "raw": {"path": rel(raw_path), "bytes": raw_path.stat().st_size,
                               "sha256": sha256(raw_path)},
                       "print": file_identity(printed_path), "stacks": file_identity(stacks_path),
                       "merge_backtraces": False})
    require(seen == expected_keys, "heaptrack coordinate coverage changed")
    return result


def load_memory_analysis() -> dict[str, Any] | None:
    path = PACKET / "memory-analysis.json"
    if not path.is_file():
        return None
    value = read_json(path)
    require(isinstance(value, dict), "memory-analysis.json is malformed")
    return {"path": rel(path), "sha256": sha256(path), "value": value}


def make_pairs(native: list[dict[str, Any]], plan: dict[str, Any]) -> list[dict[str, Any]]:
    grouped: dict[tuple[Any, ...], dict[str, dict[str, Any]]] = {}
    for row in native:
        key = (*_case_key(row["case"]), row["block"])
        require(row["leg"] in LEGS and row["leg"] not in grouped.setdefault(key, {}),
                f"duplicate native pair: {key}")
        grouped[key][row["leg"]] = row
    cases = expected_cases(plan, "native")
    output: list[dict[str, Any]] = []
    for case in cases:
        blocks: list[dict[str, Any]] = []
        for block in range(6):
            key = (*_case_key(case), block)
            require(set(grouped.get(key, {})) == set(LEGS), f"native pair incomplete: {key}")
            old, new = grouped[key]["before"], grouped[key]["after"]
            require(old["corpus_fingerprint"] == new["corpus_fingerprint"],
                    f"native corpus changed: {key}")
            require(old["resource_contract"] == new["resource_contract"],
                    f"native resource contract changed: {key}")
            blocks.append({"before": old, "after": new, "block": block})
        metrics: dict[str, Any] = {}
        fields = {"p50": "p50_ns", "p95": "p95_ns", "p99": "p99_ns",
                  "rss": "rss_kib", "minor_faults": "minor_faults",
                  "major_faults": "major_faults"}
        for name, field in fields.items():
            before = [float(row["before"]["process"][field]) if name in
                      ("rss", "minor_faults", "major_faults")
                      else float(row["before"][field]) for row in blocks]
            after = [float(row["after"]["process"][field]) if name in
                     ("rss", "minor_faults", "major_faults")
                     else float(row["after"][field]) for row in blocks]
            metrics[name] = paired_metric(before, after, name)
        output.append({"case": case, "blocks": 6, "samples_per_report": 30,
                       "metrics": metrics, "fresh_pair": case["state"] == "fresh",
                       "diagnostic_only": True})
    require(len(output) == 10, "native paired case cardinality changed")
    return output


def csv_rows(paired: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for row in paired:
        out = dict(row["case"])
        out.update({"blocks": row["blocks"], "samples_per_report": row["samples_per_report"],
                    "diagnostic_only": row["diagnostic_only"]})
        for name in NATIVE_METRICS:
            metric = row["metrics"][name]
            out[f"{name}_estimate"] = metric["estimate"]
            out[f"{name}_ci95_low"] = metric["ci95_low"]
            out[f"{name}_ci95_high"] = metric["ci95_high"]
            out[f"{name}_difference_estimate"] = metric["difference_estimate"]
            out[f"{name}_block_ratios"] = json.dumps(metric["block_ratios"], separators=(",", ":"))
            out[f"{name}_block_deltas"] = json.dumps(metric["block_deltas"], separators=(",", ":"))
        rows.append(out)
    return rows


def render_markdown(analysis: dict[str, Any], paired: list[dict[str, Any]]) -> str:
    lines = [
        "# 0788 cached-Part memory attribution",
        "",
        "This packet diagnoses the rejected 0787 cached-Part candidate. It does not",
        "revise adoption thresholds or make an adoption decision; the exact candidate",
        "is restored after capture. Native timing, whole-child process RSS and faults,",
        "phase snapshots, and heaptrack allocation profiles measure different scopes.",
        "Diagnostic probe measurements are never pooled with native timing or RSS.",
        "",
        f"- Reports/samples: {analysis['counts']['reports']} / {analysis['counts']['samples']}",
        f"- Native reports/samples: {analysis['counts']['native_reports']} / {analysis['counts']['native_samples']}",
        f"- Memory off/on samples: {analysis['counts']['memory_off_samples']} / {analysis['counts']['memory_on_samples']}",
        f"- Heaptrack reports/samples: {analysis['counts']['heaptrack_reports']} / {analysis['counts']['heaptrack_samples']}",
        f"- Bootstrap seed: {BOOTSTRAP_SEED}; endpoint indexes: {BOOTSTRAP_INDEXES}",
        "- Previous 0787 rejection remains authoritative regardless of diagnostic counters.",
        "",
        "| Shape | State | Floor | Workers | p50 | p95 | p99 | RSS | Minor faults | Major faults |",
        "|---|---|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in paired:
        values = [row["metrics"][name]["estimate"] for name in
                  ("p50", "p95", "p99", "rss", "minor_faults", "major_faults")]
        def fmt(value: Any) -> str:
            return "n/a" if value is None else f"{value:.6g}"
        lines.append("| " + " | ".join([
            row["case"]["shape"], row["case"]["state"], str(row["case"]["task_floor"]),
            str(row["case"]["workers"]), *(fmt(value) for value in values)]) + " |")
    lines.extend(["", "Raw six-block distributions and paired deltas remain in paired.csv and analysis.json.", ""])
    return "\n".join(lines)


def build_analysis() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    plan = load_plan()
    oracle = payload_oracle()
    origin_value = origin()
    base_names = production_names(origin_value["base"])
    base_files = _git_blob_hashes(origin_value["base"], base_names)
    cleanup, cleanup_verified = load_cleanup()
    builds = {leg: load_build(leg, plan, base_files, cleanup, cleanup_verified) for leg in LEGS}
    require(builds["before"]["source"]["tool"] == builds["after"]["source"]["tool"],
            "native tool source differs between builds")
    candidate = load_candidate_binding(builds["before"], builds["after"])
    quality = load_quality(builds["after"], base_files)
    architecture = load_architecture()
    native = load_native_lane("native", plan, builds, oracle)
    qualification_before = load_native_lane("qualification-before", plan, builds, oracle)
    qualification_after = load_native_lane("qualification-after", plan, builds, oracle)
    memory = load_memory(plan, builds, oracle)
    heaptrack = load_heaptrack(plan, builds)
    paired = make_pairs(native, plan)
    all_samples = sum(row["samples"] for row in native)
    all_samples += sum(row["samples"] for row in qualification_before + qualification_after + memory)
    all_samples += 480
    counts = {
        "reports": 120 + 10 + 10 + 64 + 16, "samples": all_samples,
        "native_reports": len(native), "native_samples": sum(row["samples"] for row in native),
        "qualification_reports": len(qualification_before) + len(qualification_after),
        "qualification_samples": sum(row["samples"] for row in qualification_before + qualification_after),
        "memory_reports": len(memory), "memory_off_reports": sum(row["mode"] == "off" for row in memory),
        "memory_on_reports": sum(row["mode"] == "on" for row in memory),
        "memory_off_samples": sum(row["samples"] for row in memory if row["mode"] == "off"),
        "memory_on_samples": sum(row["samples"] for row in memory if row["mode"] == "on"),
        "heaptrack_reports": len(heaptrack), "heaptrack_samples": 480,
    }
    require(counts == {"reports": 220, "samples": 5092, "native_reports": 120,
                       "native_samples": 3600, "qualification_reports": 20,
                       "qualification_samples": 20, "memory_reports": 64,
                       "memory_off_reports": 32, "memory_on_reports": 32,
                       "memory_off_samples": 496, "memory_on_samples": 496,
                       "heaptrack_reports": 16, "heaptrack_samples": 480},
            "aggregate report/sample counts changed")
    memory_analysis = load_memory_analysis()
    output = {
        "schema": ANALYSIS_SCHEMA, "plan_schema": plan["schema"],
        "report_schema": REPORT_SCHEMA,
        "purpose": plan["purpose"],
        "diagnostic_only": True, "adoption_decision": "not evaluated",
        "previous_0787_rejection_remains_authoritative": True,
        "bootstrap": {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                       "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median",
                       "endpoint_indexes": BOOTSTRAP_INDEXES},
        "counts": counts,
        "oracle": oracle,
        "origin": {"base": origin_value["base"], "branch": origin_value.get("branch"),
                   "receipt": file_identity(ORIGIN_PATH)},
        "host": {"receipt": file_identity(PACKET / "host.json")},
        "architecture": architecture,
        "source": {"before": builds["before"]["source"], "after": builds["after"]["source"],
                    "candidate": candidate, "production_file_count": EXPECTED_PRODUCTION_FILES,
                    "before_matches_base": True, "after_candidate_diff_exact": True},
        "build": builds, "quality": quality,
        "native": {"records": native, "reports": len(native),
                    "samples": sum(row["samples"] for row in native),
                    "timings_or_rss_pooled_with_diagnostic": False},
        "qualification": {"before": qualification_before, "after": qualification_after},
        "memory": {"records": memory, "reports": len(memory),
                    "timings_pooled_with_native": False,
                    "analysis": None if memory_analysis is None else
                    {"path": memory_analysis["path"], "sha256": memory_analysis["sha256"]}},
        "heaptrack": {"records": heaptrack, "reports": len(heaptrack),
                       "samples": 480, "timings_or_rss_pooled_with_native": False,
                       "global_peak_is_not_native_rss": True},
        "paired": paired,
        "verification": {
            "plan_checked": True, "payload_oracle_checked": True,
            "report_schema_checked": True, "source_metrics_checked": True,
            "finite_cpu_task_and_permit_contract_checked": True,
            "source_and_candidate_custody_checked": True,
            "native_raw_distributions_retained": True,
            "bootstrap_endpoints_checked": True,
            "diagnostic_lanes_separated": True,
            "previous_0787_rejection_not_overturned": True,
            "production_restoration_required": True,
        },
    }
    return output, paired


PAIRED_FIELDS = ["route", "shape", "state", "task_floor", "workers", "blocks",
                 "samples_per_report", "diagnostic_only"]
for _metric in NATIVE_METRICS:
    PAIRED_FIELDS.extend([f"{_metric}_estimate", f"{_metric}_ci95_low",
                          f"{_metric}_ci95_high", f"{_metric}_difference_estimate",
                          f"{_metric}_block_ratios", f"{_metric}_block_deltas"])


def write_outputs(analysis: dict[str, Any], paired: list[dict[str, Any]]) -> None:
    outputs = {"analysis.json": json.dumps(analysis, indent=2, sort_keys=True) + "\n",
               "paired.csv": csv_text(csv_rows(paired)),
               "paired.md": render_markdown(analysis, paired)}
    for name, text in outputs.items():
        path = PACKET / name
        require(not path.exists(), f"refusing to overwrite retained {name}")
        path.write_text(text)


def csv_text(rows: list[dict[str, Any]]) -> str:
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=PAIRED_FIELDS, lineterminator="\n",
                            extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue()


def check_outputs(analysis: dict[str, Any], paired: list[dict[str, Any]]) -> None:
    require(read_json(PACKET / "analysis.json") == analysis,
            "analysis.json does not replay byte-for-byte")
    require((PACKET / "paired.csv").is_file()
            and (PACKET / "paired.csv").read_text() == csv_text(csv_rows(paired)),
            "paired.csv does not replay byte-for-byte")
    require((PACKET / "paired.md").is_file()
            and (PACKET / "paired.md").read_text() == render_markdown(analysis, paired),
            "paired.md does not replay byte-for-byte")


def analyze(*, write: bool = False, check: bool = False) -> dict[str, Any]:
    result, paired = build_analysis()
    if write:
        write_outputs(result, paired)
    if check:
        check_outputs(result, paired)
    return result


def main(argv: list[str] | None = None) -> int:
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--write", action="store_true", help="write deterministic outputs")
    modes.add_argument("--check", action="store_true", help="replay retained outputs")
    args = parser.parse_args(argv)
    try:
        analyze(write=args.write or not args.check, check=args.check)
    except ReplayError as error:
        print(f"0788 analysis failed: {error}", file=sys.stderr)
        return 1
    print("0788 analysis PASS: 220 reports, 5092 samples, 10 paired native cases")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
