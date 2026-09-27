#!/usr/bin/env python3
"""Replay the 0788 point-in-time RSS phase observations.

The executable used by this module deliberately does not read ``/proc``.  The
capture parent does that work at the acknowledgement points and stores the raw
files in one gzip JSON artifact per process.  This module is therefore an
offline parser: it verifies custody, parses every smaps mapping, records the
classification and counter disagreements, and derives paired phase deltas.

There are only two repeats in this diagnostic lane.  The output consequently
contains the two observations and ordinary medians where useful; it does not
bootstrap them and it never pools these observations with the native lane.
"""

from __future__ import annotations

import gzip
import hashlib
import json
import math
import os
import re
import statistics
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN = PACKET / "plan.json"
RECEIPTS = PACKET / "memory" / "receipts.json"
OUT_JSON = PACKET / "memory-analysis.json"
OUT_MD = PACKET / "memory-analysis.md"

SCHEMA = "litchi.cached-part-memory-analysis.v1"
PLAN_SCHEMA = "litchi.cached-part-memory.0788.v1"
REPORT_SCHEMA = "litchi.execution-baseline.v1"
RECEIPT_SCHEMA = "litchi.cached-part-memory-receipts.0788.v1"
MODES = ("off", "on")
LEGS = ("before", "after")
PROTOCOLS = ((1, 0), (30, 3))
ROUTES = ("parts",)
SHAPES = ("small", "large", "mixed")
STATES = ("fresh", "primed")
FLOORS = (0, 65536)
WORKERS = (4, 32)
CASES = (
    ("parts", "small", "fresh", 0, 4),
    ("parts", "small", "primed", 0, 4),
    ("parts", "small", "primed", 65536, 4),
    ("parts", "small", "primed", 0, 32),
)
EXPECTED_RECEIPTS_PER_MODE = 32
EXPECTED_RECEIPTS = EXPECTED_RECEIPTS_PER_MODE * 2
EXPECTED_SAMPLES_PER_MODE = 496
EXPECTED_SAMPLES = EXPECTED_SAMPLES_PER_MODE * 2
EXPECTED_SNAPSHOTS_PER_PROCESS = {1: 11, 30: 16}
SENTINEL = (1 << 64) - 1
COUNTERS = (
    "Rss",
    "Pss",
    "Anonymous",
    "Private_Clean",
    "Private_Dirty",
    "Shared_Clean",
    "Shared_Dirty",
    "Swap",
)
CATEGORIES = (
    "executable",
    "file-other",
    "sharedlibs",
    "heap",
    "main-thread-named-stack",
    "namedanon",
    "unnamedanon",
    "special",
)
PHASES = (
    "package_ready",
    "after_preload",
    "after_operation",
    "after_batch_drop",
    "after_package_drop",
)
CONTROL_PHASES = (
    "startup",
    "corpus_ready",
    "warmup_done",
    "samples_done",
    "report_written",
    "report_dropped",
)
SAMPLE_BYTES = {"small": 32 * 4096, "large": 32 * 262144,
                "mixed": 31 * 262144 + 4096}
SOURCE_METRICS = {
    "small": {"requested_bytes": 73590,
               "histogram": [32, 0, 0, 32, 0, 0, 0, 0]},
    "large": {"requested_bytes": 146041,
               "histogram": [32, 0, 0, 0, 32, 0, 0, 0]},
    "mixed": {"requested_bytes": 143782,
               "histogram": [32, 0, 0, 1, 31, 0, 0, 0]},
}
METRIC_KEYS = COUNTERS + ("Size",)

HEADER_RE = re.compile(
    r"^(?P<start>[0-9a-fA-F]+)-(?P<end>[0-9a-fA-F]+)\s+"
    r"(?P<perms>[-rwxps]{4})\s+(?P<offset>[0-9a-fA-F]+)\s+"
    r"(?P<dev>[0-9a-fA-F]+:[0-9a-fA-F]+)\s+(?P<inode>\d+)"
    r"(?:\s+(?P<path>.*?))?\s*$"
)
VALUE_RE = re.compile(r"^(?P<key>[A-Za-z][A-Za-z0-9_]*):\s+"
                      r"(?P<value>[0-9]+(?:\.[0-9]+)?)\s*(?P<unit>kB)?\s*$")
STATUS_VALUE_RE = re.compile(r"^(?P<key>[A-Za-z][A-Za-z0-9_]*):\s+"
                             r"(?P<value>[0-9]+)(?:\s+(?P<unit>\S+))?\s*$")


class ReplayError(RuntimeError):
    """Raised when retained evidence is missing or contradictory."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON {path}: {error}")


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
    return (isinstance(value, str) and len(value) == 64 and
            all(char in "0123456789abcdefABCDEF" for char in value))


def nonnegative_int(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")
    return value


def positive_int(value: Any, label: str) -> int:
    value = nonnegative_int(value, label)
    require(value > 0, f"{label} is not positive")
    return value


def finite_number(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")
    return float(value)


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def path_candidates(raw: Path) -> list[Path]:
    candidates: list[Path] = []
    if raw.is_absolute():
        text = raw.as_posix()
        marker = "/docs/performance/results/change-0788/"
        if marker in text:
            candidates.append(PACKET / text.split(marker, 1)[1])
        candidates.append(raw)
    else:
        text = raw.as_posix()
        prefix = "docs/performance/results/change-0788/"
        if text.startswith(prefix):
            candidates.append(PACKET / text[len(prefix):])
        candidates.extend((PACKET / raw, ROOT / raw))
    unique: list[Path] = []
    for candidate in candidates:
        candidate = candidate.resolve(strict=False)
        if candidate not in unique:
            unique.append(candidate)
    return unique


def resolve_path(raw: Any, *, packet_bound: bool = True) -> Path:
    require(isinstance(raw, str) and raw, f"invalid artifact path {raw!r}")
    candidates = path_candidates(Path(raw))
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            if not packet_bound:
                return candidate
            try:
                candidate.relative_to(PACKET.resolve())
            except ValueError:
                continue
            return candidate
    fallback = candidates[0] if candidates else Path(raw).resolve(strict=False)
    if packet_bound:
        try:
            fallback.relative_to(PACKET.resolve())
        except ValueError:
            fail(f"artifact escaped packet: {raw}")
    return fallback


def artifact(value: Any, label: str, *, packet_bound: bool = True,
             allow_missing: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an artifact receipt")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = resolve_path(raw, packet_bound=packet_bound)
    result = {"path": rel(path), "declared_path": raw, "bytes": size,
              "sha256": digest, "present": path.is_file()}
    if path.is_file():
        require(not path.is_symlink(), f"{label} is a symlink")
        require(path.stat().st_size == size, f"{label}.bytes changed")
        require(sha256(path) == digest, f"{label}.sha256 changed")
        result["verified"] = True
        result["path"] = rel(path)
        result["_path"] = path
        return result
    if allow_missing:
        result["verified"] = False
        return result
    fail(f"missing {label}: {raw}")


def artifact_path(value: Any, label: str, *, packet_bound: bool = True) -> Path:
    checked = artifact(value, label, packet_bound=packet_bound)
    path = checked.get("_path")
    require(isinstance(path, Path), f"{label} has no local path")
    return path


def _artifact_dicts(value: Any) -> Iterable[dict[str, Any]]:
    """Yield artifact-shaped dictionaries from a cleanup witness tree."""
    if isinstance(value, dict):
        if {"path", "bytes", "sha256"}.issubset(value):
            yield value
        for child in value.values():
            yield from _artifact_dicts(child)
    elif isinstance(value, list):
        for child in value:
            yield from _artifact_dicts(child)


def verify_binary_custody(value: Any, label: str) -> dict[str, Any]:
    """Verify a live binary or its exact retained cleanup witness.

    The returned representation intentionally uses the receipt's original
    path.  This keeps ``memory-analysis.json`` byte-stable before and after the
    owned target is removed, while the witness still proves the exact bytes
    that were checked before removal.
    """
    require(isinstance(value, dict), f"{label} is not an artifact receipt")
    raw_path = value.get("path")
    size = value.get("bytes", value.get("size"))
    digest = value.get("sha256", value.get("digest"))
    require(isinstance(raw_path, str) and raw_path,
            f"{label}.path is missing")
    nonnegative_int(size, f"{label}.bytes")
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = resolve_path(raw_path, packet_bound=False)
    if path.is_file():
        require(not path.is_symlink(), f"{label} is a symlink")
        require(path.stat().st_size == size, f"{label}.bytes changed")
        require(sha256(path) == digest, f"{label}.sha256 changed")
    else:
        witness_path = PACKET / "cleanup.json"
        require(witness_path.is_file() and not witness_path.is_symlink(),
                f"{label} is absent and cleanup witness is missing")
        witness = read_json(witness_path)
        require(isinstance(witness, dict) and witness.get("verified") is True and
                witness.get("target_absent_after_removal") is True,
                "cleanup witness does not prove verified target removal")
        matches = []
        for item in _artifact_dicts(witness):
            if (item.get("path") == raw_path and
                    item.get("bytes") == size and
                    item.get("sha256") == digest):
                matches.append(item)
        require(len(matches) == 1,
                f"{label} has no unique exact cleanup artifact witness")
    return {"path": raw_path, "bytes": size, "sha256": digest,
            "custody_verified": True}


def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema changed")
    memory = plan.get("memory")
    require(isinstance(memory, dict), "memory plan is missing")
    frozen_cases = [
        {"route": route, "shape": shape, "state": state,
         "task_floor": floor, "workers": workers}
        for route, shape, state, floor, workers in CASES
    ]
    require(memory.get("cases") == frozen_cases, "memory cases changed")
    require(memory.get("repeats") == 2, "memory repeat count changed")
    require(memory.get("sample_phases") == list(PHASES), "sample phases changed")
    require(memory.get("global_phases") == list(CONTROL_PHASES),
            "global phases changed")
    require(memory.get("proc_files") == ["smaps", "smaps_rollup", "maps",
                                           "status", "stat"],
            "proc file set changed")
    require(memory.get("snapshots_expected") == 432,
            "snapshot count contract changed")
    protocols = memory.get("protocols")
    require(protocols == [{"samples": 1, "warmup": 0},
                          {"samples": 30, "warmup": 3}],
            "memory protocols changed")
    require(plan.get("expected_counts", {}).get("memory_on_reports") == 32,
            "memory on count changed")
    require(plan.get("expected_counts", {}).get("memory_off_reports") == 32,
            "memory off count changed")
    return plan


def load_receipts() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    value = read_json(RECEIPTS)
    if isinstance(value, list):
        return {}, value
    require(isinstance(value, dict), "memory receipts are not an object/list")
    rows = value.get("receipts", value.get("processes", value.get("rows")))
    require(isinstance(rows, list), "memory receipts list is missing")
    return value, rows


def case_from_receipt(row: dict[str, Any]) -> dict[str, Any]:
    nested = row.get("case")
    case = nested if isinstance(nested, dict) else row
    result = {}
    for key in ("route", "shape", "state", "task_floor", "workers"):
        result[key] = case.get(key)
    require(tuple(result.get(key) for key in
                  ("route", "shape", "state", "task_floor", "workers")) in CASES,
            f"receipt case is not frozen: {result}")
    return result


def field(row: dict[str, Any], *names: str) -> Any:
    for name in names:
        if name in row:
            return row[name]
    return None


def normalize_receipt(row: dict[str, Any], index: int) -> dict[str, Any]:
    require(isinstance(row, dict), f"receipt {index} is not an object")
    case = case_from_receipt(row)
    repeat = field(row, "repeat", "iteration")
    repeat = nonnegative_int(repeat, f"receipt {index}.repeat")
    require(repeat < 2, f"receipt {index}.repeat is outside 0..1")
    leg = field(row, "leg", "build")
    require(leg in LEGS, f"receipt {index}.leg is invalid: {leg!r}")
    mode = field(row, "mode", "probe")
    if isinstance(mode, bool):
        mode = "on" if mode else "off"
    require(mode in MODES, f"receipt {index}.mode is invalid: {mode!r}")
    samples = nonnegative_int(field(row, "samples", "sample_count"),
                              f"receipt {index}.samples")
    warmup = nonnegative_int(field(row, "warmup", "warmups"),
                             f"receipt {index}.warmup")
    require((samples, warmup) in PROTOCOLS,
            f"receipt {index} protocol is not frozen")
    exit_code = field(row, "exit_code", "status")
    require(exit_code == 0, f"receipt {index} exited {exit_code!r}")
    binary = field(row, "binary", "executable")
    report = field(row, "report", "output")
    rss = field(row, "rss", "time", "rss_time")
    time_artifact = field(row, "time", "time_output")
    log = field(row, "log", "stderr")
    snapshots = field(row, "snapshots", "snapshot")
    require(binary is not None and report is not None and rss is not None and
            time_artifact is not None and log is not None,
            f"receipt {index} is missing required artifacts")
    return {
        "index": index,
        "route": case["route"], "shape": case["shape"],
        "state": case["state"], "task_floor": case["task_floor"],
        "workers": case["workers"], "repeat": repeat, "leg": leg,
        "mode": mode, "samples": samples, "warmup": warmup,
        "command": row.get("command"), "started": row.get("started"),
        "ended": row.get("ended"), "exit_code": exit_code,
        "binary_receipt": binary, "report_receipt": report,
        "rss_receipt": rss, "log_receipt": log,
        "time_receipt": time_artifact,
        "snapshots_receipt": snapshots,
        "raw": row,
    }


def open_text_artifact(value: Any, label: str) -> tuple[str, dict[str, Any]]:
    checked = artifact(value, label, packet_bound=True)
    path = checked.get("_path")
    require(isinstance(path, Path), f"{label} is unavailable")
    try:
        text = path.read_text()
    except OSError as error:
        fail(f"cannot read {label}: {error}")
    return text, {key: value for key, value in checked.items() if key != "_path"}


def read_ru_maxrss(value: Any, label: str) -> tuple[int, dict[str, Any]]:
    text, checked = open_text_artifact(value, label)
    stripped = text.strip()
    value_int: int | None = None
    counters: dict[str, int] = {}
    if stripped.startswith("{"):
        try:
            parsed = json.loads(stripped)
        except ValueError as error:
            fail(f"{label} JSON is invalid: {error}")
        require(isinstance(parsed, dict), f"{label} JSON is not an object")
        for key in ("minor_faults", "major_faults"):
            if key in parsed:
                counters[key] = nonnegative_int(parsed[key], f"{label}.{key}")
        for key in ("ru_maxrss_kib", "rss_kib", "max_rss_kib", "maxrss"):
            if key in parsed:
                value_int = nonnegative_int(parsed[key], f"{label}.{key}")
                break
    else:
        match = re.fullmatch(r"\s*([0-9]+)\s*", stripped)
        if match:
            value_int = int(match.group(1))
    require(value_int is not None, f"{label} is not a ru_maxrss KiB value")
    require(value_int > 0, f"{label} ru_maxrss is zero")
    checked.update(counters)
    return value_int, checked


def read_time_control(value: Any, label: str) -> tuple[dict[str, int], dict[str, Any]]:
    """Decode ``/usr/bin/time -f '%M %R %F'`` without inventing elapsed time.

    ``%R`` is minor faults and ``%F`` is major faults.  The memory lane's
    separate JSON ``rss`` receipt carries the same three counters under their
    explicit names; neither file contains an elapsed-time measurement.
    """
    text, checked = open_text_artifact(value, label)
    match = re.fullmatch(r"\s*([0-9]+)\s+([0-9]+)\s+([0-9]+)\s*", text)
    require(match is not None, f"{label} is not the frozen %M %R %F format")
    parsed = {
        "ru_maxrss_kib": int(match.group(1)),
        "minor_faults": int(match.group(2)),
        "major_faults": int(match.group(3)),
    }
    require(parsed["ru_maxrss_kib"] > 0,
            f"{label} ru_maxrss is zero")
    return parsed, checked


def report_path_for_receipt(row: dict[str, Any]) -> Path:
    # Reports are packet-bound and are retained after target/worktree cleanup.
    return artifact_path(row["report_receipt"],
                         f"receipt {row['index']} report")


def expected_histogram(shape: str, state: str) -> list[int]:
    if state == "primed":
        return [0] * 8
    return SOURCE_METRICS[shape]["histogram"]


def validate_source_metrics(metrics: Any, row: dict[str, Any], sample_label: str) -> dict[str, Any]:
    require(isinstance(metrics, dict), f"{sample_label}.source_metrics is not an object")
    require(metrics.get("availability") == "source-metrics-feature",
            f"{sample_label} source metrics feature is unavailable")
    for key in ("logical_calls", "requested_bytes", "returned_bytes",
                "short_reads", "active_reads_after_operation",
                "max_simultaneous_reads"):
        nonnegative_int(metrics.get(key), f"{sample_label}.source_metrics.{key}")
    histogram = metrics.get("request_size_histogram")
    require(isinstance(histogram, list) and len(histogram) == 8,
            f"{sample_label} source request histogram changed")
    for i, value in enumerate(histogram):
        nonnegative_int(value, f"{sample_label}.source_metrics.histogram[{i}]")
    expected_calls = 0 if row["state"] == "primed" else 64
    expected_bytes = 0 if row["state"] == "primed" else SOURCE_METRICS[row["shape"]]["requested_bytes"]
    require(metrics["logical_calls"] == expected_calls,
            f"{sample_label} source call count changed")
    require(metrics["requested_bytes"] == expected_bytes and
            metrics["returned_bytes"] == expected_bytes,
            f"{sample_label} source byte count changed")
    require(metrics["short_reads"] == 0 and
            metrics["active_reads_after_operation"] == 0,
            f"{sample_label} source read completion changed")
    require(metrics["max_simultaneous_reads"] == 0 if row["state"] == "primed"
            else metrics["max_simultaneous_reads"] >= 1,
            f"{sample_label} source concurrency is invalid")
    require(histogram == expected_histogram(row["shape"], row["state"]),
            f"{sample_label} source request histogram changed")
    require(sum(histogram) == metrics["logical_calls"],
            f"{sample_label} source histogram does not sum to calls")
    return {
        "availability": metrics["availability"],
        "logical_calls": metrics["logical_calls"],
        "requested_bytes": metrics["requested_bytes"],
        "returned_bytes": metrics["returned_bytes"],
        "short_reads": metrics["short_reads"],
        "active_reads_after_operation": metrics["active_reads_after_operation"],
        "max_simultaneous_reads": metrics["max_simultaneous_reads"],
        "request_size_histogram": list(histogram),
    }


def validate_report(path: Path, row: dict[str, Any]) -> dict[str, Any]:
    report = read_json(path)
    require(isinstance(report, dict) and report.get("schema") == REPORT_SCHEMA,
            f"receipt {row['index']} report schema changed")
    config = report.get("config")
    require(isinstance(config, dict), f"receipt {row['index']} report config missing")
    expected_config = {
        "route": row["route"], "shape": row["shape"],
        "workers": row["workers"], "task_floor": row["task_floor"],
        "state": row["state"], "samples": row["samples"],
        "warmup": row["warmup"],
    }
    for key, expected in expected_config.items():
        require(config.get(key) == expected,
                f"receipt {row['index']} report config {key} changed")
    corpus = report.get("corpus")
    require(isinstance(corpus, dict) and corpus.get("shape") == row["shape"],
            f"receipt {row['index']} corpus changed")
    members = corpus.get("members")
    require(isinstance(members, list) and len(members) == 32,
            f"receipt {row['index']} member count changed")
    for index, member in enumerate(members):
        require(isinstance(member, dict) and member.get("index") == index,
                f"receipt {row['index']} member order changed")
        nonnegative_int(member.get("bytes"),
                        f"receipt {row['index']} member {index}.bytes")
        require(is_sha(member.get("sha256")),
                f"receipt {row['index']} member {index}.sha256 missing")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == row["samples"],
            f"receipt {row['index']} report sample count changed")
    source_rows: list[dict[str, Any]] = []
    sample_rows: list[dict[str, Any]] = []
    for expected_index, sample in enumerate(samples):
        label = f"receipt {row['index']} sample {expected_index}"
        require(isinstance(sample, dict) and sample.get("sample") == expected_index,
                f"{label} index changed")
        wall_ns = nonnegative_int(sample.get("wall_ns"), f"{label}.wall_ns")
        cpu_ns = sample.get("cpu_ns")
        if cpu_ns is not None:
            nonnegative_int(cpu_ns, f"{label}.cpu_ns")
        verification = sample.get("verification")
        require(isinstance(verification, dict), f"{label}.verification missing")
        require(verification.get("ordered") is True and
                verification.get("all_member_sha256_match") is True and
                verification.get("members") == 32,
                f"{label} output verification failed")
        require(verification.get("logical_bytes") == SAMPLE_BYTES[row["shape"]],
                f"{label} logical byte count changed")
        require(is_sha(verification.get("sequence_sha256")),
                f"{label} sequence digest missing")
        resources = sample.get("resources")
        require(isinstance(resources, dict), f"{label}.resources missing")
        require(resources.get("worker_and_io_released") is True and
                resources.get("cpu_tasks_within_limit") is True,
                f"{label} resource release/limit check failed")
        for resource_key in ("before_operation", "after_operation", "after_drop"):
            snap = resources.get(resource_key)
            require(isinstance(snap, dict), f"{label}.{resource_key} missing")
            for key in ("workers", "io_concurrency", "cpu_tasks"):
                nonnegative_int(snap.get(key), f"{label}.{resource_key}.{key}")
        require(resources["after_drop"]["workers"] == 0 and
                resources["after_drop"]["io_concurrency"] == 0,
                f"{label} worker/IO resources were not released")
        source = validate_source_metrics(sample.get("source_metrics"), row, label)
        source_rows.append(source)
        sample_rows.append({
            "sample": expected_index, "wall_ns": wall_ns,
            "cpu_ns": cpu_ns, "verification": {
                "logical_bytes": verification["logical_bytes"],
                "sequence_sha256": verification["sequence_sha256"],
            }, "source_metrics": source,
            "resources": resources,
        })
    return {
        "schema": report["schema"], "config": config,
        "corpus": {
            "shape": corpus["shape"], "members": len(members),
            "logical_bytes": SAMPLE_BYTES[row["shape"]],
        }, "samples": sample_rows,
    }


def parse_header(line: str, label: str) -> dict[str, Any] | None:
    match = HEADER_RE.match(line.rstrip("\n"))
    if not match:
        return None
    start = int(match.group("start"), 16)
    end = int(match.group("end"), 16)
    require(end > start, f"{label} has an empty mapping")
    return {
        "start": start, "end": end, "perms": match.group("perms"),
        "offset": int(match.group("offset"), 16), "dev": match.group("dev"),
        "inode": int(match.group("inode")),
        "pathname": (match.group("path") or "").strip(),
    }


def mapping_key(mapping: dict[str, Any]) -> tuple[Any, ...]:
    return tuple(mapping[key] for key in
                 ("start", "end", "perms", "offset", "dev", "inode", "pathname"))


def parse_maps(raw: str, label: str) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for line_number, line in enumerate(raw.splitlines(), 1):
        if not line.strip():
            continue
        header = parse_header(line, f"{label}:{line_number}")
        require(header is not None, f"{label}:{line_number} is not a maps header")
        result.append(header)
    require(result, f"{label} is empty")
    require(len({mapping_key(item) for item in result}) == len(result),
            f"{label} has duplicate mappings")
    return result


def parse_kib_value(raw: str, label: str) -> int | float:
    match = VALUE_RE.match(raw.strip())
    require(match is not None, f"{label} is not a kB value")
    require(match.group("unit") in (None, "kB"), f"{label} has an unknown unit")
    value = float(match.group("value"))
    return int(value) if value.is_integer() else value


def parse_smaps(raw: str, label: str) -> list[dict[str, Any]]:
    mappings: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for line_number, line in enumerate(raw.splitlines(), 1):
        header = parse_header(line, f"{label}:{line_number}")
        if header is not None:
            if current is not None:
                mappings.append(current)
            current = header
            current["values"] = {}
            continue
        require(current is not None, f"{label}:{line_number} precedes first mapping")
        if line.strip() == "":
            continue
        match = VALUE_RE.match(line.strip())
        if match is None:
            # VmFlags and future textual fields are not counters.  Retain no
            # interpretation of them, but do not silently accept a malformed
            # numeric field.
            if line.lstrip().startswith("VmFlags:"):
                continue
            continue
        key = match.group("key")
        current["values"][key] = parse_kib_value(line, f"{label}:{line_number}")
    if current is not None:
        mappings.append(current)
    require(mappings, f"{label} has no mappings")
    for index, mapping in enumerate(mappings):
        values = mapping["values"]
        for key in COUNTERS:
            value = values.get(key)
            finite_number(value, f"{label} mapping {index}.{key}")
            require(value >= 0, f"{label} mapping {index}.{key} is negative")
        finite_number(values.get("Size"), f"{label} mapping {index}.Size")
    require(len({mapping_key(item) for item in mappings}) == len(mappings),
            f"{label} has duplicate mappings")
    return mappings


def parse_rollup(raw: str, label: str) -> dict[str, int | float]:
    values: dict[str, int | float] = {}
    lines = raw.splitlines()
    require(lines, f"{label} is empty")
    # smaps_rollup begins with a synthetic mapping header.  It is not a
    # counter line, but validating it prevents a maps/smaps file from being
    # accidentally accepted as rollup evidence.
    header = parse_header(lines[0], f"{label}:1")
    require(header is not None and header["pathname"] == "[rollup]" and
            header["perms"] == "---p" and header["offset"] == 0 and
            header["dev"] == "00:00" and header["inode"] == 0,
            f"{label}:1 is not the expected [rollup] header")
    for line_number, line in enumerate(lines[1:], 2):
        if not line.strip():
            continue
        match = VALUE_RE.match(line.strip())
        require(match is not None, f"{label}:{line_number} is malformed")
        key = match.group("key")
        values[key] = parse_kib_value(line, f"{label}:{line_number}")
    for key in COUNTERS:
        finite_number(values.get(key), f"{label}.{key}")
    return values


def parse_status(raw: str, label: str) -> dict[str, Any]:
    values: dict[str, Any] = {}
    for line_number, line in enumerate(raw.splitlines(), 1):
        match = STATUS_VALUE_RE.match(line.strip())
        if match:
            value = int(match.group("value"))
            unit = match.group("unit")
            values[match.group("key")] = {"value": value, "unit": unit}
        elif line.startswith(("Name:", "State:")):
            key, value = line.split(":", 1)
            values[key] = value.strip()
    require("Pid" in values and "PPid" in values and "Threads" in values,
            f"{label} lacks Pid/PPid/Threads")
    return values


def parse_stat(raw: str, label: str) -> dict[str, Any]:
    line = raw.strip()
    match = re.match(r"^(\d+)\s+\((.*)\)\s+(.*)$", line)
    require(match is not None, f"{label} is malformed")
    pid = int(match.group(1))
    rest = match.group(3).split()
    require(len(rest) >= 18, f"{label} has too few fields")
    try:
        ppid = int(rest[1])
        threads = int(rest[17])
    except ValueError as error:
        fail(f"{label} numeric fields are malformed: {error}")
    return {"pid": pid, "comm": match.group(2), "state": rest[0],
            "ppid": ppid, "threads": threads}


def basename_without_deleted(path: str) -> str:
    value = path.strip()
    if value.endswith(" (deleted)"):
        value = value[:-10]
    return Path(value).name


def classify_mapping(mapping: dict[str, Any], exe: str) -> str:
    pathname = mapping["pathname"]
    if pathname == "[heap]":
        return "heap"
    if pathname == "[stack]" or pathname.startswith("[stack:"):
        return "main-thread-named-stack"
    if pathname.startswith("[anon:") or pathname.startswith("[anon_shmem:"):
        return "namedanon"
    if pathname.startswith("[") and pathname.endswith("]"):
        return "special"
    if not pathname:
        return "unnamedanon"
    if basename_without_deleted(pathname) == basename_without_deleted(exe):
        return "executable"
    lower = pathname.lower()
    if (".so" in lower or "/ld-" in lower or "/libc-" in lower or
            "/lib/" in lower or "/lib64/" in lower):
        return "sharedlibs"
    return "file-other"


def sum_metrics(mappings: Iterable[dict[str, Any]]) -> dict[str, int | float]:
    result = {key: 0 for key in COUNTERS}
    for mapping in mappings:
        values = mapping["values"]
        for key in COUNTERS:
            result[key] += values[key]
    return result


def classify_mappings(mappings: list[dict[str, Any]], exe: str,
                      rollup: dict[str, int | float], status: dict[str, Any],
                      maps: list[dict[str, Any]], label: str) -> dict[str, Any]:
    require({mapping_key(item) for item in mappings} ==
            {mapping_key(item) for item in maps}, f"{label} maps/smaps differ")
    totals = sum_metrics(mappings)
    groups: dict[str, dict[str, Any]] = {}
    full: list[dict[str, Any]] = []
    by_category: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for mapping in mappings:
        category = classify_mapping(mapping, exe)
        item = {
            "start": mapping["start"], "end": mapping["end"],
            "bytes": mapping["end"] - mapping["start"],
            "perms": mapping["perms"], "offset": mapping["offset"],
            "dev": mapping["dev"], "inode": mapping["inode"],
            "pathname": mapping["pathname"], "category": category,
            "counters_kib": {key: mapping["values"][key] for key in COUNTERS},
            "size_kib": mapping["values"]["Size"],
        }
        full.append(item)
        by_category[category].append(mapping)
    for category in CATEGORIES:
        members = by_category.get(category, [])
        groups[category] = {
            "mapping_count": len(members),
            "counters_kib": sum_metrics(members),
        }
    disagreements: list[dict[str, Any]] = []
    rollup_delta: dict[str, int | float] = {}
    for key in COUNTERS:
        delta = rollup[key] - totals[key]
        rollup_delta[key] = delta
        if key == "Pss":
            tolerance = max(1, len(mappings))
            ok = abs(delta) <= tolerance
            reason = "per-mapping PSS kernel rounding"
        elif key == "Rss":
            tolerance = 0
            ok = delta == 0
            reason = "quiescent RSS should match exactly"
        else:
            tolerance = 0
            ok = delta == 0
            reason = "smaps/rollup counter should match"
        if not ok:
            disagreements.append({"counter": key, "rollup": rollup[key],
                                  "per_mapping": totals[key], "delta": delta,
                                  "tolerance_kib": tolerance, "reason": reason})
    status_rss = status.get("VmRSS")
    status_disagreement = None
    if isinstance(status_rss, dict) and status_rss.get("unit") == "kB":
        delta = status_rss["value"] - rollup["Rss"]
        if delta != 0:
            status_disagreement = {
                "status_vmrss_kib": status_rss["value"],
                "smaps_rollup_rss_kib": rollup["Rss"], "delta_kib": delta,
                "interpretation": "proc status VmRSS is retained as a separate, asynchronous counter",
            }
    return {
        "mapping_count": len(mappings), "mappings": full,
        "categories": groups, "sum_kib": totals, "rollup_kib": rollup,
        "rollup_delta_kib": rollup_delta,
        "counter_disagreements": disagreements,
        "status_rss_disagreement": status_disagreement,
        "status_vmrss_kib": (status.get("VmRSS", {}).get("value")
                              if isinstance(status.get("VmRSS"), dict) else None),
        "status_vmhwm_kib": (status.get("VmHWM", {}).get("value")
                              if isinstance(status.get("VmHWM"), dict) else None),
        "status_threads": status["Threads"]["value"],
        "stat_threads": None,
    }


def normalize_task_ids(value: Any, label: str) -> list[int]:
    if isinstance(value, dict):
        for key in ("task_ids", "ids", "tids", "threads"):
            if key in value:
                value = value[key]
                break
    require(isinstance(value, list) and value, f"{label} task_ids missing")
    result = []
    for index, task in enumerate(value):
        if isinstance(task, dict):
            task = task.get("tid", task.get("pid", task.get("id")))
        result.append(positive_int(task, f"{label}.task_ids[{index}]"))
    return result


def expected_phase_sequence(samples: int) -> list[tuple[str, int | None]]:
    result: list[tuple[str, int | None]] = [
        ("startup", None), ("corpus_ready", None), ("warmup_done", None),
    ]
    indexes = [0] if samples == 1 else [0, samples - 1]
    for sample in indexes:
        result.extend((phase, sample) for phase in PHASES)
    result.extend((phase, None) for phase in ("samples_done", "report_written",
                                               "report_dropped"))
    return result


def normalize_phase(value: Any) -> str:
    require(isinstance(value, str) and value, "snapshot phase is missing")
    if value.startswith("parts."):
        value = value[6:]
    require(value in CONTROL_PHASES + PHASES, f"unknown snapshot phase {value!r}")
    return value


def normalize_sample(value: Any, phase: str) -> int | None:
    if value is None or value == SENTINEL or value == str(SENTINEL):
        result = None
    else:
        result = nonnegative_int(value, f"snapshot {phase}.sample")
    if phase in CONTROL_PHASES:
        require(result is None, f"control phase {phase} has a sample")
    else:
        require(result is not None, f"sample phase {phase} lacks a sample")
    return result


def parse_snapshot(snapshot: dict[str, Any], index: int, row: dict[str, Any],
                   binary_name: str) -> dict[str, Any]:
    require(isinstance(snapshot, dict), f"receipt {row['index']} snapshot {index} is not an object")
    phase = normalize_phase(snapshot.get("phase"))
    sample = normalize_sample(snapshot.get("sample"), phase)
    sequence = snapshot.get("sequence", snapshot.get("index", index))
    require(sequence == index, f"receipt {row['index']} snapshot sequence changed")
    pid = positive_int(snapshot.get("pid"), f"receipt {row['index']} snapshot pid")
    exe = snapshot.get("exe", snapshot.get("executable"))
    require(isinstance(exe, str) and exe, f"receipt {row['index']} snapshot exe missing")
    require(basename_without_deleted(exe) == basename_without_deleted(binary_name),
            f"receipt {row['index']} snapshot exe does not match binary")
    ppid = positive_int(snapshot.get("ppid"), f"receipt {row['index']} snapshot ppid")
    starttime = positive_int(snapshot.get("starttime"),
                             f"receipt {row['index']} snapshot starttime")
    task_ids = normalize_task_ids(snapshot.get("task_ids"),
                                  f"receipt {row['index']} snapshot {index}")
    require(task_ids == [pid],
            f"receipt {row['index']} snapshot {index} has unjoined worker tasks")
    raw = snapshot.get("raw")
    require(isinstance(raw, dict), f"receipt {row['index']} snapshot raw files missing")
    for key in ("smaps", "smaps_rollup", "maps", "status", "stat"):
        require(isinstance(raw.get(key), str),
                f"receipt {row['index']} snapshot raw.{key} is not text")
    maps = parse_maps(raw["maps"], f"receipt {row['index']} snapshot {index}.maps")
    smaps = parse_smaps(raw["smaps"], f"receipt {row['index']} snapshot {index}.smaps")
    rollup = parse_rollup(raw["smaps_rollup"],
                          f"receipt {row['index']} snapshot {index}.smaps_rollup")
    status = parse_status(raw["status"],
                          f"receipt {row['index']} snapshot {index}.status")
    stat = parse_stat(raw["stat"],
                      f"receipt {row['index']} snapshot {index}.stat")
    require(status["Pid"]["value"] == pid and stat["pid"] == pid,
            f"receipt {row['index']} snapshot {index} PID changed")
    require(status["PPid"]["value"] == ppid and stat["ppid"] == ppid,
            f"receipt {row['index']} snapshot {index} PPID changed")
    require(status["Threads"]["value"] == 1 and stat["threads"] == 1,
            f"receipt {row['index']} snapshot {index} still has worker threads")
    classification = classify_mappings(smaps, exe, rollup, status, maps,
                                       f"receipt {row['index']} snapshot {index}")
    classification["stat_threads"] = stat["threads"]
    return {
        "sequence": index, "phase": phase, "sample": sample, "pid": pid,
        "ppid": ppid, "starttime": starttime, "exe": exe, "task_ids": task_ids,
        "classification": classification,
    }


def snapshot_path_for_receipt(row: dict[str, Any]) -> Path:
    require(row["mode"] == "on", f"receipt {row['index']} has no on snapshot")
    value = row["snapshots_receipt"]
    require(value is not None, f"receipt {row['index']} snapshot artifact missing")
    return artifact_path(value, f"receipt {row['index']} snapshots")


def read_snapshots(row: dict[str, Any], binary_name: str) -> list[dict[str, Any]]:
    path = snapshot_path_for_receipt(row)
    try:
        with gzip.open(path, "rt", encoding="utf-8") as stream:
            value = json.load(stream)
    except (OSError, ValueError) as error:
        fail(f"receipt {row['index']} snapshots gzip/JSON invalid: {error}")
    require(isinstance(value, list), f"receipt {row['index']} snapshots is not a list")
    expected_count = EXPECTED_SNAPSHOTS_PER_PROCESS[row["samples"]]
    require(len(value) == expected_count,
            f"receipt {row['index']} snapshot count is {len(value)}, expected {expected_count}")
    result = [parse_snapshot(item, index, row, binary_name)
              for index, item in enumerate(value)]
    expected = expected_phase_sequence(row["samples"])
    actual = [(item["phase"], item["sample"]) for item in result]
    require(actual == expected, f"receipt {row['index']} phase sequence changed")
    identities = {(item["pid"], item["ppid"], item["starttime"], item["exe"])
                  for item in result}
    require(len(identities) == 1,
            f"receipt {row['index']} snapshot process identity changed")
    return result


def counters_for(snapshot: dict[str, Any], category: str | None = None) -> dict[str, int | float]:
    classification = snapshot["classification"]
    if category is None:
        return dict(classification["sum_kib"])
    return dict(classification["categories"][category]["counters_kib"])


def phase_values(snapshots: list[dict[str, Any]], samples: int) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for snapshot in snapshots:
        phase = snapshot["phase"]
        if phase not in PHASES:
            continue
        sample = str(snapshot["sample"])
        result.setdefault(sample, {})[phase] = counters_for(snapshot)
    return result


def subtract(left: dict[str, Any], right: dict[str, Any]) -> dict[str, Any]:
    return {key: left[key] - right[key] for key in left}


def operation_deltas(values: dict[str, Any]) -> list[dict[str, Any]]:
    result = []
    for sample, phases in sorted(values.items(), key=lambda item: int(item[0])):
        for before, after in zip(PHASES, PHASES[1:]):
            if before not in phases or after not in phases:
                continue
            result.append({"sample": int(sample), "from": before, "to": after,
                           "delta_kib": subtract(phases[after], phases[before])})
    return result


def verify_receipt(row: dict[str, Any]) -> dict[str, Any]:
    binary_checked = verify_binary_custody(
        row["binary_receipt"], f"receipt {row['index']} binary")
    binary_name = Path(str(row["binary_receipt"].get("path", ""))).name
    require(binary_name, f"receipt {row['index']} binary name missing")
    report_path = report_path_for_receipt(row)
    report = validate_report(report_path, row)
    rss, rss_artifact = read_ru_maxrss(row["rss_receipt"],
                                       f"receipt {row['index']} RSS/time")
    time_values, time_artifact = read_time_control(
        row["time_receipt"], f"receipt {row['index']} /usr/bin/time")
    require(rss == time_values["ru_maxrss_kib"],
            f"receipt {row['index']} rss and %M disagree")
    for key in ("minor_faults", "major_faults"):
        require(rss_artifact.get(key) == time_values[key],
                f"receipt {row['index']} rss and %{ 'R' if key == 'minor_faults' else 'F' } disagree")
    _, log_artifact = open_text_artifact(row["log_receipt"],
                                         f"receipt {row['index']} log")
    snapshots: list[dict[str, Any]] = []
    if row["mode"] == "on":
        snapshots = read_snapshots(row, binary_name)
    else:
        require(row["snapshots_receipt"] is None,
                f"receipt {row['index']} off control has snapshots")
    values = phase_values(snapshots, row["samples"]) if snapshots else {}
    disagreements = []
    for snapshot in snapshots:
        disagreements.extend({"sequence": snapshot["sequence"], **item}
                              for item in snapshot["classification"]["counter_disagreements"])
        status = snapshot["classification"].get("status_rss_disagreement")
        if status is not None:
            disagreements.append({"sequence": snapshot["sequence"], **status})
    sample_summary = []
    for sample in report["samples"]:
        sample_summary.append({
            "sample": sample["sample"], "wall_ns": sample["wall_ns"],
            "cpu_ns": sample["cpu_ns"],
            "source_metrics": sample["source_metrics"],
            "verification": sample["verification"],
        })
    return {
        "index": row["index"], "case": {key: row[key] for key in
                                            ("route", "shape", "state", "task_floor", "workers")},
        "repeat": row["repeat"], "leg": row["leg"], "mode": row["mode"],
        "samples": row["samples"], "warmup": row["warmup"],
        "command": row["command"], "started": row["started"],
        "ended": row["ended"], "exit_code": row["exit_code"],
        "binary": binary_checked,
        "report": {"path": rel(report_path), "bytes": report_path.stat().st_size,
                   "sha256": sha256(report_path)},
        "rss_time": {**rss_artifact, **time_artifact,
                     "ru_maxrss_kib": rss,
                     "minor_faults": time_values["minor_faults"],
                     "major_faults": time_values["major_faults"]},
        "log": log_artifact,
        "report_summary": {"samples": sample_summary,
                            "corpus": report["corpus"]},
        "snapshots": snapshots,
        "phase_values_kib": values,
        "operation_drop_deltas_kib": operation_deltas(values),
        "counter_disagreements": disagreements,
        "snapshot_count": len(snapshots),
    }


def key_for(row: dict[str, Any]) -> tuple[Any, ...]:
    return (row["route"], row["shape"], row["state"], row["task_floor"],
            row["workers"], row["repeat"], row["leg"], row["samples"],
            row["warmup"])


def receipt_key_from_verified(row: dict[str, Any]) -> tuple[Any, ...]:
    case = row["case"]
    return (case["route"], case["shape"], case["state"], case["task_floor"],
            case["workers"], row["repeat"], row["leg"], row["samples"],
            row["warmup"])


def paired_deltas(rows: list[dict[str, Any]]) -> dict[str, Any]:
    by_key = {receipt_key_from_verified(row) + (row["mode"],): row for row in rows}
    mode_rows = []
    leg_rows = []
    operation_rows = []
    for key in sorted({receipt_key_from_verified(row) for row in rows}):
        for mode in MODES:
            before_key = key[:6] + ("before",) + key[7:] + (mode,)
            after_key = key[:6] + ("after",) + key[7:] + (mode,)
            before = by_key.get(before_key)
            after = by_key.get(after_key)
            if before and after and before["snapshots"] and after["snapshots"]:
                for sample, phases in before["phase_values_kib"].items():
                    after_phases = after["phase_values_kib"].get(sample, {})
                    for phase in PHASES:
                        if phase in phases and phase in after_phases:
                            leg_rows.append({
                                "case": before["case"], "repeat": before["repeat"],
                                "samples": before["samples"], "warmup": before["warmup"],
                                "mode": mode, "sample": int(sample), "phase": phase,
                                "delta_kib_after_minus_before": subtract(after_phases[phase], phases[phase]),
                            })
                    for delta in before["operation_drop_deltas_kib"]:
                        match = next((item for item in after["operation_drop_deltas_kib"]
                                      if item["sample"] == delta["sample"] and
                                      item["from"] == delta["from"] and item["to"] == delta["to"]), None)
                        if match:
                            operation_rows.append({
                                "case": before["case"], "repeat": before["repeat"],
                                "samples": before["samples"], "warmup": before["warmup"],
                                "mode": mode, "sample": delta["sample"],
                                "from": delta["from"], "to": delta["to"],
                                "delta_kib_after_minus_before": subtract(match["delta_kib"], delta["delta_kib"]),
                            })
        for phase in PHASES:
            off_key = key[:6] + (key[6],) + key[7:] + ("off",)
            on_key = key[:6] + (key[6],) + key[7:] + ("on",)
            off = by_key.get(off_key)
            on = by_key.get(on_key)
            # The key includes leg, so this loop yields one handshake pair per
            # leg.  The off lane intentionally has no phase snapshots; its
            # ru_maxrss is still retained as the separate control.
            if off and on:
                mode_rows.append({
                    "case": off["case"], "repeat": off["repeat"],
                    "leg": off["leg"], "samples": off["samples"],
                    "warmup": off["warmup"],
                    "ru_maxrss_kib_on_minus_off": on["rss_time"]["ru_maxrss_kib"] - off["rss_time"]["ru_maxrss_kib"],
                    "ru_maxrss_kib_on": on["rss_time"]["ru_maxrss_kib"],
                    "ru_maxrss_kib_off": off["rss_time"]["ru_maxrss_kib"],
                    "phase_observation": "off control has no smaps phases",
                })
                break
    # The rows above deliberately keep each of the two repeats.  A compact
    # repeat summary is descriptive only and does not claim a confidence band.
    return {"mode_on_minus_off_process_controls": mode_rows,
            "leg_after_minus_before_phase_deltas": leg_rows,
            "leg_after_minus_before_operation_drop_deltas": operation_rows}


def validate_cardinality(rows: list[dict[str, Any]]) -> None:
    require(len(rows) == EXPECTED_RECEIPTS,
            f"memory receipt count is {len(rows)}, expected {EXPECTED_RECEIPTS}")
    by_mode = defaultdict(list)
    for row in rows:
        by_mode[row["mode"]].append(row)
    for mode in MODES:
        require(len(by_mode[mode]) == EXPECTED_RECEIPTS_PER_MODE,
                f"{mode} receipt count changed")
        for case in CASES:
            for repeat in range(2):
                for leg in LEGS:
                    matches = [item for item in by_mode[mode]
                               if tuple(item[key] for key in
                                        ("route", "shape", "state", "task_floor", "workers")) == case
                               and item["repeat"] == repeat and item["leg"] == leg]
                    require(len(matches) == 2,
                            f"{mode} case/repeat/leg protocol cardinality changed: {case} {repeat} {leg}")
    for mode in MODES:
        samples = sum(row["samples"] for row in by_mode[mode])
        require(samples == EXPECTED_SAMPLES_PER_MODE,
                f"{mode} measured sample count is {samples}")


def high_water_summary(rows: list[dict[str, Any]],
                       comparisons: dict[str, Any]) -> dict[str, Any]:
    """Keep GNU time high-water and acknowledged RSS observations separate.

    ``ru_maxrss`` is a process high-water counter.  The largest smaps-rollup
    value and largest status VmHWM are each derived from the retained phase
    observations and are not interchangeable with it.  Keeping one row per
    process makes a later report able to show the raw differences rather than
    silently treating one counter as a proxy for another.
    """
    process_rows: list[dict[str, Any]] = []
    for row in rows:
        if row["mode"] != "on":
            continue
        snapshots = row["snapshots"]
        smaps_values = [snapshot["classification"]["rollup_kib"]["Rss"]
                        for snapshot in snapshots]
        vmhwm_values = [snapshot["classification"]["status_vmhwm_kib"]
                        for snapshot in snapshots
                        if snapshot["classification"]["status_vmhwm_kib"] is not None]
        vmrss_values = [snapshot["classification"]["status_vmrss_kib"]
                        for snapshot in snapshots
                        if snapshot["classification"]["status_vmrss_kib"] is not None]
        require(smaps_values and vmhwm_values,
                f"receipt {row['index']} lacks RSS/VmHWM observations")
        time_rss = row["rss_time"]["ru_maxrss_kib"]
        max_smaps = max(smaps_values)
        max_vmhwm = max(vmhwm_values)
        process_rows.append({
            "receipt": row["index"], "case": row["case"],
            "repeat": row["repeat"], "leg": row["leg"],
            "samples": row["samples"], "warmup": row["warmup"],
            "ru_maxrss_kib": time_rss,
            "max_observed_smaps_rollup_rss_kib": max_smaps,
            "max_observed_status_vmhwm_kib": max_vmhwm,
            "max_observed_status_vmrss_kib": max(vmrss_values) if vmrss_values else None,
            "smaps_minus_time_kib": max_smaps - time_rss,
            "vmhwm_minus_time_kib": max_vmhwm - time_rss,
            "time_below_max_observed_smaps_rollup_rss": time_rss < max_smaps,
            "time_below_max_observed_status_vmhwm": time_rss < max_vmhwm,
            "phase_observation_count": len(smaps_values),
        })

    def summary(rows_for_metric: list[dict[str, Any]], key: str) -> dict[str, Any]:
        values = [row[key] for row in rows_for_metric]
        return {
            "count": len(values), "values_kib": values,
            "min_kib": min(values) if values else None,
            "max_kib": max(values) if values else None,
            "max_abs_kib": max((abs(value) for value in values), default=None),
            "zero_count": sum(value == 0 for value in values),
        }

    phase_rss: dict[str, list[int]] = defaultdict(list)
    for item in comparisons["leg_after_minus_before_phase_deltas"]:
        phase_rss[item["phase"]].append(item["delta_kib_after_minus_before"]["Rss"])
    operation_rss: list[int] = [
        item["delta_kib"]["Rss"]
        for row in rows if row["mode"] == "on"
        for item in row["operation_drop_deltas_kib"]
        if item["from"] == "after_preload" and item["to"] == "after_operation"
    ]
    return {
        "scope": "On-mode diagnostic children only; each max is over retained acknowledged snapshots in that child",
        "interpretation": "GNU time ru_maxrss is intended as a process high-water counter. When it is below an observed same-child smaps_rollup RSS or status VmHWM, that is a counter disagreement requiring investigation; it is not treated as noise or used to overturn the rejected 0787 decision, and it provides no causal attribution.",
        "rows": process_rows,
        "counts": {
            "on_processes": len(process_rows),
            "time_below_max_observed_smaps_rollup_rss": sum(
                row["time_below_max_observed_smaps_rollup_rss"] for row in process_rows),
            "time_below_max_observed_status_vmhwm": sum(
                row["time_below_max_observed_status_vmhwm"] for row in process_rows),
        },
        "ru_maxrss_vs_observed_smaps_rss": summary(process_rows,
                                                    "smaps_minus_time_kib"),
        "ru_maxrss_vs_observed_status_vmhwm": summary(process_rows,
                                                        "vmhwm_minus_time_kib"),
        "operation_after_preload_to_after_operation_rss_deltas_kib": {
            "count": len(operation_rss), "values": operation_rss,
            "zero_count": sum(value == 0 for value in operation_rss),
            "max_abs": max((abs(value) for value in operation_rss), default=None),
        },
        "leg_after_minus_before_phase_rss_deltas_kib": {
            phase: {"count": len(values), "values": values,
                    "min": min(values) if values else None,
                    "max": max(values) if values else None,
                    "max_abs": max((abs(value) for value in values), default=None),
                    "zero_count": sum(value == 0 for value in values)}
            for phase, values in sorted(phase_rss.items())
        },
    }


def memory_analysis() -> dict[str, Any]:
    plan = load_plan()
    receipt_meta, raw_rows = load_receipts()
    normalized = [normalize_receipt(row, index) for index, row in enumerate(raw_rows)]
    validate_cardinality(normalized)
    verified = [verify_receipt(row) for row in normalized]
    verified.sort(key=lambda row: (row["case"]["shape"], row["case"]["state"],
                                   row["case"]["task_floor"], row["case"]["workers"],
                                   row["samples"], row["warmup"], row["repeat"],
                                   LEGS.index(row["leg"]), MODES.index(row["mode"])))
    disagreements = [
        {"receipt": row["index"], "case": row["case"], "repeat": row["repeat"],
         "leg": row["leg"], "mode": row["mode"], "items": row["counter_disagreements"]}
        for row in verified if row["counter_disagreements"]
    ]
    comparisons = paired_deltas(verified)
    high_water = high_water_summary(verified, comparisons)
    return {
        "schema": SCHEMA,
        "scope": {
            "claim": "Point-in-time smaps observations at acknowledged lifecycle phases",
            "rss": "smaps and smaps_rollup counters are observations, not operation peaks",
            "ru_maxrss": "probe-off and probe-on /usr/bin/time controls are separate high-water counters",
            "anonymous": "Only explicit [heap], named anonymous mappings, unnamed anonymous mappings and special mappings are classified; generic anonymous bytes are never called heap or TLS",
            "statistics": "Two repeats are retained as paired observations; no bootstrap and no native-lane pooling",
            "handshake": "The enabled diagnostic handshake can allocate or fault pages and is compared with a disabled control",
        },
        "plan": {"path": rel(PLAN), "sha256": sha256(PLAN),
                 "schema": plan["schema"]},
        "receipts": {
            "path": rel(RECEIPTS), "sha256": sha256(RECEIPTS),
            "schema": receipt_meta.get("schema"), "count": len(verified),
            "samples": sum(row["samples"] for row in verified),
        },
        "counts": {
            "receipts": len(verified), "mode_off_receipts": sum(row["mode"] == "off" for row in verified),
            "mode_on_receipts": sum(row["mode"] == "on" for row in verified),
            "samples": sum(row["samples"] for row in verified),
            "snapshots": sum(row["snapshot_count"] for row in verified),
            "counter_disagreement_receipts": len(disagreements),
        },
        "processes": verified,
        "counter_disagreements": disagreements,
        "comparisons": comparisons,
        "high_water_summary": high_water,
        "limitations": [
            "smaps is a point-in-time page-table scan; it cannot establish an operation peak",
            "PSS comparison permits one KiB per mapping for kernel/per-mapping rounding",
            "status VmRSS is retained separately and is not required to equal smaps_rollup RSS",
            "task IDs and Threads=1 establish that worker tasks had joined at each acknowledged point; they do not identify allocator ownership",
            "file-backed, anonymous and stack labels are mapping classifications, not causal attribution",
            "the disabled control has no phase snapshots; its ru_maxrss is not substituted for a phase series",
        ],
    }


def markdown(result: dict[str, Any]) -> str:
    counts = result["counts"]
    lines = [
        "# 0788 memory phase analysis",
        "",
        "This is an offline replay of acknowledged `/proc` snapshots. It records",
        "point-in-time smaps classifications and paired deltas for the diagnostic",
        "processes. It does not revise the 0787 decision, pool native timings, or",
        "call a phase value an operation peak.",
        "",
        f"- Receipts: {counts['receipts']} ({counts['mode_off_receipts']} off, {counts['mode_on_receipts']} on)",
        f"- Report samples: {counts['samples']}",
        f"- On-mode snapshots: {counts['snapshots']}",
        f"- Receipts with counter disagreements: {counts['counter_disagreement_receipts']}",
        "",
        "## Counter checks",
        "",
        "Per-mapping sums are compared with `smaps_rollup`. RSS is required to",
        "match exactly at the retained snapshot; PSS allows the documented",
        "per-mapping kernel rounding tolerance. `status` VmRSS is retained as a",
        "separate asynchronous counter.",
        "",
    ]
    if result["counter_disagreements"]:
        lines += ["| receipt | leg | mode | phase sequence | disagreements |",
                  "| ---: | --- | --- | ---: | --- |"]
        for item in result["counter_disagreements"]:
            lines.append(f"| {item['receipt']} | {item['leg']} | {item['mode']} | "
                         f"{len(item['items'])} | `{json.dumps(item['items'], sort_keys=True)}` |")
    else:
        lines += ["No smaps-versus-rollup counter disagreement exceeded the recorded tolerance.", ""]
    high_water = result["high_water_summary"]
    lines += [
        "## GNU time versus acknowledged RSS",
        "",
        "The diagnostic children retain GNU `/usr/bin/time` `ru_maxrss` separately",
        "from the largest acknowledged `smaps_rollup` RSS and status VmHWM. The",
        f"time high-water value was below the retained smaps RSS in "
        f"{high_water['counts']['time_below_max_observed_smaps_rollup_rss']} of "
        f"{high_water['counts']['on_processes']} on-mode children and below VmHWM in "
        f"{high_water['counts']['time_below_max_observed_status_vmhwm']} children.",
        "A high-water counter below a same-child observed RSS is a counter",
        "disagreement requiring investigation. It does not provide causal",
        "attribution or reinterpret the rejected 0787 result.",
        "",
        "The complete per-process raw differences and phase deltas are retained",
        "in `high_water_summary` in `memory-analysis.json`.",
        "",
    ]
    lines += [
        "## Scope limits",
        "",
        "Anonymous mappings are classified by their actual mapping labels. The",
        "parser never infers heap or TLS from generic anonymous bytes. The enabled",
        "phase handshake is an observational perturbation, so on/off differences",
        "are retained as controls. Two repeats are shown as observations; no",
        "bootstrap interval is calculated.",
        "",
        "The complete per-mapping classification, phase values, operation/drop",
        "deltas, source-metrics checks, output checks and counter disagreements are",
        "in `memory-analysis.json`.",
        "",
    ]
    return "\n".join(lines)


def main() -> None:
    require(sys.argv[1:] in (["--write"], ["--check"]),
            "usage: memory_analysis.py --write|--check")
    result = memory_analysis()
    encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
    if sys.argv[1] == "--write":
        OUT_JSON.write_text(encoded)
        OUT_MD.write_text(markdown(result))
        print("0788 memory phase analysis written")
    else:
        require(OUT_JSON.is_file(), "memory-analysis.json is missing")
        require(OUT_JSON.read_text() == encoded,
                "memory-analysis.json is stale or non-deterministic")
        expected_md = markdown(result)
        require(OUT_MD.is_file() and OUT_MD.read_text() == expected_md,
                "memory-analysis.md is stale")
        print("0788 memory phase analysis replay PASS")


if __name__ == "__main__":
    try:
        main()
    except ReplayError as error:
        print(f"memory analysis failed: {error}", file=sys.stderr)
        raise SystemExit(1)
