#!/usr/bin/env python3
"""Fail-closed summary and replay for the 0423 PPTX lifecycle bundle.

The bundle contains one current revision, two lifecycle roles, two corpora, two
repeats, and normal/allocator lanes.  This program verifies custody and the
single-report contract before deriving statistics.  It deliberately does not
compare owned with source-backed elapsed values, normal with allocator elapsed
values, or any revision pair.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import gzip
import hashlib
import json
import math
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


CHANGE = 423
CPU = 2
REPEATS = ("R1", "R2")
ROLES = ("owned", "source")
CORPORA = ("plain", "media_rich")
LANES = ("normal", "allocator")
LANE_CONFIG = {"normal": {"samples": 100, "warmups": 10}, "allocator": {"samples": 30, "warmups": 3}}
ORDER = (("owned", "R1"), ("source", "R1"), ("source", "R2"), ("owned", "R2"))
SELECTORS = {
    ("plain", "owned"): "pptx_cross_copy_plain_lifecycle",
    ("plain", "source"): "pptx_source_backed_cross_copy_plain_lifecycle",
    ("media_rich", "owned"): "pptx_cross_copy_media_rich_lifecycle",
    ("media_rich", "source"): "pptx_source_backed_cross_copy_media_rich_lifecycle",
}
NORMAL_BINARY = "litchi-perf-baseline"
ALLOCATOR_BINARY = "litchi-perf-baseline-alloc"
ALLOCATOR_REVISION = "serialized_region_peak_v3"
TOOLCHAIN = "1.98.1"
DRIFT_CEILINGS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}
VECTOR_FIELDS = (
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
SHA256 = re.compile(r"^[0-9a-f]{64}$")
RSS = re.compile(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", re.MULTILINE)
EXIT = re.compile(r"^\s*Exit status:\s*(-?\d+)\s*$", re.MULTILINE)
MAX_ARTIFACT_BYTES = 512 * 1024 * 1024
MAX_DECODED_BYTES = 512 * 1024 * 1024


class SummaryError(RuntimeError):
    """A custody, schema, or evidence-contract failure."""


def fail(message: str) -> None:
    raise SummaryError(message)


def strict_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON number {value!r}")


def object_value(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label} must be an object")
    return value


def list_value(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        fail(f"{label} must be a list")
    return value


def string_value(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label} must be a non-empty string")
    return value


def integer_value(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label} must be an integer >= {minimum}")
    return value


def sha_value(value: Any, label: str) -> str:
    value = string_value(value, label).lower()
    if SHA256.fullmatch(value) is None:
        fail(f"{label} must be a lowercase SHA-256")
    return value


def finite_number(value: Any, label: str, minimum: float | None = None) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        fail(f"{label} must be a finite number")
    number = float(value)
    if not math.isfinite(number) or (minimum is not None and number < minimum):
        fail(f"{label} must be a finite number >= {minimum}")
    return number


def digest_json(value: Any) -> str:
    try:
        payload = json.dumps(
            value, sort_keys=True, separators=(",", ":"),
            ensure_ascii=False, allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot hash JSON identity: {error}")
    return hashlib.sha256(payload).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def relative(root: Path, value: Any, label: str) -> Path:
    raw = string_value(value, label)
    path = Path(raw)
    if path.is_absolute() or ".." in path.parts:
        fail(f"{label} must be a relative path without '..'")
    resolved = (root / path).resolve()
    try:
        resolved.relative_to(root.resolve())
    except ValueError:
        fail(f"{label} escapes the evidence root")
    return resolved


def bounded_bytes(path: Path, label: str) -> bytes:
    try:
        if not path.is_file():
            fail(f"{label} is missing: {path}")
        if path.stat().st_size > MAX_ARTIFACT_BYTES:
            fail(f"{label} exceeds the raw artifact limit")
    except OSError as error:
        fail(f"cannot inspect {label}: {error}")
    try:
        if path.name.endswith(".gz"):
            stream = gzip.open(path, "rb")
            close = stream.close
        elif path.name.endswith(".zst"):
            process = subprocess.Popen(
                ["zstd", "-q", "-dc", str(path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            raw, stderr = process.communicate(timeout=120)
            if process.returncode != 0:
                fail(f"{label} zstd decompression failed: {stderr.decode('utf-8', 'replace').strip()}")
            if len(raw) > MAX_DECODED_BYTES:
                fail(f"{label} exceeds the decoded artifact limit")
            return raw
        else:
            raw = path.read_bytes()
            if len(raw) > MAX_DECODED_BYTES:
                fail(f"{label} exceeds the decoded artifact limit")
            return raw
        chunks: list[bytes] = []
        total = 0
        try:
            while True:
                chunk = stream.read(1024 * 1024)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_DECODED_BYTES:
                    fail(f"{label} exceeds the decoded artifact limit")
                chunks.append(chunk)
        finally:
            close()
        return b"".join(chunks)
    except FileNotFoundError:
        fail(f"zstd is unavailable while reading {label}")
    except subprocess.TimeoutExpired:
        process.kill()
        process.communicate()
        fail(f"zstd decompression timed out for {label}")
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        fail(f"cannot read {label}: {error}")
    raise AssertionError("unreachable")


def load_json(path: Path, label: str) -> tuple[Any, str]:
    raw = bounded_bytes(path, label)
    try:
        value = json.loads(
            raw.decode("utf-8"), object_pairs_hook=strict_pairs,
            parse_constant=reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"{label} is invalid JSON: {error}")
    return value, hashlib.sha256(raw).hexdigest()


def write_json(path: Path, value: Any) -> None:
    try:
        path.write_text(
            json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False) + "\n",
            encoding="utf-8",
        )
    except (OSError, TypeError, ValueError, OverflowError) as error:
        fail(f"cannot write {path}: {error}")


def artifact(root: Path, value: Any, label: str, *, allow_empty: bool) -> tuple[Path, dict[str, Any]]:
    entry = object_value(value, label)
    path = relative(root, entry.get("path"), f"{label}.path")
    try:
        size = path.stat().st_size
    except OSError as error:
        fail(f"cannot inspect {label}: {error}")
    if size == 0 and not allow_empty:
        fail(f"{label} is empty")
    expected_bytes = integer_value(entry.get("bytes"), f"{label}.bytes")
    expected_sha = sha_value(entry.get("sha256"), f"{label}.sha256")
    if size != expected_bytes or sha256_file(path) != expected_sha:
        fail(f"{label} custody hash or byte count does not match")
    recorded_allow_empty = entry.get("allow_empty", allow_empty)
    if recorded_allow_empty is not allow_empty:
        fail(f"{label}.allow_empty does not match its artifact contract")
    return path, {"path": str(path.relative_to(root)), "bytes": size, "sha256": expected_sha}


def parse_time(path: Path, label: str) -> int:
    try:
        text = bounded_bytes(path, label).decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"{label} is not UTF-8: {error}")
    exits = EXIT.findall(text)
    if exits != ["0"]:
        fail(f"{label} must contain exactly one successful Exit status")
    rss = RSS.findall(text)
    if len(rss) != 1:
        fail(f"{label} must contain exactly one Maximum resident set size")
    return int(rss[0])


def find_repo_root(explicit: Path | None) -> Path:
    if explicit is not None:
        root = explicit.expanduser().resolve()
        if not (root / "tools" / "perf_abba_summary.py").is_file():
            fail(f"--repo-root lacks pinned perf_abba_summary.py: {root}")
        return root
    here = Path(__file__).resolve()
    for candidate in (here, *here.parents):
        if (candidate / "tools" / "perf_abba_summary.py").is_file():
            return candidate
    fail("cannot locate tools/perf_abba_summary.py; pass --repo-root")
    raise AssertionError("unreachable")


def shared_tools(repo_root: Path) -> tuple[Any, Any, Any]:
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    try:
        from tools import perf_abba_summary, perf_compare, validate_perf_corpus_binding
    except ImportError as error:
        fail(f"cannot import pinned shared validators: {error}")
    return perf_abba_summary, perf_compare, validate_perf_corpus_binding


def load_protocol(root: Path) -> tuple[dict[str, Any], str]:
    path = root / "protocol.json"
    protocol, _ = load_json(path, "protocol")
    protocol = object_value(protocol, "protocol")
    if protocol.get("change") != CHANGE or protocol.get("cpu") != CPU or protocol.get("workers") != 1:
        fail("protocol change/CPU/worker identity changed")
    if protocol.get("corpora") != list(CORPORA) or protocol.get("selectors") != [
        SELECTORS[("plain", "owned")], SELECTORS[("plain", "source")],
        SELECTORS[("media_rich", "owned")], SELECTORS[("media_rich", "source")],
    ]:
        fail("protocol corpus or selector order changed")
    lanes = protocol.get("lanes")
    if lanes != LANE_CONFIG:
        fail("protocol lane sample contract changed")
    if protocol.get("order") != [{"role": role, "repeat": repeat} for role, repeat in ORDER]:
        fail("protocol O/S/S/O order changed")
    if protocol.get("expected_run_count") != 16:
        fail("protocol expected_run_count must be 16")
    flags = protocol.get("common_flags")
    if not isinstance(flags, list) or not flags or any(not isinstance(flag, str) or not flag for flag in flags):
        fail("protocol.common_flags must be a non-empty string list")
    if any(flag in {"--case", "--json", "--corpus-manifest", "--samples", "--warmup", "--warmups"} for flag in flags):
        fail("protocol.common_flags contains a capture-owned flag")
    return protocol, sha256_file(path)


def build_entry_from_capture(root: Path, capture: dict[str, Any], lane: str) -> Path:
    builds = capture.get("builds")
    if isinstance(builds, dict) and isinstance(builds.get(lane), dict):
        entry = builds[lane]
        candidate = entry.get("path", entry.get("record", entry.get("build")))
        if candidate is not None:
            return relative(root, candidate, f"capture.builds.{lane}")
    for key in (f"build-{lane}.json", f"build-candidate-{lane}.json", "build.json", "build-candidate.json"):
        candidate = root / key
        if candidate.is_file():
            return candidate
    fail(f"cannot locate the {lane} build record")
    raise AssertionError("unreachable")


def build_binary(build: dict[str, Any], lane: str) -> dict[str, Any]:
    binaries = build.get("binaries")
    raw: Any = None
    if isinstance(binaries, dict):
        raw = binaries.get(lane)
        if raw is None:
            raw = binaries.get("normal" if lane == "normal" else "allocator")
    if raw is None:
        raw = build.get("binary")
    if raw is None:
        raw = build.get("binary_identity")
    binary = object_value(raw, f"{lane} build binary")
    digest = sha_value(binary.get("sha256", binary.get("binary_sha256")), f"{lane} binary sha256")
    size = integer_value(binary.get("bytes", binary.get("binary_bytes")), f"{lane} binary bytes", 1)
    path = string_value(binary.get("path"), f"{lane} binary path")
    expected_name = NORMAL_BINARY if lane == "normal" else ALLOCATOR_BINARY
    if binary.get("name") not in (None, expected_name) and Path(path).name != expected_name:
        fail(f"{lane} binary name does not identify {expected_name}")
    candidate = Path(path)
    if candidate.is_file():
        if sha256_file(candidate) != digest or candidate.stat().st_size != size:
            fail(f"{lane} binary on disk differs from its retained build identity")
    return {"path": path, "sha256": digest, "bytes": size, "name": expected_name}


def load_build(root: Path, capture: dict[str, Any], lane: str, protocol_sha: str) -> dict[str, Any]:
    path = build_entry_from_capture(root, capture, lane)
    build, _ = load_json(path, f"{lane} build")
    build = object_value(build, f"{lane} build")
    build_sha = sha256_file(path)
    if build.get("change") not in (None, CHANGE) or build.get("status") not in (None, "pass"):
        fail(f"{lane} build record is not a passing 0423 build")
    recorded_protocol = build.get("protocol_sha256")
    if recorded_protocol is not None and recorded_protocol != protocol_sha:
        fail(f"{lane} build is not bound to protocol.json")
    before = object_value(build.get("source_before", build.get("source")), f"{lane}.source_before")
    after = object_value(build.get("source_after", before), f"{lane}.source_after")
    if before.get("revision") != after.get("revision") or before.get("clean") is not True or after.get("clean") is not True:
        fail(f"{lane} build source is not a clean unchanged identity")
    if before.get("git_status_porcelain", "") != "" or after.get("git_status_porcelain", "") != "":
        fail(f"{lane} build source has a dirty status")
    binary = build_binary(build, lane)
    return {
        "path": str(path.relative_to(root)),
        "sha256": build_sha,
        "source": before,
        "source_revision": string_value(before.get("revision"), f"{lane}.source.revision"),
        "binary": binary,
    }


def expected_tool(lane: str) -> dict[str, str]:
    base = {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
    }
    if lane == "normal":
        return {**base, "binary": NORMAL_BINARY, "instrumentation": "none"}
    return {
        **base,
        "binary": ALLOCATOR_BINARY,
        "instrumentation": "system_allocator_operation_scoped",
        "allocator_counter_revision": ALLOCATOR_REVISION,
    }


def validate_elapsed(elapsed: Any, label: str, samples: int, perf_abba_summary: Any) -> tuple[list[int], list[int], dict[str, Any]]:
    elapsed = object_value(elapsed, label)
    if elapsed.get("unit") != "ns":
        fail(f"{label}.unit must be 'ns'")
    values = [integer_value(value, f"{label}.samples[{index}]", 1) for index, value in enumerate(list_value(elapsed.get("samples"), f"{label}.samples"))]
    if len(values) != samples or values != sorted(values):
        fail(f"{label}.samples must contain {samples} sorted positive values")
    order = [integer_value(value, f"{label}.sample_order[{index}]", 0) for index, value in enumerate(list_value(elapsed.get("sample_order"), f"{label}.sample_order"))]
    if len(order) != samples or sorted(order) != list(range(samples)):
        fail(f"{label}.sample_order must be a complete permutation")
    try:
        stats = perf_abba_summary.recompute_statistics(elapsed, label)
    except Exception as error:
        fail(f"{label} statistics do not recompute: {error}")
    return values, order, stats


def metric_vector(value: Any, label: str, samples: int, *, measured: bool) -> list[int] | None:
    obj = object_value(value, label)
    status = obj.get("status")
    if measured:
        if status != "measured" or obj.get("scope") != "operation_global_system_allocator":
            fail(f"{label} is not measured operation allocator evidence")
        values = [integer_value(item, f"{label}.values[{index}") for index, item in enumerate(list_value(obj.get("values"), f"{label}.values"))]
        if len(values) != samples:
            fail(f"{label}.values must contain {samples} values")
        return values
    if status not in {"unavailable", "not_applicable"}:
        fail(f"{label}.status must be unavailable/not_applicable in the normal lane")
    if obj.get("values") is not None:
        fail(f"{label}.values must be omitted when allocation is unavailable")
    return None


def chronological(order: list[int], values: list[int]) -> list[int]:
    return [value for _, value in sorted(zip(order, values), key=lambda pair: pair[0])]


def validate_allocator(operation: dict[str, Any], label: str, samples: int, order: list[int], lane: str) -> dict[str, Any]:
    allocation = object_value(operation.get("allocation"), f"{label}.allocation")
    measured = lane == "allocator"
    if not measured:
        if allocation.get("status") not in {"unavailable", "not_applicable"}:
            fail(f"{label}.allocation must be unavailable in normal reports")
        # The shared operation validator owns the exact unavailable-vector
        # shape.  Accept its full vectors or a compact status marker here.
        return {"status": allocation.get("status"), "vectors": None}
    vectors: dict[str, list[int]] = {}
    for field in VECTOR_FIELDS:
        values = metric_vector(allocation.get(field), f"{label}.allocation.{field}", samples, measured=measured)
        assert values is not None
        vectors[field] = values
    if allocation.get("status") != "measured" or allocation.get("scope") != "operation_global_system_allocator":
        fail(f"{label}.allocation status/scope is invalid")
    before_ok = all(a >= b for a, b in zip(vectors["peak_live_bytes_before"], vectors["live_bytes_before"]))
    after_ok = all(a >= b for a, b in zip(vectors["peak_live_bytes_after"], vectors["live_bytes_after"]))
    after_before_ok = all(a >= b for a, b in zip(vectors["peak_live_bytes_after"], vectors["peak_live_bytes_before"]))
    region_ok = all(
        max(before, after) <= region <= lifetime
        for before, after, region, lifetime in zip(
            vectors["live_bytes_before"], vectors["live_bytes_after"],
            vectors["region_peak_live_bytes"], vectors["peak_live_bytes_after"],
        )
    )
    before_chron = chronological(order, vectors["peak_live_bytes_before"])
    after_chron = chronological(order, vectors["peak_live_bytes_after"])
    before_monotonic = all(a <= b for a, b in zip(before_chron, before_chron[1:]))
    after_monotonic = all(a <= b for a, b in zip(after_chron, after_chron[1:]))
    boundary_monotonic = all(a >= b for a, b in zip(before_chron[1:], after_chron))
    invariants = {
        "peak_before_ge_live_before": before_ok,
        "peak_after_ge_live_after": after_ok,
        "peak_after_ge_peak_before": after_before_ok,
        "region_peak_within_endpoint_and_lifetime_bounds": region_ok,
        "peak_before_monotonic": before_monotonic,
        "peak_after_monotonic": after_monotonic,
        "peak_boundary_monotonic": boundary_monotonic,
    }
    if not all(invariants.values()):
        fail(f"{label} allocator V3 invariants failed: {invariants}")
    return {"status": "measured", "vectors": vectors, "invariants": invariants}


def vector_stats(values: list[int]) -> dict[str, Any]:
    return {"count": len(values), "min": min(values), "max": max(values), "mean": sum(values) / len(values)}


def validate_operation(result: dict[str, Any], label: str, lane: str, samples: int, order: list[int], perf_compare: Any) -> tuple[dict[str, Any], dict[str, Any]]:
    operation = object_value(result.get("operation_metrics"), f"{label}.operation_metrics")
    if operation.get("sample_count") != samples or operation.get("sample_indices") != order:
        fail(f"{label}.operation_metrics sample alignment is invalid")
    if operation.get("alignment") != "elapsed_ns.samples_by_elapsed_then_sample_index":
        fail(f"{label}.operation_metrics.alignment is invalid")
    try:
        perf_compare._validate_operation_metrics(
            operation, f"{label}.operation_metrics", result["elapsed_ns"]["samples"], 1,
            elapsed_sample_order=order, tool_identity=result["_tool"],
        )
    except Exception as error:
        fail(f"{label}.operation_metrics rejected by shared comparator: {error}")
    allocation = validate_allocator(operation, label, samples, order, lane)
    sink = object_value(operation.get("sink"), f"{label}.operation_metrics.sink")
    if sink.get("write_status") != "measured":
        fail(f"{label}.operation_metrics.sink.write_status must be measured")
    return operation, {"allocation": allocation, "sink": sink}


def core_identity(result: dict[str, Any], special: dict[str, Any], corpus: str, role: str, label: str) -> dict[str, Any]:
    shape = object_value(result.get("corpus"), f"{label}.corpus")
    if shape.get("shape") != corpus and not (corpus == "media_rich" and shape.get("shape") == "media-rich"):
        fail(f"{label}.corpus.shape does not match {corpus}")
    def required_field(name: str, *aliases: str) -> Any:
        for candidate in (name, *aliases):
            if candidate in special:
                return special[candidate]
        fail(f"{label}.{name} is missing from the lifecycle identity")
        raise AssertionError("unreachable")

    planned_part_count = "planned_part_count"
    planned_bytes = "planned_bytes"
    external_relationship_count = "external_relationship_count"
    collision_remapped_parts = "collision_remapped_parts"
    if role == "source":
        planned_part_count = "matched_owned_planned_part_count"
        planned_bytes = "matched_owned_planned_bytes"
        external_relationship_count = "matched_owned_external_relationship_count"
        collision_remapped_parts = "matched_owned_collision_remapped_parts"
    identity = {
        "shape": corpus,
        "source_archive_sha256": sha_value(special.get("source_archive_sha256"), f"{label}.source_archive_sha256"),
        "destination_archive_sha256": sha_value(special.get("destination_archive_sha256"), f"{label}.destination_archive_sha256"),
        "source_slide": integer_value(special.get("source_slide"), f"{label}.source_slide"),
        "destination_slide": integer_value(special.get("destination_slide"), f"{label}.destination_slide"),
        "insertion_position": integer_value(special.get("insertion_position"), f"{label}.insertion_position"),
        "destination_slide_count_before": integer_value(special.get("destination_slide_count_before"), f"{label}.destination_slide_count"),
        "destination_slide_count_after": integer_value(special.get("destination_slide_count_after"), f"{label}.destination_slide_count_after"),
        "planned_part_count": integer_value(required_field(planned_part_count), f"{label}.{planned_part_count}"),
        "planned_bytes": integer_value(required_field(planned_bytes), f"{label}.{planned_bytes}"),
        "external_relationship_count": integer_value(required_field(external_relationship_count, "matched_owned_external_relationships"), f"{label}.{external_relationship_count}"),
        "collision_remapped_parts": integer_value(required_field(collision_remapped_parts), f"{label}.{collision_remapped_parts}"),
    }
    if corpus == "media_rich":
        # The source-backed summary carries this marker; the older owned
        # summary is identified by the selector/corpus contract.
        if role == "source" and special.get("media_rich") is not True:
            fail(f"{label} must identify a media-rich lifecycle")
        if "media_leaf_count" in special:
            identity["media_leaf_count"] = integer_value(special.get("media_leaf_count"), f"{label}.media_leaf_count")
        if "media_leaf_bytes" in special:
            identity["media_leaf_bytes"] = integer_value(special.get("media_leaf_bytes"), f"{label}.media_leaf_bytes")
    else:
        if special.get("media_rich", False) is not False:
            fail(f"{label} plain lifecycle must not identify media_rich")
    return identity


def validate_special(result: dict[str, Any], label: str, corpus: str, role: str, samples: int, elapsed_samples: list[int]) -> dict[str, Any]:
    source = object_value(result.get("source"), f"{label}.source")
    key = "pptx_cross_copy" if role == "owned" else "pptx_source_backed_cross_copy_lifecycle"
    special = object_value(source.get(key), f"{label}.source.{key}")
    performance_claim = special.get("performance_claim")
    if not isinstance(performance_claim, str) or not performance_claim.startswith("none:"):
        fail(f"{label}.{key}.performance_claim must withhold claims")
    if "lifecycle_ns" not in special:
        fail(f"{label}.{key}.lifecycle_ns is required")
    lifecycle = [integer_value(value, f"{label}.{key}.lifecycle_ns[{index}]", 1) for index, value in enumerate(list_value(special.get("lifecycle_ns"), f"{label}.{key}.lifecycle_ns"))]
    if len(lifecycle) != samples or lifecycle != elapsed_samples:
        fail(f"{label}.{key}.lifecycle_ns must align exactly with elapsed_ns.samples")
    phase_fields = ("open_ns", "plan_ns", "publication_ns") if role == "source" else ("plan_ns", "commit_ns", "publication_ns")
    phases: dict[str, list[int]] = {}
    for field in phase_fields:
        values = [integer_value(value, f"{label}.{key}.{field}[{index}]", 0) for index, value in enumerate(list_value(special.get(field), f"{label}.{key}.{field}"))]
        if len(values) != samples or any(value > total for value, total in zip(values, lifecycle)):
            fail(f"{label}.{key}.{field} is not a bounded phase vector")
        phases[field] = values
    gates = object_value(special.get("gates"), f"{label}.{key}.gates")
    if not gates or any(value is not True for value in gates.values()):
        fail(f"{label}.{key}.gates must all be true")
    output = sha_value(result.get("output_sha256"), f"{label}.output_sha256")
    expected = sha_value(special.get("expected_output_sha256"), f"{label}.{key}.expected_output_sha256")
    if output != expected:
        fail(f"{label} output digest differs from expected output")
    output_vector = [sha_value(value, f"{label}.{key}.output_sha256[{index}") for index, value in enumerate(list_value(special.get("output_sha256"), f"{label}.{key}.output_sha256"))]
    if len(output_vector) != samples or output_vector != [output] * samples:
        fail(f"{label}.{key}.output_sha256 is not a stable per-sample vector")
    identity = core_identity(result, special, corpus, role, label)
    source_reads: dict[str, list[int]] = {}
    expected_output_bytes: int | None = None
    if role == "source":
        expected_bytes_value = special.get("expected_output_bytes", special.get("source_expected_output_bytes"))
        expected_output_bytes = integer_value(expected_bytes_value, f"{label}.{key}.expected_output_bytes", 1)
        for field in ("source_read_calls", "source_read_bytes", "destination_read_calls", "destination_read_bytes"):
            values = [integer_value(value, f"{label}.{key}.{field}[{index}]", 1) for index, value in enumerate(list_value(special.get(field), f"{label}.{key}.{field}"))]
            if len(values) != samples or not any(values):
                fail(f"{label}.{key}.{field} must contain non-zero per-sample source evidence")
            source_reads[field] = values
    return {
        "key": key,
        "identity": identity,
        "output_sha256": output,
        "output_vector_sha256": digest_json(output_vector),
        "lifecycle_stats": statistics_from_values(lifecycle),
        "phase_stats": {field: statistics_from_values(values) for field, values in phases.items()},
        "source_read_stats": {field: statistics_from_values(values) for field, values in source_reads.items()},
        "expected_output_bytes": expected_output_bytes,
        "gates": gates,
    }


def statistics_from_values(values: list[int]) -> dict[str, Any]:
    ordered = sorted(values)
    n = len(ordered)
    left = ordered[(n - 1) // 2]
    right = ordered[n // 2]
    return {
        "count": n,
        "min": ordered[0],
        "p50": left // 2 + right // 2 + (left % 2 + right % 2) // 2,
        "p95": ordered[min(((95 * n + 99) // 100) - 1, n - 1)],
        "p99": ordered[min(((99 * n + 99) // 100) - 1, n - 1)],
        "max": ordered[-1],
        "mean": sum(values) / n,
    }


def validate_sink(result: dict[str, Any], label: str, expected_output_bytes: int | None = None) -> dict[str, int]:
    sink = object_value(result.get("sink"), f"{label}.sink")
    accepted = integer_value(sink.get("accepted_bytes"), f"{label}.sink.accepted_bytes", 1)
    writes = integer_value(sink.get("write_calls"), f"{label}.sink.write_calls", 1)
    largest = integer_value(sink.get("largest_write"), f"{label}.sink.largest_write")
    if largest > 65536:
        fail(f"{label}.sink.largest_write exceeds the protocol write chunk")
    buckets = object_value(sink.get("write_size_buckets"), f"{label}.sink.write_size_buckets")
    if integer_value(buckets.get("bytes_over_65536"), f"{label}.sink.write_size_buckets.bytes_over_65536") != 0:
        fail(f"{label}.sink has a write over the protocol chunk")
    if expected_output_bytes is not None and accepted != expected_output_bytes:
        fail(f"{label}.sink.accepted_bytes differs from expected_output_bytes")
    return {"accepted_bytes": accepted, "write_calls": writes, "largest_write": largest}


def validate_report(
    report: dict[str, Any], catalog: dict[str, Any], *, root: Path, repo_root: Path,
    label: str, lane: str, corpus: str, role: str, selector: str,
    samples: int, warmups: int, journal: dict[str, Any], binary: dict[str, Any],
    perf_abba_summary: Any, perf_compare: Any, corpus_binding: Any, report_raw_sha: str,
) -> dict[str, Any]:
    if report.get("schema_version") != 1 or not isinstance(report.get("results"), list) or len(report["results"]) != 1:
        fail(f"{label} has an invalid report envelope")
    expected = expected_tool(lane)
    if report.get("tool") != expected:
        fail(f"{label}.tool does not match {lane} identity")
    identity = object_value(report.get("binary_identity"), f"{label}.binary_identity")
    if identity.get("binary_sha256") != binary["sha256"] or identity.get("binary_bytes") != binary["bytes"]:
        fail(f"{label}.binary_identity differs from its build record")
    environment = object_value(report.get("environment"), f"{label}.environment")
    if environment.get("git_revision") != journal.get("source_revision") or environment.get("git_worktree_dirty") is not False:
        fail(f"{label}.environment source identity is not clean/bound")
    configuration = object_value(report.get("configuration"), f"{label}.configuration")
    if configuration.get("samples_per_case") != samples or configuration.get("warmup_iterations_per_case") != warmups or configuration.get("cases") != [selector]:
        fail(f"{label}.configuration sample/case identity differs")
    if configuration.get("execution_workers", [1]) != [1]:
        fail(f"{label}.configuration.execution_workers must be [1]")
    try:
        perf_compare.validate_parallel_metrics(report, label)
        corpus_binding.validate_binding(report, catalog)
    except Exception as error:
        fail(f"{label} shared report/catalog validation failed: {error}")
    result = object_value(report["results"][0], f"{label}.results[0]")
    if result.get("case") != selector:
        fail(f"{label}.results[0].case differs from the journal")
    elapsed_samples, sample_order, elapsed_stats = validate_elapsed(result.get("elapsed_ns"), f"{label}.elapsed_ns", samples, perf_abba_summary)
    result_for_operation = dict(result)
    result_for_operation["_tool"] = expected
    operation, operation_proof = validate_operation(result_for_operation, f"{label}.results[0]", lane, samples, sample_order, perf_compare)
    special_proof = validate_special(result, f"{label}.results[0]", corpus, role, samples, elapsed_samples)
    sink = validate_sink(result, f"{label}.results[0]", special_proof["expected_output_bytes"])
    catalog_ref = object_value(report.get("corpus_catalog"), f"{label}.corpus_catalog")
    catalog_identity = {key: catalog.get(key) for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256")}
    if catalog_ref != catalog_identity:
        fail(f"{label}.corpus_catalog differs from the catalog sidecar")
    return {
        "label": label, "lane": lane, "corpus": corpus, "role": role,
        "repeat": journal["repeat"], "selector": selector,
        "report_sha256": report_raw_sha,
        "catalog_sha256": digest_json(catalog),
        "catalog_identity": catalog_identity,
        "catalog_identity_sha256": digest_json(catalog_identity),
        "corpus_identity": special_proof["identity"],
        "output_sha256": special_proof["output_sha256"],
        "output_vector_sha256": special_proof["output_vector_sha256"],
        "elapsed_stats": elapsed_stats,
        "elapsed_sample_sha256": digest_json(elapsed_samples),
        "sample_order_sha256": digest_json(sample_order),
        "phase_stats": special_proof["phase_stats"],
        "source_read_stats": special_proof["source_read_stats"],
        "expected_output_bytes": special_proof["expected_output_bytes"],
        "sink": sink,
        "allocation": operation_proof["allocation"],
        "allocation_stats": {
            field: vector_stats(values)
            for field, values in (operation_proof["allocation"].get("vectors") or {}).items()
        },
        "whole_process_rss_kib": parse_time(root / journal["_artifact_paths"]["time_v"], f"{label}.time-v"),
        "protocol_sha256": journal["protocol_sha256"],
        "build_sha256": journal["build_sha256"],
        "binary_sha256": binary["sha256"],
        "source_revision": journal["source_revision"],
        "source_identity_sha256": journal.get("source_identity_sha256"),
        "artifact_hashes": journal["_artifact_hashes"],
        "verification_summary_sha256": journal["_verification_summary_sha256"],
        "invariants": operation_proof["allocation"].get("invariants"),
    }


def run_verifier(root: Path, repo_root: Path, report: Path, catalog: Path, selector: str, lane: str, samples: int, warmups: int, expected_sha: str, label: str) -> None:
    verifier = root / "verify.py"
    if not verifier.is_file():
        fail(f"{label} retained verify.py is missing")
    command = [
        sys.executable, str(verifier), "--repo-root", str(repo_root),
        "--report", str(report), "--catalog", str(catalog),
        "--selector", selector, "--lane", lane,
        "--samples", str(samples), "--warmups", str(warmups),
    ]
    try:
        process = subprocess.run(command, cwd=repo_root, capture_output=True, text=True, check=False)
    except OSError as error:
        fail(f"{label} retained verifier could not start: {error}")
    if process.returncode != 0:
        fail(f"{label} retained verifier rejected the report: {(process.stderr or process.stdout).strip()}")
    try:
        value = json.loads(process.stdout, object_pairs_hook=strict_pairs, parse_constant=reject_constant)
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"{label} retained verifier emitted invalid JSON: {error}")
    value = object_value(value, f"{label}.replay_verifier")
    if digest_json(value) != expected_sha or value.get("claim_authorized") is not False or value.get("performance_claim") is not None:
        fail(f"{label} retained verifier output differs or authorizes a claim")


def verify_artifacts(root: Path, journal: dict[str, Any], label: str) -> dict[str, Path]:
    raw = object_value(journal.get("artifacts"), f"{label}.artifacts")
    required = {
        "report": False, "catalog": False, "time_v": False,
        "stdout": True, "stderr": True, "verify_stdout": False, "verify_stderr": True,
    }
    if any(name not in raw for name in required):
        fail(f"{label}.artifacts is missing a required custody artifact")
    paths: dict[str, Path] = {}
    hashes: dict[str, str] = {}
    for name, allow_empty in required.items():
        path, record = artifact(root, raw[name], f"{label}.artifacts.{name}", allow_empty=allow_empty)
        paths[name] = path
        hashes[name] = record["sha256"]
    journal["_artifact_paths"] = paths
    journal["_artifact_hashes"] = hashes
    report_decoded_sha = load_json(paths["report"], f"{label}.report")[1]
    catalog_decoded_sha = load_json(paths["catalog"], f"{label}.catalog")[1]
    verification_raw, _ = load_json(paths["verify_stdout"], f"{label}.verify_stdout")
    verification = object_value(verification_raw, f"{label}.verify_stdout")
    if verification.get("claim_authorized") is not False or verification.get("performance_claim") is not None:
        fail(f"{label}.verify_stdout must withhold claims")
    if verification.get("selector") != journal.get("selector") or verification.get("lane") != journal.get("lane"):
        fail(f"{label}.verify_stdout selector/lane identity differs")
    if verification.get("samples") != journal.get("contract", {}).get("samples") or verification.get("warmups") != journal.get("contract", {}).get("warmups"):
        fail(f"{label}.verify_stdout sample contract differs")
    reports = list_value(verification.get("reports"), f"{label}.verify_stdout.reports")
    if len(reports) != 1:
        fail(f"{label}.verify_stdout must contain one report proof")
    report_proof = object_value(reports[0], f"{label}.verify_stdout.reports[0]")
    if report_proof.get("report_sha256") != report_decoded_sha:
        fail(f"{label}.verify_stdout is not bound to the report artifact")
    if report_proof.get("catalog_sha256") not in (None, catalog_decoded_sha):
        fail(f"{label}.verify_stdout is not bound to the catalog artifact")
    recorded = object_value(journal.get("verification"), f"{label}.verification")
    verification_sha = digest_json(verification)
    if recorded.get("status") != "pass" or recorded.get("summary_sha256") != verification_sha:
        fail(f"{label}.verification custody differs from retained verifier output")
    journal["_verification_summary_sha256"] = verification_sha
    return paths


def parse_timestamp(value: Any, label: str) -> _datetime.datetime:
    raw = string_value(value, label)
    try:
        parsed = _datetime.datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError as error:
        fail(f"{label} is not an ISO timestamp: {error}")
    if parsed.tzinfo is None:
        fail(f"{label} must include a timezone")
    return parsed


def validate_journal(root: Path, manifest_run: dict[str, Any], builds: dict[str, dict[str, Any]], protocol: dict[str, Any], protocol_sha: str, verifier_sha: str, repo_root: Path, shared: tuple[Any, Any, Any]) -> dict[str, Any]:
    lane = string_value(manifest_run.get("lane"), "capture.run.lane")
    corpus = string_value(manifest_run.get("corpus"), "capture.run.corpus")
    role = string_value(manifest_run.get("role"), "capture.run.role")
    repeat = string_value(manifest_run.get("repeat"), "capture.run.repeat")
    selector = SELECTORS.get((corpus, role))
    if lane not in LANES or corpus not in CORPORA or role not in ROLES or repeat not in REPEATS or selector is None:
        fail("capture run has an unknown lane/corpus/role/repeat")
    if manifest_run.get("selector") != selector:
        fail(f"{lane}/{corpus}/{role}/{repeat} selector differs from protocol")
    journal_path = relative(root, manifest_run.get("journal"), "capture.run.journal")
    if not journal_path.is_file() or sha256_file(journal_path) != sha_value(manifest_run.get("journal_sha256"), "capture.run.journal_sha256"):
        fail(f"{lane}/{corpus}/{role}/{repeat} journal custody failed")
    journal_raw, _ = load_json(journal_path, str(journal_path))
    journal = object_value(journal_raw, str(journal_path))
    label = f"{lane}/{corpus}/{role}/{repeat}"
    if journal.get("status") != "pass" or journal.get("change") != CHANGE or journal.get("lane") != lane or journal.get("corpus") != corpus or journal.get("role") != role or journal.get("repeat") != repeat or journal.get("selector") != selector:
        fail(f"{label} journal identity/status failed")
    if journal.get("fresh_process") is not True or journal.get("cpu") != CPU or journal.get("exit_code") != 0:
        fail(f"{label} is not a successful fresh CPU-2 process")
    expected_config = LANE_CONFIG[lane]
    if journal.get("contract") != {"lane": lane, "samples": expected_config["samples"], "warmups": expected_config["warmups"]}:
        fail(f"{label} journal contract differs")
    build = builds[lane]
    if journal.get("protocol_sha256") != protocol_sha or journal.get("build_sha256") != build["sha256"]:
        fail(f"{label} protocol/build custody differs")
    journal_build = object_value(journal.get("build"), f"{label}.build")
    if journal_build.get("sha256") != build["sha256"] or journal_build.get("path") != build["path"]:
        fail(f"{label}.build custody differs")
    journal_protocol = object_value(journal.get("protocol"), f"{label}.protocol")
    if journal_protocol.get("sha256") != protocol_sha or journal_protocol.get("path") != "protocol.json":
        fail(f"{label}.protocol custody differs")
    if journal.get("binary_sha256") != build["binary"]["sha256"] or journal.get("binary_bytes") != build["binary"]["bytes"]:
        fail(f"{label} binary custody differs")
    if journal.get("source_revision") != build["source_revision"]:
        fail(f"{label} source revision differs from build")
    if journal.get("source_identity_sha256") != digest_json(build["source"]):
        fail(f"{label} source identity differs from build")
    if journal.get("source_after_identity_sha256") != journal.get("source_identity_sha256"):
        fail(f"{label} source changed during capture")
    binary_after = object_value(journal.get("binary_after"), f"{label}.binary_after")
    if binary_after.get("sha256") != build["binary"]["sha256"] or binary_after.get("bytes") != build["binary"]["bytes"]:
        fail(f"{label} binary changed during capture")
    if journal.get("common_flags") != protocol["common_flags"]:
        fail(f"{label} common flags differ from protocol")
    verifier_record = object_value(journal.get("verifier"), f"{label}.verifier")
    if Path(string_value(verifier_record.get("path"), f"{label}.verifier.path")).name != "verify.py":
        fail(f"{label}.verifier path is not verify.py")
    if sha_value(verifier_record.get("sha256"), f"{label}.verifier.sha256") != verifier_sha:
        fail(f"{label}.verifier.sha256 differs from retained verify.py")
    if sha_value(journal.get("verifier_sha256"), f"{label}.verifier_sha256") != verifier_sha:
        fail(f"{label}.verifier_sha256 differs from retained verify.py")
    environment = object_value(journal.get("environment"), f"{label}.environment")
    if environment.get("RUSTUP_TOOLCHAIN") != TOOLCHAIN:
        fail(f"{label}.environment.RUSTUP_TOOLCHAIN differs")
    artifacts = verify_artifacts(root, journal, label)
    argv = list_value(journal.get("argv"), f"{label}.argv")
    expected_argv = [
        "taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o", "<time-v>",
        build["binary"]["path"], "--case", selector, *protocol["common_flags"],
        "--samples", str(expected_config["samples"]), "--warmup", str(expected_config["warmups"]),
        "--json", "<report>", "--corpus-manifest", "<catalog>",
    ]
    if len(argv) != len(expected_argv):
        fail(f"{label}.argv has unexpected extra or missing arguments")
    for index, (actual, expected) in enumerate(zip(argv, expected_argv)):
        if expected == "<time-v>":
            matches = Path(actual).name == artifacts["time_v"].name
        elif expected == "<report>":
            matches = Path(actual).name == artifacts["report"].name
        elif expected == "<catalog>":
            matches = Path(actual).name == artifacts["catalog"].name
        else:
            matches = actual == expected
        if not matches:
            fail(f"{label}.argv[{index}] differs from the frozen capture command")
    verify_argv = list_value(journal.get("verify_argv"), f"{label}.verify_argv")
    for flag, value in (("--selector", selector), ("--lane", lane), ("--samples", str(expected_config["samples"])), ("--warmups", str(expected_config["warmups"]))):
        if verify_argv.count(flag) != 1 or verify_argv[verify_argv.index(flag) + 1] != value:
            fail(f"{label}.verify_argv {flag} differs from protocol")
    if verify_argv.count("--report") != 1 or Path(verify_argv[verify_argv.index("--report") + 1]).name != artifacts["report"].name:
        fail(f"{label}.verify_argv report path is not bound")
    if verify_argv.count("--catalog") != 1 or Path(verify_argv[verify_argv.index("--catalog") + 1]).name != artifacts["catalog"].name:
        fail(f"{label}.verify_argv catalog path is not bound")
    for key in ("started_utc", "finished_utc"):
        parse_timestamp(journal.get(key), f"{label}.{key}")
    report_raw, report_sha = load_json(artifacts["report"], f"{label}.report")
    catalog_raw, _ = load_json(artifacts["catalog"], f"{label}.catalog")
    report = object_value(report_raw, f"{label}.report")
    catalog = object_value(catalog_raw, f"{label}.catalog")
    perf_abba_summary, perf_compare, corpus_binding = shared
    proof = validate_report(
        report, catalog, root=root, repo_root=repo_root, label=label, lane=lane,
        corpus=corpus, role=role, selector=selector, samples=expected_config["samples"],
        warmups=expected_config["warmups"], journal=journal, binary=build["binary"],
        perf_abba_summary=perf_abba_summary, perf_compare=perf_compare,
        corpus_binding=corpus_binding, report_raw_sha=sha256_file(artifacts["report"]),
    )
    run_verifier(root, repo_root, artifacts["report"], artifacts["catalog"], selector, lane, expected_config["samples"], expected_config["warmups"], journal["_verification_summary_sha256"], label)
    proof["started_utc"] = journal["started_utc"]
    proof["finished_utc"] = journal["finished_utc"]
    proof["journal_sha256"] = sha256_file(journal_path)
    proof["artifact_hashes"] = journal["_artifact_hashes"]
    proof["verification_summary_sha256"] = journal["_verification_summary_sha256"]
    proof["journal_path"] = str(journal_path.relative_to(root))
    return proof


def drift(first: dict[str, Any], second: dict[str, Any]) -> dict[str, Any]:
    values: dict[str, float] = {}
    within: dict[str, bool] = {}
    for field, ceiling in DRIFT_CEILINGS.items():
        a = float(first["elapsed_stats"][field])
        b = float(second["elapsed_stats"][field])
        if a <= 0:
            fail(f"cannot calculate repeat drift from non-positive {field}")
        values[field] = (b - a) / a * 100.0
        within[field] = math.isfinite(values[field]) and abs(values[field]) <= ceiling
    return {
        "r2_minus_r1_percent": values,
        "ceilings_percent": dict(DRIFT_CEILINGS),
        "within_ceiling": within,
        "accepted_statistics": [field for field in DRIFT_CEILINGS if within[field]],
        "withheld_statistics": [field for field in DRIFT_CEILINGS if not within[field]],
        "acceptance_grade": all(within.values()),
    }


def render_table(rows: list[dict[str, Any]], drifts: dict[str, Any]) -> str:
    lines = [
        "# 0423 same-revision lifecycle baseline",
        "",
        "Normal rows retain timing statistics; allocator rows retain resource vectors and whole-process RSS. No owned/source or normal/allocator elapsed comparison is made.",
        "",
        "| Lane | Corpus | Role | Repeat | Samples | p50 ns | Mean ns | p95 ns | p99 ns | RSS KiB | Alloc calls mean | Allocated bytes mean | Region peak mean |",
        "|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in rows:
        elapsed = row["elapsed_stats"]
        alloc = row["allocation_stats"]
        lines.append(
            "| {lane} | {corpus} | {role} | {repeat} | {count} | {p50} | {mean} | {p95} | {p99} | {rss:,} | {calls} | {bytes} | {region} |".format(
                lane=row["lane"], corpus=row["corpus"], role=row["role"], repeat=row["repeat"],
                count=elapsed["sample_count"], p50=elapsed["p50"] if row["lane"] == "normal" else "—",
                mean=f"{elapsed['mean']:.3f}" if row["lane"] == "normal" else "—",
                p95=elapsed["p95"] if row["lane"] == "normal" else "—",
                p99=elapsed["p99"] if row["lane"] == "normal" else "—",
                rss=row["whole_process_rss_kib"],
                calls=f"{alloc.get('allocation_calls', {}).get('mean', '—'):.3f}" if "allocation_calls" in alloc else "—",
                bytes=f"{alloc.get('allocated_bytes', {}).get('mean', '—'):.3f}" if "allocated_bytes" in alloc else "—",
                region=f"{alloc.get('region_peak_live_bytes', {}).get('mean', '—'):.3f}" if "region_peak_live_bytes" in alloc else "—",
            )
        )
    lines.extend(["", "## Within-role repeat drift", ""])
    lines.append("| Corpus | Role | p50 | Mean | p95 | p99 | Acceptance-grade all fields |")
    lines.append("|---|---|---:|---:|---:|---:|---|")
    for key, value in sorted(drifts.items()):
        corpus, role = key.split("/", 1)
        d = value["r2_minus_r1_percent"]
        lines.append(f"| {corpus} | {role} | {d['p50']:+.3f}% | {d['mean']:+.3f}% | {d['p95']:+.3f}% | {d['p99']:+.3f}% | {value['acceptance_grade']} |")
    lines.extend([
        "",
        "All source/output and semantic/raw-preservation gates are retained per role. Allocator elapsed values are intentionally withheld from comparisons; allocator vectors are process callback-order observations, and RSS is whole-process GNU `time -v` evidence.",
        "",
    ])
    return "\n".join(lines)


def summarize(root: Path, repo_root_arg: Path | None, replay: bool) -> None:
    root = root.expanduser().resolve()
    capture_raw, _ = load_json(root / "capture.json", "capture")
    capture = object_value(capture_raw, "capture")
    if capture.get("change") != CHANGE or capture.get("status") != "pass" or capture.get("claim_authorized") is not False or capture.get("performance_claim") is not None:
        fail("capture is not a passing no-claim 0423 bundle")
    protocol, protocol_sha = load_protocol(root)
    capture_verifier = object_value(capture.get("verifier"), "capture.verifier")
    capture_verifier_sha = sha_value(capture_verifier.get("sha256"), "capture.verifier.sha256")
    retained_verifier = root / "verify.py"
    if not retained_verifier.is_file() or sha256_file(retained_verifier) != capture_verifier_sha:
        fail("capture verifier identity differs from retained verify.py")
    repo_root = find_repo_root(repo_root_arg)
    shared = shared_tools(repo_root)
    builds = {lane: load_build(root, capture, lane, protocol_sha) for lane in LANES}
    revisions = {builds[lane]["source_revision"] for lane in LANES}
    if len(revisions) != 1:
        fail("normal and allocator builds do not share one source revision")
    baseline = protocol.get("baseline_revision")
    if baseline is not None and (not isinstance(baseline, str) or not baseline):
        fail("protocol baseline_revision must be a non-empty contextual anchor")
    if capture.get("protocol", {}).get("sha256") not in (None, protocol_sha):
        fail("capture protocol custody differs")
    raw_runs = list_value(capture.get("runs"), "capture.runs")
    if len(raw_runs) != 16 or capture.get("expected_run_count") != 16:
        fail("capture must contain exactly 16 runs")
    observed_keys = {(item.get("lane"), item.get("corpus"), item.get("role"), item.get("repeat")) for item in raw_runs}
    expected_keys = {(lane, corpus, role, repeat) for lane in LANES for corpus in CORPORA for role in ROLES for repeat in REPEATS}
    if observed_keys != expected_keys:
        fail("capture runs do not cover every lane/corpus/role/repeat exactly once")
    rows: list[dict[str, Any]] = []
    for item in raw_runs:
        rows.append(validate_journal(root, item, builds, protocol, protocol_sha, capture_verifier_sha, repo_root, shared))
    order_index = {pair: index for index, pair in enumerate(ORDER)}
    rows.sort(key=lambda row: (LANES.index(row["lane"]), CORPORA.index(row["corpus"]), order_index[(row["role"], row["repeat"])]))
    # Every lane/corpus must have the frozen O/S/S/O sequence, and successful
    # journal intervals must not overlap in the capture order.
    for lane in LANES:
        for corpus in CORPORA:
            group = [row for row in rows if row["lane"] == lane and row["corpus"] == corpus]
            if [(row["role"], row["repeat"]) for row in group] != list(ORDER):
                fail(f"{lane}/{corpus} is not in frozen O/S/S/O order")
    ordered_by_time = sorted(rows, key=lambda row: parse_timestamp(row["started_utc"], f"{row['label']}.started_utc"))
    for previous, current in zip(ordered_by_time, ordered_by_time[1:]):
        if parse_timestamp(previous["finished_utc"], f"{previous['label']}.finished_utc") > parse_timestamp(current["started_utc"], f"{current['label']}.started_utc"):
            fail(f"capture journal intervals overlap: {previous['label']} and {current['label']}")
    # Normalize shared input identity and enforce it across all roles/lanes.
    input_gates: dict[str, Any] = {}
    for corpus in CORPORA:
        group = [row for row in rows if row["corpus"] == corpus]
        # Owned and source-backed summaries expose the same matched input
        # identity, while only the source report carries media-leaf fields.
        def shared_identity(row: dict[str, Any]) -> dict[str, Any]:
            return {key: value for key, value in row["corpus_identity"].items() if key not in {"media_leaf_count", "media_leaf_bytes"}}

        first = shared_identity(group[0])
        if any(shared_identity(row) != first for row in group):
            fail(f"{corpus} source/destination/index identity differs across roles or lanes")
        if corpus == "media_rich":
            source_media = {
                (row["corpus_identity"].get("media_leaf_count"), row["corpus_identity"].get("media_leaf_bytes"))
                for row in group if row["role"] == "source"
            }
            if len(source_media) != 1 or None in next(iter(source_media)):
                fail(f"{corpus} source media-leaf identity is not stable")
        owned_outputs = {row["sink"]["accepted_bytes"] for row in group if row["role"] == "owned"}
        if len(owned_outputs) != 1:
            fail(f"{corpus} owned output size is not stable across lanes/repeats")
        owned_bytes = next(iter(owned_outputs))
        ceiling = 2 * owned_bytes + 65536
        if any(row["sink"]["accepted_bytes"] > ceiling for row in group):
            fail(f"{corpus} output exceeds the protocol common sink ceiling")
        input_gates[corpus] = {
            "identity": first,
            "source_media_identity": sorted(source_media) if corpus == "media_rich" else None,
            "owned_output_bytes": owned_bytes,
            "common_sink_ceiling_bytes": ceiling,
            "all_outputs_fit_sink_ceiling": True,
            "role_outputs": {
                role: sorted({row["output_sha256"] for row in group if row["role"] == role})
                for role in ROLES
            },
        }
    drifts: dict[str, Any] = {}
    for corpus in CORPORA:
        for role in ROLES:
            group = [row for row in rows if row["lane"] == "normal" and row["corpus"] == corpus and row["role"] == role]
            first = next(row for row in group if row["repeat"] == "R1")
            second = next(row for row in group if row["repeat"] == "R2")
            drifts[f"{corpus}/{role}"] = drift(first, second)
    summary = {
        "change": CHANGE,
        "classification": "same-revision owned/source-backed lifecycle baseline; no optimization claim",
        "claim_authorized": False,
        "performance_claim": None,
        "claim_withheld_reason": "0423 permits per-role current distributions, resource vectors, logical I/O and repeat stability only",
        "protocol": {"path": "protocol.json", "sha256": protocol_sha},
        "source": {"revision": next(iter(revisions)), "baseline_revision": baseline},
        "builds": {lane: {"path": builds[lane]["path"], "sha256": builds[lane]["sha256"], "binary": builds[lane]["binary"]} for lane in LANES},
        "configuration": {"cpu": CPU, "workers": 1, "lanes": LANE_CONFIG, "corpora": list(CORPORA), "roles": list(ROLES), "order": [{"role": role, "repeat": repeat} for role, repeat in ORDER]},
        "input_gates": input_gates,
        "repeat_drift": drifts,
        "runs": rows,
        "limitations": [
            "Owned/source-backed elapsed values use different public API paths and may produce different ZIP bytes; no cross-role speedup claim is made.",
            "Normal and allocator elapsed values are not compared; allocator V3 values are callback-order process resource vectors with observer overhead.",
            "Logical in-memory ReadAt counters are not physical-I/O measurements. Whole-process RSS includes setup, validation, and teardown.",
            "Teardown is outside the lifecycle interval, so live-after and region peaks do not establish post-drop retention or an owner budget.",
            "Replay revalidates retained reports, catalogs, journals, verifier output, and custody without requiring the original worktrees or binaries.",
        ],
    }
    rendered = json.dumps(summary, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False) + "\n"
    table = render_table(rows, drifts)
    summary_path = root / "summary.json"
    table_path = root / "result-table.md"
    if replay:
        try:
            if summary_path.read_text(encoding="utf-8") != rendered or table_path.read_text(encoding="utf-8") != table:
                fail("retained summary or result-table differs from deterministic replay")
        except OSError as error:
            fail(f"replay output is missing: {error}")
    else:
        if summary_path.exists() or table_path.exists():
            fail("refusing to overwrite existing summary outputs; use --replay")
        write_json(summary_path, summary)
        try:
            table_path.write_text(table, encoding="utf-8")
        except OSError as error:
            fail(f"cannot write result-table.md: {error}")
    print(json.dumps({"status": "pass", "change": CHANGE, "runs": len(rows), "replay": replay}, sort_keys=True))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--replay", action="store_true")
    args = parser.parse_args()
    try:
        summarize(args.root, args.repo_root, args.replay)
    except SummaryError as error:
        print(f"0423 summary failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
