#!/usr/bin/env python3
"""Fail-closed summary and replay for the conditional 0424 capture.

The capture contains control and candidate source-backed lifecycle reports for
plain and media-rich corpora.  This tool rechecks every report, catalog,
journal, build receipt, pinned validator, and verifier proof before deriving
same-role repeat drift and diagnostic resource deltas.  It preserves measured
vectors and deliberately withholds speedup, release-latency, and optimization
claims.
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


CHANGE = 424
VERIFIER_CHANGE = 423
CPU = 2
WORKERS = 1
TOOLCHAIN = "1.98.1"
REPEATS = ("R1", "R2")
ROLES = ("control", "candidate")
CORPORA = ("plain", "media_rich")
LANES = ("normal", "allocator")
LANE_CONFIG = {
    "normal": {"samples": 100, "warmups": 10},
    "allocator": {"samples": 30, "warmups": 3},
}
ORDER = (("control", "R1"), ("candidate", "R1"), ("candidate", "R2"), ("control", "R2"))
SELECTORS = {
    ("plain", "control"): "pptx_source_backed_cross_copy_plain_lifecycle",
    ("plain", "candidate"): "pptx_source_backed_cross_copy_plain_lifecycle",
    ("media_rich", "control"): "pptx_source_backed_cross_copy_media_rich_lifecycle",
    ("media_rich", "candidate"): "pptx_source_backed_cross_copy_media_rich_lifecycle",
}
NORMAL_BINARY = "litchi-perf-baseline"
ALLOCATOR_BINARY = "litchi-perf-baseline-alloc"
ALLOCATOR_REVISION = "serialized_region_peak_v3"
DRIFT_CEILINGS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}
REGRESSION_THRESHOLD = 5.0
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
PROCESS_VECTOR_FIELDS = (
    "user_cpu_ticks",
    "system_cpu_ticks",
    "clock_ticks_per_second",
    "minor_faults",
    "major_faults",
    "voluntary_context_switches",
    "nonvoluntary_context_switches",
    "rss_delta_bytes",
    "peak_rss_bytes",
    "rchar",
    "wchar",
    "read_bytes",
    "write_bytes",
    "cancelled_write_bytes",
    "syscr",
    "syscw",
)
SINK_VECTOR_FIELDS = ("accepted_bytes", "write_calls", "largest_write")
SINK_BUCKET_FIELDS = (
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
)
SHA256 = re.compile(r"^[0-9a-f]{64}$")
GIT_REVISION = re.compile(r"^[0-9a-f]{40}$")
RSS = re.compile(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", re.MULTILINE)
EXIT = re.compile(r"^\s*Exit status:\s*(-?\d+)\s*$", re.MULTILINE)
MAX_ARTIFACT_BYTES = 512 * 1024 * 1024
MAX_DECODED_BYTES = 512 * 1024 * 1024
PINNED_TOOL_FILES = (
    "pinned/tools/perf_abba_summary.py",
    "pinned/tools/perf_compare.py",
    "pinned/tools/perf_resource_profile.py",
    "pinned/tools/validate_perf_corpus_binding.py",
)


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


def external_reference(root: Path, value: Any, label: str, expected: Path) -> None:
    raw = string_value(value, label)
    resolved = (root / raw).resolve() if not Path(raw).is_absolute() else Path(raw).resolve()
    if resolved != expected.resolve():
        fail(f"{label} does not resolve to {expected}")


def portable_external_reference(root: Path, value: Any, label: str, expected: Path) -> None:
    """Validate a captured path while allowing relocation of the bundle.

    The measurement manifest records the verifier's absolute path from the
    capture host.  Its content hash is checked separately, so replay accepts
    that historical absolute path when the retained basename matches the
    current pinned copy.
    """
    raw = string_value(value, label)
    path = Path(raw)
    if path.is_absolute():
        if path.name != expected.name:
            fail(f"{label} does not identify {expected.name}")
        return
    external_reference(root, raw, label, expected)


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
                stream.close()
            return b"".join(chunks)
        if path.name.endswith(".zst"):
            process = subprocess.Popen(
                ["zstd", "-q", "-dc", str(path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            try:
                raw, stderr = process.communicate(timeout=120)
            except subprocess.TimeoutExpired:
                process.kill()
                process.communicate()
                fail(f"zstd decompression timed out for {label}")
            if process.returncode != 0:
                fail(f"{label} zstd decompression failed: {stderr.decode('utf-8', 'replace').strip()}")
            if len(raw) > MAX_DECODED_BYTES:
                fail(f"{label} exceeds the decoded artifact limit")
            return raw
        raw = path.read_bytes()
        if len(raw) > MAX_DECODED_BYTES:
            fail(f"{label} exceeds the decoded artifact limit")
        return raw
    except FileNotFoundError:
        fail(f"zstd is unavailable while reading {label}")
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


def load_protocol(bundle: Path) -> tuple[dict[str, Any], str]:
    path = bundle / "measurement-protocol.json"
    protocol, raw_sha = load_json(path, "measurement protocol")
    protocol = object_value(protocol, "measurement protocol")
    if protocol.get("change") != CHANGE or protocol.get("cpu") != CPU or protocol.get("workers") != WORKERS:
        fail("measurement protocol change/CPU/worker identity changed")
    if protocol.get("corpora") != list(CORPORA):
        fail("measurement protocol corpus order changed")
    if protocol.get("selectors") != [SELECTORS[(corpus, "control")] for corpus in CORPORA]:
        fail("measurement protocol selector order changed")
    if protocol.get("lanes") != LANE_CONFIG:
        fail("measurement protocol lane sample contract changed")
    if protocol.get("order") != [{"role": role, "repeat": repeat} for role, repeat in ORDER]:
        fail("measurement protocol control/candidate order changed")
    if protocol.get("expected_run_count") != 16:
        fail("measurement protocol expected_run_count must be 16")
    expected_flags = [
        "--shape", "many-small", "--payload", "compressible", "--writer-shape", "large",
        "--xlsx-shape", "medium", "--xlsx-cell-crud-shape", "medium",
        "--xlsx-row-visibility-shape", "medium", "--semantic-shape", "medium",
        "--workers", "1", "--filesystem-cache", "warm",
    ]
    if protocol.get("common_flags") != expected_flags:
        fail("measurement protocol common_flags differ from the frozen protocol")
    if any(flag in {"--case", "--json", "--corpus-manifest", "--samples", "--warmup", "--warmups"} for flag in expected_flags):
        fail("measurement protocol common_flags contain capture-owned options")
    drift = object_value(protocol.get("normal_repeat_drift_ceilings_percent"), "measurement protocol drift ceilings")
    actual_drift = {field: finite_number(drift.get(field), f"measurement protocol drift.{field}", 0.0) for field in DRIFT_CEILINGS}
    if actual_drift != DRIFT_CEILINGS:
        fail("measurement protocol normal repeat drift ceilings changed")
    baseline = string_value(protocol.get("baseline_revision"), "measurement protocol baseline_revision").lower()
    if GIT_REVISION.fullmatch(baseline) is None:
        fail("measurement protocol baseline_revision must be a full lowercase revision")
    return protocol, raw_sha


def build_binary(build: dict[str, Any], lane: str, label: str) -> dict[str, Any]:
    binaries = object_value(build.get("binaries"), f"{label}.binaries")
    binary = object_value(binaries.get(lane), f"{label}.binaries.{lane}")
    digest = sha_value(binary.get("sha256", binary.get("binary_sha256")), f"{label}.{lane}.sha256")
    size = integer_value(binary.get("bytes", binary.get("binary_bytes")), f"{label}.{lane}.bytes", 1)
    path = string_value(binary.get("path"), f"{label}.{lane}.path")
    if not Path(path).is_absolute() or not path.startswith("/tmp/"):
        fail(f"{label}.{lane}.path must identify an absolute copied /tmp binary")
    if binary.get("binary_sha256") not in (None, digest) or binary.get("binary_bytes") not in (None, size):
        fail(f"{label}.{lane} binary aliases disagree")
    if binary.get("executable") is not None and binary.get("executable") is not True:
        fail(f"{label}.{lane} binary is not marked executable")
    return {
        "path": path,
        "sha256": digest,
        "bytes": size,
        "binary_sha256": digest,
        "binary_bytes": size,
        "label": binary.get("label"),
        "profile": binary.get("profile"),
    }


def load_build(bundle: Path, role: str, protocol: dict[str, Any], protocol_sha: str) -> dict[str, Any]:
    path = bundle / f"measurement-build-{role}.json"
    build, raw_sha = load_json(path, f"{role} measurement build")
    build = object_value(build, f"{role} measurement build")
    if build.get("change") != CHANGE or build.get("status") != "pass" or build.get("exit_code") != 0:
        fail(f"{role} measurement build is not a passing 0424 receipt")
    if build.get("role") not in (None, role):
        fail(f"{role} measurement build role does not match")
    before = object_value(build.get("source_before"), f"{role}.source_before")
    after = object_value(build.get("source_after"), f"{role}.source_after")
    if before != after or before.get("clean") is not True or before.get("git_status_porcelain", "") != "":
        fail(f"{role} source_before/source_after must be one clean unchanged identity")
    revision = string_value(before.get("revision"), f"{role}.source.revision").lower()
    if GIT_REVISION.fullmatch(revision) is None:
        fail(f"{role} source revision must be a full lowercase revision")
    if role == "control" and revision != protocol["baseline_revision"]:
        fail("control source revision must equal measurement protocol baseline_revision")
    build_protocol = sha_value(build.get("protocol_sha256"), f"{role}.protocol_sha256")
    if role == "candidate" and build_protocol != protocol_sha:
        fail("candidate measurement build is not bound to measurement-protocol.json")
    binaries = {lane: build_binary(build, lane, role) for lane in LANES}
    record = {
        "role": role,
        "path": path.name,
        "sha256": raw_sha,
        "status": build.get("status"),
        "build_kind": build.get("build_kind"),
        "build_protocol_sha256": build_protocol,
        "source_before": before,
        "source_after": after,
        "source_revision": revision,
        "source_identity_sha256": digest_json(before),
        "binaries": binaries,
        "origin_build": build.get("origin_build"),
    }
    if role == "control":
        if build.get("build_kind") != "reused_exact_prior_clean_build_artifacts":
            fail("control receipt must identify exact prior clean build reuse")
        origin = object_value(build.get("origin_build"), "control.origin_build")
        origin_path = bundle / string_value(origin.get("path"), "control.origin_build.path")
        if origin_path.resolve().parent != bundle.resolve():
            fail("control origin build must be beside the measurement bundle")
        if sha256_file(origin_path) != sha_value(origin.get("sha256"), "control.origin_build.sha256"):
            fail("control origin build receipt hash differs")
        origin_raw, _ = load_json(origin_path, "control origin build")
        origin_raw = object_value(origin_raw, "control origin build")
        if origin_raw.get("status") != "pass" or origin_raw.get("change") != VERIFIER_CHANGE:
            fail("control origin build is not the retained 0423 passing receipt")
        if object_value(origin_raw.get("source_before"), "origin.source_before") != before:
            fail("control receipt source identity differs from origin build")
        if object_value(origin_raw.get("source_after"), "origin.source_after") != after:
            fail("control receipt source-after identity differs from origin build")
        origin_binaries = object_value(origin_raw.get("binaries"), "origin.binaries")
        for lane in LANES:
            origin_binary = object_value(origin_binaries.get(lane), f"origin.binaries.{lane}")
            origin_sha = sha_value(origin_binary.get("sha256", origin_binary.get("binary_sha256")), f"origin.{lane}.sha256")
            origin_bytes = integer_value(origin_binary.get("bytes", origin_binary.get("binary_bytes")), f"origin.{lane}.bytes", 1)
            if origin_sha != binaries[lane]["sha256"] or origin_bytes != binaries[lane]["bytes"]:
                fail(f"control {lane} binary differs from origin build")
        origin_protocol = sha_value(origin_raw.get("protocol_sha256"), "origin.protocol_sha256")
        if build_protocol == protocol_sha or build_protocol != origin_protocol:
            fail("control reused receipt does not retain its origin build protocol identity")
    return record


def validate_pinned_manifest(bundle: Path, pinned: Path) -> str:
    path = bundle / "pinned" / "tools-manifest.json"
    manifest, raw_sha = load_json(path, "pinned tools manifest")
    manifest = object_value(manifest, "pinned tools manifest")
    entries = list_value(manifest.get("files"), "pinned tools manifest.files")
    by_path: dict[str, dict[str, Any]] = {}
    for entry in entries:
        obj = object_value(entry, "pinned tools manifest.files[]")
        name = string_value(obj.get("path"), "pinned tools manifest file.path")
        if name in by_path:
            fail(f"pinned tools manifest repeats {name}")
        by_path[name] = obj
    if set(by_path) != set(PINNED_TOOL_FILES):
        fail("pinned tools manifest file set differs from the four retained validators")
    for relative_name in PINNED_TOOL_FILES:
        entry = by_path[relative_name]
        if string_value(entry.get("source_path"), f"{relative_name}.source_path") == "":
            fail(f"{relative_name}.source_path is empty")
        actual = bundle / relative_name
        if not actual.is_file():
            fail(f"pinned validator is missing: {actual}")
        if sha256_file(actual) != sha_value(entry.get("sha256"), f"{relative_name}.sha256"):
            fail(f"pinned validator hash differs: {relative_name}")
    if not (pinned / "verify-report.py").is_file():
        fail("pinned verify-report.py is missing")
    return raw_sha


def shared_tools(pinned: Path) -> tuple[Any, Any, Any, Any]:
    if str(pinned) not in sys.path:
        sys.path.insert(0, str(pinned))
    for name in list(sys.modules):
        if name == "tools" or name.startswith("tools."):
            del sys.modules[name]
    try:
        from tools import perf_abba_summary, perf_compare, perf_resource_profile, validate_perf_corpus_binding
    except ImportError as error:
        fail(f"cannot import pinned shared validators: {error}")
    return perf_abba_summary, perf_compare, perf_resource_profile, validate_perf_corpus_binding


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


def vector_metric(value: Any, label: str, samples: int, *, status: str, scope: str | None = None) -> list[int]:
    metric = object_value(value, label)
    if metric.get("status") != status:
        fail(f"{label}.status must be {status!r}")
    if scope is not None and metric.get("scope") != scope:
        fail(f"{label}.scope must be {scope!r}")
    values = [integer_value(item, f"{label}.values[{index}]") for index, item in enumerate(list_value(metric.get("values"), f"{label}.values"))]
    if len(values) != samples:
        fail(f"{label}.values must contain {samples} values")
    return values


def validate_allocator(operation: dict[str, Any], label: str, samples: int, lane: str, perf_compare: Any, result: dict[str, Any]) -> dict[str, Any]:
    allocation = object_value(operation.get("allocation"), f"{label}.allocation")
    scope = "operation_global_system_allocator"
    if lane == "normal":
        if allocation.get("status") != "unavailable" or allocation.get("scope") != scope:
            fail(f"{label}.allocation must be unavailable operation-scoped evidence")
        for field in VECTOR_FIELDS:
            metric = object_value(allocation.get(field), f"{label}.allocation.{field}")
            if metric.get("status") != "unavailable" or metric.get("values") is not None:
                fail(f"{label}.allocation.{field} must omit unavailable values")
        return {"status": "unavailable", "vectors": None, "invariants": None}
    if allocation.get("status") != "measured" or allocation.get("scope") != scope:
        fail(f"{label}.allocation must be measured V3 operation-scoped evidence")
    vectors = {field: vector_metric(allocation.get(field), f"{label}.allocation.{field}", samples, status="measured", scope=scope) for field in VECTOR_FIELDS}
    before_ok = all(a >= b for a, b in zip(vectors["peak_live_bytes_before"], vectors["live_bytes_before"]))
    after_ok = all(a >= b for a, b in zip(vectors["peak_live_bytes_after"], vectors["live_bytes_after"]))
    after_before_ok = all(a >= b for a, b in zip(vectors["peak_live_bytes_after"], vectors["peak_live_bytes_before"]))
    region_ok = all(
        max(before, after) <= region <= peak
        for before, after, region, peak in zip(
            vectors["live_bytes_before"], vectors["live_bytes_after"],
            vectors["region_peak_live_bytes"], vectors["peak_live_bytes_after"],
        )
    )
    invariants = {
        "peak_before_ge_live_before": before_ok,
        "peak_after_ge_live_after": after_ok,
        "peak_after_ge_peak_before": after_before_ok,
        "region_peak_within_endpoint_and_lifetime_bounds": region_ok,
    }
    if not all(invariants.values()):
        fail(f"{label} allocator V3 invariants failed: {invariants}")
    try:
        perf_compare._validate_allocator_operation_evidence(
            dict(result), label, samples, allocator_counter_revision=ALLOCATOR_REVISION,
        )
    except Exception as error:
        fail(f"{label} allocator V3 evidence rejected by pinned comparator: {error}")
    return {"status": "measured", "vectors": vectors, "invariants": invariants}


def vector_stats(values: list[int]) -> dict[str, Any]:
    if not values:
        fail("cannot summarize an empty vector")
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


def validate_operation(result: dict[str, Any], label: str, lane: str, samples: int, order: list[int], perf_compare: Any) -> tuple[dict[str, Any], dict[str, Any]]:
    operation = object_value(result.get("operation_metrics"), f"{label}.operation_metrics")
    if operation.get("sample_count") != samples or operation.get("sample_indices") != order:
        fail(f"{label}.operation_metrics sample alignment is invalid")
    if operation.get("alignment") != "elapsed_ns.samples_by_elapsed_then_sample_index":
        fail(f"{label}.operation_metrics.alignment is invalid")
    result_for_operation = dict(result)
    result_for_operation["_tool"] = expected_tool(lane)
    try:
        perf_compare._validate_operation_metrics(
            operation, f"{label}.operation_metrics", result["elapsed_ns"]["samples"], 1,
            elapsed_sample_order=order, tool_identity=result_for_operation["_tool"],
        )
    except Exception as error:
        fail(f"{label}.operation_metrics rejected by pinned comparator: {error}")
    allocation = validate_allocator(operation, label, samples, lane, perf_compare, result_for_operation)
    sink = object_value(operation.get("sink"), f"{label}.operation_metrics.sink")
    if sink.get("write_status") != "measured":
        fail(f"{label}.operation_metrics.sink.write_status must be measured")
    return operation, {"allocation": allocation, "sink": sink}


def core_identity(special: dict[str, Any], corpus: str, label: str) -> dict[str, Any]:
    def required(name: str, *aliases: str) -> Any:
        for candidate in (name, *aliases):
            if candidate in special:
                return special[candidate]
        fail(f"{label}.{name} is missing")
        raise AssertionError("unreachable")
    identity = {
        "shape": corpus,
        "source_archive_sha256": sha_value(special.get("source_archive_sha256"), f"{label}.source_archive_sha256"),
        "destination_archive_sha256": sha_value(special.get("destination_archive_sha256"), f"{label}.destination_archive_sha256"),
        "source_slide": integer_value(special.get("source_slide"), f"{label}.source_slide"),
        "destination_slide": integer_value(special.get("destination_slide"), f"{label}.destination_slide"),
        "insertion_position": integer_value(special.get("insertion_position"), f"{label}.insertion_position"),
        "destination_slide_count_before": integer_value(special.get("destination_slide_count_before"), f"{label}.destination_slide_count_before"),
        "destination_slide_count_after": integer_value(special.get("destination_slide_count_after"), f"{label}.destination_slide_count_after"),
        "planned_part_count": integer_value(required("matched_owned_planned_part_count"), f"{label}.matched_owned_planned_part_count"),
        "planned_bytes": integer_value(required("matched_owned_planned_bytes"), f"{label}.matched_owned_planned_bytes"),
        "external_relationship_count": integer_value(required("matched_owned_external_relationship_count", "matched_owned_external_relationships"), f"{label}.matched_owned_external_relationship_count"),
        "collision_remapped_parts": integer_value(required("matched_owned_collision_remapped_parts"), f"{label}.matched_owned_collision_remapped_parts"),
        "matched_owned_output_bytes": integer_value(special.get("matched_owned_output_bytes"), f"{label}.matched_owned_output_bytes", 1),
        "matched_owned_output_ceiling": integer_value(special.get("matched_owned_output_ceiling"), f"{label}.matched_owned_output_ceiling", 1),
        "source_expected_output_bytes": integer_value(special.get("source_expected_output_bytes"), f"{label}.source_expected_output_bytes", 1),
    }
    if corpus == "media_rich":
        if special.get("media_rich") is not True:
            fail(f"{label} must identify a media-rich lifecycle")
        identity["media_leaf_count"] = integer_value(special.get("media_leaf_count"), f"{label}.media_leaf_count")
        identity["media_leaf_bytes"] = integer_value(special.get("media_leaf_bytes"), f"{label}.media_leaf_bytes")
    elif special.get("media_rich") is not False:
        fail(f"{label} plain lifecycle must not identify media_rich")
    return identity


def validate_special(result: dict[str, Any], label: str, corpus: str, samples: int, elapsed_samples: list[int]) -> dict[str, Any]:
    source = object_value(result.get("source"), f"{label}.source")
    key = "pptx_source_backed_cross_copy_lifecycle"
    special = object_value(source.get(key), f"{label}.source.{key}")
    claim = special.get("performance_claim")
    if not isinstance(claim, str) or not claim.startswith("none:"):
        fail(f"{label}.{key}.performance_claim must withhold claims")
    lifecycle = [integer_value(value, f"{label}.{key}.lifecycle_ns[{index}]", 1) for index, value in enumerate(list_value(special.get("lifecycle_ns"), f"{label}.{key}.lifecycle_ns"))]
    if len(lifecycle) != samples or lifecycle != elapsed_samples:
        fail(f"{label}.{key}.lifecycle_ns must align with elapsed_ns.samples")
    phases: dict[str, list[int]] = {}
    for field in ("open_ns", "plan_ns", "publication_ns"):
        values = [integer_value(value, f"{label}.{key}.{field}[{index}]", 1) for index, value in enumerate(list_value(special.get(field), f"{label}.{key}.{field}"))]
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
    output_vector = [sha_value(value, f"{label}.{key}.output_sha256[{index}]") for index, value in enumerate(list_value(special.get("output_sha256"), f"{label}.{key}.output_sha256"))]
    if len(output_vector) != samples or output_vector != [output] * samples:
        fail(f"{label}.{key}.output_sha256 is not a stable per-sample vector")
    source_reads: dict[str, list[int]] = {}
    for field in ("source_read_calls", "source_read_bytes", "destination_read_calls", "destination_read_bytes"):
        values = [integer_value(value, f"{label}.{key}.{field}[{index}]", 1) for index, value in enumerate(list_value(special.get(field), f"{label}.{key}.{field}"))]
        if len(values) != samples:
            fail(f"{label}.{key}.{field} must contain {samples} samples")
        source_reads[field] = values
    return {
        "identity": core_identity(special, corpus, label),
        "output_sha256": output,
        "output_vector_sha256": digest_json(output_vector),
        "lifecycle_vectors": {"lifecycle_ns": lifecycle, **phases, **source_reads, "output_sha256": output_vector},
        "lifecycle_stats": vector_stats(lifecycle),
        "phase_stats": {field: vector_stats(values) for field, values in phases.items()},
        "source_read_stats": {field: vector_stats(values) for field, values in source_reads.items()},
        "expected_output_bytes": integer_value(special.get("source_expected_output_bytes"), f"{label}.{key}.source_expected_output_bytes", 1),
        "gates": gates,
    }


def validate_sink(
    result: dict[str, Any], label: str, expected_output_bytes: int,
    matched_owned_output_bytes: int, ceiling: int,
) -> dict[str, int]:
    sink = object_value(result.get("sink"), f"{label}.sink")
    accepted = integer_value(sink.get("accepted_bytes"), f"{label}.sink.accepted_bytes", 1)
    writes = integer_value(sink.get("write_calls"), f"{label}.sink.write_calls", 1)
    largest = integer_value(sink.get("largest_write"), f"{label}.sink.largest_write")
    if accepted != expected_output_bytes:
        fail(f"{label}.sink.accepted_bytes differs from source_expected_output_bytes")
    if accepted > ceiling:
        fail(f"{label}.sink.accepted_bytes exceeds matched-owned output ceiling")
    if accepted == ceiling:
        fail(f"{label}.sink.accepted_bytes unexpectedly equals the reservation ceiling")
    if accepted == matched_owned_output_bytes:
        fail(f"{label}.source sink bytes must remain distinct from matched-owned output bytes")
    if largest > 65536:
        fail(f"{label}.sink.largest_write exceeds the protocol write chunk")
    buckets = object_value(sink.get("write_size_buckets"), f"{label}.sink.write_size_buckets")
    if integer_value(buckets.get("bytes_over_65536"), f"{label}.sink.write_size_buckets.bytes_over_65536") != 0:
        fail(f"{label}.sink has a write over the protocol chunk")
    return {"accepted_bytes": accepted, "write_calls": writes, "largest_write": largest}


def metric_vectors(group: Any, fields: tuple[str, ...], label: str, samples: int, expected_status: str) -> dict[str, list[int]]:
    group = object_value(group, label)
    if group.get("status") != expected_status:
        fail(f"{label}.status must be {expected_status!r}")
    result: dict[str, list[int]] = {}
    for field in fields:
        result[field] = vector_metric(group.get(field), f"{label}.{field}", samples, status=expected_status)
    return result


def validate_sink_operation(group: Any, label: str, samples: int) -> dict[str, Any]:
    group = object_value(group, label)
    if group.get("status") != "not_applicable" or group.get("write_status") != "measured":
        fail(f"{label} must expose measured write vectors under write_status")
    vectors = metric_vectors({"status": "measured", **{field: group.get(field) for field in SINK_VECTOR_FIELDS}}, SINK_VECTOR_FIELDS, f"{label}.write", samples, "measured")
    buckets = object_value(group.get("write_size_buckets"), f"{label}.write_size_buckets")
    if buckets.get("status") != "measured":
        fail(f"{label}.write_size_buckets.status must be measured")
    bucket_vectors = {
        field: vector_metric(buckets.get(field), f"{label}.write_size_buckets.{field}", samples, status="measured")
        for field in SINK_BUCKET_FIELDS
    }
    vectors["write_size_buckets"] = bucket_vectors
    return vectors


def validate_report(
    report: dict[str, Any], catalog: dict[str, Any], *, label: str, lane: str, corpus: str,
    role: str, selector: str, samples: int, warmups: int, journal: dict[str, Any],
    binary: dict[str, Any], perf_abba_summary: Any, perf_compare: Any,
    corpus_binding: Any, report_raw_sha: str, catalog_raw_sha: str, time_v: Path,
    artifact_hashes: dict[str, str], verification_summary_sha256: str,
) -> dict[str, Any]:
    if report.get("schema_version") != 1 or not isinstance(report.get("results"), list) or len(report["results"]) != 1:
        fail(f"{label} has an invalid report envelope")
    if report.get("tool") != expected_tool(lane):
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
    if configuration.get("execution_workers") != [WORKERS]:
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
    operation, operation_proof = validate_operation(result, f"{label}.results[0]", lane, samples, sample_order, perf_compare)
    special = validate_special(result, f"{label}.results[0]", corpus, samples, elapsed_samples)
    sink = validate_sink(
        result, f"{label}.results[0]", special["expected_output_bytes"],
        special["identity"]["matched_owned_output_bytes"],
        special["identity"]["matched_owned_output_ceiling"],
    )
    catalog_ref = object_value(report.get("corpus_catalog"), f"{label}.corpus_catalog")
    catalog_identity = {key: catalog.get(key) for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256")}
    if catalog_ref != catalog_identity:
        fail(f"{label}.corpus_catalog differs from the catalog sidecar")
    process_vectors = metric_vectors(operation.get("process"), PROCESS_VECTOR_FIELDS, f"{label}.operation_metrics.process", samples, "measured")
    sink_vectors = validate_sink_operation(operation.get("sink"), f"{label}.operation_metrics.sink", samples)
    if lane == "allocator" and operation_proof["allocation"]["status"] != "measured":
        fail(f"{label} allocator allocation vectors are unavailable")
    if lane == "normal" and operation_proof["allocation"]["status"] != "unavailable":
        fail(f"{label} normal allocation vectors must be unavailable")
    return {
        "label": label, "lane": lane, "corpus": corpus, "role": role, "repeat": journal["repeat"], "selector": selector,
        "report_sha256": report_raw_sha, "catalog_sha256": catalog_raw_sha,
        "catalog_identity": catalog_identity, "catalog_identity_sha256": digest_json(catalog_identity),
        "corpus_identity": special["identity"], "output_sha256": special["output_sha256"],
        "raw_corpus_identity_sha256": digest_json(object_value(result.get("corpus"), f"{label}.results[0].corpus")),
        "output_vector_sha256": special["output_vector_sha256"], "expected_output_bytes": special["expected_output_bytes"],
        "elapsed_stats": elapsed_stats, "elapsed_samples": elapsed_samples, "sample_order": sample_order,
        "elapsed_sample_sha256": digest_json(elapsed_samples), "sample_order_sha256": digest_json(sample_order),
        "lifecycle_vectors": special["lifecycle_vectors"], "lifecycle_stats": special["lifecycle_stats"],
        "phase_stats": special["phase_stats"], "source_read_stats": special["source_read_stats"], "gates": special["gates"],
        "sink": sink, "allocation": operation_proof["allocation"],
        "allocation_stats": {field: vector_stats(values) for field, values in (operation_proof["allocation"].get("vectors") or {}).items()},
        "allocation_vectors": operation_proof["allocation"].get("vectors"),
        "process_vectors": process_vectors, "process_stats": {field: vector_stats(values) for field, values in process_vectors.items()},
        "sink_vectors": sink_vectors,
        "whole_process_rss_kib": parse_time(time_v, f"{label}.time-v"),
        "protocol_sha256": journal["protocol_sha256"], "build_sha256": journal["build_sha256"],
        "binary_sha256": binary["sha256"], "source_revision": journal["source_revision"],
        "source_identity_sha256": journal["source_identity_sha256"],
        "source_before_identity_sha256": journal["source_before_identity_sha256"],
        "source_after_identity_sha256": journal["source_after_identity_sha256"],
        "source_worktree": journal.get("source_worktree"), "argv": journal["argv"], "verify_argv": journal["verify_argv"],
        "environment": journal["environment"], "started_utc": journal["started_utc"], "finished_utc": journal["finished_utc"],
        "binary": journal["binary"], "binary_before": journal["binary_before"], "binary_after": journal["binary_after"],
        "build": journal["build"], "protocol": journal["protocol"], "verifier": journal["verifier"],
        "artifact_hashes": artifact_hashes, "verification_summary_sha256": verification_summary_sha256,
        "invariants": operation_proof["allocation"].get("invariants"),
    }


def run_verifier(pinned: Path, report: Path, catalog: Path, selector: str, lane: str, samples: int, warmups: int, expected_sha: str, label: str) -> None:
    verifier = pinned / "verify-report.py"
    command = [
        sys.executable, "-B", str(verifier), "--repo-root", str(pinned), "--report", str(report),
        "--catalog", str(catalog), "--selector", selector, "--lane", lane, "--contract", "formal",
        "--samples", str(samples), "--warmups", str(warmups),
    ]
    try:
        process = subprocess.run(command, cwd=pinned, capture_output=True, text=True, check=False)
    except OSError as error:
        fail(f"{label} retained verifier could not start: {error}")
    if process.returncode != 0:
        fail(f"{label} retained verifier rejected the report: {(process.stderr or process.stdout).strip()}")
    try:
        value = json.loads(process.stdout, object_pairs_hook=strict_pairs, parse_constant=reject_constant)
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"{label} retained verifier emitted invalid JSON: {error}")
    value = object_value(value, f"{label}.replay_verifier")
    reports = list_value(value.get("reports"), f"{label}.replay_verifier.reports")
    if value.get("change") != VERIFIER_CHANGE or value.get("status") != "pass" or value.get("claim_authorized") is not False or value.get("performance_claim") is not None or value.get("selector") != selector or value.get("lane") != lane or value.get("samples") != samples or value.get("warmups") != warmups or value.get("report_count") != 1 or len(reports) != 1:
        fail(f"{label} retained verifier proof has the wrong contract")
    proof = object_value(reports[0], f"{label}.replay_verifier.reports[0]")
    if proof.get("report_sha256") != sha256_file(report) or proof.get("catalog_sha256") != sha256_file(catalog):
        fail(f"{label} retained verifier proof is not bound to the raw artifacts")
    if digest_json(value) != expected_sha:
        fail(f"{label} retained verifier output differs from the captured proof")


def verify_artifacts(root: Path, journal: dict[str, Any], label: str) -> tuple[dict[str, Path], dict[str, str], str]:
    raw = object_value(journal.get("artifacts"), f"{label}.artifacts")
    required = {
        "report": False, "catalog": False, "time_v": False,
        "stdout": True, "stderr": True, "verify_stdout": False, "verify_stderr": True,
    }
    if set(raw) != set(required):
        fail(f"{label}.artifacts must contain exactly the required custody artifacts")
    paths: dict[str, Path] = {}
    hashes: dict[str, str] = {}
    for name, allow_empty in required.items():
        path, record = artifact(root, raw[name], f"{label}.artifacts.{name}", allow_empty=allow_empty)
        paths[name] = path
        hashes[name] = record["sha256"]
    verification_raw, _ = load_json(paths["verify_stdout"], f"{label}.verify_stdout")
    verification = object_value(verification_raw, f"{label}.verify_stdout")
    if verification.get("change") != VERIFIER_CHANGE or verification.get("status") != "pass" or verification.get("claim_authorized") is not False or verification.get("performance_claim") is not None:
        fail(f"{label}.verify_stdout must be a passing no-claim proof")
    if verification.get("selector") != journal.get("selector") or verification.get("lane") != journal.get("lane"):
        fail(f"{label}.verify_stdout selector/lane identity differs")
    contract = object_value(journal.get("contract"), f"{label}.contract")
    if verification.get("samples") != contract.get("samples") or verification.get("warmups") != contract.get("warmups"):
        fail(f"{label}.verify_stdout sample contract differs")
    reports = list_value(verification.get("reports"), f"{label}.verify_stdout.reports")
    if verification.get("report_count") != 1 or len(reports) != 1:
        fail(f"{label}.verify_stdout must contain one report proof")
    report_proof = object_value(reports[0], f"{label}.verify_stdout.reports[0]")
    report_sha = load_json(paths["report"], f"{label}.report")[1]
    catalog_sha = load_json(paths["catalog"], f"{label}.catalog")[1]
    if report_proof.get("report_sha256") != report_sha or report_proof.get("catalog_sha256") != catalog_sha:
        fail(f"{label}.verify_stdout is not bound to raw report/catalog artifacts")
    recorded = object_value(journal.get("verification"), f"{label}.verification")
    verification_sha = digest_json(verification)
    if recorded.get("status") != "pass" or recorded.get("exit_code") != 0 or recorded.get("summary_sha256") != verification_sha or recorded.get("stdout_sha256") != hashes["verify_stdout"] or recorded.get("stderr_sha256") != hashes["verify_stderr"]:
        fail(f"{label}.verification custody differs from retained verifier output")
    if journal.get("verification_exit_code") != 0:
        fail(f"{label}.verification_exit_code must be zero")
    return paths, hashes, verification_sha


def parse_timestamp(value: Any, label: str) -> _datetime.datetime:
    raw = string_value(value, label)
    try:
        parsed = _datetime.datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError as error:
        fail(f"{label} is not an ISO timestamp: {error}")
    if parsed.tzinfo is None:
        fail(f"{label} must include a timezone")
    return parsed


def portable_artifact_path(value: Any, current: Path, output: Path, label: str) -> None:
    recorded = Path(string_value(value, label))
    try:
        expected_parts = current.resolve().relative_to(output.resolve()).parts
    except ValueError:
        fail(f"{label} is outside the output root")
    if not expected_parts or tuple(recorded.parts[-len(expected_parts):]) != tuple(expected_parts):
        fail(f"{label} is not bound to the retained run artifact")


def validate_command(argv: list[Any], artifacts: dict[str, Path], output: Path, build: dict[str, Any], selector: str, contract: dict[str, Any], protocol: dict[str, Any], label: str) -> None:
    expected = [
        "taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o",
        "<time-v>", build["path"], "--case", selector,
        *protocol["common_flags"], "--samples", str(contract["samples"]),
        "--warmup", str(contract["warmups"]), "--json", "<report>",
        "--corpus-manifest", "<catalog>",
    ]
    if len(argv) != len(expected):
        fail(f"{label}.argv differs from the frozen capture command")
    for index, (actual, wanted) in enumerate(zip(argv, expected)):
        if wanted == "<time-v>":
            portable_artifact_path(actual, artifacts["time_v"], output, f"{label}.argv[{index}]")
        elif wanted == "<report>":
            portable_artifact_path(actual, artifacts["report"], output, f"{label}.argv[{index}]")
        elif wanted == "<catalog>":
            portable_artifact_path(actual, artifacts["catalog"], output, f"{label}.argv[{index}]")
        elif actual != wanted:
            fail(f"{label}.argv[{index}] differs from the frozen capture command")


def validate_verify_command(argv: list[Any], pinned: Path, artifacts: dict[str, Path], output: Path, selector: str, lane: str, contract: dict[str, Any], label: str) -> None:
    if "-B" not in argv or "--contract" not in argv or argv[argv.index("--contract") + 1] != "formal":
        fail(f"{label}.verify_argv must use the formal pinned contract")
    if "--repo-root" not in argv or Path(argv[argv.index("--repo-root") + 1]).name != pinned.name:
        fail(f"{label}.verify_argv is not bound to the pinned root")
    if argv.count("--report") != 1 or argv.count("--catalog") != 1:
        fail(f"{label}.verify_argv must bind one report and one catalog")
    portable_artifact_path(argv[argv.index("--report") + 1], artifacts["report"], output, f"{label}.verify_argv.report")
    portable_artifact_path(argv[argv.index("--catalog") + 1], artifacts["catalog"], output, f"{label}.verify_argv.catalog")
    for flag, value in (("--selector", selector), ("--lane", lane), ("--samples", str(contract["samples"])), ("--warmups", str(contract["warmups"]))):
        if argv.count(flag) != 1 or argv[argv.index(flag) + 1] != value:
            fail(f"{label}.verify_argv {flag} differs from the contract")


def validate_journal(output: Path, pinned: Path, manifest_run: dict[str, Any], builds: dict[str, dict[str, Any]], protocol: dict[str, Any], protocol_sha: str, verifier_sha: str, shared: tuple[Any, Any, Any, Any]) -> dict[str, Any]:
    lane = string_value(manifest_run.get("lane"), "capture.run.lane")
    corpus = string_value(manifest_run.get("corpus"), "capture.run.corpus")
    role = string_value(manifest_run.get("role"), "capture.run.role")
    repeat = string_value(manifest_run.get("repeat"), "capture.run.repeat")
    selector = SELECTORS.get((corpus, role))
    if lane not in LANES or corpus not in CORPORA or role not in ROLES or repeat not in REPEATS or selector is None:
        fail("capture run has an unknown lane/corpus/role/repeat")
    if manifest_run.get("selector") != selector:
        fail(f"{lane}/{corpus}/{role}/{repeat} selector differs from protocol")
    if manifest_run.get("status") != "pass":
        fail(f"{lane}/{corpus}/{role}/{repeat} manifest run is not marked pass")
    manifest_index = integer_value(manifest_run.get("index"), f"{lane}/{corpus}/{role}/{repeat}.manifest.index", 1)
    journal_path = relative(output, manifest_run.get("journal"), "capture.run.journal")
    journal_sha = sha_value(manifest_run.get("journal_sha256"), "capture.run.journal_sha256")
    if not journal_path.is_file() or sha256_file(journal_path) != journal_sha:
        fail(f"{lane}/{corpus}/{role}/{repeat} journal custody failed")
    journal_raw, _ = load_json(journal_path, str(journal_path))
    journal = object_value(journal_raw, str(journal_path))
    label = f"{lane}/{corpus}/{role}/{repeat}"
    if journal.get("status") != "pass" or journal.get("change") != CHANGE or journal.get("lane") != lane or journal.get("corpus") != corpus or journal.get("role") != role or journal.get("repeat") != repeat or journal.get("selector") != selector:
        fail(f"{label} journal identity/status failed")
    if journal.get("fresh_process") is not True or journal.get("cpu") != CPU or journal.get("workers") != WORKERS or journal.get("exit_code") != 0:
        fail(f"{label} is not a successful fresh CPU-2 process")
    contract = {"lane": lane, **LANE_CONFIG[lane]}
    if journal.get("contract") != contract:
        fail(f"{label} journal contract differs")
    build = builds[role]
    if journal.get("protocol_sha256") != protocol_sha or journal.get("build_sha256") != build["sha256"]:
        fail(f"{label} protocol/build custody differs")
    journal_build = object_value(journal.get("build"), f"{label}.build")
    if journal_build.get("sha256") != build["sha256"] or Path(string_value(journal_build.get("path"), f"{label}.build.path")).name != build["path"] or journal_build.get("protocol_sha256") != build["build_protocol_sha256"]:
        fail(f"{label}.build custody differs")
    journal_protocol = object_value(journal.get("protocol"), f"{label}.protocol")
    if journal_protocol.get("sha256") != protocol_sha or Path(string_value(journal_protocol.get("path"), f"{label}.protocol.path")).name != "measurement-protocol.json" or journal.get("protocol_after_sha256") != protocol_sha:
        fail(f"{label}.protocol custody differs")
    binary = build["binaries"][lane]
    if journal.get("binary_sha256") != binary["sha256"] or journal.get("binary_bytes") != binary["bytes"]:
        fail(f"{label} binary custody differs")
    for field in ("source_revision",):
        if journal.get(field) != build["source_revision"]:
            fail(f"{label}.{field} differs from build")
    if journal.get("source_identity_sha256") != build["source_identity_sha256"] or journal.get("source_before_identity_sha256") != build["source_identity_sha256"] or journal.get("source_after_identity_sha256") != build["source_identity_sha256"] or journal.get("source_unchanged") is not True:
        fail(f"{label} source identity changed during capture")
    for field in ("binary_before", "binary_after"):
        value = object_value(journal.get(field), f"{label}.{field}")
        if value.get("sha256") != binary["sha256"] or value.get("bytes") != binary["bytes"]:
            fail(f"{label}.{field} differs from build")
    if journal.get("binary_unchanged") is not True or journal.get("common_flags") != protocol["common_flags"]:
        fail(f"{label} binary/common-flag custody differs")
    verifier_record = object_value(journal.get("verifier"), f"{label}.verifier")
    if Path(string_value(verifier_record.get("path"), f"{label}.verifier.path")).name != "verify-report.py" or sha_value(verifier_record.get("sha256"), f"{label}.verifier.sha256") != verifier_sha or sha_value(journal.get("verifier_sha256"), f"{label}.verifier_sha256") != verifier_sha or journal.get("verifier_after_sha256") != verifier_sha:
        fail(f"{label}.verifier custody differs from pinned verify-report.py")
    environment = object_value(journal.get("environment"), f"{label}.environment")
    if environment.get("RUSTUP_TOOLCHAIN") != TOOLCHAIN or environment.get("PYTHONDONTWRITEBYTECODE") != "1":
        fail(f"{label}.environment differs from protocol")
    artifacts, artifact_hashes, verification_summary_sha256 = verify_artifacts(output, journal, label)
    validate_command(list_value(journal.get("argv"), f"{label}.argv"), artifacts, output, binary, selector, contract, protocol, label)
    validate_verify_command(list_value(journal.get("verify_argv"), f"{label}.verify_argv"), pinned, artifacts, output, selector, lane, contract, label)
    for key in ("started_utc", "finished_utc"):
        parse_timestamp(journal.get(key), f"{label}.{key}")
    report_raw, report_sha = load_json(artifacts["report"], f"{label}.report")
    catalog_raw, catalog_sha = load_json(artifacts["catalog"], f"{label}.catalog")
    report = object_value(report_raw, f"{label}.report")
    catalog = object_value(catalog_raw, f"{label}.catalog")
    perf_abba_summary, perf_compare, _resource_profile, corpus_binding = shared
    proof = validate_report(
        report, catalog, label=label, lane=lane, corpus=corpus, role=role, selector=selector,
        samples=contract["samples"], warmups=contract["warmups"], journal=journal, binary=binary,
        perf_abba_summary=perf_abba_summary, perf_compare=perf_compare, corpus_binding=corpus_binding,
        report_raw_sha=report_sha, catalog_raw_sha=catalog_sha, time_v=artifacts["time_v"],
        artifact_hashes=artifact_hashes, verification_summary_sha256=verification_summary_sha256,
    )
    run_verifier(pinned, artifacts["report"], artifacts["catalog"], selector, lane, contract["samples"], contract["warmups"], verification_summary_sha256, label)
    proof["started_utc"] = journal["started_utc"]
    proof["finished_utc"] = journal["finished_utc"]
    proof["journal_sha256"] = journal_sha
    proof["journal_path"] = str(journal_path.relative_to(output))
    proof["source_before"] = builds[role]["source_before"]
    proof["source_after"] = builds[role]["source_after"]
    proof["build_kind"] = builds[role]["build_kind"]
    if manifest_run.get("output_sha256") != proof["output_sha256"]:
        fail(f"{label} manifest output digest differs from its report")
    if manifest_run.get("corpus_identity_sha256") != proof["raw_corpus_identity_sha256"]:
        fail(f"{label} manifest corpus identity differs from its report")
    proof["manifest_run_index"] = manifest_index
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


def delta(control: float, candidate: float, *, label: str) -> dict[str, Any]:
    if control < 0 or candidate < 0:
        fail(f"{label} cannot compare negative values")
    if control == 0:
        return {
            "control": control, "candidate": candidate, "percent": None,
            "undefined_baseline": True, "regression_trigger": candidate > 0,
        }
    percent = (candidate - control) / control * 100.0
    return {
        "control": control, "candidate": candidate, "percent": percent,
        "undefined_baseline": False, "regression_trigger": math.isfinite(percent) and percent > REGRESSION_THRESHOLD,
    }


def normal_diagnostics(rows: list[dict[str, Any]], triggers: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for corpus in CORPORA:
        for repeat in REPEATS:
            control = next(row for row in rows if row["lane"] == "normal" and row["corpus"] == corpus and row["role"] == "control" and row["repeat"] == repeat)
            candidate = next(row for row in rows if row["lane"] == "normal" and row["corpus"] == corpus and row["role"] == "candidate" and row["repeat"] == repeat)
            metrics: dict[str, Any] = {}
            for field in ("p50", "mean", "p95", "p99"):
                comparison = delta(float(control["elapsed_stats"][field]), float(candidate["elapsed_stats"][field]), label=f"normal/{corpus}/{repeat}/{field}")
                metrics[field] = comparison
                if comparison["regression_trigger"]:
                    triggers.append({"kind": "normal_diagnostic", "corpus": corpus, "repeat": repeat, "metric": field, **comparison})
            comparison = delta(float(control["whole_process_rss_kib"]), float(candidate["whole_process_rss_kib"]), label=f"normal/{corpus}/{repeat}/whole_process_rss_kib")
            metrics["whole_process_rss_kib"] = comparison
            if comparison["regression_trigger"]:
                triggers.append({"kind": "normal_diagnostic", "corpus": corpus, "repeat": repeat, "metric": "whole_process_rss_kib", **comparison})
            result.append({"corpus": corpus, "repeat": repeat, "metrics": metrics})
    return result


def allocator_resources(rows: list[dict[str, Any]], triggers: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    fields = ("allocated_bytes", "allocation_calls", "region_peak_live_bytes", "live_bytes_before", "live_bytes_after")
    for corpus in CORPORA:
        for repeat in REPEATS:
            control = next(row for row in rows if row["lane"] == "allocator" and row["corpus"] == corpus and row["role"] == "control" and row["repeat"] == repeat)
            candidate = next(row for row in rows if row["lane"] == "allocator" and row["corpus"] == corpus and row["role"] == "candidate" and row["repeat"] == repeat)
            metrics: dict[str, Any] = {}
            for field in (*fields, "rss_delta_bytes"):
                if field == "rss_delta_bytes":
                    control_value = float(control["process_stats"][field]["mean"])
                    candidate_value = float(candidate["process_stats"][field]["mean"])
                else:
                    control_value = float(control["allocation_stats"][field]["mean"])
                    candidate_value = float(candidate["allocation_stats"][field]["mean"])
                comparison = delta(control_value, candidate_value, label=f"allocator/{corpus}/{repeat}/{field}")
                metrics[field] = comparison
                if comparison["regression_trigger"]:
                    triggers.append({"kind": "allocator_resource", "corpus": corpus, "repeat": repeat, "metric": field, **comparison})
            comparison = delta(float(control["whole_process_rss_kib"]), float(candidate["whole_process_rss_kib"]), label=f"allocator/{corpus}/{repeat}/whole_process_rss_kib")
            metrics["whole_process_rss_kib"] = comparison
            if comparison["regression_trigger"]:
                triggers.append({"kind": "allocator_resource", "corpus": corpus, "repeat": repeat, "metric": "whole_process_rss_kib", **comparison})
            result.append({"corpus": corpus, "repeat": repeat, "metrics": metrics})
    return result


def render_table(rows: list[dict[str, Any]], drifts: dict[str, Any], normal: list[dict[str, Any]], allocator: list[dict[str, Any]], triggers: list[dict[str, Any]]) -> str:
    lines = [
        "# 0424 matched source-backed lifecycle baseline",
        "",
        "Rows preserve full vectors in `summary.json`. Normal elapsed and RSS comparisons are diagnostic; allocator comparisons cover resource vectors only. No result authorizes a performance claim.",
        "",
        "| Lane | Corpus | Role | Repeat | Samples | p50 ns | Mean ns | p95 ns | p99 ns | RSS KiB | Alloc calls mean | Allocated bytes mean | Region peak mean |",
        "|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in rows:
        elapsed = row["elapsed_stats"]
        alloc = row["allocation_stats"]
        lines.append(
            "| {lane} | {corpus} | {role} | {repeat} | {count} | {p50} | {mean} | {p95} | {p99} | {rss:,} | {calls} | {bytes} | {region} |".format(
                lane=row["lane"], corpus=row["corpus"], role=row["role"], repeat=row["repeat"], count=elapsed["sample_count"],
                p50=elapsed["p50"] if row["lane"] == "normal" else "—", mean=f"{elapsed['mean']:.3f}" if row["lane"] == "normal" else "—",
                p95=elapsed["p95"] if row["lane"] == "normal" else "—", p99=elapsed["p99"] if row["lane"] == "normal" else "—", rss=row["whole_process_rss_kib"],
                calls=f"{alloc.get('allocation_calls', {}).get('mean', 0):.3f}" if alloc else "—",
                bytes=f"{alloc.get('allocated_bytes', {}).get('mean', 0):.3f}" if alloc else "—",
                region=f"{alloc.get('region_peak_live_bytes', {}).get('mean', 0):.3f}" if alloc else "—",
            )
        )
    lines.extend(["", "## Within-role normal repeat drift", "", "| Corpus | Role | p50 | Mean | p95 | p99 | All fields within ceiling |", "|---|---|---:|---:|---:|---:|---|"])
    for key in sorted(drifts):
        corpus, role = key.split("/", 1)
        value = drifts[key]
        d = value["r2_minus_r1_percent"]
        lines.append(f"| {corpus} | {role} | {d['p50']:+.3f}% | {d['mean']:+.3f}% | {d['p95']:+.3f}% | {d['p99']:+.3f}% | {value['acceptance_grade']} |")
    withheld = [
        (key, value["withheld_statistics"])
        for key, value in sorted(drifts.items())
        if value["withheld_statistics"]
    ]
    lines.append("")
    if withheld:
        lines.append("Repeat statistics withheld for: " + ", ".join(f"{key} ({', '.join(fields)})" for key, fields in withheld) + ".")
    else:
        lines.append("No normal repeat statistic exceeded its frozen drift ceiling; no repeat statistic is withheld.")
    lines.extend(["", "## Normal diagnostic control/candidate deltas", "", "| Corpus | Repeat | Metric | Delta | >5% trigger |", "|---|---|---|---:|---|"])
    for entry in normal:
        for field, value in entry["metrics"].items():
            rendered = "undefined baseline" if value["percent"] is None else f"{value['percent']:+.3f}%"
            lines.append(f"| {entry['corpus']} | {entry['repeat']} | {field} | {rendered} | {value['regression_trigger']} |")
    lines.extend(["", "## Allocator resource deltas", "", "| Corpus | Repeat | Metric | Delta | >5% trigger |", "|---|---|---|---:|---|"])
    for entry in allocator:
        for field, value in entry["metrics"].items():
            rendered = "undefined baseline" if value["percent"] is None else f"{value['percent']:+.3f}%"
            lines.append(f"| {entry['corpus']} | {entry['repeat']} | {field} | {rendered} | {value['regression_trigger']} |")
    lines.extend(["", "## Review triggers", ""])
    if triggers:
        for trigger in triggers:
            lines.append(f"- `{trigger['kind']}` {trigger['corpus']}/{trigger['repeat']} `{trigger['metric']}`: regression trigger")
    else:
        lines.append("- None above the frozen 5% diagnostic threshold.")
    lines.extend(["", "Allocator elapsed comparisons are withheld. Whole-process RSS includes setup, verifier-visible process work, and teardown; operation allocator vectors include observer overhead. The 0424 experiment remains descriptive and does not establish a release latency or optimization claim.", ""])
    return "\n".join(lines)


def summarize(output: Path, bundle_arg: Path | None, pinned_arg: Path | None, replay: bool) -> None:
    output = output.expanduser().resolve()
    bundle = (bundle_arg.expanduser().resolve() if bundle_arg is not None else (output.parent if output.name == "matched" else output)).resolve()
    pinned = (pinned_arg.expanduser().resolve() if pinned_arg is not None else bundle / "pinned").resolve()
    if pinned != (bundle / "pinned").resolve():
        fail("pinned root must be the bundle's pinned directory")
    capture_raw, capture_sha = load_json(output / "capture.json", "capture")
    capture = object_value(capture_raw, "capture")
    if capture.get("change") != CHANGE or capture.get("status") != "pass" or capture.get("claim_authorized") is not False or capture.get("performance_claim") is not None:
        fail("capture is not a passing no-claim 0424 bundle")
    protocol, protocol_sha = load_protocol(bundle)
    capture_protocol = object_value(capture.get("protocol"), "capture.protocol")
    if sha_value(capture_protocol.get("sha256"), "capture.protocol.sha256") != protocol_sha:
        fail("capture protocol custody differs")
    external_reference(output, capture_protocol.get("path"), "capture.protocol.path", bundle / "measurement-protocol.json")
    verifier_path = pinned / "verify-report.py"
    verifier_record = object_value(capture.get("verifier"), "capture.verifier")
    verifier_sha = sha_value(verifier_record.get("sha256"), "capture.verifier.sha256")
    portable_external_reference(output, verifier_record.get("path"), "capture.verifier.path", verifier_path)
    if verifier_sha != sha256_file(verifier_path):
        fail("capture verifier identity differs from pinned verify-report.py")
    pinned_manifest_sha = validate_pinned_manifest(bundle, pinned)
    shared = shared_tools(pinned)
    builds = {role: load_build(bundle, role, protocol, protocol_sha) for role in ROLES}
    capture_builds = object_value(capture.get("builds"), "capture.builds")
    for role in ROLES:
        entry = object_value(capture_builds.get(role), f"capture.builds.{role}")
        if sha_value(entry.get("sha256"), f"capture.builds.{role}.sha256") != builds[role]["sha256"]:
            fail(f"capture {role} build custody differs")
        if entry.get("source_revision") != builds[role]["source_revision"] or entry.get("source_identity_sha256") != builds[role]["source_identity_sha256"] or entry.get("build_protocol_sha256") != builds[role]["build_protocol_sha256"]:
            fail(f"capture {role} build source/protocol identity differs")
        external_reference(output, entry.get("path"), f"capture.builds.{role}.path", bundle / builds[role]["path"])
    expected_selectors = [SELECTORS[(corpus, "control")] for corpus in CORPORA]
    if (
        capture.get("cpu") != CPU or capture.get("workers") != WORKERS
        or capture.get("corpora") != list(CORPORA)
        or capture.get("selectors") != expected_selectors
        or capture.get("lanes") != LANE_CONFIG
        or capture.get("order") != [{"role": role, "repeat": repeat} for role, repeat in ORDER]
        or capture.get("common_flags") != protocol["common_flags"]
    ):
        fail("capture manifest configuration differs from measurement protocol")
    raw_runs = list_value(capture.get("runs"), "capture.runs")
    if len(raw_runs) != 16 or capture.get("expected_run_count") != 16:
        fail("capture must contain exactly 16 runs")
    expected_keys = {(lane, corpus, role, repeat) for lane in LANES for corpus in CORPORA for role in ROLES for repeat in REPEATS}
    observed_keys = {(item.get("lane"), item.get("corpus"), item.get("role"), item.get("repeat")) for item in raw_runs}
    if observed_keys != expected_keys:
        fail("capture runs do not cover every lane/corpus/role/repeat exactly once")
    expected_indices = {
        (lane, corpus, role, repeat): index
        for index, (lane, corpus, role, repeat) in enumerate(
            ( (lane, corpus, role, repeat)
              for lane in LANES for corpus in CORPORA for role, repeat in ORDER ),
            start=1,
        )
    }
    if {
        (item.get("lane"), item.get("corpus"), item.get("role"), item.get("repeat"), item.get("index"))
        for item in raw_runs
    } != {
        (*key, index) for key, index in expected_indices.items()
    }:
        fail("capture run indices do not match the frozen measurement order")
    rows = [validate_journal(output, pinned, item, builds, protocol, protocol_sha, verifier_sha, shared) for item in raw_runs]
    order_index = {pair: index for index, pair in enumerate(ORDER)}
    rows.sort(key=lambda row: (LANES.index(row["lane"]), CORPORA.index(row["corpus"]), order_index[(row["role"], row["repeat"])]))
    for lane in LANES:
        for corpus in CORPORA:
            group = [row for row in rows if row["lane"] == lane and row["corpus"] == corpus]
            if [(row["role"], row["repeat"]) for row in group] != list(ORDER):
                fail(f"{lane}/{corpus} is not in frozen control/candidate/candidate/control order")
    chronological = sorted(rows, key=lambda row: parse_timestamp(row["started_utc"], f"{row['label']}.started_utc"))
    for previous, current in zip(chronological, chronological[1:]):
        if parse_timestamp(previous["finished_utc"], f"{previous['label']}.finished_utc") > parse_timestamp(current["started_utc"], f"{current['label']}.started_utc"):
            fail(f"capture journal intervals overlap: {previous['label']} and {current['label']}")
    input_gates: dict[str, Any] = {}
    for corpus in CORPORA:
        group = [row for row in rows if row["corpus"] == corpus]
        first_identity = group[0]["corpus_identity"]
        if any(row["corpus_identity"] != first_identity for row in group):
            fail(f"{corpus} corpus identity differs across control/candidate or lanes")
        if len({row["output_sha256"] for row in group}) != 1:
            fail(f"{corpus} output digest differs across control/candidate or lanes")
        if len({row["expected_output_bytes"] for row in group}) != 1 or len({row["sink"]["accepted_bytes"] for row in group}) != 1:
            fail(f"{corpus} output/sink byte identity differs across control/candidate or lanes")
        ceiling = first_identity["matched_owned_output_ceiling"]
        if any(row["sink"]["accepted_bytes"] > ceiling for row in group):
            fail(f"{corpus} output exceeds matched-owned sink ceiling")
        input_gates[corpus] = {
            "identity": first_identity,
            "output_sha256": next(iter({row["output_sha256"] for row in group})),
            "output_bytes": next(iter({row["sink"]["accepted_bytes"] for row in group})),
            "matched_owned_output_ceiling": ceiling,
            "same_corpus_output_sink_across_all_runs": True,
        }
    drifts: dict[str, Any] = {}
    drift_withholding: list[dict[str, Any]] = []
    for corpus in CORPORA:
        for role in ROLES:
            first = next(row for row in rows if row["lane"] == "normal" and row["corpus"] == corpus and row["role"] == role and row["repeat"] == "R1")
            second = next(row for row in rows if row["lane"] == "normal" and row["corpus"] == corpus and row["role"] == role and row["repeat"] == "R2")
            value = drift(first, second)
            drifts[f"{corpus}/{role}"] = value
            if value["withheld_statistics"]:
                drift_withholding.append({"corpus": corpus, "role": role, "statistics": value["withheld_statistics"], "reason": "repeat drift exceeded frozen normal-lane ceiling"})
    triggers: list[dict[str, Any]] = []
    normal = normal_diagnostics(rows, triggers)
    allocator = allocator_resources(rows, triggers)
    summary = {
        "change": CHANGE,
        "classification": "matched source-backed allocation-reuse experiment; descriptive evidence only",
        "claim_authorized": False,
        "performance_claim": None,
        "claim_withheld_reason": "0424 permits custody-bound current distributions, same-corpus correctness, repeat stability, and resource diagnostics; it does not authorize release latency or optimization claims",
        "capture": {"path": "capture.json", "sha256": capture_sha},
        "protocol": {"path": "../measurement-protocol.json", "sha256": protocol_sha},
        "pinned": {
            "verify_report": {"path": "../pinned/verify-report.py", "sha256": verifier_sha},
            "tools_manifest": {"path": "../pinned/tools-manifest.json", "sha256": pinned_manifest_sha},
        },
        "source": {role: {"revision": builds[role]["source_revision"], "identity_sha256": builds[role]["source_identity_sha256"], "worktree": builds[role]["source_before"].get("worktree")} for role in ROLES},
        "builds": {role: {"path": f"../{builds[role]['path']}", "sha256": builds[role]["sha256"], "build_kind": builds[role]["build_kind"], "build_protocol_sha256": builds[role]["build_protocol_sha256"], "binaries": builds[role]["binaries"], "origin_build": builds[role]["origin_build"]} for role in ROLES},
        "configuration": {"cpu": CPU, "workers": WORKERS, "lanes": LANE_CONFIG, "corpora": list(CORPORA), "roles": list(ROLES), "order": [{"role": role, "repeat": repeat} for role, repeat in ORDER]},
        "input_gates": input_gates,
        "repeat_drift": drifts,
        "normal_diagnostic_deltas": normal,
        "allocator_resource_deltas": allocator,
        "review": {"regression_threshold_percent": REGRESSION_THRESHOLD, "regression_triggers": triggers, "drift_withholding": drift_withholding, "release_latency_claim_withheld": True, "allocator_timing_comparison": "withheld", "auto_optimization_claim": False},
        "runs": rows,
        "limitations": [
            "Control and candidate are source-backed lifecycle roles; normal elapsed deltas are diagnostic and do not establish a speedup.",
            "Allocator elapsed samples are not compared. Allocator vectors are operation-scoped callback evidence with observer overhead.",
            "Logical source reads are in-process counters, not physical I/O. Whole-process RSS includes setup, validation, and teardown.",
            "The retained source identity is receipt-bound; replay does not require the original worktree or copied binaries to remain present.",
        ],
    }
    rendered = json.dumps(summary, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False) + "\n"
    table = render_table(rows, drifts, normal, allocator, triggers)
    summary_path = output / "summary.json"
    table_path = output / "result-table.md"
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
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent / "matched", help="fresh matched output root")
    parser.add_argument("--bundle", type=Path, help="bundle root containing measurement protocol/build receipts")
    parser.add_argument("--repo-root", dest="pinned_root", type=Path, help="pinned root containing verify-report.py and tools/")
    parser.add_argument("--replay", action="store_true")
    args = parser.parse_args()
    try:
        summarize(args.root, args.bundle, args.pinned_root, args.replay)
    except SummaryError as error:
        print(f"0424 summary failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
