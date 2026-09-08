"""Strict shared report arithmetic/schema checks adapted from change0432.
PPTX streaming corpus and semantic checks live in analyze.py; no old XLSX corpus gate is used.
"""
from __future__ import annotations

import argparse

import hashlib

import json

import math

from pathlib import Path

import re

import shlex

import sys

from typing import Any, Iterable

CHANGE = 476

SCHEMA_VERSION = 1

CASE = "pptx_streaming_create"

SHAPES = {"tiny": 8, "medium": 256, "large": 8_192}

SAMPLES = 30

WARMUPS = 3

WORKERS = 1

ALLOCATOR_REVISION = "serialized_region_peak_v3"

ALLOCATOR_SCOPE = "operation_global_system_allocator"

ALIGNMENT = "elapsed_ns.samples_by_elapsed_then_sample_index"

LATENCY_CLAIM = "comparable_timed_operation"

U64_MAX = (1 << 64) - 1

MAX_JSON_BYTES = 512 * 1024 * 1024

HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")

HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")

ALLOCATOR_FIELDS = (
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

SOURCE_VECTOR_FIELDS = (
    "logical_read_calls",
    "logical_read_requested_bytes",
    "logical_read_returned_bytes",
    "logical_read_largest_requested_bytes",
    "logical_read_largest_returned_bytes",
    "max_concurrent_reads",
)

SOURCE_BOUNDARY_FIELDS = ("compressed_bytes", "decompressed_bytes", "recompressed_bytes")

PROCESS_FIELDS = (
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

PARALLEL_METRIC_KEYS = {"status", "value", "scope", "reason"}

class VerificationError(ValueError):
    """A report violates the 0432 evidence contract."""

def fail(path: str, message: str) -> None:
    raise VerificationError(f"{path}: {message}")

def obj(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "expected an object")
    return value

def array(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "expected an array")
    return value

def text(value: Any, path: str, *, allow_empty: bool = False) -> str:
    if not isinstance(value, str) or (not allow_empty and not value):
        fail(path, "expected a non-empty string" if not allow_empty else "expected a string")
    return value

def boolean(value: Any, path: str) -> bool:
    if not isinstance(value, bool):
        fail(path, "expected a boolean")
    return value

def u64(value: Any, path: str, *, positive: bool = False) -> int:
    if (
        isinstance(value, bool)
        or not isinstance(value, int)
        or value < (1 if positive else 0)
        or value > U64_MAX
    ):
        fail(path, "expected a positive u64" if positive else "expected a u64")
    return value

def finite(value: Any, path: str, *, nonnegative: bool = True) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        fail(path, "expected a finite number")
    try:
        number = float(value)
    except (OverflowError, ValueError):
        fail(path, "number is outside the finite range")
    if not math.isfinite(number) or (nonnegative and number < 0):
        fail(path, "expected a finite non-negative number")
    return number

def digest(value: Any, path: str) -> str:
    value = text(value, path)
    if HEX64.fullmatch(value) is None:
        fail(path, "expected a SHA-256 hexadecimal digest")
    return value.lower()

def exact_keys(
    value: dict[str, Any], required: Iterable[str], optional: Iterable[str], path: str
) -> None:
    required_set = set(required)
    allowed = required_set | set(optional)
    missing = sorted(required_set - set(value))
    unknown = sorted(set(value) - allowed)
    if missing or unknown:
        detail = []
        if missing:
            detail.append(f"missing={missing}")
        if unknown:
            detail.append(f"unknown={unknown}")
        fail(path, "keys mismatch: " + ", ".join(detail))

def reject_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON object key: {key}")
        result[key] = value
    return result

def reject_constant(value: str) -> Any:
    raise VerificationError(f"non-finite JSON constant is not permitted: {value}")

def load_json(path: Path, label: str | None = None) -> Any:
    label = label or str(path)
    try:
        if not path.is_file():
            fail(label, "file is missing")
        if path.stat().st_size > MAX_JSON_BYTES:
            fail(label, f"file exceeds {MAX_JSON_BYTES} bytes")
        raw = path.read_bytes()
    except OSError as error:
        fail(label, f"cannot read file: {error}")
    try:
        return json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=reject_duplicate_pairs,
            parse_constant=reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, VerificationError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")

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
        raise VerificationError(f"cannot canonicalize JSON: {error}") from error

def sha256_json(value: Any) -> str:
    return hashlib.sha256(canonical(value)).hexdigest()

def check_hash_tree(value: Any, path: str) -> None:
    """Reject malformed digest fields in optional producer/catalog evidence."""

    if isinstance(value, dict):
        for key, child in value.items():
            lowered = key.lower()
            # ``canonicalization.hash`` is the algorithm label ``sha256``,
            # not a digest; only digest-shaped field names are checked here.
            if lowered.endswith("_sha256") or lowered == "sha256":
                if child is not None:
                    digest(child, f"{path}.{key}")
            check_hash_tree(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            check_hash_tree(child, f"{path}[{index}]")

def check_tool(report: dict[str, Any], mode: str) -> dict[str, Any]:
    tool = obj(report.get("tool"), "report.tool")
    required = {
        "name",
        "version",
        "binary",
        "profile",
        "target_os",
        "target_arch",
        "instrumentation",
    }
    if mode == "allocator":
        required.add("allocator_counter_revision")
    exact_keys(tool, required, (), "report.tool")
    expected = {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": (
            "system_allocator_operation_scoped" if mode == "allocator" else "none"
        ),
    }
    for key, value in expected.items():
        if tool.get(key) != value:
            fail(f"report.tool.{key}", f"must be {value!r}")
    expected_binary = "litchi-perf-baseline-alloc" if mode == "allocator" else "litchi-perf-baseline"
    if tool.get("binary") != expected_binary:
        fail("report.tool.binary", f"must be {expected_binary!r}")
    if mode == "allocator" and tool.get("allocator_counter_revision") != ALLOCATOR_REVISION:
        fail("report.tool.allocator_counter_revision", f"must be {ALLOCATOR_REVISION!r}")
    return tool

def check_binary_identity(report: dict[str, Any]) -> None:
    identity = obj(report.get("binary_identity"), "report.binary_identity")
    exact_keys(
        identity,
        ("path", "binary_sha256", "binary_bytes", "mode_bits", "executable", "profile"),
        (),
        "report.binary_identity",
    )
    path = text(identity["path"], "report.binary_identity.path")
    if not Path(path).is_absolute():
        fail("report.binary_identity.path", "must be absolute provenance, not a relative path")
    digest(identity["binary_sha256"], "report.binary_identity.binary_sha256")
    u64(identity["binary_bytes"], "report.binary_identity.binary_bytes", positive=True)
    mode_bits = identity["mode_bits"]
    if mode_bits is None:
        fail("report.binary_identity.mode_bits", "must be present on the Linux capture host")
    if isinstance(mode_bits, bool) or not isinstance(mode_bits, int) or not 0 <= mode_bits <= 0o7777:
        fail("report.binary_identity.mode_bits", "must be executable Unix permission bits")
    if mode_bits & 0o111 == 0:
        fail("report.binary_identity.mode_bits", "does not identify an executable")
    if identity["executable"] is not True:
        fail("report.binary_identity.executable", "must be true")
    if identity["profile"] != "release":
        fail("report.binary_identity.profile", "must match the release protocol")

def normalized_rustflags(value: str, path: str) -> list[str]:
    """Normalize Rust's split and attached ``-C`` spellings to one token form."""

    try:
        tokens = shlex.split(value, posix=True)
    except ValueError as error:
        fail(path, f"invalid shell-style RUSTFLAGS: {error}")
    normalized: list[str] = []
    index = 0
    while index < len(tokens):
        token = tokens[index]
        if token == "-C":
            if index + 1 >= len(tokens):
                fail(path, "-C is missing its codegen option")
            normalized.append(f"-C{tokens[index + 1]}")
            index += 2
        else:
            normalized.append(token)
            index += 1
    return normalized

def check_environment(report: dict[str, Any], mode: str) -> None:
    environment = obj(report.get("environment"), "report.environment")
    fields = (
        "rustc_version",
        "git_revision",
        "git_worktree_dirty",
        "logical_cpus_available",
        "allocator",
        "rustflags",
        "cargo_build_target",
        "perf_event_paranoid",
        "os",
        "kernel",
        "cpu_model",
        "total_memory_bytes",
        "page_size_bytes",
        "filesystem_type",
        "source_destination_same_device",
        "cpu_affinity",
        "storage_identifier",
    )
    exact_keys(environment, fields, (), "report.environment")
    text(environment["rustc_version"], "report.environment.rustc_version")
    revision = text(environment["git_revision"], "report.environment.git_revision")
    if HEX40.fullmatch(revision) is None:
        fail("report.environment.git_revision", "must be a 40-character revision")
    # The producer reports the complete checkout, including intentionally
    # retained evidence and user-local files.  Whether tracked source stayed
    # unchanged is bound by the outer capture custody receipt; this portable
    # report checker only validates the producer's boolean field.
    boolean(environment["git_worktree_dirty"], "report.environment.git_worktree_dirty")
    u64(environment["logical_cpus_available"], "report.environment.logical_cpus_available", positive=True)
    allocator = "CountingSystemAllocator(std::alloc::System)" if mode == "allocator" else "Rust system allocator"
    if environment["allocator"] != allocator:
        fail("report.environment.allocator", f"must be {allocator!r}")
    flags = environment["rustflags"]
    if flags is not None:
        flags = text(flags, "report.environment.rustflags", allow_empty=True)
        normalized = normalized_rustflags(flags, "report.environment.rustflags")
        pointer_flags = [token for token in normalized if token.startswith("-Cforce-frame-pointers=")]
        if pointer_flags != ["-Cforce-frame-pointers=yes"]:
            fail("report.environment.rustflags", "must bind frame-pointer capture flags")
    for key in ("cargo_build_target", "perf_event_paranoid", "os", "kernel", "cpu_model", "filesystem_type", "cpu_affinity", "storage_identifier"):
        if environment[key] is not None:
            text(environment[key], f"report.environment.{key}", allow_empty=True)
    if environment["os"] != "linux":
        fail("report.environment.os", "must be linux for this protocol")
    if environment["cpu_affinity"] != "2":
        fail("report.environment.cpu_affinity", "must bind the protocol CPU 2")
    for key in ("total_memory_bytes", "page_size_bytes"):
        if environment[key] is not None:
            u64(environment[key], f"report.environment.{key}", positive=True)
    for key in ("source_destination_same_device",):
        if environment[key] is not None:
            boolean(environment[key], f"report.environment.{key}")

def check_configuration(report: dict[str, Any], shape: str) -> None:
    config = obj(report.get("configuration"), "report.configuration")
    fields = (
        "samples_per_case",
        "warmup_iterations_per_case",
        "filesystem_cache_states",
        "filesystem_fresh_child_per_sample",
        "filesystem_process_isolated",
        "filesystem_root_selected",
        "cases",
        "corpus_shapes",
        "payload_kinds",
        "writer_shapes",
        "xlsx_shapes",
        "xlsb_shapes",
        "xlsx_cell_crud_shapes",
        "xlsx_row_visibility_shapes",
        "semantic_shapes",
        "rtf_variants",
        "range_simulation",
        "execution_workers",
        "opc_cache_lock_diagnostics",
    )
    exact_keys(config, fields, (), "report.configuration")
    if config["samples_per_case"] != SAMPLES or config["warmup_iterations_per_case"] != WARMUPS:
        fail("report.configuration", "sample and warmup counts do not match protocol")
    expected_lists = {
        # parse_options() defaults to warm plus cold-requested even though
        # this selector has no filesystem child; the configuration is still
        # serialized and must retain that producer identity.
        "filesystem_cache_states": ["warm", "cold-requested"],
        "cases": [CASE],
        "corpus_shapes": ["tiny", "many-small", "few-large", "wide-root"],
        "payload_kinds": ["compressible", "incompressible"],
        "writer_shapes": ["tiny", "large", "payload-heavy"],
        "xlsx_shapes": ["tiny", "medium", "dense-wide"],
        "xlsb_shapes": ["tiny", "medium", "large", "sparse"],
        "xlsx_cell_crud_shapes": ["medium", "dense-sparse"],
        "xlsx_row_visibility_shapes": ["medium", "large"],
        "semantic_shapes": [shape],
        "rtf_variants": ["plain"],
        "execution_workers": [WORKERS],
    }
    for key, expected in expected_lists.items():
        if config[key] != expected:
            fail(f"report.configuration.{key}", f"must be exactly {expected!r}")
    for key in expected_lists:
        array(config[key], f"report.configuration.{key}")
        for index, value in enumerate(config[key]):
            if key == "execution_workers":
                u64(value, f"report.configuration.{key}[{index}]", positive=True)
            else:
                text(value, f"report.configuration.{key}[{index}]")
    for key in ("filesystem_fresh_child_per_sample", "filesystem_process_isolated"):
        if config[key] is not True:
            fail(f"report.configuration.{key}", "must be true")
    if config["filesystem_root_selected"] is not False:
        fail("report.configuration.filesystem_root_selected", "must be false")
    simulation = obj(config["range_simulation"], "report.configuration.range_simulation")
    exact_keys(
        simulation,
        ("fixed_latency_us", "request_overhead_us", "bandwidth_bytes_per_second", "max_physical_range_bytes"),
        (),
        "report.configuration.range_simulation",
    )
    expected_simulation = {
        "fixed_latency_us": 100,
        "request_overhead_us": 25,
        "bandwidth_bytes_per_second": 50 * 1024 * 1024,
        "max_physical_range_bytes": 4 * 1024,
    }
    for key in expected_simulation:
        u64(simulation[key], f"report.configuration.range_simulation.{key}")
    if simulation != expected_simulation:
        fail("report.configuration.range_simulation", "does not match harness defaults")
    if config["opc_cache_lock_diagnostics"] is not False:
        fail("report.configuration.opc_cache_lock_diagnostics", "must be false")

def parallel_metric(value: Any, path: str, *, status: str | None = None, value_kind: str | None = None) -> dict[str, Any]:
    metric = obj(value, path)
    exact_keys(metric, ("status", "scope"), ("value", "reason"), path)
    actual_status = text(metric["status"], f"{path}.status")
    if actual_status not in {"measured", "not_applicable", "unavailable"}:
        fail(f"{path}.status", "unknown parallel metric status")
    text(metric["scope"], f"{path}.scope")
    has_value = "value" in metric
    has_reason = "reason" in metric
    if actual_status == "measured" and not has_value:
        fail(path, "measured metric must carry a value")
    if actual_status != "measured" and has_value:
        fail(path, "unavailable/not-applicable metric must omit value")
    if actual_status != "measured" and not has_reason:
        fail(path, "unavailable/not-applicable metric must carry a reason")
    if has_reason:
        text(metric["reason"], f"{path}.reason")
    if status is not None and actual_status != status:
        fail(f"{path}.status", f"must be {status!r}")
    if value_kind == "u64_list":
        values = array(metric.get("value"), f"{path}.value")
        for index, item in enumerate(values):
            u64(item, f"{path}.value[{index}]")
    elif value_kind == "u64":
        u64(metric.get("value"), f"{path}.value")
    return metric

def check_parallel(report: dict[str, Any], archive_hash: str) -> None:
    value = obj(report.get("parallel_metrics"), "report.parallel_metrics")
    exact_keys(
        value,
        ("schema_version", "scope", "claim", "configured_worker_budget", "observed_process_thread_count", "cases"),
        (),
        "report.parallel_metrics",
    )
    if u64(value["schema_version"], "report.parallel_metrics.schema_version") != 1 or value["scope"] != "explicit_local_execution_only" or value["claim"] != "descriptive":
        fail("report.parallel_metrics", "schema or claim identity differs")
    parallel_metric(value["configured_worker_budget"], "report.parallel_metrics.configured_worker_budget", status="measured", value_kind="u64_list")
    if value["configured_worker_budget"]["value"] != [WORKERS] or value["configured_worker_budget"]["scope"] != "configuration.execution_workers":
        fail("report.parallel_metrics.configured_worker_budget", "does not bind workers=[1]")
    observed = parallel_metric(value["observed_process_thread_count"], "report.parallel_metrics.observed_process_thread_count", status="unavailable")
    if observed["scope"] != "process_thread_count":
        fail("report.parallel_metrics.observed_process_thread_count.scope", "has wrong scope")
    cases = array(value["cases"], "report.parallel_metrics.cases")
    if len(cases) != 1:
        fail("report.parallel_metrics.cases", "must contain one streaming case")
    case = obj(cases[0], "report.parallel_metrics.cases[0]")
    exact_keys(case, ("case", "corpus_sha256", "configured_worker_count", "observed_local_worker_count", "deterministic_task_count", "deterministic_chunk_count", "lock_wait_ns"), (), "report.parallel_metrics.cases[0]")
    if case["case"] != CASE or digest(case["corpus_sha256"], "report.parallel_metrics.cases[0].corpus_sha256") != archive_hash:
        fail("report.parallel_metrics.cases[0]", "case or corpus identity differs")
    for key in ("configured_worker_count", "observed_local_worker_count", "deterministic_task_count", "deterministic_chunk_count"):
        metric = parallel_metric(case[key], f"report.parallel_metrics.cases[0].{key}", status="not_applicable")
        if key == "configured_worker_count" and metric["scope"] != "result.execution.worker_count":
            fail(f"report.parallel_metrics.cases[0].{key}.scope", "has wrong scope")
    lock = parallel_metric(case["lock_wait_ns"], "report.parallel_metrics.cases[0].lock_wait_ns", status="unavailable")
    if lock["scope"] != "lock_wait_ns":
        fail("report.parallel_metrics.cases[0].lock_wait_ns.scope", "has wrong scope")

def check_elapsed(value: Any, path: str = "report.results[0].elapsed_ns", count: int = SAMPLES) -> tuple[list[int], list[int], dict[str, Any]]:
    if count <= 0:
        fail(path, "sample count must be positive")
    elapsed = obj(value, path)
    exact_keys(elapsed, ("unit", "samples", "sample_order", "min", "p50", "p95", "p99", "max", "mean", "standard_deviation", "confidence_interval_95"), (), path)
    if elapsed["unit"] != "ns":
        fail(f"{path}.unit", "must be ns")
    raw = array(elapsed["samples"], f"{path}.samples")
    if len(raw) != count:
        fail(f"{path}.samples", f"must contain exactly {count} raw samples")
    samples = [u64(item, f"{path}.samples[{index}]", positive=True) for index, item in enumerate(raw)]
    if samples != sorted(samples):
        fail(f"{path}.samples", "must be sorted by elapsed time")
    order_raw = array(elapsed["sample_order"], f"{path}.sample_order")
    if len(order_raw) != count:
        fail(f"{path}.sample_order", "must align with all retained samples")
    order = [u64(item, f"{path}.sample_order[{index}]") for index, item in enumerate(order_raw)]
    if sorted(order) != list(range(count)):
        fail(f"{path}.sample_order", "must be a complete original-sample permutation")
    for index in range(1, count):
        if samples[index] == samples[index - 1] and order[index] <= order[index - 1]:
            fail(f"{path}.sample_order", "must increase across tied elapsed samples")

    mean = 0.0
    squared = 0.0
    for index, sample in enumerate(samples):
        current = float(sample)
        sample_number = float(index + 1)
        delta = current - mean
        next_mean = mean + delta / sample_number
        squared += delta * (current - next_mean)
        mean = next_mean
    standard_deviation = math.sqrt(squared / float(count - 1)) if count > 1 else 0.0
    t = (
        2.045
        if count - 1 == 29
        else 1.959963984540054
    )
    margin = t * standard_deviation / math.sqrt(float(count)) if count > 1 else 0.0
    midpoint = samples[(count - 1) // 2] // 2 + samples[count // 2] // 2 + (samples[(count - 1) // 2] % 2 + samples[count // 2] % 2) // 2
    nearest = lambda percentile: samples[min(((percentile * count + 99) // 100) - 1, count - 1)]
    expected_int = {"min": samples[0], "p50": midpoint, "p95": nearest(95), "p99": nearest(99), "max": samples[-1]}
    for key, expected in expected_int.items():
        if u64(elapsed[key], f"{path}.{key}") != expected:
            fail(f"{path}.{key}", f"does not match retained samples ({expected})")
    for key, expected in (("mean", mean), ("standard_deviation", standard_deviation)):
        actual = finite(elapsed[key], f"{path}.{key}")
        if not math.isclose(actual, expected, rel_tol=1e-12, abs_tol=1e-12):
            fail(f"{path}.{key}", f"does not match retained samples ({expected})")
    confidence = obj(elapsed["confidence_interval_95"], f"{path}.confidence_interval_95")
    exact_keys(confidence, ("method", "lower", "upper"), (), f"{path}.confidence_interval_95")
    if confidence["method"] != "two-sided Student's t interval for the mean":
        fail(f"{path}.confidence_interval_95.method", "does not match the harness")
    lower = finite(confidence["lower"], f"{path}.confidence_interval_95.lower")
    upper = finite(confidence["upper"], f"{path}.confidence_interval_95.upper")
    if not math.isclose(lower, max(mean - margin, 0.0), rel_tol=1e-12, abs_tol=1e-12) or not math.isclose(upper, mean + margin, rel_tol=1e-12, abs_tol=1e-12):
        fail(f"{path}.confidence_interval_95", "does not match the Student's t calculation")
    return samples, order, {"min": expected_int["min"], "p50": expected_int["p50"], "p95": expected_int["p95"], "p99": expected_int["p99"], "max": expected_int["max"], "mean": mean, "standard_deviation": standard_deviation, "confidence_interval_95": {"lower": max(mean - margin, 0.0), "upper": mean + margin}}

def check_metric_vector(value: Any, path: str, count: int, *, expected_status: str | None = None, expected_scope: str | None = None, values_required: bool | None = None) -> list[int] | None:
    vector = obj(value, path)
    exact_keys(vector, ("status", "scope"), ("values",), path)
    status = text(vector["status"], f"{path}.status")
    if status not in {"measured", "not_applicable", "unavailable", "overflow"}:
        fail(f"{path}.status", "unknown metric status")
    scope = text(vector["scope"], f"{path}.scope")
    if expected_status is not None and status != expected_status:
        fail(f"{path}.status", f"must be {expected_status!r}")
    if expected_scope is not None and scope != expected_scope:
        fail(f"{path}.scope", f"must be {expected_scope!r}")
    has_values = "values" in vector
    if values_required is None:
        values_required = status == "measured"
    if values_required != has_values:
        fail(path, "measured values presence does not match metric status")
    if not has_values:
        return None
    values = array(vector["values"], f"{path}.values")
    if len(values) != count:
        fail(f"{path}.values", f"must contain exactly {count} values")
    result = [u64(item, f"{path}.values[{index}]") for index, item in enumerate(values)]
    return result

def check_operation_metrics(value: Any, mode: str, elapsed: list[int], order: list[int], sink: dict[str, Any], path: str = "report.results[0].operation_metrics") -> dict[str, Any]:
    operation = obj(value, path)
    required = {"sample_count", "sample_indices", "alignment", "latency_claim", "source", "process", "sink", "publication", "materialization", "cfb_phases"}
    if mode == "allocator":
        required.add("allocation")
    exact_keys(operation, required, (), path)
    if operation["sample_count"] != SAMPLES:
        fail(f"{path}.sample_count", "must match the protocol")
    sample_indices = [u64(item, f"{path}.sample_indices[{index}]") for index, item in enumerate(array(operation["sample_indices"], f"{path}.sample_indices"))]
    if len(sample_indices) != SAMPLES or sample_indices != order:
        fail(f"{path}.sample_indices", "must equal elapsed_ns.sample_order")
    if operation["alignment"] != ALIGNMENT or operation["latency_claim"] != LATENCY_CLAIM:
        fail(path, "alignment or latency claim differs from the streaming protocol")

    source = obj(operation["source"], f"{path}.source")
    exact_keys(source, ("status", "counter_scope", *SOURCE_VECTOR_FIELDS, "logical_read_pattern", *SOURCE_BOUNDARY_FIELDS), (), f"{path}.source")
    if source["status"] != "not_applicable" or source["counter_scope"] != "not_applicable_in_process_sink":
        fail(f"{path}.source", "PPTX in-process sink source must be explicitly not applicable")
    source_scopes = {
        **{key: "operation_logical_read_at" for key in SOURCE_VECTOR_FIELDS},
        "logical_read_pattern": "operation_logical_read_at_range_order_not_physical_io",
        "compressed_bytes": "unavailable_read_at_has_no_compressed_member_boundary",
        "decompressed_bytes": "unavailable_read_at_has_no_decompressed_byte_boundary",
        "recompressed_bytes": "unavailable_atomic_save_has_no_recompressed_byte_boundary",
    }
    for key, scope in source_scopes.items():
        vector = source[key]
        if key == "logical_read_pattern":
            pattern = obj(vector, f"{path}.source.{key}")
            exact_keys(pattern, ("status", "scope"), (), f"{path}.source.{key}")
            if pattern["status"] != "not_applicable" or pattern["scope"] != scope:
                fail(f"{path}.source.{key}", "must be unavailable as an in-process sink observation")
        else:
            check_metric_vector(vector, f"{path}.source.{key}", SAMPLES, expected_status="not_applicable", expected_scope=scope, values_required=False)

    process = obj(operation["process"], f"{path}.process")
    exact_keys(process, ("status", *PROCESS_FIELDS), (), f"{path}.process")
    process_status = text(process["status"], f"{path}.process.status")
    if process_status not in {"measured", "unavailable"}:
        fail(f"{path}.process.status", "must be measured or explicitly unavailable")
    process_scopes = {
        **{key: "procfs_in_process_operation_delta_including_procfs_probe_overhead" for key in PROCESS_FIELDS if key not in {"rss_delta_bytes", "peak_rss_bytes"}},
        "rss_delta_bytes": "procfs_in_process_rss_delta_including_procfs_probe_overhead",
        "peak_rss_bytes": "process_lifetime_high_water_after_not_operation_peak",
    }
    for key in PROCESS_FIELDS:
        check_metric_vector(process[key], f"{path}.process.{key}", SAMPLES, expected_status=process_status, expected_scope=process_scopes[key])
        if key == "clock_ticks_per_second" and process_status == "measured":
            values = process[key]["values"]
            assert isinstance(values, list)
            if any(value == 0 for value in values):
                fail(f"{path}.process.clock_ticks_per_second", "clock frequency must be positive")

    sink_metrics = obj(operation["sink"], f"{path}.sink")
    exact_keys(sink_metrics, ("status", "output_bytes", "write_status", *SINK_VECTOR_FIELDS, "write_size_buckets"), (), f"{path}.sink")
    if sink_metrics["status"] != "not_applicable" or sink_metrics["write_status"] != "measured":
        fail(f"{path}.sink", "streaming discard sink statuses are inconsistent")
    check_metric_vector(sink_metrics["output_bytes"], f"{path}.sink.output_bytes", SAMPLES, expected_status="not_applicable", expected_scope="post_operation_output_length_not_sink_write_volume", values_required=False)
    for key, scope in (("accepted_bytes", "logical_sink_accepted_write_bytes"), ("write_calls", "logical_sink_accepted_write_calls"), ("largest_write", "logical_sink_largest_accepted_write")):
        values = check_metric_vector(sink_metrics[key], f"{path}.sink.{key}", SAMPLES, expected_status="measured", expected_scope=scope)
        assert values is not None
        expected = sink[key]
        if any(item != expected for item in values):
            fail(f"{path}.sink.{key}.values", "does not repeat the deterministic top-level sink summary")
    buckets = obj(sink_metrics["write_size_buckets"], f"{path}.sink.write_size_buckets")
    exact_keys(buckets, ("status", *SINK_BUCKET_FIELDS), (), f"{path}.sink.write_size_buckets")
    if buckets["status"] != "measured":
        fail(f"{path}.sink.write_size_buckets.status", "must be measured")
    for key in SINK_BUCKET_FIELDS:
        values = check_metric_vector(buckets[key], f"{path}.sink.write_size_buckets.{key}", SAMPLES, expected_status="measured", expected_scope="logical_sink_accepted_write_size_bucket_counts")
        assert values is not None
        if any(item != sink["buckets"][key] for item in values):
            fail(f"{path}.sink.write_size_buckets.{key}.values", "does not repeat the deterministic top-level sink bucket")

    publication = obj(operation["publication"], f"{path}.publication")
    exact_keys(publication, ("status", "changed_spans", "published_bytes"), (), f"{path}.publication")
    if publication["status"] != "not_applicable":
        fail(f"{path}.publication.status", "must be not_applicable")
    for key in ("changed_spans", "published_bytes"):
        check_metric_vector(publication[key], f"{path}.publication.{key}", SAMPLES, expected_status="not_applicable", expected_scope="logical_publication_counter", values_required=False)
    materialization = obj(operation["materialization"], f"{path}.materialization")
    exact_keys(materialization, ("status", "opc_parts"), (), f"{path}.materialization")
    if materialization["status"] != "not_applicable":
        fail(f"{path}.materialization.status", "must be not_applicable")
    check_metric_vector(materialization["opc_parts"], f"{path}.materialization.opc_parts", SAMPLES, expected_status="not_applicable", expected_scope="logical_materialization_counter", values_required=False)
    phases = obj(operation["cfb_phases"], f"{path}.cfb_phases")
    exact_keys(phases, ("status", "open", "plan", "atomic_publication"), (), f"{path}.cfb_phases")
    if phases["status"] != "not_applicable":
        fail(f"{path}.cfb_phases.status", "must be not_applicable")
    for phase_name in ("open", "plan", "atomic_publication"):
        phase = obj(phases[phase_name], f"{path}.cfb_phases.{phase_name}")
        exact_keys(phase, ("elapsed_ns", "logical_read_calls", "logical_read_requested_bytes", "logical_read_returned_bytes"), (), f"{path}.cfb_phases.{phase_name}")
        check_metric_vector(phase["elapsed_ns"], f"{path}.cfb_phases.{phase_name}.elapsed_ns", SAMPLES, expected_status="not_applicable", expected_scope="timed_cfb_phase_elapsed_ns", values_required=False)
        for key in ("logical_read_calls", "logical_read_requested_bytes", "logical_read_returned_bytes"):
            check_metric_vector(phase[key], f"{path}.cfb_phases.{phase_name}.{key}", SAMPLES, expected_status="not_applicable", expected_scope="timed_cfb_phase_logical_read_at", values_required=False)

    allocation: dict[str, Any] | None = operation.get("allocation")
    if mode == "allocator":
        allocation = obj(allocation, f"{path}.allocation")
        check_allocation(allocation, elapsed, f"{path}.allocation")
    elif allocation is not None:
        allocation = obj(allocation, f"{path}.allocation")
        exact_keys(allocation, ("status", "scope", *ALLOCATOR_FIELDS), (), f"{path}.allocation")
        if allocation["status"] != "unavailable" or allocation["scope"] != ALLOCATOR_SCOPE:
            fail(f"{path}.allocation", "normal instrumentation cannot publish measured allocator evidence")
        for key in ALLOCATOR_FIELDS:
            check_metric_vector(allocation[key], f"{path}.allocation.{key}", SAMPLES, expected_status="unavailable", expected_scope=ALLOCATOR_SCOPE, values_required=False)
    return {"process": process, "allocation": allocation}

def check_allocation(allocation: dict[str, Any], elapsed: list[int], path: str) -> None:
    exact_keys(allocation, ("status", "scope", *ALLOCATOR_FIELDS), (), path)
    if allocation["status"] != "measured" or allocation["scope"] != ALLOCATOR_SCOPE:
        fail(path, "allocator lane must be measured with the v3 system-allocator scope")
    vectors: dict[str, list[int]] = {}
    for key in ALLOCATOR_FIELDS:
        values = check_metric_vector(allocation[key], f"{path}.{key}", SAMPLES, expected_status="measured", expected_scope=ALLOCATOR_SCOPE)
        assert values is not None
        vectors[key] = values
    for index in range(SAMPLES):
        before = vectors["live_bytes_before"][index]
        after = vectors["live_bytes_after"][index]
        allocated = vectors["allocated_bytes"][index]
        deallocated = vectors["deallocated_bytes"][index]
        peak_before = vectors["peak_live_bytes_before"][index]
        peak_after = vectors["peak_live_bytes_after"][index]
        region = vectors["region_peak_live_bytes"][index]
        if vectors["failed_allocation_calls"][index] != 0:
            fail(f"{path}.failed_allocation_calls.values[{index}]", "timed allocation lane contains a failed allocation")
        if allocated + before < deallocated or after != before + allocated - deallocated:
            fail(f"{path}[{index}]", "checked allocation byte delta does not balance live bytes")
        if after != before:
            fail(f"{path}.live_bytes_after.values[{index}]", "timed operation must have zero live-byte exit delta")
        if vectors["allocation_calls"][index] < vectors["reallocation_calls"][index]:
            fail(f"{path}.allocation_calls.values[{index}]", "allocation calls cannot be fewer than reallocations")
        if peak_before < before or peak_after < peak_before:
            fail(f"{path}[{index}]", "absolute allocator high-water counters moved backwards")
        if region < before or region < after or region > peak_after:
            fail(f"{path}.region_peak_live_bytes.values[{index}]", "region peak violates live endpoint/high-water bounds")

def check_catalog_sidecar(report_path: Path, report: dict[str, Any], reference: dict[str, Any]) -> None:
    catalog_path = report_path.with_name("corpus-catalog.json")
    if not catalog_path.is_file():
        fail("report.corpus_catalog", f"catalog sidecar {catalog_path.name!r} is missing")
    catalog = obj(load_json(catalog_path, str(catalog_path)), "corpus_catalog_file")
    check_hash_tree(catalog, "corpus_catalog_file")
    exact_keys(catalog, ("manifest_version", "manifest_kind", "catalog_id", "canonicalization", "catalog_sha256", "content_set_sha256", "build", "corpora", "case_bindings"), (), "corpus_catalog_file")
    if catalog["manifest_version"] != 2 or catalog["manifest_kind"] != "corpus-catalog" or catalog["catalog_id"] != "litchi-perf-corpus-v2":
        fail("corpus_catalog_file", "catalog identity differs from the schema-2 producer")
    canonicalization = obj(catalog["canonicalization"], "corpus_catalog_file.canonicalization")
    exact_keys(canonicalization, ("algorithm", "hash"), (), "corpus_catalog_file.canonicalization")
    if canonicalization != {"algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"}:
        fail("corpus_catalog_file.canonicalization", "does not identify canonical hashing")
    catalog_hash = digest(catalog["catalog_sha256"], "corpus_catalog_file.catalog_sha256")
    content_hash = digest(catalog["content_set_sha256"], "corpus_catalog_file.content_set_sha256")
    if reference["manifest_version"] != catalog["manifest_version"] or reference["catalog_id"] != catalog["catalog_id"] or reference["catalog_sha256"] != catalog_hash or reference["content_set_sha256"] != content_hash:
        fail("report.corpus_catalog", "does not match the catalog sidecar")
    without_catalog_hash = dict(catalog)
    del without_catalog_hash["catalog_sha256"]
    if sha256_json(without_catalog_hash) != catalog_hash:
        fail("corpus_catalog_file.catalog_sha256", "does not match canonical catalog content")
    corpora = array(catalog["corpora"], "corpus_catalog_file.corpora")
    if not corpora:
        fail("corpus_catalog_file.corpora", "must retain the generated corpus")
    expected_archive = report["results"][0]["corpus"]["archive_sha256"].lower()
    matched = False
    content_rows = []
    for index, corpus in enumerate(corpora):
        row = obj(corpus, f"corpus_catalog_file.corpora[{index}]")
        if "id" not in row or "bytes" not in row or "members" not in row:
            fail(f"corpus_catalog_file.corpora[{index}]", "is missing content identity fields")
        bytes_row = obj(row["bytes"], f"corpus_catalog_file.corpora[{index}].bytes")
        if digest(bytes_row.get("archive_sha256"), f"corpus_catalog_file.corpora[{index}].bytes.archive_sha256") == expected_archive:
            matched = True
        members = obj(row["members"], f"corpus_catalog_file.corpora[{index}].members")
        items = array(members.get("items"), f"corpus_catalog_file.corpora[{index}].members.items")
        member_rows = []
        for item_index, item in enumerate(items):
            item = obj(item, f"corpus_catalog_file.corpora[{index}].members.items[{item_index}]")
            member_rows.append({"ordinal": u64(item.get("ordinal"), f"corpus_catalog_file.corpora[{index}].members.items[{item_index}].ordinal"), "name": text(item.get("name"), f"corpus_catalog_file.corpora[{index}].members.items[{item_index}].name"), "sha256": digest(item.get("sha256"), f"corpus_catalog_file.corpora[{index}].members.items[{item_index}].sha256")})
        content_rows.append({"id": text(row["id"], f"corpus_catalog_file.corpora[{index}].id"), "archive_sha256": digest(bytes_row["archive_sha256"], f"corpus_catalog_file.corpora[{index}].bytes.archive_sha256"), "members": member_rows})
    if not matched:
        fail("corpus_catalog_file.corpora", "does not contain the report archive identity")
    bindings = array(catalog["case_bindings"], "corpus_catalog_file.case_bindings")
    binding_rows = []
    for index, binding in enumerate(bindings):
        row = obj(binding, f"corpus_catalog_file.case_bindings[{index}]")
        exact_keys(row, ("case", "corpus_id", "legacy_name", "legacy_archive_sha256", "role"), (), f"corpus_catalog_file.case_bindings[{index}]")
        binding_rows.append({"case": text(row["case"], f"corpus_catalog_file.case_bindings[{index}].case"), "corpus_id": text(row["corpus_id"], f"corpus_catalog_file.case_bindings[{index}].corpus_id"), "role": text(row["role"], f"corpus_catalog_file.case_bindings[{index}].role")})
    expected_content = {"corpora": content_rows, "case_bindings": binding_rows}
    if sha256_json(expected_content) != content_hash:
        fail("corpus_catalog_file.content_set_sha256", "does not match canonical content identities")
    if not any(row["case"] == CASE and row["role"] == "timed" for row in binding_rows):
        fail("corpus_catalog_file.case_bindings", "does not bind the streaming case")


# Public 0476 adapter.  The strict 0474 arithmetic above is copied into this
# bundle so a portable replay does not import an earlier result directory.
ReportError = VerificationError


def read_json(path: Path) -> Any:
    return load_json(path, str(path))


def file_digest(path: Path) -> tuple[str, int]:
    digest_value = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest_value.update(block)
                size += len(block)
    except OSError as error:
        raise VerificationError(f"cannot hash {path}: {error}") from error
    return digest_value.hexdigest(), size


def vector_stats(values: Iterable[int | float], path: str) -> dict[str, Any]:
    checked = [finite(value, f"{path}[{index}]") for index, value in enumerate(values)]
    if not checked:
        fail(path, "cannot be empty")
    ordered = sorted(checked)
    count = len(ordered)
    return {
        "min": ordered[0],
        "p50": (ordered[(count - 1) // 2] + ordered[count // 2]) / 2,
        "p95": ordered[min(((95 * count + 99) // 100) - 1, count - 1)],
        "p99": ordered[min(((99 * count + 99) // 100) - 1, count - 1)],
        "max": ordered[-1],
        "mean": math.fsum(ordered) / count,
        "samples": list(values),
    }


def percent_delta(control: float, candidate: float) -> float:
    if control == 0:
        return 0.0 if candidate == 0 else math.inf
    return ((candidate / control) - 1.0) * 100.0


def compare_vectors(control: Iterable[int | float], candidate: Iterable[int | float], path: str) -> dict[str, Any]:
    left = list(control)
    right = list(candidate)
    if not left or len(left) != len(right):
        fail(path, "comparison vectors must be equally sized and non-empty")
    left_stats = vector_stats(left, f"{path}.control")
    right_stats = vector_stats(right, f"{path}.candidate")
    return {
        "control": left_stats,
        "candidate": right_stats,
        "candidate_vs_control_percent": percent_delta(float(left_stats["mean"]), float(right_stats["mean"])),
    }


def parse_counter_text(text: str, expected_events: Iterable[str]) -> dict[str, dict[str, Any]]:
    expected = list(expected_events)
    if not expected or len(set(expected)) != len(expected):
        raise ReportError("counter protocol contains duplicate or empty events")
    parsed: dict[str, dict[str, Any]] = {}
    for line_number, line in enumerate(text.splitlines(), 1):
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        fields = stripped.split(";")
        if len(fields) != 7 or fields[1] or fields[5] or fields[6]:
            fail(f"counter line {line_number}", "must use the seven-column perf stat CSV form")
        event = fields[2].strip()
        if event not in expected or event in parsed:
            fail(f"counter line {line_number}", f"unknown or duplicate event {event!r}")
        runtime_text = fields[3].strip().replace(",", "")
        try:
            runtime = int(runtime_text)
        except ValueError as error:
            raise ReportError(f"counter line {line_number}: event runtime is not an integer") from error
        if runtime < 0:
            fail(f"counter line {line_number}", "event runtime is negative")
        try:
            running_percent = float(fields[4].strip())
        except ValueError as error:
            raise ReportError(f"counter line {line_number}: running percentage is not numeric") from error
        if not math.isfinite(running_percent) or not 0 <= running_percent <= 100:
            fail(f"counter line {line_number}", "running percentage is outside 0..100")
        raw_count = fields[0].strip().replace(",", "")
        if raw_count in {"<not supported>", "<not counted>"}:
            parsed[event] = {
                "status": "not_supported" if raw_count == "<not supported>" else "not_counted",
                "count": None,
                "event_runtime_ns": runtime,
                "running_percent": running_percent,
            }
        else:
            try:
                count = int(raw_count)
            except ValueError as error:
                raise ReportError(f"counter line {line_number}: count is not an integer") from error
            if count < 0:
                fail(f"counter line {line_number}", "count is negative")
            parsed[event] = {
                "status": "measured",
                "count": count,
                "event_runtime_ns": runtime,
                "running_percent": running_percent,
            }
    if set(parsed) != set(expected):
        fail("counters", f"event set differs: expected {expected!r}, got {sorted(parsed)!r}")
    return parsed


def _strict_corpus(row: dict[str, Any], shape: str, path: str) -> dict[str, Any]:
    corpus = obj(row.get("corpus"), f"{path}.corpus")
    exact_keys(corpus, ("name", "generator", "package_format", "shape", "payload_kind", "compression", "entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "archive_sha256", "target_entry", "target_payload_bytes", "target_payload_sha256", "xlsx"), (), f"{path}.corpus")
    if corpus["shape"] != shape or corpus["compression"] != "deflate" or corpus["package_format"] != "PPTX/OOXML/ZIP":
        fail(f"{path}.corpus", "shape, package, or compression identity differs")
    for key in ("name", "generator", "package_format", "shape", "payload_kind", "compression", "target_entry"):
        text(corpus[key], f"{path}.corpus.{key}")
    for key in ("entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        u64(corpus[key], f"{path}.corpus.{key}")
    digest(corpus["archive_sha256"], f"{path}.corpus.archive_sha256")
    digest(corpus["target_payload_sha256"], f"{path}.corpus.target_payload_sha256")
    if corpus["xlsx"] is not None:
        fail(f"{path}.corpus.xlsx", "must remain null for the PPTX selector")
    if row.get("output_sha256") != corpus["archive_sha256"]:
        fail(f"{path}.output_sha256", "does not equal the archive identity")
    return corpus


def _strict_sink(row: dict[str, Any], path: str) -> dict[str, Any]:
    sink = obj(row.get("sink"), f"{path}.sink")
    exact_keys(sink, ("accepted_bytes", "write_calls", "largest_write", "write_size_buckets", "retained_output_bytes", "input_bytes", "authored_part_bytes"), (), f"{path}.sink")
    for key in ("accepted_bytes", "write_calls", "largest_write", "retained_output_bytes", "input_bytes", "authored_part_bytes"):
        u64(sink[key], f"{path}.sink.{key}")
    buckets = obj(sink["write_size_buckets"], f"{path}.sink.write_size_buckets")
    exact_keys(buckets, SINK_BUCKET_FIELDS, (), f"{path}.sink.write_size_buckets")
    for key in SINK_BUCKET_FIELDS:
        u64(buckets[key], f"{path}.sink.write_size_buckets.{key}")
    if sink["retained_output_bytes"] != 0:
        fail(f"{path}.sink.retained_output_bytes", "must be zero for the hashing discard sink")
    return dict(sink, buckets=buckets)


def _strict_source(row: dict[str, Any], shape: str, path: str) -> dict[str, Any]:
    source = obj(row.get("source"), f"{path}.source")
    exact_keys(source, ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads", "pptx_slides"), (), f"{path}.source")
    for key in ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads"):
        values = array(source[key], f"{path}.source.{key}")
        if values:
            fail(f"{path}.source.{key}", "must be empty for the in-process streaming sink")
    slides = obj(source["pptx_slides"], f"{path}.source.pptx_slides")
    exact_keys(slides, ("role", "implementation", "timing_scope", "performance_claim", "semantic_sha256", "full_text_sha256", "archive_sha256", "target_payload_sha256", "archive_member_set_verified", "semantic_reopen_verified", "deterministic_output_verified", "slide_count", "text_box_count", "input_text_bytes", "authored_part_bytes", "observed_max_slide_xml_bytes", "max_slide_xml_bytes", "structural_metadata_fixed_member_count", "structural_metadata_members_per_slide", "structural_metadata_scope", "text_contract"), (), f"{path}.source.pptx_slides")
    for key in ("role", "implementation", "timing_scope", "performance_claim", "structural_metadata_scope", "text_contract"):
        text(slides[key], f"{path}.source.pptx_slides.{key}")
    for key in ("semantic_sha256", "full_text_sha256", "archive_sha256", "target_payload_sha256"):
        digest(slides[key], f"{path}.source.pptx_slides.{key}")
    for key in ("archive_member_set_verified", "semantic_reopen_verified", "deterministic_output_verified"):
        boolean(slides[key], f"{path}.source.pptx_slides.{key}")
    for key in ("slide_count", "text_box_count", "input_text_bytes", "authored_part_bytes", "observed_max_slide_xml_bytes", "max_slide_xml_bytes", "structural_metadata_fixed_member_count", "structural_metadata_members_per_slide"):
        u64(slides[key], f"{path}.source.pptx_slides.{key}")
    if slides["slide_count"] != SHAPES[shape] or slides["text_box_count"] != SHAPES[shape]:
        fail(f"{path}.source.pptx_slides", "slide/text-box counts differ from shape")
    if slides["archive_sha256"] != row["corpus"]["archive_sha256"] or slides["archive_sha256"] != row["output_sha256"]:
        fail(f"{path}.source.pptx_slides.archive_sha256", "does not bind output archive")
    return source


def validate_report(report: Mapping[str, Any], lane: Mapping[str, Any], *, path: str = "report", samples: int = SAMPLES, warmups: int = WARMUPS, workers: int = WORKERS, case: str = CASE) -> dict[str, Any]:
    if not isinstance(report, dict):
        fail(path, "report must be an object")
    if report.get("schema_version") != SCHEMA_VERSION:
        fail(f"{path}.schema_version", "must be schema version 1")
    mode = lane.get("mode")
    shape = lane.get("shape")
    if mode not in {"normal", "allocator"} or shape not in SHAPES:
        fail("lane", "mode/shape is not a formal PPTX lane")
    if samples != SAMPLES or warmups != WARMUPS or workers != WORKERS:
        fail("lane", "formal lane sample configuration differs from strict producer schema")
    check_hash_tree(report, path)
    check_tool(report, mode)
    check_binary_identity(report)
    check_environment(report, mode)
    check_configuration(report, shape)
    results = array(report.get("results"), f"{path}.results")
    if len(results) != 1:
        fail(f"{path}.results", "must contain exactly one row")
    row = obj(results[0], f"{path}.results[0]")
    exact_keys(row, ("case", "corpus", "elapsed_ns", "sink", "source", "output_sha256", "operation_metrics"), (), f"{path}.results[0]")
    if row["case"] != case:
        fail(f"{path}.results[0].case", f"must be {case!r}")
    output = digest(row["output_sha256"], f"{path}.results[0].output_sha256")
    corpus = _strict_corpus(row, shape, f"{path}.results[0]")
    sink = _strict_sink(row, f"{path}.results[0]")
    source = _strict_source(row, shape, f"{path}.results[0]")
    elapsed, order, elapsed_summary = check_elapsed(row["elapsed_ns"], f"{path}.results[0].elapsed_ns", samples)
    operation = check_operation_metrics(row["operation_metrics"], mode, elapsed, order, sink, f"{path}.results[0].operation_metrics")
    check_parallel(report, output)
    if "corpus_catalog" not in report:
        fail(f"{path}.corpus_catalog", "top-level catalog identity is missing")
    check_catalog_sidecar(Path(path), report, obj(report["corpus_catalog"], f"{path}.corpus_catalog"))
    identity = {
        "output_sha256": output,
        "corpus.archive_sha256": corpus["archive_sha256"],
        "sink.accepted_bytes": sink["accepted_bytes"],
        "sink.write_calls": sink["write_calls"],
        "sink.largest_write": sink["largest_write"],
        "sink.input_bytes": sink["input_bytes"],
        "sink.authored_part_bytes": sink["authored_part_bytes"],
        "source": sha256_json(source),
    }
    allocation: dict[str, list[int]] | None = None
    if mode == "allocator":
        raw_allocation = operation.get("allocation")
        if not isinstance(raw_allocation, dict):
            fail(f"{path}.results[0].operation_metrics.allocation", "allocator vectors are missing")
        allocation = {}
        for field in ALLOCATOR_FIELDS:
            values = check_metric_vector(raw_allocation[field], f"{path}.results[0].operation_metrics.allocation.{field}", SAMPLES, expected_status="measured", expected_scope=ALLOCATOR_SCOPE)
            if values is None:
                fail(f"{path}.results[0].operation_metrics.allocation.{field}", "measured values are missing")
            allocation[field] = values
    return {
        "row": row,
        "elapsed": elapsed,
        "sample_order": order,
        "elapsed_summary": elapsed_summary,
        "operation": {"allocation": allocation} if allocation is not None else {},
        "identity": identity,
        "mode": mode,
        "shape": shape,
        "lane": lane.get("lane"),
    }


def allocator_requested_bytes(report: Mapping[str, Any], lane: Mapping[str, Any], *, path: str = "report") -> list[int]:
    return validate_report(report, lane, path=path)["operation"]["allocation"]["allocated_bytes"]


def _guard_sink(result: dict[str, Any], case: str, path: str) -> dict[str, Any]:
    sink = obj(result.get("sink"), f"{path}.sink")
    common = {"accepted_bytes", "write_calls", "largest_write", "write_size_buckets", "retained_output_bytes", "retained_authoring_window_bytes", "input_bytes", "authored_part_bytes"}
    extras = {"rows", "cells"} if case in {"xlsx_streaming_create", "ods_streaming_create"} else {"paragraphs", "runs"}
    exact_keys(sink, common | extras, (), f"{path}.sink")
    for key, value in sink.items():
        if key == "write_size_buckets":
            continue
        u64(value, f"{path}.sink.{key}")
    if sink["retained_output_bytes"] != 0:
        fail(f"{path}.sink.retained_output_bytes", "must be zero for the discard sink")
    buckets = obj(sink["write_size_buckets"], f"{path}.sink.write_size_buckets")
    exact_keys(buckets, SINK_BUCKET_FIELDS, (), f"{path}.sink.write_size_buckets")
    for key in SINK_BUCKET_FIELDS:
        u64(buckets[key], f"{path}.sink.write_size_buckets.{key}")
    if sum(buckets.values()) != sink["write_calls"]:
        fail(f"{path}.sink.write_size_buckets", "bucket counts do not equal write_calls")
    if sink["write_calls"] == 0 and sink["largest_write"] != 0:
        fail(f"{path}.sink.largest_write", "cannot be non-zero when write_calls is zero")
    if sink["write_calls"] and sink["largest_write"] == 0:
        fail(f"{path}.sink.largest_write", "must be positive when writes are present")
    return dict(sink, buckets=buckets)


def _guard_corpus(result: dict[str, Any], shape: str, path: str) -> dict[str, Any]:
    corpus = obj(result.get("corpus"), f"{path}.corpus")
    exact_keys(corpus, ("name", "generator", "package_format", "shape", "payload_kind", "compression", "entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "archive_sha256", "target_entry", "target_payload_bytes", "target_payload_sha256", "xlsx"), (), f"{path}.corpus")
    case = text(result.get("case"), f"{path}.case")
    expected_compression = "mimetype=stored;xml=deflate" if case in {"odt_streaming_create", "odp_streaming_create"} else "deflate"
    if corpus["shape"] != shape or corpus["compression"] != expected_compression:
        fail(f"{path}.corpus", "shape or compression identity differs")
    for key in ("name", "generator", "package_format", "shape", "payload_kind", "compression", "target_entry"):
        text(corpus[key], f"{path}.corpus.{key}")
    for key in ("entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        u64(corpus[key], f"{path}.corpus.{key}")
    digest(corpus["archive_sha256"], f"{path}.corpus.archive_sha256")
    digest(corpus["target_payload_sha256"], f"{path}.corpus.target_payload_sha256")
    if corpus["xlsx"] is not None:
        xlsx = obj(corpus["xlsx"], f"{path}.corpus.xlsx")
        exact_keys(xlsx, ("sheet_count", "rows_per_sheet", "columns_per_sheet", "one_percent_update_count", "source_members"), (), f"{path}.corpus.xlsx")
        for key in ("sheet_count", "rows_per_sheet", "columns_per_sheet", "one_percent_update_count"):
            u64(xlsx[key], f"{path}.corpus.xlsx.{key}")
        members = obj(xlsx["source_members"], f"{path}.corpus.xlsx.source_members")
        exact_keys(members, ("workbook", "worksheets", "shared_strings", "styles"), (), f"{path}.corpus.xlsx.source_members")
        text(members["workbook"], f"{path}.corpus.xlsx.source_members.workbook")
        text(members["styles"], f"{path}.corpus.xlsx.source_members.styles")
        if not isinstance(members["worksheets"], list) or not members["worksheets"] or not all(isinstance(item, str) and item for item in members["worksheets"]):
            fail(f"{path}.corpus.xlsx.source_members.worksheets", "must contain names")
        if members["shared_strings"] is not None:
            text(members["shared_strings"], f"{path}.corpus.xlsx.source_members.shared_strings")
    return corpus


def validate_guard_result(result: Mapping[str, Any], *, path: str = "guard.result") -> dict[str, Any]:
    """Validate one shared-streaming guard row and all recursive metrics."""

    if not isinstance(result, dict):
        fail(path, "result must be an object")
    case = text(result.get("case"), f"{path}.case")
    if case not in {"docx_streaming_create", "xlsx_streaming_create", "odt_streaming_create", "ods_streaming_create", "odp_streaming_create"}:
        fail(f"{path}.case", "unknown guard selector")
    allowed = {"case", "corpus", "elapsed_ns", "sink", "output_sha256", "operation_metrics", "source"}
    if not set(result).issubset(allowed):
        fail(path, "producer result contains unknown fields")
    if case == "xlsx_streaming_create":
        if "source" in result:
            fail(f"{path}.source", "must be omitted for XLSX guard rows")
    else:
        if not isinstance(result.get("source"), dict):
            fail(f"{path}.source", "is missing for this guard selector")
    shape = obj(result.get("corpus"), f"{path}.corpus").get("shape")
    if shape not in {"tiny", "large"}:
        fail(f"{path}.corpus.shape", "must be tiny or large")
    corpus = _guard_corpus(result, shape, path)
    output = digest(result.get("output_sha256"), f"{path}.output_sha256")
    if output != corpus["archive_sha256"]:
        fail(f"{path}.output_sha256", "does not equal corpus archive")
    check_hash_tree(result, path)
    sink = _guard_sink(result, case, path)
    elapsed, order, elapsed_summary = check_elapsed(result.get("elapsed_ns"), f"{path}.elapsed_ns", SAMPLES)
    operation = obj(result.get("operation_metrics"), f"{path}.operation_metrics")
    check_operation_metrics(operation, "normal", elapsed, order, sink, f"{path}.operation_metrics")
    return {"case": case, "shape": shape, "output_sha256": output, "elapsed": elapsed, "sample_order": order, "elapsed_summary": elapsed_summary, "corpus": corpus, "sink": sink}


def validate_guard_report(report: Mapping[str, Any], selectors: Iterable[str], *, path: str = "guard.report") -> list[dict[str, Any]]:
    """Validate a ten-row shared-streaming report and its parallel envelope."""

    if not isinstance(report, dict):
        fail(path, "report must be an object")
    exact_keys(report, ("schema_version", "tool", "binary_identity", "environment", "configuration", "parallel_metrics", "results", "corpus_catalog"), (), path)
    if report["schema_version"] != SCHEMA_VERSION:
        fail(f"{path}.schema_version", "must be schema version 1")
    check_hash_tree(report, path)
    check_tool(report, "normal")
    check_binary_identity(report)
    check_environment(report, "normal")
    configuration = obj(report.get("configuration"), f"{path}.configuration")
    if configuration.get("samples_per_case") != SAMPLES or configuration.get("warmup_iterations_per_case") != WARMUPS or configuration.get("execution_workers") != [WORKERS]:
        fail(f"{path}.configuration", "guard sample configuration differs")
    expected_cases = list(selectors)
    if not expected_cases or len(set(expected_cases)) != len(expected_cases):
        fail(path, "guard selectors are malformed")
    results = array(report.get("results"), f"{path}.results")
    if len(results) != len(expected_cases) * 2:
        fail(f"{path}.results", "must contain selector x shape rows")
    checked = [validate_guard_result(value, path=f"{path}.results[{index}]") for index, value in enumerate(results)]
    if {(item["case"], item["shape"]) for item in checked} != {(case, shape) for case in expected_cases for shape in ("tiny", "large")}:
        fail(path, "guard selector/shape set differs")
    parallel = obj(report.get("parallel_metrics"), f"{path}.parallel_metrics")
    exact_keys(parallel, ("schema_version", "scope", "claim", "configured_worker_budget", "observed_process_thread_count", "cases"), (), f"{path}.parallel_metrics")
    if parallel["schema_version"] != 1 or parallel["scope"] != "explicit_local_execution_only" or parallel["claim"] != "descriptive":
        fail(f"{path}.parallel_metrics", "schema or claim identity differs")
    parallel_metric(parallel["configured_worker_budget"], f"{path}.parallel_metrics.configured_worker_budget", status="measured", value_kind="u64_list")
    if parallel["configured_worker_budget"]["value"] != [WORKERS] or parallel["configured_worker_budget"]["scope"] != "configuration.execution_workers":
        fail(f"{path}.parallel_metrics.configured_worker_budget", "does not bind workers=[1]")
    parallel_metric(parallel["observed_process_thread_count"], f"{path}.parallel_metrics.observed_process_thread_count", status="unavailable")
    case_rows = array(parallel["cases"], f"{path}.parallel_metrics.cases")
    if len(case_rows) != len(checked):
        fail(f"{path}.parallel_metrics.cases", "count differs")
    expected_pairs = {(item["case"], item["shape"]): item["output_sha256"] for item in checked}
    actual_pairs: set[tuple[str, str]] = set()
    for index, value in enumerate(case_rows):
        case_row = obj(value, f"{path}.parallel_metrics.cases[{index}]")
        exact_keys(case_row, ("case", "corpus_sha256", "configured_worker_count", "observed_local_worker_count", "deterministic_task_count", "deterministic_chunk_count", "lock_wait_ns"), (), f"{path}.parallel_metrics.cases[{index}]")
        pair = (text(case_row["case"], f"{path}.parallel_metrics.cases[{index}].case"), digest(case_row["corpus_sha256"], f"{path}.parallel_metrics.cases[{index}].corpus_sha256"))
        matching = [key for key, archive in expected_pairs.items() if key[0] == pair[0] and archive == pair[1]]
        if len(matching) != 1:
            fail(f"{path}.parallel_metrics.cases[{index}]", "corpus identity differs")
        actual_pairs.add(matching[0])
        for key in ("configured_worker_count", "observed_local_worker_count", "deterministic_task_count", "deterministic_chunk_count"):
            parallel_metric(case_row[key], f"{path}.parallel_metrics.cases[{index}].{key}", status="not_applicable")
        parallel_metric(case_row["lock_wait_ns"], f"{path}.parallel_metrics.cases[{index}].lock_wait_ns", status="unavailable")
    if actual_pairs != set(expected_pairs):
        fail(f"{path}.parallel_metrics.cases", "identities differ")
    return checked


def validate_pilot_report(report: Mapping[str, Any], lane: Mapping[str, Any], *, path: str = "report") -> dict[str, Any]:
    """Validate a one-sample pilot without weakening formal-lane checks."""

    if not isinstance(report, dict):
        fail(path, "report must be an object")
    if report.get("schema_version") != SCHEMA_VERSION:
        fail(f"{path}.schema_version", "must be schema version 1")
    mode = lane.get("mode")
    shape = lane.get("shape")
    if mode not in {"normal", "allocator"} or shape not in SHAPES:
        fail("lane", "pilot mode/shape is invalid")
    check_hash_tree(report, path)
    check_tool(report, mode)
    check_binary_identity(report)
    check_environment(report, mode)
    results = array(report.get("results"), f"{path}.results")
    if len(results) != 1:
        fail(f"{path}.results", "pilot must contain one result row")
    row = obj(results[0], f"{path}.results[0]")
    exact_keys(row, ("case", "corpus", "elapsed_ns", "sink", "source", "output_sha256", "operation_metrics"), (), f"{path}.results[0]")
    if row["case"] != CASE:
        fail(f"{path}.results[0].case", "pilot selector differs")
    output = digest(row["output_sha256"], f"{path}.results[0].output_sha256")
    corpus = _strict_corpus(row, shape, f"{path}.results[0]")
    sink = _strict_sink(row, f"{path}.results[0]")
    source = _strict_source(row, shape, f"{path}.results[0]")
    configuration = obj(report.get("configuration"), f"{path}.configuration")
    sample_count = u64(configuration.get("samples_per_case"), f"{path}.configuration.samples_per_case", positive=True)
    warmups = u64(configuration.get("warmup_iterations_per_case"), f"{path}.configuration.warmup_iterations_per_case")
    if sample_count != 1 or warmups != 0:
        fail(f"{path}.configuration", "pilot must use one sample and zero warmups")
    elapsed, order, elapsed_summary = check_elapsed(row["elapsed_ns"], f"{path}.results[0].elapsed_ns", 1)
    raw_operation = obj(row["operation_metrics"], f"{path}.results[0].operation_metrics")
    if raw_operation.get("sample_count") != 1 or raw_operation.get("sample_indices") != order:
        fail(f"{path}.results[0].operation_metrics", "pilot sample alignment differs")
    allocation: dict[str, list[int]] | None = None
    if mode == "allocator":
        raw_allocation = obj(raw_operation.get("allocation"), f"{path}.results[0].operation_metrics.allocation")
        allocation = {}
        for field in ALLOCATOR_FIELDS:
            values = check_metric_vector(raw_allocation.get(field), f"{path}.results[0].operation_metrics.allocation.{field}", 1, expected_status="measured", expected_scope=ALLOCATOR_SCOPE)
            if values is None:
                fail(f"{path}.results[0].operation_metrics.allocation.{field}", "pilot allocator value is missing")
            allocation[field] = values
        before = allocation["live_bytes_before"][0]
        after = allocation["live_bytes_after"][0]
        if allocation["failed_allocation_calls"][0] != 0 or after != before:
            fail(f"{path}.results[0].operation_metrics.allocation", "pilot failed allocation or live exit delta")
        if after != before + allocation["allocated_bytes"][0] - allocation["deallocated_bytes"][0]:
            fail(f"{path}.results[0].operation_metrics.allocation", "pilot allocator byte balance differs")
    check_parallel(report, output)
    if "corpus_catalog" not in report:
        fail(f"{path}.corpus_catalog", "pilot catalog identity is missing")
    check_catalog_sidecar(Path(path), report, obj(report["corpus_catalog"], f"{path}.corpus_catalog"))
    return {"row": row, "elapsed": elapsed, "sample_order": order, "elapsed_summary": elapsed_summary, "operation": {"allocation": allocation} if allocation is not None else {}, "identity": {"output_sha256": output, "corpus.archive_sha256": corpus["archive_sha256"], "sink.accepted_bytes": sink["accepted_bytes"], "sink.write_calls": sink["write_calls"], "sink.largest_write": sink["largest_write"], "sink.input_bytes": sink["input_bytes"], "sink.authored_part_bytes": sink["authored_part_bytes"], "source": sha256_json(source)}, "mode": mode, "shape": shape, "lane": lane.get("lane")}


__all__ = [
    "ALLOCATOR_FIELDS", "CASE", "ReportError", "SAMPLES", "SHAPES", "WARMUPS", "WORKERS",
    "allocator_requested_bytes", "canonical", "check_elapsed", "compare_vectors", "digest",
    "file_digest", "finite", "parse_counter_text", "percent_delta", "read_json",
    "sha256_json", "u64", "validate_guard_report", "validate_guard_result", "validate_pilot_report", "validate_report", "vector_stats",
]
