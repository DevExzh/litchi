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

CHANGE = 474

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

def check_elapsed(value: Any, path: str = "report.results[0].elapsed_ns") -> tuple[list[int], list[int], dict[str, Any]]:
    elapsed = obj(value, path)
    exact_keys(elapsed, ("unit", "samples", "sample_order", "min", "p50", "p95", "p99", "max", "mean", "standard_deviation", "confidence_interval_95"), (), path)
    if elapsed["unit"] != "ns":
        fail(f"{path}.unit", "must be ns")
    raw = array(elapsed["samples"], f"{path}.samples")
    if len(raw) != SAMPLES:
        fail(f"{path}.samples", f"must contain exactly {SAMPLES} raw samples")
    samples = [u64(item, f"{path}.samples[{index}]", positive=True) for index, item in enumerate(raw)]
    if samples != sorted(samples):
        fail(f"{path}.samples", "must be sorted by elapsed time")
    order_raw = array(elapsed["sample_order"], f"{path}.sample_order")
    if len(order_raw) != SAMPLES:
        fail(f"{path}.sample_order", "must align with all retained samples")
    order = [u64(item, f"{path}.sample_order[{index}]") for index, item in enumerate(order_raw)]
    if sorted(order) != list(range(SAMPLES)):
        fail(f"{path}.sample_order", "must be a complete original-sample permutation")
    for index in range(1, SAMPLES):
        if samples[index] == samples[index - 1] and order[index] <= order[index - 1]:
            fail(f"{path}.sample_order", "must increase across tied elapsed samples")

    mean = 0.0
    squared = 0.0
    for index, sample in enumerate(samples):
        current = float(sample)
        count = float(index + 1)
        delta = current - mean
        next_mean = mean + delta / count
        squared += delta * (current - next_mean)
        mean = next_mean
    standard_deviation = math.sqrt(squared / float(SAMPLES - 1))
    t = (
        2.045
        if SAMPLES - 1 == 29
        else 1.959963984540054
    )
    margin = t * standard_deviation / math.sqrt(float(SAMPLES))
    midpoint = samples[(SAMPLES - 1) // 2] // 2 + samples[SAMPLES // 2] // 2 + (samples[(SAMPLES - 1) // 2] % 2 + samples[SAMPLES // 2] % 2) // 2
    nearest = lambda percentile: samples[min(((percentile * SAMPLES + 99) // 100) - 1, SAMPLES - 1)]
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
