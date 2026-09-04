#!/usr/bin/env python3
"""Validate and summarize Change 0403 normal-latency OPC ABBA evidence.

This is a deliberately narrow validator for the fixed
``opc_serial_eager_open`` corpus.  It accepts only normal (uninstrumented)
reports, checks all four raw ABBA legs, recomputes elapsed statistics, and
keeps every per-row accepted/rejected/adverse decision.  A statistic is
accepted only when the candidate is strictly lower in both paired directions
and both same-implementation drifts fit that statistic's ceiling.

The selector is an in-process constructor measurement.  Its global harness
cache selector may be either the producer default ``["warm",
"cold-requested"]`` or the explicit ``["warm"]`` form, but it is inert for
this non-filesystem operation: every result has no cache state or filesystem
evidence.  The selected claim is always ``cache_state: warm``; a global
``cold-requested`` configuration does not constitute cold execution/evidence.
RSS, process and logical/physical I/O, cold-cache, fresh-child, and allocator
evidence are explicitly unavailable or not applicable and never become
claims here.  The fixed corpus identities and semantic vectors are
independent constants; a contract cannot redefine them.  Only Python's
standard library is used.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping


SCHEMA_VERSION = 1
VALIDATOR_NAME = "litchi-opc-serial-eager-latency-abba"
CASE = "opc_serial_eager_open"
CACHE_STATE = "warm"
# ``filesystem_cache_states`` is a global harness selector envelope.  This
# selector is intentionally inert for the in-process constructor case: the
# only accepted producer configurations are the historical explicit-warm
# form and the producer's default warm-plus-cold-requested form.  Neither
# form supplies filesystem evidence for this validator.
GLOBAL_CACHE_STATES = ["warm", "cold-requested"]
ACCEPTED_GLOBAL_CACHE_STATES = (tuple([CACHE_STATE]), tuple(GLOBAL_CACHE_STATES))
CACHE_CLAIM_SCOPE = (
    "cache_state: warm; in-process constructor only; global cold-requested "
    "configuration does not constitute cold execution/evidence"
)
SAMPLE_COUNT = 500
WARMUP_COUNT = 20
WORKERS = [1]
ABBA_ORDER = ["A1_control", "B1_candidate", "B2_candidate", "A2_control"]
ROLES = ("A1", "B1", "B2", "A2")
SHAPES = ("tiny", "many-small", "few-large")
PROFILE = "release"
TOOL_NAME = "litchi-perf-baseline"
TOOL_VERSION = "0.1.0"
NORMAL_BINARY = "litchi-perf-baseline"
TARGET_OS = "linux"
TARGET_ARCH = "x86_64"
ALLOCATOR = "Rust system allocator"
CPU_AFFINITY = "2"
ALIGNMENT = "elapsed_ns.samples_by_elapsed_then_sample_index"
LATENCY_CLAIM = "comparable_timed_operation"
RUSTC_VERSION_DEFAULT = "rustc 1.98.1 (48a229cea 2026-09-01)"
CONTROL_REVISION = "cca80d89bac0aa4e2740a7879cf39cdcd8cbbb44"
CANDIDATE_REVISION = "fadf43722289fc78f565b8265a03d4763d2660b5"
SHA256 = re.compile(r"^[0-9a-f]{64}$")
REVISION = re.compile(r"^[0-9a-f]{40}$")
U64_MAX = (1 << 64) - 1
I64_MIN = -(1 << 63)
I64_MAX = (1 << 63) - 1
STATISTICS = ("p50", "mean", "p95", "p99")
DRIFT_CEILINGS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}

# The allocator model is retained by the producer's semantic summary even in
# the normal binary.  It is an expected, not-observed model and is explicitly
# excluded from this validator's claims.
COUNTER_METRICS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
)
EXPECTED_ALLOCATOR_MODEL = {
    "tiny": {
        "allocation_calls": -4,
        "deallocation_calls": -4,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": -160_640,
        "deallocated_bytes": -160_640,
    },
    "many-small": {
        "allocation_calls": -510,
        "deallocation_calls": -510,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": -20_481_600,
        "deallocated_bytes": -20_481_600,
    },
    "few-large": {
        "allocation_calls": -6,
        "deallocation_calls": -6,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": -240_960,
        "deallocated_bytes": -240_960,
    },
}

# This table intentionally does not come from report data or a caller
# contract.  The archive/name/count fields are the producer's fixed harness
# identities; target and aggregate payload digests are independently pinned.
CORPUS_ORACLE = {
    "tiny": {
        "name": "tiny-compressible",
        "generator": "litchi-opc-synthetic-v2",
        "package_format": "OPC/ZIP",
        "shape": "tiny",
        "payload_kind": "compressible",
        "compression": "deflate",
        "entry_count": 3,
        "part_count": 3,
        "archive_member_count": 5,
        "entry_bytes": 512,
        "uncompressed_payload_bytes": 1_536,
        "archive_bytes": 1_310,
        "archive_sha256": "1e28b8a9049a82f07e8ea88b2d492ef522d2da793d22fa50e2fe7f354dca3e2a",
        "target_entry": "benchmark/parts/00001.bin",
        "target_payload_bytes": 512,
        "target_payload_sha256": "630b1da45fe604eda3b5468b7c9ca7facfbd404941779786276a69ff870e4bdd",
        "part_names_sha256": "5458f5d1eb9283e10cd7057abf8f63cce9d1e0b6c57c5f9f945a9bad3b99cda4",
        "part_payload_sha256": "d1baa4a40fc63856136504f95933bcb2bb3da28f2000cabe1153eaee88c723c0",
    },
    "many-small": {
        "name": "many-small-incompressible",
        "generator": "litchi-opc-synthetic-v2",
        "package_format": "OPC/ZIP",
        "shape": "many-small",
        "payload_kind": "incompressible",
        "compression": "deflate",
        "entry_count": 256,
        "part_count": 256,
        "archive_member_count": 258,
        "entry_bytes": 1_024,
        "uncompressed_payload_bytes": 262_144,
        "archive_bytes": 303_003,
        "archive_sha256": "183178dec5b0fd578e5af04279032368598eec79da7caf0441fc979ce8fc14a0",
        "target_entry": "benchmark/parts/00128.bin",
        "target_payload_bytes": 1_024,
        "target_payload_sha256": "05fd26cad1f538b7ed415a0f525a13896823b02abcf22ad1746172f035a2149d",
        "part_names_sha256": "82415ca7ad25155c41df5d93707c95e5fcc31e66cde226ff046fc84906f56bc2",
        "part_payload_sha256": "7bdf372948a4f914aea31187d1f2813254957cd907279690b022ef00737caaa7",
    },
    "few-large": {
        "name": "few-large-incompressible",
        "generator": "litchi-opc-synthetic-v2",
        "package_format": "OPC/ZIP",
        "shape": "few-large",
        "payload_kind": "incompressible",
        "compression": "deflate",
        "entry_count": 4,
        "part_count": 4,
        "archive_member_count": 6,
        "entry_bytes": 4 * 1024 * 1024,
        "uncompressed_payload_bytes": 16 * 1024 * 1024,
        "archive_bytes": 16_783_632,
        "archive_sha256": "a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6",
        "target_entry": "benchmark/parts/00002.bin",
        "target_payload_bytes": 4 * 1024 * 1024,
        "target_payload_sha256": "3dbf6225021a99c1da8750a738bde21f57591c0be1a60aa510966c47ee25b098",
        "part_names_sha256": "d48e27d95e97a4de43e476096910540416f6e19eb54a3759d5ca081b4136166c",
        "part_payload_sha256": "ac1e942c87db2e622c1e1c2efd1046e5d791a44db73bd6255078f8816d922db3",
    },
}

CONTRACT_KEYS = {
    "schema_version",
    "case",
    "cache_state",
    "samples_per_case",
    "warmup_iterations_per_case",
    "execution_workers",
    "abba_order",
    "tool",
    "environment",
    "legs",
    "corpora",
    "expected_deltas",
}
CONTRACT_TOOL_KEYS = {
    "name",
    "version",
    "binary",
    "profile",
    "target_os",
    "target_arch",
    "instrumentation",
}
CONTRACT_ENVIRONMENT_KEYS = {
    "rustc_version",
    "allocator",
    "target_os",
    "target_arch",
    "logical_cpus_available",
    "cpu_model",
    "cpu_affinity",
    "rustflags",
    "cargo_build_target",
}
CONTRACT_LEG_KEYS = {
    "implementation",
    "revision",
    "binary_sha256",
    "binary_bytes",
    "mode_bits",
    "profile",
}
CONTRACT_CORPUS_KEYS = set(CORPUS_ORACLE["tiny"])

REPORT_KEYS = {
    "schema_version",
    "tool",
    "binary_identity",
    "environment",
    "configuration",
    "parallel_metrics",
    "results",
}
REPORT_BINARY_KEYS = {
    "path",
    "binary_sha256",
    "binary_bytes",
    "mode_bits",
    "executable",
    "profile",
}
REPORT_ENVIRONMENT_KEYS = {
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
}
CONFIGURATION_KEYS = {
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
}
RANGE_SIMULATION_KEYS = {
    "fixed_latency_us",
    "request_overhead_us",
    "bandwidth_bytes_per_second",
    "max_physical_range_bytes",
}
RESULT_KEYS = {
    "case",
    "corpus",
    "elapsed_ns",
    "sink",
    "source",
    "execution",
    "output_sha256",
    "operation_metrics",
}
OPTIONAL_NULL_RESULT_KEYS = {"cache_state"}
CORPUS_REPORT_KEYS = {
    key
    for key in CONTRACT_CORPUS_KEYS
    if key not in {"part_count", "part_names_sha256", "part_payload_sha256", "rtf_variant"}
} | {"xlsx"}
ELAPSED_KEYS = {
    "unit",
    "samples",
    "sample_order",
    "min",
    "p50",
    "p95",
    "p99",
    "max",
    "mean",
    "standard_deviation",
    "confidence_interval_95",
}
CONFIDENCE_KEYS = {"method", "lower", "upper"}
SOURCE_KEYS = {
    "read_calls",
    "read_bytes",
    "ordinary_payload_read_calls",
    "ordinary_payload_read_bytes",
    "max_in_flight_reads",
    "opc_serial_eager_open",
}
SERIAL_SUMMARY_KEYS = {
    "implementation",
    "timing_scope",
    "performance_claim",
    "predeclared_allocator_model",
    "worker_count",
    "source_archive_bytes",
    "source_archive_sha256",
    "archive_member_count",
    "part_count",
    "part_names_sha256",
    "part_payload_sha256",
    "target_name",
    "target_payload_sha256",
    "all_ordinary_parts_deflated_verified",
    "observed_part_counts",
    "observed_part_names_sha256",
    "observed_part_payload_sha256",
    "observed_content_types_verified",
    "observed_root_relationship_verified",
    "observed_main_target_verified",
    "observed_deterministic_payload_hashes_verified",
}
MODEL_KEYS = {"comparison", "status", *COUNTER_METRICS}
OPERATION_KEYS = {
    "sample_count",
    "sample_indices",
    "alignment",
    "latency_claim",
    "source",
    "process",
    "sink",
    "publication",
    "materialization",
    "cfb_phases",
    "allocation",
}
METRIC_VECTOR_KEYS = {"status", "scope"}
ABSENT_VECTOR_KEYS = {"status", "scope"}
OPERATION_SOURCE_KEYS = {
    "status",
    "counter_scope",
    "logical_read_calls",
    "logical_read_requested_bytes",
    "logical_read_returned_bytes",
    "logical_read_largest_requested_bytes",
    "logical_read_largest_returned_bytes",
    "logical_read_pattern",
    "compressed_bytes",
    "decompressed_bytes",
    "recompressed_bytes",
    "max_concurrent_reads",
}
OPERATION_PROCESS_KEYS = {
    "status",
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
}
OPERATION_SINK_KEYS = {
    "status",
    "output_bytes",
    "write_status",
    "accepted_bytes",
    "write_calls",
    "largest_write",
    "write_size_buckets",
}
SINK_BUCKET_KEYS = {
    "status",
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
}
OPERATION_PUBLICATION_KEYS = {"status", "changed_spans", "published_bytes"}
OPERATION_MATERIALIZATION_KEYS = {"status", "opc_parts"}
OPERATION_CFB_KEYS = {"status", "open", "plan", "atomic_publication"}
OPERATION_CFB_PHASE_KEYS = {
    "elapsed_ns",
    "logical_read_calls",
    "logical_read_requested_bytes",
    "logical_read_returned_bytes",
}
PARALLEL_KEYS = {
    "schema_version",
    "scope",
    "claim",
    "configured_worker_budget",
    "observed_process_thread_count",
    "cases",
}
PARALLEL_METRIC_KEYS = {"status", "scope", "value", "reason"}

SOURCE_SCOPE = "operation_logical_read_at"
SOURCE_PATTERN_SCOPE = "operation_logical_read_at_range_order_not_physical_io"
SOURCE_COMPRESSED_SCOPE = "unavailable_read_at_has_no_compressed_member_boundary"
SOURCE_DECOMPRESSED_SCOPE = "unavailable_read_at_has_no_decompressed_byte_boundary"
SOURCE_RECOMPRESSED_SCOPE = "unavailable_atomic_save_has_no_recompressed_byte_boundary"
PROCESS_SCOPE = "procfs_in_process_operation_delta_including_procfs_probe_overhead"
RSS_SCOPE = "procfs_in_process_rss_delta_including_procfs_probe_overhead"
HWM_SCOPE = "process_lifetime_high_water_after_not_operation_peak"
OUTPUT_SCOPE = "post_operation_output_length_not_sink_write_volume"
SINK_SCOPE = "logical_sink_accepted_write_bytes"
SINK_WRITE_CALLS_SCOPE = "logical_sink_accepted_write_calls"
SINK_LARGEST_WRITE_SCOPE = "logical_sink_largest_accepted_write"
SINK_BUCKET_SCOPE = "logical_sink_accepted_write_size_bucket_counts"
PUBLICATION_SCOPE = "logical_publication_counter"
MATERIALIZATION_SCOPE = "logical_materialization_counter"
CFB_ELAPSED_SCOPE = "timed_cfb_phase_elapsed_ns"
CFB_SOURCE_SCOPE = "timed_cfb_phase_logical_read_at"
ALLOCATOR_SCOPE = "operation_global_system_allocator"


class ValidationError(ValueError):
    """Raised when a contract or report cannot be accepted."""


@dataclass(frozen=True)
class Contract:
    raw_sha256: str
    tool: dict[str, Any]
    environment: dict[str, Any]
    legs: dict[str, dict[str, Any]]
    corpora: dict[str, dict[str, Any]]


@dataclass(frozen=True)
class RowEvidence:
    shape: str
    corpus: dict[str, Any]
    summary: dict[str, Any]
    output_sha256: str
    elapsed: dict[str, Any]
    sample_order: list[int]


@dataclass(frozen=True)
class ValidatedReport:
    role: str
    raw_sha256: str
    report: dict[str, Any]
    rows: dict[str, RowEvidence]


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValidationError(f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _reject_constant(token: str) -> Any:
    raise ValidationError(f"non-finite JSON number {token!r}")


def _parse_float(token: str) -> float:
    value = float(token)
    if not math.isfinite(value):
        raise ValidationError(f"non-finite JSON number {token!r}")
    return value


def _read_json(path: Path, context: str) -> tuple[dict[str, Any], str]:
    try:
        raw = path.read_bytes()
        value = json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=_reject_duplicate_keys,
            parse_constant=_reject_constant,
            parse_float=_parse_float,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValidationError) as error:
        raise ValidationError(f"cannot read {context} {path}: {error}") from error
    if not isinstance(value, dict):
        raise ValidationError(f"{context} {path} must contain a JSON object")
    return value, hashlib.sha256(raw).hexdigest()


def _object(value: Any, context: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValidationError(f"{context} must be an object")
    return value


def _array(value: Any, context: str) -> list[Any]:
    if not isinstance(value, list):
        raise ValidationError(f"{context} must be an array")
    return value


def _string(value: Any, context: str) -> str:
    if not isinstance(value, str):
        raise ValidationError(f"{context} must be a string")
    return value


def _integer(value: Any, context: str, *, minimum: int = 0, maximum: int = U64_MAX) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValidationError(f"{context} must be an integer")
    if value < minimum or value > maximum:
        raise ValidationError(f"{context} must be in [{minimum}, {maximum}]")
    return value


def _signed_integer(value: Any, context: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValidationError(f"{context} must be an integer")
    if value < I64_MIN or value > I64_MAX:
        raise ValidationError(f"{context} must fit i64")
    return value


def _finite_number(value: Any, context: str) -> int | float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValidationError(f"{context} must be a finite number")
    try:
        finite = math.isfinite(float(value))
    except (OverflowError, ValueError):
        finite = False
    if not finite:
        raise ValidationError(f"{context} must be a finite number")
    return value


def _exact_keys(value: Any, keys: set[str], context: str) -> dict[str, Any]:
    obj = _object(value, context)
    actual = set(obj)
    if actual != keys:
        raise ValidationError(
            f"{context} has unexpected keys (missing={sorted(keys - actual)}, "
            f"extra={sorted(actual - keys)})"
        )
    return obj


def _expect(actual: Any, expected: Any, context: str) -> None:
    if type(actual) is not type(expected):
        raise ValidationError(f"{context} must equal {expected!r}, got {actual!r}")
    if isinstance(expected, dict):
        if set(actual) != set(expected):
            raise ValidationError(f"{context} must equal {expected!r}, got {actual!r}")
        for key in expected:
            _expect(actual[key], expected[key], f"{context}.{key}")
        return
    if isinstance(expected, list):
        if len(actual) != len(expected):
            raise ValidationError(f"{context} must equal {expected!r}, got {actual!r}")
        for index, (left, right) in enumerate(zip(actual, expected, strict=True)):
            _expect(left, right, f"{context}[{index}]")
        return
    if actual != expected:
        raise ValidationError(f"{context} must equal {expected!r}, got {actual!r}")


def _digest(value: Any, context: str) -> str:
    value = _string(value, context)
    if SHA256.fullmatch(value) is None:
        raise ValidationError(f"{context} must be a lowercase SHA-256 digest")
    return value


def _revision(value: Any, context: str) -> str:
    value = _string(value, context)
    if REVISION.fullmatch(value) is None:
        raise ValidationError(f"{context} must be a lowercase 40-character revision")
    return value


def _validate_model(value: Any, shape: str, context: str) -> dict[str, Any]:
    model = _exact_keys(value, MODEL_KEYS, context)
    _expect(model["comparison"], "candidate-control", f"{context}.comparison")
    _expect(model["status"], "expected_not_observed", f"{context}.status")
    for metric in COUNTER_METRICS:
        _signed_integer(model[metric], f"{context}.{metric}")
        _expect(model[metric], EXPECTED_ALLOCATOR_MODEL[shape][metric], f"{context}.{metric}")
    return model


def _validate_contract_document(document: Any, raw_sha256: str) -> Contract:
    contract = _exact_keys(document, CONTRACT_KEYS, "contract")
    _expect(contract["schema_version"], SCHEMA_VERSION, "contract.schema_version")
    _expect(contract["case"], CASE, "contract.case")
    _expect(contract["cache_state"], CACHE_STATE, "contract.cache_state")
    _expect(contract["samples_per_case"], SAMPLE_COUNT, "contract.samples_per_case")
    _expect(contract["warmup_iterations_per_case"], WARMUP_COUNT, "contract.warmup_iterations_per_case")
    _expect(contract["execution_workers"], WORKERS, "contract.execution_workers")
    _expect(contract["abba_order"], ABBA_ORDER, "contract.abba_order")

    tool = _exact_keys(contract["tool"], CONTRACT_TOOL_KEYS, "contract.tool")
    _expect(
        tool,
        {
            "name": TOOL_NAME,
            "version": TOOL_VERSION,
            "binary": NORMAL_BINARY,
            "profile": PROFILE,
            "target_os": TARGET_OS,
            "target_arch": TARGET_ARCH,
            "instrumentation": "none",
        },
        "contract.tool",
    )
    environment = _exact_keys(contract["environment"], CONTRACT_ENVIRONMENT_KEYS, "contract.environment")
    _string(environment["rustc_version"], "contract.environment.rustc_version")
    _expect(environment["allocator"], ALLOCATOR, "contract.environment.allocator")
    _expect(environment["target_os"], TARGET_OS, "contract.environment.target_os")
    _expect(environment["target_arch"], TARGET_ARCH, "contract.environment.target_arch")
    _integer(environment["logical_cpus_available"], "contract.environment.logical_cpus_available", minimum=1)
    _string(environment["cpu_model"], "contract.environment.cpu_model")
    _expect(environment["cpu_affinity"], CPU_AFFINITY, "contract.environment.cpu_affinity")
    if environment["rustflags"] is not None:
        _string(environment["rustflags"], "contract.environment.rustflags")
    if environment["cargo_build_target"] is not None:
        _string(environment["cargo_build_target"], "contract.environment.cargo_build_target")

    raw_legs = _exact_keys(contract["legs"], set(ROLES), "contract.legs")
    legs: dict[str, dict[str, Any]] = {}
    for role in ROLES:
        item = _exact_keys(raw_legs[role], CONTRACT_LEG_KEYS, f"contract.legs.{role}")
        expected_implementation = "control" if role in ("A1", "A2") else "candidate"
        expected_revision = CONTROL_REVISION if expected_implementation == "control" else CANDIDATE_REVISION
        _expect(item["implementation"], expected_implementation, f"contract.legs.{role}.implementation")
        _expect(item["revision"], expected_revision, f"contract.legs.{role}.revision")
        _revision(item["revision"], f"contract.legs.{role}.revision")
        _digest(item["binary_sha256"], f"contract.legs.{role}.binary_sha256")
        _integer(item["binary_bytes"], f"contract.legs.{role}.binary_bytes", minimum=1)
        _integer(item["mode_bits"], f"contract.legs.{role}.mode_bits", minimum=1, maximum=0o7777)
        _expect(item["profile"], PROFILE, f"contract.legs.{role}.profile")
        legs[role] = item
    _expect(legs["A1"], legs["A2"], "contract control A1/A2 identity")
    _expect(legs["B1"], legs["B2"], "contract candidate B1/B2 identity")
    if legs["A1"]["binary_sha256"] == legs["B1"]["binary_sha256"]:
        raise ValidationError("contract control and candidate binaries must differ")

    raw_corpora = _array(contract["corpora"], "contract.corpora")
    if len(raw_corpora) != len(SHAPES):
        raise ValidationError(f"contract.corpora must contain exactly {len(SHAPES)} rows")
    corpora: dict[str, dict[str, Any]] = {}
    for index, raw_corpus in enumerate(raw_corpora):
        item = _exact_keys(raw_corpus, CONTRACT_CORPUS_KEYS, f"contract.corpora[{index}]")
        shape = _string(item["shape"], f"contract.corpora[{index}].shape")
        if shape != SHAPES[index]:
            raise ValidationError(
                f"contract.corpora[{index}] must retain fixed shape order; expected {SHAPES[index]!r}"
            )
        if shape in corpora:
            raise ValidationError(f"contract.corpora has duplicate shape {shape!r}")
        _expect(item, CORPUS_ORACLE[shape], f"contract.corpora[{index}] retained fixed identity")
        for key in ("archive_sha256", "target_payload_sha256", "part_names_sha256", "part_payload_sha256"):
            _digest(item[key], f"contract.corpora[{index}].{key}")
        for key in (
            "entry_count",
            "part_count",
            "archive_member_count",
            "entry_bytes",
            "uncompressed_payload_bytes",
            "archive_bytes",
            "target_payload_bytes",
        ):
            _integer(item[key], f"contract.corpora[{index}].{key}", minimum=1)
        corpora[shape] = item
    if set(corpora) != set(SHAPES):
        raise ValidationError("contract.corpora must contain the fixed three shapes")

    raw_deltas = _exact_keys(contract["expected_deltas"], set(SHAPES), "contract.expected_deltas")
    for shape in SHAPES:
        item = _exact_keys(raw_deltas[shape], set(COUNTER_METRICS), f"contract.expected_deltas.{shape}")
        for metric in COUNTER_METRICS:
            _signed_integer(item[metric], f"contract.expected_deltas.{shape}.{metric}")
            _expect(item[metric], EXPECTED_ALLOCATOR_MODEL[shape][metric], f"contract.expected_deltas.{shape}.{metric}")
    return Contract(raw_sha256, tool, environment, legs, corpora)


def _validate_contract_object(value: Contract) -> Contract:
    """Revalidate a direct Contract exactly as a serialized contract document."""
    if type(value) is not Contract:
        raise ValidationError("contract must be a Contract object")
    if not isinstance(value.raw_sha256, str):
        raise ValidationError("contract.raw_sha256 must be a string")
    if value.raw_sha256 and SHA256.fullmatch(value.raw_sha256) is None:
        raise ValidationError("contract.raw_sha256 must be a lowercase SHA-256 digest or empty")
    if not isinstance(value.corpora, dict):
        raise ValidationError("contract.corpora must be an object")
    if set(value.corpora) != set(SHAPES):
        raise ValidationError("contract.corpora must contain exactly the fixed three shapes")
    document = {
        "schema_version": SCHEMA_VERSION,
        "case": CASE,
        "cache_state": CACHE_STATE,
        "samples_per_case": SAMPLE_COUNT,
        "warmup_iterations_per_case": WARMUP_COUNT,
        "execution_workers": list(WORKERS),
        "abba_order": list(ABBA_ORDER),
        "tool": value.tool,
        "environment": value.environment,
        "legs": value.legs,
        "corpora": [value.corpora[shape] for shape in SHAPES],
        "expected_deltas": EXPECTED_ALLOCATOR_MODEL,
    }
    return _validate_contract_document(document, value.raw_sha256)


def _validate_contract(value: Path | Contract) -> Contract:
    """Validate either a contract file or a direct Contract object."""
    if isinstance(value, Contract):
        return _validate_contract_object(value)
    if not isinstance(value, Path):
        raise ValidationError("contract must be a path or Contract object")
    document, raw_sha256 = _read_json(value, "contract")
    return _validate_contract_document(document, raw_sha256)


def _contract_from_cli(args: argparse.Namespace) -> Contract:
    values = {
        "control_revision": args.control_revision,
        "candidate_revision": args.candidate_revision,
        "control_binary_sha256": args.control_binary_sha256,
        "candidate_binary_sha256": args.candidate_binary_sha256,
        "control_binary_bytes": args.control_binary_bytes,
        "candidate_binary_bytes": args.candidate_binary_bytes,
        "control_mode_bits": args.control_mode_bits,
        "candidate_mode_bits": args.candidate_mode_bits,
    }
    missing = [name for name, value in values.items() if value is None]
    if missing:
        raise ValidationError(
            "--contract or all explicit identity flags are required; missing "
            + ", ".join(f"--{name.replace('_', '-')}" for name in missing)
        )
    _expect(values["control_revision"], CONTROL_REVISION, "control_revision")
    _expect(values["candidate_revision"], CANDIDATE_REVISION, "candidate_revision")
    for name in ("control_revision", "candidate_revision"):
        _revision(values[name], name)
    for name in ("control_binary_sha256", "candidate_binary_sha256"):
        _digest(values[name], name)
    for name in (
        "control_binary_bytes",
        "candidate_binary_bytes",
        "control_mode_bits",
        "candidate_mode_bits",
    ):
        _integer(values[name], name, minimum=1, maximum=0o7777 if "mode_bits" in name else U64_MAX)
    if values["control_binary_sha256"] == values["candidate_binary_sha256"]:
        raise ValidationError("control and candidate binaries must differ")
    if args.rustc_version is None or args.logical_cpus is None or args.cpu_model is None:
        raise ValidationError("--rustc-version, --logical-cpus, and --cpu-model are required without --contract")
    _string(args.rustc_version, "--rustc-version")
    _integer(args.logical_cpus, "--logical-cpus", minimum=1)
    _string(args.cpu_model, "--cpu-model")
    cpu_affinity = CPU_AFFINITY if args.cpu_affinity is None else args.cpu_affinity
    _expect(cpu_affinity, CPU_AFFINITY, "--cpu-affinity")
    environment = {
        "rustc_version": args.rustc_version,
        "allocator": ALLOCATOR,
        "target_os": TARGET_OS,
        "target_arch": TARGET_ARCH,
        "logical_cpus_available": args.logical_cpus,
        "cpu_model": args.cpu_model,
        "cpu_affinity": cpu_affinity,
        "rustflags": args.rustflags,
        "cargo_build_target": args.cargo_build_target,
    }
    control = {
        "implementation": "control",
        "revision": CONTROL_REVISION,
        "binary_sha256": values["control_binary_sha256"],
        "binary_bytes": values["control_binary_bytes"],
        "mode_bits": values["control_mode_bits"],
        "profile": PROFILE,
    }
    candidate = {
        "implementation": "candidate",
        "revision": CANDIDATE_REVISION,
        "binary_sha256": values["candidate_binary_sha256"],
        "binary_bytes": values["candidate_binary_bytes"],
        "mode_bits": values["candidate_mode_bits"],
        "profile": PROFILE,
    }
    return Contract(
        raw_sha256="",
        tool={
            "name": TOOL_NAME,
            "version": TOOL_VERSION,
            "binary": NORMAL_BINARY,
            "profile": PROFILE,
            "target_os": TARGET_OS,
            "target_arch": TARGET_ARCH,
            "instrumentation": "none",
        },
        environment=environment,
        legs={"A1": control, "A2": dict(control), "B1": candidate, "B2": dict(candidate)},
        corpora={shape: dict(CORPUS_ORACLE[shape]) for shape in SHAPES},
    )


def _student_t_critical_95(degrees: int) -> float:
    values = (
        12.706,
        4.303,
        3.182,
        2.776,
        2.571,
        2.447,
        2.365,
        2.306,
        2.262,
        2.228,
        2.201,
        2.179,
        2.160,
        2.145,
        2.131,
        2.120,
        2.110,
        2.101,
        2.093,
        2.086,
        2.080,
        2.074,
        2.069,
        2.064,
        2.060,
        2.056,
        2.052,
        2.048,
        2.045,
        2.042,
    )
    if degrees <= 0:
        return 0.0
    if degrees <= len(values):
        return values[degrees - 1]
    z = 1.959963984540054
    d = float(degrees)
    z2 = z * z
    z3 = z2 * z
    z5 = z3 * z2
    z7 = z5 * z2
    return (
        z
        + (z3 + z) / (4.0 * d)
        + (5.0 * z5 + 16.0 * z3 + 3.0 * z) / (96.0 * d * d)
        + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z)
        / (384.0 * d * d * d)
    )


def _midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def _nearest_rank(samples: list[int], percentile: int) -> int:
    index = min((percentile * len(samples) + 99) // 100 - 1, len(samples) - 1)
    return samples[index]


def _float_close(actual: Any, expected: float, context: str) -> None:
    value = _finite_number(actual, context)
    if not math.isclose(float(value), expected, rel_tol=1e-12, abs_tol=1e-9):
        raise ValidationError(f"{context} does not match recomputed statistic")


def _validate_elapsed(value: Any, context: str) -> tuple[dict[str, Any], list[int]]:
    elapsed = _exact_keys(value, ELAPSED_KEYS, context)
    _expect(elapsed["unit"], "ns", f"{context}.unit")
    samples = [
        _integer(item, f"{context}.samples[{index}]", minimum=1)
        for index, item in enumerate(_array(elapsed["samples"], f"{context}.samples"))
    ]
    if len(samples) != SAMPLE_COUNT:
        raise ValidationError(f"{context}.samples must contain exactly {SAMPLE_COUNT} values")
    if samples != sorted(samples):
        raise ValidationError(f"{context}.samples must be sorted by elapsed time")
    order = [
        _integer(item, f"{context}.sample_order[{index}]", maximum=SAMPLE_COUNT - 1)
        for index, item in enumerate(_array(elapsed["sample_order"], f"{context}.sample_order"))
    ]
    if len(order) != SAMPLE_COUNT or sorted(order) != list(range(SAMPLE_COUNT)):
        raise ValidationError(
            f"{context}.sample_order must be a complete retained-sample permutation"
        )
    if any(
        left == right and left_index >= right_index
        for left, right, left_index, right_index in zip(
            samples[:-1],
            samples[1:],
            order[:-1],
            order[1:],
            strict=True,
        )
    ):
        raise ValidationError(
            f"{context}.sample_order must be stable by (elapsed_ns, original_sample_index)"
        )

    _expect(elapsed["min"], samples[0], f"{context}.min")
    _expect(
        elapsed["p50"],
        _midpoint(samples[(len(samples) - 1) // 2], samples[len(samples) // 2]),
        f"{context}.p50",
    )
    _expect(elapsed["p95"], _nearest_rank(samples, 95), f"{context}.p95")
    _expect(elapsed["p99"], _nearest_rank(samples, 99), f"{context}.p99")
    _expect(elapsed["max"], samples[-1], f"{context}.max")

    mean = 0.0
    squared = 0.0
    for index, sample in enumerate(samples):
        current = float(sample)
        count = float(index + 1)
        delta = current - mean
        next_mean = mean + delta / count
        squared += delta * (current - next_mean)
        mean = next_mean
    standard_deviation = math.sqrt(squared / (len(samples) - 1))
    margin = _student_t_critical_95(len(samples) - 1) * standard_deviation / math.sqrt(len(samples))
    _float_close(elapsed["mean"], mean, f"{context}.mean")
    _float_close(elapsed["standard_deviation"], standard_deviation, f"{context}.standard_deviation")
    confidence = _exact_keys(
        elapsed["confidence_interval_95"],
        CONFIDENCE_KEYS,
        f"{context}.confidence_interval_95",
    )
    _expect(
        confidence["method"],
        "two-sided Student's t interval for the mean",
        f"{context}.confidence_interval_95.method",
    )
    _float_close(
        confidence["lower"],
        max(mean - margin, 0.0),
        f"{context}.confidence_interval_95.lower",
    )
    _float_close(
        confidence["upper"],
        mean + margin,
        f"{context}.confidence_interval_95.upper",
    )
    return elapsed, order


def _validate_absent_vector(value: Any, context: str, status: str, scope: str) -> None:
    vector = _exact_keys(value, ABSENT_VECTOR_KEYS, context)
    _expect(vector["status"], status, f"{context}.status")
    _expect(vector["scope"], scope, f"{context}.scope")


def _validate_unavailable_allocation(value: Any, context: str) -> None:
    allocation = _exact_keys(value, {"status", "scope", *COUNTER_METRICS, "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after"}, context)
    _expect(allocation["status"], "unavailable", f"{context}.status")
    _expect(allocation["scope"], ALLOCATOR_SCOPE, f"{context}.scope")
    for metric in (*COUNTER_METRICS, "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after"):
        _validate_absent_vector(allocation[metric], f"{context}.{metric}", "unavailable", ALLOCATOR_SCOPE)


def _validate_source(value: Any, corpus: Mapping[str, Any], context: str) -> dict[str, Any]:
    source = _exact_keys(value, SOURCE_KEYS, context)
    for key in (
        "read_calls",
        "read_bytes",
        "ordinary_payload_read_calls",
        "ordinary_payload_read_bytes",
        "max_in_flight_reads",
    ):
        _expect(source[key], [], f"{context}.{key}")
    summary = _exact_keys(
        source["opc_serial_eager_open"],
        SERIAL_SUMMARY_KEYS,
        f"{context}.opc_serial_eager_open",
    )
    shape = corpus["shape"]
    _expect(summary["implementation"], "OpcPackage::from_bytes", f"{context}.opc_serial_eager_open.implementation")
    _expect(
        summary["timing_scope"],
        "OpcPackage::from_bytes constructor only; ZIP preflight and all package semantic oracles excluded",
        f"{context}.opc_serial_eager_open.timing_scope",
    )
    _expect(summary["performance_claim"], "none", f"{context}.opc_serial_eager_open.performance_claim")
    _validate_model(summary["predeclared_allocator_model"], shape, f"{context}.opc_serial_eager_open.predeclared_allocator_model")
    _expect(summary["worker_count"], 1, f"{context}.opc_serial_eager_open.worker_count")
    _expect(summary["source_archive_bytes"], corpus["archive_bytes"], f"{context}.opc_serial_eager_open.source_archive_bytes")
    _expect(summary["source_archive_sha256"], corpus["archive_sha256"], f"{context}.opc_serial_eager_open.source_archive_sha256")
    _expect(summary["archive_member_count"], corpus["archive_member_count"], f"{context}.opc_serial_eager_open.archive_member_count")
    _expect(summary["part_count"], CORPUS_ORACLE[shape]["part_count"], f"{context}.opc_serial_eager_open.part_count")
    for key in ("part_names_sha256", "part_payload_sha256", "target_payload_sha256"):
        expected_key = "target_payload_sha256" if key == "target_payload_sha256" else key
        _expect(summary[key], CORPUS_ORACLE[shape][expected_key], f"{context}.opc_serial_eager_open.{key}")
        _digest(summary[key], f"{context}.opc_serial_eager_open.{key}")
    _expect(summary["target_name"], corpus["target_entry"], f"{context}.opc_serial_eager_open.target_name")
    _expect(summary["all_ordinary_parts_deflated_verified"], True, f"{context}.opc_serial_eager_open.all_ordinary_parts_deflated_verified")
    expected_counts = [CORPUS_ORACLE[shape]["part_count"]] * SAMPLE_COUNT
    expected_names = [CORPUS_ORACLE[shape]["part_names_sha256"]] * SAMPLE_COUNT
    expected_payloads = [CORPUS_ORACLE[shape]["part_payload_sha256"]] * SAMPLE_COUNT
    _expect(summary["observed_part_counts"], expected_counts, f"{context}.opc_serial_eager_open.observed_part_counts")
    _expect(summary["observed_part_names_sha256"], expected_names, f"{context}.opc_serial_eager_open.observed_part_names_sha256")
    _expect(summary["observed_part_payload_sha256"], expected_payloads, f"{context}.opc_serial_eager_open.observed_part_payload_sha256")
    for key in (
        "observed_content_types_verified",
        "observed_root_relationship_verified",
        "observed_main_target_verified",
        "observed_deterministic_payload_hashes_verified",
    ):
        _expect(summary[key], [True] * SAMPLE_COUNT, f"{context}.opc_serial_eager_open.{key}")
    _digest(summary["source_archive_sha256"], f"{context}.opc_serial_eager_open.source_archive_sha256")
    return summary


def _validate_process(value: Any, context: str) -> None:
    process = _exact_keys(value, OPERATION_PROCESS_KEYS, context)
    _expect(process["status"], "unavailable", f"{context}.status")
    for key in OPERATION_PROCESS_KEYS - {"status"}:
        scope = RSS_SCOPE if key == "rss_delta_bytes" else HWM_SCOPE if key == "peak_rss_bytes" else PROCESS_SCOPE
        _validate_absent_vector(process[key], f"{context}.{key}", "unavailable", scope)


def _validate_operation(value: Any, sample_order: list[int], context: str) -> None:
    operation = _exact_keys(value, OPERATION_KEYS, context)
    _expect(operation["sample_count"], SAMPLE_COUNT, f"{context}.sample_count")
    _expect(operation["sample_indices"], sample_order, f"{context}.sample_indices")
    _expect(operation["alignment"], ALIGNMENT, f"{context}.alignment")
    _expect(operation["latency_claim"], LATENCY_CLAIM, f"{context}.latency_claim")

    source = _exact_keys(operation["source"], OPERATION_SOURCE_KEYS, f"{context}.source")
    _expect(source["status"], "not_applicable", f"{context}.source.status")
    _expect(source["counter_scope"], "not_applicable_in_process_sink", f"{context}.source.counter_scope")
    for key in OPERATION_SOURCE_KEYS - {"status", "counter_scope"}:
        scope = (
            SOURCE_PATTERN_SCOPE
            if key == "logical_read_pattern"
            else SOURCE_COMPRESSED_SCOPE
            if key == "compressed_bytes"
            else SOURCE_DECOMPRESSED_SCOPE
            if key == "decompressed_bytes"
            else SOURCE_RECOMPRESSED_SCOPE
            if key == "recompressed_bytes"
            else SOURCE_SCOPE
        )
        _validate_absent_vector(source[key], f"{context}.source.{key}", "not_applicable", scope)

    _validate_process(operation["process"], f"{context}.process")

    sink = _exact_keys(operation["sink"], OPERATION_SINK_KEYS, f"{context}.sink")
    _expect(sink["status"], "not_applicable", f"{context}.sink.status")
    _expect(sink["write_status"], "not_applicable", f"{context}.sink.write_status")
    for key, scope in {
        "output_bytes": OUTPUT_SCOPE,
        "accepted_bytes": SINK_SCOPE,
        "write_calls": SINK_WRITE_CALLS_SCOPE,
        "largest_write": SINK_LARGEST_WRITE_SCOPE,
    }.items():
        _validate_absent_vector(sink[key], f"{context}.sink.{key}", "not_applicable", scope)
    buckets = _exact_keys(sink["write_size_buckets"], SINK_BUCKET_KEYS, f"{context}.sink.write_size_buckets")
    _expect(buckets["status"], "not_applicable", f"{context}.sink.write_size_buckets.status")
    for key in SINK_BUCKET_KEYS - {"status"}:
        _validate_absent_vector(buckets[key], f"{context}.sink.write_size_buckets.{key}", "not_applicable", SINK_BUCKET_SCOPE)

    publication = _exact_keys(operation["publication"], OPERATION_PUBLICATION_KEYS, f"{context}.publication")
    _expect(publication["status"], "not_applicable", f"{context}.publication.status")
    for key in OPERATION_PUBLICATION_KEYS - {"status"}:
        _validate_absent_vector(publication[key], f"{context}.publication.{key}", "not_applicable", PUBLICATION_SCOPE)
    materialization = _exact_keys(operation["materialization"], OPERATION_MATERIALIZATION_KEYS, f"{context}.materialization")
    _expect(materialization["status"], "not_applicable", f"{context}.materialization.status")
    _validate_absent_vector(materialization["opc_parts"], f"{context}.materialization.opc_parts", "not_applicable", MATERIALIZATION_SCOPE)
    cfb = _exact_keys(operation["cfb_phases"], OPERATION_CFB_KEYS, f"{context}.cfb_phases")
    _expect(cfb["status"], "not_applicable", f"{context}.cfb_phases.status")
    for phase in ("open", "plan", "atomic_publication"):
        phase_value = _exact_keys(cfb[phase], OPERATION_CFB_PHASE_KEYS, f"{context}.cfb_phases.{phase}")
        _validate_absent_vector(phase_value["elapsed_ns"], f"{context}.cfb_phases.{phase}.elapsed_ns", "not_applicable", CFB_ELAPSED_SCOPE)
        for key in OPERATION_CFB_PHASE_KEYS - {"elapsed_ns"}:
            _validate_absent_vector(phase_value[key], f"{context}.cfb_phases.{phase}.{key}", "not_applicable", CFB_SOURCE_SCOPE)
    _validate_unavailable_allocation(operation["allocation"], f"{context}.allocation")


def _validate_parallel_metric(value: Any, status: str, scope: str, context: str, *, expected_value: Any = None) -> None:
    metric = _object(value, context)
    expected_keys = {"status", "scope"}
    if expected_value is not None:
        expected_keys.add("value")
    if status != "measured":
        expected_keys.add("reason")
    _exact_keys(metric, expected_keys, context)
    _expect(metric["status"], status, f"{context}.status")
    _expect(metric["scope"], scope, f"{context}.scope")
    if expected_value is not None:
        _expect(metric["value"], expected_value, f"{context}.value")
    if status != "measured":
        reason = _string(metric["reason"], f"{context}.reason")
        if not reason:
            raise ValidationError(f"{context}.reason must not be empty")


def _validate_parallel(value: Any, context: str) -> None:
    parallel = _exact_keys(value, PARALLEL_KEYS, context)
    _expect(parallel["schema_version"], 1, f"{context}.schema_version")
    _expect(parallel["scope"], "explicit_local_execution_only", f"{context}.scope")
    _expect(parallel["claim"], "descriptive", f"{context}.claim")
    _validate_parallel_metric(parallel["configured_worker_budget"], "measured", "configuration.execution_workers", f"{context}.configured_worker_budget", expected_value=WORKERS)
    _validate_parallel_metric(parallel["observed_process_thread_count"], "unavailable", "process_thread_count", f"{context}.observed_process_thread_count")
    cases = _array(parallel["cases"], f"{context}.cases")
    if len(cases) != len(SHAPES):
        raise ValidationError(f"{context}.cases must contain exactly {len(SHAPES)} rows")
    case_keys = {
        "case",
        "corpus_sha256",
        "configured_worker_count",
        "observed_local_worker_count",
        "deterministic_task_count",
        "deterministic_chunk_count",
        "lock_wait_ns",
    }
    for index, (item, shape) in enumerate(zip(cases, SHAPES, strict=True)):
        entry = _exact_keys(item, case_keys, f"{context}.cases[{index}]")
        _expect(entry["case"], CASE, f"{context}.cases[{index}].case")
        _expect(entry["corpus_sha256"], CORPUS_ORACLE[shape]["archive_sha256"], f"{context}.cases[{index}].corpus_sha256")
        _digest(entry["corpus_sha256"], f"{context}.cases[{index}].corpus_sha256")
        _validate_parallel_metric(entry["configured_worker_count"], "measured", "result.execution.worker_count", f"{context}.cases[{index}].configured_worker_count", expected_value=1)
        _validate_parallel_metric(entry["observed_local_worker_count"], "not_applicable", "result.source.opc_cache.worker_count_with_one_created_local_worker_team", f"{context}.cases[{index}].observed_local_worker_count")
        _validate_parallel_metric(entry["deterministic_task_count"], "measured", "result.execution.logical_tasks", f"{context}.cases[{index}].deterministic_task_count", expected_value=1)
        _validate_parallel_metric(entry["deterministic_chunk_count"], "unavailable", "result.execution.deterministic_chunk_count", f"{context}.cases[{index}].deterministic_chunk_count")
        _validate_parallel_metric(entry["lock_wait_ns"], "unavailable", "lock_wait_ns", f"{context}.cases[{index}].lock_wait_ns")


def _validate_corpus(value: Any, shape: str, context: str) -> dict[str, Any]:
    corpus = _exact_keys(value, CORPUS_REPORT_KEYS, context)
    expected = dict(CORPUS_ORACLE[shape])
    for key in ("part_count", "part_names_sha256", "part_payload_sha256", "rtf_variant"):
        expected.pop(key, None)
    expected["xlsx"] = None
    _expect(corpus, expected, context)
    for key in ("archive_sha256", "target_payload_sha256"):
        _digest(corpus[key], f"{context}.{key}")
    return corpus


def _validate_result(value: Any, role: str, index: int) -> RowEvidence:
    context = f"{role}.results[{index}]"
    result = _object(value, context)
    if "filesystem_evidence" in result:
        raise ValidationError(
            f"{context}.filesystem_evidence is forbidden for the in-process selector"
        )
    actual_keys = set(result)
    missing = RESULT_KEYS - actual_keys
    extra = actual_keys - RESULT_KEYS - OPTIONAL_NULL_RESULT_KEYS
    if missing or extra:
        raise ValidationError(
            f"{context} has unexpected keys (missing={sorted(missing)}, extra={sorted(extra)})"
        )
    if "cache_state" in result:
        _expect(result["cache_state"], None, f"{context}.cache_state")
    _expect(result["case"], CASE, f"{context}.case")
    _expect(result["sink"], None, f"{context}.sink")
    corpus_obj = _object(result["corpus"], f"{context}.corpus")
    shape = _string(corpus_obj.get("shape"), f"{context}.corpus.shape")
    if shape not in SHAPES:
        raise ValidationError(f"{context}.corpus.shape is outside the fixed matrix")
    corpus = _validate_corpus(corpus_obj, shape, f"{context}.corpus")
    elapsed, sample_order = _validate_elapsed(result["elapsed_ns"], f"{context}.elapsed_ns")
    summary = _validate_source(result["source"], corpus, f"{context}.source")
    execution = _exact_keys(result["execution"], {"worker_count", "logical_tasks", "logical_bytes"}, f"{context}.execution")
    _expect(execution["worker_count"], 1, f"{context}.execution.worker_count")
    _expect(execution["logical_tasks"], 1, f"{context}.execution.logical_tasks")
    _expect(execution["logical_bytes"], corpus["archive_bytes"], f"{context}.execution.logical_bytes")
    _expect(result["output_sha256"], corpus["archive_sha256"], f"{context}.output_sha256")
    _digest(result["output_sha256"], f"{context}.output_sha256")
    _validate_operation(result["operation_metrics"], sample_order, f"{context}.operation_metrics")
    return RowEvidence(shape, corpus, summary, result["output_sha256"], elapsed, sample_order)


def _validate_tool(value: Any, context: str) -> None:
    tool = _exact_keys(value, CONTRACT_TOOL_KEYS, context)
    _expect(
        tool,
        {
            "name": TOOL_NAME,
            "version": TOOL_VERSION,
            "binary": NORMAL_BINARY,
            "profile": PROFILE,
            "target_os": TARGET_OS,
            "target_arch": TARGET_ARCH,
            "instrumentation": "none",
        },
        context,
    )


def _validate_binary(value: Any, role: str, contract: Contract, context: str) -> None:
    binary = _exact_keys(value, REPORT_BINARY_KEYS, context)
    expected = contract.legs[role]
    path = _string(binary["path"], f"{context}.path")
    if Path(path).name != NORMAL_BINARY:
        raise ValidationError(f"{context}.path must name {NORMAL_BINARY!r}")
    _expect(binary["binary_sha256"], expected["binary_sha256"], f"{context}.binary_sha256")
    _digest(binary["binary_sha256"], f"{context}.binary_sha256")
    _expect(binary["binary_bytes"], expected["binary_bytes"], f"{context}.binary_bytes")
    _expect(binary["mode_bits"], expected["mode_bits"], f"{context}.mode_bits")
    _expect(binary["executable"], True, f"{context}.executable")
    _expect(binary["profile"], PROFILE, f"{context}.profile")


def _validate_environment(value: Any, role: str, contract: Contract, context: str) -> dict[str, Any]:
    environment = _exact_keys(value, REPORT_ENVIRONMENT_KEYS, context)
    expected = contract.environment
    _expect(environment["rustc_version"], expected["rustc_version"], f"{context}.rustc_version")
    _expect(environment["git_revision"], contract.legs[role]["revision"], f"{context}.git_revision")
    _expect(environment["git_worktree_dirty"], False, f"{context}.git_worktree_dirty")
    _expect(environment["logical_cpus_available"], expected["logical_cpus_available"], f"{context}.logical_cpus_available")
    _expect(environment["allocator"], ALLOCATOR, f"{context}.allocator")
    _expect(environment["rustflags"], expected["rustflags"], f"{context}.rustflags")
    _expect(environment["cargo_build_target"], expected["cargo_build_target"], f"{context}.cargo_build_target")
    _expect(environment["os"], TARGET_OS, f"{context}.os")
    _expect(environment["cpu_model"], expected["cpu_model"], f"{context}.cpu_model")
    _expect(environment["cpu_affinity"], expected["cpu_affinity"], f"{context}.cpu_affinity")
    for key in ("perf_event_paranoid", "kernel", "filesystem_type", "cpu_affinity", "storage_identifier"):
        if environment[key] is not None:
            _string(environment[key], f"{context}.{key}")
    for key in ("total_memory_bytes", "page_size_bytes"):
        if environment[key] is not None:
            _integer(environment[key], f"{context}.{key}")
    if environment["source_destination_same_device"] is not None and not isinstance(environment["source_destination_same_device"], bool):
        raise ValidationError(f"{context}.source_destination_same_device must be boolean or null")
    return environment


def _expected_configuration() -> dict[str, Any]:
    return {
        "samples_per_case": SAMPLE_COUNT,
        "warmup_iterations_per_case": WARMUP_COUNT,
        "filesystem_cache_states": list(GLOBAL_CACHE_STATES),
        "filesystem_fresh_child_per_sample": True,
        "filesystem_process_isolated": True,
        "filesystem_root_selected": False,
        "cases": [CASE],
        "corpus_shapes": ["tiny", "many-small", "few-large", "wide-root"],
        "payload_kinds": ["compressible", "incompressible"],
        "writer_shapes": ["tiny", "large", "payload-heavy"],
        "xlsx_shapes": ["tiny", "medium", "dense-wide"],
        "xlsb_shapes": ["tiny", "medium", "large", "sparse"],
        "xlsx_cell_crud_shapes": ["medium", "dense-sparse"],
        "xlsx_row_visibility_shapes": ["medium", "large"],
        "semantic_shapes": ["tiny", "medium", "large"],
        "rtf_variants": ["plain"],
        "range_simulation": {
            "fixed_latency_us": 100,
            "request_overhead_us": 25,
            "bandwidth_bytes_per_second": 50 * 1024 * 1024,
            "max_physical_range_bytes": 4 * 1024,
        },
        "execution_workers": WORKERS,
    }


def _validate_configuration(value: Any, context: str) -> None:
    configuration = _exact_keys(value, CONFIGURATION_KEYS, context)
    # This list belongs to the producer's global harness configuration, not
    # to the selected in-process operation.  Keep the accepted envelopes
    # closed: a caller cannot introduce a state, reorder states, or duplicate
    # a state and have it disappear in the projection.
    cache_states = _array(configuration["filesystem_cache_states"], f"{context}.filesystem_cache_states")
    if tuple(cache_states) not in ACCEPTED_GLOBAL_CACHE_STATES:
        raise ValidationError(
            f"{context}.filesystem_cache_states must be exactly one of "
            f"{[list(states) for states in ACCEPTED_GLOBAL_CACHE_STATES]!r}; got {cache_states!r}"
        )
    expected = _expected_configuration()
    expected["filesystem_cache_states"] = cache_states
    _expect(configuration, expected, context)
    _exact_keys(configuration["range_simulation"], RANGE_SIMULATION_KEYS, f"{context}.range_simulation")
    for key in RANGE_SIMULATION_KEYS:
        _integer(configuration["range_simulation"][key], f"{context}.range_simulation.{key}")


def _validate_report(path: Path, role: str, contract: Contract) -> ValidatedReport:
    report, raw_sha256 = _read_json(path, role)
    _exact_keys(report, REPORT_KEYS, role)
    _expect(report["schema_version"], SCHEMA_VERSION, f"{role}.schema_version")
    _validate_tool(report["tool"], f"{role}.tool")
    _validate_binary(report["binary_identity"], role, contract, f"{role}.binary_identity")
    _validate_environment(report["environment"], role, contract, f"{role}.environment")
    _validate_configuration(report["configuration"], f"{role}.configuration")
    results = _array(report["results"], f"{role}.results")
    if len(results) != len(SHAPES):
        raise ValidationError(f"{role}.results must contain exactly {len(SHAPES)} fixed rows")
    rows: dict[str, RowEvidence] = {}
    for index, expected_shape in enumerate(SHAPES):
        row = _validate_result(results[index], role, index)
        if row.shape != expected_shape:
            raise ValidationError(f"{role}.results[{index}] is out of fixed shape order")
        rows[expected_shape] = row
    _validate_parallel(report["parallel_metrics"], f"{role}.parallel_metrics")
    return ValidatedReport(role, raw_sha256, report, rows)


def _environment_identity(report: Mapping[str, Any]) -> dict[str, Any]:
    environment = dict(report["environment"])
    environment.pop("git_revision", None)
    return environment


def _binary_identity(report: Mapping[str, Any]) -> dict[str, Any]:
    # Each leg may be staged under a different absolute path.  The contract
    # binds the executable's content/size/mode/profile, while the path is only
    # checked for the normal basename by _validate_binary.
    binary = dict(report["binary_identity"])
    binary.pop("path", None)
    return binary


def _validate_cross_report(reports: Mapping[str, ValidatedReport]) -> None:
    first = reports["A1"]
    first_cache_states = first.report["configuration"]["filesystem_cache_states"]
    for role in ROLES:
        if role == "A1":
            continue
        _expect(
            reports[role].report["configuration"]["filesystem_cache_states"],
            first_cache_states,
            f"{role}/A1 filesystem cache selector identity",
        )
        _expect(
            _environment_identity(reports[role].report),
            _environment_identity(first.report),
            f"{role}/A1 environment identity",
        )
        _expect(
            reports[role].report["configuration"],
            first.report["configuration"],
            f"{role}/A1 configuration identity",
        )
        _expect(
            reports[role].report["parallel_metrics"],
            first.report["parallel_metrics"],
            f"{role}/A1 parallel identity",
        )
    _expect(
        _binary_identity(reports["A1"].report),
        _binary_identity(reports["A2"].report),
        "control binary identity",
    )
    _expect(
        _binary_identity(reports["B1"].report),
        _binary_identity(reports["B2"].report),
        "candidate binary identity",
    )
    for shape in SHAPES:
        rows = {role: reports[role].rows[shape] for role in ROLES}
        for role in ("B1", "B2", "A2"):
            _expect(rows["A1"].corpus, rows[role].corpus, f"{shape} A1/{role} corpus identity")
            _expect(rows["A1"].summary, rows[role].summary, f"{shape} A1/{role} semantic identity")
            _expect(rows["A1"].output_sha256, rows[role].output_sha256, f"{shape} A1/{role} output oracle")


def _delta_percent(control: float, candidate: float, context: str) -> float:
    if control <= 0:
        raise ValidationError(f"{context} control statistic must be positive")
    result = (control - candidate) / control * 100.0
    if not math.isfinite(result):
        raise ValidationError(f"{context} candidate reduction is not finite")
    return result


def _drift_percent(first: float, second: float, context: str) -> float:
    if first <= 0:
        raise ValidationError(f"{context} baseline statistic must be positive")
    result = (second - first) / first * 100.0
    if not math.isfinite(result):
        raise ValidationError(f"{context} drift is not finite")
    return result


def _row_decision(rows: Mapping[str, RowEvidence], shape: str) -> dict[str, Any]:
    elapsed = {role: rows[role].elapsed for role in ROLES}
    candidate_reduction = {
        "a1_to_b1": {
            statistic: _delta_percent(
                float(elapsed["A1"][statistic]),
                float(elapsed["B1"][statistic]),
                f"{shape} A1_to_B1 {statistic}",
            )
            for statistic in STATISTICS
        },
        "a2_to_b2": {
            statistic: _delta_percent(
                float(elapsed["A2"][statistic]),
                float(elapsed["B2"][statistic]),
                f"{shape} A2_to_B2 {statistic}",
            )
            for statistic in STATISTICS
        },
    }
    drift = {
        "control": {
            statistic: _drift_percent(
                float(elapsed["A1"][statistic]),
                float(elapsed["A2"][statistic]),
                f"{shape} control {statistic}",
            )
            for statistic in STATISTICS
        },
        "candidate": {
            statistic: _drift_percent(
                float(elapsed["B1"][statistic]),
                float(elapsed["B2"][statistic]),
                f"{shape} candidate {statistic}",
            )
            for statistic in STATISTICS
        },
    }
    drift_within = {
        implementation: {
            statistic: abs(drift[implementation][statistic]) <= DRIFT_CEILINGS[statistic]
            for statistic in STATISTICS
        }
        for implementation in ("control", "candidate")
    }
    accepted: list[str] = []
    adverse_both: list[str] = []
    rejected: dict[str, str] = {}
    for statistic in STATISTICS:
        first_reduction = candidate_reduction["a1_to_b1"][statistic]
        second_reduction = candidate_reduction["a2_to_b2"][statistic]
        reasons: list[str] = []
        if first_reduction < 0 and second_reduction < 0:
            adverse_both.append(statistic)
            reasons.append("candidate is not lower in both paired directions")
        elif first_reduction <= 0 or second_reduction <= 0:
            if (first_reduction < 0 < second_reduction) or (second_reduction < 0 < first_reduction):
                reasons.append("paired directions disagree")
            else:
                reasons.append("candidate is not lower in both paired directions")
        for implementation in ("control", "candidate"):
            if not drift_within[implementation][statistic]:
                reasons.append(
                    f"{implementation} drift {drift[implementation][statistic]:+.6f}% "
                    f"exceeds {DRIFT_CEILINGS[statistic]:g}% ceiling"
                )
        if reasons:
            rejected[statistic] = "; ".join(reasons)
        else:
            accepted.append(statistic)
    return {
        "sample_count": SAMPLE_COUNT,
        "legs_ns": {role: elapsed[role] for role in ROLES},
        "candidate_reduction_percent": candidate_reduction,
        "same_implementation_drift_percent": drift,
        "drift_ceiling_percent": dict(DRIFT_CEILINGS),
        "drift_within_ceiling": drift_within,
        "adverse_both_statistics": adverse_both,
        "accepted_statistics": accepted,
        "rejected_statistics": rejected,
    }


def _vector_digest(values: Iterable[Any]) -> str:
    raw = json.dumps(list(values), separators=(",", ":"), sort_keys=False).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _validator_source_sha256() -> str:
    try:
        source = Path(__file__).read_bytes()
    except OSError as error:
        raise ValidationError(f"cannot read validator source: {error}") from error
    return hashlib.sha256(source).hexdigest()


def _projection(reports: Mapping[str, ValidatedReport], contract: Contract) -> dict[str, Any]:
    # The producer list is retained verbatim in the projection.  The selected
    # claim remains the warm in-process constructor only, even when the global
    # envelope names a cold-requested branch.
    producer_cache_states = reports["A1"].report["configuration"]["filesystem_cache_states"]
    rows = []
    for shape in SHAPES:
        row_map = {role: reports[role].rows[shape] for role in ROLES}
        decision = _row_decision(row_map, shape)
        a1 = row_map["A1"]
        rows.append(
            {
                "shape": shape,
                "corpus": dict(a1.corpus),
                "semantic_identity": {
                    "source_archive_sha256": a1.summary["source_archive_sha256"],
                    "part_names_sha256": a1.summary["part_names_sha256"],
                    "part_payload_sha256": a1.summary["part_payload_sha256"],
                    "target_name": a1.summary["target_name"],
                    "target_payload_sha256": a1.summary["target_payload_sha256"],
                    "part_count": a1.summary["part_count"],
                    "archive_member_count": a1.summary["archive_member_count"],
                    "semantic_vectors_verified": True,
                },
                "output_sha256": a1.output_sha256,
                "elapsed_ns": decision,
            }
        )
    return {
        "schema_version": SCHEMA_VERSION,
        "validator": VALIDATOR_NAME,
        "case": CASE,
        "protocol": {
            "abba_order": ABBA_ORDER,
            "samples_per_leg": SAMPLE_COUNT,
            "warmup_iterations_per_leg": WARMUP_COUNT,
            "cache_state": CACHE_STATE,
            "configuration_cache_states": list(producer_cache_states),
            "cache_claim_scope": CACHE_CLAIM_SCOPE,
            "execution_workers": WORKERS,
            "fixed_shapes": list(SHAPES),
            "operation_vector_alignment": ALIGNMENT,
            "sample_sort_key": "(elapsed_ns, original_sample_index)",
            "in_process_sample_semantics": True,
        },
        "provenance": {
            "contract_sha256": contract.raw_sha256 or None,
            "validator_source_sha256": _validator_source_sha256(),
            "binary_identity": {
                "mode": "reported+contract-bound",
                "file_rehashed": False,
                "reason": "report-declared normal binary hash/size/mode are checked against the narrow contract; executable paths are not opened",
            },
            "control": {
                "revision": contract.legs["A1"]["revision"],
                "binary_sha256": contract.legs["A1"]["binary_sha256"],
                "binary_bytes": contract.legs["A1"]["binary_bytes"],
                "mode_bits": contract.legs["A1"]["mode_bits"],
            },
            "candidate": {
                "revision": contract.legs["B1"]["revision"],
                "binary_sha256": contract.legs["B1"]["binary_sha256"],
                "binary_bytes": contract.legs["B1"]["binary_bytes"],
                "mode_bits": contract.legs["B1"]["mode_bits"],
            },
            "raw_report_sha256": {role: reports[role].raw_sha256 for role in ROLES},
        },
        "validation": {
            "report_count": 4,
            "distinct_raw_report_byte_streams": True,
            "matrix_rows": len(SHAPES),
            "samples_per_row": SAMPLE_COUNT,
            "elapsed_statistics_recomputed": True,
            "stable_sample_order_verified": True,
            "semantic_oracles_verified": True,
            "normal_instrumentation_verified": True,
            "binary_identity_reported_and_contract_bound": True,
            "model_is_expected_not_observed": True,
            "candidate_improvement_rule": "strictly lower in both A1_to_B1 and A2_to_B2",
            "drift_ceilings_verified": dict(DRIFT_CEILINGS),
            "decisions_preserved_per_row": True,
            "pooled_claim": False,
        },
        "claimability": {
            "normal_constructor_latency": {
                "claimable": True,
                "scope": CACHE_CLAIM_SCOPE,
            },
            "pooled_latency": {
                "claimable": False,
                "reason": "rows retain independent corpus/statistic decisions; no pooled matrix claim",
            },
            "rss": {"claimable": False, "reason": "operation process RSS is unavailable"},
            "process_counters": {"claimable": False, "reason": "operation process counters are unavailable"},
            "logical_io": {"claimable": False, "reason": "timed constructor has no logical ReadAt boundary"},
            "physical_io": {"claimable": False, "reason": "no physical-storage counters are collected"},
            "cold_cache": {"claimable": False, "reason": CACHE_CLAIM_SCOPE},
            "fresh_child_per_sample": {"claimable": False, "reason": "operation metrics have no child-process identity"},
            "allocator_elapsed_ns": {"claimable": False, "reason": "allocator vectors are unavailable in the normal binary"},
            "live_bytes_before": {"claimable": False, "reason": "allocator vectors are unavailable in the normal binary"},
            "live_bytes_after": {"claimable": False, "reason": "allocator vectors are unavailable in the normal binary"},
            "peak_live_bytes_before": {"claimable": False, "reason": "allocator vectors are unavailable in the normal binary"},
            "peak_live_bytes_after": {"claimable": False, "reason": "allocator vectors are unavailable in the normal binary"},
            "allocator": {"claimable": False, "reason": "normal binary allocator vectors are unavailable"},
        },
        "rows": rows,
    }


def validate_paths(
    a1: Path,
    b1: Path,
    b2: Path,
    a2: Path,
    contract_path: Path | None = None,
    *,
    contract: Contract | None = None,
) -> dict[str, Any]:
    if contract is not None and contract_path is not None:
        raise ValidationError("provide either contract_path or contract, not both")
    if contract is not None:
        checked_contract = _validate_contract(contract)
    elif contract_path is not None:
        checked_contract = _validate_contract(contract_path)
    else:
        checked_contract = None
    if checked_contract is None:
        raise ValidationError("a narrow contract or explicit identity contract is required")
    paths = {"A1": a1, "B1": b1, "B2": b2, "A2": a2}
    if len({path.resolve() for path in paths.values()}) != len(ROLES):
        raise ValidationError("the four ABBA roles must use four distinct report paths")
    reports = {role: _validate_report(path, role, checked_contract) for role, path in paths.items()}
    if len({report.raw_sha256 for report in reports.values()}) != len(ROLES):
        raise ValidationError("the four ABBA roles must use four distinct raw report byte streams")
    _validate_cross_report(reports)
    return _projection(reports, checked_contract)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--contract", type=Path, help="narrow JSON identity/oracle contract")
    parser.add_argument("--a1", type=Path, required=True, help="control A1 normal-latency report")
    parser.add_argument("--b1", type=Path, required=True, help="candidate B1 normal-latency report")
    parser.add_argument("--b2", type=Path, required=True, help="candidate B2 normal-latency report")
    parser.add_argument("--a2", type=Path, required=True, help="control A2 normal-latency report")
    parser.add_argument("--output", type=Path, help="optional projection output path")
    parser.add_argument("--control-revision")
    parser.add_argument("--candidate-revision")
    parser.add_argument("--control-binary-sha256")
    parser.add_argument("--candidate-binary-sha256")
    parser.add_argument("--control-binary-bytes", type=int)
    parser.add_argument("--candidate-binary-bytes", type=int)
    parser.add_argument("--control-mode-bits", type=int)
    parser.add_argument("--candidate-mode-bits", type=int)
    parser.add_argument("--rustc-version")
    parser.add_argument("--logical-cpus", type=int)
    parser.add_argument("--cpu-model")
    parser.add_argument("--cpu-affinity")
    parser.add_argument("--rustflags")
    parser.add_argument("--cargo-build-target")
    return parser


def _reject_contract_identity_mix(args: argparse.Namespace) -> None:
    if args.contract is None:
        return
    names = (
        "control_revision",
        "candidate_revision",
        "control_binary_sha256",
        "candidate_binary_sha256",
        "control_binary_bytes",
        "candidate_binary_bytes",
        "control_mode_bits",
        "candidate_mode_bits",
        "rustc_version",
        "logical_cpus",
        "cpu_model",
        "cpu_affinity",
        "rustflags",
        "cargo_build_target",
    )
    supplied = [name for name in names if getattr(args, name) is not None]
    if supplied:
        raise ValidationError(
            "--contract cannot be combined with explicit identity options: "
            + ", ".join(f"--{name.replace('_', '-')}" for name in supplied)
        )


def _reject_output_collision(output: Path | None, input_paths: Iterable[Path | None]) -> None:
    if output is None:
        return
    for input_path in input_paths:
        if input_path is None:
            continue
        if output.resolve() == input_path.resolve():
            raise ValidationError("--output must differ from contract and report inputs")
        try:
            if output.exists() and input_path.exists() and output.samefile(input_path):
                raise ValidationError("--output must differ from contract and report inputs")
        except OSError as error:
            raise ValidationError(f"cannot compare --output with input {input_path}: {error}") from error


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        _reject_contract_identity_mix(args)
        checked_contract = _validate_contract(args.contract) if args.contract is not None else _contract_from_cli(args)
        input_paths = (args.contract, args.a1, args.b1, args.b2, args.a2)
        _reject_output_collision(args.output, input_paths)
        result = validate_paths(args.a1, args.b1, args.b2, args.a2, contract=checked_contract)
        encoded = json.dumps(result, sort_keys=True, separators=(",", ":"))
        if args.output is None:
            print(encoded)
        else:
            args.output.write_text(encoded + "\n", encoding="utf-8")
            print(encoded)
    except (OSError, ValidationError) as error:
        print(f"OPC serial eager normal-latency ABBA validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
