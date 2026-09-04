#!/usr/bin/env python3
"""Validate allocator-only ABBA evidence for the OPC no-op overlay matrix.

The validator intentionally has no dependencies outside Python's standard
library.  It accepts a small contract for the final revision and executable
identities, then validates four distinct raw report byte streams (A1, B1, B2,
A2).
The reports must contain the fixed three-shape by three-overlay-count matrix.

Allocator elapsed samples and absolute live/high-water snapshots are parsed
only for alignment/cardinality and are explicitly non-claimable.  The only
derived comparison is the exact operation-scoped six-counter delta required by
the matrix contract.  These in-process reports do not carry per-sample
envelopes or PIDs, so the projection does not claim independent-process proof.
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
VALIDATOR_NAME = "litchi-opc-overlay-allocator-abba"
CASE = "opc_source_overlay_multi_part_noop"
CACHE_STATE = "warm"
SAMPLE_COUNT = 30
WARMUP_COUNT = 3
WORKERS = [1]
ABBA_ORDER = ["A1_control", "B1_candidate", "B2_candidate", "A2_control"]
ABBA_ROLES = ("A1", "B1", "B2", "A2")
SHAPES = ("overlay-small", "overlay-large", "overlay-media-incompressible")
COUNTS = (2, 8, 32)
GENERATOR = "litchi-opc-source-overlay-multi-part-v1"
ALLOCATION_SCOPE = "operation_global_system_allocator"
ALLOCATOR = "CountingSystemAllocator(std::alloc::System)"
TOOL_NAME = "litchi-perf-baseline"
TOOL_VERSION = "0.1.0"
ALLOCATOR_BINARY = "litchi-perf-baseline-alloc"
PROFILE = "release"
TARGET_OS = "linux"
TARGET_ARCH = "x86_64"
ALIGNMENT = "elapsed_ns.samples_by_elapsed_then_sample_index"
LATENCY_CLAIM = "comparable_timed_operation"
CPU_AFFINITY = "2"
SHA256 = re.compile(r"[0-9a-f]{64}\Z")
REVISION = re.compile(r"[0-9a-f]{40}\Z")

# This is the retained corpus manifest for the fixed OPC generator. These
# identities were captured from the compiled multi-part changed smoke report;
# keeping them in the validator makes the contract an identity declaration,
# not an oracle that a caller can silently redefine.
CORPUS_ORACLE = {
    "overlay-small": {
        "archive_member_count": 34,
        "entry_bytes": 1_024,
        "uncompressed_payload_bytes": 32_768,
        "archive_bytes": 7_451,
        "archive_sha256": "4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509",
        "target_entry": "benchmark/parts/00016.bin",
        "target_payload_bytes": 1_024,
        "target_payload_sha256": "5b7b9793a43d08ca2c0670289d932541377407ed352ab9f6c145f63d19de9f98",
    },
    "overlay-large": {
        "archive_member_count": 34,
        "entry_bytes": 64 * 1_024,
        "uncompressed_payload_bytes": 2 * 1_024 * 1_024,
        "archive_bytes": 2_103_195,
        "archive_sha256": "8356d7467215b04a3d1c3703f50fbd6322f2002ca7c3ead1f24414c5e550ef73",
        "target_entry": "benchmark/parts/00016.bin",
        "target_payload_bytes": 64 * 1_024,
        "target_payload_sha256": "e17b543eec6b4d3534978d7d59e7240dcbf0f2a2050fd80f32ea3daec266aa73",
    },
    "overlay-media-incompressible": {
        "archive_member_count": 34,
        "entry_bytes": 256 * 1_024,
        "uncompressed_payload_bytes": 8 * 1_024 * 1_024,
        "archive_bytes": 8_396_580,
        "archive_sha256": "bf8c309af5306c6682b9df65b97246f81b022fe5e3b5e02cc2c4dcf3e1e87883",
        "target_entry": "benchmark/parts/00016.bin",
        "target_payload_bytes": 256 * 1_024,
        "target_payload_sha256": "3ad07c7e34d3dd6d9ff75b696ccbdd702777b6e4dea04b19bbe3d0aa6d21cdeb",
    },
}

ALLOCATION_METRICS = (
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
)
COUNTER_METRICS = ALLOCATION_METRICS[:6]
NONCLAIMABLE_METRICS = ALLOCATION_METRICS[6:]
EXPECTED_DELTA_BY_COUNT = {
    2: {
        "allocation_calls": -2,
        "deallocation_calls": -2,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": -80_320,
        "deallocated_bytes": -80_320,
    },
    8: {
        "allocation_calls": -14,
        "deallocation_calls": -14,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": -562_240,
        "deallocated_bytes": -562_240,
    },
    32: {
        "allocation_calls": -62,
        "deallocation_calls": -62,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": -2_489_920,
        "deallocated_bytes": -2_489_920,
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
CONTRACT_CORPUS_KEYS = {
    "shape",
    "payload_kind",
    "base_name",
    "generator",
    "package_format",
    "compression",
    "entry_count",
    "archive_member_count",
    "entry_bytes",
    "uncompressed_payload_bytes",
    "archive_bytes",
    "archive_sha256",
    "target_entry",
    "target_payload_bytes",
    "target_payload_sha256",
}
REPORT_KEYS = {
    "schema_version",
    "tool",
    "binary_identity",
    "environment",
    "configuration",
    "parallel_metrics",
    "results",
}
REPORT_TOOL_KEYS = CONTRACT_TOOL_KEYS
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
RESULT_KEYS = {
    "case",
    "corpus",
    "elapsed_ns",
    "sink",
    "source",
    "output_sha256",
    "operation_metrics",
}
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
SOURCE_KEYS = {
    "read_calls",
    "read_bytes",
    "ordinary_payload_read_calls",
    "ordinary_payload_read_bytes",
    "max_in_flight_reads",
    "opc_source_overlay",
}
SINK_KEYS = {"accepted_bytes", "write_calls", "largest_write", "write_size_buckets"}
SINK_BUCKET_KEYS = {
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
}
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
ALLOCATION_KEYS = {"status", "scope", *ALLOCATION_METRICS}
METRIC_VECTOR_KEYS = {"values", "status", "scope"}
ABSENT_METRIC_VECTOR_KEYS = {"status", "scope"}
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
OPERATION_PUBLICATION_KEYS = {"status", "changed_spans", "published_bytes"}
OPERATION_MATERIALIZATION_KEYS = {"status", "opc_parts"}
OPERATION_CFB_KEYS = {"status", "open", "plan", "atomic_publication"}
OPERATION_CFB_PHASE_KEYS = {
    "elapsed_ns",
    "logical_read_calls",
    "logical_read_requested_bytes",
    "logical_read_returned_bytes",
}
OVERLAY_KEYS = {
    "implementation",
    "timing_scope",
    "performance_claim",
    "overlay_mode",
    "replacement_semantics",
    "overlay_count",
    "source_shape",
    "payload_kind",
    "source_bytes",
    "source_sha256",
    "expected_eager_sha256",
    "source_cache_max_bytes",
    "source_cache_max_entries",
    "sink_max_bytes",
    "sink_max_write",
    "preparation_ns",
    "open_ns",
    "planning_ns",
    "publication_ns",
    "cache_before_publication_hits",
    "cache_before_publication_cold_loads",
    "cache_before_publication_retained_entries",
    "cache_before_publication_retained_bytes",
    "source_cache_after_publication_probe_hits",
    "source_cache_after_publication_probe_cold_loads",
    "source_cache_after_publication_probe_retained_entries",
    "source_cache_after_publication_probe_retained_bytes",
    "reopened_output_cache_hits",
    "reopened_output_cache_cold_loads",
    "reopened_output_cache_retained_entries",
    "reopened_output_cache_retained_bytes",
    "observed_after_publication_source_read_calls",
    "observed_after_publication_source_read_bytes",
    "observed_after_publication_ordinary_payload_read_calls",
    "observed_after_publication_ordinary_payload_read_bytes",
    "expected_eager_semantic_verified",
    "raw_members_and_order_preservation_verified",
    "equal_payload_noop_source_verified",
    "observed_output_sha256",
}
PARALLEL_KEYS = {
    "schema_version",
    "scope",
    "claim",
    "configured_worker_budget",
    "observed_process_thread_count",
    "cases",
}


class ValidationError(ValueError):
    """Raised when the contract or evidence cannot be accepted."""


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
    count: int
    elapsed_samples: list[int]
    sample_order: list[int]
    allocation: dict[str, list[int]]
    source_identity: dict[str, Any]
    sink_identity: dict[str, Any]
    output_sha256: str


@dataclass(frozen=True)
class ValidatedReport:
    role: str
    raw_sha256: str
    report: dict[str, Any]
    rows: dict[tuple[str, int], RowEvidence]


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


def _integer(value: Any, context: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise ValidationError(f"{context} must be an integer >= {minimum}")
    return value


def _finite_number(value: Any, context: str) -> int | float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValidationError(f"{context} must be a finite number")
    if isinstance(value, float) and not math.isfinite(value):
        raise ValidationError(f"{context} must be a finite number")
    return value


def _signed_integer(value: Any, context: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValidationError(f"{context} must be an integer")
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
    if actual != expected or type(actual) is not type(expected):
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


def _validate_contract(path: Path) -> Contract:
    document, raw_sha256 = _read_json(path, "contract")
    contract = _exact_keys(document, CONTRACT_KEYS, "contract")
    _expect(contract["schema_version"], SCHEMA_VERSION, "contract.schema_version")
    _expect(contract["case"], CASE, "contract.case")
    _expect(contract["cache_state"], CACHE_STATE, "contract.cache_state")
    _expect(contract["samples_per_case"], SAMPLE_COUNT, "contract.samples_per_case")
    _expect(
        contract["warmup_iterations_per_case"],
        WARMUP_COUNT,
        "contract.warmup_iterations_per_case",
    )
    _expect(contract["execution_workers"], WORKERS, "contract.execution_workers")
    _expect(contract["abba_order"], ABBA_ORDER, "contract.abba_order")

    tool = _exact_keys(contract["tool"], CONTRACT_TOOL_KEYS, "contract.tool")
    _expect(tool["name"], TOOL_NAME, "contract.tool.name")
    _expect(tool["version"], TOOL_VERSION, "contract.tool.version")
    _expect(tool["binary"], ALLOCATOR_BINARY, "contract.tool.binary")
    _expect(tool["profile"], PROFILE, "contract.tool.profile")
    _expect(tool["target_os"], TARGET_OS, "contract.tool.target_os")
    _expect(tool["target_arch"], TARGET_ARCH, "contract.tool.target_arch")
    _expect(
        tool["instrumentation"],
        "system_allocator_operation_scoped",
        "contract.tool.instrumentation",
    )

    environment = _exact_keys(
        contract["environment"], CONTRACT_ENVIRONMENT_KEYS, "contract.environment"
    )
    _string(environment["rustc_version"], "contract.environment.rustc_version")
    _expect(environment["allocator"], ALLOCATOR, "contract.environment.allocator")
    _expect(environment["target_os"], TARGET_OS, "contract.environment.target_os")
    _expect(environment["target_arch"], TARGET_ARCH, "contract.environment.target_arch")
    _integer(
        environment["logical_cpus_available"],
        "contract.environment.logical_cpus_available",
        minimum=1,
    )
    if environment["rustflags"] is not None:
        _string(environment["rustflags"], "contract.environment.rustflags")
    if environment["cargo_build_target"] is not None:
        _string(environment["cargo_build_target"], "contract.environment.cargo_build_target")

    raw_legs = _exact_keys(contract["legs"], set(ABBA_ROLES), "contract.legs")
    legs: dict[str, dict[str, Any]] = {}
    for role in ABBA_ROLES:
        item = _exact_keys(raw_legs[role], CONTRACT_LEG_KEYS, f"contract.legs.{role}")
        _expect(
            item["implementation"],
            "control" if role in ("A1", "A2") else "candidate",
            f"contract.legs.{role}.implementation",
        )
        _revision(item["revision"], f"contract.legs.{role}.revision")
        _digest(item["binary_sha256"], f"contract.legs.{role}.binary_sha256")
        _integer(item["binary_bytes"], f"contract.legs.{role}.binary_bytes", minimum=1)
        _integer(item["mode_bits"], f"contract.legs.{role}.mode_bits", minimum=1)
        _expect(item["profile"], PROFILE, f"contract.legs.{role}.profile")
        legs[role] = item
    for first, second, label in (("A1", "A2", "control"), ("B1", "B2", "candidate")):
        _expect(legs[first], legs[second], f"contract {label} A/B leg identity")
    if legs["A1"]["revision"] == legs["B1"]["revision"]:
        raise ValidationError("contract control and candidate revisions must differ")
    if legs["A1"]["binary_sha256"] == legs["B1"]["binary_sha256"]:
        raise ValidationError("contract control and candidate allocator binaries must differ")

    raw_corpora = _array(contract["corpora"], "contract.corpora")
    if len(raw_corpora) != len(SHAPES):
        raise ValidationError(f"contract.corpora must contain exactly {len(SHAPES)} shapes")
    corpora: dict[str, dict[str, Any]] = {}
    expected_payloads = {
        "overlay-small": "compressible",
        "overlay-large": "incompressible",
        "overlay-media-incompressible": "incompressible",
    }
    for index, raw_corpus in enumerate(raw_corpora):
        corpus = _exact_keys(raw_corpus, CONTRACT_CORPUS_KEYS, f"contract.corpora[{index}]")
        shape = _string(corpus["shape"], f"contract.corpora[{index}].shape")
        if shape not in SHAPES or shape in corpora:
            raise ValidationError(f"contract.corpora has an invalid or duplicate shape {shape!r}")
        _expect(corpus["payload_kind"], expected_payloads[shape], f"contract.corpora[{index}].payload_kind")
        _expect(corpus["base_name"], f"{shape}-{corpus['payload_kind']}", f"contract.corpora[{index}].base_name")
        _expect(corpus["generator"], GENERATOR, f"contract.corpora[{index}].generator")
        _expect(corpus["package_format"], "OPC/ZIP", f"contract.corpora[{index}].package_format")
        _expect(corpus["compression"], "deflate", f"contract.corpora[{index}].compression")
        _expect(corpus["entry_count"], 32, f"contract.corpora[{index}].entry_count")
        _integer(corpus["archive_member_count"], f"contract.corpora[{index}].archive_member_count", minimum=1)
        entry_bytes = _integer(corpus["entry_bytes"], f"contract.corpora[{index}].entry_bytes", minimum=1)
        _expect(corpus["uncompressed_payload_bytes"], 32 * entry_bytes, f"contract.corpora[{index}].uncompressed_payload_bytes")
        _integer(corpus["archive_bytes"], f"contract.corpora[{index}].archive_bytes", minimum=1)
        _digest(corpus["archive_sha256"], f"contract.corpora[{index}].archive_sha256")
        _string(corpus["target_entry"], f"contract.corpora[{index}].target_entry")
        _integer(corpus["target_payload_bytes"], f"contract.corpora[{index}].target_payload_bytes", minimum=1)
        _digest(corpus["target_payload_sha256"], f"contract.corpora[{index}].target_payload_sha256")
        for key, expected in CORPUS_ORACLE[shape].items():
            _expect(corpus[key], expected, f"contract.corpora[{index}].{key} retained corpus identity")
        corpora[shape] = corpus
    if set(corpora) != set(SHAPES):
        raise ValidationError("contract.corpora must contain the fixed three overlay shapes")

    raw_deltas = _exact_keys(contract["expected_deltas"], {str(count) for count in COUNTS}, "contract.expected_deltas")
    for count in COUNTS:
        actual = _exact_keys(raw_deltas[str(count)], set(COUNTER_METRICS), f"contract.expected_deltas.{count}")
        _expect(actual, EXPECTED_DELTA_BY_COUNT[count], f"contract.expected_deltas.{count}")
    return Contract(raw_sha256, tool, environment, legs, corpora)


def _validate_tool(value: Any, context: str) -> None:
    tool = _exact_keys(value, REPORT_TOOL_KEYS, context)
    expected = {
        "name": TOOL_NAME,
        "version": TOOL_VERSION,
        "binary": ALLOCATOR_BINARY,
        "profile": PROFILE,
        "target_os": TARGET_OS,
        "target_arch": TARGET_ARCH,
        "instrumentation": "system_allocator_operation_scoped",
    }
    _expect(tool, expected, context)


def _validate_binary(value: Any, role: str, contract: Contract, context: str) -> dict[str, Any]:
    binary = _exact_keys(value, REPORT_BINARY_KEYS, context)
    expected = contract.legs[role]
    path = _string(binary["path"], f"{context}.path")
    if Path(path).name != ALLOCATOR_BINARY:
        raise ValidationError(f"{context}.path must name {ALLOCATOR_BINARY!r}")
    _expect(binary["binary_sha256"], expected["binary_sha256"], f"{context}.binary_sha256")
    _digest(binary["binary_sha256"], f"{context}.binary_sha256")
    _expect(binary["binary_bytes"], expected["binary_bytes"], f"{context}.binary_bytes")
    _expect(binary["mode_bits"], expected["mode_bits"], f"{context}.mode_bits")
    _expect(binary["executable"], True, f"{context}.executable")
    _expect(binary["profile"], PROFILE, f"{context}.profile")
    return binary


def _validate_environment(value: Any, role: str, contract: Contract, context: str) -> dict[str, Any]:
    environment = _exact_keys(value, REPORT_ENVIRONMENT_KEYS, context)
    expected = contract.environment
    _expect(environment["rustc_version"], expected["rustc_version"], f"{context}.rustc_version")
    _expect(environment["git_revision"], contract.legs[role]["revision"], f"{context}.git_revision")
    _expect(environment["git_worktree_dirty"], False, f"{context}.git_worktree_dirty")
    _expect(environment["logical_cpus_available"], expected["logical_cpus_available"], f"{context}.logical_cpus_available")
    _expect(environment["allocator"], expected["allocator"], f"{context}.allocator")
    _expect(environment["rustflags"], expected["rustflags"], f"{context}.rustflags")
    _expect(environment["cargo_build_target"], expected["cargo_build_target"], f"{context}.cargo_build_target")
    _expect(environment["os"], expected["target_os"], f"{context}.os")
    for key in ("perf_event_paranoid", "kernel", "cpu_model", "filesystem_type", "cpu_affinity", "storage_identifier"):
        if environment[key] is not None:
            _string(environment[key], f"{context}.{key}")
    _expect(environment["cpu_affinity"], CPU_AFFINITY, f"{context}.cpu_affinity")
    for key in ("total_memory_bytes", "page_size_bytes"):
        if environment[key] is not None:
            _integer(environment[key], f"{context}.{key}")
    if environment["source_destination_same_device"] is not None and not isinstance(environment["source_destination_same_device"], bool):
        raise ValidationError(f"{context}.source_destination_same_device must be a boolean or null")
    return environment


def _validate_configuration(value: Any, context: str) -> dict[str, Any]:
    configuration = _exact_keys(value, CONFIGURATION_KEYS, context)
    fixed = {
        "samples_per_case": SAMPLE_COUNT,
        "warmup_iterations_per_case": WARMUP_COUNT,
        "filesystem_cache_states": [CACHE_STATE],
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
        "execution_workers": WORKERS,
        "range_simulation": {
            "fixed_latency_us": 100,
            "request_overhead_us": 25,
            "bandwidth_bytes_per_second": 50 * 1024 * 1024,
            "max_physical_range_bytes": 4 * 1024,
        },
    }
    for key, expected in fixed.items():
        _expect(configuration[key], expected, f"{context}.{key}")
    return configuration


def _validate_elapsed(value: Any, context: str) -> tuple[list[int], list[int]]:
    elapsed = _exact_keys(value, ELAPSED_KEYS, context)
    _expect(elapsed["unit"], "ns", f"{context}.unit")
    samples = [_integer(item, f"{context}.samples[{i}]") for i, item in enumerate(_array(elapsed["samples"], f"{context}.samples"))]
    if len(samples) != SAMPLE_COUNT or any(item == 0 for item in samples):
        raise ValidationError(f"{context}.samples must contain {SAMPLE_COUNT} positive integers")
    if samples != sorted(samples):
        raise ValidationError(f"{context}.samples must be sorted in elapsed order")
    order = [_integer(item, f"{context}.sample_order[{i}]") for i, item in enumerate(_array(elapsed["sample_order"], f"{context}.sample_order"))]
    if sorted(order) != list(range(SAMPLE_COUNT)):
        raise ValidationError(f"{context}.sample_order must be an exact 0..{SAMPLE_COUNT - 1} permutation")
    for index in range(1, SAMPLE_COUNT):
        if samples[index - 1] == samples[index] and order[index - 1] > order[index]:
            raise ValidationError(f"{context}.sample_order does not resolve elapsed ties by sample index")
    # The aggregate elapsed statistics are retained by the harness for schema
    # compatibility, but allocator instrumentation makes them non-claimable.
    # Parse them for finite JSON only; never use their values in acceptance.
    for key in ("min", "p50", "p95", "p99", "max", "mean", "standard_deviation"):
        _finite_number(elapsed[key], f"{context}.{key}")
    confidence = _exact_keys(elapsed["confidence_interval_95"], {"method", "lower", "upper"}, f"{context}.confidence_interval_95")
    _string(confidence["method"], f"{context}.confidence_interval_95.method")
    _finite_number(confidence["lower"], f"{context}.confidence_interval_95.lower")
    _finite_number(confidence["upper"], f"{context}.confidence_interval_95.upper")
    return samples, order


def _validate_metric_vector(value: Any, context: str, *, sample_count: int, status: str | None = None, scope: str | None = None) -> list[int]:
    vector = _exact_keys(value, METRIC_VECTOR_KEYS, context)
    if status is not None:
        _expect(vector["status"], status, f"{context}.status")
    if scope is not None:
        _expect(vector["scope"], scope, f"{context}.scope")
    values = [_integer(item, f"{context}.values[{i}]") for i, item in enumerate(_array(vector["values"], f"{context}.values"))]
    if len(values) != sample_count:
        raise ValidationError(f"{context}.values must contain {sample_count} values")
    return values


def _validate_absent_vector(value: Any, context: str, status: str) -> None:
    vector = _exact_keys(value, ABSENT_METRIC_VECTOR_KEYS, context)
    _expect(vector["status"], status, f"{context}.status")
    _string(vector["scope"], f"{context}.scope")


def _validate_sink(value: Any, context: str, *, max_bytes: int) -> dict[str, Any]:
    sink = _exact_keys(value, SINK_KEYS, context)
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        _integer(sink[key], f"{context}.{key}")
    buckets = _exact_keys(sink["write_size_buckets"], SINK_BUCKET_KEYS, f"{context}.write_size_buckets")
    for key in SINK_BUCKET_KEYS:
        _integer(buckets[key], f"{context}.write_size_buckets.{key}")
    if sink["accepted_bytes"] == 0 or sink["write_calls"] == 0:
        raise ValidationError(f"{context} must record non-empty output writes")
    if sink["accepted_bytes"] > max_bytes:
        raise ValidationError(
            f"{context}.accepted_bytes must be <= sink_max_bytes ({max_bytes})"
        )
    if sink["largest_write"] > sink["accepted_bytes"]:
        raise ValidationError(
            f"{context}.largest_write must be <= accepted_bytes ({sink['accepted_bytes']})"
        )
    if sink["largest_write"] > 65_536:
        raise ValidationError(f"{context}.largest_write exceeds the fixed sink maximum")
    if sum(buckets.values()) != sink["write_calls"]:
        raise ValidationError(f"{context}.write_size_buckets do not sum to write_calls")
    return sink


def _validate_source(value: Any, corpus: Mapping[str, Any], count: int, context: str) -> tuple[dict[str, Any], dict[str, Any]]:
    source = _exact_keys(value, SOURCE_KEYS, context)
    vectors: dict[str, list[int]] = {}
    for key in ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads"):
        vectors[key] = [_integer(item, f"{context}.{key}[{i}]") for i, item in enumerate(_array(source[key], f"{context}.{key}"))]
        if len(vectors[key]) != SAMPLE_COUNT:
            raise ValidationError(f"{context}.{key} must contain {SAMPLE_COUNT} values")
    if any(value == 0 for value in vectors["read_calls"]) or any(value == 0 for value in vectors["read_bytes"]):
        raise ValidationError(f"{context} must record ordinary source reads")
    if any(value < count for value in vectors["ordinary_payload_read_calls"]):
        raise ValidationError(
            f"{context}.ordinary_payload_read_calls must be at least overlay_count ({count})"
        )
    if any(value <= 0 for value in vectors["ordinary_payload_read_bytes"]):
        raise ValidationError(
            f"{context}.ordinary_payload_read_bytes must be positive for every sample"
        )

    overlay = _exact_keys(source["opc_source_overlay"], OVERLAY_KEYS, f"{context}.opc_source_overlay")
    _expect(overlay["implementation"], "SourceBackedPackage::write_part_overlays_to_stream", f"{context}.opc_source_overlay.implementation")
    timing_scope = _string(overlay["timing_scope"], f"{context}.opc_source_overlay.timing_scope")
    if "preparation_ns, open_ns, planning_ns, and publication_ns" not in timing_scope or "operation_metrics.allocation covers only the write_part_overlays_to_stream publication call" not in timing_scope:
        raise ValidationError(f"{context}.opc_source_overlay.timing_scope does not bind the four phases and allocation boundary")
    _expect(overlay["performance_claim"], "none", f"{context}.opc_source_overlay.performance_claim")
    _expect(overlay["overlay_mode"], "noop", f"{context}.opc_source_overlay.overlay_mode")
    _expect(overlay["replacement_semantics"], "non-empty equal-payload replacement plan; semantic no-op", f"{context}.opc_source_overlay.replacement_semantics")
    _expect(overlay["overlay_count"], count, f"{context}.opc_source_overlay.overlay_count")
    _expect(overlay["source_shape"], corpus["shape"], f"{context}.opc_source_overlay.source_shape")
    _expect(overlay["payload_kind"], corpus["payload_kind"], f"{context}.opc_source_overlay.payload_kind")
    _expect(overlay["source_bytes"], corpus["archive_bytes"], f"{context}.opc_source_overlay.source_bytes")
    _expect(overlay["source_sha256"], corpus["archive_sha256"], f"{context}.opc_source_overlay.source_sha256")
    _expect(overlay["expected_eager_sha256"], corpus["archive_sha256"], f"{context}.opc_source_overlay.expected_eager_sha256")
    _expect(overlay["source_cache_max_bytes"], corpus["uncompressed_payload_bytes"], f"{context}.opc_source_overlay.source_cache_max_bytes")
    _expect(overlay["source_cache_max_entries"], corpus["entry_count"], f"{context}.opc_source_overlay.source_cache_max_entries")
    _expect(overlay["sink_max_bytes"], corpus["archive_bytes"] * 2 + 65_536, f"{context}.opc_source_overlay.sink_max_bytes")
    _expect(overlay["sink_max_write"], 65_536, f"{context}.opc_source_overlay.sink_max_write")
    numeric_overlay_vectors = (
        "preparation_ns", "open_ns", "planning_ns", "publication_ns",
        "cache_before_publication_hits", "cache_before_publication_cold_loads",
        "cache_before_publication_retained_entries", "cache_before_publication_retained_bytes",
        "source_cache_after_publication_probe_hits", "source_cache_after_publication_probe_cold_loads",
        "source_cache_after_publication_probe_retained_entries", "source_cache_after_publication_probe_retained_bytes",
        "reopened_output_cache_hits", "reopened_output_cache_cold_loads",
        "reopened_output_cache_retained_entries", "reopened_output_cache_retained_bytes",
        "observed_after_publication_source_read_calls", "observed_after_publication_source_read_bytes",
        "observed_after_publication_ordinary_payload_read_calls", "observed_after_publication_ordinary_payload_read_bytes",
    )
    overlay_vectors: dict[str, list[int]] = {}
    for key in numeric_overlay_vectors:
        overlay_vectors[key] = [_integer(item, f"{context}.opc_source_overlay.{key}[{i}]") for i, item in enumerate(_array(overlay[key], f"{context}.opc_source_overlay.{key}"))]
        if len(overlay_vectors[key]) != SAMPLE_COUNT:
            raise ValidationError(f"{context}.opc_source_overlay.{key} must contain {SAMPLE_COUNT} values")
    _expect(overlay_vectors["cache_before_publication_hits"], [0] * SAMPLE_COUNT, f"{context}.opc_source_overlay.cache_before_publication_hits")
    _expect(overlay_vectors["cache_before_publication_cold_loads"], [0] * SAMPLE_COUNT, f"{context}.opc_source_overlay.cache_before_publication_cold_loads")
    _expect(overlay_vectors["cache_before_publication_retained_entries"], [0] * SAMPLE_COUNT, f"{context}.opc_source_overlay.cache_before_publication_retained_entries")
    _expect(overlay_vectors["cache_before_publication_retained_bytes"], [0] * SAMPLE_COUNT, f"{context}.opc_source_overlay.cache_before_publication_retained_bytes")
    _expect(overlay_vectors["source_cache_after_publication_probe_hits"], [0] * SAMPLE_COUNT, f"{context}.opc_source_overlay.source_cache_after_publication_probe_hits")
    _expect(overlay_vectors["source_cache_after_publication_probe_cold_loads"], [count] * SAMPLE_COUNT, f"{context}.opc_source_overlay.source_cache_after_publication_probe_cold_loads")
    _expect(overlay_vectors["source_cache_after_publication_probe_retained_entries"], [count] * SAMPLE_COUNT, f"{context}.opc_source_overlay.source_cache_after_publication_probe_retained_entries")
    expected_probe_bytes = count * corpus["entry_bytes"]
    _expect(overlay_vectors["source_cache_after_publication_probe_retained_bytes"], [expected_probe_bytes] * SAMPLE_COUNT, f"{context}.opc_source_overlay.source_cache_after_publication_probe_retained_bytes")
    for key in ("reopened_output_cache_hits", "reopened_output_cache_cold_loads", "reopened_output_cache_retained_entries", "reopened_output_cache_retained_bytes"):
        _expect(overlay_vectors[key], [0] * SAMPLE_COUNT, f"{context}.opc_source_overlay.{key}")
    observed_source_pairs = {
        "observed_after_publication_source_read_calls": "read_calls",
        "observed_after_publication_source_read_bytes": "read_bytes",
        "observed_after_publication_ordinary_payload_read_calls": "ordinary_payload_read_calls",
        "observed_after_publication_ordinary_payload_read_bytes": "ordinary_payload_read_bytes",
    }
    for overlay_key, source_key in observed_source_pairs.items():
        _expect(overlay_vectors[overlay_key], vectors[source_key], f"{context}.opc_source_overlay.{overlay_key}/source.{source_key}")
    for key in ("expected_eager_semantic_verified", "raw_members_and_order_preservation_verified", "equal_payload_noop_source_verified"):
        _expect(overlay[key], True, f"{context}.opc_source_overlay.{key}")
    observed_output = [_digest(item, f"{context}.opc_source_overlay.observed_output_sha256[{i}]") for i, item in enumerate(_array(overlay["observed_output_sha256"], f"{context}.opc_source_overlay.observed_output_sha256"))]
    if len(observed_output) != SAMPLE_COUNT:
        raise ValidationError(f"{context}.opc_source_overlay.observed_output_sha256 must contain {SAMPLE_COUNT} values")
    overlay_identity = {key: overlay[key] for key in OVERLAY_KEYS if key not in {"preparation_ns", "open_ns", "planning_ns", "publication_ns", "observed_output_sha256"}}
    overlay_identity["numeric_identity_vectors"] = {key: value for key, value in overlay_vectors.items() if key not in {"preparation_ns", "open_ns", "planning_ns", "publication_ns"}}
    overlay_identity["observed_output_sha256"] = observed_output
    return vectors, {"generic": vectors, "overlay": overlay_identity, "phase_vectors": {key: overlay_vectors[key] for key in ("preparation_ns", "open_ns", "planning_ns", "publication_ns")}, "observed_output_sha256": observed_output}


def _validate_operation(value: Any, elapsed_samples: list[int], sample_order: list[int], sink: Mapping[str, Any], context: str) -> dict[str, list[int]]:
    operation = _exact_keys(value, OPERATION_KEYS, context)
    _expect(operation["sample_count"], SAMPLE_COUNT, f"{context}.sample_count")
    _expect(operation["sample_indices"], sample_order, f"{context}.sample_indices")
    _expect(operation["alignment"], ALIGNMENT, f"{context}.alignment")
    _expect(operation["latency_claim"], LATENCY_CLAIM, f"{context}.latency_claim")

    source = _exact_keys(operation["source"], OPERATION_SOURCE_KEYS, f"{context}.source")
    _expect(source["status"], "not_applicable", f"{context}.source.status")
    _expect(source["counter_scope"], "not_applicable_in_process_sink", f"{context}.source.counter_scope")
    for key in OPERATION_SOURCE_KEYS - {"status", "counter_scope"}:
        _validate_absent_vector(source[key], f"{context}.source.{key}", "not_applicable")

    process = _exact_keys(operation["process"], OPERATION_PROCESS_KEYS, f"{context}.process")
    _expect(process["status"], "unavailable", f"{context}.process.status")
    for key in OPERATION_PROCESS_KEYS - {"status"}:
        _validate_absent_vector(process[key], f"{context}.process.{key}", "unavailable")

    op_sink = _exact_keys(operation["sink"], OPERATION_SINK_KEYS, f"{context}.sink")
    _expect(op_sink["status"], "not_applicable", f"{context}.sink.status")
    _expect(op_sink["write_status"], "measured", f"{context}.sink.write_status")
    _validate_absent_vector(op_sink["output_bytes"], f"{context}.sink.output_bytes", "not_applicable")
    sink_scopes = {
        "accepted_bytes": "logical_sink_accepted_write_bytes",
        "write_calls": "logical_sink_accepted_write_calls",
        "largest_write": "logical_sink_largest_accepted_write",
    }
    for key, scope in sink_scopes.items():
        values = _validate_metric_vector(op_sink[key], f"{context}.sink.{key}", sample_count=SAMPLE_COUNT, status="measured", scope=scope)
        expected = [sink[key]] * SAMPLE_COUNT
        _expect(values, expected, f"{context}.sink.{key}/top-level sink")
    bucket_metrics = _exact_keys(op_sink["write_size_buckets"], {"status", "bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536"}, f"{context}.sink.write_size_buckets")
    _expect(bucket_metrics["status"], "measured", f"{context}.sink.write_size_buckets.status")
    for key in SINK_BUCKET_KEYS:
        values = _validate_metric_vector(bucket_metrics[key], f"{context}.sink.write_size_buckets.{key}", sample_count=SAMPLE_COUNT, status="measured", scope="logical_sink_accepted_write_size_bucket_counts")
        _expect(values, [sink["write_size_buckets"][key]] * SAMPLE_COUNT, f"{context}.sink.write_size_buckets.{key}/top-level sink")

    publication = _exact_keys(operation["publication"], OPERATION_PUBLICATION_KEYS, f"{context}.publication")
    _expect(publication["status"], "not_applicable", f"{context}.publication.status")
    for key in OPERATION_PUBLICATION_KEYS - {"status"}:
        _validate_absent_vector(publication[key], f"{context}.publication.{key}", "not_applicable")
    materialization = _exact_keys(operation["materialization"], OPERATION_MATERIALIZATION_KEYS, f"{context}.materialization")
    _expect(materialization["status"], "not_applicable", f"{context}.materialization.status")
    _validate_absent_vector(materialization["opc_parts"], f"{context}.materialization.opc_parts", "not_applicable")
    cfb = _exact_keys(operation["cfb_phases"], OPERATION_CFB_KEYS, f"{context}.cfb_phases")
    _expect(cfb["status"], "not_applicable", f"{context}.cfb_phases.status")
    for phase in ("open", "plan", "atomic_publication"):
        phase_value = _exact_keys(cfb[phase], OPERATION_CFB_PHASE_KEYS, f"{context}.cfb_phases.{phase}")
        for key in OPERATION_CFB_PHASE_KEYS:
            _validate_absent_vector(phase_value[key], f"{context}.cfb_phases.{phase}.{key}", "not_applicable")

    allocation = _exact_keys(operation["allocation"], ALLOCATION_KEYS, f"{context}.allocation")
    _expect(allocation["status"], "measured", f"{context}.allocation.status")
    _expect(allocation["scope"], ALLOCATION_SCOPE, f"{context}.allocation.scope")
    vectors: dict[str, list[int]] = {}
    for metric in ALLOCATION_METRICS:
        vectors[metric] = _validate_metric_vector(allocation[metric], f"{context}.allocation.{metric}", sample_count=SAMPLE_COUNT, status="measured", scope=ALLOCATION_SCOPE)
        if metric in COUNTER_METRICS and len(set(vectors[metric])) != 1:
            raise ValidationError(
                f"{context}.allocation.{metric} must be constant across all {SAMPLE_COUNT} samples"
            )
    return vectors


def _validate_parallel(value: Any, rows: list[tuple[str, str]], context: str) -> None:
    parallel = _exact_keys(value, PARALLEL_KEYS, context)
    _expect(parallel["schema_version"], 1, f"{context}.schema_version")
    _expect(parallel["scope"], "explicit_local_execution_only", f"{context}.scope")
    _expect(parallel["claim"], "descriptive", f"{context}.claim")
    budget = _exact_keys(parallel["configured_worker_budget"], {"status", "value", "scope"}, f"{context}.configured_worker_budget")
    _expect(budget["status"], "measured", f"{context}.configured_worker_budget.status")
    _expect(budget["value"], WORKERS, f"{context}.configured_worker_budget.value")
    _expect(budget["scope"], "configuration.execution_workers", f"{context}.configured_worker_budget.scope")
    thread = _exact_keys(parallel["observed_process_thread_count"], {"status", "scope", "reason"}, f"{context}.observed_process_thread_count")
    _expect(thread["status"], "unavailable", f"{context}.observed_process_thread_count.status")
    _expect(thread["scope"], "process_thread_count", f"{context}.observed_process_thread_count.scope")
    _string(thread["reason"], f"{context}.observed_process_thread_count.reason")
    cases = _array(parallel["cases"], f"{context}.cases")
    if len(cases) != len(rows):
        raise ValidationError(f"{context}.cases must contain {len(rows)} entries")
    case_keys = {"case", "corpus_sha256", "configured_worker_count", "observed_local_worker_count", "deterministic_task_count", "deterministic_chunk_count", "lock_wait_ns"}
    for index, (item, (shape, _count)) in enumerate(zip(cases, rows, strict=True)):
        entry = _exact_keys(item, case_keys, f"{context}.cases[{index}]")
        _expect(entry["case"], CASE, f"{context}.cases[{index}].case")
        # The corpus digest is checked against the report result by the caller;
        # requiring a digest here prevents a missing parallel identity.
        _digest(entry["corpus_sha256"], f"{context}.cases[{index}].corpus_sha256")
        for key in ("configured_worker_count", "observed_local_worker_count", "deterministic_task_count"):
            metric = _exact_keys(entry[key], {"status", "scope", "reason"}, f"{context}.cases[{index}].{key}")
            _expect(metric["status"], "not_applicable", f"{context}.cases[{index}].{key}.status")
            _string(metric["scope"], f"{context}.cases[{index}].{key}.scope")
            _string(metric["reason"], f"{context}.cases[{index}].{key}.reason")
        chunk = _exact_keys(entry["deterministic_chunk_count"], {"status", "scope", "reason"}, f"{context}.cases[{index}].deterministic_chunk_count")
        _expect(chunk["status"], "not_applicable", f"{context}.cases[{index}].deterministic_chunk_count.status")
        _string(chunk["scope"], f"{context}.cases[{index}].deterministic_chunk_count.scope")
        _string(chunk["reason"], f"{context}.cases[{index}].deterministic_chunk_count.reason")
        lock = _exact_keys(entry["lock_wait_ns"], {"status", "scope", "reason"}, f"{context}.cases[{index}].lock_wait_ns")
        _expect(lock["status"], "unavailable", f"{context}.cases[{index}].lock_wait_ns.status")
        _string(lock["scope"], f"{context}.cases[{index}].lock_wait_ns.scope")
        _string(lock["reason"], f"{context}.cases[{index}].lock_wait_ns.reason")


def _validate_corpus(value: Any, contract_corpus: Mapping[str, Any], shape: str, count: int, context: str) -> None:
    corpus = _exact_keys(value, {"name", "generator", "package_format", "shape", "payload_kind", "compression", "entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "archive_sha256", "target_entry", "target_payload_bytes", "target_payload_sha256", "xlsx"}, context)
    expected = dict(contract_corpus)
    expected.pop("base_name")
    expected["name"] = f"{contract_corpus['base_name']}-count-{count}"
    expected["xlsx"] = None
    _expect(corpus, expected, context)
    _expect(corpus["shape"], shape, f"{context}.shape")


def _validate_result(value: Any, role: str, contract: Contract, index: int) -> RowEvidence:
    context = f"{role}.results[{index}]"
    result = _exact_keys(value, RESULT_KEYS, context)
    _expect(result["case"], CASE, f"{context}.case")
    corpus_obj = _object(result["corpus"], f"{context}.corpus")
    shape = _string(corpus_obj.get("shape"), f"{context}.corpus.shape")
    if shape not in SHAPES:
        raise ValidationError(f"{context}.corpus.shape is not one of the fixed shapes")
    name = _string(corpus_obj.get("name"), f"{context}.corpus.name")
    matching_counts = [count for count in COUNTS if name == f"{contract.corpora[shape]['base_name']}-count-{count}"]
    if len(matching_counts) != 1:
        raise ValidationError(f"{context}.corpus.name does not identify exactly one fixed overlay count")
    count = matching_counts[0]
    _validate_corpus(corpus_obj, contract.corpora[shape], shape, count, f"{context}.corpus")
    elapsed_samples, sample_order = _validate_elapsed(result["elapsed_ns"], f"{context}.elapsed_ns")
    sink = _validate_sink(
        result["sink"],
        f"{context}.sink",
        max_bytes=contract.corpora[shape]["archive_bytes"] * 2 + 65_536,
    )
    source_vectors, source_identity = _validate_source(result["source"], contract.corpora[shape], count, f"{context}.source")
    output_sha256 = _digest(result["output_sha256"], f"{context}.output_sha256")
    _expect(output_sha256, contract.corpora[shape]["archive_sha256"], f"{context}.output_sha256")
    overlay = source_identity["overlay"]
    _expect(source_identity["observed_output_sha256"], [output_sha256] * SAMPLE_COUNT, f"{context}.source.opc_source_overlay.observed_output_sha256")
    phase_vectors = source_identity["phase_vectors"]
    for sample_index in range(SAMPLE_COUNT):
        phase_sum = sum(phase_vectors[key][sample_index] for key in ("preparation_ns", "open_ns", "planning_ns", "publication_ns"))
        _expect(phase_sum, elapsed_samples[sample_index], f"{context} phase sum at sorted sample {sample_index}")
    allocation = _validate_operation(result["operation_metrics"], elapsed_samples, sample_order, sink, f"{context}.operation_metrics")
    return RowEvidence(shape, count, elapsed_samples, sample_order, allocation, source_identity, sink, output_sha256)


def _validate_report(path: Path, role: str, contract: Contract) -> ValidatedReport:
    report, raw_sha256 = _read_json(path, role)
    _exact_keys(report, REPORT_KEYS, role)
    _expect(report["schema_version"], SCHEMA_VERSION, f"{role}.schema_version")
    _validate_tool(report["tool"], f"{role}.tool")
    _validate_binary(report["binary_identity"], role, contract, f"{role}.binary_identity")
    environment = _validate_environment(report["environment"], role, contract, f"{role}.environment")
    _validate_configuration(report["configuration"], f"{role}.configuration")
    results = _array(report["results"], f"{role}.results")
    expected_rows = [(shape, count) for shape in SHAPES for count in COUNTS]
    if len(results) != len(expected_rows):
        raise ValidationError(f"{role}.results must contain exactly {len(expected_rows)} matrix rows")
    rows: dict[tuple[str, int], RowEvidence] = {}
    for index, expected in enumerate(expected_rows):
        row = _validate_result(results[index], role, contract, index)
        if (row.shape, row.count) != expected:
            raise ValidationError(f"{role}.results[{index}] is out of fixed shape/count order")
        rows[expected] = row
    _validate_parallel(
        report["parallel_metrics"],
        [(shape, count) for shape, count in expected_rows],
        f"{role}.parallel_metrics",
    )
    # Parallel metrics repeat one corpus digest per result. Compare it to its
    # result after parsing, rather than accepting a detached list.
    parallel_cases = report["parallel_metrics"]["cases"]
    for index, (shape, count) in enumerate(expected_rows):
        _expect(
            parallel_cases[index]["corpus_sha256"],
            contract.corpora[shape]["archive_sha256"],
            f"{role}.parallel_metrics.cases[{index}].corpus_sha256",
        )
    return ValidatedReport(role, raw_sha256, report, rows)


def _environment_identity(report: Mapping[str, Any]) -> dict[str, Any]:
    environment = dict(report["environment"])
    environment.pop("git_revision", None)
    return environment


def _validate_cross_report(reports: Mapping[str, ValidatedReport], contract: Contract) -> None:
    first = reports["A1"]
    for role in ABBA_ROLES:
        if role == "A1":
            continue
        _expect(_environment_identity(reports[role].report), _environment_identity(first.report), f"{role}/A1 environment identity")
        _expect(reports[role].report["configuration"], first.report["configuration"], f"{role}/A1 configuration identity")
        _expect(reports[role].report["parallel_metrics"]["configured_worker_budget"], first.report["parallel_metrics"]["configured_worker_budget"], f"{role}/A1 worker identity")
    expected_keys = {(shape, count) for shape in SHAPES for count in COUNTS}
    for key in expected_keys:
        row_a1 = reports["A1"].rows[key]
        row_a2 = reports["A2"].rows[key]
        row_b1 = reports["B1"].rows[key]
        row_b2 = reports["B2"].rows[key]
        # Source, sink, semantic/oracle, cache, and output evidence are
        # deterministic identity, not latency.  Phase timing vectors are
        # intentionally excluded from this equality because elapsed is not a
        # claimable allocator metric.
        for role in ABBA_ROLES:
            if role == "A1":
                continue
            row = reports[role].rows[key]
            _expect(row_a1.source_identity["generic"], row.source_identity["generic"], f"{key} A1/{role} source identity")
            _expect(row_a1.source_identity["overlay"], row.source_identity["overlay"], f"{key} A1/{role} overlay identity")
            _expect(row_a1.sink_identity, row.sink_identity, f"{key} A1/{role} sink identity")
            _expect(row_a1.output_sha256, row.output_sha256, f"{key} A1/{role} output oracle")
        for metric in COUNTER_METRICS:
            _expect(row_a1.allocation[metric], row_a2.allocation[metric], f"{key} control {metric} A1/A2 equality")
            _expect(row_b1.allocation[metric], row_b2.allocation[metric], f"{key} candidate {metric} B1/B2 equality")
            expected_delta = EXPECTED_DELTA_BY_COUNT[key[1]][metric]
            for left, right, pair in ((row_a1, row_b1, "A1_to_B1"), (row_a2, row_b2, "A2_to_B2")):
                deltas = [candidate - control for control, candidate in zip(left.allocation[metric], right.allocation[metric], strict=True)]
                if any(delta != expected_delta for delta in deltas):
                    raise ValidationError(f"{key} {pair} {metric} does not have exact delta {expected_delta}")


def _vector_digest(values: Iterable[int]) -> str:
    raw = json.dumps(list(values), separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _validator_source_sha256() -> str:
    source_path = Path(__file__)
    try:
        source = source_path.read_bytes()
    except OSError as error:
        raise ValidationError(f"cannot read validator source {source_path}: {error}") from error
    return hashlib.sha256(source).hexdigest()


def _projection(reports: Mapping[str, ValidatedReport], contract: Contract) -> dict[str, Any]:
    rows = []
    for shape in SHAPES:
        for count in COUNTS:
            key = (shape, count)
            a1, b1 = reports["A1"].rows[key], reports["B1"].rows[key]
            allocation = {}
            for metric in COUNTER_METRICS:
                control = a1.allocation[metric]
                candidate = b1.allocation[metric]
                deltas = [right - left for left, right in zip(control, candidate, strict=True)]
                allocation[metric] = {
                    "control_unique_values": sorted(set(control)),
                    "candidate_unique_values": sorted(set(candidate)),
                    "candidate_minus_control_unique": sorted(set(deltas)),
                    "delta": deltas[0],
                    "vector_sha256": _vector_digest(deltas),
                }
            rows.append(
                {
                    "shape": shape,
                    "overlay_count": count,
                    "source_sha256": a1.source_identity["overlay"]["source_sha256"],
                    "output_sha256": a1.output_sha256,
                    "allocation": allocation,
                    "claimability": {
                        "operation_counters": True,
                        "allocator_elapsed_ns": False,
                        "live_bytes_before": False,
                        "live_bytes_after": False,
                        "peak_live_bytes_before": False,
                        "peak_live_bytes_after": False,
                    },
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
            "execution_workers": WORKERS,
            "fixed_shapes": list(SHAPES),
            "overlay_counts": list(COUNTS),
            "operation_vector_alignment": "operation_metrics.sample_indices equals elapsed_ns.sample_order",
            "sample_indices_equal_elapsed_sample_order": True,
        },
        "provenance": {
            "contract_sha256": contract.raw_sha256,
            "validator_source_sha256": _validator_source_sha256(),
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
            "binary_identity": {
                "mode": "reported+contract-bound",
                "file_rehashed": False,
                "reason": "allocator binary artifacts may be absent after collection",
            },
            "raw_report_sha256": {role: reports[role].raw_sha256 for role in ABBA_ROLES},
        },
        "validation": {
            "report_count": 4,
            "four_distinct_raw_report_byte_streams": True,
            "matrix_rows": len(SHAPES) * len(COUNTS),
            "samples_per_row": SAMPLE_COUNT,
            "allocator_vector_count": len(ALLOCATION_METRICS) * len(SHAPES) * len(COUNTS) * 4,
            "a1_equals_a2_six_counters": True,
            "b1_equals_b2_six_counters": True,
            "exact_delta_policy_verified": True,
            "semantic_source_sink_oracle_phase_identity_verified": True,
            "operation_vector_aligned_by_sample_indices_and_elapsed_order": True,
            "binary_identity_reported_and_contract_bound": True,
        },
        "claimability": {
            "allocation_calls": {"claimable": True, "scope": ALLOCATION_SCOPE},
            "deallocation_calls": {"claimable": True, "scope": ALLOCATION_SCOPE},
            "reallocation_calls": {"claimable": True, "scope": ALLOCATION_SCOPE},
            "failed_allocation_calls": {"claimable": True, "scope": ALLOCATION_SCOPE},
            "allocated_bytes": {"claimable": True, "scope": ALLOCATION_SCOPE},
            "deallocated_bytes": {"claimable": True, "scope": ALLOCATION_SCOPE},
            "allocator_elapsed_ns": {"claimable": False, "reason": "allocator instrumentation changes elapsed timing"},
            "live_bytes_before": {"claimable": False, "reason": "absolute process snapshot, not operation memory"},
            "live_bytes_after": {"claimable": False, "reason": "absolute process snapshot, not operation memory"},
            "peak_live_bytes_before": {"claimable": False, "reason": "absolute process high-water snapshot, not operation peak"},
            "peak_live_bytes_after": {"claimable": False, "reason": "absolute process high-water snapshot, not operation peak"},
            "independent_process_proof": {
                "claimable": False,
                "reason": "in-process reports contain no per-sample envelope or PID evidence",
            },
        },
        "rows": rows,
    }


def validate_paths(a1: Path, b1: Path, b2: Path, a2: Path, contract_path: Path) -> dict[str, Any]:
    contract = _validate_contract(contract_path)
    paths = {"A1": a1, "B1": b1, "B2": b2, "A2": a2}
    if len({path.resolve() for path in paths.values()}) != 4:
        raise ValidationError("the four ABBA roles must use four distinct report paths")
    reports = {role: _validate_report(path, role, contract) for role, path in paths.items()}
    if len({report.raw_sha256 for report in reports.values()}) != 4:
        raise ValidationError("the four ABBA roles must use four distinct raw report byte streams")
    _validate_cross_report(reports, contract)
    return _projection(reports, contract)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--contract", type=Path, required=True, help="checked JSON contract with final revision/binary/corpus identities")
    parser.add_argument("--a1", type=Path, required=True, help="control A1 allocator report")
    parser.add_argument("--b1", type=Path, required=True, help="candidate B1 allocator report")
    parser.add_argument("--b2", type=Path, required=True, help="candidate B2 allocator report")
    parser.add_argument("--a2", type=Path, required=True, help="control A2 allocator report")
    parser.add_argument("--output", type=Path, help="optional output path for compact JSON validation/projection (must differ from every input)")
    return parser


def _reject_output_collision(output: Path | None, input_paths: Iterable[Path]) -> None:
    if output is None:
        return
    resolved_output = output.resolve()
    for input_path in input_paths:
        if resolved_output == input_path.resolve():
            raise ValidationError("--output must differ from contract and report inputs")
        try:
            same_file = output.exists() and input_path.exists() and output.samefile(input_path)
        except OSError as error:
            raise ValidationError(f"cannot compare --output with input {input_path}: {error}") from error
        if same_file:
            raise ValidationError("--output must differ from contract and report inputs")


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        _reject_output_collision(
            args.output,
            (args.contract, args.a1, args.b1, args.b2, args.a2),
        )
        result = validate_paths(args.a1, args.b1, args.b2, args.a2, args.contract)
        encoded = json.dumps(result, sort_keys=True, separators=(",", ":"))
        if args.output is None:
            print(encoded)
        else:
            args.output.write_text(encoded + "\n", encoding="utf-8")
            print(encoded)
    except (OSError, ValidationError) as error:
        print(f"OPC overlay allocator ABBA validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
