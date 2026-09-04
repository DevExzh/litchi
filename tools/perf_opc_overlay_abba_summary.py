#!/usr/bin/env python3
"""Summarize the normal ABBA publication phase for the OPC noop selector.

This is a deliberately narrow evidence boundary.  It accepts exactly four
clean, current-schema normal reports in ``A1, B1, B2, A2`` order, each carrying
the fixed three-shape by three-count noop matrix.  The harness's top-level
elapsed value is validated as a phase-sum oracle, but it is never summarized:
all latency statistics in the result are computed from the nested
``source.opc_source_overlay.publication_ns`` vectors only.  Allocator reports
and allocator timing are outside this tool's claim boundary.  The selected
normal reports contain 500 retained in-process samples after 20 warmups; the
global filesystem-child configuration fields do not establish fresh-child
semantics for this selector.

The report envelope and the format-specific source/sink gates are delegated to
the existing standard-library-only ``perf_abba_summary`` validator.  This
module adds the stricter selector, matrix, identity, and publication-phase
policy without modifying that validator or its existing evidence hashes.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import re
import sys
from pathlib import Path
from typing import Any, Mapping, Sequence

if __package__:
    from . import perf_abba_summary
else:  # pragma: no cover - exercised by the direct CLI entry point
    import perf_abba_summary


SCHEMA_VERSION = 1
TOOL_NAME = "litchi-opc-overlay-abba-summary"
TOOL_VERSION = "0.1.0"
REPORT_SCHEMA_VERSION = 1
REPORT_TOOL_NAME = "litchi-perf-baseline"
REPORT_TOOL_VERSION = "0.1.0"
REPORT_PROFILE = "current-v1"
CASE = "opc_source_overlay_multi_part_noop"
SHAPES = ("overlay-small", "overlay-large", "overlay-media-incompressible")
COUNTS = (2, 8, 32)
ROLES = ("a1", "b1", "b2", "a2")
STATISTICS = ("p50", "mean", "p95", "p99")
DRIFT_CEILINGS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}
U64_MAX = (1 << 64) - 1
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
REVISION_RE = re.compile(r"[0-9a-f]{40}\Z")
SAMPLE_COUNT = 500
WARMUP_ITERATIONS = 20
FILESYSTEM_CACHE_STATES = ["warm", "cold-requested"]
EXECUTION_WORKERS = [1]
CPU_AFFINITY = "2"
RUSTC_VERSION = "rustc 1.98.1 (48a229cea 2026-09-01)"
ALLOCATOR = "Rust system allocator"
TARGET_OS = "linux"
TARGET_ARCH = "x86_64"
TIMING_SCOPE = (
    "elapsed_ns is explicitly the sum of preparation_ns, open_ns, planning_ns, "
    "and publication_ns; operation_metrics.allocation covers only the "
    "write_part_overlays_to_stream publication call; structural setup belongs "
    "to those named phases, while interstitial checks and eager artifact, "
    "reopen/digest/preservation, source/cache evidence, configured cache/sink "
    "ceilings, and semantic oracles are excluded from the aggregate"
)
OPERATION_METRICS_KEYS = frozenset(
    {
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
)
METRIC_VECTOR_STATUSES = frozenset(
    {"measured", "not_applicable", "unavailable", "overflow"}
)
ALLOCATION_VECTOR_FIELDS = (
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
OPERATION_ALIGNMENT = "elapsed_ns.samples_by_elapsed_then_sample_index"
OPERATION_LATENCY_CLAIM = "comparable_timed_operation"
ALLOCATION_SCOPE = "operation_global_system_allocator"
OUTPUT_BYTES_SCOPE = "post_operation_output_length_not_sink_write_volume"
SINK_ACCEPTED_BYTES_SCOPE = "logical_sink_accepted_write_bytes"
SINK_WRITE_CALLS_SCOPE = "logical_sink_accepted_write_calls"
SINK_LARGEST_WRITE_SCOPE = "logical_sink_largest_accepted_write"
SINK_BUCKET_SCOPE = "logical_sink_accepted_write_size_bucket_counts"
SOURCE_SCOPE = "operation_logical_read_at"
SOURCE_PATTERN_SCOPE = "operation_logical_read_at_range_order_not_physical_io"
SOURCE_COMPRESSED_SCOPE = "unavailable_read_at_has_no_compressed_member_boundary"
SOURCE_DECOMPRESSED_SCOPE = "unavailable_read_at_has_no_decompressed_byte_boundary"
SOURCE_RECOMPRESSED_SCOPE = "unavailable_atomic_save_has_no_recompressed_byte_boundary"
PROCESS_SCOPE = "procfs_operation_delta"
PROC_IO_SCOPE = "child_process_interval_delta_including_procfs_probe_overhead"
RSS_SCOPE = "procfs_operation_delta_not_peak"
HWM_SCOPE = "process_lifetime_high_water_after_not_operation_peak"
IN_PROCESS_PROCESS_SCOPE = "procfs_in_process_operation_delta_including_procfs_probe_overhead"
IN_PROCESS_RSS_SCOPE = "procfs_in_process_rss_delta_including_procfs_probe_overhead"
CFB_PHASE_ELAPSED_SCOPE = "timed_cfb_phase_elapsed_ns"
CFB_PHASE_SOURCE_SCOPE = "timed_cfb_phase_logical_read_at"

EXPECTED_CORPORA: dict[str, dict[str, Any]] = {
    "overlay-small": {
        "generator": "litchi-opc-source-overlay-multi-part-v1",
        "package_format": "OPC/ZIP",
        "shape": "overlay-small",
        "payload_kind": "compressible",
        "compression": "deflate",
        "entry_count": 32,
        "archive_member_count": 34,
        "entry_bytes": 1024,
        "uncompressed_payload_bytes": 32 * 1024,
        "archive_bytes": 7451,
        "target_entry": "benchmark/parts/00016.bin",
        "target_payload_bytes": 1024,
        "archive_sha256": "4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509",
        "target_payload_sha256": "5b7b9793a43d08ca2c0670289d932541377407ed352ab9f6c145f63d19de9f98",
        "xlsx": None,
        "name_prefix": "overlay-small-compressible-count-",
    },
    "overlay-large": {
        "generator": "litchi-opc-source-overlay-multi-part-v1",
        "package_format": "OPC/ZIP",
        "shape": "overlay-large",
        "payload_kind": "incompressible",
        "compression": "deflate",
        "entry_count": 32,
        "archive_member_count": 34,
        "entry_bytes": 64 * 1024,
        "uncompressed_payload_bytes": 2 * 1024 * 1024,
        "archive_bytes": 2_103_195,
        "target_entry": "benchmark/parts/00016.bin",
        "target_payload_bytes": 64 * 1024,
        "archive_sha256": "8356d7467215b04a3d1c3703f50fbd6322f2002ca7c3ead1f24414c5e550ef73",
        "target_payload_sha256": "e17b543eec6b4d3534978d7d59e7240dcbf0f2a2050fd80f32ea3daec266aa73",
        "xlsx": None,
        "name_prefix": "overlay-large-incompressible-count-",
    },
    "overlay-media-incompressible": {
        "generator": "litchi-opc-source-overlay-multi-part-v1",
        "package_format": "OPC/ZIP",
        "shape": "overlay-media-incompressible",
        "payload_kind": "incompressible",
        "compression": "deflate",
        "entry_count": 32,
        "archive_member_count": 34,
        "entry_bytes": 256 * 1024,
        "uncompressed_payload_bytes": 8 * 1024 * 1024,
        "archive_bytes": 8_396_580,
        "target_entry": "benchmark/parts/00016.bin",
        "target_payload_bytes": 256 * 1024,
        "archive_sha256": "bf8c309af5306c6682b9df65b97246f81b022fe5e3b5e02cc2c4dcf3e1e87883",
        "target_payload_sha256": "3ad07c7e34d3dd6d9ff75b696ccbdd702777b6e4dea04b19bbe3d0aa6d21cdeb",
        "xlsx": None,
        "name_prefix": "overlay-media-incompressible-incompressible-count-",
    },
}

ROOT_KEYS = frozenset(
    {
        "schema_version",
        "tool",
        "binary_identity",
        "environment",
        "configuration",
        "parallel_metrics",
        "results",
    }
)
TOOL_KEYS = frozenset(
    {"name", "version", "binary", "profile", "target_os", "target_arch", "instrumentation"}
)
BINARY_KEYS = frozenset(
    {"path", "binary_sha256", "binary_bytes", "mode_bits", "executable", "profile"}
)
ENVIRONMENT_KEYS = frozenset(
    {
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
)
CONFIGURATION_KEYS = frozenset(
    {
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
)
ROW_KEYS = frozenset(
    {"case", "corpus", "elapsed_ns", "sink", "source", "output_sha256", "operation_metrics"}
)
REQUIRED_ROW_KEYS = frozenset(
    {
        "case",
        "corpus",
        "elapsed_ns",
        "sink",
        "source",
        "output_sha256",
        "operation_metrics",
    }
)
ELAPSED_KEYS = frozenset(
    {
        "unit",
        "sample_order",
        "min",
        "p50",
        "p95",
        "p99",
        "max",
        "mean",
        "standard_deviation",
        "confidence_interval_95",
        "samples",
    }
)
CONFIDENCE_KEYS = frozenset({"method", "lower", "upper"})
SOURCE_KEYS = frozenset(
    {
        "read_calls",
        "read_bytes",
        "ordinary_payload_read_calls",
        "ordinary_payload_read_bytes",
        "max_in_flight_reads",
        "opc_source_overlay",
    }
)
SINK_KEYS = frozenset(
    {"accepted_bytes", "write_calls", "largest_write", "write_size_buckets"}
)
SINK_BUCKET_KEYS = frozenset(
    {
        "bytes_0",
        "bytes_1_to_512",
        "bytes_513_to_4096",
        "bytes_4097_to_16384",
        "bytes_16385_to_65536",
        "bytes_over_65536",
    }
)
OVERLAY_KEYS = frozenset(
    {
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
)
PHASE_FIELDS = ("preparation_ns", "open_ns", "planning_ns", "publication_ns")
COUNTER_FIELDS = (
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
)
SOURCE_VECTOR_FIELDS = (
    "read_calls",
    "read_bytes",
    "ordinary_payload_read_calls",
    "ordinary_payload_read_bytes",
    "max_in_flight_reads",
)
OVERLAY_DYNAMIC_FIELDS = frozenset((*PHASE_FIELDS, *COUNTER_FIELDS, "observed_output_sha256"))

OPERATION_SOURCE_KEYS = frozenset(
    {
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
)
OPERATION_PROCESS_KEYS = frozenset(
    {
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
)
OPERATION_SINK_KEYS = frozenset(
    {
        "status",
        "output_bytes",
        "write_status",
        "accepted_bytes",
        "write_calls",
        "largest_write",
        "write_size_buckets",
    }
)
OPERATION_BUCKET_KEYS = frozenset(
    {
        "status",
        "bytes_0",
        "bytes_1_to_512",
        "bytes_513_to_4096",
        "bytes_4097_to_16384",
        "bytes_16385_to_65536",
        "bytes_over_65536",
    }
)
OPERATION_PUBLICATION_KEYS = frozenset({"status", "changed_spans", "published_bytes"})
OPERATION_MATERIALIZATION_KEYS = frozenset({"status", "opc_parts"})
OPERATION_CFB_KEYS = frozenset({"status", "open", "plan", "atomic_publication"})
OPERATION_CFB_PHASE_KEYS = frozenset(
    {
        "elapsed_ns",
        "logical_read_calls",
        "logical_read_requested_bytes",
        "logical_read_returned_bytes",
    }
)
ALLOCATION_KEYS = frozenset({"status", "scope", *ALLOCATION_VECTOR_FIELDS})


class OverlayAbbaInputError(ValueError):
    """Raised when the specialized overlay ABBA boundary is not satisfied."""


def _error(message: str) -> OverlayAbbaInputError:
    return OverlayAbbaInputError(message)


def _require_object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise _error(f"{location} must be an object")
    return value


def _exact_keys(value: Any, location: str, expected: frozenset[str]) -> dict[str, Any]:
    obj = _require_object(value, location)
    actual = set(obj)
    missing = sorted(expected - actual)
    unknown = sorted(actual - expected, key=repr)
    if missing or unknown:
        raise _error(f"{location} keys mismatch (missing={missing!r}, unknown={unknown!r})")
    return obj


def _row_keys(value: Any, location: str) -> dict[str, Any]:
    obj = _require_object(value, location)
    actual = set(obj)
    missing = sorted(REQUIRED_ROW_KEYS - actual)
    unknown = sorted(actual - ROW_KEYS)
    if missing or unknown:
        raise _error(f"{location} keys mismatch (missing={missing!r}, unknown={unknown!r})")
    return obj


def _canonical(value: Any, location: str) -> str:
    try:
        return json.dumps(
            value,
            sort_keys=True,
            separators=(",", ":"),
            ensure_ascii=True,
            allow_nan=False,
        )
    except (TypeError, ValueError, OverflowError) as error:
        raise _error(f"{location} is not canonical JSON: {error}") from error


def _sha256(value: Any, location: str) -> str:
    digest = _canonical(value, location).encode("utf-8")
    return hashlib.sha256(digest).hexdigest()


def _u64(value: Any, location: str, *, positive: bool = False) -> int:
    if (
        isinstance(value, bool)
        or not isinstance(value, int)
        or value < (1 if positive else 0)
        or value > U64_MAX
    ):
        requirement = "positive " if positive else ""
        raise _error(f"{location} must be a {requirement}unsigned 64-bit integer")
    return value


def _bool(value: Any, location: str) -> bool:
    if not isinstance(value, bool):
        raise _error(f"{location} must be a boolean")
    return value


def _digest(value: Any, location: str) -> str:
    if not isinstance(value, str) or SHA256_RE.fullmatch(value) is None:
        raise _error(f"{location} must be a lowercase SHA-256 hex string")
    return value


def _validate_root_shape(report: Mapping[str, Any], label: str) -> None:
    _exact_keys(report, label, ROOT_KEYS)
    if report["schema_version"] != REPORT_SCHEMA_VERSION:
        raise _error(f"{label}.schema_version must be {REPORT_SCHEMA_VERSION}")

    tool = _exact_keys(report["tool"], f"{label}.tool", TOOL_KEYS)
    if (
        tool["name"],
        tool["version"],
        tool["binary"],
        tool["profile"],
        tool["instrumentation"],
    ) != (
        REPORT_TOOL_NAME,
        REPORT_TOOL_VERSION,
        REPORT_TOOL_NAME,
        "release",
        "none",
    ):
        raise _error(f"{label}.tool is not the normal litchi-perf-baseline identity")
    if tool["target_os"] != TARGET_OS or tool["target_arch"] != TARGET_ARCH:
        raise _error(f"{label}.tool target platform is not the fixed normal target")

    binary = _exact_keys(report["binary_identity"], f"{label}.binary_identity", BINARY_KEYS)
    if not isinstance(binary["path"], str) or not binary["path"].startswith(("/", "\\\\")):
        raise _error(f"{label}.binary_identity.path must be absolute")
    if binary["path"].replace("\\", "/").rstrip("/").rsplit("/", 1)[-1] != REPORT_TOOL_NAME:
        raise _error(
            f"{label}.binary_identity.path must name {REPORT_TOOL_NAME!r}"
        )
    _digest(binary["binary_sha256"], f"{label}.binary_identity.binary_sha256")
    _u64(binary["binary_bytes"], f"{label}.binary_identity.binary_bytes", positive=True)
    if binary["mode_bits"] is not None:
        _u64(binary["mode_bits"], f"{label}.binary_identity.mode_bits")
    if binary["executable"] is not True:
        raise _error(f"{label}.binary_identity.executable must be true")
    if binary["profile"] != tool["profile"]:
        raise _error(f"{label}.binary_identity.profile disagrees with tool.profile")

    environment = _exact_keys(report["environment"], f"{label}.environment", ENVIRONMENT_KEYS)
    if environment["rustc_version"] != RUSTC_VERSION:
        raise _error(
            f"{label}.environment.rustc_version must be exactly {RUSTC_VERSION!r}"
        )
    if (
        not isinstance(environment["git_revision"], str)
        or REVISION_RE.fullmatch(environment["git_revision"]) is None
    ):
        raise _error(f"{label}.environment.git_revision must be a 40-character revision")
    if environment["git_worktree_dirty"] is not False:
        raise _error(f"{label}.environment.git_worktree_dirty must be false")
    _u64(environment["logical_cpus_available"], f"{label}.environment.logical_cpus_available", positive=True)
    if environment["allocator"] != ALLOCATOR:
        raise _error(f"{label}.environment.allocator must be {ALLOCATOR!r}")
    if environment["cpu_affinity"] != CPU_AFFINITY:
        raise _error(f"{label}.environment.cpu_affinity must be {CPU_AFFINITY!r}")

    configuration = _exact_keys(
        report["configuration"], f"{label}.configuration", CONFIGURATION_KEYS
    )
    samples = configuration["samples_per_case"]
    if samples != SAMPLE_COUNT:
        raise _error(
            f"{label}.configuration.samples_per_case must be exactly {SAMPLE_COUNT}"
        )
    warmups = configuration["warmup_iterations_per_case"]
    if warmups != WARMUP_ITERATIONS:
        raise _error(
            f"{label}.configuration.warmup_iterations_per_case must be exactly "
            f"{WARMUP_ITERATIONS}"
        )
    if configuration["filesystem_cache_states"] != FILESYSTEM_CACHE_STATES:
        raise _error(
            f"{label}.configuration.filesystem_cache_states must be exactly "
            f"{FILESYSTEM_CACHE_STATES!r}"
        )
    if configuration["filesystem_fresh_child_per_sample"] is not True:
        raise _error(
            f"{label}.configuration.filesystem_fresh_child_per_sample must be true"
        )
    if configuration["filesystem_process_isolated"] is not True:
        raise _error(
            f"{label}.configuration.filesystem_process_isolated must be true"
        )
    if configuration["cases"] != [CASE]:
        raise _error(f"{label}.configuration.cases must be exactly [{CASE!r}]")
    if configuration["filesystem_root_selected"] is not False:
        raise _error(f"{label}.configuration.filesystem_root_selected must be false")
    if configuration["execution_workers"] != EXECUTION_WORKERS:
        raise _error(
            f"{label}.configuration.execution_workers must be exactly "
            f"{EXECUTION_WORKERS!r}"
        )


def _validate_corpus(corpus: Any, location: str, shape: str, count: int) -> dict[str, Any]:
    expected = EXPECTED_CORPORA[shape]
    expected = {
        **expected,
        "name": f"{expected['name_prefix']}{count}",
    }
    expected.pop("name_prefix")
    actual = _exact_keys(corpus, location, frozenset(expected))
    for field, value in expected.items():
        if actual[field] != value:
            raise _error(f"{location}.{field} disagrees with the fixed corpus identity")
    return actual


def _validate_elapsed(value: Any, location: str, sample_count: int) -> tuple[list[int], list[int]]:
    elapsed = _exact_keys(value, location, ELAPSED_KEYS)
    if elapsed["unit"] != "ns":
        raise _error(f"{location}.unit must be 'ns'")
    samples = elapsed["samples"]
    if not isinstance(samples, list) or len(samples) != sample_count:
        raise _error(f"{location}.samples must contain exactly {sample_count} samples")
    samples = [_u64(item, f"{location}.samples[{index}]") for index, item in enumerate(samples)]
    if samples != sorted(samples):
        raise _error(f"{location}.samples must be sorted ascending")
    order = elapsed["sample_order"]
    if (
        not isinstance(order, list)
        or len(order) != sample_count
        or any(isinstance(item, bool) or not isinstance(item, int) or item < 0 or item >= sample_count for item in order)
        or sorted(order) != list(range(sample_count))
    ):
        raise _error(f"{location}.sample_order must be an exact sample permutation")
    confidence = _exact_keys(elapsed["confidence_interval_95"], f"{location}.confidence_interval_95", CONFIDENCE_KEYS)
    if confidence["method"] != "two-sided Student's t interval for the mean":
        raise _error(f"{location}.confidence_interval_95.method does not match the harness")
    for field in ("min", "p50", "p95", "p99", "max"):
        _u64(elapsed[field], f"{location}.{field}")
    for field in ("mean", "standard_deviation"):
        _finite(elapsed[field], f"{location}.{field}")
    _finite(confidence["lower"], f"{location}.confidence_interval_95.lower")
    _finite(confidence["upper"], f"{location}.confidence_interval_95.upper")
    # The generic report validator recomputes every top-level statistic.  Keep
    # that check here too so this narrow tool remains explicit about the
    # phase-sum binding it relies on.
    try:
        recomputed = perf_abba_summary.recompute_statistics(elapsed, location)
    except Exception as error:
        raise _error(str(error)) from error
    for field in ("min", "p50", "p95", "max"):
        if elapsed[field] != recomputed[field]:
            raise _error(f"{location}.{field} does not match samples")
    for field in ("mean", "standard_deviation"):
        if not math.isclose(elapsed[field], recomputed[field], rel_tol=1e-12, abs_tol=1e-12):
            raise _error(f"{location}.{field} does not match samples")
    for field in ("lower", "upper"):
        if not math.isclose(
            confidence[field], recomputed["confidence_interval_95"][field], rel_tol=1e-12, abs_tol=1e-12
        ):
            raise _error(f"{location}.confidence_interval_95.{field} does not match samples")
    return samples, order


def _finite(value: Any, location: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise _error(f"{location} must be a finite number")
    result = float(value)
    if not math.isfinite(result):
        raise _error(f"{location} must be a finite number")
    return result


def _validate_sink(
    value: Any,
    location: str,
    *,
    max_bytes: int,
    expected_accepted_bytes: int,
) -> dict[str, Any]:
    sink = _exact_keys(value, location, SINK_KEYS)
    accepted = _u64(sink["accepted_bytes"], f"{location}.accepted_bytes", positive=True)
    writes = _u64(sink["write_calls"], f"{location}.write_calls", positive=True)
    largest = _u64(sink["largest_write"], f"{location}.largest_write", positive=True)
    if accepted != expected_accepted_bytes:
        raise _error(
            f"{location}.accepted_bytes must equal the deterministic noop output "
            "bytes"
        )
    if accepted > max_bytes or largest > 65_536 or largest > accepted:
        raise _error(f"{location} exceeds the bounded sink policy")
    buckets = _exact_keys(sink["write_size_buckets"], f"{location}.write_size_buckets", SINK_BUCKET_KEYS)
    bucket_values = {
        field: _u64(buckets[field], f"{location}.write_size_buckets.{field}")
        for field in SINK_BUCKET_KEYS
    }
    if (
        sum(bucket_values.values()) != writes
        or bucket_values["bytes_0"] != 0
        or bucket_values["bytes_over_65536"] != 0
    ):
        raise _error(f"{location}.write_size_buckets disagrees with sink writes")
    if writes < (accepted + 65_535) // 65_536:
        raise _error(f"{location}.write_calls cannot carry accepted bytes")
    bucket_order = (
        ("bytes_1_to_512", 512),
        ("bytes_513_to_4096", 4_096),
        ("bytes_4097_to_16384", 16_384),
        ("bytes_16385_to_65536", 65_536),
    )
    largest_bucket = next(
        index for index, (_field, upper_bound) in enumerate(bucket_order) if largest <= upper_bound
    )
    if bucket_values[bucket_order[largest_bucket][0]] == 0:
        raise _error(f"{location}.write_size_buckets omits the largest write")
    if any(
        bucket_values[field] != 0
        for field, _upper_bound in bucket_order[largest_bucket + 1 :]
    ):
        raise _error(f"{location}.write_size_buckets exceeds largest_write")
    if max_bytes != accepted * 2 + 65_536:
        raise _error(f"{location} does not match 2*accepted_bytes+65536")
    return sink


def _vector(value: Any, location: str, sample_count: int, *, positive: bool = False) -> list[int]:
    if not isinstance(value, list) or len(value) != sample_count:
        raise _error(f"{location} must contain exactly {sample_count} samples")
    return [
        _u64(item, f"{location}[{index}]", positive=positive)
        for index, item in enumerate(value)
    ]


def _metric_vector(
    value: Any,
    location: str,
    sample_count: int,
    *,
    expected_status: str,
    expected_scope: str,
    expected_values: Sequence[int] | None = None,
) -> list[int] | None:
    """Validate one serialized operation-metrics MetricVector.

    The report serializer omits ``values`` for an unavailable or
    not-applicable metric.  In particular, ``values: null`` is not an
    acceptable representation of an unavailable allocator vector.
    """

    obj = _require_object(value, location)
    status = obj.get("status")
    scope = obj.get("scope")
    if not isinstance(status, str) or status not in METRIC_VECTOR_STATUSES:
        raise _error(f"{location}.status is not a recognized metric status")
    if status != expected_status:
        raise _error(
            f"{location}.status must be {expected_status!r}, got {status!r}"
        )
    if scope != expected_scope:
        raise _error(
            f"{location}.scope must be {expected_scope!r}, got {scope!r}"
        )
    if status == "measured":
        if set(obj) != {"status", "scope", "values"}:
            raise _error(f"{location} measured MetricVector keys are malformed")
        values = _vector(obj["values"], f"{location}.values", sample_count)
        if expected_values is not None and values != list(expected_values):
            raise _error(f"{location}.values disagrees with the sink envelope")
        return values
    if set(obj) != {"status", "scope"}:
        raise _error(
            f"{location} {status} MetricVector must omit its numeric values"
        )
    if expected_values is not None:
        raise _error(f"{location} cannot bind values for a non-measured metric")
    return None


def _validate_operation_metrics(
    value: Any,
    location: str,
    *,
    sample_count: int,
    elapsed_sample_order: Sequence[int],
    sink: Mapping[str, Any],
) -> dict[str, Any]:
    """Validate the normal in-process operation-metrics envelope.

    This is intentionally separate from the top-level source/cache evidence:
    the normal multi-part overlay runner emits an in-process sink envelope and
    an explicit unavailable allocator envelope, while the publication vectors
    remain the only claimable latency samples.
    """

    operation = _exact_keys(value, location, OPERATION_METRICS_KEYS)
    declared_count = operation["sample_count"]
    if (
        isinstance(declared_count, bool)
        or not isinstance(declared_count, int)
        or declared_count != sample_count
    ):
        raise _error(f"{location}.sample_count must equal {sample_count}")
    sample_indices = operation["sample_indices"]
    if sample_indices != list(elapsed_sample_order):
        raise _error(f"{location}.sample_indices must equal elapsed_ns.sample_order")
    if operation["alignment"] != OPERATION_ALIGNMENT:
        raise _error(f"{location}.alignment does not match the harness")
    if operation["latency_claim"] != OPERATION_LATENCY_CLAIM:
        raise _error(f"{location}.latency_claim must be {OPERATION_LATENCY_CLAIM!r}")

    source = _exact_keys(operation["source"], f"{location}.source", OPERATION_SOURCE_KEYS)
    if source["status"] != "not_applicable":
        raise _error(f"{location}.source.status must be 'not_applicable'")
    if source["counter_scope"] != "not_applicable_in_process_sink":
        raise _error(
            f"{location}.source.counter_scope must identify the in-process sink"
        )
    source_scopes = {
        "logical_read_calls": SOURCE_SCOPE,
        "logical_read_requested_bytes": SOURCE_SCOPE,
        "logical_read_returned_bytes": SOURCE_SCOPE,
        "logical_read_largest_requested_bytes": SOURCE_SCOPE,
        "logical_read_largest_returned_bytes": SOURCE_SCOPE,
        "logical_read_pattern": SOURCE_PATTERN_SCOPE,
        "compressed_bytes": SOURCE_COMPRESSED_SCOPE,
        "decompressed_bytes": SOURCE_DECOMPRESSED_SCOPE,
        "recompressed_bytes": SOURCE_RECOMPRESSED_SCOPE,
        "max_concurrent_reads": SOURCE_SCOPE,
    }
    for field, scope in source_scopes.items():
        _metric_vector(
            source[field],
            f"{location}.source.{field}",
            sample_count,
            expected_status="not_applicable",
            expected_scope=scope,
        )

    process = _exact_keys(operation["process"], f"{location}.process", OPERATION_PROCESS_KEYS)
    if process["status"] != "unavailable":
        raise _error(f"{location}.process.status must be 'unavailable'")
    process_scopes = {
        "user_cpu_ticks": IN_PROCESS_PROCESS_SCOPE,
        "system_cpu_ticks": IN_PROCESS_PROCESS_SCOPE,
        "clock_ticks_per_second": IN_PROCESS_PROCESS_SCOPE,
        "minor_faults": IN_PROCESS_PROCESS_SCOPE,
        "major_faults": IN_PROCESS_PROCESS_SCOPE,
        "voluntary_context_switches": IN_PROCESS_PROCESS_SCOPE,
        "nonvoluntary_context_switches": IN_PROCESS_PROCESS_SCOPE,
        "rss_delta_bytes": IN_PROCESS_RSS_SCOPE,
        "peak_rss_bytes": HWM_SCOPE,
        "rchar": IN_PROCESS_PROCESS_SCOPE,
        "wchar": IN_PROCESS_PROCESS_SCOPE,
        "read_bytes": IN_PROCESS_PROCESS_SCOPE,
        "write_bytes": IN_PROCESS_PROCESS_SCOPE,
        "cancelled_write_bytes": IN_PROCESS_PROCESS_SCOPE,
        "syscr": IN_PROCESS_PROCESS_SCOPE,
        "syscw": IN_PROCESS_PROCESS_SCOPE,
    }
    for field, scope in process_scopes.items():
        _metric_vector(
            process[field],
            f"{location}.process.{field}",
            sample_count,
            expected_status="unavailable",
            expected_scope=scope,
        )

    operation_sink = _exact_keys(
        operation["sink"], f"{location}.sink", OPERATION_SINK_KEYS
    )
    if operation_sink["status"] != "not_applicable":
        raise _error(f"{location}.sink.status must be 'not_applicable'")
    _metric_vector(
        operation_sink["output_bytes"],
        f"{location}.sink.output_bytes",
        sample_count,
        expected_status="not_applicable",
        expected_scope=OUTPUT_BYTES_SCOPE,
    )
    if operation_sink["write_status"] != "measured":
        raise _error(f"{location}.sink.write_status must be 'measured'")
    expected_sink_vectors = {
        "accepted_bytes": ([sink["accepted_bytes"]] * sample_count, SINK_ACCEPTED_BYTES_SCOPE),
        "write_calls": ([sink["write_calls"]] * sample_count, SINK_WRITE_CALLS_SCOPE),
        "largest_write": ([sink["largest_write"]] * sample_count, SINK_LARGEST_WRITE_SCOPE),
    }
    for field, (expected_values, scope) in expected_sink_vectors.items():
        _metric_vector(
            operation_sink[field],
            f"{location}.sink.{field}",
            sample_count,
            expected_status="measured",
            expected_scope=scope,
            expected_values=expected_values,
        )
    operation_buckets = _exact_keys(
        operation_sink["write_size_buckets"],
        f"{location}.sink.write_size_buckets",
        OPERATION_BUCKET_KEYS,
    )
    if operation_buckets["status"] != "measured":
        raise _error(f"{location}.sink.write_size_buckets.status must be 'measured'")
    for field in SINK_BUCKET_KEYS:
        _metric_vector(
            operation_buckets[field],
            f"{location}.sink.write_size_buckets.{field}",
            sample_count,
            expected_status="measured",
            expected_scope=SINK_BUCKET_SCOPE,
            expected_values=[sink["write_size_buckets"][field]] * sample_count,
        )

    for group_name, group, group_keys in (
        (
            "publication",
            operation["publication"],
            OPERATION_PUBLICATION_KEYS,
        ),
        (
            "materialization",
            operation["materialization"],
            OPERATION_MATERIALIZATION_KEYS,
        ),
    ):
        group_object = _exact_keys(group, f"{location}.{group_name}", group_keys)
        if group_object["status"] != "not_applicable":
            raise _error(f"{location}.{group_name}.status must be 'not_applicable'")
        for field, metric in group_object.items():
            if field == "status":
                continue
            _metric_vector(
                metric,
                f"{location}.{group_name}.{field}",
                sample_count,
                expected_status="not_applicable",
                expected_scope=(
                    "logical_publication_counter"
                    if group_name == "publication"
                    else "logical_materialization_counter"
                ),
            )

    phases = _exact_keys(operation["cfb_phases"], f"{location}.cfb_phases", OPERATION_CFB_KEYS)
    if phases["status"] != "not_applicable":
        raise _error(f"{location}.cfb_phases.status must be 'not_applicable'")
    for phase_name in ("open", "plan", "atomic_publication"):
        phase = _exact_keys(
            phases[phase_name],
            f"{location}.cfb_phases.{phase_name}",
            OPERATION_CFB_PHASE_KEYS,
        )
        for field, metric in phase.items():
            _metric_vector(
                metric,
                f"{location}.cfb_phases.{phase_name}.{field}",
                sample_count,
                expected_status="not_applicable",
                expected_scope=(
                    CFB_PHASE_ELAPSED_SCOPE
                    if field == "elapsed_ns"
                    else CFB_PHASE_SOURCE_SCOPE
                ),
            )

    allocation = _exact_keys(
        operation["allocation"], f"{location}.allocation", ALLOCATION_KEYS
    )
    if allocation["status"] != "unavailable":
        raise _error(f"{location}.allocation.status must be 'unavailable'")
    if allocation["scope"] != ALLOCATION_SCOPE:
        raise _error(f"{location}.allocation.scope must be {ALLOCATION_SCOPE!r}")
    for field in ALLOCATION_VECTOR_FIELDS:
        _metric_vector(
            allocation[field],
            f"{location}.allocation.{field}",
            sample_count,
            expected_status="unavailable",
            expected_scope=ALLOCATION_SCOPE,
        )
    return operation


def _validate_row(
    row: Any, location: str, *, shape: str, count: int, sample_count: int
) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any], str]:
    row_object = _row_keys(row, location)
    if row_object["case"] != CASE:
        raise _error(f"{location}.case must be {CASE!r}")
    corpus = _validate_corpus(row_object["corpus"], f"{location}.corpus", shape, count)
    elapsed_samples, sample_order = _validate_elapsed(
        row_object["elapsed_ns"], f"{location}.elapsed_ns", sample_count
    )
    source = _exact_keys(row_object["source"], f"{location}.source", SOURCE_KEYS)
    source_vectors = {
        field: _vector(source[field], f"{location}.source.{field}", sample_count)
        for field in SOURCE_VECTOR_FIELDS
    }
    overlay = _exact_keys(
        source["opc_source_overlay"], f"{location}.source.opc_source_overlay", OVERLAY_KEYS
    )
    expected = EXPECTED_CORPORA[shape]
    if overlay["implementation"] != "SourceBackedPackage::write_part_overlays_to_stream":
        raise _error(f"{location}.source.opc_source_overlay.implementation does not match")
    if overlay["performance_claim"] != "none":
        raise _error(f"{location}.source.opc_source_overlay.performance_claim must be 'none'")
    if overlay["overlay_mode"] != "noop":
        raise _error(f"{location}.source.opc_source_overlay.overlay_mode must be 'noop'")
    if overlay["replacement_semantics"] != "non-empty equal-payload replacement plan; semantic no-op":
        raise _error(f"{location}.source.opc_source_overlay.replacement_semantics does not match")
    if overlay["timing_scope"] != TIMING_SCOPE:
        raise _error(
            f"{location}.source.opc_source_overlay.timing_scope does not match "
            "the exact harness constant"
        )
    if overlay["overlay_count"] != count or overlay["source_shape"] != shape or overlay["payload_kind"] != expected["payload_kind"]:
        raise _error(f"{location}.source.opc_source_overlay matrix identity disagrees")
    if overlay["source_bytes"] != expected["archive_bytes"] or overlay["source_cache_max_bytes"] != expected["uncompressed_payload_bytes"] or overlay["source_cache_max_entries"] != expected["entry_count"]:
        raise _error(f"{location}.source.opc_source_overlay source limits disagree")
    if overlay["sink_max_write"] != 65_536:
        raise _error(f"{location}.source.opc_source_overlay.sink_max_write must be 65536")
    source_digest = _digest(overlay["source_sha256"], f"{location}.source.opc_source_overlay.source_sha256")
    if source_digest != expected["archive_sha256"] or _digest(overlay["expected_eager_sha256"], f"{location}.source.opc_source_overlay.expected_eager_sha256") != source_digest:
        raise _error(f"{location}.source.opc_source_overlay source/eager digest identity disagrees")
    if overlay["expected_eager_semantic_verified"] is not True or overlay["raw_members_and_order_preservation_verified"] is not True or overlay["equal_payload_noop_source_verified"] is not True:
        raise _error(f"{location}.source.opc_source_overlay correctness or noop oracle is false")
    output_digest = _digest(row_object["output_sha256"], f"{location}.output_sha256")
    if output_digest != source_digest:
        raise _error(f"{location}.output_sha256 must equal the noop source digest")
    observed_digests = overlay["observed_output_sha256"]
    if not isinstance(observed_digests, list) or len(observed_digests) != sample_count:
        raise _error(f"{location}.source.opc_source_overlay.observed_output_sha256 cardinality differs")
    if any(_digest(item, f"{location}.source.opc_source_overlay.observed_output_sha256[{index}]") != output_digest for index, item in enumerate(observed_digests)):
        raise _error(f"{location}.source.opc_source_overlay observed output digest differs")

    phases = {
        field: _vector(overlay[field], f"{location}.source.opc_source_overlay.{field}", sample_count)
        for field in PHASE_FIELDS
    }
    counters = {
        field: _vector(
            overlay[field],
            f"{location}.source.opc_source_overlay.{field}",
            sample_count,
        )
        for field in COUNTER_FIELDS
    }
    for index, elapsed in enumerate(elapsed_samples):
        phase_sum = sum(phases[field][index] for field in PHASE_FIELDS)
        if phase_sum > U64_MAX or phase_sum != elapsed:
            raise _error(f"{location} phase vectors do not sum to elapsed_ns.samples[{index}]")
    for field in (
        "cache_before_publication_hits",
        "cache_before_publication_cold_loads",
        "cache_before_publication_retained_entries",
        "cache_before_publication_retained_bytes",
        "reopened_output_cache_hits",
        "reopened_output_cache_cold_loads",
        "reopened_output_cache_retained_entries",
        "reopened_output_cache_retained_bytes",
    ):
        if any(counters[field]):
            raise _error(f"{location}.source.opc_source_overlay.{field} must be zero")
    expected_probe_bytes = count * expected["entry_bytes"]
    for field, expected_value in (
        ("source_cache_after_publication_probe_hits", 0),
        ("source_cache_after_publication_probe_cold_loads", count),
        ("source_cache_after_publication_probe_retained_entries", count),
        ("source_cache_after_publication_probe_retained_bytes", expected_probe_bytes),
    ):
        if any(value != expected_value for value in counters[field]):
            raise _error(f"{location}.source.opc_source_overlay.{field} disagrees with selected count")
    for nested, root in (
        ("observed_after_publication_source_read_calls", "read_calls"),
        ("observed_after_publication_source_read_bytes", "read_bytes"),
        ("observed_after_publication_ordinary_payload_read_calls", "ordinary_payload_read_calls"),
        ("observed_after_publication_ordinary_payload_read_bytes", "ordinary_payload_read_bytes"),
    ):
        if counters[nested] != source_vectors[root]:
            raise _error(f"{location}.source.opc_source_overlay.{nested} disagrees with source.{root}")
    if any(
        value == 0
        for value in source_vectors["read_calls"]
        + source_vectors["read_bytes"]
        + source_vectors["ordinary_payload_read_calls"]
        + source_vectors["ordinary_payload_read_bytes"]
    ):
        raise _error(f"{location}.source read oracle contains zero evidence")
    if any(value < count for value in source_vectors["ordinary_payload_read_calls"]):
        raise _error(f"{location}.source ordinary payload reads are below overlay_count")
    if any(value != 1 for value in source_vectors["max_in_flight_reads"]):
        raise _error(
            f"{location}.source.max_in_flight_reads must equal the fixed one-worker "
            "concurrency bound"
        )

    sink_max_bytes = _u64(overlay["sink_max_bytes"], f"{location}.source.opc_source_overlay.sink_max_bytes", positive=True)
    sink = _validate_sink(
        row_object["sink"],
        f"{location}.sink",
        max_bytes=sink_max_bytes,
        expected_accepted_bytes=expected["archive_bytes"],
    )
    _validate_operation_metrics(
        row_object["operation_metrics"],
        f"{location}.operation_metrics",
        sample_count=sample_count,
        elapsed_sample_order=sample_order,
        sink=sink,
    )
    return row_object, corpus, source, sink, output_digest


def _source_identity(source: Mapping[str, Any], location: str) -> str:
    identity = copy.deepcopy(dict(source))
    for field in SOURCE_VECTOR_FIELDS:
        identity.pop(field, None)
    overlay = dict(identity["opc_source_overlay"])
    for field in OVERLAY_DYNAMIC_FIELDS:
        overlay.pop(field, None)
    identity["opc_source_overlay"] = overlay
    return _canonical(identity, location)


def _source_counter_identity(source: Mapping[str, Any], location: str) -> str:
    """Return the canonical identity of all measured source/cache vectors."""

    overlay = source["opc_source_overlay"]
    vectors = {
        "source": {field: source[field] for field in SOURCE_VECTOR_FIELDS},
        "overlay": {field: overlay[field] for field in COUNTER_FIELDS},
    }
    return _canonical(vectors, location)


def _oracle_identity(row: Mapping[str, Any], location: str) -> str:
    overlay = row["source"]["opc_source_overlay"]
    oracle = {
        "expected_eager_sha256": overlay["expected_eager_sha256"],
        "expected_eager_semantic_verified": overlay[
            "expected_eager_semantic_verified"
        ],
        "raw_members_and_order_preservation_verified": overlay[
            "raw_members_and_order_preservation_verified"
        ],
        "equal_payload_noop_source_verified": overlay[
            "equal_payload_noop_source_verified"
        ],
        "output_sha256": row["output_sha256"],
    }
    return _canonical(oracle, location)


def _statistic(samples: Sequence[int]) -> dict[str, Any]:
    if not samples:
        raise _error("publication_ns must not be empty")
    ordered = sorted(
        _u64(value, f"publication_ns[{index}]", positive=True)
        for index, value in enumerate(samples)
    )
    mean = 0.0
    for index, value in enumerate(ordered, start=1):
        mean += (float(value) - mean) / float(index)
    p50_left = ordered[(len(ordered) - 1) // 2]
    p50_right = ordered[len(ordered) // 2]
    p50 = p50_left // 2 + p50_right // 2 + (p50_left % 2 + p50_right % 2) // 2

    def nearest_rank(percentile: int) -> int:
        index = ((percentile * len(ordered) + 99) // 100) - 1
        return ordered[min(index, len(ordered) - 1)]

    return {"p50": p50, "mean": mean, "p95": nearest_rank(95), "p99": nearest_rank(99)}


def _reduction(control: float, candidate: float, location: str) -> float:
    if control <= 0:
        raise _error(f"{location} control statistic must be positive")
    value = (control - candidate) / control * 100.0
    if not math.isfinite(value):
        raise _error(f"{location} reduction is not finite")
    return value


def _drift(first: float, second: float, location: str) -> float:
    if first <= 0:
        raise _error(f"{location} baseline statistic must be positive")
    value = (second - first) / first * 100.0
    if not math.isfinite(value):
        raise _error(f"{location} drift is not finite")
    return value


def _phase_summary(
    stats: Mapping[str, Mapping[str, Any]], location: str
) -> dict[str, Any]:
    reductions = {
        "a1_to_b1": {
            name: _reduction(stats["a1"][name], stats["b1"][name], f"{location}.{name}.a1_to_b1")
            for name in STATISTICS
        },
        "a2_to_b2": {
            name: _reduction(stats["a2"][name], stats["b2"][name], f"{location}.{name}.a2_to_b2")
            for name in STATISTICS
        },
    }
    drift = {
        "control": {
            name: _drift(stats["a1"][name], stats["a2"][name], f"{location}.control.{name}")
            for name in STATISTICS
        },
        "candidate": {
            name: _drift(stats["b1"][name], stats["b2"][name], f"{location}.candidate.{name}")
            for name in STATISTICS
        },
    }
    within = {
        implementation: {
            name: abs(drift[implementation][name]) <= DRIFT_CEILINGS[name]
            for name in STATISTICS
        }
        for implementation in ("control", "candidate")
    }
    authorized: dict[str, bool] = {}
    rejected: dict[str, list[str]] = {}
    for name in STATISTICS:
        reasons: list[str] = []
        if reductions["a1_to_b1"][name] <= 0 or reductions["a2_to_b2"][name] <= 0:
            reasons.append("candidate reduction is not positive in both paired directions")
        for implementation in ("control", "candidate"):
            if not within[implementation][name]:
                reasons.append(
                    f"{implementation} drift {drift[implementation][name]:+.6f}% exceeds {DRIFT_CEILINGS[name]:g}% ceiling"
                )
        authorized[name] = not reasons
        if reasons:
            rejected[name] = reasons
    return {
        "legs_ns": {role: dict(stats[role]) for role in ROLES},
        "candidate_reduction_percent": reductions,
        "same_implementation_drift_percent": drift,
        "drift_ceiling_percent": dict(DRIFT_CEILINGS),
        "drift_within_ceiling": within,
        "authorized_statistics": authorized,
        "accepted_statistics": [name for name in STATISTICS if authorized[name]],
        "rejected_statistics": rejected,
    }


def summarize_reports(
    reports: Mapping[str, Any] | Sequence[Mapping[str, Any]],
    b1: Mapping[str, Any] | None = None,
    b2: Mapping[str, Any] | None = None,
    a2: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate four reports and return publication-only ABBA evidence."""

    if any(value is not None for value in (b1, b2, a2)):
        if b1 is None or b2 is None or a2 is None:
            raise _error("all four ABBA reports are required")
        report_values = (reports, b1, b2, a2)
    elif isinstance(reports, Mapping):
        if set(reports) != set(ROLES):
            missing = sorted(set(ROLES) - set(reports))
            unknown = sorted(set(reports) - set(ROLES), key=repr)
            raise _error(
                f"reports mapping keys mismatch (missing={missing!r}, "
                f"unknown={unknown!r})"
            )
        try:
            report_values = tuple(reports[role] for role in ROLES)
        except KeyError as error:
            raise _error(f"reports mapping must contain {list(ROLES)!r}") from error
    elif isinstance(reports, Sequence) and not isinstance(reports, (str, bytes)) and len(reports) == 4:
        report_values = tuple(reports)
    else:
        raise _error("reports must contain exactly four reports in A1,B1,B2,A2 order")

    validated: dict[str, tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any], str]] = {}
    canonical_report_hashes: dict[str, str] = {}
    for role, report in zip(ROLES, report_values):
        label = role.upper()
        if not isinstance(report, Mapping):
            raise _error(f"{label} report must be an object")
        try:
            perf_abba_summary._validate_report(
                report, label, profile=REPORT_PROFILE, report_role=role
            )
        except Exception as error:
            raise _error(f"{label} generic report validation failed: {error}") from error
        _validate_root_shape(report, label)
        canonical_report_hashes[role] = _sha256(report, f"{label}.report")
        validated_rows: dict[tuple[str, int], tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any], str]] = {}
        sample_count = report["configuration"]["samples_per_case"]
        results = report["results"]
        if not isinstance(results, list) or len(results) != len(SHAPES) * len(COUNTS):
            raise _error(f"{label}.results must contain exactly 9 matrix rows")
        expected_row_order = [(shape, count) for shape in SHAPES for count in COUNTS]
        for index, row in enumerate(results):
            row_location = f"{label}.results[{index}]"
            row_object = _require_object(row, row_location)
            corpus = _require_object(row_object.get("corpus"), f"{row_location}.corpus")
            shape = corpus.get("shape")
            if shape not in EXPECTED_CORPORA:
                raise _error(f"{row_location}.corpus.shape is outside the fixed matrix")
            name = corpus.get("name")
            prefix = EXPECTED_CORPORA[shape]["name_prefix"]
            if not isinstance(name, str) or not name.startswith(prefix):
                raise _error(f"{row_location}.corpus.name does not identify the fixed matrix")
            suffix = name[len(prefix) :]
            if not suffix.isdigit() or int(suffix) not in COUNTS:
                raise _error(f"{row_location}.corpus.name has an invalid overlay count")
            count = int(suffix)
            key = (shape, count)
            if key != expected_row_order[index]:
                raise _error(
                    f"{label}.results[{index}] is not matrix cell "
                    f"{expected_row_order[index]!r}"
                )
            if key in validated_rows:
                raise _error(f"{label}.results contains duplicate matrix cell {key!r}")
            validated_rows[key] = _validate_row(
                row, row_location, shape=shape, count=count, sample_count=sample_count
            )
        expected_keys = {(shape, count) for shape in SHAPES for count in COUNTS}
        if set(validated_rows) != expected_keys:
            raise _error(f"{label}.results does not cover exactly the fixed 3-shape x 3-count matrix")
        validated[role] = (
            report["tool"],
            report["environment"],
            report["configuration"],
            validated_rows,
            canonical_report_hashes[role],
        )

    if len(set(canonical_report_hashes.values())) != len(ROLES):
        raise _error("A1/B1/B2/A2 reports must be four distinct JSON reports")
    tool_ids = {_canonical(validated[role][0], f"{role}.tool") for role in ROLES}
    if len(tool_ids) != 1:
        raise _error("tool identity differs between ABBA legs")
    stable_env_ids = {
        role: _canonical(
            {key: value for key, value in validated[role][1].items() if key != "git_revision"},
            f"{role}.environment.stable",
        )
        for role in ROLES
    }
    if len(set(stable_env_ids.values())) != 1:
        raise _error("stable environment identity differs between ABBA legs")
    if validated["a1"][1]["git_revision"] != validated["a2"][1]["git_revision"]:
        raise _error("control A1/A2 git_revision differs")
    if validated["b1"][1]["git_revision"] != validated["b2"][1]["git_revision"]:
        raise _error("candidate B1/B2 git_revision differs")
    if validated["a1"][1]["git_revision"] == validated["b1"][1]["git_revision"]:
        raise _error("control and candidate git_revision must differ")

    binary_ids = {
        role: _canonical(report_values[index]["binary_identity"], f"{role}.binary_identity")
        for index, role in enumerate(ROLES)
    }
    if binary_ids["a1"] != binary_ids["a2"] or binary_ids["b1"] != binary_ids["b2"]:
        raise _error("same-implementation binary identity differs between ABBA legs")
    if report_values[0]["binary_identity"]["binary_sha256"] == report_values[1]["binary_identity"]["binary_sha256"]:
        raise _error("control and candidate binary SHA-256 hashes must differ")
    configs = {_canonical(validated[role][2], f"{role}.configuration") for role in ROLES}
    if len(configs) != 1:
        raise _error("configuration identity differs between ABBA legs")

    first_rows = validated["a1"][3]
    results: list[dict[str, Any]] = []
    for shape in SHAPES:
        for count in COUNTS:
            key = (shape, count)
            rows = {role: validated[role][3][key] for role in ROLES}
            source_ids = {role: _source_identity(rows[role][2], f"{role}.{shape}[{count}].source") for role in ROLES}
            source_counter_ids = {
                role: _source_counter_identity(
                    rows[role][2], f"{role}.{shape}[{count}].source_counters"
                )
                for role in ROLES
            }
            sink_ids = {role: _canonical(rows[role][3], f"{role}.{shape}[{count}].sink") for role in ROLES}
            output_ids = {role: rows[role][4] for role in ROLES}
            oracle_ids = {
                role: _oracle_identity(
                    rows[role][0], f"{role}.{shape}[{count}].oracle"
                )
                for role in ROLES
            }
            if len(set(source_ids.values())) != 1:
                raise _error(f"{shape}[{count}] source identity differs between ABBA legs")
            if len(set(source_counter_ids.values())) != 1:
                raise _error(
                    f"{shape}[{count}] source/read/cache counter vectors differ "
                    "between ABBA legs"
                )
            if len(set(sink_ids.values())) != 1:
                raise _error(f"{shape}[{count}] sink identity differs between ABBA legs")
            if len(set(output_ids.values())) != 1:
                raise _error(f"{shape}[{count}] noop output identity differs between ABBA legs")
            if len(set(oracle_ids.values())) != 1:
                raise _error(f"{shape}[{count}] oracle identity differs between ABBA legs")
            pub_stats = {
                role: _statistic(rows[role][2]["opc_source_overlay"]["publication_ns"])
                for role in ROLES
            }
            phase = _phase_summary(pub_stats, f"{shape}[{count}].publication_ns")
            phase["sample_count"] = validated["a1"][2]["samples_per_case"]
            results.append(
                {
                    "case": CASE,
                    "shape": shape,
                    "overlay_count": count,
                    "corpus": first_rows[key][1],
                    "identity": {
                        "source_status": "verified_equal",
                        "source_canonical_json": source_ids["a1"],
                        "source_counter_vectors_status": "verified_equal",
                        "source_counter_vectors_canonical_json": source_counter_ids["a1"],
                        "sink_status": "verified_equal",
                        "sink_canonical_json": sink_ids["a1"],
                        "output_sha256_status": "verified_equal",
                        "output_sha256": output_ids["a1"],
                        "oracle_status": "verified_equal",
                        "oracle_canonical_json": oracle_ids["a1"],
                        "phase_sum_verified": True,
                    },
                    "publication_ns": phase,
                }
            )

    tool = copy.deepcopy(validated["a1"][0])
    environments = {role: copy.deepcopy(validated[role][1]) for role in ROLES}
    configuration = copy.deepcopy(validated["a1"][2])
    return {
        "schema_version": SCHEMA_VERSION,
        "tool": {"name": TOOL_NAME, "version": TOOL_VERSION},
        "protocol": {
            "order": ["a1_control", "b1_candidate", "b2_candidate", "a2_control"],
            "samples_per_leg": SAMPLE_COUNT,
            "warmup_iterations_per_leg": WARMUP_ITERATIONS,
            "filesystem_cache_states": list(FILESYSTEM_CACHE_STATES),
            "filesystem_root_selected": False,
            "execution_workers": list(EXECUTION_WORKERS),
            "cpu_affinity": CPU_AFFINITY,
            "rustc_version": RUSTC_VERSION,
            "allocator": ALLOCATOR,
            "instrumentation": "none",
            "timed_metric": "source.opc_source_overlay.publication_ns",
            "statistics": list(STATISTICS),
            "drift_ceiling_percent": dict(DRIFT_CEILINGS),
            "percentiles": "p50 = Rust u64 floor midpoint; p95/p99 = integer nearest-rank",
            "mean": "Welford floating-point mean over nested publication_ns samples",
            "sample_scope": "500 retained in-process samples per normal warm leg",
            "claim_scope": "normal warm ABBA in-process publication phase only",
            "filesystem_fresh_child_semantics_claimed": False,
            "excluded_metrics": ["top-level elapsed_ns", "allocator timing"],
        },
        "harness_identity": {
            "schema_version": REPORT_SCHEMA_VERSION,
            "tool": tool,
            "configuration": configuration,
        },
        "environment": {
            "stable": json.loads(next(iter(stable_env_ids.values()))),
            "legs": environments,
        },
        "implementation_identity": {
            "control": {
                "git_revision": environments["a1"]["git_revision"],
                "binary_identity": copy.deepcopy(report_values[0]["binary_identity"]),
                "legs": ["a1", "a2"],
            },
            "candidate": {
                "git_revision": environments["b1"]["git_revision"],
                "binary_identity": copy.deepcopy(report_values[1]["binary_identity"]),
                "legs": ["b1", "b2"],
            },
            "distinct": True,
        },
        "report_identity": {
            role: {"canonical_sha256": canonical_report_hashes[role]} for role in ROLES
        },
        "results": results,
        "verification": {
            "result_count": len(results),
            "matrix": "3 shapes x counts 2,8,32",
            "tool_identity_verified": True,
            "provenance_identity_verified": True,
            "environment_stable_identity_verified": True,
            "configuration_identity_verified": True,
            "formal_protocol_verified": True,
            "corpus_identity_verified": True,
            "source_identity_verified": True,
            "source_counter_vectors_verified": True,
            "sink_identity_verified": True,
            "oracle_identity_verified": True,
            "phase_sum_verified": True,
            "operation_metrics_verified": True,
            "sample_order_binding_verified": True,
            "publication_statistics_recomputed_from_nested_samples": True,
            "total_elapsed_statistics_claimed": False,
            "allocator_timing_claimed": False,
            "allocation_delta_checked": False,
            "in_process_sample_scope_verified": True,
            "filesystem_fresh_child_semantics_claimed": False,
        },
    }


build_summary = summarize_reports
summarize_abba = summarize_reports


def load_report(path: Path) -> dict[str, Any]:
    """Load one report with duplicate-key and non-finite-number rejection."""

    def reject_nonfinite(value: str) -> None:
        raise ValueError(f"non-finite JSON value {value!r}")

    def reject_duplicate(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"duplicate JSON object key {key!r}")
            result[key] = value
        return result

    try:
        with path.open(encoding="utf-8") as handle:
            value = json.load(
                handle, object_pairs_hook=reject_duplicate, parse_constant=reject_nonfinite
            )
    except (OSError, ValueError) as error:
        raise OverlayAbbaInputError(f"cannot read {path}: {error}") from error
    return _require_object(value, str(path))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("reports", nargs="*", type=Path, metavar="REPORT")
    for role in ROLES:
        parser.add_argument(f"--{role}", type=Path, help=f"{role.upper()} normal report")
    parser.add_argument("--json-out", "--output", type=Path)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = build_parser().parse_args(argv)
    try:
        named = [getattr(arguments, role) for role in ROLES]
        if arguments.reports and any(path is not None for path in named):
            raise _error("use positional reports or named --a1/--b1/--b2/--a2 reports")
        if arguments.reports:
            if len(arguments.reports) != 4:
                raise _error("exactly four positional reports are required")
            paths = arguments.reports
        else:
            if any(path is None for path in named):
                raise _error("all four named reports are required")
            paths = named
        reports = {role: load_report(path) for role, path in zip(ROLES, paths)}
        summary = summarize_reports(reports)
        encoded = json.dumps(summary, indent=2, sort_keys=True, allow_nan=False) + "\n"
        if arguments.json_out is None:
            sys.stdout.write(encoded)
        else:
            try:
                with arguments.json_out.open("x", encoding="utf-8") as handle:
                    handle.write(encoded)
            except (FileExistsError, OSError) as error:
                raise _error(f"cannot create {arguments.json_out}: {error}") from error
    except (OverlayAbbaInputError, OSError, ValueError) as error:
        print(f"{TOOL_NAME}: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":  # pragma: no cover - exercised by CLI tests
    raise SystemExit(main())
