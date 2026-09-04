#!/usr/bin/env python3
"""Fail-closed validator and projection builder for Change 0401 allocator ABBA.

This validator is intentionally separate from the Change 0400 validator.  The
0400 validator and its evidence are already retained and their bytes are part
of that change's provenance.  Change 0401 has its own fixed binary identities,
raw allocator vectors, and report contract.

Only the six operation-scoped call/byte metrics are claimable observations.
Allocator elapsed time and the absolute live/high-water snapshots are retained
for alignment and auditability, but are explicitly non-claimable.

The module uses only Python's standard library.  It rejects duplicate JSON
keys, non-finite numbers, schema drift, raw/vector misalignment, reused fresh
child processes, provenance drift, and projection drift.
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
from typing import Any


SHA256 = re.compile(r"^[0-9a-f]{64}$")
CASE = "xlsx_file_selected_cell"
CACHE_STATE = "warm"
SAMPLE_COUNT = 30
WARMUP_ITERATIONS = 3
ALLOCATION_SCOPE = "operation_global_system_allocator"
RUSTC_VERSION = "rustc 1.98.1 (48a229cea 2026-09-01)"
ALLOCATOR = "CountingSystemAllocator(std::alloc::System)"
CORPUS_SHA256 = "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036"
SEMANTIC_SHA256 = "020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e"

SELECTED_CELL = {
    "canonical_sheet_name": "Bench01",
    "sheet_position": 1,
    "prepared_selector": "bEnCh01",
    "cell_address": "M29",
    "view": "stored",
    "value_kind": "number",
    "lexical_value": "1028012",
    "digest": "36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1",
}

METRICS = (
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
CLAIMABLE_METRICS = frozenset(METRICS[:6])
NONCLAIMABLE_METRICS = frozenset(METRICS[6:])
EXPECTED_CLAIMABLE_VECTORS = {
    "control": (84221, 84206, 12, 0, 10706565, 10705182),
    "candidate": (81918, 81903, 12, 0, 10690444, 10689061),
}

LEGS = {
    "A1": {
        "implementation": "control",
        "revision": "0859063be5a67bd2aafb3531f2126020b2b5000d",
        "binary_sha256": "2173aa6e718eb056f22c9187279e632fff659850c630438376ac29c577704f71",
        "binary_bytes": 54767312,
    },
    "B1": {
        "implementation": "candidate",
        "revision": "87f26d5ee02a1903e668bf7f60fa3ef954a0c3fb",
        "binary_sha256": "068bef7e6be0824f8bde4a5494e61e02dc0feb40653e748d91c86a5058c40d66",
        "binary_bytes": 54758784,
    },
    "B2": {
        "implementation": "candidate",
        "revision": "87f26d5ee02a1903e668bf7f60fa3ef954a0c3fb",
        "binary_sha256": "068bef7e6be0824f8bde4a5494e61e02dc0feb40653e748d91c86a5058c40d66",
        "binary_bytes": 54758784,
    },
    "A2": {
        "implementation": "control",
        "revision": "0859063be5a67bd2aafb3531f2126020b2b5000d",
        "binary_sha256": "2173aa6e718eb056f22c9187279e632fff659850c630438376ac29c577704f71",
        "binary_bytes": 54767312,
    },
}

REPORT_KEYS = {
    "schema_version",
    "tool",
    "binary_identity",
    "environment",
    "configuration",
    "parallel_metrics",
    "results",
    "filesystem_evidence",
}
TOOL_KEYS = {
    "name",
    "version",
    "binary",
    "profile",
    "target_os",
    "target_arch",
    "instrumentation",
}
BINARY_KEYS = {
    "path",
    "binary_sha256",
    "binary_bytes",
    "mode_bits",
    "executable",
    "profile",
}
ENVIRONMENT_KEYS = {
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
RESULT_KEYS = {"case", "cache_state", "corpus", "elapsed_ns", "sink", "operation_metrics"}
FILESYSTEM_ENVELOPE_KEYS = {
    "case",
    "corpus",
    "warmup_iterations",
    "sample_count",
    "cache_states",
    "fresh_child_per_sample",
    "samples",
}
FILESYSTEM_SAMPLE_KEYS = {
    "child_process_id",
    "sample_index",
    "cache_state",
    "elapsed_ns",
    "parent_wall_ns",
    "cold_advice",
    "logical_read_counter_scope",
    "logical_read_calls",
    "logical_read_requested_bytes",
    "logical_read_bytes",
    "logical_read_largest_requested_bytes",
    "logical_read_largest_returned_bytes",
    "max_concurrent_reads",
    "logical_read_request_sizes",
    "logical_read_request_size_buckets",
    "process_metrics",
    "allocation_metrics",
    "output_sha256",
    "output_bytes",
    "opc_materialized_parts",
    "cfb_changed_spans",
    "cfb_published_bytes",
    "xlsx_source_sha256",
    "xlsx_semantic_sha256",
    "xlsx_selected_cell",
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
ALLOCATION_KEYS = {"status", "scope", *METRICS}
VECTOR_KEYS = {"values", "status", "scope"}

EXPECTED_CORPUS = {
    "name": "xlsx-cell-values-medium",
    "generator": "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1",
    "package_format": "XLSX/OPC/ZIP",
    "shape": "medium",
    "payload_kind": "deterministic-multi-sheet-scalar-grid-with-media",
    "compression": "deflate",
    "entry_count": 9216,
    "archive_member_count": 17,
    "entry_bytes": 4,
    "uncompressed_payload_bytes": 4231168,
    "archive_bytes": 4226429,
    "archive_sha256": CORPUS_SHA256,
    "target_entry": "Sheet1!A1",
    "target_payload_bytes": 1,
    "target_payload_sha256": "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
    "xlsx": {
        "sheet_count": 4,
        "rows_per_sheet": 48,
        "columns_per_sheet": 48,
        "one_percent_update_count": 93,
        "source_members": {
            "workbook": "xl/workbook.xml",
            "worksheets": [
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet3.xml",
                "xl/worksheets/sheet4.xml",
            ],
            "shared_strings": None,
            "styles": "xl/styles.xml",
        },
    },
}
EXPECTED_CONFIGURATION = {
    "samples_per_case": SAMPLE_COUNT,
    "warmup_iterations_per_case": WARMUP_ITERATIONS,
    "filesystem_cache_states": [CACHE_STATE],
    "filesystem_fresh_child_per_sample": True,
    "filesystem_process_isolated": True,
    "filesystem_root_selected": True,
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
        "bandwidth_bytes_per_second": 52428800,
        "max_physical_range_bytes": 4096,
    },
    "execution_workers": [1],
}
EXPECTED_ENVIRONMENT_BASE = {
    "rustc_version": RUSTC_VERSION,
    "git_worktree_dirty": False,
    "logical_cpus_available": 1,
    "allocator": ALLOCATOR,
    "rustflags": None,
    "cargo_build_target": None,
    "perf_event_paranoid": "1",
    "os": "linux",
    "kernel": "Linux 7.0.0-1011-aws",
    "cpu_model": "AMD EPYC 9R45",
    "total_memory_bytes": 132553797632,
    "page_size_bytes": 4096,
    "filesystem_type": "tmpfs",
    "source_destination_same_device": True,
    "cpu_affinity": "2",
    "storage_identifier": None,
}


class ValidationError(ValueError):
    """The evidence is malformed, incomplete, or inconsistent."""


@dataclass(frozen=True)
class ValidatedLeg:
    name: str
    raw_sha256: str
    report: dict[str, Any]
    sample_indices: list[int]
    vectors: dict[str, list[int]]
    child_process_ids: list[int]


def _no_duplicate_object_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValidationError(f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _reject_json_constant(token: str) -> Any:
    raise ValidationError(f"non-finite JSON number {token!r}")


def _finite_json_float(token: str) -> float:
    value = float(token)
    if not math.isfinite(value):
        raise ValidationError(f"non-finite JSON number {token!r}")
    return value


def _read_object(path: Path) -> tuple[dict[str, Any], str]:
    try:
        raw = path.read_bytes()
        value = json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=_no_duplicate_object_pairs,
            parse_constant=_reject_json_constant,
            parse_float=_finite_json_float,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValidationError) as error:
        raise ValidationError(f"cannot read {path}: {error}") from error
    if not isinstance(value, dict):
        raise ValidationError(f"{path} must contain a JSON object")
    return value, hashlib.sha256(raw).hexdigest()


def _object(value: Any, context: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValidationError(f"{context} must be an object")
    return value


def _array(value: Any, context: str) -> list[Any]:
    if not isinstance(value, list):
        raise ValidationError(f"{context} must be an array")
    return value


def _integer(value: Any, context: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise ValidationError(f"{context} must be an integer >= {minimum}")
    return value


def _exact_keys(value: Any, keys: set[str], context: str) -> dict[str, Any]:
    obj = _object(value, context)
    actual = set(obj)
    if actual != keys:
        missing = sorted(keys - actual)
        extra = sorted(actual - keys)
        raise ValidationError(
            f"{context} has unexpected keys (missing={missing}, extra={extra})"
        )
    return obj


def _expect(actual: Any, expected: Any, context: str) -> None:
    if actual != expected or type(actual) is not type(expected):
        raise ValidationError(f"{context} must equal {expected!r}, got {actual!r}")


def _sha256(value: Any, context: str) -> str:
    if not isinstance(value, str) or SHA256.fullmatch(value) is None:
        raise ValidationError(f"{context} must be a lowercase SHA-256 digest")
    return value


def _vector_digest(values: list[int]) -> str:
    raw = json.dumps(values, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _projection_metric_scope(metric: str) -> str:
    if metric.startswith("peak_live_bytes_"):
        return "absolute_process_high_water_snapshot"
    if metric.startswith("live_bytes_"):
        return "absolute_process_live_snapshot"
    return ALLOCATION_SCOPE


def _validate_tool(value: Any, context: str) -> None:
    tool = _exact_keys(value, TOOL_KEYS, context)
    _expect(
        tool,
        {
            "name": "litchi-perf-baseline",
            "version": "0.1.0",
            "binary": "litchi-perf-baseline-alloc",
            "profile": "release",
            "target_os": "linux",
            "target_arch": "x86_64",
            "instrumentation": "system_allocator_operation_scoped",
        },
        context,
    )


def _validate_binary(value: Any, leg: str, context: str) -> None:
    binary = _exact_keys(value, BINARY_KEYS, context)
    identity = LEGS[leg]
    if not isinstance(binary["path"], str) or not binary["path"].endswith(
        "/litchi-perf-baseline-alloc"
    ):
        raise ValidationError(f"{context}.path has the wrong allocator binary basename")
    _expect(binary["binary_sha256"], identity["binary_sha256"], f"{context}.binary_sha256")
    _sha256(binary["binary_sha256"], f"{context}.binary_sha256")
    _expect(binary["binary_bytes"], identity["binary_bytes"], f"{context}.binary_bytes")
    _expect(binary["mode_bits"], 509, f"{context}.mode_bits")
    _expect(binary["executable"], True, f"{context}.executable")
    _expect(binary["profile"], "release", f"{context}.profile")


def _validate_environment(value: Any, leg: str, context: str) -> None:
    environment = _exact_keys(value, ENVIRONMENT_KEYS, context)
    expected = dict(EXPECTED_ENVIRONMENT_BASE)
    expected["git_revision"] = LEGS[leg]["revision"]
    _expect(environment, expected, context)


def _validate_configuration(value: Any, context: str) -> None:
    configuration = _exact_keys(value, CONFIGURATION_KEYS, context)
    _expect(configuration, EXPECTED_CONFIGURATION, context)


def _validate_corpus(value: Any, context: str) -> dict[str, Any]:
    corpus = _exact_keys(
        value,
        set(EXPECTED_CORPUS),
        context,
    )
    _expect(corpus, EXPECTED_CORPUS, context)
    return corpus


def _validate_raw_allocation(value: Any, context: str) -> dict[str, int]:
    allocation = _exact_keys(value, ALLOCATION_KEYS, context)
    _expect(allocation["status"], "measured", f"{context}.status")
    _expect(allocation["scope"], ALLOCATION_SCOPE, f"{context}.scope")
    return {
        metric: _integer(allocation[metric], f"{context}.{metric}")
        for metric in METRICS
    }


def _validate_process_metrics(value: Any, context: str) -> None:
    process = _object(value, context)
    if not process:
        raise ValidationError(f"{context} must not be empty")
    for key, item in process.items():
        _integer(item, f"{context}.{key}")


def _validate_sample(
    value: Any,
    leg: str,
    ordinal: int,
) -> tuple[int, int, int, dict[str, int]]:
    context = f"{leg}.filesystem_evidence.samples[{ordinal}]"
    sample = _exact_keys(value, FILESYSTEM_SAMPLE_KEYS, context)
    sample_index = _integer(sample["sample_index"], f"{context}.sample_index")
    child_id = _integer(sample["child_process_id"], f"{context}.child_process_id", minimum=1)
    elapsed = _integer(sample["elapsed_ns"], f"{context}.elapsed_ns", minimum=1)
    _integer(sample["parent_wall_ns"], f"{context}.parent_wall_ns", minimum=1)
    fixed = {
        "cache_state": CACHE_STATE,
        "cold_advice": "not_requested",
        "logical_read_counter_scope": "not_applicable_filesystem_xlsx",
        "logical_read_calls": 0,
        "logical_read_requested_bytes": 0,
        "logical_read_bytes": 0,
        "logical_read_largest_requested_bytes": 0,
        "logical_read_largest_returned_bytes": 0,
        "max_concurrent_reads": 0,
        "logical_read_request_sizes": [],
        "output_sha256": None,
        "output_bytes": None,
        "opc_materialized_parts": None,
        "cfb_changed_spans": None,
        "cfb_published_bytes": None,
        "xlsx_source_sha256": CORPUS_SHA256,
        "xlsx_semantic_sha256": SEMANTIC_SHA256,
        "xlsx_selected_cell": SELECTED_CELL,
    }
    for key, expected in fixed.items():
        _expect(sample[key], expected, f"{context}.{key}")
    buckets = _object(
        sample["logical_read_request_size_buckets"],
        f"{context}.logical_read_request_size_buckets",
    )
    expected_buckets = {
        "bytes_0": 0,
        "bytes_1_to_512": 0,
        "bytes_513_to_4096": 0,
        "bytes_4097_to_16384": 0,
        "bytes_16385_to_65536": 0,
        "bytes_over_65536": 0,
    }
    _expect(buckets, expected_buckets, f"{context}.logical_read_request_size_buckets")
    _validate_process_metrics(sample["process_metrics"], f"{context}.process_metrics")
    allocation = _validate_raw_allocation(
        sample["allocation_metrics"], f"{context}.allocation_metrics"
    )
    return sample_index, child_id, elapsed, allocation


def _validate_operation_vectors(
    value: Any,
    leg: str,
    elapsed_order: list[int],
    samples_by_index: dict[int, tuple[int, dict[str, int]]],
) -> dict[str, list[int]]:
    context = f"{leg}.results[0].operation_metrics"
    operation = _exact_keys(value, OPERATION_KEYS, context)
    _expect(operation["sample_count"], SAMPLE_COUNT, f"{context}.sample_count")
    _expect(operation["sample_indices"], elapsed_order, f"{context}.sample_indices")
    _expect(
        operation["alignment"],
        "elapsed_ns.samples_by_elapsed_then_sample_index",
        f"{context}.alignment",
    )
    _expect(
        operation["latency_claim"],
        "evidence_only_filesystem_selector",
        f"{context}.latency_claim",
    )
    source = _object(operation["source"], f"{context}.source")
    _expect(source.get("status"), "not_applicable", f"{context}.source.status")
    _expect(
        source.get("counter_scope"),
        "not_applicable_filesystem_xlsx",
        f"{context}.source.counter_scope",
    )
    allocation = _exact_keys(operation["allocation"], ALLOCATION_KEYS, f"{context}.allocation")
    _expect(allocation["status"], "measured", f"{context}.allocation.status")
    _expect(allocation["scope"], ALLOCATION_SCOPE, f"{context}.allocation.scope")
    vectors: dict[str, list[int]] = {}
    implementation = LEGS[leg]["implementation"]
    expected_claimable = EXPECTED_CLAIMABLE_VECTORS[implementation]
    for metric_index, metric in enumerate(METRICS):
        vector_context = f"{context}.allocation.{metric}"
        wrapper = _exact_keys(allocation[metric], VECTOR_KEYS, vector_context)
        _expect(wrapper["status"], "measured", f"{vector_context}.status")
        _expect(wrapper["scope"], ALLOCATION_SCOPE, f"{vector_context}.scope")
        values = _array(wrapper["values"], f"{vector_context}.values")
        if len(values) != SAMPLE_COUNT:
            raise ValidationError(f"{vector_context}.values must contain {SAMPLE_COUNT} values")
        typed_values = [
            _integer(item, f"{vector_context}.values[{index}]")
            for index, item in enumerate(values)
        ]
        raw_values = [samples_by_index[index][1][metric] for index in elapsed_order]
        _expect(typed_values, raw_values, f"{vector_context}.values/raw alignment")
        if len(set(typed_values)) != 1:
            raise ValidationError(f"{vector_context}.values must be constant")
        if metric in CLAIMABLE_METRICS:
            _expect(
                typed_values[0],
                expected_claimable[metric_index],
                f"{vector_context}.exact Change 0401 value",
            )
        vectors[metric] = typed_values
    return vectors


def _validate_report(path: Path, leg: str) -> ValidatedLeg:
    report, raw_sha256 = _read_object(path)
    _exact_keys(report, REPORT_KEYS, leg)
    _expect(report["schema_version"], 1, f"{leg}.schema_version")
    _validate_tool(report["tool"], f"{leg}.tool")
    _validate_binary(report["binary_identity"], leg, f"{leg}.binary_identity")
    _validate_environment(report["environment"], leg, f"{leg}.environment")
    _validate_configuration(report["configuration"], f"{leg}.configuration")
    _object(report["parallel_metrics"], f"{leg}.parallel_metrics")

    results = _array(report["results"], f"{leg}.results")
    if len(results) != 1:
        raise ValidationError(f"{leg}.results must contain exactly one result")
    result = _exact_keys(results[0], RESULT_KEYS, f"{leg}.results[0]")
    _expect(result["case"], CASE, f"{leg}.results[0].case")
    _expect(result["cache_state"], CACHE_STATE, f"{leg}.results[0].cache_state")
    corpus = _validate_corpus(result["corpus"], f"{leg}.results[0].corpus")
    _expect(result["sink"], None, f"{leg}.results[0].sink")

    evidence = _array(report["filesystem_evidence"], f"{leg}.filesystem_evidence")
    if len(evidence) != 1:
        raise ValidationError(f"{leg}.filesystem_evidence must contain exactly one entry")
    envelope = _exact_keys(evidence[0], FILESYSTEM_ENVELOPE_KEYS, f"{leg}.filesystem_evidence[0]")
    _expect(envelope["case"], CASE, f"{leg}.filesystem_evidence[0].case")
    _expect(envelope["corpus"], corpus, f"{leg}.filesystem_evidence[0].corpus")
    _expect(envelope["warmup_iterations"], WARMUP_ITERATIONS, f"{leg}.filesystem_evidence[0].warmup_iterations")
    _expect(envelope["sample_count"], SAMPLE_COUNT, f"{leg}.filesystem_evidence[0].sample_count")
    _expect(envelope["cache_states"], [CACHE_STATE], f"{leg}.filesystem_evidence[0].cache_states")
    _expect(envelope["fresh_child_per_sample"], True, f"{leg}.filesystem_evidence[0].fresh_child_per_sample")
    raw_samples = _array(envelope["samples"], f"{leg}.filesystem_evidence[0].samples")
    if len(raw_samples) != SAMPLE_COUNT:
        raise ValidationError(
            f"{leg}.filesystem_evidence[0].samples must contain {SAMPLE_COUNT} samples"
        )
    samples_by_index: dict[int, tuple[int, dict[str, int]]] = {}
    child_ids: list[int] = []
    for ordinal, sample in enumerate(raw_samples):
        sample_index, child_id, elapsed, allocation = _validate_sample(sample, leg, ordinal)
        if sample_index != ordinal:
            raise ValidationError(
                f"{leg} filesystem sample at ordinal {ordinal} has sample_index {sample_index}"
            )
        if sample_index in samples_by_index:
            raise ValidationError(f"{leg} repeats filesystem sample_index {sample_index}")
        samples_by_index[sample_index] = (elapsed, allocation)
        child_ids.append(child_id)
    expected_indices = list(range(SAMPLE_COUNT))
    if sorted(samples_by_index) != expected_indices:
        raise ValidationError(f"{leg} filesystem sample indices are not 0..{SAMPLE_COUNT - 1}")
    if len(set(child_ids)) != SAMPLE_COUNT:
        raise ValidationError(f"{leg} child process IDs are not unique")

    elapsed = _exact_keys(result["elapsed_ns"], ELAPSED_KEYS, f"{leg}.results[0].elapsed_ns")
    _expect(elapsed["unit"], "ns", f"{leg}.results[0].elapsed_ns.unit")
    raw_elapsed_order = _array(elapsed["sample_order"], f"{leg}.results[0].elapsed_ns.sample_order")
    elapsed_order = [
        _integer(item, f"{leg}.results[0].elapsed_ns.sample_order[{index}]")
        for index, item in enumerate(raw_elapsed_order)
    ]
    if sorted(elapsed_order) != expected_indices:
        raise ValidationError(f"{leg} elapsed sample_order must be a permutation")
    elapsed_samples = _array(elapsed["samples"], f"{leg}.results[0].elapsed_ns.samples")
    expected_elapsed = [samples_by_index[index][0] for index in elapsed_order]
    _expect(elapsed_samples, expected_elapsed, f"{leg}.elapsed/raw sample alignment")
    if any(
        isinstance(item, bool) or not isinstance(item, int) or item <= 0
        for item in elapsed_samples
    ):
        raise ValidationError(f"{leg} elapsed samples must be positive integers")
    if elapsed_samples != sorted(elapsed_samples):
        raise ValidationError(f"{leg} elapsed samples must be in elapsed order")
    elapsed_rank = [(samples_by_index[index][0], index) for index in elapsed_order]
    if elapsed_rank != sorted(elapsed_rank):
        raise ValidationError(f"{leg} elapsed ties must be ordered by ascending sample_index")

    vectors = _validate_operation_vectors(
        result["operation_metrics"], leg, elapsed_order, samples_by_index
    )
    return ValidatedLeg(
        name=leg,
        raw_sha256=raw_sha256,
        report=report,
        sample_indices=list(elapsed_order),
        vectors=vectors,
        child_process_ids=child_ids,
    )


def _validate_cross_leg(legs: dict[str, ValidatedLeg]) -> None:
    a1 = legs["A1"].report
    for name in ("B1", "B2", "A2"):
        report = legs[name].report
        _expect(report["configuration"], a1["configuration"], f"{name}/A1 configuration identity")
        _expect(report["parallel_metrics"], a1["parallel_metrics"], f"{name}/A1 parallel identity")
        _expect(report["results"][0]["corpus"], a1["results"][0]["corpus"], f"{name}/A1 corpus identity")
    _expect(a1["binary_identity"], legs["A2"].report["binary_identity"], "control binary identity")
    _expect(legs["B1"].report["binary_identity"], legs["B2"].report["binary_identity"], "candidate binary identity")
    for metric in METRICS:
        for name in ("A1", "A2", "B1", "B2"):
            if len(set(legs[name].vectors[metric])) != 1:
                raise ValidationError(f"{name} {metric} vector is not constant")
        _expect(legs["A1"].vectors[metric], legs["A2"].vectors[metric], f"control {metric} vector equality")
        _expect(legs["B1"].vectors[metric], legs["B2"].vectors[metric], f"candidate {metric} vector equality")
    all_child_ids = [pid for leg in legs.values() for pid in leg.child_process_ids]
    if len(set(all_child_ids)) != 4 * SAMPLE_COUNT:
        raise ValidationError("fresh child process IDs are not unique across all four legs")


def _load_legs(
    a1_path: Path,
    b1_path: Path,
    b2_path: Path,
    a2_path: Path,
) -> dict[str, ValidatedLeg]:
    paths = {"A1": a1_path, "B1": b1_path, "B2": b2_path, "A2": a2_path}
    if len({path.resolve() for path in paths.values()}) != 4:
        raise ValidationError("the four ABBA roles must use four distinct report paths")
    legs = {name: _validate_report(path, name) for name, path in paths.items()}
    if len({leg.raw_sha256 for leg in legs.values()}) != 4:
        raise ValidationError("the four ABBA roles must use four distinct raw reports")
    _validate_cross_leg(legs)
    return legs


def _projection_leg(leg: ValidatedLeg) -> dict[str, Any]:
    allocation: dict[str, Any] = {}
    for metric in METRICS:
        vector = leg.vectors[metric]
        allocation[metric] = {
            "status": "measured",
            "scope": _projection_metric_scope(metric),
            "sample_count": SAMPLE_COUNT,
            "vector_sha256": _vector_digest(vector),
            "constant": len(set(vector)) == 1,
            "unique_values": sorted(set(vector)),
        }
    return {
        "raw_report_sha256": leg.raw_sha256,
        "aligned_sample_indices": {
            "sample_count": SAMPLE_COUNT,
            "vector_sha256": _vector_digest(leg.sample_indices),
            "permutation": True,
            "alignment": "elapsed_ns.samples_by_elapsed_then_sample_index",
        },
        "allocation": allocation,
    }


def _rounded_percentage(delta: int, control: int) -> tuple[float | None, float | None, str]:
    if control == 0:
        return None, None, "undefined_zero_control"
    percentage = round(100.0 * delta / control, 12)
    return percentage, -percentage, "defined"


def _projection_pair(control: ValidatedLeg, candidate: ValidatedLeg) -> dict[str, Any]:
    metrics: dict[str, Any] = {}
    for metric in METRICS:
        control_values = control.vectors[metric]
        candidate_values = candidate.vectors[metric]
        deltas = [right - left for left, right in zip(control_values, candidate_values, strict=True)]
        control_unique = sorted(set(control_values))
        candidate_unique = sorted(set(candidate_values))
        delta_unique = sorted(set(deltas))
        control_value = control_unique[0] if len(control_unique) == 1 else None
        candidate_value = candidate_unique[0] if len(candidate_unique) == 1 else None
        delta = delta_unique[0] if len(delta_unique) == 1 else None
        percentage_change: float | None = None
        reduction_percentage: float | None = None
        percentage_status = "not_constant"
        if control_value is not None and candidate_value is not None and delta is not None:
            percentage_change, reduction_percentage, percentage_status = _rounded_percentage(delta, control_value)
        metrics[metric] = {
            "control_unique_values": control_unique,
            "candidate_unique_values": candidate_unique,
            "candidate_minus_control_unique": delta_unique,
            "delta_vector_sha256": _vector_digest(deltas),
            "pairwise_sample_count": SAMPLE_COUNT,
            "pairwise_delta_constant": len(delta_unique) == 1,
            "control_value": control_value,
            "candidate_value": candidate_value,
            "delta": delta,
            "percentage_change": percentage_change,
            "reduction_percentage": reduction_percentage,
            "percentage_status": percentage_status,
        }
    return {
        "pairwise_sample_count": SAMPLE_COUNT,
        "pairwise_alignment": "same elapsed-order rank; sample_index is report-local",
        "percentage_formula": "100 * (candidate - control) / control; reduction_percentage = -percentage_change",
        "metrics": metrics,
    }


def _claimability() -> dict[str, Any]:
    result: dict[str, Any] = {}
    for metric in METRICS[:6]:
        result[metric] = {
            "status": "scoped_observation",
            "claimable": True,
            "scope": "matched candidate-minus-control operation-global system-allocator observation for the fixed warm filesystem selector and corpus",
        }
    result["allocator_elapsed_ns"] = {
        "status": "non_claimable",
        "claimable": False,
        "scope": "allocator-instrumented elapsed_ns",
        "reason": "allocator instrumentation is observational and is not an authorized elapsed-time acceptance metric",
    }
    for metric in METRICS[6:]:
        if metric.startswith("peak_live_bytes_"):
            scope = "absolute_process_high_water_snapshot"
        else:
            scope = "absolute_process_live_snapshot"
        result[metric] = {
            "status": "non_claimable",
            "claimable": False,
            "scope": scope,
            "reason": "cross-process global baseline or high-water snapshot is not an operation memory claim",
        }
    return result


def build_projection(
    a1_path: Path,
    b1_path: Path,
    b2_path: Path,
    a2_path: Path,
) -> dict[str, Any]:
    """Validate four reports and return a deterministic lossless projection."""

    legs = _load_legs(a1_path, b1_path, b2_path, a2_path)
    return _build_projection_from_legs(legs)


def _build_projection_from_legs(legs: dict[str, ValidatedLeg]) -> dict[str, Any]:
    """Build a projection from one already-validated report snapshot."""

    per_leg = {name: _projection_leg(legs[name]) for name in ("A1", "B1", "B2", "A2")}
    return {
        "schema_version": 1,
        "change": "0401",
        "scope": "Exact operation-scoped call/byte metrics plus absolute process live/high-water snapshots for allocator-enabled xlsx_file_selected_cell; per-sample vectors remain lossless in raw allocator reports.",
        "timing_status": "observational_only",
        "sample_count": SAMPLE_COUNT,
        "case": CASE,
        "cache_state": CACHE_STATE,
        "corpus_sha256": CORPUS_SHA256,
        "metrics": list(METRICS),
        "vector_digest_algorithm": "sha256(compact-json UTF-8 array)",
        "report_digest_algorithm": "sha256(raw report UTF-8 bytes)",
        "sample_index_alignment": "operation_metrics.sample_indices == elapsed_ns.sample_order; pairwise matching uses elapsed-order rank because sample_index is report-local.",
        "oracle_identity": {
            "source_sha256": CORPUS_SHA256,
            "semantic_sha256": SEMANTIC_SHA256,
            "selected_cell": SELECTED_CELL,
        },
        "protocol": {
            "abba_order": ["A1_control", "B1_candidate", "B2_candidate", "A2_control"],
            "samples_per_leg": SAMPLE_COUNT,
            "warmup_iterations_per_leg": WARMUP_ITERATIONS,
            "cache_state": CACHE_STATE,
            "fresh_child_per_sample": True,
            "process_isolated": True,
            "cpu_affinity": "2",
            "execution_workers": [1],
            "logical_cpus_available": 1,
            "filesystem_type": "tmpfs",
            "source_destination_same_device": True,
            "pairwise_alignment": "same elapsed-order rank; sample_index is report-local",
        },
        "provenance": {
            "rustc_version": RUSTC_VERSION,
            "allocator": ALLOCATOR,
            "profile": "release",
            "control": {
                "revision": LEGS["A1"]["revision"],
                "allocator_binary": {
                    "sha256": LEGS["A1"]["binary_sha256"],
                    "bytes": LEGS["A1"]["binary_bytes"],
                    "profile": "release",
                },
            },
            "candidate": {
                "revision": LEGS["B1"]["revision"],
                "allocator_binary": {
                    "sha256": LEGS["B1"]["binary_sha256"],
                    "bytes": LEGS["B1"]["binary_bytes"],
                    "profile": "release",
                },
            },
        },
        "per_leg": per_leg,
        "matched_deltas": {
            "A1_to_B1": _projection_pair(legs["A1"], legs["B1"]),
            "A2_to_B2": _projection_pair(legs["A2"], legs["B2"]),
        },
        "validation": {
            "raw_report_count": 4,
            "samples_per_leg": SAMPLE_COUNT,
            "total_sample_count": 4 * SAMPLE_COUNT,
            "unique_child_process_count": 4 * SAMPLE_COUNT,
            "fresh_child_process_ids_unique": True,
            "raw_to_aggregate_alignment_verified": True,
            "sample_index_permutation_verified": True,
            "allocator_vector_count": len(METRICS) * 4,
            "allocator_vector_status_scope_verified": True,
            "a1_equals_a2_allocation_vectors": True,
            "b1_equals_b2_allocation_vectors": True,
            "same_implementation_leg_equality_verified": True,
            "corpus_identity_verified": True,
            "selected_cell_oracle_identity_verified": True,
            "allocator_elapsed_summary_present": False,
            "allocator_elapsed_summary_claimable": False,
        },
        "claimability": _claimability(),
    }


def _reject_allocator_latency_statistics(value: Any, context: str = "projection") -> None:
    forbidden = {
        "elapsed_ns",
        "p50",
        "p95",
        "p99",
        "mean",
        "min",
        "max",
        "standard_deviation",
        "confidence_interval_95",
    }
    if isinstance(value, dict):
        overlap = forbidden & set(value)
        if overlap:
            raise ValidationError(
                f"{context} contains allocator latency statistic keys {sorted(overlap)}"
            )
        for key, child in value.items():
            _reject_allocator_latency_statistics(child, f"{context}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _reject_allocator_latency_statistics(child, f"{context}[{index}]")


def _validate_projection(
    path: Path,
    legs: dict[str, ValidatedLeg],
) -> None:
    projection, _ = _read_object(path)
    _reject_allocator_latency_statistics(projection)
    expected = _build_projection_from_legs(legs)
    _expect(projection, expected, "projection")
    # Keep this check explicit: projection equality above is intentionally
    # strict, while this assertion documents the non-claimable live/peak set.
    for metric in NONCLAIMABLE_METRICS:
        _expect(
            projection["claimability"][metric]["claimable"],
            False,
            f"projection.claimability.{metric}.claimable",
        )
    _expect(
        projection["claimability"]["allocator_elapsed_ns"]["claimable"],
        False,
        "projection.claimability.allocator_elapsed_ns.claimable",
    )
    # The parameter is deliberately consumed so callers cannot accidentally
    # validate a projection against a different set of reports.
    _expect(set(legs), set(LEGS), "projection validated leg set")


def validate_paths(
    a1_path: Path,
    b1_path: Path,
    b2_path: Path,
    a2_path: Path,
    projection_path: Path | None = None,
) -> dict[str, int]:
    """Validate the fixed Change 0401 ABBA reports and optional projection."""

    legs = _load_legs(a1_path, b1_path, b2_path, a2_path)
    if projection_path is not None:
        _validate_projection(projection_path, legs)
    return {
        "reports": 4,
        "samples": 4 * SAMPLE_COUNT,
        "unique_child_process_ids": len(
            {pid for leg in legs.values() for pid in leg.child_process_ids}
        ),
        "allocator_vectors": len(METRICS) * 4,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--a1", type=Path, required=True, help="control A1 allocator report")
    parser.add_argument("--b1", type=Path, required=True, help="candidate B1 allocator report")
    parser.add_argument("--b2", type=Path, required=True, help="candidate B2 allocator report")
    parser.add_argument("--a2", type=Path, required=True, help="control A2 allocator report")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--projection", type=Path, help="projection to validate")
    group.add_argument(
        "--write-projection",
        type=Path,
        help="write the deterministic projection to this path, then validate it",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.write_projection is not None:
            projection = build_projection(args.a1, args.b1, args.b2, args.a2)
            args.write_projection.write_text(
                json.dumps(projection, indent=2) + "\n", encoding="utf-8"
            )
            projection_path = args.write_projection
        else:
            projection_path = args.projection
        result = validate_paths(args.a1, args.b1, args.b2, args.a2, projection_path)
    except (OSError, ValidationError) as error:
        print(f"Change 0401 allocator ABBA validation failed: {error}", file=sys.stderr)
        return 1
    print(
        "validated Change 0401 allocator ABBA: "
        f"{result['reports']} reports, {result['samples']} samples, "
        f"{result['unique_child_process_ids']} unique fresh child processes, "
        f"{result['allocator_vectors']} exact allocator vectors; "
        "allocator elapsed/live/peak metrics are non-claimable"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
