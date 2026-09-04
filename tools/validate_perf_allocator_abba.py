#!/usr/bin/env python3
"""Fail-closed validator for the Change 0400 allocator ABBA evidence.

The allocator reports retain elapsed measurements because the harness schema
requires them, but allocator instrumentation changes timing.  This validator
uses elapsed ordering only to align raw samples.  It deliberately neither
computes nor prints allocator latency statistics.

The validator is standard-library-only.  It checks the fixed Change 0400
capture identity, every raw operation-scoped allocator vector, fresh-child
process uniqueness, the XLSX source/semantic/selected-cell oracles, matched
A1/B1/B2/A2 provenance, and (when supplied) the complete allocation
projection derived from those reports.
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
SAMPLE_COUNT = 30
CASE = "xlsx_file_selected_cell"
CACHE_STATE = "warm"
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
CLAIMABLE_METRICS = set(METRICS[:6])
NONCLAIMABLE_METRICS = set(METRICS[6:])

LEGS = {
    "A1": {
        "implementation": "control",
        "revision": "2e47ccebf449ef88943c0abcecd32bd9141eb520",
        "binary_sha256": "7aa3aab90fea509e5806f081cb2f1247b348d9ec1b376fad9f56899355b3112c",
        "binary_bytes": 54_781_064,
    },
    "B1": {
        "implementation": "candidate",
        "revision": "f159c0aed603672aacee8e5923586ce4aa8753f7",
        "binary_sha256": "f69200b54d294dcc83af235e2ce5d1e8b2848ea63b9ab783408bec1092310482",
        "binary_bytes": 54_756_712,
    },
    "B2": {
        "implementation": "candidate",
        "revision": "f159c0aed603672aacee8e5923586ce4aa8753f7",
        "binary_sha256": "f69200b54d294dcc83af235e2ce5d1e8b2848ea63b9ab783408bec1092310482",
        "binary_bytes": 54_756_712,
    },
    "A2": {
        "implementation": "control",
        "revision": "2e47ccebf449ef88943c0abcecd32bd9141eb520",
        "binary_sha256": "7aa3aab90fea509e5806f081cb2f1247b348d9ec1b376fad9f56899355b3112c",
        "binary_bytes": 54_781_064,
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
RAW_ALLOCATION_KEYS = {"status", "scope", *METRICS}
OPERATION_ALLOCATION_KEYS = {"status", "scope", *METRICS}
VECTOR_KEYS = {"values", "status", "scope"}
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


class ValidationError(ValueError):
    """The evidence is malformed, incomplete, or internally inconsistent."""


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


def _validate_corpus(value: Any, context: str) -> dict[str, Any]:
    corpus = _object(value, context)
    required = {
        "name",
        "generator",
        "package_format",
        "shape",
        "payload_kind",
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
        "xlsx",
    }
    _exact_keys(corpus, required, context)
    expected_scalars = {
        "name": "xlsx-cell-values-medium",
        "generator": "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1",
        "package_format": "XLSX/OPC/ZIP",
        "shape": "medium",
        "payload_kind": "deterministic-multi-sheet-scalar-grid-with-media",
        "compression": "deflate",
        "entry_count": 9_216,
        "archive_member_count": 17,
        "entry_bytes": 4,
        "uncompressed_payload_bytes": 4_231_168,
        "archive_bytes": 4_226_429,
        "archive_sha256": CORPUS_SHA256,
        "target_entry": "Sheet1!A1",
        "target_payload_bytes": 1,
        "target_payload_sha256": "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
    }
    for key, expected in expected_scalars.items():
        _expect(corpus[key], expected, f"{context}.{key}")
    xlsx = _object(corpus["xlsx"], f"{context}.xlsx")
    expected_xlsx = {
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
    }
    _expect(xlsx, expected_xlsx, f"{context}.xlsx")
    return corpus


def _validate_tool(value: Any, context: str) -> None:
    tool = _exact_keys(value, TOOL_KEYS, context)
    expected = {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": "litchi-perf-baseline-alloc",
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": "system_allocator_operation_scoped",
    }
    _expect(tool, expected, context)


def _validate_binary(value: Any, leg: str, context: str) -> dict[str, Any]:
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
    return binary


def _validate_environment(value: Any, leg: str, context: str) -> dict[str, Any]:
    environment = _exact_keys(value, ENVIRONMENT_KEYS, context)
    fixed = {
        "rustc_version": RUSTC_VERSION,
        "git_revision": LEGS[leg]["revision"],
        "git_worktree_dirty": False,
        "logical_cpus_available": 1,
        "allocator": ALLOCATOR,
        "rustflags": None,
        "cargo_build_target": None,
        "perf_event_paranoid": "1",
        "os": "linux",
        "kernel": "Linux 7.0.0-1011-aws",
        "cpu_model": "AMD EPYC 9R45",
        "total_memory_bytes": 132_553_797_632,
        "page_size_bytes": 4_096,
        "filesystem_type": "tmpfs",
        "source_destination_same_device": True,
        "cpu_affinity": "2",
        "storage_identifier": None,
    }
    _expect(environment, fixed, context)
    return environment


def _validate_configuration(value: Any, context: str) -> dict[str, Any]:
    configuration = _exact_keys(value, CONFIGURATION_KEYS, context)
    fixed = {
        "samples_per_case": SAMPLE_COUNT,
        "warmup_iterations_per_case": 3,
        "filesystem_cache_states": [CACHE_STATE],
        "filesystem_fresh_child_per_sample": True,
        "filesystem_process_isolated": True,
        "filesystem_root_selected": True,
        "cases": [CASE],
        "execution_workers": [1],
    }
    for key, expected in fixed.items():
        _expect(configuration[key], expected, f"{context}.{key}")
    return configuration


def _validate_raw_allocation(value: Any, context: str) -> dict[str, int]:
    allocation = _exact_keys(value, RAW_ALLOCATION_KEYS, context)
    _expect(allocation["status"], "measured", f"{context}.status")
    _expect(allocation["scope"], ALLOCATION_SCOPE, f"{context}.scope")
    result: dict[str, int] = {}
    for metric in METRICS:
        result[metric] = _integer(allocation[metric], f"{context}.{metric}")
    return result


def _validate_sample(value: Any, leg: str, ordinal: int) -> tuple[int, int, int, dict[str, int]]:
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
    if not buckets or any(
        isinstance(value, bool) or not isinstance(value, int) or value != 0
        for value in buckets.values()
    ):
        raise ValidationError(f"{context}.logical_read_request_size_buckets must be all zero")
    process = _object(sample["process_metrics"], f"{context}.process_metrics")
    if not process:
        raise ValidationError(f"{context}.process_metrics must not be empty")
    for key, process_value in process.items():
        _integer(process_value, f"{context}.process_metrics.{key}")
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
    operation = _object(value, context)
    required = {
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
    _exact_keys(operation, required, context)
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
    allocation = _exact_keys(
        operation["allocation"], OPERATION_ALLOCATION_KEYS, f"{context}.allocation"
    )
    _expect(allocation["status"], "measured", f"{context}.allocation.status")
    _expect(allocation["scope"], ALLOCATION_SCOPE, f"{context}.allocation.scope")
    vectors: dict[str, list[int]] = {}
    for metric in METRICS:
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
        vectors[metric] = typed_values
    return vectors


def _validate_report(path: Path, leg: str) -> ValidatedLeg:
    report, raw_sha256 = _read_object(path)
    _exact_keys(report, REPORT_KEYS, leg)
    _expect(report["schema_version"], 1, f"{leg}.schema_version")
    _validate_tool(report["tool"], f"{leg}.tool")
    _validate_binary(report["binary_identity"], leg, f"{leg}.binary_identity")
    _validate_environment(report["environment"], leg, f"{leg}.environment")
    configuration = _validate_configuration(
        report["configuration"], f"{leg}.configuration"
    )
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
    envelope = _exact_keys(
        evidence[0], FILESYSTEM_ENVELOPE_KEYS, f"{leg}.filesystem_evidence[0]"
    )
    _expect(envelope["case"], CASE, f"{leg}.filesystem_evidence[0].case")
    _expect(envelope["corpus"], corpus, f"{leg}.filesystem_evidence[0].corpus")
    _expect(
        envelope["warmup_iterations"],
        3,
        f"{leg}.filesystem_evidence[0].warmup_iterations",
    )
    _expect(
        envelope["sample_count"],
        SAMPLE_COUNT,
        f"{leg}.filesystem_evidence[0].sample_count",
    )
    _expect(
        envelope["cache_states"],
        [CACHE_STATE],
        f"{leg}.filesystem_evidence[0].cache_states",
    )
    _expect(
        envelope["fresh_child_per_sample"],
        True,
        f"{leg}.filesystem_evidence[0].fresh_child_per_sample",
    )
    raw_samples = _array(envelope["samples"], f"{leg}.filesystem_evidence[0].samples")
    if len(raw_samples) != SAMPLE_COUNT:
        raise ValidationError(
            f"{leg}.filesystem_evidence[0].samples must contain "
            f"{SAMPLE_COUNT} samples"
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

    elapsed = _exact_keys(
        result["elapsed_ns"], ELAPSED_KEYS, f"{leg}.results[0].elapsed_ns"
    )
    _expect(elapsed["unit"], "ns", f"{leg}.results[0].elapsed_ns.unit")
    raw_elapsed_order = _array(
        elapsed["sample_order"], f"{leg}.results[0].elapsed_ns.sample_order"
    )
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
    elapsed_rank = [
        (samples_by_index[index][0], index) for index in elapsed_order
    ]
    if elapsed_rank != sorted(elapsed_rank):
        raise ValidationError(
            f"{leg} elapsed ties must be ordered by ascending sample_index"
        )

    vectors = _validate_operation_vectors(
        result["operation_metrics"], leg, elapsed_order, samples_by_index
    )
    _expect(
        configuration["samples_per_case"],
        len(elapsed_order),
        f"{leg}.configuration.samples_per_case",
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
        _expect(
            report["configuration"],
            a1["configuration"],
            f"{name}/A1 configuration identity",
        )
        _expect(
            report["parallel_metrics"],
            a1["parallel_metrics"],
            f"{name}/A1 parallel identity",
        )
        _expect(
            report["results"][0]["corpus"],
            a1["results"][0]["corpus"],
            f"{name}/A1 corpus identity",
        )
    _expect(
        a1["binary_identity"],
        legs["A2"].report["binary_identity"],
        "control binary identity",
    )
    _expect(
        legs["B1"].report["binary_identity"],
        legs["B2"].report["binary_identity"],
        "candidate binary identity",
    )
    for metric in METRICS:
        if len(set(legs["A1"].vectors[metric])) != 1:
            raise ValidationError(f"A1 {metric} vector is not constant")
        if len(set(legs["A2"].vectors[metric])) != 1:
            raise ValidationError(f"A2 {metric} vector is not constant")
        if len(set(legs["B1"].vectors[metric])) != 1:
            raise ValidationError(f"B1 {metric} vector is not constant")
        if len(set(legs["B2"].vectors[metric])) != 1:
            raise ValidationError(f"B2 {metric} vector is not constant")
        _expect(
            legs["A1"].vectors[metric],
            legs["A2"].vectors[metric],
            f"control {metric} vector equality",
        )
        _expect(
            legs["B1"].vectors[metric],
            legs["B2"].vectors[metric],
            f"candidate {metric} vector equality",
        )
    all_child_ids = [pid for leg in legs.values() for pid in leg.child_process_ids]
    if len(set(all_child_ids)) != 4 * SAMPLE_COUNT:
        raise ValidationError("fresh child process IDs are not unique across all four legs")


def _validate_projection_leg(value: Any, leg: ValidatedLeg, context: str) -> None:
    projection = _exact_keys(
        value,
        {"raw_report_sha256", "aligned_sample_indices", "allocation"},
        context,
    )
    _expect(projection["raw_report_sha256"], leg.raw_sha256, f"{context}.raw_report_sha256")
    indices = _exact_keys(
        projection["aligned_sample_indices"],
        {"sample_count", "vector_sha256", "permutation", "alignment"},
        f"{context}.aligned_sample_indices",
    )
    expected_indices = {
        "sample_count": SAMPLE_COUNT,
        "vector_sha256": _vector_digest(leg.sample_indices),
        "permutation": True,
        "alignment": "elapsed_ns.samples_by_elapsed_then_sample_index",
    }
    _expect(indices, expected_indices, f"{context}.aligned_sample_indices")
    allocations = _exact_keys(
        projection["allocation"], set(METRICS), f"{context}.allocation"
    )
    for metric in METRICS:
        vector = leg.vectors[metric]
        expected = {
            "status": "measured",
            "scope": _projection_metric_scope(metric),
            "sample_count": SAMPLE_COUNT,
            "vector_sha256": _vector_digest(vector),
            "constant": len(set(vector)) == 1,
            "unique_values": sorted(set(vector)),
        }
        _expect(allocations[metric], expected, f"{context}.allocation.{metric}")


def _rounded_percentage(delta: int, control: int) -> tuple[float | None, float | None, str]:
    if control == 0:
        return None, None, "undefined_zero_control"
    percentage = round(100.0 * delta / control, 12)
    return percentage, -percentage, "defined"


def _validate_projection_pair(
    value: Any,
    control: ValidatedLeg,
    candidate: ValidatedLeg,
    context: str,
) -> None:
    pair = _exact_keys(
        value,
        {"pairwise_sample_count", "pairwise_alignment", "percentage_formula", "metrics"},
        context,
    )
    _expect(pair["pairwise_sample_count"], SAMPLE_COUNT, f"{context}.pairwise_sample_count")
    _expect(
        pair["pairwise_alignment"],
        "same elapsed-order rank; sample_index is report-local",
        f"{context}.pairwise_alignment",
    )
    _expect(
        pair["percentage_formula"],
        "100 * (candidate - control) / control; reduction_percentage = -percentage_change",
        f"{context}.percentage_formula",
    )
    metrics = _exact_keys(pair["metrics"], set(METRICS), f"{context}.metrics")
    metric_keys = {
        "control_unique_values",
        "candidate_unique_values",
        "candidate_minus_control_unique",
        "delta_vector_sha256",
        "pairwise_sample_count",
        "pairwise_delta_constant",
        "control_value",
        "candidate_value",
        "delta",
        "percentage_change",
        "reduction_percentage",
        "percentage_status",
    }
    for metric in METRICS:
        control_values = control.vectors[metric]
        candidate_values = candidate.vectors[metric]
        deltas = [
            right - left
            for left, right in zip(control_values, candidate_values, strict=True)
        ]
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
            percentage_change, reduction_percentage, percentage_status = _rounded_percentage(
                delta, control_value
            )
        expected = {
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
        actual = _exact_keys(metrics[metric], metric_keys, f"{context}.metrics.{metric}")
        _expect(actual, expected, f"{context}.metrics.{metric}")


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


def _validate_projection(path: Path, legs: dict[str, ValidatedLeg]) -> None:
    projection, _ = _read_object(path)
    top_keys = {
        "schema_version",
        "change",
        "scope",
        "timing_status",
        "sample_count",
        "case",
        "cache_state",
        "corpus_sha256",
        "metrics",
        "vector_digest_algorithm",
        "report_digest_algorithm",
        "sample_index_alignment",
        "oracle_identity",
        "protocol",
        "provenance",
        "per_leg",
        "matched_deltas",
        "validation",
        "claimability",
    }
    _exact_keys(projection, top_keys, "projection")
    fixed = {
        "schema_version": 1,
        "change": "0400",
        "timing_status": "observational_only",
        "sample_count": SAMPLE_COUNT,
        "case": CASE,
        "cache_state": CACHE_STATE,
        "corpus_sha256": CORPUS_SHA256,
        "metrics": list(METRICS),
        "vector_digest_algorithm": "sha256(compact-json UTF-8 array)",
        "report_digest_algorithm": "sha256(raw report UTF-8 bytes)",
    }
    for key, expected in fixed.items():
        _expect(projection[key], expected, f"projection.{key}")
    _expect(
        projection["scope"],
        "Exact operation-scoped call/byte metrics plus absolute process "
        "live/high-water snapshots for allocator-enabled "
        "xlsx_file_selected_cell; per-sample vectors remain lossless in raw "
        "allocator reports.",
        "projection.scope",
    )
    _expect(
        projection["sample_index_alignment"],
        "operation_metrics.sample_indices == elapsed_ns.sample_order; pairwise "
        "matching uses elapsed-order rank because sample_index is report-local.",
        "projection.sample_index_alignment",
    )
    expected_oracle = {
        "source_sha256": CORPUS_SHA256,
        "semantic_sha256": SEMANTIC_SHA256,
        "selected_cell": SELECTED_CELL,
    }
    _expect(projection["oracle_identity"], expected_oracle, "projection.oracle_identity")
    protocol = _object(projection["protocol"], "projection.protocol")
    protocol_expected = {
        "abba_order": ["A1_control", "B1_candidate", "B2_candidate", "A2_control"],
        "samples_per_leg": SAMPLE_COUNT,
        "warmup_iterations_per_leg": 3,
        "cache_state": CACHE_STATE,
        "fresh_child_per_sample": True,
        "process_isolated": True,
        "cpu_affinity": "2",
        "execution_workers": [1],
        "logical_cpus_available": 1,
        "filesystem_type": "tmpfs",
        "source_destination_same_device": True,
        "pairwise_alignment": "same elapsed-order rank; sample_index is report-local",
    }
    _expect(protocol, protocol_expected, "projection.protocol")
    provenance = _exact_keys(
        projection["provenance"],
        {"rustc_version", "allocator", "profile", "control", "candidate"},
        "projection.provenance",
    )
    _expect(
        provenance.get("rustc_version"),
        RUSTC_VERSION,
        "projection.provenance.rustc_version",
    )
    _expect(
        provenance.get("allocator"),
        ALLOCATOR,
        "projection.provenance.allocator",
    )
    _expect(provenance.get("profile"), "release", "projection.provenance.profile")
    for implementation, leg_name in (("control", "A1"), ("candidate", "B1")):
        item = _object(
            provenance.get(implementation),
            f"projection.provenance.{implementation}",
        )
        _exact_keys(
            item,
            {"revision", "allocator_binary"},
            f"projection.provenance.{implementation}",
        )
        _expect(
            item.get("revision"),
            LEGS[leg_name]["revision"],
            f"projection.provenance.{implementation}.revision",
        )
        binary = _object(
            item.get("allocator_binary"),
            f"projection.provenance.{implementation}.allocator_binary",
        )
        _exact_keys(
            binary,
            {"sha256", "bytes", "profile"},
            f"projection.provenance.{implementation}.allocator_binary",
        )
        _expect(
            binary.get("sha256"),
            LEGS[leg_name]["binary_sha256"],
            f"projection.provenance.{implementation}.allocator_binary.sha256",
        )
        _expect(
            binary.get("bytes"),
            LEGS[leg_name]["binary_bytes"],
            f"projection.provenance.{implementation}.allocator_binary.bytes",
        )
        _expect(
            binary.get("profile"),
            "release",
            f"projection.provenance.{implementation}.allocator_binary.profile",
        )

    per_leg = _exact_keys(projection["per_leg"], set(LEGS), "projection.per_leg")
    for name, leg in legs.items():
        _validate_projection_leg(per_leg[name], leg, f"projection.per_leg.{name}")
    pairs = _exact_keys(
        projection["matched_deltas"], {"A1_to_B1", "A2_to_B2"}, "projection.matched_deltas"
    )
    _validate_projection_pair(
        pairs["A1_to_B1"],
        legs["A1"],
        legs["B1"],
        "projection.matched_deltas.A1_to_B1",
    )
    _validate_projection_pair(
        pairs["A2_to_B2"],
        legs["A2"],
        legs["B2"],
        "projection.matched_deltas.A2_to_B2",
    )

    validation = _object(projection["validation"], "projection.validation")
    expected_validation = {
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
    }
    _expect(validation, expected_validation, "projection.validation")
    claimability = _object(projection["claimability"], "projection.claimability")
    expected_claim_keys = set(METRICS) | {"allocator_elapsed_ns"}
    _exact_keys(claimability, expected_claim_keys, "projection.claimability")
    for metric in CLAIMABLE_METRICS:
        item = _exact_keys(
            claimability[metric],
            {"status", "claimable", "scope"},
            f"projection.claimability.{metric}",
        )
        _expect(
            item.get("status"),
            "scoped_observation",
            f"projection.claimability.{metric}.status",
        )
        _expect(
            item.get("claimable"),
            True,
            f"projection.claimability.{metric}.claimable",
        )
        _expect(
            item.get("scope"),
            "matched candidate-minus-control operation-global "
            "system-allocator observation for the fixed warm filesystem "
            "selector and corpus",
            f"projection.claimability.{metric}.scope",
        )
    for metric in NONCLAIMABLE_METRICS | {"allocator_elapsed_ns"}:
        item = _exact_keys(
            claimability[metric],
            {"status", "claimable", "scope", "reason"},
            f"projection.claimability.{metric}",
        )
        _expect(
            item.get("status"),
            "non_claimable",
            f"projection.claimability.{metric}.status",
        )
        _expect(
            item.get("claimable"),
            False,
            f"projection.claimability.{metric}.claimable",
        )
        if metric == "allocator_elapsed_ns":
            expected_scope = "allocator-instrumented elapsed_ns"
            expected_reason = (
                "allocator instrumentation is observational and is not an "
                "authorized elapsed-time acceptance metric"
            )
        elif metric.startswith("peak_live_bytes_"):
            expected_scope = "absolute_process_high_water_snapshot"
            expected_reason = (
                "cross-process global baseline or high-water snapshot is not "
                "an operation memory claim"
            )
        else:
            expected_scope = "absolute_process_live_snapshot"
            expected_reason = (
                "cross-process global baseline or high-water snapshot is not "
                "an operation memory claim"
            )
        _expect(
            item.get("scope"),
            expected_scope,
            f"projection.claimability.{metric}.scope",
        )
        _expect(
            item.get("reason"),
            expected_reason,
            f"projection.claimability.{metric}.reason",
        )
    _reject_allocator_latency_statistics(projection)


def validate_paths(
    a1_path: Path,
    b1_path: Path,
    b2_path: Path,
    a2_path: Path,
    projection_path: Path | None = None,
) -> dict[str, int]:
    paths = {"A1": a1_path, "B1": b1_path, "B2": b2_path, "A2": a2_path}
    legs = {name: _validate_report(path, name) for name, path in paths.items()}
    if len({path.resolve() for path in paths.values()}) != 4:
        raise ValidationError("the four ABBA roles must use four distinct report paths")
    if len({leg.raw_sha256 for leg in legs.values()}) != 4:
        raise ValidationError("the four ABBA roles must use four distinct raw reports")
    _validate_cross_leg(legs)
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
    parser.add_argument("--a1", type=Path, required=True, help="control A1 raw allocator report")
    parser.add_argument("--b1", type=Path, required=True, help="candidate B1 raw allocator report")
    parser.add_argument("--b2", type=Path, required=True, help="candidate B2 raw allocator report")
    parser.add_argument("--a2", type=Path, required=True, help="control A2 raw allocator report")
    parser.add_argument(
        "--projection",
        type=Path,
        help="optional allocation-metrics.json projection to validate",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        result = validate_paths(args.a1, args.b1, args.b2, args.a2, args.projection)
    except ValidationError as error:
        print(f"allocator ABBA validation failed: {error}", file=sys.stderr)
        return 1
    print(
        "validated allocator ABBA: "
        f"{result['reports']} reports, {result['samples']} samples, "
        f"{result['unique_child_process_ids']} unique fresh child processes, "
        f"{result['allocator_vectors']} exact allocator vectors; "
        "allocator elapsed time not evaluated"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
