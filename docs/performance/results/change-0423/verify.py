#!/usr/bin/env python3
"""Fail-closed single-report verifier for the 0423 PPTX lifecycle tranche.

The verifier accepts only the two established owned-source lifecycle selectors
and the two source-backed lifecycle selectors introduced by this tranche.  It
validates correctness and measurement boundaries for one report at a time;
none of the accepted evidence authorizes a performance claim.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import subprocess
import sys
from typing import Any, Mapping


EXPECTED_CHANGE = 423
NORMAL_BINARY = "litchi-perf-baseline"
ALLOCATOR_BINARY = "litchi-perf-baseline-alloc"
NORMAL_INSTRUMENTATION = "none"
ALLOCATOR_INSTRUMENTATION = "system_allocator_operation_scoped"
ALLOCATOR_COUNTER_REVISION = "serialized_region_peak_v3"
MAX_JSON_BYTES = 512 * 1024 * 1024
ZSTD_TIMEOUT_SECONDS = 120
SINK_MAX_WRITE = 65_536
SOURCE_MEDIA_ENTRY_COUNT = 8
SOURCE_MEDIA_ENTRY_BYTES = 2 * 1024 * 1024
SOURCE_MEDIA_TOTAL_BYTES = SOURCE_MEDIA_ENTRY_COUNT * SOURCE_MEDIA_ENTRY_BYTES
SHA256_RE = re.compile(r"^[0-9a-fA-F]{64}$")

# The 0423 source-backed lifecycle intentionally reuses the two fixed owned
# lifecycle fixtures.  Keep these identities here rather than accepting any
# self-consistent archive hash supplied by a report.
FIXED_CORPORA = {
    False: {
        "source_archive_sha256": "8ab4513642f48d99fa46e83b2c2764e8b20d417d3933ce3d5efe0825b21f918e",
        "destination_archive_sha256": "85fff18da5a020ed3f26897a58cdc839255ba769f89a3dea75c1de55e859ec6f",
        "entry_count": 23,
        "archive_member_count": 41,
        "entry_bytes": 76,
        "uncompressed_payload_bytes": 84715,
        "archive_bytes": 30539,
        "target_payload_bytes": 76,
        "target_payload_sha256": "9ac62cfc7da10f3c8e81c97cbb04d527499792f9eb629336bc11378788f719aa",
        "planned_part_count": 1,
        "planned_bytes": 1419,
        "external_relationship_count": 0,
        "collision_remapped_parts": 1,
        "matched_owned_output_bytes": 31545,
        "matched_owned_output_ceiling": 128626,
        "owned_output_sha256": "3e9ae2805a5dfd8e57ceff4a3662a2772f4197850252f7a23df8e6525d58ee22",
    },
    True: {
        "source_archive_sha256": "830928aafdc3ec8a5995a0d84a82ea9e2acb7f190f45008ddc017e2edfbf684b",
        "destination_archive_sha256": "a46fc227c453cbabbe54d6ca35fcad1cf1292c9a11e8e500bb1e4ec6a2708a4d",
        "entry_count": 31,
        "archive_member_count": 49,
        "entry_bytes": 76,
        "uncompressed_payload_bytes": 16866277,
        "archive_bytes": 16814664,
        "target_payload_bytes": 76,
        "target_payload_sha256": "9ac62cfc7da10f3c8e81c97cbb04d527499792f9eb629336bc11378788f719aa",
        "planned_part_count": 9,
        "planned_bytes": 16784397,
        "external_relationship_count": 0,
        "collision_remapped_parts": 9,
        "matched_owned_output_bytes": 33599873,
        "matched_owned_output_ceiling": 67265282,
        "owned_output_sha256": "6a3536fc7d9055eefad6eb115294f0157b7f8acd576a4898b4e402e48c3e7b4f",
    },
}

COMMON_CONFIGURATION = {
    "corpus_shapes": ["many-small"],
    "payload_kinds": ["compressible"],
    "writer_shapes": ["large"],
    "xlsx_shapes": ["medium"],
    "xlsx_cell_crud_shapes": ["medium"],
    "xlsx_row_visibility_shapes": ["medium"],
    "semantic_shapes": ["medium"],
    "execution_workers": [1],
    "filesystem_cache_states": ["warm"],
    "filesystem_fresh_child_per_sample": True,
    "filesystem_process_isolated": True,
    "filesystem_root_selected": False,
    "xlsb_shapes": ["tiny", "medium", "large", "sparse"],
    "rtf_variants": ["plain"],
    "range_simulation": {
        "fixed_latency_us": 100,
        "request_overhead_us": 25,
        "bandwidth_bytes_per_second": 52_428_800,
        "max_physical_range_bytes": 4096,
    },
    "opc_cache_lock_diagnostics": False,
}

OWNED_SELECTORS = frozenset(
    {
        "pptx_cross_copy_plain_lifecycle",
        "pptx_cross_copy_media_rich_lifecycle",
    }
)
SOURCE_SELECTORS = frozenset(
    {
        "pptx_source_backed_cross_copy_plain_lifecycle",
        "pptx_source_backed_cross_copy_media_rich_lifecycle",
    }
)
SELECTORS = OWNED_SELECTORS | SOURCE_SELECTORS

OWNED_GENERATOR = "litchi-pptx-cross-slide-copy-evidence-v1"
SOURCE_GENERATOR = "litchi-pptx-source-backed-cross-slide-copy-evidence-v1"
PPTX_PACKAGE_FORMAT = "PPTX/OPC/ZIP"
TARGET_ENTRY = "source-slide:2/destination-slide:1/position:1"
SOURCE_SLIDE = 2
DESTINATION_SLIDE = 1
INSERTION_POSITION = 1
DESTINATION_SLIDE_COUNT_BEFORE = 2
DESTINATION_SLIDE_COUNT_AFTER = 3
SOURCE_SLIDE_NAME = "Slide 258"
DESTINATION_SLIDE_NAME = "Slide 257"

REPORT_KEYS = {
    "schema_version",
    "tool",
    "binary_identity",
    "environment",
    "configuration",
    "parallel_metrics",
    "results",
    "corpus_catalog",
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
CORPUS_KEYS = {
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
OWNED_GATE_FIELDS = (
    "semantic_output_verified",
    "package_topology_verified",
    "dependency_closure_verified",
    "source_immutability_verified",
    "collision_remap_verified",
    "durable_patch_round_trip_verified",
    "borrowed_provenance_refusal_verified",
    "stale_source_refusal_verified",
    "stale_destination_refusal_verified",
    "foreign_source_refusal_verified",
)
SOURCE_GATE_FIELDS = (
    "matched_owned_corpus_verified",
    "semantic_output_verified",
    "package_topology_verified",
    "dependency_boundary_verified",
    "layout_reuse_verified",
    "untouched_destination_members_verified",
    "deterministic_output_verified",
    "source_version_stability_verified",
    "source_revision_refusal_verified",
    "destination_revision_refusal_verified",
    "foreign_destination_refusal_verified",
    "added_opc_parts_verified",
    "added_zip_members_verified",
    "media_leaf_payloads_verified",
    "media_leaf_content_types_verified",
    "media_relationships_verified",
)
OWNED_VECTOR_FIELDS = (
    "plan_ns",
    "commit_ns",
    "publication_ns",
    "reopen_ns",
    "output_sha256",
    "lifecycle_ns",
)
SOURCE_VECTOR_FIELDS = (
    "lifecycle_ns",
    "open_ns",
    "plan_ns",
    "publication_ns",
    "output_sha256",
    "source_read_calls",
    "source_read_bytes",
    "destination_read_calls",
    "destination_read_bytes",
)
OWNED_SUMMARY_KEYS = {
    "implementation",
    "timing_scope",
    "performance_claim",
    "source_archive_sha256",
    "destination_archive_sha256",
    "expected_output_sha256",
    "source_slide",
    "destination_slide",
    "insertion_position",
    "source_slide_name",
    "destination_slide_name",
    "destination_slide_count_before",
    "destination_slide_count_after",
    "planned_part_count",
    "planned_bytes",
    "external_relationship_count",
    "collision_remapped_parts",
    "gates",
    *OWNED_VECTOR_FIELDS,
}
SOURCE_SUMMARY_KEYS = {
    "implementation",
    "workload",
    "timing_scope",
    "performance_claim",
    "source_archive_sha256",
    "destination_archive_sha256",
    "expected_output_sha256",
    "matched_owned_output_bytes",
    "matched_owned_output_ceiling",
    "source_expected_output_bytes",
    "media_rich",
    "source_slide",
    "destination_slide",
    "insertion_position",
    "source_slide_name",
    "destination_slide_name",
    "destination_slide_count_before",
    "destination_slide_count_after",
    "matched_owned_planned_part_count",
    "matched_owned_planned_bytes",
    "matched_owned_external_relationship_count",
    "matched_owned_collision_remapped_parts",
    "added_opc_part_count",
    "added_archive_member_count",
    "added_slide_payload_bytes",
    "media_leaf_count",
    "media_leaf_bytes",
    "gates",
    *SOURCE_VECTOR_FIELDS,
}
SOURCE_COMMON_FIELDS = (
    "read_calls",
    "read_bytes",
    "ordinary_payload_read_calls",
    "ordinary_payload_read_bytes",
    "max_in_flight_reads",
)


class VerificationError(ValueError):
    """A report or its bound evidence violates the 0423 contract."""


def fail(path: str, message: str) -> None:
    raise VerificationError(f"{path}: {message}")


def object_value(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "must be an object")
    return value


def list_value(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "must be a list")
    return value


def string_value(value: Any, path: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        fail(path, "must be a non-empty string" if nonempty else "must be a string")
    return value


def integer_value(value: Any, path: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(path, f"must be an integer at least {minimum}")
    return value


def u64_value(value: Any, path: str, *, positive: bool = False) -> int:
    minimum = 1 if positive else 0
    value = integer_value(value, path, minimum=minimum)
    if value > (1 << 64) - 1:
        fail(path, "must fit in unsigned 64-bit range")
    return value


def sha256_value(value: Any, path: str) -> str:
    value = string_value(value, path)
    if SHA256_RE.fullmatch(value) is None:
        fail(path, "must be a 64-character hexadecimal SHA-256")
    return value.lower()


def canonical(value: Any, path: str = "value") -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError, OverflowError) as error:
        fail(path, f"is not canonical JSON: {error}")
    raise AssertionError("unreachable")


def reject_json_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON value {value!r}")


def reject_duplicate_json_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def load_json(path: Path, label: str) -> tuple[dict[str, Any], str]:
    try:
        if not path.is_file():
            fail(label, "file is missing")
        if path.stat().st_size > MAX_JSON_BYTES:
            fail(label, f"compressed/raw size exceeds {MAX_JSON_BYTES} bytes")
    except OSError as error:
        fail(label, f"cannot inspect file: {error}")
    try:
        if path.name.endswith(".zst"):
            process = subprocess.Popen(
                ["zstd", "-q", "-dc", str(path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            try:
                raw, stderr = process.communicate(timeout=ZSTD_TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                process.kill()
                process.communicate()
                fail(label, "zstd decompression timed out")
            if process.returncode != 0:
                detail = stderr.decode("utf-8", "replace").strip()
                fail(label, f"zstd decompression failed{': ' + detail if detail else ''}")
        else:
            raw = path.read_bytes()
    except FileNotFoundError:
        fail(label, "zstd executable is unavailable")
    except OSError as error:
        fail(label, f"cannot read file: {error}")
    if len(raw) > MAX_JSON_BYTES:
        fail(label, f"decoded JSON exceeds {MAX_JSON_BYTES} bytes")
    try:
        value = json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=reject_duplicate_json_keys,
            parse_constant=reject_json_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as error:
        fail(label, f"invalid JSON: {error}")
    return object_value(value, label), hashlib.sha256(raw).hexdigest()


def find_repo_root(explicit: Path | None) -> Path:
    if explicit is not None:
        root = explicit.expanduser().resolve()
        if not (root / "tools" / "perf_abba_summary.py").is_file():
            fail("--repo-root", "does not contain tools/perf_abba_summary.py")
        return root
    here = Path(__file__).resolve()
    for candidate in (here, *here.parents):
        if (candidate / "tools" / "perf_abba_summary.py").is_file():
            return candidate
    fail("repo", "cannot locate tools/perf_abba_summary.py; pass --repo-root")
    raise AssertionError("unreachable")


def import_validators(repo_root: Path) -> tuple[Any, Any, Any]:
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))
    try:
        from tools import perf_abba_summary, perf_compare, validate_perf_corpus_binding
    except ImportError as error:
        fail("repo", f"cannot import stable validators: {error}")
    return perf_abba_summary, perf_compare, validate_perf_corpus_binding


def contract_samples(lane: str, contract: str) -> tuple[int, int]:
    contracts = {
        ("normal", "formal"): (100, 10),
        ("normal", "functional"): (1, 0),
        ("allocator", "formal"): (30, 3),
        ("allocator", "functional"): (1, 0),
    }
    try:
        return contracts[(lane, contract)]
    except KeyError:
        fail("contract", f"unsupported lane/contract pair {lane!r}/{contract!r}")
        raise AssertionError("unreachable")


def expected_tool(lane: str) -> dict[str, str]:
    if lane == "normal":
        return {
            "name": "litchi-perf-baseline",
            "version": "0.1.0",
            "binary": NORMAL_BINARY,
            "profile": "release",
            "target_os": "linux",
            "target_arch": "x86_64",
            "instrumentation": NORMAL_INSTRUMENTATION,
        }
    return {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": ALLOCATOR_BINARY,
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": ALLOCATOR_INSTRUMENTATION,
        "allocator_counter_revision": ALLOCATOR_COUNTER_REVISION,
    }


def _nearest_rank(samples: list[int], percentile: int) -> int:
    index = ((percentile * len(samples) + 99) // 100) - 1
    return samples[min(index, len(samples) - 1)]


def validate_elapsed(
    value: Any,
    path: str,
    expected_samples: int,
    perf_abba_summary: Any,
) -> tuple[list[int], list[int]]:
    elapsed = object_value(value, path)
    if elapsed.get("unit") != "ns":
        fail(f"{path}.unit", "must be 'ns'")
    raw_samples = list_value(elapsed.get("samples"), f"{path}.samples")
    if len(raw_samples) != expected_samples:
        fail(f"{path}.samples", f"must contain exactly {expected_samples} samples")
    samples = [
        u64_value(item, f"{path}.samples[{index}]")
        for index, item in enumerate(raw_samples)
    ]
    if any(value == 0 for value in samples):
        fail(f"{path}.samples", "must contain positive durations")
    if samples != sorted(samples):
        fail(f"{path}.samples", "must be sorted ascending")

    raw_order = elapsed.get("sample_order")
    if raw_order is None:
        if expected_samples != 1:
            fail(f"{path}.sample_order", "is required when more than one sample is retained")
        order = [0]
    else:
        order_values = list_value(raw_order, f"{path}.sample_order")
        if len(order_values) != expected_samples:
            fail(f"{path}.sample_order", f"must contain exactly {expected_samples} entries")
        order = [
            u64_value(item, f"{path}.sample_order[{index}]")
            for index, item in enumerate(order_values)
        ]
        if sorted(order) != list(range(expected_samples)):
            fail(f"{path}.sample_order", "must be a complete sample permutation")

    if expected_samples >= 15:
        try:
            perf_abba_summary.recompute_statistics(elapsed, path)
        except Exception as error:
            fail(path, f"elapsed statistics rejected: {error}")
    else:
        expected = {
            "min": samples[0],
            "p50": samples[(len(samples) - 1) // 2] // 2
            + samples[len(samples) // 2] // 2
            + (samples[(len(samples) - 1) // 2] % 2 + samples[len(samples) // 2] % 2) // 2,
            "p95": _nearest_rank(samples, 95),
            "p99": _nearest_rank(samples, 99),
            "max": samples[-1],
        }
        for field, expected_value in expected.items():
            if u64_value(elapsed.get(field), f"{path}.{field}") != expected_value:
                fail(f"{path}.{field}", "reported statistic disagrees with samples")
        mean = elapsed.get("mean")
        standard_deviation = elapsed.get("standard_deviation")
        if (
            isinstance(mean, bool)
            or not isinstance(mean, (int, float))
            or not math.isfinite(float(mean))
            or float(mean) < 0
            or isinstance(standard_deviation, bool)
            or not isinstance(standard_deviation, (int, float))
            or not math.isfinite(float(standard_deviation))
            or float(standard_deviation) < 0
        ):
            fail(path, "mean and standard_deviation must be finite non-negative numbers")
        confidence = object_value(
            elapsed.get("confidence_interval_95"),
            f"{path}.confidence_interval_95",
        )
        if confidence.get("method") != "two-sided Student's t interval for the mean":
            fail(f"{path}.confidence_interval_95.method", "does not match the harness")
        for field in ("lower", "upper"):
            value = confidence.get(field)
            if (
                isinstance(value, bool)
                or not isinstance(value, (int, float))
                or not math.isfinite(float(value))
                or float(value) < 0
            ):
                fail(f"{path}.confidence_interval_95.{field}", "must be finite and non-negative")
        if float(confidence["lower"]) > float(confidence["upper"]):
            fail(f"{path}.confidence_interval_95", "lower must not exceed upper")
    return samples, order


def validate_operation(
    result: Mapping[str, Any],
    path: str,
    lane: str,
    samples: int,
    sample_order: list[int],
    perf_compare: Any,
) -> dict[str, Any]:
    operation = object_value(result.get("operation_metrics"), f"{path}.operation_metrics")
    elapsed = object_value(result.get("elapsed_ns"), f"{path}.elapsed_ns")
    try:
        perf_compare._validate_operation_metrics(
            operation,
            f"{path}.operation_metrics",
            elapsed["samples"],
            1,
            elapsed_sample_order=sample_order,
            tool_identity=expected_tool(lane),
        )
    except Exception as error:
        fail(path, f"operation metrics rejected: {error}")
    allocation = object_value(operation.get("allocation"), f"{path}.operation_metrics.allocation")
    expected_status = "measured" if lane == "allocator" else "unavailable"
    if allocation.get("status") != expected_status:
        fail(
            f"{path}.operation_metrics.allocation.status",
            f"must be {expected_status!r} for the {lane} lane",
        )
    if lane == "allocator":
        try:
            perf_compare._validate_allocator_operation_evidence(
                dict(result),
                path,
                samples,
                allocator_counter_revision=ALLOCATOR_COUNTER_REVISION,
            )
        except Exception as error:
            fail(path, f"allocator operation evidence rejected: {error}")
    return operation


def validate_corpus(
    corpus: dict[str, Any], summary: Mapping[str, Any], selector: str, path: str
) -> None:
    if set(corpus) != CORPUS_KEYS:
        fail(path, "corpus has an unexpected schema")
    media_rich = selector.endswith("media_rich_lifecycle")
    expected_shape = "media-rich" if media_rich else "plain"
    expected_payload = (
        "deterministic-incompressible-media-and-slide-text"
        if media_rich
        else "deterministic-slide-text"
    )
    expected_generator = SOURCE_GENERATOR if selector in SOURCE_SELECTORS else OWNED_GENERATOR
    for field, expected in (
        ("name", selector),
        ("generator", expected_generator),
        ("package_format", PPTX_PACKAGE_FORMAT),
        ("shape", expected_shape),
        ("payload_kind", expected_payload),
        ("compression", "deflate"),
        ("target_entry", TARGET_ENTRY),
    ):
        if corpus.get(field) != expected:
            fail(f"{path}.{field}", f"must be {expected!r}")
    for field in (
        "entry_count",
        "archive_member_count",
        "entry_bytes",
        "uncompressed_payload_bytes",
        "archive_bytes",
        "target_payload_bytes",
    ):
        u64_value(corpus.get(field), f"{path}.{field}")
    sha256_value(corpus.get("archive_sha256"), f"{path}.archive_sha256")
    sha256_value(corpus.get("target_payload_sha256"), f"{path}.target_payload_sha256")
    destination_digest = sha256_value(
        summary.get("destination_archive_sha256"),
        "summary.destination_archive_sha256",
    )
    if corpus["archive_sha256"].lower() != destination_digest:
        fail(path, "archive_sha256 does not match destination_archive_sha256")
    expected = FIXED_CORPORA[media_rich]
    for field in (
        "entry_count",
        "archive_member_count",
        "entry_bytes",
        "uncompressed_payload_bytes",
        "archive_bytes",
        "target_payload_bytes",
    ):
        if corpus[field] != expected[field]:
            fail(f"{path}.{field}", "does not match the fixed corpus identity")
    if corpus["archive_sha256"].lower() != expected["destination_archive_sha256"]:
        fail(f"{path}.archive_sha256", "does not match the fixed destination archive")
    if corpus["target_payload_sha256"].lower() != expected["target_payload_sha256"]:
        fail(f"{path}.target_payload_sha256", "does not match the fixed target payload")
    if corpus["xlsx"] is not None:
        fail(f"{path}.xlsx", "must be null for the PPTX corpus")


def validate_gates(gates: Any, fields: tuple[str, ...], path: str) -> None:
    gates = object_value(gates, path)
    if set(gates) != set(fields):
        fail(path, "gate schema does not match the selector contract")
    for field in fields:
        if gates[field] is not True:
            fail(f"{path}.{field}", "must be true")


def validate_fixed_summary_fields(
    summary: dict[str, Any],
    selector: str,
    path: str,
    *,
    plan_fields: tuple[str, str, str, str] = (
        "planned_part_count",
        "planned_bytes",
        "external_relationship_count",
        "collision_remapped_parts",
    ),
) -> None:
    media_rich = selector.endswith("media_rich_lifecycle")
    expected = FIXED_CORPORA[media_rich]
    for field in ("implementation", "timing_scope", "performance_claim", "source_slide_name", "destination_slide_name"):
        string_value(summary.get(field), f"{path}.{field}")
    if not summary["performance_claim"].startswith("none:"):
        fail(f"{path}.performance_claim", "must begin with 'none:'")
    for field, expected_value in (
        ("source_slide", SOURCE_SLIDE),
        ("destination_slide", DESTINATION_SLIDE),
        ("insertion_position", INSERTION_POSITION),
        ("destination_slide_count_before", DESTINATION_SLIDE_COUNT_BEFORE),
        ("destination_slide_count_after", DESTINATION_SLIDE_COUNT_AFTER),
    ):
        if u64_value(summary.get(field), f"{path}.{field}") != expected_value:
            fail(f"{path}.{field}", f"must be {expected_value}")
    if summary["source_slide_name"] != SOURCE_SLIDE_NAME:
        fail(f"{path}.source_slide_name", "does not match the fixed source slide")
    if summary["destination_slide_name"] != DESTINATION_SLIDE_NAME:
        fail(f"{path}.destination_slide_name", "does not match the fixed destination slide")
    source_digest = sha256_value(summary.get("source_archive_sha256"), f"{path}.source_archive_sha256")
    destination_digest = sha256_value(summary.get("destination_archive_sha256"), f"{path}.destination_archive_sha256")
    if source_digest == destination_digest:
        fail(path, "source and destination archives must differ")
    if source_digest != expected["source_archive_sha256"]:
        fail(f"{path}.source_archive_sha256", "does not match the fixed source archive")
    if destination_digest != expected["destination_archive_sha256"]:
        fail(f"{path}.destination_archive_sha256", "does not match the fixed destination archive")
    sha256_value(summary.get("expected_output_sha256"), f"{path}.expected_output_sha256")
    plan_part_field, plan_bytes_field, relationship_field, collision_field = plan_fields
    for field in plan_fields:
        u64_value(summary.get(field), f"{path}.{field}")
    if summary[plan_part_field] == 0:
        fail(f"{path}.{plan_part_field}", "must be positive")
    for field, expected_field in zip(
        plan_fields,
        (
            "planned_part_count",
            "planned_bytes",
            "external_relationship_count",
            "collision_remapped_parts",
        ),
    ):
        if summary[field] != expected[expected_field]:
            fail(f"{path}.{field}", "does not match the fixed copy plan")
    if summary[collision_field] > summary[plan_part_field]:
        fail(f"{path}.{collision_field}", f"must not exceed {plan_part_field}")


def validate_digest_vector(
    values: Any, path: str, expected: str, samples: int
) -> None:
    vector = list_value(values, path)
    if len(vector) != samples:
        fail(path, f"must contain exactly {samples} samples")
    for index, value in enumerate(vector):
        if sha256_value(value, f"{path}[{index}]") != expected:
            fail(f"{path}[{index}]", "does not match expected_output_sha256")


def validate_u64_vector(values: Any, path: str, samples: int, *, positive: bool = False) -> None:
    vector = list_value(values, path)
    if len(vector) != samples:
        fail(path, f"must contain exactly {samples} samples")
    for index, value in enumerate(vector):
        u64_value(value, f"{path}[{index}]", positive=positive)


def validate_owned_summary(
    summary: dict[str, Any], result: dict[str, Any], selector: str, samples: int, elapsed: list[int], path: str
) -> None:
    if set(summary) != OWNED_SUMMARY_KEYS:
        fail(path, "owned summary has an unexpected schema")
    validate_fixed_summary_fields(summary, selector, path)
    media_rich = selector.endswith("media_rich_lifecycle")
    expected = FIXED_CORPORA[media_rich]
    expected_parts = 9 if media_rich else 1
    expected_collisions = 9 if media_rich else 1
    if summary["planned_part_count"] != expected_parts:
        fail(f"{path}.planned_part_count", f"must be {expected_parts} for this corpus")
    if summary["collision_remapped_parts"] != expected_collisions:
        fail(f"{path}.collision_remapped_parts", f"must be {expected_collisions} for this corpus")
    validate_gates(summary.get("gates"), OWNED_GATE_FIELDS, f"{path}.gates")
    expected_output = sha256_value(summary["expected_output_sha256"], f"{path}.expected_output_sha256")
    if expected_output != expected["owned_output_sha256"]:
        fail(f"{path}.expected_output_sha256", "does not match the fixed owned output")
    if sha256_value(result.get("output_sha256"), "result.output_sha256") != expected_output:
        fail("result.output_sha256", "does not match expected_output_sha256")
    for field in OWNED_VECTOR_FIELDS[:-1]:
        if field == "output_sha256":
            validate_digest_vector(summary[field], f"{path}.{field}", expected_output, samples)
        else:
            validate_u64_vector(summary[field], f"{path}.{field}", samples)
    lifecycle = summary["lifecycle_ns"]
    validate_u64_vector(lifecycle, f"{path}.lifecycle_ns", samples, positive=True)
    if lifecycle != elapsed:
        fail(f"{path}.lifecycle_ns", "must equal elapsed_ns.samples")
    for index in range(samples):
        phase_total = sum(summary[field][index] for field in ("plan_ns", "commit_ns", "publication_ns"))
        if phase_total > lifecycle[index]:
            fail(path, f"phase total exceeds lifecycle_ns[{index}]")


def validate_source_summary(
    summary: dict[str, Any], result: dict[str, Any], selector: str, samples: int, elapsed: list[int], path: str
) -> None:
    if set(summary) != SOURCE_SUMMARY_KEYS:
        fail(path, "source-backed lifecycle summary has an unexpected schema")
    validate_fixed_summary_fields(
        summary,
        selector,
        path,
        plan_fields=(
            "matched_owned_planned_part_count",
            "matched_owned_planned_bytes",
            "matched_owned_external_relationship_count",
            "matched_owned_collision_remapped_parts",
        ),
    )
    media_rich = selector.endswith("media_rich_lifecycle")
    expected = FIXED_CORPORA[media_rich]
    workload = string_value(summary.get("workload"), f"{path}.workload")
    expected_workload = (
        "matched media-rich owned cross-copy corpus"
        if media_rich
        else "matched plain owned cross-copy corpus"
    )
    if workload != expected_workload:
        fail(f"{path}.workload", "does not identify the fixed matched owned corpus")
    if summary.get("media_rich") is not media_rich:
        fail(f"{path}.media_rich", "does not match the selected corpus")
    # These matched-owned plan counts/bytes are the closure oracle carried by
    # the source report, not a claim that the source implementation measured
    # an independent planner.  The source-specific gate set and read vectors
    # carry the source-backed observations.
    expected_parts = 9 if media_rich else 1
    expected_zip_members = 10 if media_rich else 2
    expected_media_count = SOURCE_MEDIA_ENTRY_COUNT if media_rich else 0
    expected_media_bytes = SOURCE_MEDIA_TOTAL_BYTES if media_rich else 0
    for field, expected_value in (
        ("matched_owned_planned_part_count", expected_parts),
        ("added_opc_part_count", expected_parts),
        ("added_archive_member_count", expected_zip_members),
        ("media_leaf_count", expected_media_count),
        ("media_leaf_bytes", expected_media_bytes),
    ):
        if u64_value(summary.get(field), f"{path}.{field}") != expected_value:
            fail(f"{path}.{field}", f"must be {expected_value} for this corpus")
    if summary["matched_owned_collision_remapped_parts"] != (9 if media_rich else 1):
        fail(
            f"{path}.matched_owned_collision_remapped_parts",
            "does not match the fixed closure",
        )
    u64_value(summary.get("added_slide_payload_bytes"), f"{path}.added_slide_payload_bytes", positive=True)
    owned_bytes = u64_value(summary.get("matched_owned_output_bytes"), f"{path}.matched_owned_output_bytes", positive=True)
    ceiling = u64_value(summary.get("matched_owned_output_ceiling"), f"{path}.matched_owned_output_ceiling", positive=True)
    source_bytes = u64_value(
        summary.get("source_expected_output_bytes"),
        f"{path}.source_expected_output_bytes",
        positive=True,
    )
    if owned_bytes != expected["matched_owned_output_bytes"]:
        fail(f"{path}.matched_owned_output_bytes", "does not match the fixed owned output")
    if ceiling != expected["matched_owned_output_ceiling"]:
        fail(f"{path}.matched_owned_output_ceiling", "does not match the fixed sink ceiling")
    expected_ceiling = owned_bytes * 2 + 64 * 1024
    if ceiling != expected_ceiling:
        fail(f"{path}.matched_owned_output_ceiling", "must be twice matched output plus 64 KiB")
    if source_bytes > ceiling:
        fail(f"{path}.source_expected_output_bytes", "exceeds matched owned-output ceiling")
    validate_gates(summary.get("gates"), SOURCE_GATE_FIELDS, f"{path}.gates")
    expected_output = sha256_value(summary["expected_output_sha256"], f"{path}.expected_output_sha256")
    if sha256_value(result.get("output_sha256"), "result.output_sha256") != expected_output:
        fail("result.output_sha256", "does not match expected_output_sha256")
    validate_digest_vector(summary["output_sha256"], f"{path}.output_sha256", expected_output, samples)
    for field in ("lifecycle_ns", "open_ns", "plan_ns", "publication_ns"):
        validate_u64_vector(summary[field], f"{path}.{field}", samples, positive=True)
    if summary["lifecycle_ns"] != elapsed:
        fail(f"{path}.lifecycle_ns", "must equal elapsed_ns.samples")
    for index in range(samples):
        phase_total = sum(
            summary[field][index] for field in ("open_ns", "plan_ns", "publication_ns")
        )
        if phase_total > elapsed[index]:
            fail(f"{path}.open_ns", f"open/plan/publication total exceeds lifecycle_ns[{index}]")
    for field in ("source_read_calls", "source_read_bytes", "destination_read_calls", "destination_read_bytes"):
        validate_u64_vector(summary[field], f"{path}.{field}", samples, positive=True)


def validate_sink(
    result: dict[str, Any], summary: Mapping[str, Any], selector: str, samples: int, path: str, perf_abba_summary: Any
) -> None:
    try:
        sink = perf_abba_summary._validate_pptx_cross_copy_sink(result.get("sink"), f"{path}.sink")
        perf_abba_summary._validate_pptx_cross_copy_operation_sink_binding(
            result, sink, path, samples
        )
    except Exception as error:
        fail(path, f"sink validation failed: {error}")
    if selector in SOURCE_SELECTORS:
        # The source-backed writer has its own ZIP serialization and may
        # legitimately produce a different byte count from the eager owned
        # control.  The control bytes define only the common reservation
        # ceiling; this report's sink/output digest is the source role's own
        # output identity.
        source_bytes = u64_value(
            summary.get("source_expected_output_bytes"),
            "summary.source_expected_output_bytes",
            positive=True,
        )
        if sink["accepted_bytes"] != source_bytes:
            fail(
                f"{path}.sink.accepted_bytes",
                "does not match source_expected_output_bytes",
            )
        if sink["accepted_bytes"] > summary["matched_owned_output_ceiling"]:
            fail(f"{path}.sink.accepted_bytes", "exceeds matched owned-output ceiling")
    else:
        media_rich = selector.endswith("media_rich_lifecycle")
        expected_bytes = FIXED_CORPORA[media_rich]["matched_owned_output_bytes"]
        if sink["accepted_bytes"] != expected_bytes:
            fail(f"{path}.sink.accepted_bytes", "does not match the fixed owned output")


def validate_report(
    report: dict[str, Any],
    catalog: dict[str, Any],
    *,
    selector: str,
    lane: str,
    samples: int,
    warmups: int,
    perf_abba_summary: Any,
    perf_compare: Any,
    corpus_binding: Any,
) -> dict[str, Any]:
    if selector not in SELECTORS:
        fail("selector", f"unsupported selector {selector!r}")
    if set(report) != REPORT_KEYS:
        fail("report", "top-level schema does not match the lifecycle contract")
    if report.get("schema_version") != 1:
        fail("report.schema_version", "must be 1")
    expected = expected_tool(lane)
    tool = object_value(report.get("tool"), "report.tool")
    if tool != expected:
        fail("report.tool", f"does not match the {lane} instrumentation contract")
    try:
        if perf_abba_summary.detect_report_profile(report, "report") != "current-v1":
            fail("report.tool", "legacy reports are not accepted")
        if lane == "normal":
            perf_abba_summary._validate_tool(tool, "report", "current-v1")
        perf_abba_summary._validate_binary_identity(report.get("binary_identity"), "report", tool)
        perf_abba_summary._validate_environment(report.get("environment"), "report")
        perf_compare.validate_parallel_metrics(report, "report")
    except VerificationError:
        raise
    except Exception as error:
        fail("report", f"shared identity/parallel validator rejected report: {error}")

    configuration = object_value(report.get("configuration"), "report.configuration")
    if integer_value(configuration.get("samples_per_case"), "report.configuration.samples_per_case", minimum=1) != samples:
        fail("report.configuration.samples_per_case", f"must be {samples}")
    if integer_value(configuration.get("warmup_iterations_per_case"), "report.configuration.warmup_iterations_per_case") != warmups:
        fail("report.configuration.warmup_iterations_per_case", f"must be {warmups}")
    if configuration.get("cases") != [selector]:
        fail("report.configuration.cases", "must contain exactly the selected case")
    expected_configuration = COMMON_CONFIGURATION | {
        "samples_per_case": samples,
        "warmup_iterations_per_case": warmups,
        "cases": [selector],
    }
    if canonical(configuration) != canonical(expected_configuration):
        fail("report.configuration", "must match the complete fixed experiment configuration")
    for field, expected_value in COMMON_CONFIGURATION.items():
        if configuration.get(field) != expected_value:
            fail(f"report.configuration.{field}", f"must match {expected_value!r}")
    for field, expected_value in (
        ("filesystem_cache_states", ["warm"]),
        ("filesystem_fresh_child_per_sample", True),
        ("filesystem_process_isolated", True),
        ("filesystem_root_selected", False),
        ("execution_workers", [1]),
    ):
        if field in configuration and configuration[field] != expected_value:
            fail(f"report.configuration.{field}", f"does not match {expected_value!r}")

    catalog_ref = object_value(report.get("corpus_catalog"), "report.corpus_catalog")
    expected_catalog_ref = {
        key: catalog.get(key)
        for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256")
    }
    if catalog_ref != expected_catalog_ref:
        fail("report.corpus_catalog", "does not match catalog sidecar")
    try:
        corpus_binding.validate_binding(report, catalog)
    except Exception as error:
        fail("report", f"report/catalog binding rejected: {error}")
    catalog_build = object_value(catalog.get("build"), "catalog.build")
    environment = object_value(report.get("environment"), "report.environment")
    if catalog_build.get("git_revision") != environment.get("git_revision"):
        fail("catalog.build.git_revision", "does not match report environment revision")

    try:
        if samples >= 15:
            perf_abba_summary._validate_configuration(configuration, "report")
        indexed = perf_abba_summary._index_results(report, "report")
        if selector in OWNED_SELECTORS:
            perf_abba_summary._validate_configuration_rows(configuration, indexed, "report")
            perf_abba_summary._validate_pptx_cross_copy_result_rows(indexed, configuration, "report")
    except Exception as error:
        fail("report", f"shared row validator rejected report: {error}")
    # Source lifecycle fixtures have fixed plain/media-rich shapes independent
    # of the generic CLI shape switches. This dedicated validator binds their
    # exact configuration, single selected row, sample cardinality and pinned
    # corpus below; the historical shared fixed-case registry has no source
    # lifecycle entry and cannot validate its shape-to-selector mapping.
    results = list_value(report.get("results"), "report.results")
    if len(results) != 1:
        fail("report.results", "must contain exactly one result")
    result = object_value(results[0], "report.results[0]")
    if set(result) != RESULT_KEYS:
        fail("report.results[0]", "result schema does not match the lifecycle contract")
    if result.get("case") != selector:
        fail("report.results[0].case", "does not match selected case")

    special_key = "pptx_cross_copy" if selector in OWNED_SELECTORS else "pptx_source_backed_cross_copy_lifecycle"
    source = object_value(result.get("source"), "report.results[0].source")
    expected_source_keys = set(SOURCE_COMMON_FIELDS) | {special_key}
    if set(source) != expected_source_keys:
        fail("report.results[0].source", "source schema does not match the lifecycle contract")
    for field in SOURCE_COMMON_FIELDS:
        values = list_value(source[field], f"report.results[0].source.{field}")
        if values and len(values) != samples:
            fail(f"report.results[0].source.{field}", "must be empty or match sample count")
        for index, value in enumerate(values):
            u64_value(value, f"report.results[0].source.{field}[{index}]")

    elapsed, sample_order = validate_elapsed(
        result.get("elapsed_ns"), "report.results[0].elapsed_ns", samples, perf_abba_summary
    )
    operation = validate_operation(
        result, "report.results[0]", lane, samples, sample_order, perf_compare
    )
    if operation.get("sample_indices") != sample_order:
        fail(
            "report.results[0].operation_metrics.sample_indices",
            "must match elapsed_ns.sample_order",
        )
    summary = object_value(source.get(special_key), f"report.results[0].source.{special_key}")
    if selector in OWNED_SELECTORS:
        validate_owned_summary(summary, result, selector, samples, elapsed, f"report.results[0].source.{special_key}")
    else:
        validate_source_summary(summary, result, selector, samples, elapsed, f"report.results[0].source.{special_key}")
    corpus = object_value(result.get("corpus"), "report.results[0].corpus")
    validate_corpus(corpus, summary, selector, "report.results[0].corpus")
    validate_sink(result, summary, selector, samples, "report.results[0]", perf_abba_summary)
    if operation.get("sample_count") != samples:
        fail("report.results[0].operation_metrics.sample_count", "does not match contract")
    return {
        "change": EXPECTED_CHANGE,
        "selector": selector,
        "lane": lane,
        "samples": samples,
        "warmups": warmups,
        "claim_authorized": False,
        "performance_claim": None,
        "report": report,
        "catalog": catalog,
    }


def resolve_contract(args: argparse.Namespace) -> tuple[int, int]:
    expected_samples, expected_warmups = contract_samples(args.lane, args.contract)
    samples = expected_samples if args.samples is None else args.samples
    warmups = expected_warmups if args.warmups is None else args.warmups
    if args.samples is None and args.warmups is not None:
        fail("contract", "--samples and --warmups must be supplied together")
    if args.warmups is None and args.samples is not None:
        fail("contract", "--samples and --warmups must be supplied together")
    if (samples, warmups) != (expected_samples, expected_warmups):
        fail("contract", f"{args.lane}/{args.contract} requires {expected_samples} samples and {expected_warmups} warmups")
    return samples, warmups


def run(args: argparse.Namespace) -> dict[str, Any]:
    repo_root = find_repo_root(args.repo_root)
    perf_abba_summary, perf_compare, corpus_binding = import_validators(repo_root)
    samples, warmups = resolve_contract(args)
    report, report_sha256 = load_json(args.report.expanduser(), "report")
    catalog, catalog_sha256 = load_json(args.catalog.expanduser(), "catalog")
    validate_report(
        report,
        catalog,
        selector=args.selector,
        lane=args.lane,
        samples=samples,
        warmups=warmups,
        perf_abba_summary=perf_abba_summary,
        perf_compare=perf_compare,
        corpus_binding=corpus_binding,
    )
    # Keep stdout as a custody proof envelope.  The validated report and
    # catalog remain on disk and are bound by these hashes; embedding either
    # full object here would duplicate the captured artifacts and make the
    # verifier output unnecessarily large.
    result = {
        "change": EXPECTED_CHANGE,
        "status": "pass",
        "claim_authorized": False,
        "performance_claim": None,
        "selector": args.selector,
        "lane": args.lane,
        "samples": samples,
        "warmups": warmups,
        "report_count": 1,
        "reports": [
            {
                "report_sha256": report_sha256,
                "catalog_sha256": catalog_sha256,
            }
        ],
    }
    if args.output is not None:
        output = args.output.expanduser()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, sort_keys=True, indent=2) + "\n", encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--catalog", type=Path, required=True)
    parser.add_argument("--selector", choices=sorted(SELECTORS), required=True)
    parser.add_argument("--lane", choices=("normal", "allocator"), default="normal")
    parser.add_argument("--contract", choices=("formal", "functional"), default="formal")
    parser.add_argument("--samples", type=int)
    parser.add_argument("--warmups", type=int)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        result = run(args)
    except (VerificationError, OSError) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 2
    print(json.dumps(result, sort_keys=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
