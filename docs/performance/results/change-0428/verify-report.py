#!/usr/bin/env python3
"""Verify one portable 0428 PPTX cache and retention report.

The checker is independent of the Rust capture code. It validates complete
source-cache counters and gauges, explicit availability, managed-budget
phase state, exact output gates, and bounded refusal oracles. Consumed or
dropped cache owners must be unavailable; zero-filled diagnostics are never
accepted as a substitute.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Iterable

SCHEMA = "pptx_cache_retention_v1"
SCENARIOS = {
    "lifecycle",
    "exact-admission",
    "one-under",
    "pinned-eviction",
    "oversized-bypass",
    "repeated-publication",
}
CORPORA = {"plain", "media-rich"}
MAX_SAMPLES = 1_000
MAX_WARMUP = 1_000
U64_MAX = (1 << 64) - 1
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")

CACHE_COUNTER_FIELDS = (
    "hits", "cold_loads", "waiter_joins", "successful_loads", "failed_loads",
    "evictions", "bypasses", "oversized_bypasses", "allocation_bypasses",
    "budget_reservation_failures",
)
CACHE_GAUGE_FIELDS = (
    "retained_entries", "retained_bytes", "in_flight_loads", "budget_managed",
    "budget_memory_used", "budget_cache_reserved_bytes", "budget_memory_limit",
    "budget_input_bytes_used", "budget_input_bytes_limit",
    "budget_output_bytes_used", "budget_output_bytes_limit",
    "budget_work_used", "budget_work_limit", "budget_objects_used",
    "budget_objects_limit", "budget_catalog_reserved_objects",
    "budget_cache_reserved_objects",
)
CACHE_FIELDS = CACHE_COUNTER_FIELDS + CACHE_GAUGE_FIELDS
BUDGET_POINT_FIELDS = (
    "memory_used", "memory_limit", "input_bytes_used", "input_bytes_limit",
    "output_bytes_used", "output_bytes_limit", "work_used", "work_limit",
    "objects_used", "objects_limit", "depth_used", "depth_limit",
)
BUDGET_FIELDS = (
    "memory_used", "memory_limit", "objects_used", "objects_limit",
    "input_bytes_used", "input_bytes_limit", "output_bytes_used",
    "output_bytes_limit", "work_used", "work_limit", "depth_used", "depth_limit",
)
CUMULATIVE_BUDGET_FIELDS = ("input_bytes_used", "output_bytes_used", "work_used")
SOURCE_CACHE_PHASES = (
    "baseline", "opened", "planned", "published", "drop_result", "drop_plan",
    "drop_view", "drop_caller_sources", "drop_sink",
)
IMAGE_PHASES = ("baseline", "metadata", "payload", "drop_image", "drop_view", "drop_source")
OVERSIZED_PHASES = (
    "baseline", "metadata", "payload", "drop_image", "oversized_reload", "drop_view", "drop_source",
)
PINNED_PHASES = (
    "baseline", "metadata", "pinned_a", "pinned_b", "drop_image", "reload_b",
    "drop_view", "drop_source",
)
REPEATED_PHASES = (
    "baseline", "source_opened",
    "opened_1", "planned_1", "published_1", "drop_result_1", "drop_plan_1", "drop_sink_1",
    "opened_2", "planned_2", "published_2", "drop_result_2", "drop_plan_2", "drop_sink_2",
    "opened_3", "planned_3", "published_3", "drop_result_3", "drop_plan_3", "drop_sink_3",
    "fourth_refusal", "drop_view", "drop_sources",
)
NEAR_DESCRIPTION_PHASES = (
    "baseline", "metadata", "payload", "drop_image", "drop_view",
    "opened_1", "planned_1", "published_1", "drop_plan_1", "drop_sink_1",
)
GATE_FIELDS = (
    "matched_owned_corpus_verified", "semantic_output_verified", "package_topology_verified",
    "dependency_boundary_verified", "layout_reuse_verified", "untouched_destination_members_verified",
    "deterministic_output_verified", "source_version_stability_verified",
    "source_revision_refusal_verified", "destination_revision_refusal_verified",
    "foreign_destination_refusal_verified", "added_opc_parts_verified", "added_zip_members_verified",
    "media_leaf_payloads_verified", "media_leaf_content_types_verified", "media_relationships_verified",
)
PHASE_RECORD_FIELDS = (
    "label", "source_cache", "destination_cache", "source_reads", "destination_reads",
    "source_budget", "destination_budget", "rss",
)
CACHE_POINT_FIELDS = (
    "availability", "unavailable_reason", "interval_label", "counter_delta_checked", "event_delta",
) + CACHE_FIELDS
READ_POINT_FIELDS = ("availability", "unavailable_reason", "read_calls", "read_bytes")
RSS_POINT_FIELDS = ("availability", "unavailable_reason", "rss_bytes", "vm_hwm_bytes")
PHASE_DESCRIPTION_FIELDS = ("label", "live_owners")
CONFIGURED_LIMIT_FIELDS = (
    "cache_max_bytes", "cache_max_entries", "memory_limit", "input_bytes_limit",
    "output_bytes_limit", "work_limit", "objects_limit", "depth_limit",
)
LIFECYCLE_ROW_FIELDS = ("sample_index", "exact_output_verified", "output_sha256", "output_bytes", "phases")
NEAR_ROW_FIELDS = (
    "sample_index", "result", "phases", "exact_payload_verified", "payload_sha256", "payload_bytes",
    "payload_read_calls", "payload_read_bytes", "accepted_output_bytes", "refusal_accepted_output_bytes",
    "reload_read_calls", "reload_read_bytes", "refusal_sink_accepted_bytes", "refusal_resource",
    "typed_memory_refusal", "typed_resource_refusal", "eviction_verified",
    "pinned_bypass_verified", "oversized_bypass_verified", "repeated_output_identities_verified",
    "publication_output_sha256", "root_memory_floor", "metadata_memory_used", "payload_memory_used",
    "drop_image_memory_used", "leaf_bytes", "memory_limit", "cache_max_bytes", "cache_max_entries",
    "configured_limits", "destination_configured_limits",
)
PHASE_OWNERS = ("source_cache", "destination_cache")
FORBIDDEN_CLAIM_KEYS = {
    "object_owned_bytes", "owned_bytes", "cache_owned_bytes",
    "attributed_cache_bytes", "managed_budget_release", "cache_release_bytes",
    "leak", "leak_bytes", "cache_leak", "cache_leak_bytes", "rss_release",
    "rss_released_bytes",
}


class VerificationError(ValueError):
    """The JSON document violates the 0428 evidence contract."""


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
        fail(path, "expected a non-empty string")
    return value


def boolean(value: Any, path: str) -> bool:
    if not isinstance(value, bool):
        fail(path, "expected a boolean")
    return value


def uint(value: Any, path: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0 or value > U64_MAX:
        fail(path, "expected a u64 integer")
    return value


def optional_uint(value: Any, path: str) -> int | None:
    if value is None:
        return None
    return uint(value, path)


def digest(value: Any, path: str) -> str:
    value = text(value, path)
    if not HEX64.fullmatch(value):
        fail(path, "expected a SHA-256 hexadecimal digest")
    return value.lower()


def exact_keys(
    value: dict[str, Any],
    required: Iterable[str],
    optional: Iterable[str],
    path: str,
) -> None:
    required = tuple(required)
    allowed = set(required) | set(optional)
    missing = [name for name in required if name not in value]
    if missing:
        fail(path, f"missing fields: {', '.join(missing)}")
    unknown = sorted(set(value) - allowed)
    if unknown:
        fail(path, f"unknown fields: {', '.join(unknown)}")


def reject_unsupported_claims(value: Any, path: str) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            lowered = key.lower()
            if lowered in FORBIDDEN_CLAIM_KEYS:
                fail(f"{path}.{key}", "unsupported ownership, leak, or RSS-release claim")
            if any(token in lowered for token in (
                "attributed", "owned_memory", "cache_leak", "leak_bytes"
            )):
                fail(f"{path}.{key}", "unsupported ownership or leak attribution")
            reject_unsupported_claims(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            reject_unsupported_claims(child, f"{path}[{index}]")


def reject_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON object key: {key}")
        result[key] = value
    return result


def reject_nonfinite_constant(value: str) -> Any:
    raise VerificationError(f"non-finite JSON constant is not permitted: {value}")


def check_all_numbers(value: Any, path: str) -> None:
    if isinstance(value, bool):
        return
    if isinstance(value, int):
        uint(value, path)
        return
    if isinstance(value, float):
        fail(path, "floating-point values are not permitted; expected u64 or boolean")
    if isinstance(value, dict):
        for key, child in value.items():
            check_all_numbers(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            check_all_numbers(child, f"{path}[{index}]")


def check_hash_fields(value: Any, path: str) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            lowered = key.lower()
            if lowered == "publication_output_sha256":
                hashes = array(child, f"{path}.{key}")
                for index, item in enumerate(hashes):
                    digest(item, f"{path}.{key}[{index}]")
            elif lowered.endswith("_sha256") or lowered in {"sha256", "hash"}:
                # Optional near-row/image hashes are serialized as null for
                # refusal and publication-only rows.  Every present digest
                # remains a strict SHA-256 string.
                if child is not None:
                    digest(child, f"{path}.{key}")
            check_hash_fields(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            check_hash_fields(child, f"{path}[{index}]")


def check_corpus_manifest(report: dict[str, Any]) -> None:
    manifest = obj(report["corpus_manifest"], "report.corpus_manifest")
    if not manifest:
        fail("report.corpus_manifest", "manifest cannot be empty")
    for key, child in manifest.items():
        if key.lower().endswith("_sha256"):
            digest(child, f"report.corpus_manifest.{key}")
        elif isinstance(child, str):
            text(child, f"report.corpus_manifest.{key}", allow_empty=True)
        elif isinstance(child, list):
            for index, item in enumerate(child):
                if isinstance(item, str):
                    text(item, f"report.corpus_manifest.{key}[{index}]", allow_empty=True)


def check_limits(value: Any, path: str) -> dict[str, int]:
    limits = obj(value, path)
    exact_keys(limits, CONFIGURED_LIMIT_FIELDS, (), path)
    result = {field: uint(limits[field], f"{path}.{field}") for field in CONFIGURED_LIMIT_FIELDS}
    for field in CONFIGURED_LIMIT_FIELDS:
        if result[field] == 0:
            fail(f"{path}.{field}", "configured limit must be positive")
    return result


def check_budget_point(value: Any, path: str) -> dict[str, int]:
    point = obj(value, path)
    exact_keys(point, BUDGET_POINT_FIELDS, (), path)
    result = {field: uint(point[field], f"{path}.{field}") for field in BUDGET_POINT_FIELDS}
    for field in ("memory_limit", "input_bytes_limit", "output_bytes_limit", "work_limit",
                  "objects_limit", "depth_limit"):
        if result[field] == 0:
            fail(f"{path}.{field}", "budget limit must be positive")
    for used, limit in (
        ("memory_used", "memory_limit"), ("input_bytes_used", "input_bytes_limit"),
        ("output_bytes_used", "output_bytes_limit"), ("work_used", "work_limit"),
        ("objects_used", "objects_limit"), ("depth_used", "depth_limit"),
    ):
        if result[used] > result[limit]:
            fail(f"{path}.{used}", "budget usage exceeds its configured limit")
    return result


def check_scope(value: Any, path: str) -> None:
    scope = text(value, path).lower()
    if scope in {"wrong scope", "wrong_scope", "mutated", "unknown"}:
        fail(path, "invalid diagnostic scope")
    if not any(token in scope for token in (
        "cache", "budget", "source", "process", "operation", "rss"
    )):
        fail(path, "scope does not identify a supported evidence boundary")


def check_cache_diagnostics(value: Any, path: str) -> dict[str, Any] | None:
    raw = obj(value, path)
    exact_keys(raw, CACHE_POINT_FIELDS, (), path)
    availability = text(raw["availability"], f"{path}.availability").lower()
    if availability not in {"available", "unavailable"}:
        fail(f"{path}.availability", "unknown cache availability state")
    if availability == "unavailable":
        text(raw["unavailable_reason"], f"{path}.unavailable_reason")
        for field in CACHE_FIELDS:
            if raw[field] is not None:
                fail(f"{path}.{field}", "unavailable cache owner cannot contain fabricated diagnostics")
        for field in ("interval_label", "counter_delta_checked", "event_delta"):
            if raw[field] is not None:
                fail(f"{path}.{field}", "unavailable cache owner cannot contain interval evidence")
        return None

    if raw["unavailable_reason"] is not None:
        fail(f"{path}.unavailable_reason", "available cache owner cannot have an unavailable reason")
    for field in CACHE_COUNTER_FIELDS:
        uint(raw[field], f"{path}.{field}")
    boolean(raw["budget_managed"], f"{path}.budget_managed")
    for field in CACHE_GAUGE_FIELDS:
        if field == "budget_managed":
            continue
        uint(raw[field], f"{path}.{field}")
    if raw["counter_delta_checked"] is not None:
        boolean(raw["counter_delta_checked"], f"{path}.counter_delta_checked")
    if raw["event_delta"] is None:
        if raw["counter_delta_checked"] is not None:
            fail(path, "counter-delta marker requires an event delta")
    else:
        event = obj(raw["event_delta"], f"{path}.event_delta")
        exact_keys(event, CACHE_COUNTER_FIELDS, (), f"{path}.event_delta")
        for field in CACHE_COUNTER_FIELDS:
            uint(event[field], f"{path}.event_delta.{field}")
        if raw["counter_delta_checked"] is not True:
            fail(f"{path}.counter_delta_checked", "event delta was not marked checked")
        if raw["interval_label"] is None:
            fail(f"{path}.interval_label", "checked event delta omitted its interval label")
    if raw["interval_label"] is not None:
        text(raw["interval_label"], f"{path}.interval_label")
    return raw


def check_phase_descriptions(value: Any, path: str) -> tuple[str, ...]:
    phases = array(value, path)
    if not phases:
        fail(path, "phase descriptions cannot be empty")
    labels: list[str] = []
    for index, raw in enumerate(phases):
        phase = obj(raw, f"{path}[{index}]")
        exact_keys(phase, PHASE_DESCRIPTION_FIELDS, (), f"{path}[{index}]")
        label = text(phase["label"], f"{path}[{index}].label")
        if label in labels:
            fail(f"{path}[{index}].label", "duplicate phase label")
        labels.append(label)
        owners = text(phase["live_owners"], f"{path}[{index}].live_owners")
        if "owned bytes" in owners.lower() or "cache bytes" in owners.lower():
            fail(f"{path}[{index}].live_owners", "phase description claims byte ownership")
    return tuple(labels)


def check_read_point(value: Any, path: str) -> None:
    read = obj(value, path)
    exact_keys(read, READ_POINT_FIELDS, (), path)
    state = text(read["availability"], f"{path}.availability").lower()
    if state not in {"available", "unavailable"}:
        fail(f"{path}.availability", "unknown source-read availability")
    if state == "available":
        if read["unavailable_reason"] is not None:
            fail(f"{path}.unavailable_reason", "available source-read point has an unavailable reason")
        uint(read["read_calls"], f"{path}.read_calls")
        uint(read["read_bytes"], f"{path}.read_bytes")
    else:
        text(read["unavailable_reason"], f"{path}.unavailable_reason")
        if read["read_calls"] is not None or read["read_bytes"] is not None:
            fail(path, "unavailable source-read point cannot fabricate counters")


def check_rss_point(value: Any, path: str) -> None:
    rss = obj(value, path)
    exact_keys(rss, RSS_POINT_FIELDS, (), path)
    state = text(rss["availability"], f"{path}.availability").lower()
    if state not in {"available", "unavailable"}:
        fail(f"{path}.availability", "invalid RSS point status")
    if state == "available":
        if rss["unavailable_reason"] is not None:
            fail(f"{path}.unavailable_reason", "available RSS point has an unavailable reason")
        uint(rss["rss_bytes"], f"{path}.rss_bytes")
        uint(rss["vm_hwm_bytes"], f"{path}.vm_hwm_bytes")
    else:
        text(rss["unavailable_reason"], f"{path}.unavailable_reason")
        if rss["rss_bytes"] is not None or rss["vm_hwm_bytes"] is not None:
            fail(path, "unavailable RSS point cannot fabricate process bytes")


def check_gates(value: Any, path: str, corpus: str) -> None:
    gates = obj(value, path)
    exact_keys(gates, GATE_FIELDS, (), path)
    for key in GATE_FIELDS:
        if not boolean(gates[key], f"{path}.{key}"):
            fail(f"{path}.{key}", "correctness gate is false")


def expected_row_phase_labels(scenario: str) -> tuple[str, ...]:
    if scenario == "lifecycle":
        return SOURCE_CACHE_PHASES
    if scenario in {"exact-admission", "one-under"}:
        return IMAGE_PHASES
    if scenario == "oversized-bypass":
        return OVERSIZED_PHASES
    if scenario == "pinned-eviction":
        return PINNED_PHASES
    return REPEATED_PHASES


def check_phase_records(
    row: dict[str, Any],
    path: str,
    scenario: str,
    limits: dict[str, int],
    destination_limits: dict[str, int],
) -> tuple[int, list[dict[str, int]]]:
    phases = array(row.get("phases"), f"{path}.phases")
    expected_labels = expected_row_phase_labels(scenario)
    if tuple(
        text(obj(raw, f"{path}.phases[{index}]").get("label"), f"{path}.phases[{index}].label")
        for index, raw in enumerate(phases)
    ) != expected_labels:
        fail(f"{path}.phases", f"phase labels do not match the {scenario} journal")

    owner_points: dict[str, list[tuple[str, dict[str, Any] | None, dict[str, Any]]]] = {
        owner: [] for owner in PHASE_OWNERS
    }
    read_points: dict[str, list[tuple[str, dict[str, Any]]]] = {
        "source_reads": [], "destination_reads": [],
    }
    budgets: dict[str, list[tuple[str, dict[str, int]]]] = {
        "source_budget": [], "destination_budget": [],
    }
    counter_delta_count = 0
    for index, raw in enumerate(phases):
        phase = obj(raw, f"{path}.phases[{index}]")
        exact_keys(phase, PHASE_RECORD_FIELDS, (), f"{path}.phases[{index}]")
        label = text(phase["label"], f"{path}.phases[{index}].label")
        source_live = label not in {
            "baseline", "drop_view", "drop_source", "drop_sources", "drop_caller_sources", "drop_sink",
        }
        destination_live = label in {"opened", "planned"} or label.startswith(("opened_", "planned_"))
        for owner in PHASE_OWNERS:
            cache_path = f"{path}.phases[{index}].{owner}"
            snapshot = check_cache_diagnostics(phase[owner], cache_path)
            expected_live = source_live if owner == "source_cache" else destination_live
            if (snapshot is not None) != expected_live:
                fail(cache_path, "cache availability differs from the prescribed owner lifetime")
            owner_points[owner].append((label, snapshot, phase[owner]))
            if phase[owner]["event_delta"] is not None:
                counter_delta_count += 1
        for name in read_points:
            read_path = f"{path}.phases[{index}].{name}"
            check_read_point(phase[name], read_path)
            expected_live = label not in {
                "baseline", "drop_source", "drop_sources", "drop_caller_sources", "drop_sink",
            }
            if name == "destination_reads" and scenario not in {"lifecycle", "repeated-publication"}:
                expected_live = False
            if (phase[name]["availability"] == "available") != expected_live:
                fail(read_path, "source-read availability differs from the prescribed caller lifetime")
            read_points[name].append((label, phase[name]))
        for name in budgets:
            budget_path = f"{path}.phases[{index}].{name}"
            budgets[name].append((label, check_budget_point(phase[name], budget_path)))
        for owner in PHASE_OWNERS:
            if phase[owner]["availability"] == "available":
                budget = phase[owner.replace("_cache", "_budget")]
                for resource in ("memory", "objects", "input_bytes", "output_bytes", "work"):
                    if phase[owner][f"budget_{resource}_used"] != budget[f"{resource}_used"]:
                        fail(f"{path}.phases[{index}].{owner}", "cache usage differs from its caller budget")
        check_rss_point(phase["rss"], f"{path}.phases[{index}].rss")

    for owner, points in owner_points.items():
        previous: dict[str, Any] | None = None
        for label, current, raw_cache in points:
            if current is None:
                previous = None
                continue
            if current["budget_managed"] is not True:
                fail(f"{path}.phases.{owner}.{label}.budget_managed", "managed cache evidence must be true")
            if current["in_flight_loads"] != 0:
                fail(f"{path}.phases.{owner}.{label}.in_flight_loads", "completed phase retained an in-flight load")
            expected_limits = limits if owner == "source_cache" else destination_limits
            for field, limit_name in (
                ("retained_entries", "cache_max_entries"),
                ("retained_bytes", "cache_max_bytes"),
                ("budget_cache_reserved_bytes", "cache_max_bytes"),
                ("budget_catalog_reserved_objects", "objects_limit"),
                ("budget_cache_reserved_objects", "objects_limit"),
                ("budget_memory_limit", "memory_limit"),
                ("budget_input_bytes_limit", "input_bytes_limit"),
                ("budget_output_bytes_limit", "output_bytes_limit"),
                ("budget_work_limit", "work_limit"),
                ("budget_objects_limit", "objects_limit"),
            ):
                observed = current[field]
                expected = expected_limits[limit_name]
                if field.endswith("_limit"):
                    if observed != expected:
                        fail(f"{path}.phases.{owner}.{label}.{field}", "cache diagnostic limit differs from configured limits")
                elif observed > expected:
                    fail(f"{path}.phases.{owner}.{label}.{field}", "cache gauge exceeds configured limit")
            for used, limit_field in (
                ("budget_memory_used", "budget_memory_limit"),
                ("budget_input_bytes_used", "budget_input_bytes_limit"),
                ("budget_output_bytes_used", "budget_output_bytes_limit"),
                ("budget_work_used", "budget_work_limit"),
                ("budget_objects_used", "budget_objects_limit"),
            ):
                if current[used] > current[limit_field]:
                    fail(f"{path}.phases.{owner}.{label}.{used}", "cache usage exceeds its limit")
            if current["budget_cache_reserved_bytes"] > current["budget_memory_used"]:
                fail(f"{path}.phases.{owner}.{label}.budget_cache_reserved_bytes", "cache byte reservation exceeds managed memory")
            if current["budget_cache_reserved_objects"] > current["budget_objects_used"]:
                fail(f"{path}.phases.{owner}.{label}.budget_cache_reserved_objects", "cache object reservation exceeds managed objects")
            if current["budget_catalog_reserved_objects"] > current["budget_objects_used"]:
                fail(f"{path}.phases.{owner}.{label}.budget_catalog_reserved_objects", "catalog reservation exceeds managed objects")
            if previous is None:
                if raw_cache["event_delta"] is not None or raw_cache["counter_delta_checked"] is not None:
                    fail(f"{path}.phases.{owner}.{label}", "first available owner point cannot contain an interval delta")
            else:
                for field in CACHE_COUNTER_FIELDS:
                    if current[field] < previous[field]:
                        fail(f"{path}.phases.{owner}.{label}.{field}", "cache counter moved backwards")
                delta = raw_cache["event_delta"]
                if delta is None:
                    fail(f"{path}.phases.{owner}.{label}", "available owner interval omitted checked counter delta")
                for field in CACHE_COUNTER_FIELDS:
                    if delta[field] != current[field] - previous[field]:
                        fail(f"{path}.phases.{owner}.{label}.event_delta.{field}", "delta does not equal adjacent counters")
            previous = current

    for name, points in read_points.items():
        previous: dict[str, Any] | None = None
        for label, current in points:
            if current["availability"] == "unavailable":
                previous = None
                continue
            if previous is not None:
                if current["read_calls"] < previous["read_calls"]:
                    fail(f"{path}.phases.{name}.{label}.read_calls", "source read calls moved backwards")
                if current["read_bytes"] < previous["read_bytes"]:
                    fail(f"{path}.phases.{name}.{label}.read_bytes", "source read bytes moved backwards")
            previous = current

    for name, points in budgets.items():
        previous: dict[str, int] | None = None
        expected_limits = limits if name == "source_budget" else destination_limits
        for label, current in points:
            for used, limit_field in (
                ("memory_used", "memory_limit"), ("input_bytes_used", "input_bytes_limit"),
                ("output_bytes_used", "output_bytes_limit"), ("work_used", "work_limit"),
                ("objects_used", "objects_limit"), ("depth_used", "depth_limit"),
            ):
                if current[limit_field] != expected_limits[limit_field]:
                    fail(f"{path}.phases.{name}.{label}.{limit_field}", "budget limit differs from configured limits")
            if previous is not None:
                for field in CUMULATIVE_BUDGET_FIELDS:
                    if current[field] < previous[field]:
                        fail(f"{path}.phases.{name}.{label}.{field}", "cumulative budget charge moved backwards")
            previous = current

    # Publication consumes the destination editor.  The API cannot provide a
    # post-consumption cache snapshot; an explicit unavailable point is the
    # only valid value at and after that boundary.
    for index, raw in enumerate(phases):
        label = raw["label"]
        consumed = (
            scenario == "lifecycle" and label in {"published", "drop_result", "drop_plan", "drop_view", "drop_caller_sources", "drop_sink"}
        ) or (
            scenario == "repeated-publication"
            and (label.startswith("published_") or label.startswith("drop_result_")
                 or label.startswith("drop_plan_") or label.startswith("drop_sink_")
                 or label in {"fourth_refusal", "drop_view", "drop_sources"})
        )
        if consumed and owner_points["destination_cache"][index][1] is not None:
            fail(f"{path}.phases[{index}].destination_cache", "consumed destination cache was fabricated as available")
        if (
            scenario in {"exact-admission", "one-under", "pinned-eviction", "oversized-bypass"}
            and owner_points["destination_cache"][index][1] is not None
        ):
            fail(f"{path}.phases[{index}].destination_cache", "image boundary fabricated a destination cache owner")
        source_dropped = (
            scenario in {"exact-admission", "one-under", "oversized-bypass"}
            and label in {"drop_view", "drop_source"}
        ) or (
            scenario == "pinned-eviction" and label in {"drop_view", "drop_source"}
        ) or (
            scenario == "lifecycle" and label in {"drop_view", "drop_caller_sources", "drop_sink"}
        ) or (
            scenario == "repeated-publication" and label in {"drop_view", "drop_sources"}
        )
        if source_dropped and owner_points["source_cache"][index][1] is not None:
            fail(f"{path}.phases[{index}].source_cache", "dropped source cache was fabricated as available")

    # All rows end after their caller-owned package/cache handles have been
    # dropped.  Memory and Objects are releasable gauges; cumulative charges
    # remain recorded and are checked above for monotonicity.
    final_source = budgets["source_budget"][-1][1]
    final_destination = budgets["destination_budget"][-1][1]
    if final_source["memory_used"] != 0 or final_source["objects_used"] != 0:
        fail(f"{path}.phases[-1].source_budget", "source releasable resources did not return to zero")
    if final_destination["memory_used"] != 0 or final_destination["objects_used"] != 0:
        fail(f"{path}.phases[-1].destination_budget", "destination releasable resources did not return to zero")
    return counter_delta_count, [value for _, value in budgets["source_budget"]]


def all_dicts(value: Any, path: str = "report") -> Iterable[tuple[str, dict[str, Any]]]:
    if isinstance(value, dict):
        yield path, value
        for key, child in value.items():
            yield from all_dicts(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from all_dicts(child, f"{path}[{index}]")


def check_output_gates(report: dict[str, Any], scenario: str, corpus: str) -> None:
    check_gates(report["gates"], "report.gates", corpus)


def check_refusal_oracle(report: dict[str, Any], scenario: str) -> None:
    if scenario == "lifecycle":
        return
    rows = array(report["samples_raw"], "report.samples_raw")
    for index, raw in enumerate(rows):
        row = obj(raw, f"report.samples_raw[{index}]")
        path = f"report.samples_raw[{index}]"
        if scenario == "one-under":
            if row["result"] != "typed_memory_refusal_before_payload_read":
                fail(f"{path}.result", "one-under row did not record the typed memory refusal")
            if row["exact_payload_verified"] or row["payload_sha256"] is not None or row["payload_bytes"] is not None:
                fail(path, "one-under row fabricated a selected payload oracle")
            if row["payload_read_calls"] != 0 or row["payload_read_bytes"] != 0:
                fail(path, "one-under refusal performed target payload I/O")
            if row["accepted_output_bytes"] != 0 or row["refusal_accepted_output_bytes"] != 0:
                fail(path, "one-under refusal accepted output bytes")
            if row["refusal_resource"] != "Memory" or not row["typed_memory_refusal"] or not row["typed_resource_refusal"]:
                fail(path, "one-under refusal was not a typed Memory resource refusal")
        elif scenario == "exact-admission":
            if row["result"] != "admitted_exact_memory_ceiling" or not row["exact_payload_verified"]:
                fail(path, "exact admission did not produce its exact payload oracle")
            if row["payload_read_calls"] == 0 or row["payload_read_bytes"] == 0:
                fail(path, "exact admission did not read the selected payload")
            if row["typed_memory_refusal"] or row["typed_resource_refusal"]:
                fail(path, "exact admission was reported as refused")
        elif scenario == "oversized-bypass":
            if row["result"] != "admitted_uncached_oversized_payload" or not row["exact_payload_verified"]:
                fail(path, "oversized row omitted the exact payload oracle")
            if not row["oversized_bypass_verified"] or row["typed_resource_refusal"]:
                fail(path, "oversized row was not a successful bypass")
            if row["payload_read_calls"] == 0 or row["payload_read_bytes"] == 0:
                fail(path, "oversized row did not read the payload")
        elif scenario == "pinned-eviction":
            if row["result"] != "pinned_normal_bypass_then_clean_reload":
                fail(path, "pinned row omitted its result oracle")
            if not row["exact_payload_verified"] or not row["eviction_verified"] or not row["pinned_bypass_verified"]:
                fail(path, "pinned row omitted eviction or pinned-bypass evidence")
            if row["oversized_bypass_verified"] or row["typed_resource_refusal"]:
                fail(path, "pinned row was misclassified as oversized or refused")
        elif scenario == "repeated-publication":
            if row["result"] != "three_exact_publications_and_fourth_output_refusal":
                fail(path, "repeated row omitted its publication result oracle")
            if not row["exact_payload_verified"] or not row["repeated_output_identities_verified"]:
                fail(path, "repeated row omitted exact publication identity evidence")
            if row["typed_resource_refusal"] is not True or row["refusal_resource"] != "OutputBytes":
                fail(path, "repeated row omitted the typed fourth output refusal")
            if row["refusal_accepted_output_bytes"] != 0:
                fail(path, "fourth refusal accepted output bytes")


def check_output_identities(report: dict[str, Any], scenario: str) -> None:
    expected_digest = digest(report["expected_output_sha256"], "report.expected_output_sha256")
    output_values: list[tuple[str, str]] = []
    for path, current in all_dicts(report):
        for key, child in current.items():
            lowered = key.lower()
            if lowered.endswith("_output_sha256") or lowered in {"output_sha256", "archive_output_sha256"}:
                if isinstance(child, list):
                    for index, item in enumerate(child):
                        output_values.append((f"{path}.{key}[{index}]", digest(item, f"{path}.{key}[{index}]")))
                elif child is not None:
                    output_values.append((f"{path}.{key}", digest(child, f"{path}.{key}")))
    if expected_digest is not None:
        for path, observed in output_values:
            if observed != expected_digest:
                fail(path, "output identity differs from expected output digest")
    if scenario == "repeated-publication" and output_values:
        if len({value for _, value in output_values}) != 1:
            fail("report", "repeated publication output identities are not stable")


def check_near_row(row: dict[str, Any], path: str, report: dict[str, Any], limits: dict[str, int]) -> None:
    exact_keys(row, NEAR_ROW_FIELDS, (), path)
    uint(row["sample_index"], f"{path}.sample_index")
    text(row["result"], f"{path}.result")
    for field in (
        "exact_payload_verified", "typed_memory_refusal", "typed_resource_refusal",
        "eviction_verified", "pinned_bypass_verified", "oversized_bypass_verified",
        "repeated_output_identities_verified",
    ):
        boolean(row[field], f"{path}.{field}")
    for field in ("payload_read_calls", "payload_read_bytes", "reload_read_calls", "reload_read_bytes"):
        optional_uint(row[field], f"{path}.{field}")
    for field in (
        "accepted_output_bytes", "refusal_accepted_output_bytes", "refusal_sink_accepted_bytes",
        "memory_limit", "cache_max_bytes", "cache_max_entries",
    ):
        uint(row[field], f"{path}.{field}")
    if row["payload_sha256"] is not None:
        digest(row["payload_sha256"], f"{path}.payload_sha256")
    for field in ("payload_bytes", "root_memory_floor", "metadata_memory_used", "payload_memory_used",
                  "drop_image_memory_used", "leaf_bytes"):
        if row[field] is not None:
            uint(row[field], f"{path}.{field}")
    if row["refusal_resource"] is not None and row["refusal_resource"] not in {"Memory", "OutputBytes"}:
        fail(f"{path}.refusal_resource", "unknown refusal resource")
    publication_hashes = array(row["publication_output_sha256"], f"{path}.publication_output_sha256")
    for index, value in enumerate(publication_hashes):
        digest(value, f"{path}.publication_output_sha256[{index}]")
    observed_limits = check_limits(row["configured_limits"], f"{path}.configured_limits")
    if observed_limits != limits:
        fail(f"{path}.configured_limits", "per-row limits differ from report configuration")
    destination_limits = check_limits(
        row["destination_configured_limits"], f"{path}.destination_configured_limits"
    )
    if destination_limits != report["destination_configured_limits"]:
        fail(f"{path}.destination_configured_limits", "per-row destination limits differ from report configuration")
    for field in ("memory_limit", "cache_max_bytes", "cache_max_entries"):
        if row[field] != limits[field]:
            fail(f"{path}.{field}", "row limit differs from report configuration")
    if row["refusal_sink_accepted_bytes"] != row["refusal_accepted_output_bytes"]:
        fail(f"{path}.refusal_sink_accepted_bytes", "refusal sink and accepted output counters differ")

    scenario = report["scenario"]
    payload_read_fields = ("payload_read_calls", "payload_read_bytes")
    reload_read_fields = ("reload_read_calls", "reload_read_bytes")
    if scenario in {"exact-admission", "one-under", "oversized-bypass", "pinned-eviction"}:
        for field in payload_read_fields:
            if row[field] is None:
                fail(f"{path}.{field}", "image-boundary row omitted its payload source-read oracle")
    if scenario == "oversized-bypass":
        for field in reload_read_fields:
            if row[field] is None:
                fail(f"{path}.{field}", "oversized row omitted its reload source-read oracle")
    elif scenario in {"exact-admission", "one-under"}:
        for field in reload_read_fields:
            if row[field] is not None:
                fail(f"{path}.{field}", "row fabricated a reload source-read oracle")
    elif scenario == "pinned-eviction":
        for field in reload_read_fields:
            if row[field] is None:
                fail(f"{path}.{field}", "pinned row omitted its reload source-read oracle")
    elif scenario == "repeated-publication":
        for field in payload_read_fields + reload_read_fields:
            if row[field] is not None:
                fail(f"{path}.{field}", "publication-only row fabricated a selected-image source-read oracle")
    if scenario not in {"one-under", "repeated-publication"}:
        if row["refusal_resource"] is not None or row["refusal_accepted_output_bytes"] != 0:
            fail(path, "successful near row fabricated refusal accounting")
        if row["accepted_output_bytes"] != 0:
            fail(path, "image near row fabricated publication output bytes")
    expected_image = report["expected_image_sha256"]
    expected_image_bytes = report["expected_image_bytes"]
    if scenario in {"exact-admission", "one-under", "oversized-bypass", "pinned-eviction"}:
        if expected_image is None or expected_image_bytes is None:
            fail("report.expected_image_sha256", "image-boundary report omitted its selected payload oracle")
        if row["leaf_bytes"] is None or row["leaf_bytes"] == 0:
            fail(f"{path}.leaf_bytes", "image-boundary row omitted its positive leaf size")
    else:
        if expected_image is not None or expected_image_bytes is not None:
            fail("report.expected_image_sha256", "publication-only report fabricated an image oracle")
        for field in ("root_memory_floor", "metadata_memory_used", "payload_memory_used", "drop_image_memory_used", "leaf_bytes"):
            if row[field] is not None:
                fail(f"{path}.{field}", "publication-only row fabricated image calibration")

    if scenario == "exact-admission":
        if row["payload_sha256"] != expected_image or row["payload_bytes"] != expected_image_bytes:
            fail(path, "exact admission payload identity differs from the selected image oracle")
    elif scenario == "one-under":
        if row["payload_sha256"] is not None or row["payload_bytes"] is not None:
            fail(path, "one-under row fabricated payload identity")
    elif scenario == "oversized-bypass":
        if row["payload_sha256"] != expected_image or row["payload_bytes"] != expected_image_bytes:
            fail(path, "oversized payload identity differs from the selected image oracle")
    elif scenario == "pinned-eviction":
        if (
            row["payload_sha256"] != expected_image
            or row["payload_bytes"] != expected_image_bytes
            or row["payload_bytes"] != row["leaf_bytes"]
        ):
            fail(path, "pinned row payload identity differs from the selected image oracle")
    elif scenario == "repeated-publication":
        if row["publication_output_sha256"] != [report["expected_output_sha256"]] * 3:
            fail(f"{path}.publication_output_sha256", "publication identities are not three exact expected outputs")
        if row["accepted_output_bytes"] != report["destination_configured_limits"]["output_bytes_limit"]:
            fail(f"{path}.accepted_output_bytes", "accepted publication bytes differ from the configured three-output ceiling")
    if scenario in {"exact-admission", "one-under", "oversized-bypass"}:
        root = row["root_memory_floor"]
        metadata = row["metadata_memory_used"]
        payload = row["payload_memory_used"]
        dropped = row["drop_image_memory_used"]
        leaf = row["leaf_bytes"]
        if root is None or metadata is None or payload is None or dropped is None or leaf is None or root == 0:
            fail(path, "image admission row omitted memory-boundary evidence")
        if metadata < root:
            fail(f"{path}.metadata_memory_used", "metadata usage is below the pinned root floor")
        if scenario == "exact-admission":
            if row["memory_limit"] != root + leaf or payload != root + leaf or dropped != payload:
                fail(path, "exact admission memory boundary is inconsistent with root plus payload")
        elif scenario == "one-under":
            if row["memory_limit"] + 1 != root + leaf or payload != root or dropped != root:
                fail(path, "one-under memory refusal boundary is inconsistent with root plus payload")
        else:
            if row["cache_max_bytes"] + 1 != leaf or payload != root + leaf or dropped != root:
                fail(path, "oversized bypass boundary is inconsistent with payload and cache capacity")


def check_near_phase_semantics(
    row: dict[str, Any], path: str, scenario: str,
    limits: dict[str, int], destination_limits: dict[str, int],
) -> None:
    phases = {phase["label"]: phase for phase in row["phases"]}
    if scenario in {"exact-admission", "one-under", "oversized-bypass"}:
        before = phases["metadata"]["source_reads"]
        after = phases["payload"]["source_reads"]
        if before["availability"] != "available" or after["availability"] != "available":
            fail(path, "image payload boundary lacks available source-read points")
        calls = after["read_calls"] - before["read_calls"]
        bytes_read = after["read_bytes"] - before["read_bytes"]
        if row["payload_read_calls"] is None or row["payload_read_bytes"] is None:
            fail(path, "image payload boundary omitted its row source-read oracle")
        if calls != row["payload_read_calls"] or bytes_read != row["payload_read_bytes"]:
            fail(path, "row payload-read oracle differs from metadata-to-payload source counters")
        if scenario == "one-under" and (calls != 0 or bytes_read != 0):
            fail(path, "one-under refusal crossed the selected payload read boundary")
        if scenario != "one-under" and (calls == 0 or bytes_read == 0):
            fail(path, "successful image row performed no selected payload read")
        if scenario == "oversized-bypass":
            reload_before = phases["drop_image"]["source_reads"]
            reload_after = phases["oversized_reload"]["source_reads"]
            if (
                reload_before["availability"] != "available"
                or reload_after["availability"] != "available"
            ):
                fail(path, "oversized row lacks available source-read points for its cold reload")
            reload_calls = reload_after["read_calls"] - reload_before["read_calls"]
            reload_bytes = reload_after["read_bytes"] - reload_before["read_bytes"]
            if row["reload_read_calls"] is None or row["reload_read_bytes"] is None:
                fail(path, "oversized row omitted its reload source-read oracle")
            if (
                reload_calls != row["reload_read_calls"]
                or reload_bytes != row["reload_read_bytes"]
            ):
                fail(path, "oversized reload oracle differs from drop-image-to-reload source counters")
            if reload_calls == 0 or reload_bytes == 0:
                fail(path, "oversized row performed no cold reload source read")
    elif scenario == "pinned-eviction":
        for before_label, after_label in (("metadata", "pinned_a"), ("pinned_a", "pinned_b"), ("drop_image", "reload_b")):
            before = phases[before_label]["source_reads"]
            after = phases[after_label]["source_reads"]
            if before["availability"] != "available" or after["availability"] != "available":
                fail(path, f"pinned row lacks source-read points for {before_label}->{after_label}")
            if after["read_calls"] <= before["read_calls"] or after["read_bytes"] <= before["read_bytes"]:
                fail(path, f"pinned row omitted a cold source read for {before_label}->{after_label}")
        payload_before = phases["metadata"]["source_reads"]
        payload_after = phases["pinned_b"]["source_reads"]
        reload_before = phases["drop_image"]["source_reads"]
        reload_after = phases["reload_b"]["source_reads"]
        payload_calls = payload_after["read_calls"] - payload_before["read_calls"]
        payload_bytes = payload_after["read_bytes"] - payload_before["read_bytes"]
        reload_calls = reload_after["read_calls"] - reload_before["read_calls"]
        reload_bytes = reload_after["read_bytes"] - reload_before["read_bytes"]
        if (
            row["payload_read_calls"] is None
            or row["payload_read_bytes"] is None
            or row["reload_read_calls"] is None
            or row["reload_read_bytes"] is None
        ):
            fail(path, "pinned row omitted a source-read oracle")
        if (
            row["payload_read_calls"] != payload_calls
            or row["payload_read_bytes"] != payload_bytes
            or row["reload_read_calls"] != reload_calls
            or row["reload_read_bytes"] != reload_bytes
        ):
            fail(path, "pinned row read oracles differ from its cross-phase source counters")
        metadata = phases["metadata"]["source_cache"]
        pinned_a = phases["pinned_a"]["source_cache"]
        pinned_b = phases["pinned_b"]["source_cache"]
        dropped = phases["drop_image"]["source_cache"]
        reloaded = phases["reload_b"]["source_cache"]
        for label, point in (("metadata", metadata), ("pinned_a", pinned_a), ("pinned_b", pinned_b), ("drop_image", dropped), ("reload_b", reloaded)):
            if point["availability"] != "available":
                fail(f"{path}.{label}.source_cache", "pinned row omitted an available cache point")
        if pinned_a["retained_entries"] != 2 or pinned_a["evictions"] <= metadata["evictions"]:
            fail(path, "pinned A did not establish the two-entry eviction boundary")
        if (
            pinned_b["evictions"] != pinned_a["evictions"]
            or pinned_b["bypasses"] <= pinned_a["bypasses"]
            or pinned_b["oversized_bypasses"] != metadata["oversized_bypasses"]
            or pinned_b["retained_entries"] != pinned_a["retained_entries"]
            or pinned_b["retained_bytes"] != pinned_a["retained_bytes"]
            or pinned_b["budget_memory_used"] != pinned_a["budget_memory_used"] + row["leaf_bytes"]
        ):
            fail(path, "pinned B did not record a normal pinned bypass with a temporary payload reservation")
        if dropped["budget_memory_used"] != pinned_a["budget_memory_used"]:
            fail(path, "dropping pinned handles did not release the uncached payload reservation")
        if (
            reloaded["successful_loads"] <= dropped["successful_loads"]
            or reloaded["evictions"] <= dropped["evictions"]
            or reloaded["bypasses"] != dropped["bypasses"]
            or reloaded["retained_entries"] != 2
            or reloaded["budget_memory_used"] != pinned_a["budget_memory_used"]
        ):
            fail(path, "pinned reload did not evict and retain within the released cache slot")
    elif scenario == "repeated-publication":
        if any(row[field] is not None for field in (
            "payload_read_calls", "payload_read_bytes", "reload_read_calls", "reload_read_bytes",
        )):
            fail(path, "publication-only row fabricated a selected-image source-read oracle")
        for number in (1, 2, 3):
            phase = phases[f"drop_sink_{number}"]
            budget = check_budget_point(phase["destination_budget"], f"{path}.drop_sink_{number}.destination_budget")
            expected = destination_limits["output_bytes_limit"] * number // 3
            if budget["output_bytes_used"] != expected:
                fail(f"{path}.drop_sink_{number}.destination_budget.output_bytes_used", "publication output charge is not the expected cumulative value")
        final = check_budget_point(phases["fourth_refusal"]["destination_budget"], f"{path}.fourth_refusal.destination_budget")
        if final["output_bytes_used"] != destination_limits["output_bytes_limit"]:
            fail(f"{path}.fourth_refusal.destination_budget.output_bytes_used", "fourth refusal changed the cumulative output boundary")


def check_report(report: Any) -> dict[str, Any]:
    report = obj(report, "report")
    reject_unsupported_claims(report, "report")
    check_all_numbers(report, "report")
    check_hash_fields(report, "report")
    if text(report.get("schema"), "report.schema") != SCHEMA:
        fail("report.schema", f"expected {SCHEMA!r}")
    scenario = text(report.get("scenario"), "report.scenario")
    if scenario not in SCENARIOS:
        fail("report.scenario", f"expected one of {sorted(SCENARIOS)}")
    common = (
        "schema", "scenario", "corpus", "cache_scope", "budget_scope", "source_io_scope",
        "publication_scope", "rss_scope", "samples", "warmup", "checked_iteration_count",
        "source_revision", "binary_sha256", "binary_bytes", "current_exe", "source_archive_sha256",
        "source_archive_bytes", "destination_archive_sha256", "destination_archive_bytes",
        "expected_output_sha256", "expected_output_bytes", "corpus_manifest", "gates",
        "configured_limits", "destination_configured_limits", "phases", "samples_raw",
    )
    if scenario == "lifecycle":
        required = common + ("destination_editor_consumed_during_publish",)
        optional: tuple[str, ...] = ()
    else:
        required = common + ("expected_image_sha256", "expected_image_bytes")
        optional = ()
    exact_keys(report, required, optional, "report")

    corpus = text(report["corpus"], "report.corpus")
    if corpus not in CORPORA:
        fail("report.corpus", f"expected one of {sorted(CORPORA)}")
    if scenario in {"exact-admission", "one-under", "pinned-eviction", "oversized-bypass"} and corpus != "media-rich":
        fail("report.corpus", "image-boundary scenarios require the media-rich corpus")
    samples = uint(report["samples"], "report.samples")
    if not 1 <= samples <= MAX_SAMPLES:
        fail("report.samples", f"expected a bounded count from 1 to {MAX_SAMPLES}")
    warmup = uint(report["warmup"], "report.warmup")
    if warmup > MAX_WARMUP:
        fail("report.warmup", f"expected a bounded count from 0 to {MAX_WARMUP}")
    checked = uint(report["checked_iteration_count"], "report.checked_iteration_count")
    if checked != samples + warmup:
        fail("report.checked_iteration_count", "checked iteration count differs from samples plus warmup")
    revision = text(report["source_revision"], "report.source_revision")
    if not HEX40.fullmatch(revision):
        fail("report.source_revision", "expected a 40-character hexadecimal revision")
    digest(report["binary_sha256"], "report.binary_sha256")
    uint(report["binary_bytes"], "report.binary_bytes")
    text(report["current_exe"], "report.current_exe")
    for name in (
        "source_archive_sha256", "destination_archive_sha256", "expected_output_sha256",
    ):
        digest(report[name], f"report.{name}")
    for name in (
        "source_archive_bytes", "destination_archive_bytes", "expected_output_bytes",
    ):
        uint(report[name], f"report.{name}")
    for name in ("cache_scope", "budget_scope", "source_io_scope", "publication_scope", "rss_scope"):
        check_scope(report[name], f"report.{name}")
    if scenario == "lifecycle":
        if report["destination_editor_consumed_during_publish"] is not True:
            fail("report.destination_editor_consumed_during_publish", "consuming publication boundary is not recorded")
    else:
        if report["expected_image_sha256"] is not None:
            digest(report["expected_image_sha256"], "report.expected_image_sha256")
        if report["expected_image_bytes"] is not None:
            uint(report["expected_image_bytes"], "report.expected_image_bytes")
        if scenario in {"exact-admission", "one-under", "pinned-eviction", "oversized-bypass"}:
            if report["expected_image_sha256"] is None or report["expected_image_bytes"] is None:
                fail("report.expected_image_sha256", "image-boundary report omitted its selected payload oracle")
        elif report["expected_image_sha256"] is not None or report["expected_image_bytes"] is not None:
            fail("report.expected_image_sha256", "publication-only report fabricated an image oracle")

    limits = check_limits(report["configured_limits"], "report.configured_limits")
    destination_limits = check_limits(
        report["destination_configured_limits"], "report.destination_configured_limits"
    )
    check_corpus_manifest(report)
    check_output_gates(report, scenario, corpus)
    phase_description_labels = check_phase_descriptions(report["phases"], "report.phases")
    expected_description_labels = SOURCE_CACHE_PHASES if scenario == "lifecycle" else NEAR_DESCRIPTION_PHASES
    if phase_description_labels != expected_description_labels:
        fail("report.phases", "phase description labels do not match the frozen scenario schema")

    rows = array(report["samples_raw"], "report.samples_raw")
    if len(rows) != samples:
        fail("report.samples_raw", "raw row count differs from reported samples")
    interval_count = 0
    total_cache_observations = 0
    for index, raw in enumerate(rows):
        row = obj(raw, f"report.samples_raw[{index}]")
        path = f"report.samples_raw[{index}]"
        if scenario == "lifecycle":
            exact_keys(row, LIFECYCLE_ROW_FIELDS, (), path)
            if uint(row["sample_index"], f"{path}.sample_index") != index:
                fail(f"{path}.sample_index", "raw sample order is not preserved")
            if row["exact_output_verified"] is not True:
                fail(f"{path}.exact_output_verified", "exact output gate is false")
            if digest(row["output_sha256"], f"{path}.output_sha256") != report["expected_output_sha256"].lower():
                fail(f"{path}.output_sha256", "row output identity differs from expected output")
            if uint(row["output_bytes"], f"{path}.output_bytes") != report["expected_output_bytes"]:
                fail(f"{path}.output_bytes", "row output byte length differs from expected output")
        else:
            check_near_row(row, path, report, limits)
            if uint(row["sample_index"], f"{path}.sample_index") != index:
                fail(f"{path}.sample_index", "raw sample order is not preserved")
        deltas, _budget_values = check_phase_records(row, path, scenario, limits, destination_limits)
        if scenario != "lifecycle":
            check_near_phase_semantics(row, path, scenario, limits, destination_limits)
        interval_count += deltas
        total_cache_observations += 2 * len(row["phases"])
        if scenario == "repeated-publication":
            # Each successful publication is followed by a fresh destination
            # editor under the same destination Budget root.  Releasable
            # gauges must be zero at every drop_plan boundary; cumulative
            # OutputBytes is intentionally retained and checked as a charge.
            for phase in row["phases"]:
                if phase["label"] in {"drop_plan_1", "drop_plan_2", "drop_plan_3"}:
                    budget = check_budget_point(phase["destination_budget"], f"{path}.{phase['label']}.destination_budget")
                    if budget["memory_used"] != 0 or budget["objects_used"] != 0:
                        fail(f"{path}.{phase['label']}.destination_budget", "destination releasable gauges did not return to zero")
    if total_cache_observations == 0:
        fail("report.samples_raw", "no cache observations were retained")
    if interval_count == 0:
        fail("report.samples_raw", "no checked cache counter delta was retained")
    check_refusal_oracle(report, scenario)
    check_output_identities(report, scenario)
    return {
        "status": "valid", "schema": SCHEMA, "scenario": scenario, "corpus": corpus,
        "samples": samples, "cache_observations": total_cache_observations,
        "checked_counter_deltas": interval_count,
    }


def load_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=reject_duplicate_pairs,
            parse_constant=reject_nonfinite_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise VerificationError(f"{path}: cannot read JSON: {exc}") from exc


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", nargs="?", type=Path, help="cache-retention report JSON")
    parser.add_argument("--report", dest="report_option", type=Path, help="cache-retention report JSON")
    parser.add_argument("--protocol", type=Path, default=None, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    path = args.report_option or args.path
    if path is None:
        parser.error("a report path is required")
    try:
        result = check_report(load_json(path))
    except Exception as exc:
        print(f"INVALID: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
