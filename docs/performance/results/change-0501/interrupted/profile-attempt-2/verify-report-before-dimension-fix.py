#!/usr/bin/env python3
"""Independently validate one PPTX provider-lifecycle report."""

from __future__ import annotations

import argparse
import json
import math
import re
from pathlib import Path
from typing import Any


SCHEMA = "pptx_provider_lifecycle_v1"
PHASES = (
    "baseline",
    "opened",
    "planned",
    "published",
    "drop_result",
    "drop_plan",
    "drop_view",
    "drop_caller_sources",
    "drop_sink",
)
GATES = (
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
CACHE_COUNTERS = (
    "hits",
    "cold_loads",
    "waiter_joins",
    "successful_loads",
    "failed_loads",
    "evictions",
    "bypasses",
    "oversized_bypasses",
    "allocation_bypasses",
    "budget_reservation_failures",
)
CACHE_GAUGES = (
    "retained_entries",
    "retained_bytes",
    "in_flight_loads",
    "budget_memory_used",
    "budget_cache_reserved_bytes",
    "budget_input_bytes_used",
    "budget_output_bytes_used",
    "budget_work_used",
    "budget_objects_used",
    "budget_catalog_reserved_objects",
    "budget_cache_reserved_objects",
)
BUDGET_USED = (
    "memory_used",
    "input_bytes_used",
    "output_bytes_used",
    "work_used",
    "objects_used",
    "depth_used",
)
BUDGET_LIMITS = (
    "memory_limit",
    "input_bytes_limit",
    "output_bytes_limit",
    "work_limit",
    "objects_limit",
    "depth_limit",
)
READ_COUNTERS = (
    "logical_calls",
    "requested_bytes",
    "returned_bytes",
    "short_reads",
    "delayed_calls",
)
READ_FIELDS = READ_COUNTERS + ("min_request_bytes", "max_request_bytes", "request_size_counts")
REQUEST_BUCKETS = 18
HEX40 = re.compile(r"^[0-9a-f]{40}$")
HEX64 = re.compile(r"^[0-9a-f]{64}$")
U64_MAX = (1 << 64) - 1


class InvalidReport(ValueError):
    """The producer report does not satisfy the evidence contract."""


def fail(message: str) -> None:
    raise InvalidReport(message)


def load(path: Path) -> dict[str, Any]:
    def duplicate(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                fail(f"duplicate JSON key: {key}")
            result[key] = value
        return result

    def nonfinite(value: str) -> Any:
        fail(f"non-finite JSON constant: {value}")

    try:
        value = json.loads(path.read_text(), object_pairs_hook=duplicate, parse_constant=nonfinite)
    except InvalidReport:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON: {error}")
    if not isinstance(value, dict):
        fail("report root is not an object")
    return value


def text(value: Any, label: str, *, empty: bool = False) -> str:
    if not isinstance(value, str) or (not empty and not value):
        fail(f"{label}: expected text")
    return value


def boolean(value: Any, label: str) -> bool:
    if not isinstance(value, bool):
        fail(f"{label}: expected boolean")
    return value


def uint(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value <= U64_MAX:
        fail(f"{label}: expected u64")
    return value


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def array(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        fail(f"{label}: expected array")
    return value


def digest(value: Any, label: str, length: int = 64) -> str:
    value = text(value, label)
    pattern = HEX64 if length == 64 else HEX40
    if not pattern.fullmatch(value):
        fail(f"{label}: expected lowercase SHA-256/Git hexadecimal identity")
    return value


def check_limits(value: Any, label: str) -> dict[str, int]:
    limits = obj(value, label)
    result = {}
    for key in (
        "cache_max_bytes",
        "cache_max_entries",
        "memory_limit",
        "input_bytes_limit",
        "output_bytes_limit",
        "work_limit",
        "objects_limit",
        "depth_limit",
    ):
        result[key] = uint(limits.get(key), f"{label}.{key}")
        if result[key] == 0:
            fail(f"{label}.{key}: zero configured limit")
    return result


def check_budget(value: Any, label: str, limits: dict[str, int]) -> dict[str, int]:
    budget = obj(value, label)
    result = {key: uint(budget.get(key), f"{label}.{key}") for key in BUDGET_USED + BUDGET_LIMITS}
    for used, limit in zip(BUDGET_USED, BUDGET_LIMITS):
        if result[used] > result[limit]:
            fail(f"{label}.{used}: use exceeds limit")
    for key, expected in zip(BUDGET_LIMITS, BUDGET_LIMITS):
        limit_key = {
            "memory_limit": "memory_limit",
            "input_bytes_limit": "input_bytes_limit",
            "output_bytes_limit": "output_bytes_limit",
            "work_limit": "work_limit",
            "objects_limit": "objects_limit",
            "depth_limit": "depth_limit",
        }[key]
        if result[key] != limits[limit_key]:
            fail(f"{label}.{key}: differs from configured limit")
    return result


def check_cache(value: Any, label: str, limits: dict[str, int]) -> dict[str, Any] | None:
    cache = obj(value, label)
    state = text(cache.get("availability"), f"{label}.availability")
    if state not in {"available", "unavailable"}:
        fail(f"{label}.availability: invalid state")
    if state == "unavailable":
        text(cache.get("unavailable_reason"), f"{label}.unavailable_reason")
        for key in CACHE_COUNTERS + CACHE_GAUGES + (
            "budget_managed",
            "budget_memory_limit",
            "budget_input_bytes_limit",
            "budget_output_bytes_limit",
            "budget_work_limit",
            "budget_objects_limit",
            "budget_catalog_reserved_objects",
            "budget_cache_reserved_objects",
            "interval_label",
            "counter_delta_checked",
            "event_delta",
        ):
            if cache.get(key) is not None:
                fail(f"{label}.{key}: unavailable point contains a value")
        return None
    if cache.get("unavailable_reason") is not None:
        fail(f"{label}: available point has an unavailable reason")
    for key in CACHE_COUNTERS + CACHE_GAUGES:
        uint(cache.get(key), f"{label}.{key}")
    if cache.get("budget_managed") is not True:
        fail(f"{label}.budget_managed: cache is not managed")
    for key, limit in (
        ("retained_bytes", limits["cache_max_bytes"]),
        ("retained_entries", limits["cache_max_entries"]),
        ("budget_cache_reserved_bytes", limits["cache_max_bytes"]),
        ("budget_cache_reserved_objects", limits["objects_limit"]),
        ("budget_catalog_reserved_objects", limits["objects_limit"]),
    ):
        if cache[key] > limit:
            fail(f"{label}.{key}: exceeds configured cache/object limit")
    for used, limit in (
        ("budget_memory_used", "budget_memory_limit"),
        ("budget_input_bytes_used", "budget_input_bytes_limit"),
        ("budget_output_bytes_used", "budget_output_bytes_limit"),
        ("budget_work_used", "budget_work_limit"),
        ("budget_objects_used", "budget_objects_limit"),
    ):
        if cache[used] > cache[limit]:
            fail(f"{label}.{used}: exceeds cache diagnostic limit")
    event = cache.get("event_delta")
    if event is not None:
        event = obj(event, f"{label}.event_delta")
        for key in CACHE_COUNTERS:
            uint(event.get(key), f"{label}.event_delta.{key}")
        if cache.get("counter_delta_checked") is not True:
            fail(f"{label}: cache event delta was not marked checked")
        text(cache.get("interval_label"), f"{label}.interval_label")
    elif cache.get("counter_delta_checked") is not None or cache.get("interval_label") is not None:
        fail(f"{label}: incomplete cache delta marker")
    return cache


def check_read(value: Any, label: str, *, range_cap: int | None) -> dict[str, Any] | None:
    read = obj(value, label)
    state = text(read.get("availability"), f"{label}.availability")
    if state not in {"available", "unavailable"}:
        fail(f"{label}.availability: invalid state")
    if state == "unavailable":
        text(read.get("unavailable_reason"), f"{label}.unavailable_reason")
        for key in READ_FIELDS + ("counter_delta_checked", "delta"):
            if read.get(key) is not None:
                fail(f"{label}.{key}: unavailable point contains a value")
        return None
    if read.get("unavailable_reason") is not None:
        fail(f"{label}: available point has an unavailable reason")
    for key in READ_COUNTERS:
        uint(read.get(key), f"{label}.{key}")
    minimum = read.get("min_request_bytes")
    maximum = read.get("max_request_bytes")
    if minimum is not None:
        uint(minimum, f"{label}.min_request_bytes")
    if maximum is not None:
        uint(maximum, f"{label}.max_request_bytes")
    counts = array(read.get("request_size_counts"), f"{label}.request_size_counts")
    if len(counts) != REQUEST_BUCKETS:
        fail(f"{label}.request_size_counts: expected {REQUEST_BUCKETS} buckets")
    counts = [uint(item, f"{label}.request_size_counts[{index}]") for index, item in enumerate(counts)]
    if sum(counts) != read["logical_calls"]:
        fail(f"{label}.request_size_counts: does not sum to logical calls")
    if read["returned_bytes"] > read["requested_bytes"]:
        fail(f"{label}: returned bytes exceed requested bytes")
    if read["short_reads"] > read["logical_calls"] or read["delayed_calls"] > read["logical_calls"]:
        fail(f"{label}: read subcounter exceeds logical calls")
    if range_cap is not None and read["returned_bytes"] > read["logical_calls"] * range_cap:
        fail(f"{label}: returned bytes exceed range cap")
    delta = obj(read.get("delta"), f"{label}.delta")
    for key in READ_COUNTERS:
        uint(delta.get(key), f"{label}.delta.{key}")
    for key in ("min_request_bytes", "max_request_bytes"):
        if delta.get(key) is not None:
            uint(delta[key], f"{label}.delta.{key}")
    delta_counts = array(delta.get("request_size_counts"), f"{label}.delta.request_size_counts")
    if len(delta_counts) != REQUEST_BUCKETS:
        fail(f"{label}.delta.request_size_counts: expected {REQUEST_BUCKETS} buckets")
    if sum(uint(item, f"{label}.delta.request_size_counts[{index}]") for index, item in enumerate(delta_counts)) != delta["logical_calls"]:
        fail(f"{label}.delta.request_size_counts: does not sum to calls")
    if read.get("counter_delta_checked") is not True:
        fail(f"{label}.counter_delta_checked: missing true marker")
    return {**read, "request_size_counts": counts, "delta": {**delta, "request_size_counts": delta_counts}}


def check_rss(value: Any, label: str) -> None:
    rss = obj(value, label)
    state = text(rss.get("availability"), f"{label}.availability")
    if state == "available":
        if rss.get("unavailable_reason") is not None:
            fail(f"{label}: available RSS has a reason")
        uint(rss.get("rss_bytes"), f"{label}.rss_bytes")
        uint(rss.get("vm_hwm_bytes"), f"{label}.vm_hwm_bytes")
    elif state == "unavailable":
        text(rss.get("unavailable_reason"), f"{label}.unavailable_reason")
        if rss.get("rss_bytes") is not None or rss.get("vm_hwm_bytes") is not None:
            fail(f"{label}: unavailable RSS has bytes")
    else:
        fail(f"{label}.availability: invalid state")


def check_provider_config(report: dict[str, Any]) -> int | None:
    provider = text(report.get("provider"), "report.provider")
    if provider not in {"bytes", "file", "range"}:
        fail("report.provider: unsupported provider")
    config = obj(report.get("provider_config"), "report.provider_config")
    if config.get("provider") != provider:
        fail("report.provider_config.provider: differs from report provider")
    cap = config.get("max_range_bytes")
    delay = config.get("delay_us")
    if cap is not None:
        uint(cap, "report.provider_config.max_range_bytes")
    if delay is not None:
        uint(delay, "report.provider_config.delay_us")
    if config.get("adapter_max_range_bytes") != cap or config.get("adapter_delay_us") != delay:
        fail("report.provider_config: adapter values differ")
    if provider == "range":
        if cap is None or cap == 0 or cap > 1_048_576 or delay is None or config.get("delay_configured") is not True:
            fail("report.provider_config: invalid range controls")
        if config.get("bytes_file_unlimited_cap") is not False:
            fail("report.provider_config: range marked unlimited")
        return cap
    if cap is not None or delay is not None or config.get("delay_configured") is not False:
        fail("report.provider_config: direct provider has range controls")
    if config.get("bytes_file_unlimited_cap") is not True:
        fail("report.provider_config: direct provider is not unlimited")
    return None


def check_report(report: dict[str, Any]) -> dict[str, Any]:
    if report.get("schema") != SCHEMA:
        fail("report.schema: unexpected provider report schema")
    corpus = text(report.get("corpus"), "report.corpus")
    if corpus not in {"plain", "media-rich"}:
        fail("report.corpus: unsupported corpus")
    range_cap = check_provider_config(report)
    samples = uint(report.get("samples"), "report.samples")
    warmup = uint(report.get("warmup"), "report.warmup")
    if samples != 30 or warmup != 3:
        fail("report sample dimensions differ from the formal protocol")
    if uint(report.get("checked_iteration_count"), "report.checked_iteration_count") != samples + warmup:
        fail("report.checked_iteration_count: does not equal samples plus warmups")
    if not HEX40.fullmatch(text(report.get("source_revision"), "report.source_revision")):
        fail("report.source_revision: expected lowercase Git revision")
    for key in ("source_archive_sha256", "destination_archive_sha256", "expected_output_sha256", "binary_sha256"):
        digest(report.get(key), f"report.{key}")
    for key in ("source_archive_bytes", "destination_archive_bytes", "expected_output_bytes", "binary_bytes"):
        uint(report.get(key), f"report.{key}")
    text(report.get("current_exe"), "report.current_exe")
    limits = check_limits(report.get("configured_limits"), "report.configured_limits")
    check_limits(report.get("destination_configured_limits"), "report.destination_configured_limits")
    if report.get("destination_editor_consumed_during_publish") is not True:
        fail("report.destination_editor_consumed_during_publish: false")
    if report.get("final_memory_objects_depth_zero_checked") is not True:
        fail("report.final_memory_objects_depth_zero_checked: false")
    gates = obj(report.get("gates"), "report.gates")
    for key in GATES:
        if gates.get(key) is not True:
            fail(f"report.gates.{key}: false")

    phases = array(report.get("phases"), "report.phases")
    labels = tuple(text(obj(item, f"report.phases[{index}]").get("label"), f"report.phases[{index}].label") for index, item in enumerate(phases))
    if labels != PHASES:
        fail("report.phases: phase order differs from the lifecycle contract")
    rows = array(report.get("samples_raw"), "report.samples_raw")
    if len(rows) != samples:
        fail("report.samples_raw: row count differs from samples")
    expected_digest = report["expected_output_sha256"]
    expected_bytes = report["expected_output_bytes"]
    timing_names = ("open_source_ns", "open_destination_ns", "open_ns", "plan_ns", "publication_ns", "api_sum_ns")
    for index, row_value in enumerate(rows):
        row = obj(row_value, f"report.samples_raw[{index}]")
        if uint(row.get("sample_index"), f"sample[{index}].sample_index") != index:
            fail(f"sample[{index}].sample_index: not contiguous")
        if row.get("exact_output_verified") is not True:
            fail(f"sample[{index}].exact_output_verified: false")
        if digest(row.get("output_sha256"), f"sample[{index}].output_sha256") != expected_digest:
            fail(f"sample[{index}].output_sha256: differs from report expected output")
        if uint(row.get("output_bytes"), f"sample[{index}].output_bytes") != expected_bytes:
            fail(f"sample[{index}].output_bytes: differs from report expected output")
        timing = obj(row.get("timings"), f"sample[{index}].timings")
        for key in timing_names:
            uint(timing.get(key), f"sample[{index}].timings.{key}")
        if timing["open_ns"] != timing["open_source_ns"] + timing["open_destination_ns"]:
            fail(f"sample[{index}].timings.open_ns: sum mismatch")
        if timing["api_sum_ns"] != timing["open_ns"] + timing["plan_ns"] + timing["publication_ns"]:
            fail(f"sample[{index}].timings.api_sum_ns: sum mismatch")

        previous: dict[str, Any] = {"source": None, "destination": None}
        previous_budget: dict[str, dict[str, int] | None] = {"source": None, "destination": None}
        previous_cache: dict[str, dict[str, Any] | None] = {"source": None, "destination": None}
        for phase_index, phase_value in enumerate(array(row.get("phases"), f"sample[{index}].phases")):
            phase = obj(phase_value, f"sample[{index}].phases[{phase_index}]")
            if phase.get("label") != PHASES[phase_index]:
                fail(f"sample[{index}].phases[{phase_index}].label: differs from phase description")
            for owner in ("source", "destination"):
                budget = check_budget(phase.get(f"{owner}_budget"), f"sample[{index}].{owner}_budget", limits)
                old_budget = previous_budget[owner]
                if old_budget is not None:
                    for key in ("input_bytes_used", "output_bytes_used", "work_used"):
                        if budget[key] < old_budget[key]:
                            fail(f"sample[{index}].{owner}_budget.{key}: moved backwards")
                previous_budget[owner] = budget
                cache = check_cache(phase.get(f"{owner}_cache"), f"sample[{index}].{owner}_cache", limits)
                old_cache = previous_cache[owner]
                if cache is not None and old_cache is not None:
                    for key in CACHE_COUNTERS:
                        if cache[key] < old_cache[key]:
                            fail(f"sample[{index}].{owner}_cache.{key}: moved backwards")
                    event = cache.get("event_delta")
                    if event is None:
                        fail(f"sample[{index}].{owner}_cache: missing interval delta")
                    for key in CACHE_COUNTERS:
                        if event[key] != cache[key] - old_cache[key]:
                            fail(f"sample[{index}].{owner}_cache.event_delta.{key}: arithmetic mismatch")
                previous_cache[owner] = cache
                read = check_read(
                    phase.get(f"{owner}_reads"),
                    f"sample[{index}].{owner}_reads",
                    range_cap=range_cap,
                )
                old_read = previous[owner]
                if read is not None and old_read is not None:
                    for key in READ_COUNTERS:
                        if read[key] < old_read[key]:
                            fail(f"sample[{index}].{owner}_reads.{key}: moved backwards")
                    for key in READ_COUNTERS:
                        if read["delta"][key] != read[key] - old_read[key]:
                            fail(f"sample[{index}].{owner}_reads.delta.{key}: arithmetic mismatch")
                    for old_count, new_count, delta_count in zip(
                        old_read["request_size_counts"], read["request_size_counts"], read["delta"]["request_size_counts"]
                    ):
                        if new_count < old_count or delta_count != new_count - old_count:
                            fail(f"sample[{index}].{owner}_reads: histogram arithmetic mismatch")
                previous[owner] = read
                check_rss(phase.get("rss"), f"sample[{index}].rss")
        for owner in ("source", "destination"):
            final_budget = previous_budget[owner]
            assert final_budget is not None
            for key in ("memory_used", "objects_used", "depth_used"):
                if final_budget[key] != 0:
                    fail(f"sample[{index}].{owner}_budget.{key}: owner was not released")
    return {
        "corpus": corpus,
        "provider": report["provider"],
        "samples": samples,
        "expected_output_sha256": expected_digest,
        "expected_output_bytes": expected_bytes,
        "source_archive_sha256": report["source_archive_sha256"],
        "destination_archive_sha256": report["destination_archive_sha256"],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report", type=Path)
    args = parser.parse_args()
    result = check_report(load(args.report))
    print(json.dumps({"status": "pass", **result}, sort_keys=True))


if __name__ == "__main__":
    main()
