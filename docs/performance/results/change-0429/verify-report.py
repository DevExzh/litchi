#!/usr/bin/env python3
"""Independently verify change-0429 provider and native image journals.

The producer is deliberately not imported.  This checker treats every report
as untrusted JSON and validates the serialized contract, ownership boundaries,
counter arithmetic, and the small checked-in native fixture oracle.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Iterable

U64_MAX = (1 << 64) - 1
MAX_SAMPLES = 1_000
MAX_WARMUP = 1_000
HEX40 = re.compile(r"^[0-9a-f]{40}$")
HEX64 = re.compile(r"^[0-9a-f]{64}$")

PROVIDER_SCHEMA = "pptx_provider_lifecycle_v1"
NATIVE_SCHEMA = "pptx_native_image_lifecycle_v1"
PROVIDERS = {"bytes", "file", "range"}
CORPORA = {"plain", "media-rich"}
FIXTURES = {"poi-slide", "poi-video"}

# The source adapter records one cumulative counter for each request-size
# bucket.  Keep the bucket contract here instead of trusting a producer-side
# constant or accepting an arbitrary histogram length.
REQUEST_SIZE_BUCKETS = (
    (1, 1),
    (2, 2),
    (3, 4),
    (5, 8),
    (9, 16),
    (17, 32),
    (33, 64),
    (65, 128),
    (129, 256),
    (257, 512),
    (513, 1024),
    (1025, 2048),
    (2049, 4096),
    (4097, 8192),
    (8193, 16384),
    (16385, 32768),
    (32769, 65536),
    (65537, None),
)

CACHE_COUNTER_FIELDS = (
    "hits", "cold_loads", "waiter_joins", "successful_loads", "failed_loads",
    "evictions", "bypasses", "oversized_bypasses", "allocation_bypasses",
    "budget_reservation_failures",
)
CACHE_GAUGE_FIELDS = (
    "retained_entries", "retained_bytes", "in_flight_loads", "budget_managed",
    "budget_memory_used", "budget_cache_reserved_bytes", "budget_memory_limit",
    "budget_input_bytes_used", "budget_input_bytes_limit",
    "budget_output_bytes_used", "budget_output_bytes_limit", "budget_work_used",
    "budget_work_limit", "budget_objects_used", "budget_objects_limit",
    "budget_catalog_reserved_objects", "budget_cache_reserved_objects",
)
CACHE_FIELDS = CACHE_COUNTER_FIELDS + CACHE_GAUGE_FIELDS
CACHE_POINT_FIELDS = (
    "availability", "unavailable_reason", "interval_label", "counter_delta_checked",
    "event_delta",
) + CACHE_FIELDS
BUDGET_FIELDS = (
    "memory_used", "memory_limit", "input_bytes_used", "input_bytes_limit",
    "output_bytes_used", "output_bytes_limit", "work_used", "work_limit",
    "objects_used", "objects_limit", "depth_used", "depth_limit",
)
LIMIT_FIELDS = (
    "cache_max_bytes", "cache_max_entries", "memory_limit", "input_bytes_limit",
    "output_bytes_limit", "work_limit", "objects_limit", "depth_limit",
)
READ_FIELDS = (
    "availability", "unavailable_reason", "logical_calls", "requested_bytes",
    "returned_bytes", "request_size_counts", "min_request_bytes", "max_request_bytes",
    "short_reads", "delayed_calls", "counter_delta_checked", "delta",
)
READ_COUNTER_FIELDS = ("logical_calls", "requested_bytes", "returned_bytes", "short_reads", "delayed_calls")
READ_DELTA_COMMON_FIELDS = (
    "logical_calls", "requested_bytes", "returned_bytes", "request_size_counts",
    "short_reads", "delayed_calls",
)
READ_DELTA_BOUNDED_FIELDS = (
    "logical_calls", "requested_bytes", "returned_bytes", "min_request_bytes",
    "max_request_bytes", "request_size_counts", "short_reads", "delayed_calls",
)
RSS_FIELDS = ("availability", "unavailable_reason", "rss_bytes", "vm_hwm_bytes")

PROVIDER_PHASES = (
    "baseline", "opened", "planned", "published", "drop_result", "drop_plan",
    "drop_view", "drop_caller_sources", "drop_sink",
)
NATIVE_PHASES = (
    "baseline", "opened", "selected", "loaded", "drop_view", "drop_slide",
    "drop_image", "drop_source",
)
NATIVE_OWNER_TOKENS = {
    "baseline": ("no source adapter",),
    "opened": ("source-backed", "caller source adapter"),
    "selected": ("selected slide",),
    "loaded": ("returned image payload",),
    "drop_view": ("returned image payload", "public view"),
    "drop_slide": ("returned image payload", "selected slide"),
    "drop_image": ("caller source adapter",),
    "drop_source": ("no lifecycle-owned source",),
}
PHASE_DESCRIPTION_FIELDS = ("label", "live_owners")
PROVIDER_PHASE_FIELDS = (
    "label", "source_cache", "destination_cache", "source_reads", "destination_reads",
    "source_budget", "destination_budget", "rss",
)
NATIVE_PHASE_FIELDS = ("label", "source_cache", "source_reads", "source_budget", "rss")

GATE_FIELDS = (
    "matched_owned_corpus_verified", "semantic_output_verified", "package_topology_verified",
    "dependency_boundary_verified", "layout_reuse_verified", "untouched_destination_members_verified",
    "deterministic_output_verified", "source_version_stability_verified",
    "source_revision_refusal_verified", "destination_revision_refusal_verified",
    "foreign_destination_refusal_verified", "added_opc_parts_verified", "added_zip_members_verified",
    "media_leaf_payloads_verified", "media_leaf_content_types_verified", "media_relationships_verified",
)

NATIVE_ORACLES: dict[str, dict[str, Any]] = {
    "poi-slide": {
        "path": "test-data/poi/test-data/slideshow/bug62513.pptx",
        "name": "bug62513.pptx",
        "resave_scope": "Apache POI producer fixture; retained as supplied",
        "archive_sha256": "cd841112bd5b53f21e8434d080af8b2eeca78a07bc820e9ea0b97a0962c85c79",
        "archive_bytes": 384775,
        "slide": 4,
        "image": 0,
        "image_count": 1,
        "shape_position": 1,
        "shape_id": 37890,
        "shape_name": "Picture 2",
        "bounds": {"x": 1115616, "y": 2276872, "width": 7017380, "height": 3262858},
        "relationship_id": "rId2",
        "part": "/ppt/media/image2.jpeg",
        "content_type": "image/jpeg",
        "payload_bytes": 21997,
        "payload_sha256": "d5f10480ab75ce1175eea7432e684e9fac7319a7b9bd15512759385eae4842db",
    },
    "poi-video": {
        "path": "test-data/poi/test-data/slideshow/EmbeddedVideo.pptx",
        "name": "EmbeddedVideo.pptx",
        "resave_scope": "Apache POI producer fixture with embedded-video poster image; retained as supplied",
        "archive_sha256": "7940e3b1a339db11f00b65399a2fe77e0e85a5da3a30ac8d6c8a0a77527b2ab2",
        "archive_bytes": 201418,
        "slide": 0,
        "image": 0,
        "image_count": 1,
        "shape_position": 0,
        "shape_id": 2,
        "shape_name": "file_example_MP4_480_1_5MG_Trim",
        "bounds": {"x": 3810000, "y": 2143125, "width": 4572000, "height": 2571750},
        "relationship_id": "rId4",
        "part": "/ppt/media/image1.png",
        "content_type": "image/png",
        "payload_bytes": 65215,
        "payload_sha256": "f5516c6cae484df63ce03db77fb69b778660916b9207de5a4e04aa5e3b72908d",
    },
}


class VerificationError(ValueError):
    """A report does not satisfy the frozen evidence contract."""


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
    if value != value.lower() or not HEX64.fullmatch(value):
        fail(path, "expected a lowercase SHA-256 hexadecimal digest")
    return value


def exact_keys(value: dict[str, Any], required: Iterable[str], optional: Iterable[str], path: str) -> None:
    required = tuple(required)
    allowed = set(required) | set(optional)
    missing = [name for name in required if name not in value]
    if missing:
        fail(path, f"missing fields: {', '.join(missing)}")
    unknown = sorted(set(value) - allowed)
    if unknown:
        fail(path, f"unknown fields: {', '.join(unknown)}")


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
    elif isinstance(value, float):
        fail(path, "floating-point values are not permitted; expected u64 or boolean")
    elif isinstance(value, dict):
        for key, child in value.items():
            check_all_numbers(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            check_all_numbers(child, f"{path}[{index}]")


def check_hash_fields(value: Any, path: str) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            lowered = key.lower()
            if lowered.endswith("_sha256") or lowered in {"sha256", "hash"}:
                if child is not None:
                    digest(child, f"{path}.{key}")
            check_hash_fields(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            check_hash_fields(child, f"{path}[{index}]")


def check_scope(value: Any, path: str, tokens: tuple[str, ...] = ("source", "cache", "budget", "operation", "process", "rss")) -> None:
    scope = text(value, path).lower()
    if scope in {"wrong_scope", "wrong scope", "bogus", "mutated", "unknown"}:
        fail(path, "invalid diagnostic scope")
    if not any(token in scope for token in tokens):
        fail(path, "scope does not identify a supported evidence boundary")


def check_limits(value: Any, path: str) -> dict[str, int]:
    limits = obj(value, path)
    exact_keys(limits, LIMIT_FIELDS, (), path)
    result = {field: uint(limits[field], f"{path}.{field}") for field in LIMIT_FIELDS}
    for field, current in result.items():
        if current == 0:
            fail(f"{path}.{field}", "configured limit must be positive")
    return result


def check_budget(value: Any, path: str) -> dict[str, int]:
    point = obj(value, path)
    exact_keys(point, BUDGET_FIELDS, (), path)
    result = {field: uint(point[field], f"{path}.{field}") for field in BUDGET_FIELDS}
    for used, limit in (
        ("memory_used", "memory_limit"), ("input_bytes_used", "input_bytes_limit"),
        ("output_bytes_used", "output_bytes_limit"), ("work_used", "work_limit"),
        ("objects_used", "objects_limit"), ("depth_used", "depth_limit"),
    ):
        if result[used] > result[limit]:
            fail(f"{path}.{used}", "budget usage exceeds its configured limit")
    return result


def check_cache(value: Any, path: str) -> dict[str, Any] | None:
    cache = obj(value, path)
    exact_keys(cache, CACHE_POINT_FIELDS, (), path)
    state = text(cache["availability"], f"{path}.availability")
    if state not in {"available", "unavailable"}:
        fail(f"{path}.availability", "unknown cache availability state")
    if state == "unavailable":
        text(cache["unavailable_reason"], f"{path}.unavailable_reason")
        for field in CACHE_FIELDS + ("interval_label", "counter_delta_checked", "event_delta"):
            if cache[field] is not None:
                fail(f"{path}.{field}", "unavailable cache owner cannot fabricate diagnostics")
        return None
    if cache["unavailable_reason"] is not None:
        fail(f"{path}.unavailable_reason", "available cache has an unavailable reason")
    for field in CACHE_COUNTER_FIELDS:
        uint(cache[field], f"{path}.{field}")
    boolean(cache["budget_managed"], f"{path}.budget_managed")
    for field in CACHE_GAUGE_FIELDS:
        if field != "budget_managed":
            uint(cache[field], f"{path}.{field}")
    if cache["counter_delta_checked"] is not None:
        boolean(cache["counter_delta_checked"], f"{path}.counter_delta_checked")
    if cache["event_delta"] is None:
        if cache["counter_delta_checked"] is not None:
            fail(path, "counter-delta marker requires an event delta")
    else:
        event = obj(cache["event_delta"], f"{path}.event_delta")
        exact_keys(event, CACHE_COUNTER_FIELDS, (), f"{path}.event_delta")
        for field in CACHE_COUNTER_FIELDS:
            uint(event[field], f"{path}.event_delta.{field}")
        if cache["counter_delta_checked"] is not True:
            fail(f"{path}.counter_delta_checked", "event delta was not marked checked")
        text(cache["interval_label"], f"{path}.interval_label")
    if cache["interval_label"] is not None:
        text(cache["interval_label"], f"{path}.interval_label")
    return cache


def check_request_histogram(value: Any, logical_calls: int, minimum: int | None, maximum: int | None, path: str) -> list[int]:
    counts = array(value, path)
    if len(counts) != len(REQUEST_SIZE_BUCKETS):
        fail(path, f"expected exactly {len(REQUEST_SIZE_BUCKETS)} request-size buckets")
    result = [uint(item, f"{path}[{index}]") for index, item in enumerate(counts)]
    if sum(result) != logical_calls:
        fail(path, "request-size histogram does not sum to logical calls")
    nonempty = [index for index, count in enumerate(result) if count]
    if not nonempty:
        if logical_calls != 0 or minimum is not None or maximum is not None:
            fail(path, "empty request histogram has calls or request bounds")
        return result
    if logical_calls == 0 or minimum is None or maximum is None:
        fail(path, "nonempty request histogram omitted cumulative bounds")
    if minimum <= 0 or maximum < minimum:
        fail(path, "invalid cumulative request bounds")
    first_low, first_high = REQUEST_SIZE_BUCKETS[nonempty[0]]
    last_low, last_high = REQUEST_SIZE_BUCKETS[nonempty[-1]]
    if minimum < first_low or (first_high is not None and minimum > first_high):
        fail(path, "cumulative minimum is outside the first nonzero request bucket")
    if maximum < last_low or (last_high is not None and maximum > last_high):
        fail(path, "cumulative maximum is outside the last nonzero request bucket")
    return result


def check_interval_histogram(value: Any, logical_calls: int, path: str) -> list[int]:
    counts = array(value, path)
    if len(counts) != len(REQUEST_SIZE_BUCKETS):
        fail(path, f"expected exactly {len(REQUEST_SIZE_BUCKETS)} request-size buckets")
    result = [uint(item, f"{path}[{index}]") for index, item in enumerate(counts)]
    if sum(result) != logical_calls:
        fail(path, "request-size interval histogram does not sum to interval logical calls")
    return result


def check_read(value: Any, path: str, *, delta_has_bounds: bool) -> dict[str, Any]:
    read = obj(value, path)
    exact_keys(read, READ_FIELDS, (), path)
    state = text(read["availability"], f"{path}.availability")
    if state not in {"available", "unavailable"}:
        fail(f"{path}.availability", "unknown source-read availability state")
    if state == "unavailable":
        text(read["unavailable_reason"], f"{path}.unavailable_reason")
        for field in READ_FIELDS[2:]:
            if read[field] is not None:
                fail(f"{path}.{field}", "unavailable source cannot fabricate counters")
        return read
    if read["unavailable_reason"] is not None:
        fail(f"{path}.unavailable_reason", "available source-read point has an unavailable reason")
    for field in ("logical_calls", "requested_bytes", "returned_bytes", "min_request_bytes", "max_request_bytes", "short_reads", "delayed_calls"):
        optional_uint(read[field], f"{path}.{field}")
    if read["logical_calls"] is None or read["requested_bytes"] is None or read["returned_bytes"] is None:
        fail(path, "available source-read point omitted required cumulative counters")
    if read["returned_bytes"] > read["requested_bytes"]:
        fail(path, "returned bytes exceed requested bytes")
    if read["short_reads"] is None or read["short_reads"] > read["logical_calls"]:
        fail(path, "short-read count exceeds logical calls")
    if read["min_request_bytes"] is not None and read["max_request_bytes"] is not None:
        minimum_total = read["logical_calls"] * read["min_request_bytes"]
        maximum_total = read["logical_calls"] * read["max_request_bytes"]
        if not minimum_total <= read["requested_bytes"] <= maximum_total:
            fail(path, "requested-byte total is outside its cumulative request bounds")
    check_request_histogram(
        read["request_size_counts"],
        read["logical_calls"],
        read["min_request_bytes"],
        read["max_request_bytes"],
        f"{path}.request_size_counts",
    )
    marker = read["counter_delta_checked"]
    if marker is not None:
        boolean(marker, f"{path}.counter_delta_checked")
    if read["delta"] is None:
        if marker is not None:
            fail(path, "delta marker requires a delta object")
    else:
        delta = obj(read["delta"], f"{path}.delta")
        delta_fields = READ_DELTA_BOUNDED_FIELDS if delta_has_bounds else READ_DELTA_COMMON_FIELDS
        exact_keys(delta, delta_fields, (), f"{path}.delta")
        for field in ("logical_calls", "requested_bytes", "returned_bytes", "short_reads", "delayed_calls"):
            uint(delta[field], f"{path}.delta.{field}")
        # Interval min/max are intentionally not subtracted: the adapter
        # carries cumulative bounds on every checked delta.  The interval
        # histogram still has to account for every interval call.
        check_interval_histogram(
            delta["request_size_counts"],
            delta["logical_calls"],
            f"{path}.delta.request_size_counts",
        )
        if delta_has_bounds:
            optional_uint(delta["min_request_bytes"], f"{path}.delta.min_request_bytes")
            optional_uint(delta["max_request_bytes"], f"{path}.delta.max_request_bytes")
        if delta_has_bounds and (delta["min_request_bytes"] != read["min_request_bytes"] or delta["max_request_bytes"] != read["max_request_bytes"]):
            fail(f"{path}.delta", "delta request bounds are not the current cumulative bounds")
        if marker is not True:
            fail(f"{path}.counter_delta_checked", "read delta was not marked checked")
        if delta["returned_bytes"] > delta["requested_bytes"]:
            fail(f"{path}.delta", "delta returned bytes exceed requested bytes")
        if delta["short_reads"] > delta["logical_calls"]:
            fail(f"{path}.delta", "delta short reads exceed logical calls")
    return read


def check_rss(value: Any, path: str) -> None:
    rss = obj(value, path)
    exact_keys(rss, RSS_FIELDS, (), path)
    state = text(rss["availability"], f"{path}.availability")
    if state not in {"available", "unavailable"}:
        fail(f"{path}.availability", "unknown RSS state")
    if state == "available":
        if rss["unavailable_reason"] is not None:
            fail(f"{path}.unavailable_reason", "available RSS point has an unavailable reason")
        uint(rss["rss_bytes"], f"{path}.rss_bytes")
        uint(rss["vm_hwm_bytes"], f"{path}.vm_hwm_bytes")
    else:
        text(rss["unavailable_reason"], f"{path}.unavailable_reason")
        if rss["rss_bytes"] is not None or rss["vm_hwm_bytes"] is not None:
            fail(path, "unavailable RSS point cannot fabricate process bytes")


def check_gate_object(value: Any, path: str) -> None:
    gates = obj(value, path)
    exact_keys(gates, GATE_FIELDS, (), path)
    for field in GATE_FIELDS:
        if boolean(gates[field], f"{path}.{field}") is not True:
            fail(f"{path}.{field}", "correctness gate is false")


def check_phase_descriptions(value: Any, expected: tuple[str, ...], path: str, owner_tokens: dict[str, tuple[str, ...]] | None = None) -> None:
    phases = array(value, path)
    if tuple(text(obj(item, f"{path}[{i}]")["label"], f"{path}[{i}].label") for i, item in enumerate(phases)) != expected:
        fail(path, "phase labels do not match the frozen phase order")
    for i, item in enumerate(phases):
        phase = obj(item, f"{path}[{i}]")
        exact_keys(phase, PHASE_DESCRIPTION_FIELDS, (), f"{path}[{i}]")
        owners = text(phase["live_owners"], f"{path}[{i}].live_owners").lower()
        if owner_tokens is not None:
            for token in owner_tokens[expected[i]]:
                if token.lower() not in owners:
                    fail(f"{path}[{i}].live_owners", f"owner description omitted required evidence: {token}")


def check_manifest(value: Any, report: dict[str, Any]) -> None:
    # The provider corpus manifest describes the deterministic destination
    # package used as the cache/publication corpus.  The report keeps the
    # independently flattened source identity beside it.
    manifest = obj(value, "report.corpus_manifest")
    fields = (
        "name", "generator", "package_format", "shape", "payload_kind", "compression",
        "entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes",
        "archive_bytes", "archive_sha256", "target_entry", "target_payload_bytes",
        "target_payload_sha256", "xlsx",
    )
    exact_keys(manifest, fields, ("rtf_variant",), "report.corpus_manifest")
    for field in ("name", "generator", "package_format", "shape", "payload_kind", "compression", "target_entry"):
        text(manifest[field], f"report.corpus_manifest.{field}")
    for field in ("entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        uint(manifest[field], f"report.corpus_manifest.{field}")
    archive_hash = digest(manifest["archive_sha256"], "report.corpus_manifest.archive_sha256")
    if archive_hash != report["destination_archive_sha256"]:
        fail("report.corpus_manifest.archive_sha256", "manifest archive identity differs from destination archive")
    if manifest["archive_bytes"] != report["destination_archive_bytes"]:
        fail("report.corpus_manifest.archive_bytes", "manifest archive size differs from destination archive")
    digest(manifest["target_payload_sha256"], "report.corpus_manifest.target_payload_sha256")
    if manifest.get("rtf_variant") is not None:
        text(manifest["rtf_variant"], "report.corpus_manifest.rtf_variant")
    if manifest["xlsx"] is not None:
        xlsx = obj(manifest["xlsx"], "report.corpus_manifest.xlsx")
        exact_keys(xlsx, ("sheet_count", "rows_per_sheet", "columns_per_sheet", "one_percent_update_count", "source_members"), (), "report.corpus_manifest.xlsx")
        for field in ("sheet_count", "rows_per_sheet", "columns_per_sheet", "one_percent_update_count"):
            uint(xlsx[field], f"report.corpus_manifest.xlsx.{field}")
        members = obj(xlsx["source_members"], "report.corpus_manifest.xlsx.source_members")
        exact_keys(members, ("workbook", "worksheets", "shared_strings", "styles"), (), "report.corpus_manifest.xlsx.source_members")
        text(members["workbook"], "report.corpus_manifest.xlsx.source_members.workbook")
        array(members["worksheets"], "report.corpus_manifest.xlsx.source_members.worksheets")
        for i, item in enumerate(members["worksheets"]):
            text(item, f"report.corpus_manifest.xlsx.source_members.worksheets[{i}]")
        for field in ("shared_strings", "styles"):
            if members[field] is not None:
                text(members[field], f"report.corpus_manifest.xlsx.source_members.{field}")


def check_provider_config(value: Any, report: dict[str, Any]) -> None:
    config = obj(value, "report.provider_config")
    fields = (
        "provider", "max_range_bytes", "delay_us", "adapter_max_range_bytes",
        "adapter_delay_us", "delay_configured", "bytes_file_unlimited_cap",
        "file_scope", "range_scope",
    )
    exact_keys(config, fields, (), "report.provider_config")
    provider = text(report["provider"], "report.provider")
    if text(config["provider"], "report.provider_config.provider") != provider:
        fail("report.provider_config.provider", "provider configuration is mislabeled")
    if provider not in PROVIDERS:
        fail("report.provider", "unknown provider")
    max_range = optional_uint(config["max_range_bytes"], "report.provider_config.max_range_bytes")
    delay = optional_uint(config["delay_us"], "report.provider_config.delay_us")
    delay_configured = boolean(config["delay_configured"], "report.provider_config.delay_configured")
    unlimited = boolean(config["bytes_file_unlimited_cap"], "report.provider_config.bytes_file_unlimited_cap")
    adapter_max_range = optional_uint(config["adapter_max_range_bytes"], "report.provider_config.adapter_max_range_bytes")
    adapter_delay = optional_uint(config["adapter_delay_us"], "report.provider_config.adapter_delay_us")
    if adapter_max_range != max_range or adapter_delay != delay:
        fail("report.provider_config", "reported adapter limits differ from requested provider limits")
    if provider == "range":
        if max_range is None or max_range == 0 or max_range > 1_048_576 or delay is None or delay > 100_000:
            fail("report.provider_config", "range provider omitted a bounded cap or delay")
        if not delay_configured or unlimited:
            fail("report.provider_config", "range provider configuration flags are inconsistent")
    else:
        if max_range is not None or delay is not None or delay_configured is not False or unlimited is not True:
            fail("report.provider_config", "direct provider configuration flags are inconsistent")
    text(config["file_scope"], "report.provider_config.file_scope")
    text(config["range_scope"], "report.provider_config.range_scope")
    if config["file_scope"] != report["file_scope"] or config["range_scope"] != report["range_scope"]:
        fail("report.provider_config", "provider scope fields differ from report scopes")


def check_native_provider_config(value: Any, report: dict[str, Any]) -> tuple[int | None, bool]:
    config = obj(value, "report.provider_config")
    fields = ("provider", "max_range_bytes", "delay_us", "delay_configured", "file_scope", "range_scope")
    exact_keys(config, fields, (), "report.provider_config")
    provider = text(report["provider"], "report.provider")
    if text(config["provider"], "report.provider_config.provider") != provider:
        fail("report.provider_config.provider", "provider configuration is mislabeled")
    max_range = optional_uint(config["max_range_bytes"], "report.provider_config.max_range_bytes")
    delay = optional_uint(config["delay_us"], "report.provider_config.delay_us")
    delay_configured = boolean(config["delay_configured"], "report.provider_config.delay_configured")
    if provider == "range":
        if max_range is None or max_range == 0 or max_range > 1_048_576 or delay is None or delay > 100_000 or not delay_configured:
            fail("report.provider_config", "native range provider configuration is not bounded and enabled")
    elif provider in {"bytes", "file"}:
        if max_range is not None or delay is not None or delay_configured:
            fail("report.provider_config", "native direct provider configuration is inconsistent")
    else:
        fail("report.provider", "unknown provider")
    text(config["file_scope"], "report.provider_config.file_scope")
    text(config["range_scope"], "report.provider_config.range_scope")
    if config["file_scope"] != report["file_scope"] or config["range_scope"] != report["range_scope"]:
        fail("report.provider_config", "native provider scope fields differ from report scopes")
    return max_range, delay_configured


def check_common_identity(report: dict[str, Any], *, native: bool) -> tuple[int, int, dict[str, int]]:
    samples = uint(report["samples"], "report.samples")
    warmup = uint(report["warmup"], "report.warmup")
    if samples not in {1, 30}:
        fail("report.samples", "only one-sample controls and thirty-sample formal reports are valid")
    if warmup not in {0, 3}:
        fail("report.warmup", "only zero-warmup controls and three-warmup formal reports are valid")
    if (samples, warmup) not in {(1, 0), (30, 3)}:
        fail("report.warmup", "control reports require (samples=1,warmup=0) and formal reports require (samples=30,warmup=3)")
    checked = uint(report["checked_iteration_count"], "report.checked_iteration_count")
    if checked != samples + warmup:
        fail("report.checked_iteration_count", "does not equal samples plus warmups")
    revision = text(report["source_revision"], "report.source_revision")
    if not HEX40.fullmatch(revision):
        fail("report.source_revision", "expected forty lowercase hexadecimal characters")
    digest(report["binary_sha256"], "report.binary_sha256")
    uint(report["binary_bytes"], "report.binary_bytes")
    text(report["current_exe"], "report.current_exe")
    limits = check_limits(report["configured_limits"], "report.configured_limits")
    return samples, warmup, limits


def check_budget_limits(budget: dict[str, int], limits: dict[str, int], path: str) -> None:
    for field in LIMIT_FIELDS[2:]:
        if budget[field] != limits[field]:
            fail(f"{path}.{field}", "budget limit differs from configured limits")


def check_budget_sequence(points: list[tuple[str, dict[str, int], str]]) -> None:
    previous: dict[str, int] | None = None
    for label, current, path in points:
        if previous is not None:
            for field in ("input_bytes_used", "output_bytes_used", "work_used"):
                if current[field] < previous[field]:
                    fail(f"{path}.{field}", "cumulative budget charge moved backwards")
        previous = current


def check_cache_limits(cache: dict[str, Any], limits: dict[str, int], path: str) -> None:
    for field, limit_field in (
        ("retained_entries", "cache_max_entries"), ("retained_bytes", "cache_max_bytes"),
        ("budget_cache_reserved_bytes", "cache_max_bytes"),
        ("budget_catalog_reserved_objects", "objects_limit"),
        ("budget_cache_reserved_objects", "objects_limit"),
        ("budget_memory_limit", "memory_limit"), ("budget_input_bytes_limit", "input_bytes_limit"),
        ("budget_output_bytes_limit", "output_bytes_limit"), ("budget_work_limit", "work_limit"),
        ("budget_objects_limit", "objects_limit"),
    ):
        if field.endswith("_limit"):
            if cache[field] != limits[limit_field]:
                fail(f"{path}.{field}", "cache diagnostic limit differs from configured limits")
        elif cache[field] > limits[limit_field]:
            fail(f"{path}.{field}", "cache gauge exceeds configured limits")
    for used, limit in (
        ("budget_memory_used", "budget_memory_limit"), ("budget_input_bytes_used", "budget_input_bytes_limit"),
        ("budget_output_bytes_used", "budget_output_bytes_limit"), ("budget_work_used", "budget_work_limit"),
        ("budget_objects_used", "budget_objects_limit"),
    ):
        if cache[used] > cache[limit]:
            fail(f"{path}.{used}", "cache gauge exceeds its diagnostic limit")
    if cache["budget_cache_reserved_bytes"] > cache["budget_memory_used"]:
        fail(path, "cache reservation exceeds managed memory")
    if cache["budget_cache_reserved_objects"] > cache["budget_objects_used"]:
        fail(path, "cache reservation exceeds managed objects")
    if cache["budget_catalog_reserved_objects"] > cache["budget_objects_used"]:
        fail(path, "catalog reservation exceeds managed objects")


def check_cache_sequence(points: list[tuple[str, dict[str, Any] | None, str]], limits: dict[str, int]) -> int:
    previous: dict[str, Any] | None = None
    checked = 0
    for label, current, path in points:
        if current is None:
            previous = None
            continue
        if current["budget_managed"] is not True or current["in_flight_loads"] != 0:
            fail(path, "available cache owner is not a quiescent managed cache")
        check_cache_limits(current, limits, path)
        if previous is None:
            if current["event_delta"] is not None or current["counter_delta_checked"] is not None:
                fail(path, "first available cache point cannot contain an interval delta")
        else:
            for field in CACHE_COUNTER_FIELDS:
                if current[field] < previous[field]:
                    fail(f"{path}.{field}", "cache counter moved backwards")
            event = current["event_delta"]
            if event is None:
                fail(path, "available cache interval omitted checked counter delta")
            checked += 1
            for field in CACHE_COUNTER_FIELDS:
                if event[field] != current[field] - previous[field]:
                    fail(f"{path}.event_delta.{field}", "cache counter delta does not match adjacent points")
        previous = current
    return checked


def check_read_sequence(points: list[tuple[str, dict[str, Any]]], *, max_range: int | None, delay_configured: bool, delta_has_bounds: bool) -> int:
    previous: dict[str, Any] | None = None
    checked = 0
    for label, current in points:
        if current["availability"] == "unavailable":
            previous = None
            continue
        path = f"phase[{label}].source_reads"
        if max_range is not None and current["returned_bytes"] > current["logical_calls"] * max_range:
            fail(path, "returned bytes exceed configured range cap")
        if current["delayed_calls"] > current["logical_calls"]:
            fail(path, "delayed calls exceed logical calls")
        if current["delayed_calls"] != (current["logical_calls"] if delay_configured else 0):
            fail(path, "delayed-call count differs from the configured adapter delay")
        if previous is not None:
            for field in READ_COUNTER_FIELDS:
                if current[field] < previous[field]:
                    fail(f"{path}.{field}", "read counter moved backwards")
            for field in ("min_request_bytes", "max_request_bytes"):
                if current[field] is None or previous[field] is None:
                    continue
                if field == "min_request_bytes" and current[field] > previous[field]:
                    fail(f"{path}.{field}", "cumulative minimum request bound increased")
                if field == "max_request_bytes" and current[field] < previous[field]:
                    fail(f"{path}.{field}", "cumulative maximum request bound decreased")
            delta = current["delta"]
            if delta is None:
                fail(path, "available read interval omitted its checked delta")
            checked += 1
            for field in READ_COUNTER_FIELDS:
                if delta[field] != current[field] - previous[field]:
                    fail(f"{path}.delta.{field}", "read delta does not match adjacent snapshots")
            previous_buckets = previous["request_size_counts"]
            current_buckets = current["request_size_counts"]
            delta_buckets = delta["request_size_counts"]
            for bucket, (before, after, interval) in enumerate(zip(previous_buckets, current_buckets, delta_buckets)):
                if after < before:
                    fail(f"{path}.request_size_counts[{bucket}]", "request-size counter moved backwards")
                if interval != after - before:
                    fail(f"{path}.delta.request_size_counts[{bucket}]", "request-size delta does not match adjacent snapshots")
            if delta_has_bounds and (delta["min_request_bytes"] != current["min_request_bytes"] or delta["max_request_bytes"] != current["max_request_bytes"]):
                fail(f"{path}.delta", "delta request bounds are not cumulative after-snapshot bounds")
        else:
            # Every first available point is measured against the explicit
            # all-zero adapter baseline.  This makes the first retained
            # interval independently checkable after an owner drop.
            delta = current["delta"]
            if delta is None:
                fail(path, "first available read point omitted its checked zero-base delta")
            for field in READ_COUNTER_FIELDS:
                if delta[field] != current[field]:
                    fail(f"{path}.delta.{field}", "first read delta is not measured from the zero baseline")
            for bucket, (count, interval) in enumerate(zip(current["request_size_counts"], delta["request_size_counts"])):
                if interval != count:
                    fail(f"{path}.delta.request_size_counts[{bucket}]", "first request-size delta is not measured from the zero baseline")
            if delta_has_bounds and (delta["min_request_bytes"] != current["min_request_bytes"] or delta["max_request_bytes"] != current["max_request_bytes"]):
                fail(f"{path}.delta", "first delta request bounds differ from the cumulative snapshot")
            checked += 1
        previous = current
    return checked


def check_phase_budget_pair(phase: dict[str, Any], limits: dict[str, int], path: str, source_cache: dict[str, Any] | None, destination: bool = False) -> None:
    source_budget = check_budget(phase["destination_budget" if destination else "source_budget"], path)
    check_budget_limits(source_budget, limits, path)
    if source_cache is not None:
        for resource in ("memory", "objects", "input_bytes", "output_bytes", "work"):
            if source_cache[f"budget_{resource}_used"] != source_budget[f"{resource}_used"]:
                fail(path, "available cache gauges do not match their caller budget")


def check_provider_phases(report: dict[str, Any], limits: dict[str, int], destination_limits: dict[str, int]) -> tuple[int, int]:
    rows = array(report["samples_raw"], "report.samples_raw")
    if len(rows) != report["samples"]:
        fail("report.samples_raw", "row count differs from samples")
    total_cache_deltas = 0
    total_read_deltas = 0
    for index, raw in enumerate(rows):
        row = obj(raw, f"report.samples_raw[{index}]")
        path = f"report.samples_raw[{index}]"
        exact_keys(row, ("sample_index", "exact_output_verified", "output_sha256", "output_bytes", "timings", "phases"), (), path)
        if uint(row["sample_index"], f"{path}.sample_index") != index:
            fail(f"{path}.sample_index", "sample index is not contiguous")
        if boolean(row["exact_output_verified"], f"{path}.exact_output_verified") is not True:
            fail(f"{path}.exact_output_verified", "exact output gate is false")
        if digest(row["output_sha256"], f"{path}.output_sha256") != report["expected_output_sha256"]:
            fail(f"{path}.output_sha256", "row output identity differs from expected output")
        if uint(row["output_bytes"], f"{path}.output_bytes") != report["expected_output_bytes"]:
            fail(f"{path}.output_bytes", "row output size differs from expected output")
        timing = obj(row["timings"], f"{path}.timings")
        exact_keys(timing, ("open_source_ns", "open_destination_ns", "open_ns", "plan_ns", "publication_ns", "api_sum_ns"), (), f"{path}.timings")
        for field in timing:
            uint(timing[field], f"{path}.timings.{field}")
        if timing["open_ns"] != timing["open_source_ns"] + timing["open_destination_ns"]:
            fail(f"{path}.timings.open_ns", "open duration is not the checked API sum")
        if timing["api_sum_ns"] != timing["open_ns"] + timing["plan_ns"] + timing["publication_ns"]:
            fail(f"{path}.timings.api_sum_ns", "API duration sum is inconsistent")
        phases = array(row["phases"], f"{path}.phases")
        if len(phases) != len(PROVIDER_PHASES):
            fail(f"{path}.phases", "wrong phase count")
        source_cache_points: list[tuple[str, dict[str, Any] | None, str]] = []
        destination_cache_points: list[tuple[str, dict[str, Any] | None, str]] = []
        source_reads: list[tuple[str, dict[str, Any]]] = []
        destination_reads: list[tuple[str, dict[str, Any]]] = []
        source_budgets: list[tuple[str, dict[str, int], str]] = []
        destination_budgets: list[tuple[str, dict[str, int], str]] = []
        for pindex, raw_phase in enumerate(phases):
            phase = obj(raw_phase, f"{path}.phases[{pindex}]")
            exact_keys(phase, PROVIDER_PHASE_FIELDS, (), f"{path}.phases[{pindex}]")
            label = text(phase["label"], f"{path}.phases[{pindex}].label")
            if label != PROVIDER_PHASES[pindex]:
                fail(f"{path}.phases[{pindex}].label", "phase order differs from the frozen provider journal")
            source = check_cache(phase["source_cache"], f"{path}.phases[{pindex}].source_cache")
            destination = check_cache(phase["destination_cache"], f"{path}.phases[{pindex}].destination_cache")
            expected_source = label in {"opened", "planned", "published", "drop_result", "drop_plan"}
            expected_destination = label in {"opened", "planned"}
            if (source is not None) != expected_source:
                fail(f"{path}.phases[{pindex}].source_cache", "source cache availability differs from owner lifetime")
            if (destination is not None) != expected_destination:
                fail(f"{path}.phases[{pindex}].destination_cache", "destination cache availability differs from owner lifetime")
            source_cache_points.append((label, source, f"{path}.phases[{pindex}].source_cache"))
            destination_cache_points.append((label, destination, f"{path}.phases[{pindex}].destination_cache"))
            source_read = check_read(phase["source_reads"], f"{path}.phases[{pindex}].source_reads", delta_has_bounds=True)
            destination_read = check_read(phase["destination_reads"], f"{path}.phases[{pindex}].destination_reads", delta_has_bounds=True)
            expected_read = label in {"opened", "planned", "published", "drop_result", "drop_plan", "drop_view"}
            if (source_read["availability"] == "available") != expected_read:
                fail(f"{path}.phases[{pindex}].source_reads", "source-read availability differs from caller lifetime")
            if (destination_read["availability"] == "available") != expected_read:
                fail(f"{path}.phases[{pindex}].destination_reads", "destination-read availability differs from caller lifetime")
            source_reads.append((label, source_read))
            destination_reads.append((label, destination_read))
            source_budget = check_budget(phase["source_budget"], f"{path}.phases[{pindex}].source_budget")
            destination_budget = check_budget(phase["destination_budget"], f"{path}.phases[{pindex}].destination_budget")
            source_budgets.append((label, source_budget, f"{path}.phases[{pindex}].source_budget"))
            destination_budgets.append((label, destination_budget, f"{path}.phases[{pindex}].destination_budget"))
            check_budget_limits(source_budget, limits, f"{path}.phases[{pindex}].source_budget")
            check_budget_limits(destination_budget, destination_limits, f"{path}.phases[{pindex}].destination_budget")
            if source is not None:
                for resource in ("memory", "objects", "input_bytes", "output_bytes", "work"):
                    if source[f"budget_{resource}_used"] != source_budget[f"{resource}_used"]:
                        fail(f"{path}.phases[{pindex}].source_cache", "cache gauges differ from source budget")
            if destination is not None:
                for resource in ("memory", "objects", "input_bytes", "output_bytes", "work"):
                    if destination[f"budget_{resource}_used"] != destination_budget[f"{resource}_used"]:
                        fail(f"{path}.phases[{pindex}].destination_cache", "cache gauges differ from destination budget")
            check_rss(phase["rss"], f"{path}.phases[{pindex}].rss")
        total_cache_deltas += check_cache_sequence(source_cache_points, limits)
        total_cache_deltas += check_cache_sequence(destination_cache_points, destination_limits)
        check_budget_sequence(source_budgets)
        check_budget_sequence(destination_budgets)
        max_range = None
        if report["provider"] == "range":
            max_range = uint(report["provider_config"]["max_range_bytes"], "report.provider_config.max_range_bytes")
        delay_configured = report["provider"] == "range"
        total_read_deltas += check_read_sequence(source_reads, max_range=max_range, delay_configured=delay_configured, delta_has_bounds=True)
        total_read_deltas += check_read_sequence(destination_reads, max_range=max_range, delay_configured=delay_configured, delta_has_bounds=True)
        final_source = check_budget(phases[-1]["source_budget"], f"{path}.phases[-1].source_budget")
        final_destination = check_budget(phases[-1]["destination_budget"], f"{path}.phases[-1].destination_budget")
        for final, location in ((final_source, "source_budget"), (final_destination, "destination_budget")):
            if final["memory_used"] != 0 or final["objects_used"] != 0 or final["depth_used"] != 0:
                fail(f"{path}.phases[-1].{location}", "releasable managed gauges did not return to zero")
    if not total_cache_deltas or not total_read_deltas:
        fail("report.samples_raw", "no checked cache or source-read deltas were retained")
    return total_cache_deltas, total_read_deltas


def check_provider_report(report: Any) -> dict[str, Any]:
    report = obj(report, "report")
    check_all_numbers(report, "report")
    check_hash_fields(report, "report")
    required = (
        "schema", "corpus", "provider", "provider_scope", "timing_scope", "range_scope", "file_scope", "rss_scope",
        "samples", "warmup", "checked_iteration_count", "source_revision", "binary_sha256", "binary_bytes", "current_exe",
        "source_archive_sha256", "source_archive_bytes", "destination_archive_sha256", "destination_archive_bytes",
        "expected_output_sha256", "expected_output_bytes", "corpus_manifest", "gates", "provider_config",
        "configured_limits", "destination_configured_limits", "destination_editor_consumed_during_publish",
        "final_memory_objects_depth_zero_checked", "phases", "samples_raw",
    )
    exact_keys(report, required, (), "report")
    if text(report["schema"], "report.schema") != PROVIDER_SCHEMA:
        fail("report.schema", f"expected {PROVIDER_SCHEMA!r}")
    if text(report["corpus"], "report.corpus") not in CORPORA:
        fail("report.corpus", "unknown synthetic corpus")
    if text(report["provider"], "report.provider") not in PROVIDERS:
        fail("report.provider", "unknown source provider")
    for field in ("provider_scope", "timing_scope", "range_scope", "file_scope", "rss_scope"):
        check_scope(report[field], f"report.{field}")
    digest(report["source_archive_sha256"], "report.source_archive_sha256")
    digest(report["destination_archive_sha256"], "report.destination_archive_sha256")
    digest(report["expected_output_sha256"], "report.expected_output_sha256")
    for field in ("source_archive_bytes", "destination_archive_bytes", "expected_output_bytes"):
        uint(report[field], f"report.{field}")
    samples, _warmup, limits = check_common_identity(report, native=False)
    destination_limits = check_limits(report["destination_configured_limits"], "report.destination_configured_limits")
    for field in LIMIT_FIELDS:
        if limits[field] != destination_limits[field]:
            fail(f"report.destination_configured_limits.{field}", "source and destination limits differ")
    check_provider_config(report["provider_config"], report)
    check_manifest(report["corpus_manifest"], report)
    check_gate_object(report["gates"], "report.gates")
    check_phase_descriptions(report["phases"], PROVIDER_PHASES, "report.phases")
    if boolean(report["destination_editor_consumed_during_publish"], "report.destination_editor_consumed_during_publish") is not True:
        fail("report.destination_editor_consumed_during_publish", "consuming publication boundary is not recorded")
    if boolean(report["final_memory_objects_depth_zero_checked"], "report.final_memory_objects_depth_zero_checked") is not True:
        fail("report.final_memory_objects_depth_zero_checked", "final gauge check is not recorded")
    cache_deltas, read_deltas = check_provider_phases(report, limits, destination_limits)
    return {"status": "valid", "schema": PROVIDER_SCHEMA, "corpus": report["corpus"], "provider": report["provider"], "samples": samples, "checked_cache_deltas": cache_deltas, "checked_read_deltas": read_deltas}


def check_native_oracle(report: dict[str, Any]) -> dict[str, Any]:
    fixture = text(report["fixture"], "report.fixture")
    if fixture not in FIXTURES:
        fail("report.fixture", "unknown native fixture")
    expected = NATIVE_ORACLES[fixture]
    payload = obj(report["payload_oracle"], "report.payload_oracle")
    exact_keys(
        payload,
        (
            "slide", "image", "image_count", "shape_position", "shape_id", "shape_name",
            "bounds", "relationship_id", "part", "content_type", "payload_bytes", "payload_sha256",
        ),
        (),
        "report.payload_oracle",
    )
    for field in ("slide", "image", "image_count", "shape_position", "shape_id", "payload_bytes"):
        uint(payload[field], f"report.payload_oracle.{field}")
    text(payload["shape_name"], "report.payload_oracle.shape_name")
    text(payload["content_type"], "report.payload_oracle.content_type")
    expected_bounds = expected["bounds"]
    if expected_bounds is None:
        if payload["bounds"] is not None:
            fail("report.payload_oracle.bounds", "native oracle unexpectedly supplied bounds")
    else:
        bounds = obj(payload["bounds"], "report.payload_oracle.bounds")
        exact_keys(bounds, ("x", "y", "width", "height"), (), "report.payload_oracle.bounds")
        for field in ("x", "y", "width", "height"):
            uint(bounds[field], f"report.payload_oracle.bounds.{field}")
    text(payload["relationship_id"], "report.payload_oracle.relationship_id")
    text(payload["part"], "report.payload_oracle.part")
    digest(payload["payload_sha256"], "report.payload_oracle.payload_sha256")
    for field in (
        "slide", "image", "image_count", "shape_position", "shape_id", "shape_name", "bounds",
        "relationship_id", "part", "content_type", "payload_bytes", "payload_sha256",
    ):
        if payload[field] != expected[field]:
            fail(f"report.payload_oracle.{field}", "payload oracle differs from the frozen native oracle")
    for field, expected_value in (("native_fixture_sha256", expected["archive_sha256"]), ("native_fixture_bytes", expected["archive_bytes"]), ("native_fixture_path", expected["path"]), ("native_fixture_name", expected["name"]), ("native_fixture_resave_scope", expected["resave_scope"])):
        if field.endswith("sha256"):
            digest(report[field], f"report.{field}")
        elif field.endswith("bytes"):
            uint(report[field], f"report.{field}")
        else:
            text(report[field], f"report.{field}")
        if report[field] != expected_value:
            fail(f"report.{field}", "flattened native fixture identity differs from oracle")
    return expected


def check_native_phases(report: dict[str, Any], limits: dict[str, int], *, max_range: int | None, delay_configured: bool) -> tuple[int, int]:
    rows = array(report["samples_raw"], "report.samples_raw")
    if len(rows) != report["samples"]:
        fail("report.samples_raw", "row count differs from samples")
    total_cache = 0
    total_reads = 0
    expected_cache = {"opened", "selected", "loaded"}
    expected_read = {"opened", "selected", "loaded", "drop_view", "drop_slide", "drop_image"}
    for index, raw in enumerate(rows):
        row = obj(raw, f"report.samples_raw[{index}]")
        path = f"report.samples_raw[{index}]"
        exact_keys(row, ("sample_index", "timings", "phases", "descriptor_verified", "returned_descriptor_equal", "payload", "final_memory_objects_depth_zero", "source_owner_release_verified"), (), path)
        if uint(row["sample_index"], f"{path}.sample_index") != index:
            fail(f"{path}.sample_index", "sample index is not contiguous")
        timing = obj(row["timings"], f"{path}.timings")
        exact_keys(timing, ("open_ns", "metadata_ns", "read_ns", "api_sum_ns"), (), f"{path}.timings")
        for field in timing:
            uint(timing[field], f"{path}.timings.{field}")
        if timing["api_sum_ns"] != timing["open_ns"] + timing["metadata_ns"] + timing["read_ns"]:
            fail(f"{path}.timings.api_sum_ns", "API duration sum is inconsistent")
        if boolean(row["descriptor_verified"], f"{path}.descriptor_verified") is not True:
            fail(f"{path}.descriptor_verified", "descriptor oracle is false")
        if boolean(row["returned_descriptor_equal"], f"{path}.returned_descriptor_equal") is not True:
            fail(f"{path}.returned_descriptor_equal", "returned descriptor differs from selected descriptor")
        payload = obj(row["payload"], f"{path}.payload")
        exact_keys(payload, ("bytes", "sha256", "verified", "verified_after_slide_drop"), (), f"{path}.payload")
        expected = NATIVE_ORACLES[report["fixture"]]
        if uint(payload["bytes"], f"{path}.payload.bytes") != expected["payload_bytes"]:
            fail(f"{path}.payload.bytes", "payload size differs from native oracle")
        if digest(payload["sha256"], f"{path}.payload.sha256") != expected["payload_sha256"]:
            fail(f"{path}.payload.sha256", "payload identity differs from native oracle")
        if boolean(payload["verified"], f"{path}.payload.verified") is not True or boolean(payload["verified_after_slide_drop"], f"{path}.payload.verified_after_slide_drop") is not True:
            fail(f"{path}.payload", "returned image was not verified through owner drops")
        if boolean(row["final_memory_objects_depth_zero"], f"{path}.final_memory_objects_depth_zero") is not True or boolean(row["source_owner_release_verified"], f"{path}.source_owner_release_verified") is not True:
            fail(path, "final owner release gates are false")
        phases = array(row["phases"], f"{path}.phases")
        if len(phases) != len(NATIVE_PHASES):
            fail(f"{path}.phases", "wrong native phase count")
        cache_points: list[tuple[str, dict[str, Any] | None, str]] = []
        read_points: list[tuple[str, dict[str, Any]]] = []
        budget_points: list[tuple[str, dict[str, int], str]] = []
        for pindex, raw_phase in enumerate(phases):
            phase = obj(raw_phase, f"{path}.phases[{pindex}]")
            exact_keys(phase, NATIVE_PHASE_FIELDS, (), f"{path}.phases[{pindex}]")
            label = text(phase["label"], f"{path}.phases[{pindex}].label")
            if label != NATIVE_PHASES[pindex]:
                fail(f"{path}.phases[{pindex}].label", "native phase order differs from frozen journal")
            cache = check_cache(phase["source_cache"], f"{path}.phases[{pindex}].source_cache")
            read = check_read(phase["source_reads"], f"{path}.phases[{pindex}].source_reads", delta_has_bounds=False)
            if (cache is not None) != (label in expected_cache):
                fail(f"{path}.phases[{pindex}].source_cache", "native cache availability differs from owner lifetime")
            if (read["availability"] == "available") != (label in expected_read):
                fail(f"{path}.phases[{pindex}].source_reads", "native source-read availability differs from owner lifetime")
            budget = check_budget(phase["source_budget"], f"{path}.phases[{pindex}].source_budget")
            check_budget_limits(budget, limits, f"{path}.phases[{pindex}].source_budget")
            budget_points.append((label, budget, f"{path}.phases[{pindex}].source_budget"))
            if cache is not None:
                for resource in ("memory", "objects", "input_bytes", "output_bytes", "work"):
                    if cache[f"budget_{resource}_used"] != budget[f"{resource}_used"]:
                        fail(f"{path}.phases[{pindex}].source_cache", "cache gauges differ from caller budget")
            check_rss(phase["rss"], f"{path}.phases[{pindex}].rss")
            cache_points.append((label, cache, f"{path}.phases[{pindex}].source_cache"))
            read_points.append((label, read))
        total_cache += check_cache_sequence(cache_points, limits)
        check_budget_sequence(budget_points)
        total_reads += check_read_sequence(read_points, max_range=max_range, delay_configured=delay_configured, delta_has_bounds=False)
        final = check_budget(phases[-1]["source_budget"], f"{path}.phases[-1].source_budget")
        if final["memory_used"] != 0 or final["objects_used"] != 0 or final["depth_used"] != 0:
            fail(f"{path}.phases[-1].source_budget", "native releasable gauges did not return to zero")
    if not total_cache or not total_reads:
        fail("report.samples_raw", "no checked native cache or read deltas were retained")
    return total_cache, total_reads


def check_native_report(report: Any) -> dict[str, Any]:
    report = obj(report, "report")
    check_all_numbers(report, "report")
    check_hash_fields(report, "report")
    required = (
        "schema", "fixture", "provider", "native_fixture_sha256", "native_fixture_bytes",
        "native_fixture_path", "native_fixture_name", "native_fixture_resave_scope", "payload_oracle",
        "source_snapshot", "provider_scope", "timing_scope", "file_scope", "range_scope", "rss_scope",
        "samples", "warmup", "checked_iteration_count", "source_revision", "binary_sha256", "binary_bytes",
        "current_exe", "provider_config", "configured_limits", "phases", "samples_raw",
    )
    exact_keys(report, required, (), "report")
    if text(report["schema"], "report.schema") != NATIVE_SCHEMA:
        fail("report.schema", f"expected {NATIVE_SCHEMA!r}")
    fixture = text(report["fixture"], "report.fixture")
    if fixture not in FIXTURES:
        fail("report.fixture", "unknown native fixture")
    provider = text(report["provider"], "report.provider")
    if provider not in PROVIDERS:
        fail("report.provider", "unknown provider")
    for field in ("provider_scope", "timing_scope", "file_scope", "range_scope", "rss_scope"):
        check_scope(report[field], f"report.{field}")
    samples, _warmup, limits = check_common_identity(report, native=True)
    max_range, delay_configured = check_native_provider_config(report["provider_config"], report)
    source_snapshot = obj(report["source_snapshot"], "report.source_snapshot")
    exact_keys(source_snapshot, ("adapter", "scope", "checked_phase_delta", "unavailable_after_caller_source_drop"), (), "report.source_snapshot")
    if text(source_snapshot["adapter"], "report.source_snapshot.adapter") != "PptxRangeSource":
        fail("report.source_snapshot.adapter", "unexpected source adapter")
    check_scope(source_snapshot["scope"], "report.source_snapshot.scope")
    if boolean(source_snapshot["checked_phase_delta"], "report.source_snapshot.checked_phase_delta") is not True or boolean(source_snapshot["unavailable_after_caller_source_drop"], "report.source_snapshot.unavailable_after_caller_source_drop") is not True:
        fail("report.source_snapshot", "source snapshot contract flags are false")
    check_native_oracle(report)
    check_phase_descriptions(report["phases"], NATIVE_PHASES, "report.phases", NATIVE_OWNER_TOKENS)
    cache_deltas, read_deltas = check_native_phases(report, limits, max_range=max_range, delay_configured=delay_configured)
    return {"status": "valid", "schema": NATIVE_SCHEMA, "fixture": fixture, "provider": provider, "samples": samples, "checked_cache_deltas": cache_deltas, "checked_read_deltas": read_deltas}


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicate_pairs, parse_constant=reject_nonfinite_constant)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise VerificationError(f"{path}: cannot read JSON: {exc}") from exc


def check_report(report: Any) -> dict[str, Any]:
    root = obj(report, "report")
    schema = root.get("schema")
    if schema == PROVIDER_SCHEMA:
        return check_provider_report(root)
    if schema == NATIVE_SCHEMA:
        return check_native_report(root)
    fail("report.schema", "unknown or missing change-0429 schema")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", nargs="?", type=Path)
    parser.add_argument("--report", dest="report_option", type=Path)
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
