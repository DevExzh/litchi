#!/usr/bin/env python3
"""Independently validate and summarize the 0483 DOCX tail comparison.

The Rust harness supplies observations and route-specific proofs.  This module
does not trust producer summaries: every sample vector, corpus digest,
histogram, conservation equation, confidence interval, percentile and review
flag is recomputed from retained JSON.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import statistics
import subprocess
import zlib
from pathlib import Path
from typing import Any, Mapping


ROOT = Path(__file__).resolve().parent
DRIVER_SCRIPTS = (
    "common.py", "gate.py", "build.py", "capture.py", "analyze.py",
    "verify.py", "test_evidence.py", "fuzz.py", "fuzz-seeds.py",
    "consumer-probe.py",
)
PROTOCOL_SCHEMA = "docx-tail-append-comparison-v1"
REPORT_SCHEMAS = {"docx-bounded-tail-append-comparison-v1"}
SUMMARY_SCHEMA = "docx-tail-append-comparison-summary-v1"
COUNTS = (64, 8_192, 131_072)
INSTRUMENTATIONS = ("normal", "allocator")
ARMS = ("a1", "b1", "b2", "a2")
REPEATS = (1, 2)
ROUTES = ("materialized", "bounded")
ROUTE_NAMES = {
    "materialized": "materialized_paragraph_copy",
    "bounded": "bounded_plain_text_tail_append",
    "materialized_paragraph_copy": "materialized_paragraph_copy",
    "bounded_plain_text_tail_append": "bounded_plain_text_tail_append",
}
SAMPLES = 30
WARMUPS = 3
CPU = 2
REGRESSION_REVIEW_PERCENT = 5
MAIN_PATH = "word/document.xml"
OPAQUE_PATH = "word/perf-opaque.bin"
OPAQUE_BYTES = 32 * 1024
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
GENERATOR = "litchi-docx-bounded-tail-append-comparison-v1"
FORMAT = "DOCX/OOXML/OPC/ZIP"
HASH_SINK_MAX_WRITE = 16 * 1024
SINK_ID = "non_seek_hashing_scalar_sha256_shortwrite_16k"
SOURCE_ID = "caller_owned_arc_positional_read_at_scalar_counters_requested_returned_fixed_histogram"
SHA256 = re.compile(r"^[0-9a-f]{64}$")
LABEL = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
HIST_FIELDS = (
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
)
READ_FIELDS = ("calls", "requested_bytes", "returned_bytes", "request_histogram")
SINK_FIELDS = ("accepted_bytes", "write_calls", "largest_write", "histogram", "sha256")
ALLOC_FIELDS = (
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
PROCESS_FIELDS = (
    "rchar",
    "wchar",
    "read_bytes",
    "write_bytes",
    "cancelled_write_bytes",
    "syscr",
    "syscw",
    "minor_faults",
    "major_faults",
    "user_cpu_ticks",
    "system_cpu_ticks",
    "clock_ticks_per_second",
    "voluntary_context_switches",
    "nonvoluntary_context_switches",
    "rss_bytes",
    "peak_rss_bytes",
)


class AnalysisError(ValueError):
    """Retained evidence is malformed or contradicts the frozen protocol."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=lambda value: (_ for _ in ()).throw(ValueError(value)),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
        stream.write("\n")


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def integer(value: Any, label: str, expected: int | None = None) -> int:
    if not isinstance(value, int) or isinstance(value, bool) or value < 0:
        fail(f"{label}: expected a non-negative integer")
    if expected is not None and value != expected:
        fail(f"{label}: expected {expected}, got {value}")
    return value


def positive(value: Any, label: str) -> int:
    result = integer(value, label)
    if result == 0:
        fail(f"{label}: expected a positive integer")
    return result


def digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or SHA256.fullmatch(value) is None:
        fail(f"{label}: expected a lower-case SHA-256 digest")
    return value


def bool_true(value: Any, label: str) -> None:
    if value is not True:
        fail(f"{label}: expected true")


def source_xml(count: int) -> bytes:
    prefix = (
        '<?xml version="1.0" encoding="UTF-8"?><w:document '
        f'xmlns:w="{WORD_NAMESPACE}"><w:body>'
    ).encode()
    text = xml_escape(append_text(count))
    fragment = paragraph_xml(text)
    fragment += b"".join(
        paragraph_xml(f"paragraph-{index:06}") for index in range(1, count)
    )
    return prefix + fragment + b"</w:body></w:document>"


def candidate_xml(count: int) -> bytes:
    source = source_xml(count)
    marker = b"</w:body></w:document>"
    return source[: -len(marker)] + paragraph_xml(xml_escape(append_text(count))) + marker


def bounded_candidate_xml(count: int) -> bytes:
    """Return the bounded route's independent lexical XML oracle.

    The two routes have identical paragraph semantics, but the bounded route
    authors a fragment with a local Word namespace and an explicit preserve
    space policy. Keeping this raw oracle separate prevents the verifier from
    requiring byte-identical main XML from semantically equal routes.
    """

    source = source_xml(count)
    marker = b"</w:body></w:document>"
    fragment = (
        f'<w:p xmlns:w="{WORD_NAMESPACE}"><w:r><w:t '
        f'xml:space="preserve">{xml_escape(append_text(count))}'
        "</w:t></w:r></w:p>"
    ).encode()
    return source[: -len(marker)] + fragment + marker


def paragraph_xml(index: int) -> bytes:
    return f"<w:p><w:r><w:t>{index}</w:t></w:r></w:p>".encode()


def append_text(count: int) -> str:
    return f"tail-text-{count:06}-café <&> plain"


def xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def opaque_payload() -> bytes:
    return bytes(
        (
            ((index & 0xFF) * 37)
            + ((index >> 8) & 0xFF)
        )
        & 0xFF
        for index in range(OPAQUE_BYTES)
    )


def semantic(count: int, candidate: bool = False) -> dict[str, Any]:
    values: list[str] = [append_text(count)] + [f"paragraph-{index:06}" for index in range(1, count)]
    if candidate:
        values.append(append_text(count))
    order = hashlib.sha256(b"litchi-docx-bounded-tail-order-v1\0")
    text = hashlib.sha256(b"litchi-docx-bounded-tail-text-v1\0")
    order.update(len(values).to_bytes(8, "little"))
    text.update(len(values).to_bytes(8, "little"))
    text_bytes = 0
    for position, value in enumerate(values):
        value_text = value.encode()
        text_bytes += len(value_text)
        order.update(position.to_bytes(8, "little"))
        order.update(len(value_text).to_bytes(8, "little"))
        order.update(value_text)
        text.update(len(value_text).to_bytes(8, "little"))
        text.update(value_text)
    return {
        "paragraph_count": len(values),
        "order_sha256": order.hexdigest(),
        "text_sha256": text.hexdigest(),
        "text_bytes": text_bytes,
    }


def expected_captures(protocol: Mapping[str, Any] | None = None) -> list[dict[str, Any]]:
    attempt = str(protocol.get("attempt", "accepted")) if protocol else "accepted"
    rows: list[dict[str, Any]] = []
    for arm, repeat, route in (
        ("a1", 1, "materialized"),
        ("b1", 1, "bounded"),
        ("b2", 2, "bounded"),
        ("a2", 2, "materialized"),
    ):
        counts = COUNTS if repeat == 1 else tuple(reversed(COUNTS))
        for instrumentation in INSTRUMENTATIONS:
            for count in counts:
                rows.append({
                    "label": f"{arm}-{instrumentation}-{count}-{route}",
                    "arm": arm,
                    "repeat": repeat,
                    "route": route,
                    "route_name": ROUTE_NAMES[route],
                    "instrumentation": instrumentation,
                    "count": count,
                    "attempt": attempt,
                })
    return rows


def protocol_rows(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    if protocol.get("schema") != PROTOCOL_SCHEMA:
        fail("protocol schema differs")
    integer(protocol.get("samples"), "protocol.samples", SAMPLES)
    integer(protocol.get("warmups"), "protocol.warmups", WARMUPS)
    integer(protocol.get("cpu"), "protocol.cpu", CPU)
    if protocol.get("phase_samples") is not False:
        fail("phase_samples must be false for the initial matrix")
    for key in (
        "normal_and_allocator_timings_separate",
        "process_rss_includes_setup_oracles_and_teardown",
        "allocator_counters_are_operation_scoped",
    ):
        if protocol.get(key) is not True:
            fail(f"protocol.{key} must be true")
    if protocol.get("regression_review_percent") != REGRESSION_REVIEW_PERCENT:
        fail("protocol regression threshold differs")
    if protocol.get("performance_claim") != "none":
        fail("protocol performance claim differs")
    attempt = protocol.get("attempt")
    if not isinstance(attempt, str) or not attempt or any(
        character.isspace() or character in "/\\" for character in attempt
    ) or attempt in {".", ".."}:
        fail("protocol attempt is malformed")
    if not isinstance(protocol.get("frozen_utc"), str) or not protocol["frozen_utc"]:
        fail("protocol frozen timestamp is missing")
    if protocol.get("routes") != list(ROUTES) or protocol.get("arms") != list(ARMS):
        fail("protocol routes or arms differ")
    if protocol.get("instrumentations") != list(INSTRUMENTATIONS) or protocol.get("counts") != list(COUNTS):
        fail("protocol instrumentation or count matrix differs")
    for key in ("comparison", "scope", "route_flag"):
        if not isinstance(protocol.get(key), str) or not protocol[key]:
            fail(f"protocol.{key} is missing")
    environment = protocol.get("environment")
    if not isinstance(environment, dict) or set(environment) != {
        "RUSTUP_TOOLCHAIN", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
        "CARGO_PROFILE_RELEASE_DEBUG", "RUSTFLAGS", "DEBUGINFOD_URLS", "LC_ALL",
    } or not all(isinstance(value, str) for value in environment.values()):
        fail("protocol.environment differs")
    scripts = protocol.get("scripts")
    if not isinstance(scripts, dict) or set(scripts) != set(DRIVER_SCRIPTS):
        fail("protocol scripts differ")
    for name, value in scripts.items():
        digest(value, f"protocol.scripts.{name}")
    captures = protocol.get("captures")
    wanted = expected_captures(protocol)
    if not isinstance(captures, list) or len(captures) != len(wanted):
        fail(f"protocol must contain exactly {len(wanted)} captures")
    rows: list[dict[str, Any]] = []
    for index, (actual, expected) in enumerate(zip(captures, wanted)):
        if not isinstance(actual, dict):
            fail(f"protocol.captures[{index}] is malformed")
        for key, value in expected.items():
            if actual.get(key) != value:
                fail(f"protocol.captures[{index}].{key} differs")
        if LABEL.fullmatch(str(actual.get("label", ""))) is None:
            fail(f"protocol.captures[{index}].label is not path-safe")
        argv = actual.get("argv")
        if not isinstance(argv, list) or not argv or not all(isinstance(item, str) for item in argv):
            fail(f"protocol.captures[{index}].argv is malformed")
        rows.append(dict(actual))
    if len({row["label"] for row in rows}) != len(rows):
        fail("protocol capture labels are not unique")
    return rows


def _u64(value: Any, label: str) -> int:
    result = integer(value, label)
    if result >= 1 << 64:
        fail(f"{label}: exceeds u64")
    return result


def validate_reads(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict) or set(value) != set(READ_FIELDS):
        fail(f"{label}: source-read fields differ")
    calls = _u64(value["calls"], f"{label}.calls")
    requested = _u64(value["requested_bytes"], f"{label}.requested_bytes")
    returned = _u64(value["returned_bytes"], f"{label}.returned_bytes")
    histogram = value["request_histogram"]
    if not isinstance(histogram, dict) or set(histogram) != set(HIST_FIELDS):
        fail(f"{label}.request_histogram: fields differ")
    histogram = {key: _u64(histogram[key], f"{label}.request_histogram.{key}") for key in HIST_FIELDS}
    if sum(histogram.values()) != calls:
        fail(f"{label}: request histogram does not sum to calls")
    if returned > requested:
        fail(f"{label}: returned bytes exceed requested bytes")
    lower = (
        histogram["bytes_1_to_512"]
        + 513 * histogram["bytes_513_to_4096"]
        + 4097 * histogram["bytes_4097_to_16384"]
        + 16385 * histogram["bytes_16385_to_65536"]
        + 65537 * histogram["bytes_over_65536"]
    )
    upper = (
        512 * histogram["bytes_1_to_512"]
        + 4096 * histogram["bytes_513_to_4096"]
        + 16384 * histogram["bytes_4097_to_16384"]
        + 65536 * histogram["bytes_16385_to_65536"]
        + ((1 << 64) - 1) * histogram["bytes_over_65536"]
    )
    if not lower <= requested <= upper:
        fail(f"{label}: requested bytes outside histogram bounds")
    return {"calls": calls, "requested_bytes": requested, "returned_bytes": returned, "request_histogram": histogram}


def validate_sink(value: Any, candidate_bytes: int, candidate_sha: str, label: str) -> dict[str, Any]:
    if not isinstance(value, dict) or set(value) != set(SINK_FIELDS):
        fail(f"{label}: sink fields differ")
    accepted = _u64(value["accepted_bytes"], f"{label}.accepted_bytes")
    writes = _u64(value["write_calls"], f"{label}.write_calls")
    largest = _u64(value["largest_write"], f"{label}.largest_write")
    histogram = value["histogram"]
    if not isinstance(histogram, dict) or set(histogram) != set(HIST_FIELDS):
        fail(f"{label}.histogram: fields differ")
    histogram = {key: _u64(histogram[key], f"{label}.histogram.{key}") for key in HIST_FIELDS}
    if sum(histogram.values()) != writes:
        fail(f"{label}: sink histogram does not sum to write calls")
    if largest > HASH_SINK_MAX_WRITE:
        fail(f"{label}: sink write exceeds fixed short-write maximum")
    if accepted != candidate_bytes or digest(value["sha256"], f"{label}.sha256") != candidate_sha:
        fail(f"{label}: sink does not authenticate the candidate archive")
    return {"accepted_bytes": accepted, "write_calls": writes, "largest_write": largest, "histogram": histogram, "sha256": candidate_sha}


def validate_process(value: Any, label: str) -> dict[str, int] | None:
    if value is None:
        return None
    if not isinstance(value, dict) or set(value) != set(PROCESS_FIELDS):
        fail(f"{label}: process fields differ")
    result = {key: _u64(value[key], f"{label}.{key}") for key in PROCESS_FIELDS}
    if result["clock_ticks_per_second"] == 0 or result["peak_rss_bytes"] < result["rss_bytes"]:
        fail(f"{label}: invalid process observer values")
    return result


def validate_allocation(value: Any, instrumentation: str, label: str) -> dict[str, Any] | None:
    if value is None:
        if instrumentation == "allocator":
            fail(f"{label}: allocator sample missing")
        return None
    if not isinstance(value, dict):
        fail(f"{label}: allocation sample malformed")
    if instrumentation == "normal":
        if value not in ({"status": "unavailable", "scope": "operation_global_system_allocator"},):
            fail(f"{label}: normal allocation must be explicitly unavailable")
        return dict(value)
    if set(value) != {"status", "scope", *ALLOC_FIELDS}:
        fail(f"{label}: allocator fields differ")
    if value.get("status") != "measured" or value.get("scope") != "operation_global_system_allocator":
        fail(f"{label}: allocator identity differs")
    result = {key: _u64(value[key], f"{label}.{key}") for key in ALLOC_FIELDS}
    if result["failed_allocation_calls"] != 0:
        fail(f"{label}: failed allocation count is non-zero")
    if result["live_bytes_before"] + result["allocated_bytes"] - result["deallocated_bytes"] != result["live_bytes_after"]:
        fail(f"{label}: allocator conservation fails")
    if result["region_peak_live_bytes"] < max(result["live_bytes_before"], result["live_bytes_after"]):
        fail(f"{label}: region peak below boundary live bytes")
    return {"status": value["status"], "scope": value["scope"], **result}


def normalize_route(config: Mapping[str, Any], corpus: Mapping[str, Any], label: str) -> tuple[str, str]:
    raw = config.get("route", config.get("operation"))
    if raw is None:
        raw = corpus.get("route", corpus.get("operation"))
    if raw not in ROUTE_NAMES:
        fail(f"{label}: route identity is missing or unknown")
    short = "materialized" if raw.startswith("materialized") else "bounded"
    return short, ROUTE_NAMES[raw]


def validate_members(value: Any, label: str) -> list[dict[str, Any]]:
    if not isinstance(value, list) or len(value) != 4:
        fail(f"{label}: expected four physical members")
    fields = {"path", "compression_method", "data_descriptor", "crc32", "decoded_bytes", "decoded_sha256", "compressed_bytes", "compressed_sha256"}
    result: list[dict[str, Any]] = []
    for index, member in enumerate(value):
        if not isinstance(member, dict) or set(member) != fields:
            fail(f"{label}[{index}]: member fields differ")
        if not isinstance(member["path"], str) or not member["path"]:
            fail(f"{label}[{index}].path is malformed")
        if not isinstance(member["compression_method"], str) or not member["compression_method"]:
            fail(f"{label}[{index}].compression_method is malformed")
        if not isinstance(member["data_descriptor"], bool):
            fail(f"{label}[{index}].data_descriptor is malformed")
        for key in ("crc32", "decoded_bytes", "compressed_bytes"):
            _u64(member[key], f"{label}[{index}].{key}")
        if member["crc32"] >= 1 << 32:
            fail(f"{label}[{index}].crc32 exceeds u32")
        digest(member["decoded_sha256"], f"{label}[{index}].decoded_sha256")
        digest(member["compressed_sha256"], f"{label}[{index}].compressed_sha256")
        result.append(dict(member))
    paths = [member["path"] for member in result]
    if len(set(paths)) != 4 or MAIN_PATH not in paths or OPAQUE_PATH not in paths:
        fail(f"{label}: expected main and opaque members")
    return result


def validate_corpus(value: Any, count: int, route: str, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: corpus missing")
    required = {
        "generator", "format", "count", "append_text", "append_text_sha256",
        "source_archive_bytes", "source_archive_sha256", "source_main_xml_bytes", "source_main_xml_sha256",
        "source_semantic", "source_members", "source_member_count", "source_unchanged_verified",
        "source_main_xml_archive_verified", "source_opaque_member_exact_verified",
        "route_semantics_equal_verified", "route_untouched_members_equal_verified", "route_outputs_physical_equal",
        "materialized", "bounded", "limits",
    }
    if not required.issubset(value):
        fail(f"{label}: corpus is missing fields {sorted(required - set(value))}")
    if value["generator"] != GENERATOR or value["format"] != FORMAT or value["count"] != count:
        fail(f"{label}: corpus identity differs")
    source = source_xml(count)
    candidates = {
        "materialized": candidate_xml(count),
        "bounded": bounded_candidate_xml(count),
    }
    source_hash = hashlib.sha256(source).hexdigest()
    append_value = append_text(count)
    if value["append_text"] != append_value or value["append_text_sha256"] != hashlib.sha256(append_value.encode()).hexdigest():
        fail(f"{label}: append text oracle differs")
    if value["source_main_xml_bytes"] != len(source) or value["source_main_xml_sha256"] != source_hash:
        fail(f"{label}: source XML differs from independent oracle")
    source_archive_bytes = positive(value["source_archive_bytes"], f"{label}.source_archive_bytes")
    source_archive_sha = digest(value["source_archive_sha256"], f"{label}.source_archive_sha256")
    source_members = validate_members(value["source_members"], f"{label}.source_members")
    if value["source_member_count"] != len(source_members):
        fail(f"{label}: member counts differ")
    source_by_path = {member["path"]: member for member in source_members}
    if source_by_path[MAIN_PATH]["decoded_bytes"] != len(source) or source_by_path[MAIN_PATH]["decoded_sha256"] != source_hash:
        fail(f"{label}: source main member oracle differs")
    opaque = opaque_payload()
    opaque_hash = hashlib.sha256(opaque).hexdigest()
    opaque_crc = zlib.crc32(opaque) & 0xFFFF_FFFF
    if (
        source_by_path[OPAQUE_PATH]["decoded_bytes"] != OPAQUE_BYTES
        or source_by_path[OPAQUE_PATH]["decoded_sha256"] != opaque_hash
        or source_by_path[OPAQUE_PATH]["crc32"] != opaque_crc
    ):
        fail(f"{label}: opaque member length differs")
    semantic_expected = semantic(count)
    actual = value["source_semantic"]
    if not isinstance(actual, dict) or not all(actual.get(key) == expected for key, expected in semantic_expected.items()):
        fail(f"{label}.source_semantic: semantic oracle differs")
    for key in (
        "source_unchanged_verified", "source_main_xml_archive_verified", "source_opaque_member_exact_verified",
        "route_semantics_equal_verified", "route_untouched_members_equal_verified", "route_outputs_physical_equal",
    ):
        if key.startswith("route_"):
            if not isinstance(value.get(key), bool):
                fail(f"{label}.{key}: expected boolean")
        else:
            bool_true(value.get(key), f"{label}.{key}")
    route_records: dict[str, dict[str, Any]] = {}
    for route_key in ("materialized", "bounded"):
        route_value = value[route_key]
        candidate = candidates[route_key]
        candidate_hash = hashlib.sha256(candidate).hexdigest()
        route_records[route_key] = validate_route_corpus(
            route_value, count, route_key, source, candidate, source_members, source_hash, candidate_hash, f"{label}.{route_key}"
        )
    if route_records["materialized"]["output_semantic"] != route_records["bounded"]["output_semantic"]:
        fail(f"{label}: route semantic outputs differ")
    for route_key in ("materialized", "bounded"):
        if not route_records[route_key]["untouched_members"]:
            fail(f"{label}: {route_key} changed an untouched member")
    limits = value["limits"]
    if not isinstance(limits, dict):
        fail(f"{label}.limits is malformed")
    expected_limits = {
        "copy_max_xml_bytes": len(source) + 1024 * 1024,
        "copy_max_paragraphs": count + 1,
        "copy_max_events": (count + 1) * 8 + 128,
        "copy_max_depth": 16,
        "copy_max_output_bytes": len(source) + 2 * 1024 * 1024,
    }
    for key, expected in expected_limits.items():
        if limits.get(key) != expected:
            fail(f"{label}.limits.{key}: expected {expected}")
    if not isinstance(limits.get("bounded"), str) or not limits["bounded"]:
        fail(f"{label}.limits.bounded: expected bounded policy text")
    return {
        "source_archive_bytes": source_archive_bytes,
        "source_archive_sha256": source_archive_sha,
        "source_main_xml_bytes": len(source),
        "source_main_xml_sha256": source_hash,
        "source_semantic": semantic_expected,
        "source_members": source_members,
        "routes": route_records,
        "route": route,
        "raw": value,
    }


def validate_route_corpus(
    value: Any,
    count: int,
    route: str,
    source: bytes,
    candidate: bytes,
    source_members: list[dict[str, Any]],
    source_hash: str,
    candidate_hash: str,
    label: str,
) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: route corpus missing")
    required = {
        "route", "output_archive_bytes", "output_archive_sha256", "output_main_xml_bytes",
        "output_main_xml_sha256", "output_main_xml_expected_semantic_verified",
        "output_main_xml_expected_raw_verified", "scalar_proof", "output_semantic",
        "output_members", "output_member_count", "output_untouched_members_verified",
        "output_opaque_member_exact_verified", "output_physical_order_verified",
        "output_main_compressed_equal_source", "route_output_replay_verified",
        "route_inverse_verified", "stale_source_refusal_verified",
    }
    if not required.issubset(value):
        fail(f"{label}: missing fields {sorted(required - set(value))}")
    expected_route = ROUTE_NAMES[route]
    if value["route"] != expected_route:
        fail(f"{label}.route differs")
    output_archive_bytes = positive(value["output_archive_bytes"], f"{label}.output_archive_bytes")
    output_archive_sha = digest(value["output_archive_sha256"], f"{label}.output_archive_sha256")
    expected_xml_hash = hashlib.sha256(candidate).hexdigest()
    if value["output_main_xml_bytes"] != len(candidate) or value["output_main_xml_sha256"] != expected_xml_hash:
        fail(f"{label}: output XML differs from independent candidate oracle")
    output_members = validate_members(value["output_members"], f"{label}.output_members")
    if value["output_member_count"] != len(output_members):
        fail(f"{label}: output member count differs")
    source_by_path = {member["path"]: member for member in source_members}
    output_by_path = {member["path"]: member for member in output_members}
    if [member["path"] for member in output_members] != [member["path"] for member in source_members]:
        fail(f"{label}: output member order differs")
    if output_by_path[MAIN_PATH]["decoded_bytes"] != len(candidate) or output_by_path[MAIN_PATH]["decoded_sha256"] != expected_xml_hash:
        fail(f"{label}: output main member differs from candidate oracle")
    opaque = opaque_payload()
    if (
        output_by_path[OPAQUE_PATH] != source_by_path[OPAQUE_PATH]
        or output_by_path[OPAQUE_PATH]["decoded_bytes"] != len(opaque)
        or output_by_path[OPAQUE_PATH]["decoded_sha256"] != hashlib.sha256(opaque).hexdigest()
        or output_by_path[OPAQUE_PATH]["crc32"] != (zlib.crc32(opaque) & 0xFFFF_FFFF)
    ):
        fail(f"{label}: opaque member is not exact")
    untouched = all(
        output_by_path[path] == member
        for path, member in source_by_path.items()
        if path != MAIN_PATH
    )
    if untouched != (value["output_untouched_members_verified"] is True):
        fail(f"{label}: untouched-member proof disagrees with member records")
    if value["output_opaque_member_exact_verified"] is not True or value["output_physical_order_verified"] is not True:
        fail(f"{label}: physical preservation proof is absent")
    expected_compressed_equal = (
        output_by_path[MAIN_PATH]["compressed_sha256"]
        == source_by_path[MAIN_PATH]["compressed_sha256"]
    )
    if value["output_main_compressed_equal_source"] != expected_compressed_equal:
        fail(f"{label}: compressed-main preservation proof disagrees with member records")
    output_semantic = value["output_semantic"]
    expected_semantic = semantic(count, True)
    if not isinstance(output_semantic, dict) or not all(output_semantic.get(key) == expected for key, expected in expected_semantic.items()):
        fail(f"{label}: output semantic oracle differs")
    for key in (
        "output_main_xml_expected_semantic_verified", "output_main_xml_expected_raw_verified",
        "route_output_replay_verified",
        "route_inverse_verified", "stale_source_refusal_verified",
    ):
        bool_true(value.get(key), f"{label}.{key}")
    scalar = value["scalar_proof"]
    if route == "materialized":
        if scalar is not None:
            fail(f"{label}.scalar_proof: materialized route must omit bounded proof")
    else:
        validate_scalar_proof(scalar, count, source, candidate, source_hash, expected_xml_hash, f"{label}.scalar_proof")
    return {
        "route": expected_route,
        "output_archive_bytes": output_archive_bytes,
        "output_archive_sha256": output_archive_sha,
        "output_semantic": expected_semantic,
        "output_members": output_members,
        "untouched_members": untouched,
        "raw": value,
    }


def validate_scalar_proof(
    value: Any,
    count: int,
    source: bytes,
    candidate: bytes,
    source_hash: str,
    candidate_hash: str,
    label: str,
) -> None:
    fields = {
        "source_version_id", "source_version_revision", "source_len", "source_sha256",
        "insertion_offset", "source_paragraph_count", "source_event_count",
        "source_max_depth", "source_strict_namespace", "source_sect_pr_len",
        "source_sect_pr_sha256", "candidate_len", "candidate_sha256",
        "candidate_paragraph_count", "candidate_event_count", "candidate_max_depth",
        "generated_offset", "generated_once", "candidate_sect_pr_len",
        "candidate_sect_pr_sha256",
    }
    if not isinstance(value, dict) or set(value) != fields:
        fail(f"{label}: bounded scalar proof fields differ")
    for key in (
        "source_version_id", "source_version_revision", "source_len", "insertion_offset",
        "source_paragraph_count", "source_event_count", "source_max_depth", "candidate_len",
        "candidate_paragraph_count", "candidate_event_count", "candidate_max_depth",
        "generated_offset", "candidate_sect_pr_len",
    ):
        _u64(value[key], f"{label}.{key}")
    if value["source_version_id"] == 0 or value["source_version_revision"] != 0:
        fail(f"{label}: source version identity is invalid")
    source_close = source.find(b"</w:body>")
    if source_close < 0:
        fail(f"{label}: source body close is absent")
    expected = {
        "source_len": len(source),
        "source_sha256": source_hash,
        "insertion_offset": source_close,
        "source_paragraph_count": count,
        "source_strict_namespace": False,
        "source_sect_pr_len": 0,
        "source_sect_pr_sha256": hashlib.sha256(b"").hexdigest(),
        "candidate_len": len(candidate),
        "candidate_sha256": candidate_hash,
        "candidate_paragraph_count": count + 1,
        "generated_offset": source_close,
        "generated_once": True,
        "candidate_sect_pr_len": 0,
        "candidate_sect_pr_sha256": hashlib.sha256(b"").hexdigest(),
    }
    for key, expected_value in expected.items():
        if value.get(key) != expected_value:
            fail(f"{label}.{key}: scalar proof oracle differs")
    if value["source_event_count"] == 0 or value["candidate_event_count"] <= value["source_event_count"]:
        fail(f"{label}: scalar event counts are not increasing")
    if value["source_max_depth"] == 0 or value["candidate_max_depth"] < value["source_max_depth"]:
        fail(f"{label}: scalar depth proof is invalid")


def validate_report(
    report: Any,
    spec: Mapping[str, Any],
    label: str = "report",
    *,
    expected_samples: int = SAMPLES,
    expected_warmups: int = WARMUPS,
) -> dict[str, Any]:
    if not isinstance(report, dict):
        fail(f"{label}: report missing")
    if report.get("schema") not in REPORT_SCHEMAS or report.get("version") != 1:
        fail(f"{label}: report schema/version differs")
    binary = report.get("binary")
    if not isinstance(binary, dict):
        fail(f"{label}.binary is missing")
    expected_binary = {
        "binary": "litchi-perf-baseline-alloc" if spec["instrumentation"] == "allocator" else "litchi-perf-baseline",
        "allocator": "CountingSystemAllocator(std::alloc::System)" if spec["instrumentation"] == "allocator" else "Rust system allocator",
        "instrumentation": "system_allocator_operation_scoped" if spec["instrumentation"] == "allocator" else "none",
        "counter_revision": "serialized_region_peak_v3" if spec["instrumentation"] == "allocator" else None,
    }
    if binary != expected_binary:
        fail(f"{label}.binary: instrumentation identity differs")
    config = report.get("config")
    if not isinstance(config, dict):
        fail(f"{label}.config missing")
    if (
        config.get("counts") != [spec["count"]]
        or config.get("samples") != expected_samples
        or config.get("warmups") != expected_warmups
    ):
        fail(f"{label}.config: expected {expected_samples} samples/{expected_warmups} warmups")
    if config.get("sink") != SINK_ID or config.get("source") != SOURCE_ID:
        fail(f"{label}.config: source/sink identity differs")
    if not isinstance(config.get("lifecycle"), str) or not config["lifecycle"]:
        fail(f"{label}.config.lifecycle is missing")
    if not isinstance(config.get("text_authoring"), str) or not config["text_authoring"]:
        fail(f"{label}.config.text_authoring is missing")
    route, route_name = normalize_route(config, {}, f"{label}.config")
    if route != spec["route"] or route_name != spec["route_name"]:
        fail(f"{label}.config: route differs from capture spec")
    cases = report.get("cases")
    if not isinstance(cases, list) or len(cases) != 1:
        fail(f"{label}: expected one count case")
    case = cases[0]
    if not isinstance(case, dict) or case.get("count") != spec["count"]:
        fail(f"{label}.case: count differs")
    corpus = validate_corpus(case.get("corpus"), spec["count"], route, f"{label}.corpus")
    route_reports = case.get("routes")
    if not isinstance(route_reports, list) or len(route_reports) != 1:
        fail(f"{label}: expected one route report")
    route_report = route_reports[0]
    if not isinstance(route_report, dict) or route_report.get("route") != route_name:
        fail(f"{label}: route report identity differs")
    raw_route_corpus = corpus["raw"][route]
    if route_report.get("corpus") != raw_route_corpus:
        fail(f"{label}: route report corpus differs from shared corpus proof")
    samples = route_report.get("samples")
    if not isinstance(samples, list) or len(samples) != expected_samples:
        fail(f"{label}: sample vector differs")
    checked_samples: list[dict[str, Any]] = []
    for index, sample in enumerate(samples):
        if not isinstance(sample, dict) or not {"sample", "route", "elapsed_ns", "source_reads", "sink", "allocation", "process"}.issubset(sample):
            fail(f"{label}.samples[{index}]: fields differ")
        if sample["sample"] != index or sample["route"] != route_name:
            fail(f"{label}.samples[{index}]: sample index differs")
        elapsed_ns = positive(sample["elapsed_ns"], f"{label}.samples[{index}].elapsed_ns")
        reads = validate_reads(sample["source_reads"], f"{label}.samples[{index}].source_reads")
        route_corpus = corpus["routes"][route]
        sink = validate_sink(sample["sink"], route_corpus["output_archive_bytes"], route_corpus["output_archive_sha256"], f"{label}.samples[{index}].sink")
        allocation = validate_allocation(sample["allocation"], spec["instrumentation"], f"{label}.samples[{index}].allocation")
        process = validate_process(sample["process"], f"{label}.samples[{index}].process")
        checked_samples.append({"sample": index, "elapsed_ns": elapsed_ns, "source_reads": reads, "sink": sink, "allocation": allocation, "process": process})
    return {
        "schema": report["schema"],
        "version": report["version"],
        "config": config,
        "route": route,
        "route_name": route_name,
        "corpus": corpus,
        "samples": checked_samples,
        "raw": report,
    }


def _resource_rss(spec: Mapping[str, Any]) -> int:
    path = ROOT / "captures" / f"{spec['label']}.resource"
    if not path.is_file():
        fail(f"{spec['label']}: /usr/bin/time resource receipt missing")
    values = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("Maximum resident set size (kbytes):"):
            values.append(int(line.rsplit(":", 1)[1].strip()))
    if len(values) != 1 or values[0] < 0:
        fail(f"{spec['label']}: expected exactly one RSS observation")
    return values[0] * 1024


def _stats(values: list[int | float]) -> dict[str, float | int]:
    if not values:
        fail("cannot summarize an empty sample vector")
    ordered = sorted(values)
    mean = statistics.fmean(values)
    stdev = statistics.stdev(values) if len(values) > 1 else 0.0
    sem = stdev / math.sqrt(len(values))
    return {
        "n": len(values),
        "min": min(values),
        "max": max(values),
        "mean": mean,
        "stdev": stdev,
        "ci95_low": mean - 1.96 * sem,
        "ci95_high": mean + 1.96 * sem,
        "p50": _percentile(ordered, 0.50),
        "p95": _percentile(ordered, 0.95),
        "p99": _percentile(ordered, 0.99),
    }


def _percentile(ordered: list[int | float], fraction: float) -> float:
    if len(ordered) == 1:
        return float(ordered[0])
    position = (len(ordered) - 1) * fraction
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return float(ordered[lower])
    ratio = position - lower
    return float(ordered[lower] + (ordered[upper] - ordered[lower]) * ratio)


def metric_row(spec: Mapping[str, Any], checked: Mapping[str, Any]) -> dict[str, Any]:
    samples = checked["samples"]
    row: dict[str, Any] = {
        "label": spec["label"],
        "arm": spec["arm"],
        "repeat": spec["repeat"],
        "route": checked["route"],
        "route_name": checked["route_name"],
        "instrumentation": spec["instrumentation"],
        "count": spec["count"],
        "source_archive_sha256": checked["corpus"]["source_archive_sha256"],
        "source_main_xml_sha256": checked["corpus"]["source_main_xml_sha256"],
        "rss_bytes": _resource_rss(spec),
        "elapsed_ns": _stats([sample["elapsed_ns"] for sample in samples]),
        "throughput_bytes_per_second": _stats([
            sample["sink"]["accepted_bytes"] * 1_000_000_000 / sample["elapsed_ns"]
            for sample in samples
        ]),
        "source_reads": {
            field: _stats([sample["source_reads"][field] for sample in samples])
            for field in ("calls", "requested_bytes", "returned_bytes")
        },
        "sink": {
            field: _stats([sample["sink"][field] for sample in samples])
            for field in ("accepted_bytes", "write_calls", "largest_write")
        },
    }
    allocations = [sample["allocation"] for sample in samples]
    if spec["instrumentation"] == "allocator":
        if any(item is None for item in allocations):
            fail(f"{spec['label']}: allocator metrics missing")
        row["allocation"] = {field: _stats([item[field] for item in allocations]) for field in ALLOC_FIELDS}
        row["total_peak_live_bytes"] = row["allocation"]["region_peak_live_bytes"]
    else:
        if any(item is not None for item in allocations):
            fail(f"{spec['label']}: normal report exposes measured allocation numbers")
        row["allocation"] = {"status": "unavailable", "scope": "operation_global_system_allocator"}
    process = [sample["process"] for sample in samples]
    if any(item is not None for item in process):
        row["process"] = {field: _stats([item[field] for item in process if item is not None]) for field in PROCESS_FIELDS}
    else:
        row["process"] = None
    return row


def _percent_change(before: float, after: float) -> float:
    if before == 0:
        return 0.0 if after == 0 else math.inf
    return (after - before) / before * 100.0


def _comparison_metrics(
    first: Mapping[str, Any], second: Mapping[str, Any], instrumentation: str
) -> list[str]:
    names = [
        "elapsed_ns",
        "throughput_bytes_per_second",
        "rss_bytes",
        "source_reads.calls",
        "source_reads.requested_bytes",
        "source_reads.returned_bytes",
        "sink.accepted_bytes",
        "sink.write_calls",
        "sink.largest_write",
    ]
    if isinstance(first.get("process"), dict) and isinstance(second.get("process"), dict):
        names.extend(f"process.{field}" for field in PROCESS_FIELDS)
    if instrumentation == "allocator":
        names.append("total_peak_live_bytes")
        names.extend(f"allocation.{field}" for field in ALLOC_FIELDS)
    return names


def _metric_mean(row: Mapping[str, Any], metric: str) -> float:
    value: Any = row
    for part in metric.split("."):
        if not isinstance(value, Mapping) or part not in value:
            fail(f"comparison metric is missing: {metric}")
        value = value[part]
    if isinstance(value, Mapping):
        value = value.get("mean")
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        fail(f"comparison metric is not numeric: {metric}")
    return float(value)


def _comparison_changes(
    first: Mapping[str, Any], second: Mapping[str, Any], instrumentation: str,
    first_key: str, second_key: str,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    changes: dict[str, Any] = {}
    flags: list[dict[str, Any]] = []
    for metric in _comparison_metrics(first, second, instrumentation):
        before = _metric_mean(first, metric)
        after = _metric_mean(second, metric)
        change = _percent_change(before, after)
        changes[metric] = {first_key: before, second_key: after, "percent": change}
        if abs(change) > REGRESSION_REVIEW_PERCENT:
            flags.append({
                "metric": metric,
                "percent": change,
                "adverse": change < 0 if metric == "throughput_bytes_per_second" else change > 0,
            })
    return changes, flags


def comparisons(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_key = {(row["arm"], row["instrumentation"], row["count"]): row for row in rows}
    output: list[dict[str, Any]] = []
    for instrumentation in INSTRUMENTATIONS:
        for count in COUNTS:
            for candidate_arm, control_arm, name in (
                ("b1", "a1", "candidate_b1_vs_control_a1"),
                ("b2", "a2", "candidate_b2_vs_control_a2"),
            ):
                candidate = by_key[(candidate_arm, instrumentation, count)]
                control = by_key[(control_arm, instrumentation, count)]
                changes, flags = _comparison_changes(
                    control, candidate, instrumentation, "control_mean", "candidate_mean"
                )
                output.append({"name": name, "instrumentation": instrumentation, "count": count, "changes": changes, "flags": flags})
            for route, first_arm, second_arm, name in (
                ("materialized", "a1", "a2", "materialized_repeat_a2_vs_a1"),
                ("bounded", "b1", "b2", "bounded_repeat_b2_vs_b1"),
            ):
                first = by_key[(first_arm, instrumentation, count)]
                second = by_key[(second_arm, instrumentation, count)]
                changes, flags = _comparison_changes(
                    first, second, instrumentation, "first_mean", "second_mean"
                )
                output.append({"name": name, "instrumentation": instrumentation, "count": count, "route": route, "changes": changes, "flags": flags})
    return output


def analyze(protocol: Mapping[str, Any]) -> dict[str, Any]:
    specs = protocol_rows(protocol)
    rows: list[dict[str, Any]] = []
    checked_reports: list[dict[str, Any]] = []
    for spec in specs:
        report_path = ROOT / "captures" / f"{spec['label']}.report.json"
        report = read_json(report_path)
        checked = validate_report(report, spec, spec["label"])
        rows.append(metric_row(spec, checked))
        checked_reports.append({"label": spec["label"], "report_sha256": hashlib.sha256(report_path.read_bytes()).hexdigest()})
    if len(rows) != 24:
        fail("analysis did not retain exactly 24 captures")
    for count in COUNTS:
        source_rows = [row for row in rows if row["count"] == count]
        source_hashes = {row["source_archive_sha256"] for row in source_rows}
        source_xml_hashes = {row["source_main_xml_sha256"] for row in source_rows}
        if len(source_hashes) != 1 or len(source_xml_hashes) != 1:
            fail(f"count {count}: route/process reports do not share one deterministic source")
    by_arm = {arm: sum(1 for row in rows if row["arm"] == arm) for arm in ARMS}
    if by_arm != {arm: 6 for arm in ARMS}:
        fail(f"capture arm counts differ: {by_arm}")
    return {
        "schema": SUMMARY_SCHEMA,
        "protocol_sha256": hashlib.sha256((ROOT / "protocol.json").read_bytes()).hexdigest(),
        "source_manifest_sha256": _source_manifest_hash(protocol),
        "capture_count": len(rows),
        "sample_count": len(rows) * SAMPLES,
        "rows": rows,
        "comparisons": comparisons(rows),
        "review_threshold_percent": REGRESSION_REVIEW_PERCENT,
        "flags_include_improvements": True,
        "reports": checked_reports,
    }


def _source_manifest_hash(protocol: Mapping[str, Any]) -> str | None:
    attempt = protocol.get("attempt")
    if not isinstance(attempt, str) or not attempt:
        fail("protocol attempt is missing")
    paths = [ROOT / "builds" / f"binaries-{attempt}.json"]
    if not paths[0].is_file():
        return None
    values = {read(paths[0]).get("source_manifest_sha256")}
    values.discard(None)
    if len(values) != 1:
        fail("binary custody records do not share one source manifest")
    return next(iter(values))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--summary", default="summary.json")
    options = parser.parse_args()
    protocol = read_json(ROOT / "protocol.json")
    summary = analyze(protocol)
    write_json(ROOT / options.summary, summary)
    print(f"analyzed {summary['capture_count']} captures / {summary['sample_count']} samples")


if __name__ == "__main__":
    main()
