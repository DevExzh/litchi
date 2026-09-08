#!/usr/bin/env python3
"""Validate and summarize the retained 0479 DOCX append observations.

The executable report is intentionally a small, self describing envelope.  This
module is the independent boundary around it: it checks the corpus and output
oracles, validates every measured vector, and derives statistics without
trusting producer supplied summaries.  Total and phase mode rows stay separate;
phase high water marks are never added together or presented as a total peak.
"""

from __future__ import annotations

import hashlib
import json
import math
import re
import statistics
import datetime as _datetime
import zlib
from pathlib import Path
from typing import Any, Mapping


ROOT = Path(__file__).resolve().parent
PROTOCOL_SCHEMA = "docx-plain-paragraph-tail-append-capture-v1"
REPORT_SCHEMA = "docx-plain-paragraph-tail-append-v1"
SUMMARY_SCHEMA = "docx-plain-paragraph-tail-append-summary-v1"
CORPUS_MANIFEST_SCHEMA = "docx-plain-paragraph-tail-append-corpus-manifest-v1"
COUNTS = (64, 8_192, 131_072)
MODES = ("total", "phases")
INSTRUMENTATIONS = ("normal", "allocator")
REPEATS = (1, 2)
SAMPLES = 30
WARMUPS = 3
CPU = 2
PHASES = ("open", "snapshot", "stage", "commit", "publish", "drop")
REGRESSION_REVIEW_PERCENT = 5
MAX_DURABLE_BYTES = 64 * 1024 * 1024
OPAQUE_PATH = "word/perf-opaque.bin"
OPAQUE_BYTES = 32 * 1024
OPAQUE_SHA256 = "7c697d881b1f1e264566d6ab95c8304a77229b468642128de4f271e9d9a25856"
OPAQUE_CRC32 = 0x1F806DA7
MAIN_PATH = "word/document.xml"
GENERATOR = "litchi-docx-plain-paragraph-tail-append-v1"
FORMAT = "DOCX/OOXML/OPC/ZIP"
WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
HASH_SINK_MAX_WRITE = 16 * 1024
SINK_ID = "non_seek_hashing_scalar_sha256_no_archive_retention_shortwrite_16k"
SOURCE_ID = "caller_owned_arc_positional_read_at_scalar_counters_requested_returned_fixed_histogram"
SHA256 = re.compile(r"^[0-9a-f]{64}$")


class AnalysisError(ValueError):
    """The retained data cannot support the declared comparison."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def _reject_constant(value: str) -> Any:
    fail(f"non-finite JSON value {value}")


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            fail(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=_reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, AnalysisError) as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
        stream.write("\n")


def integer(value: Any, label: str, expected: int | None = None) -> int:
    if not isinstance(value, int) or isinstance(value, bool) or value < 0:
        fail(f"{label}: expected a non-negative integer")
    if expected is not None and value != expected:
        fail(f"{label}: expected {expected}, got {value}")
    return value


def u64(value: Any, label: str, expected: int | None = None) -> int:
    result = integer(value, label, expected)
    if result >= 1 << 64:
        fail(f"{label}: exceeds an unsigned 64-bit counter")
    return result


def positive(value: Any, label: str) -> int:
    result = integer(value, label)
    if result == 0:
        fail(f"{label}: expected a positive integer")
    return result


def digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or SHA256.fullmatch(value) is None:
        fail(f"{label}: expected a lower-case SHA-256 digest")
    return value


def _bool(value: Any, label: str) -> None:
    if value is not True:
        fail(f"{label}: expected true")


def _text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty text")
    return value


def _timestamp(value: Any, label: str) -> _datetime.datetime:
    if not isinstance(value, str) or not value:
        fail(f"{label}: timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp: {error}")
    if parsed.tzinfo is None:
        fail(f"{label}: timestamp must include a timezone")
    return parsed


def expected_captures() -> list[dict[str, Any]]:
    forward = [
        (instrumentation, count, mode)
        for instrumentation in INSTRUMENTATIONS
        for count in COUNTS
        for mode in MODES
    ]
    rows: list[dict[str, Any]] = []
    for repeat, sequence in ((1, forward), (2, list(reversed(forward)))):
        for instrumentation, count, mode in sequence:
            rows.append({
                "label": f"r{repeat}-{instrumentation}-{count}-{mode}",
                "instrumentation": instrumentation,
                "count": count,
                "mode": mode,
                "repeat": repeat,
            })
    return rows


def protocol_rows(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    if protocol.get("schema") != PROTOCOL_SCHEMA:
        fail("protocol schema differs")
    integer(protocol.get("samples"), "protocol.samples", SAMPLES)
    integer(protocol.get("warmups"), "protocol.warmups", WARMUPS)
    integer(protocol.get("cpu"), "protocol.cpu", CPU)
    integer(protocol.get("append_count"), "protocol.append_count", 1)
    if protocol.get("normal_and_allocator_timings_separate") is not True:
        fail("protocol does not separate normal and allocator timings")
    if protocol.get("phase_peaks_are_not_total_peaks") is not True:
        fail("protocol does not separate phase and total peaks")
    if protocol.get("regression_review_percent") != REGRESSION_REVIEW_PERCENT:
        fail("protocol regression review threshold differs")
    if protocol.get("performance_claim") != "none":
        fail("protocol performance claim differs")
    for key in ("comparison", "scope"):
        _text(protocol.get(key), f"protocol.{key}")
    environment = protocol.get("environment")
    if not isinstance(environment, dict):
        fail("protocol.environment is missing")
    required_env = (
        "RUSTUP_TOOLCHAIN", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
        "CARGO_PROFILE_RELEASE_DEBUG", "RUSTFLAGS", "DEBUGINFOD_URLS", "LC_ALL",
    )
    if set(environment) != set(required_env):
        fail("protocol.environment fields differ")
    scripts = protocol.get("scripts")
    if not isinstance(scripts, dict) or set(scripts) != {"common.py", "capture.py"}:
        fail("protocol.scripts fields differ")
    for name, value in scripts.items():
        digest(value, f"protocol.scripts.{name}")
    captures = protocol.get("captures")
    expected = expected_captures()
    if not isinstance(captures, list) or len(captures) != len(expected):
        fail(f"protocol must contain exactly {len(expected)} captures")
    rows: list[dict[str, Any]] = []
    for index, (actual, wanted) in enumerate(zip(captures, expected)):
        if not isinstance(actual, dict):
            fail(f"protocol.captures[{index}] must be an object")
        for key, value in wanted.items():
            if actual.get(key) != value:
                fail(f"protocol.captures[{index}].{key} differs")
        argv = actual.get("argv")
        if not isinstance(argv, list) or not argv or not all(isinstance(item, str) for item in argv):
            fail(f"protocol.captures[{index}].argv is malformed")
        rows.append(dict(actual))
    if len({row["label"] for row in rows}) != len(rows):
        fail("protocol capture labels are not unique")
    return rows


def _fragment(count: int = 0) -> bytes:
    return f"<w:p><w:r><w:t>paragraph-{count:06}</w:t></w:r></w:p>".encode()


def _source_xml(count: int) -> bytes:
    prefix = (
        '<?xml version="1.0" encoding="UTF-8"?><w:document '
        f'xmlns:w="{WORD_NAMESPACE}"><w:body>'
    ).encode()
    return prefix + b"".join(_fragment(index) for index in range(count)) + b"</w:body></w:document>"


def _candidate_xml(count: int) -> bytes:
    source = _source_xml(count)
    marker = b"</w:body></w:document>"
    if not source.endswith(marker):
        fail("internal source XML formula is malformed")
    return source[:-len(marker)] + _fragment(0) + marker


def _semantic(count: int, candidate: bool = False) -> dict[str, Any]:
    values = list(range(count))
    if candidate:
        values.append(0)
    order = hashlib.sha256()
    text = hashlib.sha256()
    order.update(b"litchi-docx-tail-order-v1\0")
    text.update(b"litchi-docx-tail-text-v1\0")
    order.update(len(values).to_bytes(8, "little"))
    text.update(len(values).to_bytes(8, "little"))
    text_bytes = 0
    for index, value in enumerate(values):
        value_text = f"paragraph-{value:06}".encode()
        text_bytes += len(value_text)
        order.update(index.to_bytes(8, "little"))
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


def _metadata_identity(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected an object")
    required = {
        "path", "compression_method", "data_descriptor", "crc32",
        "decoded_bytes", "decoded_sha256", "compressed_bytes", "compressed_sha256",
    }
    if set(value) != required:
        fail(f"{label}: member fields differ")
    _text(value.get("path"), f"{label}.path")
    if not isinstance(value.get("compression_method"), str):
        fail(f"{label}.compression_method: expected text")
    if not isinstance(value.get("data_descriptor"), bool):
        fail(f"{label}.data_descriptor: expected boolean")
    crc32 = u64(value.get("crc32"), f"{label}.crc32")
    if crc32 > 0xFFFF_FFFF:
        fail(f"{label}.crc32: exceeds a 32-bit CRC")
    u64(value.get("decoded_bytes"), f"{label}.decoded_bytes")
    digest(value.get("decoded_sha256"), f"{label}.decoded_sha256")
    u64(value.get("compressed_bytes"), f"{label}.compressed_bytes")
    digest(value.get("compressed_sha256"), f"{label}.compressed_sha256")
    if value["decoded_bytes"] == 0 or value["compressed_bytes"] == 0:
        fail(f"{label}: member byte counts must be positive")
    return dict(value)


def _limits(value: Any, count: int, source_xml_bytes: int, label: str) -> dict[str, int]:
    if not isinstance(value, dict):
        fail(f"{label}: limits are missing")
    fields = ("max_xml_bytes", "max_paragraphs", "max_events", "max_depth", "max_output_bytes", "max_durable_bytes")
    if set(value) != set(fields):
        fail(f"{label}: limit fields differ")
    result = {field: positive(value.get(field), f"{label}.{field}") for field in fields}
    expected = {
        "max_xml_bytes": source_xml_bytes + 1024 * 1024,
        "max_paragraphs": count + 1,
        "max_events": (count + 1) * 8 + 128,
        "max_depth": 16,
        "max_output_bytes": source_xml_bytes + 2 * 1024 * 1024,
        "max_durable_bytes": MAX_DURABLE_BYTES,
    }
    if result != expected:
        fail(f"{label}: limits differ from the bounded corpus policy")
    return result


def _corpus(value: Any, count: int, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: corpus is missing")
    required = {
        "generator", "format", "count", "opaque_path", "opaque_bytes",
        "source_archive_bytes", "source_archive_sha256", "candidate_archive_bytes", "candidate_archive_sha256",
        "source_main_xml_bytes", "source_main_xml_sha256", "candidate_main_xml_bytes", "candidate_main_xml_sha256",
        "source_main_xml_archive_verified", "candidate_main_xml_archive_verified", "candidate_main_xml_oracle_verified",
        "source_members", "candidate_members", "source_member_count", "candidate_member_count",
        "source_semantic", "candidate_semantic", "source_unchanged_verified",
        "source_semantic_reopen_verified", "candidate_semantic_reopen_verified",
        "tail_copy_exactly_one_verified", "untouched_members_verified", "opaque_member_exact_verified", "physical_order_verified",
        "limits", "patch",
    }
    if set(value) != required:
        fail(f"{label}: corpus fields differ")
    if value["generator"] != GENERATOR or value["format"] != FORMAT or value["count"] != count:
        fail(f"{label}: corpus generator/format/count differs")
    if value["opaque_path"] != OPAQUE_PATH or value["opaque_bytes"] != OPAQUE_BYTES:
        fail(f"{label}: opaque member identity differs")
    source_xml = _source_xml(count)
    candidate_xml = _candidate_xml(count)
    if value["source_main_xml_bytes"] != len(source_xml) or value["candidate_main_xml_bytes"] != len(candidate_xml):
        fail(f"{label}: main XML byte counts differ")
    if value["source_main_xml_sha256"] != hashlib.sha256(source_xml).hexdigest():
        fail(f"{label}: source main XML digest differs")
    if value["candidate_main_xml_sha256"] != hashlib.sha256(candidate_xml).hexdigest():
        fail(f"{label}: candidate main XML digest differs")
    digest(value["source_archive_sha256"], f"{label}.source_archive_sha256")
    digest(value["candidate_archive_sha256"], f"{label}.candidate_archive_sha256")
    if value["source_archive_sha256"] == value["candidate_archive_sha256"]:
        fail(f"{label}: source and candidate archive digests coincide")
    source_archive_bytes = u64(value["source_archive_bytes"], f"{label}.source_archive_bytes")
    candidate_archive_bytes = u64(value["candidate_archive_bytes"], f"{label}.candidate_archive_bytes")
    if source_archive_bytes == 0 or candidate_archive_bytes == 0:
        fail(f"{label}: archive byte counts must be positive")
    if candidate_archive_bytes < source_archive_bytes:
        fail(f"{label}: candidate archive is smaller than its source despite the append")
    if value["source_member_count"] != 4 or value["candidate_member_count"] != 4:
        fail(f"{label}: expected exactly four ZIP members")
    source_members = [_metadata_identity(item, f"{label}.source_members[{i}]") for i, item in enumerate(value["source_members"])] if isinstance(value["source_members"], list) else fail(f"{label}.source_members: expected list")
    candidate_members = [_metadata_identity(item, f"{label}.candidate_members[{i}]") for i, item in enumerate(value["candidate_members"])] if isinstance(value["candidate_members"], list) else fail(f"{label}.candidate_members: expected list")
    if len(source_members) != 4 or len(candidate_members) != 4:
        fail(f"{label}: member inventory length differs")
    expected_paths = {"[Content_Types].xml", "_rels/.rels", MAIN_PATH, OPAQUE_PATH}
    if {item["path"] for item in source_members} != expected_paths or {item["path"] for item in candidate_members} != expected_paths:
        fail(f"{label}: member paths differ")
    source_order = [item["path"] for item in source_members]
    candidate_order = [item["path"] for item in candidate_members]
    if len(set(source_order)) != len(source_order) or source_order != candidate_order:
        fail(f"{label}: physical member order differs or contains duplicates")
    source_by_path = {item["path"]: item for item in source_members}
    candidate_by_path = {item["path"]: item for item in candidate_members}
    if source_by_path[MAIN_PATH]["decoded_bytes"] != len(source_xml) or candidate_by_path[MAIN_PATH]["decoded_bytes"] != len(candidate_xml):
        fail(f"{label}: main member decoded lengths differ")
    if source_by_path[MAIN_PATH]["decoded_sha256"] != hashlib.sha256(source_xml).hexdigest() or candidate_by_path[MAIN_PATH]["decoded_sha256"] != hashlib.sha256(candidate_xml).hexdigest():
        fail(f"{label}: main member decoded digest differs from the XML oracle")
    source_crc32 = zlib.crc32(source_xml) & 0xFFFF_FFFF
    candidate_crc32 = zlib.crc32(candidate_xml) & 0xFFFF_FFFF
    if source_by_path[MAIN_PATH]["crc32"] != source_crc32 or candidate_by_path[MAIN_PATH]["crc32"] != candidate_crc32:
        fail(f"{label}: main member CRC differs from the XML oracle")
    if source_by_path[OPAQUE_PATH]["decoded_bytes"] != OPAQUE_BYTES or candidate_by_path[OPAQUE_PATH]["decoded_bytes"] != OPAQUE_BYTES:
        fail(f"{label}: opaque member decoded length differs")
    if source_by_path[OPAQUE_PATH]["decoded_sha256"] != OPAQUE_SHA256 or candidate_by_path[OPAQUE_PATH]["decoded_sha256"] != OPAQUE_SHA256:
        fail(f"{label}: opaque member digest differs from the fixed payload oracle")
    if source_by_path[OPAQUE_PATH]["crc32"] != OPAQUE_CRC32 or candidate_by_path[OPAQUE_PATH]["crc32"] != OPAQUE_CRC32:
        fail(f"{label}: opaque member CRC differs from the fixed payload oracle")
    if source_by_path[OPAQUE_PATH] != candidate_by_path[OPAQUE_PATH]:
        fail(f"{label}: opaque member was not preserved exactly")
    source_compressed_bytes = sum(item["compressed_bytes"] for item in source_members)
    candidate_compressed_bytes = sum(item["compressed_bytes"] for item in candidate_members)
    if source_archive_bytes <= source_compressed_bytes or candidate_archive_bytes <= candidate_compressed_bytes:
        fail(f"{label}: archive is shorter than its retained compressed member payload")
    for path in expected_paths - {MAIN_PATH}:
        if source_by_path[path] != candidate_by_path[path]:
            fail(f"{label}: untouched member {path!r} changed")
    if source_by_path[MAIN_PATH]["decoded_sha256"] == candidate_by_path[MAIN_PATH]["decoded_sha256"]:
        fail(f"{label}: main member digest did not change")
    source_semantic = value["source_semantic"]
    candidate_semantic = value["candidate_semantic"]
    for semantic, expected, semantic_label in (
        (source_semantic, _semantic(count), f"{label}.source_semantic"),
        (candidate_semantic, _semantic(count, True), f"{label}.candidate_semantic"),
    ):
        if not isinstance(semantic, dict) or set(semantic) != set(expected):
            fail(f"{semantic_label}: semantic fields differ")
        for key, wanted in expected.items():
            if semantic.get(key) != wanted:
                fail(f"{semantic_label}.{key}: expected {wanted!r}")
    for key in ("source_main_xml_archive_verified", "candidate_main_xml_archive_verified", "candidate_main_xml_oracle_verified", "source_unchanged_verified", "source_semantic_reopen_verified", "candidate_semantic_reopen_verified", "tail_copy_exactly_one_verified", "untouched_members_verified", "opaque_member_exact_verified", "physical_order_verified"):
        _bool(value.get(key), f"{label}.{key}")
    limits = _limits(value["limits"], count, len(source_xml), f"{label}.limits")
    patch = value["patch"]
    if not isinstance(patch, dict):
        fail(f"{label}.patch: expected object")
    patch_fields = {"copied_source_position", "copied_before_position", "copied_paragraphs", "copied_bytes", "durable_bytes", "durable_canonical_verified", "replay_verified", "inverse_verified", "stale_source_refusal_verified", "publication_inverse_verified"}
    if set(patch) != patch_fields:
        fail(f"{label}.patch: fields differ")
    for key in ("copied_source_position", "copied_before_position", "copied_paragraphs", "copied_bytes", "durable_bytes"):
        integer(patch.get(key), f"{label}.patch.{key}")
    if patch["copied_source_position"] != 0 or patch["copied_before_position"] != count or patch["copied_paragraphs"] != 1 or patch["copied_bytes"] != len(_fragment(0)):
        fail(f"{label}.patch: copy effect differs")
    if patch["durable_bytes"] == 0 or patch["durable_bytes"] > MAX_DURABLE_BYTES:
        fail(f"{label}.patch.durable_bytes: outside finite bound")
    for key in ("durable_canonical_verified", "replay_verified", "inverse_verified", "stale_source_refusal_verified", "publication_inverse_verified"):
        _bool(patch.get(key), f"{label}.patch.{key}")
    result = dict(value)
    result["source_members"] = source_members
    result["candidate_members"] = candidate_members
    result["limits"] = limits
    return result


PROCESS_FIELDS = (
    "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
    "syscr", "syscw", "minor_faults", "major_faults", "user_cpu_ticks",
    "system_cpu_ticks", "clock_ticks_per_second", "voluntary_context_switches",
    "nonvoluntary_context_switches", "rss_bytes", "peak_rss_bytes",
)
ALLOC_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
SINK_FIELDS = ("accepted_bytes", "write_calls", "largest_write", "histogram", "sha256")
HIST_FIELDS = ("bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536")
READ_FIELDS = ("calls", "requested_bytes", "returned_bytes", "request_histogram")


def _process(value: Any, label: str) -> dict[str, int]:
    if not isinstance(value, dict) or set(value) != set(PROCESS_FIELDS):
        fail(f"{label}: process observer fields differ")
    result = {key: u64(value[key], f"{label}.{key}") for key in PROCESS_FIELDS}
    if result["clock_ticks_per_second"] == 0:
        fail(f"{label}.clock_ticks_per_second: expected a positive clock rate")
    if result["peak_rss_bytes"] < result["rss_bytes"]:
        fail(f"{label}: RSS high-water mark is below the RSS delta")
    return result


def _reads(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict) or set(value) != set(READ_FIELDS):
        fail(f"{label}: logical ReadAt fields differ")
    result = {
        "calls": u64(value["calls"], f"{label}.calls"),
        "requested_bytes": u64(value["requested_bytes"], f"{label}.requested_bytes"),
        "returned_bytes": u64(value["returned_bytes"], f"{label}.returned_bytes"),
    }
    histogram = value["request_histogram"]
    if not isinstance(histogram, dict) or set(histogram) != set(HIST_FIELDS):
        fail(f"{label}.request_histogram: fields differ")
    for key in HIST_FIELDS:
        u64(histogram[key], f"{label}.request_histogram.{key}")
    if sum(histogram.values()) != result["calls"]:
        fail(f"{label}: logical-read request histogram does not sum to calls")
    if result["returned_bytes"] > result["requested_bytes"]:
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
        + histogram["bytes_over_65536"] * ((1 << 64) - 1)
    )
    if not lower <= result["requested_bytes"] <= upper or result["requested_bytes"] >= 1 << 64:
        fail(f"{label}: logical-read request byte total is outside histogram bounds")
    return {**result, "request_histogram": dict(histogram)}


def _read_add(left: Mapping[str, Any], right: Mapping[str, Any], label: str) -> dict[str, Any]:
    histogram = {
        field: left["request_histogram"][field] + right["request_histogram"][field]
        for field in HIST_FIELDS
    }
    return {
        "calls": left["calls"] + right["calls"],
        "requested_bytes": left["requested_bytes"] + right["requested_bytes"],
        "returned_bytes": left["returned_bytes"] + right["returned_bytes"],
        "request_histogram": histogram,
    }


def _read_equal(left: Mapping[str, Any], right: Mapping[str, Any], label: str) -> None:
    if left != right:
        fail(f"{label}: logical ReadAt observation differs")


def _allocation(value: Any, instrumentation: str, label: str, *, total: bool = False) -> dict[str, Any] | None:
    if value is None:
        if instrumentation == "allocator":
            fail(f"{label}: allocator allocation sample is missing")
        return None
    if not isinstance(value, dict):
        fail(f"{label}: allocation sample is malformed")
    if instrumentation == "normal":
        # The normal executable does not install the counting allocator and
        # normally serializes this Option as null.  If a producer chooses to
        # retain an explicit unavailable Sample, serde omits every numeric
        # Option field; accepting a null-filled numeric vector would blur the
        # distinction between unavailable and measured zero.
        if set(value) != {"status", "scope"}:
            fail(f"{label}: unavailable normal allocation must omit numeric fields")
        if value.get("status") != "unavailable" or value.get("scope") != "operation_global_system_allocator":
            fail(f"{label}: normal allocation must be explicitly unavailable")
        return dict(value)
    if set(value) != {"status", "scope", *ALLOC_FIELDS}:
        fail(f"{label}: allocation fields differ")
    if value.get("status") != "measured" or value.get("scope") != "operation_global_system_allocator":
        fail(f"{label}: allocator allocation identity differs")
    for key in ALLOC_FIELDS:
        u64(value.get(key), f"{label}.{key}")
    if value["failed_allocation_calls"] != 0:
        fail(f"{label}: allocator failure count is non-zero")
    if total and value["live_bytes_before"] != value["live_bytes_after"]:
        fail(f"{label}: total lifecycle did not return to its entry live bytes")
    if value["live_bytes_before"] + value["allocated_bytes"] - value["deallocated_bytes"] != value["live_bytes_after"]:
        fail(f"{label}: allocator live-byte conservation failed")
    if value["peak_live_bytes_before"] > value["peak_live_bytes_after"]:
        fail(f"{label}: process peak decreased across the region")
    if value["live_bytes_before"] > value["peak_live_bytes_before"] or value["live_bytes_after"] > value["peak_live_bytes_after"]:
        fail(f"{label}: process peak is below its boundary live bytes")
    if value["region_peak_live_bytes"] < value["live_bytes_before"] or value["region_peak_live_bytes"] < value["live_bytes_after"]:
        fail(f"{label}: region peak is below live bytes")
    if value["region_peak_live_bytes"] > value["peak_live_bytes_after"]:
        fail(f"{label}: region peak exceeds the process peak")
    return dict(value)


def _sink(value: Any, corpus: Mapping[str, Any], label: str) -> dict[str, Any]:
    if not isinstance(value, dict) or set(value) != set(SINK_FIELDS):
        fail(f"{label}: sink fields differ")
    accepted = u64(value["accepted_bytes"], f"{label}.accepted_bytes")
    calls = u64(value["write_calls"], f"{label}.write_calls")
    largest = u64(value["largest_write"], f"{label}.largest_write")
    if calls == 0 or largest == 0:
        fail(f"{label}: sink write counters must be positive")
    if accepted != corpus["candidate_archive_bytes"] or largest > accepted or largest > HASH_SINK_MAX_WRITE or accepted > calls * HASH_SINK_MAX_WRITE:
        fail(f"{label}: sink output bound differs from candidate archive")
    digest(value["sha256"], f"{label}.sha256")
    if value["sha256"] != corpus["candidate_archive_sha256"]:
        fail(f"{label}: sink digest differs from candidate archive")
    histogram = value["histogram"]
    if not isinstance(histogram, dict) or set(histogram) != set(HIST_FIELDS):
        fail(f"{label}.histogram: fields differ")
    for key in HIST_FIELDS:
        u64(histogram[key], f"{label}.histogram.{key}")
    if sum(histogram.values()) != calls:
        fail(f"{label}: sink histogram does not sum to write calls")
    if histogram["bytes_0"] != 0 or histogram["bytes_16385_to_65536"] != 0 or histogram["bytes_over_65536"] != 0:
        fail(f"{label}: sink write-size bound failed")
    return dict(value)


def _phase(value: Any, instrumentation: str, label: str) -> dict[str, Any]:
    if not isinstance(value, dict) or set(value) != {"elapsed_ns", "source_reads", "allocation", "process"}:
        fail(f"{label}: phase fields differ")
    positive(value["elapsed_ns"], f"{label}.elapsed_ns")
    reads = _reads(value["source_reads"], f"{label}.source_reads")
    allocation = _allocation(value["allocation"], instrumentation, f"{label}.allocation")
    process = _process(value["process"], f"{label}.process") if value["process"] is not None else fail(f"{label}.process: unavailable process observation")
    return {"elapsed_ns": value["elapsed_ns"], "source_reads": reads, "allocation": allocation, "process": process}


def _total_sample(value: Any, corpus: Mapping[str, Any], instrumentation: str, label: str) -> dict[str, Any]:
    fields = {"sample", "elapsed_ns", "source_reads", "sink", "allocation", "process"}
    if not isinstance(value, dict) or set(value) != fields:
        fail(f"{label}: total sample fields differ")
    integer(value["sample"], f"{label}.sample")
    positive(value["elapsed_ns"], f"{label}.elapsed_ns")
    reads = _reads(value["source_reads"], f"{label}.source_reads")
    if reads["calls"] == 0 or reads["requested_bytes"] == 0:
        fail(f"{label}.source_reads: total lifecycle must perform a source read")
    sink = _sink(value["sink"], corpus, f"{label}.sink")
    allocation = _allocation(value["allocation"], instrumentation, f"{label}.allocation", total=True)
    process = _process(value["process"], f"{label}.process") if value["process"] is not None else fail(f"{label}.process: unavailable process observation")
    return {"sample": value["sample"], "elapsed_ns": value["elapsed_ns"], "source_reads": reads, "sink": sink, "allocation": allocation, "process": process}


def _phase_sample(value: Any, corpus: Mapping[str, Any], instrumentation: str, label: str) -> dict[str, Any]:
    fields = {"sample", *PHASES, "source_reads", "sink"}
    if not isinstance(value, dict) or set(value) != fields:
        fail(f"{label}: phase sample fields differ")
    integer(value["sample"], f"{label}.sample")
    phases = {phase: _phase(value[phase], instrumentation, f"{label}.{phase}") for phase in PHASES}
    reads = _reads(value["source_reads"], f"{label}.source_reads")
    summed_reads = {"calls": 0, "requested_bytes": 0, "returned_bytes": 0,
                    "request_histogram": {field: 0 for field in HIST_FIELDS}}
    for phase in PHASES:
        summed_reads = _read_add(summed_reads, phases[phase]["source_reads"], f"{label}.{phase}")
    _read_equal(summed_reads, reads, f"{label}: phase ReadAt sum versus lifecycle")
    if instrumentation == "allocator":
        for previous, current in zip(PHASES, PHASES[1:]):
            left = phases[previous]["allocation"]
            right = phases[current]["allocation"]
            if left is None or right is None:
                fail(f"{label}: allocator phase sample is missing")
            if left["live_bytes_after"] != right["live_bytes_before"]:
                fail(f"{label}: phase live-byte boundaries are discontinuous")
        first = phases[PHASES[0]]["allocation"]
        last = phases[PHASES[-1]]["allocation"]
        if first is None or last is None or last["live_bytes_after"] != first["live_bytes_before"]:
            fail(f"{label}: phase lifecycle did not release to its entry live bytes")
    sink = _sink(value["sink"], corpus, f"{label}.sink")
    return {"sample": value["sample"], **phases, "source_reads": reads, "sink": sink}


def validate_report(report: Mapping[str, Any], spec: Mapping[str, Any], path: str = "report") -> dict[str, Any]:
    if report.get("schema") != REPORT_SCHEMA or report.get("version") != 1:
        fail(f"{path}: report schema/version differs")
    instrumentation = spec["instrumentation"]
    binary = report.get("binary")
    if not isinstance(binary, dict) or set(binary) != {"binary", "allocator", "instrumentation", "counter_revision"}:
        fail(f"{path}.binary: fields differ")
    expected = {
        "binary": "litchi-perf-baseline" if instrumentation == "normal" else "litchi-perf-baseline-alloc",
        "allocator": "Rust system allocator" if instrumentation == "normal" else "CountingSystemAllocator(std::alloc::System)",
        "instrumentation": "none" if instrumentation == "normal" else "system_allocator_operation_scoped",
        "counter_revision": None if instrumentation == "normal" else "serialized_region_peak_v3",
    }
    if binary != expected:
        fail(f"{path}.binary: instrumentation identity differs")
    config = report.get("config")
    required_config = {"counts", "samples", "warmups", "mode", "lifecycle_phases", "sink", "source"}
    if not isinstance(config, dict) or set(config) != required_config:
        fail(f"{path}.config: fields differ")
    if config["counts"] != [spec["count"]] or config["samples"] != SAMPLES or config["warmups"] != WARMUPS or config["mode"] != spec["mode"]:
        fail(f"{path}.config: selected count/mode/protocol differs")
    if config["lifecycle_phases"] != list(PHASES):
        fail(f"{path}.config.lifecycle_phases differs")
    if config["sink"] != SINK_ID or config["source"] != SOURCE_ID:
        fail(f"{path}.config source/sink identity differs")
    cases = report.get("cases")
    if not isinstance(cases, list) or len(cases) != 1:
        fail(f"{path}: expected one selected corpus case")
    case = cases[0]
    if not isinstance(case, dict) or set(case) != {"count", "corpus", "total_samples", "phase_samples"}:
        fail(f"{path}.cases[0]: fields differ")
    if case["count"] != spec["count"]:
        fail(f"{path}.cases[0].count differs")
    corpus = _corpus(case["corpus"], spec["count"], f"{path}.cases[0].corpus")
    if spec["mode"] == "total":
        if case["phase_samples"] is not None or not isinstance(case["total_samples"], list):
            fail(f"{path}: total mode sample vector shape differs")
        samples = case["total_samples"]
        if len(samples) != SAMPLES:
            fail(f"{path}: expected {SAMPLES} total samples")
        rows = [_total_sample(item, corpus, instrumentation, f"{path}.total_samples[{i}]") for i, item in enumerate(samples)]
    else:
        if case["total_samples"] is not None or not isinstance(case["phase_samples"], list):
            fail(f"{path}: phases mode sample vector shape differs")
        samples = case["phase_samples"]
        if len(samples) != SAMPLES:
            fail(f"{path}: expected {SAMPLES} phase samples")
        rows = [_phase_sample(item, corpus, instrumentation, f"{path}.phase_samples[{i}]") for i, item in enumerate(samples)]
    if [row["sample"] for row in rows] != list(range(SAMPLES)):
        fail(f"{path}: sample indices are not exactly 0..29")
    first_reads = rows[0]["source_reads"]
    first_sink = rows[0]["sink"]
    for index, row in enumerate(rows[1:], 1):
        _read_equal(first_reads, row["source_reads"], f"{path}.samples[{index}]: total logical-read identity drift")
        if row["sink"] != first_sink:
            fail(f"{path}.samples[{index}]: sink identity drift")
    if spec["mode"] == "phases":
        for phase in PHASES:
            first_phase_reads = rows[0][phase]["source_reads"]
            for index, row in enumerate(rows[1:], 1):
                _read_equal(first_phase_reads, row[phase]["source_reads"], f"{path}.samples[{index}].{phase}: phase logical-read identity drift")
    return {"schema": report["schema"], "version": report["version"], "binary": dict(binary), "config": dict(config), "count": spec["count"], "corpus": corpus, "samples": rows}


def _stats(values: list[int | float], *, signed: bool = False) -> dict[str, Any]:
    if len(values) != SAMPLES or any(
        isinstance(value, bool)
        or not isinstance(value, (int, float))
        or not math.isfinite(float(value))
        or (not signed and value < 0)
        for value in values
    ):
        fail("statistics input is invalid")
    ordered = sorted(values)
    mean = statistics.mean(values)
    half = 2.045 * statistics.stdev(values) / math.sqrt(len(values)) if len(set(values)) > 1 else 0.0
    return {
        "n": len(values), "mean": mean, "minimum": min(values), "maximum": max(values),
        "p50": ordered[math.ceil(len(values) * .50) - 1],
        "p95": ordered[math.ceil(len(values) * .95) - 1],
        "p99": ordered[math.ceil(len(values) * .99) - 1],
        "mean_t95_interval": [mean - half, mean + half],
    }


def _percent(before: float, after: float) -> float | None:
    return None if before == 0 else 100 * (after - before) / before


def _resource_rss(spec: Mapping[str, Any]) -> int:
    label = str(spec["label"])
    path = ROOT / "captures" / f"{label}.resource"
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as error:
        fail(f"{label}: GNU time resource file cannot be read: {error}")
    matches = re.findall(
        r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$",
        text,
        re.MULTILINE,
    )
    if len(matches) != 1:
        fail(f"{label}: expected exactly one GNU time RSS observation")
    return u64(int(matches[0]), f"{label}: process_max_rss_kib")


def _corpus_manifest() -> tuple[dict[str, Any], str]:
    path = ROOT / "corpus-manifest.json"
    if not path.is_file():
        fail("corpus-manifest.json is missing")
    value = read_json(path)
    if not isinstance(value, dict) or set(value) != {"schema", "frozen_utc", "cases", "pilots"}:
        fail("corpus-manifest.json: fields differ")
    if value["schema"] != CORPUS_MANIFEST_SCHEMA or not isinstance(value["frozen_utc"], str) or not value["frozen_utc"]:
        fail("corpus-manifest.json: schema/timestamp differs")
    corpora = value["cases"]
    if not isinstance(corpora, dict) or set(corpora) != {str(count) for count in COUNTS}:
        fail("corpus-manifest.json.cases: count keys differ")
    pilots = value["pilots"]
    expected_pilots = {f"pilot-{instrumentation}-{mode}" for instrumentation in INSTRUMENTATIONS for mode in MODES}
    if not isinstance(pilots, dict) or set(pilots) != expected_pilots:
        fail("corpus-manifest.json.pilots: labels differ")
    for count in COUNTS:
        record = corpora[str(count)]
        if not isinstance(record, dict):
            fail(f"corpus-manifest.json.cases.{count}: expected corpus object")
        # The freeze script stores the complete record directly under each
        # count.  Validate it here so a manifest cannot weaken the later
        # report-to-manifest identity comparison with a wrapper or partial
        # record.
        corpora[str(count)] = _corpus(record, count, f"corpus-manifest.json.cases.{count}")
    for label, reference in pilots.items():
        if not isinstance(reference, dict) or set(reference) != {"path", "bytes", "sha256"}:
            fail(f"corpus-manifest.json.pilots.{label}: reference differs")
        pilot_name = _text(reference["path"], f"corpus-manifest.json.pilots.{label}.path")
        if Path(pilot_name).name != pilot_name or Path(pilot_name).is_absolute():
            fail(f"corpus-manifest.json.pilots.{label}.path: report must be a local filename")
        positive(reference["bytes"], f"corpus-manifest.json.pilots.{label}.bytes")
        digest(reference["sha256"], f"corpus-manifest.json.pilots.{label}.sha256")
        pilot_path = ROOT / pilot_name
        if not pilot_path.is_file() or pilot_path.is_symlink():
            fail(f"corpus-manifest.json.pilots.{label}: retained pilot report is missing or not local")
        actual_bytes = pilot_path.stat().st_size
        actual_hash = hashlib.sha256(pilot_path.read_bytes()).hexdigest()
        if actual_bytes != reference["bytes"] or actual_hash != reference["sha256"]:
            fail(f"corpus-manifest.json.pilots.{label}: report metadata differs")
        pilot = read_json(pilot_path)
        if not isinstance(pilot, dict) or set(pilot) != {"schema", "version", "binary", "config", "cases"} or pilot.get("schema") != REPORT_SCHEMA or pilot.get("version") != 1:
            fail(f"corpus-manifest.json.pilots.{label}: report schema differs")
        instrumentation = "allocator" if label.startswith("pilot-allocator-") else "normal"
        mode = "phases" if label.endswith("-phases") else "total"
        expected_binary = {
            "binary": "litchi-perf-baseline" if instrumentation == "normal" else "litchi-perf-baseline-alloc",
            "allocator": "Rust system allocator" if instrumentation == "normal" else "CountingSystemAllocator(std::alloc::System)",
            "instrumentation": "none" if instrumentation == "normal" else "system_allocator_operation_scoped",
            "counter_revision": None if instrumentation == "normal" else "serialized_region_peak_v3",
        }
        if pilot.get("binary") != expected_binary:
            fail(f"corpus-manifest.json.pilots.{label}: binary identity differs")
        config = pilot.get("config")
        if not isinstance(config, dict) or set(config) != {"counts", "samples", "warmups", "mode", "lifecycle_phases", "sink", "source"} or config.get("counts") != list(COUNTS) or config.get("samples") != 1 or config.get("warmups") != 1 or config.get("mode") != mode or config.get("lifecycle_phases") != list(PHASES) or config.get("sink") != SINK_ID or config.get("source") != SOURCE_ID:
            fail(f"corpus-manifest.json.pilots.{label}: pilot protocol differs")
        cases = pilot.get("cases")
        if not isinstance(cases, list) or len(cases) != len(COUNTS) or any(
            not isinstance(case, dict) or set(case) != {"count", "corpus", "total_samples", "phase_samples"}
            for case in cases
        ) or [case["count"] for case in cases] != list(COUNTS):
            fail(f"corpus-manifest.json.pilots.{label}: cases differ")
        for case in cases:
            if case.get("corpus") != corpora[str(case["count"])]:
                fail(f"corpus-manifest.json.pilots.{label}: corpus case differs")
            corpus = corpora[str(case["count"])]
            if mode == "total":
                if case["phase_samples"] is not None or not isinstance(case["total_samples"], list) or len(case["total_samples"]) != 1:
                    fail(f"corpus-manifest.json.pilots.{label}: total sample vector differs")
                sample = _total_sample(case["total_samples"][0], corpus, instrumentation, f"corpus-manifest.json.pilots.{label}.count{case['count']}.total")
            else:
                if case["total_samples"] is not None or not isinstance(case["phase_samples"], list) or len(case["phase_samples"]) != 1:
                    fail(f"corpus-manifest.json.pilots.{label}: phase sample vector differs")
                sample = _phase_sample(case["phase_samples"][0], corpus, instrumentation, f"corpus-manifest.json.pilots.{label}.count{case['count']}.phases")
            if sample["sample"] != 0:
                fail(f"corpus-manifest.json.pilots.{label}: pilot sample index differs")

    # The corpus is frozen only after every pilot gate has finished and before
    # the first formal capture begins.  Keep this ordering in the portable
    # evidence itself so a later report cannot silently use a recaptured
    # corpus under an old manifest.
    pilot_finished = []
    for label in sorted(expected_pilots):
        receipt_path = ROOT / "validation" / f"{label}.json"
        receipt = read_json(receipt_path)
        if not isinstance(receipt, dict):
            fail(f"validation/{label}.json: pilot receipt must be an object")
        pilot_finished.append(_timestamp(receipt.get("finished_utc"), f"validation/{label}.finished_utc"))
    formal_started = []
    for spec in expected_captures():
        started_path = ROOT / "captures" / f"{spec['label']}.started.json"
        started = read_json(started_path)
        if not isinstance(started, dict):
            fail(f"captures/{spec['label']}.started.json: capture envelope must be an object")
        formal_started.append(_timestamp(started.get("started_utc"), f"captures/{spec['label']}.started_utc"))
    frozen = _timestamp(value["frozen_utc"], "corpus-manifest.json.frozen_utc")
    if frozen <= max(pilot_finished):
        fail("corpus-manifest.json was frozen before all pilot receipts finished")
    if frozen >= min(formal_started):
        fail("corpus-manifest.json was frozen after the first formal capture started")
    return value, hashlib.sha256(path.read_bytes()).hexdigest()


def _metric_rows(
    rows: list[dict[str, Any]],
    mode: str,
    instrumentation: str,
    corpus: Mapping[str, Any],
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "elapsed_ns": _stats([row["elapsed_ns"] for row in rows]) if mode == "total" else None,
        "phase_elapsed_ns": {phase: _stats([row[phase]["elapsed_ns"] for row in rows]) for phase in PHASES} if mode == "phases" else None,
        "source_read_calls": _stats([row["source_reads"]["calls"] for row in rows]),
        "source_read_requested_bytes": _stats([row["source_reads"]["requested_bytes"] for row in rows]),
        "source_read_returned_bytes": _stats([row["source_reads"]["returned_bytes"] for row in rows]),
        "source_read_request_histogram": {
            field: _stats([row["source_reads"]["request_histogram"][field] for row in rows]) for field in HIST_FIELDS
        },
        "sink_accepted_bytes": _stats([row["sink"]["accepted_bytes"] for row in rows]),
        "sink_write_calls": _stats([row["sink"]["write_calls"] for row in rows]),
        "sink_largest_write": _stats([row["sink"]["largest_write"] for row in rows]),
        "sink_write_histogram": {
            field: _stats([row["sink"]["histogram"][field] for row in rows]) for field in HIST_FIELDS
        },
        "process": {field: _stats([row["process"][field] for row in rows]) for field in PROCESS_FIELDS} if mode == "total" else {
            phase: {field: _stats([row[phase]["process"][field] for row in rows]) for field in PROCESS_FIELDS} for phase in PHASES
        },
        "allocation": None,
    }
    if mode == "total":
        allocations = [row["allocation"] for row in rows]
        if instrumentation == "normal":
            if any(item is not None for item in allocations):
                fail("normal total row unexpectedly exposes allocation numbers")
            result["allocation"] = {"status": "unavailable", "scope": "operation_global_system_allocator"}
        else:
            result["allocation"] = {
                field: _stats([item[field] for item in allocations if item is not None]) for field in ALLOC_FIELDS
            }
            result["total_peak_live_bytes"] = result["allocation"]["region_peak_live_bytes"]
            result["total_peak_live_bytes_scope"] = "absolute process allocator live-byte high-water mark"
            result["incremental_peak_live_bytes"] = _stats([
                item["region_peak_live_bytes"] - item["live_bytes_before"]
                for item in allocations if item is not None
            ])
        if mode == "total":
            elapsed = [row["elapsed_ns"] for row in rows]
            source_bytes = corpus["source_archive_bytes"]
            candidate_bytes = corpus["candidate_archive_bytes"]
            result["throughput"] = {
                "operations_per_second": _stats([1_000_000_000 / value for value in elapsed]),
                "source_archive_bytes_per_second": _stats([
                    source_bytes * 1_000_000_000 / value for value in elapsed
                ]),
                "candidate_archive_bytes_per_second": _stats([
                    candidate_bytes * 1_000_000_000 / value for value in elapsed
                ]),
                "source_logical_bytes_per_second": _stats([
                    corpus["source_main_xml_bytes"] * 1_000_000_000 / value for value in elapsed
                ]),
                "candidate_logical_bytes_per_second": _stats([
                    corpus["candidate_main_xml_bytes"] * 1_000_000_000 / value for value in elapsed
                ]),
            }
            result["throughput_scope"] = "one appended paragraph per lifecycle; logical bytes are source/candidate main XML bytes, archive-byte rates are retained separately, and count remains source paragraphs scanned"
    else:
        result["phase_allocation"] = {}
        result["phase_source_reads"] = {
            phase: {
                "calls": _stats([row[phase]["source_reads"]["calls"] for row in rows]),
                "requested_bytes": _stats([row[phase]["source_reads"]["requested_bytes"] for row in rows]),
                "returned_bytes": _stats([row[phase]["source_reads"]["returned_bytes"] for row in rows]),
                "request_histogram": {
                    field: _stats([row[phase]["source_reads"]["request_histogram"][field] for row in rows]) for field in HIST_FIELDS
                },
            } for phase in PHASES
        }
        for phase in PHASES:
            allocations = [row[phase]["allocation"] for row in rows]
            if instrumentation == "normal":
                if any(item is not None for item in allocations):
                    fail(f"normal phase {phase} unexpectedly exposes allocation numbers")
                result["phase_allocation"][phase] = {"status": "unavailable", "scope": "operation_global_system_allocator"}
            else:
                result["phase_allocation"][phase] = {
                    field: _stats([item[field] for item in allocations if item is not None])
                    for field in ALLOC_FIELDS
                }
                result["phase_allocation"][phase]["retained_delta"] = _stats([
                    item["live_bytes_after"] - item["live_bytes_before"]
                    for item in allocations if item is not None
                ], signed=True)
        result["phase_peaks_are_not_total_peak"] = True
    return result


def derive() -> dict[str, Any]:
    protocol = read_json(ROOT / "protocol.json")
    if not isinstance(protocol, dict):
        fail("protocol must be an object")
    specs = protocol_rows(protocol)
    corpus_manifest, corpus_manifest_sha256 = _corpus_manifest()
    reports: dict[str, dict[str, Any]] = {}
    identities: dict[int, dict[str, Any]] = {}
    rows: dict[str, dict[str, Any]] = {}
    for spec in specs:
        label = spec["label"]
        report_path = ROOT / "captures" / f"{label}.report.json"
        report = read_json(report_path)
        if not isinstance(report, dict):
            fail(f"{label}: report must be an object")
        checked = validate_report(report, spec, str(report_path))
        reports[label] = checked
        process_max_rss_kib = _resource_rss(spec)
        corpus = checked["corpus"]
        identity = json.loads(json.dumps(corpus, sort_keys=True))
        if spec["count"] in identities and identities[spec["count"]] != identity:
            fail(f"{label}: corpus identity changed across captures")
        identities[spec["count"]] = identity
        if corpus_manifest["cases"][str(spec["count"])] != identity:
            fail(f"{label}: corpus identity differs from frozen corpus-manifest.json")
        rows[label] = {
            "capture": spec,
            "corpus_identity": {
                "count": spec["count"], "source_archive_bytes": corpus["source_archive_bytes"],
                "source_archive_sha256": corpus["source_archive_sha256"],
                "candidate_archive_bytes": corpus["candidate_archive_bytes"],
                "candidate_archive_sha256": corpus["candidate_archive_sha256"],
            },
            # /usr/bin/time observes the complete child process, including
            # setup, corpus/oracle construction, warmups, report writing, and
            # teardown.  It is retained beside operation-local process
            # deltas and is never folded into latency or allocator rows.
            "process_max_rss_kib": process_max_rss_kib,
            **_metric_rows(checked["samples"], spec["mode"], spec["instrumentation"], corpus),
        }
    pairs: list[dict[str, Any]] = []
    for instrumentation in INSTRUMENTATIONS:
        for count in COUNTS:
            for repeat in REPEATS:
                total = rows[f"r{repeat}-{instrumentation}-{count}-total"]
                phases = rows[f"r{repeat}-{instrumentation}-{count}-phases"]
                # Total and phase runs must refer to the same oracle.  Their
                # latency vectors are intentionally not merged.
                if total["corpus_identity"] != phases["corpus_identity"]:
                    fail(f"{instrumentation}/{count}/r{repeat}: total and phases output identity differs")
                total_report = reports[f"r{repeat}-{instrumentation}-{count}-total"]
                phase_report = reports[f"r{repeat}-{instrumentation}-{count}-phases"]
                for index, (total_sample, phase_sample) in enumerate(zip(total_report["samples"], phase_report["samples"])):
                    _read_equal(total_sample["source_reads"], phase_sample["source_reads"], f"{instrumentation}/{count}/r{repeat}/sample{index}: total/phase lifecycle reads differ")
                    if total_sample["sink"] != phase_sample["sink"]:
                        fail(f"{instrumentation}/{count}/r{repeat}/sample{index}: total/phase sink differs")
                pairs.append({
                    "instrumentation": instrumentation, "count": count, "repeat": repeat,
                    "same_output_identity": True,
                    "total_elapsed_ns": total["elapsed_ns"],
                    "phase_elapsed_ns": phases["phase_elapsed_ns"],
                    "phase_sum_peak_claim": False,
                })
    repeat_drifts: list[dict[str, Any]] = []
    for instrumentation in INSTRUMENTATIONS:
        for count in COUNTS:
            for mode in MODES:
                first = rows[f"r1-{instrumentation}-{count}-{mode}"]
                second = rows[f"r2-{instrumentation}-{count}-{mode}"]
                if first["corpus_identity"] != second["corpus_identity"]:
                    fail(f"{instrumentation}/{count}/{mode}: repeat corpus identity differs")
                metric_changes: dict[str, float | None] = {
                    "process_max_rss_kib": _percent(first["process_max_rss_kib"], second["process_max_rss_kib"]),
                }
                if mode == "total":
                    for metric in ("mean", "p50", "p95", "p99"):
                        metric_changes[f"elapsed_{metric}"] = _percent(first["elapsed_ns"][metric], second["elapsed_ns"][metric])
                else:
                    for phase in PHASES:
                        for metric in ("mean", "p50", "p95", "p99"):
                            metric_changes[f"{phase}_elapsed_{metric}"] = _percent(first["phase_elapsed_ns"][phase][metric], second["phase_elapsed_ns"][phase][metric])
                repeat_drifts.append({
                    "instrumentation": instrumentation, "count": count, "mode": mode,
                    "percent_changes": metric_changes,
                    "absolute_review_flags": [key for key, value in metric_changes.items() if value is not None and abs(value) > REGRESSION_REVIEW_PERCENT],
                })
    growth: list[dict[str, Any]] = []
    for instrumentation in INSTRUMENTATIONS:
        for mode in MODES:
            lane = [rows[f"r{repeat}-{instrumentation}-{count}-{mode}"] for repeat in REPEATS for count in COUNTS]
            metric = "elapsed_ns" if mode == "total" else "phase_elapsed_ns"
            if mode == "total":
                growth.append({"instrumentation": instrumentation, "mode": mode, "repeat": 1, "means": {str(count): rows[f"r1-{instrumentation}-{count}-{mode}"][metric]["mean"] for count in COUNTS}, "claim": "scaling_cost_baseline"})
            else:
                growth.append({"instrumentation": instrumentation, "mode": mode, "repeat": 1, "means": {str(count): {phase: rows[f"r1-{instrumentation}-{count}-{mode}"][metric][phase]["mean"] for phase in PHASES} for count in COUNTS}, "claim": "scaling_cost_baseline"})
    return {
        "schema": SUMMARY_SCHEMA,
        "captures": len(rows),
        "samples": len(rows) * SAMPLES,
        "rows": rows,
        "corpus_manifest_sha256": corpus_manifest_sha256,
        "pairs": pairs,
        "repeat_drifts": repeat_drifts,
        "growth": growth,
        "growth_basis": "means use reverse-order repeat 1 only; repeat 2 remains an independent drift check",
        "whole_process_rss_scope": "GNU /usr/bin/time -v Maximum resident set size (kbytes), including setup, corpus/oracles, warmups, measured lifecycles, report serialization, and teardown; not operation-local RSS",
        "operation_process_rss_scope": "report process.rss_bytes is a saturating RSS delta and process.peak_rss_bytes is the absolute VmHWM after the observer interval",
        "allocation_byte_counter_scope": "allocated_bytes includes the full new_size on realloc callbacks; allocator-byte accounting carries no physical-copy claim",
        "uncertainty": "nearest-rank percentiles; mean interval uses t(29)=2.045; independent reverse-order process repeats retained; no latency or memory optimization claim",
        "phase_peak_rule": "phase high-water marks are attribution evidence and are never summed or presented as a total-operation peak",
        "normal_allocation": "unavailable; normal binary does not install the counting allocator",
    }


def main() -> int:
    try:
        write_json(ROOT / "summary.json", derive())
    except (AnalysisError, OSError) as error:
        print(f"analyze.py: FAIL: {error}")
        return 1
    print("analyze.py: wrote summary.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
