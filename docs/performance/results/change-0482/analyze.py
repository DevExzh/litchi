#!/usr/bin/env python3
"""Validate and summarize the bounded XML reader measurements.

The two routes are compared within one executable and one source revision.
The module intentionally keeps allocation and process RSS observations out of
the semantic equality check: those are the quantities being compared.  Every
other successful audit result (bytes, hashes, event/count fields, and any
format-level oracle data) must agree sample-for-sample before statistics are
written.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import functools
import hashlib
import json
import math
from pathlib import Path
import re
import statistics
from typing import Any, Mapping

from common import ROOT


# Keep the constants local so this file remains usable before a binary is built.
SCHEMA = "xml-stream-audit-comparison-v1"
REPORT_SCHEMA = "litchi.xml-stream-audit.v1"
SUMMARY_SCHEMA = "xml-stream-audit-comparison-summary-v1"
XML_ASSET_SCHEMA = "xml-stream-audit-xml-assets-v1"
XML_ASSET_INVENTORY_PATH = "xml-assets.json"
XML_ASSET_COUNT = 77
FUZZ_SCHEMA = "xml-stream-audit-fuzz-evidence-v1"
FUZZ_SEED_COUNT = 11
FUZZ_TARGET = "x86_64-unknown-linux-gnu"
FUZZ_DATA_PREFIX = "fuzz/accepted"
SIZES = (64 * 1024, 8 * 1024 * 1024, 128 * 1024 * 1024)
MODES = ("materialized", "streaming")
INSTRUMENTATIONS = ("normal", "allocator")
ARMS = ("materialized", "streaming")
REPEATS = (1, 2)
SAMPLES = 30
WARMUPS = 3
CPU = 2
REVIEW_PERCENT = 5.0
T_CRITICAL_29 = 2.045229642
REQUIRED_VALIDATION_LABELS = (
    "build-normal-accepted",
    "build-allocator-accepted",
    "format-xml-final-accepted",
    "xml-tests-final-accepted",
    "xml-no-default-final-accepted",
    "xml-clippy-final-accepted",
    "xml-rustdoc-final-accepted",
    "format-opc-final-accepted",
    "opc-tests-accepted",
    "opc-clippy-accepted",
    "opc-rustdoc-accepted",
    "workspace-check-accepted",
    "harness-tests-accepted",
    "harness-clippy-accepted",
    "boundaries-accepted",
    "registry-strict-accepted",
    "evidence-tests-final-v2-accepted",
    "analyze-final",
    "format-zip-final-accepted",
    "zip-tests-final-accepted",
    "zip-clippy-final-accepted",
    "zip-rustdoc-final-accepted",
    "xml-fuzz-build-accepted",
    "xml-fuzz-smoke-accepted",
)


class AnalysisError(ValueError):
    """Evidence is absent, malformed, or internally inconsistent."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def read_json(path: Path, label: str | None = None) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        fail(f"{label or path}: cannot read JSON: {error}")


def write_json_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def sha256_file(path: Path) -> str:
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError as error:
        fail(f"{path}: cannot hash: {error}")


def _integer(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label}: expected integer >= {minimum}")
    return value


def _number(value: Any, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        fail(f"{label}: expected number")
    if not math.isfinite(float(value)):
        fail(f"{label}: expected finite number")
    return float(value)


def _timestamp(value: Any, label: str) -> None:
    if not isinstance(value, str) or not value:
        fail(f"{label}: timestamp missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp: {error}")
    if parsed.tzinfo is None:
        fail(f"{label}: timestamp lacks timezone")


def expected_captures(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    captures = protocol.get("captures")
    if not isinstance(captures, list):
        fail("protocol.captures is missing")
    expected: list[dict[str, Any]] = []
    for instrumentation in INSTRUMENTATIONS:
        rows = [row for row in captures if isinstance(row, dict) and row.get("instrumentation") == instrumentation]
        if len(rows) != 12:
            fail(f"protocol has {len(rows)} {instrumentation} captures; expected 12")
        expected.extend(rows)
    return expected


def _validate_capture_argv(row: Mapping[str, Any], binary_path: str, index: int) -> None:
    argv = row.get("argv")
    label = f"protocol.captures[{index}].argv"
    if not isinstance(argv, list) or len(argv) != 18 or not all(isinstance(value, str) for value in argv):
        fail(f"{label}: expected the retained taskset/GNU-time command vector")
    expected = [
        "/usr/bin/time",
        "-v",
        "-o",
        None,
        "/usr/bin/taskset",
        "-c",
        str(CPU),
        binary_path,
        "--mode",
        row["mode"],
        "--sizes",
        str(row["size_bytes"]),
        "--samples",
        str(SAMPLES),
        "--warmup",
        str(WARMUPS),
        "--json",
        None,
    ]
    for position, expected_value in enumerate(expected):
        if expected_value is not None and argv[position] != expected_value:
            fail(f"{label}: command argument {position} differs")
    if not Path(argv[3]).is_absolute() or Path(argv[3]).parent.name != "captures" or Path(argv[3]).name != f"{row['label']}.resource":
        fail(f"{label}: GNU-time resource path differs")
    if not Path(argv[17]).is_absolute() or Path(argv[17]).parent.name != "captures" or Path(argv[17]).name != f"{row['label']}.report.json":
        fail(f"{label}: report path differs")


def _validate_fuzz_file_reference(value: Any, expected_path: str, label: str) -> None:
    if (
        not isinstance(value, dict)
        or set(value) != {"path", "bytes", "sha256"}
        or value.get("path") != expected_path
        or isinstance(value.get("bytes"), bool)
        or not isinstance(value.get("bytes"), int)
        or value["bytes"] <= 0
        or not isinstance(value.get("sha256"), str)
        or re.fullmatch(r"[0-9a-f]{64}", value["sha256"]) is None
    ):
        fail(f"{label}: retained file reference differs")


def _validate_fuzz_inventory_reference(value: Any, expected_path: str, minimum_files: int, label: str) -> None:
    if (
        not isinstance(value, dict)
        or set(value) != {"path", "files", "sha256"}
        or value.get("path") != expected_path
        or isinstance(value.get("files"), bool)
        or not isinstance(value.get("files"), int)
        or value["files"] < minimum_files
        or not isinstance(value.get("sha256"), str)
        or re.fullmatch(r"[0-9a-f]{64}", value["sha256"]) is None
    ):
        fail(f"{label}: retained directory inventory reference differs")


def validate_protocol(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    if protocol.get("schema") != SCHEMA or protocol.get("report_schema") != REPORT_SCHEMA:
        fail("protocol schema differs")
    if protocol.get("version") != 1:
        fail("protocol version differs")
    if protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS:
        fail("protocol sample counts differ")
    if protocol.get("cpu") != CPU or protocol.get("sizes") != list(SIZES):
        fail("protocol size or CPU contract differs")
    if protocol.get("modes") != list(MODES) or protocol.get("arms") != list(ARMS):
        fail("protocol route contract differs")
    if protocol.get("repeats") != list(REPEATS) or protocol.get("sequence") != "A1/B1/B2/A2":
        fail("protocol repeat sequence differs")
    if protocol.get("normal_and_allocator_timings_separate") is not True:
        fail("protocol does not separate normal and allocator timings")
    if protocol.get("no_cross_revision_claim") is not True:
        fail("protocol does not prohibit cross-revision claims")
    if protocol.get("performance_claim") != "primitive-enabler-only":
        fail("protocol performance scope differs")
    if protocol.get("pair_review_percent") != 5:
        fail("protocol pair-review threshold differs")
    xml_inventory = protocol.get("xml_asset_inventory")
    if (
        not isinstance(xml_inventory, dict)
        or set(xml_inventory) != {"path", "sha256", "files"}
        or xml_inventory.get("path") != XML_ASSET_INVENTORY_PATH
        or xml_inventory.get("files") != XML_ASSET_COUNT
        or not isinstance(xml_inventory.get("sha256"), str)
        or re.fullmatch(r"[0-9a-f]{64}", xml_inventory["sha256"]) is None
    ):
        fail("protocol XML asset inventory binding differs")
    fuzz = protocol.get("fuzz")
    if (
        not isinstance(fuzz, dict)
        or set(fuzz) != {
            "schema",
            "cwd",
            "helper",
            "prepared",
            "build",
            "smoke",
            "manifest",
            "lock",
            "seed_inventory",
            "post_run_inventory",
        }
        or fuzz.get("schema") != FUZZ_SCHEMA
        or not isinstance(fuzz.get("cwd"), str)
        or not Path(fuzz["cwd"]).is_absolute()
    ):
        fail("protocol fuzz evidence binding differs")
    _validate_fuzz_file_reference(fuzz.get("helper"), "fuzz.py", "protocol.fuzz.helper")
    for name, expected_path in (
        ("prepared", f"{FUZZ_DATA_PREFIX}/prepared.json"),
        ("build", f"{FUZZ_DATA_PREFIX}/build.json"),
        ("smoke", f"{FUZZ_DATA_PREFIX}/smoke.json"),
        ("manifest", f"{FUZZ_DATA_PREFIX}/build-inputs/Cargo.toml"),
        ("lock", f"{FUZZ_DATA_PREFIX}/build-inputs/Cargo.lock.txt"),
    ):
        _validate_fuzz_file_reference(fuzz.get(name), expected_path, f"protocol.fuzz.{name}")
    _validate_fuzz_inventory_reference(
        fuzz.get("seed_inventory"), "fuzz/seeds", FUZZ_SEED_COUNT, "protocol.fuzz.seed_inventory"
    )
    _validate_fuzz_inventory_reference(
        fuzz.get("post_run_inventory"), f"{FUZZ_DATA_PREFIX}/post-run", 1, "protocol.fuzz.post_run_inventory"
    )
    required_validation_labels = protocol.get("required_validation_labels")
    if (
        not isinstance(required_validation_labels, list)
        or len(set(required_validation_labels)) != len(required_validation_labels)
        or not all(isinstance(item, str) and item for item in required_validation_labels)
        or set(required_validation_labels) != set(REQUIRED_VALIDATION_LABELS)
    ):
        fail("protocol.required_validation_labels differs from the final gate policy")
    environment = protocol.get("environment")
    if not isinstance(environment, dict) or set(environment) != {"RUSTUP_TOOLCHAIN", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "CARGO_PROFILE_RELEASE_DEBUG", "RUSTFLAGS", "DEBUGINFOD_URLS", "LC_ALL"}:
        fail("protocol.environment fields differ")
    if not all(isinstance(item, str) for item in environment.values()):
        fail("protocol.environment values differ")
    binaries = protocol.get("binaries")
    if not isinstance(binaries, dict) or set(binaries) != set(INSTRUMENTATIONS):
        fail("protocol binaries differ")
    for instrumentation in INSTRUMENTATIONS:
        binary = binaries[instrumentation]
        if not isinstance(binary, dict) or set(binary) != {"path", "bytes", "sha256"}:
            fail(f"protocol.binaries.{instrumentation} fields differ")
        _integer(binary["bytes"], f"protocol.binaries.{instrumentation}.bytes", 1)
        if not isinstance(binary["path"], str) or not Path(binary["path"]).is_absolute():
            fail(f"protocol.binaries.{instrumentation}.path must be absolute")
        if not isinstance(binary["sha256"], str) or len(binary["sha256"]) != 64:
            fail(f"protocol.binaries.{instrumentation}.sha256 differs")
    builds = protocol.get("builds")
    if not isinstance(builds, dict) or set(builds) != set(INSTRUMENTATIONS):
        fail("protocol build bindings differ")
    for instrumentation in INSTRUMENTATIONS:
        reference = builds[instrumentation]
        if not isinstance(reference, dict) or set(reference) != {"path", "sha256"}:
            fail(f"protocol.builds.{instrumentation} fields differ")
        if reference["path"] != f"builds/{instrumentation}-accepted.json":
            fail(f"protocol.builds.{instrumentation}.path differs")
        if not isinstance(reference["sha256"], str) or len(reference["sha256"]) != 64:
            fail(f"protocol.builds.{instrumentation}.sha256 differs")
    captures = expected_captures(protocol)
    expected_labels = {
        f"{instrumentation}-{mode}-r{repeat}-{size}"
        for instrumentation in INSTRUMENTATIONS
        for arm, repeat, sequence in (
            ("materialized", 1, [("materialized", size) for size in SIZES]),
            ("streaming", 1, [("streaming", size) for size in SIZES]),
            ("streaming", 2, [("streaming", size) for size in reversed(SIZES)]),
            ("materialized", 2, [("materialized", size) for size in reversed(SIZES)]),
        )
        for mode, size in sequence
    }
    if {row.get("label") for row in captures} != expected_labels:
        fail("protocol capture labels differ")
    expected_order = []
    for instrumentation in INSTRUMENTATIONS:
        for mode, repeat, sizes in (
            ("materialized", 1, SIZES),
            ("streaming", 1, SIZES),
            ("streaming", 2, tuple(reversed(SIZES))),
            ("materialized", 2, tuple(reversed(SIZES))),
        ):
            expected_order.extend(f"{instrumentation}-{mode}-r{repeat}-{size}" for size in sizes)
    if [row.get("label") for row in captures] != expected_order:
        fail("protocol captures are not in the frozen A1/B1/B2/A2 order")
    for index, row in enumerate(captures):
        if set(row) != {"label", "instrumentation", "arm", "mode", "size_bytes", "repeat", "argv"}:
            fail(f"protocol.captures[{index}] fields differ")
        if row["instrumentation"] not in INSTRUMENTATIONS or row["arm"] not in ARMS:
            fail(f"protocol.captures[{index}] route differs")
        if row["mode"] != row["arm"] or row["size_bytes"] not in SIZES or row["repeat"] not in REPEATS:
            fail(f"protocol.captures[{index}] dimensions differ")
        _validate_capture_argv(row, binaries[row["instrumentation"]]["path"], index)
    return captures


def _resource_rss(path: Path) -> int:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as error:
        fail(f"{path}: cannot read GNU time resource file: {error}")
    matches = re.findall(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", text, re.MULTILINE)
    if len(matches) != 1:
        fail(f"{path}: expected exactly one GNU time RSS line")
    return int(matches[0]) * 1024


def _sample_list(case: Mapping[str, Any], label: str, expected_samples: int = SAMPLES) -> list[dict[str, Any]]:
    samples = case.get("samples")
    if not isinstance(samples, list):
        fail(f"{label}: report.samples is missing")
    if len(samples) != expected_samples:
        fail(f"{label}: expected {expected_samples} samples, got {len(samples)}")
    if not all(isinstance(sample, dict) for sample in samples):
        fail(f"{label}: sample is not an object")
    for index, sample in enumerate(samples):
        if sample.get("sample", index) != index:
            fail(f"{label}: sample index {index} differs")
    return samples  # type: ignore[return-value]


def _case(report: Mapping[str, Any], spec: Mapping[str, Any], label: str) -> Mapping[str, Any]:
    cases = report.get("cases")
    if not isinstance(cases, list) or len(cases) != 1 or not isinstance(cases[0], dict):
        fail(f"{label}: report must contain exactly one case")
    case = cases[0]
    if set(case) != {"mode", "size_bytes", "source", "limits", "warmups", "samples"}:
        fail(f"{label}: case fields differ from the live harness schema")
    if case.get("mode") != spec["mode"] or case.get("size_bytes") != spec["size_bytes"]:
        fail(f"{label}: case route or size differs")
    if not isinstance(case.get("source"), dict) or not isinstance(case.get("limits"), dict):
        fail(f"{label}: source or limits record is missing")
    return case


def _semantic(value: Any) -> Any:
    """Drop timing/instrumentation envelopes while retaining audit evidence."""

    if isinstance(value, dict):
        volatile = {
            "elapsed_ns",
            "duration_ns",
            "elapsed",
            "process",
            "allocation",
            "allocator",
            "rss_bytes",
            "peak_rss_bytes",
            "sample",
            "mode",
            "route",
            "instrumentation",
            "binary",
        }
        return {key: _semantic(item) for key, item in sorted(value.items()) if key not in volatile}
    if isinstance(value, list):
        return [_semantic(item) for item in value]
    return value


def _audit_projection(report: Mapping[str, Any], case: Mapping[str, Any], samples: list[dict[str, Any]], label: str) -> list[Any]:
    top: dict[str, Any] = {}
    for key in ("source", "limits", "corpus", "oracle", "input", "output", "audit"):
        if key in case:
            top[key] = case[key]
        elif key in report:
            top[key] = report[key]
    projections: list[Any] = []
    for sample in samples:
        selected = None
        for key in ("outcome", "audit", "result", "report", "output", "oracle", "counters"):
            if key in sample:
                selected = sample[key]
                break
        if selected is None:
            # A compact report may put its five audit counters directly in the
            # sample.  Keep only those stable fields in that case.
            names = ("bytes", "events", "attributes", "max_depth", "text_bytes", "sha256", "input_sha256", "output_sha256")
            selected = {key: sample[key] for key in names if key in sample}
        if not selected:
            fail(f"{label}: no stable audit result in sample")
        projections.append(_semantic(selected))
    if not top and not projections:
        fail(f"{label}: report has no stable result")
    return [_semantic(top), projections]


def _find_numeric(value: Any, names: set[str]) -> int | None:
    if isinstance(value, dict):
        for key, item in value.items():
            if key in names and isinstance(item, int) and not isinstance(item, bool):
                return item
        for item in value.values():
            found = _find_numeric(item, names)
            if found is not None:
                return found
    elif isinstance(value, list):
        for item in value:
            found = _find_numeric(item, names)
            if found is not None:
                return found
    return None


def _allocation_sample(sample: Mapping[str, Any]) -> Mapping[str, Any] | None:
    value = sample.get("allocation")
    if isinstance(value, dict):
        return value
    value = sample.get("counters")
    if isinstance(value, dict) and any(key in value for key in ("live_bytes_before", "live_bytes_after", "region_peak_live_bytes")):
        return value
    return None


def _expected_layout(size: int) -> dict[str, int]:
    """Return the fixed generator layout encoded by the audit harness."""

    prefix_bytes = 6  # ``<root>``
    suffix_bytes = 7  # ``</root>``
    item_open_bytes = 6  # ``<item>``
    item_close_bytes = 7  # ``</item>``
    full_record_bytes = 1024
    minimum_record_bytes = item_open_bytes + item_close_bytes
    body = size - prefix_bytes - suffix_bytes
    full_record_count, remainder = divmod(body, full_record_bytes)
    if remainder == 0:
        tail_record_bytes = 0
    elif remainder >= minimum_record_bytes:
        tail_record_bytes = remainder
    else:
        full_record_count -= 1
        tail_record_bytes = full_record_bytes + remainder
    record_count = full_record_count + int(tail_record_bytes != 0)
    return {
        "prefix_bytes": prefix_bytes,
        "suffix_bytes": suffix_bytes,
        "full_record_bytes": full_record_bytes,
        "full_record_count": full_record_count,
        "tail_record_bytes": tail_record_bytes,
        "record_count": record_count,
        "text_bytes": size - prefix_bytes - suffix_bytes - record_count * minimum_record_bytes,
        "item_open_bytes": item_open_bytes,
        "item_close_bytes": item_close_bytes,
    }


@functools.lru_cache(maxsize=None)
def _expected_source_sha256(size: int) -> str:
    """Independently hash the fixed generator pattern without a full payload.

    The Rust generator casts each record index to ``u8`` before applying the
    alphabet modulo.  Keep that truncation explicit here instead of reducing
    the unbounded Python index directly.
    """

    layout = _expected_layout(size)
    digest = hashlib.sha256()
    digest.update(b"<root>")
    for record in range(layout["record_count"]):
        record_bytes = (
            layout["full_record_bytes"]
            if record < layout["full_record_count"]
            else layout["tail_record_bytes"]
        )
        text_bytes = record_bytes - layout["item_open_bytes"] - layout["item_close_bytes"]
        digest.update(b"<item>")
        alphabet_offset = (record % 256) % 26
        digest.update(bytes((ord("a") + alphabet_offset,)) * text_bytes)
        digest.update(b"</item>")
    digest.update(b"</root>")
    return digest.hexdigest()


def _validate_source(case: Mapping[str, Any], size: int, label: str) -> None:
    source = case.get("source")
    if not isinstance(source, dict):
        fail(f"{label}: source identity is missing")
    if set(source) != {
        "requested_bytes", "generated_bytes", "sha256", "actual_generator_bytes", "actual_generator_sha256",
        "actual_generator_oracle_verified", "generator", "chunk_bytes", "source_materialization", "layout",
    }:
        fail(f"{label}: source fields differ from the live harness schema")
    if source.get("requested_bytes") != size or source.get("generated_bytes") != size:
        fail(f"{label}: generated source byte count differs")
    if source.get("actual_generator_bytes") != size:
        fail(f"{label}: actual generator byte count differs")
    if source.get("actual_generator_oracle_verified") is not True:
        fail(f"{label}: generator oracle was not verified")
    if source.get("source_materialization") != "materialized_leg_only":
        fail(f"{label}: source materialization identity differs")
    if source.get("generator") != "litchi-xml-repetitive-root-items-v1" or source.get("chunk_bytes") != 16 * 1024:
        fail(f"{label}: generator identity differs")
    for name in ("sha256", "actual_generator_sha256"):
        value = source.get(name)
        if not isinstance(value, str) or re.fullmatch(r"[0-9a-f]{64}", value) is None:
            fail(f"{label}: source {name} digest is malformed")
    if source["sha256"] != source["actual_generator_sha256"]:
        fail(f"{label}: source oracle digest differs")
    if source["sha256"] != _expected_source_sha256(size):
        fail(f"{label}: source digest does not match the independent generator pattern")
    layout = source.get("layout")
    if not isinstance(layout, dict) or set(layout) != {
        "prefix_bytes", "suffix_bytes", "full_record_bytes", "full_record_count", "tail_record_bytes",
        "record_count", "text_bytes", "item_open_bytes", "item_close_bytes",
    } or layout != _expected_layout(size):
        fail(f"{label}: deterministic source layout differs")


def _validate_limits(case: Mapping[str, Any], size: int, label: str) -> None:
    limits = case.get("limits")
    if not isinstance(limits, dict):
        fail(f"{label}: limit profile is missing")
    if set(limits) != {
        "max_bytes", "max_depth", "max_events", "max_attributes", "max_token_bytes", "max_text_bytes",
        "streaming_memory_upper_bound",
    }:
        fail(f"{label}: limit fields differ from the live harness schema")
    expected = {
        "max_bytes": size,
        "max_depth": 8,
        "max_events": 4_000_000,
        "max_attributes": 16,
        "max_token_bytes": 16 * 1024,
        "max_text_bytes": size,
    }
    if any(limits.get(name) != value for name, value in expected.items()):
        fail(f"{label}: limit profile differs")
    if not isinstance(limits.get("streaming_memory_upper_bound"), int) or limits["streaming_memory_upper_bound"] <= 0:
        fail(f"{label}: streaming memory bound is missing")


def _elapsed(sample: Mapping[str, Any], label: str) -> int:
    for key in ("elapsed_ns", "duration_ns"):
        if key in sample:
            return _integer(sample[key], f"{label}.{key}", 1)
    fail(f"{label}: elapsed_ns is missing")


def _validate_report(report: Mapping[str, Any], spec: Mapping[str, Any], resource: Path) -> dict[str, Any]:
    label = str(spec["label"])
    if set(report) != {"schema", "benchmark", "binary", "modes", "sizes", "samples", "warmups", "timed_scope", "corpus_scope", "cases"}:
        fail(f"{label}: report fields differ from the live harness schema")
    if report.get("schema") != REPORT_SCHEMA:
        fail(f"{label}: report schema differs")
    if report.get("benchmark") != "litchi-xml-repetitive-root-items-v1":
        fail(f"{label}: report benchmark differs")
    if report.get("timed_scope") != "generator_create_read_materialize_or_stream_audit_and_drop":
        fail(f"{label}: timed scope differs")
    if report.get("corpus_scope") != "deterministic_generator_hash_and_layout_are_precomputed_outside_timed_regions":
        fail(f"{label}: corpus scope differs")
    binary = report.get("binary")
    expected_instrumentation = "none" if spec["instrumentation"] == "normal" else "system_allocator_operation_scoped"
    expected_binary = {
        "binary": "litchi-perf-baseline" if spec["instrumentation"] == "normal" else "litchi-perf-baseline-alloc",
        "allocator": "Rust system allocator" if spec["instrumentation"] == "normal" else "CountingSystemAllocator(std::alloc::System)",
        "instrumentation": expected_instrumentation,
        "counter_revision": None if spec["instrumentation"] == "normal" else "serialized_region_peak_v3",
    }
    if binary != expected_binary:
        fail(f"{label}: report instrumentation differs")
    case = _case(report, spec, label)
    mode = case.get("mode")
    size = case.get("size_bytes")
    if report.get("modes") != [spec["mode"]] or report.get("sizes") != [spec["size_bytes"]]:
        fail(f"{label}: top-level mode or size vector differs")
    expected_samples = spec.get("samples", SAMPLES)
    expected_warmups = spec.get("warmups", WARMUPS)
    if (
        isinstance(expected_samples, bool)
        or not isinstance(expected_samples, int)
        or expected_samples <= 0
        or isinstance(expected_warmups, bool)
        or not isinstance(expected_warmups, int)
        or expected_warmups < 0
    ):
        fail(f"{label}: pilot/formal sample protocol is malformed")
    samples_count = report.get("samples")
    warmups = report.get("warmups")
    if mode != spec["mode"] or size != spec["size_bytes"]:
        fail(f"{label}: report route or size differs")
    _validate_source(case, size, label)
    _validate_limits(case, size, label)
    if samples_count != expected_samples or warmups != expected_warmups:
        fail(f"{label}: report sample protocol differs")
    samples = _sample_list(case, label, expected_samples)
    if case.get("warmups") != expected_warmups:
        fail(f"{label}: case warmup count differs")
    semantic = _audit_projection(report, case, samples, label)
    for index, sample in enumerate(samples):
        if set(sample) != {"sample", "elapsed_ns", "actual_bytes", "source_bytes_match", "outcome", "allocation", "operation"}:
            fail(f"{label}/sample{index}: fields differ from the live harness schema")
        if sample.get("source_bytes_match") is not True or sample.get("actual_bytes") != spec["size_bytes"]:
            fail(f"{label}/sample{index}: generated source byte identity differs")
        outcome = sample.get("outcome")
        if not isinstance(outcome, dict) or outcome.get("status") != "success" or not isinstance(outcome.get("report"), dict):
            fail(f"{label}/sample{index}: audit did not report success")
        if set(outcome) != {"status", "report"} or set(outcome["report"]) != {"attributes", "bytes", "events", "max_depth", "text_bytes"}:
            fail(f"{label}/sample{index}: audit result fields differ")
        operation = sample.get("operation")
        if not isinstance(operation, dict) or set(operation) != {"source_materialization_bytes", "generator_buffer_bytes", "generator_position", "entry_live_bytes", "exit_live_bytes", "net_live_bytes", "zero_net_live"}:
            fail(f"{label}/sample{index}: operation accounting is missing")
        expected_materialization = size if spec["mode"] == "materialized" else None
        if (
            operation.get("source_materialization_bytes") != expected_materialization
            or operation.get("generator_buffer_bytes") != 16 * 1024
            or operation.get("generator_position") != spec["size_bytes"]
        ):
            fail(f"{label}/sample{index}: generator position differs")
        net = operation.get("net_live_bytes")
        if net is not None and net != 0:
            fail(f"{label}/sample{index}: operation net live bytes are non-zero")
        if operation.get("zero_net_live") is False:
            fail(f"{label}/sample{index}: operation reports non-zero live bytes")
        if spec["instrumentation"] == "allocator" and (
            operation.get("net_live_bytes") != 0 or operation.get("zero_net_live") is not True
        ):
            fail(f"{label}/sample{index}: allocator operation lacks zero-net-live proof")
    elapsed = [_elapsed(sample, f"{label}/sample{index}") for index, sample in enumerate(samples)]
    allocation = []
    for index, sample in enumerate(samples):
        observation = _allocation_sample(sample)
        if spec["instrumentation"] == "normal":
            if sample.get("allocation") is not None or observation is not None:
                fail(f"{label}/sample{index}: normal binary unexpectedly reported allocator counters")
            allocation.append(None)
            continue
        if (
            observation is None
            or set(observation) != {
                "status", "scope", "allocation_calls", "deallocation_calls", "reallocation_calls",
                "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before",
                "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes",
            }
            or observation.get("status") != "measured"
        ):
            fail(f"{label}/sample{index}: allocator sample is not measured")
        if observation.get("scope") != "operation_global_system_allocator":
            fail(f"{label}/sample{index}: allocator scope differs")
        status = observation.get("status")
        if status in ("overflow", "Overflow"):
            fail(f"{label}/sample{index}: allocator observation overflowed")
        before = _find_numeric(observation, {"live_bytes_before"})
        after = _find_numeric(observation, {"live_bytes_after"})
        if before is None or after is None:
            fail(f"{label}/sample{index}: measured allocator sample lacks live-byte endpoints")
        if before != after:
            fail(f"{label}/sample{index}: allocator live-byte conservation failed")
        for field in (
            "allocation_calls",
            "deallocation_calls",
            "reallocation_calls",
            "failed_allocation_calls",
            "allocated_bytes",
            "deallocated_bytes",
            "peak_live_bytes_before",
            "peak_live_bytes_after",
            "region_peak_live_bytes",
        ):
            if not isinstance(observation.get(field), int) or isinstance(observation.get(field), bool) or observation[field] < 0:
                fail(f"{label}/sample{index}: allocator {field} is missing or invalid")
        region_peak = observation["region_peak_live_bytes"]
        if region_peak < before or region_peak < after:
            fail(f"{label}/sample{index}: allocator region peak precedes live endpoints")
        allocation.append(
            {
                "status": status,
                "scope": observation.get("scope"),
                "allocation_calls": observation["allocation_calls"],
                "deallocation_calls": observation["deallocation_calls"],
                "reallocation_calls": observation["reallocation_calls"],
                "failed_allocation_calls": observation["failed_allocation_calls"],
                "allocated_bytes": observation["allocated_bytes"],
                "deallocated_bytes": observation["deallocated_bytes"],
                "live_bytes_before": before,
                "live_bytes_after": after,
                "net_live_bytes": None if before is None or after is None else after - before,
                "peak_live_bytes_before": observation["peak_live_bytes_before"],
                "peak_live_bytes_after": observation["peak_live_bytes_after"],
                "region_peak_live_bytes": _find_numeric(observation, {"region_peak_live_bytes"}),
                "region_peak_increment_bytes": region_peak - before,
            }
        )
    return {
        "label": label,
        "instrumentation": spec["instrumentation"],
        "mode": spec["mode"],
        "size_bytes": spec["size_bytes"],
        "repeat": spec["repeat"],
        "report": report,
        "samples": samples,
        "semantic": semantic,
        "semantic_sha256": hashlib.sha256(json.dumps(semantic, sort_keys=True, separators=(",", ":")).encode()).hexdigest(),
        "elapsed": elapsed,
        "allocation": allocation,
        "rss_bytes": _resource_rss(resource),
    }


def load_rows(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for spec in expected_captures(protocol):
        prefix = ROOT / "captures" / spec["label"]
        terminal = read_json(prefix.with_suffix(".json"), f"{spec['label']}.receipt")
        if not isinstance(terminal, dict):
            fail(f"{spec['label']}: terminal receipt is not an object")
        if terminal.get("exit_code") != 0:
            fail(f"{spec['label']}: process failed with exit {terminal.get('exit_code')}")
        if terminal.get("source_unchanged") is not True or terminal.get("binary_unchanged") is not True:
            fail(f"{spec['label']}: source or binary changed during process")
        report_path = prefix.with_suffix(".report.json")
        resource_path = prefix.with_suffix(".resource")
        report = read_json(report_path, f"{spec['label']}.report")
        if not isinstance(report, dict):
            fail(f"{spec['label']}: report is not an object")
        artifacts = terminal.get("artifacts")
        if not isinstance(artifacts, dict):
            fail(f"{spec['label']}: receipt artifacts missing")
        report_meta = artifacts.get(report_path.name)
        if not isinstance(report_meta, dict) or report_meta.get("sha256") != sha256_file(report_path):
            fail(f"{spec['label']}: receipt/report digest mismatch")
        rows.append(_validate_report(report, spec, resource_path))
    return rows


def _percent(numerator: float, denominator: float) -> float | None:
    if denominator == 0:
        return None
    return (numerator - denominator) / denominator * 100.0


def _nearest_rank(values: list[int], quantile: float) -> int:
    ordered = sorted(values)
    index = max(0, min(len(ordered) - 1, math.ceil(quantile * len(ordered)) - 1))
    return ordered[index]


def _stats(values: list[int]) -> dict[str, Any]:
    if not values:
        fail("cannot summarize an empty sample vector")
    mean = statistics.fmean(values)
    stdev = statistics.stdev(values) if len(values) > 1 else 0.0
    half = T_CRITICAL_29 * stdev / math.sqrt(len(values)) if len(values) > 1 else 0.0
    return {
        "n": len(values),
        "mean_ns": mean,
        "stdev_ns": stdev,
        "ci95_mean_ns": [mean - half, mean + half],
        "p50_ns": _nearest_rank(values, 0.50),
        "p95_ns": _nearest_rank(values, 0.95),
        "p99_ns": _nearest_rank(values, 0.99),
        "min_ns": min(values),
        "max_ns": max(values),
    }


def _value_stats(values: list[int]) -> dict[str, Any]:
    if not values:
        fail("cannot summarize an empty counter vector")
    return {
        "n": len(values),
        "mean": statistics.fmean(values),
        "min": min(values),
        "max": max(values),
        "p50": _nearest_rank(values, 0.50),
        "p95": _nearest_rank(values, 0.95),
        "p99": _nearest_rank(values, 0.99),
    }


def _allocator_summary(row: Mapping[str, Any]) -> dict[str, Any] | None:
    observations = row["allocation"]
    if not observations or all(observation is None for observation in observations):
        return None
    if any(observation is None for observation in observations):
        fail("allocator observation vector mixes measured and absent samples")
    fields = (
        "allocation_calls",
        "deallocation_calls",
        "reallocation_calls",
        "failed_allocation_calls",
        "allocated_bytes",
        "deallocated_bytes",
        "live_bytes_before",
        "live_bytes_after",
        "net_live_bytes",
        "peak_live_bytes_before",
        "peak_live_bytes_after",
        "region_peak_live_bytes",
        "region_peak_increment_bytes",
    )
    derived = {
        field: _value_stats([int(observation[field]) for observation in observations])
        for field in fields
    }
    derived["requested_bytes"] = row["size_bytes"]
    return derived


def _row_metric(row: Mapping[str, Any]) -> dict[str, Any]:
    latency = _stats(list(row["elapsed"]))
    return {
        "instrumentation": row["instrumentation"],
        "mode": row["mode"],
        "size_bytes": row["size_bytes"],
        "repeat": row["repeat"],
        "requested_bytes": row["size_bytes"],
        "latency": latency,
        "rss_bytes": row["rss_bytes"],
        "rss_kib": row["rss_bytes"] / 1024.0,
        "semantic": row["semantic"],
        "semantic_sha256": row["semantic_sha256"],
        "allocator": _allocator_summary(row),
        "allocator_raw": row["allocation"],
    }


def _pair_metric(a: Mapping[str, Any], b: Mapping[str, Any], kind: str) -> dict[str, Any]:
    if a["semantic"] != b["semantic"]:
        fail(f"{kind}: successful report equality/hash/count contract failed")
    delta_latency = _percent(b["latency"]["mean_ns"], a["latency"]["mean_ns"])
    delta_rss = _percent(float(b["rss_bytes"]), float(a["rss_bytes"]))
    flags = []
    if delta_latency is not None and abs(delta_latency) > REVIEW_PERCENT:
        flags.append({"metric": "latency_mean", "delta_percent": delta_latency, "threshold_percent": REVIEW_PERCENT})
    if delta_rss is not None and abs(delta_rss) > REVIEW_PERCENT:
        flags.append({"metric": "process_rss", "delta_percent": delta_rss, "threshold_percent": REVIEW_PERCENT})
    return {
        "instrumentation": a["instrumentation"],
        "size_bytes": a["size_bytes"],
        "repeat": a["repeat"],
        "materialized": {"mean_ns": a["latency"]["mean_ns"], "rss_bytes": a["rss_bytes"]},
        "streaming": {"mean_ns": b["latency"]["mean_ns"], "rss_bytes": b["rss_bytes"]},
        "delta_percent_streaming_vs_materialized": {
            "latency_mean": delta_latency,
            "process_rss": delta_rss,
        },
        "flags_over_5_percent": flags,
        "semantic_sha256": a["semantic_sha256"],
    }


def _repeat_metric(rows: Mapping[tuple[str, int], Mapping[str, Any]], instrumentation: str, mode: str, size: int) -> dict[str, Any]:
    first = rows[(mode, 1)]
    second = rows[(mode, 2)]
    delta_latency = _percent(second["latency"]["mean_ns"], first["latency"]["mean_ns"])
    delta_rss = _percent(float(second["rss_bytes"]), float(first["rss_bytes"]))
    flags = []
    if delta_latency is not None and abs(delta_latency) > REVIEW_PERCENT:
        flags.append({"metric": "latency_mean", "delta_percent": delta_latency, "threshold_percent": REVIEW_PERCENT})
    if delta_rss is not None and abs(delta_rss) > REVIEW_PERCENT:
        flags.append({"metric": "process_rss", "delta_percent": delta_rss, "threshold_percent": REVIEW_PERCENT})
    return {
        "instrumentation": instrumentation,
        "mode": mode,
        "size_bytes": size,
        "repeat_1": {"mean_ns": first["latency"]["mean_ns"], "rss_bytes": first["rss_bytes"]},
        "repeat_2": {"mean_ns": second["latency"]["mean_ns"], "rss_bytes": second["rss_bytes"]},
        "delta_percent_repeat_2_vs_repeat_1": {"latency_mean": delta_latency, "process_rss": delta_rss},
        "flags_over_5_percent": flags,
    }


def summarize(protocol: Mapping[str, Any], rows: list[dict[str, Any]]) -> dict[str, Any]:
    by_key = {(row["instrumentation"], row["mode"], row["size_bytes"], row["repeat"]): row for row in rows}
    if len(by_key) != len(rows):
        fail("duplicate measurement row")
    metric_by_key = {key: _row_metric(row) for key, row in by_key.items()}
    metrics = [metric_by_key[key] for key in sorted(metric_by_key)]
    pairs = []
    repeats = []
    for instrumentation in INSTRUMENTATIONS:
        for size in SIZES:
            for repeat in REPEATS:
                materialized = metric_by_key[(instrumentation, "materialized", size, repeat)]
                streaming = metric_by_key[(instrumentation, "streaming", size, repeat)]
                pairs.append(_pair_metric(materialized, streaming, f"{instrumentation}/{size}/r{repeat}"))
            local = {(mode, repeat): metric_by_key[(instrumentation, mode, size, repeat)] for mode in MODES for repeat in REPEATS}
            for mode in MODES:
                repeats.append(_repeat_metric({(mode, repeat): local[(mode, repeat)] for repeat in REPEATS}, instrumentation, mode, size))
    return {
        "schema": SUMMARY_SCHEMA,
        "version": 1,
        "protocol_sha256": sha256_file(ROOT / "protocol.json"),
        "source_scope": "same executable and source revision; materialized versus streaming XML audit route",
        "performance_claim": "primitive-enabler-only",
        "normal_and_allocator_timings_separate": True,
        "process_rss_scope": "GNU time -v per-process maximum RSS; setup and teardown included",
        "latency_scope": "harness-reported operation elapsed_ns; warmups excluded",
        "statistics": {"samples": SAMPLES, "warmups": WARMUPS, "t_critical_df29": T_CRITICAL_29, "percentiles": "nearest-rank", "review_threshold_percent": REVIEW_PERCENT},
        "rows": metrics,
        "pairs": pairs,
        "repeat_drifts": repeats,
        "flags_over_5_percent": [
            item
            for item in pairs + repeats
            if item["flags_over_5_percent"]
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-only", action="store_true", help="recompute and print without writing summary.json")
    parser.add_argument("--output", type=Path, default=ROOT / "summary.json")
    args = parser.parse_args()
    try:
        protocol = read_json(ROOT / "protocol.json", "protocol.json")
        if not isinstance(protocol, dict):
            fail("protocol.json is not an object")
        validate_protocol(protocol)
        rows = load_rows(protocol)
        summary = summarize(protocol, rows)
        if not args.data_only:
            write_json_exclusive(args.output, summary)
        print(json.dumps(summary, indent=2, sort_keys=True))
        return 0
    except AnalysisError as error:
        print(f"analyze.py: FAIL: {error}", file=__import__("sys").stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
