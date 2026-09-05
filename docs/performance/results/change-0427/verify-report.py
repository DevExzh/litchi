#!/usr/bin/env python3
"""Verify one portable 0427 PPTX allocator-retention report.

The Rust retention subcommand emits one fresh-process lane: one API, one
corpus, and its raw measured rows.  This checker validates the phase contract,
absolute callback journal, and correctness gates.  It intentionally leaves
source, binary, corpus, and output cross-report binding to the capture
verifier.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any


SCHEMA = "pptx_retention_v1"
COUNTER_REVISION = "serialized_region_peak_v3"
COUNTER_SCOPE = "operation_global_system_allocator"
ALLOCATOR_IDENTITY = "CountingSystemAllocator(std::alloc::System)"
INSTRUMENTATION_IDENTITY = "system_allocator_operation_scoped"
CALLBACK_SCOPE = (
    "global allocator callbacks after System returns; process-wide across threads; "
    "excludes RSS, object-owned memory, allocator-internal realloc overlap, and latency"
)
OWNERSHIP_SCOPE = (
    "phase labels describe retained handles at callback-order boundaries; live bytes are "
    "process-global allocator counters and are not object-retention measurements"
)
MAX_SAMPLES = 1_000
MAX_WARMUP = 1_000
U64_MAX = (1 << 64) - 1
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")

CHECKPOINT_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes",
    "peak_live_bytes",
    "overflowed",
    "observer_invalid",
)
REGION_FIELDS = (
    "status",
    "scope",
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
    "retention_probe_region_peak_live_bytes",
)

OWNED_PHASES = (
    ("baseline_before_inputs", "region acquired; corpus bytes remain outside the lifecycle; no input clones or sink"),
    ("prepared_inputs_and_sink", "owned source and destination input Vec clones, bounded sink, and reserved sink capacity"),
    ("opened_documents", "owned source and destination Packages plus their opened presentation handles"),
    ("planned", "opened source/destination handles and the validated cross-copy plan"),
    ("published", "opened packages, plan, publication result, and sink containing the exact output"),
    ("drop_result", "opened packages, plan, and exact output sink after the publication result is released"),
    ("drop_plan", "opened packages and exact output sink after the plan is released"),
    ("drop_document_handles", "exact output sink remains after source and destination Packages and opened handles are released"),
    ("drop_sink", "no lifecycle-owned input, document, plan, publication result, or sink handles; process baseline may differ"),
)
SOURCE_PHASES = (
    ("baseline_before_inputs", "region acquired; corpus bytes remain outside the lifecycle; no source Arcs or sink"),
    ("prepared_inputs_and_sink", "source-backed caller InstrumentedSource Arcs, temporary ReadAt Arcs, and reserved bounded sink"),
    ("opened_documents", "source-backed presentation view and editor, caller source Arcs, and sink"),
    ("planned", "source-backed view, editor, source-retaining plan, caller source Arcs, and sink"),
    ("published", "source-backed view, plan, publication result, caller source Arcs, and exact output sink; editor is consumed by publication"),
    ("drop_result", "source-backed view, plan, caller source Arcs, and exact output sink"),
    ("drop_plan", "source-backed view, caller source Arcs, and exact output sink"),
    ("drop_document_handles", "source InstrumentedSource Arcs and exact output sink; source-backed view is released"),
    ("drop_caller_source_arcs", "exact output sink only; caller InstrumentedSource Arcs are released"),
    ("drop_sink", "no lifecycle-owned source, document, plan, publication result, or sink handles; process baseline may differ"),
)


class VerificationError(ValueError):
    """The JSON document violates the retention evidence contract."""


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


def text(value: Any, path: str) -> str:
    if not isinstance(value, str) or not value:
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


def digest(value: Any, path: str) -> str:
    value = text(value, path)
    if not HEX64.fullmatch(value):
        fail(path, "expected a SHA-256 hexadecimal digest")
    return value.lower()


def exact_keys(value: dict[str, Any], required: tuple[str, ...], optional: tuple[str, ...], path: str) -> None:
    allowed = set(required) | set(optional)
    missing = [name for name in required if name not in value]
    if missing:
        fail(path, f"missing fields: {', '.join(missing)}")
    unknown = sorted(set(value) - allowed)
    if unknown:
        fail(path, f"unknown fields: {', '.join(unknown)}")


def reject_unsupported_claims(value: Any, path: str) -> None:
    forbidden = {
        "rss",
        "rss_bytes",
        "peak_rss",
        "process_rss",
        "process_rss_bytes",
        "object_owned_bytes",
        "owned_bytes",
        "managed_budget",
        "managed_budget_bytes",
        "cache_eviction",
        "cache_evictions",
        "leak",
        "leak_bytes",
    }
    if isinstance(value, dict):
        for key, child in value.items():
            if key.lower() in forbidden:
                fail(f"{path}.{key}", "unsupported RSS/object/cache/leak claim")
            reject_unsupported_claims(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            reject_unsupported_claims(child, f"{path}[{index}]")


def check_hash_fields(value: Any, path: str) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            if key.lower().endswith("_sha256") or key.lower() in {"sha256", "hash"}:
                digest(child, f"{path}.{key}")
            check_hash_fields(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            check_hash_fields(child, f"{path}[{index}]")


def reject_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON object key: {key}")
        result[key] = value
    return result


def reject_nonfinite_constant(value: str) -> Any:
    raise VerificationError(f"non-finite JSON constant is not permitted: {value}")


def check_checkpoint(value: Any, path: str) -> dict[str, Any]:
    checkpoint = obj(value, path)
    exact_keys(checkpoint, CHECKPOINT_FIELDS, (), path)
    result: dict[str, Any] = {}
    for name in CHECKPOINT_FIELDS:
        if name in ("overflowed", "observer_invalid"):
            result[name] = boolean(checkpoint[name], f"{path}.{name}")
            if result[name]:
                fail(f"{path}.{name}", "allocator checkpoint is invalid")
        else:
            result[name] = uint(checkpoint[name], f"{path}.{name}")
    return result


def check_arc_counts(value: Any, path: str) -> None:
    counts = obj(value, path)
    exact_keys(counts, ("source", "destination"), (), path)
    source = uint(counts["source"], f"{path}.source")
    destination = uint(counts["destination"], f"{path}.destination")
    if source != 1 or destination != 1:
        fail(path, "after-document-drop caller Arc strong counts must be exactly one")


def check_corpus_manifest(value: Any, report: dict[str, Any]) -> None:
    manifest = obj(value, "report.corpus_manifest")
    required = (
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
    )
    exact_keys(manifest, required, ("rtf_variant",), "report.corpus_manifest")
    for name in ("name", "generator", "package_format", "shape", "payload_kind", "compression", "target_entry"):
        text(manifest[name], f"report.corpus_manifest.{name}")
    for name in ("entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        uint(manifest[name], f"report.corpus_manifest.{name}")
    digest(manifest["archive_sha256"], "report.corpus_manifest.archive_sha256")
    digest(manifest["target_payload_sha256"], "report.corpus_manifest.target_payload_sha256")
    if manifest["archive_bytes"] != report["destination_archive_bytes"]:
        fail("report.corpus_manifest.archive_bytes", "does not match destination archive bytes")
    if manifest["archive_sha256"].lower() != report["destination_archive_sha256"].lower():
        fail("report.corpus_manifest.archive_sha256", "does not match destination archive hash")
    if manifest["xlsx"] is not None:
        xlsx = obj(manifest["xlsx"], "report.corpus_manifest.xlsx")
        exact_keys(
            xlsx,
            ("sheet_count", "rows_per_sheet", "columns_per_sheet", "one_percent_update_count", "source_members"),
            (),
            "report.corpus_manifest.xlsx",
        )
        for name in ("sheet_count", "rows_per_sheet", "columns_per_sheet", "one_percent_update_count"):
            uint(xlsx[name], f"report.corpus_manifest.xlsx.{name}")
        members = obj(xlsx["source_members"], "report.corpus_manifest.xlsx.source_members")
        exact_keys(members, ("workbook", "worksheets", "shared_strings", "styles"), (), "report.corpus_manifest.xlsx.source_members")
        text(members["workbook"], "report.corpus_manifest.xlsx.source_members.workbook")
        array(members["worksheets"], "report.corpus_manifest.xlsx.source_members.worksheets")
        for item in members["worksheets"]:
            text(item, "report.corpus_manifest.xlsx.source_members.worksheets[]")
        for name in ("shared_strings", "styles"):
            if members[name] is not None:
                text(members[name], f"report.corpus_manifest.xlsx.source_members.{name}")


def check_transitions(checkpoints: list[dict[str, Any]], path: str) -> None:
    monotone = (
        "allocation_calls",
        "deallocation_calls",
        "reallocation_calls",
        "failed_allocation_calls",
        "allocated_bytes",
        "deallocated_bytes",
        "peak_live_bytes",
    )
    for index, current in enumerate(checkpoints):
        if current["allocated_bytes"] < current["deallocated_bytes"]:
            fail(f"{path}[{index}]", "absolute deallocation bytes exceed allocation bytes")
        if current["live_bytes"] != current["allocated_bytes"] - current["deallocated_bytes"]:
            fail(f"{path}[{index}].live_bytes", "absolute allocated-minus-deallocated balance is impossible")
        if current["peak_live_bytes"] < current["live_bytes"]:
            fail(f"{path}[{index}].peak_live_bytes", "lifetime peak is below endpoint live bytes")
    for index in range(1, len(checkpoints)):
        previous = checkpoints[index - 1]
        current = checkpoints[index]
        for name in monotone:
            if current[name] < previous[name]:
                fail(f"{path}[{index}].{name}", "absolute counter moved backwards")
        expected_live = previous["live_bytes"] + (
            current["allocated_bytes"] - previous["allocated_bytes"]
        ) - (current["deallocated_bytes"] - previous["deallocated_bytes"])
        if current["live_bytes"] != expected_live:
            fail(f"{path}[{index}].live_bytes", "counter transition live balance is impossible")


def check_region(value: Any, path: str, checkpoints: list[dict[str, Any]]) -> None:
    region = obj(value, path)
    exact_keys(region, REGION_FIELDS, (), path)
    if text(region["status"], f"{path}.status") != "measured":
        fail(f"{path}.status", "region must be measured")
    if text(region["scope"], f"{path}.scope") != COUNTER_SCOPE:
        fail(f"{path}.scope", f"expected {COUNTER_SCOPE!r}")
    numeric = {name: uint(region[name], f"{path}.{name}") for name in REGION_FIELDS[2:]}
    if numeric["live_bytes_before"] != checkpoints[0]["live_bytes"]:
        fail(f"{path}.live_bytes_before", "does not match baseline checkpoint")
    if numeric["live_bytes_after"] != checkpoints[-1]["live_bytes"]:
        fail(f"{path}.live_bytes_after", "does not match sink-dropped checkpoint")
    if numeric["peak_live_bytes_before"] != checkpoints[0]["peak_live_bytes"]:
        fail(f"{path}.peak_live_bytes_before", "does not match baseline high-water")
    if numeric["peak_live_bytes_after"] != checkpoints[-1]["peak_live_bytes"]:
        fail(f"{path}.peak_live_bytes_after", "does not match final high-water")
    if numeric["retention_probe_region_peak_live_bytes"] < max(item["live_bytes"] for item in checkpoints):
        fail(f"{path}.retention_probe_region_peak_live_bytes", "region peak is below a checkpoint endpoint")
    if numeric["retention_probe_region_peak_live_bytes"] > checkpoints[-1]["peak_live_bytes"]:
        fail(f"{path}.retention_probe_region_peak_live_bytes", "region peak exceeds lifetime peak")
    for name in (
        "allocation_calls",
        "deallocation_calls",
        "reallocation_calls",
        "failed_allocation_calls",
        "allocated_bytes",
        "deallocated_bytes",
    ):
        expected = checkpoints[-1][name] - checkpoints[0][name]
        if numeric[name] != expected:
            fail(f"{path}.{name}", "region counter delta does not match journal endpoints")


def check_output_and_gates(report: dict[str, Any]) -> None:
    digest(report["expected_output_sha256"], "report.expected_output_sha256")
    uint(report["expected_output_bytes"], "report.expected_output_bytes")
    if not boolean(report["corpus_gates_verified"], "report.corpus_gates_verified"):
        fail("report.corpus_gates_verified", "corpus correctness gates are false")
    if not boolean(report["all_iteration_output_bytes_verified"], "report.all_iteration_output_bytes_verified"):
        fail("report.all_iteration_output_bytes_verified", "exact output equality gate is false")


def check_report(report: Any) -> dict[str, Any]:
    report = obj(report, "report")
    reject_unsupported_claims(report, "report")
    check_hash_fields(report, "report")
    required = (
        "schema",
        "api",
        "corpus",
        "corpus_manifest",
        "samples",
        "warmup",
        "source_revision",
        "source_archive_sha256",
        "source_archive_bytes",
        "destination_archive_sha256",
        "destination_archive_bytes",
        "expected_output_sha256",
        "expected_output_bytes",
        "corpus_gates_verified",
        "all_iteration_output_bytes_verified",
        "checked_iteration_count",
        "binary_sha256",
        "current_exe",
        "binary_bytes",
        "allocator",
        "instrumentation",
        "allocator_counter_revision",
        "callback_scope",
        "ownership_scope",
        "phases",
        "samples_raw",
    )
    exact_keys(report, required, (), "report")
    if text(report["schema"], "report.schema") != SCHEMA:
        fail("report.schema", f"expected {SCHEMA!r}")
    api = text(report["api"], "report.api")
    if api not in {"owned", "source-backed"}:
        fail("report.api", "expected 'owned' or 'source-backed'")
    corpus = text(report["corpus"], "report.corpus")
    if corpus not in {"plain", "media-rich"}:
        fail("report.corpus", "expected 'plain' or 'media-rich'")
    samples = uint(report["samples"], "report.samples")
    if not 1 <= samples <= MAX_SAMPLES:
        fail("report.samples", f"expected a bounded count from 1 to {MAX_SAMPLES}")
    warmup = uint(report["warmup"], "report.warmup")
    if warmup > MAX_WARMUP:
        fail("report.warmup", f"expected a bounded count from 0 to {MAX_WARMUP}")
    if uint(report["checked_iteration_count"], "report.checked_iteration_count") != samples + warmup:
        fail("report.checked_iteration_count", "must equal warmup plus retained samples")
    revision = text(report["source_revision"], "report.source_revision")
    if not HEX40.fullmatch(revision):
        fail("report.source_revision", "expected a 40-character hexadecimal revision")
    for name in ("source_archive_sha256", "destination_archive_sha256", "binary_sha256"):
        digest(report[name], f"report.{name}")
    for name in ("source_archive_bytes", "destination_archive_bytes", "expected_output_bytes", "binary_bytes"):
        uint(report[name], f"report.{name}")
    text(report["current_exe"], "report.current_exe")
    if text(report["allocator"], "report.allocator") != ALLOCATOR_IDENTITY:
        fail("report.allocator", f"expected {ALLOCATOR_IDENTITY!r}")
    if text(report["instrumentation"], "report.instrumentation") != INSTRUMENTATION_IDENTITY:
        fail("report.instrumentation", f"expected {INSTRUMENTATION_IDENTITY!r}")
    check_corpus_manifest(report["corpus_manifest"], report)
    if text(report["allocator_counter_revision"], "report.allocator_counter_revision") != COUNTER_REVISION:
        fail("report.allocator_counter_revision", f"expected {COUNTER_REVISION!r}")
    if text(report["callback_scope"], "report.callback_scope") != CALLBACK_SCOPE:
        fail("report.callback_scope", "allocator callback scope changed")
    if text(report["ownership_scope"], "report.ownership_scope") != OWNERSHIP_SCOPE:
        fail("report.ownership_scope", "ownership scope changed")
    check_output_and_gates(report)

    expected_phases = SOURCE_PHASES if api == "source-backed" else OWNED_PHASES
    phases = array(report["phases"], "report.phases")
    if len(phases) != len(expected_phases):
        fail("report.phases", f"expected {len(expected_phases)} API-specific phase descriptions")
    for index, (raw, expected) in enumerate(zip(phases, expected_phases)):
        phase = obj(raw, f"report.phases[{index}]")
        exact_keys(phase, ("label", "live_owners"), (), f"report.phases[{index}]")
        if text(phase["label"], f"report.phases[{index}].label") != expected[0]:
            fail(f"report.phases[{index}].label", f"expected {expected[0]!r}")
        if text(phase["live_owners"], f"report.phases[{index}].live_owners") != expected[1]:
            fail(f"report.phases[{index}].live_owners", "API owner description changed")

    rows = array(report["samples_raw"], "report.samples_raw")
    if len(rows) != samples:
        fail("report.samples_raw", "raw row count differs from reported samples")
    checkpoint_names = [
        "baseline_before_inputs",
        "prepared_inputs_and_sink",
        "opened_documents",
        "planned",
        "published",
        "drop_result",
        "drop_plan",
        "drop_document_handles",
    ]
    if api == "source-backed":
        checkpoint_names.append("drop_caller_source_arcs")
    checkpoint_names.append("drop_sink")
    required_row = ("sample_index", *checkpoint_names, "retention_probe")
    optional_row = ("source_arc_counts_after_document_drop",) if api == "source-backed" else ()
    for index, raw in enumerate(rows):
        row = obj(raw, f"report.samples_raw[{index}]")
        exact_keys(row, required_row, optional_row, f"report.samples_raw[{index}]")
        if uint(row["sample_index"], f"report.samples_raw[{index}].sample_index") != index:
            fail(f"report.samples_raw[{index}].sample_index", "raw sample order is not preserved")
        checkpoints = [check_checkpoint(row[name], f"report.samples_raw[{index}].{name}") for name in checkpoint_names]
        check_transitions(checkpoints, f"report.samples_raw[{index}].checkpoints")
        if api == "source-backed":
            if "source_arc_counts_after_document_drop" not in row:
                fail(f"report.samples_raw[{index}]", "source-backed row omitted source Arc counts")
            check_arc_counts(row["source_arc_counts_after_document_drop"], f"report.samples_raw[{index}].source_arc_counts_after_document_drop")
        check_region(row["retention_probe"], f"report.samples_raw[{index}].retention_probe", checkpoints)
    return {"status": "valid", "schema": SCHEMA, "api": api, "corpus": corpus, "samples": samples}


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
    parser.add_argument("path", nargs="?", type=Path, help="retention report JSON")
    parser.add_argument("--report", dest="report_option", type=Path, help="retention report JSON")
    parser.add_argument("--protocol", type=Path, default=None, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    path = args.report_option or args.path
    if path is None:
        parser.error("a report path is required")
    try:
        result = check_report(load_json(path))
    except VerificationError as exc:
        print(f"INVALID: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
