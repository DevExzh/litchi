#!/usr/bin/env python3
"""Independently audit a completed 0968 diagnostic CFB census with per-edit stage counts.

This auditor is deliberately separate from the 0968 schema checker and reducer.
It performs no build, workload, timer, allocator, or performance calculation. It
only reads a bound plan, GO receipt, capture manifest, and child packet custody,
then writes one explicitly requested fresh audit JSON. Existing output paths are
never overwritten.
"""
from __future__ import annotations

import argparse
from collections import defaultdict
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import sys
from typing import Any


PLAN_SCHEMA = "litchi.performance.0968.cfb_commit_census_plan.v3"
REPORT_SCHEMA = "litchi.performance.0968.cfb_commit_census.v2"
GO_SCHEMA = "litchi.performance.0968.cfb_commit_census_go.v3"
CAPTURE_SCHEMA = "litchi.performance.0968.cfb_commit_census_capture.v3"
RUN_START_SCHEMA = "litchi.performance.0968.cfb_census_run_start.v3"
RUN_TERMINAL_SCHEMA = "litchi.performance.0968.cfb_census_run_terminal.v3"
CHILD_START_SCHEMA = "litchi.performance.0968.cfb_child_start.v3"
CHILD_TERMINAL_SCHEMA = "litchi.performance.0968.cfb_child_terminal.v3"
CHILD_PARSE_SCHEMA = "litchi.performance.0968.cfb_child_parse.v3"

# The corrected source seal is expected to carry one unique top-level schema id.
# v3 is intentionally absent: its retained JSON has a duplicate top-level
# ``schema`` key and must remain historical/superseded.
SEAL_SCHEMAS = {
    "litchi.performance.0968.cfb_commit_census.seal.v4",
}
PROVENANCE_SCHEMAS = {
    "litchi.performance.0968.cfb_commit_census.provenance.v3",
    "litchi.performance.0968.cfb_commit_census.provenance.v4",
}
SOURCE_HARNESS_SCHEMAS = {
    "litchi.performance.0968.cfb_commit_census_source_harness.v2",
}
SCHEMA_RECEIPT_ID = "litchi.performance.0968.cfb_commit_census.v2"
INCIDENT_SCHEMA = "litchi.performance.0968.cfb_commit_census_incident.v1"
ORACLE_SCHEMAS = {
    "litchi.performance.0968.cfb_commit_ordinary_oracle.v1",
    "litchi.performance.0968.cfb_commit_ordinary_oracle.v2",
}
EQUIVALENCE_SCHEMA = "litchi.performance.0968.ordinary_oracle_equivalence.v1"

CASES = ("docnohf", "docfloat")
CASE_FILES = {
    "docnohf": "test-data/ole/doc/NoHeadFoot.doc",
    "docfloat": "test-data/ole/doc/FloatingPictures.doc",
}
EDIT_COUNTS = (1, 2, 8, 32)
RETENTIONS = ("zero", "default", "release")
RETENTION_ARG = {"zero": "zero", "default": "default", "release": "release"}
REPORT_RETENTION = {
    "zero": "zero",
    "default": "default_8mib",
    "release": "release_8mib",
}
RETENTION_CEILING = {"zero": 0, "default": 8 * 1024 * 1024, "release": 8 * 1024 * 1024}
ALTERNATING_TEXTS = (
    "litchi copy-through baseline replacement text",
    "litchi retained render second replacement text",
)
GATE_NAMES = (
    "all_stages_change_text",
    "exact_cross_route_outputs",
    "final_paragraph_semantic_readback",
    "final_public_reader_reopen",
    "exact_noop_source",
    "noop_shares_source_allocation",
    "forward_patch_exact_output",
    "inverse_patch_exact_source",
    "foreign_patch_refused",
    "stale_patch_refused",
    "same_value_preserves_retention",
    "operation_limit_refusal_atomic",
    "retained_capacities_within_ceiling",
    "zero_route_clears_retention",
    "release_route_clears_every_stage",
    "untouched_streams_preserved",
    "directory_metadata_preserved",
    "semantic_oracles_passed",
    "semantic_change_count_matches_edit_count",
    "allocator_retained_matches_output_capacity",
)
AUDIT_SCHEMA = "litchi.performance.0968.independent_raw_census_audit.v2"

CENSUS_FIELDS = (
    "commit_attempts",
    "commit_successes",
    "commit_errors",
    "overlay_attempted",
    "overlay_success",
    "overlay_declined",
    "overlay_error",
    "overlay_not_attempted",
    "layout_attempted",
    "layout_reuse_success",
    "layout_non_reuse",
    "layout_error",
    "capture_attempted",
    "capture_success",
    "capture_error",
    "capture_streams",
    "capture_payload_bytes",
    "saturated",
)
REPORT_FIELDS = {
    "schema_version",
    "schema",
    "change",
    "case",
    "input",
    "source_sha256",
    "source_bytes",
    "selected_paragraph",
    "edits",
    "alternating_texts",
    "retention",
    "retention_ceiling_bytes",
    "timing_claim",
    "region_definition",
    "warmups",
    "samples_requested",
    "output_sha256",
    "output_bytes",
    "output_capacity",
    "reference_output_sha256",
    "reference_output_bytes",
    "exact_output_equal_reference",
    "route_outputs",
    "correctness",
    "samples",
}
ROUTE_FIELDS = {
    "retention",
    "retention_ceiling_bytes",
    "output_sha256",
    "output_bytes",
    "output_capacity",
    "semantic_change_count",
    "final_retained_before_commit",
    "stages",
    "oracle",
}
SAMPLE_FIELDS = {
    "index",
    "output_sha256",
    "output_bytes",
    "output_capacity",
    "semantic_change_count",
    "exact_output_equal_reference",
    "final_paragraph_semantic_readback",
    "final_public_reader_reopen",
    "retained_capacities_within_ceiling",
    "allocator_retained_matches_output_capacity",
    "oracle",
    "stages",
    "census",
}
ORACLE_FIELDS = {
    "oracle_ok",
    "source_structurally_valid",
    "expected_structurally_valid",
    "output_structurally_valid",
    "semantic_reopen_ok",
    "semantic_directory_metadata_matches_expected",
    "unchanged_stream_bytes_match_expected",
    "unchanged_source_output_stream_bytes_match",
    "output_stream_paths_match_expected",
    "output_stream_bytes_match_expected",
    "root_clsid_preserved",
    "storage_clsids_preserved",
    "failure_reasons",
}
STAGE_FIELDS = {"stage", "requested_text", "before", "after", "after_release"}
HEX_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
MAX_JSON_BYTES = 1 << 20
MAX_STDOUT_BYTES = 1 << 20
MAX_STDERR_BYTES = 1 << 20


class AuditError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def reject_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            raise AuditError(f"duplicate JSON key: {key}")
        value[key] = item
    return value


def path_regular(value: str | Path, label: str) -> Path:
    raw = Path(value)
    require(not raw.is_symlink(), f"{label}: symlink refused: {raw}")
    path = raw.resolve()
    require(path.is_file() and not path.is_symlink(), f"{label}: regular file required: {path}")
    return path


def path_directory(value: str | Path, label: str) -> Path:
    raw = Path(value)
    require(not raw.is_symlink(), f"{label}: symlink directory refused: {raw}")
    path = raw.resolve()
    require(path.is_dir() and not path.is_symlink(), f"{label}: directory required: {path}")
    return path


def within(path: Path, root: Path, label: str) -> Path:
    path = path.resolve()
    root = root.resolve()
    try:
        path.relative_to(root)
    except ValueError as exc:
        raise AuditError(f"{label}: path escapes {root}: {path}") from exc
    return path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path, label: str) -> dict[str, Any]:
    path = path_regular(path, label)
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_file(path)}


def check_hash(value: Any, label: str) -> None:
    require(isinstance(value, str) and HEX_RE.fullmatch(value) is not None, f"{label}: invalid SHA-256")


def check_revision(value: Any, label: str) -> None:
    require(isinstance(value, str) and REVISION_RE.fullmatch(value) is not None, f"{label}: invalid revision")


def check_nonnegative(value: Any, label: str) -> None:
    require(type(value) is int and value >= 0, f"{label}: nonnegative integer required")


def check_bool(value: Any, label: str) -> None:
    require(type(value) is bool, f"{label}: boolean required")


def load_json(path: Path, label: str, maximum: int = MAX_JSON_BYTES) -> Any:
    path = path_regular(path, label)
    require(path.stat().st_size < maximum, f"{label}: JSON exceeds bound")
    try:
        text = path.read_text(encoding="utf-8")
        return json.loads(text, object_pairs_hook=reject_duplicate_pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise AuditError(f"{label}: invalid JSON: {exc}") from exc


def verify_descriptor(value: Any, label: str, expected: Path | None = None, parent: Path | None = None) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor object required")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}: descriptor path missing")
    path = path_regular(path_value, label)
    if expected is not None:
        require(path == expected.resolve(), f"{label}: path binding changed")
    if parent is not None:
        within(path, parent, label)
    check_nonnegative(value.get("bytes"), f"{label}.bytes")
    check_hash(value.get("sha256"), f"{label}.sha256")
    require(value["bytes"] == path.stat().st_size, f"{label}: byte count changed")
    require(value["sha256"] == sha256_file(path), f"{label}: SHA-256 changed")
    return path


def verify_reference(value: Any, label: str, expected: Path | None = None, parent: Path | None = None) -> Path:
    """Verify a path/SHA reference whose byte count is intentionally omitted."""
    require(isinstance(value, dict), f"{label}: reference object required")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}: reference path missing")
    path = path_regular(path_value, label)
    if expected is not None:
        require(path == expected.resolve(), f"{label}: path binding changed")
    if parent is not None:
        within(path, parent, label)
    check_hash(value.get("sha256"), f"{label}.sha256")
    require(value["sha256"] == sha256_file(path), f"{label}: SHA-256 changed")
    if "bytes" in value:
        check_nonnegative(value["bytes"], f"{label}.bytes")
        require(value["bytes"] == path.stat().st_size, f"{label}: byte count changed")
    return path


def expected_cells() -> list[dict[str, Any]]:
    return [
        {
            "id": f"{case}/e{edits}/{retention}",
            "case": case,
            "edits": edits,
            "retention": retention,
            "retention_arg": RETENTION_ARG[retention],
        }
        for case in CASES
        for edits in EDIT_COUNTS
        for retention in RETENTIONS
    ]


def expected_rows() -> list[dict[str, Any]]:
    cells = expected_cells()
    rows: list[dict[str, Any]] = []
    position = 0
    for repeat, direction, ordered in ((1, "forward", cells), (2, "reverse", list(reversed(cells)))):
        for repeat_position, cell in enumerate(ordered, start=1):
            position += 1
            rows.append(
                {
                    **cell,
                    "child_id": f"cfb-r{repeat:02d}-{cell['case']}-e{cell['edits']:02d}-{cell['retention']}",
                    "position": position,
                    "repeat": repeat,
                    "direction": direction,
                    "repeat_position": repeat_position,
                    "warmups": 0,
                    "samples": 1,
                    "retry": False,
                }
            )
    return rows


def source_scope_digest(root: Path, files: list[Any], label: str) -> tuple[int, str, list[dict[str, Any]]]:
    root = path_directory(root, f"{label}.root")
    entries: list[dict[str, Any]] = []
    seen: set[str] = set()
    for index, item in enumerate(files):
        path = verify_descriptor(item, f"{label}[{index}]")
        relative = path.relative_to(root).as_posix() if path.is_relative_to(root) else None
        require(relative is not None, f"{label}[{index}]: path escapes source root")
        require(relative not in seen, f"{label}: duplicate path {relative}")
        seen.add(relative)
        entries.append({
            "path": str(path),
            "relative": relative,
            "bytes": path.stat().st_size,
            "sha256": sha256_file(path),
        })
    entries.sort(key=lambda item: item["relative"])
    digest = hashlib.sha256()
    for item in entries:
        digest.update(item["relative"].encode("utf-8"))
        digest.update(b"\0")
        digest.update(item["bytes"].to_bytes(8, "big"))
        digest.update(item["sha256"].encode("ascii"))
        digest.update(b"\n")
    return len(entries), digest.hexdigest(), entries


def check_provenance(value: dict[str, Any], revision: str, label: str) -> dict[str, Any]:
    schema_id = value.get("schema")
    require(schema_id in PROVENANCE_SCHEMAS, f"{label}: unsupported provenance schema {schema_id!r}")
    state = value.get("state")
    require(isinstance(state, str) and state.startswith("diagnostic_source_harness_sealed_pending_"), f"{label}: provenance state")
    require(value.get("current_head") == revision, f"{label}: source revision")
    root = path_directory(value.get("source_diagnostic_path", ""), f"{label}.source")
    files = value.get("diagnostic_files")
    require(isinstance(files, list) and files, f"{label}: diagnostic file scope")
    count, scope_hash, normalized = source_scope_digest(root, files, f"{label}.diagnostic_files")
    verify_descriptor(value.get("binary"), f"{label}.binary")
    fixtures = value.get("fixtures")
    require(isinstance(fixtures, list) and len(fixtures) == len(CASES), f"{label}: fixture descriptors")
    for index, item in enumerate(fixtures):
        verify_descriptor(item, f"{label}.fixtures[{index}]")
    verify_descriptor(value.get("code_patch"), f"{label}.code_patch")
    schema_receipt = value.get("schema_v2")
    if schema_receipt is not None:
        schema_path = verify_descriptor(schema_receipt, f"{label}.schema_v2")
        schema_value = load_json(schema_path, f"{label}.schema_v2.json")
        require(schema_value.get("schema") == SCHEMA_RECEIPT_ID and schema_value.get("schema_version") == 2, f"{label}: schema receipt")
    return {
        "root": str(root),
        "file_count": count,
        "scope_sha256": scope_hash,
        "scope_files": normalized,
        "source_revision": revision,
        "binary": value["binary"],
        "fixtures": fixtures,
    }


def check_source_harness(value: Any, revision: str, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: object required")
    schema_id = value.get("schema")
    if schema_id in SEAL_SCHEMAS:
        require(value.get("status") == "SEALED_DIAGNOSTIC_SOURCE_HARNESS", f"{label}: seal status")
        require(value.get("source_baseline_head") == revision, f"{label}: seal source revision")
        patch = value.get("code_patch")
        verify_descriptor(patch, f"{label}.code_patch")
        schema_receipt = value.get("schema_receipt")
        require(isinstance(schema_receipt, dict), f"{label}: unique schema_receipt required")
        schema_path = verify_reference(schema_receipt, f"{label}.schema_receipt")
        schema_value = load_json(schema_path, f"{label}.schema_receipt.json")
        require(schema_value.get("schema") == SCHEMA_RECEIPT_ID and schema_value.get("schema_version") == 2, f"{label}: schema receipt mismatch")
        require(value.get("scope", {}).get("timing_claim") is False, f"{label}: timing claim")
        require(value.get("scope", {}).get("allocator_instrumentation") is False, f"{label}: allocator instrumentation")
        require(value.get("scope", {}).get("metadata_projection") is False, f"{label}: metadata projection")
        require(value.get("scope", {}).get("formal_children") == 48, f"{label}: formal child count")
        seal_binary = verify_reference(value.get("binary"), f"{label}.binary")
        provenance_descriptor = value.get("provenance")
        require(isinstance(provenance_descriptor, dict), f"{label}: provenance descriptor")
        provenance_path = verify_reference(provenance_descriptor, f"{label}.provenance")
        provenance = load_json(provenance_path, f"{label}.provenance.json")
        binding = check_provenance(provenance, revision, f"{label}.provenance")
        require(value["binary"].get("path") == binding["binary"].get("path") and value["binary"].get("sha256") == binding["binary"].get("sha256"), f"{label}: binary differs from provenance")
        require(seal_binary == path_regular(value["binary"]["path"], f"{label}.binary"), f"{label}: binary path")
        return binding
    if schema_id in PROVENANCE_SCHEMAS:
        return check_provenance(value, revision, label)
    require(schema_id in SOURCE_HARNESS_SCHEMAS, f"{label}: unsupported source-harness schema {schema_id!r}")
    require(value.get("schema_version") == 2 and value.get("status") == "sealed", f"{label}: source-harness status/version")
    require(value.get("source_revision") == revision, f"{label}: source revision")
    source = value.get("source")
    require(isinstance(source, dict), f"{label}: source binding")
    files = source.get("scope_files") or value.get("harness_files")
    require(isinstance(files, list) and files, f"{label}: source scope")
    count, scope_hash, normalized = source_scope_digest(Path(source.get("root", "")), files, f"{label}.scope")
    require(count == source.get("file_count") and scope_hash == source.get("scope_sha256"), f"{label}: source scope digest")
    require(value.get("report_schema") == REPORT_SCHEMA and value.get("timing_claim") is False, f"{label}: report/timing scope")
    verify_descriptor(value.get("binary"), f"{label}.binary")
    return {"root": str(path_directory(source["root"], f"{label}.root")), "file_count": count, "scope_sha256": scope_hash, "scope_files": normalized, "source_revision": revision, "binary": value["binary"]}


def check_incident(value: Any, label: str) -> None:
    """Independently bind the preserved rejected v2 epoch to the v3 plan."""
    require(isinstance(value, dict) and value.get("schema") == INCIDENT_SCHEMA, f"{label}: schema")
    require(value.get("status") == "protocol_bug_confirmed_pending_v3_review_and_go" and value.get("change") == "0968", f"{label}: status/change")
    finding = value.get("finding")
    require(isinstance(finding, dict), f"{label}: finding")
    require(finding.get("source_or_binary_changed") is False and finding.get("workload_retry_performed") is False, f"{label}: source/retry disposition")
    require(finding.get("incorrect_v2_requirements") == ["commit_attempts == 1", "capture_attempted == 1"], f"{label}: incorrect v2 requirements")
    require(finding.get("correct_contract") == ["commit_attempts == row.edits", "commit_successes == row.edits", "capture_attempted == row.edits", "capture_success == row.edits"], f"{label}: corrected contract")
    retained = value.get("retained_v2")
    require(isinstance(retained, dict), f"{label}: retained v2")
    for key in ("plan", "capture_manifest", "run_terminal", "schema_tool", "runner_tool", "binary"):
        verify_descriptor(retained.get(key), f"{label}.{key}")
    require(retained.get("capture_status") == "complete_with_failures" and retained.get("packet_count") == 48 and retained.get("retry_count") == 0 and retained.get("raw_stdout_reports_audited") == 48, f"{label}: retained capture summary")
    require(retained.get("returncodes") == {"0": 48} and retained.get("children_by_edits") == {"1": 12, "2": 12, "8": 12, "32": 12}, f"{label}: retained child summary")
    rows = retained.get("raw_rows")
    require(isinstance(rows, list) and len(rows) == 48, f"{label}: retained raw rows")
    expected_by_id = {row["child_id"]: row for row in expected_rows()}
    seen: set[str] = set()
    for row in rows:
        require(isinstance(row, dict), f"{label}: retained row object")
        child_id = row.get("child_id")
        require(isinstance(child_id, str) and child_id in expected_by_id and child_id not in seen, f"{label}: retained row identity")
        expected = expected_by_id[child_id]
        require(type(row.get("edits")) is int and row["edits"] == expected["edits"] and row.get("returncode") == 0, f"{label}: retained row outcome")
        require(row.get("commit_attempts") == row["edits"] and row.get("capture_attempted") == row["edits"], f"{label}: retained row cardinality")
        seen.add(child_id)
    require(seen == set(expected_by_id), f"{label}: retained row set")
    offline = value.get("offline_audit")
    require(isinstance(offline, dict) and offline.get("result") == "all_48_retained_stdout_packets_pass_per_edit_cardinality_and_existing_correctness_oracles" and offline.get("matrix_edits") == [1, 2, 8, 32] and offline.get("retentions") == ["zero", "default", "release"] and offline.get("repetitions") == 2, f"{label}: offline audit")


def check_oracle(value: Any, label: str) -> None:
    require(isinstance(value, dict) and set(value) == ORACLE_FIELDS, f"{label}: oracle fields")
    for field in ORACLE_FIELDS - {"failure_reasons"}:
        require(value[field] is True, f"{label}.{field}: false")
    require(value["failure_reasons"] == [], f"{label}: failure reasons")


def check_stage_vector(value: Any, label: str) -> None:
    require(isinstance(value, list) and len(value) == 32, f"{label}: 32 stages required")
    for index, stage in enumerate(value):
        require(isinstance(stage, dict) and set(stage) == STAGE_FIELDS, f"{label}[{index}]: stage fields")
        check_nonnegative(stage["stage"], f"{label}[{index}].stage")
        require(stage["stage"] <= 32, f"{label}[{index}].stage bound")
        require(isinstance(stage["requested_text"], str), f"{label}[{index}].requested_text")
        for field in ("before", "after", "after_release"):
            require(stage[field] is None or (type(stage[field]) is int and stage[field] >= 0), f"{label}[{index}].{field}")


def check_census(value: Any, label: str, expected_edits: int) -> None:
    require(isinstance(value, dict) and set(value) == set(CENSUS_FIELDS), f"{label}: census fields")
    require(type(expected_edits) is int and expected_edits > 0, f"{label}: invalid expected edit count")
    for field in CENSUS_FIELDS[:-1]:
        check_nonnegative(value[field], f"{label}.{field}")
    require(value["saturated"] is False, f"{label}: saturated")
    # The observer is inside the per-edit commit seam. Every requested edit
    # therefore contributes one attempt and one successful stage completion;
    # the old v1 auditor incorrectly required a corpus-wide count of one.
    require(value["commit_attempts"] == expected_edits and value["commit_successes"] == expected_edits and value["commit_errors"] == 0, f"{label}: per-edit commit outcome")
    require(value["commit_attempts"] == value["commit_successes"] + value["commit_errors"], f"{label}: commit conservation")
    require(value["overlay_attempted"] + value["overlay_not_attempted"] == value["commit_attempts"], f"{label}: overlay admission conservation")
    require(value["overlay_attempted"] == value["overlay_success"] + value["overlay_declined"] + value["overlay_error"], f"{label}: overlay outcome conservation")
    require(value["overlay_error"] == 0, f"{label}: overlay error")
    require(value["layout_attempted"] == value["layout_reuse_success"] + value["layout_non_reuse"] + value["layout_error"], f"{label}: layout conservation")
    require(value["layout_error"] == 0, f"{label}: layout error")
    require(value["capture_attempted"] == value["capture_success"] + value["capture_error"], f"{label}: capture conservation")
    require(value["capture_attempted"] == expected_edits and value["capture_success"] == expected_edits and value["capture_error"] == 0, f"{label}: per-edit capture outcome")
    require(value["capture_streams"] > 0, f"{label}: capture stream count")


def check_output(value: Any, expected: dict[str, Any], label: str) -> None:
    require(isinstance(value, dict) and set(value) == {"sha256", "bytes"}, f"{label}: output fields")
    check_hash(value["sha256"], f"{label}.sha256")
    check_nonnegative(value["bytes"], f"{label}.bytes")
    require(value == expected, f"{label}: ordinary output oracle mismatch")


def check_route(value: Any, row: dict[str, Any], oracle: dict[str, dict[str, Any]], label: str) -> None:
    require(isinstance(value, dict) and set(value) == ROUTE_FIELDS, f"{label}: route fields")
    retention = value["retention"]
    require(retention in REPORT_RETENTION.values(), f"{label}: retention")
    internal = next(key for key, shown in REPORT_RETENTION.items() if shown == retention)
    require(value["retention_ceiling_bytes"] == RETENTION_CEILING[internal], f"{label}: ceiling")
    check_hash(value["output_sha256"], f"{label}.output_sha256")
    for field in ("output_bytes", "output_capacity", "semantic_change_count"):
        check_nonnegative(value[field], f"{label}.{field}")
    require(value["semantic_change_count"] == row["edits"], f"{label}: semantic change count")
    require(value["final_retained_before_commit"] is None or (type(value["final_retained_before_commit"]) is int and value["final_retained_before_commit"] >= 0), f"{label}: retained value")
    check_stage_vector(value["stages"], f"{label}.stages")
    check_oracle(value["oracle"], f"{label}.oracle")
    expected = oracle[f"{row['case']}/e{row['edits']}/{internal}"]
    require(value["output_sha256"] == expected["sha256"] and value["output_bytes"] == expected["bytes"], f"{label}: output differs from ordinary oracle")


def check_report(report: Any, row: dict[str, Any], fixtures: dict[str, dict[str, Any]], oracle: dict[str, dict[str, Any]], label: str) -> None:
    require(isinstance(report, dict) and set(report) == REPORT_FIELDS, f"{label}: report fields")
    require(report["schema"] == REPORT_SCHEMA and report["schema_version"] == 2 and report["change"] == "0968", f"{label}: schema/change")
    require(report["case"] == row["case"] and report["edits"] == row["edits"], f"{label}: cell identity")
    require(report["retention"] == REPORT_RETENTION[row["retention"]], f"{label}: retention identity")
    require(isinstance(report["input"], str) and report["input"], f"{label}: input")
    fixture = fixtures[row["case"]]
    require(report["source_sha256"] == fixture["sha256"] and report["source_bytes"] == fixture["bytes"], f"{label}: fixture identity")
    require(report["selected_paragraph"] == 0 and tuple(report["alternating_texts"]) == ALTERNATING_TEXTS, f"{label}: selection/texts")
    require(report["retention_ceiling_bytes"] == RETENTION_CEILING[row["retention"]], f"{label}: retention ceiling")
    require(report["timing_claim"] is False and isinstance(report["region_definition"], str) and "ObjectEditor::commit_candidate_with_rendered" in report["region_definition"], f"{label}: timing/scope")
    require(report["warmups"] == 0 and report["samples_requested"] == 1, f"{label}: sample policy")
    check_hash(report["output_sha256"], f"{label}.output_sha256")
    for field in ("output_bytes", "output_capacity", "reference_output_bytes"):
        check_nonnegative(report[field], f"{label}.{field}")
    check_hash(report["reference_output_sha256"], f"{label}.reference_output_sha256")
    expected = oracle[row["id"]]
    reference = oracle[f"{row['case']}/e{row['edits']}/zero"]
    require(report["output_sha256"] == expected["sha256"] and report["output_bytes"] == expected["bytes"], f"{label}: output oracle mismatch")
    require(report["reference_output_sha256"] == reference["sha256"] and report["reference_output_bytes"] == reference["bytes"], f"{label}: reference oracle mismatch")
    require(report["exact_output_equal_reference"] is True, f"{label}: reference mismatch")
    correctness = report["correctness"]
    require(isinstance(correctness, dict) and set(correctness) == set(GATE_NAMES), f"{label}: correctness fields")
    require(all(value is True for value in correctness.values()), f"{label}: correctness gate false")
    routes = report["route_outputs"]
    require(isinstance(routes, list) and len(routes) == 3, f"{label}: route count")
    require([route.get("retention") for route in routes] == list(REPORT_RETENTION.values()), f"{label}: route order")
    for index, route in enumerate(routes):
        check_route(route, row, oracle, f"{label}.route_outputs[{index}]")
    samples = report["samples"]
    require(isinstance(samples, list) and len(samples) == 1, f"{label}: sample count")
    sample = samples[0]
    require(isinstance(sample, dict) and set(sample) == SAMPLE_FIELDS, f"{label}.samples[0]: fields")
    require(sample["index"] == 0, f"{label}.samples[0]: index")
    check_hash(sample["output_sha256"], f"{label}.sample.output_sha256")
    for field in ("output_bytes", "output_capacity", "semantic_change_count"):
        check_nonnegative(sample[field], f"{label}.sample.{field}")
    require(sample["output_sha256"] == report["output_sha256"] and sample["output_bytes"] == report["output_bytes"], f"{label}: sample output mismatch")
    require(sample["semantic_change_count"] == row["edits"], f"{label}: sample semantic count")
    for field in ("exact_output_equal_reference", "final_paragraph_semantic_readback", "final_public_reader_reopen", "retained_capacities_within_ceiling", "allocator_retained_matches_output_capacity"):
        require(sample[field] is True, f"{label}.sample.{field}: false")
    check_oracle(sample["oracle"], f"{label}.sample.oracle")
    check_stage_vector(sample["stages"], f"{label}.sample.stages")
    check_census(sample["census"], f"{label}.sample.census", row["edits"])


def verify_plan(plan_path: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, dict[str, Any]]]:
    plan = load_json(plan_path, "plan")
    require(isinstance(plan, dict) and plan.get("schema") == PLAN_SCHEMA and plan.get("change") == "0968", "plan schema/change")
    require(isinstance(plan.get("status"), str) and plan["status"], "plan status")
    revision = plan.get("source_revision")
    check_revision(revision, "plan.source_revision")
    source = plan.get("source")
    require(isinstance(source, dict), "plan source")
    source_root = path_directory(source.get("root", ""), "plan.source.root")
    scope_files = source.get("scope_files")
    require(isinstance(scope_files, list) and scope_files, "plan source scope")
    count, scope_hash, normalized = source_scope_digest(source_root, scope_files, "plan.source.scope_files")
    require(count == source.get("file_count") and scope_hash == source.get("scope_sha256"), "plan source scope digest")
    require(normalized == scope_files, "plan source scope descriptor normalization")
    source_harness_path = verify_descriptor(plan.get("source_harness"), "plan.source_harness")
    source_harness = load_json(source_harness_path, "source-harness")
    binding = check_source_harness(source_harness, revision, "source-harness")
    require(binding["root"] == str(source_root), "source root differs from source-harness")
    require(binding["file_count"] == count and binding["scope_sha256"] == scope_hash and binding["scope_files"] == normalized, "source scope differs from source-harness")
    incident_path = verify_descriptor(plan.get("incident"), "plan.incident")
    check_incident(load_json(incident_path, "plan incident"), "plan incident")
    binary_path = verify_descriptor(plan.get("binary"), "plan.binary")
    require(os.access(binary_path, os.X_OK), "plan.binary executable bit missing")
    require(plan["binary"]["path"] == str(binary_path), "plan binary path")
    require(plan["binary"] == binding["binary"], "plan binary differs from source-harness")
    fixtures_value = plan.get("fixtures")
    require(isinstance(fixtures_value, dict) and set(fixtures_value) == set(CASES), "plan fixture set")
    fixture_descriptors: dict[str, dict[str, Any]] = {}
    for case, relative in CASE_FILES.items():
        path = verify_descriptor(fixtures_value[case], f"plan.fixture.{case}")
        require(path == source_root / relative, f"plan.fixture.{case} path")
        fixture_descriptors[case] = dict(fixtures_value[case])
    if binding.get("fixtures") is not None:
        bound_fixtures = {
            Path(item["path"]).resolve(): item
            for item in binding["fixtures"]
        }
        require(
            {Path(item["path"]).resolve(): item for item in fixture_descriptors.values()}
            == bound_fixtures,
            "plan fixtures differ from source-harness fixtures",
        )
    oracle_path = verify_descriptor(plan.get("ordinary_output_oracle"), "plan.ordinary_output_oracle")
    oracle_value = load_json(oracle_path, "ordinary output oracle")
    require(isinstance(oracle_value, dict) and oracle_value.get("schema") in ORACLE_SCHEMAS and oracle_value.get("status") == "byte_equivalent" and oracle_value.get("source_revision") == revision, "ordinary oracle identity")
    if oracle_value["schema"].endswith(".v1"):
        require("schema_version" not in oracle_value, "ordinary v1 oracle must omit schema_version")
    else:
        require(oracle_value.get("schema_version") == 2, "ordinary v2 oracle schema_version")
    oracle_cells = oracle_value.get("cells")
    expected_ids = {cell["id"] for cell in expected_cells()}
    require(isinstance(oracle_cells, dict) and set(oracle_cells) == expected_ids, "ordinary oracle exact 24 cells")
    for cell_id, output in oracle_cells.items():
        require(isinstance(output, dict) and set(output) == {"sha256", "bytes"}, f"ordinary oracle {cell_id}: fields")
        check_hash(output["sha256"], f"ordinary oracle {cell_id}.sha256")
        check_nonnegative(output["bytes"], f"ordinary oracle {cell_id}.bytes")
    require(plan.get("ordinary_output_oracle_cells") == oracle_cells, "plan ordinary oracle copy")
    equivalence_path = verify_descriptor(plan.get("ordinary_oracle_equivalence"), "plan.ordinary_oracle_equivalence")
    equivalence = load_json(equivalence_path, "ordinary oracle equivalence")
    require(isinstance(equivalence, dict) and equivalence.get("schema") == EQUIVALENCE_SCHEMA and equivalence.get("schema_version") == 1, "ordinary oracle equivalence schema")
    require(equivalence.get("status") == "pass_byte_equivalent_output_oracle_only", "ordinary oracle equivalence status")
    decision = equivalence.get("decision")
    require(isinstance(decision, dict) and decision.get("ordinary_0967_binary_may_be_reused_as_0968_output_oracle") is True and decision.get("ordinary_0967_binary_may_be_reused_as_0968_diagnostic_executable") is False and decision.get("performance_or_timing_equivalence_claim") is False and decision.get("allocator_or_rss_equivalence_claim") is False, "ordinary oracle equivalence decision")
    oracle_file = equivalence.get("retained_output_validation", {}).get("oracle_file")
    if not isinstance(oracle_file, dict):
        oracle_file = equivalence.get("historical_packet_source", {}).get("oracle_file")
    require(isinstance(oracle_file, dict) and Path(oracle_file.get("path", "")).resolve() == oracle_path.resolve() and oracle_file.get("sha256") == sha256_file(oracle_path), "ordinary oracle equivalence oracle binding")
    require(bool(equivalence.get("reuse_conditions")) and equivalence.get("review_constraints", {}).get("workload_started_by_review") is False, "ordinary oracle equivalence reuse constraints")
    target_revision = equivalence.get("target_0968", {}).get("source_revision")
    require(target_revision is None or target_revision == revision, "ordinary oracle equivalence source revision")
    review = plan.get("review")
    require(isinstance(review, dict), "plan review")
    review_path = verify_descriptor(review.get("document"), "plan.review.document")
    review_value = load_json(review_path, "independent review")
    require(isinstance(review_value, dict) and review_value.get("status") in {"pass", "passed", "PASS"}, "independent review status")
    require(review.get("status") == review_value.get("status"), "plan review status binding")
    tools = plan.get("tools")
    require(isinstance(tools, dict), "plan tools")
    for label in ("runner", "reducer", "schema", "tests"):
        verify_descriptor(tools.get(label), f"plan.tool.{label}")
    scope = plan.get("claim_scope")
    require(isinstance(scope, dict), "plan claim scope")
    for field in ("diagnostic_only", "timing_claim", "allocator_total_claim", "elapsed_time_claim", "optimization_candidate", "adoption_claim"):
        expected = field == "diagnostic_only"
        require(scope.get(field) is expected, f"plan claim scope {field}")
    contract = plan.get("report_contract")
    require(isinstance(contract, dict), "plan report contract")
    require(contract.get("elapsed_time_field") == "forbidden" and contract.get("metadata_projection_counter") == "forbidden", "plan report contract permits forbidden fields")
    require(plan.get("report_schema") == REPORT_SCHEMA, "plan report schema")
    policy = plan.get("run_policy")
    require(isinstance(policy, dict), "plan run policy")
    expected_policy = {
        "serial": True,
        "cargo_processes": 0,
        "retry": False,
        "warmups": 0,
        "samples": 1,
        "deadline_seconds": 300,
        "max_stdout_bytes": MAX_STDOUT_BYTES,
        "max_stderr_bytes": MAX_STDERR_BYTES,
        "terminal_before_parse": True,
        "raw_stdout_stderr_retained": True,
        "cpu_affinity": None,
    }
    require(policy == expected_policy, "plan run policy changed")
    require(plan.get("environment") == {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC", "RUST_BACKTRACE": "0", "PAGER": "cat"}, "plan environment changed")
    rows = plan.get("schedule", {}).get("rows") if isinstance(plan.get("schedule"), dict) else None
    require(rows == expected_rows(), "plan fixed 48-row schedule changed")
    require(Path(plan.get("capture_root", "")).is_absolute(), "plan capture root absolute path")
    argv_template = plan.get("argv_template")
    expected_argv_tail = ["--input", "{fixture}", "--case", "{case}", "--edits", "{edits}", "--retention", "{retention_arg}", "--warmup", "0", "--samples", "1"]
    require(isinstance(argv_template, list) and argv_template == [str(binary_path), *expected_argv_tail], "plan argv template changed")
    require(plan.get("no_retry") is True and plan.get("no_candidate") is True, "plan retry/candidate policy")
    return plan, {"source": source, "source_harness": source_harness, "binary": dict(plan["binary"]), "fixtures": fixture_descriptors, "review": review_value}, oracle_cells


def verify_go(go_path: Path, plan_path: Path, plan: dict[str, Any]) -> dict[str, Any]:
    go = load_json(go_path, "GO")
    require(isinstance(go, dict) and go.get("schema") == GO_SCHEMA and go.get("status") == "authorized" and go.get("retry") is False, "GO schema/status/retry")
    require(go.get("plan_sha256") == sha256_file(plan_path), "GO plan binding")
    require(go.get("source_scope_sha256") == plan["source"]["scope_sha256"], "GO source binding")
    require(go.get("binary_sha256") == plan["binary"]["sha256"], "GO binary binding")
    require(go.get("source_harness_sha256") == plan["source_harness"]["sha256"], "GO source-harness binding")
    return go


def expected_argv(plan: dict[str, Any], row: dict[str, Any]) -> list[str]:
    fixture = plan["fixtures"][row["case"]]["path"]
    substitutions = {"{fixture}": fixture, "{case}": row["case"], "{edits}": str(row["edits"]), "{retention_arg}": row["retention_arg"]}
    result: list[str] = []
    for item in plan["argv_template"]:
        value = str(item)
        for key, replacement in substitutions.items():
            value = value.replace(key, replacement)
        require("{" not in value and "}" not in value, f"{row['child_id']}: unbound argv placeholder")
        result.append(value)
    return result


def check_child(item: Any, expected: dict[str, Any], plan: dict[str, Any], fixtures: dict[str, dict[str, Any]], oracle: dict[str, dict[str, Any]], plan_path: Path, capture_root: Path) -> dict[str, Any]:
    label = expected["child_id"]
    require(isinstance(item, dict), f"{label}: manifest child object")
    required_item = {"child_id", "row", "root", "status", "problems", "start", "terminal", "stdout", "stderr", "report", "failure"}
    require(set(item) == required_item, f"{label}: manifest child fields")
    require(item["child_id"] == label and item["row"] == expected and item["status"] == "passed" and item["problems"] == [] and item["failure"] is None, f"{label}: manifest child identity/status")
    child_root = path_directory(item["root"], f"{label}.root")
    require(child_root == capture_root / label, f"{label}: child path")
    expected_names = {"start.json", "terminal.json", "stdout.raw", "stderr.raw", "report.json", "parse.json"}
    require({p.name for p in child_root.iterdir()} == expected_names, f"{label}: unexpected child files")
    start_path = verify_descriptor(item["start"], f"{label}.start", child_root / "start.json", child_root)
    terminal_path = verify_descriptor(item["terminal"], f"{label}.terminal", child_root / "terminal.json", child_root)
    stdout_path = verify_descriptor(item["stdout"], f"{label}.stdout", child_root / "stdout.raw", child_root)
    stderr_path = verify_descriptor(item["stderr"], f"{label}.stderr", child_root / "stderr.raw", child_root)
    report_path = verify_descriptor(item["report"], f"{label}.report", child_root / "report.json", child_root)
    parse_path = path_regular(child_root / "parse.json", f"{label}.parse")
    require(parse_path.stat().st_size < MAX_JSON_BYTES, f"{label}: parse exceeds bound")
    parse = load_json(parse_path, f"{label}.parse")
    require(isinstance(parse, dict) and parse.get("schema") == CHILD_PARSE_SCHEMA and parse.get("status") == "parsed" and parse.get("child_id") == label, f"{label}: parse identity")
    require(parse.get("report") == item["report"], f"{label}: parse report binding")
    start = load_json(start_path, f"{label}.start")
    require(isinstance(start, dict) and start.get("schema") == CHILD_START_SCHEMA and start.get("status") == "started" and start.get("child_id") == label, f"{label}: start identity")
    require(start.get("row") == expected and start.get("plan_sha256") == sha256_file(plan_path), f"{label}: start plan/row binding")
    require(start.get("source") == plan["source"] and start.get("source_harness") == plan["source_harness"] and start.get("binary") == plan["binary"], f"{label}: start custody binding")
    require(start.get("fixture") == plan["fixtures"][expected["case"]], f"{label}: start fixture binding")
    require(start.get("cwd") == plan["source"]["root"] and start.get("environment") == plan["environment"], f"{label}: start execution binding")
    require(start.get("argv") == expected_argv(plan, expected), f"{label}: argv binding")
    terminal = load_json(terminal_path, f"{label}.terminal")
    require(isinstance(terminal, dict) and terminal.get("schema") == CHILD_TERMINAL_SCHEMA and terminal.get("status") == "terminal" and terminal.get("child_id") == label, f"{label}: terminal identity")
    require(terminal.get("returncode") == 0 and terminal.get("timed_out") is False and terminal.get("launch_error") is None, f"{label}: terminal outcome")
    require(terminal.get("stdout") == item["stdout"] and terminal.get("stderr") == item["stderr"], f"{label}: terminal stream descriptor binding")
    require(stdout_path.stat().st_size <= MAX_STDOUT_BYTES and stderr_path.stat().st_size <= MAX_STDERR_BYTES, f"{label}: stream bound")
    try:
        stdout_report = json.loads(stdout_path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicate_pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise AuditError(f"{label}: stdout report JSON: {exc}") from exc
    report = load_json(report_path, f"{label}.report")
    require(stdout_report == report, f"{label}: stdout/report mismatch")
    check_report(report, expected, fixtures, oracle, label)
    return {"child_id": label, "row": expected, "report": report, "census": report["samples"][0]["census"]}


def empty_counters() -> dict[str, int]:
    return {field: 0 for field in CENSUS_FIELDS if field != "saturated"}


def add_counters(target: dict[str, int], census: dict[str, Any]) -> None:
    for field in target:
        target[field] += census[field]


def summarize(records: list[dict[str, Any]]) -> dict[str, Any]:
    overall = empty_counters()
    by_case: dict[str, dict[str, int]] = {case: empty_counters() for case in CASES}
    by_retention: dict[str, dict[str, int]] = {retention: empty_counters() for retention in REPORT_RETENTION.values()}
    by_cell: dict[str, dict[str, int]] = {cell["id"]: empty_counters() for cell in expected_cells()}
    counts = {"overall": len(records), "by_case": {case: 0 for case in CASES}, "by_retention": {name: 0 for name in REPORT_RETENTION.values()}, "by_cell": {cell["id"]: 0 for cell in expected_cells()}}
    for record in records:
        row = record["row"]
        census = record["census"]
        add_counters(overall, census)
        add_counters(by_case[row["case"]], census)
        report_retention = record["report"]["retention"]
        add_counters(by_retention[report_retention], census)
        add_counters(by_cell[row["id"]], census)
        counts["by_case"][row["case"]] += 1
        counts["by_retention"][report_retention] += 1
        counts["by_cell"][row["id"]] += 1
    return {"counts": counts, "overall": overall, "by_case": by_case, "by_retention": by_retention, "by_cell": by_cell}


def audit(plan_path: Path, go_path: Path, manifest_path: Path) -> dict[str, Any]:
    plan_path = path_regular(plan_path, "plan")
    go_path = path_regular(go_path, "GO")
    manifest_path = path_regular(manifest_path, "capture manifest")
    plan, custody, oracle = verify_plan(plan_path)
    go = verify_go(go_path, plan_path, plan)
    capture_root = manifest_path.parent
    require(Path(plan["capture_root"]).resolve() == capture_root.resolve(), "capture root differs from plan")
    manifest = load_json(manifest_path, "capture manifest")
    require(isinstance(manifest, dict) and manifest.get("schema") == CAPTURE_SCHEMA and manifest.get("status") == "complete", "capture manifest schema/status")
    require(manifest.get("plan") == descriptor(plan_path, "plan"), "capture plan descriptor")
    require(manifest.get("go") == descriptor(go_path, "GO"), "capture GO descriptor")
    require(manifest.get("source") == plan["source"] and manifest.get("source_harness") == plan["source_harness"] and manifest.get("binary") == plan["binary"], "capture source/binary custody")
    require(manifest.get("packet_count") == 48 and manifest.get("retries") == 0 and manifest.get("failures") == [], "capture packet/retry/failure policy")
    run_start_path = path_regular(capture_root / "run-start.json", "run-start")
    run_terminal_path = path_regular(capture_root / "run-terminal.json", "run-terminal")
    run_start = load_json(run_start_path, "run-start")
    require(isinstance(run_start, dict) and run_start.get("schema") == RUN_START_SCHEMA and run_start.get("status") == "started", "run-start identity")
    require(run_start.get("plan") == descriptor(plan_path, "plan") and run_start.get("go") == descriptor(go_path, "GO"), "run-start plan/GO custody")
    require(run_start.get("source") == plan["source"] and run_start.get("source_harness") == plan["source_harness"] and run_start.get("binary") == plan["binary"], "run-start source/binary custody")
    require(run_start.get("children") == 48 and run_start.get("cells") == 24 and run_start.get("retries") == 0, "run-start matrix/retry policy")
    run_terminal = load_json(run_terminal_path, "run-terminal")
    require(isinstance(run_terminal, dict) and run_terminal.get("schema") == RUN_TERMINAL_SCHEMA and run_terminal.get("status") == "completed", "run-terminal identity")
    require(run_terminal.get("children") == 48 and run_terminal.get("failures") == [] and run_terminal.get("retries") == 0, "run-terminal completion/retry policy")
    expected = expected_rows()
    expected_by_id = {row["child_id"]: row for row in expected}
    items = manifest.get("children")
    require(isinstance(items, list) and len(items) == 48, "capture child list")
    ids = [item.get("child_id") if isinstance(item, dict) else None for item in items]
    require(set(ids) == set(expected_by_id) and len(ids) == len(set(ids)), "capture child identity set")
    top_names = {entry.name for entry in capture_root.iterdir()}
    require(top_names == {"run-start.json", "run-terminal.json", manifest_path.name, *expected_by_id}, "capture root has unexpected files")
    records = []
    for item in items:
        row = expected_by_id[item["child_id"]]
        records.append(check_child(item, row, plan, custody["fixtures"], oracle, plan_path, capture_root))
    cell_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        cell_groups[record["row"]["id"]].append(record)
    repeat_results: dict[str, Any] = {}
    for cell in expected_cells():
        rows = sorted(cell_groups[cell["id"]], key=lambda record: record["row"]["repeat"])
        require(len(rows) == 2 and {record["row"]["repeat"] for record in rows} == {1, 2}, f"{cell['id']}: repetition set")
        equal = rows[0]["report"] == rows[1]["report"]
        require(equal, f"{cell['id']}: repeat report mismatch")
        repeat_results[cell["id"]] = {"children": [record["child_id"] for record in rows], "equal": equal, "output": {"sha256": rows[0]["report"]["output_sha256"], "bytes": rows[0]["report"]["output_bytes"]}}
    summary = summarize(records)
    return {
        "schema": AUDIT_SCHEMA,
        "status": "pass",
        "decision": "diagnostic_only_no_adoption",
        "created_utc": utc_now(),
        "tool": {"path": str(Path(__file__).resolve()), "bytes": Path(__file__).stat().st_size, "sha256": sha256_file(Path(__file__).resolve())},
        "scope": {"read_only": True, "new_build": False, "new_capture": False, "timing_claim": False, "allocator_total_claim": False, "speedup_claim": False, "adoption_claim": False},
        "inputs": {"plan": descriptor(plan_path, "plan"), "go": descriptor(go_path, "GO"), "capture_manifest": descriptor(manifest_path, "capture manifest")},
        "custody": {"source_revision": plan["source_revision"], "source_root": plan["source"]["root"], "source_scope_files": plan["source"]["file_count"], "source_scope_sha256": plan["source"]["scope_sha256"], "binary": plan["binary"], "ordinary_oracle_cells": 24, "children": 48, "cells": 24, "retries": 0, "terminal_before_parse": True},
        "checks": {"fixed_schedule": True, "source_executable_plan_go_manifest_terminal_bound": True, "all_child_packets_custodied": True, "all_reports_and_routes_schema_valid": True, "all_20_correctness_gates_true": True, "ordinary_output_sha_length_match": True, "all_24_repeats_equal": True, "per_packet_stage_conservation": True, "no_timer_or_allocator_measurement_fields": True},
        "repeat_results": repeat_results,
        "route_payload_summary": summary,
        "findings": [],
    }


def write_exclusive(path: Path, value: Any) -> None:
    raw = Path(path)
    require(not raw.exists() and not raw.is_symlink(), f"output already exists: {raw}")
    raw.parent.mkdir(parents=True, exist_ok=True)
    payload = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(payload) < MAX_JSON_BYTES, "audit output exceeds bound")
    with raw.open("xb") as stream:
        stream.write(payload)
        stream.flush()
        os.fsync(stream.fileno())


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--go", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output = Path(args.output).resolve()
    require(not output.exists() and not output.is_symlink(), f"output already exists: {output}")
    try:
        result = audit(args.plan, args.go, args.manifest)
        code = 0
    except (AuditError, OSError) as exc:
        result = {
            "schema": AUDIT_SCHEMA,
            "status": "rejected",
            "decision": "diagnostic_only_no_adoption",
            "created_utc": utc_now(),
            "tool": {"path": str(Path(__file__).resolve()), "bytes": Path(__file__).stat().st_size, "sha256": sha256_file(Path(__file__).resolve())},
            "scope": {"read_only": True, "new_build": False, "new_capture": False, "timing_claim": False, "allocator_total_claim": False, "speedup_claim": False, "adoption_claim": False},
            "inputs": {"plan": str(args.plan.resolve()), "go": str(args.go.resolve()), "capture_manifest": str(args.manifest.resolve())},
            "findings": [str(exc)],
        }
        code = 2
    write_exclusive(output, result)
    print(json.dumps({"schema": result["schema"], "status": result["status"], "output": str(output), "findings": len(result["findings"])}, sort_keys=True))
    return code


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AuditError, OSError) as exc:
        raise SystemExit(f"FAIL: {exc}")
