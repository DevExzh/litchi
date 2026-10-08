#!/usr/bin/env python3
"""Schema and custody checks for the 0968 staged CFB census v2.

The v2 packet follows the sealed diagnostic harness directly.  It deliberately
keeps the meaningful route counters only: commit, overlay, layout and package
capture.  The harness guarantees that a successful commit scope is closed only
after discovery succeeds; no separate low-level open, allocation or discovery
counter is projected into the packet.

This module never launches a process and never writes an artifact.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
from pathlib import Path
from typing import Any

# The child report remains the sealed harness's v2 packet.  The plan and
# custody envelopes are v3 so the corrected per-edit cardinality contract
# cannot accidentally consume the retained v2 plan or capture.
PLAN_SCHEMA = "litchi.performance.0968.cfb_commit_census_plan.v3"
REPORT_SCHEMA = "litchi.performance.0968.cfb_commit_census.v2"
SOURCE_HARNESS_SCHEMA = "litchi.performance.0968.cfb_commit_census_source_harness.v2"
PROVENANCE_SCHEMAS = {
    "litchi.performance.0968.cfb_commit_census.provenance.v2",
    "litchi.performance.0968.cfb_commit_census.provenance.v3",
}
SEAL_SCHEMAS = {"litchi.performance.0968.cfb_commit_census.seal.v4"}
ORACLE_SCHEMAS = {
    "litchi.performance.0968.cfb_commit_ordinary_oracle.v1",
    "litchi.performance.0968.cfb_commit_ordinary_oracle.v2",
}
EQUIVALENCE_SCHEMA = "litchi.performance.0968.ordinary_oracle_equivalence.v1"
ORACLE_SCHEMA = "litchi.performance.0968.cfb_commit_ordinary_oracle.v2"
GO_SCHEMA = "litchi.performance.0968.cfb_commit_census_go.v3"
CAPTURE_SCHEMA = "litchi.performance.0968.cfb_commit_census_capture.v3"
REDUCTION_SCHEMA = "litchi.performance.0968.cfb_commit_census_reduction.v3"
INCIDENT_SCHEMA = "litchi.performance.0968.cfb_commit_census_incident.v1"
CHILD_START_SCHEMA = "litchi.performance.0968.cfb_child_start.v3"
CHILD_TERMINAL_SCHEMA = "litchi.performance.0968.cfb_child_terminal.v3"
CHILD_PARSE_SCHEMA = "litchi.performance.0968.cfb_child_parse.v3"
CHILD_FAILURE_SCHEMA = "litchi.performance.0968.cfb_child_failure.v3"
RUN_START_SCHEMA = "litchi.performance.0968.cfb_census_run_start.v3"
RUN_TERMINAL_SCHEMA = "litchi.performance.0968.cfb_census_run_terminal.v3"

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
    # This is a retained 0967 oracle name.  It is a boolean correctness
    # check, not an allocator measurement or claim in the v2 packet.
    "allocator_retained_matches_output_capacity",
)
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
HEX = set("0123456789abcdef")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
MAX_JSON_BYTES = 1 << 20
MAX_STDOUT_BYTES = 1 << 20
MAX_STDERR_BYTES = 1 << 20
SKIP_TREE_DIRS = {".git", "target", "__pycache__"}


class CensusError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CensusError(message)


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        require(key not in value, f"JSON duplicate key: {key}")
        value[key] = item
    return value


def load_json(path: Path, *, maximum: int = MAX_JSON_BYTES) -> Any:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    require(path.stat().st_size < maximum, f"JSON exceeds bound: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=_reject_duplicate_keys)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise CensusError(f"invalid JSON {path}: {exc}") from exc


def regular_file(path: Path, label: str) -> Path:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"{label}: not a regular file: {path}")
    return path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with regular_file(path, "hash").open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, label: str) -> dict[str, Any]:
    path = regular_file(path, label)
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_file(path)}


def verify_descriptor(value: Any, label: str, *, executable: bool = False) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}: path missing")
    path = regular_file(Path(path_value), label)
    require(type(value.get("bytes")) is int and value["bytes"] == path.stat().st_size,
            f"{label}: byte count changed")
    expected = value.get("sha256")
    check_hex(expected, f"{label}.sha256")
    require(sha256_file(path) == expected, f"{label}: SHA-256 changed")
    if executable:
        require(os.access(path, os.X_OK), f"{label}: executable bit missing")
    return path


def verify_reference(value: Any, label: str) -> Path:
    """Verify a sealed path/SHA reference whose producer omitted byte count."""
    require(isinstance(value, dict), f"{label}: reference missing")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}: path missing")
    path = regular_file(Path(path_value), label)
    if "bytes" in value:
        require(type(value["bytes"]) is int and value["bytes"] == path.stat().st_size,
                f"{label}: byte count changed")
    check_hex(value.get("sha256"), f"{label}.sha256")
    require(sha256_file(path) == value["sha256"], f"{label}: SHA-256 changed")
    return path


def source_tree_digest(root: Path) -> tuple[int, str]:
    """Digest source/harness files while excluding generated build caches."""
    root = root.resolve()
    require(root.is_dir() and not root.is_symlink(), f"source root missing: {root}")
    entries: list[tuple[str, int, str]] = []
    for current, directories, files in os.walk(root, topdown=True, followlinks=False):
        current_path = Path(current)
        directories[:] = sorted(name for name in directories if name not in SKIP_TREE_DIRS)
        files.sort()
        for name in files:
            path = current_path / name
            regular_file(path, "source tree")
            entries.append((path.relative_to(root).as_posix(), path.stat().st_size, sha256_file(path)))
        for name in directories:
            require(not (current_path / name).is_symlink(), f"source tree symlink: {current_path / name}")
    digest = hashlib.sha256()
    for relative, size, file_hash in entries:
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        digest.update(size.to_bytes(8, "big"))
        digest.update(file_hash.encode("ascii"))
        digest.update(b"\n")
    return len(entries), digest.hexdigest()


def source_scope_digest(root: Path, files: list[dict[str, Any]]) -> tuple[int, str, list[dict[str, Any]]]:
    """Bind the explicit diagnostic/runtime file scope without walking docs or targets."""
    root = root.resolve()
    entries: list[dict[str, Any]] = []
    seen: set[str] = set()
    for index, item in enumerate(files):
        require(isinstance(item, dict), f"source scope[{index}]: descriptor")
        path = verify_descriptor(item, f"source scope[{index}]")
        try:
            relative = path.relative_to(root).as_posix()
        except ValueError as exc:
            raise CensusError(f"source scope[{index}]: path escapes source root") from exc
        require(relative not in seen, f"source scope: duplicate {relative}")
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
            rows.append({
                **cell,
                "child_id": f"cfb-r{repeat:02d}-{cell['case']}-e{cell['edits']:02d}-{cell['retention']}",
                "position": position,
                "repeat": repeat,
                "direction": direction,
                "repeat_position": repeat_position,
                "warmups": 0,
                "samples": 1,
                "retry": False,
            })
    return rows


def check_hex(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label}: invalid SHA-256")


def check_int(value: Any, label: str) -> None:
    require(type(value) is int and value >= 0, f"{label}: invalid nonnegative integer")


def check_bool(value: Any, label: str) -> None:
    require(type(value) is bool, f"{label}: invalid boolean")


def check_census(value: Any, label: str, expected_edits: int) -> None:
    require(isinstance(value, dict) and set(value) == set(CENSUS_FIELDS), f"{label}: census field set")
    require(type(expected_edits) is int and expected_edits in EDIT_COUNTS,
            f"{label}: invalid expected edit count")
    for field in CENSUS_FIELDS[:-1]:
        check_int(value[field], f"{label}.{field}")
    check_bool(value["saturated"], f"{label}.saturated")
    require(value["saturated"] is False, f"{label}: saturated census")
    # One child invokes the diagnostic commit loop once per requested edit.
    # The v2 checker incorrectly treated every child as a one-edit packet;
    # retain the measured per-edit cardinality explicitly in this contract.
    require(value["commit_attempts"] == expected_edits,
            f"{label}: expected {expected_edits} commit attempts")
    require(value["commit_attempts"] == value["commit_successes"] + value["commit_errors"],
            f"{label}: commit conservation")
    require(value["commit_successes"] == expected_edits and value["commit_errors"] == 0,
            f"{label}: expected {expected_edits} successful commits")
    require(value["overlay_attempted"] + value["overlay_not_attempted"] == value["commit_attempts"],
            f"{label}: overlay stage conservation")
    require(value["overlay_attempted"] == value["overlay_success"] + value["overlay_declined"] + value["overlay_error"],
            f"{label}: overlay outcome conservation")
    require(value["overlay_error"] == 0, f"{label}: overlay error")
    require(value["layout_attempted"] == value["layout_reuse_success"] + value["layout_non_reuse"] + value["layout_error"],
            f"{label}: layout conservation")
    require(value["layout_error"] == 0, f"{label}: layout error")
    require(value["capture_attempted"] == value["capture_success"] + value["capture_error"],
            f"{label}: capture conservation")
    require(value["capture_attempted"] == expected_edits and
            value["capture_success"] == expected_edits and value["capture_error"] == 0,
            f"{label}: expected {expected_edits} successful captures")
    require(value["capture_streams"] > 0, f"{label}: successful capture has no streams")


def check_correctness(value: Any, label: str) -> None:
    require(isinstance(value, dict) and set(value) == set(GATE_NAMES), f"{label}: gate field set")
    for field in GATE_NAMES:
        require(value[field] is True, f"{label}.{field}: gate failed")


def check_oracle(value: Any, label: str) -> None:
    require(isinstance(value, dict) and set(value) == ORACLE_FIELDS, f"{label}: oracle field set")
    for field in ORACLE_FIELDS - {"failure_reasons"}:
        require(value[field] is True, f"{label}.{field}: oracle failed")
    require(value["failure_reasons"] == [], f"{label}: oracle failure reasons")


def check_stages(value: Any, label: str) -> None:
    require(isinstance(value, list) and len(value) == 32, f"{label}: stage vector")
    fields = {"stage", "requested_text", "before", "after", "after_release"}
    for index, stage in enumerate(value):
        require(isinstance(stage, dict) and set(stage) == fields, f"{label}[{index}]: stage fields")
        check_int(stage["stage"], f"{label}[{index}].stage")
        require(stage["stage"] <= 32, f"{label}[{index}].stage: bound")
        require(isinstance(stage["requested_text"], str), f"{label}[{index}].requested_text")
        for field in ("before", "after", "after_release"):
            require(stage[field] is None or (type(stage[field]) is int and stage[field] >= 0),
                    f"{label}[{index}].{field}")


def check_output(value: Any, expected: dict[str, Any], label: str) -> None:
    require(isinstance(value, dict) and set(value) == {"sha256", "bytes"}, f"{label}: output fields")
    check_hex(value["sha256"], f"{label}.sha256")
    check_int(value["bytes"], f"{label}.bytes")
    require(value == expected, f"{label}: differs from ordinary byte-equivalence oracle")


def check_route(value: Any, row: dict[str, Any], plan: dict[str, Any], oracle: dict[str, dict[str, Any]], label: str) -> None:
    require(isinstance(value, dict) and set(value) == ROUTE_FIELDS, f"{label}: route fields")
    retention_name = value["retention"]
    require(retention_name in REPORT_RETENTION.values(), f"{label}: route retention")
    retention = next(key for key, name in REPORT_RETENTION.items() if name == retention_name)
    require(value["retention_ceiling_bytes"] == RETENTION_CEILING[retention], f"{label}: retention ceiling")
    check_hex(value["output_sha256"], f"{label}.output_sha256")
    for field in ("output_bytes", "output_capacity", "semantic_change_count"):
        check_int(value[field], f"{label}.{field}")
    require(value["semantic_change_count"] == row["edits"], f"{label}: semantic change count")
    require(value["final_retained_before_commit"] is None or type(value["final_retained_before_commit"]) is int,
            f"{label}: retained value")
    check_stages(value["stages"], f"{label}.stages")
    check_oracle(value["oracle"], f"{label}.oracle")
    expected = oracle[f"{row['case']}/e{row['edits']}/{retention}"]
    require(value["output_sha256"] == expected["sha256"] and value["output_bytes"] == expected["bytes"],
            f"{label}: route output differs from ordinary oracle")


def check_sample(value: Any, row: dict[str, Any], plan: dict[str, Any], oracle: dict[str, dict[str, Any]], expected: dict[str, Any], label: str) -> None:
    require(isinstance(value, dict) and set(value) == SAMPLE_FIELDS, f"{label}: sample fields")
    require(value["index"] == 0, f"{label}: sample index")
    check_hex(value["output_sha256"], f"{label}.output_sha256")
    for field in ("output_bytes", "output_capacity", "semantic_change_count"):
        check_int(value[field], f"{label}.{field}")
    require(value["output_sha256"] == expected["output_sha256"] and value["output_bytes"] == expected["output_bytes"],
            f"{label}: selected output mismatch")
    require(value["semantic_change_count"] == row["edits"], f"{label}: semantic change count")
    for field in (
        "exact_output_equal_reference",
        "final_paragraph_semantic_readback",
        "final_public_reader_reopen",
        "retained_capacities_within_ceiling",
        "allocator_retained_matches_output_capacity",
    ):
        require(value[field] is True, f"{label}.{field}: false")
    check_oracle(value["oracle"], f"{label}.oracle")
    check_stages(value["stages"], f"{label}.stages")
    check_census(value["census"], f"{label}.census", row["edits"])


def check_report(report: Any, row: dict[str, Any], plan: dict[str, Any], oracle: dict[str, dict[str, Any]]) -> None:
    require(isinstance(report, dict) and set(report) == REPORT_FIELDS, "report field set")
    require(report["schema"] == REPORT_SCHEMA and report["schema_version"] == 2 and report["change"] == "0968",
            "report schema/change")
    require(report["case"] == row["case"] and report["edits"] == row["edits"], "report cell identity")
    require(report["retention"] == REPORT_RETENTION[row["retention"]], "report retention identity")
    require(isinstance(report["input"], str) and report["input"], "report input")
    fixture = plan["fixtures"][row["case"]]
    require(report["source_sha256"] == fixture["sha256"] and report["source_bytes"] == fixture["bytes"],
            "report fixture identity")
    require(report["selected_paragraph"] == 0, "report selected paragraph")
    require(tuple(report["alternating_texts"]) == ALTERNATING_TEXTS, "report alternating texts")
    require(report["retention_ceiling_bytes"] == RETENTION_CEILING[row["retention"]], "report retention ceiling")
    require(report["timing_claim"] is False and isinstance(report["region_definition"], str), "report timing scope")
    require("ObjectEditor::commit_candidate_with_rendered" in report["region_definition"], "report census boundary")
    require(report["warmups"] == 0 and report["samples_requested"] == 1, "report sample policy")
    check_hex(report["output_sha256"], "report.output_sha256")
    for field in ("output_bytes", "output_capacity", "reference_output_bytes"):
        check_int(report[field], f"report.{field}")
    check_hex(report["reference_output_sha256"], "report.reference_output_sha256")
    expected = oracle[row["id"]]
    require(report["output_sha256"] == expected["sha256"] and report["output_bytes"] == expected["bytes"],
            "report output differs from ordinary oracle")
    reference = oracle[f"{row['case']}/e{row['edits']}/zero"]
    require(report["reference_output_sha256"] == reference["sha256"] and report["reference_output_bytes"] == reference["bytes"],
            "report reference output differs from ordinary oracle")
    require(report["exact_output_equal_reference"] is True, "report output/reference mismatch")
    check_correctness(report["correctness"], "report correctness")
    require(isinstance(report["route_outputs"], list) and len(report["route_outputs"]) == 3, "report route outputs")
    route_names = [route["retention"] for route in report["route_outputs"]]
    require(route_names == [REPORT_RETENTION[name] for name in RETENTIONS], "report route order")
    for index, route in enumerate(report["route_outputs"]):
        check_route(route, row, plan, oracle, f"report.route_outputs[{index}]")
    require(isinstance(report["samples"], list) and len(report["samples"]) == 1, "report samples")
    check_sample(report["samples"][0], row, plan, oracle, report, "report.samples[0]")


def check_source_harness_manifest(value: Any, label: str, source_revision: str) -> dict[str, Any]:
    """Validate either the sealed source-harness manifest or provenance v2.

    The current author receipt is provenance-shaped: it binds the immutable
    baseline archive and an explicit list of changed/runtime/harness files,
    instead of claiming custody over the generated target tree.  Normalize
    both accepted shapes to the same source binding for the plan.
    """
    require(isinstance(value, dict), f"{label}: object")
    seal_id = value.get("schema")
    if value.get("status") == "SEALED_DIAGNOSTIC_SOURCE_HARNESS" and seal_id in SEAL_SCHEMAS:
        require(value.get("status") == "SEALED_DIAGNOSTIC_SOURCE_HARNESS", f"{label}: seal status")
        require(value.get("source_baseline_head") == source_revision, f"{label}: seal source revision")
        provenance_receipt = value.get("provenance")
        require(isinstance(provenance_receipt, dict), f"{label}: provenance receipt")
        provenance_path = verify_reference(provenance_receipt, f"{label}.provenance")
        provenance = load_json(provenance_path)
        provenance_binding = check_source_harness_manifest(provenance, f"{label}.provenance", source_revision)
        patch = value.get("code_patch")
        require(isinstance(patch, dict), f"{label}: code patch binding")
        verify_descriptor(patch, f"{label}.code_patch")
        schema_receipt = value.get("schema_receipt")
        require(isinstance(schema_receipt, dict), f"{label}: schema receipt")
        schema_path = verify_reference(schema_receipt, f"{label}.schema_receipt")
        schema_value = load_json(schema_path)
        require(schema_value.get("schema") == REPORT_SCHEMA and schema_value.get("schema_version") == 2,
                f"{label}: schema receipt mismatch")
        require(value.get("scope", {}).get("timing_claim") is False and
                value.get("scope", {}).get("allocator_instrumentation") is False and
                value.get("scope", {}).get("metadata_projection") is False,
                f"{label}: seal claim scope")
        seal_binary = value.get("binary")
        provenance_binary = provenance.get("binary")
        require(isinstance(seal_binary, dict) and isinstance(provenance_binary, dict), f"{label}: binary receipt")
        require(seal_binary.get("path") == provenance_binary.get("path") and
                seal_binary.get("sha256") == provenance_binary.get("sha256"),
                f"{label}: binary receipt changed")
        require(value.get("scope", {}).get("formal_children") == 48, f"{label}: formal child count")
        return provenance_binding

    if value.get("schema") in PROVENANCE_SCHEMAS:
        require(value.get("state") in {
            "diagnostic_source_harness_sealed_pending_formal_go",
            "diagnostic_source_harness_sealed_pending_oracle_review_and_formal_go",
        },
                f"{label}: provenance state")
        revision = value.get("current_head")
        require(revision == source_revision and REVISION_RE.fullmatch(revision or "") is not None,
                f"{label}: source revision")
        root = Path(value.get("source_diagnostic_path", "")).resolve()
        require(root.is_dir(), f"{label}: diagnostic source root")
        files = value.get("diagnostic_files")
        require(isinstance(files, list) and files, f"{label}: explicit diagnostic file scope")
        count, scope_hash, normalized_files = source_scope_digest(root, files)
        binary = value.get("binary")
        require(isinstance(binary, dict), f"{label}: binary binding")
        verify_descriptor(binary, f"{label}.binary", executable=True)
        fixtures = value.get("fixtures")
        require(isinstance(fixtures, list) and len(fixtures) == len(CASES), f"{label}: fixture binding")
        for index, fixture in enumerate(fixtures):
            verify_descriptor(fixture, f"{label}.fixtures[{index}]")
        schema_receipt = value.get("schema_v2")
        if schema_receipt is not None:
            schema_path = verify_descriptor(schema_receipt, f"{label}.schema_v2")
            schema_value = load_json(schema_path)
            require(schema_value.get("schema") == REPORT_SCHEMA and schema_value.get("schema_version") == 2,
                    f"{label}: schema receipt mismatch")
        patch = value.get("code_patch")
        require(isinstance(patch, dict), f"{label}: code patch binding")
        verify_descriptor(patch, f"{label}.code_patch")
        return {
            "root": str(root),
            "file_count": count,
            "scope_sha256": scope_hash,
            "scope_files": normalized_files,
            "source_revision": revision,
            "baseline_archive_sha256": value.get("source_baseline_archive_sha256"),
            "diagnostic_boundary": {
                "commit_success_after": "object_discovery_returns_successfully",
                "capture_scope": "post_render_package_capture_stream_reads_only",
            },
        }

    require(value.get("schema") == SOURCE_HARNESS_SCHEMA, f"{label}: schema")
    require(value.get("schema_version") == 2 and value.get("status") == "sealed", f"{label}: status/version")
    require(value.get("source_revision") == source_revision and REVISION_RE.fullmatch(source_revision or "") is not None,
            f"{label}: source revision")
    source = value.get("source")
    require(isinstance(source, dict), f"{label}: source binding")
    root = Path(source.get("root", "")).resolve()
    require(root.is_dir(), f"{label}: source root")
    files = source.get("scope_files") or value.get("harness_files")
    require(isinstance(files, list) and files, f"{label}: explicit source scope")
    count, scope_hash, normalized_files = source_scope_digest(root, files)
    require(count == source.get("file_count") and scope_hash == source.get("scope_sha256"),
            f"{label}: source scope changed")
    require(value.get("report_schema") == REPORT_SCHEMA, f"{label}: report schema")
    require(value.get("timing_claim") is False, f"{label}: timing claim")
    if "allocator_total_claim" in value:
        require(value["allocator_total_claim"] is False, f"{label}: allocator claim")
    patch = value.get("diagnostic_patch")
    if patch is not None:
        verify_descriptor(patch, f"{label}.diagnostic_patch")
    boundary = value.get("diagnostic_boundary")
    require(isinstance(boundary, dict), f"{label}: diagnostic boundary")
    require(boundary.get("commit_success_after") in {"object_discovery_returns_successfully", "object_discovery"},
            f"{label}: commit success guard")
    require(boundary.get("capture_scope") in {"post_render_package_capture_stream_reads_only", "post_render_package_capture"},
            f"{label}: capture scope")
    return {
        "root": str(root),
        "file_count": count,
        "scope_sha256": scope_hash,
        "scope_files": normalized_files,
        "source_revision": source_revision,
        "diagnostic_boundary": boundary,
    }


def load_ordinary_oracle(value: Any, label: str, source_revision: str) -> dict[str, dict[str, Any]]:
    require(isinstance(value, dict) and value.get("schema") in ORACLE_SCHEMAS, f"{label}: schema")
    if "schema_version" in value:
        require(value["schema_version"] in {1, 2}, f"{label}: schema version")
    require(value.get("status") == "byte_equivalent" and value.get("source_revision") == source_revision,
            f"{label}: status/source")
    cells = value.get("cells")
    expected = {cell["id"] for cell in expected_cells()}
    require(isinstance(cells, dict) and set(cells) == expected, f"{label}: exact cell set")
    result: dict[str, dict[str, Any]] = {}
    for cell_id, output in cells.items():
        require(isinstance(output, dict) and set(output) == {"sha256", "bytes"}, f"{label}.{cell_id}: output fields")
        check_hex(output["sha256"], f"{label}.{cell_id}.sha256")
        check_int(output["bytes"], f"{label}.{cell_id}.bytes")
        result[cell_id] = output
    return result


def check_oracle_equivalence(value: Any, label: str, oracle_path: Path, oracle_sha256: str, source_revision: str) -> None:
    require(isinstance(value, dict) and value.get("schema") == EQUIVALENCE_SCHEMA and value.get("schema_version") == 1,
            f"{label}: schema")
    require(value.get("status") == "pass_byte_equivalent_output_oracle_only", f"{label}: status")
    decision = value.get("decision")
    require(isinstance(decision, dict) and decision.get("ordinary_0967_binary_may_be_reused_as_0968_output_oracle") is True,
            f"{label}: ordinary oracle reuse decision")
    require(decision.get("ordinary_0967_binary_may_be_reused_as_0968_diagnostic_executable") is False,
            f"{label}: diagnostic executable reuse decision")
    require(decision.get("performance_or_timing_equivalence_claim") is False and
            decision.get("allocator_or_rss_equivalence_claim") is False,
            f"{label}: prohibited equivalence claim")
    oracle_file = value.get("retained_output_validation", {}).get("oracle_file")
    if not isinstance(oracle_file, dict):
        oracle_file = value.get("historical_packet_source", {}).get("oracle_file")
    require(isinstance(oracle_file, dict), f"{label}: oracle file receipt")
    require(Path(oracle_file.get("path", "")).resolve() == oracle_path.resolve() and
            oracle_file.get("sha256") == oracle_sha256, f"{label}: oracle file binding")
    require(value.get("reuse_conditions") and value.get("review_constraints", {}).get("workload_started_by_review") is False,
            f"{label}: reuse constraints")
    target_revision = value.get("target_0968", {}).get("source_revision")
    if target_revision is not None:
        require(target_revision == source_revision, f"{label}: target source revision")


def check_review(value: Any, label: str) -> None:
    require(isinstance(value, dict) and value.get("status") in {"pass", "passed", "PASS"},
            f"{label}: review status")
    if str(value.get("schema", "")).startswith("litchi.performance.0968.review.diagnostic_source_approval"):
        seal = value.get("seal_v4", value.get("seal"))
        require(isinstance(seal, dict) and str(seal.get("path", "")).endswith("/seal-v4.json"),
                f"{label}: seal-v4 binding")
        seal_path = verify_reference(seal, f"{label}.seal_v4")
        seal_value = load_json(seal_path)
        require(seal_value.get("schema") == "litchi.performance.0968.cfb_commit_census.seal.v4" and
                seal_value.get("status") == "SEALED_DIAGNOSTIC_SOURCE_HARNESS",
                f"{label}: seal-v4 content")


def check_incident(value: Any, label: str) -> None:
    """Bind the retained v2 failure without treating it as a new run."""
    require(isinstance(value, dict) and value.get("schema") == INCIDENT_SCHEMA,
            f"{label}: schema")
    require(value.get("status") == "protocol_bug_confirmed_pending_v3_review_and_go" and
            value.get("change") == "0968", f"{label}: status/change")
    finding = value.get("finding")
    require(isinstance(finding, dict), f"{label}: finding")
    require(finding.get("source_or_binary_changed") is False and
            finding.get("workload_retry_performed") is False,
            f"{label}: source/retry disposition")
    require(finding.get("incorrect_v2_requirements") ==
            ["commit_attempts == 1", "capture_attempted == 1"],
            f"{label}: incorrect v2 requirements")
    require(finding.get("correct_contract") == [
        "commit_attempts == row.edits",
        "commit_successes == row.edits",
        "capture_attempted == row.edits",
        "capture_success == row.edits",
    ], f"{label}: corrected contract")
    retained = value.get("retained_v2")
    require(isinstance(retained, dict), f"{label}: retained v2")
    for key in ("plan", "capture_manifest", "run_terminal", "schema_tool", "runner_tool", "binary"):
        verify_descriptor(retained.get(key), f"{label}.{key}")
    require(retained.get("capture_status") == "complete_with_failures" and
            retained.get("packet_count") == 48 and retained.get("retry_count") == 0 and
            retained.get("raw_stdout_reports_audited") == 48,
            f"{label}: retained capture summary")
    require(retained.get("returncodes") == {"0": 48} and
            retained.get("children_by_edits") == {"1": 12, "2": 12, "8": 12, "32": 12},
            f"{label}: retained child summary")
    rows = retained.get("raw_rows")
    require(isinstance(rows, list) and len(rows) == 48, f"{label}: retained raw rows")
    ids: set[str] = set()
    for row in rows:
        require(isinstance(row, dict), f"{label}: retained row object")
        child_id = row.get("child_id")
        edits = row.get("edits")
        require(isinstance(child_id, str) and child_id not in ids and
                type(edits) is int and edits in EDIT_COUNTS and
                row.get("returncode") == 0 and
                row.get("commit_attempts") == edits and
                row.get("capture_attempted") == edits,
                f"{label}: retained row cardinality")
        ids.add(child_id)
    audit = value.get("offline_audit")
    require(isinstance(audit, dict) and
            audit.get("result") == "all_48_retained_stdout_packets_pass_per_edit_cardinality_and_existing_correctness_oracles" and
            audit.get("matrix_edits") == [1, 2, 8, 32] and
            audit.get("retentions") == ["zero", "default", "release"] and
            audit.get("repetitions") == 2,
            f"{label}: offline audit")


def check_go(value: Any, label: str, plan_path: Path, plan: dict[str, Any]) -> None:
    require(isinstance(value, dict) and value.get("schema") == GO_SCHEMA, f"{label}: schema")
    require(value.get("status") == "authorized" and value.get("retry") is False, f"{label}: authorization/no-retry")
    require(value.get("plan_sha256") == sha256_file(plan_path), f"{label}: plan binding")
    require(value.get("source_scope_sha256") == plan["source"]["scope_sha256"], f"{label}: source binding")
    require(value.get("binary_sha256") == plan["binary"]["sha256"], f"{label}: binary binding")
    require(value.get("source_harness_sha256") == plan["source_harness"]["sha256"], f"{label}: source-harness binding")


def normalized_report(report: dict[str, Any]) -> dict[str, Any]:
    # Child identity lives in the capture manifest row, not in the harness
    # packet.  The v2 packet itself must be byte-for-byte semantically equal.
    return json.loads(json.dumps(report, sort_keys=True))
