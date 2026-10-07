#!/usr/bin/env python3
"""Independently reduce the external 0966 DOC SPRM census packet.

The reducer performs its own report arithmetic rather than importing the
capture runner or the scenario author's checker.  It accepts a packet only
when all 48 terminal receipts are successful, every 0942 correctness gate is
true, every phase bucket conserves into the total bucket, and the forward and
reverse fresh repetitions are equal.  It reports logical parser counters only.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
from pathlib import Path
import re
from typing import Any


PLAN_SCHEMA = "litchi.performance.0966.doc_sprm_census_plan.v1"
REPORT_SCHEMA = "litchi.performance.0966.doc_sprm_census.v1"
REDUCTION_SCHEMA = "litchi.performance.0966.doc_sprm_census_reduction.v1"
CASES = ("docnohf", "docfloat")
EDIT_COUNTS = (1, 2, 8, 32)
RETENTIONS = ("zero", "default_8mib", "release_8mib")
RETENTION_ARG = {"zero": "zero", "default_8mib": "default", "release_8mib": "release"}
PHASES = ("snapshot_open", "edit_open", "staged_replacement", "final_commit", "output_copy", "other")
HISTOGRAM_LABELS = ("0", "1", "2-3", "4-7", "8-15", "16-31", "32-63", "64-127", "128-255", "256+")
ROUTE_SPECS = (("zero", 0), ("default_8mib", 8 * 1024 * 1024), ("release_8mib", 8 * 1024 * 1024))
STAGE_TEXTS = (
    "litchi copy-through baseline replacement text",
    "litchi retained render second replacement text",
)
MAX_EDIT_STAGES = 32
ORACLE_BOOL_FIELDS = (
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
BUCKET_FIELDS = (
    "calls", "successes", "errors", "input_empty", "input_nonempty", "input_bytes_total",
    "input_bytes_min", "input_bytes_max", "output_count_total", "output_capacity_total",
    "output_capacity_min", "output_capacity_max", "logical_vec_capacity_bytes",
)
ADDITIVE_FIELDS = (
    "calls", "successes", "errors", "input_empty", "input_nonempty", "input_bytes_total",
    "output_count_total", "output_capacity_total", "logical_vec_capacity_bytes",
)
HISTOGRAM_FIELDS = ("input_length_histogram", "output_count_histogram", "errors_by_kind")
HEX = set("0123456789abcdef")
ID_RE = re.compile(r"^[A-Za-z0-9_.-]+$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
MAX_JSON_BYTES = 1 << 20


class ReducerError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReducerError(message)


def read_json(path: Path, *, maximum: int = MAX_JSON_BYTES) -> Any:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    require(path.stat().st_size < maximum, f"JSON too large: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ReducerError(f"invalid JSON {path}: {error}") from error


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


def verify_descriptor(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}: path missing")
    path = regular_file(Path(path_value), label)
    require(type(value.get("bytes")) is int and value["bytes"] == path.stat().st_size,
            f"{label}: byte count changed")
    expected = value.get("sha256")
    require(isinstance(expected, str) and len(expected) == 64 and set(expected) <= HEX,
            f"{label}: SHA-256 missing")
    require(sha256_file(path) == expected, f"{label}: SHA-256 changed")
    return path


def source_tree_digest(root: Path) -> tuple[int, str]:
    root = root.resolve()
    require(root.is_dir() and not root.is_symlink(), f"source root missing: {root}")
    entries: list[tuple[str, int, str]] = []
    for current, directories, files in os.walk(root, topdown=True, followlinks=False):
        current_path = Path(current)
        directories.sort()
        files.sort()
        for name in files:
            path = current_path / name
            regular_file(path, "source tree")
            entries.append((path.relative_to(root).as_posix(), path.stat().st_size, sha256_file(path)))
        for name in directories:
            require(not (current_path / name).is_symlink(), f"source tree contains symlink directory: {current_path / name}")
    digest = hashlib.sha256()
    for relative, size, file_hash in entries:
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        digest.update(size.to_bytes(8, "big"))
        digest.update(file_hash.encode("ascii"))
        digest.update(b"\n")
    return len(entries), digest.hexdigest()


def write_exclusive(path: Path, value: object) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(data) < MAX_JSON_BYTES, f"reduction exceeds {MAX_JSON_BYTES} bytes")
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise ReducerError(f"refusing to overwrite retained reduction: {path}") from error


def expected_rows() -> list[dict[str, Any]]:
    cells = [(case, edits, retention) for case in CASES for edits in EDIT_COUNTS for retention in RETENTIONS]
    rows: list[dict[str, Any]] = []
    position = 0
    for repeat, direction, ordered in ((1, "forward", cells), (2, "reverse", list(reversed(cells)))):
        for repeat_position, (case, edits, retention) in enumerate(ordered, start=1):
            position += 1
            rows.append({
                "id": f"sprm-r{repeat:02d}-{case}-e{edits:02d}-{retention}",
                "position": position,
                "repeat": repeat,
                "direction": direction,
                "repeat_position": repeat_position,
                "case": case,
                "edits": edits,
                "retention": retention,
                "retention_arg": RETENTION_ARG[retention],
                "warmups": 0,
                "samples": 1,
                "retry": False,
            })
    return rows


def check_plan(plan_path: Path) -> dict[str, Any]:
    plan = read_json(plan_path)
    require(isinstance(plan, dict) and plan.get("schema") == PLAN_SCHEMA, "plan schema")
    require(plan.get("change") == "0966", "plan change")
    require(isinstance(plan.get("source_revision"), str) and
            REVISION_RE.fullmatch(plan["source_revision"]) is not None,
            "source revision must be 40 lowercase hexadecimal characters")
    source = plan.get("source")
    require(isinstance(source, dict), "source binding")
    root = Path(source.get("root", ""))
    count, digest = source_tree_digest(root)
    require(count == source.get("file_count") and digest == source.get("tree_sha256"), "source tree changed")
    verify_descriptor(source.get("parser"), "parser")
    harness = source.get("harness_files")
    require(isinstance(harness, list) and len(harness) == 8, "harness allowlist")
    for index, item in enumerate(harness):
        verify_descriptor(item, f"harness[{index}]")
    executable = verify_descriptor(plan.get("executable"), "executable")
    require(os.access(executable, os.X_OK), "executable bit missing")
    fixtures = plan.get("fixtures")
    require(isinstance(fixtures, dict) and set(fixtures) == set(CASES), "fixtures")
    for case in CASES:
        verify_descriptor(fixtures[case], f"fixture {case}")
    review = plan.get("review")
    require(isinstance(review, dict) and review.get("role") == "doc_census_review", "review role")
    review_path = verify_descriptor(review.get("document"), "review")
    review_value = read_json(review_path)
    require(isinstance(review_value, dict) and review_value.get("status") in {"pass", "passed"},
            "review status is not pass/passed")
    require(review.get("document_status") == review_value.get("status"), "review status binding")
    verify_descriptor(plan.get("checker"), "checker")
    tools = plan.get("tools")
    require(isinstance(tools, dict), "tools")
    verify_descriptor(tools.get("runner"), "runner")
    verify_descriptor(tools.get("reducer"), "reducer")
    scope = plan.get("claim_scope")
    require(isinstance(scope, dict) and scope.get("timing_claim") is False, "timing claim scope")
    policy = plan.get("run_policy")
    require(isinstance(policy, dict) and policy.get("serial") is True and policy.get("retry") is False,
            "serial/no-retry policy")
    require(policy.get("warmups") == 0 and policy.get("samples") == 1 and policy.get("deadline_seconds") == 300,
            "sample/deadline policy")
    require(policy.get("cpu_affinity") is None, "unexpected CPU affinity")
    rows = plan.get("schedule", {}).get("rows") if isinstance(plan.get("schedule"), dict) else None
    require(rows == expected_rows(), "schedule mismatch")
    require(plan.get("report_schema") == REPORT_SCHEMA, "report schema binding")
    for row in rows:
        require(ID_RE.fullmatch(row["id"]) is not None, "invalid child id")
    return plan


def check_int(value: Any, label: str) -> None:
    require(type(value) is int and value >= 0, f"{label}: invalid nonnegative integer")


def check_array(value: Any, length: int, label: str) -> list[int]:
    require(isinstance(value, list) and len(value) == length, f"{label}: histogram shape")
    for index, item in enumerate(value):
        check_int(item, f"{label}[{index}]")
    return value


def check_optional_int(value: Any, label: str) -> None:
    require(value is None or (type(value) is int and value >= 0), f"{label}: invalid optional integer")


def check_oracle(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label}: oracle missing")
    for field in ORACLE_BOOL_FIELDS:
        require(type(value.get(field)) is bool and value[field] is True,
                f"{label}.{field}: oracle gate failed")
    require(value.get("failure_reasons") == [], f"{label}: oracle failure reasons present")


def check_stages(value: Any, edits: int, retention: str, label: str) -> None:
    require(isinstance(value, list) and len(value) == MAX_EDIT_STAGES,
            f"{label}: stage array shape")
    expected_keys = {"stage", "requested_text", "before", "after", "after_release"}
    for index, stage in enumerate(value):
        require(isinstance(stage, dict) and set(stage) == expected_keys,
                f"{label}[{index}]: stage shape")
        if index < edits:
            require(stage["stage"] == index + 1, f"{label}[{index}]: stage number")
            require(stage["requested_text"] == STAGE_TEXTS[index % 2],
                    f"{label}[{index}]: requested text")
            for field in ("before", "after", "after_release"):
                check_optional_int(stage[field], f"{label}[{index}].{field}")
            if retention == "zero":
                require(stage["before"] is None and stage["after"] is None and
                        stage["after_release"] is None,
                        f"{label}[{index}]: zero-retention stage retained bytes")
            if retention != "release":
                require(stage["after_release"] is None,
                        f"{label}[{index}]: non-release after_release populated")
            else:
                require(stage["after_release"] is None,
                        f"{label}[{index}]: release after_release not cleared")
        else:
            require(stage == {
                "stage": 0,
                "requested_text": "",
                "before": None,
                "after": None,
                "after_release": None,
            }, f"{label}[{index}]: inactive stage is not empty")


def check_route(value: Any, index: int, edits: int, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: route missing")
    retention, ceiling = ROUTE_SPECS[index]
    required = {
        "retention", "retention_ceiling_bytes", "output_sha256", "output_bytes",
        "output_capacity", "semantic_change_count", "final_retained_before_commit",
        "stages", "oracle",
    }
    require(set(value) == required, f"{label}: route field set")
    require(value["retention"] == retention, f"{label}: route retention")
    require(value["retention_ceiling_bytes"] == ceiling, f"{label}: route ceiling")
    output_hash = value["output_sha256"]
    require(isinstance(output_hash, str) and len(output_hash) == 64 and
            set(output_hash) <= HEX, f"{label}: output hash")
    for field in ("output_bytes", "output_capacity", "semantic_change_count"):
        check_int(value[field], f"{label}.{field}")
    require(value["semantic_change_count"] == edits,
            f"{label}: semantic change count does not match edit count")
    check_optional_int(value["final_retained_before_commit"],
                       f"{label}.final_retained_before_commit")
    if retention in {"zero", "release"}:
        require(value["final_retained_before_commit"] is None,
                f"{label}: cleared retention route retained bytes")
    check_stages(value["stages"], edits, retention, f"{label}.stages")
    check_oracle(value["oracle"], f"{label}.oracle")
    return value


def check_bucket(bucket: Any, label: str, sprm_size: int) -> None:
    require(isinstance(bucket, dict), f"{label}: bucket missing")
    for field in BUCKET_FIELDS:
        check_int(bucket.get(field), f"{label}.{field}")
    input_histogram = check_array(bucket.get("input_length_histogram"), 10, f"{label}.input_length_histogram")
    output_histogram = check_array(bucket.get("output_count_histogram"), 10, f"{label}.output_count_histogram")
    errors = check_array(bucket.get("errors_by_kind"), 5, f"{label}.errors_by_kind")
    require(bucket["calls"] == bucket["successes"] + bucket["errors"], f"{label}: call conservation")
    require(bucket["calls"] == bucket["input_empty"] + bucket["input_nonempty"], f"{label}: empty/nonempty conservation")
    require(sum(input_histogram) == bucket["calls"], f"{label}: input histogram conservation")
    require(sum(output_histogram) == bucket["successes"], f"{label}: output histogram conservation")
    require(sum(errors) == bucket["errors"], f"{label}: error histogram conservation")
    require(bucket["logical_vec_capacity_bytes"] == bucket["output_capacity_total"] * sprm_size,
            f"{label}: logical Vec capacity bytes")
    if bucket["calls"] == 0:
        require(bucket["input_bytes_min"] == 0 and bucket["input_bytes_max"] == 0,
                f"{label}: empty input bounds")
    else:
        require(bucket["input_bytes_min"] <= bucket["input_bytes_max"], f"{label}: input bounds")
    if bucket["successes"] == 0:
        require(bucket["output_capacity_min"] == 0 and bucket["output_capacity_max"] == 0,
                f"{label}: empty output bounds")
    else:
        require(bucket["output_capacity_min"] <= bucket["output_capacity_max"], f"{label}: output bounds")


def census_from_report(report: dict[str, Any], path: Path) -> dict[str, Any]:
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1, f"{path}: sample count")
    sample = samples[0]
    require(isinstance(sample, dict), f"{path}: sample object")
    census = sample.get("sprm_census")
    require(isinstance(census, dict) and census.get("schema_version") == 1, f"{path}: census schema")
    require(census.get("saturated") is False, f"{path}: saturated counters")
    require(census.get("histogram_labels") == list(HISTOGRAM_LABELS), f"{path}: histogram labels")
    sprm_size = census.get("sprm_size_bytes")
    check_int(sprm_size, f"{path}: SPRM size")
    require(sprm_size > 0, f"{path}: SPRM size is zero")
    buckets: dict[str, dict[str, Any]] = {}
    for phase in PHASES:
        bucket = census.get(phase)
        check_bucket(bucket, f"{path}:{phase}", sprm_size)
        buckets[phase] = bucket
    total = census.get("total")
    check_bucket(total, f"{path}:total", sprm_size)
    for field in ADDITIVE_FIELDS:
        require(sum(bucket[field] for bucket in buckets.values()) == total[field],
                f"{path}: phase conservation for {field}")
    for field, length in (("input_length_histogram", 10), ("output_count_histogram", 10), ("errors_by_kind", 5)):
        expected = [sum(bucket[field][index] for bucket in buckets.values()) for index in range(length)]
        require(expected == total[field], f"{path}: phase conservation for {field}")
    input_buckets = [bucket for bucket in buckets.values() if bucket["calls"]]
    output_buckets = [bucket for bucket in buckets.values() if bucket["successes"]]
    if total["calls"]:
        require(total["input_bytes_min"] == min(bucket["input_bytes_min"] for bucket in input_buckets) and
                total["input_bytes_max"] == max(bucket["input_bytes_max"] for bucket in input_buckets),
                f"{path}: phase input bounds")
    if total["successes"]:
        require(total["output_capacity_min"] == min(bucket["output_capacity_min"] for bucket in output_buckets) and
                total["output_capacity_max"] == max(bucket["output_capacity_max"] for bucket in output_buckets),
                f"{path}: phase output bounds")
    require(buckets["other"]["calls"] == 0, f"{path}: unclassified phase calls")
    return census


def check_report(report: Any, row: dict[str, Any], plan: dict[str, Any], path: Path) -> dict[str, Any]:
    require(isinstance(report, dict), f"{path}: report object")
    require(report.get("schema") == REPORT_SCHEMA and report.get("schema_version") == 1,
            f"{path}: report schema")
    require(report.get("change") == "0966", f"{path}: report change")
    require(report.get("timing_claim") is False and report.get("allocator_timing_claim") is False,
            f"{path}: timing claim")
    require(report.get("warmups") == 0 and report.get("samples_requested") == 1,
            f"{path}: warmup/sample policy")
    require(report.get("case") == row["case"] and report.get("edits") == row["edits"] and
            report.get("retention") == row["retention"], f"{path}: cell identity")
    require(report.get("selected_paragraph") == 0, f"{path}: selected paragraph")
    require(report.get("alternating_texts") == list(STAGE_TEXTS), f"{path}: alternating texts")
    expected_ceiling = dict(ROUTE_SPECS)[row["retention"]]
    require(report.get("retention_ceiling_bytes") == expected_ceiling,
            f"{path}: selected retention ceiling")
    require(report.get("allocator_instrumented") is False,
            f"{path}: ordinary census lane has allocator instrumentation")
    fixture = plan["fixtures"][row["case"]]
    require(report.get("source_sha256") == fixture["sha256"] and report.get("source_bytes") == fixture["bytes"],
            f"{path}: fixture identity")
    require(report.get("exact_output_equal_reference") is True, f"{path}: output parity")
    for field in (
        "selected_paragraph", "alternating_texts", "retention_ceiling_bytes", "allocator_instrumented",
        "output_sha256", "output_bytes", "output_capacity",
        "reference_output_sha256", "reference_output_bytes", "route_outputs",
    ):
        require(field in report, f"{path}: missing report field {field}")
    correctness = report.get("correctness")
    require(isinstance(correctness, dict) and set(correctness) == set(GATE_NAMES), f"{path}: correctness gate set")
    require(all(type(correctness[name]) is bool and correctness[name] for name in GATE_NAMES),
            f"{path}: correctness gate failed")
    require(report.get("elapsed_ns") == [] and report.get("allocation") is None,
            f"{path}: timing/allocation payload")
    require(isinstance(report.get("route_outputs"), list) and len(report["route_outputs"]) == 3,
            f"{path}: route output shape")
    routes = [check_route(route, index, row["edits"], f"{path}:route[{index}]")
              for index, route in enumerate(report["route_outputs"])]
    require(routes[0]["output_sha256"] == report["reference_output_sha256"] and
            routes[0]["output_bytes"] == report["reference_output_bytes"],
            f"{path}: reference route identity")
    require(all(route["output_sha256"] == routes[0]["output_sha256"] and
                route["output_bytes"] == routes[0]["output_bytes"] for route in routes),
            f"{path}: cross-route output identity")
    selected_index = {retention: index for index, (retention, _) in enumerate(ROUTE_SPECS)}[row["retention"]]
    selected_route = routes[selected_index]
    for field in ("output_sha256", "output_bytes", "output_capacity"):
        require(report.get(field) == selected_route[field],
                f"{path}: selected top-level {field} mismatch")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1, f"{path}: sample count")
    sample = samples[0]
    require(isinstance(sample, dict) and sample.get("index") == 0, f"{path}: sample identity")
    require(sample.get("elapsed_ns") is None and sample.get("allocation") is None,
            f"{path}: sample timing/allocation payload")
    for field in ("output_sha256", "output_bytes", "output_capacity"):
        require(sample.get(field) == report.get(field), f"{path}: sample/root {field} mismatch")
    require(sample.get("semantic_change_count") == row["edits"],
            f"{path}: sample semantic change count")
    require(sample.get("exact_output_equal_reference") is True and
            sample.get("final_paragraph_semantic_readback") is True and
            sample.get("final_public_reader_reopen") is True and
            sample.get("retained_capacities_within_ceiling") is True and
            sample.get("allocator_retained_matches_output_capacity") is True,
            f"{path}: sample correctness")
    check_oracle(sample.get("oracle"), f"{path}:sample.oracle")
    check_stages(sample.get("stages"), row["edits"], row["retention"], f"{path}:sample.stages")
    require(sample.get("stages") == selected_route["stages"],
            f"{path}: selected sample/route stages differ")
    require(sample.get("oracle") == selected_route["oracle"],
            f"{path}: selected sample/route oracle differs")
    census = census_from_report(report, path)
    return {"report": report, "census": census}


def canonical_report(report: dict[str, Any]) -> dict[str, Any]:
    sample = copy.deepcopy(report["samples"][0])
    sample.pop("elapsed_ns", None)
    sample.pop("index", None)
    return {
        "source_sha256": report["source_sha256"],
        "source_bytes": report["source_bytes"],
        "selected_paragraph": report["selected_paragraph"],
        "edits": report["edits"],
        "alternating_texts": report["alternating_texts"],
        "retention": report["retention"],
        "retention_ceiling_bytes": report["retention_ceiling_bytes"],
        "timing_claim": report["timing_claim"],
        "allocator_instrumented": report["allocator_instrumented"],
        "allocator_timing_claim": report["allocator_timing_claim"],
        "warmups": report["warmups"],
        "samples_requested": report["samples_requested"],
        "output_sha256": report["output_sha256"],
        "output_bytes": report["output_bytes"],
        "output_capacity": report["output_capacity"],
        "reference_output_sha256": report["reference_output_sha256"],
        "reference_output_bytes": report["reference_output_bytes"],
        "exact_output_equal_reference": report["exact_output_equal_reference"],
        "route_outputs": report["route_outputs"],
        "correctness": report["correctness"],
        "sample": sample,
    }


def canonical_hash(value: object) -> str:
    data = json.dumps(value, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(data).hexdigest()


def empty_success_calls(bucket: dict[str, Any]) -> int:
    # The sealed parser returns Ok(Vec::new()) for exactly empty input.  The
    # report exposes that input class as input_empty; malformed nonempty calls
    # are separately counted in errors and errors_by_kind.
    require(bucket["input_empty"] <= bucket["successes"], "empty input exceeds successful calls")
    return bucket["input_empty"]


def reduce_packet(plan_path: Path, capture_root: Path, output_path: Path) -> None:
    plan = check_plan(plan_path)
    capture_root = capture_root.resolve()
    require(capture_root == Path(plan["capture_root"]).resolve(), "capture root is not bound by plan")
    run_start_path = capture_root / "run-start.json"
    run_terminal_path = capture_root / "run-terminal.json"
    run_start = read_json(run_start_path)
    run_terminal = read_json(run_terminal_path)
    require(run_start.get("schema") == "litchi.performance.0966.run_start.v1", "run start schema")
    require(run_terminal.get("schema") == "litchi.performance.0966.run_terminal.v1" and
            run_terminal.get("status") == "completed" and run_terminal.get("children") == 48,
            "run terminal is incomplete")
    require(run_terminal.get("failures") == [] and run_terminal.get("retries") == 0,
            "failed child or retry present")
    require(run_start.get("plan", {}).get("sha256") == sha256_file(plan_path), "run start plan binding")
    require(run_start.get("source_tree_sha256") == plan["source"]["tree_sha256"], "run start source binding")
    require(run_start.get("executable", {}).get("sha256") == plan["executable"]["sha256"], "run start executable binding")

    rows = plan["schedule"]["rows"]
    expected_ids = {row["id"] for row in rows}
    actual_dirs = {path.name for path in capture_root.iterdir() if path.is_dir()}
    require(actual_dirs == expected_ids, "capture has missing or extra child directories")
    grouped: dict[tuple[str, int, str], list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]]] = {}
    for row in rows:
        child = capture_root / row["id"]
        start = read_json(child / "start.json")
        terminal = read_json(child / "terminal.json")
        require(start.get("schema") == "litchi.performance.0966.child_start.v1" and
                start.get("child_id") == row["id"] and start.get("row") == row,
                f"{row['id']}: start receipt binding")
        require(start.get("plan_sha256") == sha256_file(plan_path), f"{row['id']}: plan hash")
        require(terminal.get("schema") == "litchi.performance.0966.child_terminal.v1" and
                terminal.get("child_id") == row["id"] and terminal.get("returncode") == 0 and
                terminal.get("timed_out") is False and terminal.get("launch_error") is None,
                f"{row['id']}: terminal failure")
        stdout_path = verify_descriptor(terminal.get("stdout"), f"{row['id']}: stdout")
        stderr_path = verify_descriptor(terminal.get("stderr"), f"{row['id']}: stderr")
        require(stdout_path == (child / "stdout.raw").resolve() and stderr_path == (child / "stderr.raw").resolve(),
                f"{row['id']}: raw stream path")
        require(stdout_path.stat().st_size <= plan["run_policy"]["max_stdout_bytes"] and
                stderr_path.stat().st_size <= plan["run_policy"]["max_stderr_bytes"],
                f"{row['id']}: raw stream bound")
        parse = read_json(child / "parse.json")
        require(parse.get("schema") == "litchi.performance.0966.child_parse.v1" and
                parse.get("status") == "parsed" and parse.get("child_id") == row["id"],
                f"{row['id']}: parse receipt")
        report_path = verify_descriptor(parse.get("report"), f"{row['id']}: report")
        require(report_path == (child / "report.json").resolve(), f"{row['id']}: report path")
        report = read_json(report_path)
        try:
            raw_report = json.loads(stdout_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError) as error:
            raise ReducerError(f"{row['id']}: raw stdout is not JSON: {error}") from error
        require(raw_report == report, f"{row['id']}: raw stdout/report mismatch")
        checked = check_report(report, row, plan, report_path)
        grouped.setdefault((row["case"], row["edits"], row["retention"]), []).append((row, report, checked["census"]))

    expected_cells = {(case, edits, retention) for case in CASES for edits in EDIT_COUNTS for retention in RETENTIONS}
    require(set(grouped) == expected_cells and all(len(items) == 2 for items in grouped.values()),
            "the 24 cells do not each have exactly two reports")
    phase_totals = {
        phase: {"calls": 0, "successes": 0, "errors": 0, "empty_success_calls": 0,
                "logical_vec_capacity_bytes": 0}
        for phase in PHASES
    }
    cell_summaries: list[dict[str, Any]] = []
    for key in sorted(grouped):
        items = grouped[key]
        items.sort(key=lambda item: item[0]["repeat"])
        first_row, first_report, first_census = items[0]
        second_row, second_report, second_census = items[1]
        require(first_row["repeat"] == 1 and first_row["direction"] == "forward" and
                second_row["repeat"] == 2 and second_row["direction"] == "reverse",
                f"{key}: repeat direction/order")
        first_canonical = canonical_report(first_report)
        second_canonical = canonical_report(second_report)
        require(first_canonical == second_canonical, f"{key}: forward/reverse reports differ")
        cell_summaries.append({
            "case": key[0], "edits": key[1], "retention": key[2],
            "children": [first_row["id"], second_row["id"]],
            "canonical_sha256": canonical_hash(first_canonical),
            "calls": first_census["total"]["calls"],
            "successes": first_census["total"]["successes"],
            "errors": first_census["total"]["errors"],
        })
        for phase in PHASES:
            bucket = first_census[phase]
            phase_totals[phase]["calls"] += bucket["calls"]
            phase_totals[phase]["successes"] += bucket["successes"]
            phase_totals[phase]["errors"] += bucket["errors"]
            phase_totals[phase]["empty_success_calls"] += empty_success_calls(bucket)
            phase_totals[phase]["logical_vec_capacity_bytes"] += bucket["logical_vec_capacity_bytes"]
    empty_rank = sorted(
        ({"phase": phase, "calls": values["empty_success_calls"]} for phase, values in phase_totals.items()),
        key=lambda item: (-item["calls"], PHASES.index(item["phase"])),
    )
    capacity_rank = sorted(
        ({"phase": phase, "bytes": values["logical_vec_capacity_bytes"]}
         for phase, values in phase_totals.items()),
        key=lambda item: (-item["bytes"], PHASES.index(item["phase"])),
    )
    result = {
        "schema": REDUCTION_SCHEMA,
        "schema_version": 1,
        "change": "0966",
        "status": "diagnostic_only_pass",
        "diagnostic_only": True,
        "timing_claim": False,
        "allocator_total_claim": False,
        "elapsed_time_claim": False,
        "removable_work_claim": False,
        "plan": descriptor(plan_path, "plan"),
        "run": {"path": str(capture_root), "children": 48},
        "children": 48,
        "cells": 24,
        "repeat_equal": True,
        "conservation": True,
        "correctness_gates_all_true": True,
        "aggregation": "one representative per repeat-equal cell",
        "phase_totals": phase_totals,
        "rankings": {
            "empty_success_calls_by_phase": empty_rank,
            "logical_vec_capacity_bytes_by_phase": capacity_rank,
        },
        "cell_summaries": cell_summaries,
        "interpretation_boundary": "Logical parser census only; capacity fields are Vec capacity bytes, not allocator bytes, RSS, elapsed time, or removable-work evidence.",
    }
    write_exclusive(output_path, result)
    print(json.dumps({"schema": REDUCTION_SCHEMA, "status": result["status"], "children": 48, "cells": 24}, sort_keys=True))


def self_test() -> None:
    rows = expected_rows()
    require(len(rows) == 48, "self-test schedule count")
    require(rows[0]["direction"] == "forward" and rows[23]["direction"] == "forward" and
            rows[24]["direction"] == "reverse" and rows[47]["direction"] == "reverse",
            "self-test repeat direction")
    cells = {(row["case"], row["edits"], row["retention"]) for row in rows}
    require(len(cells) == 24, "self-test cell count")
    require(all(sum(1 for row in rows if (row["case"], row["edits"], row["retention"]) == cell) == 2 for cell in cells),
            "self-test repeat multiplicity")
    valid = {
        "calls": 2, "successes": 1, "errors": 1, "input_empty": 1, "input_nonempty": 1,
        "input_bytes_total": 3, "input_bytes_min": 0, "input_bytes_max": 3,
        "input_length_histogram": [1, 0, 0, 0, 0, 0, 0, 0, 0, 1],
        "output_count_total": 0, "output_count_histogram": [1, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        "output_capacity_total": 0, "output_capacity_min": 0, "output_capacity_max": 0,
        "logical_vec_capacity_bytes": 0, "errors_by_kind": [1, 0, 0, 0, 0],
    }
    check_bucket(valid, "self-test-valid", 48)
    invalid = dict(valid)
    invalid["errors"] = 0
    try:
        check_bucket(invalid, "self-test-invalid", 48)
    except ReducerError:
        print(json.dumps({"self_test": "passed", "invalid_control_rejected": True}, sort_keys=True))
        return
    raise ReducerError("self-test invalid control was accepted")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--capture", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.self_test:
        require(args.plan is None and args.capture is None and args.output is None, "self-test takes no packet arguments")
        self_test()
        return
    require(args.plan is not None and args.capture is not None and args.output is not None,
            "--plan, --capture, and --output are required")
    reduce_packet(args.plan.resolve(), args.capture.resolve(), args.output.resolve())


if __name__ == "__main__":
    try:
        main()
    except ReducerError as error:
        raise SystemExit(f"FAIL: {error}")
