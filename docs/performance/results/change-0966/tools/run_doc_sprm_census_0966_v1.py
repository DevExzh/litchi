#!/usr/bin/env python3
"""Run the sealed 0966 DOC SPRM census serially.

The default operation is ``--check-plan`` and does not launch anything.
``--capture`` requires a separate root-owned GO receipt.  It launches one
fresh process per row, never retries a row, writes raw stdout/stderr and the
terminal receipt before parsing JSON, and retains failed children.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
import time
from typing import Any


PLAN_SCHEMA = "litchi.performance.0966.doc_sprm_census_plan.v1"
REPORT_SCHEMA = "litchi.performance.0966.doc_sprm_census.v1"
GO_SCHEMA = "litchi.performance.0966.doc_sprm_census_go.v1"
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
HEX = set("0123456789abcdef")
MAX_JSON_BYTES = 1 << 20
ID_RE = re.compile(r"^[A-Za-z0-9_.-]+$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")


class RunnerError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RunnerError(message)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def read_json(path: Path, *, maximum: int = MAX_JSON_BYTES) -> Any:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    require(path.stat().st_size < maximum, f"JSON is too large: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"invalid JSON {path}: {error}") from error


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
    require(isinstance(expected, str) and len(expected) == 64 and set(expected) <= HEX,
            f"{label}: SHA-256 missing")
    require(sha256_file(path) == expected, f"{label}: SHA-256 changed")
    if executable:
        require(os.access(path, os.X_OK), f"{label}: executable bit is missing")
    return path


def source_tree_digest(root: Path) -> tuple[int, str]:
    root = root.resolve()
    require(root.is_dir() and not root.is_symlink(), f"source root is not a directory: {root}")
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
    require(len(data) < MAX_JSON_BYTES, f"JSON output exceeds {MAX_JSON_BYTES} bytes: {path}")
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise RunnerError(f"refusing to overwrite retained file: {path}") from error


def write_bytes_exclusive(path: Path, data: bytes) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise RunnerError(f"refusing to overwrite retained file: {path}") from error


def check_int(value: Any, label: str, *, minimum: int = 0) -> None:
    require(type(value) is int and value >= minimum, f"{label}: invalid integer")


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
            require(stage["after_release"] is None,
                    f"{label}[{index}]: after_release must be cleared/absent")
        else:
            require(stage == {
                "stage": 0,
                "requested_text": "",
                "before": None,
                "after": None,
                "after_release": None,
            }, f"{label}[{index}]: inactive stage is not empty")


def check_bucket(bucket: Any, label: str, sprm_size: int) -> None:
    require(isinstance(bucket, dict), f"{label}: bucket missing")
    fields = (
        "calls", "successes", "errors", "input_empty", "input_nonempty", "input_bytes_total",
        "input_bytes_min", "input_bytes_max", "output_count_total", "output_capacity_total",
        "output_capacity_min", "output_capacity_max", "logical_vec_capacity_bytes",
    )
    for field in fields:
        check_int(bucket.get(field), f"{label}.{field}")
    require(isinstance(bucket.get("input_length_histogram"), list) and
            len(bucket["input_length_histogram"]) == 10,
            f"{label}: input histogram shape")
    require(isinstance(bucket.get("output_count_histogram"), list) and
            len(bucket["output_count_histogram"]) == 10,
            f"{label}: output histogram shape")
    require(isinstance(bucket.get("errors_by_kind"), list) and
            len(bucket["errors_by_kind"]) == 5,
            f"{label}: error histogram shape")
    for name in ("input_length_histogram", "output_count_histogram", "errors_by_kind"):
        for index, item in enumerate(bucket[name]):
            check_int(item, f"{label}.{name}[{index}]")
    require(bucket["calls"] == bucket["successes"] + bucket["errors"],
            f"{label}: call conservation")
    require(bucket["calls"] == bucket["input_empty"] + bucket["input_nonempty"],
            f"{label}: empty/nonempty conservation")
    require(sum(bucket["input_length_histogram"]) == bucket["calls"],
            f"{label}: input histogram conservation")
    require(sum(bucket["output_count_histogram"]) == bucket["successes"],
            f"{label}: output histogram conservation")
    require(sum(bucket["errors_by_kind"]) == bucket["errors"],
            f"{label}: error histogram conservation")
    require(bucket["logical_vec_capacity_bytes"] == bucket["output_capacity_total"] * sprm_size,
            f"{label}: logical capacity equation")
    if bucket["calls"] == 0:
        require(bucket["input_bytes_min"] == 0 and bucket["input_bytes_max"] == 0,
                f"{label}: empty input bounds")
    else:
        require(bucket["input_bytes_min"] <= bucket["input_bytes_max"],
                f"{label}: input bounds")
    if bucket["successes"] == 0:
        require(bucket["output_capacity_min"] == 0 and bucket["output_capacity_max"] == 0,
                f"{label}: empty output bounds")
    else:
        require(bucket["output_capacity_min"] <= bucket["output_capacity_max"],
                f"{label}: output bounds")


def check_census(value: Any, label: str) -> None:
    require(isinstance(value, dict) and value.get("schema_version") == 1,
            f"{label}: census schema")
    require(value.get("saturated") is False, f"{label}: saturated counters")
    require(value.get("histogram_labels") == list(HISTOGRAM_LABELS),
            f"{label}: histogram labels")
    sprm_size = value.get("sprm_size_bytes")
    check_int(sprm_size, f"{label}: SPRM size")
    require(sprm_size > 0, f"{label}: SPRM size is zero")
    buckets = {}
    for phase in PHASES:
        check_bucket(value.get(phase), f"{label}:{phase}", sprm_size)
        buckets[phase] = value[phase]
    check_bucket(value.get("total"), f"{label}:total", sprm_size)
    total = value["total"]
    for field in (
        "calls", "successes", "errors", "input_empty", "input_nonempty", "input_bytes_total",
        "output_count_total", "output_capacity_total", "logical_vec_capacity_bytes",
    ):
        require(sum(bucket[field] for bucket in buckets.values()) == total[field],
                f"{label}: phase conservation for {field}")
    for field, length in (("input_length_histogram", 10), ("output_count_histogram", 10),
                          ("errors_by_kind", 5)):
        require([sum(bucket[field][index] for bucket in buckets.values())
                 for index in range(length)] == total[field],
                f"{label}: phase conservation for {field}")
    input_buckets = [bucket for bucket in buckets.values() if bucket["calls"]]
    output_buckets = [bucket for bucket in buckets.values() if bucket["successes"]]
    if total["calls"]:
        require(total["input_bytes_min"] == min(bucket["input_bytes_min"] for bucket in input_buckets) and
                total["input_bytes_max"] == max(bucket["input_bytes_max"] for bucket in input_buckets),
                f"{label}: phase input bounds")
    if total["successes"]:
        require(total["output_capacity_min"] == min(bucket["output_capacity_min"] for bucket in output_buckets) and
                total["output_capacity_max"] == max(bucket["output_capacity_max"] for bucket in output_buckets),
                f"{label}: phase output bounds")
    require(buckets["other"]["calls"] == 0, f"{label}: unclassified phase calls")


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
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema mismatch")
    require(plan.get("change") == "0966", "plan change mismatch")
    require(isinstance(plan.get("source_revision"), str) and
            REVISION_RE.fullmatch(plan["source_revision"]) is not None,
            "source revision must be 40 lowercase hexadecimal characters")
    require(plan.get("status") in {"prospective_only_no_build_no_capture", "sealed_ready_for_go"},
            "plan is not a prospective/sealed plan")
    scope = plan.get("claim_scope")
    require(isinstance(scope, dict), "claim scope missing")
    for key in ("timing_claim", "allocator_total_claim", "elapsed_time_claim", "removable_work_claim"):
        require(scope.get(key) is False, f"claim scope {key} must be false")
    source = plan.get("source")
    require(isinstance(source, dict), "source binding missing")
    source_root = Path(source.get("root", ""))
    require(source_root.is_dir() and not source_root.is_symlink(), "source root missing")
    file_count, tree_hash = source_tree_digest(source_root)
    require(file_count == source.get("file_count") and tree_hash == source.get("tree_sha256"),
            "source tree binding changed")
    verify_descriptor(source.get("parser"), "SPRM parser")
    harness_files = source.get("harness_files")
    require(isinstance(harness_files, list) and len(harness_files) == 8, "harness allowlist shape")
    for index, item in enumerate(harness_files):
        verify_descriptor(item, f"harness[{index}]")
    executable = verify_descriptor(plan.get("executable"), "census executable", executable=True)
    require(plan.get("argv_template", [None])[0] == str(executable), "argv does not bind executable")
    fixtures = plan.get("fixtures")
    require(isinstance(fixtures, dict) and set(fixtures) == set(CASES), "fixture binding shape")
    for case in CASES:
        verify_descriptor(fixtures[case], f"fixture {case}")
    review = plan.get("review")
    require(isinstance(review, dict) and review.get("role") == "doc_census_review", "review role missing")
    review_path = verify_descriptor(review.get("document"), "doc_census_review")
    review_value = read_json(review_path)
    require(isinstance(review_value, dict) and review_value.get("status") in {"pass", "passed"},
            "doc_census_review has not passed")
    require(review.get("document_status") == review_value.get("status"),
            "review status binding changed")
    verify_descriptor(plan.get("checker"), "independent checker")
    tools = plan.get("tools")
    require(isinstance(tools, dict), "tool bindings missing")
    for label in ("runner", "reducer"):
        verify_descriptor(tools.get(label), label)
    policy = plan.get("run_policy")
    require(isinstance(policy, dict), "run policy missing")
    require(policy.get("serial") is True and policy.get("retry") is False, "run must be serial/no-retry")
    require(policy.get("warmups") == 0 and policy.get("samples") == 1, "warmup/sample policy mismatch")
    require(policy.get("deadline_seconds") == 300, "deadline policy mismatch")
    require(policy.get("terminal_before_parse") is True and policy.get("raw_stdout_stderr_retained") is True,
            "receipt policy mismatch")
    require(policy.get("cpu_affinity") is None, "0966 must not claim CPU affinity")
    for key in ("max_stdout_bytes", "max_stderr_bytes"):
        require(policy.get(key) == 1 << 20, f"{key} policy mismatch")
    environment = plan.get("environment")
    require(environment == {
        "PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
        "RUST_BACKTRACE": "0", "PAGER": "cat",
    }, "environment is not the frozen minimal environment")
    rows = plan.get("schedule", {}).get("rows") if isinstance(plan.get("schedule"), dict) else None
    require(rows == expected_rows(), "schedule rows differ from the 48-row forward/reverse matrix")
    require(plan.get("schedule", {}).get("children") == 48 and plan.get("schedule", {}).get("cells") == 24,
            "schedule count mismatch")
    capture_root = Path(plan.get("capture_root", ""))
    require(capture_root.is_absolute(), "capture root must be absolute")
    for row in rows:
        require(ID_RE.fullmatch(row["id"]) is not None, f"invalid child id: {row['id']}")
    require(plan.get("report_schema") == REPORT_SCHEMA, "report schema binding mismatch")
    return plan


def verify_go(go_path: Path, plan_path: Path, plan: dict[str, Any]) -> dict[str, Any]:
    go = read_json(go_path)
    require(isinstance(go, dict) and go.get("schema") == GO_SCHEMA, "GO schema mismatch")
    require(go.get("status") == "authorized", "root GO is not authorized")
    require(go.get("plan_sha256") == sha256_file(plan_path), "GO does not bind this plan")
    require(go.get("source_tree_sha256") == plan["source"]["tree_sha256"], "GO source binding mismatch")
    require(go.get("executable_sha256") == plan["executable"]["sha256"], "GO executable binding mismatch")
    return go


def fixture_path(plan: dict[str, Any], case: str) -> Path:
    descriptor_value = plan["fixtures"][case]
    return verify_descriptor(descriptor_value, f"fixture {case}")


def basic_report_check(report: Any, row: dict[str, Any], plan: dict[str, Any]) -> None:
    require(isinstance(report, dict), "report is not an object")
    require(report.get("schema") == REPORT_SCHEMA and report.get("schema_version") == 1 and
            report.get("change") == "0966", "report schema/change")
    require(report.get("timing_claim") is False and report.get("allocator_timing_claim") is False,
            "report makes a timing claim")
    require(report.get("warmups") == 0 and report.get("samples_requested") == 1, "report sample policy")
    require(report.get("case") == row["case"] and report.get("edits") == row["edits"], "report cell identity")
    require(report.get("retention") == row["retention"], "report retention identity")
    require(report.get("selected_paragraph") == 0 and
            report.get("alternating_texts") == list(STAGE_TEXTS), "report lifecycle identity")
    require(report.get("retention_ceiling_bytes") == dict(ROUTE_SPECS)[row["retention"]],
            "report retention ceiling")
    require(report.get("allocator_instrumented") is False,
            "ordinary census lane has allocator instrumentation")
    fixture = plan["fixtures"][row["case"]]
    require(report.get("source_sha256") == fixture["sha256"] and report.get("source_bytes") == fixture["bytes"],
            "report fixture identity")
    require(report.get("exact_output_equal_reference") is True, "report output parity failed")
    correctness = report.get("correctness")
    require(isinstance(correctness, dict) and set(correctness) == set(GATE_NAMES) and
            all(type(correctness[name]) is bool and correctness[name] for name in GATE_NAMES),
            "report correctness gate failed")
    require(report.get("elapsed_ns") == [] and report.get("allocation") is None,
            "report timing/allocation payload")
    route_values = report.get("route_outputs")
    require(isinstance(route_values, list) and len(route_values) == 3, "report route output shape")
    routes = [check_route(route, index, row["edits"], f"route[{index}]")
              for index, route in enumerate(route_values)]
    require(routes[0]["output_sha256"] == report.get("reference_output_sha256") and
            routes[0]["output_bytes"] == report.get("reference_output_bytes"),
            "report reference route identity")
    require(all(route["output_sha256"] == routes[0]["output_sha256"] and
                route["output_bytes"] == routes[0]["output_bytes"] for route in routes),
            "report cross-route output identity")
    selected_index = {retention: index for index, (retention, _) in enumerate(ROUTE_SPECS)}[row["retention"]]
    selected_route = routes[selected_index]
    for field in ("output_sha256", "output_bytes", "output_capacity"):
        require(report.get(field) == selected_route[field], f"report selected {field}")
    require(isinstance(report.get("samples"), list) and len(report["samples"]) == 1, "report sample count")
    sample = report["samples"][0]
    require(isinstance(sample, dict) and sample.get("index") == 0 and
            sample.get("elapsed_ns") is None and sample.get("allocation") is None,
            "report sample identity/timing payload")
    for field in ("output_sha256", "output_bytes", "output_capacity"):
        require(sample.get(field) == report.get(field), f"report sample/root {field}")
    require(sample.get("semantic_change_count") == row["edits"],
            "report sample semantic change count")
    require(sample.get("exact_output_equal_reference") is True and
            sample.get("final_paragraph_semantic_readback") is True and
            sample.get("final_public_reader_reopen") is True and
            sample.get("retained_capacities_within_ceiling") is True and
            sample.get("allocator_retained_matches_output_capacity") is True,
            "report sample correctness")
    check_oracle(sample.get("oracle"), "sample.oracle")
    check_stages(sample.get("stages"), row["edits"], row["retention"], "sample.stages")
    require(sample.get("stages") == selected_route["stages"] and
            sample.get("oracle") == selected_route["oracle"],
            "report selected sample/route identity")
    check_census(sample.get("sprm_census"), "sample.sprm_census")


def run_child(plan: dict[str, Any], plan_path: Path, row: dict[str, Any], capture_root: Path) -> tuple[bool, str]:
    child_root = capture_root / row["id"]
    require(not child_root.exists(), f"child already retained: {child_root}")
    child_root.mkdir(parents=True, exist_ok=False)
    binary_path = Path(plan["executable"]["path"])
    input_path = fixture_path(plan, row["case"])
    argv = [
        str(binary_path), "--case", row["case"], "--input", str(input_path),
        "--edits", str(row["edits"]), "--retention", row["retention_arg"],
        "--warmup", "0", "--samples", "1",
    ]
    start = {
        "schema": "litchi.performance.0966.child_start.v1",
        "status": "started",
        "child_id": row["id"],
        "row": row,
        "plan_sha256": sha256_file(plan_path),
        "source_tree_sha256": plan["source"]["tree_sha256"],
        "executable": plan["executable"],
        "fixture": plan["fixtures"][row["case"]],
        "argv": argv,
        "cwd": plan["source"]["root"],
        "environment": plan["environment"],
        "started_utc": utc_now(),
    }
    write_exclusive(child_root / "start.json", start)
    stdout = b""
    stderr = b""
    returncode: int | None = None
    timed_out = False
    launch_error: str | None = None
    process: subprocess.Popen[bytes] | None = None
    try:
        process = subprocess.Popen(
            argv,
            cwd=plan["source"]["root"],
            env=dict(plan["environment"]),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            start_new_session=True,
        )
        try:
            stdout, stderr = process.communicate(timeout=plan["run_policy"]["deadline_seconds"])
        except subprocess.TimeoutExpired:
            timed_out = True
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
            stdout, stderr = process.communicate()
        returncode = process.returncode
    except (OSError, ValueError) as error:
        launch_error = str(error)
        if process is not None:
            returncode = process.returncode
    write_bytes_exclusive(child_root / "stdout.raw", stdout)
    write_bytes_exclusive(child_root / "stderr.raw", stderr)
    stdout_descriptor = descriptor(child_root / "stdout.raw", "raw stdout")
    stderr_descriptor = descriptor(child_root / "stderr.raw", "raw stderr")
    terminal = {
        "schema": "litchi.performance.0966.child_terminal.v1",
        "status": "terminal",
        "child_id": row["id"],
        "returncode": returncode,
        "timed_out": timed_out,
        "launch_error": launch_error,
        "stdout": stdout_descriptor,
        "stderr": stderr_descriptor,
        "ended_utc": utc_now(),
    }
    # This is deliberately before any JSON parsing.
    write_exclusive(child_root / "terminal.json", terminal)
    problems: list[str] = []
    if returncode != 0:
        problems.append(f"returncode={returncode}")
    if timed_out:
        problems.append("timeout")
    if launch_error:
        problems.append(f"launch_error={launch_error}")
    if len(stdout) > plan["run_policy"]["max_stdout_bytes"]:
        problems.append("stdout exceeds bound")
    if len(stderr) > plan["run_policy"]["max_stderr_bytes"]:
        problems.append("stderr exceeds bound")
    report: Any = None
    if not problems:
        try:
            report = json.loads(stdout.decode("utf-8"))
            basic_report_check(report, row, plan)
            write_exclusive(child_root / "report.json", report)
            write_exclusive(child_root / "parse.json", {
                "schema": "litchi.performance.0966.child_parse.v1",
                "status": "parsed",
                "child_id": row["id"],
                "report": descriptor(child_root / "report.json", "parsed report"),
                "parsed_utc": utc_now(),
            })
        except (UnicodeError, json.JSONDecodeError, RunnerError, OSError) as error:
            problems.append(f"report parse/identity failure: {error}")
    if problems:
        write_exclusive(child_root / "failure.json", {
            "schema": "litchi.performance.0966.child_failure.v1",
            "status": "rejected_no_retry",
            "child_id": row["id"],
            "problems": problems,
            "recorded_utc": utc_now(),
        })
        return False, "; ".join(problems)
    return True, "ok"


def capture(plan_path: Path, plan: dict[str, Any], go_path: Path) -> None:
    verify_go(go_path, plan_path, plan)
    capture_root = Path(plan["capture_root"]).resolve()
    require(not capture_root.exists(), f"capture root is retained; refusing retry: {capture_root}")
    capture_root.mkdir(parents=True, exist_ok=False)
    run_start = {
        "schema": "litchi.performance.0966.run_start.v1",
        "status": "started",
        "plan": descriptor(plan_path, "plan"),
        "go": descriptor(go_path, "GO"),
        "source_tree_sha256": plan["source"]["tree_sha256"],
        "executable": plan["executable"],
        "children": 48,
        "started_utc": utc_now(),
    }
    write_exclusive(capture_root / "run-start.json", run_start)
    results: list[dict[str, Any]] = []
    for row in plan["schedule"]["rows"]:
        ok, detail = run_child(plan, plan_path, row, capture_root)
        results.append({"id": row["id"], "ok": ok, "detail": detail})
    failures = [item for item in results if not item["ok"]]
    write_exclusive(capture_root / "run-terminal.json", {
        "schema": "litchi.performance.0966.run_terminal.v1",
        "status": "completed" if not failures else "completed_with_failures",
        "children": len(results),
        "failures": failures,
        "retries": 0,
        "ended_utc": utc_now(),
    })
    if failures:
        raise RunnerError(f"{len(failures)} child rows failed; no retries were attempted")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--check-plan", action="store_true")
    parser.add_argument("--capture", action="store_true")
    parser.add_argument("--go", type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    require(args.check_plan ^ args.capture, "choose exactly one of --check-plan or --capture")
    plan = check_plan(args.plan)
    if args.check_plan:
        print(json.dumps({"schema": PLAN_SCHEMA, "status": "plan_valid", "children": 48, "cells": 24}, sort_keys=True))
        return
    require(args.go is not None, "--capture requires --go")
    capture(args.plan.resolve(), plan, args.go.resolve())
    print(json.dumps({"schema": PLAN_SCHEMA, "status": "capture_complete", "children": 48}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except RunnerError as error:
        raise SystemExit(f"FAIL: {error}")
