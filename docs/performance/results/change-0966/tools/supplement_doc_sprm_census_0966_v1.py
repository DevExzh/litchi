#!/usr/bin/env python3
"""Derive compact, read-only 0966 cell and custody supplements.

This helper consumes the sealed plan, completed capture packet, and retained
reduction.  It never invokes the executable, imports the runner/reducer, or
changes any sealed artifact.  Logical capacity values are explicitly labeled
as a parser model; this packet makes no allocator or timing claim.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any


PLAN_SCHEMA = "litchi.performance.0966.doc_sprm_census_plan.v1"
REDUCTION_SCHEMA = "litchi.performance.0966.doc_sprm_census_reduction.v1"
DETAILS_SCHEMA = "litchi.performance.0966.doc_sprm_census_cell_details.v1"
MANIFEST_SCHEMA = "litchi.performance.0966.doc_sprm_census_child_manifest.v1"
PHASES = ("snapshot_open", "edit_open", "staged_replacement", "final_commit", "output_copy", "other")
MAX_JSON_BYTES = 1 << 20
HEX = set("0123456789abcdef")


class SupplementError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise SupplementError(message)


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


def read_json(path: Path, label: str) -> Any:
    path = regular_file(path, label)
    require(path.stat().st_size < MAX_JSON_BYTES, f"{label}: JSON exceeds bound")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise SupplementError(f"{label}: invalid JSON: {error}") from error


def write_exclusive(path: Path, value: object) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(data) < MAX_JSON_BYTES, f"output exceeds {MAX_JSON_BYTES} bytes: {path}")
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            import os
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise SupplementError(f"refusing to overwrite retained output: {path}") from error


def check_descriptor(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    path = regular_file(Path(value.get("path", "")), label)
    require(type(value.get("bytes")) is int and value["bytes"] == path.stat().st_size,
            f"{label}: byte count changed")
    expected = value.get("sha256")
    require(isinstance(expected, str) and len(expected) == 64 and set(expected) <= HEX,
            f"{label}: SHA-256 missing")
    require(sha256_file(path) == expected, f"{label}: SHA-256 changed")
    return path


def check_plan(plan_path: Path) -> dict[str, Any]:
    plan = read_json(plan_path, "plan")
    require(isinstance(plan, dict) and plan.get("schema") == PLAN_SCHEMA, "plan schema")
    require(plan.get("status") == "sealed_ready_for_go", "plan is not sealed")
    require(plan.get("schedule", {}).get("children") == 48 and
            plan.get("schedule", {}).get("cells") == 24, "plan schedule count")
    require(isinstance(plan.get("schedule", {}).get("rows"), list) and
            len(plan["schedule"]["rows"]) == 48, "plan rows")
    check_descriptor(plan.get("executable"), "plan executable")
    return plan


def check_child(capture_root: Path, row: dict[str, Any], plan_hash: str) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    child = capture_root / row["id"]
    require(child.is_dir() and not child.is_symlink(), f"{row['id']}: child directory missing")
    start_path = child / "start.json"
    terminal_path = child / "terminal.json"
    parse_path = child / "parse.json"
    report_path = child / "report.json"
    start = read_json(start_path, f"{row['id']}: start")
    terminal = read_json(terminal_path, f"{row['id']}: terminal")
    parse = read_json(parse_path, f"{row['id']}: parse")
    report = read_json(report_path, f"{row['id']}: report")
    require(start.get("child_id") == row["id"] and start.get("row") == row,
            f"{row['id']}: start row binding")
    require(start.get("plan_sha256") == plan_hash, f"{row['id']}: plan binding")
    require(terminal.get("status") == "terminal" and terminal.get("child_id") == row["id"] and
            terminal.get("returncode") == 0 and terminal.get("timed_out") is False and
            terminal.get("launch_error") is None, f"{row['id']}: terminal is not successful")
    require(parse.get("status") == "parsed" and parse.get("child_id") == row["id"],
            f"{row['id']}: parse receipt")
    report_descriptor = parse.get("report")
    report_file = check_descriptor(report_descriptor, f"{row['id']}: report descriptor")
    require(report_file == report_path.resolve(), f"{row['id']}: report path binding")
    return start, terminal, report


def count_fields(bucket: dict[str, Any]) -> dict[str, Any]:
    return {
        "calls": bucket["calls"],
        "successes": bucket["successes"],
        "errors": bucket["errors"],
        "input_empty_calls": bucket["input_empty"],
        "input_nonempty_calls": bucket["input_nonempty"],
        "logical_vec_capacity_bytes_model": bucket["logical_vec_capacity_bytes"],
    }


def cell_detail(row: dict[str, Any], child_ids: list[str], report: dict[str, Any], r1_path: Path, r2_path: Path) -> dict[str, Any]:
    require(report.get("timing_claim") is False and report.get("allocator_timing_claim") is False and
            report.get("allocator_instrumented") is False and report.get("allocation") is None,
            f"{row['id']}: report timing/allocator boundary")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1, f"{row['id']}: sample shape")
    census = samples[0].get("sprm_census")
    require(isinstance(census, dict) and census.get("schema_version") == 1 and
            census.get("saturated") is False, f"{row['id']}: census shape")
    labels = census.get("histogram_labels")
    require(isinstance(labels, list) and len(labels) == 10, f"{row['id']}: histogram labels")
    sprm_size = census.get("sprm_size_bytes")
    require(type(sprm_size) is int and sprm_size > 0, f"{row['id']}: SPRM size")
    total = census.get("total")
    require(isinstance(total, dict), f"{row['id']}: total bucket")
    for field in ("calls", "successes", "errors", "input_empty", "input_nonempty",
                  "input_bytes_total", "output_count_total", "output_capacity_total",
                  "logical_vec_capacity_bytes"):
        require(type(total.get(field)) is int and total[field] >= 0,
                f"{row['id']}: invalid total {field}")
    require(total["logical_vec_capacity_bytes"] == total["output_capacity_total"] * sprm_size,
            f"{row['id']}: logical model equation")
    phase_totals = {}
    for phase in PHASES:
        bucket = census.get(phase)
        require(isinstance(bucket, dict), f"{row['id']}: missing phase {phase}")
        phase_totals[phase] = count_fields(bucket)
    return {
        "case": row["case"],
        "edits": row["edits"],
        "retention": row["retention"],
        "children": child_ids,
        "repeat_equal_verified_by_reduction": True,
        "r1_report": descriptor(r1_path, f"{row['id']}: R1 report"),
        "r2_report": descriptor(r2_path, f"{row['id']}: R2 report"),
        "census": {
            "sprm_size_bytes": sprm_size,
            "histogram_labels": labels,
            "total": {
                "calls": total["calls"],
                "successes": total["successes"],
                "errors": total["errors"],
                "input_empty_calls": total["input_empty"],
                "input_nonempty_calls": total["input_nonempty"],
                "input_length_histogram": total["input_length_histogram"],
                "output_count_histogram": total["output_count_histogram"],
                "output_capacity_total_elements": total["output_capacity_total"],
                "logical_vec_capacity_bytes_model": total["logical_vec_capacity_bytes"],
            },
            "phase_totals": phase_totals,
        },
        "interpretation_boundary": "Logical Vec capacity model only; model bytes are not allocator traffic, RSS, elapsed time, or removable-work evidence.",
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path, default=Path("/home/zhuhe/litchi-performance-0966/protocol/plan-doc-sprm-census-0966-v1.json"))
    parser.add_argument("--capture", type=Path, default=Path("/home/zhuhe/litchi-performance-0966/capture/doc-sprm-census-v1"))
    parser.add_argument("--reduction", type=Path, default=Path("/home/zhuhe/litchi-performance-0966/analysis/doc-sprm-census-reduction-v1.json"))
    parser.add_argument("--details-output", type=Path, default=Path("/home/zhuhe/litchi-performance-0966/analysis/census-cell-details-v1.json"))
    parser.add_argument("--manifest-output", type=Path, default=Path("/home/zhuhe/litchi-performance-0966/analysis/child-terminal-report-manifest-v1.json"))
    args = parser.parse_args()

    plan = check_plan(args.plan.resolve())
    plan_hash = sha256_file(args.plan)
    reduction = read_json(args.reduction.resolve(), "reduction")
    require(isinstance(reduction, dict) and reduction.get("schema") == REDUCTION_SCHEMA and
            reduction.get("status") == "diagnostic_only_pass" and
            reduction.get("repeat_equal") is True and reduction.get("children") == 48 and
            reduction.get("cells") == 24, "reduction is not the completed passing packet")
    capture_root = args.capture.resolve()
    require(capture_root == Path(plan["capture_root"]).resolve(), "capture root is not plan-bound")
    run_start_path = capture_root / "run-start.json"
    run_terminal_path = capture_root / "run-terminal.json"
    run_start = read_json(run_start_path, "run-start")
    run_terminal = read_json(run_terminal_path, "run-terminal")
    require(run_start.get("status") == "started" and run_start.get("plan", {}).get("sha256") == plan_hash,
            "run-start binding")
    require(run_terminal.get("status") == "completed" and run_terminal.get("children") == 48 and
            run_terminal.get("failures") == [] and run_terminal.get("retries") == 0,
            "run-terminal is not a complete no-retry packet")

    rows = plan["schedule"]["rows"]
    expected_ids = {row["id"] for row in rows}
    actual_ids = {path.name for path in capture_root.iterdir() if path.is_dir()}
    require(actual_ids == expected_ids, "capture child set differs from sealed schedule")
    checked: dict[str, tuple[dict[str, Any], dict[str, Any], dict[str, Any]]] = {}
    for row in rows:
        checked[row["id"]] = check_child(capture_root, row, plan_hash)

    grouped: dict[tuple[str, int, str], list[dict[str, Any]]] = {}
    manifest_children = []
    for row in rows:
        start, terminal, report = checked[row["id"]]
        child = capture_root / row["id"]
        grouped.setdefault((row["case"], row["edits"], row["retention"]), []).append({
            "row": row, "report": report, "report_path": child / "report.json",
        })
        manifest_children.append({
            "id": row["id"],
            "repeat": row["repeat"],
            "direction": row["direction"],
            "case": row["case"],
            "edits": row["edits"],
            "retention": row["retention"],
            "terminal": descriptor(child / "terminal.json", f"{row['id']}: terminal"),
            "report": descriptor(child / "report.json", f"{row['id']}: report"),
        })

    details = []
    for key in sorted(grouped):
        items = sorted(grouped[key], key=lambda item: item["row"]["repeat"])
        require(len(items) == 2 and items[0]["row"]["repeat"] == 1 and items[1]["row"]["repeat"] == 2,
                f"{key}: repeat pair missing")
        r1 = items[0]["row"]
        details.append(cell_detail(
            r1,
            [items[0]["row"]["id"], items[1]["row"]["id"]],
            items[0]["report"], items[0]["report_path"].resolve(), items[1]["report_path"].resolve(),
        ))

    reduction_descriptor = descriptor(args.reduction, "reduction")
    write_exclusive(args.details_output, {
        "schema": DETAILS_SCHEMA,
        "schema_version": 1,
        "change": "0966",
        "status": "derived_from_completed_pass",
        "plan": descriptor(args.plan, "plan"),
        "reduction": reduction_descriptor,
        "capture_root": str(capture_root),
        "cells": 24,
        "fresh_children": 48,
        "repeat_equal": True,
        "timing_claim": False,
        "allocator_total_claim": False,
        "cell_rows": details,
        "interpretation_boundary": "Logical Vec capacity model only; model bytes are not allocator traffic, RSS, elapsed time, or removable-work evidence.",
    })
    write_exclusive(args.manifest_output, {
        "schema": MANIFEST_SCHEMA,
        "schema_version": 1,
        "change": "0966",
        "status": "custody_manifest_for_completed_pass",
        "plan": descriptor(args.plan, "plan"),
        "reduction": reduction_descriptor,
        "run_start": descriptor(run_start_path, "run-start"),
        "run_terminal": descriptor(run_terminal_path, "run-terminal"),
        "children": manifest_children,
        "children_count": len(manifest_children),
        "repeat_equal_verified_by_reduction": True,
        "interpretation_boundary": "Custody hashes only; no timing or allocator claim.",
    })
    print(json.dumps({
        "schema": DETAILS_SCHEMA,
        "status": "supplement_written",
        "cells": len(details),
        "children": len(manifest_children),
        "details": str(args.details_output.resolve()),
        "manifest": str(args.manifest_output.resolve()),
    }, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except SupplementError as error:
        raise SystemExit(f"FAIL: {error}")
