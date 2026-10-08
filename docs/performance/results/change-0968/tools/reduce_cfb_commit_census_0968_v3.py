#!/usr/bin/env python3
"""Independently reduce the retained 0968 diagnostic CFB census.

The reducer launches nothing and performs no build.  It checks raw terminal
custody, report schema, route conservation, the fixed 20 correctness gates,
ordinary output byte equivalence, and exact equality of the two fresh
repetitions.  A successful reduction is a valid diagnostic census only; it is
never an optimization or adoption decision.
"""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import json
import os
from pathlib import Path
from typing import Any

import cfb_census_schema_0968_v3 as schema


class ReducerError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReducerError(message)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def write_exclusive(path: Path, value: object) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(payload) < schema.MAX_JSON_BYTES, f"reduction exceeds bound: {path}")
    try:
        with path.open("xb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as exc:
        raise ReducerError(f"refusing to overwrite retained reduction: {path}") from exc


def check_plan(plan_path: Path) -> dict[str, Any]:
    plan = schema.load_json(plan_path)
    require(isinstance(plan, dict) and plan.get("schema") == schema.PLAN_SCHEMA and
            plan.get("schema_version") == 3, "plan schema/version")
    require(plan.get("change") == "0968" and plan.get("source_revision") and schema.REVISION_RE.fullmatch(plan["source_revision"]), "plan identity")
    source = plan.get("source")
    require(isinstance(source, dict), "source binding")
    source_harness_path = schema.verify_descriptor(plan.get("source_harness"), "source-harness manifest")
    source_harness = schema.load_json(source_harness_path)
    source_binding = schema.check_source_harness_manifest(source_harness, "source-harness", plan["source_revision"])
    require(source_binding["root"] == str(Path(source["root"]).resolve()), "source-harness root changed")
    require(source_binding["file_count"] == source.get("file_count") and source_binding["scope_sha256"] == source.get("scope_sha256"), "source-harness scope changed")
    require(source_binding["scope_files"] == source.get("scope_files"), "source-harness file scope changed")
    schema.verify_descriptor(plan.get("binary"), "diagnostic binary", executable=True)
    fixtures = plan.get("fixtures")
    require(isinstance(fixtures, dict) and set(fixtures) == set(schema.CASES), "fixture set")
    for case, relative in schema.CASE_FILES.items():
        path = schema.verify_descriptor(fixtures[case], f"fixture {case}")
        require(path == Path(source["root"]).resolve() / relative, f"fixture {case} path")
    oracle_path = schema.verify_descriptor(plan.get("ordinary_output_oracle"), "ordinary oracle")
    oracle = schema.load_ordinary_oracle(schema.load_json(oracle_path), "ordinary oracle", plan["source_revision"])
    require(plan.get("ordinary_output_oracle_cells") == oracle, "ordinary oracle copy changed")
    equivalence_path = schema.verify_descriptor(plan.get("ordinary_oracle_equivalence"), "ordinary oracle equivalence")
    schema.check_oracle_equivalence(schema.load_json(equivalence_path), "ordinary oracle equivalence", oracle_path, schema.sha256_file(oracle_path), plan["source_revision"])
    review = plan.get("review")
    require(isinstance(review, dict), "review binding")
    review_path = schema.verify_descriptor(review.get("document"), "review")
    review_value = schema.load_json(review_path)
    schema.check_review(review_value, "review")
    require(review.get("status") == review_value.get("status"), "review status binding")
    incident_descriptor = schema.verify_descriptor(plan.get("incident"), "incident")
    schema.check_incident(schema.load_json(incident_descriptor), "incident")
    tools = plan.get("tools")
    require(isinstance(tools, dict), "tool binding")
    for label in ("runner", "reducer", "schema", "tests"):
        schema.verify_descriptor(tools.get(label), f"tool {label}")
    require(plan.get("report_schema") == schema.REPORT_SCHEMA, "report schema binding")
    require(plan.get("schedule", {}).get("rows") == schema.expected_rows(), "fixed schedule changed")
    require(plan.get("claim_scope", {}).get("diagnostic_only") is True and plan["claim_scope"].get("timing_claim") is False and plan["claim_scope"].get("optimization_candidate") is False, "claim scope")
    return plan


def load_child_record(manifest_path: Path, plan_path: Path, item: dict[str, Any], plan: dict[str, Any], oracle: dict[str, dict[str, Any]]) -> dict[str, Any]:
    require(isinstance(item, dict), "manifest child is not an object")
    child_id = item.get("child_id")
    require(isinstance(child_id, str), "manifest child id")
    row = item.get("row")
    require(isinstance(row, dict), f"{child_id}: row missing")
    root = Path(item.get("root", ""))
    require(root.is_dir() and not root.is_symlink(), f"{child_id}: child root missing")
    for name in ("start", "terminal", "stdout", "stderr"):
        schema.verify_descriptor(item.get(name), f"{child_id}:{name}")
    start_path = schema.verify_descriptor(item["start"], f"{child_id}:start")
    start = schema.load_json(start_path)
    require(start.get("schema") == schema.CHILD_START_SCHEMA and start.get("status") == "started" and start.get("child_id") == child_id,
            f"{child_id}: start schema")
    require(start.get("row") == row, f"{child_id}: start row binding")
    require(start.get("plan_sha256") == schema.sha256_file(plan_path), f"{child_id}: start plan binding")
    require(start.get("source") == plan["source"] and start.get("source_harness") == plan["source_harness"] and
            start.get("binary") == plan["binary"] and start.get("fixture") == plan["fixtures"][row["case"]],
            f"{child_id}: start custody binding")
    terminal_path = schema.verify_descriptor(item["terminal"], f"{child_id}:terminal")
    terminal = schema.load_json(terminal_path)
    require(terminal.get("schema") == schema.CHILD_TERMINAL_SCHEMA and terminal.get("child_id") == child_id, f"{child_id}: terminal schema")
    require(terminal.get("returncode") == 0 and terminal.get("timed_out") is False and terminal.get("launch_error") is None, f"{child_id}: terminal failure")
    report_path = schema.verify_descriptor(item.get("report"), f"{child_id}:report")
    report = schema.load_json(report_path)
    schema.check_report(report, row, plan, oracle)
    return {"child_id": child_id, "row": row, "report": report, "root": str(root), "terminal": terminal}


def aggregate(reports: list[dict[str, Any]]) -> dict[str, Any]:
    overlay = collections.Counter()
    totals = {field: 0 for field in schema.CENSUS_FIELDS if field != "saturated"}
    for report in reports:
        census = report["samples"][0]["census"]
        for field in totals:
            totals[field] += census[field]
        overlay["success"] += census["overlay_success"]
        overlay["declined"] += census["overlay_declined"]
        overlay["error"] += census["overlay_error"]
        overlay["not_attempted"] += census["overlay_not_attempted"]
    return {"overlay_outcomes": dict(sorted(overlay.items())), **totals}


def reduce_records(plan: dict[str, Any], oracle: dict[str, dict[str, Any]], records: list[dict[str, Any]]) -> tuple[dict[str, Any], list[str]]:
    errors: list[str] = []
    expected_rows = {row["child_id"]: row for row in schema.expected_rows()}
    if len(records) != 48:
        errors.append(f"expected 48 records, got {len(records)}")
    by_id: dict[str, dict[str, Any]] = {}
    for record in records:
        child_id = record.get("child_id")
        if child_id in by_id:
            errors.append(f"duplicate child {child_id}")
        by_id[child_id] = record
        if child_id not in expected_rows:
            errors.append(f"unexpected child {child_id}")
            continue
        if record.get("row") != expected_rows[child_id]:
            errors.append(f"{child_id}: row identity mismatch")
    missing = sorted(set(expected_rows) - set(by_id))
    errors.extend(f"missing child {child_id}" for child_id in missing)
    valid_records: list[dict[str, Any]] = []
    for child_id, record in by_id.items():
        if child_id in expected_rows and record.get("row") == expected_rows[child_id] and isinstance(record.get("report"), dict):
            try:
                schema.check_report(record["report"], expected_rows[child_id], plan, oracle)
                valid_records.append(record)
            except schema.CensusError as exc:
                errors.append(f"{child_id}: {exc}")
    cells: dict[str, list[dict[str, Any]]] = collections.defaultdict(list)
    for record in valid_records:
        cells[record["row"]["id"]].append(record)
    cell_results: dict[str, Any] = {}
    for cell in schema.expected_cells():
        cell_id = cell["id"]
        rows = sorted(cells.get(cell_id, []), key=lambda item: item["row"]["repeat"])
        if len(rows) != 2 or {item["row"]["repeat"] for item in rows} != {1, 2}:
            errors.append(f"{cell_id}: repetition set is incomplete or duplicated")
            continue
        first, second = rows
        equal = schema.normalized_report(first["report"]) == schema.normalized_report(second["report"])
        if not equal:
            errors.append(f"{cell_id}: repetition route/counter/output/correctness mismatch")
        cell_results[cell_id] = {
            "repetitions": [first["child_id"], second["child_id"]],
            "repeat_equal": equal,
            "route_census": first["report"]["samples"][0]["census"],
            "output": {
                "sha256": first["report"]["output_sha256"],
                "bytes": first["report"]["output_bytes"],
            },
            "correctness": first["report"]["correctness"],
        }
    summary = {
        "schema": schema.REDUCTION_SCHEMA,
        "status": "pass" if not errors else "rejected",
        "decision": "diagnostic_only_no_adoption",
        "timing_claim": False,
        "allocator_total_claim": False,
        "packet_count": len(records),
        "cell_count": len(cell_results),
        "repeat_count_per_cell": 2,
        "errors": errors,
        "cells": cell_results,
        "aggregate": aggregate([record["report"] for record in valid_records]),
    }
    return summary, errors


def reduce_capture(plan_path: Path, manifest_path: Path, output_path: Path) -> dict[str, Any]:
    plan = check_plan(plan_path)
    manifest = schema.load_json(manifest_path)
    errors: list[str] = []
    require(isinstance(manifest, dict) and manifest.get("schema") == schema.CAPTURE_SCHEMA, "capture manifest schema")
    plan_descriptor = manifest.get("plan")
    require(isinstance(plan_descriptor, dict) and plan_descriptor.get("path") == str(plan_path.resolve()) and plan_descriptor.get("sha256") == schema.sha256_file(plan_path), "capture plan binding")
    go_descriptor = schema.verify_descriptor(manifest.get("go"), "capture GO")
    go = schema.load_json(go_descriptor)
    schema.check_go(go, "capture GO", plan_path, plan)
    require(manifest.get("packet_count") == 48 and manifest.get("retries") == 0, "capture count/retry policy")
    require(manifest.get("failures") == [] and manifest.get("status") == "complete", "capture contains failed children")
    source_harness = manifest.get("source_harness")
    require(manifest.get("source") == plan["source"], "capture source binding")
    require(source_harness == plan["source_harness"], "capture source-harness binding")
    require(manifest.get("binary") == plan["binary"], "capture binary binding")
    oracle = plan["ordinary_output_oracle_cells"]
    items = manifest.get("children")
    require(isinstance(items, list), "capture child list")
    records: list[dict[str, Any]] = []
    for item in items:
        try:
            records.append(load_child_record(manifest_path, plan_path, item, plan, oracle))
        except (schema.CensusError, OSError, TypeError, ValueError) as exc:
            errors.append(str(exc))
    summary, record_errors = reduce_records(plan, oracle, records)
    errors.extend(record_errors)
    summary["status"] = "pass" if not errors else "rejected"
    summary["errors"] = errors
    summary["plan"] = schema.descriptor(plan_path, "plan")
    summary["capture_manifest"] = schema.descriptor(manifest_path, "capture manifest")
    summary["source"] = plan["source"]
    summary["source_harness"] = plan["source_harness"]
    summary["binary"] = plan["binary"]
    summary["created_utc"] = utc_now()
    return summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    summary = reduce_capture(args.plan.resolve(), args.manifest.resolve(), args.output.resolve())
    write_exclusive(args.output.resolve(), summary)
    print(json.dumps({"schema": schema.REDUCTION_SCHEMA, "status": summary["status"], "output": str(args.output.resolve()), "errors": len(summary["errors"])}, sort_keys=True))
    return 0 if summary["status"] == "pass" else 2


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (ReducerError, schema.CensusError) as exc:
        raise SystemExit(f"FAIL: {exc}")
