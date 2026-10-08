#!/usr/bin/env python3
"""Run the sealed 0968 diagnostic CFB census serially.

The default operation is ``--check-plan`` and has no output side effect.
``--capture`` requires a root-owned GO receipt, launches exactly the 48 fixed
children serially, keeps terminal/stdout/stderr evidence before parsing, and
never retries a child.  This runner makes no timing or allocator claim.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
from typing import Any

import cfb_census_schema_0968_v3 as schema


class RunnerError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RunnerError(message)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def write_exclusive(path: Path, value: object) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(payload) < schema.MAX_JSON_BYTES, f"JSON exceeds bound: {path}")
    try:
        with path.open("xb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as exc:
        raise RunnerError(f"refusing to overwrite retained file: {path}") from exc


def write_bytes_exclusive(path: Path, value: bytes) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(value)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as exc:
        raise RunnerError(f"refusing to overwrite retained file: {path}") from exc


def check_plan(plan_path: Path) -> dict[str, Any]:
    plan = schema.load_json(plan_path)
    require(isinstance(plan, dict) and plan.get("schema") == schema.PLAN_SCHEMA and
            plan.get("schema_version") == 3, "plan schema/version")
    require(plan.get("change") == "0968", "plan change")
    require(plan.get("status") in {"prospective_only_no_build_no_capture", "sealed_ready_for_go"}, "plan status")
    revision = plan.get("source_revision")
    require(isinstance(revision, str) and schema.REVISION_RE.fullmatch(revision), "source revision")
    source = plan.get("source")
    require(isinstance(source, dict), "source binding")
    source_root = Path(source.get("root", ""))
    require(source_root.is_dir(), "source root missing")

    source_harness_descriptor = plan.get("source_harness")
    source_harness_path = schema.verify_descriptor(source_harness_descriptor, "source-harness manifest")
    source_harness = schema.load_json(source_harness_path)
    source_binding = schema.check_source_harness_manifest(source_harness, "source-harness", revision)
    require(source_binding["root"] == str(source_root.resolve()), "source root differs from source-harness manifest")
    require(source_binding["file_count"] == source.get("file_count") and source_binding["scope_sha256"] == source.get("scope_sha256"), "source-harness scope binding changed")
    require(source_binding["scope_files"] == source.get("scope_files"), "source-harness file scope changed")

    binary = schema.verify_descriptor(plan.get("binary"), "diagnostic executable", executable=True)
    require(plan["binary"]["path"] == str(binary), "binary path binding")
    fixtures = plan.get("fixtures")
    require(isinstance(fixtures, dict) and set(fixtures) == set(schema.CASES), "fixture binding")
    for case, relative in schema.CASE_FILES.items():
        path = schema.verify_descriptor(fixtures[case], f"fixture {case}")
        require(path == source_root / relative, f"fixture {case} path binding")

    oracle_path = schema.verify_descriptor(plan.get("ordinary_output_oracle"), "ordinary output oracle")
    oracle = schema.load_ordinary_oracle(schema.load_json(oracle_path), "ordinary oracle", revision)
    require(plan.get("ordinary_output_oracle_cells") == oracle, "ordinary output oracle copy changed")
    equivalence_path = schema.verify_descriptor(plan.get("ordinary_oracle_equivalence"), "ordinary oracle equivalence")
    schema.check_oracle_equivalence(schema.load_json(equivalence_path), "ordinary oracle equivalence", oracle_path, schema.sha256_file(oracle_path), revision)

    review = plan.get("review")
    require(isinstance(review, dict), "review binding")
    review_path = schema.verify_descriptor(review.get("document"), "independent review")
    review_value = schema.load_json(review_path)
    schema.check_review(review_value, "independent review")
    require(review.get("status") == review_value.get("status"), "review status binding")

    incident_descriptor = schema.verify_descriptor(plan.get("incident"), "incident")
    schema.check_incident(schema.load_json(incident_descriptor), "incident")

    tools = plan.get("tools")
    require(isinstance(tools, dict), "tool bindings")
    for label in ("runner", "reducer", "schema", "tests"):
        schema.verify_descriptor(tools.get(label), f"tool {label}")

    require(plan.get("report_schema") == schema.REPORT_SCHEMA, "report schema binding")
    scope = plan.get("claim_scope")
    require(isinstance(scope, dict), "claim scope")
    for field in ("diagnostic_only", "timing_claim", "allocator_total_claim", "elapsed_time_claim", "optimization_candidate", "adoption_claim"):
        expected = True if field == "diagnostic_only" else False
        require(scope.get(field) is expected, f"claim scope {field}")
    contract = plan.get("report_contract")
    require(isinstance(contract, dict) and contract.get("elapsed_time_field") == "forbidden" and contract.get("metadata_projection_counter") == "forbidden", "report contract")

    policy = plan.get("run_policy")
    require(policy == {
        "serial": True,
        "cargo_processes": 0,
        "retry": False,
        "warmups": 0,
        "samples": 1,
        "deadline_seconds": 300,
        "max_stdout_bytes": schema.MAX_STDOUT_BYTES,
        "max_stderr_bytes": schema.MAX_STDERR_BYTES,
        "terminal_before_parse": True,
        "raw_stdout_stderr_retained": True,
        "cpu_affinity": None,
    }, "run policy changed")
    require(plan.get("environment") == {
        "PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
        "RUST_BACKTRACE": "0", "PAGER": "cat",
    }, "environment changed")
    rows = plan.get("schedule", {}).get("rows") if isinstance(plan.get("schedule"), dict) else None
    require(rows == schema.expected_rows(), "fixed 48-row schedule changed")
    require(Path(plan.get("capture_root", "")).is_absolute(), "capture root must be absolute")
    argv_template = plan.get("argv_template")
    require(isinstance(argv_template, list) and argv_template[0] == str(binary), "argv template does not bind executable")
    require(plan.get("no_retry") is True and plan.get("no_candidate") is True, "claim policy changed")
    return plan


def verify_go(go_path: Path, plan_path: Path, plan: dict[str, Any]) -> dict[str, Any]:
    go = schema.load_json(go_path)
    schema.check_go(go, "GO", plan_path, plan)
    return go


def argv_for(plan: dict[str, Any], row: dict[str, Any]) -> list[str]:
    fixture = plan["fixtures"][row["case"]]["path"]
    substitutions = {
        "{fixture}": fixture,
        "{case}": row["case"],
        "{edits}": str(row["edits"]),
        "{retention_arg}": row["retention_arg"],
    }
    result = []
    for item in plan["argv_template"]:
        value = str(item)
        for key, replacement in substitutions.items():
            value = value.replace(key, replacement)
        require("{" not in value and "}" not in value, f"unbound argv placeholder: {value}")
        result.append(value)
    return result


def run_child(plan_path: Path, plan: dict[str, Any], row: dict[str, Any], capture_root: Path, oracle: dict[str, dict[str, Any]]) -> dict[str, Any]:
    child_root = capture_root / row["child_id"]
    require(not child_root.exists(), f"retained child exists: {child_root}")
    child_root.mkdir(parents=True, exist_ok=False)
    argv = argv_for(plan, row)
    start = {
        "schema": schema.CHILD_START_SCHEMA,
        "status": "started",
        "child_id": row["child_id"],
        "row": row,
        "plan_sha256": schema.sha256_file(plan_path),
        "source": plan["source"],
        "source_harness": plan["source_harness"],
        "binary": plan["binary"],
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
            close_fds=True,
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
    except (OSError, ValueError) as exc:
        launch_error = str(exc)
        if process is not None:
            returncode = process.returncode

    write_bytes_exclusive(child_root / "stdout.raw", stdout)
    write_bytes_exclusive(child_root / "stderr.raw", stderr)
    stdout_descriptor = schema.descriptor(child_root / "stdout.raw", "raw stdout")
    stderr_descriptor = schema.descriptor(child_root / "stderr.raw", "raw stderr")
    terminal = {
        "schema": schema.CHILD_TERMINAL_SCHEMA,
        "status": "terminal",
        "child_id": row["child_id"],
        "returncode": returncode,
        "timed_out": timed_out,
        "launch_error": launch_error,
        "stdout": stdout_descriptor,
        "stderr": stderr_descriptor,
        "ended_utc": utc_now(),
    }
    write_exclusive(child_root / "terminal.json", terminal)
    terminal_descriptor = schema.descriptor(child_root / "terminal.json", "terminal")
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
    report_descriptor = None
    if not problems:
        try:
            report = json.loads(stdout.decode("utf-8"))
            schema.check_report(report, row, plan, oracle)
            report_path = child_root / "report.json"
            write_exclusive(report_path, report)
            report_descriptor = schema.descriptor(report_path, "report")
            write_exclusive(child_root / "parse.json", {
                "schema": schema.CHILD_PARSE_SCHEMA,
                "status": "parsed",
                "child_id": row["child_id"],
                "report": report_descriptor,
                "parsed_utc": utc_now(),
            })
        except (UnicodeError, json.JSONDecodeError, schema.CensusError, OSError, TypeError, AttributeError) as exc:
            problems.append(f"report parse/identity failure: {exc}")
    failure_descriptor = None
    if problems:
        failure_path = child_root / "failure.json"
        write_exclusive(failure_path, {
            "schema": schema.CHILD_FAILURE_SCHEMA,
            "status": "rejected_no_retry",
            "child_id": row["child_id"],
            "problems": problems,
            "recorded_utc": utc_now(),
        })
        failure_descriptor = schema.descriptor(failure_path, "failure")
    return {
        "child_id": row["child_id"],
        "row": row,
        "root": str(child_root),
        "status": "passed" if not problems else "failed",
        "problems": problems,
        "start": schema.descriptor(child_root / "start.json", "start"),
        "terminal": terminal_descriptor,
        "stdout": stdout_descriptor,
        "stderr": stderr_descriptor,
        "report": report_descriptor,
        "failure": failure_descriptor,
    }


def capture(plan_path: Path, plan: dict[str, Any], go_path: Path) -> None:
    verify_go(go_path, plan_path, plan)
    capture_root = Path(plan["capture_root"]).resolve()
    require(not capture_root.exists(), f"capture root retained; refusing retry: {capture_root}")
    capture_root.mkdir(parents=True, exist_ok=False)
    oracle = plan["ordinary_output_oracle_cells"]
    run_start = {
        "schema": schema.RUN_START_SCHEMA,
        "status": "started",
        "plan": schema.descriptor(plan_path, "plan"),
        "go": schema.descriptor(go_path, "GO"),
        "source": plan["source"],
        "source_harness": plan["source_harness"],
        "binary": plan["binary"],
        "children": 48,
        "cells": 24,
        "retries": 0,
        "started_utc": utc_now(),
    }
    write_exclusive(capture_root / "run-start.json", run_start)
    results = []
    for row in plan["schedule"]["rows"]:
        results.append(run_child(plan_path, plan, row, capture_root, oracle))
    failures = [item for item in results if item["status"] != "passed"]
    terminal_value = {
        "schema": schema.RUN_TERMINAL_SCHEMA,
        "status": "completed" if not failures else "completed_with_failures",
        "children": len(results),
        "failures": [item["child_id"] for item in failures],
        "retries": 0,
        "ended_utc": utc_now(),
    }
    write_exclusive(capture_root / "run-terminal.json", terminal_value)
    manifest = {
        "schema": schema.CAPTURE_SCHEMA,
        "status": "complete" if not failures else "complete_with_failures",
        "plan": schema.descriptor(plan_path, "plan"),
        "go": schema.descriptor(go_path, "GO"),
        "source": plan["source"],
        "source_harness": plan["source_harness"],
        "binary": plan["binary"],
        "packet_count": len(results),
        "children": results,
        "failures": [item["child_id"] for item in failures],
        "retries": 0,
        "created_utc": utc_now(),
    }
    manifest_path = capture_root / "capture-manifest-v3.json"
    write_exclusive(manifest_path, manifest)
    if failures:
        raise RunnerError(f"{len(failures)} child rows failed; no retries were attempted")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check-plan", action="store_true")
    mode.add_argument("--capture", action="store_true")
    parser.add_argument("--go", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    plan_path = args.plan.resolve()
    plan = check_plan(plan_path)
    if args.check_plan:
        print(json.dumps({"schema": schema.PLAN_SCHEMA, "status": "plan_valid", "children": 48, "cells": 24}, sort_keys=True))
        return 0
    require(args.go is not None, "--capture requires --go")
    capture(plan_path, plan, args.go.resolve())
    print(json.dumps({"schema": schema.PLAN_SCHEMA, "status": "capture_complete", "children": 48}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (RunnerError, schema.CensusError) as exc:
        raise SystemExit(f"FAIL: {exc}")
