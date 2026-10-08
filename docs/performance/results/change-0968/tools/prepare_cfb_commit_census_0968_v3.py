#!/usr/bin/env python3
"""Bind a prospective 0968 diagnostic CFB census plan.

This helper hashes sealed external artifacts only.  It never builds, launches,
opens a fixture, or overwrites a retained plan.  All output paths are explicit
so ``--help`` and argument errors are side-effect free.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
from pathlib import Path
from typing import Any

import cfb_census_schema_0968_v3 as schema

ARTIFACT_ROOT = Path("/home/zhuhe/litchi-performance-0968")
ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
    "PAGER": "cat",
}


class PlanError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise PlanError(message)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def json_descriptor(path: Path, label: str) -> dict[str, Any]:
    value = schema.descriptor(path, label)
    require(value["bytes"] < schema.MAX_JSON_BYTES, f"{label}: JSON too large")
    return value


def write_exclusive(path: Path, value: object) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(payload) < schema.MAX_JSON_BYTES, "plan exceeds JSON size bound")
    try:
        with path.open("xb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as exc:
        raise PlanError(f"refusing to overwrite retained plan: {path}") from exc


def load_review(path: Path) -> dict[str, Any]:
    value = schema.load_json(path)
    schema.check_review(value, "review")
    return value


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-harness", type=Path, required=True,
                        help="sealed diagnostic source/harness manifest JSON")
    parser.add_argument("--binary", type=Path, required=True,
                        help="sealed diagnostic executable")
    parser.add_argument("--ordinary-oracle", type=Path, required=True,
                        help="ordinary 0967 byte-equivalence output oracle JSON")
    parser.add_argument("--oracle-equivalence", type=Path, required=True,
                        help="independent ordinary-output equivalence receipt JSON")
    parser.add_argument("--review", type=Path, required=True,
                        help="independent census review JSON with status pass/passed")
    parser.add_argument("--incident", type=Path, required=True,
                        help="immutable incident receipt for the rejected v2 epoch")
    parser.add_argument("--output", type=Path, required=True,
                        help="new plan JSON path; existing paths are never overwritten")
    parser.add_argument("--capture-root", type=Path, required=True,
                        help="new external capture directory; it must not exist")
    parser.add_argument("--runner", type=Path, default=ARTIFACT_ROOT / "tools/run_cfb_commit_census_0968_v3.py")
    parser.add_argument("--reducer", type=Path, default=ARTIFACT_ROOT / "tools/reduce_cfb_commit_census_0968_v3.py")
    parser.add_argument("--schema-tool", type=Path, default=ARTIFACT_ROOT / "tools/cfb_census_schema_0968_v3.py")
    parser.add_argument("--tests", type=Path, default=ARTIFACT_ROOT / "tools/test_cfb_commit_census_0968_v3.py")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    source_harness_path = args.source_harness.resolve()
    source_harness = schema.load_json(source_harness_path)
    require(isinstance(source_harness, dict), "source-harness manifest is not an object")
    source_revision = source_harness.get(
        "source_revision",
        source_harness.get("current_head", source_harness.get("source_baseline_head")),
    )
    require(isinstance(source_revision, str) and schema.REVISION_RE.fullmatch(source_revision), "source revision is not 40 lowercase hex characters")
    source = schema.check_source_harness_manifest(source_harness, "source-harness", source_revision)
    source_root = Path(source["root"]).resolve()
    file_count = source["file_count"]
    scope_sha256 = source["scope_sha256"]
    fixtures = {}
    for case, relative in schema.CASE_FILES.items():
        fixtures[case] = schema.descriptor(source_root / relative, f"fixture {case}")

    binary = schema.descriptor(args.binary.resolve(), "diagnostic executable")
    require(os.access(Path(binary["path"]), os.X_OK), "diagnostic executable is not executable")
    provenance_binary = source_harness.get("binary")
    if isinstance(provenance_binary, dict):
        require(binary["path"] == provenance_binary.get("path") and binary["sha256"] == provenance_binary.get("sha256"),
                "diagnostic executable differs from provenance binding")
    oracle_path = args.ordinary_oracle.resolve()
    oracle = schema.load_ordinary_oracle(schema.load_json(oracle_path), "ordinary oracle", source_revision)
    oracle_descriptor = json_descriptor(oracle_path, "ordinary oracle")
    equivalence_path = args.oracle_equivalence.resolve()
    equivalence = schema.load_json(equivalence_path)
    equivalence_descriptor = json_descriptor(equivalence_path, "ordinary oracle equivalence")
    schema.check_oracle_equivalence(equivalence, "ordinary oracle equivalence", oracle_path, oracle_descriptor["sha256"], source_revision)
    review_path = args.review.resolve()
    review_value = load_review(review_path)
    review_descriptor = json_descriptor(review_path, "independent review")
    incident_path = args.incident.resolve()
    incident_value = schema.load_json(incident_path)
    schema.check_incident(incident_value, "incident")
    incident_descriptor = json_descriptor(incident_path, "incident")

    output = args.output.resolve()
    capture_root = args.capture_root.resolve()
    require(output != capture_root and not capture_root.exists(), "capture root must be a new external directory")
    require(capture_root.is_absolute(), "capture root must be absolute")
    require(output.is_absolute(), "plan output must be absolute")

    tools = {}
    for key, path in {
        "runner": args.runner,
        "reducer": args.reducer,
        "schema": args.schema_tool,
        "tests": args.tests,
    }.items():
        tools[key] = json_descriptor(path.resolve(), f"tool {key}")

    rows = schema.expected_rows()
    plan = {
        "schema": schema.PLAN_SCHEMA,
        "schema_version": 3,
        "change": "0968",
        "created_utc": utc_now(),
        "status": "prospective_only_no_build_no_capture",
        "status_note": "Diagnostic census only; root GO is required before any child launch.",
        "source_revision": source_revision,
        "source": {
            "root": str(source_root),
            "file_count": file_count,
            "scope_sha256": scope_sha256,
            "scope_files": source["scope_files"],
            "baseline_archive_sha256": source.get("baseline_archive_sha256"),
            "diagnostic_boundary": source["diagnostic_boundary"],
        },
        "source_harness": json_descriptor(source_harness_path, "source-harness manifest"),
        "binary": binary,
        "fixtures": fixtures,
        "ordinary_output_oracle": oracle_descriptor,
        "ordinary_oracle_equivalence": equivalence_descriptor,
        "review": {
            "role": review_value.get("reviewer_role", "independent_census_review"),
            "status": review_value["status"],
            "document": review_descriptor,
        },
        "incident": incident_descriptor,
        "tools": tools,
        "claim_scope": {
            "diagnostic_only": True,
            "timing_claim": False,
            "allocator_total_claim": False,
            "elapsed_time_claim": False,
            "optimization_candidate": False,
            "adoption_claim": False,
        },
        "report_schema": schema.REPORT_SCHEMA,
        "report_contract": {
            "timing_claim": False,
            "elapsed_time_field": "forbidden",
            "allocator_total_field": "forbidden",
            "metadata_projection_counter": "forbidden",
            "binary_or_source_revision_fields": "forbidden_in_child_packet; custody is in plan/capture manifest",
            "route_counter_scope": "ObjectEditor::commit_candidate_with_rendered through successful discovery; capture stream reads only",
            "layout_reuse_definition": "layout_reuse_success is the writer's successful last_sector_layout report and covers the validated/emitted reuse path",
            "payload_bytes_definition": "sum of stream payload lengths opened by the one full Package::capture",
            "legacy_correctness_allocator_name": "allocator_retained_matches_output_capacity is a retained correctness oracle boolean, not allocator measurement",
            "correctness_gate_set": list(schema.GATE_NAMES),
        },
        "matrix": {
            "cases": list(schema.CASES),
            "edits": list(schema.EDIT_COUNTS),
            "retentions": list(schema.RETENTIONS),
            "cells": 24,
            "repetitions_per_cell": 2,
            "children": 48,
            "fresh_serial_child": True,
        },
        "schedule": {"rows": rows},
        "argv_template": [
            str(binary["path"]),
            "--input", "{fixture}",
            "--case", "{case}",
            "--edits", "{edits}",
            "--retention", "{retention_arg}",
            "--warmup", "0",
            "--samples", "1",
        ],
        "environment": ENVIRONMENT,
        "run_policy": {
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
        },
        "capture_root": str(capture_root),
        "no_retry": True,
        "no_candidate": True,
        "ordinary_output_oracle_cells": oracle,
    }
    write_exclusive(output, plan)
    print(json.dumps({"schema": schema.PLAN_SCHEMA, "status": "plan_written", "path": str(output), "children": 48, "cells": 24}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (PlanError, schema.CensusError) as exc:
        raise SystemExit(f"FAIL: {exc}")
