#!/usr/bin/env python3
"""Generate one external, custody-bound 0973 pre-capture plan."""
from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path
from typing import Any

import protocol_common_0973_v2 as p

ROOT = Path(__file__).resolve().parents[1]
PROPOSAL = ROOT / "diagnostic-proposal-v1.md"
ADR = ROOT / "receipts/adr-goal-revalidation-0973-v1.json"
REVIEW_CONTEXT = Path("/home/zhuhe/litchi-performance-0972/receipts/dwfl-replay-author-seal-v6.json")
DESIGN = ROOT / "design-v1.md"
METRIC_PROPOSAL = ROOT / "analysis/metric-proposal-v1.md"
DESIGN_PRECEDENCE = ROOT / "design-precedence-clarification-v1.md"
REVIEW_FINDINGS = ROOT / "receipts/independent-review-0973-v1.json"
ASLR_CONTROL = ROOT / "receipts/aslr-translation-0973-final-v4.json"
SYNTHETIC_E2E = ROOT / "receipts/synthetic-packet-pipeline-author-v2-v4.json"


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def binding(path: str | Path) -> dict[str, Any]:
    return p.file_binding(path)


def build(helper_seal_path: Path) -> dict[str, Any]:
    seal = p.load_json(helper_seal_path)
    if seal.get("schema") != p.HELPER_SCHEMA or seal.get("status") != "sealed":
        raise p.ProtocolError("helper seal is not sealed 0973 v2")
    helpers = seal.get("helpers")
    if not isinstance(helpers, dict):
        raise p.ProtocolError("helper seal lacks helper bindings")
    helper_bindings = {name: binding(row["path"]) for name, row in helpers.items()}
    for name, row in helpers.items():
        if helper_bindings[name]["sha256"] != row.get("sha256"):
            raise p.ProtocolError(f"helper changed after seal: {name}")
    shim_so = Path(helpers["shim_so"]["path"])
    plan_path = ROOT / "plans/plan-0973-pre-capture-v2.json"
    plan: dict[str, Any] = {
        "schema": p.PLAN_SCHEMA,
        "status": "pre_capture",
        "created_utc": utc(),
        "plan_path": str(plan_path),
        "no_build_or_workload": True,
        "execution": {"serial": True, "retry": False, "timeout_seconds": 900, "children": p.CHILDREN},
        "matrix": {"children": p.CHILDREN, "warmups": p.WARMUPS, "samples": p.SAMPLES, "cpu": p.CPU, "cells": [{"case": "docfloat", "edits": 1, "retention": "default"}], "serial_order": list(range(1, p.CHILDREN + 1))},
        "perf": {"path": p.PERF_PATH, "bytes": Path(p.PERF_PATH).stat().st_size, "sha256": p.sha256(Path(p.PERF_PATH)), "event": "instructions:u", "period": p.PERIOD, "callgraph": p.CALLGRAPH, "stack_cap": p.STACK_CAP, "sample_flags": ["--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache", "--user-callchains"], "buildid_export": ["buildid-list", "-i", "<child>.perf.data"]},
        "marker": {"symbol": p.MARKER, "dso_path": p.BINARY_PATH, "build_id": p.BINARY_BUILD_ID, "calls": p.CALLS, "p50_scope": "80 measured samples only", "marker_scope": "5 warmups + 80 measured calls combined", "owner_scope": "all 12 children; missing owner membership unknown"},
        "pilot_gate": {"child": 1, "min_marker_samples": 100, "lost_events": 0, "marker_period_share_positive": True, "on_failure": "stop_without_retry"},
        "child_gate": {"all_children": True, "min_marker_samples": 100, "lost_events": 0, "marker_period_share_positive": True, "on_failure": "stop_without_retry"},
        "binary": {**binding(p.BINARY_PATH), "build_id": p.BINARY_BUILD_ID},
        "fixture": binding(p.FIXTURE_PATH),
        "source_manifest": binding(p.SOURCE_MANIFEST_PATH),
        "source_file_manifest": binding(p.SOURCE_FILE_MANIFEST_PATH),
        "source_proof": binding(p.SOURCE_PROOF_PATH),
        "proposal": binding(PROPOSAL),
        "design": binding(DESIGN),
        "metric_proposal": binding(METRIC_PROPOSAL),
        "design_precedence": binding(DESIGN_PRECEDENCE),
        "adr_revalidation": binding(ADR),
        "validator": {**binding(p.VALIDATOR_PATH), "retained_schema": "litchi.performance.0969.protocol_common.v4"},
        "launcher": binding(p.LAUNCHER_PATH),
        "python": binding(p.PYTHON_PATH),
        "taskset": binding(p.TASKSET_PATH),
        "libdw": {**binding(p.LIBDW_PATH), "version": p.LIBDW_VERSION},
        "libelf": {**binding(p.LIBELF_PATH), "version": p.LIBELF_VERSION},
        "compiler": {**binding(p.COMPILER_PATH), "version": "cc (Ubuntu 15.2.0-16ubuntu1)"},
        "header": binding(p.HEADER_PATH),
        "shim_source": binding(helpers["shim_source"]["path"]),
        "shim_version_script": binding(helpers["shim_map"]["path"]),
        "shim_build": {"argv": p.SHIM_BUILD_COMMAND, "compiler": binding(p.COMPILER_PATH), "header": binding(p.HEADER_PATH), "libdw": {**binding(p.LIBDW_PATH), "version": p.LIBDW_VERSION}, "libelf": {**binding(p.LIBELF_PATH), "version": p.LIBELF_VERSION}, "shared_object": binding(shim_so)},
        "shim": {"shared_object": binding(shim_so), "target_only": True, "offline_only": True, "native_child_preload": "absent", "env_values": ["LITCHI_0973_TARGET_PATH", "LITCHI_0973_TARGET_NAME", "LITCHI_0973_CORRECTED_BASE", "LITCHI_0973_PERF_EXPECTED_START"]},
        "helper_seal": binding(helper_seal_path),
        "helpers_seal_sha256": p.sha256(helper_seal_path),
        "helpers": helper_bindings,
        "review_context": binding(REVIEW_CONTEXT),
        "review_findings": binding(REVIEW_FINDINGS),
        "aslr_control": binding(ASLR_CONTROL),
        "synthetic_e2e": binding(SYNTHETIC_E2E),
        "environment": p.ENVIRONMENT,
        "native_environment": {"LD_PRELOAD": "absent", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"},
        "process_policy": {"require_sample_pid_tid_equal_exec_pid": True, "reject_any_other_thread_or_descendant": True, "sampling_monitor": False},
        "dynamic_layout": {"source": "exact per-child perf MMAP2 + ELF64 PT_LOAD", "hardcoded_aslr_addresses": False, "requires_all_pt_load_mappings": True, "ambiguity": "fail child without retry"},
        "report_contract": {"case": "docfloat", "edits": 1, "retention": "default_8mib", "warmups": p.WARMUPS, "samples_requested": p.SAMPLES, "sample_indices": list(range(p.SAMPLES)), "elapsed_ns_identity": "samples[i].elapsed_ns == elapsed_ns[i]"},
        "raw_stack_contract": {"schema": p.RAW_STACK_SCHEMA, "sample_type": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER", "period": p.PERIOD, "stack_cap": p.STACK_CAP, "bounds": "0 <= dyn_size <= captured_stack_size <= 32768", "missing_or_short": "missing fields reject; bounded short captured_stack_size is retained in its distribution", "raw_scope": "all PERF_RECORD_SAMPLE records without misc filtering"},
        "raw_artifacts": p.RAW_ARTIFACTS,
        "outputs": {"capture_root": str(ROOT / "capture/instruction-profile-v2"), "analysis_root": str(ROOT / "analysis")},
        "interpretation": {"diagnostic_only": True, "qualification_claim": False, "speed_claim": False, "exact_timed_region_claim": False, "marker_period_share": "raw decoder coverage only"},
    }
    return plan


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--helper-seal", type=Path, required=True)
    ap.add_argument("--output", type=Path, required=True)
    args = ap.parse_args()
    try:
        plan = build(args.helper_seal.resolve())
        out = args.output.resolve()
        plan["plan_path"] = str(out)
        p.check_plan_bindings(plan, None)  # checks all immutable fields before self path/output binding
        p.write_json_exclusive(out, plan)
        # Recheck with the final path now that the self-binding is present.
        p.check_plan_bindings(p.load_json(out), out)
        print(json.dumps({"schema": p.PLAN_SCHEMA, "status": "pass", "path": str(out), "sha256": p.sha256(out)}, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"0973 plan generator rejected: {exc}", file=__import__("sys").stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
