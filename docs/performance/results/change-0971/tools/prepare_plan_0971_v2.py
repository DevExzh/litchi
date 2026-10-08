#!/usr/bin/env python3
"""Emit the immutable, pre-capture 0971 instruction-profile plan."""
from __future__ import annotations

import argparse
import datetime as dt
import importlib.util
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"


def module():
    path = TOOLS / "protocol_common_0971_v2.py"
    spec = importlib.util.spec_from_file_location("protocol_common_0971_v2", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot import protocol helper")
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    p = module()

    binary = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
    fixture = Path("/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc")
    source_manifest = Path("/home/zhuhe/litchi-performance-0969/manifests/source-manifest-v1.json")
    source_file_manifest = Path("/home/zhuhe/litchi-performance-0969/manifests/source-file-manifest-v1.json")
    proposal = ROOT / "attribution-proposal-v1.md"
    capability = ROOT / "receipts/perf-record-capability-probe-v1.json"
    variation = ROOT / "source-work-variation-v1.md"
    adr = ROOT / "receipts/adr-revalidation-0971-v1.json"
    validator = Path(p.VALIDATOR_PATH)
    launcher = Path("/home/zhuhe/litchi-performance-0970/tools/launch_baseline_v2.py")
    python = Path("/usr/bin/python3.14")
    perf = Path("/usr/bin/perf")
    taskset = Path("/usr/bin/taskset")
    harness_source = Path("/home/zhuhe/litchi-performance-0969/source-baseline/tools/perf-cfb-profile/src/lib.rs")
    harness_cargo = Path("/home/zhuhe/litchi-performance-0969/source-baseline/tools/perf-cfb-profile/Cargo.toml")
    source_proof = ROOT / "receipts/source-content-proof-0971-v2.json"
    helper_seal = ROOT / "receipts/helper-seal-0971-v2-correction-v2.json"
    review_context = ROOT / "review/independent-final-review-0971-v1.json"
    correction_receipt = ROOT / "receipts/correction-receipt-0971-v3.json"
    for path in (binary, fixture, source_manifest, source_file_manifest, source_proof, proposal, capability, variation, adr, validator, launcher, python, perf, taskset, harness_source, harness_cargo, helper_seal, review_context, correction_receipt):
        if not path.is_file():
            raise SystemExit(f"required binding is not ready: {path}")

    bind = p.file_binding
    plan_path = args.output.resolve()
    plan = {
        "schema": p.PLAN_SCHEMA,
        "status": "pre_capture",
        "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "plan_path": str(plan_path),
        "no_build_or_workload": True,
        "scope": {
            "claim": "diagnostic instruction attribution only",
            "qualification": False,
            "source_or_candidate_effect": False,
            "timing_claim": False,
            "native_binary": "retained historical 0967 baseline only",
        },
        "protocol_document": bind(proposal),
        "proposal": bind(proposal),
        "proposal_binding": {"path": str(proposal.resolve()), "sha256": p.sha256(proposal)},
        "capability": bind(capability),
        "capability_receipt": bind(capability),
        "source_variation": bind(variation),
        "adr_revalidation": bind(adr),
        "source_manifest": bind(source_manifest),
        "source_file_manifest": bind(source_file_manifest),
        "source_proof": bind(source_proof),
        "review_context": {**bind(review_context), "status": "blocked_prelaunch", "role": "prior independent blocker receipt; fresh pass required in GO"},
        "correction_receipt": {**bind(correction_receipt), "status": "pass"},
        "source": {
            "baseline_root": "/home/zhuhe/litchi-performance-0969/source-baseline",
            "manifest": bind(source_manifest),
            "file_manifest": bind(source_file_manifest),
            "content_proof": bind(source_proof),
            "harness_source": bind(harness_source),
            "harness_cargo": bind(harness_cargo),
            "interpretation": "source custody only; no current-HEAD or candidate executable is admitted",
        },
        "binary": {
            **bind(binary),
            "role": "retained_0967_historical_baseline",
            "build_id": p.BINARY_BUILD_ID,
            "executable": True,
        },
        "fixture": {**bind(fixture), "case": "docfloat", "edits": 1, "retention": "default", "harness_retention": "default_8mib"},
        "marker": {
            "symbol": p.MARKER,
            "dso_path": str(binary.absolute()),
            "dso_basename": binary.name,
            "build_id": p.BINARY_BUILD_ID,
            "address_range": "0xec4f0..0xec56d",
            "source": "tools/perf-cfb-profile/src/lib.rs:338",
            "calls": p.CALLS,
            "marker_scope": "5 warmups + 80 measured calls combined",
            "p50_scope": "80 measured samples only",
            "lifecycle_scope": "85-call lifecycle subset; marker owner percentage is not exact timed-region attribution",
            "missing_membership": "unknown",
            "unresolved_ancestry": "reported descriptively",
        },
        "matrix": {
            "cells": [{"case": "docfloat", "edits": 1, "retention": "default"}],
            "children": p.CHILDREN,
            "serial": True,
            "serial_order": list(range(1, p.CHILDREN + 1)),
            "cpu": p.CPU,
            "warmups": p.WARMUPS,
            "samples": p.SAMPLES,
            "first_child_is_pilot": True,
        },
        "perf": {
            "path": str(perf.resolve()),
            "sha256": p.sha256(perf),
            "event": "instructions:u",
            "period": p.PERIOD,
            "callgraph": "dwarf,16384",
            "sample_flags": ["--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache", "--user-callchains"],
            "decode_no_inline": ["--show-lost-events", "--no-inline", "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso"],
            "decode_inline": ["--show-lost-events", "--inline"],
            "buildid_export": ["buildid-list", "-i", "<child>.perf.data"],
            "whole_process_scope": "perf-launched child process tree (Python launcher and execed baseline); excludes perf bookkeeping",
            "timed_scope": "harness elapsed_ns only; never conflated with perf whole-process counts",
        },
        "report_contract": {"case": "docfloat", "edits": 1, "retention": "default_8mib", "warmups": 5, "samples_requested": 80, "sample_indices": list(range(80)), "elapsed_ns_identity": "samples[i].elapsed_ns == elapsed_ns[i]"},
        "raw_artifacts": ["report", "stderr", "perf_data", "launch", "command", "terminal", "no_inline", "inline", "header", "buildid"],
        "environment": {"LC_ALL": "C", "LANG": "C", "TZ": "UTC"},
        "launcher": {**bind(launcher), "schema": p.LAUNCH_SCHEMA, "execve_required": True, "requested_affinity": [p.CPU], "actual_affinity": [p.CPU]},
        "python": bind(python),
        "taskset": bind(taskset),
        "validator": {**bind(validator), "schema": "litchi.performance.0969.protocol_common.v4", "report_schema": "litchi.performance.0917.doc_retention.v1"},
        "helpers": {
            "protocol": bind(TOOLS / "protocol_common_0971_v2.py"),
            "runner": bind(TOOLS / "run_0971_v2.py"),
            "reducer": bind(TOOLS / "reduce_0971_v2.py"),
            "plan_generator": bind(TOOLS / "prepare_plan_0971_v2.py"),
        },
        "helper_seal": bind(helper_seal),
        "helpers_seal_sha256": p.sha256(helper_seal),
        "pilot_gate": {"child": 1, "min_marker_samples": 100, "lost_events": 0, "marker_period_share_positive": True, "on_failure": "stop_without_retry"},
        "execution": {"serial": True, "retry": False, "timeout_seconds": 1800, "quiescence_required": True, "no_build": True, "no_cargo": True, "no_other_workload": True},
        "outputs": {"capture_root": str((ROOT / "capture/instruction-profile-v2").resolve()), "analysis_root": str((ROOT / "analysis/instruction-profile-v2").resolve()), "raw_retained": True, "exclusive_artifacts": True},
        "interpretation": {
            "diagnostic_only": True,
            "no_speed_claim": True,
            "no_adoption_decision": True,
            "all_12_children_required": True,
            "marker_correlations": "descriptive across all 12; no fast/slow subset",
            "marker_membership_missing": "unknown",
            "unresolved_and_truncated": "reported and never silently qualified",
        },
    }
    p.check_plan_bindings(plan, plan_path)
    p.write_json_exclusive(plan_path, plan)
    print(plan_path)
    print(p.sha256(plan_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
