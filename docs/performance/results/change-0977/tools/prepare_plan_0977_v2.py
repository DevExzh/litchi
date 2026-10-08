#!/usr/bin/env python3
"""Prepare a fresh 0977 helper seal and pre-capture plan."""
from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import plan_binding_0977_v2 as b

ROOT = b.ROOT
TOOLS = ROOT / "tools"
DESIGN = ROOT / "design-0977-v1.md"
ADR_RECEIPT = ROOT / "receipts/adr-goal-revalidation-0977-v1.json"
DECODER_REVIEW = b.DECODER_REVIEW
NOOP_REVIEW = b.NOOP_REVIEW


def binding(path: Path) -> dict[str, Any]:
    return b.binding(path)


def dependency(path: Path, *, scope: str | None = None) -> dict[str, Any]:
    row = binding(path)
    if scope is not None:
        row["scope"] = scope
    return row


def load(path: Path) -> dict[str, Any]:
    value = b.strict_load(path)
    if not isinstance(value, dict):
        raise b.BindingError(f"{path}: expected object")
    return value


def make_helpers(decoder: Path, reducer: Path) -> dict[str, dict[str, Any]]:
    return {
        "bridge": binding(TOOLS / "affinity_bridge_0977.py"),
        "runner": binding(TOOLS / "run_0977_v2.py"),
        "plan_binding": binding(TOOLS / "plan_binding_0977_v2.py"),
        "plan_generator": binding(TOOLS / "prepare_plan_0977_v2.py"),
        "launcher": binding(b.LAUNCHER),
        "decoder": binding(decoder),
        "reducer": binding(reducer),
    }


def make_seal(decoder: Path, reducer: Path) -> dict[str, Any]:
    helpers = make_helpers(decoder, reducer)
    return {
        "schema": b.HELPER_SCHEMA,
        "status": "sealed",
        "version": "0977-v2",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "scope": "fresh serial Office capture bridge/runner/plan and decoder/reducer interface; no workload started",
        "helpers": helpers,
        "retained_lineage": {
            "affinity_bridge_0976": dependency(b.BRIDGE_LINEAGE, scope="handoff lineage; 0977 bridge is a fresh target-generic adaptation"),
            "decoder_final_v2": {
                "source": dependency(b.DECODER_SOURCE),
                "map": dependency(b.DECODER_MAP),
                "so": dependency(b.DECODER_SO),
                "validator": dependency(b.DECODER_VALIDATOR),
            },
        },
        "workload_started": False,
        "build_started": False,
    }


def make_plan(seal_path: Path, decoder: Path, reducer: Path, plan_path: Path) -> dict[str, Any]:
    seal = load(seal_path)
    helpers = make_helpers(decoder, reducer)
    if seal.get("helpers") != helpers:
        raise b.BindingError("helper seal does not match current helper files")
    dependencies = {
        "perf": binding(b.PERF),
        "taskset": binding(b.TASKSET),
        "python": binding(b.PYTHON),
        "launcher": binding(b.LAUNCHER),
        "bridge_lineage": dependency(b.BRIDGE_LINEAGE, scope="handoff lineage; 0977 bridge is a fresh target-generic adaptation"),
        "decoder_so": binding(b.DECODER_SO),
        "decoder_source": binding(b.DECODER_SOURCE),
        "decoder_map": binding(b.DECODER_MAP),
        "decoder_validator": binding(b.DECODER_VALIDATOR),
        "libdw": binding(b.LIBDW),
        "libelf": binding(b.LIBELF),
        "decoder_header": binding(b.HEADER),
        "control_target": binding(b.CONTROL_TARGET),
    }
    aslr_review = DECODER_REVIEW
    noop_receipt = NOOP_REVIEW
    plan = {
        "schema": b.PLAN_SCHEMA,
        "status": "pre_capture",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "plan_path": str(plan_path.resolve()),
        "design": binding(DESIGN),
        "adr_goal_revalidation": binding(ADR_RECEIPT),
        "capability_probe": binding(b.CAPABILITY),
        "perf_requalification": binding(b.PERF_REQUALIFICATION_SOURCE),
        "perf_requalification_posthoc": binding(b.PERF_REQUALIFICATION),
        "perf_tool_drift": binding(b.PERF_TOOL_DRIFT),
        "source_manifest": binding(b.SOURCE_MANIFEST),
        "source_file_manifest": binding(b.SOURCE_FILE_MANIFEST),
        "binary": binding(b.BINARY),
        "fixture": binding(b.FIXTURE),
        "python": binding(b.PYTHON),
        "taskset": binding(b.TASKSET),
        "launcher": binding(b.LAUNCHER),
        "dependencies": dependencies,
        "helpers": helpers,
        "helper_seal": binding(seal_path),
        "helper_seal_sha256": b.sha256(seal_path),
        "decoder_reviews": {
            "final_v2_review": binding(aslr_review),
            "noop_posthoc": binding(noop_receipt),
            "focused_controls": binding(b.FOCUSED_CONTROLS),
        },
        "batch": {
            "children": b.CHILDREN, "serial": True, "cpu": b.CPU, "recorder_cpu": b.RECORDER_CPU,
            "period": b.PERIOD, "stack_cap": b.STACK_CAP, "warmups": b.WARMUPS, "samples": b.SAMPLES,
            "minimum_marker_samples": 100, "marker_period_positive": True, "raw_lost_sum": 0,
            "terminal_lost_samples": 0, "empty_callchain_count": 0, "stop_on_first_failure": True, "retry": False,
        },
        "execution": {"children": b.CHILDREN, "serial": True, "retry": False, "timeout_seconds": b.TIMEOUT_SECONDS,
                       "stop_on_first_failure": True, "quiescence_required": True},
        "perf": {
            "path": str(b.PERF.resolve()), "sha256": b.sha256(b.PERF), "event": "instructions:u", "period": b.PERIOD,
            "callgraph": "dwarf,32768", "stack_cap": b.STACK_CAP, "record_mode": ["--per-thread"],
            "record_memory": ["-m", "16M"],
            "sample_flags": ["--period", "--timestamp", "--sample-cpu", "--buildid-all", "--no-buildid-cache", "--user-callchains"],
        },
        # The current perf's report header is retained for provenance only:
        # its known HEADER_EVENT_DESC warning can omit the descriptor.  The
        # decoder's independent ``evlist -v`` record is the admission source
        # for the event attributes below.
        "perf_contract": b.PERF_CONTRACT,
        "environment": b.environment(),
        "workload": {
            "case": "docfloat", "edits": 1, "retention": "default",
            "target_argv_template": ["<binary>", "--case", "docfloat", "--input", "<fixture>", "--edits", "1", "--retention", "default", "--warmup", "5", "--samples", "80"],
            "same_binary_fixture_all_children": True,
            "same_pid_exec_chain": ["perf", "affinity_bridge_0977.py", "launch_baseline_v2.py", "bound_workload"],
        },
        "control": {
            "target_script": binding(b.CONTROL_TARGET),
            "seconds": 0.25,
            "check_interval": 0.02,
            "output_root": str((ROOT / "capture/affinity-control-v1").resolve()),
            "scope": "harmless stdlib target through the same fresh 0977 bridge/Popen boundary; no Office marker decoder",
            "raw_gates": {"raw_lost_sum": 0, "terminal_lost_samples": 0, "cpu": b.CPU, "single_pid": True},
        },
        "marker": {
            "build_id": "e5c457384eeebd057932c444efda77618f137399",
            "calls": 85,
            "dso_path": str(b.BINARY.resolve()),
            "marker_scope": "5 warmups + 80 measured calls combined",
            "owner_scope": "all 12 children; missing owner membership unknown",
            "p50_scope": "80 measured samples only",
            "symbol": "doc_retention_probe_0942_cfb_profile::profiled_measured_case",
        },
        "pilot_gate": {"child": 1, "lost_events": 0, "marker_period_share_positive": True,
                        "min_marker_samples": 100, "on_failure": "stop_without_retry"},
        "child_gate": {"all_children": True, "raw_lost_sum": 0, "terminal_lost_samples": 0,
                        "min_marker_samples": 100, "marker_period_share_positive": True,
                        "empty_callchain_count": 0, "on_failure": "stop_without_retry"},
        "report_contract": {
            "case": "docfloat", "edits": 1, "retention": "default_8mib", "warmups": 5,
            "samples_requested": 80, "sample_indices": list(range(80)),
            "elapsed_ns_identity": "samples[i].elapsed_ns == elapsed_ns[i]",
        },
        "raw_stack_contract": {
            "sample_type": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
            "period": b.PERIOD, "stack_cap": b.STACK_CAP, "bounds": "0 <= dyn_size <= captured_stack_size <= 32768",
            "raw_lost_field": "raw_lost_sum", "terminal_lost_field": "terminal_lost_samples",
            "raw_scope": "all PERF_RECORD_SAMPLE records without misc filtering",
            "task_stream_authority": "zero-loss raw task stream is lifetime authority",
        },
        "terminal_contract": {
            "field": "PERF_RECORD_LOST_SAMPLES",
            "required_value": 0,
            "source": "decoder-owned offline perf report --stdio",
            "runner_precomputes": False,
            "raw_loss_not_added": True,
        },
        "task_stream_contract": {
            "header_required": ["task = 1", "comm = 1", "comm_exec = 1", "enable_on_exec = 1", "sample_stack_user = 32768", "read_format = ID|LOST"],
            "inheritance": "inherit absent or zero; inherit=1 rejects",
            "thread_map": "exactly one expected target PID/TID",
            "COMM": "perf-exec, python3.14, launcher, and bound binary may share one PID/TID",
            "FORK": "none", "CLONE": "none", "EXIT": "main PID/TID only",
            "sample_gate": "every admitted sample has expected PID/TID and CPU31",
            "raw_authority": "raw zero-loss task stream; bounded /proc observations are descriptive",
        },
        "bridge_handoff": {
            "target_cpu": b.CPU, "recorder_cpu": b.RECORDER_CPU, "parent_all_tids": True,
            "same_pid_exec": True, "startup_handoff_samples_retained": True,
            "target_side_stdlib_checks": False,
            "target_identity": "0970 launcher receipt and raw COMM/THREAD_MAP/EXIT records",
        },
        "decoder_interface": {
            "function": "decode_child(plan, child_dir, terminal, launch)",
            "returns": "complete packet with summary and gate fields; raises DecodeFailure on malformed or unbound input",
            "offline_only": True, "native_ld_preload": "absent",
            "filenames": ["<child>.perf-header.txt", "<child>.perf-mmap-dump.txt", "<child>.target-layout.json", "<child>.raw-stack.json", "<child>.corrected.perf-script-no-inline.txt", "<child>.corrected.perf.stderr", "<child>.corrected.dwfl.jsonl", "<child>.decoder.json"],
            "native_runner_filenames": ["<child>.perf.data", "<child>.report.json", "<child>.launch.json", "<child>.command.json", "<child>.terminal.json", "<child>.perf.stderr", "<child>.bridge.json", "<child>.affinity-observations.json"],
            "decoder_terminal_report": "decoder invokes perf report --stdio offline and binds the exact terminal LOST_SAMPLES text; runner does not precompute it",
            "decoder_owned_receipts": ["<child>.terminal-report.txt", "<child>.terminal-summary.json", "<child>.loss.json"],
            "modes": ["corrected_final_v2_only"],
        },
        "outputs": {"capture_root": str((ROOT / "capture/instruction-profile-v2").resolve()), "analysis_root": str((ROOT / "analysis").resolve())},
        "native_environment": {"LC_ALL": "C", "LANG": "C", "TZ": "UTC", "LD_PRELOAD": "absent"},
        "command_receipt_contract": {"cwd": str(b.REPO.resolve()), "no_secret_environment": True, "LD_PRELOAD": "absent"},
        "interpretation": {"fresh_cohort": True, "not_a_retry_of": "/home/zhuhe/litchi-performance-0975", "diagnostic_only_until_reduction": True, "no_workload_started": True},
        "no_build_or_workload": True,
    }
    return plan


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--seal-output", type=Path)
    ap.add_argument("--helper-seal", type=Path)
    ap.add_argument("--decoder", type=Path, required=True)
    ap.add_argument("--reducer", type=Path, required=True)
    ap.add_argument("--output", type=Path)
    args = ap.parse_args()
    try:
        decoder, reducer = args.decoder.resolve(), args.reducer.resolve()
        if args.seal_output:
            if args.helper_seal or args.output:
                raise b.BindingError("--seal-output is standalone")
            b.write_exclusive(args.seal_output.resolve(), make_seal(decoder, reducer))
            print(json.dumps({"status": "pass", "path": str(args.seal_output.resolve()), "sha256": b.sha256(args.seal_output)}, sort_keys=True))
            return 0
        if args.helper_seal is None or args.output is None:
            raise b.BindingError("--helper-seal, --decoder, --reducer, and --output are required")
        plan = make_plan(args.helper_seal.resolve(), decoder, reducer, args.output.resolve())
        b.validate_seal(plan, args.output.resolve())
        b.write_exclusive(args.output.resolve(), plan)
        b.validate_plan(args.output.resolve())
        print(json.dumps({"status": "pass", "path": str(args.output.resolve()), "sha256": b.sha256(args.output)}, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"0977 plan generator rejected: {exc}", file=__import__("sys").stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
