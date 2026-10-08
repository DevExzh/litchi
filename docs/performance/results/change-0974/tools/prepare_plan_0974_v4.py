#!/usr/bin/env python3
"""Create the external 0974 v4 helper seal or per-thread pre-capture plan."""
from __future__ import annotations

import argparse
import copy
import json
import sys
from pathlib import Path

import protocol_adapter_0974_v3 as a

ROOT = Path("/home/zhuhe/litchi-performance-0974")
ADR = ROOT / "receipts/adr-goal-revalidation-0974-v1.json"
OLD_PLAN = ROOT / "plans/plan-0974-pre-capture-v1.json"


def binding(path: str | Path) -> dict:
    return a.file_binding(path)


def helper_rows() -> dict[str, dict]:
    return {
        "protocol": binding(ROOT / "tools/protocol_adapter_0974_v3.py"),
        "runner": binding(ROOT / "tools/run_0974_v3.py"),
        "reducer": binding(ROOT / "tools/reduce_0974_v4.py"),
        "plan_generator": binding(ROOT / "tools/prepare_plan_0974_v4.py"),
        "raw_stack": binding(ROOT / "tools/protocol_adapter_0974_v3.py"),
        "shim_source": binding(a.V1.RETAINED_ROOT / "tools-v2/dwfl_dynamic_shim_v2.c"),
        "shim_map": binding(a.V1.RETAINED_ROOT / "tools-v2/dwfl_dynamic_shim_v2.map"),
        "shim_so": binding(a.V1.RETAINED_ROOT / "replay-dwfl-v2/dwfl_dynamic_shim_v2.so"),
    }


def make_helper_seal() -> dict:
    return {
        "schema": a.HELPER_SCHEMA,
        "status": "sealed",
        "version": "0974-v4-reducer-correction",
        "created_utc": a.V1.utc(),
        "scope": "v1 retained; native perf argv adds --per-thread and -m 16M; raw task/lifecycle audit is fail-closed",
        "helpers": helper_rows(),
        "retained_v1_helper_seal": binding(a.V1_HELPER_SEAL),
        "retained_v1_plan": binding(a.V1_PLAN),
        "retained_0973": {name: a.V1.retained_binding(name) for name in a.V1.RETAINED_FILES},
        "record_mode": {"tokens": ["--per-thread"], "inheritance": "disabled_by_per_thread", "forbidden_tokens": ["-a", "--all-cpus", "-C", "--cpu", "--inherit"]},
        "record_memory": {"tokens": a.MEMORY_TOKENS, "bytes": a.MEMORY_BYTES, "position": "immediately before --call-graph"},
        "task_scope": "same launcher/native PID and one thread; raw task stream is lifetime authority",
        "loss_contract": {"raw_field": a.RAW_LOST_FIELD, "terminal_field": a.TERMINAL_LOST_FIELD, "double_counting": "the two zero gates are not added"},
    }


def make_plan(helper_seal_path: Path, output: Path) -> dict:
    old = a.load_json(OLD_PLAN)
    seal = a.load_json(helper_seal_path)
    if seal.get("schema") != a.HELPER_SCHEMA or seal.get("status") != "sealed" or seal.get("helpers") != helper_rows():
        raise a.ProtocolError("v4 helper files changed since seal")
    plan = copy.deepcopy(old)
    plan.update({
        "schema": a.PLAN_SCHEMA,
        "status": "pre_capture",
        "created_utc": a.V1.utc(),
        "plan_path": str(output.resolve()),
        "design": binding(a.DESIGN_PATH),
        "helper_seal": binding(helper_seal_path),
        "helpers_seal_sha256": a.sha256(helper_seal_path),
        "helpers": seal["helpers"],
        "outputs": {"capture_root": str((ROOT / "capture/instruction-profile-v1").resolve()), "analysis_root": str((ROOT / "analysis").resolve())},
        "capability_probe": binding(a.CAPABILITY_PATH),
        "record_spec": binding(a.RECORD_SPEC_PATH),
        "failed_mmap_probe": binding(a.FAILED_MMAP_PATH),
        "failed_cpu_probe": binding(a.FAILED_CPU_PATH),
        "lifecycle_control": binding(a.LIFECYCLE_CONTROL_PATH),
        "capability_lineage": {"successful_per_thread": binding(a.CAPABILITY_PATH), "prior_failed_implicit_mmap": binding(a.FAILED_MMAP_PATH), "prior_failed_cpu31_mmap": binding(a.FAILED_CPU_PATH), "no_fallback": True},
        "interpretation": {"diagnostic_only": True, "qualification_claim": False, "speed_claim": False, "exact_timed_region_claim": False, "marker_period_share": "raw decoder coverage only", "fresh_cohort": True, "not_a_retry_of": "/home/zhuhe/litchi-performance-0973", "native_argv_deltas_only": ["--per-thread", "-m 16M"], "per_thread_scope": "single launcher/native PID and one thread; raw task stream gate"},
    })
    plan["perf"] = copy.deepcopy(plan["perf"])
    plan["perf"]["record_memory"] = {"tokens": a.MEMORY_TOKENS, "bytes": a.MEMORY_BYTES, "position": "immediately before --call-graph"}
    plan["perf"]["record_mode"] = {"tokens": ["--per-thread"], "inheritance": "disabled_by_per_thread", "forbidden_tokens": ["-a", "--all-cpus", "-C", "--cpu", "--inherit"], "scope": "one launcher/native exec PID and one native thread; task/FORK audit required per child"}
    plan["perf"]["cpu_selector"] = {"tokens": [], "affinity_scope": "outer taskset --cpu-list 31"}
    plan["raw_stack_contract"] = {"schema": a.BASE.RAW_STACK_SCHEMA, "sample_type": "IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER", "period": a.BASE.PERIOD, "stack_cap": a.BASE.STACK_CAP, "bounds": "0 <= dyn_size <= captured_stack_size <= 32768", "missing_or_short": "missing fields reject; bounded short captured_stack_size is retained in its distribution", "raw_scope": "all PERF_RECORD_SAMPLE records without misc filtering", "raw_lost_field": a.RAW_LOST_FIELD, "raw_lost_scope": "sum of PERF_RECORD_LOST lost fields; PERF_RECORD_LOST_SAMPLES is not added", "task_stream_field": "task_stream"}
    plan["terminal_contract"] = {"field": a.TERMINAL_LOST_FIELD, "record_type": "PERF_RECORD_LOST_SAMPLES", "required_value": 0, "source": "perf report --stdio terminal summary", "raw_loss_not_added": True}
    plan["command_receipt_contract"] = {"environment": {"LD_PRELOAD": "absent", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}, "cwd": "record inherited runner cwd before launch", "no_secret_environment": True}
    plan["task_stream_contract"] = {"header_required": ["task = 1", "comm = 1", "comm_exec = 1", "enable_on_exec = 1", "sample_stack_user = 32768", "read_format = ID|LOST"], "inheritance": "inherit field absent or zero; inherit=1 rejects", "thread_map": "exactly nr=1 with expected launcher/native PID", "COMM": "same PID/TID only; require perf-exec, python3.14, and native binary identity", "FORK": "none", "CLONE": "none", "EXIT": "main PID/TID only; descendant exit rejects", "CPU_MAP": "record descriptively; sample CPU31 remains strict", "raw_authority": "zero-loss raw task stream is lifetime authority"}
    plan["proc_task_contract"] = {"scope": "only /proc/<exec_pid>/task", "snapshots": ["before_packet_decode", "after_packet_decode"], "absence": "explicit post-exit absence allowed; absence does not prove history", "present": "task_ids must equal [exec_pid]", "permission_error": "explicit error and fail closed", "snapshot_scope": "bounded post-exit observations; not live execution history", "full_proc_scan": False}
    plan["batch_0974"] = {"fresh_cohort": True, "previous_cohort": "/home/zhuhe/litchi-performance-0973", "native_argv_deltas": ["--per-thread", "-m 16M"], "children": a.BASE.CHILDREN, "warmups": a.BASE.WARMUPS, "samples": a.BASE.SAMPLES, "cpu": a.BASE.CPU, "minimum_marker_samples_each_child": 100, "marker_period_share_positive_each_child": True, "required_raw_lost_sum": 0, "required_terminal_lost_samples": 0, "no_retry": True}
    return plan


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--seal-output", type=Path)
    ap.add_argument("--helper-seal", type=Path)
    ap.add_argument("--output", type=Path)
    args = ap.parse_args()
    try:
        if args.seal_output:
            if args.helper_seal or args.output:
                raise a.ProtocolError("--seal-output is standalone")
            out = args.seal_output.resolve()
            a.BASE.write_json_exclusive(out, make_helper_seal())
            print(json.dumps({"schema": a.HELPER_SCHEMA, "status": "pass", "path": str(out), "sha256": a.sha256(out)}, indent=2, sort_keys=True))
            return 0
        if args.helper_seal is None or args.output is None:
            raise a.ProtocolError("--helper-seal and --output are required")
        out = args.output.resolve()
        plan = make_plan(args.helper_seal.resolve(), out)
        a.check_plan_bindings(plan, None)
        a.BASE.write_json_exclusive(out, plan)
        a.check_plan_bindings(a.load_json(out), out)
        print(json.dumps({"schema": a.PLAN_SCHEMA, "status": "pass", "path": str(out), "sha256": a.sha256(out)}, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"0974 v4 plan generator rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
