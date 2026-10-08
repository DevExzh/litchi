#!/usr/bin/env python3
"""Seal and check the 0976 continuous-mode affinity control plan.

Only hashes and JSON custody records are produced.  This module never starts
perf or the stdlib target; the plan remains ``pre_review`` until independent
review and a root GO receipt bind it.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
import time
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0976")
TOOLS = ROOT / "tools"
PLAN_SCHEMA = "litchi.performance.0976.affinity_control_plan.v6"
SEAL_SCHEMA = "litchi.performance.0976.affinity_helper_seal.v6"
ADR = ROOT / "receipts/adr-goal-revalidation-0976-v1.json"
DESIGN = ROOT / "design-affinity-ordering-v1.md"
DESIGN_REVIEW = ROOT / "analysis/independent-affinity-ordering-review-v1.json"
PRIOR_REVIEW = ROOT / "analysis/independent-design-review-0976-v1.json"
PERF = Path("/usr/bin/perf")
PERF_MAN = Path("/usr/share/man/man1/perf-record.1.gz")
BRIDGE = TOOLS / "affinity_bridge_0976_v2.py"
TARGET = TOOLS / "stdlib_cpu_burn_0976_v2.py"
TASKSET = Path("/usr/bin/taskset")
PYTHON = Path("/usr/bin/python3.14")
GENERATOR = TOOLS / "prepare_control_plan_0976_v6.py"
CONTROL_ROOT = ROOT / "control/affinity-control-v6"


class PlanError(RuntimeError):
    pass


def sha256(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: str | Path) -> dict[str, Any]:
    resolved = Path(path).resolve()
    if not resolved.is_file():
        raise PlanError(f"missing binding: {resolved}")
    return {"path": str(resolved), "bytes": resolved.stat().st_size, "sha256": sha256(resolved)}


def utc_now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def load(path: str | Path) -> dict[str, Any]:
    try:
        value = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise PlanError(f"cannot load JSON {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise PlanError(f"JSON root is not an object: {path}")
    return value


def helper_rows() -> dict[str, dict[str, Any]]:
    return {
        "bridge": binding(BRIDGE),
        "stdlib_target": binding(TARGET),
        "control_plan_generator": binding(GENERATOR),
        "taskset": binding(TASKSET),
        "python": binding(PYTHON),
    }


def make_seal() -> dict[str, Any]:
    return {
        "schema": SEAL_SCHEMA,
        "status": "sealed",
        "created_utc": utc_now(),
        "scope": "0976 harmless continuous affinity handoff control only; no perf/workload execution",
        "helpers": helper_rows(),
        "protocol": {
            "mode": "continuous_event_mode",
            "target_cpu": 31,
            "recorder_cpu": 30,
            "all_recorder_tids": True,
            "parent_identity": ["pid", "uid", "starttime", "comm", "exe", "sha256"],
            "startup_samples": "retained and classified as startup/handoff; no FIFO filter",
            "native_settings": {
                "event": "instructions:u",
                "period": 1000000,
                "record_mode": ["--per-thread"],
                "record_memory": ["-m", "16M"],
                "callgraph": "dwarf,32768",
            },
        },
        "self_test_contract": {
            "perf_started": False,
            "workload_started": False,
            "identity_negative_control": True,
            "affinity_negative_control": True,
            "periodic_parent_checks": True,
        },
    }


def output_paths() -> dict[str, str]:
    root = CONTROL_ROOT.resolve()
    return {
        "control_root": str(root),
        "perf_data": str((root / "control.perf.data").resolve()),
        "perf_header": str((root / "perf-header.txt").resolve()),
        "perf_record_stdout": str((root / "perf-record.stdout").resolve()),
        "perf_record_stderr": str((root / "perf-record.stderr").resolve()),
        "bridge_receipt": str((root / "bridge-receipt.json").resolve()),
        "target_receipt": str((root / "target-receipt.json").resolve()),
        "command_receipt": str((root / "command-receipt.json").resolve()),
        "terminal_receipt": str((root / "terminal-receipt.json").resolve()),
    }


def future_paths_are_fresh(paths: dict[str, str]) -> None:
    for name, value in paths.items():
        if name != "control_root" and Path(value).exists():
            raise PlanError(f"fresh control output already exists: {name}={value}")


def command(paths: dict[str, str]) -> list[str]:
    return [
        str(TASKSET),
        "--cpu-list",
        "31",
        str(PERF),
        "record",
        "--per-thread",
        "-e",
        "instructions:u",
        "-c",
        "1000000",
        "-m",
        "16M",
        "--call-graph",
        "dwarf,32768",
        "--period",
        "--timestamp",
        "--sample-cpu",
        "--buildid-all",
        "--no-buildid-cache",
        "--user-callchains",
        "-o",
        paths["perf_data"],
        "--",
        str(PYTHON),
        str(BRIDGE),
        "--expected-perf-exe",
        str(PERF),
        "--expected-perf-sha256",
        sha256(PERF),
        "--expected-interpreter-exe",
        str(PYTHON),
        "--expected-interpreter-sha256",
        sha256(PYTHON),
        "--target-cpu",
        "31",
        "--recorder-cpu",
        "30",
        "--bridge-receipt",
        paths["bridge_receipt"],
        "--target-exe",
        str(PYTHON),
        "--target-script",
        str(TARGET),
        "--target-receipt",
        paths["target_receipt"],
        "--",
        "--seconds",
        "0.25",
    ]


def make_plan(seal_path: Path, plan_path: Path) -> dict[str, Any]:
    paths = output_paths()
    future_paths_are_fresh(paths)
    seal = load(seal_path)
    if seal.get("schema") != SEAL_SCHEMA or seal.get("status") != "sealed" or seal.get("helpers") != helper_rows():
        raise PlanError("helper seal is stale or has the wrong schema")
    return {
        "schema": PLAN_SCHEMA,
        "status": "pre_review",
        "created_utc": utc_now(),
        "plan_path": str(plan_path.resolve()),
        "scope": "one harmless stdlib CPU-burn continuous affinity control; no Office/native workload, build, or qualification",
        "design": binding(DESIGN),
        "design_review": binding(DESIGN_REVIEW),
        "prior_design_review": binding(PRIOR_REVIEW),
        "adr_goal_revalidation": binding(ADR),
        "helper_seal": binding(seal_path),
        "helpers": seal["helpers"],
        "perf": binding(PERF),
        "perf_manual": binding(PERF_MAN),
        "toolchain": {"taskset": binding(TASKSET), "python": binding(PYTHON)},
        "outputs": paths,
        "fresh_output_contract": "all files below control_root are created exclusively after root GO; no overwrite or retry",
        "command": command(paths),
        "environment": {
            "whitelist": {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "TZ": "UTC", "PYTHONDONTWRITEBYTECODE": "1"},
            "LD_PRELOAD": "absent",
            "secrets": "not copied into child environment",
        },
        "topology": {
            "mode": "continuous_event_mode",
            "outer_affinity": [31],
            "bridge_affinity_before_and_after": [31],
            "target_affinity_before_exec_and_after_exec": [31],
            "perf_parent_affinity_before_handoff": [31],
            "perf_parent_affinity_after_handoff": [30],
            "perf_tids": "enumerate /proc/<perf_pid>/task and move/verify every TID; leader-only move rejects",
            "same_pid": "bridge execs /usr/bin/python3.14 target; target PID/starttime equals bridge PID/starttime",
            "parent_identity": "exact PID plus UID/starttime, comm=perf, resolved /usr/bin/perf path, and bound SHA-256",
        },
        "handoff_sequence": [
            "outer taskset starts perf and bridge on CPU31 with continuous enable_on_exec",
            "bridge verifies own CPU31 and exact perf parent PID/UID/starttime/comm/exe",
            "bridge snapshots every perf TID and verifies each starts on CPU31",
            "bridge moves every observed perf TID to CPU30 and reads every mask back",
            "bridge rejects any TID set/starttime change or mask mismatch",
            "bridge rechecks itself and every perf TID immediately before same-PID target exec",
            "bridge execs the stdlib target without changing PID/TID",
            "target periodically rechecks its CPU31 mask and every perf TID CPU30 mask",
        ],
        "startup_semantics": {
            "events_continuously_enabled": True,
            "enable_on_exec": True,
            "startup_handoff_samples": "retained; classify and report startup/handoff records rather than filter",
            "fifo_mode": "optional diagnostic only; not part of this plan",
            "target_sample_cpu_gate": "all admitted target samples must still be CPU31",
        },
        "target": {
            "kind": "stdlib_only_cpu_burn",
            "path": str(TARGET.resolve()),
            "interpreter": binding(PYTHON),
            "seconds": 0.25,
            "check_interval_seconds": 0.02,
            "cpu": 31,
            "recorder_cpu": 30,
            "no_office_or_baseline_binary": True,
            "target_receipt": paths["target_receipt"],
        },
        "forbidden_perf_options": ["--threads", "--aio", "--affinity", "-C", "--cpu", "-a", "--all-cpus", "BPF sideband", "--off-cpu"],
        "control_gates": {
            "all_perf_tids_moved_and_verified": True,
            "parent_identity_exact": True,
            "target_cpu_exact": 31,
            "sample_cpu_exact": 31,
            "positive_decoded_sample_count": "> 0",
            "positive_period_count": "> 0",
            "sample_period_exact": 1000000,
            "same_pid_and_tid": True,
            "periodic_parent_checks": True,
            "no_unexpected_fork_clone_descendant": True,
            "raw_perf_record_lost_sum": 0,
            "terminal_perf_record_lost_samples": 0,
            "raw_loss_domains_are_separate": True,
            "stop_on_first_failure": True,
            "no_retry": True,
        },
        "future_native_gates": {
            "minimum_marker_samples_each_child": 100,
            "positive_marker_period_share": True,
            "scope": "future Office/native cohort only; impossible and inapplicable to this stdlib control",
        },
        "periodic_observation_window": {
            "interval_seconds": 0.02,
            "scope": "start, periodic, and end observations of parent identity/TID masks",
            "claim": "all recorded observations passed",
            "not_proof_of_continuous_absence": True,
            "between_poll_transient": "outside this control proof",
            "native_requirement": "final identity/mask checks plus decoded sample, loss, and task-scope gates remain required",
        },
        "negative_controls": {
            "offline_self_test": "identity starttime mismatch and recorder mask mismatch must reject",
            "runtime_identity_negative": "not run before review/GO; mismatched parent must produce failed_pre_exec receipt and no target exec",
        },
        "admission": {
            "independent_review_required": True,
            "root_go_required": True,
            "native_workload_started": False,
            "control_started": False,
            "build_started": False,
            "qualification_claim": False,
            "speed_claim": False,
        },
    }


def write_exclusive(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def check(plan_path: Path) -> dict[str, Any]:
    plan = load(plan_path)
    if plan.get("schema") != PLAN_SCHEMA or plan.get("status") != "pre_review":
        raise PlanError("plan schema/status mismatch")
    if plan.get("plan_path") != str(plan_path.resolve()):
        raise PlanError("plan self-binding mismatch")
    seal_path = Path(plan["helper_seal"]["path"]).resolve()
    if plan["helper_seal"] != binding(seal_path):
        raise PlanError("helper seal binding mismatch")
    seal = load(seal_path)
    if seal.get("schema") != SEAL_SCHEMA or seal.get("status") != "sealed" or seal.get("helpers") != helper_rows():
        raise PlanError("helper seal schema or helpers changed")
    if plan.get("helpers") != seal.get("helpers"):
        raise PlanError("plan/helper map mismatch")
    for name, row in plan["helpers"].items():
        if row != binding(row["path"]):
            raise PlanError(f"helper changed: {name}")
    for field, expected in (("design", DESIGN), ("design_review", DESIGN_REVIEW), ("prior_design_review", PRIOR_REVIEW), ("adr_goal_revalidation", ADR), ("perf", PERF), ("perf_manual", PERF_MAN)):
        if plan.get(field) != binding(expected):
            raise PlanError(f"binding mismatch: {field}")
    if plan.get("toolchain") != {"taskset": binding(TASKSET), "python": binding(PYTHON)}:
        raise PlanError("taskset/python toolchain binding mismatch")
    if plan.get("topology", {}).get("mode") != "continuous_event_mode":
        raise PlanError("continuous mode is not explicit")
    if plan.get("startup_semantics", {}).get("fifo_mode") != "optional diagnostic only; not part of this plan":
        raise PlanError("FIFO mode was accidentally adopted")
    if plan.get("command") != command(plan["outputs"]):
        raise PlanError("exact continuous command drift")
    if plan.get("outputs") != output_paths():
        raise PlanError("output path drift")
    control_gates = plan.get("control_gates", {})
    if control_gates.get("positive_decoded_sample_count") != "> 0" or control_gates.get("positive_period_count") != "> 0":
        raise PlanError("control positive sample/period gates missing")
    if "minimum_marker_samples_each_child" in control_gates or "positive_marker_period_share" in control_gates:
        raise PlanError("future marker gates leaked into harmless control gates")
    observation = plan.get("periodic_observation_window", {})
    if observation.get("interval_seconds") != 0.02 or observation.get("not_proof_of_continuous_absence") is not True:
        raise PlanError("periodic observation scope is not explicit")
    return {"schema": PLAN_SCHEMA, "status": "pass", "plan": str(plan_path.resolve()), "sha256": sha256(plan_path), "helper_seal": binding(seal_path), "control_started": False, "workload_started": False}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--seal-output", type=Path)
    parser.add_argument("--plan-output", type=Path)
    parser.add_argument("--check-plan", type=Path)
    args = parser.parse_args()
    try:
        if args.check_plan:
            if args.seal_output or args.plan_output:
                raise PlanError("--check-plan is standalone")
            print(json.dumps(check(args.check_plan.resolve()), indent=2, sort_keys=True))
            return 0
        if not args.seal_output or not args.plan_output:
            raise PlanError("--seal-output and --plan-output are required")
        seal_path = args.seal_output.resolve()
        plan_path = args.plan_output.resolve()
        write_exclusive(seal_path, make_seal())
        write_exclusive(plan_path, make_plan(seal_path, plan_path))
        print(json.dumps(check(plan_path), indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"0976 v6 control plan rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
