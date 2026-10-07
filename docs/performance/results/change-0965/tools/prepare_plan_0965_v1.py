#!/usr/bin/env python3
"""Prepare the hash-bound prospective 0965 ingress diagnostic plan.

This helper only reads the inherited 0964 plan, retained source/build
receipts, retained executables, the diagnostic runner, its reducer, and the
affinity-control receipt.  It writes one immutable plan.  It never edits a
source tree, binary, or capture packet and never starts a workload.
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0965").resolve()
BASE_PLAN = Path("/home/zhuhe/litchi-performance-0964/protocol/plan-final-epoch2-v1.json").resolve()
BASE_AUDIT = Path("/home/zhuhe/litchi-performance-0964/review/independent-epoch2-raw-audit-v1.json").resolve()
DESIGN = ROOT / "diagnostic-design-v3.md"
RUNNER = ROOT / "tools/run_ingress_diagnostic_0965_v1.py"
REDUCER = ROOT / "tools/reduce_ingress_0965_v1.py"
CONTROLS = ROOT / "protocol/affinity-controls-0965-v5.json"
HOST_BEFORE = ROOT / "protocol/host-before-v1.json"
OUTPUT = ROOT / "protocol/plan-diagnostic-0965-v6.json"
CPU_MASK = [12, 13, 14, 15]
DEADLINE = 900
WARMUP = 20
SAMPLES = 256
CASES = ("medium-one-percent", "dense-sparse-one-percent", "medium-one-edit")
PAIR_SCHEDULE = {
    "medium-one-percent": ("RR", "RC", "CC", "CR"),
    "dense-sparse-one-percent": ("RC", "CC", "CR", "RR"),
    "medium-one-edit": ("CC", "CR", "RR", "RC"),
}
EXPECTED_BASELINE_SHA = "5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3"
EXPECTED_CANDIDATE_SHA = "d21863a6613800391a01914865fc4eaecdf54ad3061447be8b6be7d47d0dc910"
EXPECTED_DESIGN_SHA = "0dbff34e930aec3d63350ebf6cce9d0c8a86afdcd8f176f50e52a7e6fe2c06bf"


class PlanError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise PlanError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as exc:
        raise PlanError(f"cannot hash {path}: {exc}") from exc
    return digest.hexdigest()


def descriptor(path: Path, executable: bool | None = None) -> dict[str, Any]:
    path = path.expanduser().resolve()
    require(path.is_file() and not path.is_symlink(), f"regular file required: {path}")
    mode = path.stat().st_mode & 0o777
    is_executable = bool(mode & 0o111)
    if executable is not None:
        require(is_executable is executable, f"executable mismatch: {path}")
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "mode": mode,
        "executable": is_executable,
        "sha256": sha256(path),
    }


def directory(path: Path, label: str) -> str:
    path = path.expanduser().resolve()
    require(path.is_dir() and not path.is_symlink(), f"{label} directory missing: {path}")
    return str(path)


def read_json(path: Path) -> Any:
    path = path.expanduser().resolve()
    require(path.is_file() and not path.is_symlink(), f"JSON missing: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise PlanError(f"invalid JSON {path}: {exc}") from exc


def checked_descriptor(path: Path, expected: dict[str, Any], label: str, executable: bool | None = None) -> dict[str, Any]:
    actual = descriptor(path, executable)
    for key in ("path", "bytes", "mode", "executable", "sha256"):
        if key in expected:
            require(actual[key] == expected[key], f"{label}.{key} changed")
    return actual


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="microseconds")


def build_rows(cases: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    case_by_id = {case["case_id"]: case for case in cases}
    require(tuple(case_by_id) == CASES, "inherited case order changed")
    children: list[dict[str, Any]] = []
    pairs: list[dict[str, Any]] = []
    for case_id in CASES:
        case = case_by_id[case_id]
        for pair_index, pair_type in enumerate(PAIR_SCHEDULE[case_id]):
            require(len(pair_type) == 2 and set(pair_type) <= {"R", "C"}, f"invalid pair {pair_type}")
            pair_id = f"{case_id}__pair{pair_index:02d}__{pair_type}"
            pair_children: list[dict[str, Any]] = []
            for position, role in enumerate(pair_type):
                binary_role = "baseline" if role == "R" else "candidate"
                row = {
                    "child_id": f"native__{pair_id}__p{position}",
                    "lane": "native",
                    "case_id": case_id,
                    "pair_id": pair_id,
                    "pair_type": pair_type,
                    "pair_index": pair_index,
                    "position": position,
                    "role": role,
                    "binary_role": binary_role,
                    "comparison_id": "retained-baseline-vs-retained-candidate",
                    "shape": case["shape"],
                    "workload": case["workload"],
                    "workers": 1,
                    "warmup_iterations": WARMUP,
                    "measured_samples": SAMPLES,
                    "cpu_mask": list(CPU_MASK),
                }
                children.append(row)
                pair_children.append(row)
            pairs.append({
                "pair_id": pair_id,
                "case_id": case_id,
                "pair_index": pair_index,
                "pair_type": pair_type,
                "children": [row["child_id"] for row in pair_children],
                "roles": list(pair_type),
                "adjacent": True,
            })
    require(len(children) == 24 and len(pairs) == 12, "0965 schedule is not 24 children/12 pairs")
    for case_id in CASES:
        case_pairs = [pair for pair in pairs if pair["case_id"] == case_id]
        require([pair["pair_type"] for pair in case_pairs] == list(PAIR_SCHEDULE[case_id]), f"pair rotation changed for {case_id}")
        require(sum(pair["pair_type"] == "RC" for pair in case_pairs) == 1, f"RC count changed for {case_id}")
        require(sum(pair["pair_type"] == "CR" for pair in case_pairs) == 1, f"CR count changed for {case_id}")
        require(sum(pair["pair_type"] == "RR" for pair in case_pairs) == 1, f"RR count changed for {case_id}")
        require(sum(pair["pair_type"] == "CC" for pair in case_pairs) == 1, f"CC count changed for {case_id}")
    return children, pairs


def custody(base: dict[str, Any], base_path: Path) -> dict[str, Any]:
    source = base["source"]
    quality = Path("/home/zhuhe/litchi-performance-0964/receipts/quality-0964-final-v1.json")
    binding = Path("/home/zhuhe/litchi-performance-0964/receipts/baseline-binding-v2.json")
    freeze = Path("/home/zhuhe/litchi-performance-0964/receipts/candidate-freeze-v4.json")
    build_result = Path("/home/zhuhe/litchi-performance-0964/receipts/candidate-managed-v3-build-result.json")
    lock_path = Path(source["baseline_source_root"]) / "tools/perf-baseline/Cargo.lock"
    candidate_perf_lock_path = Path(source["candidate_source_root"]) / "tools/perf-baseline/Cargo.lock"
    root_lock_path = Path(source["candidate_source_root"]) / "Cargo.lock"
    allowlist_path = Path("/home/zhuhe/litchi-performance-0964/protocol/candidate-test-allowlist-v3.json")
    baseline_manifest_path = Path(source["baseline_manifest"]["path"])
    candidate_manifest_path = Path(source["candidate_manifest"]["path"])
    base_binding = checked_descriptor(binding, {}, "baseline binding")
    base_freeze = checked_descriptor(freeze, {}, "candidate freeze")
    base_quality = checked_descriptor(quality, {}, "candidate quality")
    base_build = checked_descriptor(build_result, {}, "candidate build result")
    base_manifest = checked_descriptor(baseline_manifest_path, source["baseline_manifest"], "baseline manifest")
    candidate_manifest = checked_descriptor(candidate_manifest_path, source["candidate_manifest"], "candidate manifest")
    perf_lock = descriptor(lock_path)
    require(perf_lock["sha256"] == source["required_perf_baseline_lock_sha256"], "candidate perf-baseline Cargo.lock differs from required lock")
    candidate_perf_lock = descriptor(candidate_perf_lock_path)
    require(candidate_perf_lock["sha256"] == perf_lock["sha256"], "candidate perf-baseline Cargo.lock differs from baseline lock")
    root_lock = descriptor(root_lock_path)
    allowlist = descriptor(allowlist_path)
    return {
        "inherited_0964_plan": descriptor(base_path),
        "inherited_0964_plan_revision": base["plan_revision"],
        "inherited_0964_raw_audit": descriptor(BASE_AUDIT),
        "baseline_source_root": directory(Path(source["baseline_source_root"]), "baseline source"),
        "candidate_source_root": directory(Path(source["candidate_source_root"]), "candidate source"),
        "baseline_manifest": base_manifest,
        "candidate_manifest": candidate_manifest,
        "baseline_binding": base_binding,
        "candidate_freeze": base_freeze,
        "candidate_quality": base_quality,
        "candidate_build_result": base_build,
        "candidate_allowlist": allowlist,
        "required_perf_baseline_lock_sha256": source["required_perf_baseline_lock_sha256"],
        "required_perf_baseline_lock": perf_lock,
        "candidate_perf_baseline_lock": candidate_perf_lock,
        "candidate_root_lock_overlay": root_lock,
        "candidate_change_allowlist": list(source["candidate_change_allowlist"]),
        "production_adoption": False,
        "source_rebuilt": False,
    }


def binaries(base: dict[str, Any]) -> dict[str, Any]:
    native = base["lanes"]["native"]
    baseline = checked_descriptor(Path(native["baseline_binary"]["path"]), native["baseline_binary"], "baseline executable", True)
    candidate = checked_descriptor(Path(native["candidate_binary"]["path"]), native["candidate_binary"], "candidate executable", True)
    require(baseline["sha256"] == EXPECTED_BASELINE_SHA, "retained baseline hash changed")
    require(candidate["sha256"] == EXPECTED_CANDIDATE_SHA, "retained candidate hash changed")
    return {
        "R": baseline,
        "C": candidate,
        "roles": {
            "R": "retained 0957 baseline executable",
            "C": "retained 0964 candidate executable",
        },
    }


def controls(path: Path) -> dict[str, Any]:
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0965.affinity-launch-controls.v1", "affinity control schema changed")
    require(value.get("status") == "pass" and value.get("workload_started") is False, "affinity controls did not pass before workload")
    require(value.get("cpu_mask") == CPU_MASK, "affinity control mask changed")
    require(value.get("positive", {}).get("observed_affinity") == CPU_MASK, "positive affinity control changed")
    require(value.get("positive", {}).get("guard_passed") is True, "positive affinity guard missing")
    require(value.get("negative", {}).get("guard_rejected") is True, "negative affinity guard missing")
    require(value.get("positive", {}).get("cleanup", {}).get("reaped") is True, "positive control not reaped")
    require(value.get("negative", {}).get("cleanup", {}).get("reaped") is True, "negative control not reaped")
    return {
        "descriptor": descriptor(path),
        "schema": value["schema"],
        "status": value["status"],
        "workload_started": value["workload_started"],
        "cpu_mask": value["cpu_mask"],
        "runner": value.get("runner"),
    }


def make_plan(args: argparse.Namespace) -> dict[str, Any]:
    base_path = args.base_plan.expanduser().resolve()
    base = read_json(base_path)
    require(base.get("schema") == "litchi.performance.0964.inline-number-screen-plan.v1", "inherited plan schema changed")
    require(base.get("status") == "authorized" and base.get("plan_revision") == "epoch2-root-final-v1", "inherited 0964 plan is not final epoch2")
    require(base.get("scope", {}).get("capture_epoch") == "epoch2-affinity-hardened", "inherited 0964 epoch changed")
    require([case.get("case_id") for case in base.get("cases", [])] == list(CASES), "inherited case set changed")
    runner = descriptor(args.runner, True)
    reducer = descriptor(args.reducer, True)
    affinity = controls(args.affinity_controls)
    host_before = descriptor(args.host_before)
    children, pairs = build_rows(copy.deepcopy(base["cases"]))
    native_contract = copy.deepcopy(base["report_contract"]["native"])
    native_contract.update({"warmup_count": WARMUP, "warmup_iterations": WARMUP, "measured_samples": SAMPLES, "raw_sample_count": SAMPLES})
    design_descriptor = descriptor(DESIGN)
    require(design_descriptor["sha256"] == EXPECTED_DESIGN_SHA, "sealed diagnostic design v3 hash changed")
    source = custody(base, base_path)
    binary = binaries(base)
    root_go: dict[str, Any]
    if args.root_go is None:
        root_go = {"status": "pending_root_GO", "path": str((ROOT / "protocol/root-go-0965-v1.json").resolve())}
    else:
        go_path = args.root_go.expanduser().resolve()
        go = read_json(go_path)
        require(go.get("status") == "GO", "root GO status is not GO")
        root_go = {"status": "GO", "descriptor": descriptor(go_path), "scope": go.get("scope"), "bindings": go.get("bindings", [])}
    status = args.status
    if status == "authorized":
        require(root_go.get("status") == "GO", "authorized plan requires root GO")
    plan = {
        "schema": "litchi.performance.0965.ingress-diagnostic-plan.v1",
        "status": status,
        "plan_id": "xlsx-ingress-tail-diagnostic-0965",
        "plan_revision": "diagnostic-root-final-v1" if status == "authorized" else "diagnostic-prospective-v1",
        "created_utc": now(),
        "status_note": "Descriptive diagnostic only; does not reopen the 0964 screen or authorize adoption.",
        "purpose": "Classify retained-executable ingress-tail behavior with same-binary and launch-order controls.",
        "scope": {
            "format": "XLSX managed-compression benchmark",
            "lane": "native",
            "cases": list(CASES),
            "children": 24,
            "pairs": 12,
            "children_per_case": 8,
            "pair_types_per_case": ["RR", "RC", "CC", "CR"],
            "warmup_iterations": WARMUP,
            "measured_samples": SAMPLES,
            "fresh_process": True,
            "serial": True,
            "retry_or_replacement": False,
            "iwork_excluded": True,
            "source_rebuilt": False,
            "diagnostic_only": True,
        },
        "inherited_0964": {
            "plan": source["inherited_0964_plan"],
            "plan_revision": source["inherited_0964_plan_revision"],
            "raw_audit": source["inherited_0964_raw_audit"],
            "rejection_remains_binding": True,
            "safety_threshold_unchanged": 1.05,
        },
        "design": {
            "descriptor": design_descriptor,
            "schema": "prospective-design-v3",
            "status": "sealed",
        },
        "source": source,
        "binaries": binary,
        "cases": copy.deepcopy(base["cases"]),
        "report_contract": {
            "native": native_contract,
            "cleanup_resources": list(base["report_contract"]["cleanup_resources"]),
            "shared_oracles": list(base["report_contract"]["shared_oracles"]),
            "sink_max_write_bytes": base["report_contract"]["sink_max_write_bytes"],
        },
        "schedule": {
            "pair_order_by_case": {key: list(value) for key, value in PAIR_SCHEDULE.items()},
            "pairs": pairs,
            "children": children,
            "adjacent_pair_boundary": "second child starts only after first terminal receipt and process-group reaping",
        },
        "runner": {
            "descriptor": runner,
            "schema": "litchi.performance.0965.diagnostic-capture.v1",
            "cpu_mask": CPU_MASK,
            "deadline_seconds": DEADLINE,
            "preexec_affinity": "os.sched_setaffinity(0, {12,13,14,15})",
            "pid_affinity_getter": "os.sched_getaffinity(pid)",
            "terminal_before_parse": True,
        },
        "reducer": {
            "descriptor": reducer,
            "schema": "litchi.performance.0965.ingress-diagnostic-reduction.v1",
            "command": ["python3", str(reducer["path"]), "--plan", "<PLAN>", "--packet", "<PACKET>", "--output", "<OUTPUT>"],
            "quantile_method": "nearest-rank; rank=ceil(q*n), one-indexed",
            "n": SAMPLES,
            "p99_rank": 254,
            "adoption_decision": False,
        },
        "affinity_controls": affinity,
        "execution": {
            "cpu_mask": CPU_MASK,
            "deadline_seconds": DEADLINE,
            "max_parallel_children": 1,
            "retry_or_replacement": False,
            "terminal_before_parse": True,
            "taskset_path": "/usr/bin/taskset",
            "time_path": "/usr/bin/time",
            "time_argv_prefix": ["/usr/bin/time", "-v", "-o", "time-v.txt", "--"],
            "workload_cwd": base["lanes"]["native"]["execution"]["workload_cwd"],
            "environment": {
                "PATH": "/usr/bin:/bin",
                "LANG": "C",
                "LC_ALL": "C",
                "RUST_BACKTRACE": "0",
                "TZ": "UTC",
                "approval_variable": "LITCHI_PERF_0965_CAPTURE_APPROVED",
            },
            "argv_template": ["/usr/bin/taskset", "--cpu-list", "12-15", "<BINARY>", "--shape", "<SHAPE>", "--workload", "<WORKLOAD>", "--workers", "1", "--warmup", str(WARMUP), "--samples", str(SAMPLES), "--output", "<OUTPUT>"],
        },
        "capture": {
            "root": str((ROOT / "capture-ingress-diagnostic-v1").resolve()),
            "authorization": "root_GO_required",
            "measurements_started": False,
            "host_snapshot_before": host_before,
            "host_snapshot_after": {
                "status": "required_after_terminal",
                "path": str((ROOT / "protocol/host-after-v1.json").resolve()),
            },
            "host_snapshot_required": {"before": True, "after": True, "fields": ["cpu_load", "online_affinity", "memory", "filesystem", "scheduling_settings"]},
        },
        "analysis": {
            "root": str((ROOT / "analysis/ingress-diagnostic-v1").resolve()),
            "screen_reopened": False,
            "adoption_allowed": False,
            "causal_claim": False,
            "same_binary_controls": ["RR", "CC"],
            "cross_binary_orders": ["RC", "CR"],
            "cross_binary_pairs_per_case": {"RC": 1, "CR": 1},
            "packet_wide_pair_counts": {"RR": 3, "RC": 3, "CC": 3, "CR": 3},
        },
        "diagnostic_rule": {
            "type": "descriptive_classification_only",
            "threshold": 1.05,
            "frozen_0964_screen_rule_relaxed": False,
            "labels": ["concordant", "nonreproduced", "mixed"],
            "cross_order_spread_reported_only": True,
            "invalid_packet_no_diagnostic_label": True,
            "no_adoption_or_confirmation": True,
        },
        "root_go": root_go,
    }
    return plan


def write_once(path: Path, value: Any) -> None:
    path = path.expanduser().resolve()
    require(not path.exists() and not path.is_symlink(), f"output already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write((json.dumps(value, indent=2, sort_keys=True) + "\n").encode())
        stream.flush()
        os.fsync(stream.fileno())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base-plan", type=Path, default=BASE_PLAN)
    parser.add_argument("--runner", type=Path, default=RUNNER)
    parser.add_argument("--reducer", type=Path, default=REDUCER)
    parser.add_argument("--affinity-controls", type=Path, default=CONTROLS)
    parser.add_argument("--host-before", type=Path, default=HOST_BEFORE)
    parser.add_argument("--output", type=Path, default=OUTPUT)
    parser.add_argument("--status", choices=("prospective", "authorized"), default="prospective")
    parser.add_argument("--root-go", type=Path)
    args = parser.parse_args()
    try:
        write_once(args.output, make_plan(args))
    except PlanError as exc:
        parser.error(str(exc))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
