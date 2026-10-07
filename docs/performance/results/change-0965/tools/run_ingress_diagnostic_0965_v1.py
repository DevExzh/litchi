#!/usr/bin/env python3
"""Run the sealed 0965 retained-executable ingress diagnostic.

The default operations are read-only plan checks and the explicit affinity
controls.  Workload capture requires an authorized plan, a bound root GO, and
``LITCHI_PERF_0965_CAPTURE_APPROVED=1``.  Capture is serial, fresh-process,
terminal-first, fail-stop, and never retries or replaces a child.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0965").resolve()
PLAN_PATH = ROOT / "protocol/plan-diagnostic-0965-v6.json"
PLAN_SCHEMA = "litchi.performance.0965.ingress-diagnostic-plan.v1"
RUNNER_SCHEMA = "litchi.performance.0965.diagnostic-capture.v1"
CONTROL_SCHEMA = "litchi.performance.0965.affinity-launch-controls.v1"
DESIGN_SCHEMA = "prospective-design-v3"
EXPECTED_DESIGN_SHA = "0dbff34e930aec3d63350ebf6cce9d0c8a86afdcd8f176f50e52a7e6fe2c06bf"
APPROVAL_ENV = "LITCHI_PERF_0965_CAPTURE_APPROVED"
CPU_MASK = [12, 13, 14, 15]
DEADLINE = 900
WARMUP = 20
SAMPLES = 256
FIXED_ENV = {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "RUST_BACKTRACE": "0", "TZ": "UTC"}
EXPECTED_CASES = ("medium-one-percent", "dense-sparse-one-percent", "medium-one-edit")
PAIR_SCHEDULE = {
    "medium-one-percent": ("RR", "RC", "CC", "CR"),
    "dense-sparse-one-percent": ("RC", "CC", "CR", "RR"),
    "medium-one-edit": ("CC", "CR", "RR", "RC"),
}
PHASES = ("ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns")
TIME_KEYS = {
    "Maximum resident set size (kbytes)": "max_rss_kib",
    "User time (seconds)": "user_seconds",
    "System time (seconds)": "system_seconds",
    "Major (requiring I/O) page faults": "major_page_faults",
    "Minor (reclaiming a frame) page faults": "minor_page_faults",
    "Voluntary context switches": "voluntary_context_switches",
    "Involuntary context switches": "involuntary_context_switches",
}
AFFINITY_CONTROL_COMMAND = ["/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/true"]


class DiagnosticError(RuntimeError):
    pass


class DiagnosticInterrupted(DiagnosticError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise DiagnosticError(message)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="microseconds")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as exc:
        raise DiagnosticError(f"cannot hash {path}: {exc}") from exc
    return digest.hexdigest()


def descriptor(path: Path, executable: bool | None = None) -> dict[str, Any]:
    path = path.expanduser().resolve()
    require(path.is_file() and not path.is_symlink(), f"regular file required: {path}")
    mode = path.stat().st_mode & 0o777
    is_executable = bool(mode & 0o111)
    if executable is not None:
        require(is_executable is executable, f"executable mismatch: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "mode": mode, "executable": is_executable, "sha256": sha256(path)}


def bind(value: Any, label: str, executable: bool | None = None) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("path"), str), f"{label} binding missing")
    actual = descriptor(Path(value["path"]), executable)
    for key in ("path", "bytes", "mode", "executable", "sha256"):
        if key in value:
            require(value[key] == actual[key], f"{label}.{key} changed")
    return actual


def read_json(path: Path) -> Any:
    path = path.expanduser().resolve()
    require(path.is_file() and not path.is_symlink(), f"JSON missing: {path}")
    require(path.stat().st_size <= 128 * 1024 * 1024, f"JSON too large: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise DiagnosticError(f"invalid JSON {path}: {exc}") from exc


def write_once(path: Path, value: Any) -> None:
    path = path.expanduser().resolve()
    require(not path.exists() and not path.is_symlink(), f"artifact already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write((json.dumps(value, indent=2, sort_keys=True) + "\n").encode())
        stream.flush()
        os.fsync(stream.fileno())


def affinity_setter(mask: list[int]):
    frozen = set(mask)

    def set_affinity() -> None:
        os.sched_setaffinity(0, frozen)

    return set_affinity


def exact_process_affinity(pid: int) -> list[int]:
    try:
        return sorted(os.sched_getaffinity(pid))
    except (OSError, ProcessLookupError) as exc:
        raise DiagnosticError(f"cannot read launched process affinity for pid {pid}: {exc}") from exc


def group_clear(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return True
    except PermissionError:
        return False
    return False


def cleanup(process: subprocess.Popen[bytes] | None, pgid: int | None) -> dict[str, Any]:
    actions: list[str] = []
    if process is None:
        return {"actions": actions, "reaped": True, "process_group_clear": True}
    if pgid is None:
        pgid = process.pid
    if pgid is not None and not group_clear(pgid):
        try:
            os.killpg(pgid, signal.SIGTERM)
            actions.append("SIGTERM")
        except ProcessLookupError:
            pass
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(pgid, signal.SIGKILL)
                actions.append("SIGKILL")
            except ProcessLookupError:
                pass
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        return {"actions": actions, "reaped": False, "process_group_clear": False}
    end = time.monotonic() + 5
    while pgid is not None and time.monotonic() < end and not group_clear(pgid):
        time.sleep(0.05)
    return {"actions": actions, "reaped": True, "process_group_clear": pgid is None or group_clear(pgid)}


def run_process(argv: list[str], cwd: Path, child: Path, environment: dict[str, str], child_id: str | None = None, launched: list[str] | None = None) -> dict[str, Any]:
    process: subprocess.Popen[bytes] | None = None
    pgid: int | None = None
    error: str | None = None
    timed_out = False
    returncode: int | None = None
    observed: list[int] | None = None
    affinity_verified = False
    started = now()
    clean: dict[str, Any] = {"actions": [], "reaped": True, "process_group_clear": True}
    try:
        with (child / "stdout.log").open("xb") as stdout, (child / "stderr.log").open("xb") as stderr:
            started = now()
            process = subprocess.Popen(
                argv,
                cwd=str(cwd),
                env=environment,
                stdin=subprocess.DEVNULL,
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
                preexec_fn=affinity_setter(CPU_MASK),
            )
            if launched is not None and child_id is not None:
                launched.append(child_id)
            pgid = process.pid
            observed = exact_process_affinity(process.pid)
            require(observed == CPU_MASK, f"child affinity differs from frozen mask: {observed}")
            affinity_verified = True
            try:
                returncode = process.wait(timeout=DEADLINE)
            except subprocess.TimeoutExpired:
                timed_out = True
                clean = cleanup(process, pgid)
                returncode = process.returncode
            else:
                clean = cleanup(process, pgid)
    except BaseException as exc:
        error = f"{type(exc).__name__}: {exc}"
        if process is not None and launched is not None and child_id is not None and child_id not in launched:
            launched.append(child_id)
        clean = cleanup(process, pgid)
        if process is not None:
            returncode = process.returncode
    return {
        "started_utc": started,
        "terminal_utc": now(),
        "pid": process.pid if process else None,
        "process_group_id": pgid,
        "returncode": returncode,
        "timed_out": timed_out,
        "launch_error": error,
        "cleanup": clean,
        "reaped": clean["reaped"],
        "process_group_clear": clean["process_group_clear"],
        "observed_affinity": observed,
        "affinity_verified": affinity_verified and error is None and observed == CPU_MASK,
    }


def launch_affinity_control(mask: list[int]) -> tuple[subprocess.Popen[bytes], list[int], dict[str, Any]]:
    process = subprocess.Popen(
        AFFINITY_CONTROL_COMMAND,
        cwd="/tmp",
        env=FIXED_ENV,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
        preexec_fn=affinity_setter(mask),
    )
    observed = exact_process_affinity(process.pid)
    return process, observed, {"pid": process.pid, "command": list(AFFINITY_CONTROL_COMMAND), "requested_preexec_affinity": list(mask)}


def run_affinity_controls(output: Path | None) -> dict[str, Any]:
    wrong_mask = [CPU_MASK[0]]
    positive_process: subprocess.Popen[bytes] | None = None
    negative_process: subprocess.Popen[bytes] | None = None
    positive: dict[str, Any] = {}
    negative: dict[str, Any] = {}
    try:
        positive_process, observed, positive = launch_affinity_control(CPU_MASK)
        positive["observed_affinity"] = observed
        require(observed == CPU_MASK, f"positive control observed affinity {observed}")
        positive["guard_passed"] = True
        positive["returncode"] = positive_process.wait(timeout=5)
        require(positive["returncode"] == 0, f"positive control return code {positive['returncode']}")
        positive["cleanup"] = cleanup(positive_process, positive_process.pid)
        require(positive["cleanup"]["reaped"] and positive["cleanup"]["process_group_clear"], "positive control cleanup failed")
        positive_process = None

        negative_process, observed, negative = launch_affinity_control(wrong_mask)
        negative["observed_affinity"] = observed
        try:
            require(observed == CPU_MASK, f"negative control observed affinity {observed}")
        except DiagnosticError as exc:
            negative["guard_rejected"] = True
            negative["guard_error"] = str(exc)
        require(negative.get("guard_rejected") is True, "negative control was not rejected")
        negative["cleanup"] = cleanup(negative_process, negative_process.pid)
        require(negative["cleanup"]["reaped"] and negative["cleanup"]["process_group_clear"], "negative control cleanup failed")
        negative_process = None
    finally:
        if positive_process is not None:
            cleanup(positive_process, positive_process.pid)
        if negative_process is not None:
            cleanup(negative_process, negative_process.pid)
    result = {
        "schema": CONTROL_SCHEMA,
        "status": "pass",
        "created_utc": now(),
        "runner": descriptor(Path(__file__), True),
        "cpu_mask": CPU_MASK,
        "command": list(AFFINITY_CONTROL_COMMAND),
        "positive": positive,
        "negative": negative,
        "workload_started": False,
    }
    if output is not None:
        write_once(output, result)
    return result


def schedule_rows(plan: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for case_id in EXPECTED_CASES:
        for pair_index, pair_type in enumerate(PAIR_SCHEDULE[case_id]):
            pair_id = f"{case_id}__pair{pair_index:02d}__{pair_type}"
            for position, role in enumerate(pair_type):
                rows.append({
                    "child_id": f"native__{pair_id}__p{position}",
                    "lane": "native",
                    "case_id": case_id,
                    "pair_id": pair_id,
                    "pair_type": pair_type,
                    "pair_index": pair_index,
                    "position": position,
                    "role": role,
                    "binary_role": "baseline" if role == "R" else "candidate",
                    "comparison_id": "retained-baseline-vs-retained-candidate",
                    "shape": next(case["shape"] for case in plan["cases"] if case["case_id"] == case_id),
                    "workload": next(case["workload"] for case in plan["cases"] if case["case_id"] == case_id),
                    "workers": 1,
                    "warmup_iterations": WARMUP,
                    "measured_samples": SAMPLES,
                    "cpu_mask": list(CPU_MASK),
                })
    return rows


def validate_controls(plan: dict[str, Any], runner_binding: dict[str, Any]) -> None:
    controls = plan.get("affinity_controls")
    require(isinstance(controls, dict), "affinity controls binding missing")
    control_descriptor = bind(controls.get("descriptor"), "affinity controls")
    value = read_json(Path(control_descriptor["path"]))
    require(value.get("schema") == CONTROL_SCHEMA and value.get("status") == "pass", "affinity controls status/schema changed")
    require(value.get("workload_started") is False and value.get("cpu_mask") == CPU_MASK, "affinity controls workload/mask changed")
    require(value.get("runner") == runner_binding, "affinity controls runner binding changed")
    positive = value.get("positive", {})
    negative = value.get("negative", {})
    require(positive.get("requested_preexec_affinity") == CPU_MASK and positive.get("observed_affinity") == CPU_MASK, "positive affinity control changed")
    require(positive.get("guard_passed") is True and positive.get("returncode") == 0, "positive affinity guard failed")
    require(negative.get("requested_preexec_affinity") == [CPU_MASK[0]] and negative.get("observed_affinity") != CPU_MASK, "negative affinity control did not exercise mismatch")
    require(negative.get("guard_rejected") is True and isinstance(negative.get("guard_error"), str), "negative affinity guard failed")
    for item in (positive, negative):
        require(item.get("cleanup", {}).get("reaped") is True and item.get("cleanup", {}).get("process_group_clear") is True, "affinity control cleanup failed")


def validate_plan(plan_path: Path, capture: bool = False) -> tuple[dict[str, Any], dict[str, Any]]:
    plan_path = plan_path.expanduser().resolve()
    plan = read_json(plan_path)
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema mismatch")
    require(plan.get("status") in {"prospective", "authorized"}, "plan status invalid")
    require(plan.get("plan_id") == "xlsx-ingress-tail-diagnostic-0965", "plan id changed")
    require(plan.get("scope", {}).get("children") == 24 and plan.get("scope", {}).get("pairs") == 12, "child/pair scope changed")
    require(plan.get("scope", {}).get("warmup_iterations") == WARMUP and plan.get("scope", {}).get("measured_samples") == SAMPLES, "sample scope changed")
    require(plan.get("scope", {}).get("retry_or_replacement") is False, "retry policy changed")
    require(plan.get("scope", {}).get("diagnostic_only") is True, "diagnostic-only scope missing")
    require(plan.get("capture", {}).get("authorization") == "root_GO_required", "root GO policy missing")
    require(plan.get("capture", {}).get("measurements_started") is False, "measurements already started")
    require(plan.get("analysis", {}).get("adoption_allowed") is False and plan.get("analysis", {}).get("screen_reopened") is False, "analysis decision policy changed")
    require(plan.get("diagnostic_rule", {}).get("frozen_0964_screen_rule_relaxed") is False, "0964 rule was relaxed")
    design = plan.get("design", {})
    require(design.get("schema") == DESIGN_SCHEMA and design.get("status") == "sealed", "sealed design v3 binding missing")
    design_binding = bind(design.get("descriptor"), "design v3")
    require(design_binding["sha256"] == EXPECTED_DESIGN_SHA, "sealed design v3 hash changed")
    runner = bind(plan.get("runner", {}).get("descriptor"), "runner", True)
    require(plan.get("runner", {}).get("schema") == RUNNER_SCHEMA, "runner schema changed")
    require(plan.get("runner", {}).get("cpu_mask") == CPU_MASK and plan.get("runner", {}).get("deadline_seconds") == DEADLINE, "runner binding changed")
    reducer = bind(plan.get("reducer", {}).get("descriptor"), "reducer", True)
    require(plan.get("reducer", {}).get("schema") == "litchi.performance.0965.ingress-diagnostic-reduction.v1", "reducer schema changed")
    require(plan.get("reducer", {}).get("n") == SAMPLES and plan.get("reducer", {}).get("p99_rank") == 254, "reducer sample contract changed")
    validate_controls(plan, runner)

    inherited = plan.get("inherited_0964", {})
    inherited_plan = bind(inherited.get("plan"), "inherited 0964 plan")
    inherited_audit = bind(inherited.get("raw_audit"), "inherited 0964 raw audit")
    require(inherited.get("plan_revision") == "epoch2-root-final-v1", "inherited 0964 revision changed")
    base = read_json(Path(inherited_plan["path"]))
    require(base.get("schema") == "litchi.performance.0964.inline-number-screen-plan.v1" and base.get("status") == "authorized", "inherited 0964 plan is not final")
    require(base.get("plan_revision") == "epoch2-root-final-v1" and base.get("scope", {}).get("capture_epoch") == "epoch2-affinity-hardened", "inherited 0964 epoch changed")
    source = plan.get("source", {})
    for key in ("baseline_manifest", "candidate_manifest", "candidate_allowlist", "baseline_binding", "candidate_freeze", "candidate_quality", "candidate_build_result", "required_perf_baseline_lock", "candidate_perf_baseline_lock", "candidate_root_lock_overlay"):
        bind(source.get(key), f"source.{key}")
    require(source.get("required_perf_baseline_lock_sha256") == "db62bdd9174484613acae1f49f8187cf105d0b4976e2fea4de91df7372a893d5", "required lock changed")
    require(source.get("production_adoption") is False and source.get("source_rebuilt") is False, "source custody policy changed")
    for key in ("baseline_source_root", "candidate_source_root"):
        root = Path(source.get(key, ""))
        require(root.is_dir() and not root.is_symlink(), f"source root missing: {key}")
    binaries = plan.get("binaries", {})
    baseline = bind(binaries.get("R"), "baseline executable", True)
    candidate = bind(binaries.get("C"), "candidate executable", True)
    require(baseline["sha256"] == "5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3", "baseline executable hash changed")
    require(candidate["sha256"] == "d21863a6613800391a01914865fc4eaecdf54ad3061447be8b6be7d47d0dc910", "candidate executable hash changed")

    cases = plan.get("cases")
    require(isinstance(cases, list) and [case.get("case_id") for case in cases] == list(EXPECTED_CASES), "case order changed")
    require(cases == base.get("cases"), "case oracle binding differs from final 0964 plan")
    contract = plan.get("report_contract", {}).get("native", {})
    require(contract.get("schema") == "litchi.xlsx-managed-compression-benchmark.v1", "report schema changed")
    require(contract.get("warmup_iterations") == WARMUP and contract.get("measured_samples") == SAMPLES and contract.get("raw_sample_count") == SAMPLES, "report sample contract changed")
    require(contract.get("sample_phase_fields") == list(PHASES), "report phase contract changed")
    require(plan.get("execution", {}).get("cpu_mask") == CPU_MASK and plan.get("execution", {}).get("max_parallel_children") == 1, "execution serial/affinity contract changed")
    cwd = Path(plan.get("execution", {}).get("workload_cwd", ""))
    require(cwd.is_dir() and not cwd.is_symlink(), "workload cwd missing")
    require(plan.get("execution", {}).get("environment", {}).get("approval_variable") == APPROVAL_ENV, "approval variable changed")

    expected_rows = schedule_rows(plan)
    actual_rows = plan.get("schedule", {}).get("children")
    require(actual_rows == expected_rows, "exact 0965 child schedule changed")
    expected_pairs = []
    for case_id in EXPECTED_CASES:
        for pair_index, pair_type in enumerate(PAIR_SCHEDULE[case_id]):
            pair_id = f"{case_id}__pair{pair_index:02d}__{pair_type}"
            expected_pairs.append({"pair_id": pair_id, "case_id": case_id, "pair_index": pair_index, "pair_type": pair_type, "children": [f"native__{pair_id}__p0", f"native__{pair_id}__p1"], "roles": list(pair_type), "adjacent": True})
    require(plan.get("schedule", {}).get("pairs") == expected_pairs, "exact pair schedule changed")
    require(plan.get("analysis", {}).get("cross_binary_pairs_per_case") == {"RC": 1, "CR": 1}, "cross-order pair counts changed")
    require(plan.get("analysis", {}).get("packet_wide_pair_counts") == {"RR": 3, "RC": 3, "CC": 3, "CR": 3}, "packet-wide pair counts changed")
    require(plan.get("capture", {}).get("host_snapshot_before", {}).get("path"), "host-before binding missing")
    bind(plan["capture"]["host_snapshot_before"], "host-before")
    host_after = plan.get("capture", {}).get("host_snapshot_after", {})
    require(host_after.get("status") == "required_after_terminal", "host-after policy changed")
    require(host_after.get("path") == str((ROOT / "protocol/host-after-v1.json").resolve()), "host-after path changed")

    root_go = plan.get("root_go", {})
    if capture:
        require(plan.get("status") == "authorized", "capture requires authorized plan")
        require(root_go.get("status") == "GO" and isinstance(root_go.get("descriptor"), dict), "capture requires bound root GO")
        go_descriptor = bind(root_go["descriptor"], "root GO")
        go = read_json(Path(go_descriptor["path"]))
        require(go.get("status") == "GO", "root GO receipt status changed")
        require(os.environ.get(APPROVAL_ENV) == "1", f"set {APPROVAL_ENV}=1 only for an authorized capture")
    return plan, {"plan": descriptor(plan_path), "runner": runner, "reducer": reducer, "baseline": baseline, "candidate": candidate}


def case_for(plan: dict[str, Any], case_id: str) -> dict[str, Any]:
    return next(case for case in plan["cases"] if case["case_id"] == case_id)


def positive_int(value: Any, label: str) -> int:
    require(type(value) is int and value > 0, f"{label} must be a positive integer")
    return value


def parse_time(path: Path) -> dict[str, float | int]:
    require(path.is_file() and not path.is_symlink(), f"time output missing: {path}")
    values: dict[str, float | int] = {}
    for line in path.read_text(errors="replace").splitlines():
        if ":" not in line:
            continue
        key, raw = line.split(":", 1)
        name = TIME_KEYS.get(key.strip())
        if name is None:
            continue
        token = raw.strip().split()[0] if raw.strip() else ""
        try:
            value: float | int = float(token) if name in {"user_seconds", "system_seconds"} else int(token)
        except ValueError as exc:
            raise DiagnosticError(f"invalid /usr/bin/time metric {name}") from exc
        require(math.isfinite(float(value)) and value >= 0, f"invalid /usr/bin/time metric {name}")
        values[name] = value
    require(set(TIME_KEYS.values()) <= values.keys(), "complete /usr/bin/time metric set missing")
    require(type(values.get("max_rss_kib")) is int and values["max_rss_kib"] > 0, "positive max RSS missing")
    return values


def validate_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any]) -> None:
    contract = plan["report_contract"]["native"]
    require(report.get("schema") == contract["schema"] and report.get("status") == "completed", f"{row['child_id']} report status/schema changed")
    case_cfg = case_for(plan, row["case_id"])
    for key, expected in (("shape", row["shape"]), ("workload", row["workload"]), ("workers", 1), ("warmup_iterations", WARMUP), ("measured_samples", SAMPLES)):
        require(report.get(key) == expected, f"{row['child_id']} report root {key} changed")
    require(report.get("performance_claim") == contract["performance_claim"], f"{row['child_id']} performance claim changed")
    case = report.get("case")
    require(isinstance(case, dict), f"{row['child_id']} case missing")
    expected = {
        "id": f"{row['shape']}-{row['workload']}-workers1",
        "shape": row["shape"], "workload": row["workload"], "workers": 1,
        "update_count": case_cfg["update_count"], "source_bytes": case_cfg["source_bytes"], "source_sha256": case_cfg["source_sha256"],
        "output_bytes": case_cfg["output_bytes"], "output_sha256": case_cfg["output_sha256"], "semantic_sha256": case_cfg["semantic_sha256"],
        "expected_cpu_tasks": case_cfg["expected_cpu_tasks"], "admission_expected_worker_permits": 0, "admission_observed_worker_permits": 0,
        "admission_cpu_tasks_used": case_cfg["expected_cpu_tasks"], "preflight_actual_commit_inverse_verified": True,
        "context_limits": case_cfg["context_limits"], "sink_capacity_bytes": case_cfg["context_limits"]["output_bytes"],
        "sink_max_write_bytes": 65536, "warmup_iterations": WARMUP, "measured_samples": SAMPLES,
    }
    for key, expected_value in expected.items():
        require(case.get(key) == expected_value, f"{row['child_id']} case oracle {key} changed")
    samples = case.get("samples")
    require(isinstance(samples, list) and len(samples) == SAMPLES, f"{row['child_id']} sample count changed")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("sample") == index, f"{row['child_id']} sample index changed")
        values = [positive_int(sample.get(name), f"{row['child_id']} {name}") for name in PHASES]
        require(sum(values[:4]) <= values[4], f"{row['child_id']} phase sum exceeds total")
        for key, wanted in (("output_bytes", case_cfg["output_bytes"]), ("output_bytes_used", case_cfg["output_bytes"]), ("cpu_tasks_used", case_cfg["expected_cpu_tasks"]), ("output_sha256", case_cfg["output_sha256"]), ("semantic_sha256", case_cfg["semantic_sha256"])):
            require(sample.get(key) == wanted, f"{row['child_id']} sample oracle {key} changed")
        require(sample.get("fixture_inverse_restored_exact_source") is True, f"{row['child_id']} inverse oracle failed")
        require(type(sample.get("sink_write_calls")) is int and sample["sink_write_calls"] > 0, f"{row['child_id']} sink calls missing")
        require(type(sample.get("sink_largest_write")) is int and 0 < sample["sink_largest_write"] <= 65536, f"{row['child_id']} sink bound changed")
        resources = sample.get("resources_after_drop")
        require(isinstance(resources, dict) and all(resources.get(name) == 0 for name in ("memory", "workers", "io_concurrency", "objects")), f"{row['child_id']} resource cleanup failed")


def binary_for(plan: dict[str, Any], row: dict[str, Any]) -> Path:
    key = "R" if row["role"] == "R" else "C"
    return Path(plan["binaries"][key]["path"]).resolve()


def child_environment(plan: dict[str, Any]) -> dict[str, str]:
    configured = plan["execution"]["environment"]
    environment = {key: str(value) for key, value in configured.items() if key != "approval_variable"}
    environment[configured["approval_variable"]] = "1"
    return environment


def command_for(plan: dict[str, Any], row: dict[str, Any], child: Path) -> tuple[list[str], list[str], Path, Path, dict[str, str]]:
    binary = binary_for(plan, row)
    cwd = Path(plan["execution"]["workload_cwd"])
    output = child / "output"
    time_file = child / "time-v.txt"
    require(not output.exists() and not output.is_symlink(), "benchmark output must be absent before launch")
    direct = [
        "/usr/bin/taskset", "--cpu-list", "12-15", str(binary),
        "--shape", row["shape"], "--workload", row["workload"],
        "--workers", "1", "--warmup", str(WARMUP), "--samples", str(SAMPLES), "--output", str(output),
    ]
    wrapped = [
        "/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/time", "-v", "-o", str(time_file), "--", *direct[3:],
    ]
    return wrapped, direct, cwd, output, child_environment(plan)


def run_child(plan: dict[str, Any], plan_path: Path, row: dict[str, Any], packet: Path, index: int, launched: list[str]) -> None:
    child = packet / "children" / row["child_id"]
    child.mkdir(parents=True, exist_ok=False)
    argv, workload_argv, cwd, output, environment = command_for(plan, row, child)
    binary = binary_for(plan, row)
    binary_key = "R" if row["role"] == "R" else "C"
    binary_binding = bind(plan["binaries"][binary_key], f"{row['child_id']} binary {binary_key}", True)
    require(binary_binding["path"] == str(binary), f"{row['child_id']} binary path changed")
    start = {
        "schema": RUNNER_SCHEMA + ".child-start",
        "index": index,
        "child": row,
        "plan": descriptor(plan_path),
        "binary": binary_binding,
        "argv": argv,
        "workload_argv": workload_argv,
        "cwd": str(cwd),
        "environment": environment,
        "deadline_seconds": DEADLINE,
        "output_absent_before_launch": True,
        "started_utc": now(),
    }
    write_once(child / "start.json", start)
    run = run_process(argv, cwd, child, environment, row["child_id"], launched)
    ok = run["launch_error"] is None and run["returncode"] == 0 and not run["timed_out"] and run["reaped"] and run["process_group_clear"] and run["affinity_verified"]
    write_once(child / "launch.json", {
        "schema": RUNNER_SCHEMA + ".child-launch",
        "pid": run["pid"], "observed_affinity": run["observed_affinity"], "expected_affinity": CPU_MASK,
        "affinity_verified_before_wait": run["affinity_verified"], "launched_utc": run["started_utc"],
    })
    terminal = {
        "schema": RUNNER_SCHEMA + ".child-terminal",
        "index": index, "child_id": row["child_id"], "status": "completed" if ok else "failed",
        "return_code": run["returncode"], "timed_out": run["timed_out"], "process_group_clear": run["process_group_clear"],
        "pid": run["pid"], "run": run, "output": {"path": str(output), "exists_after_terminal": output.is_dir() and not output.is_symlink()},
        "terminal_before_report_parse": True, "ended_utc": now(),
    }
    write_once(child / "terminal.json", terminal)
    require(ok, f"{row['child_id']} process failed")
    try:
        report = output / "report.json"
        require(report.is_file() and not report.is_symlink(), f"{row['child_id']} report missing")
        report_value = read_json(report)
        validate_report(report_value, row, plan)
        validation = {
            "schema": RUNNER_SCHEMA + ".child-validation", "status": "passed", "index": index, "child": row,
            "plan": descriptor(plan_path), "terminal": descriptor(child / "terminal.json"), "report": descriptor(report),
            "time": descriptor(child / "time-v.txt"), "time_metrics": parse_time(child / "time-v.txt"),
        }
    except BaseException as exc:
        validation = {
            "schema": RUNNER_SCHEMA + ".child-validation", "status": "failed", "index": index, "child": row,
            "terminal": descriptor(child / "terminal.json"), "error": f"{type(exc).__name__}: {exc}",
        }
        write_once(child / "validation.json", validation)
        raise
    write_once(child / "validation.json", validation)
    write_once(child / "post-terminal-artifacts.json", {
        "schema": RUNNER_SCHEMA + ".post-terminal-artifacts",
        "terminal": descriptor(child / "terminal.json"), "stdout": descriptor(child / "stdout.log"),
        "stderr": descriptor(child / "stderr.log"), "time": descriptor(child / "time-v.txt"),
        "report": descriptor(report), "validation": descriptor(child / "validation.json"),
    })


def _raise_for_signal(signum: int, _frame: Any) -> None:
    raise DiagnosticInterrupted(f"received signal {signum}")


def _install_signal_handlers() -> dict[int, Any]:
    previous: dict[int, Any] = {}
    for signum in (signal.SIGINT, signal.SIGTERM):
        previous[signum] = signal.getsignal(signum)
        signal.signal(signum, _raise_for_signal)
    return previous


def _restore_signal_handlers(previous: dict[int, Any]) -> None:
    for signum, handler in previous.items():
        signal.signal(signum, handler)


def capture(plan_path: Path, packet: Path) -> int:
    plan, bindings = validate_plan(plan_path, capture=True)
    packet = packet.expanduser().resolve()
    require(not packet.exists() and not packet.is_symlink(), f"capture root already exists: {packet}")
    packet.mkdir(parents=True, exist_ok=False)
    (packet / "children").mkdir()
    rows = plan["schedule"]["children"]
    launched: list[str] = []
    completed: list[str] = []
    error: str | None = None
    previous = _install_signal_handlers()
    try:
        write_once(packet / "capture-start.json", {
            "schema": RUNNER_SCHEMA + ".capture-start", "status": "started", "lane": "native",
            "plan": bindings["plan"], "runner": bindings["runner"], "serial": True,
            "children_expected": len(rows), "warmup_iterations": WARMUP, "measured_samples": SAMPLES,
            "retry_or_replacement": False, "host_snapshot_before": plan["capture"]["host_snapshot_before"],
            "started_utc": now(),
        })
        for index, row in enumerate(rows):
            run_child(plan, plan_path, row, packet, index, launched)
            completed.append(row["child_id"])
    except BaseException as exc:
        error = f"{type(exc).__name__}: {exc}"
    finally:
        terminal_status = "completed" if error is None and len(completed) == len(rows) else "invalid"
        terminal = {
            "schema": RUNNER_SCHEMA + ".capture-terminal", "status": terminal_status, "lane": "native",
            "children_started": len(launched), "children_expected": len(rows), "children_completed": len(completed),
            "retry_or_replacement": False, "terminal_before_parse": True,
            "capture_start": descriptor(packet / "capture-start.json") if (packet / "capture-start.json").exists() else None,
            "host_snapshot_after": plan["capture"]["host_snapshot_after"], "ended_utc": now(),
        }
        if error is not None:
            terminal["error"] = error
        write_once(packet / "capture-terminal.json", terminal)
        _restore_signal_handlers(previous)
    if error is not None or len(completed) != len(rows):
        return 1
    write_once(packet / "completed-packet.json", {
        "schema": RUNNER_SCHEMA + ".completed-packet", "status": "completed", "lane": "native",
        "plan": bindings["plan"], "runner": bindings["runner"],
        "children": [{"child_id": row["child_id"], "validation": descriptor(packet / "children" / row["child_id"] / "validation.json")} for row in rows],
        "capture_terminal": descriptor(packet / "capture-terminal.json"),
    })
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, default=PLAN_PATH)
    parser.add_argument("--check-plan", action="store_true")
    parser.add_argument("--affinity-controls", action="store_true")
    parser.add_argument("--affinity-controls-output", type=Path)
    parser.add_argument("--capture", action="store_true")
    parser.add_argument("--packet", type=Path)
    args = parser.parse_args(argv)
    try:
        if args.affinity_controls:
            run_affinity_controls(args.affinity_controls_output)
            return 0
        if args.check_plan:
            validate_plan(args.plan.resolve(), capture=False)
            return 0
        if args.capture:
            plan = read_json(args.plan.resolve())
            packet = args.packet or Path(plan["capture"]["root"])
            return capture(args.plan.resolve(), packet)
        parser.error("choose --check-plan, --affinity-controls, or --capture")
    except DiagnosticError as exc:
        parser.error(str(exc))
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
