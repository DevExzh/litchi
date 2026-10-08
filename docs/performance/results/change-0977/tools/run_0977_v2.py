#!/usr/bin/env python3
"""Serial 0977 Office qualification capture runner.

This runner owns process launch and packet custody.  It does not patch any
retained runner module.  A real run requires an explicit root GO receipt; the
default and self-test paths never start perf or an Office process.
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import signal
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Mapping

import plan_binding_0977_v2 as custody

ROOT = custody.ROOT
TOOLS = ROOT / "tools"
BRIDGE = TOOLS / "affinity_bridge_0977.py"
LAUNCHER = custody.LAUNCHER
PACKET_SCHEMA = "litchi.performance.0977.packet.v1"
TERMINAL_SCHEMA = "litchi.performance.0977.terminal.v1"
MANIFEST_SCHEMA = "litchi.performance.0977.capture_manifest.v1"
CHILD_PROGRESS_SCHEMA = "litchi.performance.0977.child_progress.v1"


class RunnerError(RuntimeError):
    pass


def utc() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def strict_load(path: Path) -> Any:
    return custody.strict_load(path)


def write_exclusive(path: Path, value: Any) -> None:
    custody.write_exclusive(path, value)


def native_env() -> dict[str, str]:
    # Keep only normal execution environment values; no decoder preload or
    # LITCHI diagnostic variables may reach the Office target.
    env = dict(os.environ)
    env.pop("LD_PRELOAD", None)
    for key in list(env):
        if key.startswith("LITCHI_"):
            env.pop(key, None)
    env.update({"LANG": "C", "LC_ALL": "C", "TZ": "UTC", "PYTHONDONTWRITEBYTECODE": "1"})
    return env


def recorded_environment(env: Mapping[str, str]) -> dict[str, Any]:
    allowed = {"LANG", "LC_ALL", "TZ", "PYTHONDONTWRITEBYTECODE"}
    return {key: env.get(key) for key in sorted(allowed)} | {
        "LD_PRELOAD": "absent",
        "secret_environment_recorded": False,
    }


def terminate_group(proc: subprocess.Popen[Any]) -> None:
    try:
        os.killpg(proc.pid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    try:
        proc.wait(timeout=1.0)
        return
    except subprocess.TimeoutExpired:
        pass
    try:
        os.killpg(proc.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass
    proc.wait(timeout=2.0)


def verify_group_gone(pid: int | None) -> dict[str, Any]:
    """Ensure the runner-owned process group is gone, with bounded recheck."""
    if pid is None:
        return {"status": "unavailable", "checked": False}
    try:
        os.killpg(pid, 0)
    except ProcessLookupError:
        return {"status": "gone", "checked": True, "passed": True}
    except PermissionError as exc:
        return {"status": "error", "checked": True, "passed": False, "error": str(exc)}
    # The runner's session is private.  A surviving group after wait is an
    # orphaned target; terminate it, then verify that the group really went.
    try:
        os.killpg(pid, signal.SIGKILL)
    except ProcessLookupError:
        return {"status": "gone_after_cleanup", "checked": True, "passed": True}
    for _ in range(20):
        try:
            os.killpg(pid, 0)
        except ProcessLookupError:
            return {"status": "gone_after_cleanup", "checked": True, "passed": True}
        time.sleep(0.025)
    return {"status": "survived_cleanup", "checked": True, "passed": False, "pid": pid}


def write_text_exclusive(path: Path, text: str) -> None:
    if path.exists():
        raise RunnerError(f"refusing to overwrite {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="")


def target_argv(plan: Mapping[str, Any], child_dir: Path) -> list[str]:
    binary = Path(plan["binary"]["path"]).resolve()
    fixture = Path(plan["fixture"]["path"]).resolve()
    workload = plan["workload"]
    launch = child_dir / f"{child_dir.name}.launch.json"
    return [
        str(Path(plan["python"]["path"]).resolve()),
        str(Path(plan["launcher"]["path"]).resolve()),
        "--metadata", str(launch),
        "--requested-affinity", "[31]",
        "--",
        str(binary),
        "--case", str(workload["case"]),
        "--input", str(fixture),
        "--edits", str(workload["edits"]),
        "--retention", str(workload["retention"]),
        "--warmup", str(custody.WARMUPS),
        "--samples", str(custody.SAMPLES),
    ]


def record_argv(plan: Mapping[str, Any], child_dir: Path, *, control: bool = False) -> list[str]:
    data = child_dir / f"{child_dir.name}.perf.data"
    bridge_receipt = child_dir / f"{child_dir.name}.bridge.json"
    target = target_argv(plan, child_dir) if not control else []
    target_script = plan["control"]["target_script"] if control else plan["launcher"]
    bridge_args = [
        str(Path(plan["python"]["path"]).resolve()), str(BRIDGE.resolve()),
        "--expected-perf-exe", str(Path(plan["perf"]["path"]).resolve()),
        "--expected-perf-sha256", plan["perf"]["sha256"],
        "--expected-target-exe", str(Path(plan["python"]["path"]).resolve()),
        "--expected-target-sha256", plan["python"]["sha256"],
        "--expected-target-script", str(Path(target_script["path"]).resolve()),
        "--expected-target-script-sha256", target_script["sha256"],
        "--bridge-receipt", str(bridge_receipt),
        "--target-cpu", str(custody.CPU), "--recorder-cpu", str(custody.RECORDER_CPU),
    ]
    if control:
        bridge_args += ["--control-target", "--control-receipt", str(child_dir / f"{child_dir.name}.control-target.json"),
                        "--control-seconds", str(plan["control"]["seconds"]),
                        "--control-check-interval", str(plan["control"]["check_interval"])]
    else:
        bridge_args += ["--", *target]
    return [
        str(Path(plan["taskset"]["path"]).resolve()), "--cpu-list", str(custody.CPU),
        str(Path(plan["perf"]["path"]).resolve()), "record", "--per-thread",
        "-e", "instructions:u", "-c", str(custody.PERIOD), "-m", "16M",
        "--call-graph", "dwarf,32768", "--period", "--timestamp", "--sample-cpu",
        "--buildid-all", "--no-buildid-cache", "--user-callchains", "-o", str(data),
        "--", *bridge_args,
    ]


def task_observation(pid: int | None, phase: str) -> dict[str, Any]:
    """Record a bounded /proc observation; raw perf remains authoritative."""
    row: dict[str, Any] = {"phase": phase, "pid": pid, "scope": "targeted perf/target task only",
                           "lifetime_authority": "raw zero-loss perf task stream", "live_history_claim": False}
    if pid is None:
        row.update({"status": "unavailable", "reason": "target PID not known at observation"})
        return row
    try:
        tids = sorted(int(entry.name) for entry in Path(f"/proc/{pid}/task").iterdir() if entry.name.isdecimal())
        rows = []
        for tid in tids:
            rows.append({"tid": tid, "starttime": proc_starttime(tid), "affinity": sorted(os.sched_getaffinity(tid))})
        row.update({"status": "present", "task_ids": tids, "tasks": rows,
                    "identity": {"pid": pid, "uid": os.stat(f"/proc/{pid}").st_uid,
                                 "starttime": proc_starttime(pid),
                                 "comm": Path(f"/proc/{pid}/comm").read_text(encoding="utf-8").strip(),
                                 "exe": os.readlink(f"/proc/{pid}/exe")}})
    except FileNotFoundError:
        row.update({"status": "absent", "task_ids": [], "reason": "not present at bounded observation"})
    except PermissionError as exc:
        row.update({"status": "error", "task_ids": None, "reason": "permission_denied", "error": str(exc)})
    except OSError as exc:
        row.update({"status": "error", "task_ids": None, "reason": "proc_observation_error", "error": str(exc)})
    return row


def proc_starttime(pid: int) -> int:
    """Return the Linux start-time field for a PID/TID without broad scans."""
    text = Path(f"/proc/{pid}/stat").read_text(encoding="utf-8")
    close = text.rfind(")")
    fields = text[close + 2 :].split() if close >= 0 else []
    if len(fields) <= 19:
        raise RunnerError(f"malformed /proc/{pid}/stat")
    return int(fields[19])


def target_pid_from_launch(path: Path, bridge_path: Path | None = None) -> int | None:
    try:
        value = strict_load(path)
        pid = value.get("exec_pid")
        return pid if type(pid) is int and pid > 0 else None
    except (OSError, RunnerError, ValueError, TypeError, json.JSONDecodeError):
        pass
    if bridge_path is not None:
        try:
            value = strict_load(bridge_path)
            pid = value.get("bridge_pid")
            return pid if type(pid) is int and pid > 0 else None
        except (OSError, RunnerError, ValueError, TypeError, json.JSONDecodeError):
            return None
    return None


def observe_processes(proc_pid: int, launch_path: Path, observations: dict[str, Any], phase: str,
                      bridge_path: Path) -> None:
    bridge_ready = False
    if bridge_path.is_file():
        try:
            bridge_ready = strict_load(bridge_path).get("status") == "ready_to_exec"
        except Exception:
            bridge_ready = False
    observations["rows"].append({
        "monotonic_ns": time.monotonic_ns(),
        "phase": phase,
        "bridge_handoff_ready": bridge_ready,
        "perf_parent": task_observation(proc_pid, phase),
        "target": task_observation(target_pid_from_launch(launch_path, bridge_path), phase),
    })


def _task_keys(rows: Any) -> set[tuple[int, int]]:
    if not isinstance(rows, list):
        return set()
    result: set[tuple[int, int]] = set()
    for row in rows:
        if isinstance(row, Mapping) and type(row.get("tid")) is int and type(row.get("starttime")) is int:
            result.add((row["tid"], row["starttime"]))
    return result


def affinity_observation_gate(observations: Mapping[str, Any], bridge: Mapping[str, Any],
                              launch: Mapping[str, Any] | None, *, control: bool) -> dict[str, Any]:
    """Gate only observations made after the bridge says handoff completed.

    The raw perf task stream remains the lifetime authority.  These rows are a
    bounded sampled check: pre-handoff CPU31 observations are allowed, and an
    absent post-exit /proc task is not treated as a live-history claim.
    """
    if bridge.get("status") != "ready_to_exec" or bridge.get("native_exec_same_pid") is not True:
        raise RunnerError("bridge receipt is not a completed handoff")
    expected_parent = _task_keys(bridge.get("tasks_after"))
    if not expected_parent:
        raise RunnerError("bridge receipt has no parent task identity set")
    target_pid = bridge.get("bridge_pid")
    target_start = bridge.get("bridge_starttime")
    if not control and isinstance(launch, Mapping):
        target_pid = launch.get("exec_pid")
        if target_pid != bridge.get("bridge_pid"):
            raise RunnerError("launch PID is not the bridge same-PID exec")
    checked = 0
    active_target = 0
    postexit_target = 0
    errors: list[str] = []
    for row in observations.get("rows", []):
        if not isinstance(row, Mapping) or row.get("bridge_handoff_ready") is not True:
            continue
        checked += 1
        parent = row.get("perf_parent")
        if not isinstance(parent, Mapping):
            errors.append("handoff_row_missing_parent")
            continue
        if parent.get("status") == "present":
            identity = parent.get("identity")
            expected_identity = bridge.get("parent_after") or bridge.get("parent_before") or {}
            if not isinstance(identity, Mapping) or any(identity.get(key) != expected_identity.get(key) for key in ("pid", "uid", "starttime", "comm", "exe")):
                errors.append("parent_identity_changed")
            tasks = parent.get("tasks")
            keys = _task_keys(tasks)
            if keys != expected_parent:
                errors.append("parent_task_identity_changed")
            if not isinstance(tasks, list) or any(task.get("affinity") != [custody.RECORDER_CPU] for task in tasks if isinstance(task, Mapping)):
                errors.append("parent_task_not_on_recorder_cpu")
        elif parent.get("status") not in {"absent", "unavailable"}:
            errors.append("parent_observation_error")
        target = row.get("target")
        if not isinstance(target, Mapping):
            errors.append("handoff_row_missing_target")
            continue
        if target.get("status") == "present":
            active_target += 1
            identity = target.get("identity")
            if not isinstance(identity, Mapping) or target_pid is not None and (identity.get("pid") != target_pid or identity.get("starttime") != target_start):
                errors.append("target_identity_changed")
            tasks = target.get("tasks")
            if not isinstance(tasks, list) or len(tasks) != 1:
                errors.append("target_task_count_not_one")
                continue
            task = tasks[0]
            if target_pid is not None and (task.get("tid") != target_pid or task.get("starttime") != target_start):
                errors.append("target_task_identity_changed")
            if task.get("affinity") != [custody.CPU]:
                errors.append("target_task_not_on_target_cpu")
            if parent.get("status") != "present":
                errors.append("target_present_without_parent_observation")
        elif target.get("status") == "absent":
            postexit_target += 1
        elif target.get("status") not in {"unavailable"}:
            errors.append("target_observation_error")
    if checked == 0:
        raise RunnerError("no post-handoff affinity observation was sampled")
    if errors:
        raise RunnerError("affinity observation gate failed: " + ", ".join(sorted(set(errors))))
    return {"passed": True, "post_handoff_rows": checked, "active_target_rows": active_target,
            "postexit_target_rows": postexit_target, "parent_task_keys": sorted(expected_parent),
            "scope": "bounded sampled observations; raw task stream remains lifetime authority"}


def load_decoder(path: Path):
    spec = importlib.util.spec_from_file_location("litchi_decode_0977_runtime", path)
    if spec is None or spec.loader is None:
        raise RunnerError(f"cannot load decoder {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if not callable(getattr(module, "decode_child", None)):
        raise RunnerError("decoder does not expose decode_child(plan, child_dir, terminal, launch)")
    return module


def packet_gate(packet: Mapping[str, Any]) -> tuple[bool, list[str]]:
    corrected = ((packet.get("perf") or {}).get("corrected")
                 if isinstance(packet.get("perf"), Mapping) else None)
    summary = corrected if isinstance(corrected, Mapping) else (
        packet.get("summary") if isinstance(packet.get("summary"), Mapping) else packet)
    decoder_gate = packet.get("gate") if isinstance(packet.get("gate"), Mapping) else {}
    reasons: list[str] = []
    if packet.get("status") != "complete":
        reasons.append("decoder_status_not_complete")
    checks = {
        "marker_sample_count": lambda value: isinstance(value, int) and value >= 100,
        "marker_period_share": lambda value: isinstance(value, (int, float)) and not isinstance(value, bool) and value > 0,
        "empty_callchain_count": lambda value: type(value) is int and value == 0,
        "sample_count": lambda value: type(value) is int and value > 0,
    }
    for key, predicate in checks.items():
        if not predicate(summary.get(key)):
            reasons.append(f"{key}_gate_failed")
    for key in ("raw_lost_sum", "terminal_lost_samples"):
        if type(decoder_gate.get(key)) is not int or decoder_gate.get(key) != 0:
            reasons.append(f"{key}_gate_failed")
    if decoder_gate.get("passed") is not True:
        reasons.append("decoder_gate_not_passed")
    return not reasons, reasons


def control_packet_gate(packet: Mapping[str, Any]) -> tuple[bool, list[str]]:
    summary = packet.get("summary") if isinstance(packet.get("summary"), Mapping) else packet
    reasons: list[str] = []
    if packet.get("status") != "complete":
        reasons.append("control_status_not_complete")
    for key in ("raw_lost_sum", "terminal_lost_samples"):
        if type(summary.get(key)) is not int or summary.get(key) != 0:
            reasons.append(f"{key}_gate_failed")
    if summary.get("sample_count", 0) <= 0:
        reasons.append("control_sample_count_missing")
    if summary.get("cpu_gate") is not True:
        reasons.append("control_cpu_gate_failed")
    if summary.get("task_gate") is not True:
        reasons.append("control_task_gate_failed")
    return not reasons, reasons


def packet_summary(packet: Mapping[str, Any]) -> Mapping[str, Any]:
    perf = packet.get("perf")
    if isinstance(perf, Mapping) and isinstance(perf.get("corrected"), Mapping):
        return perf["corrected"]
    if isinstance(packet.get("summary"), Mapping):
        return packet["summary"]
    return packet


def write_status(child_dir: Path, value: dict[str, Any]) -> None:
    write_exclusive(child_dir / f"{child_dir.name}.status.json", value)


def run_child(plan: Mapping[str, Any], child_number: int, capture_root: Path, decoder_path: Path, *, control: bool = False) -> dict[str, Any]:
    child_id = f"child-{child_number:02d}"
    expected_cwd = str(Path(plan["command_receipt_contract"]["cwd"]).resolve())
    if str(Path.cwd().resolve()) != expected_cwd:
        raise RunnerError(f"runner cwd drifted: {Path.cwd().resolve()} != {expected_cwd}")
    child_dir = capture_root / child_id
    try:
        child_dir.mkdir()
    except FileExistsError as exc:
        raise RunnerError(f"refusing to reuse child directory {child_dir}") from exc
    command = record_argv(plan, child_dir, control=control)
    command_path = child_dir / f"{child_id}.command.json"
    launch_path = child_dir / f"{child_id}.launch.json"
    bridge_path = child_dir / f"{child_id}.bridge.json"
    observations: dict[str, Any] = {
        "schema": "litchi.performance.0977.affinity_observations.v1",
        "interval_seconds": 0.02,
        "scope": "bounded observations of the perf parent and same-PID target",
        "lifetime_authority": "raw zero-loss perf task stream",
        "live_history_claim": False,
        "rows": [],
    }
    write_exclusive(command_path, {
        "schema": "litchi.performance.0977.command.v1",
        "packet_id": child_id,
        "argv": command,
        "cwd": str(Path.cwd()),
        "environment": recorded_environment(native_env()),
        "native_ld_preload": "absent",
        "same_pid_exec_chain": ["perf", "affinity_bridge_0977.py", "stdlib_cpu_burn_0976_v2.py"] if control else ["perf", "affinity_bridge_0977.py", "launch_baseline_v2.py", "bound_workload"],
    })
    report_path = child_dir / f"{child_id}.report.json"
    stderr_path = child_dir / f"{child_id}.perf.stderr"
    data_path = child_dir / f"{child_id}.perf.data"
    started = utc()
    process_pid: int | None = None
    timed_out = False
    launch_error = False
    returncode = 127
    with report_path.open("xb") as report_handle, stderr_path.open("xb") as stderr_handle:
        try:
            proc = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=report_handle, stderr=stderr_handle,
                                    env=native_env(), start_new_session=True)
            process_pid = proc.pid
            observe_processes(process_pid, launch_path, observations, "start", bridge_path)
            deadline = time.monotonic() + int(plan["execution"]["timeout_seconds"])
            while proc.poll() is None:
                observe_processes(process_pid, launch_path, observations, "periodic", bridge_path)
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    terminate_group(proc)
                    returncode = proc.returncode if proc.returncode is not None else 124
                    stderr_handle.write(b"timeout: process group terminated and reaped\n")
                    timed_out = True
                    break
                time.sleep(min(0.02, remaining))
            if not timed_out:
                returncode = proc.returncode if proc.returncode is not None else 127
            observe_processes(process_pid, launch_path, observations, "end", bridge_path)
        except OSError as exc:
            stderr_handle.write((f"launch error: {exc}\n").encode())
            launch_error = True
    group_cleanup = verify_group_gone(process_pid)
    terminal = {
        "schema": TERMINAL_SCHEMA, "packet_id": child_id, "started_utc": started, "finished_utc": utc(),
        "returncode": returncode, "timed_out": timed_out, "launch_error": launch_error,
        "process_pid": process_pid, "command": command, "native_ld_preload": "absent",
        "process_group_cleanup": group_cleanup,
    }
    terminal_path = child_dir / f"{child_id}.terminal.json"
    observation_path = child_dir / f"{child_id}.affinity-observations.json"
    write_exclusive(observation_path, observations)
    if returncode != 0 or timed_out or not data_path.is_file() or not report_path.is_file() or group_cleanup.get("passed") is not True:
        write_exclusive(terminal_path, terminal)
        status = {"schema": CHILD_PROGRESS_SCHEMA, "packet_id": child_id, "status": "terminal_failure", "terminal": terminal, "no_retry": True}
        write_status(child_dir, status)
        raise RunnerError(f"{child_id}: terminal/cleanup failure rc={returncode} cleanup={group_cleanup}")

    # The decoder owns all offline perf subcommands, including the terminal
    # LOST_SAMPLES report.  The runner records only the native target stdout
    # oracle, perf stderr/data, terminal status, and process custody.
    if not bridge_path.is_file() or (not control and not launch_path.is_file()):
        terminal["affinity_gate"] = {"passed": False, "reason": "bridge_or_launch_receipt_missing"}
        write_exclusive(terminal_path, terminal)
        status = {"schema": CHILD_PROGRESS_SCHEMA, "packet_id": child_id, "status": "custody_failure", "reason": "bridge_or_launch_receipt_missing", "no_retry": True}
        write_status(child_dir, status)
        raise RunnerError(f"{child_id}: same-PID bridge/launch receipt missing")
    bridge = strict_load(bridge_path)
    if control:
        launch = {}
    else:
        launch = strict_load(launch_path)
    try:
        terminal["affinity_gate"] = affinity_observation_gate(observations, bridge, launch, control=control)
    except Exception as exc:
        terminal["affinity_gate"] = {"passed": False, "error": str(exc)}
        write_exclusive(terminal_path, terminal)
        status = {"schema": CHILD_PROGRESS_SCHEMA, "packet_id": child_id, "status": "affinity_failure", "terminal": terminal, "no_retry": True}
        write_status(child_dir, status)
        raise
    write_exclusive(terminal_path, terminal)
    if bridge.get("status") != "ready_to_exec" or bridge.get("native_exec_same_pid") is not True:
        raise RunnerError(f"{child_id}: bridge handoff did not complete")
    if not control and (launch.get("exec_pid") != launch.get("pid") or launch.get("actual_affinity") != [custody.CPU] or launch.get("execve") is not True):
        raise RunnerError(f"{child_id}: launch receipt identity/affinity gate failed")
    decoder = load_decoder(decoder_path)
    if control:
        control_path = child_dir / f"{child_id}.control-target.json"
        if not control_path.is_file():
            raise RunnerError(f"{child_id}: control target receipt missing")
        control_receipt = strict_load(control_path)
        if control_receipt.get("status") != "pass" or control_receipt.get("control_only") is not True or control_receipt.get("native_workload_started") is not False:
            raise RunnerError(f"{child_id}: harmless control receipt failed")
        if control_receipt.get("target_cpu") != custody.CPU or control_receipt.get("recorder_cpu") != custody.RECORDER_CPU:
            raise RunnerError(f"{child_id}: harmless control CPU contract failed")
        validate_control = getattr(decoder, "validate_control_child", None)
        if not callable(validate_control):
            raise RunnerError("decoder does not expose validate_control_child for the harmless control")
        try:
            packet = validate_control(plan, child_dir, terminal, control_receipt)
        except Exception as exc:
            raise RunnerError(f"{child_id}: control raw gate failed: {exc}") from exc
    else:
        packet = None
    try:
        if not control:
            packet = decoder.decode_child(plan, child_dir, terminal, launch)
    except Exception as exc:
        status = {"schema": CHILD_PROGRESS_SCHEMA, "packet_id": child_id, "status": "decoder_failure", "error": str(exc), "no_retry": True}
        write_status(child_dir, status)
        raise RunnerError(f"{child_id}: decoder failed: {exc}") from exc
    if not isinstance(packet, dict):
        raise RunnerError(f"{child_id}: decoder returned non-object packet")
    packet = dict(packet)
    packet.setdefault("schema", PACKET_SCHEMA)
    packet.setdefault("packet_id", child_id)
    # Decoder owns the offline perf report, exact terminal LOST_SAMPLES parse,
    # and the terminal-summary/loss receipts.  The runner only binds files it
    # actually received from that decoder; it never reconstructs those gates.
    dump_path = child_dir / f"{child_id}.perf-mmap-dump.txt"
    packet["runner_bindings"] = {name: custody.binding(child_dir / name) for name in
                                  (f"{child_id}.command.json", f"{child_id}.terminal.json", f"{child_id}.perf.data",
                                   f"{child_id}.perf.stderr", f"{child_id}.affinity-observations.json",
                                   f"{child_id}.bridge.json", f"{child_id}.report.json") + (() if control else (f"{child_id}.launch.json",))}
    if (child_dir / f"{child_id}.terminal-summary.json").is_file():
        packet["runner_bindings"]["terminal_summary"] = custody.binding(child_dir / f"{child_id}.terminal-summary.json")
    if (child_dir / f"{child_id}.loss.json").is_file():
        packet["runner_bindings"]["loss"] = custody.binding(child_dir / f"{child_id}.loss.json")
    if dump_path.is_file():
        packet["runner_bindings"]["raw_task"] = custody.binding(dump_path)
    passed, reasons = control_packet_gate(packet) if control else packet_gate(packet)
    packet["runner_gate"] = {"passed": passed, "reasons": reasons,
                              "affinity": terminal.get("affinity_gate"),
                              "process_group_cleanup": terminal.get("process_group_cleanup")}
    packet_path = child_dir / f"{child_id}.packet.json"
    write_exclusive(packet_path, packet)
    summary = packet_summary(packet)
    status = {"schema": CHILD_PROGRESS_SCHEMA, "packet_id": child_id, "status": "gated" if passed else "gate_failure",
              "marker_sample_count": summary.get("marker_sample_count"),
              "empty_callchain_count": summary.get("empty_callchain_count"),
              "raw_lost_sum": packet.get("gate", {}).get("raw_lost_sum"),
              "terminal_lost_samples": packet.get("gate", {}).get("terminal_lost_samples"),
              "decoder_gate": packet.get("gate"), "runner_gate": packet["runner_gate"], "no_retry": True}
    write_status(child_dir, status)
    if not passed:
        raise RunnerError(f"{child_id}: child gate failed: {reasons}")
    return packet


def check_go(plan_path: Path, go_path: Path, *, control: bool) -> dict[str, Any]:
    plan = strict_load(plan_path)
    custody.validate_seal(plan, plan_path)
    go = strict_load(go_path)
    if go.get("schema") != custody.GO_SCHEMA or go.get("authorized") is not True or go.get("status") != "authorized":
        raise RunnerError("root GO is not authorized")
    expected_mode = "control" if control else "office"
    if go.get("mode") != expected_mode:
        raise RunnerError(f"root GO mode is not {expected_mode}")
    if go.get("plan") != custody.binding(plan_path):
        raise RunnerError("root GO does not bind exact plan")
    # The helper seal is already part of the immutable plan.  Requiring the
    # GO to repeat the exact binding prevents a stale or substituted seal from
    # authorizing either mode.
    if go.get("helper_seal") != plan.get("helper_seal"):
        raise RunnerError("root GO does not bind the plan's exact helper seal")
    custody.check_binding(go["helper_seal"], "root_go.helper_seal")
    independent_review = go.get("independent_review")
    if not isinstance(independent_review, Mapping):
        raise RunnerError("root GO is missing independent_review")
    review_path = custody.check_binding(independent_review, "root_go.independent_review")
    review = strict_load(review_path)
    if review.get("status") != "pass":
        raise RunnerError("independent review is not passing")
    if review.get("plan") != custody.binding(plan_path):
        raise RunnerError("independent review does not bind the exact plan")
    if (go.get("quiescence") or {}).get("status") != "pass":
        raise RunnerError("root GO quiescence is not passing")
    if go.get("synthetic_boundary", {}).get("enabled") is True:
        raise RunnerError("synthetic boundary is never admissible for --run")
    if not control:
        for field in ("control_manifest", "control_audit"):
            row = go.get(field)
            if not isinstance(row, Mapping):
                raise RunnerError(f"office GO is missing post-control {field}")
            path = custody.check_binding(row, f"root_go.{field}")
            value = strict_load(path)
            if value.get("status") not in {"pass", "complete"}:
                raise RunnerError(f"post-control {field} is not passing")
        manifest = strict_load(custody.check_binding(go["control_manifest"], "root_go.control_manifest"))
        if manifest.get("mode") != "control" or manifest.get("failure") is not None or manifest.get("children_completed") != 1:
            raise RunnerError("post-control manifest is not a complete single-control packet")
    return plan


def run(plan_path: Path, go_path: Path, *, control: bool = False) -> int:
    plan = check_go(plan_path.resolve(), go_path.resolve(), control=control)
    capture_root = Path((plan["control"]["output_root"] if control else plan["outputs"]["capture_root"])).resolve()
    if capture_root.exists():
        raise RunnerError(f"fresh capture root already exists: {capture_root}")
    if custody.CPU not in sorted(os.sched_getaffinity(0)):
        raise RunnerError(f"CPU {custody.CPU} is not in runner affinity")
    capture_root.mkdir(parents=True)
    decoder_path = Path(plan["helpers"]["decoder"]["path"]).resolve()
    packet_paths: list[str] = []
    failure: dict[str, Any] | None = None
    requested_children = 1 if control else custody.CHILDREN
    for child in range(1, requested_children + 1):
        try:
            packet = run_child(plan, child, capture_root, decoder_path, control=control)
            packet_path = capture_root / f"child-{child:02d}/child-{child:02d}.packet.json"
            packet_paths.append(str(packet_path))
            summary = packet_summary(packet)
            print(json.dumps({"schema": CHILD_PROGRESS_SCHEMA, "child": child, "status": "gated",
                "marker_sample_count": summary.get("marker_sample_count"),
                              "empty_callchain_count": summary.get("empty_callchain_count"),
                              "raw_lost_sum": summary.get("raw_lost_sum", 0),
                              "terminal_lost_samples": summary.get("terminal_lost_samples", 0)}, sort_keys=True), flush=True)
        except Exception as exc:
            failure = {"child": child, "type": type(exc).__name__, "error": str(exc), "no_retry": True}
            print(json.dumps({"schema": CHILD_PROGRESS_SCHEMA, "child": child, "status": "failed", "error": str(exc)}, sort_keys=True), flush=True)
            break
    manifest = {"schema": MANIFEST_SCHEMA, "status": "failed" if failure else "complete", "created_utc": utc(),
                "plan": custody.binding(plan_path), "root_go": custody.binding(go_path), "children_requested": requested_children,
                "children_completed": len(packet_paths), "serial": True, "no_retry": True, "packets": packet_paths, "failure": failure}
    manifest["mode"] = "control" if control else "office"
    write_exclusive(capture_root / "capture-manifest.json", manifest)
    print(json.dumps(manifest, indent=2, sort_keys=True))
    return 2 if failure else 0


def self_test() -> dict[str, Any]:
    command_plan = {
        "binary": {"path": str(custody.BINARY)}, "fixture": {"path": str(custody.FIXTURE)},
        "python": {"path": str(custody.PYTHON), "sha256": custody.sha256(custody.PYTHON)},
        "launcher": {"path": str(LAUNCHER), "sha256": custody.sha256(LAUNCHER)},
        "perf": {"path": str(custody.PERF), "sha256": custody.sha256(custody.PERF)},
        "taskset": {"path": str(custody.TASKSET)},
        "workload": {"case": "docfloat", "edits": 1, "retention": "default"},
    }
    import tempfile
    with tempfile.TemporaryDirectory(prefix="litchi-0977-command-") as temp:
        child = Path(temp) / "child-01"
        child.mkdir()
        command = record_argv(command_plan, child)
        if command[0:3] != [str(custody.TASKSET), "--cpu-list", "31"]:
            raise RunnerError("taskset command self-test failed")
        if "--per-thread" not in command or command[command.index("-m") + 1] != "16M":
            raise RunnerError("record mode self-test failed")
        if "-a" in command or "--all-cpus" in command or "--inherit" in command:
            raise RunnerError("forbidden perf token self-test failed")
        if command.count(str(BRIDGE.resolve())) != 1 or command.count(str(LAUNCHER.resolve())) != 2:
            raise RunnerError("same-PID bridge/launcher chain self-test failed")
    good_packet = {"status": "complete", "perf": {"corrected": {"marker_sample_count": 100,
                    "marker_period_share": 0.1, "empty_callchain_count": 0, "sample_count": 1}},
                   "gate": {"passed": True, "raw_lost_sum": 0, "terminal_lost_samples": 0}}
    passed, _ = packet_gate(good_packet)
    if not passed:
        raise RunnerError("complete packet gate self-test failed")
    bad_packet = {**good_packet, "perf": {"corrected": {**good_packet["perf"]["corrected"], "empty_callchain_count": 1}}}
    passed, reasons = packet_gate(bad_packet)
    if passed or "empty_callchain_count_gate_failed" not in reasons:
        raise RunnerError("empty-callchain gate self-test failed")
    bridge = {"status": "ready_to_exec", "native_exec_same_pid": True, "bridge_pid": 50,
              "bridge_starttime": 500, "parent_after": {"pid": 10, "uid": 1, "starttime": 100,
              "comm": "perf", "exe": "/usr/bin/perf"},
              "tasks_after": [{"tid": 10, "starttime": 100, "affinity": [30]}]}
    observation = {"rows": [{"phase": "periodic", "bridge_handoff_ready": True,
                  "perf_parent": {"status": "present", "identity": bridge["parent_after"],
                                   "tasks": [{"tid": 10, "starttime": 100, "affinity": [30]}]},
                  "target": {"status": "present", "identity": {"pid": 50, "uid": 1,
                              "starttime": 500, "comm": "python3.14", "exe": "/usr/bin/python3.14"},
                             "tasks": [{"tid": 50, "starttime": 500, "affinity": [31]}]}}]}
    launch = {"exec_pid": 50}
    if not affinity_observation_gate(observation, bridge, launch, control=False).get("passed"):
        raise RunnerError("affinity observation self-test failed")
    bad_observation = json.loads(json.dumps(observation))
    bad_observation["rows"][0]["target"]["tasks"][0]["affinity"] = [30]
    try:
        affinity_observation_gate(bad_observation, bridge, launch, control=False)
    except RunnerError:
        pass
    else:
        raise RunnerError("affinity mutation was accepted")
    return {"schema": "litchi.performance.0977.runner_self_test.v1", "status": "pass", "workload_started": False,
            "bridge_exec_chain": "same_pid_bridge_then_0970_launcher_then_target", "serial_children": custody.CHILDREN,
            "decoder_owns_loss_and_perf_decode": True}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--check-plan", type=Path)
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--control", action="store_true", help="run the separately bound harmless stdlib control")
    ap.add_argument("--plan", type=Path)
    ap.add_argument("--go", type=Path)
    args = ap.parse_args()
    try:
        if args.self_test:
            if any((args.check_plan, args.run, args.plan, args.go, args.control)):
                raise RunnerError("--self-test is standalone")
            print(json.dumps(self_test(), indent=2, sort_keys=True))
            return 0
        if args.check_plan:
            if args.run or args.plan or args.go or args.control:
                raise RunnerError("--check-plan is standalone")
            result = custody.validate_plan(args.check_plan.resolve())
            print(json.dumps(result, indent=2, sort_keys=True))
            return 0
        if args.run:
            if args.plan is None or args.go is None:
                raise RunnerError("--run requires --plan and --go")
            return run(args.plan, args.go, control=args.control)
        raise RunnerError("choose --self-test, --check-plan, or --run")
    except Exception as exc:
        print(f"0977 runner rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
