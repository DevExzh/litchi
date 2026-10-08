#!/usr/bin/env python3
"""Generic 0977 same-PID affinity handoff for an Office target.

The outer command is ``taskset --cpu-list 31 perf record --per-thread ... --
python3.14 this_file``.  The bridge verifies that parent perf is the bound
process, moves every observed parent perf TID to CPU30, verifies the complete
set again, and then execve's the retained 0970 launcher.  The launcher later
execve's the bound workload in this same PID.  No target-specific self-checks
are required: raw perf task records and the launch receipt are the lifetime
authority.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any

SCHEMA = "litchi.performance.0977.affinity_bridge.v1"
TARGET_CPU = 31
RECORDER_CPU = 30
PERF_COMM = "perf"
STAT_RE = re.compile(r"^\d+ \(.*\) ")


class BridgeError(RuntimeError):
    pass


def utc() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def mono_ns() -> int:
    return time.monotonic_ns()


def sha256(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def stat_fields(pid: int) -> list[str]:
    path = Path(f"/proc/{pid}/stat")
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise BridgeError(f"cannot read {path}: {exc}") from exc
    close = text.rfind(")")
    if close < 0 or not STAT_RE.match(text) or len(text[close + 2 :].split()) <= 19:
        raise BridgeError(f"malformed proc stat for {pid}")
    return text[close + 2 :].split()


def starttime(pid: int) -> int:
    try:
        return int(stat_fields(pid)[19])
    except (TypeError, ValueError) as exc:
        raise BridgeError(f"invalid starttime for {pid}") from exc


def proc_exe(pid: int) -> str:
    try:
        return os.readlink(f"/proc/{pid}/exe")
    except OSError as exc:
        raise BridgeError(f"cannot read exe for {pid}: {exc}") from exc


def proc_comm(pid: int) -> str:
    try:
        return Path(f"/proc/{pid}/comm").read_text(encoding="utf-8").strip()
    except OSError as exc:
        raise BridgeError(f"cannot read comm for {pid}: {exc}") from exc


def affinity(pid: int) -> list[int]:
    try:
        return sorted(os.sched_getaffinity(pid))
    except OSError as exc:
        raise BridgeError(f"cannot read affinity for {pid}: {exc}") from exc


def identity(pid: int) -> dict[str, Any]:
    return {
        "pid": pid,
        "uid": os.stat(f"/proc/{pid}").st_uid,
        "starttime": starttime(pid),
        "comm": proc_comm(pid),
        "exe": proc_exe(pid),
    }


def task_ids(pid: int) -> list[int]:
    try:
        result = sorted(int(entry.name) for entry in Path(f"/proc/{pid}/task").iterdir())
    except (OSError, ValueError) as exc:
        raise BridgeError(f"cannot enumerate perf tasks: {exc}") from exc
    if not result:
        raise BridgeError("perf has no observable task")
    return result


def task_snapshot(pid: int) -> list[dict[str, Any]]:
    return [{"tid": tid, "starttime": starttime(tid), "affinity": affinity(tid)} for tid in task_ids(pid)]


def task_keys(rows: list[dict[str, Any]]) -> list[list[int]]:
    return [[int(row["tid"]), int(row["starttime"])] for row in sorted(rows, key=lambda row: (int(row["tid"]), int(row["starttime"]))) ]


def require_identity(actual: dict[str, Any], expected: dict[str, Any], label: str) -> None:
    for key in ("pid", "uid", "starttime", "comm", "exe"):
        if actual.get(key) != expected.get(key):
            raise BridgeError(f"{label} identity changed at {key}")


def require_file(path: str | Path, expected_sha: str, label: str) -> None:
    actual = sha256(path)
    if actual != expected_sha:
        raise BridgeError(f"{label} SHA mismatch: {actual} != {expected_sha}")


def require_recorder(parent: dict[str, Any], expected_exe: str, expected_sha: str) -> None:
    if parent["uid"] != os.getuid() or parent["comm"] != PERF_COMM:
        raise BridgeError("parent is not the expected same-UID perf process")
    if os.path.realpath(parent["exe"]) != os.path.realpath(expected_exe):
        raise BridgeError(f"perf executable path drifted: {parent['exe']}")
    require_file(expected_exe, expected_sha, "perf")


def require_rows(rows: list[dict[str, Any]], expected: list[int]) -> None:
    if not rows or any(sorted(row.get("affinity", [])) != expected for row in rows):
        raise BridgeError(f"recorder task affinity is not exactly {expected}")


def write_exclusive(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n").encode()
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
    except OSError as exc:
        raise BridgeError(f"cannot create bridge receipt {path}: {exc}") from exc
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
    except Exception:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def clean_env() -> dict[str, str]:
    return {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "TZ": "UTC", "PYTHONDONTWRITEBYTECODE": "1"}


def handoff(args: argparse.Namespace) -> int:
    if args.target_cpu != TARGET_CPU or args.recorder_cpu != RECORDER_CPU:
        raise BridgeError("0977 is frozen to target CPU31 and recorder CPU30")
    own = os.getpid()
    own_start = starttime(own)
    if affinity(own) != [TARGET_CPU]:
        raise BridgeError("bridge did not inherit exactly CPU31")
    target = list(args.target)
    if target[:1] == ["--"]:
        target = target[1:]
    if args.control_target:
        if target:
            raise BridgeError("control target cannot also provide free-form argv")
        if not args.control_receipt:
            raise BridgeError("control target requires a receipt")
        target = [args.expected_target_exe, args.expected_target_script,
                  "--receipt", args.control_receipt, "--seconds", str(args.control_seconds),
                  "--check-interval", str(args.control_check_interval), "--cpu", str(TARGET_CPU),
                  "--recorder-cpu", str(RECORDER_CPU), "--expected-pid", str(own),
                  "--expected-starttime", str(own_start)]
    if not target:
        raise BridgeError("target argv is empty")
    if not target or os.path.realpath(target[0]) != os.path.realpath(args.expected_target_exe):
        raise BridgeError("target interpreter path is not bound")
    require_file(args.expected_target_exe, args.expected_target_sha256, "target interpreter")
    if args.expected_target_script:
        try:
            script_index = target.index(args.expected_target_script)
        except ValueError as exc:
            raise BridgeError("target launcher script is absent from target argv") from exc
        if script_index != 1:
            raise BridgeError("target launcher script is not argv[1]")
        require_file(args.expected_target_script, args.expected_target_script_sha256, "target launcher")

    parent_pid = os.getppid()
    before_identity = identity(parent_pid)
    require_recorder(before_identity, args.expected_perf_exe, args.expected_perf_sha256)
    before = task_snapshot(parent_pid)
    require_rows(before, [TARGET_CPU])
    before_keys = task_keys(before)
    if args.control_target:
        target.extend([
            "--expected-parent-pid", str(parent_pid),
            "--expected-parent-uid", str(before_identity["uid"]),
            "--expected-parent-starttime", str(before_identity["starttime"]),
            "--expected-parent-exe", before_identity["exe"],
            "--expected-exe", args.expected_target_exe,
            "--expected-exe-sha256", args.expected_target_sha256,
            "--expected-task-keys", json.dumps([tuple(row) for row in before_keys], separators=(",", ":")),
        ])
    events: list[dict[str, Any]] = [{"name": "bridge_started", "monotonic_ns": mono_ns()}]
    receipt: dict[str, Any] = {
        "schema": SCHEMA,
        "status": "failed_pre_exec",
        "created_utc": utc(),
        "bridge_pid": own,
        "bridge_starttime": own_start,
        "target_cpu": TARGET_CPU,
        "recorder_cpu": RECORDER_CPU,
        "parent_before": before_identity,
        "tasks_before": before,
        "target_argv": target,
        "native_exec_same_pid": False,
        "raw_authority": "raw zero-loss perf task stream; bridge masks are bounded observations",
        "events": events,
    }
    try:
        for row in before:
            os.sched_setaffinity(int(row["tid"]), {RECORDER_CPU})
        events.append({"name": "parent_tasks_moved", "monotonic_ns": mono_ns()})
        after_identity = identity(parent_pid)
        require_identity(after_identity, before_identity, "perf parent")
        after = task_snapshot(parent_pid)
        if task_keys(after) != before_keys:
            raise BridgeError("perf task set changed during handoff")
        require_rows(after, [RECORDER_CPU])
        if affinity(own) != [TARGET_CPU]:
            raise BridgeError("bridge affinity changed during handoff")
        final_identity = identity(parent_pid)
        require_identity(final_identity, before_identity, "perf parent before exec")
        final = task_snapshot(parent_pid)
        if task_keys(final) != before_keys:
            raise BridgeError("perf task set changed before exec")
        require_rows(final, [RECORDER_CPU])
        if starttime(own) != own_start or affinity(own) != [TARGET_CPU]:
            raise BridgeError("bridge PID/starttime/affinity changed before exec")
        receipt.update({
            "status": "ready_to_exec",
            "parent_after": after_identity,
            "tasks_after": after,
            "parent_final": final_identity,
            "tasks_final": final,
            "events": events,
            "native_exec_same_pid": True,
            "exec_preserves_pid_starttime_affinity": True,
            "startup_handoff_samples_retained": True,
        })
        write_exclusive(Path(args.bridge_receipt), receipt)
        os.execve(target[0], target, clean_env())
    except Exception as exc:
        receipt["error"] = str(exc)
        try:
            write_exclusive(Path(args.bridge_receipt), receipt)
        except Exception:
            pass
        raise
    raise BridgeError("execve returned unexpectedly")


def self_test() -> dict[str, Any]:
    expected = {"pid": 1, "uid": 2, "starttime": 3, "comm": "perf", "exe": "/usr/bin/perf"}
    require_identity(dict(expected), expected, "self-test")
    changed = dict(expected, starttime=4)
    try:
        require_identity(changed, expected, "self-test")
    except BridgeError:
        pass
    else:
        raise BridgeError("identity mutation accepted")
    require_rows([{"tid": 1, "starttime": 3, "affinity": [30]}], [30])
    try:
        require_rows([{"tid": 1, "starttime": 3, "affinity": [31]}], [30])
    except BridgeError:
        pass
    else:
        raise BridgeError("affinity mutation accepted")
    return {"schema": SCHEMA, "status": "pass", "workload_started": False, "same_pid_exec": True, "all_parent_tids_required": True}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--expected-perf-exe", default="/usr/bin/perf")
    parser.add_argument("--expected-perf-sha256")
    parser.add_argument("--expected-target-exe", default="/usr/bin/python3.14")
    parser.add_argument("--expected-target-sha256")
    parser.add_argument("--expected-target-script")
    parser.add_argument("--expected-target-script-sha256")
    parser.add_argument("--bridge-receipt")
    parser.add_argument("--control-target", action="store_true")
    parser.add_argument("--control-receipt")
    parser.add_argument("--control-seconds", type=float, default=0.25)
    parser.add_argument("--control-check-interval", type=float, default=0.02)
    parser.add_argument("--target-cpu", type=int, default=TARGET_CPU)
    parser.add_argument("--recorder-cpu", type=int, default=RECORDER_CPU)
    parser.add_argument("target", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    if args.self_test:
        print(json.dumps(self_test(), indent=2, sort_keys=True))
        return 0
    required = (args.expected_perf_sha256, args.expected_target_sha256, args.bridge_receipt, args.expected_target_script, args.expected_target_script_sha256)
    if any(not value for value in required):
        raise BridgeError("real handoff requires bound perf, launcher, and receipt arguments")
    try:
        if args.control_target and (not args.control_receipt or args.control_seconds <= 0 or args.control_seconds > 2 or args.control_check_interval <= 0):
            raise BridgeError("invalid bounded control target arguments")
        return handoff(args)
    except Exception as exc:
        print(f"0977 affinity bridge rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
