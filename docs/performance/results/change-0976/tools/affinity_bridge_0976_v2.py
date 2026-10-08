#!/usr/bin/env python3
"""Continuous-mode perf affinity handoff bridge for the 0976 control.

The bridge is intended to be started as the child command of an outer
``taskset --cpu-list 31 perf record`` invocation.  It performs one bounded
same-PID handoff: it verifies the perf parent, moves every currently observed
perf TID to CPU 30, verifies the complete set again, and then execs the
bounded stdlib target while retaining CPU 31.  It never moves the target and
it never silently falls back to a leader-only operation.  Continuous perf
enable-on-exec remains in force; startup/handoff samples are retained and
classified as startup records rather than filtered by a FIFO boundary.

No perf process is started by this module's self-test.  A real control must
be admitted by the separately sealed 0976 plan and root GO receipt.
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
from typing import Any, Iterable


SCHEMA = "litchi.performance.0976.affinity_bridge.v2"
PERF_COMM = "perf"
DEFAULT_TARGET_CPU = 31
DEFAULT_RECORDER_CPU = 30
_STAT_RE = re.compile(r"^\d+ \(.*\) ")


class BridgeError(RuntimeError):
    """A fail-closed bridge admission or handoff error."""


def utc_now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def mono_ns() -> int:
    return time.monotonic_ns()


def sha256_file(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _stat_fields(pid: int) -> list[str]:
    path = Path(f"/proc/{pid}/stat")
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise BridgeError(f"cannot read {path}: {exc}") from exc
    close = text.rfind(")")
    if close < 0 or not _STAT_RE.match(text):
        raise BridgeError(f"malformed proc stat for pid {pid}")
    fields = text[close + 2 :].split()
    # fields[0] is field 3 (state); field 22 is therefore index 19.
    if len(fields) <= 19:
        raise BridgeError(f"short proc stat for pid {pid}")
    return fields


def starttime(pid: int) -> int:
    try:
        return int(_stat_fields(pid)[19])
    except (TypeError, ValueError) as exc:
        raise BridgeError(f"invalid starttime for pid {pid}") from exc


def proc_exe(pid: int) -> str:
    try:
        return os.readlink(f"/proc/{pid}/exe")
    except OSError as exc:
        raise BridgeError(f"cannot read exe for pid {pid}: {exc}") from exc


def proc_comm(pid: int) -> str:
    try:
        return Path(f"/proc/{pid}/comm").read_text(encoding="utf-8").strip()
    except OSError as exc:
        raise BridgeError(f"cannot read comm for pid {pid}: {exc}") from exc


def exact_mask(pid: int) -> list[int]:
    try:
        return sorted(os.sched_getaffinity(pid))
    except OSError as exc:
        raise BridgeError(f"cannot read affinity for tid {pid}: {exc}") from exc


def identity(pid: int) -> dict[str, Any]:
    return {
        "pid": pid,
        "uid": os.stat(f"/proc/{pid}").st_uid,
        "starttime": starttime(pid),
        "comm": proc_comm(pid),
        "exe": proc_exe(pid),
    }


def task_ids(pid: int) -> list[int]:
    root = Path(f"/proc/{pid}/task")
    try:
        values = sorted(int(entry.name) for entry in root.iterdir())
    except (OSError, ValueError) as exc:
        raise BridgeError(f"cannot enumerate {root}: {exc}") from exc
    if not values:
        raise BridgeError(f"empty perf task set for pid {pid}")
    return values


def task_snapshot(pid: int) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for tid in task_ids(pid):
        rows.append({"tid": tid, "starttime": starttime(tid), "affinity": exact_mask(tid)})
    return rows


def _task_keys(rows: Iterable[dict[str, Any]]) -> list[tuple[int, int]]:
    return sorted((int(row["tid"]), int(row["starttime"])) for row in rows)


def require_exact_identity(observed: dict[str, Any], expected: dict[str, Any]) -> None:
    for key in ("pid", "uid", "starttime", "comm", "exe"):
        if observed.get(key) != expected.get(key):
            raise BridgeError(f"perf parent identity changed or mismatched at {key}")


def require_perf_parent(observed: dict[str, Any], expected_exe: str, expected_sha256: str) -> None:
    if observed["uid"] != os.getuid():
        raise BridgeError(f"unexpected perf parent uid: {observed['uid']}")
    if observed["comm"] != PERF_COMM:
        raise BridgeError(f"unexpected perf parent comm: {observed['comm']!r}")
    if os.path.realpath(observed["exe"]) != os.path.realpath(expected_exe):
        raise BridgeError(f"unexpected perf parent exe: {observed['exe']!r}")
    actual = sha256_file(expected_exe)
    if actual != expected_sha256:
        raise BridgeError("perf executable SHA-256 does not match the bound plan")


def require_bound_executable(observed: dict[str, Any], expected_exe: str, expected_sha256: str, label: str) -> None:
    if os.path.realpath(observed["exe"]) != os.path.realpath(expected_exe):
        raise BridgeError(f"{label} executable path is not bound: {observed['exe']!r}")
    if sha256_file(expected_exe) != expected_sha256:
        raise BridgeError(f"{label} executable SHA-256 does not match the bound plan")


def require_rows(rows: list[dict[str, Any]], expected_mask: list[int]) -> None:
    if not rows:
        raise BridgeError("empty recorder task observation")
    for row in rows:
        observed = sorted(int(value) for value in row.get("affinity", []))
        if observed != expected_mask:
            raise BridgeError(
                f"recorder TID {row.get('tid')} affinity {observed} != {expected_mask}"
            )


def write_exclusive(path: str | Path, payload: dict[str, Any]) -> None:
    output = Path(path)
    if not output.parent.is_dir():
        raise BridgeError(f"receipt parent does not exist: {output.parent}")
    encoded = json.dumps(payload, indent=2, sort_keys=True) + "\n"
    try:
        with output.open("x", encoding="utf-8") as handle:
            handle.write(encoded)
    except FileExistsError as exc:
        raise BridgeError(f"receipt already exists: {output}") from exc
    except OSError as exc:
        raise BridgeError(f"cannot write receipt {output}: {exc}") from exc


def clean_exec_environment() -> dict[str, str]:
    """Return the small non-secret environment admitted by the control plan."""
    return {
        "PATH": "/usr/bin:/bin",
        "LANG": "C",
        "LC_ALL": "C",
        "TZ": "UTC",
        "PYTHONDONTWRITEBYTECODE": "1",
    }


def handoff(args: argparse.Namespace) -> int:
    if args.target_cpu != DEFAULT_TARGET_CPU or args.recorder_cpu != DEFAULT_RECORDER_CPU:
        raise BridgeError("0976 control is frozen to target CPU31 and recorder CPU30")
    own_pid = os.getpid()
    own_start = starttime(own_pid)
    own_before = identity(own_pid)
    require_bound_executable(
        own_before,
        args.expected_interpreter_exe,
        args.expected_interpreter_sha256,
        "bridge interpreter",
    )
    if exact_mask(own_pid) != [args.target_cpu]:
        raise BridgeError("bridge did not start with exactly target CPU31 affinity")

    parent_pid = os.getppid()
    parent_before_identity = identity(parent_pid)
    require_perf_parent(parent_before_identity, args.expected_perf_exe, args.expected_perf_sha256)
    parent_before_tasks = task_snapshot(parent_pid)
    require_rows(parent_before_tasks, [args.target_cpu])
    parent_before_keys = _task_keys(parent_before_tasks)
    args.target = [
        args.target_exe,
        args.target_script,
        "--receipt",
        args.target_receipt,
        "--expected-pid",
        str(own_pid),
        "--expected-starttime",
        str(own_start),
        "--expected-parent-pid",
        str(parent_pid),
        "--expected-parent-starttime",
        str(parent_before_identity["starttime"]),
        "--expected-parent-uid",
        str(parent_before_identity["uid"]),
        "--expected-parent-exe",
        parent_before_identity["exe"],
        "--expected-exe",
        args.expected_interpreter_exe,
        "--expected-exe-sha256",
        args.expected_interpreter_sha256,
        "--expected-task-keys",
        json.dumps(parent_before_keys, separators=(",", ":")),
        "--recorder-cpu",
        str(args.recorder_cpu),
        "--cpu",
        str(args.target_cpu),
        *args.target_tail,
    ]

    events: list[dict[str, Any]] = [
        {"name": "bridge_started", "monotonic_ns": mono_ns()},
        {"name": "parent_identified", "monotonic_ns": mono_ns(), "identity": parent_before_identity},
        {"name": "parent_tasks_before", "monotonic_ns": mono_ns(), "tasks": parent_before_tasks},
    ]
    receipt: dict[str, Any] = {
        "schema": SCHEMA,
        "status": "failed_pre_exec",
        "created_utc": utc_now(),
        "bridge_pid": own_pid,
        "bridge_starttime": own_start,
        "expected_target_cpu": args.target_cpu,
        "expected_recorder_cpu": args.recorder_cpu,
        "expected_perf_exe": str(Path(args.expected_perf_exe).resolve()),
        "expected_perf_sha256": args.expected_perf_sha256,
        "bridge_interpreter": {
            "path": str(Path(args.expected_interpreter_exe).resolve()),
            "sha256": args.expected_interpreter_sha256,
            "observed": own_before,
        },
        "parent_before": parent_before_identity,
        "tasks_before": parent_before_tasks,
        "events": events,
        "target_argv": list(args.target),
        "native_exec_same_pid": False,
    }
    try:
        for row in parent_before_tasks:
            try:
                os.sched_setaffinity(int(row["tid"]), {args.recorder_cpu})
            except OSError as exc:
                raise BridgeError(f"cannot move perf TID {row['tid']} to CPU30: {exc}") from exc
        events.append({"name": "parent_affinity_set", "monotonic_ns": mono_ns()})

        parent_after_identity = identity(parent_pid)
        require_exact_identity(parent_after_identity, parent_before_identity)
        parent_after_tasks = task_snapshot(parent_pid)
        if _task_keys(parent_after_tasks) != parent_before_keys:
            raise BridgeError("perf TID set changed during affinity handoff")
        require_rows(parent_after_tasks, [args.recorder_cpu])
        events.append({"name": "parent_tasks_after", "monotonic_ns": mono_ns(), "tasks": parent_after_tasks})
        if exact_mask(own_pid) != [args.target_cpu]:
            raise BridgeError("bridge affinity changed during recorder handoff")

        parent_final_identity = identity(parent_pid)
        require_exact_identity(parent_final_identity, parent_before_identity)
        parent_final_tasks = task_snapshot(parent_pid)
        if _task_keys(parent_final_tasks) != parent_before_keys:
            raise BridgeError("perf TID set changed before native exec")
        require_rows(parent_final_tasks, [args.recorder_cpu])
        if exact_mask(own_pid) != [args.target_cpu] or starttime(own_pid) != own_start:
            raise BridgeError("bridge PID/starttime/affinity changed before exec")

        receipt.update(
            {
                "status": "ready_to_exec",
                "parent_after": parent_after_identity,
                "tasks_after": parent_after_tasks,
                "parent_final": parent_final_identity,
                "tasks_final": parent_final_tasks,
                "events": events,
                "continuous_event_mode": True,
                "startup_handoff_samples_retained": True,
                "native_exec_same_pid": True,
            }
        )
        write_exclusive(args.bridge_receipt, receipt)
        os.execve(args.target[0], list(args.target), clean_exec_environment())
    except Exception as exc:
        receipt["error"] = str(exc)
        receipt["events"] = events
        try:
            write_exclusive(args.bridge_receipt, receipt)
        except Exception as write_exc:
            print(f"bridge failed and could not write receipt: {write_exc}", file=sys.stderr)
        raise
    raise BridgeError("execve unexpectedly returned")


def self_test() -> dict[str, Any]:
    accepted = {"pid": 123, "uid": 1000, "starttime": 456, "comm": "perf", "exe": "/usr/bin/perf"}
    require_exact_identity(dict(accepted), accepted)
    rejected = dict(accepted)
    rejected["starttime"] = 457
    try:
        require_exact_identity(rejected, accepted)
    except BridgeError:
        identity_negative = "pass"
    else:
        raise BridgeError("identity starttime negative control was accepted")
    require_rows(
        [{"tid": 123, "starttime": 456, "affinity": [30]}, {"tid": 124, "starttime": 457, "affinity": [30]}],
        [30],
    )
    try:
        require_rows([{"tid": 123, "starttime": 456, "affinity": [31]}], [30])
    except BridgeError:
        affinity_negative = "pass"
    else:
        raise BridgeError("affinity mismatch negative control was accepted")
    return {
        "schema": SCHEMA,
        "status": "pass",
        "perf_started": False,
        "workload_started": False,
        "identity_starttime_negative": identity_negative,
        "affinity_mask_negative": affinity_negative,
        "contract": {
            "target_cpu": DEFAULT_TARGET_CPU,
            "recorder_cpu": DEFAULT_RECORDER_CPU,
            "all_tids_required": True,
            "continuous_event_mode": True,
            "startup_handoff_samples_retained": True,
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--expected-perf-exe", default="/usr/bin/perf")
    parser.add_argument("--expected-perf-sha256")
    parser.add_argument("--expected-interpreter-exe", default="/usr/bin/python3.14")
    parser.add_argument("--expected-interpreter-sha256")
    parser.add_argument("--target-cpu", type=int, default=DEFAULT_TARGET_CPU)
    parser.add_argument("--recorder-cpu", type=int, default=DEFAULT_RECORDER_CPU)
    parser.add_argument("--bridge-receipt")
    parser.add_argument("--target-exe")
    parser.add_argument("--target-script")
    parser.add_argument("--target-receipt")
    parser.add_argument("target_args", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    if args.self_test:
        print(json.dumps(self_test(), indent=2, sort_keys=True))
        return 0
    if not args.expected_perf_sha256:
        raise BridgeError("--expected-perf-sha256 is required for a real handoff")
    if not args.expected_interpreter_sha256:
        raise BridgeError("--expected-interpreter-sha256 is required for a real handoff")
    required = {
        "--bridge-receipt": args.bridge_receipt,
        "--target-exe": args.target_exe,
        "--target-receipt": args.target_receipt,
        "--target-script": args.target_script,
    }
    missing = [name for name, value in required.items() if not value]
    if missing:
        raise BridgeError(f"missing required bridge arguments: {missing}")
    # argparse leaves the separator in REMAINDER; it is never part of target argv.
    target_tail = list(args.target_args)
    if target_tail[:1] == ["--"]:
        target_tail = target_tail[1:]
    args.target_tail = target_tail
    try:
        return handoff(args)
    except Exception as exc:
        # Preserve a bounded failed-pre-exec identity receipt even when the
        # mismatch occurs before handoff() has collected its normal snapshot.
        if args.bridge_receipt:
            fallback = {
                "schema": SCHEMA,
                "status": "failed_pre_exec",
                "created_utc": utc_now(),
                "bridge_pid": os.getpid(),
                "parent_pid_observed": os.getppid(),
                "expected_target_cpu": args.target_cpu,
                "expected_recorder_cpu": args.recorder_cpu,
                "expected_perf_exe": str(Path(args.expected_perf_exe).resolve()),
                "expected_perf_sha256": args.expected_perf_sha256,
                "expected_interpreter_exe": str(Path(args.expected_interpreter_exe).resolve()),
                "expected_interpreter_sha256": args.expected_interpreter_sha256,
                "error": str(exc),
                "native_exec_same_pid": False,
            }
            try:
                write_exclusive(args.bridge_receipt, fallback)
            except Exception:
                pass
        print(f"0976 affinity bridge rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
