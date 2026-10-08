#!/usr/bin/env python3
"""Bounded stdlib-only target and continuous recorder-affinity check.

The target remains on CPU31 after the bridge exec.  During its short control
run it periodically verifies the exact perf parent identity and every perf
TID's CPU30 mask.  These observations prove the requested handoff only for
the control; they are not Office or baseline evidence.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import time
from pathlib import Path
from typing import Any


SCHEMA = "litchi.performance.0976.stdlib_control_target.v2"


class TargetError(RuntimeError):
    pass


def sha256_file(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def starttime(pid: int) -> int:
    text = Path(f"/proc/{pid}/stat").read_text(encoding="utf-8")
    close = text.rfind(")")
    fields = text[close + 2 :].split()
    if len(fields) <= 19:
        raise TargetError(f"short /proc/{pid}/stat")
    return int(fields[19])


def identity(pid: int) -> dict[str, Any]:
    return {
        "pid": pid,
        "uid": os.stat(f"/proc/{pid}").st_uid,
        "starttime": starttime(pid),
        "comm": Path(f"/proc/{pid}/comm").read_text(encoding="utf-8").strip(),
        "exe": os.readlink(f"/proc/{pid}/exe"),
        "affinity": sorted(os.sched_getaffinity(pid)),
    }


def task_snapshot(pid: int) -> list[dict[str, Any]]:
    tids = sorted(int(entry.name) for entry in Path(f"/proc/{pid}/task").iterdir())
    if not tids:
        raise TargetError("perf task set is empty")
    return [{"tid": tid, "starttime": starttime(tid), "affinity": sorted(os.sched_getaffinity(tid))} for tid in tids]


def check_parent(
    parent_pid: int,
    expected: dict[str, Any],
    recorder_cpu: int,
    expected_keys: list[tuple[int, int]],
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    parent = identity(parent_pid)
    for key in ("pid", "uid", "starttime", "comm", "exe"):
        if parent[key] != expected[key]:
            raise TargetError(f"perf parent identity changed at {key}")
    tasks = task_snapshot(parent_pid)
    keys = sorted((row["tid"], row["starttime"]) for row in tasks)
    if keys != expected_keys:
        raise TargetError("perf TID set changed during continuous check")
    for row in tasks:
        if row["affinity"] != [recorder_cpu]:
            raise TargetError(f"perf TID {row['tid']} is not exactly CPU{recorder_cpu}")
    return parent, tasks


def write_exclusive(path: Path, payload: dict[str, Any]) -> None:
    with path.open("x", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, sort_keys=True)
        handle.write("\n")


def run(args: argparse.Namespace) -> int:
    if args.cpu != 31 or args.recorder_cpu != 30:
        raise TargetError("continuous control is frozen to target CPU31 and recorder CPU30")
    own = identity(os.getpid())
    if os.path.realpath(own["exe"]) != os.path.realpath(args.expected_exe):
        raise TargetError(f"target interpreter path is {own['exe']!r}, expected {args.expected_exe!r}")
    if sha256_file(args.expected_exe) != args.expected_exe_sha256:
        raise TargetError("target interpreter SHA-256 does not match the bound plan")
    if own["affinity"] != [args.cpu]:
        raise TargetError(f"target affinity is {own['affinity']}, expected [{args.cpu}]")
    if own["pid"] != args.expected_pid or own["starttime"] != args.expected_starttime:
        raise TargetError("target PID/starttime does not match bridge identity")
    parent_pid = os.getppid()
    expected_parent = {
        "pid": args.expected_parent_pid,
        "uid": args.expected_parent_uid,
        "starttime": args.expected_parent_starttime,
        "comm": "perf",
        "exe": args.expected_parent_exe,
    }
    if parent_pid != args.expected_parent_pid:
        raise TargetError("target parent PID changed across exec")
    expected_keys = [tuple(item) for item in json.loads(args.expected_task_keys)]
    if not expected_keys:
        raise TargetError("empty expected perf TID set")
    parent_before, tasks_before = check_parent(parent_pid, expected_parent, args.recorder_cpu, expected_keys)
    started = time.monotonic_ns()
    checks: list[dict[str, Any]] = [
        {"name": "target_start", "monotonic_ns": started, "parent": parent_before, "tasks": tasks_before}
    ]
    checksum = 0x9E3779B1
    deadline = time.monotonic() + args.seconds
    iterations = 0
    next_check = time.monotonic()
    while time.monotonic() < deadline:
        checksum = (checksum * 1664525 + 1013904223 + iterations) & 0xFFFFFFFF
        checksum ^= ((checksum << 13) & 0xFFFFFFFF)
        checksum ^= checksum >> 17
        checksum ^= (checksum << 5) & 0xFFFFFFFF
        iterations += 1
        now = time.monotonic()
        if now >= next_check:
            parent, tasks = check_parent(parent_pid, expected_parent, args.recorder_cpu, expected_keys)
            checks.append({"name": "periodic", "monotonic_ns": time.monotonic_ns(), "parent": parent, "tasks": tasks})
            next_check = now + args.check_interval
    parent_after, tasks_after = check_parent(parent_pid, expected_parent, args.recorder_cpu, expected_keys)
    finished = identity(os.getpid())
    if finished["affinity"] != [args.cpu] or finished["pid"] != own["pid"] or finished["starttime"] != own["starttime"]:
        raise TargetError("target identity/affinity changed during burn")
    checks.append({"name": "target_end", "monotonic_ns": time.monotonic_ns(), "parent": parent_after, "tasks": tasks_after})
    payload = {
        "schema": SCHEMA,
        "status": "pass",
        "control_only": True,
        "native_workload_started": False,
        "target": finished,
        "target_executable": {
            "path": str(Path(args.expected_exe).resolve()),
            "sha256": args.expected_exe_sha256,
            "observed": finished,
        },
        "parent_perf": parent_after,
        "recorder_cpu": args.recorder_cpu,
        "target_cpu": args.cpu,
        "seconds_requested": args.seconds,
        "check_interval_seconds": args.check_interval,
        "parent_checks": checks,
        "elapsed_ns": time.monotonic_ns() - started,
        "iterations": iterations,
        "checksum_sha256": hashlib.sha256(str(checksum).encode("ascii")).hexdigest(),
    }
    write_exclusive(Path(args.receipt), payload)
    print(json.dumps(payload, sort_keys=True))
    return 0


def self_test() -> dict[str, Any]:
    value = 0x9E3779B1
    for index in range(100):
        value = (value * 1664525 + 1013904223 + index) & 0xFFFFFFFF
    return {
        "schema": SCHEMA,
        "status": "pass",
        "workload_started": False,
        "continuous_parent_checks": True,
        "checksum_probe": value,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--receipt")
    parser.add_argument("--seconds", type=float, default=0.25)
    parser.add_argument("--check-interval", type=float, default=0.02)
    parser.add_argument("--cpu", type=int, default=31)
    parser.add_argument("--recorder-cpu", type=int, default=30)
    parser.add_argument("--expected-pid", type=int)
    parser.add_argument("--expected-starttime", type=int)
    parser.add_argument("--expected-parent-pid", type=int)
    parser.add_argument("--expected-parent-uid", type=int)
    parser.add_argument("--expected-parent-starttime", type=int)
    parser.add_argument("--expected-parent-exe")
    parser.add_argument("--expected-exe")
    parser.add_argument("--expected-exe-sha256")
    parser.add_argument("--expected-task-keys")
    args = parser.parse_args()
    if args.self_test:
        print(json.dumps(self_test(), indent=2, sort_keys=True))
        return 0
    if not args.receipt or args.seconds <= 0 or args.seconds > 2 or args.check_interval <= 0:
        raise SystemExit("--receipt, bounded --seconds, and positive --check-interval are required")
    required = {
        "--expected-pid": args.expected_pid,
        "--expected-starttime": args.expected_starttime,
        "--expected-parent-pid": args.expected_parent_pid,
        "--expected-parent-uid": args.expected_parent_uid,
        "--expected-parent-starttime": args.expected_parent_starttime,
        "--expected-parent-exe": args.expected_parent_exe,
        "--expected-exe": args.expected_exe,
        "--expected-exe-sha256": args.expected_exe_sha256,
        "--expected-task-keys": args.expected_task_keys,
    }
    missing = [name for name, value in required.items() if value is None]
    if missing:
        raise SystemExit(f"missing required identity bindings: {missing}")
    return run(args)


if __name__ == "__main__":
    raise SystemExit(main())
