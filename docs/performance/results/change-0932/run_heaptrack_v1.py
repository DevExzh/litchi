#!/usr/bin/env python3
import hashlib
import json
import os
import signal
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path("/home/zhuhe/litchi-performance-0932")
PACKET = ROOT / "heaptrack-packet-v1"
SOURCE = Path("/home/zhuhe/litchi-performance-0931/source-final")
BINARY = Path("/home/zhuhe/litchi-performance-0931/release-target/release/litchi-perf-baseline-alloc")
REVISION = "a17d87da48504e8d1c1a6240a6c5d1b11219031f"
CPU = "12"
DEADLINE_SECONDS = 600
ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "RUST_BACKTRACE": "0",
    "LITCHI_PERF_0931_MEMORY_DIAGNOSTIC": "1",
}
OWNERSHIPS = ("shared_snapshot_clones", "independent_sources")


def now():
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


def write_xb(path, value):
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def group_closed(pid):
    if pid is None:
        return True
    try:
        os.killpg(pid, 0)
    except ProcessLookupError:
        return True
    except PermissionError:
        return False
    return False


def wait_group_closed(pid, seconds):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if group_closed(pid):
            return True
        time.sleep(0.05)
    return group_closed(pid)


def run_case(ownership):
    ident = f"s8192-c08-exact-{ownership}-heaptrack-v1"
    child = PACKET / "children" / ident
    child.mkdir(parents=True, exist_ok=False)
    report = child / "report.json"
    profile = child / "heaptrack-data.gz"
    stdout = child / "stdout.log"
    stderr = child / "stderr.log"
    argv = [
        "/usr/bin/time", "-v",
        "/usr/bin/taskset", "-c", CPU,
        "/usr/bin/heaptrack", "--record-only", "-o", str(profile),
        str(BINARY),
        "docx-section-retention",
        "--sections", "8192",
        "--concurrent-commits", "8",
        "--policy", "exact",
        "--ownership", ownership,
        "--source-revision", REVISION,
        "--output", str(report),
    ]
    started = now()
    monotonic_start = time.monotonic()
    pid = None
    exit_code = None
    timed_out = False
    runner_error = None
    with stdout.open("xb") as out, stderr.open("xb") as err:
        try:
            process = subprocess.Popen(
                argv,
                cwd=str(SOURCE),
                env=ENVIRONMENT,
                stdout=out,
                stderr=err,
                start_new_session=True,
            )
            pid = process.pid
            try:
                exit_code = process.wait(timeout=DEADLINE_SECONDS)
            except subprocess.TimeoutExpired:
                timed_out = True
                try:
                    os.killpg(pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass
                try:
                    exit_code = process.wait(timeout=15)
                except subprocess.TimeoutExpired:
                    try:
                        os.killpg(pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                    exit_code = process.wait(timeout=15)
        except Exception as exc:
            runner_error = f"{type(exc).__name__}: {exc}"
    group_was_closed = wait_group_closed(pid, 30)
    if not group_was_closed and pid is not None:
        try:
            os.killpg(pid, signal.SIGTERM)
        except ProcessLookupError:
            pass
        group_was_closed = wait_group_closed(pid, 5)
    ended = now()
    terminal = {
        "schema": "litchi.performance.0932.heaptrack-terminal.v1",
        "status": "terminal",
        "id": ident,
        "sections": 8192,
        "concurrent_commits": 8,
        "policy": "exact",
        "ownership": ownership,
        "source_revision": REVISION,
        "cwd": str(SOURCE),
        "argv": argv,
        "environment": ENVIRONMENT,
        "profiler": {
            "tool": "/usr/bin/heaptrack",
            "version": "1.5.0",
            "mode": "--record-only",
            "injection": "heaptrack-managed; no caller-provided LD_PRELOAD",
        },
        "started_utc": started,
        "ended_utc": ended,
        "wall_seconds": time.monotonic() - monotonic_start,
        "deadline_seconds": DEADLINE_SECONDS,
        "pid": pid,
        "exit_code": exit_code,
        "timed_out": timed_out,
        "runner_error": runner_error,
        "process_group_closed": group_was_closed,
        "binary": {
            "path": str(BINARY),
            "bytes": BINARY.stat().st_size,
            "sha256": "ad4e6a43c81af7057a961abb5c5893ddde61c7823620a22327f672290d2208cf",
        },
        "source": {
            "root": str(SOURCE),
            "base_revision": REVISION,
            "source_receipt": "/home/zhuhe/litchi-performance-0931/measurement-source-final-receipt.json",
            "source_receipt_sha256": "00ea8b9687265abb210938d7f30ebb0af64bb0303e6ac1fab08af18e61de3b3f",
            "full_source_listing": "/home/zhuhe/litchi-performance-0931/source-final-list.jsonl",
            "full_source_listing_sha256": "c5f194fa5aeb6cee0d7abaa67eed29e88dd8e9c1aaf022dc7d38a0a1b130db8a",
            "overlay_receipt": "/home/zhuhe/litchi-performance-0931/receipts/source-closure-addendum-v3.json",
            "overlay_receipt_sha256": "31f931a671e806bcf9da1c3e7cb937e17df37910799a31c3f96c1b1f03676de5",
        },
        "artifacts": {
            "report": str(report),
            "profile": str(profile),
            "stdout": str(stdout),
            "stderr": str(stderr),
        },
    }
    # This is intentionally the first read/write of no report/profile contents.
    write_xb(child / "terminal.json", terminal)
    return terminal


def main():
    if PACKET.exists():
        raise SystemExit(f"refusing existing packet: {PACKET}")
    PACKET.mkdir(parents=True)
    (PACKET / "children").mkdir()
    started = now()
    monotonic_start = time.monotonic()
    terminals = []
    runner_error = None
    exit_code = 0
    try:
        for ownership in OWNERSHIPS:
            terminals.append(run_case(ownership))
    except BaseException as exc:
        exit_code = 1
        runner_error = f"{type(exc).__name__}: {exc}"
    ended = now()
    outer = {
        "schema": "litchi.performance.0932.heaptrack-runner-terminal.v1",
        "status": "terminal",
        "started_utc": started,
        "ended_utc": ended,
        "wall_seconds": time.monotonic() - monotonic_start,
        "exit_code": exit_code,
        "runner_error": runner_error,
        "children": [item["id"] for item in terminals],
        "child_terminals_written_before_analysis": True,
        "analysis_started": False,
    }
    write_xb(ROOT / "runner-terminal.json", outer)
    print(json.dumps({"packet": str(PACKET), "children": [item["id"] for item in terminals], "exit_code": exit_code}, sort_keys=True))
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
