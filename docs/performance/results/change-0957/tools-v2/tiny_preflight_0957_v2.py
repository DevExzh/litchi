#!/usr/bin/env python3
"""One real-binary, one-sample 0957 v2 contract preflight.

It validates the authorized formal schedule first, executes exactly one selected
fresh process, seals terminal custody before report parsing, and retains any
failure without retry.  It is not the formal 27-child runner.
"""
from __future__ import annotations

import argparse
import json
import os
import signal
import subprocess
import sys
import time
from pathlib import Path

from common_0957_v2 import (
    ContractError,
    exclusive_json,
    file_descriptor,
    load_json,
    sha256_file,
    validate_plan_schedule,
    validate_report,
)
from capture_0957_v2 import (
    load_and_check_plan,
    child_argv,
    child_environment,
    external_descriptor,
    process_group_closed,
)


DEADLINE = 900


def fail(message: str) -> "NoReturn":
    raise ContractError(message)


def require_report_directory_absent(report_dir: Path) -> None:
    """Require the child, rather than the runner, to create its output dir."""
    if report_dir.exists():
        fail("tiny preflight report directory unexpectedly exists")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--case-id", default="medium-one-percent")
    parser.add_argument("--workers", type=int, default=1)
    args = parser.parse_args()
    plan_path = args.plan.resolve()
    packet = args.packet.resolve()
    try:
        plan, children = load_and_check_plan(plan_path)
        if plan.get("status") != "authorized":
            fail("real tiny preflight requires an authorized formal plan")
        for key in ("source", "build", "binary", "runner"):
            external_descriptor(key, plan[key])
        selected = [
            child for child in children
            if child["case_id"] == args.case_id and child["workers"] == args.workers
        ]
        if not selected:
            fail("requested tiny preflight case/worker is not in the formal schedule")
        child = dict(selected[0], index=0, repeat=0)
        tiny = json.loads(json.dumps(plan))
        tiny["revision"] = "real-tiny-preflight-v2"
        tiny["schedule"]["repeats"] = 1
        tiny["schedule"]["measured_samples"] = 1
        tiny["children"] = [child]
        if packet.exists():
            fail(f"preflight packet already exists: {packet}")
        packet.mkdir(parents=True)
        for name in ("custody", "children", "reports"):
            (packet / name).mkdir()
        custody = packet / "children" / child["child_id"]
        custody.mkdir()
        report_dir = packet / "reports" / child["child_id"]
        # The benchmark owns creation of its output directory.  Creating it
        # here would make the preflight exercise a different filesystem
        # contract from the formal runner, which only creates reports/.
        require_report_directory_absent(report_dir)
        exclusive_json(packet / "custody" / "plan.json", tiny)
        plan_descriptor = file_descriptor(packet / "custody" / "plan.json", relative_to=packet)
        argv = child_argv(child, tiny, report_dir)
        env = child_environment(tiny)
        cwd = Path(tiny["execution"]["cwd"])
        if cwd.is_symlink() or not cwd.is_dir():
            fail("preflight cwd is not a regular directory")
        start_record = {
            "schema": "litchi.performance.0957.tiny-start.v2",
            "child": child,
            "index": 0,
            "argv": argv,
            "cwd": str(cwd),
            "environment": env,
            "deadline_seconds": DEADLINE,
            "output_directory": str(report_dir),
            "process_group": True,
            "plan": plan_descriptor,
            "started_unix_ns": time.time_ns(),
        }
        exclusive_json(custody / "start.json", start_record)
        stdout_path = custody / "stdout.log"
        stderr_path = custody / "stderr.log"
        process: subprocess.Popen[bytes] | None = None
        returncode: int | None = None
        timed_out = False
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            process = subprocess.Popen(
                argv,
                cwd=str(cwd),
                env=env,
                stdin=subprocess.DEVNULL,
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
            )
            deadline = time.monotonic() + DEADLINE
            while True:
                left = deadline - time.monotonic()
                if left <= 0:
                    timed_out = True
                    break
                try:
                    returncode = process.wait(timeout=min(1.0, left))
                    break
                except subprocess.TimeoutExpired:
                    continue
            if timed_out and process.poll() is None:
                os.killpg(process.pid, signal.SIGTERM)
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    os.killpg(process.pid, signal.SIGKILL)
                    process.wait(timeout=5)
                returncode = process.returncode
        forced_group_cleanup = False
        group_closed = process is not None and process_group_closed(process)
        if process is not None and not group_closed:
            # Keep the one-child preflight contained just like the formal
            # runner.  A live descendant is a terminal failure, even when the
            # leader already returned; cleanup is recorded rather than hidden.
            forced_group_cleanup = True
            try:
                os.killpg(process.pid, signal.SIGTERM)
                process.wait(timeout=5)
            except Exception:
                try:
                    os.killpg(process.pid, signal.SIGKILL)
                    process.wait(timeout=5)
                except Exception:
                    pass
            group_closed = process_group_closed(process)
        report_path = report_dir / "report.json"
        report_descriptor = file_descriptor(report_path, relative_to=packet) if report_path.is_file() else None
        terminal = {
            "schema": "litchi.performance.0957.tiny-terminal.v2",
            "child": child,
            "index": 0,
            "argv": argv,
            "cwd": str(cwd),
            "environment": env,
            "deadline_seconds": DEADLINE,
            "returncode": returncode,
            "timed_out": timed_out,
            "leader_reaped": process is not None and process.poll() is not None,
            "process_group_reaped": group_closed,
            "forced_group_cleanup": forced_group_cleanup,
            "terminal_unix_ns": time.time_ns(),
            "stdout": file_descriptor(stdout_path, relative_to=packet),
            "stderr": file_descriptor(stderr_path, relative_to=packet),
            "report_path": str(report_path),
            "report_directory": str(report_dir),
            "report": report_descriptor,
            "plan": plan_descriptor,
        }
        exclusive_json(custody / "terminal.json", terminal)
        terminal_descriptor = file_descriptor(custody / "terminal.json", relative_to=packet)
        if timed_out or returncode != 0 or not group_closed or forced_group_cleanup:
            result = {
                "schema": "litchi.performance.0957.tiny-preflight.v2",
                "status": "failed",
                "terminal": terminal_descriptor,
                "report_parsed": False,
                "error": f"returncode={returncode} timed_out={timed_out} group_closed={group_closed}",
            }
            exclusive_json(packet / "preflight-result.json", result)
            return 1
        if not report_path.is_file():
            result = {
                "schema": "litchi.performance.0957.tiny-preflight.v2",
                "status": "failed",
                "terminal": terminal_descriptor,
                "report_parsed": False,
                "error": "successful child exited without report.json",
            }
            exclusive_json(packet / "preflight-result.json", result)
            return 1
        report = load_json(report_path)
        validate_report(report, tiny, child)
        validation = {
            "schema": "litchi.performance.0957.tiny-validation.v2",
            "status": "passed",
            "child": child,
            "index": 0,
            "plan": plan_descriptor,
            "terminal": terminal_descriptor,
            "terminal_sha256": sha256_file(custody / "terminal.json"),
            "report": file_descriptor(report_path, relative_to=packet),
        }
        exclusive_json(custody / "validation.json", validation)
        result = {
            "schema": "litchi.performance.0957.tiny-preflight.v2",
            "status": "passed",
            "plan": file_descriptor(plan_path),
            "terminal": terminal_descriptor,
            "report": file_descriptor(report_path, relative_to=packet),
            "validation": file_descriptor(custody / "validation.json", relative_to=packet),
            "child": child,
            "warmup_iterations": 2,
            "measured_samples": 1,
        }
        exclusive_json(packet / "preflight-result.json", result)
        print(json.dumps(result, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"tiny_preflight_0957_v2: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
