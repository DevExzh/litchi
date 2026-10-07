#!/usr/bin/env python3
"""Terminal-first 0957 allocator capture runner.

The runner is intentionally small: it executes only a plan whose external
source/build/binary identities and report contract are already frozen. It
writes exclusive custody records, stops on the first child failure, and never
retries a child. ``--dry-run`` validates the 27-child schedule without
starting a process.
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
from typing import Any

from common_0957_v2 import (
    ContractError,
    canonical_json,
    descriptor_matches,
    expected_children,
    exclusive_json,
    file_descriptor,
    load_json,
    packet_path,
    sha256_file,
    validate_plan_schedule,
    validate_report,
)

DEADLINE = 900


def fail(message: str) -> "NoReturn":
    raise ContractError(message)


def copy_exclusive(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with source.open("rb") as inp, destination.open("xb") as out:
        while True:
            block = inp.read(1024 * 1024)
            if not block:
                break
            out.write(block)
        out.flush()
        os.fsync(out.fileno())


def external_descriptor(label: str, value: Any) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"plan.{label} must be an object")
    path_value = value.get("path")
    if not isinstance(path_value, str) or not os.path.isabs(path_value):
        fail(f"plan.{label}.path must be absolute")
    path = Path(path_value)
    if path.is_symlink() or not path.is_file():
        fail(f"plan.{label}.path is not a regular file: {path}")
    actual = file_descriptor(path)
    for key in ("bytes", "sha256"):
        if value.get(key) != actual[key]:
            fail(f"plan.{label}.{key} does not match {path}")
    return actual


def load_and_check_plan(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    plan = load_json(path, max_bytes=4 * 1024 * 1024)
    if not isinstance(plan, dict):
        fail("plan must be an object")
    children = validate_plan_schedule(plan)
    if plan.get("schema") != "litchi.performance.0957.allocator-capture-plan.v1":
        fail("unexpected 0957 plan schema")
    if plan.get("status") not in ("prospective", "authorized"):
        fail("plan status must be prospective or authorized")
    for key in ("source", "binary", "build", "runner"):
        actual = external_descriptor(key, plan.get(key))
        if key == "binary" and plan.get("status") == "authorized":
            binary_path = Path(plan["binary"]["path"])
            if plan["binary"].get("executable") is not True or not os.access(binary_path, os.X_OK):
                fail("authorized plan binary must be executable")
    executing_runner = file_descriptor(Path(__file__).resolve())
    if any(plan["runner"].get(key) != executing_runner[key] for key in ("path", "bytes", "sha256")):
        fail("plan runner differs from the executing capture helper")
    contract = plan.get("report_contract")
    if not isinstance(contract, dict) or not isinstance(contract.get("schema"), str):
        fail("plan.report_contract.schema is missing")
    execution = plan.get("execution")
    if not isinstance(execution, dict):
        fail("plan.execution is missing")
    expected_template = [
        "/usr/bin/taskset", "--cpu-list", "12-15", "<BINARY>",
        "--shape", "<SHAPE>", "--workload", "<WORKLOAD>",
        "--workers", "<WORKERS>", "--warmup", "<WARMUP>",
        "--samples", "<SAMPLES>", "--output", "<OUTPUT>",
    ]
    if execution.get("taskset_path") != "/usr/bin/taskset" or execution.get("argv_template") != expected_template:
        fail("plan execution taskset or argv template is not the frozen command")
    expected_environment = {
        "LANG": "C",
        "LC_ALL": "C",
        "LITCHI_PERF_0957_CAPTURE_APPROVED": "1",
        "PATH": "/usr/bin:/bin",
        "RUST_BACKTRACE": "0",
        "TZ": "UTC",
    }
    if execution.get("environment") != expected_environment:
        fail("plan execution environment is not the frozen explicit environment")
    return plan, children


def expand_token(token: str, child: dict[str, Any], plan: dict[str, Any], output_dir: Path) -> str:
    case = next(case for case in plan["cases"] if case["case_id"] == child["case_id"])
    replacements = {
        "<BINARY>": str(Path(plan["binary"]["path"])),
        "<OUTPUT>": str(output_dir),
        "<SHAPE>": str(case["shape"]),
        "<WORKLOAD>": str(case["workload"]),
        "<WORKERS>": str(child["workers"]),
        "<WARMUP>": str(plan["schedule"]["warmup_iterations"]),
        "<SAMPLES>": str(plan["schedule"]["measured_samples"]),
    }
    result = token
    for marker, replacement in replacements.items():
        result = result.replace(marker, replacement)
    if "<" in result or ">" in result:
        fail(f"unexpanded command token: {token!r}")
    return result


def child_argv(child: dict[str, Any], plan: dict[str, Any], output_dir: Path) -> list[str]:
    execution = plan.get("execution")
    if not isinstance(execution, dict):
        fail("plan.execution must be an object")
    template = execution.get("argv_template")
    if not isinstance(template, list) or not template or not all(isinstance(x, str) for x in template):
        fail("plan.execution.argv_template must be a nonempty string list")
    argv = [expand_token(token, child, plan, output_dir) for token in template]
    if argv[0] != str(Path(plan["execution"].get("taskset_path", "/usr/bin/taskset"))):
        fail("argv must begin with the frozen taskset executable")
    return argv


def child_environment(plan: dict[str, Any]) -> dict[str, str]:
    execution = plan.get("execution")
    env = execution.get("environment") if isinstance(execution, dict) else None
    if not isinstance(env, dict) or not env:
        fail("plan.execution.environment must be an explicit nonempty mapping")
    result: dict[str, str] = {}
    for key, value in env.items():
        if not isinstance(key, str) or not isinstance(value, str):
            fail("plan.execution.environment keys and values must be strings")
        result[key] = value
    return result


def write_progress(packet: Path, index: int, child: dict[str, Any], status: str, error: str | None = None) -> None:
    value: dict[str, Any] = {
        "schema": "litchi.performance.0957.capture-progress.v1",
        "index": index,
        "child_id": child["child_id"],
        "status": status,
    }
    if error is not None:
        value["error"] = error
    exclusive_json(packet / "custody" / f"progress-{index:03d}.json", value)


def process_group_closed(process: subprocess.Popen[bytes], timeout: float = 5.0) -> bool:
    if process.poll() is None:
        return False
    deadline = time.monotonic() + timeout
    while True:
        try:
            os.killpg(process.pid, 0)
        except ProcessLookupError:
            return True
        except PermissionError:
            return False
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.05)


def run_child(index: int, child: dict[str, Any], plan: dict[str, Any], packet: Path) -> None:
    custody = packet / "children" / child["child_id"]
    report_dir = packet / "reports" / child["child_id"]
    custody.mkdir(parents=True, exist_ok=False)
    report_dir.parent.mkdir(parents=True, exist_ok=True)
    argv = child_argv(child, plan, report_dir)
    env = child_environment(plan)
    cwd = Path(plan["execution"]["cwd"])
    if cwd.is_symlink() or not cwd.is_dir():
        fail(f"frozen child cwd is not a directory: {cwd}")
    started = time.time_ns()
    start_record = {
        "schema": "litchi.performance.0957.child-start.v1",
        "index": index,
        "child": child,
        "argv": argv,
        "cwd": str(cwd),
        "environment": env,
        "deadline_seconds": DEADLINE,
        "output_directory": str(report_dir),
        "started_unix_ns": started,
        "process_group": True,
        "plan": file_descriptor(packet / "custody" / "plan.json", relative_to=packet),
    }
    exclusive_json(custody / "start.json", start_record)
    stdout_path = custody / "stdout.log"
    stderr_path = custody / "stderr.log"
    timed_out = False
    returncode: int | None = None
    process: subprocess.Popen[bytes] | None = None
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        try:
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
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    timed_out = True
                    break
                try:
                    returncode = process.wait(timeout=min(1.0, remaining))
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
        except BaseException:
            if process is not None and process.poll() is None:
                try:
                    os.killpg(process.pid, signal.SIGKILL)
                    process.wait(timeout=5)
                except Exception:
                    pass
            raise
    forced_group_cleanup = False
    group_closed = process is not None and process_group_closed(process)
    if process is not None and not group_closed:
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
        "schema": "litchi.performance.0957.child-terminal.v1",
        "index": index,
        "child": child,
        "argv": argv,
        "cwd": str(cwd),
        "environment": env,
        "deadline_seconds": DEADLINE,
        "returncode": returncode,
        "timed_out": timed_out,
        "terminal_unix_ns": time.time_ns(),
        "stdout": file_descriptor(stdout_path, relative_to=packet),
        "stderr": file_descriptor(stderr_path, relative_to=packet),
        "report_path": str(report_path),
        "report_directory": str(report_dir),
        "report": report_descriptor,
        "leader_reaped": process is not None and process.poll() is not None,
        "process_group_reaped": group_closed,
        "forced_group_cleanup": forced_group_cleanup,
        "plan": file_descriptor(packet / "custody" / "plan.json", relative_to=packet),
    }
    # This is the terminal-first boundary: report parsing happens only after
    # the immutable terminal receipt has been sealed.
    exclusive_json(custody / "terminal.json", terminal)
    if timed_out or returncode != 0 or not group_closed:
        write_progress(packet, index, child, "failed", f"child returncode={returncode} timed_out={timed_out}")
        fail(f"child {child['child_id']} failed before report validation")
    if not report_path.exists():
        write_progress(packet, index, child, "failed", "report.json is missing")
        fail(f"child {child['child_id']} did not produce report.json")
    report = load_json(report_path)
    validate_report(report, plan, child)
    validation = {
        "schema": "litchi.performance.0957.child-validation.v1",
        "status": "passed",
        "child": child,
        "terminal": file_descriptor(custody / "terminal.json", relative_to=packet),
        "report": file_descriptor(report_path, relative_to=packet),
        "terminal_sha256": sha256_file(custody / "terminal.json"),
        "plan": file_descriptor(packet / "custody" / "plan.json", relative_to=packet),
    }
    exclusive_json(custody / "validation.json", validation)
    write_progress(packet, index, child, "passed")


def packet_manifest(packet: Path) -> dict[str, Any]:
    entries = []
    for path in sorted(packet.rglob("*")):
        if path.name == "packet-manifest.json":
            continue
        if path.is_symlink():
            fail(f"packet contains symlink: {path}")
        if path.is_file():
            entries.append(file_descriptor(path, relative_to=packet))
        elif not path.is_dir():
            fail(f"packet contains nonregular artifact: {path}")
    return {
        "schema": "litchi.performance.0957.completed-packet-manifest.v1",
        "packet_root": str(packet),
        "entries": entries,
    }


def run(plan_path: Path, packet: Path, dry_run: bool) -> int:
    plan, children = load_and_check_plan(plan_path)
    if dry_run:
        print(json.dumps({"status": "dry_run_pass", "children": children}, sort_keys=True))
        return 0
    if plan.get("status") != "authorized":
        fail("formal capture requires plan.status=authorized")
    if packet.exists():
        fail(f"capture packet already exists: {packet}")
    packet.mkdir(parents=True)
    (packet / "children").mkdir()
    (packet / "reports").mkdir()
    (packet / "custody").mkdir()
    copy_exclusive(plan_path, packet / "custody" / "plan.json")
    exclusive_json(packet / "custody" / "runner.json", {
        "schema": "litchi.performance.0957.runner-binding.v1",
        "path": str(Path(__file__).resolve()),
        "bytes": Path(__file__).stat().st_size,
        "sha256": sha256_file(Path(__file__)),
    })
    runner_path = Path(__file__).resolve()
    runner_descriptor = file_descriptor(runner_path)
    exclusive_json(packet / "custody" / "capture-start.json", {
        "schema": "litchi.performance.0957.capture-start.v1",
        "plan": file_descriptor(plan_path),
        "copied_plan": file_descriptor(packet / "custody" / "plan.json", relative_to=packet),
        "runner": runner_descriptor,
        "children": children,
        "started_unix_ns": time.time_ns(),
    })
    completed = 0
    try:
        for index, child in enumerate(children):
            run_child(index, child, plan, packet)
            completed += 1
        exclusive_json(packet / "capture-state.json", {
            "schema": "litchi.performance.0957.capture-state.v1",
            "status": "completed",
            "children_expected": len(children),
            "children_completed": completed,
            "failed": 0,
        })
        exclusive_json(packet / "custody" / "outer-terminal.json", {
            "schema": "litchi.performance.0957.outer-terminal.v1",
            "status": "completed",
            "children_expected": len(children),
            "children_completed": completed,
            "plan": file_descriptor(plan_path),
            "copied_plan": file_descriptor(packet / "custody" / "plan.json", relative_to=packet),
            "runner": runner_descriptor,
            "terminal_unix_ns": time.time_ns(),
        })
        exclusive_json(packet / "packet-manifest.json", packet_manifest(packet))
        return 0
    except Exception as exc:
        try:
            exclusive_json(packet / "capture-state.json", {
                "schema": "litchi.performance.0957.capture-state.v1",
                "status": "failed",
                "children_expected": len(children),
                "children_completed": completed,
                "failed": 1,
                "error": str(exc),
            })
            exclusive_json(packet / "custody" / "outer-terminal.json", {
                "schema": "litchi.performance.0957.outer-terminal.v1",
                "status": "failed",
                "children_expected": len(children),
                "children_completed": completed,
                "plan": file_descriptor(plan_path),
                "copied_plan": file_descriptor(packet / "custody" / "plan.json", relative_to=packet),
                "runner": runner_descriptor,
                "terminal_unix_ns": time.time_ns(),
                "error": str(exc),
            })
        except Exception:
            pass
        raise


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    try:
        return run(args.plan.resolve(), args.packet.resolve(), args.dry_run)
    except Exception as exc:
        print(f"capture_0957_v2: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
