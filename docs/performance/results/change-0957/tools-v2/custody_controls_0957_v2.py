#!/usr/bin/env python3
"""Offline custody controls for the 0957 v2 child boundary.

This file builds one synthetic, already-terminal child record.  It never
starts the allocator binary.  The positive path exercises the same reducer
``validate_child`` call used for a completed packet; the negative paths mutate
one sealed custody fact at a time and must be rejected.
"""
from __future__ import annotations

import copy
import json
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools-v2"
REVIEW = ROOT / "review"
PLAN_PATH = REVIEW / "plan-schedule-v5-v2-draft.json"
OUT = REVIEW / "custody-controls-v2-result-v3.json"
sys.path.insert(0, str(TOOLS))

from capture_0957_v2 import child_argv, child_environment  # noqa: E402
from common_0957_v2 import (  # noqa: E402
    ContractError,
    exclusive_json,
    file_descriptor,
    load_json,
    sha256_file,
    validate_plan_schedule,
)
from controls_0957_v2 import report as synthetic_report  # noqa: E402
from reduce_0957_v2 import validate_child  # noqa: E402
from tiny_preflight_0957_v2 import require_report_directory_absent  # noqa: E402


def _write_json(path: Path, value: Any) -> None:
    exclusive_json(path, value)


def _fixture(root: Path) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    plan = load_json(PLAN_PATH)
    children = validate_plan_schedule(plan)
    child = copy.deepcopy(children[0])
    packet = root / "packet"
    custody = packet / "custody"
    child_root = packet / "children" / child["child_id"]
    report_dir = packet / "reports" / child["child_id"]
    custody.mkdir(parents=True)
    child_root.mkdir(parents=True)
    report_dir.mkdir(parents=True)
    shutil.copy2(PLAN_PATH, custody / "plan.json")
    packet_plan = file_descriptor(custody / "plan.json", relative_to=packet)
    stdout = child_root / "stdout.log"
    stderr = child_root / "stderr.log"
    stdout.write_bytes(b"synthetic stdout\n")
    stderr.write_bytes(b"")
    report_path = report_dir / "report.json"
    _write_json(report_path, synthetic_report(plan, child))
    report_descriptor = file_descriptor(report_path, relative_to=packet)
    argv = child_argv(child, plan, report_dir)
    env = child_environment(plan)
    cwd = str(Path(plan["execution"]["cwd"]))
    start = {
        "schema": "litchi.performance.0957.child-start.v1",
        "index": child["index"],
        "child": child,
        "argv": argv,
        "cwd": cwd,
        "environment": env,
        "deadline_seconds": plan["schedule"]["deadline_seconds"],
        "output_directory": str(report_dir),
        "started_unix_ns": 1,
        "process_group": True,
        "plan": packet_plan,
    }
    _write_json(child_root / "start.json", start)
    terminal = {
        "schema": "litchi.performance.0957.child-terminal.v1",
        "index": child["index"],
        "child": child,
        "argv": argv,
        "cwd": cwd,
        "environment": env,
        "deadline_seconds": plan["schedule"]["deadline_seconds"],
        "returncode": 0,
        "timed_out": False,
        "terminal_unix_ns": 2,
        "stdout": file_descriptor(stdout, relative_to=packet),
        "stderr": file_descriptor(stderr, relative_to=packet),
        "report_path": str(report_path),
        "report_directory": str(report_dir),
        "report": report_descriptor,
        "leader_reaped": True,
        "process_group_reaped": True,
        "forced_group_cleanup": False,
        "plan": packet_plan,
    }
    _write_json(child_root / "terminal.json", terminal)
    terminal_descriptor = file_descriptor(child_root / "terminal.json", relative_to=packet)
    validation = {
        "schema": "litchi.performance.0957.child-validation.v1",
        "status": "passed",
        "child": child,
        "index": child["index"],
        "terminal": terminal_descriptor,
        "report": report_descriptor,
        "terminal_sha256": sha256_file(child_root / "terminal.json"),
        "plan": packet_plan,
    }
    _write_json(child_root / "validation.json", validation)
    return packet, plan, child


def _mutate_terminal(packet: Path, update: Callable[[dict[str, Any]], None]) -> None:
    path = packet / "children" / "medium-one-percent-w1-r0" / "terminal.json"
    value = load_json(path)
    update(value)
    path.unlink()
    _write_json(path, value)


def _mutate_start(packet: Path, update: Callable[[dict[str, Any]], None]) -> None:
    path = packet / "children" / "medium-one-percent-w1-r0" / "start.json"
    value = load_json(path)
    update(value)
    path.unlink()
    _write_json(path, value)


def _mutate_validation(packet: Path, update: Callable[[dict[str, Any]], None]) -> None:
    path = packet / "children" / "medium-one-percent-w1-r0" / "validation.json"
    value = load_json(path)
    update(value)
    path.unlink()
    _write_json(path, value)


def main() -> int:
    outcomes: dict[str, str] = {}

    with tempfile.TemporaryDirectory(prefix="litchi-0957-v2-custody-") as directory:
        root = Path(directory)
        absent = root / "tiny-output"
        try:
            require_report_directory_absent(absent)
            absent.mkdir()
            try:
                require_report_directory_absent(absent)
            except ContractError:
                outcomes["tiny_output_directory_precondition"] = "rejected"
            else:
                outcomes["tiny_output_directory_precondition"] = "unexpected_pass"
        finally:
            shutil.rmtree(absent, ignore_errors=True)
        packet, plan, child = _fixture(root / "positive")
        try:
            validate_child(packet, plan, child, str(packet))
            outcomes["positive_terminal_packet"] = "passed"
        except Exception as exc:  # pragma: no cover - receipt retains detail
            outcomes["positive_terminal_packet"] = f"failed: {exc}"

        mutations: dict[str, tuple[str, Callable[[dict[str, Any]], None]]] = {
            "argv_binding": ("terminal", lambda value: value["argv"].__setitem__(3, "/wrong/binary")),
            "cwd_binding": ("terminal", lambda value: value.__setitem__("cwd", "/wrong/cwd")),
            "deadline_binding": ("terminal", lambda value: value.__setitem__("deadline_seconds", 899)),
            "output_directory_binding": ("start", lambda value: value.__setitem__("output_directory", "/wrong/output")),
            "stdout_artifact_hash": ("terminal", lambda value: value["stdout"].__setitem__("sha256", "0" * 64)),
            "validation_terminal_hash": ("validation", lambda value: value.__setitem__("terminal_sha256", "0" * 64)),
            "process_group_reap": ("terminal", lambda value: value.__setitem__("process_group_reaped", False)),
            "report_artifact_path": ("terminal", lambda value: value["report"].__setitem__("path", "reports/other/report.json")),
        }
        for name, (record, update) in mutations.items():
            case_root = root / name
            packet, plan, child = _fixture(case_root)
            if record == "terminal":
                _mutate_terminal(packet, update)
            elif record == "start":
                _mutate_start(packet, update)
            else:
                _mutate_validation(packet, update)
            try:
                validate_child(packet, plan, child, str(packet))
            except ContractError:
                outcomes[name] = "rejected"
            except Exception as exc:  # pragma: no cover - preserve unexpected tool failures
                outcomes[name] = f"unexpected_error: {exc}"
            else:
                outcomes[name] = "unexpected_pass"

    result = {
        "schema": "litchi.performance.0957.custody-controls.v2",
        "status": "passed" if outcomes.get("positive_terminal_packet") == "passed" and all(
            value == "rejected" for key, value in outcomes.items() if key != "positive_terminal_packet"
        ) else "failed",
        "validator": "reduce_0957_v2.validate_child",
        "outcomes": outcomes,
        "synthetic_only": True,
        "workload_started": False,
        "scope": "expected argv/cwd/environment/output/report bindings, terminal deadline/reap, and sealed artifact hashes",
    }
    _write_json(OUT, result)
    print(json.dumps(result, sort_keys=True))
    return 0 if result["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
