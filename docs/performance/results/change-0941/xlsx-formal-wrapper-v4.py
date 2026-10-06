#!/usr/bin/env python3
"""Sequential outer wrapper for the authorized v4 96-child capture.

The default mode is a read-only schedule check.  ``--execute`` requires an
explicitly authorized copy of the v4 plan, records an outer start receipt for
each lane, runs native (72 children) and then allocator (24 children), and
records that lane's terminal receipt before inspecting any runner output.
The v4 runner itself records every child terminal before parsing its report.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import signal
import subprocess


ROOT = Path(__file__).resolve().parent
RUNNER = ROOT / "xlsx-capture-v4.py"
DEFAULT_PLAN = ROOT / "xlsx-plan-v4.json"
OUT = ROOT / "formal-wrapper-v4"
EXPECTED_ENV = {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "TZ": "UTC", "RUST_BACKTRACE": "0"}
EXPECTED_UNSET = ["RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "CARGO_HOME", "RUSTUP_TOOLCHAIN", "CARGO_INCREMENTAL"]


def now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def artifact(path: Path) -> dict:
    if not path.is_file():
        raise RuntimeError(f"artifact missing: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def read(path: Path) -> dict:
    value = json.loads(path.read_text())
    if not isinstance(value, dict):
        raise RuntimeError(f"JSON object required: {path}")
    return value


def write(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def kill_group(process: subprocess.Popen[bytes]) -> None:
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        process.wait()


def group_clear(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return True
    except PermissionError:
        return False
    return False


def validate_plan(plan: dict) -> None:
    if plan.get("schema") != "litchi.performance.0941.xlsx-controls-plan.v4":
        raise RuntimeError("v4 plan schema changed")
    if plan.get("status") not in {"prospective_only_pending_root_approval", "authorized"}:
        raise RuntimeError("unexpected v4 plan status")
    if plan.get("scope", {}).get("native", {}).get("total_children") != 72 or plan.get("scope", {}).get("allocator", {}).get("total_children") != 24:
        raise RuntimeError("frozen 96-child scope changed")
    if plan.get("runtime_environment", {}).get("PATH") != "/usr/bin:/bin" or plan.get("runtime_environment", {}).get("unset") != EXPECTED_UNSET:
        raise RuntimeError("runtime environment policy changed")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path, default=DEFAULT_PLAN)
    parser.add_argument("--execute", action="store_true")
    args = parser.parse_args()
    plan = read(args.plan)
    validate_plan(plan)
    commands = []
    for lane, expected in (("native", 72), ("allocator", 24)):
        commands.append({"lane": lane, "expected_children": expected, "command": ["python3", "-B", str(RUNNER), "--plan", str(args.plan), "--lane", lane, "--execute"]})
    if not args.execute:
        print(json.dumps({"status": plan["status"], "execute": False, "workload_started": False, "children": 96, "lanes": commands}, indent=2, sort_keys=True))
        return 0
    if plan.get("status") != "authorized":
        raise RuntimeError("formal wrapper requires an authorized plan copy")
    if OUT.exists():
        raise RuntimeError(f"refusing to overwrite {OUT}")
    OUT.mkdir(parents=True)
    receipts = []
    try:
        for lane_info in commands:
            lane = lane_info["lane"]
            stem = f"{len(receipts):02d}-{lane}"
            stdout = OUT / f"{stem}.stdout.log"
            stderr = OUT / f"{stem}.stderr.log"
            start = OUT / f"{stem}.start.json"
            terminal = OUT / f"{stem}.terminal.json"
            command = lane_info["command"]
            write(start, {"schema": "litchi.performance.0941.xlsx-formal-wrapper-start.v4", "lane": lane, "expected_children": lane_info["expected_children"], "command": command, "plan": artifact(args.plan), "runner": artifact(RUNNER), "environment": EXPECTED_ENV, "unset": EXPECTED_UNSET, "started_utc": now(), "status": "launching"})
            with stdout.open("wb") as out_stream, stderr.open("wb") as err_stream:
                process = subprocess.Popen(command, cwd=str(ROOT), stdout=out_stream, stderr=err_stream, start_new_session=True, env=EXPECTED_ENV)
                return_code = process.wait()
            clear = group_clear(process.pid)
            if not clear:
                kill_group(process)
                clear = group_clear(process.pid)
            terminal_record = {"schema": "litchi.performance.0941.xlsx-formal-wrapper-terminal.v4", "lane": lane, "expected_children": lane_info["expected_children"], "command": command, "plan": artifact(args.plan), "runner": artifact(RUNNER), "environment": EXPECTED_ENV, "unset": EXPECTED_UNSET, "pid": process.pid, "process_group": process.pid, "return_code": return_code, "process_group_clear": clear, "stdout": artifact(stdout), "stderr": artifact(stderr), "terminal_recorded_before_output_parse": True, "ended_utc": now()}
            write(terminal, terminal_record)
            receipts.append({"lane": lane, "expected_children": lane_info["expected_children"], "start": artifact(start), "terminal": artifact(terminal), "return_code": return_code})
            if return_code != 0 or not clear:
                write(OUT / "aborted.json", {"schema": "litchi.performance.0941.xlsx-formal-wrapper-aborted.v4", "receipts": receipts, "error": f"{lane} runner failed", "workload_started": True})
                raise RuntimeError(f"{lane} v4 runner failed")
        complete = {"schema": "litchi.performance.0941.xlsx-formal-wrapper-complete.v4", "lanes": receipts, "children": 96, "workload_started": True, "ended_utc": now()}
        write(OUT / "complete.json", complete)
        print(json.dumps(complete, indent=2, sort_keys=True))
        return 0
    except Exception:
        raise


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except RuntimeError as exc:
        raise SystemExit(f"formal wrapper refused: {exc}") from exc
