#!/usr/bin/env python3
"""Static controls for the prospective 0959 runner, including retained-report positives.

Each control mutates only a prospective plan copy and invokes the runner's
``--dry-run`` path.  A zero exit code is a control failure.  No candidate,
binary workload, report, or build is started by this helper.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import runpy
import subprocess
import sys
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0959").resolve()
PLAN = ROOT / "plan-template-0959-v4.json"
RUNNER = ROOT / "tools/capture_0959_v7.py"
SCHEMA = "litchi.performance.0959.controls.v7"
POSITIVE_ALLOCATOR_REPORTS = (
    ("medium-one-percent", "medium-one-percent-w1-r0"),
    ("dense-sparse-one-percent", "dense-sparse-one-percent-w1-r0"),
    ("medium-one-edit", "medium-one-edit-w1-r0"),
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if path.is_symlink() or not path.is_file():
        raise ValueError(f"regular file required: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def write_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def utc_stamp() -> str:
    return dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")


def mutate(name: str, plan: dict[str, Any]) -> None:
    if name == "native-role-order":
        children = plan["lanes"]["native"]["children"]
        children[0]["role"], children[1]["role"] = children[1]["role"], children[0]["role"]
        children[0]["binary_role"], children[1]["binary_role"] = children[1]["binary_role"], children[0]["binary_role"]
    elif name == "allocator-count":
        plan["lanes"]["allocator"]["children"].pop()
    elif name == "native-cpu-mask":
        plan["lanes"]["native"]["children"][0]["cpu_mask"] = [12, 13, 14]
    elif name == "source-oracle":
        plan["cases"][0]["output_sha256"] = "0" * 64
    else:
        raise ValueError(f"unknown control {name}")


def positive_allocator_checks(plan_path: Path, root: Path) -> list[dict[str, Any]]:
    """Run every allocator report check against the three retained w1 reports.

    This is offline evidence only: it reads sealed 0957 reports and never
    invokes a candidate or baseline executable.
    """
    sys.dont_write_bytecode = True
    validator = runpy.run_path(str(root / "tools/capture_0959_v7.py"))
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    report_root = Path("/home/zhuhe/litchi-performance-0957/capture-allocation-v1/reports")
    rows: list[dict[str, Any]] = []
    for case_id, report_name in POSITIVE_ALLOCATOR_REPORTS:
        row = next(
            child for child in plan["lanes"]["allocator"]["children"]
            if child.get("case_id") == case_id
            and child.get("role") == "A"
            and child.get("block") == 0
            and child.get("position") == 0
        )
        report_path = report_root / report_name / "report.json"
        report = json.loads(report_path.read_text(encoding="utf-8"))
        binary_path = Path(report["binary_identity"]["path"])
        validator["validate_allocator"](report, row, plan, binary_path)
        rows.append({
            "case_id": case_id,
            "child_id": row["child_id"],
            "report": descriptor(report_path),
            "binary": descriptor(binary_path),
            "status": "passed",
        })
    return rows


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--receipt", type=Path)
    args = parser.parse_args()
    root = args.root.resolve()
    plan_path = root / "plan-template-0959-v4.json"
    runner = root / "tools/capture_0959_v7.py"
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    control_dir = root / "tmp" / f"controls-0959-v7-{utc_stamp()}-{os.getpid()}"
    control_dir.mkdir(parents=True, exist_ok=False)
    rows = []
    for name, lane in (
        ("native-role-order", "native"),
        ("allocator-count", "allocator"),
        ("native-cpu-mask", "native"),
        ("source-oracle", "native"),
    ):
        candidate = json.loads(json.dumps(plan))
        mutate(name, candidate)
        mutated = control_dir / f"{name}.json"
        write_exclusive(mutated, candidate)
        command = [sys.executable, str(runner), "--plan", str(mutated), "--lane", lane, "--dry-run"]
        result = subprocess.run(command, cwd=str(root), capture_output=True, text=True, check=False)
        passed = result.returncode != 0
        rows.append({
            "name": name,
            "lane": lane,
            "command": command,
            "return_code": result.returncode,
            "rejected": passed,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "plan": descriptor(mutated),
        })
        if not passed:
            raise SystemExit(f"control unexpectedly accepted: {name}")
    positive = positive_allocator_checks(plan_path, root)
    summary = {
        "schema": SCHEMA,
        "status": "passed",
        "plan": descriptor(plan_path),
        "runner": descriptor(runner),
        "controls": rows,
        "positive_allocator_reports": positive,
    }
    summary_path = control_dir / "summary.json"
    write_exclusive(summary_path, summary)
    receipt_path = (args.receipt or (root / "receipts" / "controls-0959-v7.json")).resolve()
    receipt = {
        "schema": "litchi.performance.0959.controls-receipt.v7",
        "status": "passed",
        "summary": descriptor(summary_path),
        "plan": descriptor(plan_path),
        "runner": descriptor(runner),
        "control_dir": str(control_dir),
        "controls": rows,
        "positive_allocator_reports": positive,
    }
    write_exclusive(receipt_path, receipt)
    print(control_dir)
    print(receipt_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
