#!/usr/bin/env python3
"""Export calls text and allocation-cost flamegraph stacks from captures."""

from __future__ import annotations

import argparse
import datetime as _datetime
import json
from pathlib import Path
import subprocess
from typing import Any

from common import (  # type: ignore[import-not-found]
    ENVIRONMENT,
    EXPORT_SCHEMA,
    PACKET,
    descriptor,
    load_plan,
    read_json,
    render_print_args,
    require,
    sha,
    utc_now,
    write_json,
)


def _time(value: str) -> _datetime.datetime:
    return _datetime.datetime.fromisoformat(value)


def _capture_rows(capture_root: Path) -> list[dict[str, Any]]:
    receipt = read_json(capture_root / "capture.json")
    require(receipt.get("status") == "pass", "capture receipt is not passing")
    rows = receipt.get("rows")
    require(type(rows) is list and len(rows) == 4, "capture row population is not four")
    return rows


def _receipt_artifact(root: Path, row: dict[str, Any], name: str) -> Path:
    receipt_path = root / row["path"]
    value = read_json(receipt_path)
    raw = value.get("artifacts", {}).get(name)
    require(isinstance(raw, dict), f"capture artifact is missing: {name}")
    path = Path(raw["path"])
    require(path.is_file() and not path.is_symlink(), f"capture artifact path is missing: {path}")
    actual = descriptor(path)
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        require(raw.get(key) == actual[key], f"capture artifact changed: {name}")
    return path.resolve()


def _argv(plan: dict[str, Any], *, profile: Path, stacks: Path, histogram: Path) -> list[str]:
    args = render_print_args(plan["print_argv_template"], profile=profile, stacks=stacks, histogram=histogram)
    return [plan["tools"]["heaptrack_print"]["path"], *args]


def _artifacts(folder: Path) -> dict[str, dict[str, Any]]:
    result = {}
    for name in ("stdout.txt", "stderr.txt", "stacks.txt", "histogram.txt"):
        path = folder / name
        if path.is_file():
            result[name] = descriptor(path)
    return result


def export(plan_path: Path, capture_root: Path, output_root: Path) -> dict[str, Any]:
    plan = load_plan(plan_path)
    capture_root = capture_root.resolve()
    output_root = output_root.resolve()
    require(capture_root.is_dir() and not capture_root.is_symlink(), f"capture root is missing: {capture_root}")
    require(not output_root.exists() and not output_root.is_symlink(), f"export output already exists or is symlinked: {output_root}")
    rows = _capture_rows(capture_root)
    output_root.mkdir(parents=True)
    previous_finished: _datetime.datetime | None = None
    exported = []
    plan_sha = sha(plan_path)
    for item in rows:
        row = item["row"]
        ordinal = row["ordinal"]
        profile = _receipt_artifact(capture_root, item, "heaptrack.zst")
        folder = output_root / "runs" / f"{ordinal:02d}-{row['role']}"
        folder.mkdir(parents=True, exist_ok=False)
        stacks = folder / "stacks.txt"
        histogram = folder / "histogram.txt"
        stdout_path = folder / "stdout.txt"
        stderr_path = folder / "stderr.txt"
        command_path = folder / "command.json"
        receipt_path = folder / "receipt.json"
        argv = _argv(plan, profile=profile, stacks=stacks, histogram=histogram)
        started = utc_now()
        command = {
            "schema": f"{EXPORT_SCHEMA}-command",
            "row": row,
            "argv": argv,
            "cwd": str(PACKET),
            "environment": dict(ENVIRONMENT),
            "tool": plan["tools"]["heaptrack_print"],
            "profile": descriptor(profile),
            "plan_sha256": plan_sha,
            "started_utc": started,
        }
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            process = subprocess.run(argv, cwd=PACKET, env=ENVIRONMENT.copy(), stdout=stdout, stderr=stderr, check=False)
        finished = utc_now()
        command.update({"finished_utc": finished, "exit_code": process.returncode})
        write_json(command_path, command)
        artifacts = _artifacts(folder)
        receipt = {
            "schema": EXPORT_SCHEMA,
            "status": "pass" if process.returncode == 0 else "fail",
            "row": row,
            "argv": argv,
            "cwd": str(PACKET),
            "environment": dict(ENVIRONMENT),
            "tool": plan["tools"]["heaptrack_print"],
            "profile": descriptor(profile),
            "plan_sha256": plan_sha,
            "started_utc": started,
            "finished_utc": finished,
            "exit_code": process.returncode,
            "command": descriptor(command_path),
            "artifacts": artifacts,
            "stdout_sha256": artifacts.get("stdout.txt", {}).get("sha256"),
            "stderr_sha256": artifacts.get("stderr.txt", {}).get("sha256"),
            "calls_ranked_text": artifacts.get("stdout.txt"),
            "allocation_flamegraph_stacks": artifacts.get("stacks.txt"),
            "allocation_histogram": artifacts.get("histogram.txt"),
            "workload_launched": False,
        }
        write_json(receipt_path, receipt)
        if process.returncode != 0:
            raise RuntimeError(f"heaptrack_print export {ordinal} exited {process.returncode}; fail-fast export stopped")
        require(stacks.is_file() and histogram.is_file(), f"heaptrack_print did not emit stacks/histogram: {ordinal}")
        require(set(artifacts) == {"stdout.txt", "stderr.txt", "stacks.txt", "histogram.txt"}, f"export artifact set changed: {ordinal}")
        if previous_finished is not None:
            require(_time(started) >= previous_finished, f"exports were not sequential: {ordinal}")
        previous_finished = _time(finished)
        exported.append({"path": str(receipt_path.relative_to(output_root)), "row": row})
    result = {
        "schema": EXPORT_SCHEMA,
        "status": "pass",
        "plan": descriptor(plan_path),
        "capture": descriptor(capture_root / "capture.json"),
        "rows": exported,
        "profiles": len(exported),
        "workload_launched": False,
        "performance_claim": None,
    }
    write_json(output_root / "export.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--capture-root", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(export(args.plan.resolve(), args.capture_root.resolve(), args.output_root.resolve()), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
