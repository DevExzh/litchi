#!/usr/bin/env python3
"""Replay 0895 custody without launching a workload or profiler command."""

from __future__ import annotations

import argparse
import datetime as _datetime
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any

from common import (  # type: ignore[import-not-found]
    CAPTURE_SCHEMA,
    ENVIRONMENT,
    EXPORT_SCHEMA,
    PACKET,
    PRIMARY_CASE,
    REDUCTION_SCHEMA,
    VERIFY_SCHEMA,
    descriptor,
    expected_binary_for_experiment,
    expected_corpora,
    load_base_plan,
    load_plan,
    read_json,
    render_print_args,
    require,
    sha,
    write_json,
    experiment,
)


def _time(value: str) -> _datetime.datetime:
    return _datetime.datetime.fromisoformat(value)


def _descriptor_matches(raw: Any, path: Path, label: str) -> None:
    require(isinstance(raw, dict), f"{label} descriptor is missing")
    actual = descriptor(path)
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        require(raw.get(key) == actual[key], f"{label} {key} mismatch")


def _capture_argv(plan: dict[str, Any], role: str, report: Path, producer: Path, profile_base: Path, binary: Path) -> list[str]:
    values = {
        "heaptrack": plan["tools"]["heaptrack"]["path"],
        "profile": str(profile_base),
        "binary": str(binary),
        "report": str(report),
        "producer": str(producer),
    }
    return [token.format(**values) for token in plan["record_argv_template"]]


def _capture_verify(plan_path: Path, plan: dict[str, Any], capture_root: Path) -> dict[str, Any]:
    receipt_path = capture_root / "capture.json"
    receipt = read_json(receipt_path)
    require(receipt.get("schema") == CAPTURE_SCHEMA and receipt.get("status") == "pass", "capture receipt is not passing")
    _descriptor_matches(receipt.get("plan"), plan_path, "capture plan")
    require(receipt.get("plan", {}).get("sha256") == sha(plan_path), "capture plan hash mismatch")
    rows = receipt.get("rows")
    expected_rows = [{"path": f"runs/{i:02d}-{role}/receipt.json", "row": plan["rows"][i]} for i, role in enumerate(("baseline", "candidate", "candidate", "baseline"))]
    require(rows == expected_rows and receipt.get("children") == 4 and receipt.get("measured_samples") == 80, "capture schedule/count mismatch")
    base_plan = load_base_plan()
    expected = expected_corpora()
    binaries = {role: expected_binary_for_experiment(base_plan, role) for role in ("baseline", "candidate")}
    previous_finished: _datetime.datetime | None = None
    for item in rows:
        row = item["row"]
        folder = capture_root / "runs" / f"{row['ordinal']:02d}-{row['role']}"
        child_receipt_path = capture_root / item["path"]
        child = read_json(child_receipt_path)
        require(child.get("schema") == CAPTURE_SCHEMA and child.get("status") == "pass" and child.get("exit_code") == 0, f"capture child is not passing: {row['ordinal']}")
        require(child.get("row") == row and child.get("plan_sha256") == sha(plan_path), f"capture child binding mismatch: {row['ordinal']}")
        binary = binaries[row["role"]]
        require(child.get("binary") == binary, f"capture binary binding mismatch: {row['ordinal']}")
        command_path = folder / "command.json"
        command = read_json(command_path)
        require(command.get("row") == row and command.get("exit_code") == 0, f"capture command status mismatch: {row['ordinal']}")
        require(command.get("environment") == ENVIRONMENT and command.get("cwd") == str(PACKET), f"capture environment/cwd mismatch: {row['ordinal']}")
        report = folder / "report.json"
        producer = folder / "producer.json"
        profile = folder / ("heaptrack" + plan["profile_suffix"])
        expected_argv = _capture_argv(plan, row["role"], report, producer, folder / "heaptrack", Path(binary["path"]))
        require(command.get("argv") == expected_argv and child.get("argv") == expected_argv, f"capture argv mismatch: {row['ordinal']}")
        _descriptor_matches(child.get("command"), command_path, f"capture command {row['ordinal']}")
        semantic = read_json(folder / "semantic.json")
        require(semantic.get("status") == "pass" and semantic.get("report_sha256") == sha(report) and semantic.get("producer_sha256") == sha(producer) and semantic.get("profile_sha256") == sha(profile), f"capture semantic mismatch: {row['ordinal']}")
        artifacts = child.get("artifacts")
        require(isinstance(artifacts, dict) and set(artifacts) == {"stdout.txt", "stderr.txt", "report.json", "producer.json", "heaptrack.zst"}, f"capture artifact set mismatch: {row['ordinal']}")
        for name in artifacts:
            _descriptor_matches(artifacts[name], folder / name, f"capture {row['ordinal']} {name}")
        experiment._validate_child(read_json(report), read_json(producer), PRIMARY_CASE, "native", 0, 20, expected[PRIMARY_CASE], binary)
        started = _time(command["started_utc"])
        finished = _time(command["finished_utc"])
        require(finished >= started, f"capture interval invalid: {row['ordinal']}")
        if previous_finished is not None:
            require(started >= previous_finished, f"capture children were not sequential: {row['ordinal']}")
        previous_finished = finished
    expected_dirs = {f"{i:02d}-{role}" for i, role in enumerate(("baseline", "candidate", "candidate", "baseline"))}
    require({p.name for p in (capture_root / "runs").iterdir() if p.is_dir()} == expected_dirs, "capture contains unexpected directories")
    return {"receipt": descriptor(receipt_path), "children": 4, "guard_children": 4}


def _export_verify(plan_path: Path, plan: dict[str, Any], capture_root: Path, export_root: Path) -> dict[str, Any]:
    receipt_path = export_root / "export.json"
    receipt = read_json(receipt_path)
    require(receipt.get("schema") == EXPORT_SCHEMA and receipt.get("status") == "pass", "export receipt is not passing")
    _descriptor_matches(receipt.get("plan"), plan_path, "export plan")
    rows = receipt.get("rows")
    expected_rows = [{"path": f"runs/{i:02d}-{role}/receipt.json", "row": plan["rows"][i]} for i, role in enumerate(("baseline", "candidate", "candidate", "baseline"))]
    require(rows == expected_rows and receipt.get("profiles") == 4, "export schedule/count mismatch")
    previous_finished: _datetime.datetime | None = None
    for item in rows:
        row = item["row"]
        folder = export_root / "runs" / f"{row['ordinal']:02d}-{row['role']}"
        child_path = export_root / item["path"]
        child = read_json(child_path)
        require(child.get("schema") == EXPORT_SCHEMA and child.get("status") == "pass" and child.get("exit_code") == 0, f"export child is not passing: {row['ordinal']}")
        capture_receipt = read_json(capture_root / f"runs/{row['ordinal']:02d}-{row['role']}/receipt.json")
        profile_raw = capture_receipt["artifacts"]["heaptrack.zst"]
        require(child.get("profile") == profile_raw, f"export profile binding mismatch: {row['ordinal']}")
        profile = Path(profile_raw["path"])
        stacks = folder / "stacks.txt"
        histogram = folder / "histogram.txt"
        expected_argv = [plan["tools"]["heaptrack_print"]["path"], *render_print_args(plan["print_argv_template"], profile=profile, stacks=stacks, histogram=histogram)]
        command_path = folder / "command.json"
        command = read_json(command_path)
        require(command.get("argv") == expected_argv and command.get("row") == row and command.get("exit_code") == 0, f"export argv/status mismatch: {row['ordinal']}")
        require(command.get("environment") == ENVIRONMENT and command.get("cwd") == str(PACKET), f"export environment/cwd mismatch: {row['ordinal']}")
        _descriptor_matches(child.get("command"), command_path, f"export command {row['ordinal']}")
        artifacts = child.get("artifacts")
        require(isinstance(artifacts, dict) and set(artifacts) == {"stdout.txt", "stderr.txt", "stacks.txt", "histogram.txt"}, f"export artifact set mismatch: {row['ordinal']}")
        for name in artifacts:
            _descriptor_matches(artifacts[name], folder / name, f"export {row['ordinal']} {name}")
        started = _time(command["started_utc"])
        finished = _time(command["finished_utc"])
        require(finished >= started, f"export interval invalid: {row['ordinal']}")
        if previous_finished is not None:
            require(started >= previous_finished, f"exports were not sequential: {row['ordinal']}")
        previous_finished = finished
    return {"receipt": descriptor(receipt_path), "profiles": 4}


def _load_reducer():
    path = PACKET / "reduce.py"
    spec = importlib.util.spec_from_file_location("litchi_selected_metadata_heaptrack_reduce_0895_verify", path)
    if spec is None or spec.loader is None:
        raise ImportError(path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def verify(plan_path: Path, capture_root: Path, export_root: Path, reduction_path: Path, output: Path) -> dict[str, Any]:
    plan = load_plan(plan_path)
    capture_root = capture_root.resolve()
    export_root = export_root.resolve()
    reduction_path = reduction_path.resolve()
    require(capture_root.is_dir() and export_root.is_dir(), "capture/export roots are missing")
    capture_result = _capture_verify(plan_path, plan, capture_root)
    export_result = _export_verify(plan_path, plan, capture_root, export_root)
    reduction = read_json(reduction_path)
    require(reduction.get("schema") == REDUCTION_SCHEMA and reduction.get("status") == "pass", "reduction is not passing")
    _descriptor_matches(reduction.get("plan"), plan_path, "reduction plan")
    _descriptor_matches(reduction.get("export"), export_root / "export.json", "reduction export")
    reducer = _load_reducer()
    exported = read_json(export_root / "export.json")["rows"]
    current_profiles = [reducer.reduce_one(plan, export_root, row) for row in exported]
    require(reduction.get("profiles") == current_profiles, "reduction is not an exact replay of current stacks")
    result = {
        "schema": VERIFY_SCHEMA,
        "status": "pass",
        "plan": descriptor(plan_path),
        "capture": capture_result,
        "export": export_result,
        "reduction": descriptor(reduction_path),
        "children": 4,
        "measured_samples": 80,
        "guard_contract": plan["guard_contract"],
        "workload_launched": False,
        "profiler_workload_launched": False,
        "performance_claim": None,
    }
    write_json(output, result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--capture-root", type=Path, required=True)
    parser.add_argument("--export-root", type=Path, required=True)
    parser.add_argument("--reduction", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(verify(args.plan.resolve(), args.capture_root.resolve(), args.export_root.resolve(), args.reduction.resolve(), args.output.resolve()), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
