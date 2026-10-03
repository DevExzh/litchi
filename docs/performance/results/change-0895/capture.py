#!/usr/bin/env python3
"""Fail-fast four-child heaptrack capture for the retained 0894 binaries."""

from __future__ import annotations

import argparse
import datetime as _datetime
import json
from pathlib import Path
import subprocess
from typing import Any

from common import (  # type: ignore[import-not-found]
    CAPTURE_SCHEMA,
    ENVIRONMENT,
    PACKET,
    PRIMARY_CASE,
    descriptor,
    expected_binary_for_experiment,
    expected_corpora,
    load_base_plan,
    load_plan,
    read_json,
    require,
    same_descriptor,
    sha,
    utc_now,
    write_json,
    experiment,
)


def _time(value: str) -> _datetime.datetime:
    return _datetime.datetime.fromisoformat(value)


def _artifacts(folder: Path) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for name in ("stdout.txt", "stderr.txt", "report.json", "producer.json", "heaptrack.zst"):
        path = folder / name
        if path.is_file():
            result[name] = descriptor(path)
    return result


def _argv(plan: dict[str, Any], role: str, *, profile: Path, binary: Path, report: Path, producer: Path) -> list[str]:
    values = {
        "heaptrack": plan["tools"]["heaptrack"]["path"],
        "profile": str(profile),
        "binary": str(binary),
        "report": str(report),
        "producer": str(producer),
    }
    result = []
    for token in plan["record_argv_template"]:
        result.append(token.format(**values))
    return result


def _write_failure_receipt(
    path: Path,
    row: dict[str, Any],
    command: dict[str, Any],
    command_path: Path,
    folder: Path,
    error: str,
) -> None:
    receipt = {
        "schema": CAPTURE_SCHEMA,
        "status": "fail",
        "row": row,
        "command": descriptor(command_path),
        "argv": command["argv"],
        "cwd": command["cwd"],
        "environment": command["environment"],
        "plan_sha256": command["plan_sha256"],
        "started_utc": command["started_utc"],
        "finished_utc": command.get("finished_utc"),
        "exit_code": command.get("exit_code"),
        "artifacts": _artifacts(folder),
        "error": error,
        "workload_launched": True,
    }
    write_json(path, receipt)


def capture(plan_path: Path, output_root: Path) -> dict[str, Any]:
    plan = load_plan(plan_path)
    base_plan = load_base_plan()
    expected = expected_corpora()
    output_root = output_root.resolve()
    require(not output_root.exists() and not output_root.is_symlink(), f"capture output already exists or is symlinked: {output_root}")
    output_root.mkdir(parents=True)
    plan_sha = sha(plan_path)
    previous_finished: _datetime.datetime | None = None
    rows = []
    binaries = {role: expected_binary_for_experiment(base_plan, role) for role in ("baseline", "candidate")}

    for row in plan["rows"]:
        ordinal = row["ordinal"]
        folder = output_root / "runs" / f"{ordinal:02d}-{row['role']}"
        folder.mkdir(parents=True, exist_ok=False)
        report = folder / "report.json"
        producer = folder / "producer.json"
        profile_base = folder / "heaptrack"
        profile = folder / ("heaptrack" + plan["profile_suffix"])
        stdout_path = folder / "stdout.txt"
        stderr_path = folder / "stderr.txt"
        command_path = folder / "command.json"
        receipt_path = folder / "receipt.json"
        binary = binaries[row["role"]]
        argv = _argv(plan, row["role"], profile=profile_base, binary=Path(binary["path"]), report=report, producer=producer)
        started = utc_now()
        command = {
            "schema": f"{CAPTURE_SCHEMA}-command",
            "row": row,
            "argv": argv,
            "cwd": str(PACKET),
            "environment": dict(ENVIRONMENT),
            "binary": binary,
            "plan_sha256": plan_sha,
            "started_utc": started,
            "profile_basename": str(profile_base),
            "profile_suffix": plan["profile_suffix"],
        }
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            process = subprocess.run(argv, cwd=PACKET, env=ENVIRONMENT.copy(), stdout=stdout, stderr=stderr, check=False)
        finished = utc_now()
        command.update({"finished_utc": finished, "exit_code": process.returncode})
        write_json(command_path, command)
        if process.returncode != 0:
            _write_failure_receipt(receipt_path, row, command, command_path, folder, f"heaptrack child exited {process.returncode}")
            raise RuntimeError(f"capture child {ordinal} exited {process.returncode}; fail-fast capture stopped")
        try:
            require(profile.is_file() and not profile.is_symlink(), f"heaptrack profile is missing: {profile}")
            require(report.is_file(), f"child report is missing: {report}")
            require(producer.is_file(), f"primary producer sidecar is missing: {producer}")
            checked = experiment._validate_child(
                read_json(report),
                read_json(producer),
                PRIMARY_CASE,
                "native",
                row["warmup"],
                row["samples"],
                expected[PRIMARY_CASE],
                binary,
            )
            require(checked["result"].get("source", {}).get("producer_full_cell_scan", {}).get("guard_contract") == plan["guard_contract"], "0893 guard contract mismatch")
            if previous_finished is not None:
                require(_time(started) >= previous_finished, f"children were not sequential: {ordinal}")
            previous_finished = _time(finished)
        except Exception as error:
            _write_failure_receipt(receipt_path, row, command, command_path, folder, str(error))
            raise
        semantic_path = folder / "semantic.json"
        write_json(semantic_path, {
            "status": "pass",
            "report_sha256": sha(report),
            "producer_sha256": sha(producer),
            "profile_sha256": sha(profile),
            "guard_contract": plan["guard_contract"],
        })
        receipt = {
            "schema": CAPTURE_SCHEMA,
            "status": "pass",
            "row": row,
            "argv": argv,
            "cwd": str(PACKET),
            "environment": dict(ENVIRONMENT),
            "binary": binary,
            "plan_sha256": plan_sha,
            "started_utc": started,
            "finished_utc": finished,
            "exit_code": 0,
            "command": descriptor(command_path),
            "semantic": descriptor(semantic_path),
            "artifacts": _artifacts(folder),
            "workload_launched": True,
        }
        require(set(receipt["artifacts"]) == {"stdout.txt", "stderr.txt", "report.json", "producer.json", "heaptrack.zst"}, f"capture artifact set changed: {ordinal}")
        write_json(receipt_path, receipt)
        rows.append({"path": str(receipt_path.relative_to(output_root)), "row": row})

    result = {
        "schema": CAPTURE_SCHEMA,
        "status": "pass",
        "plan": descriptor(plan_path),
        "rows": rows,
        "children": len(rows),
        "measured_samples": len(rows) * plan["samples"],
        "fresh_child": True,
        "fail_fast": True,
        "workload_launched": True,
        "performance_claim": None,
    }
    write_json(output_root / "capture.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(capture(args.plan.resolve(), args.output_root.resolve()), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
