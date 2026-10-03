#!/usr/bin/env python3
"""Capture the separate runner-scoped Callgrind diagnostic.

Callgrind guest instructions are retained as an independent diagnostic.  This
driver never combines them with native sampled cycles and stops after the
first command, semantic, or zero-collection failure while preserving that
failed attempt.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import subprocess

import verify


P = Path(__file__).resolve().parent


def _artifacts(folder: Path) -> dict[str, dict[str, object]]:
    result: dict[str, dict[str, object]] = {}
    for path in sorted(folder.iterdir()):
        if path.is_file() and path.name != "command.json":
            result[path.name] = verify.descriptor(path)
    return result


def capture(plan_path: Path, output_root: Path) -> dict[str, object]:
    if output_root.exists():
        raise FileExistsError(f"Callgrind output already exists; preserve the prior attempt: {output_root}")
    plan = verify.read(plan_path)
    checked = verify.validate_plan(plan, "callgrind")
    lane = checked["lane"]
    binary = checked["binary"]["identity"]
    output_root.mkdir(parents=True)
    records: list[dict[str, object]] = []
    plan_digest = verify.sha256(plan_path)
    for index in range(lane["children"]):
        folder = output_root / f"callgrind-{index:02}"
        folder.mkdir()
        raw = folder / "callgrind.out"
        report = folder / "report.json"
        evidence = folder / "producer.json"
        argv = verify.callgrind_argv(binary["path"], raw, report, evidence, plan)
        command_start = {
            "schema": "litchi.xlsx-scan-callgrind-command-start-0892.v1",
            "mode": "callgrind",
            "child_index": index,
            "argv": argv,
            "cwd": str(P),
            "environment": verify.CHILD_ENV,
            "profile": verify.PROFILE,
            "binary": binary,
            "plan_sha256": plan_digest,
            "started_utc": verify.utc_now(),
        }
        verify.write_once(folder / "command-start.json", command_start)
        with (folder / "stdout.txt").open("x", encoding="utf-8") as stdout, (folder / "stderr.txt").open("x", encoding="utf-8") as stderr:
            code = subprocess.run(argv, cwd=P, env=verify.CHILD_ENV, stdout=stdout, stderr=stderr, check=False).returncode
        command = dict(command_start)
        command.update({"schema": "litchi.xlsx-scan-callgrind-command-0892.v1", "finished_utc": verify.utc_now(), "exit_code": code})
        command["artifacts"] = _artifacts(folder)
        verify.write_once(folder / "command.json", command)
        records.append({"path": f"callgrind-{index:02}/command.json", "command_sha256": verify.sha256(folder / "command.json")})
        if code != 0:
            raise RuntimeError(f"Callgrind child {index} exited {code}; failed attempt retained at {folder}")
        verify.validate_child_semantics(folder, "callgrind", lane, binary)
        total = verify.validate_callgrind_totals(raw)
        print(f"callgrind-{index:02} complete; {total} guest instructions", flush=True)
    result = {
        "schema": verify.CALLGRIND_CAPTURE_SCHEMA,
        "status": "pass",
        "plan": str(plan_path.resolve()),
        "plan_sha256": plan_digest,
        "binary": binary,
        "records": records,
        "children": len(records),
        "profile": verify.PROFILE,
        "cpu": verify.CPU,
        "toggle_function": plan.get("toggle_function", verify.TOGGLE_FUNCTION),
        "warmup": verify.CALLGRIND_WARMUP,
        "samples": verify.CALLGRIND_SAMPLES,
        "schedule": "sequential",
        "no_retries": True,
        "performance_claim": False,
    }
    verify.write_once(output_root / "capture.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, default=P / "callgrind-plan.json")
    parser.add_argument("--output-root", type=Path, default=P / "callgrind")
    args = parser.parse_args()
    result = capture(args.plan.resolve(), args.output_root.resolve())
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
