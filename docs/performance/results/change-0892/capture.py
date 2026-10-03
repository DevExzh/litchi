#!/usr/bin/env python3
"""Capture the frozen current-source native cycle profile.

This helper only launches the already retained 0891 native executable.  It
does not build, alter source, retry a child, or compare elapsed values.  Every
failed child leaves its command and partial outputs in its attempt directory.
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


def _write(path: Path, value: object) -> None:
    verify.write_once(path, value)


def capture(plan_path: Path, output_root: Path) -> dict[str, object]:
    if output_root.exists():
        raise FileExistsError(f"capture output already exists; preserve the prior attempt: {output_root}")
    plan = verify.read(plan_path)
    checked = verify.validate_plan(plan, "native")
    lane = checked["lane"]
    binary = checked["binary"]["identity"]
    output_root.mkdir(parents=True)
    records: list[dict[str, object]] = []
    plan_digest = verify.sha256(plan_path)
    for index in range(lane["children"]):
        folder = output_root / f"child-{index:02}"
        folder.mkdir()
        report = folder / "report.json"
        evidence = folder / "producer.json"
        argv = verify.native_argv(binary["path"], report, evidence, plan)
        command_start = {
            "schema": "litchi.xlsx-scan-cycle-profile-command-start-0892.v1",
            "mode": "native",
            "child_index": index,
            "argv": argv,
            "cwd": str(P),
            "environment": verify.CHILD_ENV,
            "profile": verify.PROFILE,
            "binary": binary,
            "plan_sha256": plan_digest,
            "started_utc": verify.utc_now(),
        }
        _write(folder / "command-start.json", command_start)
        code: int
        with (folder / "stdout.txt").open("x", encoding="utf-8") as stdout, (folder / "stderr.txt").open("x", encoding="utf-8") as stderr:
            code = subprocess.run(argv, cwd=P, env=verify.CHILD_ENV, stdout=stdout, stderr=stderr, check=False).returncode
        command = dict(command_start)
        command.update({"schema": "litchi.xlsx-scan-cycle-profile-command-0892.v1", "finished_utc": verify.utc_now(), "exit_code": code})
        command["artifacts"] = _artifacts(folder)
        _write(folder / "command.json", command)
        records.append({"path": f"child-{index:02}/command.json", "command_sha256": verify.sha256(folder / "command.json")})
        if code != 0:
            raise RuntimeError(f"native child {index} exited {code}; failed attempt retained at {folder}")
        verify.validate_child_semantics(folder, "native", lane, binary)
        print(f"child-{index:02} complete; semantic replay passed", flush=True)
    result = {
        "schema": verify.CAPTURE_SCHEMA,
        "status": "pass",
        "plan": str(plan_path.resolve()),
        "plan_sha256": plan_digest,
        "binary": binary,
        "records": records,
        "children": len(records),
        "profile": verify.PROFILE,
        "cpu": verify.CPU,
        "event": verify.EVENT,
        "frequency_hz": verify.FREQUENCY_HZ,
        "call_graph": verify.CALL_GRAPH,
        "warmup": verify.NATIVE_WARMUP,
        "samples": verify.NATIVE_SAMPLES,
        "schedule": "sequential",
        "no_retries": True,
        "performance_claim": False,
    }
    _write(output_root / "capture.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, default=P / "plan.json")
    parser.add_argument("--output-root", type=Path, default=P / "capture")
    args = parser.parse_args()
    result = capture(args.plan.resolve(), args.output_root.resolve())
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
