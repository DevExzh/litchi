#!/usr/bin/env python3
"""Run the frozen native/read ABBA cycles, then the first allocator cycle.

Each command is source-bound and resumable.  The receipt inventory records the
lane and stage role for every command, so primary, read-control, and allocator
receipts cannot be mixed when a partially completed capture is resumed.
"""

from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys

sys.dont_write_bytecode = True

from custody import P, ROOT, census, sha  # noqa: E402


def read(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def write_progress(path: Path, value) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


plan = read(P / "plan.json")
freeze = read(P / "capture-freeze.json")
source_candidate = read(P / "source-candidate.json")


def check_frozen_state() -> None:
    for name, expected in freeze["files"].items():
        target = ROOT / name
        assert target.is_file() and sha(target) == expected, name
    assert census() == source_candidate


def command_for(kind: str, stage: str) -> list[str]:
    if kind == "primary-native":
        return [sys.executable, str(P / "pilot.py"), stage, "native"]
    if kind == "read-controls":
        return [sys.executable, str(P / "read-controls.py"), "capture", stage]
    if kind == "primary-allocator":
        return [sys.executable, str(P / "pilot.py"), stage, "allocator"]
    raise AssertionError(f"unknown capture kind: {kind}")


def jobs() -> list[dict[str, object]]:
    stages = [row["label"] for row in plan["stages"]]
    rows: list[dict[str, object]] = []
    for stage in stages:
        rows.append({"kind": "primary-native", "stage": stage})
        rows.append({"kind": "read-controls", "stage": stage})
    for stage in plan["allocator_stages"]:
        rows.append({"kind": "primary-allocator", "stage": stage})
    assert len(rows) == 20
    return rows


def main() -> int:
    capture_path = P / "capture.json"
    rows = read(capture_path) if capture_path.exists() else []
    assert isinstance(rows, list), "capture.json is not a list"
    expected = jobs()
    assert len(rows) <= len(expected), "capture receipt inventory is longer than the plan"

    for index, job in enumerate(expected):
        check_frozen_state()
        command = command_for(str(job["kind"]), str(job["stage"]))
        if index < len(rows):
            receipt = rows[index]
            assert isinstance(receipt, dict), f"capture receipt {index} is malformed"
            assert receipt.get("job") == job, f"capture receipt {index} job identity changed"
            assert receipt.get("command") == command, f"capture receipt {index} command changed"
            assert receipt.get("exit_code") == 0, f"capture receipt {index} was not successful"
            continue

        result = subprocess.run(command, cwd=ROOT, check=False)
        check_frozen_state()
        receipt = {
            "job": job,
            "command": command,
            "exit_code": result.returncode,
        }
        rows.append(receipt)
        write_progress(capture_path, rows)
        assert result.returncode == 0, f"capture command failed: {command!r}"

    assert len(rows) == len(expected)
    assert all(row.get("exit_code") == 0 for row in rows)
    print(f"captured {len(rows)} orchestrator commands", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
