#!/usr/bin/env python3
"""Prove 0423 replay rejects a modified pinned shared validator."""

from __future__ import annotations

import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Any


ROOT = Path(__file__).resolve().parent


def run() -> int:
    checks: list[dict[str, Any]] = []
    try:
        with tempfile.TemporaryDirectory(prefix="litchi-goal-0423-tool-binding-") as temporary:
            bundle = Path(temporary) / "bundle"
            shutil.copytree(ROOT, bundle, ignore=shutil.ignore_patterns("__pycache__"))
            command = [sys.executable, str(bundle / "portable-replay.py")]
            environment = os.environ.copy()
            environment["PYTHONDONTWRITEBYTECODE"] = "1"
            environment.pop("PYTHONPATH", None)

            valid = subprocess.run(command, cwd=temporary, env=environment, capture_output=True, text=True, check=False)
            if valid.returncode != 0:
                raise RuntimeError(f"unmodified bundle replay failed: {valid.stderr.strip()}")
            receipt = json.loads((bundle / "checks/portable-replay.json").read_text(encoding="utf-8"))
            if receipt.get("status") != "pass" or not receipt.get("temporary_export_removed"):
                raise RuntimeError("unmodified replay did not produce a passing isolated receipt")
            checks.append({"name": "unmodified_bundle_replay", "status": "pass"})

            validator = bundle / "replay-tools/perf_compare.py"
            validator.write_bytes(validator.read_bytes() + b"\n# deliberate binding mutation\n")
            invalid = subprocess.run(command, cwd=temporary, env=environment, capture_output=True, text=True, check=False)
            receipt = json.loads((bundle / "checks/portable-replay.json").read_text(encoding="utf-8"))
            if invalid.returncode == 0 or receipt.get("status") != "failed" or "pinned validator hash mismatch" not in receipt.get("error", ""):
                raise RuntimeError("modified pinned validator was not rejected before replay")
            checks.append({"name": "modified_pinned_validator_rejected", "status": "pass", "error": receipt["error"]})
    except (OSError, json.JSONDecodeError, RuntimeError, subprocess.SubprocessError) as error:
        record = {"change": 423, "status": "failed", "checks": checks, "error": str(error)}
        output = ROOT / "checks/portable-tool-binding.json"
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps(record, sort_keys=True))
        return 1

    record = {
        "change": 423,
        "status": "pass",
        "checks": checks,
        "temporary_export_removed": True,
    }
    output = ROOT / "checks/portable-tool-binding.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps(record, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
