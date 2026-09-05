#!/usr/bin/env python3
"""Check standalone bundle replay and rejection of a changed validator."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent
results = []
with tempfile.TemporaryDirectory(prefix="litchi-goal-0422-tool-binding-") as temporary:
    bundle = Path(temporary) / "bundle"
    shutil.copytree(ROOT, bundle, ignore=shutil.ignore_patterns("__pycache__"))
    command = [sys.executable, str(bundle / "portable-replay.py")]
    env = os.environ | {"PYTHONDONTWRITEBYTECODE": "1"}
    env.pop("PYTHONPATH", None)
    valid = subprocess.run(command, cwd=temporary, env=env, capture_output=True, text=True)
    assert valid.returncode == 0, valid.stderr
    receipt = json.loads((bundle / "checks/portable-replay.json").read_text())
    assert receipt["status"] == "pass" and len(receipt["checks"]) == 3
    results.append({"check": "standalone bundle without repository layout", "status": "pass"})
    validator = bundle / "replay-tools/perf_compare.py"
    validator.write_bytes(validator.read_bytes() + b"\n# deliberately altered test copy\n")
    invalid = subprocess.run(command, cwd=temporary, env=env, capture_output=True, text=True)
    assert invalid.returncode != 0, invalid.stdout
    receipt = json.loads((bundle / "checks/portable-replay.json").read_text())
    assert receipt["status"] == "failed" and "pinned validator hash mismatch" in receipt["error"]
    assert not receipt["checks"]
    results.append({"check": "modified pinned validator rejected before replay", "status": "pass"})
record = {"change": 422, "status": "pass", "checks": results, "temporary_export_removed": not Path(temporary).exists()}
(ROOT / "checks/portable-tool-binding.json").write_text(json.dumps(record, indent=2, sort_keys=True) + "\n")
print(json.dumps(record))
