#!/usr/bin/env python3
"""Reject an altered summary even when the copied bundle is resealed."""
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

root = Path(__file__).resolve().parent
with tempfile.TemporaryDirectory(prefix="litchi-0458-negative-") as directory:
    target = Path(directory) / root.name
    shutil.copytree(root, target)
    summary = target / "summary.json"
    value = json.loads(summary.read_text())
    value["lanes"][0]["lifecycle_ns"]["p50"] += 1
    summary.write_text(json.dumps(value))
    subprocess.run([sys.executable, "-B", str(target / "seal.py")], check=True, capture_output=True)
    result = subprocess.run([sys.executable, "-B", str(target / "verify.py"), "--portable"], cwd=target, capture_output=True, text=True)
    output = json.loads(result.stdout)
    assert result.returncode == 1 and output.get("error") == "summary.json: recomputed JSON differs from retained file", output
assert not Path(directory).exists()
receipt = {"schema": "litchi-0458-negative-v1", "status": "pass", "mutation": "increment first lane p50 by one nanosecond and reseal copied bundle", "exit_code": result.returncode, "verifier_output": output, "temporary_directory_absent": True}
with (root / "negative-verification.json").open("x") as stream:
    json.dump(receipt, stream, indent=2)
    stream.write("\n")
print(json.dumps(receipt))
