#!/usr/bin/env python3
"""Replay the bundle in an isolated export without original build artifacts."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent
PINNED_TOOLS = ROOT / "replay-tools"
FILES = ("perf_abba_summary.py", "perf_compare.py", "perf_resource_profile.py", "validate_perf_corpus_binding.py")
record = {"change": 421, "status": "running", "tools": [], "checks": []}
try:
    manifest = json.loads((PINNED_TOOLS / "manifest.json").read_text())
    expected = {entry["path"]: entry["sha256"] for entry in manifest["files"]}
    if len(manifest["files"]) != len(FILES) or set(expected) != set(FILES):
        raise ValueError("pinned validator manifest differs from the required module set")
    for name in FILES:
        actual = hashlib.sha256((PINNED_TOOLS / name).read_bytes()).hexdigest()
        if actual != expected[name]:
            raise ValueError(f"pinned validator hash mismatch: {name}")
except (OSError, ValueError, KeyError, TypeError) as error:
    record.update(status="failed", error=str(error), temporary_export_removed=True)
    (ROOT / "checks/portable-replay.json").write_text(json.dumps(record, indent=2, sort_keys=True) + "\n")
    raise SystemExit(str(error))
with tempfile.TemporaryDirectory(prefix="litchi-goal-0421-portable-") as temporary:
    export = Path(temporary)
    bundle = export / "docs/performance/results/change-0421"
    shutil.copytree(ROOT, bundle, ignore=shutil.ignore_patterns("__pycache__"))
    (export / "tools").mkdir()
    for name in FILES:
        source = PINNED_TOOLS / name
        shutil.copy2(source, export / "tools" / name)
        record["tools"].append({"path": f"tools/{name}", "sha256": hashlib.sha256(source.read_bytes()).hexdigest()})
    commands = (
        ("summarize.py", ["--replay"]),
        ("check-report-guards.py", []),
        ("verify.py", ["--repo-root", str(export), "--report", str(bundle / "checks/normal-identity/report.json"), "--catalog", str(bundle / "checks/normal-identity/catalog.json"), "--lane", "normal", "--samples", "1", "--warmups", "0", "--selector", "pptx_cross_copy_plain_lifecycle"]),
    )
    for script, arguments in commands:
        command = [sys.executable, str(bundle / script), *arguments]
        result = subprocess.run(command, cwd=export, env=os.environ | {"PYTHONDONTWRITEBYTECODE": "1", "PYTHONPATH": str(export)}, capture_output=True, text=True)
        record["checks"].append({"script": script, "args": [value.replace(temporary, "<EXPORT>") for value in arguments], "exit_code": result.returncode, "stdout": result.stdout.replace(temporary, "<EXPORT>"), "stderr": result.stderr.replace(temporary, "<EXPORT>")})
        if result.returncode:
            record["status"] = "failed"
            break
    else:
        record["status"] = "pass"
record["temporary_export_removed"] = not Path(temporary).exists()
(ROOT / "checks/portable-replay.json").write_text(json.dumps(record, indent=2, sort_keys=True) + "\n")
print(json.dumps({"status": record["status"], "checks": len(record["checks"]), "temporary_export_removed": record["temporary_export_removed"]}))
raise SystemExit(0 if record["status"] == "pass" else 1)
