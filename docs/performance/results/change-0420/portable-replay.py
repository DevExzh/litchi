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
REPO = ROOT.parents[3]
FILES = ("perf_abba_summary.py", "perf_compare.py", "perf_resource_profile.py", "validate_perf_corpus_binding.py")
record = {"change": 420, "status": "running", "tools": [], "checks": []}
with tempfile.TemporaryDirectory(prefix="litchi-goal-0420-portable-") as temporary:
    export = Path(temporary)
    bundle = export / "docs/performance/results/change-0420"
    shutil.copytree(ROOT, bundle, ignore=shutil.ignore_patterns("__pycache__"))
    (export / "tools").mkdir()
    for name in FILES:
        source = REPO / "tools" / name
        shutil.copy2(source, export / "tools" / name)
        record["tools"].append({"path": f"tools/{name}", "sha256": hashlib.sha256(source.read_bytes()).hexdigest()})
    for script in ("summarize.py",):
        command = [sys.executable, str(bundle / script), "--replay"]
        result = subprocess.run(command, cwd=export, env=os.environ | {"PYTHONDONTWRITEBYTECODE": "1", "PYTHONPATH": str(export)}, capture_output=True, text=True)
        record["checks"].append({"script": script, "args": ["--replay"], "exit_code": result.returncode, "stdout": result.stdout.replace(temporary, "<EXPORT>"), "stderr": result.stderr.replace(temporary, "<EXPORT>")})
        if result.returncode:
            record["status"] = "failed"
            break
    else:
        record["status"] = "pass"
record["temporary_export_removed"] = not Path(temporary).exists()
(ROOT / "checks/portable-replay.json").write_text(json.dumps(record, indent=2, sort_keys=True) + "\n")
print(json.dumps({"status": record["status"], "checks": len(record["checks"]), "temporary_export_removed": record["temporary_export_removed"]}))
raise SystemExit(0 if record["status"] == "pass" else 1)
