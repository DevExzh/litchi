"""Retain source snapshots and immutable logs for root-owned offline checks."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def artifact(path):
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def main():
    assert len(sys.argv) >= 2
    name = sys.argv[1]
    assert name in {"abort_validate.py", "reader_preflight.py", "analysis.py", "root_audit.py", "frame_audit.py", "stack_diagnostics.py", "validate.py"}
    attempts = P / "reader-attempts"
    attempts.mkdir(exist_ok=True)
    index = 0
    while (attempts / str(index)).exists():
        index += 1
    out = attempts / str(index)
    out.mkdir()
    source_dir = out / "sources"
    source_dir.mkdir()
    sources = []
    for source in sorted(P.glob("*.py")):
        target = source_dir / source.name
        shutil.copy2(source, target)
        sources.append({"original": artifact(source), "snapshot": artifact(target)})
    command = [sys.executable, "-B", str(P / name), *sys.argv[2:]]
    log = out / "console.log"
    started = time.time()
    with log.open("x") as stream:
        result = subprocess.run(command, cwd=ROOT, env=os.environ | {"PYTHONDONTWRITEBYTECODE": "1"},
                                stdout=stream, stderr=subprocess.STDOUT)
    receipt = {"schema": "litchi.performance.0828.reader-attempt.v1", "command": command,
               "started": started, "ended": time.time(), "exit_code": result.returncode,
               "sources": sources, "log": artifact(log)}
    (out / "receipt.json").write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
    print(log.read_text(), end="")
    print(f"0828 retained reader attempt {index}: exit {result.returncode}")
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
