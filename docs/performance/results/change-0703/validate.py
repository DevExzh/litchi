#!/usr/bin/env python3
"""Run existing evidence gates for a production-unchanged diagnostic packet."""
import hashlib
import json
import subprocess
import time
from pathlib import Path
P = Path(__file__).resolve().parent
ROOT = P.parents[3]

def main():
    prior = json.loads((P.parent / "change-0697/validation.json").read_text())
    rows = []
    (P / "validation").mkdir(exist_ok=True)
    for old in prior:
        name = old["name"]
        if name.startswith("probe-"):
            continue
        command = old["command"]
        path = P / "validation" / (name + ".log")
        start = time.monotonic()
        with path.open("w") as out:
            result = subprocess.run(command, cwd=ROOT, stdout=out, stderr=subprocess.STDOUT)
        rows.append(dict(name=name, command=command, exit_code=result.returncode,
                         seconds=time.monotonic()-start,
                         log_sha256=hashlib.sha256(path.read_bytes()).hexdigest()))
        (P / "validation.json").write_text(json.dumps(rows, indent=2) + "\n")
        print(name, result.returncode, flush=True)
        if result.returncode:
            raise SystemExit(path.read_text())

if __name__ == "__main__":
    main()
