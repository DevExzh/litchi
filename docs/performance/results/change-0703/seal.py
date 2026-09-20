#!/usr/bin/env python3
"""Bind the final diagnostic artifacts, audit them, and rerun lightweight doc gates."""
import hashlib
import json
import subprocess
from pathlib import Path
P = Path(__file__).resolve().parent
ROOT = P.parents[3]
EXCLUDED = {"artifact-hashes.json", "audit-before-cleanup.log", "final-validation.json"}

def files():
    return [path for path in sorted(P.rglob("*")) if path.is_file()
            and path.relative_to(P).parts[0] != "seal"
            and str(path.relative_to(P)) not in EXCLUDED]

def main():
    manifest = {str(path.relative_to(P)): hashlib.sha256(path.read_bytes()).hexdigest() for path in files()}
    report = P / "../../0703-pptx-capture-projection-reuse-diagnostic.md"
    manifest["../../0703-pptx-capture-projection-reuse-diagnostic.md"] = hashlib.sha256(report.read_bytes()).hexdigest()
    (P / "artifact-hashes.json").write_text(json.dumps(manifest,indent=2)+"\n")
    commands = [("audit", ["python3", str(P / "audit.py")])]
    commands += [(r["name"],r["command"]) for r in json.loads((P/"validation.json").read_text())
                 if r["name"] in {"claims-structural","report","coverage","non-iwork"}]
    (P / "seal").mkdir(exist_ok=True)
    rows=[]
    for name,command in commands:
        log=P/"seal"/(name+".log")
        with log.open("w") as out:
            result=subprocess.run(command,cwd=ROOT,stdout=out,stderr=subprocess.STDOUT)
        rows.append(dict(name=name,command=command,exit_code=result.returncode,
                         log_sha256=hashlib.sha256(log.read_bytes()).hexdigest()))
        (P/"final-validation.json").write_text(json.dumps(rows,indent=2)+"\n")
        print(name,result.returncode,flush=True)
        if result.returncode:
            raise SystemExit(log.read_text())

if __name__ == "__main__":
    main()
