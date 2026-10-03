#!/usr/bin/env python3
"""Replay report semantics and retained artifact custody without profiling."""
import argparse
import hashlib
import json
from pathlib import Path
import sys

P = Path(__file__).resolve().parent
sys.path.insert(0, str(P.parent / "change-0885"))
import verify

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--output", type=Path, required=True)
args = parser.parse_args()

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def read(path):
    return json.loads(path.read_text())

review = read(P / "review.json")
if review["status"] != "pass":
    raise ValueError("independent review did not pass")
for name, expected in review["summary_file_hashes"].items():
    if sha(P / name) != expected:
        raise ValueError("review binding mismatch: " + name)
plan = read(P / "plan.json")
if sha(Path(plan["binary"]["path"])) != plan["binary"]["sha256"]:
    raise ValueError("captured binary identity mismatch")
children = []
for prefix, count, samples, warmup in (("child", 3, 2000, 10), ("callgrind", 2, 10, 0)):
    for i in range(count):
        folder = P / f"{prefix}-{i:02}"
        command = read(folder / "command.json")
        if command["exit_code"] != 0:
            raise ValueError("failed profile child")
        for name, identity in command["artifacts"].items():
            path = folder / name
            if path.stat().st_size != identity["bytes"] or sha(path) != identity["sha256"]:
                raise ValueError("capture artifact mismatch: " + str(path))
        verify._validate_report(read(folder / "report.json"), read(folder / "producer.json"),
                                "native", {"samples": samples, "warmup": warmup})
        if prefix == "child":
            if command["plan_sha256"] != sha(P / "plan.json"):
                raise ValueError("native profile plan mismatch")
            if command["capture_helper_sha256"] != sha(P / "capture.py"):
                raise ValueError("native capture helper mismatch")
            for export in read(folder / "exports.json"):
                if export["exit_code"] != 0 or sha(folder / export["output"]) != export["sha256"]:
                    raise ValueError("profile export mismatch")
        else:
            annotation = read(folder / "annotate-command.json")
            if annotation["exit_code"] != 0 or sha(folder / "inclusive-canonical.txt") != annotation["output_sha256"]:
                raise ValueError("Callgrind annotation mismatch")
        children.append({"child": folder.name, "measured_samples": samples, "warmups": warmup})
result = {"status": "pass", "children": children, "profiled_measured_scans": 6000,
          "callgrind_measured_scans": 20, "production_changes": False,
          "scope": "Custody and semantic replay only; counter and IR reductions remain separate."}
with args.output.open("x") as f:
    json.dump(result, f, indent=2)
    f.write("\n")
print("All five retained profile children pass custody and semantic replay.")
