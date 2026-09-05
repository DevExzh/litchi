#!/usr/bin/env python3
"""Build a clean revision with the established 0418 identity/flag helpers."""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "change-0418" / "scripts"))
from _common import source_identity, binary_identity, build_environment, host_identity, sha256_file, utc_now

parser = argparse.ArgumentParser()
parser.add_argument("role", choices=["control", "candidate"])
parser.add_argument("worktree", type=Path)
parser.add_argument("--root", type=Path, default=ROOT, help="fresh output directory containing protocol.json")
parser.add_argument("--binary-prefix", type=Path, default=Path("/tmp/litchi-goal-0421"))
args = parser.parse_args()
target = ROOT.parents[3] / "tools/perf-baseline/target"
ROOT = args.root.resolve()
output = ROOT / f"build-{args.role}.json"
if output.exists():
    raise SystemExit(f"refusing overwrite: {output}")
environment = build_environment(target)
argv = ["cargo", "+1.98.1", "build", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--release", "--features", "allocator-metrics", "--bin", "litchi-perf-baseline", "--bin", "litchi-perf-baseline-alloc"]
record = {"change": 421, "role": args.role, "started_utc": utc_now(), "protocol_sha256": sha256_file(ROOT / "protocol.json"), "source_before": source_identity(args.worktree), "environment": environment, "host": host_identity(), "argv": argv, "status": "running"}
output.write_text(json.dumps(record, indent=2) + "\n")
with (ROOT / f"build-{args.role}.log").open("wb") as log:
    result = subprocess.run(argv, cwd=args.worktree, env=os.environ | environment, stdout=log, stderr=subprocess.STDOUT)
record["exit_code"] = result.returncode
record["source_after"] = source_identity(args.worktree)
record["binaries"] = {}
if result.returncode == 0 and record["source_before"] == record["source_after"]:
    for mode, name in [("normal", "litchi-perf-baseline"), ("allocator", "litchi-perf-baseline-alloc")]:
        destination = Path(f"{args.binary_prefix}-{args.role}-{mode}")
        if destination.exists():
            raise SystemExit(f"refusing overwrite: {destination}")
        shutil.copy2(target / "release" / name, destination)
        record["binaries"][mode] = binary_identity(destination, label=f"{args.role}/{mode}")
    record["status"] = "pass"
else:
    record["status"] = "failed"
record["finished_utc"] = utc_now()
output.write_text(json.dumps(record, indent=2) + "\n")
print(json.dumps({"role": args.role, "status": record["status"]}))
raise SystemExit(0 if record["status"] == "pass" else 1)
