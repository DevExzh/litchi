#!/usr/bin/env python3
"""Capture a whole-command heaptrack diagnostic using a bound clean build."""
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "change-0418" / "scripts"))
from _common import source_identity, binary_identity, sha256_file, utc_now

parser = argparse.ArgumentParser()
parser.add_argument("role", choices=["control", "candidate"])
parser.add_argument("--root", type=Path, default=ROOT)
args = parser.parse_args()
ROOT = args.root.resolve()
build_path = ROOT / f"build-{args.role}.json"
build = json.loads(build_path.read_text())
assert build["status"] == "pass"
source = build["source_after"]
assert source_identity(Path(source["worktree"])) == source
binary = build["binaries"]["normal"]
assert binary_identity(Path(binary["path"]), label=binary["label"]) == binary
protocol = json.loads((ROOT / "protocol.json").read_text())
folder = ROOT / "heaptrack" / args.role
folder.mkdir(parents=True, exist_ok=False)
argv = ["taskset", "-c", str(protocol["cpu"]), "heaptrack", "--record-only", "-o", str(folder / "trace"), binary["path"], "--case", protocol["selector"], *protocol["common_flags"], "--samples", str(protocol["samples"]), "--warmup", str(protocol["warmups"]), "--json", str(folder / "report.json"), "--corpus-manifest", str(folder / "catalog.json")]
record = {"role": args.role, "started_utc": utc_now(), "scope": protocol["scope"], "protocol_sha256": sha256_file(ROOT / "protocol.json"), "build_sha256": sha256_file(build_path), "argv": argv, "status": "running"}
record_path = folder / "capture.json"
record_path.write_text(json.dumps(record, indent=2) + "\n")
with (folder / "stdout.log").open("wb") as out, (folder / "stderr.log").open("wb") as err:
    result = subprocess.run(argv, cwd=source["worktree"], env=os.environ | {"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": ""}, stdout=out, stderr=err)
record["exit_code"] = result.returncode
record["source_unchanged"] = source_identity(Path(source["worktree"])) == source
record["binary_unchanged"] = binary_identity(Path(binary["path"]), label=binary["label"]) == binary
record["files"] = [{"name": p.name, "bytes": p.stat().st_size, "sha256": sha256_file(p)} for p in sorted(folder.iterdir()) if p != record_path]
record["status"] = "pass" if result.returncode == 0 and record["source_unchanged"] and record["binary_unchanged"] else "failed"
record["finished_utc"] = utc_now()
record_path.write_text(json.dumps(record, indent=2) + "\n")
print(json.dumps({"role": args.role, "status": record["status"]}))
raise SystemExit(0 if record["status"] == "pass" else 1)
