#!/usr/bin/env python3
"""Build and retain one after executable with its completed source gate."""
import argparse
import os
from pathlib import Path
import shutil
import subprocess
import sys

from support import ROOT, REPO, TEMP, ENV, environment, meta, now, read, write


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("role", choices=("normal", "allocator"))
    args = parser.parse_args()
    output = ROOT / f"build-{args.role}.json"
    destination = TEMP / "after1" / args.role / "docx_replayable_tail_append"
    if output.exists() or destination.exists():
        raise ValueError("after build already exists")
    command = ["cargo", "build", "--release", "--locked", "--manifest-path",
               "tools/perf-baseline/Cargo.toml", "--bin", "docx_replayable_tail_append"]
    if args.role == "allocator":
        command += ["--features", "allocator-metrics"]
    label = f"build-{args.role}-after1"
    subprocess.run([sys.executable, "-B", str(ROOT / "gate.py"), label, *command],
                   cwd=REPO, env=ENV, check=True)
    gate_path = ROOT / "validation" / f"{label}.json"
    gate = read(gate_path)
    if gate["exit_code"] != 0 or not gate["source_unchanged"]:
        raise ValueError("after build source changed or build failed")
    original = REPO / "tools/perf-baseline/target/release/docx_replayable_tail_append"
    destination.parent.mkdir(parents=True, exist_ok=True)
    with original.open("rb") as source, destination.open("xb") as target:
        shutil.copyfileobj(source, target)
    destination.chmod(original.stat().st_mode)
    binary = {"path": str(destination), **meta(destination), "executable": os.access(destination, os.X_OK)}
    if meta(original) != meta(destination) or not binary["executable"]:
        raise ValueError("executable copy differs")
    write(output, {"schema": "docx-replayable-tail-append-build-v1", "version": 1,
                   "role": args.role, "attempt": "after1", "command": command,
                   "binary": binary, "original_binary": {"path": str(original), **meta(original), "executable": True},
                   "gate": {"path": str(gate_path), "sha256": meta(gate_path)["sha256"]},
                   "source_before": gate["source_before"], "source_after": gate["source_after"],
                   "source_unchanged": True, "environment": environment(), "copied_utc": now()})
    print(output, binary["sha256"], flush=True)


if __name__ == "__main__":
    main()
