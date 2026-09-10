#!/usr/bin/env python3
"""Retain a successful coordinator build and its immutable source custody."""

import argparse
import os
from pathlib import Path
import shutil
import subprocess

from support import ROOT, REPO, TARGET_DIR, TEMP, environment, meta, now, read, sha, snapshot, write


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("role", choices=("normal", "allocator"))
    parser.add_argument("attempt")
    parser.add_argument("gate_label")
    args = parser.parse_args()
    assert args.attempt and all(c.isalnum() or c in "-_" for c in args.attempt)
    assert args.gate_label and all(c.isalnum() or c in "-_" for c in args.gate_label)
    gate_path = ROOT / "validation" / f"{args.gate_label}.json"
    gate = read(gate_path)
    assert gate["exit_code"] == 0 and gate["source_unchanged"] is True
    assert gate["source_before"] == gate["source_after"] == snapshot()
    assert gate["cwd"] == str(REPO) and gate["environment"] == environment()
    name = "litchi-perf-baseline" + ("-alloc" if args.role == "allocator" else "")
    command = ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml"]
    if args.role == "allocator":
        command += ["--features", "allocator-metrics"]
    command += ["--bin", name]
    assert gate["argv"] == command
    original = TARGET_DIR / "release" / name
    destination = TEMP / args.attempt / args.role / name
    assert not destination.exists() and not (ROOT / f"build-{args.role}.json").exists()
    assert original.is_file() and not original.is_symlink() and os.access(original, os.X_OK)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(original, destination)
    assert meta(original) == meta(destination)
    write(ROOT / f"build-{args.role}.json", {
        "schema": "docx-provider-lifecycle-build-v1", "version": 1,
        "role": args.role, "attempt": args.attempt, "copied_utc": now(),
        "command": command, "environment": environment(),
        "gate": {"path": str(gate_path), "sha256": sha(gate_path)},
        "binary": {"path": str(destination), "executable": True, **meta(destination)},
        "original_binary": {"path": str(original), "executable": True, **meta(original)},
        "source_before": gate["source_before"], "source_after": gate["source_after"],
        "source_unchanged": True,
        "git_revision": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip(),
    })
    print(destination)


if __name__ == "__main__":
    main()
