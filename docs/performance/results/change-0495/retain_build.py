#!/usr/bin/env python3
"""Retain one immutable 0495 before/after role build."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import shutil
import subprocess

from support import ROOT, REPO, TARGET_DIR, TEMP, environment, meta, now, read, sha, snapshot, write


PHASES = ("before", "after")
ROLES = ("normal", "allocator")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def _command(role: str) -> list[str]:
    name = "litchi-perf-baseline" + ("-alloc" if role == "allocator" else "")
    command = ["cargo", "build", "--release", "--locked", "--manifest-path",
               "tools/perf-baseline/Cargo.toml"]
    if role == "allocator":
        command += ["--features", "allocator-metrics"]
    return command + ["--bin", name]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("phase", choices=PHASES)
    parser.add_argument("role", choices=ROLES)
    parser.add_argument("attempt")
    parser.add_argument("gate_label")
    args = parser.parse_args()
    require(args.attempt and all(c.isalnum() or c in "-_" for c in args.attempt),
            "attempt must be a nonempty path-safe token")
    require(args.gate_label and all(c.isalnum() or c in "-_" for c in args.gate_label),
            "gate label must be a nonempty path-safe token")

    gate_path = ROOT / "validation" / f"{args.gate_label}.json"
    require(gate_path.is_file() and not gate_path.is_symlink(), f"missing gate receipt: {gate_path}")
    gate = read(gate_path)
    require(gate["exit_code"] == 0 and gate["timed_out"] is False
            and gate["termination"] is None and gate["source_unchanged"] is True,
            "build gate did not pass with unchanged source")
    current = snapshot()
    require(gate["source_before"] == gate["source_after"] == current,
            "build gate source custody differs from current source")
    require(gate["cwd"] == str(REPO) and gate["environment"] == environment(),
            "build gate cwd or environment differs")

    role = args.role
    command = _command(role)
    require(gate["argv"] == command, "build gate command differs from requested role")
    name = command[-1]
    original = TARGET_DIR / "release" / name
    destination = TEMP / "retained" / args.attempt / args.phase / role / name
    receipt_path = ROOT / f"build-{args.phase}-{role}.json"
    require(not destination.exists() and not receipt_path.exists(),
            "refusing to replace retained build artifacts")
    require(original.is_file() and not original.is_symlink() and os.access(original, os.X_OK),
            f"missing executable Cargo output: {original}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(original, destination)
    require(meta(original) == meta(destination), "retained executable identity differs from Cargo output")
    source_revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip()
    require(len(source_revision) == 40 and all(char in "0123456789abcdef" for char in source_revision),
            "current source revision is malformed")
    write(receipt_path, {
        "schema": "docx-provider-managed-build-v1", "version": 1,
        "phase": args.phase, "role": role, "attempt": args.attempt,
        "copied_utc": now(), "command": command, "environment": environment(),
        "gate": {"path": str(gate_path), "sha256": sha(gate_path)},
        "binary": {"path": str(destination), "executable": True, **meta(destination)},
        "original_binary": {"path": str(original), "executable": True, **meta(original)},
        "source_before": gate["source_before"], "source_after": gate["source_after"],
        "source_unchanged": True, "git_revision": source_revision,
        "retainer_sha256": sha(Path(__file__)),
    })
    print(destination)


if __name__ == "__main__":
    main()
