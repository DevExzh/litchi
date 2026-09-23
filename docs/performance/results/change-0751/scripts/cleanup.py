#!/usr/bin/env python3
"""Record binary identities, then remove change 0751's build, scratch and
before-leg trees (adapted from change 0742's cleanup script).

Writes `cleanup.json` in the packet: for every staged binary its SHA-256 and
size, taken before removal; for every removed tree its path, size and file
count; and what was kept. The before leg's detached worktree is removed with
`git worktree remove --force`. The worktree and branch of this change are
kept.

Usage: cleanup.py [--dry-run]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path("/home/zhuhe/code/litchi-worktrees")
TREES = [
    ROOT / "targets" / "0751",
    ROOT / "targets" / "0751-before",
    ROOT / "targets" / "0751-before-fp",
    ROOT / "targets" / "0751-fp",
]
SCRATCH = ROOT / "scratch" / "0751"
BEFORE_WORKTREE = ROOT / "0751-before-src"
MAIN_CHECKOUT = Path("/home/zhuhe/code/litchi")
PACKET = Path(__file__).resolve().parent.parent
BINARIES = sorted((SCRATCH / "bin").glob("*")) if (SCRATCH / "bin").exists() else []


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def census(path: Path) -> dict:
    files = 0
    size = 0
    for directory, _dirs, names in os.walk(path):
        for name in names:
            try:
                size += (Path(directory) / name).lstat().st_size
            except OSError:
                continue
            files += 1
    return {"path": str(path), "files": files, "bytes": size}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    record: dict = {"binaries": [], "removed": [], "kept": []}
    for binary in BINARIES:
        record["binaries"].append(
            {"name": binary.name, "sha256": sha256(binary), "bytes": binary.stat().st_size}
        )
    for tree in TREES:
        if tree.exists():
            record["removed"].append(census(tree))
            if not args.dry_run:
                shutil.rmtree(tree)
    if SCRATCH.exists():
        entries = sorted(SCRATCH.iterdir())
        summary = census(SCRATCH)
        summary["entries"] = [entry.name for entry in entries]
        record["removed"].append(summary)
        if not args.dry_run:
            for entry in entries:
                if entry.is_dir() and not entry.is_symlink():
                    shutil.rmtree(entry)
                else:
                    entry.unlink()
    if BEFORE_WORKTREE.exists():
        entry = census(BEFORE_WORKTREE)
        entry["note"] = "detached worktree of 6d989cad63 for the before leg, removed with git worktree remove --force"
        record["removed"].append(entry)
        if not args.dry_run:
            subprocess.run(
                ["git", "-C", str(MAIN_CHECKOUT), "worktree", "remove", "--force", str(BEFORE_WORKTREE)],
                check=True,
            )
    record["kept"] = [
        str(ROOT / "0751-pptx-cross-copy-apply-digest-reuse"),
        "branch perf/0751-pptx-cross-copy-apply-digest-reuse",
        f"{SCRATCH} (empty directory)",
        "the packet's raw reports (gzipped), counters, analyses, attribution summaries and golden transcripts",
    ]
    record["not_retained"] = [
        "perf.data and perf script text of both frame-pointer profiles (53 MB); attribution/*.json summarize them",
        "the gate command outputs; gates.txt summarizes them",
    ]
    record["dry_run"] = args.dry_run
    text = json.dumps(record, indent=2) + "\n"
    if not args.dry_run:
        (PACKET / "cleanup.json").write_text(text, encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
