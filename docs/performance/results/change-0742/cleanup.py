#!/usr/bin/env python3
"""Record binary identities, then remove change 0742's build and scratch trees.

Writes `cleanup.json` beside this script: for every removed tree its path,
size and file count, and for every retained-identity binary its SHA-256,
taken before removal. The worktree, the branch and the shared base build
(`targets/base-009d515bef`) are kept.

This is the third review's cleanup (its commits `34255fea84` and
`ddefc2cfe8`, which built no measurement binaries). Earlier versions of this
script wrote what are now `cleanup-b2132486af.json` (first review) and
`cleanup-d2b2aa3d75.json` (second review).

Usage: cleanup.py [--dry-run]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
from pathlib import Path

ROOT = Path("/home/zhuhe/code/litchi-worktrees")
TREES = [
    ROOT / "targets" / "0742",
]
SCRATCH = ROOT / "scratch" / "0742"
PACKET = Path(__file__).resolve().parent
SUPERSEDED_RAW: list[Path] = []
BINARIES: list[Path] = []


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
        if binary.exists():
            record["binaries"].append(
                {"path": str(binary), "sha256": sha256(binary), "bytes": binary.stat().st_size}
            )
    for tree in TREES + SUPERSEDED_RAW:
        if tree.exists():
            entry = census(tree)
            if tree in SUPERSEDED_RAW:
                entry["path"] = str(tree.relative_to(PACKET))
                entry["note"] = "raw reports of a superseded matrix; summaries and receipts kept"
            record["removed"].append(entry)
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
    record["kept"] = [
        str(ROOT / "0742-pptx-owned-cross-copy-media-transfer"),
        "branch perf/0742-pptx-owned-cross-copy-media-transfer",
        str(ROOT / "targets" / "base-009d515bef") + " (shared base build, not owned by this change)",
        str(SCRATCH) + " (empty directory)",
    ]
    record["dry_run"] = args.dry_run
    out = Path(__file__).with_name("cleanup.json")
    out.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(record, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
