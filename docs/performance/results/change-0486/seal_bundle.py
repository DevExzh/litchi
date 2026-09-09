#!/usr/bin/env python3
"""Seal or verify the exact 0486 diagnostic artifact inventory."""
import argparse
from pathlib import Path

from support import REPO, ROOT, meta, read, write

EXCLUDED = {"seal.json"}
REFERENCES = (
    "docs/performance/changes/0486-docx-replay-metadata-callers.md",
    "docs/performance/CRUD_COVERAGE.md",
    "docs/performance/GOAL_AUDIT.md",
    "docs/performance/REPORT.md",
    "docs/performance/results/change-0485/build-normal.json",
    "docs/performance/results/change-0484/profile_input_metadata.py",
    "docs/performance/results/change-0484/profile_routes.py",
)


def inventory():
    paths = sorted(ROOT.rglob("*"))
    if any(path.is_symlink() or path.suffix == ".pyc" for path in paths):
        raise ValueError("symlink or bytecode in evidence")
    return {path.relative_to(ROOT).as_posix(): meta(path) for path in paths
            if path.is_file() and path.relative_to(ROOT).as_posix() not in EXCLUDED}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--create", action="store_true")
    args = parser.parse_args()
    current = {"schema": "docx-caller-evidence-seal-v1", "files": inventory(),
               "references": {name: meta(REPO / name) for name in REFERENCES}}
    path = ROOT / "seal.json"
    if args.create:
        write(path, current)
    if read(path) != current:
        raise ValueError("sealed inventory/reference mismatch")
    print(f'PASS: {len(current["files"])} artifacts, {len(REFERENCES)} references')


if __name__ == "__main__":
    main()
