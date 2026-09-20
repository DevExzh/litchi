#!/usr/bin/env python3
"""Write or verify the exact terminal 0720 evidence-packet census.

The manifest is intentionally excluded from its own census.  Every other
packet file, including receipts, raw stdout/stderr, source maps, oracle
outputs, reports, and gate logs, is covered by its relative path, byte count,
and SHA-256.  Frozen binaries may be removed during cleanup; their exact
identity must then be present in the packet's cleanup witness, while the
external binary itself is naturally outside this packet census.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
MANIFEST = HERE / "artifact-manifest.json"


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def census() -> dict[str, dict[str, int | str]]:
    files: dict[str, dict[str, int | str]] = {}
    for path in sorted(HERE.rglob("*")):
        if path == MANIFEST:
            continue
        if "__pycache__" in path.parts:
            raise SystemExit(f"refusing to seal Python cache: {path}")
        if path.is_symlink():
            raise SystemExit(f"refusing to seal symlink: {path}")
        if not path.is_file():
            continue
        relative = str(path.relative_to(HERE))
        files[relative] = {"bytes": path.stat().st_size, "sha256": digest(path)}
    return files


def write_manifest(files: dict[str, dict[str, int | str]]) -> None:
    MANIFEST.write_text(json.dumps({"files": files}, indent=2) + "\n")


def check_manifest(files: dict[str, dict[str, int | str]]) -> None:
    if not MANIFEST.is_file():
        raise SystemExit("artifact-manifest.json is missing")
    try:
        value: Any = json.loads(MANIFEST.read_text())
    except json.JSONDecodeError as error:
        raise SystemExit(f"invalid artifact manifest: {error}") from error
    if not isinstance(value, dict) or set(value) != {"files"} or value["files"] != files:
        raise SystemExit("artifact census or hashes changed")


def main() -> int:
    if sys.argv[1:] == ["--write"]:
        write_manifest(census())
    elif sys.argv[1:] == ["--check"]:
        check_manifest(census())
    else:
        raise SystemExit("usage: artifact-seal.py --write|--check")
    print(f"PASS: {len(census())} exact packet artifacts")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
