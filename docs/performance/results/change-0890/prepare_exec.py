#!/usr/bin/env python3
"""Create content-identical executable copies for the 0890 capture.

The 0888 copy control showed that pathname, inode, and executable-file cache
history can affect whole-process RSS.  The 0890 plan therefore launches each
role/mode through one separately managed copy, created once and reused across
the scheduled children.  This helper copies each source exactly once,
preserves its mode, verifies its bytes, and writes a create-only custody
receipt.  It never runs an executable and refuses existing destinations.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import stat
from typing import Any


PACKET = Path(__file__).resolve().parent
DESTINATION = PACKET / "executables"
ROLES = ("before", "after")
MODES = ("native", "allocator")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    info = path.stat()
    if not stat.S_ISREG(info.st_mode):
        raise ValueError(f"expected regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": sha(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
        "device": info.st_dev,
        "inode": info.st_ino,
    }


def write(path: Path, value: Any) -> None:
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    for role in ROLES:
        for mode in MODES:
            parser.add_argument(f"--{role}-{mode}", type=Path, required=True)
    parser.add_argument("--output", type=Path, default=PACKET / "execution-files.json")
    args = parser.parse_args()
    if DESTINATION.exists():
        raise SystemExit(f"managed executable directory already exists: {DESTINATION}")
    if args.output.exists():
        raise SystemExit(f"custody receipt already exists: {args.output}")

    DESTINATION.mkdir()
    records: dict[str, dict[str, dict[str, Any]]] = {}
    for role in ROLES:
        records[role] = {}
        for mode in MODES:
            source = (getattr(args, f"{role}_{mode}")).resolve()
            source_record = descriptor(source)
            if not source_record["executable"]:
                raise SystemExit(f"source is not executable: {source}")
            destination = DESTINATION / f"{role}-{mode}"
            shutil.copyfile(source, destination)
            destination.chmod(source_record["mode"])
            copy_record = descriptor(destination)
            if copy_record["sha256"] != source_record["sha256"] or copy_record["bytes"] != source_record["bytes"]:
                raise RuntimeError(f"managed copy differs from source: {source}")
            if copy_record["mode"] != source_record["mode"] or not copy_record["executable"]:
                raise RuntimeError(f"managed copy mode mismatch: {destination}")
            if (copy_record["device"], copy_record["inode"]) == (source_record["device"], source_record["inode"]):
                raise RuntimeError(f"managed copy aliases its source inode: {destination}")
            records[role][mode] = {"source": source_record, "copy": copy_record}

    receipt = {
        "schema": "litchi.mce-managed-executables-0890.v1",
        "status": "pass",
        "purpose": "all timed children launch content-identical fresh executable copies",
        "records": records,
    }
    write(args.output, receipt)
    print(json.dumps(receipt, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
