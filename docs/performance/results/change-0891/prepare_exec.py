#!/usr/bin/env python3
"""Create the four managed executable copies for the 0891 capture.

One copy is made for each role/mode and reused by its scheduled fresh
children. The copy controls inode and pathname history while retaining the
source hash and mode. It does not establish a causal RSS result.
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


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    info = path.stat()
    if not stat.S_ISREG(info.st_mode) or path.is_symlink():
        raise ValueError(f"expected regular non-symlink executable: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": sha256(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
        "device": info.st_dev,
        "inode": info.st_ino,
    }


def write_once(path: Path, value: Any) -> None:
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
    if DESTINATION.exists() or args.output.exists():
        raise SystemExit("managed executable custody already exists; use a new packet/attempt")

    DESTINATION.mkdir()
    records: dict[str, dict[str, dict[str, Any]]] = {}
    for role in ROLES:
        records[role] = {}
        for mode in MODES:
            source = getattr(args, f"{role}_{mode}").resolve()
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
                raise RuntimeError(f"managed copy mode differs from source: {destination}")
            if (copy_record["device"], copy_record["inode"]) == (source_record["device"], source_record["inode"]):
                raise RuntimeError(f"managed copy aliases source inode: {destination}")
            records[role][mode] = {"source": source_record, "copy": copy_record}

    receipt = {
        "schema": "litchi.xlsx-x14ac-managed-executables-0891.v1",
        "status": "pass",
        "purpose": "content-identical role/mode copies reused by fresh timed children",
        "rss_claim": False,
        "records": records,
    }
    write_once(args.output, receipt)
    print(json.dumps(receipt, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
