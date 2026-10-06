#!/usr/bin/env python3
"""Write a deterministic JSONL hash listing for an isolated measurement source."""
from __future__ import annotations
import argparse, hashlib, json, os
from pathlib import Path

def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("source", type=Path)
    ap.add_argument("listing", type=Path)
    ns = ap.parse_args()
    rows = []
    for path in sorted(ns.source.rglob("*")):
        rel = path.relative_to(ns.source).as_posix()
        if any(part in {".git", "target"} for part in path.parts):
            continue
        if path.is_symlink():
            row = {"kind": "symlink", "path": rel, "target": os.readlink(path)}
        elif path.is_file():
            digest = hashlib.sha256()
            size = 0
            with path.open("rb") as stream:
                for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                    digest.update(chunk)
                    size += len(chunk)
            row = {"kind": "file", "path": rel, "sha256": digest.hexdigest(), "size_bytes": size}
        elif path.is_dir():
            continue
        else:
            row = {"kind": "other", "path": rel}
        rows.append(row)
    ns.listing.parent.mkdir(parents=True, exist_ok=True)
    with ns.listing.open("w", encoding="utf-8") as stream:
        for row in rows:
            stream.write(json.dumps(row, sort_keys=True, separators=(",", ":")) + "\n")
    print(json.dumps({"file_count": len(rows), "listing": str(ns.listing), "listing_sha256": hashlib.sha256(ns.listing.read_bytes()).hexdigest(), "listing_bytes": ns.listing.stat().st_size}, sort_keys=True))

if __name__ == "__main__":
    main()
