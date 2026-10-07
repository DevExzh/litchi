#!/usr/bin/env python3
"""Regenerate a bounded source BOM for a prepared 0960 source root."""
from __future__ import annotations

import argparse
import hashlib
import json
import stat
from datetime import datetime, timezone
from pathlib import Path


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--role", choices=("baseline", "candidate"), required=True)
    parser.add_argument("--source-commit", required=True)
    args = parser.parse_args()
    if not args.root.is_dir() or args.root.is_symlink():
        raise SystemExit(f"source root is not a real directory: {args.root}")
    if args.output.exists():
        raise SystemExit(f"refusing to overwrite: {args.output}")
    files = []
    total = 0
    symlinks = []
    for path in sorted(args.root.rglob("*")):
        rel = path.relative_to(args.root).as_posix()
        if path.is_symlink():
            symlinks.append(rel)
            continue
        if not path.is_file():
            continue
        st = path.stat()
        total += st.st_size
        files.append({
            "bytes": st.st_size,
            "executable": bool(st.st_mode & stat.S_IXUSR),
            "mode": stat.S_IMODE(st.st_mode),
            "path": rel,
            "sha256": digest(path),
        })
    result = {
        "schema": "litchi.performance.0960.source-manifest.v2",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "role": args.role,
        "source_commit": args.source_commit,
        "source_root": str(args.root),
        "excluded_prefixes": ["docs/performance/", "docs/report/"],
        "tracked_head_export": True,
        "candidate_status": "quality_frozen_candidate" if args.role == "candidate" else None,
        "file_count": len(files),
        "bytes": total,
        "symlinks": symlinks,
        "files": files,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
