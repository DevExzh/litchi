#!/usr/bin/env python3
"""Small source and artifact custody helpers for the 0501 bundle."""

from __future__ import annotations

import hashlib
import json
import subprocess
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]


def sha_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def sha_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, indent=2) + "\n").encode()


def source_snapshot() -> dict[str, str]:
    """Hash source-like files visible to Git, including intentional untracked files."""

    raw = subprocess.check_output(
        ["git", "ls-files", "--cached", "--others", "--exclude-standard", "-z"],
        cwd=REPO,
    )
    names = set(raw.decode().split("\0"))
    names.add("Cargo.lock")
    result: dict[str, str] = {}
    for name in sorted(names):
        path = REPO / name
        if not path.is_file() or not name.endswith((".rs", ".toml", ".lock")):
            continue
        result[name] = sha_file(path)
    return result


def source_identity(snapshot: dict[str, str]) -> dict[str, Any]:
    raw = canonical_json(snapshot)
    return {"files": snapshot, "files_count": len(snapshot), "sha256": sha_bytes(raw)}


def artifact(path: Path, *, root: Path = HERE) -> dict[str, Any]:
    raw = path.read_bytes()
    return {
        "path": str(path.relative_to(root)),
        "bytes": len(raw),
        "sha256": sha_bytes(raw),
    }


def safe_relative(path: str, root: Path = HERE) -> Path:
    candidate = Path(path)
    if candidate.is_absolute() or ".." in candidate.parts:
        raise ValueError(f"unsafe bundle path: {path}")
    resolved = (root / candidate).resolve()
    if not resolved.is_relative_to(root.resolve()):
        raise ValueError(f"bundle path escapes root: {path}")
    return resolved


def now() -> str:
    import datetime

    return datetime.datetime.now(datetime.timezone.utc).isoformat()
