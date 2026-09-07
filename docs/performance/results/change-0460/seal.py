#!/usr/bin/env python3
"""Seal every regular bundle file except the root seal itself."""
import hashlib
from pathlib import Path

root = Path(__file__).resolve().parent
rows = []
for path in sorted(root.rglob("*")):
    assert not path.is_symlink(), f"refuse symlink: {path}"
    if path.is_file() and path != root / "SHA256SUMS":
        rows.append(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(root).as_posix()}\n")
(root / "SHA256SUMS").write_text("".join(rows))
print(f"sealed {len(rows)} files")
