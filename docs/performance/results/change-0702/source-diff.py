#!/usr/bin/env python3
"""Freeze the exact candidate source diff against the recorded batch baseline."""

from __future__ import annotations

import hashlib
import json
import subprocess
from pathlib import Path


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
ALLOWED = (
    "crates/litchi-ooxml-common/src/mce/codec.rs",
)


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


base = json.loads((P / "baseline.json").read_text())
head = base["baseline_head"]
name_command = ["git", "diff", "--name-only", head, "--", *ALLOWED]
names = subprocess.check_output(name_command, cwd=ROOT, text=True).splitlines()
if set(names) != set(ALLOWED):
    raise SystemExit(f"candidate source census mismatch: {names!r}")

diff_command = ["git", "diff", "--no-ext-diff", "--binary", "--full-index", head, "--", *ALLOWED]
patch = subprocess.check_output(diff_command, cwd=ROOT)
if not patch:
    raise SystemExit("candidate source diff is empty")

patch_path = P / "source-diff.patch"
patch_path.write_bytes(patch)
receipt = {
    "baseline_head": head,
    "paths": list(ALLOWED),
    "name_command": name_command,
    "diff_command": diff_command,
    "patch": str(patch_path.relative_to(P)),
    "patch_bytes": len(patch),
    "patch_sha256": sha_bytes(patch),
}
(P / "source-diff.json").write_text(json.dumps(receipt, indent=2) + "\n")
print(f"recorded exact candidate diff ({len(patch)} bytes)")
