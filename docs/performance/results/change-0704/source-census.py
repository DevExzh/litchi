#!/usr/bin/env python3
"""Freeze the PPTX-only source census for the 0704 experiment.

The baseline is the exact 602-file map recorded by ``baseline.json``.  The
candidate may add or modify Rust under ``crates/litchi-pptx`` (including a
memo module and focused tests), but the shared OOXML and OPC owners must stay
byte-identical.  This receipt is deliberately independent of the build
driver, so a later audit can verify the source boundary without rebuilding.
"""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BASELINE = json.loads((P / "baseline.json").read_text())
OWNERS = ("litchi-pptx", "litchi-ooxml-common", "litchi-opc")
SHARED_PREFIXES = ("crates/litchi-ooxml-common/", "crates/litchi-opc/")


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def source_map() -> dict[str, str]:
    return {
        str(path.relative_to(ROOT)): sha(path)
        for owner in OWNERS
        for path in (ROOT / "crates" / owner).rglob("*.rs")
    }


def main() -> None:
    if len(sys.argv) != 2 or sys.argv[1] not in {"baseline", "candidate"}:
        raise SystemExit("usage: source-census.py baseline|candidate")
    phase = sys.argv[1]
    current = source_map()
    expected = BASELINE["source_sha256"]
    removed = sorted(set(expected) - set(current))
    added = sorted(set(current) - set(expected))
    changed = sorted(
        name for name in set(expected) & set(current) if expected[name] != current[name]
    )
    if phase == "baseline":
        if current != expected:
            raise AssertionError(
                f"baseline source census mismatch: removed={removed[:5]} "
                f"added={added[:5]} changed={changed[:5]}"
            )
    else:
        if removed:
            raise AssertionError(f"candidate removed baseline sources: {removed[:5]}")
        outside = [
            name for name in changed + added
            if not name.startswith("crates/litchi-pptx/")
        ]
        if outside:
            raise AssertionError(f"candidate source change leaves PPTX lane: {outside[:5]}")
        shared_changed = [name for name in changed if name.startswith(SHARED_PREFIXES)]
        if shared_changed:
            raise AssertionError(f"candidate changed shared source: {shared_changed[:5]}")

    names_command = ["git", "diff", "--name-only", BASELINE["baseline_head"], "--"]
    names = subprocess.check_output(names_command, cwd=ROOT, text=True).splitlines()
    receipt = {
        "phase": phase,
        "baseline_head": BASELINE["baseline_head"],
        "source_file_count": len(current),
        "source_sha256": current,
        "baseline_source_file_count": len(expected),
        "added_paths": added,
        "changed_paths": changed,
        "removed_paths": removed,
        "git_diff_name_only_command": names_command,
        "git_diff_name_only": names,
        "shared_prefixes": list(SHARED_PREFIXES),
        "candidate_boundary": "PPTX Rust sources only; shared OOXML and OPC sources unchanged",
    }
    (P / f"source-census-{phase}.json").write_text(json.dumps(receipt, indent=2) + "\n")
    print(phase, "source census", len(current), "files", "added", len(added), "changed", len(changed))


if __name__ == "__main__":
    main()
