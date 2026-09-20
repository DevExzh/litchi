#!/usr/bin/env python3
"""Stage the reference-metadata profile into a clean baseline checkout.

The working tree contains the candidate edits while the isolated checkout is
created at the recorded baseline commit.  This script copies only the source
closure declared by ``performance/run_profile.py`` and replaces its lockfile
with the retained gate lock.  It writes the candidate hash map consumed by
the final gate and performance verifiers; it does not run any Rust command.
"""

from __future__ import annotations

import argparse
import ast
import hashlib
import json
from pathlib import Path
import shutil
import subprocess


HERE = Path(__file__).resolve().parent
EVIDENCE = HERE.parent
ROOT = HERE.parents[4]
PREFIX = str(EVIDENCE.relative_to(ROOT))
GATE_LOCK_SHA256 = "58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3"


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def declaration(name: str) -> tuple[str, ...]:
    module = ast.parse((EVIDENCE / "performance/run_profile.py").read_text(encoding="utf-8"))
    for node in module.body:
        if not isinstance(node, ast.Assign):
            continue
        if not any(isinstance(target, ast.Name) and target.id == name for target in node.targets):
            continue
        value = ast.literal_eval(node.value)
        if not isinstance(value, tuple) or not all(isinstance(item, str) for item in value):
            raise RuntimeError(f"{name} must be a tuple of strings")
        return value
    raise RuntimeError(f"performance {name} declaration is missing")


def profile_sources() -> set[str]:
    paths = set(declaration("SOURCE_FILES"))
    for pattern in declaration("SOURCE_FILE_GLOBS"):
        paths.update(
            str(path.relative_to(ROOT))
            for path in ROOT.glob(pattern)
            if path.is_file()
        )
    for relative in paths:
        path = Path(relative)
        if path.is_absolute() or ".." in path.parts:
            raise RuntimeError(f"profile path escapes repository: {relative}")
    return paths


def batch_files(profile: set[str]) -> list[str]:
    """Select Rust files whose formatting belongs to this profile.

    Cargo's crate-wide format gate remains authoritative.  This narrower list
    gives the retained batch log an auditable list of reference metadata owners and
    their focused tests, while newly landed reference metadata modules are picked up
    by the profile globs automatically.
    """

    selected = []
    for relative in profile:
        if not relative.endswith(".rs"):
            continue
        if relative.startswith("crates/litchi-ods/src/codec/formula/evaluation"):
            selected.append(relative)
        elif relative.startswith("crates/litchi-ods/tests/ods_formula_reference_metadata_"):
            selected.append(relative)
    return sorted(set(selected))


def source_for(relative: str) -> Path:
    return HERE / "Cargo.lock" if relative == "Cargo.lock" else ROOT / relative


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("checkout", type=Path, help="clean checkout at the recorded baseline")
    args = parser.parse_args()
    checkout = args.checkout.resolve()
    if checkout == ROOT:
        raise RuntimeError("refusing to stage into the working tree")

    baseline = json.loads((EVIDENCE / "baseline.json").read_text(encoding="utf-8"))["commit"]
    observed_head = subprocess.check_output(
        ["git", "rev-parse", "HEAD"], cwd=checkout, text=True
    ).strip()
    if observed_head != baseline:
        raise RuntimeError(f"isolated checkout must be {baseline}, observed {observed_head}")

    profile = profile_sources()
    if "Cargo.lock" not in profile:
        raise RuntimeError("performance profile must declare Cargo.lock")
    gate_lock = HERE / "Cargo.lock"
    if digest(gate_lock) != GATE_LOCK_SHA256:
        raise RuntimeError("retained gate Cargo.lock does not match the byte-profile lock")

    sources = {relative: source_for(relative) for relative in sorted(profile)}
    missing = [relative for relative, source in sources.items() if not source.is_file()]
    if missing:
        raise RuntimeError("profile source files are missing:\n" + "\n".join(missing))

    for relative, source in sources.items():
        target = checkout / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)

    hashes = {relative: digest(source) for relative, source in sources.items()}
    for relative, expected in hashes.items():
        observed = digest(checkout / relative)
        if observed != expected:
            raise RuntimeError(f"isolated source hash mismatch: {relative}")

    rust = batch_files(profile)
    (HERE / "batch-files.json").write_text(
        json.dumps(rust, indent=2) + "\n", encoding="utf-8"
    )
    (HERE / "staged-profile-sources.json").write_text(
        json.dumps(hashes, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    (HERE / "freeze.json").write_text(
        json.dumps(
            {
                "status": "staged; final gates not run",
                "base_commit": baseline,
                "profile_source_count": len(hashes),
                "selected_files": hashes,
            },
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )
    print(
        json.dumps(
            {
                "base_commit": baseline,
                "staged_sources": len(hashes),
                "batch_files": len(rust),
                "gate_lock_sha256": GATE_LOCK_SHA256,
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
