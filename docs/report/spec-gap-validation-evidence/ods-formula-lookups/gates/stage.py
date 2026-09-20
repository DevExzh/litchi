#!/usr/bin/env python3
"""Stage the lookup profile into a clean baseline checkout.

This script only copies the source closure declared by the retained
performance profile.  It never runs Cargo and refuses to stage into the live
working tree.
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


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def safe_checkout_path(checkout: Path, relative: str) -> Path:
    if not isinstance(relative, str) or not relative or Path(relative).is_absolute() or ".." in Path(relative).parts:
        raise RuntimeError(f"profile path is not safely checkout-relative: {relative!r}")
    path = (checkout / relative).resolve()
    try:
        path.relative_to(checkout.resolve())
    except ValueError as error:
        raise RuntimeError(f"profile path escapes checkout: {relative}") from error
    return path


def baseline() -> dict:
    value = json.loads((EVIDENCE / "baseline.json").read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError("baseline.json is not an object")
    return value


def declaration(profile: Path, name: str) -> tuple[str, ...]:
    module = ast.parse(profile.read_text(encoding="utf-8"))
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


def profile_sources(profile: Path) -> set[str]:
    paths = set(declaration(profile, "SOURCE_FILES"))
    for pattern in declaration(profile, "SOURCE_FILE_GLOBS"):
        paths.update(str(path.relative_to(ROOT)) for path in ROOT.glob(pattern) if path.is_file())
    for relative in paths:
        path = Path(relative)
        if path.is_absolute() or ".." in path.parts:
            raise RuntimeError(f"profile path escapes repository: {relative}")
    return paths


def batch_files(profile: set[str]) -> list[str]:
    selected: list[str] = []
    for relative in profile:
        if not relative.endswith(".rs"):
            continue
        if relative.startswith("crates/litchi-ods/src/codec/formula/evaluation"):
            selected.append(relative)
        elif relative.startswith("crates/litchi-ods/tests/ods_formula_lookup"):
            selected.append(relative)
    return sorted(set(selected))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("checkout", type=Path, help="clean checkout at the recorded baseline")
    args = parser.parse_args()
    checkout = args.checkout.resolve()
    if checkout == ROOT:
        raise RuntimeError("refusing to stage into the working tree")
    config = baseline()
    commit = config.get("commit")
    if not isinstance(commit, str) or not commit:
        raise RuntimeError("baseline commit is missing")
    observed_head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=checkout, text=True).strip()
    if observed_head != commit:
        raise RuntimeError(f"isolated checkout must be {commit}, observed {observed_head}")
    profile = EVIDENCE / "performance" / "run_profile.py"
    if not profile.is_file():
        raise RuntimeError(f"performance profile is absent: {profile}")
    sources = {relative: (HERE / "Cargo.lock" if relative == "Cargo.lock" else ROOT / relative) for relative in sorted(profile_sources(profile))}
    if "Cargo.lock" not in sources:
        raise RuntimeError("performance profile must declare Cargo.lock")
    lock_hash = config.get("gate_lock_sha256")
    if not isinstance(lock_hash, str) or digest(HERE / "Cargo.lock") != lock_hash:
        raise RuntimeError("retained gate Cargo.lock does not match baseline")
    missing = [relative for relative, source in sources.items() if not source.is_file()]
    if missing:
        raise RuntimeError("profile source files are missing:\n" + "\n".join(missing))
    for relative, source in sources.items():
        target = safe_checkout_path(checkout, relative)
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)
    hashes = {relative: digest(source) for relative, source in sources.items()}
    for relative, expected in hashes.items():
        if digest(safe_checkout_path(checkout, relative)) != expected:
            raise RuntimeError(f"isolated source hash mismatch: {relative}")
    rust = batch_files(set(hashes))
    (HERE / "batch-files.json").write_text(json.dumps(rust, indent=2) + "\n", encoding="utf-8")
    (HERE / "staged-profile-sources.json").write_text(json.dumps(hashes, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    (HERE / "freeze.json").write_text(
        json.dumps(
            {
                "status": "staged; final gates not run",
                "base_commit": commit,
                "profile_source_count": len(hashes),
                "selected_files": hashes,
            },
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )
    print(json.dumps({"base_commit": commit, "staged_sources": len(hashes), "batch_files": len(rust), "gate_lock_sha256": lock_hash}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
