#!/usr/bin/env python3
"""Build the candidate-only 0704 retention observer.

This lane has its own standalone Cargo workspace and never rebuilds or
relabels the frozen native 0704 binaries.  The receipt binds the complete
production trees used by the observer, every probe file, the build inputs,
the resulting binary, and a deterministic source hash.
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import time
from pathlib import Path


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
TARGET = ROOT.parent / "litchi-target-0704"
BINARY_DIR = ROOT.parent / "litchi-0704-bin"
BINARY = BINARY_DIR / "retention-observer"
MANIFEST = PACKET / "retention-probe" / "Cargo.toml"
LOCK = PACKET / "retention-probe" / "Cargo.lock"
CANDIDATE_CENSUS = PACKET / "source-census-candidate.json"
PRODUCTION_ROOTS = (
    ROOT / "crates" / "litchi-pptx",
    ROOT / "crates" / "litchi-ooxml-common",
    ROOT / "crates" / "litchi-opc",
)
BUILD_INPUTS = (
    ROOT / "Cargo.toml",
    ROOT / "Cargo.lock",
    ROOT / "crates" / "litchi-pptx" / "Cargo.toml",
    ROOT / "crates" / "litchi-ooxml-common" / "Cargo.toml",
    ROOT / "crates" / "litchi-opc" / "Cargo.toml",
    MANIFEST,
    LOCK,
    CANDIDATE_CENSUS,
)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def file_map(roots: tuple[Path, ...], relative_to: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    for root in roots:
        if not root.is_dir():
            raise AssertionError(f"missing census root: {root}")
        for path in sorted(root.rglob("*")):
            if path.is_file() and not path.is_symlink():
                result[str(path.relative_to(relative_to))] = sha(path)
    return dict(sorted(result.items()))


def source_hash(mapping: dict[str, str]) -> str:
    digest = hashlib.sha256()
    for name, value in sorted(mapping.items()):
        digest.update(name.encode())
        digest.update(b"\0")
        digest.update(bytes.fromhex(value))
    return digest.hexdigest()


def rust_map(mapping: dict[str, str]) -> dict[str, str]:
    return {name: digest for name, digest in mapping.items() if name.endswith(".rs")}


def other_map(mapping: dict[str, str]) -> dict[str, str]:
    return {name: digest for name, digest in mapping.items() if not name.endswith(".rs")}


def candidate_census_map() -> dict[str, str]:
    if not CANDIDATE_CENSUS.is_file():
        raise AssertionError(
            f"run source-census.py candidate before building the observer: {CANDIDATE_CENSUS}"
        )
    census = json.loads(CANDIDATE_CENSUS.read_text())
    return dict(sorted(census["source_sha256"].items()))


def probe_map() -> dict[str, str]:
    root = PACKET / "retention-probe"
    return {
        name: digest
        for name, digest in file_map((root,), PACKET).items()
        if not name.startswith("retention-probe/results/")
    }


def input_map() -> dict[str, str]:
    return {str(path.relative_to(ROOT)): sha(path) for path in BUILD_INPUTS}


def check_stable(
    production: dict[str, str], probe: dict[str, str], inputs: dict[str, str], rust: dict[str, str]
) -> None:
    current_production = file_map(PRODUCTION_ROOTS, ROOT)
    current_probe = probe_map()
    current_inputs = input_map()
    if current_production != production:
        raise AssertionError("production source census changed during retention build")
    if current_probe != probe:
        raise AssertionError("retention probe census changed during retention build")
    if current_inputs != inputs:
        raise AssertionError("build input census changed during retention build")
    if rust_map(current_production) != rust:
        raise AssertionError("Rust production census changed during retention build")
    if candidate_census_map() != rust:
        raise AssertionError("Rust production census no longer matches source-census-candidate.json")


def main() -> None:
    if not LOCK.is_file():
        raise AssertionError(
            f"standalone lockfile is required before building the observer: {LOCK}"
        )
    production = file_map(PRODUCTION_ROOTS, ROOT)
    probe = probe_map()
    inputs = input_map()
    rust = rust_map(production)
    other = other_map(production)
    if candidate_census_map() != rust:
        raise AssertionError("Rust production census does not match source-census-candidate.json")
    command = [
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        str(MANIFEST),
        "--target-dir",
        str(TARGET),
        "--bin",
        "retention-observer",
        "-j",
        "2",
    ]
    BINARY_DIR.mkdir(exist_ok=True)
    log_path = PACKET / "build-retention.log"
    started = time.monotonic()
    environment = {**os.environ, "RUSTFLAGS": "-D warnings"}
    with log_path.open("w") as log:
        result = subprocess.run(
            command,
            cwd=ROOT,
            env=environment,
            stdout=log,
            stderr=subprocess.STDOUT,
        )
    if result.returncode != 0:
        raise SystemExit(result.returncode)
    built = TARGET / "release" / "retention-observer"
    if not built.is_file():
        raise AssertionError(f"Cargo did not produce {built}")
    check_stable(production, probe, inputs, rust)
    shutil.copy2(built, BINARY)
    binary_sha256 = sha(BINARY)
    receipt = {
        "probe": "0704-retention-observer",
        "command": command,
        "cwd": str(ROOT),
        "rustflags": "-D warnings",
        "seconds": time.monotonic() - started,
        "target_dir": str(TARGET),
        "binary": str(BINARY),
        "binary_sha256": binary_sha256,
        "production_roots": [str(path.relative_to(ROOT)) for path in PRODUCTION_ROOTS],
        "production_file_count": len(production),
        "production_source_sha256": production,
        "production_source_hash": source_hash(production),
        "production_rust_file_count": len(rust),
        "production_rust_source_sha256": rust,
        "production_rust_source_hash": source_hash(rust),
        "production_other_file_count": len(other),
        "production_other_sha256": other,
        "production_other_hash": source_hash(other),
        "probe_file_count": len(probe),
        "probe_sha256": probe,
        "probe_source_hash": source_hash(probe),
        "build_inputs_sha256": inputs,
        "source_hash": source_hash({**production, **probe}),
    }
    (PACKET / "retention-build.json").write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps(receipt, indent=2))


if __name__ == "__main__":
    main()
