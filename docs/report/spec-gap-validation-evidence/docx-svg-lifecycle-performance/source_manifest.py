#!/usr/bin/env python3
"""Emit a small, deterministic source-and-fixture hash manifest."""

from __future__ import annotations

import argparse
import hashlib
import subprocess
from pathlib import Path


BASELINE = "892441d95db29da4351390716ef5c65b4c7c97de"
SOURCE_FILES = (
    "Cargo.toml",
    "crates/litchi-docx/src/error.rs",
    "crates/litchi-docx/src/source_backed.rs",
    "crates/litchi-docx/src/source_backed/svg_lifecycle.rs",
    "crates/litchi-docx/src/drawing/source.rs",
    "crates/litchi-docx/tests/drawing_svg_lifecycle.rs",
    "crates/litchi-opc/src/source_backed.rs",
    "crates/litchi-opc/src/error.rs",
    "crates/litchi-opc/src/phys_pkg.rs",
)
GENERATED_EVIDENCE_FILES = frozenset(
    {
        "report.md",
        "smoke-report.md",
        "full-report.md",
        "dominant-costs.md",
        "smoke-dominant-costs.md",
        "full-dominant-costs.md",
        "verification.json",
        "smoke-verification.json",
        "full-verification.json",
    }
)


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise SystemExit(message)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--evidence", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    root = args.root.resolve()
    evidence = args.evidence.resolve()
    snapshot = evidence / "baseline-source"
    head = subprocess.run(
        ["git", "-C", str(root), "rev-parse", "HEAD"],
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()
    require(head == BASELINE, f"baseline drift: expected {BASELINE}, got {head}")
    paths: list[tuple[Path, Path]] = []
    for relative in SOURCE_FILES:
        source = root / relative
        frozen = snapshot / relative
        require(source.is_file(), f"committed source input is missing: {source}")
        require(frozen.is_file(), f"retained source snapshot is missing: {frozen}")
        source_hash = digest(source)
        frozen_hash = digest(frozen)
        require(
            source_hash == frozen_hash,
            f"retained source snapshot differs from {relative}: {frozen_hash} != {source_hash}",
        )
        paths.append((evidence / "baseline-source" / relative, frozen))
    paths.extend(
        (path, path)
        for path in sorted(
            path
            for path in evidence.iterdir()
            if path.is_file() and path.name not in GENERATED_EVIDENCE_FILES
        )
    )
    paths.extend(
        (path, path)
        for path in sorted(evidence.joinpath("harness").rglob("*"))
    )
    paths.extend(
        (path, path)
        for path in sorted(evidence.joinpath("fixtures").rglob("*"))
    )
    lines = [
        "format=docx-svg-lifecycle-build-source-v1",
        f"git_head={head}",
        f"baseline_commit={BASELINE}",
        "source_snapshot=baseline-source",
    ]
    for display, path in paths:
        require(path.is_file(), f"manifest input is missing: {path}")
        lines.append(f"file={display.relative_to(evidence)}\t{digest(path)}")
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
