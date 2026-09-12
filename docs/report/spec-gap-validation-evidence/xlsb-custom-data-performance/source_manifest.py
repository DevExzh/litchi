#!/usr/bin/env python3
"""Record the committed source closure and candidate harness inputs."""

from __future__ import annotations

import argparse
import hashlib
import subprocess
from pathlib import Path


SOURCE_PATHS = (
    "Cargo.toml",
    "rust-toolchain.toml",
    ".cargo/config.toml",
    "crates/litchi-xlsb",
    "crates/litchi-opc",
    "crates/litchi-core",
    "crates/litchi-xldm",
    "crates/litchi-sheet",
    "crates/litchi-ooxml-common",
    "crates/litchi-drawingml",
    "crates/litchi-spreadsheet-drawing",
    "crates/soapberry-zip",
    "crates/xml-minifier",
    "crates/xml-minifier-macros",
    "docs/report/spec-gap-validation-evidence/xlsb-custom-data-design.md",
    "docs/report/spec-gap-validation-evidence/xlsb-custom-data-owner-final",
)


def digest(path: Path) -> tuple[str, int]:
    hasher = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
            size += len(chunk)
    return hasher.hexdigest(), size


def tracked(root: Path, path: str) -> list[str]:
    result = subprocess.run(
        ["git", "-C", str(root), "ls-files", "--", path],
        check=True,
        capture_output=True,
        text=True,
    )
    return [line for line in result.stdout.splitlines() if (root / line).is_file()]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--commit", required=True)
    parser.add_argument("--extra", action="append", type=Path, default=[])
    args = parser.parse_args()
    root = args.root.resolve()
    output = args.output.resolve()
    profile_root = Path(__file__).resolve().parent

    head = subprocess.check_output(
        ["git", "-C", str(root), "rev-parse", "HEAD"], text=True
    ).strip()
    if head != args.commit:
        raise SystemExit(f"source checkout is {head}, expected {args.commit}")

    source_files = sorted({item for path in SOURCE_PATHS for item in tracked(root, path)})
    if not source_files:
        raise SystemExit("source checkout contains no tracked source files")
    lines = [
        "format=xlsb-custom-data-performance-source-v1",
        f"source_commit={args.commit}",
        f"source_head={head}",
    ]
    for shown in source_files:
        digest_value, size = digest(root / shown)
        lines.append(f"source_file={shown}\t{size}\t{digest_value}")

    extras: list[Path] = []
    for extra in args.extra:
        path = extra.resolve()
        if not path.is_file():
            raise SystemExit(f"retained harness input is missing: {path}")
        extras.append(path)
    for path in sorted(set(extras), key=str):
        digest_value, size = digest(path)
        try:
            shown = Path("profile") / path.relative_to(profile_root)
        except ValueError:
            shown = Path("external") / path.name
        lines.append(f"extra_file={shown}\t{size}\t{digest_value}")

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("\n".join(lines) + "\n")
    print(f"source_files={len(source_files)} extras={len(set(extras))}")


if __name__ == "__main__":
    main()
