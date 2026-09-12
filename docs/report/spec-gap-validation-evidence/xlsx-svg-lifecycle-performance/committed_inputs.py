#!/usr/bin/env python3
"""Fail closed when profile production inputs are not committed source bytes."""

from __future__ import annotations

import argparse
import hashlib
import subprocess
from pathlib import Path


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def run(root: Path, *args: str) -> str:
    return subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), *args], text=True
    ).strip()


def require_tracked(root: Path, relative: str) -> None:
    result = subprocess.run(
        ["git", "-C", str(root), "ls-files", "--error-unmatch", "--", relative],
        check=False,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise SystemExit(f"profile input is not tracked in the committed checkout: {relative}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--commit", required=True)
    parser.add_argument("--path", action="append", required=True)
    args = parser.parse_args()

    root = args.root.resolve()
    commit = args.commit
    if len(commit) != 40 or any(character not in "0123456789abcdef" for character in commit):
        raise SystemExit("profile source pin must be a full lowercase commit id")
    try:
        current = run(root, "rev-parse", "HEAD")
        subprocess.run(
            ["git", "-C", str(root), "merge-base", "--is-ancestor", commit, current],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except subprocess.CalledProcessError as error:
        raise SystemExit(
            f"approved XLSX SVG source pin is not an ancestor of HEAD: {commit} -> {current}"
        ) from error

    changed: list[str] = []
    for relative in sorted(set(args.path)):
        require_tracked(root, relative)
        local = root / relative
        if not local.is_file():
            raise SystemExit(f"profile input is missing: {relative}")
        try:
            committed = subprocess.check_output(
                ["git", "--no-replace-objects", "-C", str(root), "show", f"{commit}:{relative}"]
            )
        except subprocess.CalledProcessError as error:
            raise SystemExit(f"approved source pin does not contain profile input: {relative}") from error
        if sha256_bytes(local.read_bytes()) != sha256_bytes(committed):
            changed.append(relative)
    if changed:
        raise SystemExit(
            "profile inputs differ from the approved committed source pin: "
            + ", ".join(changed)
        )


if __name__ == "__main__":
    main()
