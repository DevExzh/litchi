#!/usr/bin/env python3
"""Bind every production/helper input to the approved owner commit."""

from __future__ import annotations

import argparse
import hashlib
import subprocess
from pathlib import Path


def git(root: Path, *args: str) -> str:
    return subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), *args],
        text=True,
    ).strip()


def blob_hash(root: Path, commit: str, relative: str) -> str:
    data = subprocess.check_output(
        [
            "git",
            "--no-replace-objects",
            "-C",
            str(root),
            "cat-file",
            "blob",
            f"{commit}:{relative}",
        ]
    )
    return hashlib.sha256(data).hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--commit", required=True)
    parser.add_argument("--helper", required=True)
    parser.add_argument("--helper-sha256", required=True)
    parser.add_argument("--path", action="append", required=True)
    args = parser.parse_args()

    if len(args.commit) != 40 or any(char not in "0123456789abcdef" for char in args.commit):
        raise SystemExit("owner source pin must be a full lowercase commit id")
    root = args.root.resolve()
    head = git(root, "rev-parse", "HEAD")
    subprocess.run(
        ["git", "-C", str(root), "merge-base", "--is-ancestor", args.commit, head],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    paths = sorted(set(args.path + [args.helper]))
    changed: list[str] = []
    for relative in paths:
        tracked = subprocess.run(
            ["git", "-C", str(root), "ls-files", "--error-unmatch", "--", relative],
            check=False,
            capture_output=True,
            text=True,
        )
        if tracked.returncode != 0:
            raise SystemExit(f"approved input is not tracked: {relative}")
        local = root / relative
        if not local.is_file():
            raise SystemExit(f"approved input is missing: {relative}")
        if hashlib.sha256(local.read_bytes()).hexdigest() != blob_hash(root, args.commit, relative):
            changed.append(relative)
    if changed:
        raise SystemExit("approved inputs differ from owner commit: " + ", ".join(changed))

    helper_digest = blob_hash(root, args.commit, args.helper)
    if helper_digest != args.helper_sha256:
        raise SystemExit(
            f"helper SHA-256 differs: expected {args.helper_sha256}, committed {helper_digest}"
        )
    print(
        f"owner_commit={args.commit}\n"
        f"head={head}\n"
        f"helper={args.helper}\n"
        f"helper_sha256={helper_digest}\n"
        f"checked_paths={len(paths)}"
    )


if __name__ == "__main__":
    main()
