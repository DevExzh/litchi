#!/usr/bin/env python3
"""Bind every production/helper input to the approved owner commit."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path


def git(root: Path, *args: str) -> str:
    return subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), *args],
        text=True,
    ).strip()


def blob_hash(root: Path, commit: str, relative: str) -> str:
    try:
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
    except subprocess.CalledProcessError as error:
        raise SystemExit(f"owner commit lacks transitive path-package input: {relative}") from error
    return hashlib.sha256(data).hexdigest()


def package_files(package: dict[str, object]) -> list[Path]:
    """Return every file below one local path package.

    Cargo metadata lists targets, not all files reached by ``mod``,
    ``include_*``, build scripts, tests, examples, or fixture readers.  The
    guard therefore captures the complete package tree and rejects generated
    VCS/build files rather than trusting a selected PPTX source list.
    """
    manifest = Path(str(package["manifest_path"])).resolve()
    package_root = manifest.parent
    ignored = {".git", "target"}
    return sorted(
        {
            path.resolve()
            for path in package_root.rglob("*")
            if path.is_file()
            and not ignored.intersection(path.relative_to(package_root).parts)
        },
        key=str,
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--commit", required=True)
    parser.add_argument("--helper", required=True)
    parser.add_argument("--helper-sha256", required=True)
    parser.add_argument("--path", action="append", required=True)
    parser.add_argument("--metadata", type=Path)
    parser.add_argument("--manifest-path", type=Path)
    parser.add_argument("--exclude-package", action="append", default=[])
    parser.add_argument(
        "--owner-extra",
        action="append",
        default=[],
        help="workspace/build input pinned to the approved owner commit",
    )
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

    paths = {Path(value) for value in args.path + [args.helper]}
    paths.update(Path(value) for value in args.owner_extra)
    if args.metadata is not None and args.manifest_path is not None:
        raise SystemExit("pass either --metadata or --manifest-path, not both")
    if args.metadata is not None:
        metadata = json.loads(args.metadata.read_text())
    elif args.manifest_path is not None:
        metadata = json.loads(
            subprocess.check_output(
                [
                    "cargo",
                    "metadata",
                    "--format-version=1",
                    "--locked",
                    "--offline",
                    "--manifest-path",
                    str(args.manifest_path.resolve()),
                ]
            )
        )
    else:
        metadata = None
    if metadata is not None:
        excluded = set(args.exclude_package)
        for package in metadata["packages"]:
            if package.get("source") is None and package["name"] not in excluded:
                paths.update(package_files(package))
    paths = sorted(paths, key=str)
    changed: list[str] = []
    for path in paths:
        local = path if path.is_absolute() else root / path
        relative = local.resolve().relative_to(root).as_posix()
        tracked = subprocess.run(
            ["git", "-C", str(root), "ls-files", "--error-unmatch", "--", relative],
            check=False,
            capture_output=True,
            text=True,
        )
        if tracked.returncode != 0:
            raise SystemExit(f"approved input is not tracked: {relative}")
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
        f"checked_paths={len(paths)}\n"
        f"transitive_path_packages={sum(1 for package in (metadata or {}).get('packages', []) if package.get('source') is None and package['name'] not in set(args.exclude_package))}"
    )


if __name__ == "__main__":
    main()
