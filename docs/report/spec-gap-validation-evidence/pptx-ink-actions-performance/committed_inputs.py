#!/usr/bin/env python3
"""Bind semantic and transitive production inputs to their explicit pins."""

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
        raise SystemExit(f"pinned commit lacks input: {relative}") from error
    return hashlib.sha256(data).hexdigest()


def blob_id(root: Path, commit: str, relative: str) -> str:
    try:
        return git(root, "rev-parse", f"{commit}:{relative}")
    except subprocess.CalledProcessError as error:
        raise SystemExit(f"pinned commit lacks input: {relative}") from error


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
    parser.add_argument("--semantic-owner-commit")
    parser.add_argument("--production-source-baseline-commit")
    # The legacy option remains parseable for old preflight callers, but the
    # production baseline is still mandatory and all capture paths use the
    # explicit dual-pin spelling.
    parser.add_argument("--commit", dest="legacy_commit")
    parser.add_argument("--helper", required=True)
    parser.add_argument("--helper-sha256", required=True)
    parser.add_argument("--path", action="append", required=True)
    parser.add_argument("--semantic-owner-extra", action="append", default=[])
    parser.add_argument("--semantic-owner-extra-sha256", action="append", default=[])
    parser.add_argument("--semantic-owner-extra-blob", action="append", default=[])
    parser.add_argument("--metadata", type=Path)
    parser.add_argument("--manifest-path", type=Path)
    parser.add_argument("--exclude-package", action="append", default=[])
    parser.add_argument(
        "--production-extra",
        "--owner-extra",
        dest="production_extra",
        action="append",
        default=[],
        help="workspace/build input pinned to --production-source-baseline-commit",
    )
    args = parser.parse_args()

    semantic_owner_commit = args.semantic_owner_commit or args.legacy_commit
    if semantic_owner_commit is None:
        raise SystemExit("--semantic-owner-commit is required")
    if len(semantic_owner_commit) != 40 or any(
        char not in "0123456789abcdef" for char in semantic_owner_commit
    ):
        raise SystemExit("semantic owner pin must be a full lowercase commit id")
    if args.legacy_commit and args.legacy_commit != semantic_owner_commit:
        raise SystemExit("legacy --commit disagrees with --semantic-owner-commit")
    if args.production_source_baseline_commit is None:
        raise SystemExit("--production-source-baseline-commit is required")
    production_source_baseline_commit = args.production_source_baseline_commit
    if len(production_source_baseline_commit) != 40 or any(
        char not in "0123456789abcdef" for char in production_source_baseline_commit
    ):
        raise SystemExit("production source baseline pin must be a full lowercase commit id")
    root = args.root.resolve()
    head = git(root, "rev-parse", "HEAD")
    for label, commit in (
        ("semantic owner pin", semantic_owner_commit),
        ("production baseline pin", production_source_baseline_commit),
    ):
        if subprocess.run(
            ["git", "-C", str(root), "merge-base", "--is-ancestor", commit, head],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        ).returncode != 0:
            raise SystemExit(f"{label} is not an ancestor of capture HEAD: {commit}")

    semantic_owner_extras = [Path(value) for value in args.semantic_owner_extra]
    if not (
        len(semantic_owner_extras)
        == len(args.semantic_owner_extra_sha256)
        == len(args.semantic_owner_extra_blob)
    ):
        raise SystemExit(
            "semantic owner input paths, SHA-256 values, and Git blobs must have equal counts"
        )
    semantic_paths = {Path(value) for value in args.path + [args.helper]}
    semantic_paths.update(semantic_owner_extras)
    production_paths = {Path(value) for value in args.production_extra}
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
                production_paths.update(package_files(package))

    def check_paths(paths: set[Path], commit: str, label: str) -> list[str]:
        changed: list[str] = []
        for path in sorted(paths, key=str):
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
            if hashlib.sha256(local.read_bytes()).hexdigest() != blob_hash(root, commit, relative):
                changed.append(relative)
        if changed:
            raise SystemExit(f"{label} inputs differ from pinned tree: " + ", ".join(changed))
        return changed

    check_paths(semantic_paths, semantic_owner_commit, "semantic owner")
    check_paths(production_paths, production_source_baseline_commit, "production baseline")

    for path, expected_sha, expected_blob in zip(
        semantic_owner_extras,
        args.semantic_owner_extra_sha256,
        args.semantic_owner_extra_blob,
        strict=True,
    ):
        local = path if path.is_absolute() else root / path
        relative = local.resolve().relative_to(root).as_posix()
        actual_sha = hashlib.sha256(local.read_bytes()).hexdigest()
        actual_blob = blob_id(root, semantic_owner_commit, relative)
        if actual_sha != expected_sha:
            raise SystemExit(f"semantic owner input SHA-256 differs: {relative}")
        if actual_blob != expected_blob:
            raise SystemExit(f"semantic owner input Git blob differs: {relative}")

    helper_digest = blob_hash(root, semantic_owner_commit, args.helper)
    if helper_digest != args.helper_sha256:
        raise SystemExit(
            f"helper SHA-256 differs: expected {args.helper_sha256}, committed {helper_digest}"
        )
    output = [
        f"semantic_owner_commit={semantic_owner_commit}",
        f"production_source_baseline_commit={production_source_baseline_commit}",
        f"source_commit={semantic_owner_commit}",
        f"capture_head={head}",
        f"helper={args.helper}",
        f"helper_sha256={helper_digest}",
        f"semantic_owner_extras={len(semantic_owner_extras)}",
        f"semantic_paths={len(semantic_paths)}",
        f"production_paths={len(production_paths)}",
        f"transitive_path_packages={sum(1 for package in (metadata or {}).get('packages', []) if package.get('source') is None and package['name'] not in set(args.exclude_package))}",
    ]
    for path, digest, blob in zip(
        semantic_owner_extras,
        args.semantic_owner_extra_sha256,
        args.semantic_owner_extra_blob,
        strict=True,
    ):
        relative = (path if path.is_absolute() else root / path).resolve().relative_to(root).as_posix()
        output.append(f"semantic_owner_extra={relative}\t{digest}\t{blob}")
    print(
        "\n".join(output)
    )


if __name__ == "__main__":
    main()
