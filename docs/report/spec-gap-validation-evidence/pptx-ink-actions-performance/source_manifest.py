#!/usr/bin/env python3
"""Capture the complete Cargo source closure used by the isolated harness."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path


FORMAT = "pptx-ink-actions-profile-build-source-v1"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def shown(path: Path, root: Path) -> str:
    path = path.resolve()
    try:
        return path.relative_to(root).as_posix()
    except ValueError:
        return str(path)


def require_file(path: Path) -> Path:
    path = path.resolve()
    if not path.is_file():
        raise SystemExit(f"source manifest input is missing: {path}")
    return path


def committed_blob_sha256(root: Path, commit: str, relative: str) -> str:
    try:
        value = subprocess.check_output(
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
        raise SystemExit(f"cannot read committed source blob: {relative}") from error
    return hashlib.sha256(value).hexdigest()


def committed_blob_id(root: Path, commit: str, relative: str) -> str:
    try:
        return subprocess.check_output(
            [
                "git",
                "--no-replace-objects",
                "-C",
                str(root),
                "rev-parse",
                f"{commit}:{relative}",
            ],
            text=True,
        ).strip()
    except subprocess.CalledProcessError as error:
        raise SystemExit(f"cannot read semantic owner blob: {relative}") from error


def verify_local_inputs(paths: list[Path], root: Path, commit: str) -> None:
    relative = sorted(
        {
            path.resolve().relative_to(root.resolve()).as_posix()
            for path in paths
        }
    )
    if not relative:
        return
    current = subprocess.check_output(
        ["git", "-C", str(root), "rev-parse", "HEAD"], text=True
    ).strip()
    if current != commit:
        raise SystemExit(f"source HEAD changed during capture: {commit} -> {current}")
    for name in relative:
        path = root / name
        require_file(path)
        tracked = subprocess.run(
            ["git", "-C", str(root), "ls-files", "--error-unmatch", "--", name],
            capture_output=True,
            text=True,
            check=False,
        )
        if tracked.returncode != 0:
            raise SystemExit(f"local source input is not tracked: {name}")
        if sha256(path) != committed_blob_sha256(root, commit, name):
            raise SystemExit(f"local source input differs from committed tree: {name}")


def verify_production_inputs(paths: list[Path], root: Path, commit: str) -> None:
    """Pin every transitive production input to the frozen source baseline."""
    relative = sorted(
        {
            path.resolve().relative_to(root.resolve()).as_posix()
            for path in paths
        }
    )
    for name in relative:
        path = root / name
        require_file(path)
        if sha256(path) != committed_blob_sha256(root, commit, name):
            raise SystemExit(f"production baseline input differs: {name}")


def verify_semantic_owner_inputs(
    paths: list[Path],
    root: Path,
    commit: str,
    expected_sha256: list[str],
    expected_blobs: list[str],
) -> None:
    if not (len(paths) == len(expected_sha256) == len(expected_blobs)):
        raise SystemExit(
            "semantic owner input paths, SHA-256 values, and Git blobs must have equal counts"
        )
    for path, expected_sha, expected_blob in zip(
        paths, expected_sha256, expected_blobs, strict=True
    ):
        relative = path.resolve().relative_to(root.resolve()).as_posix()
        require_file(path)
        actual_sha = sha256(path)
        actual_blob = committed_blob_id(root, commit, relative)
        if actual_sha != expected_sha:
            raise SystemExit(f"semantic owner input SHA-256 differs: {relative}")
        if actual_blob != expected_blob:
            raise SystemExit(f"semantic owner input Git blob differs: {relative}")
        if actual_sha != committed_blob_sha256(root, commit, relative):
            raise SystemExit(f"semantic owner input differs from owner blob: {relative}")


def validate_commit(value: str, label: str) -> str:
    if len(value) != 40 or any(char not in "0123456789abcdef" for char in value):
        raise SystemExit(f"{label} must be a full lowercase commit id")
    return value


def require_ancestor(root: Path, ancestor: str, descendant: str, label: str) -> None:
    if subprocess.run(
        [
            "git",
            "--no-replace-objects",
            "-C",
            str(root),
            "merge-base",
            "--is-ancestor",
            ancestor,
            descendant,
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    ).returncode != 0:
        raise SystemExit(f"{label} is not an ancestor of capture HEAD: {ancestor}")


def package_files(package: dict[str, object]) -> list[Path]:
    """Capture the complete local package tree used by Cargo.

    Metadata targets identify entry points, but they do not enumerate every
    module, test fixture, example, build input, or included source file.  A
    path-package closure that only follows target ``src_path`` values can
    silently miss a transitive input.  The isolated run is deliberately
    fail-closed, so capture every file below the package root while excluding
    generated VCS/build trees.
    """
    manifest = Path(str(package["manifest_path"])).resolve()
    package_root = manifest.parent
    ignored = {".git", "target"}
    files = [
        path.resolve()
        for path in package_root.rglob("*")
        if path.is_file()
        and not ignored.intersection(path.relative_to(package_root).parts)
    ]
    return sorted(set(files), key=str)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--metadata", type=Path, required=True)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--git-commit", required=True)
    parser.add_argument("--semantic-owner-commit")
    parser.add_argument("--production-source-baseline-commit")
    # Keep the old spelling as a compatibility input, while requiring the
    # separate production baseline for every real capture.
    parser.add_argument("--owner-commit", dest="legacy_owner_commit")
    parser.add_argument(
        "--production-exclude-package",
        "--owner-exclude-package",
        dest="production_exclude_package",
        action="append",
        default=[],
    )
    parser.add_argument("--extra", type=Path, action="append", default=[])
    parser.add_argument("--context-extra", type=Path, action="append", default=[])
    parser.add_argument("--semantic-owner-extra", type=Path, action="append", default=[])
    parser.add_argument("--semantic-owner-extra-sha256", action="append", default=[])
    parser.add_argument("--semantic-owner-extra-blob", action="append", default=[])
    parser.add_argument(
        "--production-extra",
        "--owner-extra",
        dest="production_extra",
        type=Path,
        action="append",
        default=[],
        help="workspace/build input pinned to --production-source-baseline-commit",
    )
    args = parser.parse_args()

    semantic_owner_commit = args.semantic_owner_commit or args.legacy_owner_commit
    if semantic_owner_commit is None:
        raise SystemExit("--semantic-owner-commit is required")
    semantic_owner_commit = validate_commit(semantic_owner_commit, "semantic owner pin")
    if args.legacy_owner_commit and args.legacy_owner_commit != semantic_owner_commit:
        raise SystemExit("legacy --owner-commit disagrees with --semantic-owner-commit")
    if args.production_source_baseline_commit is None:
        raise SystemExit("--production-source-baseline-commit is required")
    capture_head = validate_commit(args.git_commit, "capture HEAD")
    production_source_baseline_commit = validate_commit(
        args.production_source_baseline_commit,
        "production source baseline pin",
    )

    root = args.root.resolve()
    metadata_path = args.metadata.resolve()
    metadata = json.loads(metadata_path.read_text())
    packages: list[tuple[str, str, str, Path, list[Path]]] = []
    local_inputs: list[Path] = []
    production_inputs: list[Path] = []
    excluded_production_packages = set(args.production_exclude_package)
    for package in metadata["packages"]:
        files = package_files(package)
        source = str(package.get("source") or "path")
        packages.append(
            (
                str(package["name"]),
                str(package["version"]),
                source,
                Path(str(package["manifest_path"])).resolve(),
                files,
            )
        )
        if package.get("source") is None:
            local_inputs.extend(files)
            if package["name"] not in excluded_production_packages:
                production_inputs.extend(files)

    extras = [require_file(path) for path in args.extra]
    context_extras = [require_file(path) for path in args.context_extra]
    semantic_owner_extras = [require_file(path) for path in args.semantic_owner_extra]
    production_extras = [require_file(path) for path in args.production_extra]
    local_inputs.extend(extras)
    local_inputs.extend(context_extras)
    local_inputs.extend(semantic_owner_extras)
    local_inputs.extend(production_extras)
    production_inputs.extend(production_extras)
    verify_local_inputs(local_inputs, root, capture_head)
    current = subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), "rev-parse", "HEAD"],
        text=True,
    ).strip()
    require_ancestor(root, semantic_owner_commit, current, "semantic owner pin")
    require_ancestor(root, production_source_baseline_commit, current, "production baseline pin")
    verify_semantic_owner_inputs(
        semantic_owner_extras,
        root,
        semantic_owner_commit,
        args.semantic_owner_extra_sha256,
        args.semantic_owner_extra_blob,
    )
    verify_production_inputs(production_inputs, root, production_source_baseline_commit)

    lines = [
        f"format={FORMAT}",
        f"git_commit={capture_head}",
        f"semantic_owner_commit={semantic_owner_commit}",
        f"production_source_baseline_commit={production_source_baseline_commit}",
        f"source_commit={semantic_owner_commit}",
        # Retained for readers of the v1 manifest; it is an alias only and is
        # never used for the transitive production closure.
        f"owner_commit={semantic_owner_commit}",
        f"metadata_sha256={sha256(metadata_path)}",
    ]
    for name, version, source, manifest, files in sorted(
        packages, key=lambda item: (item[0], item[1], str(item[3]))
    ):
        entries = [
            (shown(path, root), sha256(path))
            for path in files
        ]
        tree = hashlib.sha256(
            "\n".join(f"{name}\t{digest}" for name, digest in entries).encode()
        ).hexdigest()
        manifest_shown = shown(manifest, root)
        lines.append(
            "package="
            + "\t".join(
                (
                    name,
                    version,
                    source,
                    manifest_shown,
                    sha256(manifest),
                    str(len(entries)),
                    tree,
                )
            )
        )
        for file_name, digest in entries:
            lines.append(f"file={name}\t{version}\t{manifest_shown}\t{file_name}\t{digest}")
    for extra in extras:
        lines.append(f"extra=\t{shown(extra, root)}\t{sha256(extra)}")
    for extra in context_extras:
        lines.append(f"context_extra=\t{shown(extra, root)}\t{sha256(extra)}")
    for extra, expected_sha, expected_blob in zip(
        semantic_owner_extras,
        args.semantic_owner_extra_sha256,
        args.semantic_owner_extra_blob,
        strict=True,
    ):
        lines.append(
            f"semantic_extra=\t{shown(extra, root)}\t{expected_sha}\t{expected_blob}"
        )
    for extra in production_extras:
        lines.append(f"production_extra=\t{shown(extra, root)}\t{sha256(extra)}")
    for package_name in sorted(excluded_production_packages):
        lines.append(f"production_exclude_package={package_name}")
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
