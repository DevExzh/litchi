#!/usr/bin/env python3
"""Hash the transitive Cargo source closure and bind local inputs to Git blobs."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path


FORMAT = "xlsb-model-identity-cargo-source-closure-v3"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def display(path: Path, root: Path) -> str:
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


def committed_blob_sha256(root: Path, commit: str, shown: str) -> str:
    """Hash committed bytes without consulting the index or worktree."""

    try:
        committed = subprocess.check_output(
            [
                "git",
                "--no-replace-objects",
                "-C",
                str(root),
                "cat-file",
                "blob",
                f"{commit}:{shown}",
            ]
        )
    except subprocess.CalledProcessError as error:
        raise SystemExit(f"cannot read committed source input blob: {shown}") from error
    return hashlib.sha256(committed).hexdigest()


def verify_git_inputs(paths: list[Path], root: Path, commit: str) -> None:
    """Require local build inputs to equal blobs in the pinned Git tree.

    Comparing working bytes with ``git cat-file`` catches unstaged, staged, and
    ``assume-unchanged`` edits.  A status-only check cannot provide that
    guarantee because index flags can hide a modified worktree file.
    """

    root = root.resolve()
    relative: list[str] = []
    for path in sorted({path.resolve() for path in paths}, key=str):
        require_file(path)
        try:
            relative.append(path.relative_to(root).as_posix())
        except ValueError as error:
            raise SystemExit(
                "non-Git local input requires a retained source snapshot: "
                f"{path}"
            ) from error
    if not relative:
        return
    try:
        current = subprocess.check_output(
            ["git", "--no-replace-objects", "-C", str(root), "rev-parse", "HEAD"],
            text=True,
        ).strip()
    except subprocess.CalledProcessError as error:
        raise SystemExit(f"cannot read committed Git head for source inputs: {root}") from error
    if current != commit:
        raise SystemExit(f"source input Git head changed: {commit} -> {current}")

    missing: list[str] = []
    for shown in relative:
        tracked = subprocess.run(
            [
                "git",
                "--no-replace-objects",
                "-C",
                str(root),
                "ls-files",
                "--error-unmatch",
                "--",
                shown,
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        if tracked.returncode != 0:
            missing.append(shown)
    if missing:
        raise SystemExit(
            "local build inputs are not tracked in the committed checkout: "
            + ", ".join(sorted(missing))
        )

    changed = [
        shown
        for shown in relative
        if sha256(root / shown) != committed_blob_sha256(root, commit, shown)
    ]
    if changed:
        raise SystemExit(
            "local build inputs differ from committed Git blobs: "
            + ", ".join(sorted(set(changed)))
        )


def package_files(package: dict[str, object]) -> list[Path]:
    manifest = Path(str(package["manifest_path"])).resolve()
    files = [manifest]
    package_root = manifest.parent
    source_dir = package_root / "src"
    if source_dir.is_dir():
        files.extend(path.resolve() for path in source_dir.rglob("*") if path.is_file())
    build_script = package_root / "build.rs"
    if build_script.is_file():
        files.append(build_script.resolve())
    for target in package.get("targets", []):
        target = dict(target)
        source = Path(str(target["src_path"])).resolve()
        if source.is_file() and not source.is_relative_to(source_dir.resolve()):
            files.append(source)
    return sorted(set(files), key=str)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--metadata", type=Path, required=True)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--git-commit", required=True)
    parser.add_argument("--extra", type=Path, action="append", default=[])
    args = parser.parse_args()
    if len(args.git_commit) != 40 or any(
        character not in "0123456789abcdef" for character in args.git_commit
    ):
        raise SystemExit("source manifest Git commit must be a full hexadecimal object ID")

    root = args.root.resolve()
    metadata = args.metadata.resolve()
    data = json.loads(metadata.read_text())
    packages: list[tuple[str, str, str, Path, list[Path]]] = []
    local_inputs: list[Path] = []
    for package in data["packages"]:
        package = dict(package)
        manifest = Path(str(package["manifest_path"])).resolve()
        files = package_files(package)
        if package.get("source") is None:
            local_inputs.extend(files)
        packages.append(
            (
                str(package["name"]),
                str(package["version"]),
                str(package.get("source") or "path"),
                manifest,
                files,
            )
        )

    for extra in args.extra:
        local_inputs.append(require_file(extra))

    verify_git_inputs(local_inputs, root, args.git_commit)

    head = subprocess.run(
        ["git", "--no-replace-objects", "-C", str(root), "rev-parse", "HEAD"],
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()
    lines = [
        f"format={FORMAT}",
        f"git_commit={args.git_commit}",
        f"git_head={head}",
        f"metadata_sha256={sha256(metadata)}",
    ]
    for name, version, source, manifest, files in sorted(
        packages, key=lambda item: (item[0], item[1], str(item[3]))
    ):
        file_lines = [
            f"{display(path, root)}\t{sha256(path)}" for path in files
        ]
        tree_hash = hashlib.sha256("\n".join(file_lines).encode()).hexdigest()
        manifest_display = display(manifest, root)
        lines.append(
            "package="
            + "\t".join(
                (
                    name,
                    version,
                    source,
                    manifest_display,
                    sha256(manifest),
                    str(len(files)),
                    tree_hash,
                )
            )
        )
        for file_line in file_lines:
            shown, digest = file_line.split("\t", 1)
            lines.append(
                "file="
                + "\t".join((name, version, manifest_display, shown, digest))
            )
    for extra in args.extra:
        extra = extra.resolve()
        lines.append(f"extra=\t{display(extra, root)}\t{sha256(extra)}")
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
