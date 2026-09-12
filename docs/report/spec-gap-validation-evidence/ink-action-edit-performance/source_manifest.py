#!/usr/bin/env python3
"""Hash every Cargo source input used by the isolated ink-action profile."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path


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


FORMAT = "ink-action-edit-build-source-v2"


def require_file(path: Path) -> Path:
    path = path.resolve()
    if not path.is_file():
        raise SystemExit(f"source manifest input is missing: {path}")
    return path


def committed_blob_sha256(root: Path, commit: str, shown: str) -> str:
    """Hash the committed blob bytes without consulting the Git index."""

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
    """Require every local build input to equal its committed Git blob.

    ``git diff`` is deliberately not used here.  Its result can be affected by
    index flags such as ``assume-unchanged``; the local SHA-256 is compared with
    bytes read by ``git cat-file`` from the committed tree, so a mutated file
    is caught even when Git's index metadata says it is clean.
    """

    root = root.resolve()
    relative: list[str] = []
    for path in sorted({path.resolve() for path in paths}, key=str):
        require_file(path)
        try:
            shown = path.relative_to(root).as_posix()
        except ValueError as error:
            raise SystemExit(
                "non-Git local input requires a retained source snapshot: "
                f"{path}"
            ) from error
        relative.append(shown)

    if not relative:
        return
    try:
        current = subprocess.check_output(
            ["git", "-C", str(root), "rev-parse", "HEAD"], text=True
        ).strip()
    except subprocess.CalledProcessError as error:
        raise SystemExit(f"cannot read committed Git head for source inputs: {root}") from error
    if current != commit:
        raise SystemExit(
            f"source input Git head changed during manifest capture: {commit} -> {current}"
        )

    missing: list[str] = []
    for shown in relative:
        tracked = subprocess.run(
            ["git", "-C", str(root), "ls-files", "--error-unmatch", "--", shown],
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

    changed: list[str] = []
    for shown in relative:
        committed_blob = committed_blob_sha256(root, commit, shown)
        working_blob = sha256(root / shown)
        if working_blob != committed_blob:
            changed.append(shown)
    if changed:
        raise SystemExit(
            "local build inputs differ from committed Git blobs: "
            + ", ".join(sorted(set(changed)))
        )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--metadata", type=Path, required=True)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--git-commit", required=True)
    parser.add_argument("--extra", type=Path, action="append", default=[])
    args = parser.parse_args()

    root = args.root.resolve()
    metadata_path = args.metadata.resolve()
    metadata = json.loads(metadata_path.read_text())
    packages: list[tuple[str, str, Path, str | None, list[Path]]] = []
    local_inputs: list[Path] = []

    for package in metadata["packages"]:
        manifest = Path(package["manifest_path"]).resolve()
        source_dir = manifest.parent / "src"
        package_files = [manifest]
        if source_dir.is_dir():
            package_files.extend(path.resolve() for path in source_dir.rglob("*") if path.is_file())
        build_script = manifest.parent / "build.rs"
        if build_script.is_file():
            package_files.append(build_script.resolve())
        for target in package["targets"]:
            target_source = Path(target["src_path"])
            if target_source.is_file() and not target_source.resolve().is_relative_to(source_dir.resolve()):
                package_files.append(target_source.resolve())
        packages.append(
            (
                package["name"],
                package["version"],
                manifest,
                package.get("source"),
                sorted(set(package_files), key=str),
            )
        )
        if package.get("source") is None:
            local_inputs.extend(package_files)

    for extra in args.extra:
        local_inputs.append(require_file(extra))

    verify_git_inputs(local_inputs, root, args.git_commit)

    lines = [
        f"format={FORMAT}",
        f"git_commit={args.git_commit}",
        f"metadata_sha256={sha256(metadata_path)}",
    ]
    for name, version, manifest, source, package_files in sorted(
        packages, key=lambda value: (value[0], value[1], str(value[2]))
    ):
        source_text = source or "path"
        package_hash_lines = [
            f"{display(path, root)}\t{sha256(path)}" for path in package_files
        ]
        tree_hash = hashlib.sha256("\n".join(package_hash_lines).encode()).hexdigest()
        manifest_display = display(manifest, root)
        lines.append(
            "package="
            + "\t".join(
                (
                    name,
                    version,
                    source_text,
                    manifest_display,
                    sha256(manifest),
                    str(len(package_files)),
                    tree_hash,
                )
            )
        )
        for shown, digest in (line.split("\t", 1) for line in package_hash_lines):
            lines.append(
                "file="
                + "\t".join((name, version, manifest_display, shown, digest))
            )
    for path in args.extra:
        resolved = path.resolve()
        lines.append(f"extra=\t{display(resolved, root)}\t{sha256(resolved)}")
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
