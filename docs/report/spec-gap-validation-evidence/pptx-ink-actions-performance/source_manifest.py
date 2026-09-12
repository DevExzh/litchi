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


def package_files(package: dict[str, object]) -> list[Path]:
    manifest = Path(str(package["manifest_path"])).resolve()
    source_dir = manifest.parent / "src"
    files = [manifest]
    if source_dir.is_dir():
        files.extend(path for path in source_dir.rglob("*") if path.is_file())
    build_script = manifest.parent / "build.rs"
    if build_script.is_file():
        files.append(build_script)
    for target in package["targets"]:  # type: ignore[index]
        source = Path(str(target["src_path"])).resolve()  # type: ignore[index]
        if source.is_file():
            files.append(source)
    return sorted(set(path.resolve() for path in files), key=str)


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
    packages: list[tuple[str, str, str, Path, list[Path]]] = []
    local_inputs: list[Path] = []
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

    extras = [require_file(path) for path in args.extra]
    local_inputs.extend(extras)
    verify_local_inputs(local_inputs, root, args.git_commit)

    lines = [
        f"format={FORMAT}",
        f"git_commit={args.git_commit}",
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
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
