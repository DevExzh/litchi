#!/usr/bin/env python3
"""Hash the transitive Cargo source closure and bind production inputs to Git.

The production closure is compared with the approved source commit through
``git cat-file``. The captured evidence subtree (harness files and copied
native fixtures) is compared with the descendant HEAD's blobs, including when
those paths already exist at the approved source commit. Working-tree
cleanliness alone cannot establish either source identity.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path


FORMAT = "docx-styles-effects-cargo-source-closure-v1"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def require_file(path: Path) -> Path:
    path = path.resolve()
    if not path.is_file():
        raise SystemExit(f"source manifest input is missing: {path}")
    return path


def git_output(root: Path, *arguments: str) -> str:
    return subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), *arguments],
        text=True,
    ).strip()


def committed_blob_sha256(root: Path, commit: str, shown: str) -> str | None:
    result = subprocess.run(
        [
            "git",
            "--no-replace-objects",
            "-C",
            str(root),
            "cat-file",
            "-e",
            f"{commit}:{shown}",
        ],
        capture_output=True,
        check=False,
    )
    if result.returncode != 0:
        return None
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
    return hashlib.sha256(committed).hexdigest()


def display(path: Path, root: Path) -> str:
    path = path.resolve()
    try:
        return path.relative_to(root).as_posix()
    except ValueError as error:
        raise SystemExit(f"manifest input is outside staged checkout: {path}") from error


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
        source = Path(str(dict(target)["src_path"])).resolve()
        if source.is_file() and not source.is_relative_to(source_dir.resolve()):
            files.append(source)
    return sorted(set(files), key=str)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--metadata", type=Path, required=True)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--evidence", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--source-commit", required=True)
    parser.add_argument("--extra", type=Path, action="append", default=[])
    args = parser.parse_args()
    # A failed capture must never leave a stale receipt that a later verifier
    # could accidentally consume.  The runner supplies a fresh output path,
    # and this also makes direct replay fail closed.
    args.output.unlink(missing_ok=True)
    if len(args.source_commit) != 40 or any(
        character not in "0123456789abcdef" for character in args.source_commit
    ):
        raise SystemExit("source commit must be a full hexadecimal object ID")

    root = args.root.resolve()
    evidence = args.evidence.resolve()
    metadata = args.metadata.resolve()
    evidence_shown = display(evidence, root).rstrip("/")

    def is_evidence_path(shown: str) -> bool:
        return shown.startswith(evidence_shown + "/")

    current_head = git_output(root, "rev-parse", "HEAD")
    ancestry = subprocess.run(
        [
            "git",
            "--no-replace-objects",
            "-C",
            str(root),
            "merge-base",
            "--is-ancestor",
            args.source_commit,
            current_head,
        ],
        check=False,
    ).returncode
    if ancestry != 0:
        raise SystemExit(
            f"checkout {current_head} does not descend from source {args.source_commit}"
        )

    data = json.loads(metadata.read_text())
    packages: list[tuple[str, str, str, Path, list[Path]]] = []
    production_inputs: list[Path] = []
    local_inputs: list[Path] = []
    for package_data in data["packages"]:
        package = dict(package_data)
        source_value = package.get("source")
        # Cargo metadata includes registry packages whose manifests live in the
        # user cache, outside the staged checkout.  The reproducible closure
        # here is the local path-package closure; the lockfile freezes registry
        # resolution separately.
        if source_value is not None:
            continue
        manifest = Path(str(package["manifest_path"])).resolve()
        files = package_files(package)
        source = "path"
        packages.append(
            (
                str(package["name"]),
                str(package["version"]),
                source,
                manifest,
                files,
            )
        )
        for path in files:
            shown = display(path, root)
            # The production pin can contain an older copy of the evidence
            # harness. Evidence remains bound to the captured descendant HEAD,
            # even when a path also exists at the production source commit.
            if is_evidence_path(shown) or committed_blob_sha256(root, args.source_commit, shown) is None:
                local_inputs.append(path)
            else:
                production_inputs.append(path)

    extras = [require_file(path) for path in args.extra]
    for path in extras:
        shown = display(path, root)
        if is_evidence_path(shown) or committed_blob_sha256(root, args.source_commit, shown) is None:
            local_inputs.append(path)
        else:
            production_inputs.append(path)

    changed: list[str] = []
    for path in sorted(set(production_inputs), key=str):
        shown = display(path, root)
        expected = committed_blob_sha256(root, args.source_commit, shown)
        if expected is None or sha256(path) != expected:
            changed.append(shown)
    if changed:
        raise SystemExit(
            "production build inputs differ from the approved Git blobs: "
            + ", ".join(changed)
        )

    for path in sorted(set(local_inputs), key=str):
        shown = display(path, root)
        # A new path in the production closure would silently expand the
        # approved source pin. Captured evidence paths are admissible only
        # when their descendant HEAD blobs match the staged files.
        if not is_evidence_path(shown):
            raise SystemExit(
                "production build input is outside the approved source pin: "
                + shown
            )
        # Use the captured object ID, not a moving HEAD ref. Git status can
        # also hide changed bytes behind an assume-unchanged index bit.
        expected = committed_blob_sha256(root, current_head, shown)
        if expected is None:
            raise SystemExit(
                "evidence input is not committed at current Git HEAD: " + shown
            )
        if sha256(path) != expected:
            raise SystemExit(
                "evidence input differs from current Git HEAD: " + shown
            )

    lines = [
        f"format={FORMAT}",
        f"source_commit={args.source_commit}",
        f"git_head={current_head}",
        f"metadata_sha256={sha256(metadata)}",
        f"evidence_root={display(evidence, root)}",
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
    package_paths = {
        display(path, root)
        for _, _, _, _, files in packages
        for path in files
    }
    extra_paths: set[str] = set()
    for path in extras:
        shown = display(path, root)
        if shown in package_paths:
            continue
        if shown in extra_paths:
            raise SystemExit(f"source manifest extra is listed twice: {shown}")
        extra_paths.add(shown)
        lines.append(f"extra=\t{shown}\t{sha256(path)}")
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
