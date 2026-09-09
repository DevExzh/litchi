#!/usr/bin/env python3
"""Create and verify a portable, content-addressed evidence-bundle seal.

``verify.py`` checks the meaning of the 0483 evidence records.  This helper
adds the outer file-set boundary: ``SHA256SUMS`` contains one deterministic
SHA-256 entry for every regular file in the copied bundle.  The manifest is
excluded from its own inventory, and is written with exclusive creation so a
seal can never be silently replaced.

The module deliberately keeps inventory work independent from the structured
verifier.  That makes the inventory functions useful for small synthetic
tests, while the command-line modes compose them in the required order:
``--seal`` verifies structured evidence before creating the manifest, and
``--verify`` checks the exact file set before invoking the structured
verifier.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import stat
from typing import Any, Iterable, NoReturn

import verify


ROOT = Path(__file__).resolve().parent
MANIFEST_NAME = "SHA256SUMS"
MANIFEST_SCHEMA = "docx-tail-append-sha256sums-v1"
SHA256 = re.compile(r"^[0-9a-f]{64}$")


class SealError(ValueError):
    """Raised when a bundle cannot be safely sealed or verified."""


def fail(message: str) -> NoReturn:
    raise SealError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _root_path(root: Path | str) -> Path:
    """Return an absolute bundle root, rejecting a symlinked root itself."""

    path = Path(root)
    require(not path.is_symlink(), "bundle root must not be a symlink")
    require(path.is_dir(), "bundle root must be a directory")
    try:
        return path.resolve(strict=True)
    except OSError as error:
        fail(f"bundle root cannot be resolved: {error}")


def safe_relative(value: str, label: str = "path") -> str:
    """Validate the unambiguous POSIX path representation used in a seal."""

    require(isinstance(value, str) and value, f"{label}: relative path required")
    require("\\" not in value, f"{label}: backslashes are forbidden")
    require(
        not any(character in value for character in "\x00\r\n\u2028\u2029")
        and not any(ord(character) < 0x20 or ord(character) == 0x7F for character in value),
        f"{label}: control/newline characters are forbidden",
    )
    path = PurePosixPath(value)
    require(not path.is_absolute() and value not in {".", ".."}, f"{label}: absolute/parent path forbidden")
    require(all(part not in {"", ".", ".."} for part in path.parts), f"{label}: path escapes bundle")
    require(path.as_posix() == value, f"{label}: non-canonical path")
    return value


def _relative_path(root: Path, path: Path) -> str:
    try:
        relative = path.relative_to(root)
    except ValueError:
        fail(f"path escapes bundle: {path}")
    return safe_relative(PurePosixPath(*relative.parts).as_posix(), "bundle path")


def _sha256_file(path: Path, label: str) -> str:
    """Hash one regular, non-symlink file without following a link."""

    try:
        info = path.stat(follow_symlinks=False)
    except OSError as error:
        fail(f"{label}: cannot stat: {error}")
    require(stat.S_ISREG(info.st_mode), f"{label}: regular file required")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"{label}: cannot read: {error}")
    return digest.hexdigest()


def _iter_files(root: Path) -> Iterable[tuple[str, Path]]:
    """Yield every regular bundle file in deterministic traversal order.

    ``os.scandir`` with ``follow_symlinks=False`` lets us reject directory
    links before a recursive walk could accidentally leave the bundle.  Empty
    directories do not belong to the file-set seal and are therefore ignored.
    """

    def walk(directory: Path, parent_parts: tuple[str, ...]) -> Iterable[tuple[str, Path]]:
        try:
            entries = sorted(os.scandir(directory), key=lambda entry: entry.name)
        except OSError as error:
            fail(f"cannot scan {directory}: {error}")
        for entry in entries:
            relative = safe_relative(PurePosixPath(*(parent_parts + (entry.name,))).as_posix(), "bundle path")
            current = directory / entry.name
            try:
                entry_info = entry.stat(follow_symlinks=False)
            except OSError as error:
                fail(f"{relative}: cannot stat: {error}")
            mode = entry_info.st_mode
            if stat.S_ISLNK(mode):
                fail(f"{relative}: symlinks are forbidden")
            if stat.S_ISDIR(mode):
                if entry.name == "__pycache__" or "__pycache__" in parent_parts:
                    fail(f"{relative}: __pycache__ is forbidden")
                yield from walk(current, parent_parts + (entry.name,))
                continue
            require(stat.S_ISREG(mode), f"{relative}: regular file required")
            require(entry.name != "__pycache__" and "__pycache__" not in parent_parts, f"{relative}: __pycache__ is forbidden")
            require(not entry.name.endswith(".pyc"), f"{relative}: .pyc files are forbidden")
            if relative == MANIFEST_NAME:
                continue
            yield relative, current

    yield from walk(root, ())


def inventory(root: Path | str) -> dict[str, str]:
    """Return the sorted relative-path to SHA-256 mapping for a bundle.

    The manifest itself is excluded.  The returned dictionary is insertion
    ordered by path so callers can use it directly for deterministic output.
    """

    bundle = _root_path(root)
    entries = sorted(_iter_files(bundle), key=lambda item: item[0])
    result: dict[str, str] = {}
    for relative, path in entries:
        require(relative not in result, f"duplicate bundle path: {relative}")
        result[relative] = _sha256_file(path, relative)
    return result


def render_manifest(entries: dict[str, str]) -> bytes:
    """Render a canonical GNU-style SHA256SUMS file."""

    lines: list[str] = []
    for relative in sorted(entries):
        safe_relative(relative, "manifest path")
        digest = entries[relative]
        require(isinstance(digest, str) and SHA256.fullmatch(digest) is not None, f"manifest digest for {relative} is malformed")
        lines.append(f"{digest}  {relative}")
    return ("\n".join(lines) + ("\n" if lines else "")).encode("utf-8")


def parse_manifest(data: bytes) -> dict[str, str]:
    """Parse and validate canonical-looking SHA256SUMS records."""

    require(isinstance(data, bytes), "manifest data must be bytes")
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"SHA256SUMS is not UTF-8: {error}")
    if not text:
        return {}
    require(text.endswith("\n"), "SHA256SUMS must end with a newline")
    result: dict[str, str] = {}
    for number, line in enumerate(text[:-1].split("\n"), 1):
        require(line and "\r" not in line, f"SHA256SUMS line {number} is malformed")
        fields = line.split("  ", 1)
        require(len(fields) == 2 and SHA256.fullmatch(fields[0]) is not None, f"SHA256SUMS line {number} has malformed digest")
        relative = safe_relative(fields[1], f"SHA256SUMS line {number} path")
        require(relative != MANIFEST_NAME, f"SHA256SUMS line {number} includes the manifest itself")
        require(relative not in result, f"SHA256SUMS line {number} duplicates {relative}")
        result[relative] = fields[0]
    return result


def _read_manifest(bundle: Path) -> bytes:
    path = bundle / MANIFEST_NAME
    try:
        info = path.stat(follow_symlinks=False)
    except OSError as error:
        fail(f"{MANIFEST_NAME}: cannot stat: {error}")
    require(stat.S_ISREG(info.st_mode), f"{MANIFEST_NAME}: regular file required")
    try:
        return path.read_bytes()
    except OSError as error:
        fail(f"{MANIFEST_NAME}: cannot read: {error}")


def seal_write(root: Path | str) -> dict[str, Any]:
    """Write a new exclusive SHA256SUMS file and return its receipt.

    This function only seals the file boundary.  The command-line ``--seal``
    caller performs the structured evidence verification immediately before
    invoking it, which keeps this function usable with synthetic test roots.
    """

    bundle = _root_path(root)
    entries = inventory(bundle)
    encoded = render_manifest(entries)
    manifest = bundle / MANIFEST_NAME
    try:
        with manifest.open("xb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError:
        fail(f"{MANIFEST_NAME}: refusing to replace an existing seal")
    except OSError as error:
        fail(f"{MANIFEST_NAME}: cannot create exclusive seal: {error}")
    return {
        "schema": MANIFEST_SCHEMA,
        "status": "pass",
        "manifest": MANIFEST_NAME,
        "files": len(entries),
        "bytes": len(encoded),
        "sha256": hashlib.sha256(encoded).hexdigest(),
    }


def verify_inventory(root: Path | str) -> dict[str, Any]:
    """Verify exact current bundle files and bytes against SHA256SUMS."""

    bundle = _root_path(root)
    raw = _read_manifest(bundle)
    expected = parse_manifest(raw)
    actual = inventory(bundle)
    require(set(actual) == set(expected), "SHA256SUMS file set differs from the bundle")
    for relative in sorted(actual):
        require(actual[relative] == expected[relative], f"SHA256SUMS digest differs: {relative}")
    require(raw == render_manifest(expected), "SHA256SUMS is not in canonical deterministic form")
    return {
        "schema": MANIFEST_SCHEMA,
        "status": "pass",
        "manifest": MANIFEST_NAME,
        "files": len(actual),
        "bytes": len(raw),
        "sha256": hashlib.sha256(raw).hexdigest(),
    }


def _run(mode: str, root: Path) -> dict[str, Any]:
    bundle = _root_path(root)
    if mode == "seal":
        structured = verify.verify_bundle(bundle)
        sealed = seal_write(bundle)
        return {"mode": "seal", "root": str(bundle), "structured": structured, "seal": sealed}
    checked = verify_inventory(bundle)
    structured = verify.verify_bundle(bundle)
    return {"mode": "verify", "root": str(bundle), "seal": checked, "structured": structured}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--seal", action="store_true", help="verify structured evidence and create SHA256SUMS")
    modes.add_argument("--verify", action="store_true", help="verify SHA256SUMS and then structured evidence")
    parser.add_argument("--root", type=Path, default=ROOT, help="copied evidence-bundle root")
    options = parser.parse_args()
    try:
        result = _run("seal" if options.seal else "verify", options.root)
    except (SealError, verify.VerificationError) as error:
        print(f"seal verification failed: {error}")
        raise SystemExit(1)
    print(json.dumps(result, sort_keys=True))


if __name__ == "__main__":
    main()
