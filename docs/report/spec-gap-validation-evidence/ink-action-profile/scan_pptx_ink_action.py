#!/usr/bin/env python3
"""Bounded, deterministic scan for PPTX InkAction package evidence.

The scanner reads ZIP metadata first and then reads only members below the
declared limits. It records a SHA-256 for every candidate package so a receipt
can be tied to an exact checkout. It searches raw member bytes for namespace
and element markers; a hit is evidence to inspect, while no hit is only a
negative result for the recorded corpus.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from collections import Counter
from pathlib import Path
from zipfile import BadZipFile, ZipFile


PACKAGE_SUFFIXES = frozenset({".pptx", ".pptm", ".ppsx", ".potx", ".ppsm", ".potm"})
ARCHIVE_SUFFIXES = frozenset({".zip", ".jar"})
ROOTS = ("3rdparty", "test-data", "docs")
TOKENS = (
    b"http://schemas.microsoft.com/office/powerpoint/2014/inkAction",
    b"iact:",
    b"inkAction",
    b"2014/inkAction",
)
MAX_PACKAGE_BYTES = 64 * 1024 * 1024
MAX_PACKAGE_UNCOMPRESSED_BYTES = 64 * 1024 * 1024
MAX_MEMBER_BYTES = 16 * 1024 * 1024
MAX_ZIP_MEMBERS = 10_000
CHUNK_BYTES = 64 * 1024


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        while chunk := source.read(CHUNK_BYTES):
            digest.update(chunk)
    return digest.hexdigest()


def find_packages(repo: Path) -> list[Path]:
    paths: set[Path] = set()
    for root_name in ROOTS:
        root = repo / root_name
        if not root.is_dir():
            continue
        paths.update(
            path
            for path in root.rglob("*")
            if path.is_file() and path.suffix.lower() in PACKAGE_SUFFIXES
        )
    return sorted(paths)


def find_archives(repo: Path) -> list[Path]:
    paths: set[Path] = set()
    for root_name in ROOTS:
        root = repo / root_name
        if not root.is_dir():
            continue
        paths.update(
            path
            for path in root.rglob("*")
            if path.is_file() and path.suffix.lower() in ARCHIVE_SUFFIXES
        )
    return sorted(paths)


def member_hits(source, tokens: tuple[bytes, ...]) -> list[str]:
    """Search a member without materializing more than one bounded chunk."""

    overlap = max(len(token) for token in tokens) - 1
    carry = b""
    found: set[bytes] = set()
    while chunk := source.read(CHUNK_BYTES):
        window = carry + chunk
        for token in tokens:
            if token in window:
                found.add(token)
        carry = window[-overlap:]
    return [token.decode("ascii") for token in tokens if token in found]


def scan_package(path: Path, repo: Path) -> dict:
    entry = {
        "path": path.relative_to(repo).as_posix(),
        "size_bytes": path.stat().st_size,
        "sha256": sha256_file(path),
        "status": "scanned",
        "members_total": 0,
        "members_scanned": 0,
        "members_skipped_size": 0,
        "hits": [],
    }
    if entry["size_bytes"] > MAX_PACKAGE_BYTES:
        entry["status"] = "skipped_package_size"
        return entry
    try:
        with ZipFile(path) as archive:
            infos = archive.infolist()
            entry["members_total"] = len(infos)
            uncompressed = sum(info.file_size for info in infos)
            entry["uncompressed_bytes"] = uncompressed
            if len(infos) > MAX_ZIP_MEMBERS:
                entry["status"] = "skipped_member_count"
                return entry
            if uncompressed > MAX_PACKAGE_UNCOMPRESSED_BYTES:
                entry["status"] = "skipped_uncompressed_size"
                return entry
            for info in infos:
                if info.is_dir():
                    continue
                if info.file_size > MAX_MEMBER_BYTES:
                    entry["members_skipped_size"] += 1
                    continue
                entry["members_scanned"] += 1
                with archive.open(info, "r") as source:
                    matches = member_hits(source, TOKENS)
                if matches:
                    entry["hits"].append(
                        {
                            "member": info.filename,
                            "size_bytes": info.file_size,
                            "tokens": matches,
                        }
                    )
    except (OSError, BadZipFile, RuntimeError, ValueError) as error:
        entry["status"] = "error"
        entry["error"] = f"{type(error).__name__}: {error}"
    return entry


def scan_archive_metadata(path: Path, repo: Path) -> dict:
    entry = {
        "path": path.relative_to(repo).as_posix(),
        "size_bytes": path.stat().st_size,
        "sha256": sha256_file(path),
        "status": "scanned",
        "members_total": 0,
        "nested_package_entries": [],
    }
    if entry["size_bytes"] > MAX_PACKAGE_BYTES:
        entry["status"] = "skipped_archive_size"
        return entry
    try:
        with ZipFile(path) as archive:
            infos = archive.infolist()
            entry["members_total"] = len(infos)
            if len(infos) > MAX_ZIP_MEMBERS:
                entry["status"] = "skipped_member_count"
                return entry
            for info in infos:
                if info.is_dir():
                    continue
                if Path(info.filename).suffix.lower() in PACKAGE_SUFFIXES:
                    entry["nested_package_entries"].append(
                        {"member": info.filename, "size_bytes": info.file_size}
                    )
    except (OSError, BadZipFile, RuntimeError, ValueError) as error:
        entry["status"] = "error"
        entry["error"] = f"{type(error).__name__}: {error}"
    return entry


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", type=Path, default=Path(__file__).parents[4])
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    repo = args.repo.resolve()
    packages = find_packages(repo)
    entries = [scan_package(path, repo) for path in packages]
    archives = find_archives(repo)
    archive_entries = [scan_archive_metadata(path, repo) for path in archives]
    status_counts = Counter(entry["status"] for entry in entries)
    suffix_counts = Counter(Path(entry["path"]).suffix.lower() for entry in entries)
    root_counts = Counter(entry["path"].split("/", 1)[0] for entry in entries)
    hit_count = sum(len(entry["hits"]) for entry in entries)
    archive_status_counts = Counter(entry["status"] for entry in archive_entries)
    nested_package_count = sum(
        len(entry["nested_package_entries"]) for entry in archive_entries
    )
    receipt = {
        "scanner": "scan_pptx_ink_action.py",
        "version": 1,
        "scanner_sha256": sha256_file(Path(__file__).resolve()),
        "roots": list(ROOTS),
        "suffixes": sorted(PACKAGE_SUFFIXES),
        "archive_suffixes": sorted(ARCHIVE_SUFFIXES),
        "tokens": [token.decode("ascii") for token in TOKENS],
        "limits": {
            "max_package_bytes": MAX_PACKAGE_BYTES,
            "max_package_uncompressed_bytes": MAX_PACKAGE_UNCOMPRESSED_BYTES,
            "max_member_bytes": MAX_MEMBER_BYTES,
            "max_zip_members": MAX_ZIP_MEMBERS,
            "chunk_bytes": CHUNK_BYTES,
        },
        "summary": {
            "package_count": len(entries),
            "package_count_by_root": dict(sorted(root_counts.items())),
            "package_count_by_suffix": dict(sorted(suffix_counts.items())),
            "status_counts": dict(sorted(status_counts.items())),
            "member_count": sum(entry["members_total"] for entry in entries),
            "member_scanned_count": sum(
                entry["members_scanned"] for entry in entries
            ),
            "member_skipped_size_count": sum(
                entry["members_skipped_size"] for entry in entries
            ),
            "member_hit_count": hit_count,
            "archive_count": len(archive_entries),
            "archive_status_counts": dict(sorted(archive_status_counts.items())),
            "nested_package_entry_count": nested_package_count,
        },
        "packages": entries,
        "archives": archive_entries,
    }
    output = json.dumps(receipt, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.write_text(output, encoding="utf-8")
    else:
        print(output, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
