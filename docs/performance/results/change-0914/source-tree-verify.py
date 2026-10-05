#!/usr/bin/env python3
"""Check that the frozen extracted source is exactly the selected Git tree."""

from __future__ import annotations

import hashlib
import os
from pathlib import Path
import subprocess
import sys
import tarfile


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
EXTERNAL = Path(os.environ.get("LITCHI_PERF_0914_ROOT", "/home/zhuhe/litchi-performance-0914"))
SOURCE = EXTERNAL / "source" / "litchi"
ARCHIVE = EXTERNAL / "source-HEAD-code.tar"
REVISION = os.environ.get("LITCHI_PERF_0914_REVISION", "a4ca85e28cdea151176534b7f8dc912253a1ca1e")
PATHS = (
    "Cargo.toml", "clippy.toml", "deny.toml", "rust-toolchain.toml", "rustfmt.toml",
    ".cargo", "crates", "test-data/poi/test-data/spreadsheet/54016.xls",
    "test-data/rtf/watermark.rtf", "tools/perf-baseline",
)


def tree() -> dict[str, tuple[str, str]]:
    raw = subprocess.check_output(
        ["git", "-c", "core.quotepath=false", "ls-tree", "-r", "-z", "--full-tree",
         f"--format=%(objectmode) %(objectname) %(path)", REVISION, "--", *PATHS],
        cwd=ROOT,
    )
    entries = {}
    for record in raw.rstrip(b"\0").split(b"\0"):
        mode, object_id, path = record.split(b" ", 2)
        entries[path.decode()] = (mode.decode(), object_id.decode())
    return entries


def files(root: Path) -> set[str]:
    found: set[str] = set()

    def visit(directory: Path) -> None:
        for entry in os.scandir(directory):
            relative = str(Path(os.path.relpath(entry.path, root)))
            if entry.is_symlink():
                found.add(relative)
            elif entry.is_dir(follow_symlinks=False):
                visit(Path(entry.path))
            elif entry.is_file(follow_symlinks=False):
                found.add(relative)

    visit(root)
    return found


def main() -> int:
    entries = tree()
    query = "".join(f"{oid}\n" for _, oid in entries.values()).encode()
    output = subprocess.check_output(["git", "cat-file", "--batch"], cwd=ROOT, input=query)
    blobs: dict[str, tuple[str, bytes]] = {}
    offset = 0
    for relative, (mode, oid) in entries.items():
        line_end = output.index(b"\n", offset)
        object_id, kind, size_text = output[offset:line_end].split()
        offset = line_end + 1
        size = int(size_text)
        data = output[offset:offset + size]
        offset += size + 1
        if object_id.decode() != oid or kind != b"blob":
            raise RuntimeError(f"Git object lookup failed for {relative}")
        blobs[relative] = (mode, data)

    missing = sorted(set(entries) - files(SOURCE))
    extra = sorted(files(SOURCE) - set(entries))
    mismatches: list[str] = []
    for relative, (mode, expected) in blobs.items():
        path = SOURCE / relative
        if mode == "120000":
            if not path.is_symlink() or os.readlink(path).encode() != expected:
                mismatches.append(relative)
        elif not path.is_file() or path.is_symlink() or path.read_bytes() != expected:
            mismatches.append(relative)
        elif (os.stat(path).st_mode & 0o111) != (int(mode, 8) & 0o111):
            mismatches.append(relative)

    with tarfile.open(ARCHIVE) as stream:
        archived = set()
        archive_mismatches = []
        for member in stream.getmembers():
            if not (member.isfile() or member.issym()):
                continue
            relative = member.name.removeprefix("litchi/")
            archived.add(relative)
            if relative not in blobs:
                continue
            _, expected = blobs[relative]
            actual = member.linkname.encode() if member.issym() else stream.extractfile(member).read()
            if actual != expected or (member.mode & 0o111) != (int(entries[relative][0], 8) & 0o111):
                archive_mismatches.append(relative)
    tar_missing = sorted(set(entries) - archived)
    tar_extra = sorted(archived - set(entries))
    result = {
        "revision": REVISION,
        "tree_entries": len(entries),
        "extracted_entries": len(files(SOURCE)),
        "tar_files": len(archived),
        "missing": len(missing),
        "extra": len(extra),
        "content_or_mode_mismatches": len(mismatches),
        "archive_content_or_mode_mismatches": len(archive_mismatches),
        "tar_missing": len(tar_missing),
        "tar_extra": len(tar_extra),
        "archive_bytes": ARCHIVE.stat().st_size,
        "archive_sha256": hashlib.sha256(ARCHIVE.read_bytes()).hexdigest(),
    }
    print(result)
    if any(result[key] for key in ("missing", "extra", "content_or_mode_mismatches", "archive_content_or_mode_mismatches", "tar_missing", "tar_extra")):
        print({"missing": missing[:10], "extra": extra[:10], "mismatches": mismatches[:10],
               "archive_mismatches": archive_mismatches[:10], "tar_missing": tar_missing[:10],
               "tar_extra": tar_extra[:10]}, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
