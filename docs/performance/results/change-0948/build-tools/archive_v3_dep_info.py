#!/usr/bin/env python3
"""Copy v3 release dep-info files before target cleanup.

This is read-only with respect to source and build targets. It copies every
release ``*.d`` file from the native and allocator v3 targets into a packet
archive and writes per-lane JSON metadata plus a create-only archive receipt.
"""
from __future__ import annotations

import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import sys
from typing import Any

PACKET = Path(__file__).resolve().parents[1]
ARCHIVE = PACKET / "archives" / "dep-info-v3"
RECEIPT = PACKET / "receipts" / "dep-info-v3-archive-v1.json"
LANES = {
    "native": PACKET / "target-candidate-native-v3",
    "allocator": PACKET / "target-candidate-allocator-v3",
}
EXPECTED_COUNT = 236
EXPECTED_BUILD_COUNT = 23
EXPECTED_DEPS_COUNT = 212


class ArchiveError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ArchiveError(message)


def digest(path: Path) -> tuple[int, str]:
    require(path.is_file() and not path.is_symlink(), f"invalid regular file: {path}")
    h = hashlib.sha256()
    size = 0
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            h.update(chunk)
            size += len(chunk)
    return size, h.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    size, sha = digest(path)
    mode = path.stat().st_mode & 0o777
    return {
        "path": path.relative_to(PACKET).as_posix(),
        "bytes": size,
        "mode": mode,
        "executable": bool(mode & 0o111),
        "sha256": sha,
    }


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite retained file: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False) + "\n").encode()
    with path.open("xb") as output:
        output.write(data)
        output.flush()
        os.fsync(output.fileno())
    fd = os.open(path.parent, os.O_RDONLY | os.O_DIRECTORY)
    try:
        os.fsync(fd)
    finally:
        os.close(fd)


def main() -> int:
    require(not ARCHIVE.exists(), f"archive exists: {ARCHIVE}")
    require(not RECEIPT.exists(), f"receipt exists: {RECEIPT}")
    ARCHIVE.mkdir(parents=True)
    lanes: list[dict[str, Any]] = []
    for lane, target in LANES.items():
        release = target / "release"
        require(release.is_dir() and not release.is_symlink(), f"missing release target: {release}")
        source_paths = sorted(release.rglob("*.d"))
        require(len(source_paths) == EXPECTED_COUNT,
                f"{lane} dep-info count changed: {len(source_paths)}")
        lane_root = ARCHIVE / lane / "release"
        records: list[dict[str, Any]] = []
        build_count = 0
        deps_count = 0
        for source in source_paths:
            require(source.is_file() and not source.is_symlink(), f"invalid dep-info: {source}")
            relative = source.relative_to(release).as_posix()
            parts = source.relative_to(release).parts
            category = "build-script" if parts and parts[0] == "build" else \
                       "crate-dependency" if parts and parts[0] == "deps" else "other"
            if category == "build-script":
                build_count += 1
            if category == "crate-dependency":
                deps_count += 1
            destination = lane_root / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            require(not destination.exists(), f"archive destination exists: {destination}")
            shutil.copy2(source, destination)
            source_size, source_sha = digest(source)
            copied_size, copied_sha = digest(destination)
            require((source_size, source_sha) == (copied_size, copied_sha),
                    f"archive copy mismatch: {relative}")
            records.append({
                "source_path": source.relative_to(PACKET).as_posix(),
                "archive_path": destination.relative_to(PACKET).as_posix(),
                "release_relative_path": relative,
                "category": category,
                "bytes": source_size,
                "sha256": source_sha,
                "copy_verified": True,
            })
        require(build_count == EXPECTED_BUILD_COUNT and deps_count == EXPECTED_DEPS_COUNT,
                f"{lane} dep-info categories changed: {build_count}/{deps_count}")
        metadata = {
            "schema": "litchi.performance.0948.dep-info-lane-metadata.v1",
            "status": "copied-and-verified",
            "created_utc": _datetime.datetime.now(_datetime.timezone.utc).isoformat(),
            "lane": lane,
            "target_root": target.relative_to(PACKET).as_posix(),
            "release_root": release.relative_to(PACKET).as_posix(),
            "archive_root": (ARCHIVE / lane / "release").relative_to(PACKET).as_posix(),
            "dep_info_count": len(records),
            "build_script_dep_info_count": build_count,
            "crate_dependency_dep_info_count": deps_count,
            "files": records,
        }
        metadata_path = ARCHIVE / f"{lane}-dep-info.json"
        write_once(metadata_path, metadata)
        lanes.append({
            "lane": lane,
            "dep_info_count": len(records),
            "build_script_dep_info_count": build_count,
            "crate_dependency_dep_info_count": deps_count,
            "metadata": descriptor(metadata_path),
            "archive_root": (ARCHIVE / lane / "release").relative_to(PACKET).as_posix(),
            "files": records,
        })
    receipt = {
        "schema": "litchi.performance.0948.dep-info-v3-archive.v1",
        "status": "passed",
        "created_utc": _datetime.datetime.now(_datetime.timezone.utc).isoformat(),
        "source_build_revision": "candidate-build-v3",
        "archive_root": ARCHIVE.relative_to(PACKET).as_posix(),
        "lanes": lanes,
    }
    write_once(RECEIPT, receipt)
    size, sha = digest(RECEIPT)
    print(json.dumps({
        "status": receipt["status"],
        "receipt": RECEIPT.relative_to(PACKET).as_posix(),
        "bytes": size,
        "sha256": sha,
        "lanes": [
            {k: lane[k] for k in ("lane", "dep_info_count", "build_script_dep_info_count", "crate_dependency_dep_info_count", "metadata", "archive_root")}
            for lane in lanes
        ],
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, UnicodeError, ValueError, ArchiveError) as error:
        print(f"dep-info archive refused: {error}", file=sys.stderr)
        raise SystemExit(2)
