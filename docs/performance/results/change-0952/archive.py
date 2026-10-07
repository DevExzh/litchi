#!/usr/bin/env python3
"""Seal the completed 0952 forensic packet; requires explicit --seal --go."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
import subprocess
from datetime import datetime, timezone
from pathlib import Path

PACKET_DEFAULT = Path("/home/zhuhe/litchi-performance-0952")
ARCHIVE_DEFAULT = Path("/home/zhuhe/litchi-performance-0952-archive")
EXCLUDE_DIRS = {"__pycache__"}


def digest(path: Path) -> dict[str, object]:
    if not path.is_file() or path.is_symlink():
        raise RuntimeError(f"not a regular file: {path}")
    h = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block)
            size += len(block)
    mode = stat.S_IMODE(path.stat().st_mode)
    return {"path": path.as_posix(), "bytes": size, "mode": mode,
            "executable": bool(path.stat().st_mode & 0o111),
            "sha256": h.hexdigest()}


def collect(packet: Path) -> tuple[list[dict[str, object]], list[dict[str, object]], list[str]]:
    if not packet.is_dir() or packet.is_symlink():
        raise RuntimeError(f"invalid packet root: {packet}")
    entries: list[dict[str, object]] = []
    directories: list[dict[str, object]] = []
    excluded: list[str] = []
    for current, dirnames, filenames in os.walk(packet, topdown=True, followlinks=False):
        current_path = Path(current)
        kept: list[str] = []
        for name in sorted(dirnames):
            path = current_path / name
            rel = path.relative_to(packet).as_posix()
            if path.is_symlink():
                raise RuntimeError(f"symlink directory: {rel}")
            if name in EXCLUDE_DIRS:
                excluded.append(rel + "/")
            else:
                kept.append(name)
        dirnames[:] = kept
        for name in sorted(filenames):
            path = current_path / name
            rel = path.relative_to(packet).as_posix()
            if path.is_symlink():
                raise RuntimeError(f"symlink file: {rel}")
            if path.suffix == ".pyc":
                excluded.append(rel)
                continue
            item = digest(path)
            item["path"] = rel
            entries.append(item)
        if current_path != packet and not kept and not filenames:
            directories.append({"path": current_path.relative_to(packet).as_posix(),
                                "mode": stat.S_IMODE(current_path.stat().st_mode)})
    entries.sort(key=lambda item: str(item["path"]))
    directories.sort(key=lambda item: str(item["path"]))
    excluded.sort()
    return entries, directories, excluded


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet", type=Path, default=PACKET_DEFAULT)
    parser.add_argument("--archive-root", type=Path, default=ARCHIVE_DEFAULT)
    parser.add_argument("--seal", action="store_true")
    parser.add_argument("--go", action="store_true")
    args = parser.parse_args()
    if args.seal and not args.go:
        raise SystemExit("--seal requires --go")
    packet = args.packet.resolve()
    root = args.archive_root.resolve()
    root.mkdir(parents=True, exist_ok=True)
    entries, directories, excluded = collect(packet)
    manifest = {
        "schema": "litchi.performance.0952.external-archive-manifest.v1",
        "status": "sealed" if args.seal else "planning_only_no_archive",
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "packet_root": str(packet),
        "entry_count": len(entries),
        "logical_bytes": sum(int(item["bytes"]) for item in entries),
        "entries": entries,
        "directories": directories,
        "excluded": excluded,
        "symlinks": [],
        "dependencies": {
            "0950": {"path": "/home/zhuhe/litchi-performance-0950", "copied": False},
            "0951": {"path": "/home/zhuhe/litchi-performance-0951", "copied": False},
            "0942": {"path": "/home/zhuhe/litchi-performance-0942", "copied": False},
        },
        "cleanup_requires_explicit_go": True,
    }
    manifest_path = root / "manifest-v1.json"
    if manifest_path.exists() and args.seal:
        raise RuntimeError(f"refusing existing sealed manifest: {manifest_path}")
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    if not args.seal:
        print(json.dumps({"status": manifest["status"], "entries": len(entries), "bytes": manifest["logical_bytes"]}))
        return 0
    for name in ("archive-0952-v1.tar", "archive-0952-v1.tar.zst", "selection-v1.list", "archive-receipt-v1.json"):
        if (root / name).exists():
            raise RuntimeError(f"refusing existing archive artifact: {root / name}")
    selection = root / "selection-v1.list"
    with selection.open("wb") as stream:
        for item in directories + entries:
            stream.write(str(item["path"]).encode() + b"\0")
    tar_path = root / "archive-0952-v1.tar"
    zst_path = root / "archive-0952-v1.tar.zst"
    tmp_zst = root / ".archive-0952-v1.tar.zst.tmp"
    subprocess.run([
        "/usr/bin/tar", "--create", "--no-recursion", "--sort=name", "--mtime=@0",
        "--owner=0", "--group=0", "--numeric-owner", "--format=pax",
        "--file", str(tar_path), "--null", "--verbatim-files-from", "--files-from", str(selection)
    ], cwd=packet, check=True)
    subprocess.run(["/usr/bin/zstd", "-3", "--no-progress", "--force", str(tar_path), "-o", str(tmp_zst)], check=True)
    tmp_zst.replace(zst_path)
    tar_desc = digest(tar_path)
    tar_path.unlink()
    archive_desc = digest(zst_path)
    receipt = {
        "schema": "litchi.performance.0952.external-archive-receipt.v1",
        "status": "sealed",
        "manifest": digest(manifest_path),
        "selection": digest(selection),
        "compressed_archive": archive_desc,
        "uncompressed_tar": {**tar_desc, "deleted_after_compression": True},
        "helper": digest(Path(__file__).resolve()),
        "packet_census": {"entries": len(entries), "logical_bytes": manifest["logical_bytes"], "directories": len(directories), "excluded": excluded, "symlinks": []},
        "dependencies": manifest["dependencies"],
        "regeneration": f"python3 -B {Path(__file__).resolve()} --packet {packet} --archive-root {root} --seal --go",
    }
    receipt_path = root / "archive-receipt-v1.json"
    receipt_path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "sealed", "manifest": str(manifest_path), "archive": str(zst_path), "entries": len(entries), "logical_bytes": manifest["logical_bytes"], "archive_bytes": archive_desc["bytes"], "archive_sha256": archive_desc["sha256"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
