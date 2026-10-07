#!/usr/bin/env python3
"""Create a deterministic external 0948 evidence archive.

The source packet is read-only.  Target build trees are excluded except for the
historical v3 release executables whose receipt paths are dereferenced by the
qualification reuse step.  All other packet files, source trees, receipts,
raw reports, and historical attempts are retained.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import stat
import subprocess
import sys

SCHEMA = "litchi.performance.0948.external-archive.v2"
ALLOW_TARGET_FILES = {
    "target-candidate-native-v3/release/litchi-perf-baseline",
    "target-candidate-allocator-v3/release/litchi-perf-baseline-alloc",
}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            b = f.read(1024 * 1024)
            if not b:
                return h.hexdigest()
            h.update(b)


def disposable_target_dir(rel: Path) -> bool:
    parts = rel.parts
    if not parts:
        return False
    # Preserve the two explicitly allowlisted v3 executable paths.  Everything
    # else under the top-level build targets is an intermediate.
    if parts[0] in {
        "target-candidate-native-v2",
        "target-candidate-native-v3",
        "target-candidate-allocator-v2",
        "target-candidate-allocator-v3",
    }:
        return True
    # Quality runs have several independently named Cargo target trees.  Their
    # source trees and receipts remain included; only these build trees go.
    if parts[0] == "quality" and any(
        p == "target" or p.startswith("target-") or p.endswith("-target")
        for p in parts[1:]
    ):
        return True
    return False


def included(rel: Path) -> bool:
    s = rel.as_posix()
    if s in ALLOW_TARGET_FILES:
        return True
    if disposable_target_dir(rel):
        return False
    return True


def collect(root: Path, archive_root: Path, archive_name: str) -> tuple[list[dict], Path]:
    entries: list[dict] = []
    paths: list[str] = []
    skipped: list[str] = []
    for dirpath, dirnames, filenames in os.walk(root, topdown=True, followlinks=False):
        d = Path(dirpath)
        rel_d = d.relative_to(root)
        # Do not descend into disposable targets.  Keep parent names out of the
        # list; the selected executable exception is added explicitly below.
        keep_dirs = []
        for name in sorted(dirnames):
            rel = rel_d / name
            if (d / name).is_symlink():
                raise RuntimeError(f"symlink directory is not admissible: {rel}")
            if disposable_target_dir(rel):
                skipped.append(rel.as_posix() + "/")
            else:
                keep_dirs.append(name)
        dirnames[:] = keep_dirs
        for name in sorted(filenames):
            p = d / name
            rel = p.relative_to(root)
            if not included(rel):
                skipped.append(rel.as_posix())
                continue
            if p.is_symlink():
                raise RuntimeError(f"symlink is not admissible: {rel}")
            st = p.stat()
            if not stat.S_ISREG(st.st_mode):
                raise RuntimeError(f"non-regular file is not admissible: {rel}")
            entries.append({
                "path": rel.as_posix(),
                "bytes": st.st_size,
                "mode_bits": stat.S_IMODE(st.st_mode),
                "sha256": sha256(p),
            })
            paths.append(rel.as_posix())
            if len(entries) % 1000 == 0:
                print(f"hashed {len(entries)} files", file=sys.stderr, flush=True)
    # Explicitly preserve the historical executable paths even if their parent
    # target directory was pruned above.
    for rel_s in sorted(ALLOW_TARGET_FILES):
        p = root / rel_s
        if not p.is_file() or p.is_symlink():
            raise RuntimeError(f"required historical executable missing: {rel_s}")
        if not any(e["path"] == rel_s for e in entries):
            st = p.stat()
            entries.append({
                "path": rel_s,
                "bytes": st.st_size,
                "mode_bits": stat.S_IMODE(st.st_mode),
                "sha256": sha256(p),
            })
            paths.append(rel_s)
    entries.sort(key=lambda e: e["path"])
    paths = sorted(set(paths))
    if len(paths) != len(entries):
        raise RuntimeError("duplicate selected paths")
    # Selection list is NUL-delimited for tar and remains external custody data.
    list_path = archive_root / "selection-v2.list"
    list_path.write_bytes(b"".join((p.encode() + b"\0") for p in paths))
    helper = Path(__file__)
    helper_sha = sha256(helper)
    manifest = {
        "schema": SCHEMA,
        "version": 2,
        "source_root": str(root),
        "source_root_basename": root.name,
        "archive_name": archive_name,
        "selection": {
            "file_count": len(entries),
            "total_bytes": sum(e["bytes"] for e in entries),
            "entries": entries,
            "skipped_path_count": len(skipped),
            "skipped_path_sha256": hashlib.sha256(("\n".join(sorted(skipped)) + "\n").encode()).hexdigest(),
            "rules": [
                "all source-packet files are retained",
                "top-level target-candidate-v2/v3 native and allocator trees are excluded except the two v3 release executables",
                "quality Cargo target trees and final-v3-target are excluded",
                "no symlinks or non-regular files are accepted",
            ],
            "allow_target_files": sorted(ALLOW_TARGET_FILES),
        },
        "helper": {
            "path": str(helper),
            "bytes": helper.stat().st_size,
            "sha256": helper_sha,
        },
        "archive_metadata": {
            "manifest_path": "ARCHIVE_METADATA/manifest-v2.json",
            "helper_path": "ARCHIVE_METADATA/archive-0948-v2.py",
            "selection_path": "ARCHIVE_METADATA/selection-v2.list",
        },
    }
    manifest_path = archive_root / "manifest-v2.json"
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n")
    # Keep copies under names that can be included inside the archive without
    # changing the source packet.
    metadata = archive_root / "ARCHIVE_METADATA"
    metadata.mkdir(exist_ok=True)
    (metadata / "manifest-v2.json").write_bytes(manifest_path.read_bytes())
    (metadata / "archive-0948-v2.py").write_bytes(helper.read_bytes())
    (metadata / "selection-v2.list").write_bytes(list_path.read_bytes())
    return entries, list_path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=Path, required=True)
    ap.add_argument("--archive-root", type=Path, required=True)
    ap.add_argument("--archive-name", default="archive-0948-v2.tar.zst")
    ap.add_argument("--no-compress", action="store_true")
    args = ap.parse_args()
    root = args.root.resolve()
    ar = args.archive_root.resolve()
    ar.mkdir(parents=True, exist_ok=True)
    entries, list_path = collect(root, ar, args.archive_name)
    tar_path = ar / args.archive_name.removesuffix(".zst")
    if tar_path.suffix != ".tar":
        tar_path = ar / (args.archive_name + ".tar")
    # Metadata is added before source paths; fixed owner/time/order options make
    # the archive reproducible for identical source bytes and tool versions.
    tar_cmd = [
        "tar", "--create", "--file", str(tar_path), "--format=posix",
        "--sort=name", "--mtime=@0", "--owner=0", "--group=0",
        "--numeric-owner", "--pax-option=delete=atime,delete=ctime",
        "--directory", str(ar),
        "ARCHIVE_METADATA/manifest-v2.json",
        "ARCHIVE_METADATA/archive-0948-v2.py",
        "ARCHIVE_METADATA/selection-v2.list",
        "--directory", str(root), "--null", "--files-from", str(list_path),
    ]
    print("running:", " ".join(tar_cmd), file=sys.stderr, flush=True)
    subprocess.run(tar_cmd, check=True)
    result = tar_path
    if not args.no_compress:
        zstd_cmd = ["zstd", "--force", "-T4", "-19", str(tar_path), "-o", str(ar / args.archive_name)]
        print("running:", " ".join(zstd_cmd), file=sys.stderr, flush=True)
        subprocess.run(zstd_cmd, check=True)
        tar_path.unlink()
        result = ar / args.archive_name
    result_sha = sha256(result)
    receipt = {
        "schema": "litchi.performance.0948.external-archive-receipt.v1",
        "helper": {"path": str(Path(__file__)), "sha256": sha256(Path(__file__))},
        "manifest": {"path": str(ar / "manifest-v2.json"), "sha256": sha256(ar / "manifest-v2.json"), "file_count": len(entries), "source_bytes": sum(e["bytes"] for e in entries)},
        "archive": {"path": str(result), "bytes": result.stat().st_size, "sha256": result_sha, "format": "tar.zst" if result.suffix == ".zst" else "tar"},
        "tar": {"version": subprocess.check_output(["tar", "--version"], text=True).splitlines()[0]},
        "compression": subprocess.check_output(["zstd", "--version"], text=True).splitlines()[0],
        "cleanup": {"status": "not_started", "policy": "only disposable target intermediates after moved replay; retain source trees, historical binaries, role copies, and dep-info"},
    }
    rec_path = ar / "archive-receipt-v2.json"
    rec_path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
    print(json.dumps(receipt, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
