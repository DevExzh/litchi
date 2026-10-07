#!/usr/bin/env python3
"""Bounded 0953 archive-manifest positive and negative controls.

This validates one extracted tree against the sealed manifest and exercises
actual rejection paths against hard-link/synthetic mutations. It never mutates
the canonical packet or extracted positive tree.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
import os
import shutil
import stat
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def digest(path: Path) -> tuple[int, int, str]:
    st = path.lstat()
    if not stat.S_ISREG(st.st_mode):
        raise ValueError(f"not regular: {path}")
    h = hashlib.sha256(); size = 0
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block); size += len(block)
    return size, stat.S_IMODE(st.st_mode), h.hexdigest()


def no_dupes(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_bytes(), object_pairs_hook=no_dupes)


def safe_rel(value: Any) -> str:
    if not isinstance(value, str) or not value or "\\" in value or "\x00" in value:
        raise ValueError(f"unsafe relative path: {value!r}")
    rel = Path(value)
    if rel.is_absolute() or any(part in ("", ".", "..") for part in rel.parts):
        raise ValueError(f"unsafe relative path: {value!r}")
    return rel.as_posix()


def expected(manifest: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    entries = manifest.get("entries")
    directories = manifest.get("directories")
    if not isinstance(entries, list) or not isinstance(directories, list):
        raise ValueError("manifest entries/directories must be arrays")
    files: dict[str, dict[str, Any]] = {}
    dirs: dict[str, dict[str, Any]] = {}
    for item in entries:
        if not isinstance(item, dict): raise ValueError("bad entry")
        path = safe_rel(item.get("path"))
        if path in files or path in dirs: raise ValueError(f"duplicate path: {path}")
        if not isinstance(item.get("bytes"), int) or not isinstance(item.get("mode"), int):
            raise ValueError(f"bad metadata: {path}")
        if not isinstance(item.get("sha256"), str) or len(item["sha256"]) != 64:
            raise ValueError(f"bad hash: {path}")
        files[path] = item
    for item in directories:
        if not isinstance(item, dict): raise ValueError("bad directory")
        path = safe_rel(item.get("path"))
        if path in files or path in dirs: raise ValueError(f"duplicate path: {path}")
        if not isinstance(item.get("mode"), int): raise ValueError(f"bad directory mode: {path}")
        dirs[path] = item
    # Every non-root parent must be represented as a directory.
    for path in [*files, *dirs]:
        parent = Path(path).parent
        while str(parent) != ".":
            if parent.as_posix() not in dirs:
                raise ValueError(f"missing parent directory: {parent.as_posix()}")
            parent = parent.parent
    return files, dirs


def actual(tree: Path) -> tuple[set[str], set[str], set[str]]:
    if tree.is_symlink() or not tree.is_dir():
        raise ValueError("tree root must be a real directory")
    files: set[str] = set(); dirs: set[str] = set(); symlinks: set[str] = set()
    for current, dirnames, filenames in os.walk(tree, topdown=True, followlinks=False):
        current_path = Path(current)
        current_rel = current_path.relative_to(tree)
        for name in list(dirnames):
            path = current_path / name
            rel = (current_rel / name).as_posix()
            if path.is_symlink():
                symlinks.add(rel); dirnames.remove(name)
            else:
                dirs.add(rel)
        for name in filenames:
            path = current_path / name
            rel = (current_rel / name).as_posix()
            if path.is_symlink(): symlinks.add(rel)
            else: files.add(rel)
    return files, dirs, symlinks


def verify(manifest: dict[str, Any], tree: Path) -> None:
    files, dirs = expected(manifest)
    afiles, adirs, symlinks = actual(tree)
    if symlinks: raise ValueError(f"symlink paths: {sorted(symlinks)}")
    if afiles != set(files):
        raise ValueError(f"file set mismatch missing={sorted(set(files)-afiles)} extra={sorted(afiles-set(files))}")
    if adirs != set(dirs):
        raise ValueError(f"directory set mismatch missing={sorted(set(dirs)-adirs)} extra={sorted(adirs-set(dirs))}")
    for path, item in files.items():
        size, mode, sha = digest(tree / path)
        if (size, mode, sha) != (item["bytes"], item["mode"], item["sha256"]):
            raise ValueError(f"file descriptor mismatch: {path}")
    for path, item in dirs.items():
        if stat.S_IMODE((tree / path).stat().st_mode) != item["mode"]:
            raise ValueError(f"directory mode mismatch: {path}")


def clone_hardlinks(source: Path, destination: Path, manifest: dict[str, Any]) -> None:
    files, dirs = expected(manifest)
    destination.mkdir(parents=True)
    for path, item in sorted(dirs.items(), key=lambda kv: (kv[0].count("/"), kv[0])):
        out = destination / path; out.mkdir(); os.chmod(out, item["mode"])
    for path in files:
        out = destination / path; out.parent.mkdir(parents=True, exist_ok=True)
        os.link(source / path, out)


def mutate_manifest(manifest: dict[str, Any], case: str) -> dict[str, Any]:
    mutated = copy.deepcopy(manifest)
    if case == "wrong_hash":
        mutated["entries"][0]["sha256"] = "0" * 64
    elif case == "wrong_mode":
        mutated["entries"][0]["mode"] ^= 0o1
    elif case == "duplicate_entry":
        mutated["entries"].append(copy.deepcopy(mutated["entries"][0]))
    elif case == "path_traversal":
        extra = copy.deepcopy(mutated["entries"][0]); extra["path"] = "../escape"; mutated["entries"].append(extra)
    elif case == "duplicate_json_key":
        # This case is handled by feeding a separately serialized duplicate-key document.
        pass
    return mutated


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--manifest", type=Path, required=True)
    ap.add_argument("--tree", type=Path, required=True)
    ap.add_argument("--work", type=Path, required=True)
    ap.add_argument("--output", type=Path, required=True)
    args = ap.parse_args()
    manifest = load_json(args.manifest)
    verify(manifest, args.tree)
    cases = ["missing_file", "extra_file", "symlink_file", "symlink_dir", "wrong_hash", "wrong_mode", "duplicate_entry", "path_traversal", "duplicate_json_key"]
    results: list[dict[str, Any]] = []
    args.work.mkdir(parents=True, exist_ok=True)
    files, dirs = expected(manifest)
    first_file = sorted(files)[0]; first_dir = sorted(dirs)[0]
    for case in cases:
        with tempfile.TemporaryDirectory(prefix=f"0953-{case}-", dir=args.work) as temp:
            tree = Path(temp) / "tree"
            clone_hardlinks(args.tree, tree, manifest)
            case_manifest = args.manifest
            rejected = False; reason = ""
            try:
                if case == "missing_file":
                    (tree / first_file).unlink()
                elif case == "extra_file":
                    extra = tree / "unexpected-extra.bin"; extra.write_bytes(b"extra")
                elif case == "symlink_file":
                    victim = tree / first_file; victim.unlink(); victim.symlink_to(tree / sorted(files)[1])
                elif case == "symlink_dir":
                    # Add a directory symlink without disturbing a non-empty
                    # expected subtree; the validator must reject it as an
                    # unexpected/symlink path.
                    (tree / "unexpected-directory").symlink_to(tree / first_dir, target_is_directory=True)
                elif case in {"wrong_hash", "wrong_mode", "duplicate_entry", "path_traversal"}:
                    case_manifest = Path(temp) / "manifest.json"
                    case_manifest.write_text(json.dumps(mutate_manifest(manifest, case)), encoding="utf-8")
                elif case == "duplicate_json_key":
                    case_manifest = Path(temp) / "manifest.json"
                    case_manifest.write_text('{"schema":"x","schema":"y"}', encoding="utf-8")
                verify(load_json(Path(case_manifest)), tree)
            except Exception as exc:
                rejected = True; reason = type(exc).__name__ + ": " + str(exc)
            if not rejected:
                raise RuntimeError(f"negative control was accepted: {case}")
            results.append({"case": case, "rejected": True, "reason": reason[:240]})
    receipt = {
        "schema": "litchi.performance.0953.archive-manifest-controls.v2",
        "status": "pass", "created_utc": datetime.now(timezone.utc).isoformat(),
        "helper": {"path": str(Path(__file__).resolve()), "bytes": Path(__file__).stat().st_size, "sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest()},
        "manifest": {"path": str(args.manifest.resolve()), "bytes": args.manifest.stat().st_size, "sha256": hashlib.sha256(args.manifest.read_bytes()).hexdigest()},
        "tree": {"path": str(args.tree.resolve()), "positive_status": "pass", "entries": len(files), "directories": len(dirs)},
        "cases": results,
        "command": "python3 -B " + str(Path(__file__).resolve()) + " --manifest <sealed-manifest> --tree <extracted-tree> --work <owned-temp> --output <receipt>",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "entries": len(files), "directories": len(dirs), "negative_controls": len(results), "output": str(args.output)}))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
