#!/usr/bin/env python3
"""Export the frozen 0948 candidate source for the later build lanes.

The quality runner keeps two boundary-policy files beside its raw source
census.  They are quality-only inputs and are deliberately excluded from the
standalone build source.  This helper copies exactly the candidate raw census,
checks every byte against the quality freeze, and writes a create-only full
source manifest plus a candidate freeze receipt.  It never invokes Cargo.
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
QUALITY_FREEZE = PACKET / "quality" / "final" / "source-freeze-v1.json"
QUALITY_ROOT = PACKET / "quality" / "final" / "candidate-source"
BASELINE_MANIFEST = PACKET / "baseline-source-manifest-v1.json"
SOURCE = PACKET / "candidate-build-v1"
MANIFEST = PACKET / "candidate-build-manifest-v1.json"
RECEIPT = PACKET / "receipts" / "candidate-freeze-v1.json"

QUALITY_FREEZE_SHA256 = "4040f1367dbc75fff3cc383bc187cf9d5550ff4936d0a9863d963fcebc28e7a3"
SOURCE_COMMIT = "b6c5c80ffbcb60736d6c4c2f3b590838640008a9"
QUALITY_ONLY_FILES = {
    "tools/check_crate_boundaries.py",
    "tools/crate_boundaries.json",
}


class ExportError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ExportError(message)


def strict_json(path: Path) -> Any:
    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    try:
        return json.loads(path.read_text(encoding="utf-8"),
                          object_pairs_hook=reject_duplicates)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ExportError(f"cannot read JSON {path}: {error}") from error


def sha256_file(path: Path) -> tuple[int, str]:
    digest = hashlib.sha256()
    total = 0
    try:
        with path.open("rb") as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(chunk)
                total += len(chunk)
    except OSError as error:
        raise ExportError(f"cannot hash {path}: {error}") from error
    return total, digest.hexdigest()


def normalized_relative(name: Any, label: str) -> str:
    require(isinstance(name, str) and name, f"{label} path is missing")
    path = Path(name)
    require(not path.is_absolute() and path.as_posix() == name
            and ".." not in path.parts,
            f"{label} path is not normalized: {name!r}")
    return name


def descriptor(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(),
            f"descriptor target is not a regular file: {path}")
    size, digest = sha256_file(path)
    mode = path.stat().st_mode & 0o777
    return {"path": path.relative_to(PACKET).as_posix(),
            "bytes": size, "mode": mode,
            "executable": bool(mode & 0o111), "sha256": digest}


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite retained artifact: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False,
                          allow_nan=False) + "\n").encode("utf-8")
    try:
        with path.open("xb") as output:
            output.write(encoded)
            output.flush()
            os.fsync(output.fileno())
    except OSError as error:
        raise ExportError(f"cannot write {path}: {error}") from error
    directory = os.open(path.parent, os.O_RDONLY | os.O_DIRECTORY)
    try:
        os.fsync(directory)
    finally:
        os.close(directory)


def load_inputs() -> tuple[dict[str, Any], dict[str, Any], dict[str, dict[str, Any]],
                            dict[str, dict[str, Any]], list[dict[str, Any]]]:
    require(QUALITY_FREEZE.is_file(), f"missing quality freeze: {QUALITY_FREEZE}")
    size, digest = sha256_file(QUALITY_FREEZE)
    require(digest == QUALITY_FREEZE_SHA256,
            f"quality freeze SHA changed: {digest}")
    freeze = strict_json(QUALITY_FREEZE)
    require(freeze.get("schema") == "litchi.performance.0948-quality-source-freeze.v1",
            "quality freeze schema changed")
    require(freeze.get("root_head") == SOURCE_COMMIT,
            "quality freeze root commit changed")
    candidate = freeze.get("candidate")
    baseline = freeze.get("baseline")
    require(isinstance(candidate, dict) and isinstance(baseline, dict),
            "quality freeze role bindings are missing")
    require(candidate.get("quality_root") == str(QUALITY_ROOT),
            "quality candidate root changed")
    require(candidate.get("source_root") == str(PACKET / "candidate-source"),
            "quality candidate source root changed")
    require(candidate.get("raw_file_count") == 10789
            and candidate.get("raw_total_bytes") == 260075614,
            "quality candidate census changed")
    raw_entries = candidate.get("raw_manifest")
    require(isinstance(raw_entries, list) and len(raw_entries) == 10789,
            "quality candidate raw manifest count changed")
    quality_entries: dict[str, dict[str, Any]] = {}
    for item in raw_entries:
        require(isinstance(item, dict), "quality raw manifest item is not an object")
        name = normalized_relative(item.get("path"), "quality raw manifest")
        require(name not in quality_entries, f"duplicate quality raw path: {name}")
        require(isinstance(item.get("bytes"), int) and item["bytes"] >= 0
                and isinstance(item.get("sha256"), str)
                and len(item["sha256"]) == 64,
                f"quality raw manifest digest is malformed: {name}")
        quality_entries[name] = item

    baseline_raw = strict_json(BASELINE_MANIFEST)
    require(isinstance(baseline_raw, dict)
            and baseline_raw.get("source_commit") == SOURCE_COMMIT,
            "baseline source manifest commit changed")
    baseline_items = baseline_raw.get("files")
    require(isinstance(baseline_items, list) and len(baseline_items) == len(quality_entries),
            "baseline source manifest census count changed")
    baseline_entries: dict[str, dict[str, Any]] = {}
    for item in baseline_items:
        require(isinstance(item, dict), "baseline manifest item is not an object")
        name = normalized_relative(item.get("path"), "baseline manifest")
        require(name not in baseline_entries, f"duplicate baseline path: {name}")
        baseline_entries[name] = item
    require(set(baseline_entries) == set(quality_entries),
            "quality and baseline source path sets differ")
    changed_files = freeze.get("changed_files")
    require(isinstance(changed_files, list) and len(changed_files) == 10,
            "quality changed-file list changed")
    changes: list[dict[str, Any]] = []
    listed: set[str] = set()
    for item in changed_files:
        require(isinstance(item, dict), "quality changed-file item is not an object")
        name = normalized_relative(item.get("path"), "quality changed-file")
        require(name not in listed, f"duplicate changed file: {name}")
        listed.add(name)
        base = item.get("baseline")
        cand = item.get("candidate")
        require(isinstance(base, list) and len(base) == 2
                and isinstance(cand, list) and len(cand) == 2,
                f"quality changed-file tuple malformed: {name}")
        require([baseline_entries[name].get("bytes"), baseline_entries[name].get("sha256")] == base,
                f"quality changed-file baseline mismatch: {name}")
        require([quality_entries[name].get("bytes"), quality_entries[name].get("sha256")] == cand,
                f"quality changed-file candidate mismatch: {name}")
        changes.append({
            "path": name,
            "baseline_bytes": base[0],
            "baseline_sha256": base[1],
            "candidate_bytes": cand[0],
            "candidate_sha256": cand[1],
        })
    require(listed == set(name for name in baseline_entries
                          if (baseline_entries[name].get("bytes"), baseline_entries[name].get("sha256"))
                          != (quality_entries[name].get("bytes"), quality_entries[name].get("sha256"))),
            "quality changed-file list is not the manifest delta")
    # Keep these values in the return object so the caller can bind the exact
    # quality receipt and changed-file tuples into both output receipts.
    _ = size
    return freeze, candidate, quality_entries, baseline_entries, changes


def verify_quality_tree(entries: dict[str, dict[str, Any]]) -> None:
    require(QUALITY_ROOT.is_dir() and not QUALITY_ROOT.is_symlink(),
            f"quality candidate root is not a directory: {QUALITY_ROOT}")
    actual: dict[str, Path] = {}
    for path in QUALITY_ROOT.rglob("*"):
        require(not path.is_symlink(), f"quality source contains symlink: {path}")
        if path.is_dir():
            continue
        require(path.is_file(), f"quality source contains non-regular file: {path}")
        name = normalized_relative(path.relative_to(QUALITY_ROOT).as_posix(),
                                   "quality source")
        require(name not in actual, f"duplicate quality source path: {name}")
        actual[name] = path
    require(set(actual) == set(entries) | QUALITY_ONLY_FILES,
            "quality root differs from frozen raw census")
    for name, item in entries.items():
        path = actual[name]
        size, digest = sha256_file(path)
        require(size == item["bytes"] and digest == item["sha256"],
                f"quality source differs from frozen census: {name}")


def copy_source(entries: dict[str, dict[str, Any]]) -> None:
    require(not SOURCE.exists(), f"refusing to overwrite source export: {SOURCE}")
    temporary = PACKET / f".candidate-build-v1.tmp-{os.getpid()}"
    require(not temporary.exists(), f"temporary export path already exists: {temporary}")
    try:
        temporary.mkdir()
        for name in sorted(entries):
            source = QUALITY_ROOT / name
            target = temporary / name
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target, follow_symlinks=False)
        os.rename(temporary, SOURCE)
    except OSError as error:
        shutil.rmtree(temporary, ignore_errors=True)
        raise ExportError(f"candidate source copy failed: {error}") from error


def verify_export(entries: dict[str, dict[str, Any]]) -> None:
    require(SOURCE.is_dir() and not SOURCE.is_symlink(),
            f"candidate export is not a directory: {SOURCE}")
    observed: dict[str, Path] = {}
    for path in SOURCE.rglob("*"):
        require(not path.is_symlink(), f"candidate export contains symlink: {path}")
        if path.is_dir():
            continue
        require(path.is_file(), f"candidate export contains non-regular file: {path}")
        name = normalized_relative(path.relative_to(SOURCE).as_posix(), "candidate export")
        observed[name] = path
    require(set(observed) == set(entries), "candidate export file set differs from manifest")
    for name, item in entries.items():
        size, digest = sha256_file(observed[name])
        require(size == item["bytes"] and digest == item["sha256"],
                f"candidate export differs from manifest: {name}")


def main() -> int:
    freeze, candidate, entries, baseline_entries, changes = load_inputs()
    verify_quality_tree(entries)
    copy_source(entries)
    verify_export(entries)

    manifest = {
        "schema": "litchi.performance.0948.candidate-build-source-manifest.v1",
        "source_root": SOURCE.name,
        "base_commit": SOURCE_COMMIT,
        "source_commit": SOURCE_COMMIT,
        "file_count": len(entries),
        "bytes": sum(item["bytes"] for item in entries.values()),
        "quality_source_freeze": {
            "path": QUALITY_FREEZE.relative_to(PACKET).as_posix(),
            "bytes": QUALITY_FREEZE.stat().st_size,
            "sha256": QUALITY_FREEZE_SHA256,
        },
        "changes": changes,
        "files": [
            {
                "path": name,
                "bytes": entries[name]["bytes"],
                "mode_bits": entries[name].get("mode_bits"),
                "sha256": entries[name]["sha256"],
            }
            for name in sorted(entries)
        ],
    }
    write_once(MANIFEST, manifest)
    manifest_size, manifest_sha = sha256_file(MANIFEST)
    receipt = {
        "schema": "litchi.performance.0948.candidate-freeze.v1",
        "status": "frozen",
        "source_root": SOURCE.name,
        "base_commit": SOURCE_COMMIT,
        "source_commit": SOURCE_COMMIT,
        "file_count": manifest["file_count"],
        "source_bytes": manifest["bytes"],
        "changes": changes,
        "source_manifest": {
            "path": MANIFEST.relative_to(PACKET).as_posix(),
            "bytes": manifest_size,
            "sha256": manifest_sha,
        },
        "quality_source_freeze": {
            "path": QUALITY_FREEZE.relative_to(PACKET).as_posix(),
            "bytes": QUALITY_FREEZE.stat().st_size,
            "sha256": QUALITY_FREEZE_SHA256,
        },
        "quality_candidate_root": str(QUALITY_ROOT),
        "created_utc": _datetime.datetime.now(_datetime.timezone.utc).isoformat(),
        "source_copy": {
            "input": QUALITY_ROOT.relative_to(PACKET).as_posix(),
            "raw_census_verified": True,
            "quality_only_files_excluded": sorted(QUALITY_ONLY_FILES),
        },
    }
    write_once(RECEIPT, receipt)
    print(json.dumps({
        "status": "pass",
        "source": SOURCE.relative_to(PACKET).as_posix(),
        "manifest": {"path": MANIFEST.relative_to(PACKET).as_posix(),
                     "bytes": manifest_size, "sha256": manifest_sha},
        "freeze_receipt": descriptor(RECEIPT),
        "file_count": manifest["file_count"],
        "source_bytes": manifest["bytes"],
        "quality_freeze_sha256": QUALITY_FREEZE_SHA256,
        "quality_source_commit": freeze["root_head"],
        "candidate_changed_file_count": len(changes),
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, ExportError) as error:
        print(f"candidate source export refused: {error}", file=sys.stderr)
        raise SystemExit(2)
