#!/usr/bin/env python3
"""Export the reviewed 0948 candidate-source-v2 tree for build custody.

This prepares the immutable ``candidate-build-v1`` packet only.  It compares
every source file with the frozen baseline census, pins the ten expected
candidate deltas, and writes a create-only full census and source-freeze
receipt.  The quality-v1 receipt is retained as provenance only: the receipt
explicitly records that final quality-v2 validation is still required before a
build may start.  This helper never invokes Cargo.
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
INPUT_SOURCE = PACKET / "candidate-source-v2"
BASELINE_MANIFEST = PACKET / "baseline-source-manifest-v1.json"
QUALITY_PREDECESSOR = PACKET / "quality" / "final" / "source-freeze-v1.json"
SOURCE = PACKET / "candidate-build-v1"
MANIFEST = PACKET / "candidate-build-manifest-v1.json"
RECEIPT = PACKET / "receipts" / "candidate-freeze-v1.json"

SOURCE_COMMIT = "b6c5c80ffbcb60736d6c4c2f3b590838640008a9"
QUALITY_PREDECESSOR_SHA256 = (
    "4040f1367dbc75fff3cc383bc187cf9d5550ff4936d0a9863d963fcebc28e7a3"
)
EXPECTED_FILE_COUNT = 10789
EXPECTED_SOURCE_BYTES = 260075521

# These are the only candidate-source-v2 changes relative to the frozen
# baseline.  Pinning the complete delta prevents an accidental copy of a
# stale v1 candidate or a newer unreviewed checkout.
EXPECTED_CHANGED = {
    "crates/litchi-opc/src/lib.rs":
        (4507, "1a9c1cc6810b343aa5647aab9aa60e09adf14fc63b6cc0459196316ffac442a2"),
    "crates/litchi-opc/src/source_backed.rs":
        (950512, "571ec3e2e08d11626b5ce489c5d7de498445c1dd72441692306772c1893c3498"),
    "crates/litchi-xlsx/src/cell_values/facts_oracle_tests.rs":
        (22277, "d944229e02497fffb2d1307141a7b77646e8b61fb3e6a11a9864a59be9b57792"),
    "crates/litchi-xlsx/src/cell_values/snapshot.rs":
        (142066, "d1afeb511def3ab46d55b77ee067d5937d90ead82b38b8d58b414ddf3f1238e4"),
    "crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs":
        (34873, "8b2079cd3c5437d7e7113e5d4486bed9ed8eb677e87ace82c63b8ce1559f5084"),
    "crates/litchi-xlsx/src/raw/worksheet/edit/package.rs":
        (29071, "6903258ce4162c5208995d0436bd3f6b0c163ca6ab206bafad6b3f7bb87247ee"),
    "crates/litchi-xlsx/src/raw/worksheet/edit/tests.rs":
        (56639, "82a921f56b589e555f64c12d74595011bb6562e2ad9503bb95d99d38ab664cc1"),
    "crates/xml-minifier/src/audit.rs":
        (139194, "29efd9aa99c9491a1e86cfc25f1bbb31b3da7b6cbfc4438a1638c15b18b95cee"),
    "crates/xml-minifier/src/audit/namespaces.rs":
        (32468, "2b24bf90377ba75db3a15eeaa53d010bf817b17b2741b8b20e6e377fd34bbdfe"),
    "crates/xml-minifier/tests/source_replacement.rs":
        (47326, "600712ff53ad145fa82659989104a9b496e2b03afd1295f20605f29ca7a5615e"),
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


def source_files(root: Path, label: str) -> dict[str, Path]:
    require(root.is_dir() and not root.is_symlink(),
            f"{label} source root is not a directory: {root}")
    result: dict[str, Path] = {}
    for path in root.rglob("*"):
        require(not path.is_symlink(), f"{label} source contains symlink: {path}")
        if path.is_dir():
            continue
        require(path.is_file(), f"{label} source contains non-regular file: {path}")
        name = normalized_relative(path.relative_to(root).as_posix(), label)
        require(name not in result, f"duplicate {label} source path: {name}")
        result[name] = path
    return result


def load_baseline() -> dict[str, dict[str, Any]]:
    raw = strict_json(BASELINE_MANIFEST)
    require(isinstance(raw, dict) and raw.get("source_commit") == SOURCE_COMMIT,
            "baseline source manifest commit changed")
    items = raw.get("files")
    require(isinstance(items, list) and len(items) == EXPECTED_FILE_COUNT,
            "baseline source manifest census count changed")
    result: dict[str, dict[str, Any]] = {}
    for item in items:
        require(isinstance(item, dict), "baseline manifest entry is not an object")
        name = normalized_relative(item.get("path"), "baseline manifest")
        require(name not in result, f"duplicate baseline path: {name}")
        require(isinstance(item.get("bytes"), int) and item["bytes"] >= 0
                and isinstance(item.get("sha256"), str)
                and len(item["sha256"]) == 64,
                f"baseline manifest digest is malformed: {name}")
        result[name] = item
    return result


def load_candidate(baseline: dict[str, dict[str, Any]],
                   observed: dict[str, Path]) -> tuple[dict[str, dict[str, Any]],
                                                         list[dict[str, Any]]]:
    require(set(observed) == set(baseline),
            "candidate-source-v2 file set differs from the baseline census")
    entries: dict[str, dict[str, Any]] = {}
    changes: list[dict[str, Any]] = []
    for name in sorted(observed):
        size, digest = sha256_file(observed[name])
        entries[name] = {
            "path": name,
            "bytes": size,
            "mode_bits": observed[name].stat().st_mode & 0o777,
            "sha256": digest,
        }
        old = baseline[name]
        if (size, digest) != (old["bytes"], old["sha256"]):
            changes.append({
                "path": name,
                "baseline_bytes": old["bytes"],
                "baseline_sha256": old["sha256"],
                "candidate_bytes": size,
                "candidate_sha256": digest,
            })
    require(len(entries) == EXPECTED_FILE_COUNT
            and sum(item["bytes"] for item in entries.values()) == EXPECTED_SOURCE_BYTES,
            "candidate-source-v2 census count or bytes changed")
    actual_changed = {
        item["path"]: (item["candidate_bytes"], item["candidate_sha256"])
        for item in changes
    }
    require(actual_changed == EXPECTED_CHANGED,
            f"candidate-source-v2 delta changed: {sorted(actual_changed)}")
    require(len(changes) == len(EXPECTED_CHANGED),
            "candidate-source-v2 changed-file count changed")
    return entries, changes


def copy_source(source_files_by_name: dict[str, Path]) -> None:
    require(not SOURCE.exists(), f"refusing to overwrite source export: {SOURCE}")
    temporary = PACKET / f".candidate-build-v1.tmp-{os.getpid()}"
    require(not temporary.exists(), f"temporary export path already exists: {temporary}")
    try:
        temporary.mkdir()
        for name in sorted(source_files_by_name):
            target = temporary / name
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_files_by_name[name], target, follow_symlinks=False)
        os.rename(temporary, SOURCE)
    except OSError as error:
        shutil.rmtree(temporary, ignore_errors=True)
        raise ExportError(f"candidate source copy failed: {error}") from error


def verify_export(entries: dict[str, dict[str, Any]]) -> None:
    observed = source_files(SOURCE, "candidate-build-v1")
    require(set(observed) == set(entries), "candidate export file set differs from manifest")
    for name, item in entries.items():
        size, digest = sha256_file(observed[name])
        require(size == item["bytes"] and digest == item["sha256"],
                f"candidate export differs from manifest: {name}")


def main() -> int:
    require(not SOURCE.exists() and not MANIFEST.exists() and not RECEIPT.exists(),
            "candidate-build-v1 output is already present; refusing a second export")
    require(QUALITY_PREDECESSOR.is_file(),
            f"missing retained quality predecessor: {QUALITY_PREDECESSOR}")
    quality_size, quality_sha = sha256_file(QUALITY_PREDECESSOR)
    require(quality_sha == QUALITY_PREDECESSOR_SHA256,
            f"quality predecessor SHA changed: {quality_sha}")
    baseline = load_baseline()
    source_map = source_files(INPUT_SOURCE, "candidate-source-v2")
    entries, changes = load_candidate(baseline, source_map)
    copy_source(source_map)
    verify_export(entries)

    manifest = {
        "schema": "litchi.performance.0948.candidate-build-source-manifest.v1",
        "source_root": SOURCE.name,
        "base_commit": SOURCE_COMMIT,
        "source_commit": SOURCE_COMMIT,
        "file_count": len(entries),
        "bytes": sum(item["bytes"] for item in entries.values()),
        "candidate_input_root": INPUT_SOURCE.name,
        "quality_predecessor": {
            "path": QUALITY_PREDECESSOR.relative_to(PACKET).as_posix(),
            "bytes": quality_size,
            "sha256": quality_sha,
            "scope": "v1 provenance only; v2 quality compile remains required",
        },
        "changes": changes,
        "files": [entries[name] for name in sorted(entries)],
    }
    write_once(MANIFEST, manifest)
    manifest_size, manifest_sha = sha256_file(MANIFEST)
    receipt = {
        "schema": "litchi.performance.0948.candidate-freeze.v1",
        "status": "frozen-source-pending-quality-v2",
        "source_root": SOURCE.name,
        "candidate_input_root": INPUT_SOURCE.name,
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
        "quality_predecessor": {
            "path": QUALITY_PREDECESSOR.relative_to(PACKET).as_posix(),
            "bytes": quality_size,
            "sha256": quality_sha,
            "scope": "provenance only; does not qualify candidate-source-v2",
        },
        "quality_v2_required_before_build": True,
        "created_utc": _datetime.datetime.now(_datetime.timezone.utc).isoformat(),
        "source_copy": {
            "input": INPUT_SOURCE.relative_to(PACKET).as_posix(),
            "full_census_verified": True,
            "changed_file_count": len(changes),
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
        "changed_file_count": len(changes),
        "quality_v2_required_before_build": True,
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, ExportError) as error:
        print(f"candidate source export refused: {error}", file=sys.stderr)
        raise SystemExit(2)
