#!/usr/bin/env python3
"""Verify that candidate-source-v6 is test-only relative to compiled v5.

The v3 release binaries were built from candidate-build-v3, which is a full
copy of candidate-source-v5.  This read-only bridge proves that v6 changes
exactly one XML-minifier test file and that neither native nor allocator v3
release dep-info (including build-script inputs) names that file.  It writes a
create-only receipt and never edits either source tree or a build target.
"""
from __future__ import annotations

import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import stat
import sys
from typing import Any

PACKET = Path(__file__).resolve().parents[1]
SOURCE_V5 = PACKET / "candidate-source-v5"
SOURCE_V6 = PACKET / "candidate-source-v6"
COMPILED_SOURCE = PACKET / "candidate-build-v3"
SOURCE_MANIFEST = PACKET / "candidate-build-manifest-v3.json"
SOURCE_FREEZE = PACKET / "receipts" / "candidate-freeze-v3.json"
PRESERVATION = PACKET / "receipts" / "candidate-v3-preservation-v1.json"
RECEIPT = PACKET / "receipts" / "source-v5-v6-reuse-bridge-v1.json"
CHANGED_PATH = "crates/xml-minifier/tests/source_replacement.rs"

EXPECTED_V5 = {
    "file_count": 10789,
    "bytes": 260076537,
    "tree_digest": "228e74ebb6fee5bbd64d1ee81a4ec8334a4d1277cbf89d9b365c3bbe7cae8a28",
}
EXPECTED_V6 = {
    "file_count": 10789,
    "bytes": 260076604,
    "tree_digest": "e432a06166c4fa615ab8f9b4a69f979fb88590debfee13d5d9a46f00b5cbf985",
}
EXPECTED_DELTA = {
    "v5_bytes": 47326,
    "v5_sha256": "600712ff53ad145fa82659989104a9b496e2b03afd1295f20605f29ca7a5615e",
    "v6_bytes": 47393,
    "v6_sha256": "cca934098bb7307bc61df450852b3f97f02cfdf1f537469cef580686b7498ace",
}


class BridgeError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise BridgeError(message)


def sha256_file(path: Path) -> tuple[int, str]:
    require(path.is_file() and not path.is_symlink(), f"not a regular file: {path}")
    digest = hashlib.sha256()
    total = 0
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
            total += len(chunk)
    return total, digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    size, digest = sha256_file(path)
    mode = path.stat().st_mode & 0o777
    return {
        "path": path.relative_to(PACKET).as_posix(),
        "bytes": size,
        "mode": mode,
        "executable": bool(mode & 0o111),
        "sha256": digest,
    }


def source_census(root: Path) -> tuple[dict[str, tuple[int, str, int]], dict[str, Any]]:
    require(root.is_dir() and not root.is_symlink(), f"source root missing: {root}")
    entries: dict[str, tuple[int, str, int]] = {}
    total = 0
    for path in root.rglob("*"):
        require(not path.is_symlink(), f"source contains symlink: {path}")
        if path.is_dir():
            continue
        require(path.is_file(), f"source contains non-regular entry: {path}")
        relative = path.relative_to(root).as_posix()
        require(relative not in entries, f"duplicate source path: {relative}")
        size, digest = sha256_file(path)
        mode = path.stat().st_mode & 0o777
        entries[relative] = (size, digest, mode)
        total += size

    tree = hashlib.sha256()
    # Match the frozen source exporter: Path ordering is component-wise,
    # unlike lexical ordering of slash-separated strings.
    for relative in sorted(entries, key=lambda name: Path(name).parts):
        size, digest, _mode = entries[relative]
        encoded = relative.encode("utf-8")
        tree.update(len(encoded).to_bytes(8, "big"))
        tree.update(encoded)
        tree.update(size.to_bytes(8, "big"))
        tree.update(bytes.fromhex(digest))
    summary = {
        "root": root.relative_to(PACKET).as_posix(),
        "file_count": len(entries),
        "bytes": total,
        "tree_digest": tree.hexdigest(),
    }
    return entries, summary


def json_no_duplicates(path: Path) -> Any:
    def reject(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    with path.open(encoding="utf-8") as source:
        return json.load(source, object_pairs_hook=reject)


def compare_sources(
    v5: dict[str, tuple[int, str, int]],
    v6: dict[str, tuple[int, str, int]],
) -> dict[str, Any]:
    require(set(v5) == set(v6), "v5/v6 source file sets differ")
    changed: list[dict[str, Any]] = []
    mode_changes: list[str] = []
    for relative in sorted(v5):
        old = v5[relative]
        new = v6[relative]
        if old[2] != new[2]:
            mode_changes.append(relative)
        if old[:2] != new[:2]:
            changed.append({
                "path": relative,
                "v5_bytes": old[0],
                "v5_sha256": old[1],
                "v6_bytes": new[0],
                "v6_sha256": new[1],
            })
    require(not mode_changes, f"source mode changes found: {mode_changes}")
    require(len(changed) == 1 and changed[0]["path"] == CHANGED_PATH,
            f"unexpected v5/v6 changed paths: {[x['path'] for x in changed]}")
    require({k: changed[0][k] for k in EXPECTED_DELTA} == EXPECTED_DELTA,
            f"test-only delta changed: {changed[0]}")
    return {
        "same_file_set": True,
        "mode_changes": [],
        "changed_file_count": len(changed),
        "changed_files": changed,
    }


def compare_compiled_source(
    v5: dict[str, tuple[int, str, int]],
    compiled: dict[str, tuple[int, str, int]],
) -> dict[str, Any]:
    require(set(v5) == set(compiled), "compiled source file set differs from v5")
    mismatches: list[str] = []
    for relative in sorted(v5):
        if v5[relative] != compiled[relative]:
            mismatches.append(relative)
    require(not mismatches, f"compiled source differs from v5: {mismatches[:8]}")
    return {"same_file_set": True, "same_file_metadata": True, "mismatches": []}


def dep_info_scan(target: Path) -> dict[str, Any]:
    release = target / "release"
    require(release.is_dir() and not release.is_symlink(), f"release target missing: {release}")
    needle_path = CHANGED_PATH.encode("utf-8")
    needle_name = Path(CHANGED_PATH).name.encode("utf-8")
    records: list[dict[str, Any]] = []
    build_count = 0
    deps_count = 0
    path_hits: list[str] = []
    for path in sorted(release.rglob("*.d")):
        require(path.is_file() and not path.is_symlink(), f"bad dep-info file: {path}")
        payload = path.read_bytes()
        relative = path.relative_to(PACKET).as_posix()
        parts = path.relative_to(release).parts
        is_build_script = bool(parts and parts[0] == "build")
        is_dependency = bool(parts and parts[0] == "deps")
        if is_build_script:
            build_count += 1
        if is_dependency:
            deps_count += 1
        has_path = needle_path in payload
        has_name = needle_name in payload
        if has_path or has_name:
            path_hits.append(relative)
        size, digest = sha256_file(path)
        records.append({
            "path": relative,
            "bytes": size,
            "mode": path.stat().st_mode & 0o777,
            "sha256": digest,
            "category": "build-script" if is_build_script else
                        "crate-dependency" if is_dependency else "other",
            "mentions_changed_test_path": has_path,
            "mentions_changed_test_basename": has_name,
        })
    require(records, f"no release dep-info files found under {release}")
    require(not path_hits, f"changed test file appears in release dep-info: {path_hits}")
    require(build_count > 0 and deps_count > 0,
            f"dep-info categories incomplete: build={build_count}, deps={deps_count}")
    return {
        "target_root": target.relative_to(PACKET).as_posix(),
        "release_root": release.relative_to(PACKET).as_posix(),
        "dep_info_count": len(records),
        "build_script_dep_info_count": build_count,
        "crate_dependency_dep_info_count": deps_count,
        "changed_test_path_hits": [],
        "changed_test_basename_hits": [],
        "all_release_dep_info_excludes_changed_test": True,
        "files": records,
    }


def binary_binding(kind: str, binary_relative: str) -> dict[str, Any]:
    build = PACKET / "builds" / f"candidate-{kind}-v3"
    result_path = build / "result.json"
    terminal_path = build / "terminal.json"
    binary_receipt_path = build / "binary.json"
    result = json_no_duplicates(result_path)
    terminal = json_no_duplicates(terminal_path)
    binary_receipt = json_no_duplicates(binary_receipt_path)
    require(result.get("status") == "passed", f"{kind} v3 result is not passed")
    require(terminal.get("returncode") == 0 and terminal.get("timed_out") is False,
            f"{kind} v3 terminal is not successful")
    require(terminal.get("cleanup", {}).get("process_group_clear") is True
            and terminal.get("cleanup", {}).get("reaped") is True,
            f"{kind} v3 terminal custody is incomplete")
    binary = PACKET / binary_relative
    actual_size, actual_sha = sha256_file(binary)
    expected = result.get("binary", {})
    require(expected.get("status") == "present"
            and expected.get("bytes") == actual_size
            and expected.get("sha256") == actual_sha,
            f"{kind} v3 result binary binding mismatch")
    require(binary_receipt.get("binary", {}).get("bytes") == actual_size
            and binary_receipt.get("binary", {}).get("sha256") == actual_sha,
            f"{kind} v3 binary receipt mismatch")
    return {
        "kind": kind,
        "binary": descriptor(binary),
        "result": descriptor(result_path),
        "terminal": descriptor(terminal_path),
        "binary_receipt": descriptor(binary_receipt_path),
        "status": result["status"],
        "terminal_returncode": terminal["returncode"],
        "process_group_clear": terminal["cleanup"]["process_group_clear"],
        "reaped": terminal["cleanup"]["reaped"],
    }


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite retained receipt: {path}")
    payload = (json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False) + "\n").encode()
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as output:
        output.write(payload)
        output.flush()
        os.fsync(output.fileno())
    directory = os.open(path.parent, os.O_RDONLY | os.O_DIRECTORY)
    try:
        os.fsync(directory)
    finally:
        os.close(directory)


def main() -> int:
    require(not RECEIPT.exists(), f"receipt already exists: {RECEIPT}")
    v5, v5_summary = source_census(SOURCE_V5)
    v6, v6_summary = source_census(SOURCE_V6)
    compiled, compiled_summary = source_census(COMPILED_SOURCE)
    require({key: v5_summary[key] for key in EXPECTED_V5} == EXPECTED_V5,
            f"v5 census changed: {v5_summary}")
    require({key: v6_summary[key] for key in EXPECTED_V6} == EXPECTED_V6,
            f"v6 census changed: {v6_summary}")
    compare = compare_sources(v5, v6)
    compiled_compare = compare_compiled_source(v5, compiled)

    manifest = json_no_duplicates(SOURCE_MANIFEST)
    freeze = json_no_duplicates(SOURCE_FREEZE)
    require(manifest.get("source_root") == COMPILED_SOURCE.name
            and manifest.get("source_tree_digest") == EXPECTED_V5["tree_digest"]
            and manifest.get("file_count") == EXPECTED_V5["file_count"]
            and manifest.get("bytes") == EXPECTED_V5["bytes"],
            "v3 source manifest does not bind compiled v5 source")
    require(freeze.get("source_root") == COMPILED_SOURCE.name
            and freeze.get("source_tree_digest") == EXPECTED_V5["tree_digest"]
            and freeze.get("file_count") == EXPECTED_V5["file_count"]
            and freeze.get("source_bytes") == EXPECTED_V5["bytes"],
            "v3 source freeze does not bind compiled v5 source")

    binaries = [
        binary_binding("native", "target-candidate-native-v3/release/litchi-perf-baseline"),
        binary_binding("allocator", "target-candidate-allocator-v3/release/litchi-perf-baseline-alloc"),
    ]
    dep_infos = [
        dep_info_scan(PACKET / "target-candidate-native-v3"),
        dep_info_scan(PACKET / "target-candidate-allocator-v3"),
    ]
    require(all(item["all_release_dep_info_excludes_changed_test"] for item in dep_infos),
            "a v3 release dep-info file names the v6-only test")

    receipt = {
        "schema": "litchi.performance.0948.source-v5-v6-reuse-bridge.v1",
        "status": "passed",
        "created_utc": _datetime.datetime.now(_datetime.timezone.utc).isoformat(),
        "changed_test_file": CHANGED_PATH,
        "source_v5": v5_summary,
        "source_v6": v6_summary,
        "source_v5_to_v6": compare,
        "compiled_source": compiled_summary,
        "compiled_source_matches_source_v5": compiled_compare,
        "compiled_source_binding": {
            "source_manifest": descriptor(SOURCE_MANIFEST),
            "source_freeze": descriptor(SOURCE_FREEZE),
            "build_revision": COMPILED_SOURCE.name,
            "source_v6_is_not_compiled": True,
        },
        "release_dep_info": dep_infos,
        "binaries": binaries,
        "v3_binary_preservation": descriptor(PRESERVATION),
        "reuse_decision": {
            "reuse_candidate_build_v3": True,
            "reason": "candidate-source-v6 changes only an XML-minifier test lint expression; v3 release dep-info and build-script input records exclude that test file",
            "new_build_required_for_v6_test_only_change": False,
            "measurement_source_binding_remains": "candidate-build-v3",
        },
    }
    write_once(RECEIPT, receipt)
    size, digest = sha256_file(RECEIPT)
    print(json.dumps({
        "status": receipt["status"],
        "receipt": RECEIPT.relative_to(PACKET).as_posix(),
        "bytes": size,
        "sha256": digest,
        "v5": v5_summary,
        "v6": v6_summary,
        "changed": compare["changed_files"],
        "dep_info": [
            {k: item[k] for k in ("target_root", "dep_info_count", "build_script_dep_info_count", "crate_dependency_dep_info_count", "all_release_dep_info_excludes_changed_test")}
            for item in dep_infos
        ],
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, UnicodeError, ValueError, BridgeError) as error:
        print(f"v5/v6 reuse bridge refused: {error}", file=sys.stderr)
        raise SystemExit(2)
