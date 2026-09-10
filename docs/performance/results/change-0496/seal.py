#!/usr/bin/env python3
"""Create and verify the immutable 0496 phase-evidence seal.

The seal is deliberately the last read-only evidence step.  It imports the
frozen 0496 capture driver without executing its receipt-writing commands,
replays the formal collection, recomputes ``analysis/formal1.json``, checks
the existing verification and cleanup receipts, and records hashes for every
regular evidence file.  ``seal.json`` is excluded from its own inventory;
transient ``__pycache__`` directories are ignored and bytecode elsewhere is
rejected.  The retained four binaries, source manifests, build gates, current
primary Rust sources, and accepted ADR hashes are all bound by the manifest.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import stat
import sys
from typing import Any, Mapping


sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0496")
SEAL_NAME = "seal.json"
SCHEMA = "docx-phase-diagnostic-evidence-seal-v1"
VERIFICATION_SCHEMA = "docx-phase-diagnostic-seal-verification-v1"
VERSION = 1
FORMAL_ATTEMPT = "formal1"
BYTECODE_SUFFIXES = frozenset({".pyc", ".pyo"})
EXCLUDED = {
    "self": SEAL_NAME,
    "transient": "__pycache__/**",
    "bytecode": "reject outside transient directories",
}
ADR_REFRESH = "adr-refresh.json"
PROTECTED_PRIMARY = "protected-primary.json"
HARNESS_SOURCE = "harness-source.json"
PRODUCTION_SOURCE = "builds/after-final1-source.json"

# These are the four measured harness files and the production source files
# used by that harness.  The protected-primary manifest is historical cleanup
# evidence; unrelated concurrent primary files never become seal bindings.
MEASURED_HARNESS_PATHS = (
    "tools/perf-baseline/src/docx_managed_edit.rs",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/main.rs",
    "tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs",
)
SELECTED_PRODUCTION_PATHS = (
    "crates/litchi-docx/src/source_backed.rs",
    "crates/litchi-opc/src/source_backed.rs",
    "crates/litchi-opc/src/source_backed/read_ahead.rs",
    "crates/litchi-opc/src/source_backed/splice.rs",
)
PRIMARY_RUST_PATHS = MEASURED_HARNESS_PATHS + SELECTED_PRODUCTION_PATHS


class SealError(RuntimeError):
    """A malformed, changed, incomplete, or unsafe evidence binding."""


def fail(message: str) -> None:
    raise SealError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _finite(value: Any, label: str = "json") -> None:
    if isinstance(value, float):
        require(value == value and abs(value) != float("inf"), f"{label}: non-finite number")
    elif isinstance(value, Mapping):
        for key, child in value.items():
            require(isinstance(key, str), f"{label}: non-string object key")
            _finite(child, f"{label}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{label}[{index}]")


def _sha(path: Path) -> str:
    try:
        state = path.lstat()
    except OSError as error:
        fail(f"cannot stat {path}: {error}")
    require(stat.S_ISREG(state.st_mode), f"regular file required: {path}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot read {path}: {error}")
    return digest.hexdigest()


def _meta(path: Path, label: str) -> dict[str, int | str]:
    try:
        state = path.lstat()
    except OSError as error:
        fail(f"{label}: cannot stat {path}: {error}")
    require(stat.S_ISREG(state.st_mode), f"{label}: regular file required: {path}")
    return {"bytes": state.st_size, "sha256": _sha(path)}


def _read_json(path: Path, label: str) -> Any:
    _meta(path, label)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")
    _finite(value, label)
    return value


def _no_symlink_components(path: Path, label: str) -> None:
    absolute = Path(os.path.abspath(path))
    current = Path(absolute.anchor)
    for component in absolute.parts[1:]:
        current /= component
        try:
            state = current.lstat()
        except FileNotFoundError:
            continue
        except OSError as error:
            fail(f"{label}: cannot inspect {current}: {error}")
        require(not stat.S_ISLNK(state.st_mode), f"{label}: symlink component: {current}")


def _rooted(value: Path) -> Path:
    _no_symlink_components(value, "evidence root")
    try:
        state = value.lstat()
    except OSError as error:
        fail(f"evidence root: cannot stat {value}: {error}")
    require(stat.S_ISDIR(state.st_mode), f"evidence root: directory required: {value}")
    try:
        root = value.resolve(strict=True)
    except OSError as error:
        fail(f"evidence root: cannot resolve {value}: {error}")
    require(root == value.absolute(), f"evidence root is not canonical: {value}")
    return root


def _inside(root: Path, value: str | Path, label: str) -> tuple[Path, str]:
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = root / candidate
    _no_symlink_components(candidate, label)
    try:
        resolved = candidate.resolve(strict=True)
        relative = resolved.relative_to(root).as_posix()
    except (OSError, ValueError) as error:
        fail(f"{label}: path is missing or outside evidence root: {candidate} ({error})")
    _meta(resolved, label)
    return resolved, relative


def _descriptor(root: Path, path: Path, label: str) -> dict[str, int | str]:
    resolved, relative = _inside(root, path, label)
    return {"path": relative, **_meta(resolved, label)}


def _timestamp(value: Any, label: str) -> None:
    require(isinstance(value, str) and value, f"{label}: timestamp missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        fail(f"{label}: malformed timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label}: timestamp has no timezone")


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat(timespec="microseconds").replace("+00:00", "Z")


def _inventory(root: Path) -> dict[str, dict[str, int | str]]:
    """Hash all regular evidence files, excluding only seal output/bytecode."""

    result: dict[str, dict[str, int | str]] = {}

    def visit(directory: Path, relative: str) -> None:
        try:
            children = sorted(directory.iterdir(), key=lambda item: item.name)
        except OSError as error:
            fail(f"cannot enumerate evidence directory {directory}: {error}")
        for child in children:
            child_relative = f"{relative}/{child.name}" if relative else child.name
            try:
                state = child.lstat()
            except OSError as error:
                fail(f"cannot stat evidence path {child}: {error}")
            require(not stat.S_ISLNK(state.st_mode), f"evidence contains symlink: {child_relative}")
            if stat.S_ISDIR(state.st_mode):
                if child.name == "__pycache__":
                    continue
                visit(child, child_relative)
                continue
            require(stat.S_ISREG(state.st_mode), f"evidence contains special path: {child_relative}")
            if child_relative == SEAL_NAME:
                continue
            require(child.suffix.lower() not in BYTECODE_SUFFIXES, f"bytecode is not allowed: {child_relative}")
            result[child_relative] = _meta(child, f"evidence file {child_relative}")

    visit(root, "")
    return result


def _load_capture(root: Path):
    path = root / "capture.py"
    _meta(path, "capture driver")
    module_name = "change0496_capture_for_seal"
    spec = importlib.util.spec_from_file_location(module_name, path)
    require(spec is not None and spec.loader is not None, "cannot load capture driver")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except (OSError, ImportError, KeyError, TypeError, ValueError, RuntimeError) as error:
        fail(f"capture driver import failed: {error}")
    return module


def _retained_tree(capture: Any, builds: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    """Require exactly the four authenticated binaries and their parents."""

    temporary = Path(capture.TEMP)
    _no_symlink_components(temporary, "retained temporary root")
    try:
        state = temporary.lstat()
    except OSError as error:
        fail(f"retained temporary root is missing: {error}")
    require(stat.S_ISDIR(state.st_mode), "retained temporary root is not a directory")
    expected: dict[str, Path] = {}
    binary_names = {"normal": "litchi-perf-baseline", "allocator": "litchi-perf-baseline-alloc"}
    expected_keys = {f"{phase}/{role}" for phase in ("before", "after") for role in ("normal", "allocator")}
    require(set(builds) == expected_keys, "retained build inventory differs")
    for key, build in builds.items():
        require(key in {"before/normal", "before/allocator", "after/normal", "after/allocator"}, f"unknown retained build: {key}")
        phase, role = key.split("/", 1)
        binary = build.get("binary")
        require(isinstance(binary, Mapping), f"{key}: binary binding is missing")
        raw = binary.get("path")
        require(isinstance(raw, str) and raw, f"{key}: binary path is missing")
        path = Path(raw)
        expected_path = temporary / "retained" / phase / role / binary_names[role]
        _no_symlink_components(expected_path, f"{key}: retained binary")
        require(path.resolve(strict=False) == expected_path.absolute(), f"{key}: retained binary escaped its role directory")
        actual = _meta(expected_path, f"{key}: retained binary")
        require(dict(binary) == {"path": str(expected_path), **actual}, f"{key}: retained binary hash changed")
        try:
            mode = expected_path.stat().st_mode
        except OSError as error:
            fail(f"{key}: cannot inspect executable permission: {error}")
        require(bool(mode & 0o111) and os.access(expected_path, os.X_OK), f"{key}: retained binary is not executable")
        expected[key] = expected_path

    files: set[str] = set()
    directories: set[str] = set()

    def walk(directory: Path) -> None:
        try:
            entries = sorted(directory.iterdir(), key=lambda item: item.name)
        except OSError as error:
            fail(f"cannot enumerate retained tree {directory}: {error}")
        for entry in entries:
            _no_symlink_components(entry, "retained tree")
            try:
                entry_state = entry.lstat()
            except OSError as error:
                fail(f"cannot stat retained path {entry}: {error}")
            relative = str(entry.relative_to(temporary))
            if stat.S_ISDIR(entry_state.st_mode):
                directories.add(relative)
                walk(entry)
            elif stat.S_ISREG(entry_state.st_mode):
                files.add(relative)
            else:
                fail(f"retained tree contains special path: {relative}")

    walk(temporary)
    expected_files = {str(path.relative_to(temporary)) for path in expected.values()}
    expected_dirs: set[str] = set()
    for path in expected.values():
        for parent in path.parents:
            if parent == temporary:
                break
            try:
                relative_parent = parent.relative_to(temporary)
            except ValueError:
                break
            expected_dirs.add(str(relative_parent))
    require(files == expected_files, f"retained tree has unexplained files: {sorted(files ^ expected_files)}")
    require(directories == expected_dirs, f"retained tree has unexplained directories: {sorted(directories ^ expected_dirs)}")
    return {
        key: {"path": str(path), **_meta(path, f"{key}: retained binary")}
        for key, path in sorted(expected.items())
    }


def _adr_binding(root: Path = ROOT, repo: Path = REPO) -> dict[str, Any]:
    path = root / ADR_REFRESH
    value = _read_json(path, ADR_REFRESH)
    require(isinstance(value, Mapping), "adr-refresh.json must be an object")
    files = value.get("files")
    require(isinstance(files, Mapping) and files, "ADR hash inventory is empty")
    actual: dict[str, str] = {}
    for name, digest in sorted(files.items()):
        require(isinstance(name, str) and name.startswith("docs/adr/") and ".." not in Path(name).parts, f"unsafe ADR path: {name}")
        require(isinstance(digest, str) and len(digest) == 64 and all(char in "0123456789abcdef" for char in digest), f"malformed ADR hash: {name}")
        source = repo / name
        _no_symlink_components(source, f"ADR {name}")
        actual_digest = _sha(source)
        require(actual_digest == digest, f"ADR {name}: content changed")
        actual[name] = actual_digest
    return {"refresh": _descriptor(root, path, "ADR refresh"), "files": actual}


def _source_record(value: Any, label: str) -> dict[str, int | str]:
    require(isinstance(value, Mapping), f"{label}: source record is missing")
    require(set(value) == {"path", "bytes", "sha256"}, f"{label}: source record fields differ")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and Path(raw_path).is_absolute(), f"{label}.path: absolute path required")
    raw_bytes = value.get("bytes")
    require(type(raw_bytes) is int and raw_bytes >= 0, f"{label}.bytes: unsigned integer required")
    digest = value.get("sha256")
    require(isinstance(digest, str) and len(digest) == 64 and all(char in "0123456789abcdef" for char in digest), f"{label}.sha256: malformed hash")
    return {"path": raw_path, "bytes": raw_bytes, "sha256": digest}


def _primary_binding(root: Path = ROOT, repo: Path = REPO, paths: tuple[str, ...] | None = None) -> dict[str, Any]:
    manifest_path = root / PROTECTED_PRIMARY
    value = _read_json(manifest_path, PROTECTED_PRIMARY)
    require(isinstance(value, Mapping) and value, "protected primary manifest is empty")
    for name, digest in value.items():
        require(isinstance(name, str) and not Path(name).is_absolute() and ".." not in Path(name).parts, f"unsafe protected primary path: {name}")
        require(isinstance(digest, str) and len(digest) == 64 and all(char in "0123456789abcdef" for char in digest), f"malformed protected primary hash: {name}")
    selected = set(PRIMARY_RUST_PATHS if paths is None else paths)
    result: dict[str, dict[str, int | str]] = {}
    for name in sorted(selected):
        require(name.endswith(".rs"), f"primary binding is not Rust source: {name}")
        source = repo / name
        _no_symlink_components(source, f"primary Rust source {name}")
        meta = _meta(source, f"primary Rust source {name}")
        result[name] = meta
    binding: dict[str, Any] = {
        "historical_manifest": _descriptor(root, manifest_path, "protected primary manifest"),
        "files": result,
    }
    if paths is not None:
        return binding

    harness_path = root / HARNESS_SOURCE
    harness_value = _read_json(harness_path, HARNESS_SOURCE)
    require(isinstance(harness_value, Mapping), "harness source inventory is malformed")
    require(set(harness_value) == set(MEASURED_HARNESS_PATHS), "harness source inventory differs")
    harness_files: dict[str, dict[str, int | str]] = {}
    for name in MEASURED_HARNESS_PATHS:
        digest = harness_value.get(name)
        require(isinstance(digest, str) and len(digest) == 64 and all(char in "0123456789abcdef" for char in digest), f"harness source hash is malformed: {name}")
        require(result[name]["sha256"] == digest, f"measured harness source changed: {name}")
        harness_files[name] = result[name]

    production_path = root / PRODUCTION_SOURCE
    production_value = _read_json(production_path, PRODUCTION_SOURCE)
    require(isinstance(production_value, Mapping), "production source manifest is malformed")
    production_files: dict[str, dict[str, int | str]] = {}
    for name in SELECTED_PRODUCTION_PATHS:
        recorded = _source_record(production_value.get(name), f"{PRODUCTION_SOURCE}.{name}")
        current = result[name]
        require(recorded["bytes"] == current["bytes"] and recorded["sha256"] == current["sha256"], f"production source changed: {name}")
        production_files[name] = {"bytes": recorded["bytes"], "sha256": recorded["sha256"]}
    binding["harness_source"] = _descriptor(root, harness_path, "harness source inventory")
    binding["harness_files"] = harness_files
    binding["production_source"] = _descriptor(root, production_path, "production source manifest")
    binding["production_files"] = production_files
    return binding


def _cleanup_binding(root: Path, capture: Any) -> dict[str, Any]:
    path = root / "cleanup.json"
    value = _read_json(path, "cleanup receipt")
    require(isinstance(value, Mapping), "cleanup receipt must be an object")
    require(set(value) == {
        "finished_ns", "free_after", "free_before", "patches_reverse_checked",
        "protected_files_unchanged", "removed", "retained_builds", "schema",
        "scope", "source_manifests_verified", "status",
    }, "cleanup receipt fields differ")
    require(value.get("schema") == "docx-phase-cleanup-v1" and value.get("status") == "pass", "cleanup receipt is not passing")
    for field in ("source_manifests_verified", "patches_reverse_checked", "protected_files_unchanged"):
        require(value.get(field) is True, f"cleanup receipt does not prove {field}")
    removed = value.get("removed")
    require(isinstance(removed, list), "cleanup receipt removed inventory is missing")
    require(type(value.get("finished_ns")) is int and value["finished_ns"] > 0, "cleanup receipt completion time is malformed")
    for field in ("free_before", "free_after"):
        require(type(value.get(field)) is int and value[field] >= 0, f"cleanup receipt {field} is malformed")
    require(isinstance(value.get("scope"), str) and bool(value["scope"]), "cleanup receipt scope is missing")
    require(value.get("retained_builds") == [
        "before/normal", "before/allocator", "after/normal", "after/allocator",
    ], "cleanup receipt retained build inventory differs")
    temporary = Path(capture.TEMP)
    allowed = {temporary / name for name in ("before", "after", "target", "tmp", "projections", "runs")}
    for index, item in enumerate(removed):
        require(isinstance(item, Mapping), f"cleanup receipt removed[{index}] is malformed")
        require(set(item) == {"path", "allocated_bytes", "files"}, f"cleanup receipt removed[{index}] fields differ")
        raw_path = item.get("path")
        require(isinstance(raw_path, str), f"cleanup receipt removed[{index}].path is malformed")
        removed_path = Path(raw_path)
        _no_symlink_components(removed_path, f"cleanup receipt removed[{index}]")
        require(removed_path.absolute() in allowed, f"cleanup receipt removed path escaped temporary root: {raw_path}")
        require(type(item.get("allocated_bytes")) is int and item["allocated_bytes"] >= 0,
                f"cleanup receipt removed[{index}].allocated_bytes is malformed")
        require(type(item.get("files")) is int and item["files"] >= 0,
                f"cleanup receipt removed[{index}].files is malformed")
    require(len({item["path"] for item in removed}) == len(removed), "cleanup receipt removed inventory has duplicates")
    require(not (temporary / "target").exists(), "Cargo target remains after cleanup")
    require(not (temporary / "before").exists() and not (temporary / "after").exists(), "source checkout remains after cleanup")
    return {
        "path": "cleanup.json",
        **_meta(path, "cleanup receipt"),
        "schema": value["schema"],
        "status": value["status"],
        "removed": [dict(item) for item in removed],
        "source_manifests_verified": True,
        "patches_reverse_checked": True,
        "protected_files_unchanged": True,
    }


def _remove_empty_projection_parent(capture: Any) -> None:
    """Remove only the empty replay directory created by report validation."""

    projection_root = Path(capture.TEMP) / "projections"
    try:
        state = projection_root.lstat()
    except FileNotFoundError:
        return
    except OSError as error:
        fail(f"projection scratch cannot be inspected: {error}")
    require(stat.S_ISDIR(state.st_mode), "projection scratch is not a directory")
    _no_symlink_components(projection_root, "projection scratch")
    try:
        children = tuple(projection_root.iterdir())
    except OSError as error:
        fail(f"projection scratch cannot be enumerated: {error}")
    require(not children, "projection scratch contains unexplained files")
    try:
        projection_root.rmdir()
    except OSError as error:
        fail(f"projection scratch cannot be removed: {error}")


def _verification_binding(root: Path, protocol_hash: str, entries: list[Mapping[str, Any]]) -> dict[str, Any]:
    path = root / "verification" / f"{FORMAL_ATTEMPT}.json"
    value = _read_json(path, "formal verification receipt")
    require(isinstance(value, Mapping), "formal verification receipt must be an object")
    require(value.get("schema") == "docx-phase-diagnostic-verification-v1", "formal verification schema differs")
    require(value.get("version") == 1 and value.get("status") == "pass", "formal verification is not passing")
    require(value.get("attempt") == FORMAL_ATTEMPT, "formal verification attempt differs")
    require(value.get("children") == len(entries) == 32, "formal verification child count differs")
    require(value.get("samples") == len(entries) * 30 == 960, "formal verification sample count differs")
    protocol = value.get("protocol")
    require(isinstance(protocol, Mapping) and protocol.get("path") == "protocol.json" and protocol.get("sha256") == protocol_hash, "formal verification protocol binding differs")
    return {"path": "verification/formal1.json", **_meta(path, "formal verification receipt"), "receipt": dict(value)}


def _analysis_binding(root: Path, capture: Any, protocol: dict[str, Any], protocol_hash: str, builds: dict[str, dict[str, Any]]) -> tuple[dict[str, Any], list[Mapping[str, Any]]]:
    try:
        entries = capture._collect(FORMAL_ATTEMPT, protocol, builds)
        require(len(entries) == 32, "formal capture count differs")
        expected = capture.analyze_data(entries, protocol={"path": "protocol.json", "sha256": protocol_hash})
        path = root / "analysis" / f"{FORMAL_ATTEMPT}.json"
        actual = _read_json(path, "formal analysis")
        require(actual == expected, "formal analysis does not recompute from retained reports")
        require(actual.get("schema") == capture.ANALYSIS_SCHEMA and actual.get("child_count") == 32 and actual.get("sample_count") == 960, "formal analysis inventory differs")
        return ({"path": "analysis/formal1.json", **_meta(path, "formal analysis")}, entries)
    finally:
        _remove_empty_projection_parent(capture)


def _evidence(root: Path) -> dict[str, Any]:
    capture = _load_capture(root)
    try:
        protocol, protocol_hash, builds = capture._load_protocol()
        require(isinstance(protocol, Mapping) and isinstance(builds, Mapping), "capture protocol/build inventory is malformed")
        retained = _retained_tree(capture, builds)
        analysis, entries = _analysis_binding(root, capture, dict(protocol), protocol_hash, dict(builds))
        verification = _verification_binding(root, protocol_hash, entries)
        cleanup = _cleanup_binding(root, capture)
    except (OSError, ImportError, KeyError, TypeError, ValueError, RuntimeError) as error:
        fail(f"formal evidence validation failed: {error}")
    protocol_descriptor = _descriptor(root, root / "protocol.json", "frozen protocol")
    require(protocol_descriptor["sha256"] == protocol_hash, "frozen protocol hash differs")
    builds_binding = capture._builds_binding(dict(builds))
    primary = _primary_binding(root, REPO)
    adr = _adr_binding(root, REPO)
    return {
        "protocol": {**protocol_descriptor, "sha256": protocol_hash},
        "builds": builds_binding,
        "retained_binaries": retained,
        "formal": {"attempt": FORMAL_ATTEMPT, "children": len(entries), "samples": len(entries) * 30},
        "analysis": analysis,
        "verification": verification,
        "cleanup": cleanup,
        "primary_rust": primary,
        "adr": adr,
    }


def _manifest(root: Path) -> dict[str, Any]:
    require(not (root / SEAL_NAME).exists(), f"refusing to replace existing seal: {root / SEAL_NAME}")
    evidence = _evidence(root)
    return {
        "schema": SCHEMA,
        "version": VERSION,
        "root": ".",
        "sealed_utc": _now(),
        "excluded": EXCLUDED,
        **evidence,
        "files": _inventory(root),
    }


def _write_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace existing seal: {path}")
    try:
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
    except OSError as error:
        fail(f"cannot write seal {path}: {error}")


def _verify_manifest(root: Path, manifest: Any) -> dict[str, Any]:
    require(isinstance(manifest, Mapping), "seal manifest must be an object")
    expected_keys = {
        "schema", "version", "root", "sealed_utc", "excluded", "protocol", "builds",
        "retained_binaries", "formal", "analysis", "verification", "cleanup",
        "primary_rust", "adr", "files",
    }
    require(set(manifest) == expected_keys, "seal manifest fields differ")
    require(manifest.get("schema") == SCHEMA and manifest.get("version") == VERSION, "seal schema/version differs")
    require(manifest.get("root") == "." and manifest.get("excluded") == EXCLUDED, "seal root/exclusion policy differs")
    _timestamp(manifest.get("sealed_utc"), "seal timestamp")
    require(manifest.get("files") == _inventory(root), "sealed evidence inventory or hash differs")
    current = _evidence(root)
    for key, value in current.items():
        require(manifest.get(key) == value, f"sealed {key} binding changed")
    return dict(current)


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("seal", "verify"), nargs="?", default="verify")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        root = _rooted(args.root)
        require(root == ROOT, "0496 seal is restricted to its fixed evidence root")
        seal_path = root / SEAL_NAME
        if args.command == "seal":
            value = _manifest(root)
            _write_exclusive(seal_path, value)
            print(json.dumps({"schema": SCHEMA, "status": "sealed", "files": len(value["files"]), "seal": {"path": SEAL_NAME, **_meta(seal_path, "seal manifest")}}, sort_keys=True))
            return 0
        manifest = _read_json(seal_path, "seal manifest")
        checked = _verify_manifest(root, manifest)
        print(json.dumps({
            "schema": VERIFICATION_SCHEMA,
            "version": VERSION,
            "status": "pass",
            "verified_utc": _now(),
            "files": len(manifest["files"]),
            "formal": checked["formal"],
            "seal": {"path": SEAL_NAME, **_meta(seal_path, "seal manifest")},
        }, sort_keys=True))
        return 0
    except (SealError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"seal.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
