#!/usr/bin/env python3
"""Remove only authenticated 0487 scratch artifacts after all jobs finish.

The command is intentionally fail-closed.  It first revalidates the retained
comparison, profile, build, fuzz, and replay-cleanup evidence, then archives
the temporary fuzz manifest inputs and removes only the 0487 scratch paths.
The default action is destructive for those exact paths; ``--dry-run`` runs
the same custody checks without creating the archive or removing anything.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import sys
from typing import Any, Mapping


# Cleanup itself must not create a bytecode file while it imports the retained
# verifier and support helpers.  Bytecode already present under ROOT is part
# of the explicitly removable inventory and is checked before deletion.
sys.dont_write_bytecode = True

from support import REPO, ROOT, TEMP, meta, now, read, sha, write  # noqa: E402
from verify_seal import SealError, _comparison_record, _fuzz_record  # noqa: E402


SCHEMA = "docx-opc-splice-audit-cleanup-v1"
VERSION = 1
FUZZ_ATTEMPT = "candidate1"
FUZZ_ROOT = ROOT / "fuzz-stream" / FUZZ_ATTEMPT
FUZZ_WORK = TEMP / f"fuzz-docx-stream-{FUZZ_ATTEMPT}"
FUZZ_BINARY_NAME = "source_backed_tail_append_stream"
FUZZ_INPUT_ARCHIVE = ROOT / "fuzz-build-inputs"
FUZZ_INPUT_NAMES = ("Cargo.toml", "Cargo.lock")
FUZZ_INPUT_ARCHIVE_NAMES = {
    "Cargo.toml": "Cargo.toml.txt",
    "Cargo.lock": "Cargo.lock.txt",
}
BUILD_ROLES = ("normal", "allocator")
SHARED_CACHES = (
    REPO / "target",
    REPO / "target" / "fuzz-asan",
    REPO / "tools" / "perf-baseline" / "target",
)
TERMINAL_HOST_OBSERVATIONS = (
    "before-start.json",
    "before-end.json",
    "after-start.json",
    "after-end.json",
)


class CleanupError(RuntimeError):
    """A custody or cleanup precondition failed."""


def fail(message: str) -> None:
    raise CleanupError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _regular_file(path: Path, label: str) -> Path:
    require(not path.is_symlink(), f"{label}: symlink is not allowed: {path}")
    require(path.is_file(), f"{label}: regular file is missing: {path}")
    return path


def _directory(path: Path, label: str) -> Path:
    require(not path.is_symlink(), f"{label}: symlink is not allowed: {path}")
    require(path.is_dir(), f"{label}: directory is missing: {path}")
    return path


def _json(path: Path, label: str) -> Any:
    _regular_file(path, label)
    try:
        return read(path)
    except (OSError, TypeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def _file_descriptor(path: Path, label: str, *, relative_to_root: bool = False) -> dict[str, Any]:
    _regular_file(path, label)
    value: dict[str, Any] = {"path": str(path.resolve()), **meta(path)}
    if relative_to_root:
        try:
            value["relative_path"] = path.resolve().relative_to(ROOT.resolve()).as_posix()
        except ValueError as error:
            fail(f"{label}: path is outside evidence root: {path} ({error})")
    return value


def _inventory(path: Path, label: str) -> dict[str, Any] | dict[str, int | str]:
    """Return a deterministic file inventory without following symlinks."""

    require(not path.is_symlink(), f"{label}: symlink is not allowed: {path}")
    if path.is_file():
        return meta(path)
    _directory(path, label)
    result: dict[str, Any] = {}
    for child in sorted(path.iterdir(), key=lambda item: item.name):
        require(not child.is_symlink(), f"{label}: symlink is not allowed: {child}")
        result[child.name] = _inventory(child, f"{label}/{child.name}")
    return result


def _remove(path: Path, label: str) -> None:
    """Remove one already-authenticated file or directory tree."""

    if path.is_symlink():
        fail(f"{label}: refusing to remove symlink: {path}")
    if not path.exists():
        return
    if path.is_dir():
        for child in sorted(path.iterdir(), key=lambda item: item.name, reverse=True):
            _remove(child, f"{label}/{child.name}")
        path.rmdir()
    else:
        path.unlink()


def _archive_file(source: Path, destination: Path, expected: Mapping[str, Any], label: str) -> dict[str, Any]:
    """Create an exclusive text-named copy, or verify an existing copy."""

    expected_meta = {key: expected[key] for key in ("bytes", "sha256")}
    if source.exists() or source.is_symlink():
        _regular_file(source, label)
        require(meta(source) == expected_meta, f"{label}: source metadata differs from prepared receipt")
    if destination.exists() or destination.is_symlink():
        _regular_file(destination, f"{label} archive")
        require(meta(destination) == expected_meta, f"{label} archive: existing bytes differ")
    else:
        destination.parent.mkdir(parents=True, exist_ok=True)
        try:
            with source.open("rb") as input_stream, destination.open("xb") as output_stream:
                while True:
                    chunk = input_stream.read(1024 * 1024)
                    if not chunk:
                        break
                    output_stream.write(chunk)
        except OSError as error:
            fail(f"{label}: cannot archive input: {error}")
        require(meta(destination) == expected_meta, f"{label} archive: copied bytes differ")
    return _file_descriptor(destination, f"{label} archive", relative_to_root=True)


def _verify_validation_terminal() -> dict[str, Any]:
    validation = _directory(ROOT / "validation", "validation receipts")
    unfinished: list[str] = []
    terminal: list[dict[str, Any]] = []
    for started in sorted(validation.glob("*.started.json")):
        _regular_file(started, f"validation start receipt {started.name}")
        result = validation / (started.name.removesuffix(".started.json") + ".json")
        if not result.is_file() or result.is_symlink():
            unfinished.append(started.name)
            continue
        value = _json(result, f"validation receipt {result.name}")
        require(isinstance(value, Mapping), f"validation receipt {result.name}: expected an object")
        require(isinstance(value.get("exit_code"), int) and bool(value.get("finished_utc")),
                f"validation receipt {result.name}: terminal status is missing")
        terminal.append(_file_descriptor(result, f"validation receipt {result.name}", relative_to_root=True))
    require(not unfinished, f"root validation jobs are not terminal: {unfinished}")
    for name in TERMINAL_HOST_OBSERVATIONS:
        value = _json(ROOT / "host-observations" / name, f"host observation {name}")
        require(isinstance(value, Mapping), f"host observation {name}: expected an object")
    return {"validation_receipts": terminal, "host_observations": list(TERMINAL_HOST_OBSERVATIONS)}


def _verify_build(role: str) -> dict[str, Any]:
    path = ROOT / f"build-{role}.json"
    value = _json(path, f"0487 {role} build")
    require(isinstance(value, Mapping), f"0487 {role} build: expected an object")
    require(value.get("schema") == "docx-replayable-tail-append-build-v1", f"0487 {role} build: schema differs")
    require(value.get("version") == 1 and value.get("role") == role and value.get("attempt") == "after1", f"0487 {role} build: identity differs")
    require(value.get("source_unchanged") is True and value.get("source_before") == value.get("source_after"), f"0487 {role} build: source custody failed")

    binaries: dict[str, Any] = {}
    # Cargo's shared output is overwritten by the next role's build. Only
    # the copied executable is a retained input used by captures/profiles.
    for field in ("binary",):
        record = value.get(field)
        require(isinstance(record, Mapping) and isinstance(record.get("path"), str), f"0487 {role} build: {field} binding is missing")
        binary_path = Path(record["path"])
        _regular_file(binary_path, f"0487 {role} {field}")
        actual = meta(binary_path)
        require(record.get("bytes") == actual["bytes"] and record.get("sha256") == actual["sha256"], f"0487 {role} {field}: binary metadata changed")
        require(record.get("executable") is True and os.access(binary_path, os.X_OK), f"0487 {role} {field}: executable binding failed")
        binaries[field] = {"path": str(binary_path.resolve()), **actual, "executable": True}

    gate = value.get("gate")
    require(isinstance(gate, Mapping) and isinstance(gate.get("path"), str), f"0487 {role} build: gate binding is missing")
    gate_path = Path(gate["path"])
    _regular_file(gate_path, f"0487 {role} build gate")
    require(gate.get("sha256") == sha(gate_path), f"0487 {role} build gate: hash changed")
    gate_value = _json(gate_path, f"0487 {role} build gate")
    require(isinstance(gate_value, Mapping) and gate_value.get("exit_code") == 0 and gate_value.get("source_unchanged") is True, f"0487 {role} build gate: command or source custody failed")

    return {
        "record": _file_descriptor(path, f"0487 {role} build record", relative_to_root=True),
        "binary": binaries["binary"],
        "original_binary_at_build_time": value["original_binary"],
        "gate": _file_descriptor(gate_path, f"0487 {role} build gate", relative_to_root=True),
    }


def _verify_fuzz() -> dict[str, Any]:
    smoke_path = FUZZ_ROOT / "smoke.json"
    _regular_file(smoke_path, "0487 fuzz smoke")
    try:
        checked = _fuzz_record(ROOT, smoke_path)
    except SealError as error:
        fail(f"0487 fuzz custody failed: {error}")
    smoke = _json(smoke_path, "0487 fuzz smoke")
    build_path = FUZZ_ROOT / "build.json"
    prepared_path = FUZZ_ROOT / "prepared.json"
    build = _json(build_path, "0487 fuzz build")
    prepared = _json(prepared_path, "0487 fuzz prepared")
    require(isinstance(build, Mapping) and isinstance(prepared, Mapping), "0487 fuzz receipts: expected objects")
    binary = checked["binary"]
    binary_path = Path(binary["path"])
    require(binary_path.resolve() == (FUZZ_WORK / FUZZ_BINARY_NAME).resolve(), "0487 fuzz binary: retained path differs from candidate work binary")
    require(os.access(binary_path, os.X_OK), "0487 fuzz binary: executable bit is missing")
    require(build.get("binary", {}).get("path") == smoke.get("binary", {}).get("path"), "0487 fuzz build/smoke binary paths differ")
    return {
        "smoke": _file_descriptor(smoke_path, "0487 fuzz smoke", relative_to_root=True),
        "build": _file_descriptor(build_path, "0487 fuzz build", relative_to_root=True),
        "prepared": _file_descriptor(prepared_path, "0487 fuzz prepared", relative_to_root=True),
        "binary": binary,
        "seed_count": len(prepared.get("seed_records", {})),
        "positive_case_count": checked["positive_case_count"],
        "run_count": checked["run_count"],
    }


def _verify_comparison() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    summary_path = ROOT / "comparison-summary.json"
    _regular_file(summary_path, "0487 comparison summary")
    try:
        checked = _comparison_record(ROOT, summary_path)
    except SealError as error:
        fail(f"0487 comparison custody failed: {error}")
    summary = _json(summary_path, "0487 comparison summary")
    require(isinstance(summary, Mapping), "0487 comparison summary: expected an object")
    processes = summary.get("processes")
    require(isinstance(processes, list) and len(processes) == 144, "0487 comparison summary: formal process inventory is incomplete")
    expected: list[dict[str, Any]] = []
    for row in processes:
        require(isinstance(row, Mapping), "0487 comparison process row: expected an object")
        if row.get("route") != "file_store":
            continue
        directory = Path(str(row.get("directory")))
        try:
            directory.resolve().relative_to((ROOT / "captures").resolve())
        except ValueError as error:
            fail(f"0487 file-store row escapes captures: {directory} ({error})")
        receipt_path = directory / "replay-cleanup.json"
        value = _json(receipt_path, f"file-store replay cleanup {receipt_path}")
        require(value.get("schema") == "docx-opc-splice-audit-replay-cleanup-v1", f"{receipt_path}: replay cleanup schema differs")
        require(value.get("removed") is True and value.get("empty_before_removal") is True and value.get("entries_before_removal") == [], f"{receipt_path}: replay directory was not proven empty")
        replay_path = Path(str(value.get("path")))
        require(replay_path.resolve() == (directory / "replay").resolve(), f"{receipt_path}: replay path differs")
        require(not replay_path.exists(), f"{receipt_path}: replay directory still exists")
        expected.append({
            "path": receipt_path.relative_to(ROOT).as_posix(),
            "sha256": sha(receipt_path),
            "bytes": meta(receipt_path)["bytes"],
        })
    receipt_paths = sorted(
        path.relative_to(ROOT).as_posix()
        for path in (ROOT / "captures").rglob("replay-cleanup.json")
        if path.is_file() and not path.is_symlink()
    )
    expected_paths = sorted(item["path"] for item in expected)
    require(receipt_paths == expected_paths and expected_paths, "0487 file-store replay cleanup receipt inventory differs")
    replay_dirs = [path for path in (ROOT / "captures").rglob("replay") if path.is_dir() and not path.is_symlink()]
    require(not replay_dirs, f"0487 file-store replay directories remain: {replay_dirs}")
    return checked, expected


def _verify_profiles() -> dict[str, Any]:
    summary_path = ROOT / "profiles" / "profiles1-summary.json"
    markdown_path = ROOT / "profiles" / "profiles1-summary.md"
    summary = _json(summary_path, "0487 profile summary")
    require(isinstance(summary, Mapping) and summary.get("status") == "complete" and summary.get("validation_status") == "pass", "0487 profile summary is not terminal and passing")
    _regular_file(markdown_path, "0487 profile Markdown summary")
    return {
        "summary": _file_descriptor(summary_path, "0487 profile summary", relative_to_root=True),
        "markdown": _file_descriptor(markdown_path, "0487 profile Markdown summary", relative_to_root=True),
        "inventory": summary.get("inventory"),
    }


def _verify_caches() -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for path in SHARED_CACHES:
        _directory(path, f"shared Cargo cache {path}")
        result.append({"path": str(path.resolve()), "retained": True})
    return result


def _verify_temp_and_bytecode(fuzz: Mapping[str, Any]) -> dict[str, Any]:
    _directory(TEMP, "0487 temporary root")
    removed: dict[str, Any] = {}
    for path, label in ((TEMP / "draft", "0487 draft scratch"), (TEMP / "test-tmp", "0487 test scratch")):
        if path.exists() or path.is_symlink():
            inventory = _inventory(path, label)
            if path.name == "test-tmp":
                require(inventory == {}, f"{label}: expected an empty directory")
            removed[str(path)] = inventory

    require(FUZZ_WORK.is_dir() and not FUZZ_WORK.is_symlink(), f"0487 fuzz work directory is missing: {FUZZ_WORK}")
    expected_binary = Path(str(fuzz["binary"]["path"])).resolve()
    require(expected_binary == (FUZZ_WORK / FUZZ_BINARY_NAME).resolve(), "0487 fuzz work binary path differs")
    direct_names = {item.name for item in FUZZ_WORK.iterdir()}
    allowed_names = {FUZZ_BINARY_NAME, "Cargo.toml", "Cargo.lock", "corpus", "artifacts", "build-inputs"}
    require(direct_names <= allowed_names, f"0487 fuzz work contains an unauthorized path: {sorted(direct_names - allowed_names)}")
    require(expected_binary.is_file() and not expected_binary.is_symlink(), "0487 fuzz work binary is missing")
    prepared = _json(FUZZ_ROOT / "prepared.json", "0487 fuzz prepared")
    require(isinstance(prepared, Mapping), "0487 fuzz prepared: expected an object")
    for source_name in FUZZ_INPUT_NAMES:
        source = FUZZ_WORK / source_name
        expected = prepared.get("manifest" if source_name == "Cargo.toml" else "lock")
        require(isinstance(expected, Mapping), f"0487 fuzz prepared: {source_name} metadata is missing")
        archive = FUZZ_INPUT_ARCHIVE / FUZZ_INPUT_ARCHIVE_NAMES[source_name]
        if source.exists() or source.is_symlink():
            _regular_file(source, f"0487 fuzz work {source_name}")
            require(meta(source) == {key: expected[key] for key in ("bytes", "sha256")}, f"0487 fuzz work {source_name}: metadata differs")
        else:
            _regular_file(archive, f"0487 archived fuzz {source_name}")
            require(meta(archive) == {key: expected[key] for key in ("bytes", "sha256")}, f"0487 archived fuzz {source_name}: metadata differs")

    for name in ("corpus", "artifacts", "build-inputs"):
        path = FUZZ_WORK / name
        if path.exists() or path.is_symlink():
            removed[str(path)] = _inventory(path, f"0487 fuzz scratch {name}")
    for name in FUZZ_INPUT_NAMES:
        path = FUZZ_WORK / name
        if path.exists() or path.is_symlink():
            removed[str(path)] = _inventory(path, f"0487 fuzz work {name}")

    bytecode: list[dict[str, Any]] = []
    bytecode_dirs: set[Path] = set()
    for path in sorted(ROOT.rglob("*"), key=lambda item: item.as_posix()):
        if path.is_symlink():
            continue
        if path.is_file() and ("__pycache__" in path.parts or path.suffix.lower() in {".pyc", ".pyo"}):
            bytecode.append({"path": path.relative_to(ROOT).as_posix(), **meta(path)})
            for parent in path.parents:
                if parent.name == "__pycache__":
                    bytecode_dirs.add(parent)
                    break
    for directory in sorted(bytecode_dirs):
        contents = _inventory(directory, f"bytecode directory {directory}")
        require(all(Path(name).suffix.lower() in {".pyc", ".pyo"} for name in contents), f"bytecode directory contains a non-bytecode file: {directory}")
    return {"removed": removed, "bytecode": bytecode, "bytecode_dirs": [str(path.relative_to(ROOT)) for path in sorted(bytecode_dirs)]}


def _archive_inputs() -> dict[str, Any]:
    prepared = _json(FUZZ_ROOT / "prepared.json", "0487 fuzz prepared")
    _directory(FUZZ_WORK, "0487 fuzz work")
    FUZZ_INPUT_ARCHIVE.mkdir(parents=True, exist_ok=True)
    require(not FUZZ_INPUT_ARCHIVE.is_symlink(), "0487 fuzz input archive: symlink is not allowed")
    current = {path.name for path in FUZZ_INPUT_ARCHIVE.iterdir()}
    require(current <= set(FUZZ_INPUT_ARCHIVE_NAMES.values()), f"0487 fuzz input archive contains unauthorized files: {sorted(current - set(FUZZ_INPUT_ARCHIVE_NAMES.values()))}")
    records: dict[str, Any] = {}
    for source_name in FUZZ_INPUT_NAMES:
        expected = prepared.get("manifest" if source_name == "Cargo.toml" else "lock")
        require(isinstance(expected, Mapping), f"0487 fuzz prepared: {source_name} metadata is missing")
        source = FUZZ_WORK / source_name
        destination = FUZZ_INPUT_ARCHIVE / FUZZ_INPUT_ARCHIVE_NAMES[source_name]
        records[FUZZ_INPUT_ARCHIVE_NAMES[source_name]] = _archive_file(source, destination, expected, f"0487 fuzz {source_name}")
    require({path.name for path in FUZZ_INPUT_ARCHIVE.iterdir()} == set(FUZZ_INPUT_ARCHIVE_NAMES.values()), "0487 fuzz input archive inventory differs")
    return records


def _remove_authenticated(plan: Mapping[str, Any]) -> None:
    removed = plan["temporary"]["removed"]
    for path_value in removed:
        path = Path(path_value)
        # The fuzz work binary and the three measurement binaries are never in
        # this map, so this loop cannot remove retained executables.
        _remove(path, f"remove {path}")
    for path_value in plan["temporary"]["bytecode"]:
        _remove(ROOT / path_value["path"], f"remove bytecode {path_value['path']}")
    for directory in sorted((ROOT / value for value in plan["temporary"]["bytecode_dirs"]), key=lambda item: len(item.parts), reverse=True):
        if directory.exists() and not directory.is_symlink():
            directory.rmdir()


def _plan() -> dict[str, Any]:
    cleanup_path = ROOT / "cleanup.json"
    require(not cleanup_path.exists(), f"refusing to replace existing cleanup receipt: {cleanup_path}")
    terminal = _verify_validation_terminal()
    comparison, replay_receipts = _verify_comparison()
    profiles = _verify_profiles()
    builds = {role: _verify_build(role) for role in BUILD_ROLES}
    fuzz = _verify_fuzz()
    caches = _verify_caches()
    temporary = _verify_temp_and_bytecode(fuzz)
    return {
        "terminal": terminal,
        "comparison": comparison,
        "replay_receipts": replay_receipts,
        "profiles": profiles,
        "builds": builds,
        "fuzz": fuzz,
        "caches": caches,
        "temporary": temporary,
    }


def _receipt(plan: Mapping[str, Any], archive: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "schema": SCHEMA,
        "version": VERSION,
        "status": "pass",
        "completed_utc": now(),
        "driver": _file_descriptor(Path(__file__).resolve(), "0487 cleanup driver", relative_to_root=True),
        "all_root_jobs_terminal": True,
        "comparison": {
            "summary": _file_descriptor(ROOT / "comparison-summary.json", "0487 comparison summary", relative_to_root=True),
            "formal_processes": 144,
            "formal_samples": 4320,
        },
        "profiles": plan["profiles"],
        "builds": plan["builds"],
        "fuzz": plan["fuzz"],
        "replay_cleanup_receipts": plan["replay_receipts"],
        "all_replay_directories_removed": True,
        "removed": plan["temporary"]["removed"],
        "removed_bytecode": plan["temporary"]["bytecode"],
        "retained_external_binaries": [
            plan["builds"]["normal"]["binary"],
            plan["builds"]["allocator"]["binary"],
            plan["fuzz"]["binary"],
        ],
        "retained_fuzz_build_inputs": archive,
        "shared_caches_retained": plan["caches"],
        "temporary_root_retained_for_binaries": str(TEMP.resolve()),
        "scratch_files_remaining": 0,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true", help="validate custody and print the plan without archiving or deleting")
    args = parser.parse_args(argv)
    try:
        plan = _plan()
        if args.dry_run:
            print({"schema": SCHEMA, "status": "ready", "replay_cleanup_receipts": len(plan["replay_receipts"]), "retained_binaries": 3})
            return 0
        archive = _archive_inputs()
        _remove_authenticated(plan)
        receipt = _receipt(plan, archive)
        write(ROOT / "cleanup.json", receipt)
        print({"schema": SCHEMA, "status": "pass", "replay_cleanup_receipts": len(plan["replay_receipts"]), "retained_binaries": 3})
        return 0
    except (CleanupError, SealError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"cleanup.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
