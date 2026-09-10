#!/usr/bin/env python3
"""Create and verify the immutable 0497 publication evidence seal.

The verifier is the final read-only custody step.  It imports the frozen
measurement helper with bytecode disabled, revalidates the formal and pilot
raw reports, recomputes the formal analysis, checks profile and fuzz evidence,
and delegates the retained-build/cleanup checks to the bounded cleanup
validator.  The seal inventory hashes every regular evidence file while
excluding only this seal and transient bytecode directories.
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
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0497")
TARGET = TEMP / "target"
FUZZ_TEMP = TEMP / "fuzz-target"
SEAL_NAME = "seal.json"
SCHEMA = "docx-tail-append-publication-evidence-seal-v1"
VERIFY_SCHEMA = "docx-tail-append-publication-seal-verification-v1"
VERSION = 1
FORMAL_ATTEMPT = "formal1"
PILOT_ATTEMPT = "pilot1"
BYTECODE_SUFFIXES = frozenset({".pyc", ".pyo"})
EXCLUDED = {
    "self": SEAL_NAME,
    "transient": "**/__pycache__/**",
    "bytecode": "reject outside transient directories",
}
ADR_REFRESH = "adr-refresh.json"
PROTECTED_PRIMARY = "protected-primary.json"
FINAL_GATES = "final-gates.json"
MANDATORY_GATE_NAMES = frozenset({
    "focused",
    "docx-default",
    "docx-features",
    "opc-atomic",
    "opc-preservation",
    "doc-tests",
    "workspace-check",
    "production-clippy",
    "production-rustdoc",
    "harness-allocator",
    "harness-clippy",
    "harness-rustdoc",
    "fmt",
    "harness-fmt",
    "boundaries",
})
OPTIONAL_GATE_NAMES = frozenset({"helpers"})
PROFILE_HELPERS = frozenset({"profile.py", "profile_v2.py"})
PROFILE_MODES = frozenset({"counting", "atomic"})
FAILED_PROFILE_ATTEMPT = "strace1"
FAILED_PROFILE_SYSCALL = "fstatat"
FAILED_PREFLIGHT_ATTEMPT = "strace2"
FAILED_PREFLIGHT_HELPER = "profile_v2.py"
FAILED_PREFLIGHT_ERROR = (
    'Error: TailAppend(Scan("XML Events limit 6807808 exceeds hard ceiling 4000000"))\n'
)
FAILED_PREFLIGHT4_ATTEMPT = "strace4"
FAILED_PREFLIGHT4_ERROR = (
    'Error: TailAppend(Scan("XML Events limit 6742784 exceeds hard ceiling 4000000"))\n'
)
FAILED_CLI_ATTEMPT = "strace3"
FAILED_CLI_ERROR = (
    'Error: "unsupported --source-counts count 4096; expected one of [64, 8192, 131072]"\n'
)
FAILED_CLI_RAW_MARKER = (
    "unsupported --source-counts count 4096; expected one of [64, 8192, 131072]"
)
FAILED_PROFILE_ATTEMPTS = frozenset({
    FAILED_PROFILE_ATTEMPT, FAILED_PREFLIGHT_ATTEMPT, FAILED_PREFLIGHT4_ATTEMPT,
    FAILED_CLI_ATTEMPT,
})
SUCCESSFUL_PROFILE_ATTEMPT = "strace5"
EARLY_TARGET_CLEANUP = "early-target-cleanup.json"
EARLY_FUZZ_CLEANUP = "early-fuzz-target-cleanup.json"
RESUME_DRIVER = "resume.py"
RESUME_TEST_DRIVER = "test_resume.py"
RESUME_ATTEMPT = FORMAL_ATTEMPT
RESUME_START = "resume-start.json"
RESUME_TERMINAL = "resume-terminal.json"
INTERRUPTED_RECEIPT = "interrupted.json"
RESUME_REQUIRED_INPUTS = (
    "measure.py", "protocol.json", "builds.json", "provenance.json", "machine.json",
    "candidate-source.json", "interruption-observation.json", RESUME_DRIVER,
    RESUME_TEST_DRIVER,
)
GATE_ENVIRONMENT = {
    "RUSTUP_TOOLCHAIN": "1.98.1",
    "CARGO_BUILD_JOBS": "4",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
    "CARGO_TARGET_DIR": str(TEMP / "target"),
    "TMPDIR": str(TEMP / "tmp"),
    "DEBUGINFOD_URLS": "",
    "LC_ALL": "C",
    "RUSTDOCFLAGS": "-Dwarnings",
    "PYTHONDONTWRITEBYTECODE": "1",
}
GATE_RECEIPT_KEYS = frozenset({
    "argv", "cwd", "driver", "environment", "exit_code", "finished_ns", "pid",
    "source_manifest", "source_unchanged", "started_ns", "stderr", "stdout",
    "termination", "timed_out", "timeout_seconds",
})


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


def _meta(path: Path, label: str = "file") -> dict[str, int | str]:
    try:
        state = path.lstat()
    except OSError as error:
        fail(f"{label}: cannot stat {path}: {error}")
    require(stat.S_ISREG(state.st_mode), f"{label}: regular file required: {path}")
    return {"path": str(path), "bytes": state.st_size, "sha256": _sha(path)}


def _read_json(path: Path, label: str) -> Any:
    _meta(path, label)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")
    _finite(value, label)
    return value


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat(timespec="microseconds").replace("+00:00", "Z")


def _timestamp(value: Any, label: str) -> None:
    require(isinstance(value, str) and value, f"{label}: timestamp missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        fail(f"{label}: malformed timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label}: timestamp has no timezone")
    require(parsed.utcoffset() == _datetime.timedelta(0),
            f"{label}: timestamp is not UTC")


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


def _inside(root: Path, path: Path, label: str) -> str:
    _no_symlink_components(path, label)
    try:
        relative = path.resolve(strict=True).relative_to(root.resolve(strict=True)).as_posix()
    except (OSError, ValueError) as error:
        fail(f"{label}: path is missing or outside evidence root: {error}")
    return relative


def _descriptor(root: Path, path: Path, label: str) -> dict[str, int | str]:
    relative = _inside(root, path, label)
    resolved = path.resolve(strict=True)
    metadata = _meta(resolved, label)
    return {"path": relative, "bytes": metadata["bytes"], "sha256": metadata["sha256"]}


def _inventory(root: Path) -> dict[str, dict[str, int | str]]:
    """Hash every regular evidence file except seal output/transient bytecode."""

    result: dict[str, dict[str, int | str]] = {}

    def visit(directory: Path, relative: str) -> None:
        _no_symlink_components(directory, "evidence inventory")
        try:
            children = sorted(directory.iterdir(), key=lambda item: item.name)
        except OSError as error:
            fail(f"cannot enumerate evidence directory {directory}: {error}")
        for child in children:
            child_relative = f"{relative}/{child.name}" if relative else child.name
            state = child.lstat()
            require(not stat.S_ISLNK(state.st_mode), f"evidence contains symlink: {child_relative}")
            if stat.S_ISDIR(state.st_mode):
                if child.name == "__pycache__":
                    continue
                visit(child, child_relative)
                continue
            require(stat.S_ISREG(state.st_mode), f"evidence contains special path: {child_relative}")
            if child_relative == SEAL_NAME:
                continue
            require(child.suffix.lower() not in BYTECODE_SUFFIXES,
                    f"bytecode is not allowed: {child_relative}")
            result[child_relative] = _meta(child, f"evidence file {child_relative}")

    visit(root, "")
    return result


def _load_module(name: str, path: Path) -> Any:
    _meta(path, f"helper {name}")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load helper {name}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        fail(f"helper {name} import failed: {error}")
    return module


def _remove_empty(path: Path, label: str) -> None:
    try:
        state = path.lstat()
    except FileNotFoundError:
        return
    except OSError as error:
        fail(f"{label}: cannot inspect: {error}")
    require(stat.S_ISDIR(state.st_mode), f"{label}: directory required")
    _no_symlink_components(path, label)
    try:
        children = tuple(path.iterdir())
    except OSError as error:
        fail(f"{label}: cannot enumerate: {error}")
    require(not children, f"{label}: unexplained contents remain")
    try:
        path.rmdir()
    except OSError as error:
        fail(f"{label}: cannot remove empty directory: {error}")


def _capture_files(root: Path, entries: list[Mapping[str, Any]], label: str) -> dict[str, Any]:
    records: list[dict[str, Any]] = []
    for entry in entries:
        spec = entry.get("spec")
        directory = Path(entry["directory"])
        require(directory.is_dir() and not directory.is_symlink(), f"{label}: capture directory missing")
        files: dict[str, dict[str, int | str]] = {}
        for child in sorted(directory.iterdir(), key=lambda item: item.name):
            state = child.lstat()
            require(stat.S_ISREG(state.st_mode) and not stat.S_ISLNK(state.st_mode),
                    f"{label}: capture contains special path: {child}")
            files[child.name] = _descriptor(root, child, f"{label} {child.name}")
        require(files, f"{label}: capture has no raw artifacts")
        records.append({
            "label": spec.get("label") if isinstance(spec, Mapping) else directory.name,
            "directory": _inside(root, directory, f"{label} directory"),
            "files": files,
        })
    records.sort(key=lambda item: str(item["label"]))
    return {"attempt": label, "children": len(records), "captures": records}


def _formal_binding(measure: Any, protocol: Mapping[str, Any], protocol_hash: str,
                    builds: Mapping[str, Any], root: Path) -> dict[str, Any]:
    try:
        entries = measure._collect(FORMAL_ATTEMPT, dict(protocol), dict(builds))
        pilot = measure._collect_lane(PILOT_ATTEMPT, dict(protocol), dict(builds), pilot=True)
    finally:
        _remove_empty(Path(measure.TEMP) / "projections", "measurement projections")
    require(len(entries) == 288, "formal capture child count differs")
    require(len(pilot) == 72, "pilot capture child count differs")
    protocol_ref = {"path": str(measure.PROTOCOL_FILE), "sha256": protocol_hash}
    expected = measure.analyze_data(entries, protocol=protocol_ref)
    analysis_path = root / "analysis" / f"{FORMAL_ATTEMPT}.json"
    actual = _read_json(analysis_path, "formal analysis")
    require(actual == expected, "formal analysis does not recompute from raw reports")
    require(actual.get("schema") == measure.ANALYSIS_SCHEMA
            and actual.get("child_count") == 288
            and actual.get("sample_count") == 8640,
            "formal analysis counts differ")
    verification_path = root / "verification" / f"{FORMAL_ATTEMPT}.json"
    verification = _read_json(verification_path, "formal verification")
    require(isinstance(verification, Mapping)
            and verification.get("schema") == "docx-replayable-tail-publication-verification-v1"
            and verification.get("version") == 1
            and verification.get("attempt") == FORMAL_ATTEMPT
            and verification.get("children") == 288
            and verification.get("samples") == 8640
            and verification.get("status") == "pass"
            and verification.get("protocol") == protocol_ref,
            "formal verification receipt differs")
    return {
        "formal": _capture_files(root, entries, FORMAL_ATTEMPT),
        "pilot": _capture_files(root, pilot, PILOT_ATTEMPT),
        "analysis": _descriptor(root, analysis_path, "formal analysis"),
        "verification": {"descriptor": _descriptor(root, verification_path, "formal verification"),
                         "receipt": dict(verification)},
    }


def _resume_meta(root: Path, value: Any, path: Path, label: str) -> dict[str, int | str]:
    """Authenticate a resume.py absolute-path file descriptor and normalize it."""

    require(isinstance(value, Mapping) and set(value) == {"path", "bytes", "sha256"},
            f"{label}: file metadata is malformed")
    expected = _meta(path, label)
    require(dict(value) == expected, f"{label}: file metadata changed")
    return _descriptor(root, path, label)


def _resume_tree_inventory(root: Path) -> list[dict[str, Any]]:
    """Inventory the archived interrupted tree, rejecting links and specials."""

    require(root.is_dir() and not root.is_symlink(), f"resume archive directory missing: {root}")
    result: list[dict[str, Any]] = []
    for path in sorted(root.rglob("*"), key=lambda item: str(item)):
        require(not path.is_symlink(), f"resume archive contains symlink: {path}")
        relative = path.relative_to(root).as_posix()
        if path.is_dir():
            result.append({"path": str(path), "relative": relative, "kind": "directory"})
        elif path.is_file():
            result.append({"path": str(path), "relative": relative, "kind": "file",
                           "bytes": path.stat().st_size, "sha256": _sha(path)})
        else:
            fail(f"resume archive contains special path: {path}")
    return result


def _resume_optional_directory_inventory(root: Path, declared: Any, label: str) -> list[dict[str, Any]]:
    """Check a declared empty-directory tree when Git may omit its directories."""

    require(isinstance(declared, list) and declared,
            f"{label}: directory inventory is missing")
    expected: list[dict[str, Any]] = []
    for item in declared:
        require(isinstance(item, Mapping)
                and set(item) == {"path", "relative", "kind"}
                and isinstance(item.get("path"), str)
                and isinstance(item.get("relative"), str)
                and item.get("kind") == "directory",
                f"{label}: directory inventory is malformed")
        relative = item["relative"]
        require(relative and not Path(relative).is_absolute()
                and ".." not in Path(relative).parts,
                f"{label}: directory inventory escapes archive")
        require(Path(item["path"]).resolve() == (root / relative).resolve(),
                f"{label}: directory inventory path differs")
        expected.append({"path": str(root / relative), "relative": relative, "kind": "directory"})
    expected.sort(key=lambda item: item["relative"])
    require(expected == sorted(expected, key=lambda item: item["relative"]),
            f"{label}: directory inventory order differs")
    _no_symlink_components(root, label)
    if not root.exists():
        return []
    require(root.is_dir() and not root.is_symlink(), f"{label}: archive root is not a directory")
    actual = _resume_tree_inventory(root)
    require(all(item.get("kind") == "directory" for item in actual),
            f"{label}: unexpected archived private files")
    expected_by_relative = {item["relative"]: item for item in expected}
    require(all(item["relative"] in expected_by_relative
                and item["path"] == expected_by_relative[item["relative"]]["path"]
                for item in actual),
            f"{label}: unexplained archived private directory")
    return actual


def _resume_input_paths(root: Path) -> dict[str, Path]:
    return {
        "measure.py": root / "measure.py",
        "protocol.json": root / "protocol.json",
        "builds.json": root / "builds.json",
        "provenance.json": root / "provenance.json",
        "machine.json": root / "machine.json",
        "candidate-source.json": root / "candidate-source.json",
        "interruption-observation.json": root / "interruption-observation.json",
        RESUME_DRIVER: root / RESUME_DRIVER,
        RESUME_TEST_DRIVER: root / RESUME_TEST_DRIVER,
    }


def _resume_timestamp(value: Any, label: str) -> None:
    _timestamp(value, label)


def _resume_formal_entries(measure: Any, protocol: Mapping[str, Any], formal_capture: Mapping[str, Any],
                           root: Path) -> list[dict[str, Any]]:
    """Recover lightweight chronology entries from the already validated formal capture."""

    captures = formal_capture.get("captures")
    require(isinstance(captures, list) and len(captures) == 288,
            "resume formal capture binding is not complete")
    by_label: dict[str, Mapping[str, Any]] = {}
    for item in captures:
        require(isinstance(item, Mapping) and isinstance(item.get("label"), str),
                "resume formal capture label is malformed")
        require(item["label"] not in by_label, "resume formal capture labels are duplicated")
        by_label[item["label"]] = item
    specs = protocol.get("formal_runs")
    require(isinstance(specs, list) and len(specs) == 288,
            "resume formal protocol inventory is incomplete")
    entries: list[dict[str, Any]] = []
    for raw_spec in specs:
        require(isinstance(raw_spec, Mapping), "resume formal specification is malformed")
        label = raw_spec.get("label")
        require(isinstance(label, str) and label in by_label,
                f"resume formal capture is missing: {label}")
        directory_value = by_label[label].get("directory")
        require(isinstance(directory_value, str), f"resume formal directory is malformed: {label}")
        directory = root / directory_value
        require(directory.is_dir() and not directory.is_symlink(),
                f"resume formal directory is missing: {directory}")
        started = _read_json(directory / "started.json", f"resume {label} started")
        terminal = _read_json(directory / "terminal.json", f"resume {label} terminal")
        require(isinstance(started, Mapping) and isinstance(terminal, Mapping),
                f"resume {label}: terminal receipts are malformed")
        entries.append({
            "spec": dict(raw_spec, attempt=RESUME_ATTEMPT),
            "directory": directory,
            "started": started,
            "terminal": terminal,
            "started_at": measure._timestamp(started.get("started_utc"),
                                               f"{directory}.started_utc"),
            "finished_at": measure._timestamp(terminal.get("finished_utc"),
                                                f"{directory}.finished_utc"),
        })
    return entries


def _resume_prefix_receipts(root: Path, formal_capture: Mapping[str, Any],
                            specs: list[Mapping[str, Any]], start: Mapping[str, Any]) -> list[dict[str, Any]]:
    raw = start.get("prefix_receipts")
    require(isinstance(raw, list) and len(raw) == len(specs),
            "resume prefix receipt count differs")
    capture_by_label = {item["label"]: item for item in formal_capture["captures"]}
    result: list[dict[str, Any]] = []
    names = ("started.json", "terminal.json", "report.json", "resource.txt", "replay-cleanup.json")
    for ordinal, (spec, receipt) in enumerate(zip(specs, raw, strict=True)):
        require(isinstance(spec, Mapping) and isinstance(receipt, Mapping),
                f"resume prefix receipt {ordinal} is malformed")
        require(set(receipt) == {"ordinal", "label", "started", "terminal", "report",
                                 "resource", "replay_cleanup"},
                f"resume prefix receipt {ordinal} fields differ")
        label = spec.get("label")
        require(receipt.get("ordinal") == ordinal and receipt.get("label") == label
                and isinstance(label, str) and label in capture_by_label,
                f"resume prefix receipt {ordinal} identity differs")
        capture = capture_by_label[label]
        directory_value = capture.get("directory")
        require(isinstance(directory_value, str), f"resume prefix directory is malformed: {label}")
        directory = root / directory_value
        files = capture.get("files")
        require(isinstance(files, Mapping), f"resume prefix files are malformed: {label}")
        normalized: dict[str, Any] = {"ordinal": ordinal, "label": label}
        for name, key in zip(names, ("started", "terminal", "report", "resource", "replay_cleanup"), strict=True):
            require(name in files, f"resume prefix artifact is missing: {directory / name}")
            actual = directory / name
            require(actual.is_file() and not actual.is_symlink(),
                    f"resume prefix artifact is missing: {actual}")
            actual_descriptor = files[name]
            require(isinstance(actual_descriptor, Mapping)
                    and set(actual_descriptor) == {"path", "bytes", "sha256"},
                    f"resume prefix artifact descriptor is malformed: {actual}")
            expected = {"path": str(actual.resolve()), "bytes": actual_descriptor["bytes"],
                        "sha256": actual_descriptor["sha256"]}
            require(dict(receipt[key]) == expected,
                    f"resume prefix artifact binding changed: {actual}")
            require(dict(actual_descriptor) == {
                "path": str(actual.relative_to(root).as_posix()),
                "bytes": actual_descriptor["bytes"], "sha256": actual_descriptor["sha256"]},
                    f"resume prefix formal inventory changed: {actual}")
            normalized[key] = _descriptor(root, actual, f"resume prefix {label} {name}")
        result.append(normalized)
    return result


def _resume_live_roots(capture: Path, private: Path) -> None:
    """The completed suffix reuses its capture path, but releases private scratch.

    The formal collector authenticates the replacement reports and terminals;
    the interrupted binding independently authenticates the archived raw files.
    Requiring the capture path to be absent would reject every completed resume.
    """

    require(capture.is_dir() and not capture.is_symlink(),
            "resumed formal capture root is missing or unsafe")
    require(not private.exists() and not private.is_symlink(),
            "resumed private scratch remains")


def _resume_interrupted_binding(root: Path, resume: Any, measure: Any,
                                protocol: Mapping[str, Any], protocol_hash: str,
                                builds: Mapping[str, Any], formal_capture: Mapping[str, Any],
                                start: Mapping[str, Any]) -> dict[str, Any]:
    interrupted_root = root / "interrupted" / "formal1-enospc"
    receipt_path = interrupted_root / INTERRUPTED_RECEIPT
    capture_root = interrupted_root / "capture"
    private_root = interrupted_root / "private-root"
    receipt = _read_json(receipt_path, "interrupted resume receipt")
    require(isinstance(receipt, Mapping)
            and set(receipt) == {
                "schema", "version", "status", "attempt", "ordinal", "label", "scope",
                "driver_error", "protocol", "measure_driver", "builds", "observation",
                "original_capture_directory", "archived_capture_directory", "original_private_root",
                "archived_private_root", "capture_files", "capture_inventory_before",
                "private_inventory_before", "private_inventory_after", "inputs",
            }, "interrupted resume receipt fields differ")
    require(receipt.get("schema") == resume.INTERRUPTED_SCHEMA and receipt.get("version") == 1
            and receipt.get("status") == "archived_incomplete" and receipt.get("attempt") == RESUME_ATTEMPT
            and receipt.get("ordinal") == 191,
            "interrupted resume receipt identity differs")
    specs = protocol.get("formal_runs")
    require(isinstance(specs, list) and len(specs) == 288 and isinstance(specs[191], Mapping),
            "interrupted formal specification is missing")
    failed_spec = specs[191]
    require(receipt.get("label") == failed_spec.get("label"),
            "interrupted formal label differs")
    require(receipt.get("scope") ==
            "raw interrupted evidence only; no terminal receipt or measurement is reconstructed",
            "interrupted resume scope differs")
    observation = _read_json(root / "interruption-observation.json", "interruption observation")
    require(isinstance(observation, Mapping)
            and observation.get("schema") == "docx-publication-interruption-observation-v1"
            and observation.get("completed_terminal_children") == 191
            and observation.get("incomplete_label") == failed_spec.get("label")
            and isinstance(observation.get("child_exit_code"), str)
            and observation["child_exit_code"].startswith("unknown"),
            "interruption observation binding differs")
    driver_error = receipt.get("driver_error")
    require(isinstance(driver_error, Mapping)
            and set(driver_error) == {"kind", "message", "coordinator_exit_code", "child_exit_code", "scope"}
            and driver_error.get("kind") == "ENOSPC"
            and driver_error.get("message") == observation.get("observed_error")
            and driver_error.get("coordinator_exit_code") == observation.get("coordinator_exit_code")
            and driver_error.get("child_exit_code") == "unknown"
            and driver_error.get("scope") ==
            "the child stopped before report/resource/terminal completion",
            "interrupted driver error binding differs")
    original_capture = Path(str(receipt.get("original_capture_directory")))
    require(original_capture == root / "captures" / "formal1" / str(failed_spec["label"]),
            "interrupted original capture path differs")
    require(Path(str(receipt.get("archived_capture_directory"))) == capture_root,
            "interrupted archive capture path differs")
    original_private = Path(str(receipt.get("original_private_root")))
    archived_private = Path(str(receipt.get("archived_private_root")))
    require(Path(str(receipt.get("archived_private_root"))) == private_root,
            "interrupted archive private path differs")
    _resume_live_roots(original_capture, original_private)
    require(capture_root.is_dir(), "interrupted archive capture root is missing")

    input_paths = _resume_input_paths(root)
    expected_inputs = {name: _meta(path, f"resume input {name}")
                       for name, path in input_paths.items()}
    expected_input_keys = set(RESUME_REQUIRED_INPUTS)
    require(set(receipt.get("inputs", {})) == expected_input_keys,
            "interrupted input inventory differs")
    for name in RESUME_REQUIRED_INPUTS:
        require(receipt["inputs"].get(name) == expected_inputs[name],
                f"interrupted input binding changed: {name}")
    require(receipt.get("protocol") == expected_inputs["protocol.json"]
            and receipt.get("measure_driver") == expected_inputs["measure.py"]
            and receipt.get("builds") == expected_inputs["builds.json"]
            and receipt.get("observation") == expected_inputs["interruption-observation.json"],
            "interrupted source custody bindings differ")

    capture_files = receipt.get("capture_files")
    require(isinstance(capture_files, list) and len(capture_files) == len(resume.REQUIRED_RAW_FILES),
            "interrupted raw artifact inventory differs")
    by_name: dict[str, Mapping[str, Any]] = {}
    for item in capture_files:
        require(isinstance(item, Mapping)
                and set(item) == {"original", "archive", "original_name"}
                and isinstance(item.get("original_name"), str),
                "interrupted raw artifact record is malformed")
        name = item["original_name"]
        require(name in resume.REQUIRED_RAW_FILES and name not in by_name,
                f"interrupted raw artifact name differs: {name}")
        by_name[name] = item
        archive_path = capture_root / resume.raw_archive_name(name)
        original_path = original_capture / name
        require(isinstance(item.get("original"), Mapping)
                and set(item["original"]) == {"path", "bytes", "sha256"}
                and isinstance(item.get("archive"), Mapping)
                and set(item["archive"]) == {"path", "bytes", "sha256"},
                f"interrupted raw artifact metadata is malformed: {name}")
        require(Path(str(item["original"]["path"])).resolve() == original_path.resolve(),
                f"interrupted original artifact path differs: {name}")
        require(dict(item["archive"]) == _meta(archive_path, f"interrupted raw archive {name}"),
                f"interrupted raw archive changed: {name}")
        require(item["original"]["bytes"] == item["archive"]["bytes"]
                and item["original"]["sha256"] == item["archive"]["sha256"],
                f"interrupted raw archive hash differs: {name}")
    require(set(by_name) == set(resume.REQUIRED_RAW_FILES),
            "interrupted raw artifact set is incomplete")
    require({path.name for path in capture_root.iterdir()} == {
        resume.raw_archive_name(name) for name in resume.REQUIRED_RAW_FILES
    }, "interrupted archive contains unexplained capture files")

    capture_inventory = receipt.get("capture_inventory_before")
    require(isinstance(capture_inventory, list) and len(capture_inventory) == len(resume.REQUIRED_RAW_FILES),
            "interrupted pre-archive capture inventory differs")
    expected_capture_inventory = []
    for name in sorted(resume.REQUIRED_RAW_FILES):
        item = by_name[name]["original"]
        expected_capture_inventory.append({
            "path": str(original_capture / name), "relative": name, "kind": "file",
            "bytes": item["bytes"], "sha256": item["sha256"],
        })
    require(capture_inventory == expected_capture_inventory,
            "interrupted pre-archive capture inventory changed")

    private_before = receipt.get("private_inventory_before")
    private_after = receipt.get("private_inventory_after")
    require(isinstance(private_before, list) and isinstance(private_after, list),
            "interrupted private inventory is malformed")
    expected_private_after = []
    for item in private_before:
        require(isinstance(item, Mapping) and isinstance(item.get("relative"), str)
                and item.get("kind") == "directory",
                "interrupted private pre-archive inventory is malformed")
        relative = item["relative"]
        require(relative and not Path(relative).is_absolute()
                and ".." not in Path(relative).parts
                and Path(str(item.get("path"))).resolve() == (original_private / relative).resolve(),
                "interrupted private pre-archive path differs")
        expected_private_after.append({
            **dict(item), "path": str(private_root / relative),
        })
    require(private_after == expected_private_after,
            "interrupted private inventory path binding differs")
    require(private_after == [{"path": str(private_root / "tmp"),
                               "relative": "tmp", "kind": "directory"}],
            "interrupted private scratch inventory differs")
    actual_private_after = _resume_optional_directory_inventory(
        private_root, private_after, "interrupted archived private inventory")
    if actual_private_after:
        require(private_after == actual_private_after,
                "interrupted archived private inventory changed")

    partial_started = _read_json(capture_root / "started.json.raw", "interrupted started raw")
    require(isinstance(partial_started, Mapping)
            and partial_started.get("schema") == measure.CAPTURE_SCHEMA
            and partial_started.get("version") == measure.VERSION
            and partial_started.get("status") == "running"
            and partial_started.get("attempt") == RESUME_ATTEMPT,
            "interrupted started raw receipt identity differs")
    run = partial_started.get("run")
    run_fields = ("ordinal", "repeat", "phase", "publication", "role", "arm", "workload",
                  "route", "input_mode", "samples", "warmups", "label")
    require(isinstance(run, Mapping)
            and dict(run) == {key: failed_spec[key] for key in run_fields},
            "interrupted started raw run binding differs")
    build = builds[measure._build_key(failed_spec["phase"], failed_spec["role"])]
    require(partial_started.get("protocol") == {
                "path": str(root / "protocol.json"), "sha256": protocol_hash}
            and partial_started.get("driver") == protocol["driver"]
            and partial_started.get("build") == measure._build_binding(build)
            and partial_started.get("binary") == build["binary"]
            and partial_started.get("source_manifest") == build["source_manifest"],
            "interrupted started raw source/build binding differs")
    expected_private = measure._run_root(RESUME_ATTEMPT, failed_spec["label"])
    require(Path(str(partial_started.get("private_root"))) == expected_private
            and Path(str(partial_started.get("tmpdir"))) == expected_private / "tmp"
            and partial_started.get("replay_dir") is None,
            "interrupted started raw private binding differs")
    expected_argv = measure._command(failed_spec, build, original_capture / "report.json",
                                     original_capture / "resource.txt", None, None)
    require(partial_started.get("argv") == expected_argv,
            "interrupted started raw command binding differs")
    _resume_timestamp(partial_started.get("started_utc"), "interrupted started timestamp")

    prefix_specs = [dict(item) for item in specs[:191] if isinstance(item, Mapping)]
    require(len(prefix_specs) == 191, "resume prefix specification count differs")
    prefix_entries = _resume_formal_entries(measure, protocol, formal_capture, root)
    ordered = measure._chronological_entries(prefix_entries, [dict(item) for item in specs],
                                             RESUME_ATTEMPT)
    require(ordered[190]["terminal"]["finished_utc"] == start.get("prefix_last_terminal_utc"),
            "resume prefix terminal timestamp differs")
    _resume_timestamp(start.get("prefix_last_terminal_utc"), "resume prefix terminal timestamp")
    _resume_timestamp(start.get("interrupted_started_utc"), "interrupted started timestamp")
    require(measure._timestamp(start["interrupted_started_utc"], "interrupted started timestamp")
            >= ordered[190]["finished_at"],
            "resume interrupted start precedes prefix completion")
    normalized_prefix = _resume_prefix_receipts(root, formal_capture, prefix_specs, start)
    return {
        "interrupted": {
            "receipt": _descriptor(root, receipt_path, "interrupted resume receipt"),
            "capture": [_descriptor(root, capture_root / resume.raw_archive_name(name),
                                     f"interrupted raw {name}")
                        for name in sorted(resume.REQUIRED_RAW_FILES)],
            "private_inventory": private_after,
            "ordinal": 191,
            "label": failed_spec["label"],
        },
        "prefix": {"children": 191, "last_terminal_utc": start["prefix_last_terminal_utc"],
                   "receipts": normalized_prefix},
        "suffix": {"children": 97,
                   "first_started_utc": ordered[191]["started"]["started_utc"],
                   "last_terminal_utc": ordered[-1]["terminal"]["finished_utc"]},
    }


def _resume_binding(root: Path, measure: Any, protocol: Mapping[str, Any], protocol_hash: str,
                    builds: Mapping[str, Any], formal_capture: Mapping[str, Any]) -> dict[str, Any]:
    """Bind immutable resume receipts and the completed frozen suffix."""

    resume_path = root / RESUME_DRIVER
    test_path = root / RESUME_TEST_DRIVER
    resume = _load_module("change0497_resume_for_seal", resume_path)
    require(resume.ATTEMPT == RESUME_ATTEMPT and resume.START_ORDINAL == 191
            and Path(resume.FORMAL_ROOT).resolve() == (root / "captures" / "formal1").resolve()
            and Path(resume.INTERRUPTED_ROOT).resolve() ==
            (root / "interrupted" / "formal1-enospc").resolve()
            and Path(resume.RESUME_ROOT).resolve() == (root / "resume" / "formal1").resolve(),
            "resume driver roots or ordinal differ")
    start_path = root / "resume" / "formal1" / RESUME_START
    terminal_path = root / "resume" / "formal1" / RESUME_TERMINAL
    require(start_path.is_file() and terminal_path.is_file(),
            "resume start/terminal receipts are missing")
    start = _read_json(start_path, "resume start receipt")
    terminal = _read_json(terminal_path, "resume terminal receipt")
    require(isinstance(start, Mapping) and isinstance(terminal, Mapping),
            "resume start/terminal receipts are malformed")
    require(set(start) == {
        "schema", "version", "status", "attempt", "started_utc", "exit_code", "start_ordinal",
        "remaining_children", "protocol", "driver", "test_driver", "inputs", "archive",
        "prefix_children", "prefix_last_terminal_utc", "interrupted_started_utc", "prefix_receipts",
        "remaining_labels", "scope",
    }, "resume start receipt fields differ")
    require(set(terminal) == {
        "schema", "version", "status", "attempt", "started_utc", "finished_utc", "exit_code",
        "start_ordinal", "remaining_children", "completed_children", "start_receipt", "scope",
    }, "resume terminal receipt fields differ")
    require(start.get("schema") == resume.RESUME_START_SCHEMA and start.get("version") == 1
            and start.get("status") == "running" and start.get("attempt") == RESUME_ATTEMPT
            and start.get("exit_code") is None and start.get("start_ordinal") == 191
            and start.get("remaining_children") == 97 and start.get("prefix_children") == 191,
            "resume start status/count binding differs")
    require(terminal.get("schema") == resume.RESUME_TERMINAL_SCHEMA and terminal.get("version") == 1
            and terminal.get("status") == "pass" and terminal.get("attempt") == RESUME_ATTEMPT
            and terminal.get("exit_code") == 0 and terminal.get("start_ordinal") == 191
            and terminal.get("remaining_children") == 97 and terminal.get("completed_children") == 97,
            "resume terminal status/count binding differs")
    _resume_timestamp(start.get("started_utc"), "resume start timestamp")
    _resume_timestamp(terminal.get("started_utc"), "resume terminal start timestamp")
    _resume_timestamp(terminal.get("finished_utc"), "resume terminal finish timestamp")
    require(terminal.get("started_utc") == start.get("started_utc")
            and measure._timestamp(terminal["finished_utc"], "resume terminal finish timestamp")
            > measure._timestamp(start["started_utc"], "resume start timestamp"),
            "resume terminal chronology differs")
    require(start.get("protocol") == {"path": str(root / "protocol.json"), "sha256": protocol_hash},
            "resume protocol binding differs")
    input_paths = _resume_input_paths(root)
    expected_inputs = {name: _meta(path, f"resume input {name}")
                       for name, path in input_paths.items()}
    interrupted_path = root / "interrupted" / "formal1-enospc" / INTERRUPTED_RECEIPT
    expected_inputs[INTERRUPTED_RECEIPT] = _meta(interrupted_path, "interrupted resume receipt")
    require(set(start.get("inputs", {})) == set(RESUME_REQUIRED_INPUTS) | {INTERRUPTED_RECEIPT},
            "resume start input inventory differs")
    for name, expected in expected_inputs.items():
        require(start["inputs"].get(name) == expected,
                f"resume start input binding changed: {name}")
    require(start.get("driver") == expected_inputs[RESUME_DRIVER]
            and start.get("test_driver") == expected_inputs[RESUME_TEST_DRIVER],
            "resume driver/test binding differs")
    require(terminal.get("start_receipt") == _meta(start_path, "resume start receipt"),
            "resume terminal start receipt metadata differs")
    specs = protocol.get("formal_runs")
    require(isinstance(specs, list) and len(specs) == 288,
            "resume protocol formal inventory differs")
    expected_labels = [spec["label"] for spec in specs[191:] if isinstance(spec, Mapping)]
    require(len(expected_labels) == 97 and start.get("remaining_labels") == expected_labels,
            "resume remaining label inventory differs")
    interrupted_path = root / "interrupted" / "formal1-enospc" / INTERRUPTED_RECEIPT
    require(start.get("archive") == _meta(interrupted_path, "interrupted resume receipt"),
            "resume archive receipt binding differs")
    require(start.get("scope") == "continuation of frozen formal1 suffix under shared CPU lock"
            and terminal.get("scope") ==
            "resume receipts do not replace the frozen formal protocol or child terminals",
            "resume receipt scope differs")
    interrupted = _resume_interrupted_binding(root, resume, measure, protocol, protocol_hash,
                                              builds, formal_capture, start)
    require(measure._timestamp(interrupted["suffix"]["first_started_utc"],
                               "resume suffix first start")
            >= measure._timestamp(start["started_utc"], "resume start timestamp"),
            "resume suffix started before resume start")
    require(measure._timestamp(terminal["finished_utc"], "resume terminal finish timestamp")
            >= measure._timestamp(interrupted["suffix"]["last_terminal_utc"],
                                  "resume suffix last terminal"),
            "resume terminal finished before final child")
    return {
        "driver": _descriptor(root, resume_path, "resume driver"),
        "test_driver": _descriptor(root, test_path, "resume tests"),
        "protocol": {"descriptor": _descriptor(root, root / "protocol.json", "resume protocol"),
                     "sha256": protocol_hash},
        "inputs": {name: _descriptor(root, path, f"resume input {name}")
                   for name, path in input_paths.items()},
        "start": {"descriptor": _descriptor(root, start_path, "resume start receipt"),
                   "started_utc": start["started_utc"], "remaining_children": 97,
                   "prefix_children": 191, "remaining_labels": expected_labels},
        "terminal": {"descriptor": _descriptor(root, terminal_path, "resume terminal receipt"),
                      "status": "pass", "exit_code": 0, "completed_children": 97,
                      "finished_utc": terminal["finished_utc"]},
        **interrupted,
    }


def _profile_helper_binding(root: Path, value: Any, label: str) -> tuple[Path, dict[str, int | str]]:
    require(isinstance(value, Mapping) and set(value) == {"path", "bytes", "sha256"},
            f"{label}: helper metadata is malformed")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}: helper path is missing")
    try:
        helper_path = Path(raw_path).resolve(strict=True)
    except OSError as error:
        fail(f"{label}: helper path is missing: {error}")
    require(helper_path.parent == root and helper_path.name in PROFILE_HELPERS,
            f"{label}: helper is not allowlisted")
    metadata = _meta(helper_path, label)
    require(dict(value) == metadata, f"{label}: helper hash changed")
    return helper_path, _descriptor(root, helper_path, label)


def _failed_profile_disposition(root: Path, directory: Path, value: Mapping[str, Any],
                                helper: Path) -> None:
    """Authenticate the known strace1 tool-rejection attempt without replaying it."""

    attempt = directory.name
    require(attempt == FAILED_PROFILE_ATTEMPT and helper.name == "profile.py",
            "unexpected failed profile attempt")
    require(value.get("status") == "incomplete"
            and value.get("modes") == ["counting", "atomic"],
            "failed profile disposition was silently changed")
    records = value.get("records")
    require(isinstance(records, list) and len(records) == 2,
            "failed profile record inventory differs")
    for mode, record in zip(("counting", "atomic"), records, strict=True):
        label = f"strace-{mode}"
        require(isinstance(record, Mapping)
                and record.get("label") == label and record.get("mode") == mode
                and record.get("status") == "failed"
                and record.get("error") == "profile child failed"
                and record.get("report") is None and record.get("trace") is None,
                f"{label}: failed disposition differs")
        child = directory / label
        require(child.is_dir() and not child.is_symlink(), f"{label}: failed artifact directory is missing")
        require({item.name for item in child.iterdir()} == {
            "started.json", "terminal.json", "stdout.txt", "stderr.txt", "replay-cleanup.json",
        }, f"{label}: failed artifact inventory differs")
        started_path = child / "started.json"
        terminal_path = child / "terminal.json"
        stdout_path = child / "stdout.txt"
        stderr_path = child / "stderr.txt"
        cleanup_path = child / "replay-cleanup.json"
        started = _read_json(started_path, f"{label} failed started receipt")
        terminal = _read_json(terminal_path, f"{label} failed terminal receipt")
        cleanup = _read_json(cleanup_path, f"{label} failed cleanup receipt")
        require(isinstance(started, Mapping) and started.get("schema") ==
                "docx-tail-append-publication-profile-v1"
                and started.get("attempt") == attempt and started.get("label") == label
                and started.get("mode") == mode,
                f"{label}: failed started identity differs")
        require(isinstance(terminal, Mapping)
                and terminal.get("schema") == "docx-tail-append-publication-profile-terminal-v1"
                and terminal.get("attempt") == attempt and terminal.get("label") == label
                and terminal.get("mode") == mode,
                f"{label}: failed terminal identity differs")
        for key in (
            "attempt", "binary", "build", "caller_gate", "cpu", "helper",
            "operation_attribution_claim", "performance_claim", "samples", "scope",
            "source_revision", "version", "warmups",
        ):
            require(started.get(key) == value.get(key),
                    f"{label}: result {key} binding differs")
        require(isinstance(cleanup, Mapping)
                and cleanup.get("schema") == "docx-tail-append-publication-profile-cleanup-v1"
                and cleanup.get("status") == "pass" and cleanup.get("remaining") == [],
                f"{label}: failed cleanup disposition differs")
        for key in (
            "attempt", "label", "mode", "benchmark_argv", "binary", "build", "caller_gate", "cpu",
            "destination_contract", "environment", "helper", "operation_attribution_claim",
            "performance_claim", "replay_dir", "run", "samples", "scope", "source_revision",
            "source_revision_before", "tmpdir", "tool", "tool_executable", "timeout_seconds",
            "version", "warmups",
        ):
            if key in started:
                require(terminal.get(key) == started.get(key), f"{label}: terminal {key} binding differs")
        require(started.get("status") == "running" and terminal.get("status") == "failed"
                and terminal.get("exit_code") == 1 and terminal.get("timed_out") is False
                and terminal.get("termination") is None
                and terminal.get("missing_artifacts") == ["report.json", "strace.raw"],
                f"{label}: failed terminal status differs")
        argv = terminal.get("argv")
        require(isinstance(argv, list) and any(FAILED_PROFILE_SYSCALL in str(item) for item in argv),
                f"{label}: unsupported syscall disposition is missing")
        require(terminal.get("benchmark_argv") == started.get("benchmark_argv"),
                f"{label}: benchmark launch vector differs")
        require(terminal.get("cleanup") == cleanup, f"{label}: cleanup receipt binding differs")
        require(stdout_path.read_text(encoding="utf-8") == "",
                f"{label}: failed profiler emitted benchmark stdout")
        require(stderr_path.read_text(encoding="utf-8") ==
                "/usr/bin/strace: invalid system call 'fstatat'\n",
                f"{label}: unsupported syscall stderr differs")
        require(terminal.get("artifacts") == {
            "replay-cleanup.json": _meta(cleanup_path, f"{label} cleanup artifact"),
            "stderr.txt": _meta(stderr_path, f"{label} stderr artifact"),
            "stdout.txt": _meta(stdout_path, f"{label} stdout artifact"),
        }, f"{label}: failed artifact metadata differs")
        require(record.get("terminal") == _meta(terminal_path, f"{label} terminal artifact"),
                f"{label}: result terminal binding differs")


def _failed_preflight_disposition(root: Path, directory: Path, value: Mapping[str, Any],
                                  helper: Path, *, expected_attempt: str,
                                  expected_helper: str, expected_error: str,
                                  raw_marker: str, expected_source_count: int) -> None:
    """Authenticate a retained v2 preflight failure without replaying it."""

    attempt = directory.name
    require(attempt == expected_attempt and helper.name == expected_helper,
            "unexpected failed preflight attempt")
    require(value.get("status") == "incomplete"
            and value.get("modes") == ["counting", "atomic"],
            "failed preflight disposition was silently changed")
    records = value.get("records")
    require(isinstance(records, list) and len(records) == 2,
            "failed preflight record inventory differs")
    for mode, record in zip(("counting", "atomic"), records, strict=True):
        label = f"strace-{mode}"
        require(isinstance(record, Mapping)
                and record.get("label") == label and record.get("mode") == mode
                and record.get("status") == "failed"
                and record.get("error") == "profile child failed"
                and record.get("report") is None,
                f"{label}: failed preflight disposition differs")
        child = directory / label
        require(child.is_dir() and not child.is_symlink(),
                f"{label}: failed preflight artifact directory is missing")
        require({item.name for item in child.iterdir()} == {
            "started.json", "terminal.json", "stdout.txt", "stderr.txt",
            "strace.raw", "replay-cleanup.json",
        }, f"{label}: failed preflight artifact inventory differs")
        started_path = child / "started.json"
        terminal_path = child / "terminal.json"
        stdout_path = child / "stdout.txt"
        stderr_path = child / "stderr.txt"
        raw_path = child / "strace.raw"
        cleanup_path = child / "replay-cleanup.json"
        started = _read_json(started_path, f"{label} failed preflight started receipt")
        terminal = _read_json(terminal_path, f"{label} failed preflight terminal receipt")
        cleanup = _read_json(cleanup_path, f"{label} failed preflight cleanup receipt")
        require(isinstance(started, Mapping)
                and started.get("schema") == "docx-tail-append-publication-profile-v1"
                and started.get("attempt") == attempt and started.get("label") == label
                and started.get("mode") == mode,
                f"{label}: failed preflight started identity differs")
        require(isinstance(terminal, Mapping)
                and terminal.get("schema") == "docx-tail-append-publication-profile-terminal-v1"
                and terminal.get("attempt") == attempt and terminal.get("label") == label
                and terminal.get("mode") == mode,
                f"{label}: failed preflight terminal identity differs")
        for key in (
            "attempt", "binary", "build", "caller_gate", "cpu", "helper",
            "operation_attribution_claim", "performance_claim", "samples", "scope",
            "source_revision", "version", "warmups",
        ):
            require(started.get(key) == value.get(key),
                    f"{label}: preflight result {key} binding differs")
        require(isinstance(cleanup, Mapping)
                and cleanup.get("schema") == "docx-tail-append-publication-profile-cleanup-v1"
                and cleanup.get("status") == "pass" and cleanup.get("remaining") == [],
                f"{label}: failed preflight cleanup disposition differs")
        for key in (
            "attempt", "label", "mode", "benchmark_argv", "binary", "build", "caller_gate", "cpu",
            "destination_contract", "environment", "helper", "operation_attribution_claim",
            "performance_claim", "replay_dir", "run", "samples", "scope", "source_revision",
            "source_revision_before", "tmpdir", "tool", "tool_executable", "timeout_seconds",
            "version", "warmups",
        ):
            if key in started:
                require(terminal.get(key) == started.get(key),
                        f"{label}: preflight terminal {key} binding differs")
        require(started.get("status") == "running" and terminal.get("status") == "failed"
                and terminal.get("exit_code") == 1 and terminal.get("timed_out") is False
                and terminal.get("termination") is None
                and terminal.get("missing_artifacts") == ["report.json"]
                and terminal.get("report_summary") is None
                and terminal.get("profiler") is None,
                f"{label}: failed preflight terminal status differs")
        require(terminal.get("benchmark_argv") == started.get("benchmark_argv"),
                f"{label}: preflight benchmark launch vector differs")
        for receipt in (started, terminal):
            run = receipt.get("run")
            require(isinstance(run, Mapping) and run.get("source_count") == expected_source_count,
                    f"{label}: preflight source-count binding differs")
        benchmark_argv = started.get("benchmark_argv")
        require(isinstance(benchmark_argv, list)
                and "--source-counts" in benchmark_argv,
                f"{label}: preflight source-count argument is missing")
        source_index = benchmark_argv.index("--source-counts")
        require(source_index + 1 < len(benchmark_argv)
                and benchmark_argv[source_index + 1] == str(expected_source_count),
                f"{label}: preflight source-count argument differs")
        argv = terminal.get("argv")
        require(isinstance(argv, list),
                f"{label}: profiler argv is malformed")
        trace_filters = [item[6:].split(",") for item in argv
                         if isinstance(item, str) and item.startswith("trace=")]
        require(len(trace_filters) == 1
                and "fstatat" not in trace_filters[0]
                and "newfstatat" in trace_filters[0],
                f"{label}: v2 profiler syscall filter differs")
        require(stdout_path.read_text(encoding="utf-8") == "",
                f"{label}: failed preflight emitted benchmark stdout")
        require(stderr_path.read_text(encoding="utf-8") == expected_error,
                f"{label}: preflight error text differs")
        raw_text = raw_path.read_text(encoding="utf-8", errors="replace")
        require(raw_text and raw_marker in raw_text
                and "+++ exited with 1 +++" in raw_text,
                f"{label}: retained preflight trace does not show typed failure")
        require(terminal.get("cleanup") == cleanup,
                f"{label}: preflight cleanup receipt binding differs")
        require(terminal.get("artifacts") == {
            "replay-cleanup.json": _meta(cleanup_path, f"{label} cleanup artifact"),
            "stderr.txt": _meta(stderr_path, f"{label} stderr artifact"),
            "stdout.txt": _meta(stdout_path, f"{label} stdout artifact"),
            "strace.raw": _meta(raw_path, f"{label} raw strace artifact"),
        }, f"{label}: failed preflight artifact metadata differs")
        require(record.get("terminal") == _meta(terminal_path, f"{label} terminal artifact")
                and record.get("trace") == _meta(raw_path, f"{label} raw strace artifact")
                and record.get("report") is None,
                f"{label}: preflight result artifact binding differs")


def _profile_binding(root: Path) -> list[dict[str, Any]]:
    profile_root = root / "profiles"
    require(profile_root.is_dir() and not profile_root.is_symlink(), "profile evidence root is missing")
    attempts: list[dict[str, Any]] = []
    modules: dict[Path, Any] = {}
    successful: list[str] = []
    failed_seen: set[str] = set()
    for directory in sorted(profile_root.iterdir(), key=lambda item: item.name):
        require(directory.is_dir() and not directory.is_symlink(),
                f"profile attempt is not a directory: {directory}")
        attempt = directory.name
        result_path = directory / "result.json"
        value = _read_json(result_path, f"profile {attempt} result")
        require(isinstance(value, Mapping)
                and value.get("schema") == "docx-tail-append-publication-profile-result-v1"
                and value.get("version") == 1 and value.get("attempt") == attempt,
                f"profile {attempt} result schema/identity differs")
        helper_path, helper_binding = _profile_helper_binding(root, value.get("helper"),
                                                              f"profile {attempt} helper")
        modes = value.get("modes")
        require(isinstance(modes, list) and len(modes) == len(PROFILE_MODES)
                and set(modes) == PROFILE_MODES,
                f"profile {attempt}: both publication routes are required")
        status = value.get("status")
        disposition = "pass"
        if attempt == FAILED_PROFILE_ATTEMPT:
            require(status != "pass", "known failed strace1 attempt was silently marked successful")
            _failed_profile_disposition(root, directory, value, helper_path)
            failed_seen.add(attempt)
            disposition = "failed-tool-invocation"
        elif attempt == FAILED_PREFLIGHT_ATTEMPT:
            require(status != "pass", "known failed strace2 attempt was silently marked successful")
            _failed_preflight_disposition(
                root, directory, value, helper_path,
                expected_attempt=FAILED_PREFLIGHT_ATTEMPT,
                expected_helper=FAILED_PREFLIGHT_HELPER,
                expected_error=FAILED_PREFLIGHT_ERROR,
                raw_marker="XML Events limit 6807808 exceeds hard ceiling 4000000",
                expected_source_count=8192,
            )
            failed_seen.add(attempt)
            disposition = "failed-preflight"
        elif attempt == FAILED_CLI_ATTEMPT:
            require(status != "pass", "known failed strace3 attempt was silently marked successful")
            _failed_preflight_disposition(
                root, directory, value, helper_path,
                expected_attempt=FAILED_CLI_ATTEMPT,
                expected_helper=FAILED_PREFLIGHT_HELPER,
                expected_error=FAILED_CLI_ERROR,
                raw_marker=FAILED_CLI_RAW_MARKER,
                expected_source_count=4096,
            )
            failed_seen.add(attempt)
            disposition = "failed-cli-validation"
        elif attempt == FAILED_PREFLIGHT4_ATTEMPT:
            require(status != "pass", "known failed strace4 attempt was silently marked successful")
            _failed_preflight_disposition(
                root, directory, value, helper_path,
                expected_attempt=FAILED_PREFLIGHT4_ATTEMPT,
                expected_helper=FAILED_PREFLIGHT_HELPER,
                expected_error=FAILED_PREFLIGHT4_ERROR,
                raw_marker="XML Events limit 6742784 exceeds hard ceiling 4000000",
                expected_source_count=64,
            )
            failed_seen.add(attempt)
            disposition = "failed-preflight"
        else:
            require(status == "pass", f"profile {attempt} is incomplete")
            module = modules.get(helper_path)
            if module is None:
                module = _load_module(f"change0497_{helper_path.stem}_for_seal", helper_path)
                modules[helper_path] = module
            try:
                result = module._verify(attempt)
            except Exception as error:
                fail(f"profile {attempt} verification failed: {error}")
            require(result == 0, f"profile {attempt} did not verify")
            successful.append(attempt)
        files: dict[str, dict[str, int | str]] = {}
        for child in sorted(directory.rglob("*"), key=lambda item: str(item)):
            if child.is_dir():
                require(not child.is_symlink(), f"profile contains symlinked directory: {child}")
                continue
            state = child.lstat()
            require(stat.S_ISREG(state.st_mode) and not stat.S_ISLNK(state.st_mode),
                    f"profile contains special path: {child}")
            files[child.relative_to(root).as_posix()] = _meta(child, f"profile {attempt} file")
        attempts.append({"attempt": attempt, "status": status, "modes": list(modes),
                         "disposition": disposition, "helper": helper_binding,
                         "result": _descriptor(root, result_path, f"profile {attempt} result"),
                         "files": files})
    require(failed_seen == FAILED_PROFILE_ATTEMPTS,
            "known failed profile custody is missing")
    require(SUCCESSFUL_PROFILE_ATTEMPT in successful,
            "required complete strace5 profile attempt was not retained")
    return attempts


def _early_target_binding(root: Path) -> dict[str, Any]:
    """Bind the authenticated pre-seal Cargo-target removal receipt.

    The target may already be absent when the final cleanup validator runs.
    This receipt records why that earlier removal was authorized and prevents
    a later seal from silently treating an unaccounted-for missing target as
    ordinary cleanup output.
    """

    path = root / EARLY_TARGET_CLEANUP
    value = _read_json(path, "early target cleanup")
    require(isinstance(value, Mapping)
            and set(value) == {"path", "reason", "allocated_bytes", "free_before",
                                "free_after", "completed_ns", "active_build_refs"},
            "early target cleanup schema differs")
    require(value.get("path") == str(TARGET.resolve(strict=False)),
            "early target cleanup path differs")
    require(isinstance(value.get("reason"), str) and value["reason"].strip(),
            "early target cleanup reason is missing")
    for key in ("allocated_bytes", "free_before", "free_after", "completed_ns"):
        require(type(value.get(key)) is int and value[key] >= 0,
                f"early target cleanup {key} is malformed")
    require(value["allocated_bytes"] > 0 and value["completed_ns"] > 0,
            "early target cleanup counters are empty")
    require(value.get("active_build_refs") == [],
            "early target cleanup has active build references")
    return {
        "descriptor": _descriptor(root, path, "early target cleanup"),
        "path": value["path"],
        "reason": value["reason"],
        "allocated_bytes": value["allocated_bytes"],
        "free_before": value["free_before"],
        "free_after": value["free_after"],
        "completed_ns": value["completed_ns"],
        "active_build_refs": [],
    }


def _early_fuzz_binding(root: Path) -> dict[str, Any]:
    """Bind the separately removed ASan target before final cleanup."""

    path = root / EARLY_FUZZ_CLEANUP
    value = _read_json(path, "early fuzz target cleanup")
    require(isinstance(value, Mapping)
            and set(value) == {"path", "allocated_bytes", "completed_ns", "reason"},
            "early fuzz target cleanup schema differs")
    expected_path = FUZZ_TEMP / "target"
    require(value.get("path") == str(expected_path.resolve(strict=False)),
            "early fuzz target cleanup path differs")
    require(isinstance(value.get("reason"), str) and value["reason"].strip(),
            "early fuzz target cleanup reason is missing")
    require(type(value.get("allocated_bytes")) is int and value["allocated_bytes"] > 0
            and type(value.get("completed_ns")) is int and value["completed_ns"] > 0,
            "early fuzz target cleanup counters are malformed")
    return {
        "descriptor": _descriptor(root, path, "early fuzz target cleanup"),
        "path": value["path"],
        "allocated_bytes": value["allocated_bytes"],
        "completed_ns": value["completed_ns"],
        "reason": value["reason"],
    }


def _adr_binding(root: Path, repo: Path) -> dict[str, Any]:
    path = root / ADR_REFRESH
    value = _read_json(path, ADR_REFRESH)
    require(isinstance(value, Mapping) and isinstance(value.get("files"), Mapping)
            and value["files"], "ADR refresh inventory is empty")
    actual: dict[str, str] = {}
    for name, digest in sorted(value["files"].items()):
        require(isinstance(name, str) and name.startswith("docs/adr/")
                and ".." not in Path(name).parts, f"unsafe ADR path: {name}")
        require(isinstance(digest, str) and len(digest) == 64
                and all(char in "0123456789abcdef" for char in digest), f"malformed ADR hash: {name}")
        source = repo / name
        _no_symlink_components(source, f"ADR {name}")
        actual_digest = _sha(source)
        require(actual_digest == digest, f"ADR {name}: content changed")
        actual[name] = actual_digest
    return {"refresh": _descriptor(root, path, "ADR refresh"), "files": actual}


def _source_binding(measure: Any, root: Path, repo: Path) -> dict[str, Any]:
    candidate_binding, candidate = measure._candidate_source_binding()
    current: dict[str, dict[str, int | str]] = {}
    for name, expected in sorted(candidate.items()):
        require(isinstance(name, str) and not Path(name).is_absolute()
                and ".." not in Path(name).parts, f"candidate path is unsafe: {name}")
        source = repo / name
        meta = _meta(source, f"candidate source {name}")
        require(meta["bytes"] == expected["bytes"] and meta["sha256"] == expected["sha256"],
                f"candidate source changed: {name}")
        current[name] = meta
    fixture_path = root / "fixture-inputs.json"
    fixture = measure._fixture_inputs_binding()
    fixture_value = _read_json(fixture_path, "fixture inputs")
    require(isinstance(fixture_value, Mapping) and isinstance(fixture_value.get("files"), Mapping),
            "fixture input inventory is malformed")
    verification_path = root / "fixture-verification.json"
    verification_value = _read_json(verification_path, "fixture verification")
    require(isinstance(verification_value, Mapping)
            and set(verification_value) == {"verified_utc", "manifest_sha256", "phases"},
            "fixture verification schema differs")
    _timestamp(verification_value.get("verified_utc"), "fixture verification timestamp")
    require(verification_value.get("manifest_sha256") == _sha(fixture_path),
            "fixture verification manifest binding differs")
    phases = verification_value.get("phases")
    require(isinstance(phases, Mapping) and set(phases) == {"before", "after"},
            "fixture verification phase inventory differs")
    phase_summary: dict[str, Any] = {}
    for phase in ("before", "after"):
        record = phases[phase]
        require(isinstance(record, Mapping)
                and set(record) == {"source_root", "checked_files", "checked_bytes", "status"},
                f"fixture verification {phase} schema differs")
        require(record.get("source_root") == str((TEMP / phase).resolve(strict=False))
                and type(record.get("checked_files")) is int and record["checked_files"] > 0
                and type(record.get("checked_bytes")) is int and record["checked_bytes"] >= 0
                and record.get("status") == "pass",
                f"fixture verification {phase} is incomplete")
        phase_summary[phase] = dict(record)
    return {
        "candidate": {"manifest": _descriptor(root, root / "candidate-source.json", "candidate source manifest"),
                      "patch": _descriptor(root, root / "candidate.patch", "candidate patch"),
                      "metadata": candidate_binding, "files": current},
        "fixture": {"descriptor": _descriptor(root, fixture_path, "fixture inputs"),
                    "metadata": fixture, "scope": fixture_value.get("scope"),
                    "file_count": len(fixture_value["files"]),
                    "verification": {"descriptor": _descriptor(root, verification_path,
                                                                  "fixture verification"),
                                     "manifest_sha256": verification_value["manifest_sha256"],
                                     "phases": phase_summary}},
    }


def _reviews_binding(root: Path) -> list[dict[str, int | str]]:
    paths = sorted({path for path in root.rglob("*.md")
                    if path.is_file() and "review" in path.name.lower()})
    require(paths, "review evidence is missing")
    return [_descriptor(root, path, f"review {path.name}") for path in paths]


def _file_binding(root: Path, value: Any, path: Path, label: str) -> dict[str, int | str]:
    """Compare a gate.py descriptor with the current regular file."""

    require(isinstance(value, Mapping) and set(value) == {"path", "bytes", "sha256"},
            f"{label}: file metadata is malformed")
    _no_symlink_components(path, label)
    expected = _meta(path, label)
    require(dict(value) == expected, f"{label}: file metadata changed")
    return {"path": _inside(root, path, label), "bytes": expected["bytes"],
            "sha256": expected["sha256"]}


def _gate_receipt_binding(root: Path, name: str, command: list[str], value: Any) -> dict[str, Any]:
    label = f"gate receipt {name}"
    require(isinstance(value, Mapping), f"{label}: receipt metadata is malformed")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}: receipt path is missing")
    path = Path(raw_path)
    if not path.is_absolute():
        path = root / path
    path = path.resolve(strict=True)
    validation = (root / "validation").resolve(strict=True)
    require(path.parent == validation and path.suffix == ".json" and path.stem,
            f"{label}: receipt path is outside validation custody")
    receipt_meta = _descriptor(root, path, label)
    require(value.get("bytes") == receipt_meta["bytes"] and value.get("sha256") == receipt_meta["sha256"],
            f"{label}: receipt hash changed")
    receipt = _read_json(path, label)
    require(isinstance(receipt, Mapping) and set(receipt) == GATE_RECEIPT_KEYS,
            f"{label}: gate.py receipt fields differ")
    require(receipt.get("argv") == command, f"{label}: command differs")
    require(receipt.get("cwd") == str((TEMP / "after").resolve(strict=False)),
            f"{label}: cwd differs")
    require(receipt.get("environment") == GATE_ENVIRONMENT, f"{label}: environment differs")
    require(receipt.get("exit_code") == 0 and receipt.get("source_unchanged") is True
            and receipt.get("timed_out") is False and receipt.get("termination") is None,
            f"{label}: terminal did not pass")
    for field in ("started_ns", "finished_ns", "pid", "timeout_seconds"):
        require(type(receipt.get(field)) is int and receipt[field] > 0,
                f"{label}.{field}: positive integer required")
    require(receipt["finished_ns"] > receipt["started_ns"], f"{label}: terminal chronology differs")
    require(receipt["timeout_seconds"] == 3600, f"{label}: timeout differs")

    validation = root / "validation"
    artifact = path.stem
    driver = validation.parent / "gate.py"
    source_manifest = validation / f"{artifact}.source.json"
    stdout = validation / f"{artifact}.stdout"
    stderr = validation / f"{artifact}.stderr"
    require(Path(receipt["driver"]["path"]).resolve(strict=False) == driver.resolve(strict=True),
            f"{label}: driver path differs")
    driver_binding = _file_binding(root, receipt["driver"], driver, f"{label}.driver")
    source_binding = _file_binding(root, receipt["source_manifest"], source_manifest,
                                   f"{label}.source_manifest")
    stdout_binding = _file_binding(root, receipt["stdout"], stdout, f"{label}.stdout")
    stderr_binding = _file_binding(root, receipt["stderr"], stderr, f"{label}.stderr")
    source_value = _read_json(source_manifest, f"{label}.source_manifest")
    require(isinstance(source_value, Mapping) and source_value,
            f"{label}: source manifest is empty")

    started_path = validation / f"{artifact}.started.json"
    started = _read_json(started_path, f"{label}.started")
    require(isinstance(started, Mapping)
            and set(started) == {"argv", "cwd", "driver", "environment", "source_manifest",
                                 "started_ns", "timeout_seconds"},
            f"{label}: started receipt fields differ")
    require(started.get("argv") == command and started.get("cwd") == receipt["cwd"]
            and started.get("environment") == receipt["environment"]
            and started.get("started_ns") == receipt["started_ns"]
            and started.get("timeout_seconds") == receipt["timeout_seconds"],
            f"{label}: started receipt differs")
    _file_binding(root, started["driver"], driver, f"{label}.started.driver")
    _file_binding(root, started["source_manifest"], source_manifest,
                  f"{label}.started.source_manifest")
    return {
        "path": receipt_meta["path"], "bytes": receipt_meta["bytes"], "sha256": receipt_meta["sha256"],
        "argv": list(command), "started_ns": receipt["started_ns"],
        "finished_ns": receipt["finished_ns"], "pid": receipt["pid"],
        "driver": driver_binding, "source_manifest": source_binding,
        "stdout": stdout_binding, "stderr": stderr_binding,
        "exit_code": 0, "source_unchanged": True, "timed_out": False,
        "termination": None, "timeout_seconds": receipt["timeout_seconds"],
    }


def _named_gate_receipts(root: Path, value: Mapping[str, Any],
                         commands: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    raw = value.get("receipts")
    require(isinstance(raw, Mapping) and raw, "final gates named receipt inventory is missing")
    names = set(commands)
    require(names == MANDATORY_GATE_NAMES or names == MANDATORY_GATE_NAMES | OPTIONAL_GATE_NAMES,
            "final gates command inventory differs")
    require(set(raw) == names, "final gates named receipt inventory differs from commands")
    return {name: _gate_receipt_binding(root, name, list(commands[name]), raw[name])
            for name in sorted(names)}


def _gates_binding(root: Path) -> dict[str, Any]:
    path = root / FINAL_GATES
    value = _read_json(path, FINAL_GATES)
    require(isinstance(value, Mapping), "final gates receipt is not an object")
    require("scope" in value and value["scope"], "final gates scope is missing")
    commands = value.get("commands")
    require(isinstance(commands, Mapping) and commands, "final gates command inventory is empty")
    require(all(isinstance(name, str) and name and isinstance(command, list) and command
                and all(isinstance(item, str) and item for item in command)
                for name, command in commands.items()), "final gates command inventory is malformed")
    receipts = _named_gate_receipts(root, value, commands)
    historical = []
    for candidate in sorted(root.glob("gates-*.json")):
        historical.append(_descriptor(root, candidate, f"gate receipt {candidate.name}"))
    final = {"descriptor": _descriptor(root, path, "final gates"),
             "schema": value.get("schema"), "version": value.get("version"),
             "commands": sorted(commands), "scope": value["scope"], "receipts": receipts}
    return {"final": final,
            "historical": historical}


def _cleanup_binding(root: Path, cleanup: Any) -> dict[str, Any]:
    receipts = cleanup._coerce_build_receipts()
    verification = cleanup.verify(root=root, temp=TEMP, target=TARGET, fuzz_temp=FUZZ_TEMP,
                                  build_receipts=receipts)
    require(verification.get("status") == "pass", "cleanup verification did not pass")
    receipt_path = root / "cleanup.json"
    receipt = _read_json(receipt_path, "cleanup receipt")
    require(isinstance(receipt, Mapping) and receipt.get("schema") == cleanup.SCHEMA
            and receipt.get("version") == cleanup.VERSION and receipt.get("status") == "pass",
            "cleanup receipt schema/status differs")
    return {"descriptor": _descriptor(root, receipt_path, "cleanup receipt"),
            "scope": receipt.get("scope"), "removed": receipt.get("removed_paths"),
            "retained_binaries": receipt.get("retained_binaries"),
            "temporary_scratch_remaining": receipt.get("temporary_scratch_remaining")}


def _evidence(root: Path) -> dict[str, Any]:
    measure = _load_module("change0497_measure_for_seal", root / "measure.py")
    cleanup = _load_module("change0497_cleanup_for_seal", root / "cleanup.py")
    protocol, protocol_hash, builds = measure._load_protocol()
    require(isinstance(protocol, Mapping) and isinstance(builds, Mapping),
            "measurement protocol/build inventory is malformed")
    captures = _formal_binding(measure, protocol, protocol_hash, builds, root)
    resume_binding = _resume_binding(root, measure, protocol, protocol_hash, builds,
                                     captures["formal"])
    profiles = _profile_binding(root)
    early_target = _early_target_binding(root)
    early_fuzz = _early_fuzz_binding(root)
    cleanup_binding = _cleanup_binding(root, cleanup)
    fuzz = cleanup._validate_fuzz_receipts(root, FUZZ_TEMP, present=False,
                                           retained_root=TEMP / "retained")
    return {
        "protocol": {"descriptor": _descriptor(root, measure.PROTOCOL_FILE, "frozen protocol"),
                     "sha256": protocol_hash},
        "builds": measure._builds_binding(dict(builds)),
        "formal": captures["formal"],
        "pilot": captures["pilot"],
        "analysis": captures["analysis"],
        "verification": captures["verification"],
        "resume": resume_binding,
        "profiles": profiles,
        "early_target_cleanup": early_target,
        "early_fuzz_target_cleanup": early_fuzz,
        "fuzz": fuzz,
        "cleanup": cleanup_binding,
        "source": _source_binding(measure, root, REPO),
        "adr": _adr_binding(root, REPO),
        "reviews": _reviews_binding(root),
        "gates": _gates_binding(root),
        "protected_primary": _descriptor(root, root / PROTECTED_PRIMARY, "protected primary manifest"),
    }


def _manifest(root: Path) -> dict[str, Any]:
    require(not (root / SEAL_NAME).exists(), f"refusing to replace existing seal: {root / SEAL_NAME}")
    return {"schema": SCHEMA, "version": VERSION, "root": ".", "sealed_utc": _now(),
            "excluded": EXCLUDED, **_evidence(root), "files": _inventory(root)}


def _write_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace seal: {path}")
    try:
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
    except OSError as error:
        fail(f"cannot write seal: {error}")


def _verify_manifest(root: Path, manifest: Any) -> None:
    require(isinstance(manifest, Mapping), "seal manifest is not an object")
    expected = {"schema", "version", "root", "sealed_utc", "excluded", "protocol", "builds",
                "formal", "pilot", "analysis", "verification", "resume", "profiles", "fuzz", "cleanup",
                "early_target_cleanup", "early_fuzz_target_cleanup", "source", "adr", "reviews", "gates",
                "protected_primary", "files"}
    require(set(manifest) == expected, "seal manifest fields differ")
    require(manifest.get("schema") == SCHEMA and manifest.get("version") == VERSION
            and manifest.get("root") == "." and manifest.get("excluded") == EXCLUDED,
            "seal schema/root policy differs")
    _timestamp(manifest.get("sealed_utc"), "seal timestamp")
    require(manifest.get("files") == _inventory(root), "sealed evidence inventory or hash changed")
    current = _evidence(root)
    for key, value in current.items():
        require(manifest.get(key) == value, f"sealed {key} binding changed")


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("seal", "verify"), nargs="?", default="verify")
    parser.add_argument("--root", type=Path, default=ROOT, help=argparse.SUPPRESS)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        root = args.root.resolve(strict=True)
        require(root == ROOT, "0497 seal is restricted to its fixed evidence root")
        seal_path = root / SEAL_NAME
        if args.command == "seal":
            value = _manifest(root)
            _write_exclusive(seal_path, value)
            print(json.dumps({"schema": SCHEMA, "status": "sealed", "files": len(value["files"])}, sort_keys=True))
            return 0
        manifest = _read_json(seal_path, "seal manifest")
        _verify_manifest(root, manifest)
        print(json.dumps({"schema": VERIFY_SCHEMA, "version": VERSION, "status": "pass",
                          "verified_utc": _now(), "files": len(manifest["files"])}, sort_keys=True))
        return 0
    except (SealError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"seal.py: error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
