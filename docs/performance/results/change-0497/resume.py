#!/usr/bin/env python3
"""Fail-closed continuation for the interrupted 0497 formal capture.

The original coordinator stopped while writing the ordinal-191 cleanup
receipt after the disk filled.  This driver archives that raw, incomplete
directory without inventing a terminal receipt, validates ordinals 0--190
through the frozen 0497 validators, and can then run only ordinals 191--287
in the frozen order under the same CPU lock.  It never replaces a capture
directory or a resume receipt.

The normal ``measure.py`` and frozen ``protocol.json`` are inputs.  This file
does not modify either one.  ``--check`` is read-only; ``--run`` is the only
mode that archives the interrupted raw directory and launches the remaining
children.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import sys
from typing import Any, Mapping


sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parent
ATTEMPT = "formal1"
START_ORDINAL = 191
FORMAL_ROOT = ROOT / "captures" / ATTEMPT
INTERRUPTED_ROOT = ROOT / "interrupted" / "formal1-enospc"
RESUME_ROOT = ROOT / "resume" / ATTEMPT
OBSERVATION_FILE = ROOT / "interruption-observation.json"
MEASURE_FILE = ROOT / "measure.py"
PROTOCOL_FILE = ROOT / "protocol.json"
BUILDS_FILE = ROOT / "builds.json"
PROVENANCE_FILE = ROOT / "provenance.json"
MACHINE_FILE = ROOT / "machine.json"
CANDIDATE_FILE = ROOT / "candidate-source.json"
TEST_FILE = ROOT / "test_resume.py"
RESUME_SCHEMA = "docx-replayable-tail-publication-resume-v1"
INTERRUPTED_SCHEMA = "docx-replayable-tail-publication-interrupted-v1"
RESUME_START_SCHEMA = "docx-replayable-tail-publication-resume-start-v1"
RESUME_TERMINAL_SCHEMA = "docx-replayable-tail-publication-resume-terminal-v1"
REQUIRED_RAW_FILES = ("started.json", "stdout.txt", "stderr.txt", "resource.txt",
                      "replay-cleanup.json")


class ResumeError(RuntimeError):
    """A continuation custody or safety precondition failed."""


def fail(message: str) -> None:
    raise ResumeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat(
        timespec="microseconds").replace("+00:00", "Z")


def _sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"regular file required: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _meta(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"regular file required: {path}")
    return {"path": str(path.resolve()), "bytes": path.stat().st_size,
            "sha256": _sha256(path)}


def _read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"JSON file required: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        fail(f"invalid JSON {path}: {error}")
    raise AssertionError("unreachable")


def _write_new(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(),
            f"refusing to replace existing receipt: {path}")
    missing: list[Path] = []
    current = path.parent
    while not current.exists() and not current.is_symlink():
        missing.append(current)
        current = current.parent
    require(current.is_dir() and not current.is_symlink(),
            f"receipt parent is not a regular directory: {current}")
    for directory in reversed(missing):
        directory.mkdir()
        require(directory.is_dir() and not directory.is_symlink(),
                f"receipt parent changed while creating: {directory}")
    try:
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError as error:
        raise ResumeError(f"refusing to replace existing receipt: {path}") from error


def _load_measure() -> Any:
    """Load the frozen driver without putting its directory on ``sys.path``."""

    name = "measure0497_resume_unique"
    spec = importlib.util.spec_from_file_location(name, MEASURE_FILE)
    require(spec is not None and spec.loader is not None,
            f"cannot load measurement driver: {MEASURE_FILE}")
    module = importlib.util.module_from_spec(spec)
    require(name not in sys.modules, f"measurement module name is already loaded: {name}")
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(name, None)
        raise
    return module


def _load_frozen() -> tuple[Any, dict[str, Any], str, dict[str, dict[str, Any]]]:
    measure = _load_measure()
    protocol, protocol_hash, builds = measure._load_protocol()
    require(protocol["schema"] == measure.PROTOCOL_SCHEMA,
            "frozen protocol schema differs")
    require(protocol["capture_root"] == str(measure.CAPTURE_ROOT),
            "frozen capture root differs")
    require(protocol["temporary_root"] == str(measure.TEMP),
            "frozen temporary root differs")
    return measure, protocol, protocol_hash, builds


def _formal_specs(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    raw = protocol.get("formal_runs")
    require(isinstance(raw, list) and len(raw) == 288,
            "frozen formal inventory must contain 288 runs")
    result: list[dict[str, Any]] = []
    labels: set[str] = set()
    for ordinal, item in enumerate(raw):
        require(isinstance(item, Mapping), f"formal run {ordinal} is malformed")
        spec = dict(item)
        require(spec.get("ordinal") == ordinal,
                f"formal ordinal {ordinal} is not frozen at its position")
        label = spec.get("label")
        require(isinstance(label, str) and label and label not in labels,
                f"formal label {label!r} is missing or duplicated")
        labels.add(label)
        result.append(spec)
    return result


def remaining_specs(protocol: Mapping[str, Any], start: int = START_ORDINAL) -> list[dict[str, Any]]:
    """Return the exact frozen suffix, rejecting any ordinal gap or reorder."""

    require(start == START_ORDINAL, "this continuation only supports ordinal 191")
    specs = _formal_specs(protocol)
    require(0 < start < len(specs), "continuation ordinal is outside the formal inventory")
    for ordinal, spec in enumerate(specs):
        require(spec["ordinal"] == ordinal,
                f"formal inventory has an ordinal gap at {ordinal}")
    return specs[start:]


def launch_specs(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Bind the fixed attempt to each exact frozen suffix specification."""

    result = [dict(spec, attempt=ATTEMPT) for spec in remaining_specs(protocol)]
    require([spec["ordinal"] for spec in result] == list(range(START_ORDINAL, 288)),
            "resume suffix ordinals are not contiguous")
    require(all(spec.get("attempt") == ATTEMPT for spec in result),
            "resume suffix attempt binding is missing")
    return result


def raw_archive_name(name: str) -> str:
    """Map a live artifact basename to a cleanup-invisible raw basename."""

    require(name in REQUIRED_RAW_FILES, f"unexpected interrupted artifact: {name}")
    return f"{name}.raw"


def _tree_inventory(root: Path) -> list[dict[str, Any]]:
    """Describe files and directories without treating raw JSON as receipts."""

    require(root.is_dir() and not root.is_symlink(), f"directory required: {root}")
    result: list[dict[str, Any]] = []
    for path in sorted(root.rglob("*")):
        require(not path.is_symlink(), f"symlink in interrupted tree: {path}")
        relative = path.relative_to(root).as_posix()
        if path.is_dir():
            result.append({"path": str(path), "relative": relative, "kind": "directory"})
        elif path.is_file():
            result.append({"path": str(path), "relative": relative, "kind": "file",
                           "bytes": path.stat().st_size, "sha256": _sha256(path)})
        else:
            fail(f"special file in interrupted tree: {path}")
    return result


def _prefix_receipt(measure: Any, protocol: Mapping[str, Any], builds: Mapping[str, dict[str, Any]],
                    entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for entry in entries:
        directory = Path(entry["directory"])
        result.append({
            "ordinal": entry["spec"]["ordinal"],
            "label": entry["spec"]["label"],
            "started": _meta(directory / "started.json"),
            "terminal": _meta(directory / "terminal.json"),
            "report": _meta(directory / "report.json"),
            "resource": _meta(directory / "resource.txt"),
            "replay_cleanup": _meta(directory / "replay-cleanup.json"),
        })
    return result


def _partial_spec(measure: Any, protocol: Mapping[str, Any],
                  builds: Mapping[str, dict[str, Any]],
                  spec: Mapping[str, Any], directory: Path) -> dict[str, Any]:
    require(directory.is_dir() and not directory.is_symlink(),
            f"interrupted capture directory is missing: {directory}")
    names = {path.name for path in directory.iterdir()}
    require(names == set(REQUIRED_RAW_FILES),
            f"interrupted capture artifact set differs: {sorted(names)}")
    for name in REQUIRED_RAW_FILES:
        path = directory / name
        require(path.is_file() and not path.is_symlink(),
                f"interrupted artifact is not a regular file: {path}")
    started = _read_json(directory / "started.json")
    require(isinstance(started, Mapping) and started.get("status") == "running",
            "interrupted started receipt is not a running capture")
    run = started.get("run")
    require(isinstance(run, Mapping), "interrupted started run metadata is missing")
    for key in ("ordinal", "repeat", "phase", "publication", "role", "arm", "workload",
                "route", "input_mode", "samples", "warmups", "label"):
        require(run.get(key) == spec.get(key),
                f"interrupted run field differs at {key}: {run.get(key)!r} != {spec.get(key)!r}")
    require(started.get("attempt") == ATTEMPT,
            "interrupted attempt differs")
    require(started.get("protocol") == {"path": str(measure.PROTOCOL_FILE),
                                         "sha256": _sha256(measure.PROTOCOL_FILE)},
            "interrupted protocol binding differs")
    require(started.get("driver") == protocol["driver"],
            "interrupted measurement-driver binding differs")
    build = builds[measure._build_key(spec["phase"], spec["role"])]
    require(started.get("build") == measure._build_binding(build),
            "interrupted build binding differs")
    require(started.get("binary") == build["binary"]
            and started.get("source_manifest") == build["source_manifest"],
            "interrupted binary/source binding differs")
    expected_private = measure._run_root(ATTEMPT, spec["label"])
    require(Path(str(started.get("private_root"))).resolve() == expected_private.resolve(),
            "interrupted private-root binding differs")
    expected_tmp = expected_private / "tmp"
    require(Path(str(started.get("tmpdir"))).resolve() == expected_tmp.resolve(),
            "interrupted tmpdir binding differs")
    require(started.get("replay_dir") is None,
            "interrupted latency run unexpectedly has a replay directory")
    expected_argv = measure._command(spec, build, directory / "report.json",
                                     directory / "resource.txt", None, None)
    require(started.get("argv") == expected_argv,
            "interrupted command binding differs")
    require(expected_private.is_dir() and not expected_private.is_symlink(),
            f"interrupted private root is missing: {expected_private}")
    private_inventory = _tree_inventory(expected_private)
    require(private_inventory == [{"path": str(expected_tmp), "relative": "tmp",
                                  "kind": "directory"}],
            "interrupted private root is not the recorded empty tmp structure")
    return {"started": dict(started), "private_inventory": private_inventory,
            "artifact_inventory": _tree_inventory(directory)}


def validate_prefix(measure: Any, protocol: Mapping[str, Any], protocol_hash: str,
                    builds: Mapping[str, dict[str, Any]]) -> dict[str, Any]:
    """Validate the passing prefix, chronology, and one raw interrupted child."""

    specs = _formal_specs(protocol)
    prefix_specs = [dict(item, attempt=ATTEMPT) for item in specs[:START_ORDINAL]]
    failed_spec = dict(specs[START_ORDINAL], attempt=ATTEMPT)
    require(FORMAL_ROOT.is_dir() and not FORMAL_ROOT.is_symlink(),
            f"formal capture root is missing: {FORMAL_ROOT}")
    expected_names = {item["label"] for item in prefix_specs} | {failed_spec["label"]}
    actual_directories = {path.name for path in FORMAL_ROOT.iterdir()
                          if path.is_dir() and not path.is_symlink()}
    require(all(path.is_dir() and not path.is_symlink()
                for path in FORMAL_ROOT.iterdir()),
            "formal capture root contains a non-directory entry")
    require(actual_directories == expected_names,
            "formal namespace must contain only the passing prefix and ordinal 191 raw child")
    entries: list[dict[str, Any]] = []
    for spec in prefix_specs:
        directory = FORMAL_ROOT / spec["label"]
        started = measure._read_json(directory / "started.json")
        terminal = measure._read_json(directory / "terminal.json")
        build = builds[measure._build_key(spec["phase"], spec["role"])]
        measure._validate_terminal(started, terminal, spec, build, protocol, protocol_hash, directory)
        raw = measure._read_json(directory / "report.json")
        measure._validate_report(directory / "report.json", spec, build, started["argv"],
                                 started["input_file"],
                                 Path(started["replay_dir"]) if started["replay_dir"] else None,
                                 Path(started["tmpdir"]))
        resource = measure._resource(directory / "resource.txt")
        entries.append({"spec": spec, "directory": directory, "started": started,
                        "terminal": terminal, "report": raw, "resource": resource,
                        "started_at": measure._timestamp(started["started_utc"],
                                                         f"{directory}.started_utc"),
                        "finished_at": measure._timestamp(terminal["finished_utc"],
                                                           f"{directory}.finished_utc")})
    ordered = measure._chronological_entries(entries, prefix_specs, ATTEMPT)
    partial_directory = FORMAL_ROOT / failed_spec["label"]
    partial = _partial_spec(measure, protocol, builds, failed_spec, partial_directory)
    partial_started_at = measure._timestamp(partial["started"]["started_utc"],
                                            f"{partial_directory}.started_utc")
    require(partial_started_at >= ordered[-1]["finished_at"],
            "interrupted ordinal 191 started before passing ordinal 190 finished")
    return {"specs": specs, "prefix_specs": prefix_specs, "failed_spec": failed_spec,
            "entries": ordered, "partial_directory": partial_directory,
            "partial": partial, "partial_started_at": partial_started_at}


def _input_bindings(measure: Any) -> dict[str, dict[str, Any]]:
    files = {"measure.py": MEASURE_FILE, "protocol.json": PROTOCOL_FILE,
             "builds.json": BUILDS_FILE, "provenance.json": PROVENANCE_FILE,
             "machine.json": MACHINE_FILE, "candidate-source.json": CANDIDATE_FILE,
             "interruption-observation.json": OBSERVATION_FILE,
             "resume.py": Path(__file__).resolve(), "test_resume.py": TEST_FILE}
    result = {}
    for label, path in files.items():
        result[label] = _meta(path)
    return result


def _archive_interrupted(preflight: Mapping[str, Any], inputs: Mapping[str, Any]) -> dict[str, Any]:
    require(not INTERRUPTED_ROOT.exists() and not INTERRUPTED_ROOT.is_symlink(),
            f"refusing to replace interrupted archive: {INTERRUPTED_ROOT}")
    partial = Path(preflight["partial_directory"])
    started = preflight["partial"]["started"]
    private = Path(str(started["private_root"]))
    require(partial.is_dir() and private.is_dir(),
            "interrupted capture/private roots changed before archive")
    capture_inventory = _tree_inventory(partial)
    private_inventory = _tree_inventory(private)
    archive_capture = INTERRUPTED_ROOT / "capture"
    archive_private = INTERRUPTED_ROOT / "private-root"
    INTERRUPTED_ROOT.mkdir(parents=True, exist_ok=False)
    archive_capture.mkdir(exist_ok=False)
    archive_private.parent.mkdir(parents=True, exist_ok=True)
    raw_files: list[dict[str, Any]] = []
    for item in sorted(partial.iterdir()):
        require(item.is_file() and not item.is_symlink(),
                f"interrupted capture contains unexpected non-file: {item}")
        before = _meta(item)
        target = archive_capture / raw_archive_name(item.name)
        shutil.move(str(item), str(target))
        after = _meta(target)
        require(after["bytes"] == before["bytes"] and after["sha256"] == before["sha256"],
                f"raw interrupted artifact changed while archiving: {item.name}")
        raw_files.append({"original": before, "archive": after,
                          "original_name": item.name})
    partial.rmdir()
    private_before = _tree_inventory(private)
    shutil.move(str(private), str(archive_private))
    private_after = _tree_inventory(archive_private)
    require([{k: v for k, v in item.items() if k not in ("path",)}
              for item in private_before]
             == [{k: v for k, v in item.items() if k not in ("path",)}
                 for item in private_after],
             "private interrupted scratch changed while archiving")
    observation = _read_json(OBSERVATION_FILE)
    require(isinstance(observation, Mapping)
            and observation.get("schema") == "docx-publication-interruption-observation-v1",
            "interruption observation receipt schema differs")
    require(observation.get("completed_terminal_children") == START_ORDINAL
            and observation.get("incomplete_label") == preflight["failed_spec"]["label"]
            and observation.get("child_exit_code", "").startswith("unknown"),
            "interruption observation does not bind ordinal 191 unknown child exit")
    receipt = {
        "schema": INTERRUPTED_SCHEMA,
        "version": 1,
        "status": "archived_incomplete",
        "attempt": ATTEMPT,
        "ordinal": START_ORDINAL,
        "label": preflight["failed_spec"]["label"],
        "scope": "raw interrupted evidence only; no terminal receipt or measurement is reconstructed",
        "driver_error": {
            "kind": "ENOSPC",
            "message": observation.get("observed_error"),
            "coordinator_exit_code": observation.get("coordinator_exit_code"),
            "child_exit_code": "unknown",
            "scope": "the child stopped before report/resource/terminal completion",
        },
        "protocol": inputs["protocol.json"],
        "measure_driver": inputs["measure.py"],
        "builds": inputs["builds.json"],
        "observation": inputs["interruption-observation.json"],
        "original_capture_directory": str(partial),
        "archived_capture_directory": str(archive_capture),
        "original_private_root": str(private),
        "archived_private_root": str(archive_private),
        "capture_files": raw_files,
        "capture_inventory_before": capture_inventory,
        "private_inventory_before": private_inventory,
        "private_inventory_after": private_after,
        "inputs": dict(inputs),
    }
    _write_new(INTERRUPTED_ROOT / "interrupted.json", receipt)
    receipt["receipt"] = _meta(INTERRUPTED_ROOT / "interrupted.json")
    return receipt


def _resume_start(measure: Any, protocol: Mapping[str, Any], protocol_hash: str,
                  builds: Mapping[str, dict[str, Any]], preflight: Mapping[str, Any],
                  archive: Mapping[str, Any]) -> dict[str, Any]:
    inputs = _input_bindings(measure)
    inputs["interrupted.json"] = _meta(INTERRUPTED_ROOT / "interrupted.json")
    return {
        "schema": RESUME_START_SCHEMA,
        "version": 1,
        "status": "running",
        "attempt": ATTEMPT,
        "started_utc": _now(),
        "exit_code": None,
        "start_ordinal": START_ORDINAL,
        "remaining_children": len(preflight["specs"]) - START_ORDINAL,
        "protocol": {"path": str(PROTOCOL_FILE), "sha256": protocol_hash},
        "driver": inputs["resume.py"],
        "test_driver": inputs["test_resume.py"],
        "inputs": inputs,
        "archive": archive["receipt"],
        "prefix_children": len(preflight["entries"]),
        "prefix_last_terminal_utc": preflight["entries"][-1]["terminal"]["finished_utc"],
        "interrupted_started_utc": preflight["partial"]["started"]["started_utc"],
        "prefix_receipts": _prefix_receipt(measure, protocol, builds, preflight["entries"]),
        "remaining_labels": [spec["label"] for spec in preflight["specs"][START_ORDINAL:]],
        "scope": "continuation of frozen formal1 suffix under shared CPU lock",
    }


def _resume_terminal(start_receipt: Mapping[str, Any], status: str,
                     completed: int, error: str | None = None) -> dict[str, Any]:
    result = {
        "schema": RESUME_TERMINAL_SCHEMA,
        "version": 1,
        "status": status,
        "attempt": ATTEMPT,
        "started_utc": start_receipt["started_utc"],
        "finished_utc": _now(),
        "exit_code": 0 if status == "pass" else 1,
        "start_ordinal": START_ORDINAL,
        "remaining_children": start_receipt["remaining_children"],
        "completed_children": completed,
        "start_receipt": _meta(RESUME_ROOT / "resume-start.json"),
        "scope": "resume receipts do not replace the frozen formal protocol or child terminals",
    }
    if error is not None:
        result["error"] = error
    return result


def launch_remaining(measure: Any, protocol: Mapping[str, Any],
                     protocol_hash: str, builds: Mapping[str, dict[str, Any]],
                     specs: list[dict[str, Any]],
                     completed: list[str] | None = None) -> list[str]:
    """Launch a pre-bound suffix and return labels whose terminals completed."""

    expected = launch_specs(protocol)
    require(specs == expected, "resume suffix differs from frozen order")
    completed = [] if completed is None else completed
    for spec in specs:
        require(spec.get("attempt") == ATTEMPT,
                f"resume spec lacks {ATTEMPT} attempt binding")
        measure._launch(spec, builds[measure._build_key(spec["phase"], spec["role"])],
                        protocol, protocol_hash, measure.DEFAULT_TIMEOUT_SECONDS)
        completed.append(spec["label"])
    return completed


def check() -> dict[str, Any]:
    measure, protocol, protocol_hash, builds = _load_frozen()
    preflight = validate_prefix(measure, protocol, protocol_hash, builds)
    return {
        "schema": RESUME_SCHEMA,
        "version": 1,
        "status": "ready",
        "attempt": ATTEMPT,
        "prefix_children": len(preflight["entries"]),
        "failed_ordinal": START_ORDINAL,
        "remaining_children": len(preflight["specs"]) - START_ORDINAL,
        "failed_label": preflight["failed_spec"]["label"],
        "protocol": {"path": str(PROTOCOL_FILE), "sha256": protocol_hash},
        "prefix_last_terminal_utc": preflight["entries"][-1]["terminal"]["finished_utc"],
        "interrupted_started_utc": preflight["partial"]["started"]["started_utc"],
        "scope": "read-only prefix and raw interruption audit; no capture launched",
    }


def run() -> dict[str, Any]:
    """Archive the raw failure and launch the exact frozen suffix."""

    measure, protocol, protocol_hash, builds = _load_frozen()

    def locked() -> dict[str, Any]:
        preflight = validate_prefix(measure, protocol, protocol_hash, builds)
        require(not RESUME_ROOT.exists() and not RESUME_ROOT.is_symlink(),
                f"refusing to reuse resume namespace: {RESUME_ROOT}")
        archive = _archive_interrupted(preflight, _input_bindings(measure))
        start = _resume_start(measure, protocol, protocol_hash, builds, preflight, archive)
        _write_new(RESUME_ROOT / "resume-start.json", start)
        completed_labels: list[str] = []
        try:
            specs = launch_specs(protocol)
            launch_remaining(measure, protocol, protocol_hash, builds, specs,
                             completed=completed_labels)
        except Exception as error:
            terminal = _resume_terminal(start, "failed", len(completed_labels),
                                        f"{type(error).__name__}: {error}")
            _write_new(RESUME_ROOT / "resume-terminal.json", terminal)
            if isinstance(error, ResumeError):
                raise
            raise ResumeError(f"resume child suffix failed: {error}") from error
        terminal = _resume_terminal(start, "pass", len(completed_labels))
        _write_new(RESUME_ROOT / "resume-terminal.json", terminal)
        return terminal

    return measure._cpu_lock(locked)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_mutually_exclusive_group(required=True)
    commands.add_argument("--check", action="store_true",
                          help="validate the passing prefix and raw interruption only")
    commands.add_argument("--run", action="store_true",
                          help="archive the raw interruption and launch the frozen suffix")
    args = parser.parse_args(argv)
    try:
        result = check() if args.check else run()
        json.dump(result, sys.stdout, indent=2, sort_keys=True)
        sys.stdout.write("\n")
        return 0
    except ResumeError as error:
        print(f"0497 resume: FAIL: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
