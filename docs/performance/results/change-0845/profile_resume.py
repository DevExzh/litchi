"""Resume the 0845 Callgrind side population after the v1 reader failure.

This module is additive.  It never edits ``profile.py``, ``reader.py``, the
frozen plan, a retained raw capture, or a completed command.  ``freeze``
records the state that existed after the first four commands (the fourth has
raw output but no profile receipt).  ``resume`` first validates that frozen
state with :mod:`reader_v2`, seals only the missing fourth receipt, and then
executes the remaining twelve plan rows in order.  ``analyze`` validates all
sixteen rows with the same reader before compacting them into the v2 analysis
file.
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path
from typing import Any, Mapping, Sequence

import driver as d
import profile as base
import reader_v2 as r


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PLAN_PATH = P / "profile-plan.json"
PROFILE_DIR = P / "profile-runs"
CAPTURE_PATH = P / "capture-profile.json"
FREEZE_PATH = P / "profile-resume-freeze.json"
RESULT_PATH = P / "profile-resume.json"
ANALYSIS_PATH = P / "profile-analysis-v2.json"
READER_PATH = P / "reader_v2.py"
REPAIR_SPEC_PATH = P / "reader-repair-spec.json"
RUNNER_PATH = P / "profile_resume.py"
FAILED_CAPTURE_DIR = P / "commands" / "profile-capture"
PREFIX_COUNT = 4
PROFILE_FILES = {
    "profile.callgrind",
    "profile.callgrind.1",
    "report.json",
    "started.json",
    "time.txt",
}
PROFILE_FILES_WITH_RECEIPT = PROFILE_FILES | {"receipt.json"}
COMMAND_FILES = {"started.json", "receipt.json", "output.log"}
AUXILIARY_PROFILE_LABELS = {"profile-binding-preflight"}


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read(path: Path) -> Any:
    return d.read(path)


def descriptor(path: Path) -> dict[str, Any]:
    return d.desc(path)


def bind(value: Mapping[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(value, Mapping), f"{label}: descriptor is not an object")
    path = Path(value.get("path", ""))
    actual = descriptor(path)
    require(actual == dict(value), f"{label}: descriptor changed: {path}")
    return actual


def plan() -> dict[str, Any]:
    value = base.plan()
    require(value.get("reports") == 16 and value.get("formal_samples") == 16,
            "resume plan population changed")
    require(value.get("capture_manifest") == "capture-profile.json",
            "resume must retain the frozen capture manifest name")
    require(value.get("analysis_output") == "profile-analysis.json",
            "resume must not rewrite the frozen analysis name")
    return value


def repair_witness() -> dict[str, Any]:
    value = read(REPAIR_SPEC_PATH)
    require(value.get("status") == "reviewed", "reader repair specification is not reviewed")
    require(value.get("official_url") == "https://valgrind.org/docs/manual/cl-format.html",
            "reader repair specification URL changed")
    require(value.get("repair") ==
            "Recognize fe as a current source-file switch and cfl as the cfi callee-file alias; preserve function-identity and cost conservation assertions.",
            "reader repair specification changed")
    installed = value.get("installed_reference")
    regression = value.get("regression_tests")
    require(isinstance(installed, Mapping) and isinstance(regression, Mapping),
            "reader repair specification witnesses are missing")
    bind(installed, "installed Callgrind format reference")
    bind(regression, "reader v2 regression result")
    return {
        "descriptor": descriptor(REPAIR_SPEC_PATH),
        "installed_reference": dict(installed),
        "official_url": value["official_url"],
        "repair": value["repair"],
        "regression_tests": dict(regression),
        "status": value["status"],
    }


def bindings_summary(bindings: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    return {
        leg: {
            "source": bindings[leg]["source"],
            "build": bindings[leg]["build"],
            "freeze": bindings[leg]["freeze"],
            "binary": bindings[leg]["binary"],
        }
        for leg in base.LEGS
    }


def exact_inventory(path: Path, names: set[str], label: str) -> None:
    require(path.is_dir() and not path.is_symlink(), f"{label}: directory is missing")
    require({item.name for item in path.iterdir()} == names,
            f"{label}: artifact inventory changed")
    require(all(item.is_file() and not item.is_symlink() for item in path.iterdir()),
            f"{label}: directory contains a non-file artifact")


def file_inventory(paths: Sequence[Path]) -> list[dict[str, Any]]:
    return [descriptor(path) for path in sorted(paths, key=lambda item: str(item))]


def inventory_from_directory(path: Path) -> list[dict[str, Any]]:
    return file_inventory([item for item in path.iterdir() if item.is_file()])


def auxiliary_profile_inventory() -> list[dict[str, Any]]:
    result = []
    for label in sorted(AUXILIARY_PROFILE_LABELS):
        path = P / "commands" / label
        exact_inventory(path, COMMAND_FILES, f"auxiliary profile command {label}")
        result.append({"label": label, "files": inventory_from_directory(path)})
    return result


def require_inventory(value: Sequence[Mapping[str, Any]], label: str) -> None:
    paths = [str(item.get("path", "")) for item in value]
    require(paths == sorted(paths), f"{label}: descriptor order changed")
    require(len(paths) == len(set(paths)), f"{label}: duplicate descriptor path")
    for index, item in enumerate(value):
        bind(item, f"{label}[{index}]")


def profile_folder(row: Mapping[str, Any]) -> Path:
    return base.folder(row)


def profile_command_path(row: Mapping[str, Any]) -> Path:
    return P / "commands" / ("profile-" + profile_folder(row).name) / "receipt.json"


def _finite(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{label}: timestamp is not numeric")
    result = float(value)
    require(result >= 0 and result == result and result != float("inf"),
            f"{label}: timestamp is not finite")
    return result


def validate_failure_witness() -> dict[str, Any]:
    """Bind the original v1 capture failure as an explicit witness."""

    require(not CAPTURE_PATH.exists(), "capture-profile.json must remain absent")
    exact_inventory(FAILED_CAPTURE_DIR, COMMAND_FILES, "original profile-capture command")
    started_path = FAILED_CAPTURE_DIR / "started.json"
    receipt_path = FAILED_CAPTURE_DIR / "receipt.json"
    log_path = FAILED_CAPTURE_DIR / "output.log"
    started = read(started_path)
    receipt = read(receipt_path)
    expected_argv = ["python3", "-B", str(P / "profile.py"), "capture"]
    require(started.get("argv") == expected_argv and receipt.get("argv") == expected_argv,
            "original profile-capture argv changed")
    require(started.get("cwd") == str(ROOT) and receipt.get("cwd") == str(ROOT),
            "original profile-capture cwd changed")
    require(receipt.get("exit_code") == 1, "original profile-capture was not a failed run")
    require(receipt.get("started_unix") == started.get("started_unix"),
            "original profile-capture start time changed")
    require(_finite(receipt.get("finished_unix"), "original profile-capture"),
            "original profile-capture finish time is malformed")
    require(receipt["finished_unix"] >= receipt["started_unix"],
            "original profile-capture finish precedes start")
    log = receipt.get("log")
    require(isinstance(log, Mapping), "original profile-capture log descriptor is missing")
    bind(log, "original profile-capture log")
    require(dict(log) == descriptor(log_path), "original profile-capture log descriptor changed")
    text = log_path.read_text(encoding="utf-8")
    require("function source file changed" in text
            and "reader.ReaderError" in text
            and "profile-01-shared-4096-before" in text,
            "original profile-capture failure reason changed")
    return {
        "status": "controlled-reader-failure",
        "exit_code": 1,
        "argv": expected_argv,
        "started": descriptor(started_path),
        "receipt": descriptor(receipt_path),
        "log": descriptor(log_path),
    }


def qualification_cache(value: Mapping[str, Any], bindings: Mapping[str, Any],
                        admission: Mapping[str, Mapping[str, Any]]) -> tuple[
                            dict[tuple[str, str, int], dict[str, Any]], list[dict[str, Any]]]:
    indices = {leg: base.qualification_index(leg, admission) for leg in base.LEGS}
    cache: dict[tuple[str, str, int], dict[str, Any]] = {}
    identities: list[dict[str, Any]] = []
    for row in value["rows"]:
        key = (row["leg"], row["shape"], row["parts"])
        if key in cache:
            continue
        item = base.expected_qualification(row, indices[row["leg"]], bindings, admission)
        cache[key] = item
        identities.append({
            "leg": row["leg"],
            "shape": row["shape"],
            "parts": row["parts"],
            "folder": str(item["folder"]),
            "report": descriptor(item["report_path"]),
        })
    require(len(cache) == 8 and len(identities) == 8,
            "qualification identity set is incomplete")
    return cache, sorted(identities, key=lambda item: (item["leg"], item["shape"], item["parts"]))


def validate_command(row: Mapping[str, Any], out: Path, started: Mapping[str, Any],
                     bindings: Mapping[str, Any], qualification: Mapping[str, Any],
                     receipt: Mapping[str, Any] | None) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    """Validate the already executed command and every descriptor it owns."""

    command_path = profile_command_path(row)
    exact_inventory(command_path.parent, COMMAND_FILES, f"profile command {out.name}")
    actual_command_ref = descriptor(command_path)
    if receipt is not None:
        command_ref = receipt.get("command")
        require(isinstance(command_ref, Mapping), f"profile receipt {out}: command descriptor missing")
        bind(command_ref, f"profile receipt command {out}")
        require(dict(command_ref) == actual_command_ref,
                f"profile receipt command descriptor changed: {out}")
    command = read(command_path)
    command_started = read(command_path.parent / "started.json")
    log = command.get("log")
    require(isinstance(log, Mapping), f"profile command log descriptor missing: {out}")
    bind(log, f"profile command log {out}")
    require(dict(log) == descriptor(command_path.parent / "output.log"),
            f"profile command log descriptor changed: {out}")
    argv = base.profile_argv(row, bindings[row["leg"]]["binary"], out)
    require(command.get("exit_code") == 0 and command.get("argv") == argv,
            f"profile command receipt changed: {out}")
    require(command_started.get("argv") == argv,
            f"profile command started argv changed: {out}")
    require(command.get("cwd") == str(ROOT) and command_started.get("cwd") == str(ROOT),
            f"profile command cwd changed: {out}")
    require(command.get("started_unix") == command_started.get("started_unix"),
            f"profile command start time changed: {out}")
    require(_finite(command.get("started_unix"), f"profile command {out}"),
            f"profile command start time is malformed: {out}")
    require(_finite(command.get("finished_unix"), f"profile command {out}"),
            f"profile command finish time is malformed: {out}")
    require(command["finished_unix"] >= command["started_unix"],
            f"profile command finished before start: {out}")
    require(command["started_unix"] >= started.get("started_unix", 0),
            f"profile command started before profile row: {out}")
    if receipt is not None:
        require(_finite(receipt.get("finished_unix"), f"profile receipt {out}"),
                f"profile receipt finish time is malformed: {out}")
        require(receipt["finished_unix"] >= command["finished_unix"],
                f"profile receipt finished before command: {out}")
    return command_path, command, command_started


def validate_profile_row(row: Mapping[str, Any], bindings: Mapping[str, Any],
                         qualification: Mapping[str, Any], admission: Mapping[str, Mapping[str, Any]],
                         require_receipt: bool) -> dict[str, Any]:
    """Validate a profile row with reader_v2 before any compaction."""

    out = profile_folder(row)
    require(out.is_dir() and not out.is_symlink(), f"missing profile folder: {out}")
    expected_files = PROFILE_FILES_WITH_RECEIPT if require_receipt else PROFILE_FILES
    exact_inventory(out, expected_files, f"profile folder {out.name}")
    started_path = out / "started.json"
    started = read(started_path)
    binary = bindings[row["leg"]]["binary"]
    build = bindings[row["leg"]]["build"]
    argv = base.profile_argv(row, binary, out)
    require(started.get("row") == dict(row), f"profile row identity changed: {out}")
    require(started.get("argv") == argv, f"profile started argv changed: {out}")
    require(started.get("binary") == binary and started.get("build") == build,
            f"profile binary/build binding changed: {out}")
    require(started.get("plan") == descriptor(PLAN_PATH), f"profile plan binding changed: {out}")
    require(started.get("source") == bindings[row["leg"]]["source"],
            f"profile source binding changed: {out}")
    require(started.get("qualification") == descriptor(qualification["report_path"]),
            f"profile qualification binding changed: {out}")
    require(_finite(started.get("started_unix"), f"profile row {out}"),
            f"profile row start time is malformed: {out}")

    receipt_path = out / "receipt.json"
    receipt = read(receipt_path) if require_receipt else None
    if receipt is not None:
        require(receipt.get("row") == dict(row) and receipt.get("exit_code") == 0,
                f"profile receipt identity/exit changed: {out}")
        artifacts = receipt.get("artifacts")
        require(isinstance(artifacts, list), f"profile artifact receipt is missing: {out}")
        artifact_paths = [str(item.get("path", "")) for item in artifacts]
        require(artifact_paths == sorted(artifact_paths), f"profile artifact order changed: {out}")
        require({Path(path).name for path in artifact_paths} == PROFILE_FILES - {"receipt.json"},
                f"profile artifact names changed: {out}")
        require(len(artifact_paths) == len(set(artifact_paths)),
                f"profile artifact descriptors are duplicated: {out}")
        for item in artifacts:
            bind(item, f"profile artifact {out}")
            require(dict(item) == descriptor(Path(item["path"])),
                    f"profile artifact descriptor changed: {out}")

    command_path, command, command_started = validate_command(
        row, out, started, bindings, qualification, receipt)
    counts = (out / "time.txt").read_text(encoding="utf-8").split()
    require(len(counts) == 3 and all(token.isdigit() for token in counts)
            and int(counts[0]) > 0, f"profile time receipt malformed: {out}")
    report = r.validate_report(
        out / "report.json", row["parts"], row["mode"], 1, 0,
        expected=qualification["report"], lane="profiles", shape=row["shape"],
    )
    profile = r.validate_profile(
        out / "profile.callgrind.1", out / "profile.callgrind",
        owner=base.PROFILE_OWNER, command=argv,
    )
    return {
        "row": dict(row),
        "folder": str(out),
        "started": descriptor(started_path),
        "report": report,
        "profile": profile,
        "command": descriptor(command_path),
        "command_started": descriptor(command_path.parent / "started.json"),
        "receipt": descriptor(receipt_path) if receipt is not None else None,
        "receipt_value": receipt,
        "raw_inventory": file_inventory([
            out / "profile.callgrind",
            out / "profile.callgrind.1",
        ]),
        "command_value": command,
        "command_started_value": command_started,
    }


def prefix_inventory(row: Mapping[str, Any], validated: Mapping[str, Any],
                     require_receipt: bool) -> dict[str, Any]:
    out = profile_folder(row)
    command_dir = profile_command_path(row).parent
    return {
        "row": dict(row),
        "folder": str(out),
        "files": inventory_from_directory(out),
        "command_files": inventory_from_directory(command_dir),
        "receipt_expected": require_receipt,
        "raw_callgrind": list(validated["raw_inventory"]),
        "profile": {
            "sha256": validated["profile"]["sha256"],
            "summary_ir": validated["profile"]["summary_ir"],
            "owner": dict(validated["profile"]["owner"]),
            "validation": dict(validated["profile"]["validation"]),
        },
    }


def static_freeze(value: Mapping[str, Any], admission_value: Mapping[str, Any],
                  bindings: Mapping[str, Any], qualification: Sequence[Mapping[str, Any]],
                  failure: Mapping[str, Any], prefix: Sequence[Mapping[str, Any]],
                  repair: Mapping[str, Any], auxiliary: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    return {
        "schema": "litchi.opc-preservation-profile-resume-freeze-0845.v2",
        "status": "pre-resume-frozen",
        "scope": "additive Callgrind resume; no native or Heaptrack pooling",
        "plan": descriptor(PLAN_PATH),
        "original_runner": descriptor(P / "profile.py"),
        "resume_runner": descriptor(RUNNER_PATH),
        "reader_v1": descriptor(P / "reader.py"),
        "reader_v2": descriptor(READER_PATH),
        "reader_repair": dict(repair),
        "admission": descriptor(P / "admission.json"),
        "admission_status": admission_value["status"],
        "bindings": dict(bindings_summary(bindings)),
        "qualification": list(qualification),
        "original_capture_failure": dict(failure),
        "auxiliary_profile_commands": list(auxiliary),
        "prefix_rows": list(prefix),
        "prefix_count": PREFIX_COUNT,
        "remaining_rows": [dict(row) for row in value["rows"][PREFIX_COUNT:]],
        "remaining_count": len(value["rows"]) - PREFIX_COUNT,
        "rules": {
            "preserve_raw_prefix": True,
            "seal_missing_receipt_after_reader_v2": True,
            "run_remaining_rows_once": True,
            "capture_manifest_written": False,
            "native_or_heap_overlap": False,
            "native_cycle_or_phase_time_claim": False,
        },
    }


def prepare(*, require_current_after: bool = True) -> tuple[
        dict[str, Any], Mapping[str, Any], dict[str, Any],
        dict[tuple[str, str, int], dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    value = plan()
    require(not CAPTURE_PATH.exists(), "capture-profile.json must remain absent")
    _, admission = base.admission_index()
    admission_value = read(P / "admission.json")
    bindings = base.validate_bindings(value, admission, require_current_after=require_current_after)
    qcache, qualifications = qualification_cache(value, bindings, admission)
    failure = validate_failure_witness()
    repair = repair_witness()
    return value, admission, bindings, qcache, qualifications, {
        "admission": admission_value,
        "failure": failure,
        "repair": repair,
    }


def freeze() -> dict[str, Any]:
    """Freeze the four existing rows before any new receipt or workload."""

    require(not RESULT_PATH.exists(), "resume result already exists; refusing overwrite")
    require(not ANALYSIS_PATH.exists(), "v2 analysis already exists; refusing overwrite")
    value, admission, bindings, qcache, qualifications, prepared = prepare()
    require(PROFILE_DIR.is_dir(), "profile-runs is missing")
    expected_prefix = {profile_folder(row).name for row in value["rows"][:PREFIX_COUNT]}
    require({item.name for item in PROFILE_DIR.iterdir()} == expected_prefix,
            "profile-runs contains rows outside the four frozen prefix rows")
    profile_labels = {
        path.parent.name for path in (P / "commands").glob("profile-*/receipt.json")
    }
    require(profile_labels == {"profile-capture"}
            | AUXILIARY_PROFILE_LABELS
            | {"profile-" + profile_folder(row).name for row in value["rows"][:PREFIX_COUNT]},
            "profile command inventory is not the expected partial prefix")
    auxiliary = auxiliary_profile_inventory()
    prefix: list[dict[str, Any]] = []
    for index, row in enumerate(value["rows"][:PREFIX_COUNT]):
        key = (row["leg"], row["shape"], row["parts"])
        validated = validate_profile_row(
            row, bindings, qcache[key], admission, require_receipt=index < PREFIX_COUNT - 1,
        )
        prefix.append(prefix_inventory(row, validated, index < PREFIX_COUNT - 1))
    result = static_freeze(value, prepared["admission"], bindings, qualifications,
                           prepared["failure"], prefix, prepared["repair"], auxiliary)
    if FREEZE_PATH.exists():
        require(read(FREEZE_PATH) == result,
                "profile-resume-freeze.json already exists with different content")
    else:
        d.write(FREEZE_PATH, result)
    print("Profile resume pre-resume freeze PASS", PREFIX_COUNT, "rows", flush=True)
    return result


def validate_frozen_artifacts(frozen: Mapping[str, Any]) -> None:
    require(frozen.get("schema") == "litchi.opc-preservation-profile-resume-freeze-0845.v2",
            "resume freeze schema changed")
    require(frozen.get("status") == "pre-resume-frozen", "resume freeze is not passing")
    require(frozen.get("prefix_count") == PREFIX_COUNT and frozen.get("remaining_count") == 12,
            "resume freeze row counts changed")
    require(frozen.get("plan") == descriptor(PLAN_PATH), "resume freeze plan changed")
    require(frozen.get("original_runner") == descriptor(P / "profile.py"),
            "resume freeze original runner changed")
    require(frozen.get("resume_runner") == descriptor(RUNNER_PATH),
            "resume freeze resume runner changed")
    require(frozen.get("reader_v1") == descriptor(P / "reader.py"),
            "resume freeze v1 reader changed")
    require(frozen.get("reader_v2") == descriptor(READER_PATH),
            "resume freeze v2 reader changed")
    repair = frozen.get("reader_repair")
    require(isinstance(repair, Mapping), "resume freeze reader repair witness is missing")
    require(repair.get("descriptor") == descriptor(REPAIR_SPEC_PATH),
            "resume freeze reader repair specification changed")
    bind(repair["installed_reference"], "frozen Callgrind format reference")
    bind(repair["regression_tests"], "frozen reader v2 regression result")
    require(frozen.get("admission") == descriptor(P / "admission.json"),
            "resume freeze admission changed")
    for category in ("prefix_rows",):
        entries = frozen.get(category)
        require(isinstance(entries, list) and len(entries) == PREFIX_COUNT,
                f"resume freeze {category} changed")
        for entry in entries:
            require_inventory(entry.get("files", []), f"resume freeze {category} files")
            require_inventory(entry.get("command_files", []),
                              f"resume freeze {category} command files")
            require_inventory(entry.get("raw_callgrind", []),
                              f"resume freeze {category} raw files")
            for item in entry["files"] + entry["command_files"]:
                bind(item, "resume frozen artifact")
    failure = frozen.get("original_capture_failure")
    require(isinstance(failure, Mapping), "resume freeze failure witness is missing")
    bind(failure["started"], "resume frozen failure start")
    bind(failure["receipt"], "resume frozen failure receipt")
    bind(failure["log"], "resume frozen failure log")
    auxiliary = frozen.get("auxiliary_profile_commands")
    require(isinstance(auxiliary, list), "resume freeze auxiliary command inventory is missing")
    require({item.get("label") for item in auxiliary} == AUXILIARY_PROFILE_LABELS,
            "resume freeze auxiliary command inventory changed")
    for item in auxiliary:
        require_inventory(item.get("files", []),
                          f"resume frozen auxiliary command {item.get('label')}")
    require(not CAPTURE_PATH.exists(), "capture-profile.json must remain absent")


def verify_prefix_against_freeze(frozen: Mapping[str, Any], value: Mapping[str, Any],
                                 bindings: Mapping[str, Any],
                                 qcache: Mapping[tuple[str, str, int], Mapping[str, Any]],
                                 admission: Mapping[str, Mapping[str, Any]],
                                 *, receipt_for_missing: bool) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    prefix = frozen["prefix_rows"]
    for index, row in enumerate(value["rows"][:PREFIX_COUNT]):
        key = (row["leg"], row["shape"], row["parts"])
        require(prefix[index]["row"] == dict(row), "resume frozen prefix row order changed")
        has_receipt = index < PREFIX_COUNT - 1 or receipt_for_missing
        validated = validate_profile_row(row, bindings, qcache[key], admission, has_receipt)
        if not has_receipt:
            current_files = inventory_from_directory(profile_folder(row))
            require(current_files == prefix[index]["files"],
                    f"frozen raw prefix changed before receipt sealing: {row}")
        else:
            frozen_files = prefix[index]["files"]
            current_raw = {
                item["path"]: item for item in inventory_from_directory(profile_folder(row))
                if item["path"].endswith("profile.callgrind")
                or item["path"].endswith("profile.callgrind.1")
            }
            frozen_raw = {
                item["path"]: item for item in prefix[index]["raw_callgrind"]
            }
            require(current_raw == frozen_raw,
                    f"frozen raw prefix changed after receipt sealing: {row}")
            if index < PREFIX_COUNT - 1:
                require(frozen_files == inventory_from_directory(profile_folder(row)),
                        f"frozen prefix artifact changed: {row}")
        rows.append(validated)
    return rows


def write_receipt(row: Mapping[str, Any], out: Path, command_path: Path, code: int) -> dict[str, Any]:
    require(code == 0, f"profile command failed: {out.name}")
    receipt_path = out / "receipt.json"
    require(not receipt_path.exists(), f"profile receipt already exists: {out}")
    artifacts = [descriptor(path) for path in sorted(out.rglob("*"), key=lambda item: str(item))
                 if path.is_file() and path.name != "receipt.json"]
    result = {
        "row": dict(row),
        "exit_code": code,
        "command": descriptor(command_path),
        "artifacts": artifacts,
        "finished_unix": max(time.time(), read(command_path).get("finished_unix", 0)),
    }
    d.write(receipt_path, result)
    return result


def resume() -> dict[str, Any]:
    """Seal row three, then run rows four through fifteen exactly once."""

    require(FREEZE_PATH.is_file(), "run freeze before resume")
    require(not RESULT_PATH.exists(), "resume result already exists; refusing overwrite")
    require(not ANALYSIS_PATH.exists(), "v2 analysis already exists; refusing overwrite")
    frozen = read(FREEZE_PATH)
    validate_frozen_artifacts(frozen)
    value, admission, bindings, qcache, qualifications, prepared = prepare()
    require(frozen == freeze(), "pre-resume freeze changed or is not reproducible")
    prefix = verify_prefix_against_freeze(
        frozen, value, bindings, qcache, admission, receipt_for_missing=False,
    )

    # This is the only write to the incomplete fourth row.  Every raw artifact
    # and the report was validated above with reader_v2 before this call.
    missing_row = value["rows"][PREFIX_COUNT - 1]
    missing_out = profile_folder(missing_row)
    missing_command = profile_command_path(missing_row)
    write_receipt(missing_row, missing_out, missing_command, 0)
    prefix = verify_prefix_against_freeze(
        frozen, value, bindings, qcache, admission, receipt_for_missing=True,
    )

    for row in value["rows"][PREFIX_COUNT:]:
        out = profile_folder(row)
        command_path = profile_command_path(row)
        require(not out.exists(), f"refusing overwrite of profile row: {out}")
        require(not command_path.parent.exists(),
                f"refusing overwrite of profile command: {command_path.parent}")
        out.mkdir(parents=False, exist_ok=False)
        qualification = qcache[(row["leg"], row["shape"], row["parts"])]
        binary = bindings[row["leg"]]["binary"]
        argv = base.profile_argv(row, binary, out)
        d.write(out / "started.json", {
            "row": dict(row),
            "argv": argv,
            "binary": binary,
            "build": bindings[row["leg"]]["build"],
            "plan": descriptor(PLAN_PATH),
            "qualification": descriptor(qualification["report_path"]),
            "source": bindings[row["leg"]]["source"],
            "started_unix": time.time(),
        })
        code = d.run("profile-" + out.name, argv)
        require(code == 0, f"profile command failed: {out.name}")
        # The full v2 report and profile checks happen before this row receipt.
        validate_profile_row(row, bindings, qualification, admission, require_receipt=False)
        write_receipt(row, out, command_path, code)
        validate_profile_row(row, bindings, qualification, admission, require_receipt=True)
        print("profile-" + out.name, "captured", flush=True)

    final_rows = validate_complete(value, admission, bindings, qcache)
    result = resume_result(value, prepared["admission"], bindings, qualifications,
                           frozen, prepared["failure"], prepared["repair"], final_rows)
    d.write(RESULT_PATH, result)
    print("Profile resume PASS", result["reports"], "reports;", result["samples"], "samples", flush=True)
    return result


def validate_complete(value: Mapping[str, Any], admission: Mapping[str, Mapping[str, Any]],
                      bindings: Mapping[str, Any],
                      qcache: Mapping[tuple[str, str, int], Mapping[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for row in value["rows"]:
        key = (row["leg"], row["shape"], row["parts"])
        rows.append(validate_profile_row(row, bindings, qcache[key], admission, True))
    require({item["folder"] for item in rows} == {str(profile_folder(row)) for row in value["rows"]},
            "profile folder inventory differs from frozen plan")
    require({item.name for item in PROFILE_DIR.iterdir() if item.is_dir()}
            == {profile_folder(row).name for row in value["rows"]},
            "profile run directory inventory differs from frozen plan")
    require({item.name for item in PROFILE_DIR.iterdir() if item.is_file()} == set(),
            "profile run root contains an unexpected file")
    expected_labels = {"profile-" + profile_folder(row).name for row in value["rows"]}
    actual_labels = {
        path.parent.name for path in (P / "commands").glob("profile-*/receipt.json")
        if path.parent.name not in {"profile-capture"} | AUXILIARY_PROFILE_LABELS
    }
    require(actual_labels == expected_labels,
            "profile command receipt inventory differs from frozen plan")
    ordered = sorted(rows, key=lambda item: read(Path(item["command"]["path"]))["started_unix"])
    require([item["row"] for item in ordered] == value["rows"],
            "profile receipt order differs from frozen plan")
    base.check_no_overlap(rows)
    return rows


def resume_result(value: Mapping[str, Any], admission_value: Mapping[str, Any],
                  bindings: Mapping[str, Any], qualifications: Sequence[Mapping[str, Any]],
                  frozen: Mapping[str, Any], failure: Mapping[str, Any],
                  repair: Mapping[str, Any],
                  rows: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    receipts = [descriptor(Path(item["folder"]) / "receipt.json") for item in rows]
    return {
        "schema": "litchi.opc-preservation-profile-resume-0845.v2",
        "status": "pass",
        "scope": "diagnostic-only-callgrind; separate from native and Heaptrack populations",
        "reports": len(rows),
        "samples": len(rows),
        "plan": descriptor(PLAN_PATH),
        "original_runner": descriptor(P / "profile.py"),
        "runner": descriptor(RUNNER_PATH),
        "reader": descriptor(READER_PATH),
        "reader_v1": descriptor(P / "reader.py"),
        "reader_repair": dict(repair),
        "admission": descriptor(P / "admission.json"),
        "admission_status": admission_value["status"],
        "freeze": descriptor(FREEZE_PATH),
        "original_capture_failure": dict(failure),
        "auxiliary_profile_commands": auxiliary_profile_inventory(),
        "bindings": dict(bindings_summary(bindings)),
        "qualification": list(qualifications),
        "prefix_count": PREFIX_COUNT,
        "resumed_count": len(rows) - PREFIX_COUNT,
        "receipts": receipts,
        "rules": {
            "preserved_prefix_raw_captures": True,
            "missing_receipt_sealed_after_reader_v2": True,
            "remaining_rows_run_once": True,
            "capture_manifest_written": False,
            "native_or_heap_overlap": False,
            "native_cycle_or_phase_time_claim": False,
        },
    }


def compact_functions_v2(parsed: Mapping[str, Any], prefixes: Sequence[str]) -> list[dict[str, Any]]:
    return base.compact_functions(parsed, prefixes)


def compact_profile_v2(raw_path: Path, terminal_path: Path, profile: Mapping[str, Any],
                       value: Mapping[str, Any], selected_symbols: Mapping[str, Sequence[str]]) -> dict[str, Any]:
    """Compact only after reader_v2 has validated the complete raw pair."""

    parsed = r.parse_raw(raw_path)
    require(parsed["sha256"] == profile["sha256"], "raw Callgrind hash changed during reduction")
    selected: dict[str, Any] = {}
    for category, prefixes in selected_symbols.items():
        functions = compact_functions_v2(parsed, prefixes)
        selected[category] = {
            "prefixes": list(prefixes),
            "functions": functions,
            "function_count": len(functions),
            # Self costs partition the validated total within this parsed
            # function table.  Inclusive edge costs remain as concrete rows;
            # they are deliberately not summed because regions can overlap.
            "self_ir": sum(item["self_ir"] for item in functions),
            "edge_count": sum(len(item["edges"]) for item in functions),
            "inclusive_ir_rows": [
                {
                    "caller_id": edge.get("caller_id"),
                    "callee_id": edge.get("callee_id"),
                    "inclusive_ir": edge.get("inclusive_ir"),
                }
                for function in functions
                for edge in function.get("edges", [])
            ],
        }
    owner = profile["owner"]
    return {
        "raw_callgrind": {
            "numbered": descriptor(raw_path),
            "terminal": descriptor(terminal_path),
        },
        "summary_ir": profile["summary_ir"],
        "owner": {
            "id": owner["id"],
            "name": owner["name"],
            "file": owner["file"],
            "self_ir": owner["self_ir"],
            "direct_children_ir": owner["direct_children_ir"],
            "inclusive_ir": profile["summary_ir"],
            "incoming": [dict(owner["incoming"])],
        },
        "selected": selected,
        "validation": dict(profile["validation"]),
        "interpretation": {
            "file_and_line": "parser-emitted context only; no source or line timing claim",
            "inclusive_regions": "retained as concrete edge rows; overlapping regions are not summed",
            "missing_symbols": "an absent or inlined selected symbol is valid evidence",
        },
        "report": {
            "path": value["report"]["report"]["path"],
            "sha256": value["report"]["report"]["sha256"],
            "source": value["report"]["source"],
            "output": value["report"]["output"],
            "logical": value["report"]["logical"],
        },
    }


def selected_totals_v2(compact: Mapping[str, Any], category: str) -> dict[str, Any]:
    selected = compact["selected"][category]
    return {
        "function_count": selected["function_count"],
        "self_ir": selected["self_ir"],
        "edge_count": selected["edge_count"],
        "inclusive_ir_rows": selected["inclusive_ir_rows"],
    }


def compare_profiles_v2(rows: Sequence[Mapping[str, Any]], value: Mapping[str, Any]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, int, int], dict[str, Mapping[str, Any]]] = {}
    for item in rows:
        row = item["row"]
        groups.setdefault((row["shape"], row["parts"], row["repeat"]), {})[row["leg"]] = item
    result: list[dict[str, Any]] = []
    for shape in base.SHAPES:
        for parts in base.PARTS:
            for repeat in range(2):
                group = groups[(shape, parts, repeat)]
                require(set(group) == set(base.LEGS),
                        f"missing profile leg for {shape}/{parts}/{repeat}")
                before = group["before"]["compact"]
                after = group["after"]["compact"]
                selected: dict[str, Any] = {}
                for category in value["selected_symbols"]:
                    b = selected_totals_v2(before, category)
                    a = selected_totals_v2(after, category)
                    selected[category] = {
                        "before": b,
                        "after": a,
                        "delta_self_ir": a["self_ir"] - b["self_ir"],
                    }
                result.append({
                    "shape": shape,
                    "parts": parts,
                    "repeat": repeat,
                    "owner": {
                        "before_summary_ir": before["summary_ir"],
                        "after_summary_ir": after["summary_ir"],
                        "before_self_ir": before["owner"]["self_ir"],
                        "after_self_ir": after["owner"]["self_ir"],
                        "delta_summary_ir": after["summary_ir"] - before["summary_ir"],
                    },
                    "selected": selected,
                })
    return result


def analyze() -> dict[str, Any]:
    """Replay all sixteen rows and write the deterministic v2 analysis."""

    require(FREEZE_PATH.is_file(), "run freeze before analysis")
    require(RESULT_PATH.is_file(), "run resume before analysis")
    frozen = read(FREEZE_PATH)
    validate_frozen_artifacts(frozen)
    result_manifest = read(RESULT_PATH)
    require(result_manifest.get("status") == "pass", "resume result is not passing")
    value, admission, bindings, qcache, _, prepared = prepare(require_current_after=False)
    require(result_manifest.get("plan") == descriptor(PLAN_PATH), "resume result plan changed")
    require(result_manifest.get("runner") == descriptor(RUNNER_PATH),
            "resume result runner changed")
    require(result_manifest.get("reader") == descriptor(READER_PATH),
            "resume result reader changed")
    require(result_manifest.get("reader_repair") == repair_witness(),
            "resume result reader repair witness changed")
    require(result_manifest.get("freeze") == descriptor(FREEZE_PATH),
            "resume result freeze changed")
    require(result_manifest.get("original_capture_failure") == prepared["failure"],
            "resume result failure witness changed")
    rows = validate_complete(value, admission, bindings, qcache)
    receipts = [descriptor(Path(item["folder"]) / "receipt.json") for item in rows]
    require(result_manifest.get("receipts") == receipts,
            "resume result receipt inventory/order changed")
    compacted: list[dict[str, Any]] = []
    for validated in rows:
        row = validated["row"]
        compacted.append({
            "row": dict(row),
            "folder": validated["folder"],
            "receipt": validated["receipt"],
            "command": validated["command"],
            "compact": compact_profile_v2(
                P / validated["folder"] / "profile.callgrind.1",
                P / validated["folder"] / "profile.callgrind",
                validated["profile"],
                validated,
                value["selected_symbols"],
            ),
        })
    result = {
        "schema": "litchi.opc-preservation-profile-analysis-0845.v2",
        "status": "pass",
        "disposition": "diagnostic-only-no-native-pooling",
        "scope": "Callgrind Ir counts inside profile_publish; no cycle, elapsed-phase, or native timing claim",
        "reports": len(compacted),
        "samples": len(compacted),
        "plan": descriptor(PLAN_PATH),
        "original_runner": descriptor(P / "profile.py"),
        "runner": descriptor(RUNNER_PATH),
        "reader": descriptor(READER_PATH),
        "reader_v1": descriptor(P / "reader.py"),
        "reader_repair": repair_witness(),
        "admission": descriptor(P / "admission.json"),
        "admission_status": prepared["admission"]["status"],
        "freeze": descriptor(FREEZE_PATH),
        "resume": descriptor(RESULT_PATH),
        "original_capture_failure": prepared["failure"],
        "rows": compacted,
        "comparisons": compare_profiles_v2(compacted, value),
        "selected_symbols": value["selected_symbols"],
        "limits": value["limits"],
    }
    if ANALYSIS_PATH.exists():
        require(read(ANALYSIS_PATH) == result,
                "profile-analysis-v2.json already exists with different content")
    else:
        d.write(ANALYSIS_PATH, result)
    print("Profile v2 analysis PASS", result["reports"], "reports;", result["samples"], "samples", flush=True)
    return result


if __name__ == "__main__":
    require(len(sys.argv) == 2, "usage: profile_resume.py freeze|resume|analyze")
    if sys.argv[1] == "freeze":
        freeze()
    elif sys.argv[1] == "resume":
        resume()
    elif sys.argv[1] == "analyze":
        analyze()
    else:
        fail(f"unknown mode: {sys.argv[1]}")
