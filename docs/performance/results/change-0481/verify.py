#!/usr/bin/env python3
"""Fail-closed verification for the portable 0481 borrowed-name bundle."""

from __future__ import annotations

import argparse
import ast
import datetime as _datetime
import hashlib
import json
import math
import re
import sys
from pathlib import Path, PurePosixPath
from typing import Any, Mapping, NoReturn

import analyze


ROOT = Path(__file__).resolve().parent
SHA256 = re.compile(r"^[0-9a-f]{64}$")
LABEL = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
VERIFICATION_SCHEMA = "docx-borrowed-names-comparison-verification-v1"
VALIDATION_SCHEMA = "docx-borrowed-names-validation-v1"
EMBEDDED_INPUT_SCHEMA = "docx-embedded-inputs-v1"
EMBEDDED_INPUTS_PATH = "embedded-inputs.json"
SOURCE_CHANGE_PATH = "crates/litchi-docx/src/source_backed/paragraph_copy.rs"

# These are the gates owned by this change.  Profile gates are optional
# whole-process observations and are checked when present, but they do not
# weaken the reproducibility boundary if perf is unavailable on another host.
REQUIRED_LABELS = (
    "format-docx", "docx-tests", "opc-shared-tests", "harness-tests", "clippy-all-features",
    "rustdoc", "boundaries", "registry-strict", "evidence-tests-final", "analyze-final",
    "build-control-normal", "build-control-allocator", "build-candidate-normal",
    "build-candidate-allocator", "pilot-control-normal-total", "pilot-control-normal-phases",
    "pilot-control-allocator-total", "pilot-control-allocator-phases",
    "pilot-candidate-normal-total", "pilot-candidate-normal-phases",
    "pilot-candidate-allocator-total", "pilot-candidate-allocator-phases",
)
REQUIRED_FINAL_LABELS = REQUIRED_LABELS
BUILD_LABELS = {
    "build-control-normal", "build-control-allocator",
    "build-candidate-normal", "build-candidate-allocator",
}
PILOT_LABELS = {
    f"pilot-{arm}-{instrumentation}-{mode}"
    for arm in analyze.ARMS
    for instrumentation in analyze.INSTRUMENTATIONS
    for mode in analyze.MODES
}
RUN_LABELS = set(REQUIRED_LABELS) - BUILD_LABELS - PILOT_LABELS


class VerificationError(ValueError):
    """A retained artifact failed an independent check."""


def fail(message: str) -> NoReturn:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path, label: str = "JSON") -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label}: regular file required")
    try:
        return json.loads(
            path.read_text(encoding="utf-8"), object_pairs_hook=_pairs,
            parse_constant=lambda value: (_ for _ in ()).throw(ValueError(value)),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")


def sha256_file(path: Path, label: str = "file") -> tuple[str, int]:
    require(path.is_file() and not path.is_symlink(), f"{label}: regular file required")
    digest = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except OSError as error:
        fail(f"{label}: cannot read: {error}")
    return digest.hexdigest(), size


def digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256.fullmatch(value) is not None, f"{label}: expected SHA-256 digest")
    return value


def integer(value: Any, label: str, expected: int | None = None) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0, f"{label}: expected a non-negative integer")
    if expected is not None:
        require(value == expected, f"{label}: expected {expected}, got {value}")
    return value


def timestamp(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{label}: timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label}: timezone is required")
    return parsed


def interval(start: Any, finish: Any, label: str) -> None:
    require(timestamp(finish, f"{label}.finished_utc") >= timestamp(start, f"{label}.started_utc"), f"{label}: finished before started")


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label}: expected relative POSIX path")
    require("\\" not in value, f"{label}: backslashes are forbidden")
    path = PurePosixPath(value)
    require(not path.is_absolute() and value not in {".", ".."}, f"{label}: absolute/parent path is forbidden")
    require(all(part not in {"", ".", ".."} for part in path.parts), f"{label}: path escapes bundle root")
    return path.as_posix()


def bundle_file(root: Path, value: Any, label: str) -> Path:
    relative = safe_relative(value, label)
    current = root
    for part in PurePosixPath(relative).parts:
        current /= part
        require(not current.is_symlink(), f"{label}: symlink component is forbidden")
    try:
        resolved = current.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve: {error}")
    base = root.resolve()
    require(resolved == base or base in resolved.parents, f"{label}: path escapes bundle root")
    require(current.is_file(), f"{label}: file is missing")
    return current


def metadata(path: Path, label: str = "file") -> dict[str, Any]:
    value, size = sha256_file(path, label)
    return {"bytes": size, "sha256": value}


def check_metadata(path: Path, expected: Any, label: str) -> None:
    require(isinstance(expected, dict) and set(expected) == {"bytes", "sha256"}, f"{label}: metadata fields differ")
    integer(expected.get("bytes"), f"{label}.bytes")
    digest(expected.get("sha256"), f"{label}.sha256")
    require(metadata(path, label) == expected, f"{label}: metadata differs")


def _manifest(root: Path, reference: Any, label: str) -> tuple[str, dict[str, str]]:
    require(isinstance(reference, dict) and set(reference) == {"path", "sha256", "files"}, f"{label}: source manifest reference differs")
    path = bundle_file(root, reference["path"], f"{label}.path")
    actual_hash, _ = sha256_file(path, label)
    require(actual_hash == digest(reference["sha256"], f"{label}.sha256"), f"{label}: manifest digest differs")
    value = read_json(path, label)
    require(isinstance(value, dict), f"{label}: manifest must be an object")
    file_count = integer(reference["files"], f"{label}.files")
    require(file_count == len(value) == 7048, f"{label}: expected flat 7048-file source manifest")
    result: dict[str, str] = {}
    for source, source_hash in value.items():
        require(isinstance(source, str), f"{label}: source path is malformed")
        safe_relative(source, f"{label}.{source}")
        result[source] = digest(source_hash, f"{label}.{source}")
    return actual_hash, result


def _check_protocol_scripts(root: Path, protocol: Mapping[str, Any]) -> None:
    for name, expected in protocol["scripts"].items():
        path = bundle_file(root, name, f"protocol.scripts.{name}")
        require(sha256_file(path, name)[0] == digest(expected, f"protocol.scripts.{name}"), f"protocol script digest differs: {name}")


def _check_argv(argv: Any, spec: Mapping[str, Any], label: str) -> None:
    require(isinstance(argv, list) and len(argv) == 18 and all(isinstance(item, str) for item in argv), f"{label}.argv: expected 18 strings")
    fixed = {0: "/usr/bin/time", 1: "-v", 2: "-o", 4: "/usr/bin/taskset", 5: "-c", 6: "2", 8: "--mode", 10: "--counts", 12: "--samples", 13: "30", 14: "--warmups", 15: "3", 16: "--json"}
    for index, expected in fixed.items():
        require(argv[index] == expected, f"{label}.argv[{index}]: expected {expected!r}")
    require(argv[9] == spec["mode"] and argv[11] == str(spec["count"]), f"{label}: mode/count are not bound")
    executable = Path(argv[7])
    require(executable.is_absolute() and executable.as_posix().endswith(f"/{spec['arm']}-{spec['instrumentation']}/docx_plain_paragraph_tail_append"), f"{label}.argv[7]: binary binding differs")
    require(Path(argv[3]).is_absolute() and argv[3].endswith(f"/captures/{spec['label']}.resource"), f"{label}.argv[3]: resource path does not bind label")
    require(Path(argv[17]).is_absolute() and argv[17].endswith(f"/captures/{spec['label']}.report.json"), f"{label}.argv[17]: report path does not bind label")


def check_protocol(root: Path) -> tuple[dict[str, Any], str]:
    path = bundle_file(root, "protocol.json", "protocol.json")
    protocol = read_json(path, "protocol.json")
    require(isinstance(protocol, dict), "protocol.json: expected object")
    try:
        rows = analyze.protocol_rows(protocol)
    except analyze.AnalysisError as error:
        fail(str(error))
    _check_protocol_scripts(root, protocol)
    for spec in rows:
        _check_argv(spec["argv"], spec, f"protocol.captures[{spec['label']}]" )
    return protocol, sha256_file(path, "protocol.json")[0]


def check_environment(root: Path, protocol: Mapping[str, Any]) -> None:
    path = bundle_file(root, "environment.json", "environment.json")
    value = read_json(path, "environment.json")
    require(isinstance(value, dict), "environment.json: expected object")
    filesystem = value.get("filesystem")
    require(isinstance(filesystem, dict), "environment.filesystem: missing")
    require(filesystem.get("exit_code") == 0 and filesystem.get("stderr") == "", "environment filesystem probe failed")
    stdout = filesystem.get("stdout")
    require(isinstance(stdout, str) and any(line.split()[:2] == ["tmpfs", "tmpfs"] for line in stdout.splitlines()[1:]), "environment filesystem is not tmpfs")
    for key, expected in protocol["environment"].items():
        require(isinstance(expected, str), f"protocol.environment.{key}: expected text")


def _check_embedded_inputs(root: Path, reference: Any, label: str) -> dict[str, Any]:
    require(isinstance(reference, dict) and set(reference) == {"path", "bytes", "sha256"}, f"{label}: reference differs")
    require(reference["path"] == EMBEDDED_INPUTS_PATH, f"{label}.path differs")
    path = bundle_file(root, reference["path"], f"{label}.path")
    check_metadata(path, {"bytes": reference["bytes"], "sha256": reference["sha256"]}, label)
    value = read_json(path, label)
    require(isinstance(value, dict) and value.get("schema") == EMBEDDED_INPUT_SCHEMA and isinstance(value.get("files"), dict), f"{label}: manifest schema differs")
    require(len(value["files"]) == 8, f"{label}: expected eight embedded inputs")
    files: dict[str, Any] = {}
    for source, item in value["files"].items():
        require(isinstance(source, str) and isinstance(item, dict) and set(item) == {"bytes", "path", "sha256"}, f"{label}.{source}: entry differs")
        destination = bundle_file(root, item["path"], f"{label}.{source}.path")
        check_metadata(destination, {"bytes": item["bytes"], "sha256": item["sha256"]}, f"{label}.{source}")
        files[source] = dict(item)
    return {"path": EMBEDDED_INPUTS_PATH, "bytes": reference["bytes"], "sha256": digest(reference["sha256"], f"{label}.sha256"), "files": files}


def _check_build_gate(root: Path, raw: Mapping[str, Any], arm: str, instrumentation: str, label: str) -> tuple[str, dict[str, str]]:
    gate_path = bundle_file(root, raw["build_gate"], f"{label}.build_gate")
    require(sha256_file(gate_path, label)[0] == digest(raw["build_gate_sha256"], f"{label}.build_gate_sha256"), f"{label}: build gate digest differs")
    gate = read_json(gate_path, label)
    require(isinstance(gate, dict) and gate.get("exit_code") == 0 and gate.get("source_unchanged") is True, f"{label}: build gate failed")
    expected_binary = {"path": raw["path"], "bytes": raw["bytes"], "sha256": raw["sha256"]}
    require(gate.get("binary") == expected_binary, f"{label}: copied binary metadata differs")
    original = gate.get("original_binary")
    require(isinstance(original, dict) and original.get("bytes") == raw["bytes"] and original.get("sha256") == raw["sha256"] and isinstance(original.get("path"), str) and Path(original["path"]).is_absolute(), f"{label}: original binary binding differs")
    require(gate.get("embedded_inputs") == raw["embedded_inputs"] and gate.get("embedded_inputs_unchanged") is True, f"{label}: embedded input binding differs")
    require(gate.get("driver_sha256") == sha256_file(bundle_file(root, "gate.py", "gate.py"), "gate.py")[0], f"{label}: gate driver binding differs")
    require(gate.get("common_sha256") == sha256_file(bundle_file(root, "common.py", "common.py"), "common.py")[0], f"{label}: common binding differs")
    validation_gate = gate.get("validation_gate")
    require(isinstance(validation_gate, dict) and set(validation_gate) == {"path", "sha256"}, f"{label}: validation gate binding missing")
    validation_path = bundle_file(root, validation_gate["path"], f"{label}.validation_gate.path")
    require(sha256_file(validation_path, f"{label}.validation_gate")[0] == digest(validation_gate["sha256"], f"{label}.validation_gate.sha256"), f"{label}: validation receipt digest differs")
    before_hash, before = _manifest(root, gate.get("source_before"), f"{label}.source_before")
    after_hash, after = _manifest(root, gate.get("source_after"), f"{label}.source_after")
    require(before_hash == after_hash == raw["source_manifest_sha256"], f"{label}: source manifest binding differs")
    return raw["source_manifest_sha256"], before if before == after else after


def check_binaries(root: Path, protocol: Mapping[str, Any]) -> dict[str, dict[str, dict[str, Any]]]:
    result: dict[str, dict[str, dict[str, Any]]] = {}
    for arm in analyze.ARMS:
        value = read_json(root / f"{arm}-binaries.json", f"{arm}-binaries.json")
        require(isinstance(value, dict) and set(value) == set(analyze.INSTRUMENTATIONS), f"{arm}-binaries: instrumentation identities differ")
        result[arm] = {}
        for instrumentation in analyze.INSTRUMENTATIONS:
            raw = value[instrumentation]
            required = {"path", "bytes", "sha256", "embedded_inputs", "embedded_inputs_unchanged", "build_gate", "build_gate_sha256", "source_manifest_sha256"}
            require(isinstance(raw, dict) and set(raw) == required, f"{arm}-binaries.{instrumentation}: fields differ")
            require(isinstance(raw["path"], str) and Path(raw["path"]).is_absolute() and raw["path"].endswith(f"/{arm}-{instrumentation}/docx_plain_paragraph_tail_append"), f"{arm}-binaries.{instrumentation}.path differs")
            integer(raw["bytes"], f"{arm}-binaries.{instrumentation}.bytes")
            require(raw["bytes"] > 0, f"{arm}-binaries.{instrumentation}.bytes must be positive")
            digest(raw["sha256"], f"{arm}-binaries.{instrumentation}.sha256")
            digest(raw["build_gate_sha256"], f"{arm}-binaries.{instrumentation}.build_gate_sha256")
            _check_embedded_inputs(root, raw["embedded_inputs"], f"{arm}-binaries.{instrumentation}.embedded_inputs")
            require(raw["embedded_inputs_unchanged"] is True, f"{arm}-binaries.{instrumentation}: embedded inputs changed")
            _check_build_gate(root, raw, arm, instrumentation, f"{arm}-binaries.{instrumentation}")
            result[arm][instrumentation] = dict(raw)
        require(result[arm]["normal"]["embedded_inputs"] == result[arm]["allocator"]["embedded_inputs"], f"{arm}: normal/allocator embedded inputs differ")
        require(result[arm]["normal"]["source_manifest_sha256"] == result[arm]["allocator"]["source_manifest_sha256"], f"{arm}: normal/allocator source manifests differ")
    require(result["control"]["normal"]["embedded_inputs"] == result["candidate"]["normal"]["embedded_inputs"], "control/candidate embedded inputs differ")
    return result


def _artifact_check(directory: Path, artifacts: Any, names: set[str], label: str) -> None:
    require(isinstance(artifacts, dict) and set(artifacts) == names, f"{label}: artifact inventory differs")
    for name, expected in artifacts.items():
        require(name == Path(name).name and "\\" not in name, f"{label}.{name}: unsafe artifact name")
        check_metadata(directory / name, expected, f"{label}/{name}")


def check_capture_receipts(root: Path, protocol: Mapping[str, Any], protocol_hash: str, binaries: Mapping[str, Mapping[str, Mapping[str, Any]]]) -> dict[str, dict[str, Any]]:
    directory = root / "captures"
    require(directory.is_dir() and not directory.is_symlink(), "captures directory is missing")
    reports: dict[str, dict[str, Any]] = {}
    previous_finished: _datetime.datetime | None = None
    for spec in analyze.protocol_rows(protocol):
        label = spec["label"]
        started = read_json(directory / f"{label}.started.json", f"captures/{label}.started.json")
        finished = read_json(directory / f"{label}.json", f"captures/{label}.json")
        expected_started = {"capture", "cwd", "started_utc", "protocol_sha256", "binary", "environment"}
        expected_finished = expected_started | {"exit_code", "finished_utc", "artifacts"}
        require(isinstance(started, dict) and set(started) == expected_started, f"{label}: started envelope differs")
        require(isinstance(finished, dict) and set(finished) == expected_finished, f"{label}: final envelope differs")
        require(started["capture"] == spec and started["protocol_sha256"] == protocol_hash, f"{label}: capture/protocol binding differs")
        require(started["environment"] == protocol["environment"], f"{label}: environment binding differs")
        require(isinstance(started["cwd"], str) and Path(started["cwd"]).is_absolute(), f"{label}: cwd provenance differs")
        require(started["binary"] == binaries[spec["arm"]][spec["instrumentation"]], f"{label}: binary binding differs")
        _check_argv(spec["argv"], spec, f"captures/{label}")
        require(spec["argv"][7] == started["binary"]["path"], f"{label}: executable does not bind binary")
        for key in ("capture", "cwd", "started_utc", "protocol_sha256", "binary", "environment"):
            require(finished[key] == started[key], f"{label}: started/final {key} differs")
        require(finished["exit_code"] == 0, f"{label}: capture failed")
        interval(started["started_utc"], finished["finished_utc"], f"capture {label}")
        started_at = timestamp(started["started_utc"], f"{label}.started_utc")
        finished_at = timestamp(finished["finished_utc"], f"{label}.finished_utc")
        if previous_finished is not None:
            require(started_at >= previous_finished, f"{label}: protocol chronology overlaps the preceding capture")
        previous_finished = finished_at
        names = {f"{label}.stdout", f"{label}.stderr", f"{label}.resource", f"{label}.report.json"}
        _artifact_check(directory, finished["artifacts"], names, f"capture {label}")
        resource = directory / f"{label}.resource"
        text = resource.read_text(encoding="utf-8")
        require(len(re.findall(r"^\s*Maximum resident set size \(kbytes\):\s*\d+\s*$", text, re.MULTILINE)) == 1, f"{label}: RSS resource observation differs")
        reports[label] = read_json(directory / f"{label}.report.json", f"captures/{label}.report.json")
    require(len(reports) == 48, "capture count differs")
    return reports


def _check_source_change(root: Path, binaries: Mapping[str, Mapping[str, Mapping[str, Any]]]) -> dict[str, Any]:
    initial = read_json(root / "initial-state.json", "initial-state.json")
    require(isinstance(initial, dict), "initial-state.json: expected object")
    initial_schema = digest(initial.get("report_schema_sha256"), "initial-state.report_schema_sha256")
    require(initial_schema == sha256_file(bundle_file(root, "report_schema.py", "report_schema.py"), "report_schema.py")[0], "report_schema.py changed")
    change = read_json(root / "source-change.json", "source-change.json")
    require(isinstance(change, dict) and set(change) == {"after_sha256", "before_path", "before_sha256", "patch_path", "path"}, "source-change.json fields differ")
    require(change["path"] == SOURCE_CHANGE_PATH and change["before_path"] == "before-source.txt" and change["patch_path"] == "source-change.patch", "source-change paths differ")
    before_file = bundle_file(root, change["before_path"], "before-source.txt")
    require(sha256_file(before_file, "before-source.txt")[0] == digest(change["before_sha256"], "source-change.before_sha256"), "before-source digest differs")
    before_text = before_file.read_text(encoding="utf-8")
    replacements = (
        (
            'let local = checked_clone(start.local_name().as_ref(), "XML local name")?;',
            'let local = start.local_name();',
        ),
        (
            'let local = checked_clone(empty.local_name().as_ref(), "XML local name")?;',
            'let local = empty.local_name();',
        ),
        (
            'let local = checked_clone(end.local_name().as_ref(), "XML local name")?;',
            'let local = end.local_name();',
        ),
        ('local.as_slice() == b"document"', 'local.as_ref() == b"document"'),
        ('next_scope(&stack, local.as_slice(), false)', 'next_scope(&stack, local.as_ref(), false)'),
        ('next_scope(&stack, local.as_slice(), true)', 'next_scope(&stack, local.as_ref(), true)'),
        ('scope_local(scope) != local.as_slice()', 'scope_local(scope) != local.as_ref()'),
    )
    after_text = before_text
    for old, new in replacements:
        require(after_text.count(old) == 1, f"before-source does not contain exactly one expected scanner expression: {old}")
        after_text = after_text.replace(old, new)
    require(hashlib.sha256(after_text.encode()).hexdigest() == digest(change["after_sha256"], "source-change.after_sha256"), "source-change after digest cannot be reconstructed")
    patch = bundle_file(root, change["patch_path"], "source-change.patch").read_text(encoding="utf-8").splitlines()
    require(patch[:1] == ["diff --git a/crates/litchi-docx/src/source_backed/paragraph_copy.rs b/crates/litchi-docx/src/source_backed/paragraph_copy.rs"], "production patch path differs")
    require(len([line for line in patch if line.startswith("-") and not line.startswith("---")]) == len(replacements), "production patch has unexpected deletion count")
    require(len([line for line in patch if line.startswith("+") and not line.startswith("+++")]) == len(replacements), "production patch has unexpected addition count")
    minus = [line[1:] for line in patch if line.startswith("-") and not line.startswith("---")]
    plus = [line[1:] for line in patch if line.startswith("+") and not line.startswith("+++")]
    expected_minus = [
        '                let local = checked_clone(start.local_name().as_ref(), "XML local name")?;',
        '                } else if local.as_slice() == b"document" && stack.is_empty() {',
        '                let scope = next_scope(&stack, local.as_slice(), false)?;',
        '                let local = checked_clone(empty.local_name().as_ref(), "XML local name")?;',
        '                let scope = next_scope(&stack, local.as_slice(), true)?;',
        '                let local = checked_clone(end.local_name().as_ref(), "XML local name")?;',
        '                if scope_local(scope) != local.as_slice() {',
    ]
    expected_plus = [
        '                let local = start.local_name();',
        '                } else if local.as_ref() == b"document" && stack.is_empty() {',
        '                let scope = next_scope(&stack, local.as_ref(), false)?;',
        '                let local = empty.local_name();',
        '                let scope = next_scope(&stack, local.as_ref(), true)?;',
        '                let local = end.local_name();',
        '                if scope_local(scope) != local.as_ref() {',
    ]
    require(minus == expected_minus, "production patch deletes unexpected lines")
    require(plus == expected_plus, "production patch adds unexpected lines")
    control_hash = binaries["control"]["normal"]["source_manifest_sha256"]
    candidate_hash = binaries["candidate"]["normal"]["source_manifest_sha256"]
    require(control_hash == initial.get("prior_source_sha256"), "control source is not the initial source")
    allowed = {SOURCE_CHANGE_PATH}
    control_manifest_path = next(root.glob("validation-sources/*.json"), None)
    require(control_manifest_path is not None, "source manifest files are missing")
    manifests: dict[str, dict[str, str]] = {}
    for arm in analyze.ARMS:
        manifest_hash = binaries[arm]["normal"]["source_manifest_sha256"]
        path = bundle_file(root, f"validation-sources/{manifest_hash}.json", f"{arm} source manifest")
        value = read_json(path, f"{arm} source manifest")
        require(isinstance(value, dict) and len(value) == 7048, f"{arm} source manifest must contain 7048 files")
        manifests[arm] = {str(key): digest(item, f"{arm} source manifest.{key}") for key, item in value.items()}
    require(manifests["control"].get(SOURCE_CHANGE_PATH) == digest(change["before_sha256"], "source-change.before_sha256"), "control production source differs")
    require(manifests["candidate"].get(SOURCE_CHANGE_PATH) == digest(change["after_sha256"], "source-change.after_sha256"), "candidate production source differs")
    differences = {path for path in set(manifests["control"]) | set(manifests["candidate"]) if manifests["control"].get(path) != manifests["candidate"].get(path)}
    require(differences == allowed, f"candidate source differs outside allowed paths: {sorted(differences - allowed)}")
    return {"initial": initial, "source_change": change, "control_manifest": manifests["control"], "candidate_manifest": manifests["candidate"]}


def _check_pilots(root: Path, binaries: Mapping[str, Mapping[str, Mapping[str, Any]]]) -> None:
    # The frozen baseline manifest is independently checked by analyze.  Its
    # four references are retained, while arm-specific pilot reports bind the
    # same corpus to both fresh binaries.
    original = analyze.ROOT
    try:
        analyze.ROOT = root
        manifest, _ = analyze._baseline_corpus()
    except analyze.AnalysisError as error:
        fail(str(error))
    finally:
        analyze.ROOT = original
    for arm in analyze.ARMS:
        for instrumentation in analyze.INSTRUMENTATIONS:
            for mode in analyze.MODES:
                label = f"pilot-{arm}-{instrumentation}-{mode}"
                path = root / f"{label}.report.json"
                require(path.is_file() and not path.is_symlink(), f"{label}: report is missing")
                value = read_json(path, label)
                # Pilot validation has no arm field in the report itself; the
                # binary record and gate receipt establish the arm binding.
                analyze._validate_pilot_report(value, label, manifest["cases"])
                expected_binary = binaries[arm][instrumentation]
                receipt = read_json(root / "validation" / f"{label}.json", f"validation/{label}.json")
                require(receipt.get("exit_code") == 0 and receipt.get("source_unchanged") is True, f"{label}: gate failed")
                argv = receipt.get("argv")
                provenance_root = Path(receipt.get("cwd", ""))
                require(provenance_root.is_absolute(), f"{label}: pilot cwd provenance differs")
                expected_argv = [
                    expected_binary["path"], "--counts", "64,8192,131072", "--mode", mode,
                    "--samples", "1", "--warmups", "1", "--json",
                    str(provenance_root / "docs/performance/results/change-0481" / f"{label}.report.json"),
                ]
                require(argv == expected_argv, f"{label}: pilot argv binding differs")


def _check_pilot_formal_order(root: Path, protocol: Mapping[str, Any]) -> None:
    """Check historic baseline and arm-local pilot/formal chronology."""
    baseline = read_json(root / "baseline-corpus.json", "baseline-corpus.json")
    frozen = timestamp(baseline.get("frozen_utc"), "baseline-corpus.json.frozen_utc")
    first_formal: dict[str, _datetime.datetime] = {}
    for spec in analyze.protocol_rows(protocol):
        started = read_json(root / "captures" / f"{spec['label']}.started.json", f"captures/{spec['label']}.started.json")
        value = timestamp(started.get("started_utc"), f"{spec['label']}.started_utc")
        first_formal[spec["arm"]] = min(first_formal.get(spec["arm"], value), value)
    require(frozen < min(first_formal.values()), "historic baseline was frozen after a formal capture started")
    for arm in analyze.ARMS:
        pilot_finished: list[_datetime.datetime] = []
        for instrumentation in analyze.INSTRUMENTATIONS:
            for mode in analyze.MODES:
                label = f"pilot-{arm}-{instrumentation}-{mode}"
                receipt = read_json(root / "validation" / f"{label}.json", f"validation/{label}.json")
                pilot_finished.append(timestamp(receipt.get("finished_utc"), f"{label}.finished_utc"))
        require(max(pilot_finished) < first_formal[arm], f"{arm} pilots were not complete before that arm's first formal capture")


def _run_gate_commands(root: Path) -> dict[str, list[str]]:
    """Read the frozen literal GATES table without executing the driver."""
    path = bundle_file(root, "run-gates.py", "run-gates.py")
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    except (OSError, SyntaxError) as error:
        fail(f"run-gates.py: cannot parse GATES table: {error}")
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(isinstance(target, ast.Name) and target.id == "GATES" for target in node.targets):
            try:
                value = ast.literal_eval(node.value)
            except (ValueError, SyntaxError) as error:
                fail(f"run-gates.py: GATES table is not literal: {error}")
            require(isinstance(value, dict), "run-gates.py: GATES table must be an object")
            result: dict[str, list[str]] = {}
            for label, argv in value.items():
                require(isinstance(label, str) and LABEL.fullmatch(label) is not None and isinstance(argv, list) and all(isinstance(item, str) for item in argv), f"run-gates.py: malformed GATES entry {label!r}")
                result[label] = list(argv)
            return result
    fail("run-gates.py: GATES table is missing")


def _expected_validation_argv(provenance_cwd: str, label: str, binaries: Mapping[str, Mapping[str, Mapping[str, Any]]]) -> list[str] | None:
    if label in BUILD_LABELS:
        arm = "candidate" if "candidate" in label else "control"
        instrumentation = "allocator" if label.endswith("-allocator") else "normal"
        result = ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin", "docx_plain_paragraph_tail_append"]
        if instrumentation == "allocator":
            result += ["--features", "allocator-metrics"]
        return result
    if label in PILOT_LABELS:
        provenance_root = Path(provenance_cwd)
        require(provenance_root.is_absolute(), f"{label}: pilot cwd provenance differs")
        parts = label.split("-")
        arm = parts[1]
        instrumentation = parts[2]
        mode = parts[3]
        return [
            binaries[arm][instrumentation]["path"], "--counts", "64,8192,131072", "--mode", mode,
            "--samples", "1", "--warmups", "1", "--json",
            str(provenance_root / "docs/performance/results/change-0481" / f"{label}.report.json"),
        ]
    return None


def check_validation_receipts(root: Path, binaries: Mapping[str, Mapping[str, Mapping[str, Any]]], protocol: Mapping[str, Any]) -> dict[str, Any]:
    directory = root / "validation"
    gate_path = bundle_file(root, "gate.py", "gate.py")
    common_path = bundle_file(root, "common.py", "common.py")
    driver_hash = sha256_file(gate_path, "gate.py")[0]
    common_hash = sha256_file(common_path, "common.py")[0]
    gate_commands = _run_gate_commands(root)
    require(RUN_LABELS <= set(gate_commands), "run-gates.py: required run gate command is missing")
    actual_labels = {
        path.stem for path in directory.glob("*.json")
        if path.is_file() and not path.name.endswith(".started.json")
    }
    require(set(REQUIRED_LABELS) <= actual_labels, f"validation receipts omit required labels: {sorted(set(REQUIRED_LABELS) - actual_labels)}")
    successes = 0
    for label in sorted(actual_labels):
        path = bundle_file(root, f"validation/{label}.json", f"validation/{label}.json")
        value = read_json(path, f"validation/{label}.json")
        require(isinstance(value, dict), f"{label}: receipt must be an object")
        if label in REQUIRED_LABELS:
            require(value.get("exit_code") == 0 and value.get("source_unchanged") is True, f"{label}: required receipt failed")
        interval(value.get("started_utc"), value.get("finished_utc"), label)
        require(value.get("driver_sha256") == driver_hash and value.get("common_sha256") == common_hash, f"{label}: driver/common binding differs")
        require(value.get("environment") == protocol["environment"], f"{label}: environment binding differs")
        started_path = bundle_file(root, f"validation/{label}.started.json", f"validation/{label}.started.json")
        started = read_json(started_path, f"validation/{label}.started.json")
        require(isinstance(started, dict), f"{label}: started receipt is malformed")
        for field in ("argv", "cwd", "environment", "driver_sha256", "common_sha256", "started_utc", "source_before"):
            require(value.get(field) == started.get(field), f"{label}: terminal receipt does not preserve started {field}")
        if label in RUN_LABELS:
            require(value.get("argv") == gate_commands[label], f"{label}: argv differs from run-gates.py GATES")
        expected_argv = _expected_validation_argv(value.get("cwd", ""), label, binaries)
        if expected_argv is not None:
            require(value.get("argv") == expected_argv, f"{label}: argv differs from frozen build/pilot binding")
        expected_artifacts = {f"{label}.stdout", f"{label}.stderr"}
        _artifact_check(directory, value.get("artifacts"), expected_artifacts, label)
        source = value.get("source_before")
        after = value.get("source_after")
        require(isinstance(source, dict) and isinstance(after, dict) and source == after, f"{label}: source changed during gate")
        require(value.get("source_exclusions", []) == [], f"{label}: source exclusions are forbidden")
        if label in BUILD_LABELS or label in PILOT_LABELS:
            arm = "candidate" if "candidate" in label else "control"
            expected_hash = binaries[arm]["normal"]["source_manifest_sha256"]
            require(source.get("sha256") == expected_hash, f"{label}: source manifest arm differs")
        elif label in REQUIRED_LABELS or label.startswith("profile-"):
            require(source.get("sha256") == binaries["candidate"]["normal"]["source_manifest_sha256"], f"{label}: final gate is not bound to candidate source")
        else:
            require(source.get("sha256") in {
                binaries["control"]["normal"]["source_manifest_sha256"],
                binaries["candidate"]["normal"]["source_manifest_sha256"],
            }, f"{label}: source manifest is not one of the two frozen arms")
        if label in REQUIRED_LABELS:
            successes += 1
    return {"required_success": successes, "required": len(REQUIRED_LABELS), "attempts": len(actual_labels)}


def _compare(expected: Any, actual: Any, label: str) -> None:
    if isinstance(expected, bool) or isinstance(actual, bool):
        require(expected == actual, f"{label}: value differs")
    elif isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        require(expected == actual or abs(float(expected) - float(actual)) <= 1e-9 * max(1.0, abs(float(expected)), abs(float(actual))), f"{label}: numeric value differs")
    elif isinstance(expected, dict) and isinstance(actual, dict):
        require(set(expected) == set(actual), f"{label}: object fields differ")
        for key in expected:
            _compare(expected[key], actual[key], f"{label}.{key}")
    elif isinstance(expected, list) and isinstance(actual, list):
        require(len(expected) == len(actual), f"{label}: list length differs")
        for index, (left, right) in enumerate(zip(expected, actual)):
            _compare(left, right, f"{label}[{index}]")
    else:
        require(expected == actual, f"{label}: value differs")


def check_summary(root: Path) -> None:
    summary = read_json(root / "summary.json", "summary.json")
    original = analyze.ROOT
    try:
        analyze.ROOT = root
        expected = analyze.derive()
    except analyze.AnalysisError as error:
        fail(str(error))
    finally:
        analyze.ROOT = original
    _compare(expected, summary, "summary")


def check_validation_ledger(root: Path, binaries: Mapping[str, Mapping[str, Mapping[str, Any]]]) -> dict[str, Any]:
    path = bundle_file(root, "rust-validation.json", "rust-validation.json")
    value = read_json(path, "rust-validation.json")
    require(isinstance(value, dict) and value.get("schema") == VALIDATION_SCHEMA, "rust-validation schema differs")
    require(value.get("required") == list(REQUIRED_LABELS), "rust-validation required labels differ")
    source = value.get("source_sha256")
    require(isinstance(source, dict) and source == {"control": binaries["control"]["normal"]["source_manifest_sha256"], "candidate": binaries["candidate"]["normal"]["source_manifest_sha256"]}, "rust-validation source hashes differ")
    attempts = value.get("attempts")
    validation_dir = root / "validation"
    actual_labels = {
        path.stem for path in validation_dir.glob("*.json")
        if path.is_file() and not path.name.endswith(".started.json")
    }
    require(isinstance(attempts, dict) and set(attempts) == actual_labels, "rust-validation attempts do not cover every retained receipt")
    require(set(REQUIRED_LABELS) <= set(attempts), "rust-validation attempts omit a required gate")
    for label in sorted(attempts):
        item = attempts[label]
        require(isinstance(item, dict), f"rust-validation.{label}: attempt is malformed")
        if label in REQUIRED_LABELS:
            require(item.get("exit_code") == 0 and item.get("source_unchanged") is True, f"rust-validation.{label}: required attempt failed")
        require(item.get("source_exclusions", []) == [], f"rust-validation.{label}: source exclusions are forbidden")
        receipt = bundle_file(root, item.get("path"), f"rust-validation.{label}.path")
        require(sha256_file(receipt, f"rust-validation.{label}")[0] == digest(item.get("sha256"), f"rust-validation.{label}.sha256"), f"rust-validation.{label}: receipt digest differs")
        receipt_value = read_json(receipt, label)
        require(item.get("path") == receipt.relative_to(root).as_posix(), f"rust-validation.{label}: path binding differs")
        for field in ("argv", "exit_code", "source_unchanged"):
            require(item.get(field) == receipt_value.get(field), f"rust-validation.{label}: {field} binding differs")
        if label in REQUIRED_LABELS:
            require(receipt_value.get("exit_code") == 0 and receipt_value.get("source_unchanged") is True, f"rust-validation.{label}: receipt failed")
    return {"required_success": len(REQUIRED_LABELS), "receipts": len(attempts), "required": len(REQUIRED_LABELS)}


def _check_reference(root: Path, reference: Any, label: str) -> Path:
    require(isinstance(reference, dict) and set(reference) == {"path", "bytes", "sha256"}, f"{label}: metadata reference differs")
    path = bundle_file(root, reference["path"], f"{label}.path")
    check_metadata(path, {"bytes": reference["bytes"], "sha256": reference["sha256"]}, label)
    return path


def check_profiles(root: Path, binaries: Mapping[str, Mapping[str, Mapping[str, Any]]]) -> dict[str, Any] | None:
    path = root / "profiles.json"
    if not path.exists():
        return None
    value = read_json(path, "profiles.json")
    require(isinstance(value, dict) and value.get("schema") == "docx-borrowed-names-paired-pmu-v1", "profiles schema differs")
    require(value.get("formal_samples") is False and value.get("excluded_samples") == 60, "profiles sample scope differs")
    records = value.get("records")
    require(isinstance(records, dict) and set(records) == set(analyze.ARMS), "profiles arm records differ")
    for arm in analyze.ARMS:
        record = records[arm]
        require(record.get("exit_code") == 0, f"profiles.{arm}: perf process failed")
        for field in ("binary", "receipt", "counters", "report", "values"):
            require(field in record, f"profiles.{arm}: {field} missing")
        require(record["binary"] == binaries[arm]["normal"], f"profiles.{arm}: binary identity differs")
        receipt_path = _check_reference(root, record["receipt"], f"profiles.{arm}.receipt")
        counters_path = _check_reference(root, record["counters"], f"profiles.{arm}.counters")
        report_path = _check_reference(root, record["report"], f"profiles.{arm}.report")
        receipt = read_json(receipt_path, f"profiles.{arm}.receipt")
        require(receipt.get("exit_code") == 0 and receipt.get("source_unchanged") is True, f"profiles.{arm}: perf receipt failed")
        provenance_root = Path(receipt.get("cwd", ""))
        require(provenance_root.is_absolute(), f"profiles.{arm}: receipt cwd provenance differs")
        expected_argv = [
            "/usr/bin/taskset", "-c", "2", "perf", "stat", "-x,", "-o",
            str(provenance_root / "docs/performance/results/change-0481" / "profiles" / f"{arm}.csv"),
            "-e", "cycles,instructions,branches,branch-misses,cache-misses,page-faults", "--",
            record["binary"]["path"], "--mode", "total", "--counts", "131072", "--samples", "30", "--warmups", "3", "--json",
            str(provenance_root / "docs/performance/results/change-0481" / "profiles" / f"{arm}.report.json"),
        ]
        require(receipt.get("argv") == expected_argv, f"profiles.{arm}: perf stat argv differs")
        report = read_json(report_path, f"profiles.{arm}.report")
        spec = {"arm": arm, "instrumentation": "normal", "count": 131072, "mode": "total"}
        analyze.validate_report(report, spec, f"profiles.{arm}.report")
        expected_events = {"cycles", "instructions", "branches", "branch-misses", "cache-misses", "page-faults"}
        require(set(record["values"]) == expected_events, f"profiles.{arm}: PMU events differ")
        import csv
        observed: dict[str, dict[str, Any]] = {}
        with counters_path.open(newline="", encoding="utf-8") as stream:
            for row in csv.reader(stream):
                if not row or row[0].startswith("#"):
                    continue
                require(len(row) >= 5 and row[2] in expected_events and row[2] not in observed, f"profiles.{arm}: malformed/duplicate PMU row")
                try:
                    counter_value = int(row[0])
                    running = float(row[4])
                except (ValueError, IndexError) as error:
                    fail(f"profiles.{arm}: PMU value is not numeric: {error}")
                require(counter_value >= 0 and math.isfinite(running), f"profiles.{arm}: PMU value is invalid")
                observed[row[2]] = {"value": counter_value, "running_percent": running}
        require(observed == record["values"], f"profiles.{arm}: recorded PMU values differ from CSV")
    return {"status": value.get("status"), "arms": len(records)}


def check_seal(root: Path, required: bool = True) -> bool:
    path = root / "SHA256SUMS"
    if not path.exists():
        require(not required, "SHA256SUMS is missing")
        return False
    require(path.is_file() and not path.is_symlink(), "SHA256SUMS: regular file required")
    entries: dict[str, str] = {}
    for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        fields = line.split(maxsplit=1)
        require(len(fields) == 2, f"SHA256SUMS:{number}: malformed line")
        name = fields[1][1:] if fields[1].startswith("*") else fields[1]
        name = safe_relative(name, f"SHA256SUMS:{number}.path")
        require(name != "SHA256SUMS" and name not in entries, f"SHA256SUMS:{number}: duplicate/self reference")
        entries[name] = digest(fields[0], f"SHA256SUMS:{number}.sha256")
    require(entries, "SHA256SUMS: no entries")
    for name, expected in entries.items():
        require(sha256_file(bundle_file(root, name, f"SHA256SUMS.{name}"), name)[0] == expected, f"SHA256SUMS: digest differs for {name}")
    actual = {candidate.relative_to(root).as_posix() for candidate in root.rglob("*") if candidate.is_file() and candidate.name != "SHA256SUMS"}
    require(actual == set(entries), "SHA256SUMS: file inventory differs")
    require({"protocol.json", "summary.json"} <= set(entries), "SHA256SUMS: protocol/summary missing")
    return True


def verify_bundle(root: Path = ROOT, *, data_only: bool = False) -> dict[str, Any]:
    root = Path(root).resolve()
    require(root.is_dir() and not root.is_symlink(), "evidence root is not a directory")
    protocol, protocol_hash = check_protocol(root)
    check_environment(root, protocol)
    binaries = check_binaries(root, protocol)
    _check_source_change(root, binaries)
    _check_pilots(root, binaries)
    _check_pilot_formal_order(root, protocol)
    reports = check_capture_receipts(root, protocol, protocol_hash, binaries)
    require(len(reports) == 48, "capture count differs")
    require((root / "summary.json").is_file(), "summary.json is missing")
    check_summary(root)
    if data_only:
        return {"schema": VERIFICATION_SCHEMA, "protocol_sha256": protocol_hash, "captures": len(reports), "samples": len(reports) * analyze.SAMPLES, "data_only": True, "profiles": check_profiles(root, binaries), "sealed": check_seal(root, required=False), "status": "pass"}
    validation = check_validation_receipts(root, binaries, protocol)
    ledger = check_validation_ledger(root, binaries)
    profiles = check_profiles(root, binaries)
    import closure
    closure_result = closure.check(root, complete=True)
    sealed = check_seal(root, required=True)
    return {"schema": VERIFICATION_SCHEMA, "protocol_sha256": protocol_hash, "captures": len(reports), "samples": len(reports) * analyze.SAMPLES, "data_only": False, "validation": validation, "rust_validation": ledger, "profiles": profiles, "sealed": sealed, "closure": closure_result, "status": "pass"}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--data-only", action="store_true")
    args = parser.parse_args(argv)
    try:
        print(json.dumps(verify_bundle(args.root, data_only=args.data_only), sort_keys=True))
    except VerificationError as error:
        print(f"verify.py: FAIL: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
