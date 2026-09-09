#!/usr/bin/env python3
"""Fail-closed portable verifier for the 0483 evidence bundle.

It checks the frozen protocol, binary custody, all capture receipts and the
independently recomputed summary.  It never needs the temporary executables;
their absence after cleanup is expected.  A copied bundle can therefore be
validated from a different directory without touching the repository source.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
from typing import Any, NoReturn
import zipfile

import analyze


ROOT = Path(__file__).resolve().parent
SHA256 = re.compile(r"^[0-9a-f]{64}$")
LABEL = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
VERIFICATION_SCHEMA = "docx-tail-append-verification-v1"


class VerificationError(ValueError):
    pass


def fail(message: str) -> NoReturn:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def timestamp(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str) and value, f"{label}: timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: invalid timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label}: timezone is required")
    return parsed


def interval(start: Any, finish: Any, label: str) -> tuple[_datetime.datetime, _datetime.datetime]:
    started = timestamp(start, f"{label}.started_utc")
    ended = timestamp(finish, f"{label}.finished_utc")
    require(ended >= started, f"{label}: finished before started")
    return started, ended


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label}: regular file required")
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=lambda value: (_ for _ in ()).throw(ValueError(value)),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"{label}: invalid JSON: {error}")


def sha256_file(path: Path, label: str) -> tuple[str, int]:
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


def metadata(path: Path, label: str) -> dict[str, Any]:
    value, size = sha256_file(path, label)
    return {"bytes": size, "sha256": value}


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label}: relative path required")
    require("\\" not in value, f"{label}: backslashes are forbidden")
    path = PurePosixPath(value)
    require(not path.is_absolute() and value not in {".", ".."}, f"{label}: absolute/parent path forbidden")
    require(all(part not in {"", ".", ".."} for part in path.parts), f"{label}: path escapes bundle")
    return path.as_posix()


def bundle_file(value: Any, label: str) -> Path:
    relative = safe_relative(value, label)
    current = ROOT
    for part in PurePosixPath(relative).parts:
        current /= part
        require(not current.is_symlink(), f"{label}: symlink component forbidden")
    try:
        resolved = current.resolve(strict=True)
    except OSError as error:
        fail(f"{label}: cannot resolve: {error}")
    base = ROOT.resolve()
    require(resolved == base or base in resolved.parents, f"{label}: path escapes bundle")
    require(current.is_file(), f"{label}: missing file")
    return current


def check_digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256.fullmatch(value) is not None, f"{label}: SHA-256 required")
    return value


def check_capture_argv(argv: Any, spec: dict[str, Any], label: str) -> None:
    require(
        isinstance(argv, list) and len(argv) == 18 and all(isinstance(item, str) for item in argv),
        f"{label}.argv: expected 18 strings",
    )
    fixed = {
        0: "/usr/bin/time", 1: "-v", 2: "-o", 4: "/usr/bin/taskset",
        5: "-c", 6: "2", 8: "--route", 10: "--counts", 12: "--samples",
        13: "30", 14: "--warmups", 15: "3", 16: "--json",
    }
    for index, expected in fixed.items():
        require(argv[index] == expected, f"{label}.argv[{index}]: expected {expected!r}")
    require(argv[9] == spec["route"] and argv[11] == str(spec["count"]), f"{label}: route/count are not bound")
    executable = Path(argv[7])
    require(
        executable.is_absolute()
        and executable.as_posix().endswith(
            f"/{spec['attempt']}/{spec['instrumentation']}/docx_bounded_tail_append_compare"
        ),
        f"{label}.argv[7]: binary binding differs",
    )
    require(
        Path(argv[3]).is_absolute() and argv[3].endswith(f"/captures/{spec['label']}.resource"),
        f"{label}.argv[3]: resource path does not bind label",
    )
    require(
        Path(argv[17]).is_absolute() and argv[17].endswith(f"/captures/{spec['label']}.report.json"),
        f"{label}.argv[17]: report path does not bind label",
    )


def check_metadata(path: Path, value: Any, label: str) -> None:
    require(isinstance(value, dict) and set(value) == {"bytes", "sha256"}, f"{label}: metadata fields differ")
    require(isinstance(value["bytes"], int) and value["bytes"] >= 0, f"{label}.bytes malformed")
    check_digest(value["sha256"], f"{label}.sha256")
    require(metadata(path, label) == value, f"{label}: metadata differs")


def check_protocol(protocol: dict[str, Any]) -> list[dict[str, Any]]:
    try:
        rows = analyze.protocol_rows(protocol)
    except analyze.AnalysisError as error:
        fail(str(error))
    for name, expected in protocol["scripts"].items():
        path = bundle_file(name, f"protocol.scripts.{name}")
        actual, _ = sha256_file(path, name)
        require(actual == expected, f"protocol script digest differs: {name}")
    plans = protocol.get("plan_files")
    require(isinstance(plans, dict) and set(plans) == {"validation-plan.json", "fuzz-plan.json"}, "protocol plan-file custody differs")
    for name, expected in plans.items():
        path = bundle_file(name, f"protocol.plan_files.{name}")
        require(sha256_file(path, name)[0] == check_digest(expected, f"protocol.plan_files.{name}"), f"protocol plan digest differs: {name}")
    return rows


def check_source_manifest(manifest_sha: Any, label: str) -> dict[str, Any]:
    digest_value = check_digest(manifest_sha, f"{label}.sha256")
    path = bundle_file(f"validation-sources/{digest_value}.json", label)
    actual, _ = sha256_file(path, label)
    require(actual == digest_value, f"{label}: digest differs")
    value = read_json(path, label)
    require(isinstance(value, dict) and value, f"{label}: manifest is empty")
    for source, source_sha in value.items():
        require(isinstance(source, str), f"{label}: source path is malformed")
        safe_relative(source, f"{label}.{source}")
        check_digest(source_sha, f"{label}.{source}")
    return value


def check_source_reference(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and set(value) == {"path", "sha256", "files"}, f"{label}: source reference fields differ")
    manifest = check_source_manifest(value.get("sha256"), label)
    require(value.get("path") == f"validation-sources/{value['sha256']}.json", f"{label}: source path is not content-addressed")
    require(value.get("files") == len(manifest), f"{label}: source file count differs")
    return manifest


def check_gate_receipt(
    build: dict[str, Any],
    protocol: dict[str, Any],
    instrumentation: str,
    label: str,
) -> dict[str, Any]:
    gate_path = bundle_file(build["gate_path"], f"{label}.gate_path")
    gate_sha, _ = sha256_file(gate_path, f"{label}.gate")
    require(gate_sha == build["gate_sha256"], f"{label}: gate digest differs")
    gate = read_json(gate_path, f"{label}.gate")
    require(build.get("gate") == gate, f"{label}: embedded gate receipt differs")
    expected_fields = {
        "schema", "label", "attempt", "argv", "cwd", "environment", "driver_sha256",
        "common_sha256", "started_utc", "source_before", "exit_code", "finished_utc",
        "source_after", "source_unchanged", "artifacts",
    }
    require(isinstance(gate, dict) and set(gate) == expected_fields, f"{label}: gate receipt fields differ")
    require(gate.get("schema") == "docx-tail-append-gate-v1", f"{label}: gate schema differs")
    require(gate.get("label") == build["gate_label"] and gate.get("attempt") == build["attempt"], f"{label}: gate identity differs")
    command = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", build["binary_name"],
    ]
    if instrumentation == "allocator":
        command.extend(["--features", "allocator-metrics"])
    require(gate.get("argv") == command == build.get("command"), f"{label}: build argv differs")
    require(isinstance(gate.get("cwd"), str) and Path(gate["cwd"]).is_absolute(), f"{label}: gate cwd differs")
    require(gate.get("environment") == protocol.get("environment"), f"{label}: gate environment differs")
    require(gate.get("driver_sha256") == hashlib.sha256((ROOT / "gate.py").read_bytes()).hexdigest(), f"{label}: gate driver binding differs")
    require(gate.get("common_sha256") == hashlib.sha256((ROOT / "common.py").read_bytes()).hexdigest(), f"{label}: common binding differs")
    started_path = gate_path.with_suffix(".started.json")
    started = read_json(started_path, f"{label}.started")
    expected_started = expected_fields - {"exit_code", "finished_utc", "source_after", "source_unchanged", "artifacts"}
    require(isinstance(started, dict) and set(started) == expected_started, f"{label}: started gate fields differ")
    for field in expected_started:
        require(started.get(field) == gate.get(field), f"{label}: terminal gate changed {field}")
    _, finished = interval(gate.get("started_utc"), gate.get("finished_utc"), label)
    require(gate.get("exit_code") == 0 and gate.get("source_unchanged") is True, f"{label}: build gate failed")
    require(gate.get("source_before") == gate.get("source_after"), f"{label}: source changed during build gate")
    source_before = gate.get("source_before")
    require(isinstance(source_before, dict), f"{label}: source manifest reference missing")
    check_source_reference(source_before, f"{label}.source_before")
    check_source_reference(gate.get("source_after"), f"{label}.source_after")
    artifacts = gate.get("artifacts")
    expected_artifacts = {f"{build['gate_label']}.stdout", f"{build['gate_label']}.stderr"}
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts, f"{label}: gate artifact inventory differs")
    for name, expected in artifacts.items():
        artifact_path = ROOT / "validation" / name
        require(artifact_path.is_file() and not artifact_path.is_symlink(), f"{label}: gate artifact missing {name}")
        check_metadata(artifact_path, expected, f"{label}.{name}")
    copied = timestamp(build.get("copied_utc"), f"{label}.copied_utc")
    require(copied >= finished, f"{label}: binary copied before build gate finished")
    return gate


def check_builds(protocol: dict[str, Any]) -> dict[str, Any]:
    attempt = protocol["attempt"]
    path = bundle_file(f"builds/binaries-{attempt}.json", "binary custody")
    value = read_json(path, "binary custody")
    require(value.get("schema") == "docx-tail-append-binaries-v1" and value.get("attempt") == attempt, "binary custody schema differs")
    require(value.get("binary_name") == "docx_bounded_tail_append_compare", "binary custody executable differs")
    binaries = value.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"normal", "allocator"}, "binary instrumentation identities differ")
    require(value.get("source_manifest_sha256") == binaries["normal"].get("source_manifest_sha256"), "binary custody source differs")
    source_hashes = set()
    for instrumentation, item in binaries.items():
        require(isinstance(item, dict), f"binary custody {instrumentation} malformed")
        for key in ("path", "bytes", "sha256", "build_path", "build_sha256", "source_manifest_sha256"):
            require(key in item, f"binary custody {instrumentation}.{key} missing")
        require(isinstance(item["path"], str) and Path(item["path"]).is_absolute(), f"binary custody {instrumentation}.path malformed")
        require(
            Path(item["path"]).as_posix().endswith(
                f"/litchi-goal-0483/{attempt}/{instrumentation}/docx_bounded_tail_append_compare"
            ),
            f"binary custody {instrumentation}.path is not the isolated copied executable",
        )
        require(isinstance(item["bytes"], int) and item["bytes"] > 0, f"binary custody {instrumentation}.bytes malformed")
        check_digest(item["sha256"], f"binary custody {instrumentation}.sha256")
        check_digest(item["build_sha256"], f"binary custody {instrumentation}.build_sha256")
        check_digest(item["source_manifest_sha256"], f"binary custody {instrumentation}.source_manifest_sha256")
        build_path = bundle_file(item["build_path"], f"binary custody {instrumentation}.build_path")
        actual_build_sha, _ = sha256_file(build_path, f"binary build {instrumentation}")
        require(actual_build_sha == item["build_sha256"], f"binary build {instrumentation}: digest differs")
        build = read_json(build_path, f"binary build {instrumentation}")
        required_build = {
            "schema", "attempt", "instrumentation", "binary_name", "command", "gate_label",
            "gate_path", "gate_sha256", "gate", "source_manifest_sha256", "copied_utc",
            "original_binary", "binary",
        }
        require(isinstance(build, dict) and set(build) == required_build, f"binary build {instrumentation}: fields differ")
        require(build.get("attempt") == attempt and build.get("instrumentation") == instrumentation, f"binary build {instrumentation}: identity differs")
        require(build.get("binary_name") == value["binary_name"], f"binary build {instrumentation}: executable differs")
        require(build.get("source_manifest_sha256") == item["source_manifest_sha256"] == value["source_manifest_sha256"], f"binary build {instrumentation}: source differs")
        require(build.get("binary", {}).get("path") == item["path"], f"binary build {instrumentation}: copied path differs")
        require(build.get("binary", {}).get("bytes") == item["bytes"] and build.get("binary", {}).get("sha256") == item["sha256"], f"binary build {instrumentation}: copied metadata differs")
        original = build.get("original_binary")
        require(isinstance(original, dict) and set(original) == {"path", "bytes", "sha256"}, f"binary build {instrumentation}: original metadata differs")
        require(isinstance(original["path"], str) and Path(original["path"]).is_absolute(), f"binary build {instrumentation}: original path differs")
        require(
            Path(original["path"]).as_posix().endswith(
                f"/tools/perf-baseline/target/release/docx_bounded_tail_append_compare"
            ),
            f"binary build {instrumentation}: original Cargo output differs",
        )
        require(original["bytes"] == item["bytes"] and original["sha256"] == item["sha256"], f"binary build {instrumentation}: original metadata differs")
        check_gate_receipt(build, protocol, instrumentation, f"binary build {instrumentation}")
        check_source_manifest(item["source_manifest_sha256"], f"binary build {instrumentation}.source_manifest")
        if Path(item["path"]).is_file():
            require(not Path(item["path"]).is_symlink(), f"binary build {instrumentation}: copied binary is symlinked")
            require(metadata(Path(item["path"]), f"binary build {instrumentation}.copied binary") == {"bytes": item["bytes"], "sha256": item["sha256"]}, f"binary build {instrumentation}: copied binary changed")
        source_hashes.add(item["source_manifest_sha256"])
    require(len(source_hashes) == 1, "normal and allocator source manifests differ")
    manifest_path = bundle_file(f"validation-sources/{next(iter(source_hashes))}.json", "source manifest")
    actual, _ = sha256_file(manifest_path, "source manifest")
    require(actual == next(iter(source_hashes)), "source manifest digest differs")
    return value


def check_captures(
    protocol: dict[str, Any], rows: list[dict[str, Any]], binaries: dict[str, Any]
) -> dict[str, tuple[_datetime.datetime, _datetime.datetime]]:
    capture_dir = ROOT / "captures"
    require(capture_dir.is_dir() and not capture_dir.is_symlink(), "captures directory is missing")
    previous_finished: _datetime.datetime | None = None
    capture_times: dict[str, tuple[_datetime.datetime, _datetime.datetime]] = {}
    for spec in rows:
        label = spec["label"]
        base = capture_dir / label
        require(not base.is_symlink(), f"{label}: capture prefix must not be a symlink")
        started = read_json(base.with_suffix(".started.json"), f"{label}.started")
        receipt = read_json(base.with_suffix(".json"), f"{label}.receipt")
        expected_started = {"schema", "capture", "cwd", "started_utc", "protocol_sha256", "binary", "environment"}
        expected_receipt = expected_started | {"exit_code", "finished_utc", "artifacts"}
        require(isinstance(started, dict) and set(started) == expected_started, f"{label}: started envelope differs")
        require(isinstance(receipt, dict) and set(receipt) == expected_receipt, f"{label}: final envelope differs")
        protocol_hash = hashlib.sha256((ROOT / "protocol.json").read_bytes()).hexdigest()
        require(started.get("capture") == spec, f"{label}: started capture spec differs")
        require(receipt.get("capture") == spec, f"{label}: final capture spec differs")
        require(started.get("protocol_sha256") == protocol_hash, f"{label}: started protocol digest differs")
        require(receipt.get("protocol_sha256") == protocol_hash, f"{label}: final protocol digest differs")
        require(started.get("environment") == protocol.get("environment"), f"{label}: started environment differs")
        require(receipt.get("environment") == started.get("environment"), f"{label}: final environment differs")
        require(started.get("binary") == binaries["binaries"][spec["instrumentation"]], f"{label}: started binary differs")
        require(receipt.get("binary") == started.get("binary"), f"{label}: final binary differs")
        require(isinstance(started.get("cwd"), str) and Path(started["cwd"]).is_absolute(), f"{label}: cwd provenance differs")
        require(receipt.get("cwd") == started.get("cwd"), f"{label}: final cwd differs")
        check_capture_argv(spec.get("argv"), spec, label)
        require(spec["argv"][7] == started["binary"]["path"], f"{label}: executable does not bind binary")
        started_at, finished_at = interval(started.get("started_utc"), receipt.get("finished_utc"), label)
        capture_times[label] = (started_at, finished_at)
        if previous_finished is not None:
            require(started_at >= previous_finished, f"{label}: capture chronology overlaps the preceding capture")
        previous_finished = finished_at
        require(receipt.get("exit_code") == 0, f"{label}: process failed")
        artifacts = receipt.get("artifacts")
        expected_artifacts = {f"{label}.stdout", f"{label}.stderr", f"{label}.resource", f"{label}.report.json"}
        require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts, f"{label}: artifact inventory differs")
        for suffix in (".stdout", ".stderr", ".resource", ".report.json"):
            path = base.with_suffix(suffix)
            require(path.is_file(), f"{label}: missing {suffix}")
            require(not path.is_symlink(), f"{label}: symlink artifact {suffix}")
            check_metadata(path, artifacts[path.name], f"{label}.{suffix}")
        resource = base.with_suffix(".resource").read_text(encoding="utf-8")
        require(
            len(re.findall(r"^\s*Maximum resident set size \(kbytes\):\s*\d+\s*$", resource, re.MULTILINE)) == 1,
            f"{label}: RSS resource observation differs",
        )
        report = read_json(base.with_suffix(".report.json"), f"{label}.report")
        try:
            analyze.validate_report(report, spec, label)
        except analyze.AnalysisError as error:
            fail(str(error))
    return capture_times


def validation_plan(protocol: dict[str, Any]) -> dict[str, Any]:
    value = protocol.get("validation")
    require(isinstance(value, dict), "protocol.validation is missing")
    required = value.get("required_labels")
    pilots = value.get("pilot_labels")
    argv = value.get("argv")
    pilot_reports = value.get("pilot_reports")
    developmental = value.get("developmental")
    require(isinstance(required, list) and required, "protocol.validation.required_labels is missing")
    require(isinstance(pilots, list), "protocol.validation.pilot_labels is missing")
    require(isinstance(argv, dict), "protocol.validation.argv is missing")
    require(isinstance(pilot_reports, dict), "protocol.validation.pilot_reports is missing")
    require(isinstance(developmental, dict), "protocol.validation.developmental is missing")
    labels = list(required) + list(pilots)
    require(all(isinstance(label, str) and LABEL.fullmatch(label) for label in labels), "protocol.validation labels are not path-safe")
    require(len(set(labels)) == len(labels), "protocol.validation labels are not unique")
    require(set(argv) == set(labels), "protocol.validation.argv does not cover final labels")
    require(all(isinstance(command, list) and all(isinstance(item, str) for item in command) for command in argv.values()), "protocol.validation.argv is malformed")
    require(set(pilot_reports) == set(pilots), "protocol.validation.pilot_reports does not cover pilot labels")
    for label, report in pilot_reports.items():
        require(
            isinstance(report, dict)
            and set(report) == {"path", "bytes", "sha256", "spec", "samples", "warmups"},
            f"protocol.validation.pilot_reports.{label} is malformed",
        )
        require(isinstance(report["spec"], dict), f"protocol.validation.pilot_reports.{label}.spec is malformed")
        require(type(report["samples"]) is int and report["samples"] > 0, f"protocol.validation.pilot_reports.{label}.samples is invalid")
        require(type(report["warmups"]) is int and report["warmups"] > 0, f"protocol.validation.pilot_reports.{label}.warmups is invalid")
        require(isinstance(report["path"], str), f"protocol.validation.pilot_reports.{label}.path is malformed")
        check_digest(report["sha256"], f"protocol.validation.pilot_reports.{label}.sha256")
    for label, item in developmental.items():
        require(
            isinstance(label, str) and LABEL.fullmatch(label) and isinstance(item, dict)
            and set(item) == {"classification", "reason", "source_before_sha256", "source_after_sha256", "current_source_differs"},
            f"protocol.validation.developmental.{label} is malformed",
        )
        require(item["classification"] in {"developmental", "historical"}, f"protocol.validation.developmental.{label}.classification is invalid")
        require(isinstance(item["reason"], str) and item["reason"], f"protocol.validation.developmental.{label}.reason is missing")
        check_digest(item["source_before_sha256"], f"protocol.validation.developmental.{label}.source_before_sha256")
        check_digest(item["source_after_sha256"], f"protocol.validation.developmental.{label}.source_after_sha256")
        require(isinstance(item["current_source_differs"], bool), f"protocol.validation.developmental.{label}.current_source_differs is invalid")
    require(not set(developmental) & set(labels), "protocol.validation.developmental overlaps final labels")
    return {
        "required": list(required),
        "pilots": list(pilots),
        "argv": argv,
        "labels": labels,
        "pilot_reports": pilot_reports,
        "developmental": developmental,
    }


def check_pilot_argv(
    argv: Any,
    receipt: dict[str, Any],
    pilot: dict[str, Any],
    binary: dict[str, Any],
    report_path: Path,
    label: str,
) -> None:
    """Bind a pilot receipt to the exact harness configuration it proves.

    Pilot reports are deliberately validated with the same independent schema
    and corpus oracles as formal captures.  This second check binds that
    report to the process command as well, so a valid report cannot be copied
    into a receipt for a different route, count, sample shape, or executable.
    """

    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv), f"{label}: pilot argv is malformed")
    binary_path = str(binary["path"])
    if len(argv) == 11:
        executable_index = 0
        option_start = 1
        report_index = 10
    elif len(argv) == 18:
        # This is the formal capture wrapper with pilot-sized sample counts.
        # Keep the wrapper shape fixed so an arbitrary shell command cannot be
        # smuggled through the pilot plan.
        require(
            argv[:7] == ["/usr/bin/time", "-v", "-o", argv[3], "/usr/bin/taskset", "-c", "2"]
            and Path(argv[3]).is_absolute()
            and argv[3].endswith(
                f"/{(pilot['path'][:-len('.report.json')] + '.resource' if pilot['path'].endswith('.report.json') else Path(pilot['path']).with_suffix('.resource').as_posix())}"
            ),
            f"{label}: timed pilot wrapper differs",
        )
        executable_index = 7
        option_start = 8
        report_index = 17
        require(argv[16] == "--json", f"{label}: timed pilot report flag differs")
    else:
        fail(f"{label}: pilot argv must contain either 11 direct or 18 timed/taskset arguments")
    require(argv[executable_index] == binary_path, f"{label}: pilot executable does not bind accepted binary")
    expected_options = [
        "--route",
        pilot["spec"]["route"],
        "--counts",
        str(pilot["spec"]["count"]),
        "--samples",
        str(pilot["samples"]),
        "--warmups",
        str(pilot["warmups"]),
        "--json",
    ]
    require(argv[option_start:option_start + 9] == expected_options, f"{label}: pilot argv is not bound to its report spec")
    require(Path(argv[report_index]).is_absolute(), f"{label}: pilot report argument must be absolute")
    expected_suffix = f"/{pilot['path']}"
    require(argv[report_index].endswith(expected_suffix), f"{label}: pilot report argument does not bind its report path")
    require(report_path.resolve() == ROOT.joinpath(pilot["path"]).resolve(), f"{label}: pilot report path escaped the bundle")
    require(isinstance(receipt.get("cwd"), str) and Path(receipt["cwd"]).is_absolute(), f"{label}: pilot cwd is malformed")


def _historical_helper(path_name: str, expected: str) -> Path | None:
    current = ROOT / path_name
    if current.is_file() and hashlib.sha256(current.read_bytes()).hexdigest() == expected:
        return current
    candidates = sorted((ROOT / "driver-history").glob(f"{expected}/{path_name}"))
    for candidate in candidates:
        if candidate.is_file() and hashlib.sha256(candidate.read_bytes()).hexdigest() == expected:
            return candidate
    return None


def check_validation_receipts(protocol: dict[str, Any], binaries: dict[str, Any]) -> dict[str, Any]:
    plan = validation_plan(protocol)
    directory = ROOT / "validation"
    require(directory.is_dir() and not directory.is_symlink(), "validation directory is missing")
    terminal_paths = sorted(
        path for path in directory.glob("*.json")
        if path.is_file() and not path.is_symlink() and not path.name.endswith(".started.json")
    )
    started_paths = sorted(
        path for path in directory.glob("*.started.json")
        if path.is_file() and not path.is_symlink()
    )
    terminal_labels = {path.stem for path in terminal_paths}
    started_labels = {path.name.removesuffix(".started.json") for path in started_paths}
    require(started_labels == terminal_labels, "validation receipts have orphaned started/terminal records")
    require(set(plan["labels"]) <= terminal_labels, f"validation receipts omit final labels: {sorted(set(plan['labels']) - terminal_labels)}")
    developmental_labels = set(plan["developmental"])
    require(
        terminal_labels == set(plan["labels"]) | developmental_labels,
        f"validation receipts are not fully classified: {sorted(terminal_labels - (set(plan['labels']) | developmental_labels))}",
    )
    final_source = binaries["binaries"]["normal"]["source_manifest_sha256"]
    final_times: dict[str, tuple[_datetime.datetime, _datetime.datetime]] = {}
    for path in terminal_paths:
        label = path.stem
        receipt = read_json(path, f"validation/{label}.json")
        started = read_json(directory / f"{label}.started.json", f"validation/{label}.started.json")
        fields = {
            "schema", "label", "attempt", "argv", "cwd", "environment", "driver_sha256",
            "common_sha256", "started_utc", "source_before", "exit_code", "finished_utc",
            "source_after", "source_unchanged", "artifacts",
        }
        start_fields = fields - {"exit_code", "finished_utc", "source_after", "source_unchanged", "artifacts"}
        require(isinstance(receipt, dict) and set(receipt) == fields, f"validation/{label}: receipt fields differ")
        require(isinstance(started, dict) and set(started) == start_fields, f"validation/{label}.started: fields differ")
        require(receipt.get("schema") == "docx-tail-append-gate-v1", f"validation/{label}: schema differs")
        require(receipt.get("label") == label and started.get("label") == label, f"validation/{label}: label binding differs")
        require(isinstance(receipt.get("attempt"), (str, type(None))), f"validation/{label}: attempt is malformed")
        require(isinstance(receipt.get("argv"), list) and receipt["argv"] and all(isinstance(item, str) for item in receipt["argv"]), f"validation/{label}: argv is malformed")
        require(isinstance(receipt.get("cwd"), str) and Path(receipt["cwd"]).is_absolute(), f"validation/{label}: cwd is malformed")
        require(type(receipt.get("exit_code")) is int, f"validation/{label}: exit code is malformed")
        require(isinstance(receipt.get("source_unchanged"), bool), f"validation/{label}: source-unchanged flag is malformed")
        for key in start_fields:
            require(receipt.get(key) == started.get(key), f"validation/{label}: started field {key} changed")
        final_times[label] = interval(receipt.get("started_utc"), receipt.get("finished_utc"), f"validation/{label}")
        require(receipt.get("environment") == protocol.get("environment"), f"validation/{label}: environment differs")
        driver = check_digest(receipt.get("driver_sha256"), f"validation/{label}.driver_sha256")
        common = check_digest(receipt.get("common_sha256"), f"validation/{label}.common_sha256")
        require(_historical_helper("gate.py", driver) is not None, f"validation/{label}: gate helper custody is missing")
        require(_historical_helper("common.py", common) is not None, f"validation/{label}: common helper custody is missing")
        source_before = receipt.get("source_before")
        source_after = receipt.get("source_after")
        require(isinstance(source_before, dict) and isinstance(source_after, dict), f"validation/{label}: source manifests are missing")
        check_source_reference(source_before, f"validation/{label}.source_before")
        check_source_reference(source_after, f"validation/{label}.source_after")
        artifacts = receipt.get("artifacts")
        expected_artifacts = {f"{label}.stdout", f"{label}.stderr"}
        require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts, f"validation/{label}: artifact inventory differs")
        for name, expected in artifacts.items():
            artifact_path = directory / name
            require(artifact_path.is_file() and not artifact_path.is_symlink(), f"validation/{label}: missing artifact {name}")
            check_metadata(artifact_path, expected, f"validation/{label}.{name}")
        if label in plan["labels"]:
            require(receipt.get("exit_code") == 0 and receipt.get("source_unchanged") is True, f"validation/{label}: required final gate failed")
            require(receipt.get("attempt") == protocol["attempt"], f"validation/{label}: accepted attempt binding differs")
            require(receipt.get("argv") == plan["argv"][label], f"validation/{label}: argv differs from protocol")
            require(receipt.get("source_before") == receipt.get("source_after"), f"validation/{label}: source changed")
            require(receipt["source_after"].get("sha256") == final_source, f"validation/{label}: final source binding differs")
            if label in plan["pilots"]:
                pilot = plan["pilot_reports"][label]
                report_path = bundle_file(pilot["path"], f"validation/{label}.pilot_report.path")
                check_metadata(report_path, {"bytes": pilot["bytes"], "sha256": pilot["sha256"]}, f"validation/{label}.pilot_report")
                report = read_json(report_path, f"validation/{label}.pilot_report")
                spec = pilot["spec"]
                require(
                    isinstance(spec, dict)
                    and set(spec) == {"label", "route", "route_name", "instrumentation", "count", "attempt"}
                    and spec["label"] == label,
                    f"validation/{label}.pilot_report.spec differs",
                )
                require(spec["route"] in analyze.ROUTES, f"validation/{label}.pilot_report.spec.route is invalid")
                require(spec["route_name"] == analyze.ROUTE_NAMES[spec["route"]], f"validation/{label}.pilot_report.spec.route_name differs")
                require(spec["instrumentation"] in analyze.INSTRUMENTATIONS, f"validation/{label}.pilot_report.spec.instrumentation is invalid")
                require(spec["count"] in analyze.COUNTS, f"validation/{label}.pilot_report.spec.count is invalid")
                require(isinstance(spec["attempt"], str) and spec["attempt"] == protocol["attempt"], f"validation/{label}.pilot_report.spec.attempt is invalid")
                check_pilot_argv(
                    receipt["argv"],
                    receipt,
                    pilot,
                    binaries["binaries"][spec["instrumentation"]],
                    report_path,
                    f"validation/{label}",
                )
                try:
                    analyze.validate_report(
                        report,
                        spec,
                        f"validation/{label}.pilot_report",
                        expected_samples=pilot["samples"],
                        expected_warmups=pilot["warmups"],
                    )
                except analyze.AnalysisError as error:
                    fail(str(error))
        else:
            item = plan["developmental"][label]
            source_before_sha = receipt["source_before"]["sha256"]
            source_after_sha = receipt["source_after"]["sha256"]
            require(source_before_sha == item["source_before_sha256"], f"validation/{label}: developmental source-before differs")
            require(source_after_sha == item["source_after_sha256"], f"validation/{label}: developmental source-after differs")
            require(item["current_source_differs"] is (source_after_sha != final_source), f"validation/{label}: current-source difference classification differs")
    required_finished = [final_times[label][1] for label in plan["required"]]
    pilot_finished = [final_times[label][1] for label in plan["pilots"]]
    require(required_finished and pilot_finished, "protocol.validation requires both final and pilot receipts")
    return {
        "required": len(plan["required"]),
        "pilots": len(plan["pilots"]),
        "developmental": len(developmental_labels),
        "retained": len(terminal_paths),
        "times": final_times,
    }


def check_fuzz_custody(protocol: dict[str, Any]) -> dict[str, Any]:
    fuzz = protocol.get("fuzz")
    require(isinstance(fuzz, dict), "protocol.fuzz is missing")
    seed_manifest_ref = fuzz.get("seed_manifest")
    generator_ref = fuzz.get("generator")
    receipts = fuzz.get("receipts")
    require(isinstance(seed_manifest_ref, dict) and isinstance(generator_ref, dict), "protocol.fuzz seed/generator references are missing")
    require(isinstance(receipts, list) and receipts, "protocol.fuzz.receipts is missing")

    def reference(value: Any, label: str) -> Path:
        require(isinstance(value, dict) and set(value) == {"path", "bytes", "sha256"}, f"{label}: reference fields differ")
        path = bundle_file(value["path"], f"{label}.path")
        check_metadata(path, {"bytes": value["bytes"], "sha256": value["sha256"]}, label)
        return path

    seed_manifest_path = reference(seed_manifest_ref, "protocol.fuzz.seed_manifest")
    generator_path = reference(generator_ref, "protocol.fuzz.generator")
    seed_manifest = read_json(seed_manifest_path, "fuzz/seed-manifest.json")
    require(isinstance(seed_manifest, dict) and seed_manifest, "fuzz seed manifest is empty")
    seed_root_value = fuzz.get("seed_root", "fuzz/seeds")
    seed_root = safe_relative(seed_root_value, "protocol.fuzz.seed_root")
    seed_files: set[str] = set()
    for name, item in seed_manifest.items():
        require(
            isinstance(name, str)
            and Path(name).name == name
            and Path(name).suffix == ".docx"
            and LABEL.fullmatch(Path(name).stem) is not None,
            f"fuzz seed name is malformed: {name!r}",
        )
        require(isinstance(item, dict) and set(item) == {"bytes", "sha256", "main_xml_sha256"}, f"fuzz seed {name}: manifest fields differ")
        path = bundle_file(f"{seed_root}/{name}", f"fuzz seed {name}")
        check_metadata(path, {"bytes": item["bytes"], "sha256": item["sha256"]}, f"fuzz seed {name}")
        check_digest(item["main_xml_sha256"], f"fuzz seed {name}.main_xml_sha256")
        try:
            with zipfile.ZipFile(path) as archive:
                require(archive.testzip() is None, f"fuzz seed {name}: CRC check failed")
                require(archive.namelist().count("word/document.xml") == 1, f"fuzz seed {name}: main member differs")
                main = archive.read("word/document.xml")
        except (OSError, zipfile.BadZipFile, KeyError) as error:
            fail(f"fuzz seed {name}: ZIP validation failed: {error}")
        require(hashlib.sha256(main).hexdigest() == item["main_xml_sha256"], f"fuzz seed {name}: main XML digest differs")
        seed_files.add(name)
    actual_seed_files = {
        path.relative_to(ROOT / seed_root).as_posix()
        for path in (ROOT / seed_root).rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    require(actual_seed_files == seed_files, "fuzz seed directory contains unexpected or missing files")
    generator = read_json(generator_path, "fuzz/generator.json")
    require(isinstance(generator, dict), "fuzz generator record is malformed")
    require(generator.get("seeds") == len(seed_files) and generator.get("zip_crc_and_main_hashes_verified") is True, "fuzz generator seed checks differ")
    require(generator.get("generator_sha256") == hashlib.sha256((ROOT / "fuzz-seeds.py").read_bytes()).hexdigest(), "fuzz generator source digest differs")

    seen: set[str] = set()
    records_by_kind: dict[str, list[dict[str, Any]]] = {"prepared": [], "build": [], "smoke": []}
    receipt_times: dict[str, tuple[_datetime.datetime, _datetime.datetime]] = {}
    for reference_value in receipts:
        require(isinstance(reference_value, dict) and set(reference_value) == {"label", "kind", "path", "bytes", "sha256"}, "fuzz receipt reference fields differ")
        label = reference_value["label"]
        require(isinstance(label, str) and LABEL.fullmatch(label), "fuzz receipt label is not path-safe")
        require(label not in seen, f"fuzz receipt label is duplicated: {label}")
        seen.add(label)
        receipt_path = reference(reference_value, f"fuzz receipt {label}")
        receipt = read_json(receipt_path, f"fuzz receipt {label}")
        kind = reference_value["kind"]
        require(kind in {"prepared", "build", "smoke"}, f"fuzz receipt {label}: kind is unknown")
        records_by_kind[kind].append(receipt)
        if kind in {"build", "smoke"}:
            _, finished = interval(receipt.get("started_utc"), receipt.get("finished_utc"), f"fuzz receipt {label}")
            receipt_times[label] = (timestamp(receipt["started_utc"], f"fuzz receipt {label}.started_utc"), finished)
        if kind == "prepared":
            prepared = timestamp(receipt.get("prepared_utc"), f"fuzz receipt {label}.prepared_utc")
            receipt_times[label] = (prepared, prepared)
            require(isinstance(receipt.get("seeds"), dict) and receipt["seeds"] == {
                name: {"bytes": item["bytes"], "sha256": item["sha256"]} for name, item in seed_manifest.items()
            }, f"fuzz receipt {label}: seed custody differs")
            for field in ("target_source", "manifest", "lock"):
                require(isinstance(receipt.get(field), dict), f"fuzz receipt {label}: {field} metadata missing")
        elif kind == "build":
            require(isinstance(receipt.get("binary"), dict), f"fuzz receipt {label}: binary metadata missing")
            binary = receipt["binary"]
            require(set(binary) == {"path", "bytes", "sha256"} and Path(binary["path"]).is_absolute(), f"fuzz receipt {label}: binary metadata differs")
            integer(binary["bytes"], f"fuzz receipt {label}.binary.bytes")
            check_digest(binary["sha256"], f"fuzz receipt {label}.binary.sha256")
            if Path(binary["path"]).is_file():
                require(not Path(binary["path"]).is_symlink(), f"fuzz receipt {label}: retained binary is symlinked")
                check_metadata(Path(binary["path"]), {"bytes": binary["bytes"], "sha256": binary["sha256"]}, f"fuzz receipt {label}.binary")
            source = receipt.get("source_snapshot")
            require(isinstance(source, dict), f"fuzz receipt {label}: source snapshot missing")
            check_source_reference(source, f"fuzz receipt {label}.source_snapshot")
        else:
            require(receipt.get("exit_code") == 0, f"fuzz receipt {label}: smoke failed")
            require(receipt.get("binary_before") == receipt.get("binary_after"), f"fuzz receipt {label}: fuzz binary changed")
            require(isinstance(receipt.get("retained"), dict), f"fuzz receipt {label}: retained crash inventory missing")
            require(isinstance(receipt.get("argv"), list) and receipt["argv"] and all(isinstance(item, str) for item in receipt["argv"]), f"fuzz receipt {label}: smoke argv is malformed")
    required_labels = fuzz.get("required_labels")
    require(isinstance(required_labels, list) and required_labels, "protocol.fuzz.required_labels is missing")
    require(set(required_labels) == seen and all(isinstance(label, str) and LABEL.fullmatch(label) for label in required_labels), "protocol.fuzz.required_labels differs")
    if "smoke" in {item["kind"] for item in receipts}:
        smoke_times = [receipt_times[item["label"]][0] for item in receipts if item["kind"] == "smoke"]
        build_finished = [receipt_times[item["label"]][1] for item in receipts if item["kind"] == "build"]
        require(build_finished and max(build_finished) <= min(smoke_times), "fuzz smoke started before its build completed")
    require({item["kind"] for item in receipts} == {"prepared", "build", "smoke"}, "protocol.fuzz must retain prepared/build/smoke receipts")
    require(all(len(records_by_kind[kind]) == 1 for kind in records_by_kind), "protocol.fuzz must retain one receipt of each kind")
    build_binary = records_by_kind["build"][0]["binary"]
    smoke = records_by_kind["smoke"][0]
    require(smoke.get("binary_before") == build_binary and smoke.get("binary_after") == build_binary, "fuzz smoke is not bound to the accepted fuzz binary")
    require(Path(smoke["argv"][0]).is_absolute() and smoke["argv"][0] == build_binary["path"], "fuzz smoke executable binding differs")
    return {"receipts": len(receipts), "seeds": len(seed_files), "labels": sorted(seen), "times": receipt_times}


def check_summary(protocol: dict[str, Any]) -> dict[str, Any]:
    path = bundle_file("summary.json", "summary")
    value = read_json(path, "summary")
    require(value.get("schema") == analyze.SUMMARY_SCHEMA, "summary schema differs")
    require(value.get("protocol_sha256") == hashlib.sha256((ROOT / "protocol.json").read_bytes()).hexdigest(), "summary protocol digest differs")
    require(value.get("capture_count") == 24 and value.get("sample_count") == 720, "summary counts differ")
    try:
        recomputed = analyze.analyze(protocol)
    except analyze.AnalysisError as error:
        fail(str(error))
    compare_summary(recomputed, value, "summary")
    return value


def compare_summary(expected: Any, actual: Any, label: str) -> None:
    """Compare the retained summary with a fresh raw recomputation.

    The analyzer emits no wall-clock field, so every field is expected to be
    stable. Numeric types are kept distinct from booleans and non-finite
    values are rejected by the JSON readers before this comparison.
    """

    if isinstance(expected, bool) or isinstance(actual, bool):
        require(type(expected) is type(actual) and expected == actual, f"{label}: value differs")
    elif isinstance(expected, (int, float)) or isinstance(actual, (int, float)):
        require(type(expected) is type(actual) and expected == actual, f"{label}: numeric value differs")
    elif isinstance(expected, dict) or isinstance(actual, dict):
        require(isinstance(expected, dict) and isinstance(actual, dict), f"{label}: object type differs")
        require(set(expected) == set(actual), f"{label}: object fields differ")
        for key in expected:
            compare_summary(expected[key], actual[key], f"{label}.{key}")
    elif isinstance(expected, list) or isinstance(actual, list):
        require(isinstance(expected, list) and isinstance(actual, list), f"{label}: list type differs")
        require(len(expected) == len(actual), f"{label}: list length differs")
        for index, (left, right) in enumerate(zip(expected, actual)):
            compare_summary(left, right, f"{label}[{index}]")
    else:
        require(expected == actual, f"{label}: value differs")


def verify() -> dict[str, Any]:
    protocol = read_json(bundle_file("protocol.json", "protocol"), "protocol")
    rows = check_protocol(protocol)
    binaries = check_builds(protocol)
    capture_times = check_captures(protocol, rows, binaries)
    validation = check_validation_receipts(protocol, binaries)
    fuzz = check_fuzz_custody(protocol)
    frozen = timestamp(protocol.get("frozen_utc"), "protocol.frozen_utc")
    build_finished = []
    for instrumentation, item in binaries["binaries"].items():
        build_path = bundle_file(item["build_path"], f"binary build {instrumentation}")
        build = read_json(build_path, f"binary build {instrumentation}")
        build_finished.append(timestamp(build["copied_utc"], f"binary build {instrumentation}.copied_utc"))
    retained_finished = [finish for _, finish in validation["times"].values()]
    retained_finished.extend(finish for _, finish in fuzz["times"].values())
    retained_finished.extend(build_finished)
    require(max(retained_finished) <= frozen, "accepted build, validation, or fuzz custody finished after protocol freeze")
    pilot_finished = [validation["times"][label][1] for label in validation_plan(protocol)["pilots"]]
    formal_started = [capture_times[label][0] for label in capture_times]
    require(max(pilot_finished) <= frozen <= min(formal_started), "pilot validation did not complete before protocol freeze and formal captures")
    check_summary(protocol)
    return {
        "schema": VERIFICATION_SCHEMA,
        "status": "pass",
        "protocol_sha256": hashlib.sha256((ROOT / "protocol.json").read_bytes()).hexdigest(),
        "capture_count": 24,
        "sample_count": 720,
        "validation": {"required": validation["required"], "pilots": validation["pilots"], "retained": validation["retained"]},
        "fuzz": {"receipts": fuzz["receipts"], "seeds": fuzz["seeds"]},
        "temporary_binaries_required": False,
    }


def verify_bundle(bundle_root: Path = ROOT) -> dict[str, Any]:
    """Verify a copied bundle after temporary binaries have been removed."""

    global ROOT
    original_root = ROOT
    original_analyze_root = analyze.ROOT
    ROOT = Path(bundle_root).resolve()
    analyze.ROOT = ROOT
    try:
        require(ROOT.is_dir() and not ROOT.is_symlink(), "evidence root is not a directory")
        return verify()
    finally:
        analyze.ROOT = original_analyze_root
        ROOT = original_root


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--output")
    options = parser.parse_args()
    try:
        value = verify_bundle(options.root)
    except VerificationError as error:
        print(f"verification failed: {error}")
        raise SystemExit(1)
    if options.output:
        path = Path(options.root).resolve() / safe_relative(options.output, "--output")
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True)
            stream.write("\n")
    print(json.dumps(value, sort_keys=True))


if __name__ == "__main__":
    main()
