#!/usr/bin/env python3
"""Run one-factor route arms as diagnostic correctness evidence.

This helper deliberately uses a separate output tree from the formal route
protocol.  It does not freeze a protocol, build a binary, or choose a source
fixture by running Rust.  The coordinator supplies an already-built binary
and invokes the helper through ``gate.py`` so the gate owns CPU affinity and
source custody.

The file-input arm is staged from the retained Rust-produced short fixture
into the gate scratch area.  The staged path is absolute because
``measure_routes._axis_argv`` resolves it relative to ``ROOT``; ``Path``
joining with an absolute operand preserves that explicit caller capability.
Failed runs retain every report, stream, resource, and staged input.  A
fully successful run removes only the staged file and the empty directories
created for it.
"""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path
import shutil
import subprocess
import zipfile
from typing import Any

import measure_routes as routes
from common import ENV, REPO, ROOT, TEMP, meta, now, read, write


WORKLOAD = "s64-a64-short-c64"
SAMPLES = 1
WARMUPS = 1
RETAINED_FIXTURE_DIR = ROOT / "consumer" / "dev56" / "short"
RETAINED_SOURCE = RETAINED_FIXTURE_DIR / "source.docx"
RETAINED_FIXTURE_HASHES = RETAINED_FIXTURE_DIR / "fixture-hashes.json"

# Keep this inventory explicit and small.  It is diagnostic coverage for the
# three one-factor axes, not a replacement for the later formal matrix.
ARM_LABELS = tuple(
    f"axis-{axis}-{value}-{WORKLOAD}"
    for axis, values in (
        ("input", ("file", "short-read", "latency")),
        ("sink", (512, routes.SINK_WRITE_BYTES, 64 * 1024)),
        ("compression", ("current", "store", "deflate")),
    )
    for value in values
)


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def _source_reference() -> dict[str, Any]:
    """Validate the retained Rust fixture and return its expected identity."""

    _require(RETAINED_SOURCE.is_file() and not RETAINED_SOURCE.is_symlink(),
             f"retained source fixture is missing or not regular: {RETAINED_SOURCE}")
    _require(RETAINED_FIXTURE_HASHES.is_file() and not RETAINED_FIXTURE_HASHES.is_symlink(),
             f"retained fixture manifest is missing: {RETAINED_FIXTURE_HASHES}")
    fixture = read(RETAINED_FIXTURE_HASHES)
    _require(fixture.get("source_count") == 64, "fixture source count is not the selected workload")
    _require(fixture.get("authored_count") == 64, "fixture authored count is not the selected workload")
    _require(fixture.get("chunk_mode") == "fixed64", "fixture chunk mode is not the selected workload")
    _require(fixture.get("text_mode") == "short", "fixture text mode is not the selected workload")
    source = fixture.get("source")
    _require(isinstance(source, dict), "fixture manifest source identity is missing")
    expected_archive = {
        "bytes": source.get("bytes"),
        "sha256": source.get("sha256"),
    }
    _require(type(expected_archive["bytes"]) is int and expected_archive["bytes"] > 0,
             "fixture manifest source byte count is invalid")
    _require(isinstance(expected_archive["sha256"], str)
             and len(expected_archive["sha256"]) == 64,
             "fixture manifest source SHA-256 is invalid")
    actual_archive = meta(RETAINED_SOURCE)
    _require(actual_archive == expected_archive,
             "retained source differs from the Rust fixture archive identity")

    # The independent corpus oracle supplies the current source XML contract;
    # comparing it here makes the fixture binding explicit without executing
    # a Rust generator during this diagnostic helper.
    oracle = routes.corpus_oracle
    _require(oracle is not None, "independent corpus oracle is unavailable")
    expected = oracle.expected_case(64, 64, "fixed64", "short")["source"]
    xml_expected = {
        "bytes": expected["main_xml_bytes"],
        "sha256": fixture.get("source_main_xml_sha256"),
    }
    _require(isinstance(xml_expected["sha256"], str) and len(xml_expected["sha256"]) == 64,
             "fixture manifest source XML SHA-256 is invalid")
    with zipfile.ZipFile(RETAINED_SOURCE) as archive:
        try:
            document = archive.read("word/document.xml")
        except KeyError as error:
            raise RuntimeError("retained source has no word/document.xml") from error
    actual_xml = {"bytes": len(document), "sha256": hashlib.sha256(document).hexdigest()}
    _require(actual_xml == xml_expected,
             "retained source XML differs from the Rust fixture identity")
    _require(
        actual_xml == {
            "bytes": expected["main_xml_bytes"],
            "sha256": expected["main_xml_sha256"],
        },
        "retained source is not the current independent Rust corpus contract",
    )
    return {
        "retained_path": str(RETAINED_SOURCE.relative_to(ROOT)),
        "fixture_manifest": str(RETAINED_FIXTURE_HASHES.relative_to(ROOT)),
        "archive": actual_archive,
        "main_xml": actual_xml,
        "rust_fixture": expected_archive,
        "current_oracle": {
            "main_xml_bytes": expected["main_xml_bytes"],
            "main_xml_sha256": expected["main_xml_sha256"],
        },
    }


def _stage_source(scratch: Path, reference: dict[str, Any]) -> tuple[Path, dict[str, Any]]:
    """Copy the verified fixture to an exclusively owned scratch path."""

    input_directory = scratch / "input"
    input_directory.mkdir()
    staged = input_directory / "source.docx"
    with RETAINED_SOURCE.open("rb") as source, staged.open("xb") as destination:
        shutil.copyfileobj(source, destination)
    staged_identity = meta(staged)
    _require(staged_identity == reference["archive"],
             "staged diagnostic source differs from retained fixture")
    binding = {
        **staged_identity,
        "retained_path": reference["retained_path"],
        "staged_path": str(staged.resolve()),
        "main_xml": reference["main_xml"],
        "identity_validation": "retained_rust_fixture_and_current_oracle_match",
    }
    return staged, binding


def _axis_arm(label: str, staged_source: Path) -> dict[str, Any]:
    try:
        arm = dict(routes.AXIS_ARM_BY_LABEL[label])
    except KeyError as error:
        raise RuntimeError(f"axis arm is absent from measure_routes: {label}") from error
    if arm["input_mode"] == "file":
        # measure_routes intentionally accepts both repository-relative and
        # absolute paths: ROOT / absolute_path remains the same absolute path.
        absolute = str(staged_source.resolve())
        arm["input_file"] = absolute
        arm["input_profile"] = dict(arm["input_profile"], file=absolute)
    return arm


def _cleanup_owned_source(staged: Path, expected: dict[str, Any]) -> tuple[bool, list[str]]:
    """Remove exactly the staged file and its empty private directories."""

    failures: list[str] = []
    if staged.is_symlink() or not staged.is_file():
        failures.append(f"owned source is not a regular file: {staged}")
        return False, failures
    try:
        if meta(staged) != {
            "bytes": expected["bytes"],
            "sha256": expected["sha256"],
        }:
            failures.append("owned source changed before successful cleanup")
            return False, failures
        staged.unlink()
    except OSError as error:
        failures.append(f"owned source cleanup failed: {type(error).__name__}: {error}")
        return False, failures

    # Only these two directories were created by _stage_source.  rmdir is
    # deliberately used so an unexpected artifact is retained rather than
    # recursively removed.
    for directory in (staged.parent, staged.parent.parent):
        try:
            directory.rmdir()
        except FileNotFoundError:
            continue
        except OSError as error:
            failures.append(f"owned scratch directory retained: {directory}: {error}")
    return not failures, failures


def _artifact_metadata(directory: Path) -> dict[str, dict[str, int | str]]:
    return {
        path.name: meta(path)
        for path in (
            directory / "report.json",
            directory / "resource.txt",
            directory / "stdout.txt",
            directory / "stderr.txt",
        )
        if path.is_file()
    }


def run(binary_path: Path, role: str, attempt: str) -> int:
    """Run all nine diagnostic arms for one explicit binary role."""

    attempt = routes.base._attempt(attempt)
    _require(role in routes.ROLES, f"unknown binary role: {role}")
    binary_path = binary_path.resolve(strict=True)
    binary = {"path": str(binary_path), **meta(binary_path)}
    destination = ROOT / "route-axis-diagnostics" / attempt
    destination.mkdir(parents=True, exist_ok=False)
    scratch = TEMP / "route-axis-diagnostics" / attempt
    scratch.mkdir(parents=True, exist_ok=False)

    bindings: dict[str, Any] = {
        "schema": "docx-replayable-tail-append-axis-diagnostic-v1",
        "version": 1,
        "binary": binary,
        "role": role,
        "attempt": attempt,
        "diagnostic_only": True,
        "protocol_frozen": False,
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "workload": WORKLOAD,
        "arms": list(ARM_LABELS),
        "driver": meta(Path(__file__)),
        "validators": routes._script_hashes(),
        "machine": routes._machine_binding(required=True),
    }
    failures: list[dict[str, str]] = []
    receipts: list[dict[str, int | str]] = []
    staged: Path | None = None
    source_binding: dict[str, Any] | None = None
    source_reference: dict[str, Any] | None = None
    try:
        source_reference = _source_reference()
        bindings["source_reference"] = source_reference
        staged, source_binding = _stage_source(scratch, source_reference)
        bindings["source"] = source_binding
    except Exception as error:
        failures.append({"label": "source", "error": f"{type(error).__name__}: {error}"})
    write(destination / "started.json", dict(bindings, started_utc=now()))

    if staged is not None and source_binding is not None:
        for label in ARM_LABELS:
            arm_directory = destination / label
            arm_directory.mkdir()
            report = arm_directory / "report.json"
            resource = arm_directory / "resource.txt"
            stdout = arm_directory / "stdout.txt"
            stderr = arm_directory / "stderr.txt"
            arm: dict[str, Any] | None = None
            argv: list[str] | None = None
            input_metadata: dict[str, Any] | None = None
            error: str | None = None
            exit_code: int | None = None
            started: dict[str, Any] = {
                "schema": "docx-replayable-tail-append-axis-diagnostic-run-v1",
                "version": 1,
                "status": "running",
                "attempt": attempt,
                "role": role,
                "label": label,
                "workload": WORKLOAD,
                "samples": SAMPLES,
                "warmups": WARMUPS,
                "source": dict(source_binding),
                "started_utc": now(),
            }
            try:
                arm = _axis_arm(label, staged)
                case = routes._axis_case(arm)
                if arm["input_mode"] == "file":
                    input_metadata = routes._axis_input_metadata(arm)
                    _require(input_metadata is not None, f"{label}: file input metadata missing")
                argv = routes._axis_argv(
                    binary,
                    case,
                    arm,
                    samples=SAMPLES,
                    warmups=WARMUPS,
                    report=report,
                    resource=resource,
                )
                started.update({
                    "argv": argv,
                    "axis": dict(arm),
                    "input_file": input_metadata,
                })
                write(arm_directory / "started.json", started)
                with stdout.open("xb") as out, stderr.open("xb") as err:
                    process = subprocess.run(
                        argv,
                        cwd=REPO,
                        env=ENV,
                        stdout=out,
                        stderr=err,
                        check=False,
                    )
                exit_code = process.returncode
                if exit_code:
                    raise RuntimeError(f"child exited {exit_code}")
                routes._check_axis_report(
                    report,
                    role,
                    arm,
                    samples=SAMPLES,
                    warmups=WARMUPS,
                    binary=binary,
                    argv=argv,
                    input_metadata=input_metadata,
                )
            except Exception as caught:
                error = f"{type(caught).__name__}: {caught}"
                failures.append({"label": label, "error": error})
            finally:
                if not (arm_directory / "started.json").exists():
                    write(arm_directory / "started.json", started)
                receipt = dict(
                    started,
                    status="pass" if error is None else "failed",
                    exit_code=exit_code,
                    finished_utc=now(),
                    passed=error is None,
                    artifacts=_artifact_metadata(arm_directory),
                )
                if error is not None:
                    receipt["error"] = error
                write(arm_directory / "receipt.json", receipt)
                receipts.append({"label": label, **meta(arm_directory / "receipt.json")})
                print(f"{label}: {'PASS' if error is None else error}", flush=True)

    unchanged = meta(binary_path) == {key: binary[key] for key in ("bytes", "sha256")}
    if not unchanged:
        failures.append({"label": "binary", "error": "executable changed during diagnostics"})

    cleanup_failures: list[str] = []
    scratch_removed = False
    if not failures and staged is not None and source_binding is not None:
        cleaned, cleanup_failures = _cleanup_owned_source(staged, source_binding)
        scratch_removed = cleaned and not scratch.exists()
        if cleanup_failures:
            failures.extend({"label": "scratch", "error": message} for message in cleanup_failures)
        if not scratch_removed and not cleanup_failures:
            failures.append({"label": "scratch", "error": "diagnostic scratch directory remains"})
    elif scratch.exists():
        scratch_removed = False

    result = dict(
        bindings,
        finished_utc=now(),
        passed=not failures,
        failures=failures,
        receipts=receipts,
        binary_unchanged=unchanged,
        scratch_removed=scratch_removed,
        failure_artifacts_retained=bool(failures),
    )
    if cleanup_failures:
        result["cleanup_failures"] = cleanup_failures
    write(destination / "result.json", result)
    return int(bool(failures))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--role", choices=routes.ROLES, required=True)
    parser.add_argument("--attempt", required=True)
    args = parser.parse_args()
    raise SystemExit(run(args.binary, args.role, args.attempt))


if __name__ == "__main__":
    main()
