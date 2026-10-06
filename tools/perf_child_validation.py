#!/usr/bin/env python3
"""Run one controlled performance child and construct its validation result.

This is the small integration boundary shared by future capture drivers.  A
child's terminal record is durably written before its report is parsed or a
validator is called.  The report's nested output identity is then copied into
the result returned to the driver, even when the older validator returned raw
vectors without that field.  A bounded descriptor may stat an oversized
report before the terminal record; it records the size and fails closed
without reading that report into memory.

The module deliberately does not know a benchmark command, corpus, route, or
statistical gate.  Callers provide those frozen values in :class:`ChildSpec`.
It also never retries a child.  A failed terminal or report validation is a
terminal observation for the caller to retain.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import re
import signal
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Mapping, Protocol, Sequence


OUTPUT_SHA256_RE = re.compile(r"\A[0-9a-f]{64}\Z")
SCHEMA = "litchi.performance.child-validation.v1"
DEFAULT_MAX_REPORT_BYTES = 1 << 20
REPORT_CONTRACT_LEGACY = "legacy"
REPORT_CONTRACT_DOCX_OWNED_WORKLOAD = "docx-owned-workload"
REPORT_CONTRACT_DOCX_OWNED_CALIBRATION = "docx-owned-calibration"
REPORT_CONTRACTS = frozenset(
    {
        REPORT_CONTRACT_LEGACY,
        REPORT_CONTRACT_DOCX_OWNED_WORKLOAD,
        REPORT_CONTRACT_DOCX_OWNED_CALIBRATION,
    }
)
DOCX_OWNED_SCHEMA = "docx_owned_ingress_attribution_v1"


class ChildValidationError(ValueError):
    """The child report cannot satisfy the caller's frozen contract."""


class ReportValidator(Protocol):
    """The existing report validator shape used by capture drivers."""

    def validate_report(
        self,
        report: dict[str, Any],
        route: str,
        *,
        warmups: int,
        samples: int,
        label: str,
        binary: Mapping[str, Any] | None,
    ) -> Mapping[str, Any]:
        ...


@dataclass(frozen=True)
class ChildPaths:
    """Immutable output locations for one child."""

    report: Path
    stdout: Path
    stderr: Path
    terminal: Path
    validation: Path


@dataclass(frozen=True)
class ChildSpec:
    """Frozen inputs needed to run one child exactly once."""

    command: tuple[str, ...]
    cwd: Path
    environment: Mapping[str, str]
    paths: ChildPaths
    route: str
    label: str
    warmups: int
    samples: int
    expected_output_sha256: str | None
    binary: Mapping[str, Any] | None
    report_contract: str = REPORT_CONTRACT_LEGACY
    deadline_seconds: float = 120.0
    observation_seconds: float = 0.25
    max_report_bytes: int = DEFAULT_MAX_REPORT_BYTES


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def _json_bytes(value: Mapping[str, Any]) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode("utf-8")


def write_json_exclusive(path: Path, value: Mapping[str, Any], *, max_bytes: int | None = None) -> None:
    """Write one JSON receipt with create-only publication semantics."""

    data = _json_bytes(value)
    if max_bytes is not None and len(data) >= max_bytes:
        raise ChildValidationError(f"JSON receipt is at least the configured limit: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def descriptor(path: Path, *, max_bytes: int | None = None) -> dict[str, Any]:
    """Describe a regular file without following a symlink at the leaf."""

    original = Path(path)
    if original.is_symlink() or not original.is_file():
        raise ChildValidationError(f"missing regular file: {original}")
    resolved = original.resolve()
    size = original.stat().st_size
    if max_bytes is not None and size >= max_bytes:
        return {
            "path": str(resolved),
            "bytes": size,
            "sha256": None,
            "error": "file is at least the configured size limit",
        }
    digest = hashlib.sha256()
    with original.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return {"path": str(resolved), "bytes": size, "sha256": digest.hexdigest()}


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ChildValidationError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def _read_report(path: Path, *, max_bytes: int) -> tuple[dict[str, Any], dict[str, Any]]:
    original = Path(path)
    if original.is_symlink() or not original.is_file():
        raise ChildValidationError(f"missing regular report: {original}")
    path = original.resolve()
    if original.stat().st_size >= max_bytes:
        raise ChildValidationError(f"report is at least the configured limit: {path}")
    with original.open("rb") as stream:
        raw = stream.read(max_bytes)
    if len(raw) >= max_bytes:
        raise ChildValidationError(f"report grew to the configured limit: {path}")
    try:
        value = json.loads(raw.decode("utf-8"), object_pairs_hook=_reject_duplicate_keys)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ChildValidationError(f"report is not strict JSON: {error}") from error
    if not isinstance(value, dict):
        raise ChildValidationError("report must be a JSON object")
    return value, {"path": str(path), "bytes": len(raw), "sha256": _sha256_bytes(raw)}


def _require_digest(value: Any, location: str) -> str:
    if not isinstance(value, str) or OUTPUT_SHA256_RE.fullmatch(value) is None:
        raise ChildValidationError(f"{location} must be a lowercase SHA-256 hex digest")
    return value


def extract_output_sha256(report: Mapping[str, Any]) -> str:
    """Extract the authoritative nested output identity from one report.

    The top-level report identity is intentionally insufficient.  The
    selected result must carry its own digest so a report with a forged or
    stale top-level convenience field cannot satisfy this boundary.
    """

    results = report.get("results")
    if not isinstance(results, list) or not results:
        raise ChildValidationError("report.results must contain at least one result")
    first = results[0]
    if not isinstance(first, dict):
        raise ChildValidationError("report.results[0] must be an object")
    return _require_digest(first.get("output_sha256"), "report.results[0].output_sha256")


def _require_exact_int(value: Any, expected: int, location: str) -> None:
    if type(value) is not int or value != expected:
        raise ChildValidationError(f"{location} must equal {expected}")


def _require_list(value: Any, location: str) -> list[Any]:
    if not isinstance(value, list):
        raise ChildValidationError(f"{location} must be a list")
    return value


def _validate_docx_owned_common(
    report: Mapping[str, Any], spec: ChildSpec, *, case: str, route: str
) -> None:
    if report.get("schema") != DOCX_OWNED_SCHEMA:
        raise ChildValidationError("DOCX owned report schema does not match report contract")
    if report.get("scenario") != "docx-owned-ingress-attribution":
        raise ChildValidationError("DOCX owned report scenario does not match report contract")
    if report.get("case") != case:
        raise ChildValidationError(f"DOCX owned report case must be {case!r}")
    if report.get("route") != route:
        raise ChildValidationError(f"DOCX owned report route must be {route!r}")
    _require_exact_int(report.get("warmup_iterations"), spec.warmups, "warmup_iterations")
    _require_exact_int(report.get("sample_count"), spec.samples, "sample_count")


def _extract_docx_owned_workload_output_sha256(
    report: Mapping[str, Any], spec: ChildSpec
) -> str:
    expected = _require_digest(spec.expected_output_sha256, "expected_output_sha256")
    route = spec.route
    _validate_docx_owned_common(report, spec, case="workload", route=route)
    report_expected = _require_digest(
        report.get("expected_output_sha256"), "report.expected_output_sha256"
    )
    if report_expected != expected:
        raise ChildValidationError("report.expected_output_sha256 does not match expected output")
    output_values = _require_list(report.get("output_sha256"), "report.output_sha256")
    if len(output_values) != spec.samples:
        raise ChildValidationError("report.output_sha256 length does not match samples")
    for index, value in enumerate(output_values):
        observed = _require_digest(value, f"report.output_sha256[{index}]")
        if observed != expected:
            raise ChildValidationError(
                f"report.output_sha256[{index}] does not match expected output"
            )
    return expected


def _validate_docx_owned_calibration(report: Mapping[str, Any], spec: ChildSpec) -> None:
    _validate_docx_owned_common(report, spec, case="calibration", route=spec.route)
    if report.get("diagnostic_enabled") is not True:
        raise ChildValidationError("calibration must have diagnostic_enabled=true")
    if report.get("expected_output_sha256") is not None:
        raise ChildValidationError("calibration must not publish an expected output digest")
    output_values = _require_list(report.get("output_sha256"), "report.output_sha256")
    if output_values:
        raise ChildValidationError("calibration must not publish per-sample output digests")
    if report.get("correctness_status") != "not_applicable":
        raise ChildValidationError("calibration correctness_status must be not_applicable")
    for key in (
        "correctness_verified",
        "inverse_source_verified",
        "correctness_gates",
        "sink",
        "owned_ingress_read_calls",
        "owned_ingress_read_bytes",
        "owned_input_bytes",
    ):
        values = _require_list(report.get(key), f"report.{key}")
        if values:
            raise ChildValidationError(f"calibration report.{key} must be empty")


def _extract_contract_output_sha256(report: Mapping[str, Any], spec: ChildSpec) -> str | None:
    if spec.report_contract == REPORT_CONTRACT_LEGACY:
        return extract_output_sha256(report)
    if spec.report_contract == REPORT_CONTRACT_DOCX_OWNED_WORKLOAD:
        return _extract_docx_owned_workload_output_sha256(report, spec)
    if spec.report_contract == REPORT_CONTRACT_DOCX_OWNED_CALIBRATION:
        _validate_docx_owned_calibration(report, spec)
        return None
    raise ChildValidationError(f"unknown report contract: {spec.report_contract}")


def _invoke_validator(
    validator: ReportValidator | Callable[..., Mapping[str, Any]],
    report: dict[str, Any],
    spec: ChildSpec,
) -> Mapping[str, Any]:
    method = getattr(validator, "validate_report", None)
    if method is None:
        method = validator
    result = method(
        report,
        spec.route,
        warmups=spec.warmups,
        samples=spec.samples,
        label=str(spec.paths.report.resolve()),
        binary=spec.binary,
    )
    if not isinstance(result, Mapping):
        raise ChildValidationError("report validator did not return an object")
    return result


def validate_and_construct_result(
    report: dict[str, Any],
    validator: ReportValidator | Callable[..., Mapping[str, Any]],
    spec: ChildSpec,
) -> dict[str, Any]:
    """Run the existing validator and return a result with a bound digest.

    The validator remains authoritative for report semantics and raw vectors.
    This wrapper adds the integration contract that the historical 0938
    runner assumed: the nested output digest is present, valid, and equal to
    the frozen expected identity.  A validator result that already includes a
    digest must agree with the report; a result that omits it receives the
    verified value.
    """

    observed = _extract_contract_output_sha256(report, spec)
    if spec.report_contract != REPORT_CONTRACT_DOCX_OWNED_CALIBRATION:
        expected = _require_digest(spec.expected_output_sha256, "expected_output_sha256")
        if observed != expected:
            if spec.report_contract == REPORT_CONTRACT_LEGACY:
                raise ChildValidationError(
                    "report.results[0].output_sha256 does not match expected_output_sha256"
                )
            raise ChildValidationError(
                "report.expected_output_sha256 does not match expected output"
            )
    validated = _invoke_validator(validator, report, spec)
    result = dict(validated)
    if observed is None:
        if "output_sha256" in result:
            raise ChildValidationError("calibration validator returned output_sha256")
    else:
        if "output_sha256" in result and result["output_sha256"] != observed:
            raise ChildValidationError("validator output_sha256 disagrees with report")
        result["output_sha256"] = observed
    elapsed = result.get("elapsed_ns")
    if elapsed is not None:
        if not isinstance(elapsed, Sequence) or isinstance(elapsed, (str, bytes)):
            raise ChildValidationError("validator elapsed_ns must be a sequence")
        result["sample_count"] = len(elapsed)
    else:
        result["sample_count"] = spec.samples
    return result


def _validate_spec(spec: ChildSpec) -> None:
    if not spec.command or any(not isinstance(item, str) or not item for item in spec.command):
        raise ChildValidationError("child command must be a non-empty string sequence")
    if not spec.cwd.is_dir():
        raise ChildValidationError(f"child cwd is not a directory: {spec.cwd}")
    if not spec.environment or any(
        not isinstance(key, str) or not key or not isinstance(value, str)
        for key, value in spec.environment.items()
    ):
        raise ChildValidationError("child environment must be an explicit string mapping")
    if spec.warmups < 0 or spec.samples <= 0:
        raise ChildValidationError("warmups must be nonnegative and samples positive")
    if spec.deadline_seconds <= 0 or spec.observation_seconds <= 0:
        raise ChildValidationError("child deadlines must be positive")
    if spec.max_report_bytes <= 0:
        raise ChildValidationError("max_report_bytes must be positive")
    if spec.report_contract not in REPORT_CONTRACTS:
        raise ChildValidationError(f"unknown report contract: {spec.report_contract}")
    if spec.report_contract == REPORT_CONTRACT_DOCX_OWNED_CALIBRATION:
        if spec.route != "route_independent":
            raise ChildValidationError(
                "docx-owned-calibration route must be route_independent"
            )
        if spec.expected_output_sha256 is not None:
            raise ChildValidationError(
                "docx-owned-calibration must omit expected_output_sha256"
            )
    elif spec.report_contract == REPORT_CONTRACT_DOCX_OWNED_WORKLOAD:
        if spec.route not in {"owned_vec", "owned_reader"}:
            raise ChildValidationError(
                "docx-owned-workload route must be owned_vec or owned_reader"
            )
        _require_digest(spec.expected_output_sha256, "expected_output_sha256")
    else:
        _require_digest(spec.expected_output_sha256, "expected_output_sha256")
    output_paths = (
        spec.paths.report,
        spec.paths.stdout,
        spec.paths.stderr,
        spec.paths.terminal,
        spec.paths.validation,
    )
    resolved_paths = [path.resolve() for path in output_paths]
    if len(set(resolved_paths)) != len(resolved_paths):
        raise ChildValidationError("child output paths must be distinct")
    for path in output_paths:
        if path.exists():
            raise ChildValidationError(f"refusing to overwrite child artifact: {path}")


def _group_exists(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def _close_process_group(pgid: int) -> tuple[bool, int, bool]:
    forced_cleanup = _group_exists(pgid)
    if forced_cleanup:
        # A normally exiting child may have left a descendant behind.  Close
        # that group so a failed trial cannot leak into later trials, while
        # returning the cleanup marker so closure cannot turn the trial green.
        try:
            os.killpg(pgid, signal.SIGTERM)
        except ProcessLookupError:
            pass
    for probe in range(1, 101):
        if not _group_exists(pgid):
            return True, probe, forced_cleanup
        if probe == 50 and forced_cleanup:
            try:
                os.killpg(pgid, signal.SIGKILL)
            except ProcessLookupError:
                pass
        time.sleep(0.01)
    return not _group_exists(pgid), 100, forced_cleanup


def run_child_and_validate(
    spec: ChildSpec,
    validator: ReportValidator | Callable[..., Mapping[str, Any]],
) -> dict[str, Any]:
    """Run one child, persist terminal state, then validate exactly once.

    The returned dictionary contains the raw validated result only on success.
    Both success and failure return after an immutable terminal and validation
    receipt have been written.  No path is replaced and no retry is attempted.
    """

    _validate_spec(spec)
    for path in (spec.paths.report, spec.paths.stdout, spec.paths.stderr,
                 spec.paths.terminal, spec.paths.validation):
        path.parent.mkdir(parents=True, exist_ok=True)

    started = _utc_now()
    monotonic_start = time.monotonic()
    process: subprocess.Popen[bytes] | None = None
    exit_code: int | None = None
    runner_error: str | None = None
    deadline_exceeded = False
    observation_timeouts = 0
    pgid: int | None = None

    with spec.paths.stdout.open("xb") as stdout, spec.paths.stderr.open("xb") as stderr:
        try:
            process = subprocess.Popen(
                list(spec.command),
                cwd=str(spec.cwd.resolve()),
                env=dict(spec.environment),
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
                close_fds=True,
            )
            pgid = process.pid
            while True:
                try:
                    exit_code = process.wait(timeout=spec.observation_seconds)
                    break
                except subprocess.TimeoutExpired:
                    observation_timeouts += 1
                    if time.monotonic() - monotonic_start <= spec.deadline_seconds:
                        continue
                    deadline_exceeded = True
                    try:
                        os.killpg(process.pid, signal.SIGTERM)
                    except ProcessLookupError:
                        pass
                    try:
                        exit_code = process.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        try:
                            os.killpg(process.pid, signal.SIGKILL)
                        except ProcessLookupError:
                            pass
                        exit_code = process.wait(timeout=5)
                    break
        except Exception as error:  # terminal receipt must describe setup failures too
            runner_error = f"{type(error).__name__}: {error}"
    ended = _utc_now()

    group_closed = True
    group_probe_count = 0
    forced_group_cleanup = False
    if pgid is not None:
        group_closed, group_probe_count, forced_group_cleanup = _close_process_group(pgid)

    report_descriptor: dict[str, Any] | None = None
    if spec.paths.report.is_file() and not spec.paths.report.is_symlink():
        report_descriptor = descriptor(spec.paths.report, max_bytes=spec.max_report_bytes)

    terminal = {
        "schema": SCHEMA,
        "status": "terminal",
        "argv": list(spec.command),
        "cwd": str(spec.cwd.resolve()),
        "environment": dict(sorted(spec.environment.items())),
        "started_utc": started,
        "ended_utc": ended,
        "exit_code": exit_code,
        "observation_timeout_count": observation_timeouts,
        "deadline_seconds": spec.deadline_seconds,
        "deadline_exceeded": deadline_exceeded,
        "runner_error": runner_error,
        "process_group_id": pgid,
        "process_group_closed": group_closed,
        "process_group_forced_cleanup": forced_group_cleanup,
        "process_group_probe_count": group_probe_count,
        "stdout": descriptor(spec.paths.stdout),
        "stderr": descriptor(spec.paths.stderr),
        "report": report_descriptor,
        "report_error": None if report_descriptor is None else report_descriptor.get("error"),
        "route": spec.route,
        "label": spec.label,
        "report_contract": spec.report_contract,
        "expected_output_sha256": spec.expected_output_sha256,
        "binary": None if spec.binary is None else dict(spec.binary),
    }
    # This is intentionally before _read_report and before validator invocation.
    write_json_exclusive(spec.paths.terminal, terminal)

    validation: dict[str, Any]
    validated_result: dict[str, Any] | None = None
    try:
        if exit_code != 0:
            raise ChildValidationError(f"child exited with code {exit_code!r}")
        if deadline_exceeded or runner_error is not None:
            raise ChildValidationError("child did not reach a successful terminal state")
        if not group_closed:
            raise ChildValidationError("child process group did not close")
        if forced_group_cleanup:
            raise ChildValidationError(
                "child process group required forced cleanup after child exit"
            )
        if report_descriptor is None:
            raise ChildValidationError("child report is missing")
        report, current_descriptor = _read_report(
            spec.paths.report, max_bytes=spec.max_report_bytes
        )
        if current_descriptor != report_descriptor:
            raise ChildValidationError("child report changed after terminal receipt")
        validated_result = validate_and_construct_result(report, validator, spec)
        validation = {
            "schema": SCHEMA,
            "status": "passed",
            "route": spec.route,
            "label": spec.label,
            "report_contract": spec.report_contract,
            "report": report_descriptor,
            "sample_count": validated_result["sample_count"],
        }
        if "output_sha256" in validated_result:
            validation["output_sha256"] = validated_result["output_sha256"]
    except Exception as error:
        validation = {
            "schema": SCHEMA,
            "status": "failed",
            "route": spec.route,
            "label": spec.label,
            "report_contract": spec.report_contract,
            "report": report_descriptor,
            "error": f"{type(error).__name__}: {error}",
        }
    write_json_exclusive(spec.paths.validation, validation)
    return {
        "status": validation["status"],
        "exit_code": exit_code,
        "terminal": terminal,
        "validation": validation,
        "result": validated_result,
    }


def _load_validator(path: Path, expected_sha256: str) -> Any:
    actual = hashlib.sha256(path.read_bytes()).hexdigest()
    if actual != _require_digest(expected_sha256, "validator_sha256"):
        raise ChildValidationError("validator module hash does not match frozen identity")
    module_spec = importlib.util.spec_from_file_location(
        "litchi_child_validation_dependency", path
    )
    if module_spec is None or module_spec.loader is None:
        raise ChildValidationError(f"cannot load validator module: {path}")
    module = importlib.util.module_from_spec(module_spec)
    module_spec.loader.exec_module(module)
    if not callable(getattr(module, "validate_report", None)):
        raise ChildValidationError("validator module has no validate_report function")
    return module


def main(argv: Sequence[str] | None = None) -> int:
    """Run one explicitly bound child from a future capture driver.

    The CLI accepts only an explicit child environment; it never forwards the
    caller's environment wholesale.  It prints a scalar outcome so raw child
    stdout remains in the separately bound stream file.
    """

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--validator", required=True, type=Path)
    parser.add_argument("--validator-sha256", required=True)
    parser.add_argument("--cwd", required=True, type=Path)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--stdout", required=True, type=Path)
    parser.add_argument("--stderr", required=True, type=Path)
    parser.add_argument("--terminal", required=True, type=Path)
    parser.add_argument("--validation", required=True, type=Path)
    parser.add_argument("--route", required=True)
    parser.add_argument("--label", required=True)
    parser.add_argument("--warmups", required=True, type=int)
    parser.add_argument("--samples", required=True, type=int)
    parser.add_argument(
        "--report-contract",
        choices=sorted(REPORT_CONTRACTS),
        default=REPORT_CONTRACT_LEGACY,
    )
    parser.add_argument("--expected-output-sha256")
    parser.add_argument("--binary-json", type=Path)
    parser.add_argument("--env", action="append", default=[], metavar="KEY=VALUE")
    parser.add_argument("--deadline-seconds", type=float, default=120.0)
    parser.add_argument("--observation-seconds", type=float, default=0.25)
    parser.add_argument("--command", nargs=argparse.REMAINDER, required=True)
    args = parser.parse_args(argv)
    try:
        environment: dict[str, str] = {}
        for item in args.env:
            key, separator, value = item.partition("=")
            if not separator or not key:
                raise ChildValidationError("--env must use KEY=VALUE")
            if key in environment:
                raise ChildValidationError(f"duplicate --env key: {key}")
            environment[key] = value
        binary: Mapping[str, Any] | None = None
        if args.binary_json is not None:
            binary_value = json.loads(args.binary_json.read_text(encoding="utf-8"))
            if not isinstance(binary_value, Mapping):
                raise ChildValidationError("--binary-json must contain an object")
            binary = dict(binary_value)
        validator = _load_validator(args.validator, args.validator_sha256)
        spec = ChildSpec(
            command=tuple(args.command),
            cwd=args.cwd,
            environment=environment,
            paths=ChildPaths(
                report=args.report,
                stdout=args.stdout,
                stderr=args.stderr,
                terminal=args.terminal,
                validation=args.validation,
            ),
            route=args.route,
            label=args.label,
            warmups=args.warmups,
            samples=args.samples,
            expected_output_sha256=args.expected_output_sha256,
            binary=binary,
            report_contract=args.report_contract,
            deadline_seconds=args.deadline_seconds,
            observation_seconds=args.observation_seconds,
        )
        outcome = run_child_and_validate(spec, validator)
    except Exception as error:
        print(json.dumps({"status": "configuration_failed", "error": f"{type(error).__name__}: {error}"}, sort_keys=True))
        return 2
    print(json.dumps({
        "status": outcome["status"],
        "exit_code": outcome["exit_code"],
        "validation": outcome["validation"],
    }, sort_keys=True))
    return 0 if outcome["status"] == "passed" else 1


if __name__ == "__main__":  # pragma: no cover - exercised by capture drivers
    raise SystemExit(main())
