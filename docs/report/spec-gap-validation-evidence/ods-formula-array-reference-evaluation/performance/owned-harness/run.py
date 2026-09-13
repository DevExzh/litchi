#!/usr/bin/env python3
"""Serial CPU-6 runner for the candidate ``Evaluated::to_owned`` harness.

The runner does not build anything.  It invokes an already-built executable
once for every selected case and phase under ``taskset -c 6`` and GNU
``time -v``.  The output directory must not exist, which prevents a new run
from overwriting a retained receipt.  Binary and harness SHA-256 values are
recorded before and after the serial run and a changed input makes the run
fail.  Source identity and a source manifest are explicit command-line
inputs; they are never inferred from a possibly dirty workspace HEAD.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import json
import platform
import shlex
import subprocess
import sys
from pathlib import Path

SIZES = (1, 16, 256, 4096)
HARNESS_FILES = ("Cargo.toml", "Cargo.lock", "run.py", "src/main.rs", "README.md")
PHASES = ("setup", "own", "parse-evaluate-own")

CASES = (
    ("scalar-number", 128, "scalar"),
    ("unicode-text", 128, "scalar"),
    *((f"array-{size}", {1: 128, 16: 32, 256: 8, 4096: 1}[size], "array") for size in SIZES),
    *((f"duplicate-list-{size}", {1: 128, 16: 32, 256: 8, 4096: 1}[size], "reference") for size in SIZES),
    *((f"three-d-list-{size}", {1: 128, 16: 32, 256: 8, 4096: 1}[size], "reference") for size in SIZES),
    ("unicode-limit", 1, "limits"),
    ("array-limit-memory", 1, "limits"),
    ("array-limit-work", 1, "limits"),
    ("array-cancelled", 1, "limits"),
)

CONFIG_FIELDS = (
    "revision",
    "workload",
    "group",
    "phase",
    "case",
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
    "expected_success",
    "expected_failure",
    "rows",
    "columns",
    "elements",
    "mode",
    "value_kind",
)
RESULT_FIELDS = (
    "mean_ns",
    "p50_ns",
    "p95_ns",
    "p99_ns",
    "alloc_calls_p50",
    "alloc_calls_max",
    "dealloc_calls_p50",
    "dealloc_calls_max",
    "requested_bytes_p50",
    "requested_bytes_max",
    "released_bytes_p50",
    "released_bytes_max",
    "live_before_p50",
    "live_after_p50",
    "live_after_max",
    "peak_live_delta_p50",
    "peak_live_delta_max",
    "work_used_p50",
    "work_used_max",
    "memory_retained_used_p50",
    "memory_retained_used_max",
    "owned_reserved_bytes_p50",
    "owned_reserved_bytes_max",
    "successes_p50",
    "successes_max",
    "refusals_p50",
    "refusals_max",
    "checksum_p50",
    "checksum_max",
    "failure",
)
FIELDS = (
    "group",
    "revision",
    "workload",
    "phase",
    "case",
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
    "expected_success",
    "expected_failure",
    "rows",
    "columns",
    "elements",
    "mode",
    "value_kind",
    *RESULT_FIELDS,
    "max_rss_kib",
    "started_at",
    "finished_at",
    "status",
)
NUMERIC_CONFIG_FIELDS = {
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
    "rows",
    "columns",
    "elements",
}
NUMERIC_RESULT_FIELDS = set(RESULT_FIELDS) - {"failure"}


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="milliseconds").replace(
        "+00:00", "Z"
    )


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def harness_hashes() -> dict[str, str]:
    root = Path(__file__).resolve().parent
    result: dict[str, str] = {}
    for relative in HARNESS_FILES:
        path = root / relative
        if not path.is_file():
            raise SystemExit(f"harness input is missing: {path}")
        result[relative] = sha256_file(path)
    return result


def write_hashes(path: Path, entries: dict[str, str]) -> None:
    path.write_text(
        "".join(f"{digest}  {name}\n" for name, digest in sorted(entries.items())),
        encoding="utf-8",
    )


def parse_key_values(line: str | None) -> dict[str, str]:
    if line is None:
        return {}
    values: dict[str, str] = {}
    for token in line.split()[1:]:
        if "=" not in token:
            raise ValueError(f"protocol token has no '=': {token!r}")
        key, value = token.split("=", 1)
        if not key or key in values:
            raise ValueError(f"duplicate or empty protocol key: {key!r}")
        values[key] = value
    return values


def rss_kib(path: Path) -> str:
    if not path.is_file():
        return ""
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if line.lstrip().startswith("Maximum resident set size (kbytes):"):
            return line.split(":", 1)[1].strip()
    return ""


def source_binding(identity: str, manifest: Path) -> dict[str, str]:
    manifest = manifest.resolve()
    if not manifest.is_file():
        raise SystemExit(f"source manifest is not a file: {manifest}")
    return {
        "identity": identity,
        "manifest": str(manifest),
        "manifest_sha256": sha256_file(manifest),
    }


def selected_cases(group: str) -> tuple[tuple[str, int, str], ...]:
    if group == "all":
        return CASES
    return tuple(case for case in CASES if case[2] == group)


def expected_failure(case: str, phase: str) -> str:
    if phase == "setup":
        return "none"
    return {
        "unicode-limit": "resource-memory",
        "array-limit-memory": "resource-memory",
        "array-limit-work": "resource-work",
        "array-cancelled": "cancelled",
    }.get(case, "none")


def expected_success(case: str, phase: str) -> bool:
    return phase == "setup" or case not in {
        "unicode-limit",
        "array-limit-memory",
        "array-limit-work",
        "array-cancelled",
    }


def child_validation(
    config: dict[str, str],
    result: dict[str, str],
    parse_errors: list[str],
    expected: dict[str, str],
    status: int,
    max_rss: str,
) -> list[str]:
    errors = list(parse_errors)
    missing_config = sorted(set(CONFIG_FIELDS) - config.keys())
    missing_result = sorted(set(RESULT_FIELDS) - result.keys())
    if status == 0 and missing_config:
        errors.append(f"missing config fields: {','.join(missing_config)}")
    if status == 0 and missing_result:
        errors.append(f"missing result fields: {','.join(missing_result)}")
    for key, value in expected.items():
        if key in config and config[key] != value:
            errors.append(f"config {key}={config[key]!r}, expected {value!r}")
    for key in NUMERIC_CONFIG_FIELDS:
        if key in config:
            try:
                if int(config[key]) < 0:
                    raise ValueError
            except ValueError:
                errors.append(f"config {key} is not nonnegative: {config[key]!r}")
    if status != 0:
        errors.append(f"child exited with status {status}")
        return errors
    if not max_rss:
        errors.append("/usr/bin/time did not report maximum RSS")
    else:
        try:
            if int(max_rss) < 0:
                raise ValueError
        except ValueError:
            errors.append(f"max_rss_kib is not nonnegative: {max_rss!r}")
    for key in NUMERIC_RESULT_FIELDS:
        if key in result:
            try:
                if int(result[key]) < 0:
                    raise ValueError
            except ValueError:
                errors.append(f"result {key} is not nonnegative: {result[key]!r}")
    if missing_config or missing_result:
        return errors
    if config.get("expected_success") not in {"true", "false"}:
        errors.append("expected_success is not true or false")
        return errors
    success_expected = expected["expected_success"] == "true"
    try:
        repeat = int(config["repeat"])
        successes = int(result["successes_p50"])
        refusals = int(result["refusals_p50"])
    except (KeyError, ValueError):
        errors.append("success/refusal counts are not integers")
        return errors
    failure = result.get("failure", "")
    if success_expected:
        if failure != "none" or successes != repeat or refusals != 0:
            errors.append(
                f"successful lane has failure={failure!r}, successes={successes}, refusals={refusals}"
            )
    else:
        if failure != expected["expected_failure"]:
            errors.append(
                f"refusal lane failure={failure!r}, expected {expected['expected_failure']!r}"
            )
        if successes != 0 or refusals != repeat:
            errors.append(
                f"refusal lane has successes={successes}, refusals={refusals}, expected 0,{repeat}"
            )
    return errors


def run_one(
    *,
    binary: Path,
    output: Path,
    revision: str,
    group: str,
    phase: str,
    case: str,
    repeat: int,
    warmups: int,
    iterations: int,
) -> tuple[dict[str, str], dict[str, object]]:
    stem = f"{revision}-{phase}-{case}"
    stdout_path = output / f"{stem}.stdout"
    stderr_path = output / f"{stem}.stderr"
    time_path = output / f"{stem}.time"
    command = [
        "taskset",
        "-c",
        "6",
        "/usr/bin/time",
        "-v",
        "-o",
        str(time_path),
        str(binary),
        "--revision",
        revision,
        "--workload",
        "owned-evaluation",
        "--group",
        group,
        "--phase",
        phase,
        "--case",
        case,
        "--warmups",
        str(warmups),
        "--iterations",
        str(iterations),
        "--repeat",
        str(repeat),
    ]
    started_at = utc_now()
    with (output / "commands.txt").open("a", encoding="utf-8") as stream:
        stream.write(shlex.join(command) + "\n")
    with stdout_path.open("w", encoding="utf-8") as stdout, stderr_path.open(
        "w", encoding="utf-8"
    ) as stderr:
        completed = subprocess.run(command, stdout=stdout, stderr=stderr, check=False)
    finished_at = utc_now()
    lines = stdout_path.read_text(encoding="utf-8", errors="replace").splitlines()
    config_error: str | None = None
    result_error: str | None = None
    try:
        config = parse_key_values(next((line for line in lines if line.startswith("config ")), None))
    except ValueError as error:
        config = {}
        config_error = str(error)
    try:
        result = parse_key_values(next((line for line in lines if line.startswith("result ")), None))
    except ValueError as error:
        result = {}
        result_error = str(error)
    max_rss = rss_kib(time_path)
    row = {
        **config,
        **result,
        "group": group,
        "revision": revision,
        "workload": "owned-evaluation",
        "phase": phase,
        "case": case,
        "repeat": str(repeat),
        "max_rss_kib": max_rss,
        "started_at": started_at,
        "finished_at": finished_at,
        "status": str(completed.returncode),
    }
    expected = {
        "revision": revision,
        "workload": "owned-evaluation",
        "group": group,
        "phase": phase,
        "case": case,
        "repeat": str(repeat),
        "expected_success": str(expected_success(case, phase)).lower(),
        "expected_failure": expected_failure(case, phase),
    }
    errors = child_validation(
        config,
        result,
        [error for error in (config_error, result_error) if error],
        expected,
        completed.returncode,
        max_rss,
    )
    sequence = {
        "phase": phase,
        "case": case,
        "repeat": repeat,
        "started_at": started_at,
        "finished_at": finished_at,
        "status": completed.returncode,
        "stdout": stdout_path.name,
        "stderr": stderr_path.name,
        "time": time_path.name,
        "config_present": bool(config),
        "result_present": bool(result),
        "validation_errors": errors,
    }
    return ({field: row.get(field, "") for field in FIELDS}, sequence)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--revision", choices=("candidate",), default="candidate")
    parser.add_argument(
        "--group", choices=("all", "scalar", "array", "reference", "limits"), default="all"
    )
    parser.add_argument("--phase", choices=("all", *PHASES), default="all")
    parser.add_argument("--source-identity", required=True)
    parser.add_argument("--source-manifest", type=Path, required=True)
    parser.add_argument("--compiler", default=None)
    parser.add_argument("--build-command", default=None)
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=15)
    args = parser.parse_args()
    if not 0 <= args.warmups <= 100 or not 1 <= args.iterations <= 1000:
        raise SystemExit("warmups must be in 0..100 and iterations in 1..1000")

    binary = args.binary.resolve()
    if not binary.is_file():
        raise SystemExit(f"benchmark binary is not a file: {binary}")
    output = args.output.resolve()
    if output.exists():
        raise SystemExit(f"output directory must be new and absent: {output}")
    try:
        binary.relative_to(output)
    except ValueError:
        pass
    else:
        raise SystemExit(f"output directory contains benchmark binary: {output}")
    output.mkdir(parents=True)
    cases = selected_cases(args.group)
    if not cases:
        raise SystemExit(f"group has no cases: {args.group}")
    phases = PHASES if args.phase == "all" else (args.phase,)
    source = source_binding(args.source_identity, args.source_manifest)
    manifest_bytes = args.source_manifest.read_bytes()
    (output / "source-manifest.json").write_bytes(manifest_bytes)
    if sha256_file(output / "source-manifest.json") != source["manifest_sha256"]:
        raise SystemExit("source manifest changed before capture")
    binary_before = sha256_file(binary)
    harness_before = harness_hashes()
    (output / "commands.txt").write_text("", encoding="utf-8")
    (output / "binary-sha256-before.txt").write_text(
        f"{binary_before}  {binary}\n", encoding="utf-8"
    )
    write_hashes(output / "harness-sha256-before.txt", harness_before)
    (output / "environment.json").write_text(
        json.dumps(
            {
                "captured_at": utc_now(),
                "python": sys.version,
                "platform": platform.platform(aliased=True),
                "machine": platform.machine(),
                "cpu_affinity": [6],
                "timer": "/usr/bin/time -v",
                "execution": "serial child processes",
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    metadata: dict[str, object] = {
        "candidate_only": True,
        "revision": args.revision,
        "workload": "owned-evaluation",
        "group": args.group,
        "phase_selection": args.phase,
        "phases": list(phases),
        "cases": [{"case": case, "repeat": repeat, "group": case_group} for case, repeat, case_group in cases],
        "warmups": args.warmups,
        "iterations": args.iterations,
        "binary": str(binary),
        "binary_sha256_before": binary_before,
        "harness_sha256_before": harness_before,
        "source": source,
        "compiler": args.compiler,
        "build_command": args.build_command,
        "cpu_affinity": [6],
        "capture_status": "running",
    }
    (output / "group.json").write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")
    rows: list[dict[str, str]] = []
    sequence: list[dict[str, object]] = []
    failures: list[str] = []
    for phase in phases:
        for case, repeat, case_group in cases:
            row, item = run_one(
                binary=binary,
                output=output,
                revision=args.revision,
                group=args.group,
                phase=phase,
                case=case,
                repeat=repeat,
                warmups=args.warmups,
                iterations=args.iterations,
            )
            rows.append(row)
            sequence.append(item)
            if item["status"] != 0 or item["validation_errors"]:
                failures.append(f"{phase}/{case}: {', '.join(item['validation_errors'])}")

    binary_after = sha256_file(binary)
    source_after = (
        sha256_file(args.source_manifest) if args.source_manifest.is_file() else None
    )
    if source_after != source["manifest_sha256"]:
        failures.append("source manifest changed during capture")
    harness_after = harness_hashes()
    (output / "binary-sha256-after.txt").write_text(
        f"{binary_after}  {binary}\n", encoding="utf-8"
    )
    write_hashes(output / "harness-sha256-after.txt", harness_after)
    if binary_after != binary_before:
        failures.append(f"binary changed during capture: {binary_before} -> {binary_after}")
    if harness_after != harness_before:
        failures.append("harness inputs changed during capture")
    with (output / "raw.csv").open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    (output / "sequence.json").write_text(
        json.dumps(
            {
                "binary_sha256_before": binary_before,
                "binary_sha256_after": binary_after,
                "harness_sha256_before": harness_before,
                "harness_sha256_after": harness_after,
                "rows": sequence,
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    (output / "runner-sha256.txt").write_text(
        f"{harness_after['run.py']}  run.py\n", encoding="utf-8"
    )
    metadata.update(
        {
            "binary_sha256_after": binary_after,
            "harness_sha256_after": harness_after,
            "source_manifest_sha256_after": source_after,
            "source_manifest_unchanged": source_after == source["manifest_sha256"],
            "binary_unchanged": binary_after == binary_before,
            "harness_unchanged": harness_after == harness_before,
            "capture_status": "complete" if not failures else "failed",
            "failure_count": len(failures),
        }
    )
    (output / "group.json").write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")
    if failures:
        detail = "; ".join(failures[:8])
        suffix = "" if len(failures) <= 8 else f"; ... ({len(failures)} total)"
        raise SystemExit(f"{len(failures)} lanes failed validation: {detail}{suffix}")


if __name__ == "__main__":
    main()
