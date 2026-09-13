#!/usr/bin/env python3
"""Run paired value-evaluator reference lanes against two retained ELFs.

This runner is deliberately capture-only.  It does not build, copy, or clean
an ELF or a Cargo target.  Each case is invoked as its own value-harness
process, with the harness's existing ``/usr/bin/time -v`` wrapper preserving
the process maximum RSS and its raw stdout, stderr, status, and CSV files.

The sequence is AB/BA/AB for every case, where A is ``--before-binary`` and B
is ``--after-binary``.  ``p50_ns`` and the other elapsed fields are timed batch
values: the explicit per-case repeat is inside each measured invocation.  The
runner compares result, status, read, work, checksum, and failure fields while
leaving allocation, memory, elapsed, RSS, and timestamps available for later
analysis.  Any mismatch or custody failure makes the runner exit nonzero.

The outer runner's logs live in ``runner-logs`` beside the child directories;
the child output path itself is left absent until ``value-harness/run.py``
creates it, because that harness requires a new or empty output directory.

Use a new on-disk output directory for every capture.  The runner does not
create one when imported or when only ``--help`` is requested.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import importlib.util
import json
import os
import platform
import re
import shlex
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable


PERFORMANCE = Path(__file__).resolve().parent
DEFAULT_HARNESS = PERFORMANCE / "value-harness"
HARNESS_FILES = ("Cargo.toml", "Cargo.lock", "run.py", "src/main.rs")
PROFILE_CPU = 6
WARMUPS = 3
ITERATIONS = 31
WORKLOAD = "value-evaluation"
ADAPTER = "value"
REVISION = "candidate"
GROUP = "all"
PHASE = "evaluate"

# Keep the comparison bounded while covering direct scalar references, range
# and matrix shapes, text/empty/error controls, lazy matrix paths, and every
# existing reference refusal path.  The case names are owned here, while the
# repeats and complete CSV field list are loaded from the frozen value runner
# at invocation time so this wrapper cannot silently drift from its protocol.
SCALED_SIZES = (1, 16, 256, 4096)
CASE_NAMES = (
    *[f"reference-repeat-{size}" for size in SCALED_SIZES],
    *[f"reference-distinct-{size}" for size in SCALED_SIZES],
    *[f"reference-range-{size}" for size in SCALED_SIZES],
    "reference-cell",
    "reference-text",
    "reference-text-arithmetic",
    "reference-empty",
    "reference-empty-arithmetic",
    "reference-error",
    "reference-matrix",
    "reference-background-4096",
    "reference-lazy",
    "matrix-lazy-inline-4096",
    "matrix-lazy-aggregate-4096",
    "reference-limit-cells",
    "reference-limit-work",
    "reference-limit-memory",
    "reference-cancelled",
)

REFUSAL_CASES = {
    "reference-limit-cells",
    "reference-limit-work",
    "reference-limit-memory",
    "reference-cancelled",
}

ELAPSED_FIELDS = {"mean_ns", "p50_ns", "p95_ns", "p99_ns"}
MEMORY_FIELDS = {
    "max_rss_kib",
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
    "memory_retained_used_p50",
    "memory_retained_used_max",
    "adapter_index_reserved_bytes_p50",
    "adapter_index_reserved_bytes_max",
}
TIMESTAMP_FIELDS = {"started_at", "finished_at"}
EXCLUDED_COMPARISON_FIELDS = ELAPSED_FIELDS | MEMORY_FIELDS | TIMESTAMP_FIELDS

PAIR_FIELDS = (
    "round",
    "order",
    "case",
    "repeat",
    "before_output",
    "after_output",
    "before_status",
    "after_status",
    "before_p50_ns",
    "after_p50_ns",
    "before_max_rss_kib",
    "after_max_rss_kib",
    "deterministic_equal",
    "mismatch_fields",
)


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


def hash_harness(harness: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    for relative in HARNESS_FILES:
        path = harness / relative
        if not path.is_file():
            raise RuntimeError(f"required value-harness file is missing: {path}")
        result[relative] = sha256_file(path)
    return result


def load_harness_protocol(
    harness: Path,
) -> tuple[tuple[str, ...], tuple[tuple[str, int], ...]]:
    """Load frozen runner fields and exact corpus repeats without executing it."""
    path = harness / "run.py"
    spec = importlib.util.spec_from_file_location("ods_value_harness_protocol", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot import value-harness protocol: {path}")
    module = importlib.util.module_from_spec(spec)
    old_dont_write = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    try:
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = old_dont_write
    fields = tuple(getattr(module, "FIELDS", ()))
    all_cases = dict(getattr(module, "ALL_CASES", ()))
    if not fields:
        raise RuntimeError("value-harness run.py has no FIELDS protocol")
    missing = [name for name in CASE_NAMES if name not in all_cases]
    if missing:
        raise RuntimeError(
            "value-harness run.py is missing selected cases: " + ",".join(missing)
        )
    return fields, tuple((name, all_cases[name]) for name in CASE_NAMES)


def binary_state(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise RuntimeError(f"benchmark ELF is not a regular file: {path}")
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
    }


def receipt_state(paths: Iterable[Path]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for path in paths:
        resolved = path.resolve()
        item: dict[str, Any] = {"path": str(resolved), "exists": resolved.is_file()}
        if resolved.is_file():
            item["bytes"] = resolved.stat().st_size
            item["sha256"] = sha256_file(resolved)
        else:
            item["bytes"] = None
            item["sha256"] = None
        result.append(item)
    return result


def selected_environment() -> dict[str, str]:
    names = (
        "CARGO_TARGET_DIR",
        "TMPDIR",
        "CARGO_INCREMENTAL",
        "RUSTFLAGS",
        "RUSTC",
        "RUSTC_WRAPPER",
        "PATH",
        "LC_ALL",
        "LANG",
        "PYTHONDONTWRITEBYTECODE",
    )
    return {name: os.environ[name] for name in names if name in os.environ}


def host_metadata() -> dict[str, Any]:
    # Keep metadata lightweight and avoid process tables or /proc command
    # lines, which may disclose unrelated command arguments.
    return {
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
        },
        "python": platform.python_version(),
        "cpu_count": os.cpu_count(),
        "profile_cpu": PROFILE_CPU,
        "profile_affinity": [PROFILE_CPU],
        "taskset": "/usr/bin/taskset",
        "time": "/usr/bin/time",
    }


def write_json(path: Path, value: Any) -> None:
    path.write_text(
        json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n",
        encoding="utf-8",
    )


def ensure_fresh_output(path: Path) -> Path:
    resolved = path.resolve()
    if not resolved.is_absolute():
        raise RuntimeError("--output must resolve to an absolute path")
    forbidden = (Path("/tmp"), Path("/var/tmp"))
    if any(resolved == root or root in resolved.parents for root in forbidden):
        raise RuntimeError(f"output must be an on-disk capture path, not temporary storage: {resolved}")
    if resolved.exists():
        raise RuntimeError(f"output directory must be new and absent: {resolved}")
    resolved.parent.mkdir(parents=True, exist_ok=True)
    resolved.mkdir()
    return resolved


def ensure_distinct_binaries(before: Path, after: Path) -> None:
    if before.resolve() == after.resolve():
        raise RuntimeError("before and after ELFs must have distinct paths")
    try:
        if before.samefile(after):
            raise RuntimeError("before and after ELFs must not be the same file")
    except OSError:
        pass


def parse_raw_row(
    path: Path, protocol_fields: tuple[str, ...]
) -> tuple[dict[str, str], list[str]]:
    errors: list[str] = []
    raw = path / "raw.csv"
    if not raw.is_file():
        return {}, ["raw.csv is missing"]
    try:
        with raw.open(newline="", encoding="utf-8") as stream:
            reader = csv.DictReader(stream)
            header = tuple(reader.fieldnames or ())
            expected_fields = set(protocol_fields)
            actual_fields = set(header)
            if len(header) != len(actual_fields):
                errors.append("raw.csv has duplicate header fields")
            missing_header = sorted(expected_fields - actual_fields)
            unexpected_header = sorted(actual_fields - expected_fields)
            if missing_header:
                errors.append(
                    "raw.csv header missing protocol fields: "
                    + ",".join(missing_header)
                )
            if unexpected_header:
                errors.append(
                    "raw.csv header has unexpected fields: "
                    + ",".join(unexpected_header)
                )
            rows = list(reader)
    except (OSError, UnicodeError, csv.Error) as error:
        return {}, [f"raw.csv cannot be read: {error}"]
    if len(rows) != 1:
        errors.append(f"raw.csv has {len(rows)} rows, expected one")
        return (rows[0] if rows else {}), errors
    row = rows[0]
    if None in row:
        errors.append("raw.csv has extra data columns")
    if any(value is None for value in row.values()):
        errors.append("raw.csv has missing field values")
    missing = [field for field in protocol_fields if field not in row]
    if missing:
        errors.append(f"raw.csv missing protocol fields: {','.join(missing)}")
    return row, errors


def time_rss(path: Path) -> str:
    files = sorted(path.glob("*.time"))
    if len(files) != 1:
        return ""
    match = re.search(
        r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$",
        files[0].read_text(encoding="utf-8"),
        re.MULTILINE,
    )
    return match.group(1) if match else ""


def validate_child(
    path: Path,
    case: str,
    repeat: int,
    process_status: int,
    protocol_fields: tuple[str, ...],
) -> tuple[dict[str, str], list[str]]:
    row, errors = parse_raw_row(path, protocol_fields)
    status_files = sorted(path.glob("*.status"))
    if len(status_files) != 1:
        errors.append(f"expected one status file, found {len(status_files)}")
    else:
        try:
            status_file = int(status_files[0].read_text(encoding="utf-8").strip())
            if status_file != process_status:
                errors.append(
                    f"status file={status_file}, subprocess status={process_status}"
                )
        except (OSError, ValueError) as error:
            errors.append(f"invalid status file: {error}")
    if not list(path.glob("*.stdout")):
        errors.append("child stdout is missing")
    if not list(path.glob("*.stderr")):
        errors.append("child stderr is missing")
    rss = time_rss(path)
    if not rss:
        errors.append("/usr/bin/time did not report maximum RSS")
    if process_status != 0:
        errors.append(f"child process exited {process_status}")
    if not row:
        return row, errors

    expected = {
        "revision": REVISION,
        "workload": WORKLOAD,
        "phase": PHASE,
        "case": case,
        "repeat": str(repeat),
        "warmups": str(WARMUPS),
        "iterations": str(ITERATIONS),
    }
    for field, value in expected.items():
        if row.get(field) != value:
            errors.append(f"{field}={row.get(field)!r}, expected {value!r}")
    if row.get("status") != str(process_status):
        errors.append(
            f"raw status={row.get('status')!r}, subprocess status={process_status}"
        )
    if row.get("max_rss_kib") != rss:
        errors.append(
            f"raw max_rss_kib={row.get('max_rss_kib')!r}, /usr/bin/time={rss!r}"
        )
    expected_success = row.get("expected_success") == "true"
    try:
        successes = int(row["successes_p50"])
        refusals = int(row["refusals_p50"])
    except (KeyError, ValueError) as error:
        errors.append(f"invalid success/refusal counts: {error}")
    else:
        if expected_success:
            if row.get("failure") != "none":
                errors.append(f"successful case has failure={row.get('failure')!r}")
            if successes != repeat or refusals != 0:
                errors.append(
                    f"successful counts={successes},{refusals}, expected {repeat},0"
                )
        else:
            if row.get("failure") in {None, "", "none"}:
                errors.append("refusal case has no failure label")
            if successes != 0 or refusals != repeat:
                errors.append(
                    f"refusal counts={successes},{refusals}, expected 0,{repeat}"
                )
    if case in REFUSAL_CASES and row.get("expected_success") != "false":
        errors.append("refusal case did not declare expected_success=false")
    if case not in REFUSAL_CASES and row.get("expected_success") != "true":
        errors.append("successful case did not declare expected_success=true")
    return row, errors


def deterministic_mismatches(
    before: dict[str, str], after: dict[str, str]
) -> dict[str, dict[str, str]]:
    fields = sorted(
        (set(before) | set(after)) - EXCLUDED_COMPARISON_FIELDS
    )
    mismatches: dict[str, dict[str, str]] = {}
    for field in fields:
        left = before.get(field, "")
        right = after.get(field, "")
        if left != right:
            mismatches[field] = {"before": left, "after": right}
    return mismatches


def invoke_case(
    *,
    binary: Path,
    harness: Path,
    output: Path,
    case: str,
    repeat: int,
    round_name: str,
    slot: str,
    environment: dict[str, str],
    protocol_fields: tuple[str, ...],
) -> dict[str, Any]:
    child = output / round_name / slot / case
    if child.exists():
        raise RuntimeError(f"child output already exists: {child}")
    child.parent.mkdir(parents=True, exist_ok=True)
    child.mkdir()
    spec = importlib.util.spec_from_file_location("ods_value_harness_child", harness / "run.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot import frozen value harness")
    module = importlib.util.module_from_spec(spec)
    sys.dont_write_bytecode = True
    spec.loader.exec_module(module)
    started = utc_now()
    row, sequence = module.run_one(
        binary, child, WORKLOAD, REVISION, GROUP, PHASE, case, repeat,
        WARMUPS, ITERATIONS, False,
    )
    process_status = int(row["status"])
    with (child / "raw.csv").open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=protocol_fields)
        writer.writeheader()
        writer.writerow(row)
    write_json(child / "sequence.json", sequence)
    command_text = (child / "commands.txt").read_text(encoding="utf-8")
    command = shlex.split(command_text.strip())
    with (output / "commands.txt").open("a", encoding="utf-8") as stream:
        stream.write(command_text)
    row, errors = validate_child(child, case, repeat, process_status, protocol_fields)
    errors.extend(sequence["validation_errors"])

    return {
        "round": round_name,
        "slot": slot,
        "case": case,
        "repeat": repeat,
        "binary": str(binary),
        "output": str(child),
        "command": command,
        "started_at": started,
        "finished_at": utc_now(),
        "status": process_status,
        "validation_errors": errors,
        "row": row,
    }


def pair_csv_row(
    round_name: str,
    order: str,
    case: str,
    repeat: int,
    before: dict[str, Any],
    after: dict[str, Any],
    mismatches: dict[str, dict[str, str]],
) -> dict[str, str]:
    left = before["row"]
    right = after["row"]
    return {
        "round": round_name,
        "order": order,
        "case": case,
        "repeat": str(repeat),
        "before_output": before["output"],
        "after_output": after["output"],
        "before_status": str(before["status"]),
        "after_status": str(after["status"]),
        "before_p50_ns": left.get("p50_ns", ""),
        "after_p50_ns": right.get("p50_ns", ""),
        "before_max_rss_kib": left.get("max_rss_kib", ""),
        "after_max_rss_kib": right.get("max_rss_kib", ""),
        "deterministic_equal": str(not mismatches).lower(),
        "mismatch_fields": ",".join(sorted(mismatches)),
    }


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Capture bounded AB/BA/AB value reference lanes using two retained ELFs."
    )
    parser.add_argument("--before-binary", type=Path, required=True)
    parser.add_argument("--after-binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--harness-dir",
        type=Path,
        default=DEFAULT_HARNESS,
        help="value-harness directory matching the ELFs (default: %(default)s)",
    )
    parser.add_argument(
        "--source-receipt",
        action="append",
        type=Path,
        default=[],
        help="source/build receipt path to hash and retain; may be repeated",
    )
    parser.add_argument(
        "--before-source-receipt",
        action="append",
        type=Path,
        default=[],
        help="receipt for the before ELF; may be repeated",
    )
    parser.add_argument(
        "--after-source-receipt",
        action="append",
        type=Path,
        default=[],
        help="receipt for the after ELF; may be repeated",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_arguments()
    before = args.before_binary.resolve()
    after = args.after_binary.resolve()
    harness = args.harness_dir.resolve()
    ensure_distinct_binaries(before, after)
    before_state = binary_state(before)
    after_state = binary_state(after)
    if not harness.is_dir():
        raise SystemExit(f"value-harness directory is missing: {harness}")
    harness_before = hash_harness(harness)
    protocol_fields, case_definitions = load_harness_protocol(harness)

    receipts = {
        "shared": receipt_state(args.source_receipt),
        "before": receipt_state(args.before_source_receipt),
        "after": receipt_state(args.after_source_receipt),
    }
    for label, states in receipts.items():
        missing = [item["path"] for item in states if not item["exists"]]
        if missing:
            raise SystemExit(f"{label} source receipt is missing: {', '.join(missing)}")

    output = ensure_fresh_output(args.output)

    runner_before = sha256_file(Path(__file__).resolve())
    environment = os.environ.copy()
    environment["PYTHONDONTWRITEBYTECODE"] = "1"

    manifest: dict[str, Any] = {
        "schema": 1,
        "started_at": utc_now(),
        "status": "running",
        "automatic_acceptance": False,
        "scope": "paired direct scalar-cell reference allocation diagnostic",
        "sequence": "AB/BA/AB per case; A=before ELF, B=after ELF",
        "workload": WORKLOAD,
        "adapter": ADAPTER,
        "revision": REVISION,
        "group": GROUP,
        "phase": PHASE,
        "profile_cpu": PROFILE_CPU,
        "warmups": WARMUPS,
        "iterations": ITERATIONS,
        "repeat_semantics": "explicit repeat is inside each timed batch; p50_ns is batch duration",
        "case_count": len(case_definitions),
        "cases": [{"case": case, "repeat": repeat} for case, repeat in case_definitions],
        "harness_protocol_fields": list(protocol_fields),
        "rounds": ["AB", "BA", "AB"],
        "expected_children": len(case_definitions) * 6,
        "before_binary": before_state,
        "after_binary": after_state,
        "harness_dir": str(harness),
        "output": str(output),
        "harness_files": list(HARNESS_FILES),
        "harness_sha256_before": harness_before,
        "runner_sha256_before": runner_before,
        "source_receipts_before": receipts,
        "host": host_metadata(),
        "environment": selected_environment() | {"PYTHONDONTWRITEBYTECODE": "1"},
        "invocation": sys.argv,
        "commands_file": "commands.txt",
        "children": [],
        "pairs": [],
        "failures": [],
    }
    write_json(output / "run.json", manifest)
    (output / "commands.txt").write_text("", encoding="utf-8")

    failures: list[str] = []
    children: list[dict[str, Any]] = []
    pairs: list[dict[str, str]] = []
    stop_after_custody_failure = False
    rounds = (("round-01-AB", "A", "B"), ("round-02-BA", "B", "A"), ("round-03-AB", "A", "B"))
    binary_for_slot = {"A": before, "B": after}
    try:
        for round_name, first_slot, second_slot in rounds:
            order = first_slot + second_slot
            for case, repeat in case_definitions:
                if stop_after_custody_failure:
                    break
                first = invoke_case(
                    binary=binary_for_slot[first_slot],
                    harness=harness,
                    output=output,
                    case=case,
                    repeat=repeat,
                    round_name=round_name,
                    slot=first_slot,
                    environment=environment,
                    protocol_fields=protocol_fields,
                )
                children.append(first)
                if first["validation_errors"]:
                    failures.extend(
                        f"{round_name}/{first_slot}/{case}: {error}"
                        for error in first["validation_errors"]
                    )
                first_binary_after = binary_state(binary_for_slot[first_slot])
                first_harness_after = hash_harness(harness)
                expected_first = before_state if first_slot == "A" else after_state
                if first_binary_after != expected_first:
                    failures.append(f"{round_name}/{first_slot}/{case}: ELF changed during capture")
                    stop_after_custody_failure = True
                if first_harness_after != harness_before:
                    failures.append(f"{round_name}/{first_slot}/{case}: value-harness changed during capture")
                    stop_after_custody_failure = True
                if stop_after_custody_failure:
                    break

                second = invoke_case(
                    binary=binary_for_slot[second_slot],
                    harness=harness,
                    output=output,
                    case=case,
                    repeat=repeat,
                    round_name=round_name,
                    slot=second_slot,
                    environment=environment,
                    protocol_fields=protocol_fields,
                )
                children.append(second)
                if second["validation_errors"]:
                    failures.extend(
                        f"{round_name}/{second_slot}/{case}: {error}"
                        for error in second["validation_errors"]
                    )
                second_binary_after = binary_state(binary_for_slot[second_slot])
                second_harness_after = hash_harness(harness)
                expected_second = after_state if second_slot == "B" else before_state
                if second_binary_after != expected_second:
                    failures.append(f"{round_name}/{second_slot}/{case}: ELF changed during capture")
                    stop_after_custody_failure = True
                if second_harness_after != harness_before:
                    failures.append(f"{round_name}/{second_slot}/{case}: value-harness changed during capture")
                    stop_after_custody_failure = True

                mismatches = deterministic_mismatches(first["row"], second["row"])
                if first["validation_errors"] or second["validation_errors"]:
                    mismatches["validation"] = {
                        "before": "; ".join(first["validation_errors"]),
                        "after": "; ".join(second["validation_errors"]),
                    }
                if mismatches:
                    failures.append(
                        f"{round_name}/{order}/{case}: deterministic mismatch in "
                        + ",".join(sorted(mismatches))
                    )
                pairs.append(pair_csv_row(round_name, order, case, repeat, first, second, mismatches))
            if stop_after_custody_failure:
                break
    except Exception as error:  # retain a final manifest for interrupted setup/launches
        failures.append(f"runner exception: {type(error).__name__}: {error}")

    # Retain all pair rows in a stable LF CSV; each child directory keeps the
    # harness raw.csv and the individual stdout/stderr/status/time files.
    with (output / "pairs.csv").open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=PAIR_FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(pairs)

    harness_after: dict[str, str] | None = None
    runner_after: str | None = None
    final_binary_states: dict[str, Any] = {}
    final_receipts: dict[str, Any] = {}
    try:
        harness_after = hash_harness(harness)
        runner_after = sha256_file(Path(__file__).resolve())
        final_binary_states = {"before": binary_state(before), "after": binary_state(after)}
        final_receipts = {
            "shared": receipt_state(args.source_receipt),
            "before": receipt_state(args.before_source_receipt),
            "after": receipt_state(args.after_source_receipt),
        }
    except Exception as error:
        failures.append(f"final custody check failed: {type(error).__name__}: {error}")

    if harness_after is not None and harness_after != harness_before:
        failures.append("value-harness changed during paired capture")
    if runner_after is not None and runner_after != runner_before:
        failures.append("paired runner source changed during capture")
    if final_binary_states:
        if final_binary_states["before"] != before_state:
            failures.append("before ELF changed during paired capture")
        if final_binary_states["after"] != after_state:
            failures.append("after ELF changed during paired capture")
    if final_receipts and final_receipts != receipts:
        failures.append("source receipt changed during paired capture")
    if len(children) != manifest["expected_children"]:
        failures.append(
            f"completed {len(children)} child processes, expected {manifest['expected_children']}"
        )
    if len(pairs) != len(case_definitions) * 3:
        failures.append(f"completed {len(pairs)} pairs, expected {len(case_definitions) * 3}")

    # Strip no child details: raw paths and validation errors remain fully
    # inspectable in run.json, while the manifest also gives a concise status.
    manifest.update(
        {
            "status": "complete" if not failures else "failed",
            "finished_at": utc_now(),
            "children": children,
            "pairs": pairs,
            "failures": failures,
            "child_count": len(children),
            "pair_count": len(pairs),
            "harness_sha256_after": harness_after,
            "runner_sha256_after": runner_after,
            "source_receipts_after": final_receipts,
            "final_binaries": final_binary_states,
            "harness_unchanged": harness_after == harness_before,
            "runner_unchanged": runner_after == runner_before,
            "binaries_unchanged": bool(final_binary_states)
            and final_binary_states.get("before") == before_state
            and final_binary_states.get("after") == after_state,
            "source_receipts_unchanged": final_receipts == receipts,
            "deterministic_mismatch_pairs": sum(
                item["deterministic_equal"] != "true" for item in pairs
            ),
        }
    )
    write_json(output / "run.json", manifest)
    if failures:
        print(
            f"paired capture failed with {len(failures)} issue(s); see {output / 'run.json'}",
            file=sys.stderr,
        )
        raise SystemExit(1)
    print(f"paired capture complete: {output}")


if __name__ == "__main__":
    main()
