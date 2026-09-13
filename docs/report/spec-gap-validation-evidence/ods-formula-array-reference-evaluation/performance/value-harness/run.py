#!/usr/bin/env python3
"""Run bounded candidate-only formula value/reference/array lanes.

The value evaluator is intentionally candidate-only until a comparable baseline
API exists.  The corpus keeps small scalar controls alongside rectangular
reference and array cases so output, budget, resolver, and allocation behavior
are visible in one run.  Every timed child is pinned to CPU 6; resolver setup
is measured separately by the ``setup`` phase.  ``--adapter worksheet`` uses
the production resolver without benchmark counters, while ``--instrumented``
adds the resolver/borrow counters and therefore has separate instrumentation
overhead.  Each output directory must be new or empty; the runner records
before/after SHA-256 hashes for the executable and all four harness files and
rejects a capture if either changes during execution.  A caller may also bind
an external production source identity or manifest with the corresponding
options.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import json
import re
import shlex
import subprocess
from pathlib import Path

SCALED_SIZES = (1, 4, 16, 256, 1024, 4096)

HARNESS_FILES = ("Cargo.toml", "Cargo.lock", "run.py", "src/main.rs")
CONFIG_FIELDS = (
    "revision",
    "workload",
    "instrumented",
    "group",
    "phase",
    "case",
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
    "expected_success",
    "rows",
    "columns",
    "elements",
    "mode",
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
    "adapter_index_reserved_bytes_p50",
    "adapter_index_reserved_bytes_max",
    "resolver_reads_p50",
    "resolver_reads_max",
    "resolver_distinct_reads_p50",
    "resolver_distinct_reads_max",
    "resolver_extent_calls_p50",
    "resolver_extent_calls_max",
    "resolver_sheet_index_calls_p50",
    "resolver_sheet_index_calls_max",
    "resolver_sheet_name_calls_p50",
    "resolver_sheet_name_calls_max",
    "resolver_sheet_count_calls_p50",
    "resolver_sheet_count_calls_max",
    "resolver_borrowed_text_bytes_p50",
    "resolver_borrowed_text_bytes_max",
    "resolver_copied_bytes_p50",
    "resolver_copied_bytes_max",
    "resolver_pointer_checks_p50",
    "resolver_pointer_checks_max",
    "resolver_pointer_matches_p50",
    "resolver_pointer_matches_max",
    "successes_p50",
    "successes_max",
    "refusals_p50",
    "refusals_max",
    "checksum_p50",
    "checksum_max",
    "failure",
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


def scaled_repeat(size: int) -> int:
    return {1: 128, 4: 64, 16: 32, 256: 8, 1024: 2, 4096: 1}[size]


SCALAR_CASES = (
    ("scalar-number", 128),
    ("scalar-not", 128),
    ("scalar-bitand", 128),
    ("scalar-roman", 128),
    ("scalar-and", 128),
    ("scalar-or", 128),
    *[(f"sequence-and-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"sequence-or-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
)

REFERENCE_CASES = (
    ("reference-cell", 128),
    ("reference-text", 128),
    ("reference-empty", 128),
    ("reference-logical", 128),
    ("reference-text-arithmetic", 128),
    ("reference-empty-arithmetic", 128),
    ("reference-matrix", 128),
    ("reference-error", 128),
    *[(f"reference-range-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"reference-background-{size}", 128) for size in (8, 64, 1024, 4096)],
    *[(f"reference-repeat-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"reference-distinct-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    ("reference-lazy", 128),
    ("reference-limit-cells", 128),
    ("reference-limit-work", 128),
    ("reference-limit-memory", 128),
    ("reference-cancelled", 128),
)

ARRAY_CASES = (
    *[(f"array-literal-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"array-arithmetic-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"array-not-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"array-bitand-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"matrix-lazy-inline-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"matrix-lazy-aggregate-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    ("array-broadcast-mismatch", 128),
    ("array-empty", 128),
    ("array-error", 128),
    ("array-iferror", 128),
    ("array-ifna", 128),
    *[(f"array-limit-{size}", 128) for size in (1, 16, 256, 1024, 4096)],
)

LIMIT_CASES = tuple(
    case
    for case in (*REFERENCE_CASES, *ARRAY_CASES)
    if "limit" in case[0] or case[0].endswith("cancelled")
)
ALL_CASES = (*SCALAR_CASES, *REFERENCE_CASES, *ARRAY_CASES)
PHASES = ("setup", "parse", "evaluate", "parse-evaluate")
WORKSHEET_PHASES = ("setup", "construct", "evaluate", "parse-evaluate")
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
    "rows",
    "columns",
    "elements",
    "mode",
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
    "adapter_index_reserved_bytes_p50",
    "adapter_index_reserved_bytes_max",
    "resolver_reads_p50",
    "resolver_reads_max",
    "resolver_distinct_reads_p50",
    "resolver_distinct_reads_max",
    "resolver_extent_calls_p50",
    "resolver_extent_calls_max",
    "resolver_sheet_index_calls_p50",
    "resolver_sheet_index_calls_max",
    "resolver_sheet_name_calls_p50",
    "resolver_sheet_name_calls_max",
    "resolver_sheet_count_calls_p50",
    "resolver_sheet_count_calls_max",
    "resolver_borrowed_text_bytes_p50",
    "resolver_borrowed_text_bytes_max",
    "resolver_copied_bytes_p50",
    "resolver_copied_bytes_max",
    "resolver_pointer_checks_p50",
    "resolver_pointer_checks_max",
    "resolver_pointer_matches_p50",
    "resolver_pointer_matches_max",
    "successes_p50",
    "successes_max",
    "refusals_p50",
    "refusals_max",
    "checksum_p50",
    "checksum_max",
    "failure",
    "max_rss_kib",
    "started_at",
    "finished_at",
    "status",
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


def harness_hashes() -> dict[str, str]:
    root = Path(__file__).resolve().parent
    hashes: dict[str, str] = {}
    for relative in HARNESS_FILES:
        path = root / relative
        if not path.is_file():
            raise SystemExit(f"harness source file is missing: {path}")
        hashes[relative] = sha256_file(path)
    return hashes


def write_hash_file(path: Path, entries: dict[str, str]) -> None:
    path.write_text(
        "".join(f"{digest}  {name}\n" for name, digest in sorted(entries.items())),
        encoding="utf-8",
    )


def parse_key_values(line: str | None) -> dict[str, str]:
    if not line:
        return {}
    values: dict[str, str] = {}
    for token in line.split()[1:]:
        if "=" not in token:
            raise ValueError(f"result token has no '=': {token!r}")
        key, value = token.split("=", 1)
        if not key or key in values:
            raise ValueError(f"duplicate or empty result key: {key!r}")
        values[key] = value
    return values


def rss_kib(path: Path) -> str:
    if not path.exists():
        return ""
    match = re.search(
        r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$",
        path.read_text(encoding="utf-8"),
        re.MULTILINE,
    )
    return match.group(1) if match else ""


def prepare_output(path: Path) -> None:
    if path.exists():
        if not path.is_dir():
            raise SystemExit(f"output path is not a directory: {path}")
        existing = sorted(path.iterdir())
        if existing:
            names = ", ".join(item.name for item in existing[:5])
            suffix = "" if len(existing) <= 5 else f", ... ({len(existing)} entries)"
            raise SystemExit(
                f"output directory must be new or empty: {path} contains {names}{suffix}"
            )
    else:
        path.mkdir(parents=True)


def write_group(path: Path, metadata: dict[str, object]) -> None:
    path.write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")


def source_binding(
    source_identity: str | None,
    source_manifest: Path | None,
) -> dict[str, str | None]:
    if source_manifest is not None:
        source_manifest = source_manifest.resolve()
        if not source_manifest.is_file():
            raise SystemExit(f"source manifest is not a file: {source_manifest}")
    return {
        "identity": source_identity or None,
        "manifest": str(source_manifest) if source_manifest is not None else None,
        "manifest_sha256": sha256_file(source_manifest) if source_manifest is not None else None,
    }


def cases_for_group(group: str) -> tuple[tuple[str, int], ...]:
    if group == "scalar":
        return SCALAR_CASES
    if group == "reference":
        return REFERENCE_CASES
    if group == "array":
        return ARRAY_CASES
    if group == "limits":
        return LIMIT_CASES
    return ALL_CASES


WORKSHEET_CASES = (
    ("adapter-cell", 128),
    ("adapter-text", 128),
    ("adapter-empty", 128),
    ("adapter-missing", 128),
    *[(f"adapter-repeated-row-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"adapter-repeated-cell-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
    *[(f"adapter-background-{size}", scaled_repeat(size)) for size in SCALED_SIZES],
)


def validate_child_row(
    row: dict[str, str],
    config: dict[str, str],
    result: dict[str, str],
    config_error: str | None,
    result_error: str | None,
    *,
    expected: dict[str, str],
    status: int,
) -> list[str]:
    """Check protocol fields without hiding a failed child lane.

    A preflight failure may happen before the binary prints either protocol
    line.  The caller still retains its replay identity in ``row`` and the
    sequence record.  A zero exit status, however, is accepted only when the
    complete result protocol and its success/refusal counts agree with the
    binary's declared expectation.
    """
    errors: list[str] = []
    if config_error:
        errors.append(f"malformed config: {config_error}")
    if result_error:
        errors.append(f"malformed result: {result_error}")
    missing_config = sorted(set(CONFIG_FIELDS) - config.keys())
    missing_result = sorted(set(RESULT_FIELDS) - result.keys())
    if status == 0 and missing_config:
        errors.append(f"missing config fields: {','.join(missing_config)}")
    if status == 0 and missing_result:
        errors.append(f"missing result fields: {','.join(missing_result)}")

    for field, expected_value in expected.items():
        actual = config.get(field)
        if actual is not None and actual != expected_value:
            errors.append(f"config {field}={actual!r}, expected {expected_value!r}")

    if config.get("expected_success") not in {"true", "false"}:
        if status == 0:
            errors.append("config expected_success is not true or false")
        return errors

    for field in NUMERIC_CONFIG_FIELDS:
        value = config.get(field)
        if value is None:
            continue
        try:
            if int(value) < 0:
                raise ValueError("negative")
        except ValueError:
            errors.append(f"config {field} is not a nonnegative integer: {value!r}")

    if status != 0 or missing_result:
        return errors

    if not row.get("max_rss_kib"):
        errors.append("/usr/bin/time did not report maximum RSS")
    else:
        try:
            if int(row["max_rss_kib"]) < 0:
                raise ValueError("negative")
        except ValueError:
            errors.append(f"max_rss_kib is not a nonnegative integer: {row['max_rss_kib']!r}")

    for field in NUMERIC_RESULT_FIELDS:
        value = result.get(field)
        if value is None:
            continue
        try:
            if int(value) < 0:
                raise ValueError("negative")
        except ValueError:
            errors.append(f"result {field} is not a nonnegative integer: {value!r}")

    expected_success = config["expected_success"] == "true"
    try:
        repeat = int(config["repeat"])
        successes = int(result["successes_p50"])
        refusals = int(result["refusals_p50"])
    except (KeyError, ValueError):
        return errors
    failure = result.get("failure", "")
    if expected_success:
        if failure != "none":
            errors.append(f"successful lane reported failure={failure!r}")
        if successes != repeat or refusals != 0:
            errors.append(
                f"successful lane counts successes={successes}, refusals={refusals}; "
                f"expected {repeat},0"
            )
    else:
        if failure in {"", "none"}:
            errors.append("expected refusal lane reported no failure label")
        if successes != 0 or refusals != repeat:
            errors.append(
                f"refusal lane counts successes={successes}, refusals={refusals}; "
                f"expected 0,{repeat}"
            )
    return errors


def run_one(
    binary: Path,
    out_dir: Path,
    workload: str,
    revision: str,
    group: str,
    phase: str,
    case: str,
    repeat: int,
    warmups: int,
    iterations: int,
    instrumented: bool,
) -> tuple[dict[str, str], dict[str, str]]:
    stem = f"{revision}-{phase}-{case}"
    stdout_path = out_dir / f"{stem}.stdout"
    stderr_path = out_dir / f"{stem}.stderr"
    time_path = out_dir / f"{stem}.time"
    status_path = out_dir / f"{stem}.status"
    command = [
        "taskset",
        "-c",
        "6",
        "/usr/bin/time",
        "-v",
        "-o",
        str(time_path),
        str(binary),
        "--workload",
        workload,
        "--revision",
        revision,
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
    if instrumented:
        command.append("--instrumented")
    started_at = utc_now()
    with (out_dir / "commands.txt").open("a", encoding="utf-8") as stream:
        stream.write(shlex.join(command) + "\n")
    with stdout_path.open("w", encoding="utf-8") as stdout, stderr_path.open(
        "w", encoding="utf-8"
    ) as stderr:
        completed = subprocess.run(command, stdout=stdout, stderr=stderr, check=False)
    finished_at = utc_now()
    status_path.write_text(f"{completed.returncode}\n", encoding="utf-8")

    lines = stdout_path.read_text(encoding="utf-8").splitlines()
    config_line = next((line for line in lines if line.startswith("config ")), None)
    result_line = next((line for line in lines if line.startswith("result ")), None)
    config_error = None
    result_error = None
    try:
        config = parse_key_values(config_line)
    except ValueError as error:
        config = {}
        config_error = str(error)
    try:
        result = parse_key_values(result_line)
    except ValueError as error:
        result = {}
        result_error = str(error)
    row = {
        **config,
        **result,
        # Keep replay identity even when the child exits before printing its
        # config/result (for example a preflight validation failure).
        "group": group,
        "revision": revision,
        "workload": workload,
        "phase": phase,
        "case": case,
        "repeat": str(repeat),
        "max_rss_kib": rss_kib(time_path),
        "started_at": started_at,
        "finished_at": finished_at,
        "status": str(completed.returncode),
    }
    expected_config = {
        "revision": revision,
        "workload": workload,
        "instrumented": str(instrumented).lower(),
        "group": group,
        "phase": phase,
        "case": case,
        "repeat": str(repeat),
    }
    validation_errors = validate_child_row(
        row,
        config,
        result,
        config_error,
        result_error,
        expected=expected_config,
        status=completed.returncode,
    )
    sequence = {
        "adapter": "worksheet" if workload == "worksheet-adapter" else "value",
        "instrumented": instrumented,
        "workload": workload,
        "revision": revision,
        "group": group,
        "phase": phase,
        "case": case,
        "repeat": str(repeat),
        "started_at": started_at,
        "finished_at": finished_at,
        "status": str(completed.returncode),
        "stdout": stdout_path.name,
        "stderr": stderr_path.name,
        "time": time_path.name,
        "config_present": bool(config),
        "result_present": bool(result),
        "config_parse_error": config_error,
        "result_parse_error": result_error,
        "validation_errors": validation_errors,
    }
    return ({field: row.get(field, "") for field in FIELDS}, sequence)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--adapter", choices=("value", "worksheet"), default="value"
    )
    parser.add_argument(
        "--instrumented",
        action="store_true",
        help="wrap worksheet adapter calls with resolver/borrow counters",
    )
    parser.add_argument("--revision", choices=("candidate",), default="candidate")
    parser.add_argument(
        "--group", choices=("all", "scalar", "reference", "array", "limits"), default="all"
    )
    parser.add_argument(
        "--phase", choices=("all", *PHASES, *WORKSHEET_PHASES), default="all"
    )
    parser.add_argument(
        "--source-identity",
        help="optional caller-supplied production source identity (for example a commit)",
    )
    parser.add_argument(
        "--source-manifest",
        type=Path,
        help="optional source manifest whose path and SHA-256 are retained in group.json",
    )
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=15)
    args = parser.parse_args()
    if args.warmups < 0 or args.iterations <= 0:
        raise SystemExit("warmups must be nonnegative and iterations must be positive")

    if args.adapter == "worksheet":
        if args.group != "all":
            raise SystemExit("worksheet adapter runs require --group all")
        cases = WORKSHEET_CASES
        available_phases = WORKSHEET_PHASES
        workload = "worksheet-adapter"
    else:
        cases = cases_for_group(args.group)
        available_phases = PHASES
        workload = "value-evaluation"
    if args.phase != "all" and args.phase not in available_phases:
        raise SystemExit(f"phase {args.phase!r} is not available for --adapter {args.adapter}")
    binary = args.binary.resolve()
    if not binary.is_file():
        raise SystemExit(f"benchmark binary is not a file: {binary}")
    output = args.output.resolve()
    prepare_output(output)
    binary_sha256_before = sha256_file(binary)
    harness_sha256_before = harness_hashes()
    source = source_binding(args.source_identity, args.source_manifest)
    phases = available_phases if args.phase == "all" else (args.phase,)
    (output / "commands.txt").write_text("", encoding="utf-8")
    (output / "binary-sha256-before.txt").write_text(
        f"{binary_sha256_before}  {binary}\n", encoding="utf-8"
    )
    write_hash_file(output / "harness-sha256-before.txt", harness_sha256_before)
    group_metadata: dict[str, object] = {
        "group": args.group,
        "adapter": args.adapter,
        "instrumented": args.instrumented,
        # The revision is a command-line label for this candidate-only
        # harness.  It is retained alongside hashes and is not inferred from
        # the executable or used as a source revision claim.
        "revision": args.revision,
        "revision_label_source": "command-line",
        "candidate_only": True,
        "workload": workload,
        "phase_selection": args.phase,
        "phases": list(phases),
        "case_count": len(cases),
        "cases": [{"case": case, "repeat": repeat} for case, repeat in cases],
        "warmups": args.warmups,
        "iterations": args.iterations,
        "binary": str(binary),
        "binary_sha256_before": binary_sha256_before,
        "harness_sha256_before": harness_sha256_before,
        "source": source,
        "cpu_affinity": [6],
        "capture_status": "running",
    }
    write_group(output / "group.json", group_metadata)
    rows: list[dict[str, str]] = []
    sequence: list[dict[str, str]] = []
    failures: list[str] = []
    for phase in phases:
        for case, repeat in cases:
            row, item = run_one(
                binary,
                output,
                workload,
                args.revision,
                args.group,
                phase,
                case,
                repeat,
                args.warmups,
                args.iterations,
                args.instrumented,
            )
            rows.append(row)
            sequence.append(item)
            validation_errors = item["validation_errors"]
            if row["status"] != "0" or validation_errors:
                details = ", ".join(validation_errors)
                failures.append(
                    f"{phase}/{case} status={row['status']} "
                    f"config={item['config_present']} result={item['result_present']}"
                    + (f" ({details})" if details else "")
                )
    binary_sha256_after = sha256_file(binary)
    harness_sha256_after = harness_hashes()
    (output / "binary-sha256-after.txt").write_text(
        f"{binary_sha256_after}  {binary}\n", encoding="utf-8"
    )
    write_hash_file(output / "harness-sha256-after.txt", harness_sha256_after)
    if binary_sha256_after != binary_sha256_before:
        failures.append(
            "benchmark binary changed during capture: "
            f"{binary_sha256_before} -> {binary_sha256_after}"
        )
    if harness_sha256_after != harness_sha256_before:
        failures.append("harness source changed during capture")
    with (output / "raw.csv").open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    (output / "sequence.json").write_text(
        json.dumps(
            {
                "group": args.group,
                "adapter": args.adapter,
                "instrumented": args.instrumented,
                "revision": args.revision,
                "workload": workload,
                "started_at": sequence[0]["started_at"] if sequence else None,
                "finished_at": sequence[-1]["finished_at"] if sequence else None,
                "binary_sha256_before": binary_sha256_before,
                "binary_sha256_after": binary_sha256_after,
                "harness_sha256_before": harness_sha256_before,
                "harness_sha256_after": harness_sha256_after,
                "rows": sequence,
            },
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    (output / "runner-sha256.txt").write_text(
        f"{harness_sha256_after['run.py']}  run.py\n",
        encoding="utf-8",
    )
    group_metadata.update(
        {
            "binary_sha256_after": binary_sha256_after,
            "harness_sha256_after": harness_sha256_after,
            "binary_unchanged": binary_sha256_after == binary_sha256_before,
            "harness_unchanged": harness_sha256_after == harness_sha256_before,
            "capture_status": "complete" if not failures else "failed",
            "failure_count": len(failures),
        }
    )
    write_group(output / "group.json", group_metadata)
    if failures:
        details = "; ".join(failures[:8])
        suffix = "" if len(failures) <= 8 else f"; ... ({len(failures)} total)"
        raise SystemExit(f"{len(failures)} child lanes failed or lacked results: {details}{suffix}")


if __name__ == "__main__":
    main()
