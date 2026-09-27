#!/usr/bin/env python3
"""Fail-closed verifier for the bounded 0773 durability syscall replay.

The probe emits a ``statx`` marker before and after each save.  This verifier
requires the complete ordered marker set for both legs, rejects duplicate,
nested, mismatched, incomplete, or unparsed windows, checks the recorded probe
exit status, and then compares normalized syscall sequences.

Usage::

    verify_trace.py BEFORE.strace AFTER.strace RUN.json WINDOWS.json
"""

from __future__ import annotations

import difflib
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Iterable


ROUTES = (
    "opc",
    "docx",
    "xlsx",
    "pptx",
    "xlsb",
    "cfb-writer",
    "cfb-sequential",
    "cfb-overlay-owned",
    "cfb-overlay-generic",
    "doc",
    "xls",
    "ppt",
)
DEFAULT_LEVELS = ("save",)
CANDIDATE_LEVELS = ("save", "full", "file-only", "no-sync")
MEMORY_CALLS = {"brk", "mmap", "munmap", "mremap", "madvise", "mprotect"}

MARKER = re.compile(r'"/litchi-0773-marker/([^/"]+)/([^/"]+)/(begin|end)"')
SYSCALL = re.compile(
    r"^(?:\d+\s+)?(?P<call>[a-z_0-9]+)\((?P<args>.*)\)\s+=\s+(?P<ret>.*)$"
)
STDOUT = re.compile(r"^(?P<route>[^ ]+) (?P<level>[^ ]+) (?P<size>[0-9]+) bytes$")

NORMALIZE = (
    (re.compile(r"\.litchi-[A-Za-z0-9]{6}\.tmp"), ".litchi-RANDOM.tmp"),
    (re.compile(r"\.litchi-cfb-\d+-\d+\.tmp"), ".litchi-cfb-PID-N.tmp"),
    (re.compile(r"0x[0-9A-Fa-f]+"), "ADDR"),
    (re.compile(r"getpid\(\) = \d+"), "getpid() = PID"),
    (re.compile(r"-(save|full|file-only|no-sync)\.out"), "-LEVEL.out"),
)


class VerificationError(RuntimeError):
    """A malformed or semantically incomplete replay artifact."""


def expected_order(levels: Iterable[str]) -> list[tuple[str, str]]:
    levels = tuple(levels)
    # The probe emits all default saves first, then each opt-in level in the
    # order declared by its LEVELS loop.
    groups = (levels,)
    if levels == CANDIDATE_LEVELS:
        groups = (DEFAULT_LEVELS, CANDIDATE_LEVELS[1:])
    return [
        (route_name, level)
        for group in groups
        for level in group
        for route in ROUTES
        for route_name in (route, f"{route}+create")
    ]


def normalize(text: str) -> str:
    for pattern, replacement in NORMALIZE:
        text = pattern.sub(replacement, text)
    return text


def parse_trace(path: Path) -> tuple[dict[tuple[str, str], list[tuple[str, str]]], list[tuple[str, str]]]:
    """Parse every marked window, rejecting every malformed in-window line."""

    windows: dict[tuple[str, str], list[tuple[str, str]]] = {}
    order: list[tuple[str, str]] = []
    current_key: tuple[str, str] | None = None
    current: list[tuple[str, str]] = []

    try:
        lines = path.read_text(encoding="utf-8", errors="strict").splitlines()
    except (OSError, UnicodeError) as error:
        raise VerificationError(f"cannot read {path}: {error}") from error

    for line_number, line in enumerate(lines, start=1):
        marker = MARKER.search(line)
        if marker:
            syscall = SYSCALL.match(line)
            if syscall is None or syscall["call"] != "statx":
                raise VerificationError(
                    f"{path}:{line_number}: marker is not a complete statx syscall"
                )
            key = (marker[1], marker[2])
            edge = marker[3]
            if edge == "begin":
                if current_key is not None:
                    raise VerificationError(
                        f"{path}:{line_number}: nested begin for {key!r} inside {current_key!r}"
                    )
                if key in windows or key in order:
                    raise VerificationError(
                        f"{path}:{line_number}: duplicate begin for {key!r}"
                    )
                current_key = key
                current = []
            else:
                if current_key is None:
                    raise VerificationError(
                        f"{path}:{line_number}: end for {key!r} without begin"
                    )
                if key != current_key:
                    raise VerificationError(
                        f"{path}:{line_number}: end for {key!r} closes {current_key!r}"
                    )
                windows[key] = current
                order.append(key)
                current_key = None
                current = []
            continue

        if current_key is None:
            continue
        syscall = SYSCALL.match(line)
        if syscall is None:
            raise VerificationError(
                f"{path}:{line_number}: unparsed syscall line inside {current_key!r}: {line!r}"
            )
        call = syscall["call"]
        current.append((call, normalize(f"{call}({syscall['args']}) = {syscall['ret']}")))

    if current_key is not None:
        raise VerificationError(f"{path}: incomplete window {current_key!r} at end of trace")
    return windows, order


def parse_stdout(path: Path, levels: Iterable[str]) -> tuple[list[tuple[str, str]], dict[tuple[str, str], int]]:
    expected = expected_order(levels)
    observed: list[tuple[str, str]] = []
    sizes: dict[tuple[str, str], int] = {}
    try:
        lines = path.read_text(encoding="utf-8", errors="strict").splitlines()
    except (OSError, UnicodeError) as error:
        raise VerificationError(f"cannot read {path}: {error}") from error
    for line_number, line in enumerate(lines, start=1):
        if not line:
            raise VerificationError(f"{path}:{line_number}: blank probe output line")
        match = STDOUT.fullmatch(line)
        if match is None:
            raise VerificationError(f"{path}:{line_number}: unparsed probe output: {line!r}")
        key = (match["route"], match["level"])
        if key in sizes:
            raise VerificationError(f"{path}:{line_number}: duplicate probe output for {key!r}")
        observed.append(key)
        sizes[key] = int(match["size"])
    if observed != expected:
        raise VerificationError(
            f"{path}: output window order/count mismatch; expected {len(expected)}, got {len(observed)}"
        )
    return observed, sizes


def output_filename(key: tuple[str, str]) -> str:
    return f"{key[0]}-{key[1]}.out"


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_published_outputs(
    run: dict[str, object],
    run_path: Path,
    leg: str,
    sizes: dict[tuple[str, str], int],
) -> dict[tuple[str, str], str]:
    entry = run["legs"][leg]
    published_value = entry.get("published")
    rows = entry.get("output_sha256")
    if not isinstance(published_value, str) or not isinstance(rows, dict):
        raise VerificationError(f"{run_path}: {leg} published output hashes are missing")
    published = Path(published_value)
    if not published.is_absolute():
        published = run_path.parent / published
    expected_names = {output_filename(key) for key in sizes}
    if set(rows) != expected_names:
        raise VerificationError(
            f"{run_path}: {leg} published output set mismatch; "
            f"expected {len(expected_names)}, got {len(rows)}"
        )
    result: dict[tuple[str, str], str] = {}
    for key, expected_size in sizes.items():
        name = output_filename(key)
        row = rows.get(name)
        if not isinstance(row, dict) or not isinstance(row.get("bytes"), int) or not isinstance(row.get("sha256"), str):
            raise VerificationError(f"{run_path}: {leg} output hash row is malformed for {name}")
        path = published / name
        try:
            actual_size = path.stat().st_size
            actual_sha = file_sha256(path)
        except OSError as error:
            raise VerificationError(f"{run_path}: {leg} published output is missing: {path}: {error}") from error
        if actual_size != expected_size or row["bytes"] != actual_size:
            raise VerificationError(
                f"{run_path}: {leg} published output size mismatch for {name}: "
                f"stdout={expected_size}, manifest={row['bytes']}, file={actual_size}"
            )
        if row["sha256"] != actual_sha:
            raise VerificationError(f"{run_path}: {leg} published output hash mismatch for {name}")
        result[key] = actual_sha
    return result


def without_memory(calls: list[tuple[str, str]]) -> list[str]:
    return [text for call, text in calls if call not in MEMORY_CALLS]


def syscall_summary(calls: list[tuple[str, str]]) -> dict[str, object]:
    names = [call for call, _ in calls]
    return {
        "syscalls": len(calls),
        "fsync": names.count("fsync"),
        "fdatasync": names.count("fdatasync"),
        "sync_file_range": names.count("sync_file_range"),
        "rename": sum(names.count(name) for name in ("rename", "renameat", "renameat2")),
        "directory_opens": sum(
            1
            for call, text in calls
            if call in ("open", "openat") and re.search(r"/tracedir>", text)
        ),
        "fsync_calls": [text for call, text in calls if call in ("fsync", "fdatasync")],
        "rename_calls": [text for call, text in calls if call.startswith("rename")],
    }


def expected_keys_or_fail(
    windows: dict[tuple[str, str], list[tuple[str, str]]],
    order: list[tuple[str, str]],
    levels: Iterable[str],
    label: str,
) -> None:
    expected = expected_order(levels)
    if order != expected:
        missing = [key for key in expected if key not in windows]
        extra = [key for key in order if key not in expected]
        raise VerificationError(
            f"{label}: expected ordered {len(expected)} windows, got {len(order)}; "
            f"missing={missing!r} extra={extra!r}"
        )
    if len(windows) != len(expected):
        raise VerificationError(f"{label}: duplicate or absent windows")


def directory_fd(calls: list[tuple[str, str]]) -> str:
    for call, text in calls:
        if call in ("open", "openat") and re.search(r"/tracedir>", text):
            match = re.search(r"=\s+(\d+)<[^>]+/tracedir>", text)
            if match:
                return match[1]
    raise VerificationError("Full window has no parent-directory open")


def without_directory_sync(calls: list[tuple[str, str]]) -> list[str]:
    filtered = without_memory(calls)
    fd = directory_fd(calls)
    result: list[str] = []
    for text in filtered:
        is_directory_open = text.startswith(("open(", "openat(")) and "/tracedir>" in text
        is_directory_fsync = text.startswith(f"fsync({fd}</") and "/tracedir>" in text
        is_directory_close = text.startswith(f"close({fd}</") and "/tracedir>" in text
        if not (is_directory_open or is_directory_fsync or is_directory_close):
            result.append(text)
    return result


def without_one_file_sync(calls: list[tuple[str, str]]) -> list[str]:
    filtered = without_memory(calls)
    sync_indices = [index for index, text in enumerate(filtered) if text.startswith(("fsync(", "fdatasync("))]
    if len(sync_indices) != 1:
        raise VerificationError(
            f"FileOnly window must contain exactly one file sync before NoSync, got {len(sync_indices)}"
        )
    return [text for index, text in enumerate(filtered) if index != sync_indices[0]]


def check_run_status(run_path: Path) -> dict[str, object]:
    try:
        run = json.loads(run_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise VerificationError(f"cannot read run manifest {run_path}: {error}") from error
    if not isinstance(run, dict) or run.get("schema") != 1 or not isinstance(run.get("legs"), dict):
        raise VerificationError(f"{run_path}: unsupported or incomplete run manifest")
    expected_windows = {"before": 24, "after": 96}
    for leg in ("before", "after"):
        entry = run["legs"].get(leg)
        if not isinstance(entry, dict):
            raise VerificationError(f"{run_path}: missing {leg} leg status")
        if entry.get("returncode") != 0:
            raise VerificationError(f"{run_path}: {leg} probe exit status is {entry.get('returncode')!r}")
        if entry.get("expected_windows") != expected_windows[leg]:
            raise VerificationError(
                f"{run_path}: {leg} expected window count is {entry.get('expected_windows')!r}, "
                f"not {expected_windows[leg]}"
            )
        if not isinstance(entry.get("stdout"), str) or not isinstance(entry.get("trace"), str):
            raise VerificationError(f"{run_path}: {leg} stdout/trace paths are missing")
    return run


def verify(before_path: Path, after_path: Path, run_path: Path, output_path: Path) -> dict[str, object]:
    report: dict[str, object] = {
        "schema": 2,
        "ok": False,
        "before_trace": str(before_path),
        "after_trace": str(after_path),
        "errors": [],
    }
    try:
        run = check_run_status(run_path)
        report["run_status"] = {
            leg: run["legs"][leg]["returncode"] for leg in ("before", "after")
        }
        for leg, trace_path in (("before", before_path), ("after", after_path)):
            declared = Path(run["legs"][leg]["trace"])
            if not declared.is_absolute():
                declared = run_path.parent / declared
            if declared.resolve() != trace_path.resolve():
                raise VerificationError(
                    f"{run_path}: {leg} trace path does not match verifier input "
                    f"({declared} != {trace_path})"
                )
        before, before_order = parse_trace(before_path)
        after, after_order = parse_trace(after_path)
        before_stdout = Path(run["legs"]["before"]["stdout"])
        after_stdout = Path(run["legs"]["after"]["stdout"])
        if not before_stdout.is_absolute():
            before_stdout = run_path.parent / before_stdout
        if not after_stdout.is_absolute():
            after_stdout = run_path.parent / after_stdout
        before_output_order, before_sizes = parse_stdout(before_stdout, DEFAULT_LEVELS)
        after_output_order, after_sizes = parse_stdout(after_stdout, CANDIDATE_LEVELS)
        before_output_hashes = validate_published_outputs(run, run_path, "before", before_sizes)
        after_output_hashes = validate_published_outputs(run, run_path, "after", after_sizes)
        expected_keys_or_fail(before, before_order, DEFAULT_LEVELS, "before")
        expected_keys_or_fail(after, after_order, CANDIDATE_LEVELS, "after")

        if before_output_order != before_order or after_output_order != after_order:
            raise VerificationError("probe stdout order does not match trace marker order")
        report["window_counts"] = {"before": len(before), "after": len(after)}

        route_reports: dict[str, object] = {}
        all_default = True
        all_default_without_memory = True
        all_save_full = True
        all_save_full_without_memory = True
        all_expectations = True
        all_subsequences = True
        all_output_bytes_identical = True

        expected_syncs = {"save": 2, "full": 2, "file-only": 1, "no-sync": 0}
        expected_directory_opens = {"save": 1, "full": 1, "file-only": 0, "no-sync": 0}
        for route in ROUTES:
            entry: dict[str, object] = {"destinations": {}}
            destination_reports: dict[str, object] = {}
            for suffix in (route, f"{route}+create"):
                base = before[(suffix, "save")]
                mine = after[(suffix, "save")]
                full = after[(suffix, "full")]
                destination: dict[str, object] = {"levels": {}}
                destination["default_identical"] = [text for _, text in base] == [text for _, text in mine]
                destination["default_identical_without_memory"] = without_memory(base) == without_memory(mine)
                destination["save_equals_full"] = [text for _, text in mine] == [text for _, text in full]
                destination["save_equals_full_without_memory"] = without_memory(mine) == without_memory(full)
                if not destination["default_identical"]:
                    destination["default_diff"] = list(
                        difflib.unified_diff(
                            [text for _, text in base],
                            [text for _, text in mine],
                            "before",
                            "after",
                            lineterm="",
                            n=1,
                        )
                    )[:80]

                levels: dict[str, object] = {}
                for level in CANDIDATE_LEVELS:
                    calls = after[(suffix, level)]
                    summary = syscall_summary(calls)
                    syncs = int(summary["fsync"]) + int(summary["fdatasync"])
                    matches = (
                        syncs == expected_syncs[level]
                        and summary["rename"] == 1
                        and summary["directory_opens"] == expected_directory_opens[level]
                    )
                    summary["matches_expectation"] = matches
                    levels[level] = summary
                    all_expectations &= matches
                destination["levels"] = levels
                destination["before_save"] = syscall_summary(base)

                file_only_expected = without_directory_sync(full)
                no_sync_expected = without_one_file_sync(after[(suffix, "file-only")])
                file_only_actual = without_memory(after[(suffix, "file-only")])
                no_sync_actual = without_memory(after[(suffix, "no-sync")])
                destination["file_only_is_full_minus_directory_sync"] = file_only_actual == file_only_expected
                destination["no_sync_is_file_only_minus_file_sync"] = no_sync_actual == no_sync_expected

                output_hashes = {
                    level: (before_output_hashes if level == "before-save" else after_output_hashes)[(suffix, "save")]
                    for level in ("before-save", "save")
                }
                output_hashes.update(
                    {level: after_output_hashes[(suffix, level)] for level in ("full", "file-only", "no-sync")}
                )
                destination["output_sha256"] = output_hashes
                destination["before_after_output_identical"] = output_hashes["before-save"] == output_hashes["save"]
                destination["all_level_outputs_identical"] = len(set(output_hashes.values())) == 1
                destination["output_bytes_identical"] = (
                    destination["before_after_output_identical"]
                    and destination["all_level_outputs_identical"]
                )
                all_output_bytes_identical &= destination["output_bytes_identical"]

                all_subsequences &= destination["file_only_is_full_minus_directory_sync"]
                all_subsequences &= destination["no_sync_is_file_only_minus_file_sync"]
                all_default &= destination["default_identical"]
                all_default_without_memory &= destination["default_identical_without_memory"]
                all_save_full &= destination["save_equals_full"]
                all_save_full_without_memory &= destination["save_equals_full_without_memory"]
                destination_reports[suffix] = destination

            entry["destinations"] = destination_reports
            entry["default_identical"] = all(
                destination["default_identical"] for destination in destination_reports.values()
            )
            entry["default_identical_without_memory"] = all(
                destination["default_identical_without_memory"] for destination in destination_reports.values()
            )
            entry["save_equals_full"] = all(
                destination["save_equals_full"] for destination in destination_reports.values()
            )
            entry["save_equals_full_without_memory"] = all(
                destination["save_equals_full_without_memory"] for destination in destination_reports.values()
            )
            entry["levels"] = {suffix: destination["levels"] for suffix, destination in destination_reports.items()}
            route_reports[route] = entry

        report.update(
            {
                "all_default_identical": all_default,
                "all_default_identical_without_memory": all_default_without_memory,
                "all_save_equals_full": all_save_full,
                "all_save_equals_full_without_memory": all_save_full_without_memory,
                "level_expectations_hold": all_expectations,
                "all_levels_are_exact_subsequences": all_subsequences,
                "all_output_bytes_identical": all_output_bytes_identical,
                "output_hashes": {
                    "before": {output_filename(key): value for key, value in before_output_hashes.items()},
                    "after": {output_filename(key): value for key, value in after_output_hashes.items()},
                },
                "routes": route_reports,
            }
        )
        checks = (
            all_default
            and all_default_without_memory
            and all_save_full
            and all_save_full_without_memory
            and all_expectations
            and all_subsequences
            and all_output_bytes_identical
        )
        if not checks:
            raise VerificationError("one or more syscall equivalence/count checks failed")
        report["ok"] = True
    except VerificationError as error:
        report.setdefault("errors", []).append(str(error))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return report


def main(argv: list[str]) -> int:
    if len(argv) != 5:
        print(f"usage: {argv[0]} BEFORE.strace AFTER.strace RUN.json WINDOWS.json", file=sys.stderr)
        return 2
    report = verify(*(Path(argument) for argument in argv[1:]))
    if not report["ok"]:
        for error in report.get("errors", []):
            print(f"verify_trace: {error}", file=sys.stderr)
        return 1
    print(json.dumps({key: value for key, value in report.items() if key not in {"routes", "errors"}}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
