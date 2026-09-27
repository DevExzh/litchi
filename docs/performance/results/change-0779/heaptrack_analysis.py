#!/usr/bin/env python3
"""Offline Heaptrack attribution for the 0779 XLSX ingress experiment.

This module only reads already-published Heaptrack artifacts.  It does not run
the profiled program or heaptrack_print.  The compressed interpreted trace is
decoded with ``zstd --decompress --stdout`` and its allocation events are
checked against the retained histogram before attribution is published.

The open probe deliberately reopens the source once after the timed region to
verify the workbook.  Both opens are therefore part of the whole-process
Heaptrack scope.  The result records the timed open and post-clock verification
open separately as well as the whole-process total.  They are two open calls
inside one ``run_one`` sample, not two samples or two ``run_one`` invocations.
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
from typing import Any, Iterable


CHUNK_BYTES = 8 * 1024
MAX_TRACE_BYTES = 512 * 1024 * 1024
MAX_TRACE_RECORDS = 20_000_000
MAX_HISTOGRAM_ROWS = 10_000_000
TARGET = "litchi_opc::phys_pkg::read_limited"
RUN_ONE = "xlsx_allocation_probe::run_one::"


class AnalysisError(RuntimeError):
    """Raised when a retained artifact is malformed or fails a binding."""


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as exc:
        raise AnalysisError(f"cannot read {path}: {exc}") from exc
    return digest.hexdigest()


def artifact(path: Path, root: Path) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        raise AnalysisError(f"expected a regular artifact: {path}")
    try:
        relative = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError as exc:
        raise AnalysisError(f"artifact escapes analysis root: {path}") from exc
    return {
        "path": relative,
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
    }


def mapped_receipt_path(value: Any, owned_worktree: Path, current_repo: Path) -> Path:
    """Map only the captured owned-worktree prefix to the current checkout.

    Absolute strings in the retained capture/report JSON remain untouched in
    the result.  This mapping is used only for checking whether a referenced
    file is present after the packet is moved from its isolated worktree to
    the main checkout.
    """
    if not isinstance(value, str) or not value:
        raise AnalysisError("receipt path is missing")
    raw = Path(value)
    if not raw.is_absolute():
        return raw
    try:
        relative = raw.resolve(strict=False).relative_to(owned_worktree.resolve(strict=False))
    except ValueError:
        return raw
    return current_repo / relative


def captured_path(
    value: Any,
    expected: Path,
    label: str,
    *,
    owned_worktree: Path,
    current_repo: Path,
) -> Path:
    if not isinstance(value, str) or not value:
        raise AnalysisError(f"{label} path is missing")
    observed = mapped_receipt_path(value, owned_worktree, current_repo)
    try:
        if observed.resolve() != expected.resolve():
            raise AnalysisError(
                f"{label} path binding differs: {observed} != {expected}"
            )
    except OSError as exc:
        raise AnalysisError(f"cannot resolve {label} path {observed}: {exc}") from exc
    return observed


def verify_captured_artifact(
    entry: Any,
    expected_path: Path,
    root: Path,
    label: str,
    *,
    owned_worktree: Path,
    current_repo: Path,
) -> dict[str, Any]:
    if not isinstance(entry, dict):
        raise AnalysisError(f"{label} binding is not an object")
    path = captured_path(
        entry.get("path"),
        expected_path,
        label,
        owned_worktree=owned_worktree,
        current_repo=current_repo,
    )
    observed = artifact(expected_path, root)
    if entry.get("bytes") != observed["bytes"]:
        raise AnalysisError(f"{label} byte binding differs")
    if entry.get("sha256") != observed["sha256"]:
        raise AnalysisError(f"{label} SHA-256 binding differs")
    return observed


def cleanup_has_receipt(cleanup: Any, receipt: dict[str, Any]) -> bool:
    """Find an exact path/size/SHA witness recursively in cleanup.json."""
    expected = (
        receipt.get("path"),
        receipt.get("bytes"),
        receipt.get("sha256"),
    )
    if not isinstance(expected[0], str) or not isinstance(expected[1], int):
        return False
    if not isinstance(expected[2], str) or len(expected[2]) != 64:
        return False
    if isinstance(cleanup, dict):
        candidate = (
            cleanup.get("path"),
            cleanup.get("bytes", cleanup.get("size")),
            cleanup.get("sha256", cleanup.get("digest")),
        )
        if candidate == expected:
            return True
        return any(cleanup_has_receipt(value, receipt) for value in cleanup.values())
    if isinstance(cleanup, list):
        return any(cleanup_has_receipt(value, receipt) for value in cleanup)
    return False


def verify_binary_receipt(
    entry: Any,
    command: list[str],
    *,
    cleanup: Any,
    cleanup_verified: bool,
    label: str,
) -> dict[str, Any]:
    if not isinstance(entry, dict):
        raise AnalysisError("capture binary receipt is not an object")
    path = entry.get("path")
    if not isinstance(path, str) or not path:
        raise AnalysisError("capture binary path is missing")
    if len(command) < 7 or command[6] != path:
        raise AnalysisError("capture command binary does not match binary receipt")
    if not isinstance(entry.get("bytes"), int) or entry["bytes"] <= 0:
        raise AnalysisError("capture binary byte receipt is invalid")
    digest = entry.get("sha256")
    if not isinstance(digest, str) or len(digest) != 64:
        raise AnalysisError("capture binary SHA-256 receipt is invalid")
    binary = Path(path)
    # The target may be removed after the final packet seal.  When present,
    # verify it now; the retained receipt remains the binding after cleanup.
    if binary.exists():
        if not binary.is_file() or binary.is_symlink():
            raise AnalysisError(f"capture binary is not a regular file: {binary}")
        if binary.stat().st_size != entry["bytes"] or sha256_file(binary) != digest:
            raise AnalysisError(f"capture binary receipt differs: {binary}")
    else:
        if not cleanup_verified:
            raise AnalysisError(f"{label} is missing without verified cleanup.json")
        if not cleanup_has_receipt(cleanup, entry):
            raise AnalysisError(f"{label} lacks an exact cleanup.json witness")
    return {
        "path": path,
        "bytes": entry["bytes"],
        "sha256": digest,
    }


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as exc:
        raise AnalysisError(f"cannot parse JSON {path}: {exc}") from exc


def require_json_object(path: Path) -> dict[str, Any]:
    value = read_json(path)
    if not isinstance(value, dict):
        raise AnalysisError(f"JSON root is not an object: {path}")
    return value


def parse_histogram(path: Path) -> dict[str, Any]:
    rows = 0
    count = 0
    requested_bytes = 0
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as exc:
        raise AnalysisError(f"cannot read histogram {path}: {exc}") from exc
    for line in lines:
        if not line:
            continue
        rows += 1
        if rows > MAX_HISTOGRAM_ROWS:
            raise AnalysisError(f"histogram has more than {MAX_HISTOGRAM_ROWS} rows")
        fields = line.split("\t")
        if len(fields) != 2 or any(not field.isdecimal() for field in fields):
            raise AnalysisError(f"malformed histogram row {rows}: {line!r}")
        size, occurrences = (int(field) for field in fields)
        count += occurrences
        requested_bytes += size * occurrences
    if rows == 0:
        raise AnalysisError(f"empty histogram: {path}")
    return {
        "rows": rows,
        "allocation_events": count,
        "requested_bytes": requested_bytes,
    }


def trace_lines(path: Path) -> Iterable[bytes]:
    if path.suffix != ".zst":
        raise AnalysisError(f"expected a .zst interpreted trace: {path}")
    try:
        process = subprocess.Popen(
            ["zstd", "--decompress", "--stdout", "--", str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except OSError as exc:
        raise AnalysisError(f"cannot start zstd for {path}: {exc}") from exc
    assert process.stdout is not None
    assert process.stderr is not None
    total = 0
    try:
        for raw in process.stdout:
            total += len(raw)
            if total > MAX_TRACE_BYTES:
                process.kill()
                raise AnalysisError(
                    f"decompressed trace exceeds {MAX_TRACE_BYTES} bytes"
                )
            yield raw
    finally:
        process.stdout.close()
    error = process.stderr.read().decode("utf-8", "replace")
    process.stderr.close()
    if process.wait() != 0:
        raise AnalysisError(f"zstd failed for {path}: {error.strip()}")


def parse_hex(token: str, label: str, record: int) -> int:
    try:
        return int(token, 16)
    except ValueError as exc:
        raise AnalysisError(
            f"invalid hexadecimal {label} at trace record {record}: {token!r}"
        ) from exc


def parse_trace(path: Path) -> dict[str, Any]:
    """Parse an interpreted Heaptrack v3 trace and attribute target events."""

    strings: dict[int, str] = {}
    instructions: dict[int, tuple[int, ...]] = {}
    traces: dict[int, tuple[int, int]] = {}
    allocation_info: list[tuple[int, int]] = []
    plus: Counter[int] = Counter()
    minus: Counter[int] = Counter()
    rows = 0

    for raw in trace_lines(path):
        rows += 1
        if rows > MAX_TRACE_RECORDS:
            raise AnalysisError(f"trace has more than {MAX_TRACE_RECORDS} records")
        try:
            line = raw.rstrip(b"\n").decode("utf-8")
        except UnicodeDecodeError as exc:
            raise AnalysisError(f"trace is not UTF-8 at record {rows}") from exc
        if not line:
            continue
        fields = line.split()
        mode = fields[0]
        if mode == "s":
            if len(fields) < 3:
                raise AnalysisError(f"malformed string record at {rows}")
            declared = parse_hex(fields[1], "string length", rows)
            text = line.split(None, 2)[2]
            if len(text.encode("utf-8")) != declared:
                raise AnalysisError(f"string length mismatch at record {rows}")
            strings[len(strings) + 1] = text
        elif mode == "i":
            if len(fields) < 3:
                raise AnalysisError(f"malformed instruction record at {rows}")
            tokens = fields[1:]
            function_ids: list[int] = []
            if len(tokens) >= 3:
                function_id = parse_hex(tokens[2], "function index", rows)
                if function_id:
                    function_ids.append(function_id)
                remaining = tokens[3:]
                if remaining:
                    if len(remaining) < 2 or (len(remaining) - 2) % 3:
                        raise AnalysisError(f"malformed inline frames at {rows}")
                    for index in range(2, len(remaining), 3):
                        function_id = parse_hex(
                            remaining[index], "inline function index", rows
                        )
                        if function_id:
                            function_ids.append(function_id)
            instructions[len(instructions) + 1] = tuple(function_ids)
        elif mode == "t":
            if len(fields) != 3:
                raise AnalysisError(f"malformed trace record at {rows}")
            trace_id = len(traces) + 1
            instruction = parse_hex(fields[1], "instruction index", rows)
            parent = parse_hex(fields[2], "trace parent", rows)
            if instruction and instruction not in instructions:
                raise AnalysisError(f"unknown instruction at trace record {rows}")
            if parent and parent >= trace_id:
                raise AnalysisError(f"forward trace parent at record {rows}")
            traces[trace_id] = (instruction, parent)
        elif mode == "a":
            if len(fields) != 3:
                raise AnalysisError(f"malformed allocation-info record at {rows}")
            size = parse_hex(fields[1], "allocation size", rows)
            trace_id = parse_hex(fields[2], "allocation trace", rows)
            if trace_id and trace_id not in traces:
                raise AnalysisError(f"unknown allocation trace at record {rows}")
            allocation_info.append((size, trace_id))
        elif mode in {"+", "-"}:
            if len(fields) != 2:
                raise AnalysisError(f"malformed allocation event at {rows}")
            index = parse_hex(fields[1], "allocation-info index", rows)
            if index >= len(allocation_info):
                raise AnalysisError(f"unknown allocation-info index at record {rows}")
            (plus if mode == "+" else minus)[index] += 1
        # Header and process metadata records (v, X, I, c, ...) are retained
        # by Heaptrack but do not participate in the allocation histogram.

    names_cache: dict[int, tuple[str, ...]] = {}

    def names_for_trace(trace_id: int) -> tuple[str, ...]:
        if trace_id in names_cache:
            return names_cache[trace_id]
        names: list[str] = []
        seen: set[int] = set()
        current = trace_id
        while current:
            if current in seen:
                raise AnalysisError("trace ancestry cycle detected")
            seen.add(current)
            instruction, parent = traces.get(current, (0, 0))
            for string_id in instructions.get(instruction, ()):
                name = strings.get(string_id)
                if name is None:
                    raise AnalysisError(f"unknown string {string_id} in trace")
                names.append(name)
            current = parent
        names_cache[trace_id] = tuple(names)
        return names_cache[trace_id]

    def open_call_key(trace_id: int) -> str:
        current = trace_id
        while current:
            instruction, parent = traces[current]
            if any(RUN_ONE in name for name in instructions.get(instruction, ())
                   for name in [strings.get(name, "")]):
                # The trace id of run_one is stable for each distinct call
                # site in this one process, even though the symbol itself is
                # shared by the timed and verification opens.
                return f"run_one_call_trace_{current}"
            current = parent
        # The target is still attributable if a future probe changes its
        # caller symbol, but such a trace cannot be safely split by call site.
        return "unclassified_open_call"

    target_rows: list[dict[str, Any]] = []
    all_requested_bytes = 0
    all_events = 0
    all_deallocated_bytes = 0
    all_deallocation_events = 0
    for index, (size, trace_id) in enumerate(allocation_info):
        occurrences = plus[index]
        deallocations = minus[index]
        all_events += occurrences
        all_requested_bytes += size * occurrences
        all_deallocation_events += deallocations
        all_deallocated_bytes += size * deallocations
        if not occurrences:
            continue
        names = names_for_trace(trace_id)
        if not any(TARGET in name for name in names):
            continue
        alloc_kind = "unknown"
        if any("GlobalAlloc7realloc" in name for name in names):
            alloc_kind = "realloc"
        elif any("GlobalAlloc5alloc" in name for name in names):
            alloc_kind = "alloc"
        target_rows.append(
            {
                "allocation_info_index": index,
                "size": size,
                "allocation_events": occurrences,
                "deallocation_events": deallocations,
                "requested_bytes": size * occurrences,
                "deallocated_bytes": size * deallocations,
                "allocator_kind": alloc_kind,
                "open_call_group": open_call_key(trace_id),
                "trace_id": trace_id,
            }
        )

    by_call: dict[str, dict[str, Any]] = {}
    for row in target_rows:
        key = row["open_call_group"]
        summary = by_call.setdefault(
            key,
            {
                "open_call_group": key,
                "first_allocation_info_index": row["allocation_info_index"],
                "allocation_info_records": 0,
                "allocation_events": 0,
                "deallocation_events": 0,
                "requested_bytes": 0,
                "deallocated_bytes": 0,
                "allocator_kinds": {},
            },
        )
        summary["allocation_info_records"] += 1
        summary["allocation_events"] += row["allocation_events"]
        summary["deallocation_events"] += row["deallocation_events"]
        summary["requested_bytes"] += row["requested_bytes"]
        summary["deallocated_bytes"] += row["deallocated_bytes"]
        kind = row["allocator_kind"]
        kind_summary = summary["allocator_kinds"].setdefault(
            kind,
            {
                "allocation_info_records": 0,
                "allocation_events": 0,
                "requested_bytes": 0,
                "deallocation_events": 0,
                "deallocated_bytes": 0,
            },
        )
        kind_summary["allocation_info_records"] += 1
        kind_summary["allocation_events"] += row["allocation_events"]
        kind_summary["requested_bytes"] += row["requested_bytes"]
        kind_summary["deallocation_events"] += row["deallocation_events"]
        kind_summary["deallocated_bytes"] += row["deallocated_bytes"]

    target_requested = sum(row["requested_bytes"] for row in target_rows)
    target_events = sum(row["allocation_events"] for row in target_rows)
    target_deallocated = sum(row["deallocated_bytes"] for row in target_rows)
    target_deallocation_events = sum(row["deallocation_events"] for row in target_rows)
    calls = sorted(
        by_call.values(),
        key=lambda row: (row["first_allocation_info_index"], row["open_call_group"]),
    )
    for order, call in enumerate(calls, start=1):
        if order == 1:
            call["open_call"] = "timed_open"
        elif order == 2:
            call["open_call"] = "post_clock_verification_open"
        else:
            call["open_call"] = f"unexpected_open_call_{order}"
    return {
        "format": "interpreted_heaptrack_v3",
        "trace_records": len(traces),
        "instruction_records": len(instructions),
        "allocation_info_records": len(allocation_info),
        "trace_records_read": rows,
        "all_allocation_events": all_events,
        "all_requested_bytes": all_requested_bytes,
        "all_deallocation_events": all_deallocation_events,
        "all_deallocated_bytes": all_deallocated_bytes,
        "target": TARGET,
        "target_allocation_info_records": len(target_rows),
        "target_allocation_events": target_events,
        "target_requested_bytes": target_requested,
        "target_deallocation_events": target_deallocation_events,
        "target_deallocated_bytes": target_deallocated,
        "target_rows": target_rows,
        "open_calls": calls,
    }


def arithmetic_model(source_bytes: int) -> dict[str, Any]:
    """Model the pre-0779 exact-reserve sequence for one read_limited call."""
    if source_bytes < 0:
        raise AnalysisError("source byte count cannot be negative")
    initial = min(CHUNK_BYTES, source_bytes)
    capacity = initial
    length = 0
    growth_capacities: list[int] = []
    while length < source_bytes:
        read = min(CHUNK_BYTES, source_bytes - length)
        required = length + read
        if required > capacity:
            capacity = required
            growth_capacities.append(capacity)
        length = required
    return {
        "source_bytes": source_bytes,
        "chunk_bytes": CHUNK_BYTES,
        "initial_allocation_bytes": initial,
        "growth_reallocation_count": len(growth_capacities),
        "growth_requested_bytes": sum(growth_capacities),
        "requested_bytes": initial + sum(growth_capacities),
        "total_allocation_events": 1 + len(growth_capacities),
        "last_capacity": capacity,
    }


def parse_print_summary(path: Path) -> dict[str, Any]:
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as exc:
        raise AnalysisError(f"cannot read Heaptrack print log {path}: {exc}") from exc
    result: dict[str, Any] = {}
    for line in text.splitlines():
        if line.startswith("calls to allocation functions:"):
            result["allocation_events"] = int(line.split(":", 1)[1].split("(", 1)[0].strip())
        elif line.startswith("temporary memory allocations:"):
            result["temporary_allocations"] = int(
                line.split(":", 1)[1].split("(", 1)[0].strip()
            )
        elif line.startswith("total memory leaked:"):
            result["total_memory_leaked"] = int(
                line.split(":", 1)[1].strip().rstrip("B")
            )
    if "allocation_events" not in result:
        raise AnalysisError(f"Heaptrack print log has no allocation total: {path}")
    return result


def path_context(root: Path) -> tuple[Path, Path]:
    """Return captured owned-worktree and current-repository roots."""
    origin_path = root / "origin.json"
    origin = require_json_object(origin_path)
    owned = origin.get("owned_worktree")
    if not isinstance(owned, str) or not owned:
        raise AnalysisError("origin.json owned_worktree is missing")
    owned_worktree = Path(owned).resolve(strict=False)
    resolved_root = root.resolve(strict=False)
    parts = resolved_root.parts
    if (
        len(parts) < 4
        or resolved_root.name != "change-0779"
        or resolved_root.parent.name != "results"
        or resolved_root.parent.parent.name != "performance"
        or resolved_root.parent.parent.parent.name != "docs"
    ):
        raise AnalysisError(f"cannot derive current repository root from {root}")
    current_repo = resolved_root.parents[3]
    return owned_worktree, current_repo


def load_cleanup(root: Path) -> tuple[Any, bool]:
    path = root / "cleanup.json"
    if not path.is_file() or path.is_symlink():
        return None, False
    cleanup = read_json(path)
    if not isinstance(cleanup, dict):
        raise AnalysisError("cleanup.json is malformed")
    verified = cleanup.get("verified") is True or cleanup.get(
        "executables_verified_before_removal"
    ) is True
    return cleanup, verified


def verify_role(
    root: Path,
    role: str,
    *,
    owned_worktree: Path,
    current_repo: Path,
    cleanup: Any,
    cleanup_verified: bool,
) -> dict[str, Any]:
    folder = root / f"heaptrack-{role}"
    if not folder.is_dir():
        raise AnalysisError(f"missing Heaptrack role directory: {folder}")
    capture_path = folder / "capture.json"
    decode_path = folder / "decode.json"
    report_path = folder / "report.json"
    histogram_path = folder / "histogram"
    print_path = folder / "print.log"
    trace_path = folder / "open.zst"
    capture = require_json_object(capture_path)
    decode = require_json_object(decode_path)
    report = require_json_object(report_path)
    for path in (capture_path, decode_path, report_path, histogram_path, print_path, trace_path):
        artifact(path, root)

    if capture.get("exit_code") != 0:
        raise AnalysisError(f"Heaptrack capture failed for {role}")
    if decode.get("exit_code") != 0:
        raise AnalysisError(f"heaptrack_print decode failed for {role}")
    source = report.get("source")
    if (
        not isinstance(source, dict)
        or not isinstance(source.get("path"), str)
        or not isinstance(source.get("bytes"), int)
        or not isinstance(source.get("sha256"), str)
    ):
        raise AnalysisError(f"report source identity is missing for {role}")
    if len(source["sha256"]) != 64 or source["bytes"] < 0:
        raise AnalysisError(f"report source identity is malformed for {role}")
    samples = report.get("samples")
    if (
        report.get("phase") != "open"
        or report.get("samples_requested") != 1
        or report.get("warmup") != 0
        or not isinstance(samples, list)
        or len(samples) != 1
        or not isinstance(samples[0], dict)
        or samples[0].get("index") != 0
        or not isinstance(samples[0].get("elapsed_ns"), int)
    ):
        raise AnalysisError(f"Heaptrack report is not a one-sample open run: {role}")
    if (
        report.get("schema") != "litchi.xlsx.allocation-probe.v1"
        or report.get("tool") != "litchi-xlsx-allocation-attribution-probe-0779"
        or report.get("sheet") != "Sheet1"
        or report.get("address") != "A1"
        or report.get("marker") != "litchi-perf-0638-ordinary-save"
        or report.get("phase") != "open"
        or report.get("timing_scope") != "Workbook::open(source) only"
    ):
        raise AnalysisError(f"Heaptrack report input flags differ for {role}")

    capture_command = capture.get("command")
    if not isinstance(capture_command, list) or not all(
        isinstance(value, str) for value in capture_command
    ):
        raise AnalysisError(f"capture command is malformed for {role}")
    binary_receipt = verify_binary_receipt(
        capture.get("binary"),
        capture_command,
        cleanup=cleanup,
        cleanup_verified=cleanup_verified,
        label=f"{role} capture binary",
    )
    expected_capture_log = folder / "capture.log"
    expected_trace_stem = folder / "open"
    expected_decode_log = folder / "print.log"
    expected_histogram = folder / "histogram"
    expected_report = folder / "report.json"
    expected_source = mapped_receipt_path(
        source["path"], owned_worktree, current_repo
    )
    expected_capture_command = [
        "taskset",
        "-c",
        "12",
        "heaptrack",
        "-o",
        str(expected_trace_stem),
        capture_command[6],
        "--source",
        str(expected_source),
        "--sheet",
        "Sheet1",
        "--phase",
        "open",
        "--samples",
        "1",
        "--warmup",
        "0",
        "--output",
        str(expected_report),
    ]
    normalized_capture_command = list(capture_command)
    for index in (5, 8, 18):
        if index >= len(normalized_capture_command):
            raise AnalysisError(f"capture command is truncated for {role}")
        normalized_capture_command[index] = str(
            mapped_receipt_path(
                normalized_capture_command[index], owned_worktree, current_repo
            )
        )
    if normalized_capture_command != expected_capture_command:
        raise AnalysisError(f"capture command flags or paths differ for {role}")
    if capture.get("log") is None:
        raise AnalysisError(f"capture log binding is missing for {role}")
    capture_log_binding = verify_captured_artifact(
        capture["log"],
        expected_capture_log,
        root,
        f"{role} capture log",
        owned_worktree=owned_worktree,
        current_repo=current_repo,
    )

    decode_command = decode.get("command")
    if not isinstance(decode_command, list) or not all(
        isinstance(value, str) for value in decode_command
    ):
        raise AnalysisError(f"decode command is malformed for {role}")
    normalized_decode_command = list(decode_command)
    for index in (2, 4):
        if index >= len(normalized_decode_command):
            raise AnalysisError(f"decode command is truncated for {role}")
        normalized_decode_command[index] = str(
            mapped_receipt_path(
                normalized_decode_command[index], owned_worktree, current_repo
            )
        )
    expected_decode_command = [
        "heaptrack_print",
        "-f",
        str(folder / "open.zst"),
        "-H",
        str(expected_histogram),
        "-n",
        "15",
        "-s",
        "3",
    ]
    if normalized_decode_command != expected_decode_command:
        raise AnalysisError(f"decode command flags or paths differ for {role}")
    if not isinstance(decode.get("trace"), dict):
        raise AnalysisError(f"decode trace binding is missing for {role}")
    if not isinstance(decode.get("histogram"), dict):
        raise AnalysisError(f"decode histogram binding is missing for {role}")
    if not isinstance(decode.get("log"), dict):
        raise AnalysisError(f"decode log binding is missing for {role}")
    trace_binding = verify_captured_artifact(
        decode["trace"],
        folder / "open.zst",
        root,
        f"{role} decode trace",
        owned_worktree=owned_worktree,
        current_repo=current_repo,
    )
    histogram_binding = verify_captured_artifact(
        decode["histogram"],
        expected_histogram,
        root,
        f"{role} decode histogram",
        owned_worktree=owned_worktree,
        current_repo=current_repo,
    )
    decode_log_binding = verify_captured_artifact(
        decode["log"],
        expected_decode_log,
        root,
        f"{role} decode log",
        owned_worktree=owned_worktree,
        current_repo=current_repo,
    )
    if (
        mapped_receipt_path(capture_command[-1], owned_worktree, current_repo).resolve()
        != expected_report.resolve()
    ):
        raise AnalysisError(f"capture report path differs for {role}")
    if expected_source.is_file() and not expected_source.is_symlink():
        if expected_source.stat().st_size != source["bytes"]:
            raise AnalysisError(f"report source byte identity differs for {role}")
        if sha256_file(expected_source) != source["sha256"]:
            raise AnalysisError(f"report source SHA-256 identity differs for {role}")
    if report.get("allocator", {}).get("binary") != "litchi-xlsx-allocation-probe-0779":
        raise AnalysisError(f"report allocator binary identity differs for {role}")
    histogram = parse_histogram(histogram_path)
    trace = parse_trace(trace_path)
    print_summary = parse_print_summary(print_path)
    if trace["all_allocation_events"] != histogram["allocation_events"]:
        raise AnalysisError(
            f"{role}: trace event count {trace['all_allocation_events']} "
            f"does not match histogram {histogram['allocation_events']}"
        )
    if trace["all_requested_bytes"] != histogram["requested_bytes"]:
        raise AnalysisError(
            f"{role}: trace requested bytes {trace['all_requested_bytes']} "
            f"do not match histogram {histogram['requested_bytes']}"
        )
    if print_summary["allocation_events"] != histogram["allocation_events"]:
        raise AnalysisError(
            f"{role}: print-log allocation count does not match histogram"
        )
    if len(trace["open_calls"]) == 0:
        raise AnalysisError(f"{role}: no target read_limited open call found")
    model = arithmetic_model(source["bytes"])
    target = {
        "run_one_samples": 1,
        "open_call_count": len(trace["open_calls"]),
        "whole_process": {
            "allocation_info_records": trace["target_allocation_info_records"],
            "allocation_events": trace["target_allocation_events"],
            "requested_bytes": trace["target_requested_bytes"],
            "initial_alloc_events": sum(
                row["allocator_kinds"].get("alloc", {}).get("allocation_events", 0)
                for row in trace["open_calls"]
            ),
            "growth_realloc_events": sum(
                row["allocator_kinds"].get("realloc", {}).get("allocation_events", 0)
                for row in trace["open_calls"]
            ),
            "deallocation_events": trace["target_deallocation_events"],
            "deallocated_bytes": trace["target_deallocated_bytes"],
        },
        "per_open_call": trace["open_calls"],
        "arithmetic_model_for_one_call": model,
    }
    for open_call in target["per_open_call"]:
        open_call["arithmetic_delta_requested_bytes"] = (
            open_call["requested_bytes"] - model["requested_bytes"]
        )
        open_call["arithmetic_delta_growth_realloc_events"] = (
            open_call["allocator_kinds"].get("realloc", {}).get("allocation_events", 0)
            - model["growth_reallocation_count"]
        )
    return {
        "role": role,
        "artifacts": {
            name: artifact(path, root)
            for name, path in {
                "capture": capture_path,
                "decode": decode_path,
                "report": report_path,
                "histogram": histogram_path,
                "print_log": print_path,
                "trace": trace_path,
            }.items()
        },
        "capture": {
            "command": capture.get("command"),
            "binary": binary_receipt,
            "exit_code": capture.get("exit_code"),
        },
        "verification": {
            "capture_log": capture_log_binding,
            "decode_trace": trace_binding,
            "decode_histogram": histogram_binding,
            "decode_log": decode_log_binding,
            "capture_command_shape": "exact",
            "decode_command_shape": "exact",
            "source_identity": source,
            "report_identity": artifact(report_path, root),
        },
        "report": {
            "schema": report.get("schema"),
            "tool": report.get("tool"),
            "phase": report.get("phase"),
            "samples": report.get("samples"),
            "source": source,
        },
        "histogram": histogram,
        "print_summary": print_summary,
        "trace": {
            key: value
            for key, value in trace.items()
            if key not in {"target_rows"}
        },
        "read_limited": target,
    }


def write_json(path: Path, value: Any) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(
        json.dumps(value, indent=2, sort_keys=False, ensure_ascii=True) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


def analyze(root: Path | str) -> dict[str, Any]:
    """Purely replay and return the deterministic analysis result.

    This API never writes ``heaptrack-analysis.json``.  A validator can import
    the module, call ``analyze(packet_root)``, and compare the returned object
    with the retained JSON after the packet has been relocated or its build
    binaries have been removed under an exact cleanup witness.
    """
    root = Path(root).resolve()
    owned_worktree, current_repo = path_context(root)
    cleanup, cleanup_verified = load_cleanup(root)
    roles = [role for role in ("before", "after") if (root / f"heaptrack-{role}").is_dir()]
    if not roles:
        raise AnalysisError("no heaptrack-before or heaptrack-after directory found")
    role_results = {
        role: verify_role(
            root,
            role,
            owned_worktree=owned_worktree,
            current_repo=current_repo,
            cleanup=cleanup,
            cleanup_verified=cleanup_verified,
        )
        for role in roles
    }
    if "before" in role_results and "after" in role_results:
        before_source = role_results["before"]["report"]["source"]
        after_source = role_results["after"]["report"]["source"]
        if before_source != after_source:
            raise AnalysisError("before/after report source identities differ")
    comparison: dict[str, Any] = {}
    if "before" in role_results and "after" in role_results:
        before = role_results["before"]["read_limited"]["whole_process"]
        after = role_results["after"]["read_limited"]["whole_process"]
        comparison = {
            "before_role": "before",
            "after_role": "after",
            "whole_process_delta": {
                "allocation_events": after["allocation_events"] - before["allocation_events"],
                "growth_realloc_events": after["growth_realloc_events"]
                - before["growth_realloc_events"],
                "requested_bytes": after["requested_bytes"] - before["requested_bytes"],
            },
            "whole_process_after_to_before_ratio": {
                "allocation_events": after["allocation_events"] / before["allocation_events"],
                "growth_realloc_events": after["growth_realloc_events"]
                / before["growth_realloc_events"],
                "requested_bytes": after["requested_bytes"] / before["requested_bytes"],
            },
        }
    result = {
        "schema": "litchi.xlsx.heaptrack-analysis.v1",
        "status": "pass",
        "scope": {
            "experiment": "0779",
            "phase": "open",
            "source_scope": "whole process, including post-clock workbook verification",
            "attribution_basis": (
                "Heaptrack allocation request size multiplied by live '+' events "
                "whose interpreted ancestry contains read_limited"
            ),
            "does_not_claim": [
                "physical bytes copied by realloc",
                "resident memory or RSS improvement",
                "latency improvement",
            ],
        },
        "roles": role_results,
        "comparison": comparison,
    }
    return result


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="change-0779 result directory",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="analysis JSON output (default: ROOT/heaptrack-analysis.json)",
    )
    args = parser.parse_args(argv)
    root = args.root.resolve()
    output = (args.output or (root / "heaptrack-analysis.json")).resolve()
    if output.parent != root:
        raise AnalysisError("analysis output must be directly inside --root")
    result = analyze(root)
    write_json(output, result)
    print(json.dumps({"status": "pass", "roles": roles, "output": str(output)}))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except AnalysisError as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(2)
