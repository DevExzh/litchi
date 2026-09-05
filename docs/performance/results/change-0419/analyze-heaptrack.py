#!/usr/bin/env python3
"""Replayable Heaptrack 1.5.0 postprocessing for change 0419.

The input capture and its capture.json are immutable.  The first pass runs
heaptrack_print once without a filter and once with the symbol discovered in
the unfiltered report.  Later ``--replay`` passes verify retained files and
recompute the numerical aggregate from the compressed trace; they do not need
the profiled binary or its worktree.

Heaptrack's v1.5.0 source is retained as an explicit semantic binding:
https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/src/analyze/print/heaptrack_print.cpp
``handleAllocation`` fills ``sizeHistogram`` while reading, while
``finalize`` applies ``filterAllocations`` afterwards.  Consequently -H is a
whole-trace histogram even when --filter-bt-function is present.  The script
reports sum(size * count) only for the unfiltered histogram.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time
from typing import Any


DEFAULT_ROOT = Path(__file__).resolve().parent
HEAPTRACK_SOURCE = (
    "https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/"
    "src/analyze/print/heaptrack_print.cpp"
)
MAX_HISTOGRAM_ROWS = 10_000_000
MAX_TRACE_BYTES = 512 * 1024 * 1024
MAX_TRACE_ROWS = 20_000_000
MAX_PER_TRACE_ROWS = 100
DIRECT_WRITER_PREFIX = "_RNvXs4_"


class AnalysisError(RuntimeError):
    pass


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON constant: {value}")


def reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def load_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=reject_duplicate_keys,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, ValueError, json.JSONDecodeError) as exc:
        raise AnalysisError(f"cannot load JSON {path}: {exc}") from exc


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as exc:
        raise AnalysisError(f"cannot hash {path}: {exc}") from exc
    return digest.hexdigest()


def artifact(path: Path, root: Path) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        raise AnalysisError(f"artifact is not a regular file: {path}")
    try:
        name = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError as exc:
        raise AnalysisError(f"artifact is outside {root}: {path}") from exc
    return {
        "path": name,
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
    }


def verify_gzip_sidecar(
    sidecar: Path,
    root: Path,
    *,
    expected_bytes: int,
    expected_sha256: str,
) -> dict[str, Any]:
    """Validate one compressed replacement for a missing capture artifact."""
    digest = hashlib.sha256()
    total = 0
    try:
        with gzip.open(sidecar, "rb") as stream:
            remaining = expected_bytes + 1
            while remaining:
                block = stream.read(min(1024 * 1024, remaining))
                if not block:
                    break
                total += len(block)
                digest.update(block)
                remaining -= len(block)
    except (OSError, EOFError) as exc:
        raise AnalysisError(f"cannot read gzip sidecar {sidecar}: {exc}") from exc
    if total != expected_bytes:
        raise AnalysisError(
            f"gzip sidecar logical size mismatch for {sidecar}: "
            f"expected {expected_bytes}, observed at most {total}"
        )
    observed_sha256 = digest.hexdigest()
    if observed_sha256 != expected_sha256:
        raise AnalysisError(f"gzip sidecar logical SHA256 mismatch for {sidecar}")
    return {
        "logical_name": sidecar.name.removesuffix(".gz"),
        "compressed": artifact(sidecar, root),
        "logical_bytes": total,
        "logical_sha256": observed_sha256,
    }


def utc_now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def write_json(path: Path, value: Any) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(
        json.dumps(value, indent=2, sort_keys=False, ensure_ascii=True) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


def capture_trace(folder: Path, capture: dict[str, Any]) -> tuple[Path, dict[str, Any]]:
    if capture.get("status") != "pass" or capture.get("exit_code") != 0:
        raise AnalysisError("capture.json is not a successful capture")
    if capture.get("source_unchanged") is not True or capture.get("binary_unchanged") is not True:
        raise AnalysisError("capture source/binary binding is not unchanged")
    argv = capture.get("argv")
    if not isinstance(argv, list) or not argv or not all(isinstance(item, str) for item in argv):
        raise AnalysisError("capture argv is missing or malformed")
    files = capture.get("files")
    if not isinstance(files, list) or not files:
        raise AnalysisError("capture file inventory is missing")
    listed: dict[str, dict[str, Any]] = {}
    sidecars: list[dict[str, Any]] = []
    for item in files:
        if not isinstance(item, dict) or not isinstance(item.get("name"), str):
            raise AnalysisError("capture file inventory contains a malformed entry")
        name = item["name"]
        path = Path(name)
        if path.is_absolute() or path.name != name or name in listed:
            raise AnalysisError(f"unsafe or duplicate capture artifact name: {name!r}")
        if not isinstance(item.get("bytes"), int) or item["bytes"] < 0:
            raise AnalysisError(f"invalid byte count for capture artifact {name!r}")
        expected_hash = item.get("sha256")
        if not isinstance(expected_hash, str) or len(expected_hash) != 64:
            raise AnalysisError(f"invalid hash for capture artifact {name!r}")
        candidate = folder / name
        if not candidate.is_file() or candidate.is_symlink():
            # The capture record names the original producer stderr.  A
            # publication step may replace only that noisy file with a gzip
            # sidecar; no other capture artifact gets this fallback.
            if name == "stderr.log":
                sidecar = folder / "stderr.log.gz"
                if sidecar.is_file() and not sidecar.is_symlink():
                    sidecars.append(
                        verify_gzip_sidecar(
                            sidecar,
                            folder,
                            expected_bytes=item["bytes"],
                            expected_sha256=expected_hash,
                        )
                    )
                    listed[name] = item
                    continue
            raise AnalysisError(f"capture artifact is missing: {candidate}")
        observed = artifact(candidate, folder)
        if observed["bytes"] != item["bytes"] or observed["sha256"] != expected_hash:
            raise AnalysisError(f"capture artifact hash mismatch: {candidate}")
        listed[name] = item
    traces = [
        folder / name
        for name in listed
        if (name == "trace" or name.startswith("trace."))
        and Path(name).suffix in {"", ".zst", ".gz", ".data", ".bin"}
    ]
    if len(traces) != 1:
        raise AnalysisError(f"expected one retained Heaptrack trace, found {traces!r}")
    return traces[0], {
        "capture": artifact(folder / "capture.json", folder),
        "capture_sha256": sha256_file(folder / "capture.json"),
        "trace": artifact(traces[0], folder),
        "trace_argv": argv,
        "trace_exit_code": capture["exit_code"],
        "trace_status": capture["status"],
        "protocol_sha256": capture.get("protocol_sha256"),
        "build_sha256": capture.get("build_sha256"),
        "capture_sidecars": sidecars,
    }


def verify_sidecar_bindings(folder: Path, binding: dict[str, Any]) -> None:
    if not isinstance(binding, dict):
        raise AnalysisError("capture sidecar binding is not an object")
    entries = binding.get("capture_sidecars", [])
    if not isinstance(entries, list):
        raise AnalysisError("capture sidecar binding is not a list")
    for expected in entries:
        if not isinstance(expected, dict):
            raise AnalysisError("malformed capture sidecar binding")
        compressed = expected.get("compressed")
        if not isinstance(compressed, dict) or not isinstance(compressed.get("path"), str):
            raise AnalysisError("capture sidecar compressed artifact binding is missing")
        sidecar = folder / compressed["path"]
        if artifact(sidecar, folder) != compressed:
            raise AnalysisError(f"capture sidecar changed: {sidecar}")
        logical_bytes = expected.get("logical_bytes")
        logical_sha256 = expected.get("logical_sha256")
        if not isinstance(logical_bytes, int) or logical_bytes < 0 or not isinstance(logical_sha256, str):
            raise AnalysisError("capture sidecar logical binding is malformed")
        observed = verify_gzip_sidecar(
            sidecar,
            folder,
            expected_bytes=logical_bytes,
            expected_sha256=logical_sha256,
        )
        if observed["logical_name"] != expected.get("logical_name"):
            raise AnalysisError(f"capture sidecar logical name changed: {sidecar}")


def capture_binding_for_manifest(
    folder: Path,
    manifest: dict[str, Any],
    *,
    require_sidecar_match: bool,
) -> tuple[Path, dict[str, Any]]:
    capture_path = folder / "capture.json"
    capture = load_json(capture_path)
    if not isinstance(capture, dict):
        raise AnalysisError("capture.json must be an object")
    trace, observed = capture_trace(folder, capture)
    expected = manifest.get("capture")
    if not isinstance(expected, dict):
        raise AnalysisError("analysis capture binding is missing")
    if expected.get("capture_sha256") != observed["capture_sha256"]:
        raise AnalysisError("capture.json changed since analysis")
    if expected.get("trace") != observed["trace"]:
        raise AnalysisError("Heaptrack trace binding changed since analysis")
    if require_sidecar_match and expected.get("capture_sidecars", []) != observed.get("capture_sidecars", []):
        raise AnalysisError("capture sidecar binding changed; run --refresh")
    return trace, observed


def optional_binding(root: Path, role: str, capture: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    protocol = root / "protocol.json"
    expected_protocol = capture.get("protocol_sha256")
    if protocol.is_file():
        observed = sha256_file(protocol)
        result["protocol"] = {"path": str(protocol), "sha256": observed}
        if isinstance(expected_protocol, str) and observed != expected_protocol:
            raise AnalysisError("protocol hash differs from capture.json binding")
    else:
        result["protocol"] = {"present": False, "captured_sha256": expected_protocol}
    build = root / f"build-{role}.json"
    expected_build = capture.get("build_sha256")
    if build.is_file():
        observed = sha256_file(build)
        result["build"] = {"path": str(build), "sha256": observed}
        if isinstance(expected_build, str) and observed != expected_build:
            raise AnalysisError("build hash differs from capture.json binding")
        raw = load_json(build)
        if isinstance(raw, dict):
            source = raw.get("source_after")
            binary = raw.get("binaries", {}).get("normal") if isinstance(raw.get("binaries"), dict) else None
            if isinstance(source, dict):
                result["source_revision"] = source.get("revision")
                result["source_worktree_at_capture"] = source.get("worktree")
            if isinstance(binary, dict):
                result["binary"] = {
                    "label": binary.get("label"),
                    "sha256": binary.get("sha256", binary.get("binary_sha256")),
                    "path_at_capture": binary.get("path"),
                }
    else:
        result["build"] = {"present": False, "captured_sha256": expected_build}
    return result


def histogram(path: Path) -> dict[str, Any]:
    total = 0
    counts = 0
    rows = 0
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
            raise AnalysisError(f"invalid Heaptrack histogram row: {line!r}")
        size, count = (int(field) for field in fields)
        total += size * count
        counts += count
    if rows == 0:
        raise AnalysisError(f"empty Heaptrack histogram: {path}")
    return {"rows": rows, "sum_counts": counts, "sum_size_times_count": total}


def trace_lines(path: Path):
    """Yield interpreted Heaptrack lines without requiring the debuggee."""
    if path.suffix == ".zst":
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
        total_bytes = 0
        try:
            for raw in process.stdout:
                total_bytes += len(raw)
                if total_bytes > MAX_TRACE_BYTES:
                    process.kill()
                    raise AnalysisError(f"decompressed trace exceeds {MAX_TRACE_BYTES} bytes")
                yield raw
        finally:
            process.stdout.close()
        error = process.stderr.read().decode("utf-8", "replace")
        process.stderr.close()
        if process.wait() != 0:
            raise AnalysisError(f"zstd failed for {path}: {error.strip()}")
        return
    if path.suffix == ".gz":
        import gzip

        stream = gzip.open(path, "rb")
    else:
        stream = path.open("rb")
    total_bytes = 0
    try:
        for raw in stream:
            total_bytes += len(raw)
            if total_bytes > MAX_TRACE_BYTES:
                raise AnalysisError(f"decompressed trace exceeds {MAX_TRACE_BYTES} bytes")
            yield raw
    finally:
        stream.close()


def hex_token(token: str, label: str) -> int:
    try:
        return int(token, 16)
    except ValueError as exc:
        raise AnalysisError(f"invalid hexadecimal {label}: {token!r}") from exc


def raw_trace_aggregate(path: Path, symbol: str) -> dict[str, Any]:
    """Aggregate + events by interpreted trace ancestry.

    Heaptrack v1.5.0 writes ``a size trace`` allocation-info records and ``+
    allocation-info-index`` events.  Summing sizes for + events reproduces
    the unfiltered -H histogram.  This parser deliberately does not inspect
    the binary or resolve addresses; all function names needed for the
    filter are already present in the interpreted trace.
    """
    strings: dict[int, str] = {}
    ips: dict[int, tuple[int, ...]] = {}
    traces: dict[int, tuple[int, int]] = {}
    allocation_infos: list[tuple[int, int]] = []
    event_counts: dict[int, int] = {}
    deallocation_counts: dict[int, int] = {}
    plus_events = 0
    minus_events = 0
    rows = 0
    for raw in trace_lines(path):
        rows += 1
        if rows > MAX_TRACE_ROWS:
            raise AnalysisError(f"trace has more than {MAX_TRACE_ROWS} records")
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
            index = len(strings) + 1
            declared = hex_token(fields[1], "string length")
            text = line.split(None, 2)[2]
            if len(text.encode("utf-8")) != declared:
                raise AnalysisError(f"string length mismatch at record {rows}")
            strings[index] = text
        elif mode == "i":
            if len(fields) < 3:
                raise AnalysisError(f"malformed instruction record at {rows}")
            tokens = fields[1:]
            # address and module are followed by an optional function/file/
            # line tuple and zero or more inline function/file/line tuples.
            function_ids: list[int] = []
            if len(tokens) >= 3:
                function_id = hex_token(tokens[2], "function index")
                if function_id:
                    function_ids.append(function_id)
                remaining = tokens[3:]
                if remaining:
                    if len(remaining) < 2 or (len(remaining) - 2) % 3:
                        raise AnalysisError(f"malformed instruction frames at {rows}")
                    for index in range(2, len(remaining), 3):
                        function_id = hex_token(remaining[index], "inline function index")
                        if function_id:
                            function_ids.append(function_id)
            ips[len(ips) + 1] = tuple(function_ids)
        elif mode == "t":
            if len(fields) != 3:
                raise AnalysisError(f"malformed trace record at {rows}")
            trace_id = len(traces) + 1
            ip_id = hex_token(fields[1], "instruction index")
            parent = hex_token(fields[2], "trace parent")
            if ip_id and ip_id not in ips:
                raise AnalysisError(f"trace instruction index is unknown at record {rows}")
            if parent and parent >= trace_id:
                raise AnalysisError(f"trace parent is forward or self-referential at record {rows}")
            traces[trace_id] = (ip_id, parent)
        elif mode == "a":
            if len(fields) != 3:
                raise AnalysisError(f"malformed allocation-info record at {rows}")
            size = hex_token(fields[1], "allocation size")
            trace_id = hex_token(fields[2], "allocation trace")
            if trace_id and trace_id not in traces:
                raise AnalysisError(f"allocation trace is unknown at record {rows}")
            allocation_infos.append((size, trace_id))
        elif mode == "+":
            if len(fields) != 2:
                raise AnalysisError(f"malformed allocation event at {rows}")
            index = hex_token(fields[1], "allocation-info index")
            if index >= len(allocation_infos):
                raise AnalysisError(f"allocation event index is unknown at record {rows}")
            event_counts[index] = event_counts.get(index, 0) + 1
            plus_events += 1
        elif mode == "-":
            if len(fields) != 2:
                raise AnalysisError(f"malformed deallocation event at {rows}")
            index = hex_token(fields[1], "deallocation-info index")
            if index >= len(allocation_infos):
                raise AnalysisError(f"deallocation event index is unknown at record {rows}")
            deallocation_counts[index] = deallocation_counts.get(index, 0) + 1
            minus_events += 1

    trace_names_cache: dict[int, tuple[str, ...]] = {}

    def names_for_trace(trace_id: int) -> tuple[str, ...]:
        if trace_id in trace_names_cache:
            return trace_names_cache[trace_id]
        names: list[str] = []
        seen: set[int] = set()
        current = trace_id
        while current:
            if current in seen:
                raise AnalysisError("trace ancestry cycle detected")
            seen.add(current)
            ip_id, parent = traces.get(current, (0, 0))
            for string_id in ips.get(ip_id, ()):
                name = strings.get(string_id)
                if name is None:
                    raise AnalysisError(f"instruction references unknown string {string_id}")
                names.append(name)
            current = parent
        trace_names_cache[trace_id] = tuple(names)
        return trace_names_cache[trace_id]

    total_bytes = 0
    total_deallocated_bytes = 0
    matched_bytes = 0
    matched_deallocated_bytes = 0
    matched_events = 0
    matched_deallocation_events = 0
    rows_by_trace: dict[int, dict[str, Any]] = {}
    direct_bytes = 0
    direct_deallocated_bytes = 0
    direct_events = 0
    direct_deallocation_events = 0
    direct_rows_by_trace: dict[int, dict[str, Any]] = {}

    def direct_writer_frame(name: str) -> bool:
        # This is the concrete Rust Write::write implementation in the
        # current trace, rather than generic Counted/Chunked callers that
        # merely carry BoundedVecWriter in a type parameter.
        return (
            symbol == "BoundedVecWriter"
            and name.startswith(DIRECT_WRITER_PREFIX)
            and "16BoundedVecWriterNt" in name
            and "cross_copy_plan" in name
            and name.endswith("Write5write")
        )

    def add_row(
        table: dict[int, dict[str, Any]],
        trace_id: int,
        names: tuple[str, ...],
        count: int,
        deallocation_count: int,
        bytes_requested: int,
        bytes_deallocated: int,
    ) -> None:
        row = table.setdefault(
            trace_id,
            {
                "trace_id": trace_id,
                "allocation_info_records": 0,
                "allocation_events": 0,
                "deallocation_events": 0,
                "requested_bytes": 0,
                "deallocated_bytes": 0,
                "matching_frames": list(names),
                "trace_frames": list(names_for_trace(trace_id)),
            },
        )
        row["allocation_info_records"] += 1
        row["allocation_events"] += count
        row["deallocation_events"] += deallocation_count
        row["requested_bytes"] += bytes_requested
        row["deallocated_bytes"] += bytes_deallocated

    for index in sorted(set(event_counts) | set(deallocation_counts)):
        size, trace_id = allocation_infos[index]
        count = event_counts.get(index, 0)
        deallocation_count = deallocation_counts.get(index, 0)
        bytes_requested = size * count
        bytes_deallocated = size * deallocation_count
        total_bytes += bytes_requested
        total_deallocated_bytes += bytes_deallocated
        names = names_for_trace(trace_id)
        matching = tuple(name for name in names if symbol in name)
        if matching:
            matched_events += count
            matched_deallocation_events += deallocation_count
            matched_bytes += bytes_requested
            matched_deallocated_bytes += bytes_deallocated
            add_row(
                rows_by_trace, trace_id, matching, count, deallocation_count,
                bytes_requested, bytes_deallocated,
            )
        direct = tuple(name for name in names if direct_writer_frame(name))
        if direct:
            direct_events += count
            direct_deallocation_events += deallocation_count
            direct_bytes += bytes_requested
            direct_deallocated_bytes += bytes_deallocated
            add_row(
                direct_rows_by_trace, trace_id, direct, count, deallocation_count,
                bytes_requested, bytes_deallocated,
            )

    per_trace = sorted(
        rows_by_trace.values(),
        key=lambda row: (-row["requested_bytes"], row["trace_id"]),
    )
    direct_per_trace = sorted(
        direct_rows_by_trace.values(),
        key=lambda row: (-row["requested_bytes"], row["trace_id"]),
    )
    for row in per_trace:
        row["share_of_all_requested_percent"] = (
            row["requested_bytes"] * 100.0 / total_bytes if total_bytes else None
        )
    for row in direct_per_trace:
        row["share_of_all_requested_percent"] = (
            row["requested_bytes"] * 100.0 / total_bytes if total_bytes else None
        )
    bounded = per_trace[:MAX_PER_TRACE_ROWS]
    direct_bounded = direct_per_trace[:MAX_PER_TRACE_ROWS]
    return {
        "format": "interpreted_heaptrack_v3",
        "records": rows,
        "instruction_records": len(ips),
        "trace_records": len(traces),
        "allocation_info_records": len(allocation_infos),
        "allocation_events": plus_events,
        "deallocation_events": minus_events,
        "unfiltered_requested_bytes": total_bytes,
        "unfiltered_deallocated_bytes": total_deallocated_bytes,
        "writer_ancestor_events": matched_events,
        "writer_ancestor_deallocation_events": matched_deallocation_events,
        "writer_ancestor_requested_bytes": matched_bytes,
        "writer_ancestor_deallocated_bytes": matched_deallocated_bytes,
        "writer_ancestor_net_live_delta_bytes": matched_bytes - matched_deallocated_bytes,
        "writer_ancestor_share_of_all_requested_percent": (
            matched_bytes * 100.0 / total_bytes if total_bytes else None
        ),
        "writer_ancestor_trace_count": len(per_trace),
        "writer_ancestor_traces_retained": len(bounded),
        "writer_ancestor_traces_truncated": len(per_trace) > len(bounded),
        "writer_ancestor_traces": bounded,
        "direct_writer_events": direct_events,
        "direct_writer_deallocation_events": direct_deallocation_events,
        "direct_writer_requested_bytes": direct_bytes,
        "direct_writer_deallocated_bytes": direct_deallocated_bytes,
        "direct_writer_net_live_delta_bytes": direct_bytes - direct_deallocated_bytes,
        "direct_writer_share_of_all_requested_percent": (
            direct_bytes * 100.0 / total_bytes if total_bytes else None
        ),
        "direct_writer_trace_count": len(direct_per_trace),
        "direct_writer_traces_retained": len(direct_bounded),
        "direct_writer_traces_truncated": len(direct_per_trace) > len(direct_bounded),
        "direct_writer_traces": direct_bounded,
    }


def run_print(
    tool: Path,
    trace: Path,
    folder: Path,
    stem: str,
    *,
    filter_symbol: str | None,
    timeout: float,
) -> dict[str, Any]:
    stdout = folder / f"{stem}.stdout.txt"
    stderr = folder / f"{stem}.stderr.txt"
    hist = folder / f"{stem}.histogram.tsv"
    if any(path.exists() for path in (stdout, stderr, hist)):
        raise AnalysisError(f"refusing to overwrite existing analysis output for {stem}")
    command = [
        str(tool),
        "--file",
        str(trace),
        "--merge-backtraces",
        "0",
        "--print-allocators",
        "--print-temporary",
        "--print-peaks",
        "--peak-limit",
        "50",
        "--sub-peak-limit",
        "20",
        "--print-histogram",
        str(hist),
    ]
    if filter_symbol is not None:
        command.extend(("--filter-bt-function", filter_symbol))
    started = time.monotonic()
    env = os.environ.copy()
    env["DEBUGINFOD_URLS"] = ""
    try:
        with stdout.open("xb") as out, stderr.open("xb") as err:
            completed = subprocess.run(
                command,
                cwd=folder,
                env=env,
                stdout=out,
                stderr=err,
                check=False,
                timeout=timeout,
            )
    except subprocess.TimeoutExpired as exc:
        raise AnalysisError(f"heaptrack_print timed out for {stem}: {exc}") from exc
    elapsed = time.monotonic() - started
    if completed.returncode != 0:
        raise AnalysisError(f"heaptrack_print failed for {stem}: exit {completed.returncode}")
    if not hist.is_file():
        raise AnalysisError(f"heaptrack_print did not produce {hist}")
    return {
        "argv": command,
        "cwd": str(folder),
        "env_overrides": {"DEBUGINFOD_URLS": ""},
        "returncode": completed.returncode,
        "elapsed_seconds": elapsed,
        "stdout": artifact(stdout, folder),
        "stderr": artifact(stderr, folder),
        "histogram": artifact(hist, folder),
        "histogram_parse": histogram(hist),
    }


def symbol_lines(paths: list[Path], symbol: str) -> list[str]:
    found: list[str] = []
    seen: set[str] = set()
    for path in paths:
        try:
            lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
        except OSError as exc:
            raise AnalysisError(f"cannot inspect report {path}: {exc}") from exc
        for line in lines:
            stripped = line.strip()
            if symbol in stripped and stripped not in seen:
                seen.add(stripped)
                found.append(stripped)
    return found[:20]


def verify_replay(folder: Path) -> dict[str, Any]:
    manifest_path = folder / "analysis.json"
    manifest = load_json(manifest_path)
    if not isinstance(manifest, dict) or manifest.get("status") != "pass":
        raise AnalysisError(f"analysis manifest is not successful: {manifest_path}")
    trace, _ = capture_binding_for_manifest(folder, manifest, require_sidecar_match=True)
    outputs = manifest.get("outputs")
    if not isinstance(outputs, list) or not outputs:
        raise AnalysisError("analysis output inventory is missing")
    for expected in outputs:
        if not isinstance(expected, dict) or not isinstance(expected.get("path"), str):
            raise AnalysisError("malformed analysis output inventory")
        path = folder / expected["path"]
        observed = artifact(path, folder)
        if observed != expected:
            raise AnalysisError(f"analysis output changed: {path}")
    raw_expected = manifest.get("raw_trace")
    symbol = manifest.get("filter", {}).get("requested")
    if not isinstance(raw_expected, dict) or not isinstance(symbol, str) or not symbol:
        raise AnalysisError(
            "analysis manifest has no raw_trace aggregate; run without --replay or use --refresh"
        )
    raw_observed = raw_trace_aggregate(trace, symbol)
    if raw_observed != raw_expected:
        raise AnalysisError("compressed-trace numerical aggregate differs from analysis.json")
    overall = manifest.get("overall")
    if not isinstance(overall, dict) or raw_observed["unfiltered_requested_bytes"] != overall.get("cumulative_requested_bytes"):
        raise AnalysisError("raw unfiltered bytes differ from recorded histogram total")
    histogram_parse = overall.get("histogram_parse")
    if isinstance(histogram_parse, dict) and raw_observed["allocation_events"] != histogram_parse.get("sum_counts"):
        raise AnalysisError("raw allocation-event count differs from recorded histogram count")
    return manifest


def refresh_raw_manifest(folder: Path) -> dict[str, Any]:
    """Add the raw trace aggregate to an older manifest without rerunning tools."""
    manifest = verify_replay_without_raw(folder)
    trace, binding = capture_binding_for_manifest(folder, manifest, require_sidecar_match=False)
    symbol = manifest["filter"]["requested"]
    raw = raw_trace_aggregate(trace, symbol)
    overall = manifest.get("overall")
    if not isinstance(overall, dict) or raw["unfiltered_requested_bytes"] != overall.get("cumulative_requested_bytes"):
        raise AnalysisError("raw unfiltered bytes differ from recorded histogram total")
    manifest["capture"] = binding
    manifest["raw_trace"] = raw
    replay = manifest.get("replay")
    if not isinstance(replay, dict):
        raise AnalysisError("analysis replay metadata is missing")
    replay["raw_trace_refreshed_utc"] = utc_now()
    write_json(folder / "analysis.json", manifest)
    return manifest


def verify_replay_without_raw(folder: Path) -> dict[str, Any]:
    """Verify retained files for a pre-parser manifest before --refresh."""
    manifest_path = folder / "analysis.json"
    manifest = load_json(manifest_path)
    if not isinstance(manifest, dict) or manifest.get("status") != "pass":
        raise AnalysisError(f"analysis manifest is not successful: {manifest_path}")
    capture_binding_for_manifest(folder, manifest, require_sidecar_match=False)
    outputs = manifest.get("outputs")
    if not isinstance(outputs, list) or not outputs:
        raise AnalysisError("analysis output inventory is missing")
    for expected in outputs:
        if not isinstance(expected, dict) or not isinstance(expected.get("path"), str):
            raise AnalysisError("malformed analysis output inventory")
        path = folder / expected["path"]
        if artifact(path, folder) != expected:
            raise AnalysisError(f"analysis output changed: {path}")
    return manifest


def analyze_role(root: Path, role: str, args: argparse.Namespace) -> dict[str, Any]:
    folder = root / "heaptrack" / role
    if not folder.is_dir():
        raise AnalysisError(f"missing Heaptrack role directory: {folder}")
    manifest_path = folder / "analysis.json"
    if args.replay:
        return verify_replay(folder)
    if manifest_path.exists():
        if args.refresh:
            return refresh_raw_manifest(folder)
        return verify_replay(folder)
    capture_path = folder / "capture.json"
    capture = load_json(capture_path)
    if not isinstance(capture, dict):
        raise AnalysisError("capture.json must be an object")
    trace, binding = capture_trace(folder, capture)
    tool_text = shutil.which(args.heaptrack_print)
    if tool_text is None:
        tool = Path(args.heaptrack_print)
        if not tool.is_file() or not os.access(tool, os.X_OK):
            raise AnalysisError(f"heaptrack_print is unavailable: {args.heaptrack_print}")
    else:
        tool = Path(tool_text)
    try:
        version = subprocess.run(
            [str(tool), "--version"],
            capture_output=True,
            text=True,
            check=False,
            timeout=30,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise AnalysisError(f"cannot query heaptrack_print version: {exc}") from exc
    if version.returncode != 0:
        raise AnalysisError("heaptrack_print --version failed")
    tool_identity = {
        "path": str(tool.resolve()),
        "version": (version.stdout or version.stderr).strip(),
        "sha256": sha256_file(tool),
    }
    overall = run_print(
        tool,
        trace,
        folder,
        "heaptrack-print-overall",
        filter_symbol=None,
        timeout=args.timeout,
    )
    overall_report = folder / "heaptrack-print-overall.stdout.txt"
    symbol = args.filter_symbol
    hits = symbol_lines([overall_report], symbol)
    if not hits:
        raise AnalysisError(
            f"{symbol!r} did not appear in the unfiltered report; refusing filtered attribution"
        )
    filtered = run_print(
        tool,
        trace,
        folder,
        "heaptrack-print-boundedvecwriter",
        filter_symbol=symbol,
        timeout=args.timeout,
    )
    raw_trace = raw_trace_aggregate(trace, symbol)
    if raw_trace["unfiltered_requested_bytes"] != overall["histogram_parse"]["sum_size_times_count"]:
        raise AnalysisError(
            "interpreted trace + events do not reproduce unfiltered Heaptrack histogram bytes"
        )
    if raw_trace["allocation_events"] != overall["histogram_parse"]["sum_counts"]:
        raise AnalysisError(
            "interpreted trace + events do not reproduce unfiltered Heaptrack histogram count"
        )
    filtered_hist = folder / "heaptrack-print-boundedvecwriter.histogram.tsv"
    overall_hist = folder / "heaptrack-print-overall.histogram.tsv"
    hist_equal = sha256_file(filtered_hist) == sha256_file(overall_hist)
    outputs = [
        artifact(path, folder)
        for path in sorted(folder.iterdir())
        if path.name.startswith("heaptrack-print-") and path.name != "analysis.json"
    ]
    # Include the source/build binding even when the original worktree and
    # profiled binary have been deleted.  Their paths remain descriptive only.
    manifest: dict[str, Any] = {
        "status": "pass",
        "change": 419,
        "role": role,
        "created_utc": utc_now(),
        "capture": binding,
        "bindings": optional_binding(root, role, capture),
        "tool": tool_identity,
        "filter": {
            "requested": symbol,
            "matched_report_lines": hits,
            "status": "symbol_found_and_filtered_stack_report_written",
        },
        "overall": {
            "scope": "whole Heaptrack command",
            "histogram": overall["histogram"],
            "histogram_parse": overall["histogram_parse"],
            "cumulative_requested_bytes": overall["histogram_parse"]["sum_size_times_count"],
            "basis": "sum(size * count) from unfiltered Heaptrack -H allocation-info histogram",
            "interpretation": "Heaptrack allocation request events; not resident memory, output bytes, or physical copies",
        },
        "raw_trace": raw_trace,
        "filtered": {
            "report": filtered,
            "histogram": filtered["histogram"],
            "histogram_matches_overall_bytes": hist_equal,
            "cumulative_requested_bytes": None,
            "status": "not_filter_specific",
            "reason": (
                "Heaptrack v1.5.0 handleAllocation populates sizeHistogram before "
                "finalize applies filterAllocations; -H therefore remains whole-trace."
            ),
        },
        "semantics_source": {
            "version": "v1.5.0",
            "url": HEAPTRACK_SOURCE,
            "interpreted_format_url": (
                "https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/"
                "src/interpret/heaptrack_interpret.cpp"
            ),
            "filter_allocations_lines": "160-190",
            "histogram_callback_lines": "471-475",
            "finalize_order_lines": "108-115",
            "allocation_event_lines": "2624-2705",
            "direct_writer_frame_predicate": (
                "Rust mangled frame starts _RNvXs4_, contains "
                "16BoundedVecWriterNt and cross_copy_plan, ends Write5write"
            ),
            "filtered_histogram_supported": False,
        },
        "replay": {
            "requires": ["capture.json", "trace artifact", "analysis output files"],
            "does_not_require": ["profiled binary", "source worktree", "heaptrack_print executable"],
        },
        "commands": {"overall": overall["argv"], "filtered": filtered["argv"]},
        "outputs": outputs,
    }
    write_json(manifest_path, manifest)
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--role", action="append", choices=("control", "candidate"))
    parser.add_argument("--replay", action="store_true", help="verify existing analysis outputs without invoking tools")
    parser.add_argument(
        "--refresh",
        action="store_true",
        help="add raw-trace numerical aggregates to an older manifest without invoking tools",
    )
    parser.add_argument("--filter-symbol", default="BoundedVecWriter")
    parser.add_argument("--heaptrack-print", default="/usr/bin/heaptrack_print")
    parser.add_argument("--timeout", type=float, default=900.0)
    args = parser.parse_args()
    root = args.root.resolve()
    roles = args.role or tuple(role for role in ("control", "candidate") if (root / "heaptrack" / role).is_dir())
    if not roles:
        print("error: no Heaptrack role directory found", file=sys.stderr)
        return 2
    for role in roles:
        try:
            manifest = analyze_role(root, role, args)
        except AnalysisError as exc:
            folder = root / "heaptrack" / role
            failure = folder / "analysis-failure.json"
            if not failure.exists():
                try:
                    write_json(failure, {"status": "failed", "role": role, "error": str(exc), "created_utc": utc_now()})
                except OSError:
                    pass
            print(f"{role}: failed: {exc}", file=sys.stderr)
            return 1
        print(json.dumps({"role": role, "status": manifest.get("status"), "analysis": str(root / "heaptrack" / role / "analysis.json")}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
