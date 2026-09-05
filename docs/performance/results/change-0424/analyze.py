#!/usr/bin/env python3
"""Replayable Heaptrack 1.5.0 postprocessing for the 0424 source lifecycle.

The input capture and its capture.json are immutable.  For each plain and
media-rich source-backed lifecycle, the first pass runs ``heaptrack_print``
once without a filter, discovers concrete source/publication symbols, and
then writes one filtered stack report per discovered symbol.  Later
``--replay`` passes verify retained files and recompute every numerical
aggregate from the compressed trace; they do not need the profiled binary or
its worktree.

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


CHANGE = 424
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
MAX_FILTERS = 8
ROLE_RUN_ROOTS = {
    "control": "runs",
    "candidate": "candidate-profile/runs",
}
ROLE_BUILD_FILES = {
    "control": "build-control.json",
    "candidate": "candidate-profile/profile-build-candidate.json",
}
ROLE_PROTOCOL_FILES = {
    "control": "protocol.json",
    "candidate": "candidate-profile/protocol.json",
}
AUTHORIZED_PROFILE_FIELDS = frozenset({
    "role",
    "profile_role",
    "profile_source_protocol_sha256",
})

# These are source-level hypotheses, not claims that the frames must be
# present.  The unfiltered report chooses the first concrete term found for
# each family, and absent families are retained as explicit ``not_found``
# results.  Raw ancestry totals, rather than filtered -H output, are the
# numerical attribution source.
FILTER_SPECS = (
    {
        "name": "source_cross_copy_prepare",
        "terms": (
            # Rust v0 symbol names are what heaptrack_print exposes for this
            # optimized private function; keep the demangled spelling as a
            # portability fallback for a future symbolizer.
            "source_cross_copy7prepare",
            "source_cross_copy::prepare",
            "source_cross_copy::Prepared",
        ),
        "purpose": "source-backed planner preparation and staging",
    },
    {
        "name": "clone_bytes_checked",
        "terms": ("clone_bytes_checked",),
        "purpose": "checked source/media byte cloning",
    },
    {
        "name": "publish_cross_slide_copy",
        "terms": ("publish_cross_slide_copy_to_stream",),
        "purpose": "PPTX cross-slide publication entry point",
    },
    {
        "name": "opc_overlay_publication",
        "terms": (
            "write_part_overlays_shared_to_stream",
            "write_part_overlays_to_stream",
        ),
        "purpose": "OPC source overlay publication",
    },
    {
        "name": "bounded_vec_writer",
        "terms": ("BoundedVecWriter",),
        "purpose": "supplemental sink growth context",
    },
)


class AnalysisError(RuntimeError):
    pass


def role_paths(root: Path, role: str) -> tuple[Path, Path, Path]:
    try:
        run_root = root / ROLE_RUN_ROOTS[role]
        build_path = root / ROLE_BUILD_FILES[role]
        if role == "candidate" and not build_path.is_file():
            for parent in (run_root.parent, root):
                for fallback in (
                    "profile-build-candidate.json", "build-candidate.json",
                    "measurement-build-candidate.json",
                ):
                    candidate = parent / fallback
                    if candidate.is_file():
                        build_path = candidate
                        break
                if build_path.is_file():
                    break
        return run_root, build_path, root / ROLE_PROTOCOL_FILES[role]
    except KeyError as exc:
        raise AnalysisError(f"unknown profile role: {role}") from exc


def build_source(build: dict[str, Any], role: str) -> dict[str, Any]:
    source = build.get("source")
    if not isinstance(source, dict):
        source = build.get("source_after")
    if not isinstance(source, dict):
        raise AnalysisError(f"{role} build has no source binding")
    return source


def build_binary(build: dict[str, Any], role: str) -> dict[str, Any]:
    binary = build.get("binary")
    if isinstance(binary, dict):
        return binary
    binaries = build.get("binaries")
    if isinstance(binaries, dict) and isinstance(binaries.get("normal"), dict):
        return binaries["normal"]
    raise AnalysisError(f"{role} build has no normal binary binding")


def require_candidate_baseline_ancestry(
    build: dict[str, Any], source: dict[str, Any], control_revision: Any
) -> None:
    """Require a retained source-binding proof for the candidate base.

    Candidate profiling may use a descendant source revision.  The profile
    analyzer must therefore avoid the control equality check used by the
    default role.  The profile build receipt is expected to retain the base
    revision (or an explicit ancestry record) alongside its source identity;
    this keeps replay independent of a deleted worktree and avoids guessing
    ancestry from a short label.
    """
    if not isinstance(control_revision, str) or not control_revision:
        raise AnalysisError("protocol control_revision is missing")
    candidates: list[Any] = [
        build.get("baseline_revision"),
        build.get("control_revision"),
        source.get("baseline_revision"),
        source.get("control_revision"),
        source.get("ancestor_revision"),
    ]
    for container in (build.get("source_ancestry"), build.get("ancestry"), source.get("ancestry")):
        if isinstance(container, dict):
            candidates.extend(
                container.get(key)
                for key in ("baseline_revision", "control_revision", "ancestor_revision", "base_revision")
            )
            revisions = container.get("revisions")
            if isinstance(revisions, list):
                candidates.extend(revisions)
        elif isinstance(container, list):
            candidates.extend(container)
    if control_revision not in candidates:
        raise AnalysisError(
            "candidate source binding does not retain control_revision ancestry proof"
        )
    revision = source.get("revision")
    if not isinstance(revision, str) or not revision:
        raise AnalysisError("candidate source binding has no revision")


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


def validate_role_protocol(
    root: Path, protocol: dict[str, Any], role: str
) -> None:
    """Bind a candidate profile to the unchanged root protocol.

    A candidate profile may change source and binary identity, but its
    workload, scope, and all other protocol fields are inherited from the
    frozen root protocol.  The three role/hash fields are the only permitted
    additions.  Control replay keeps the historical root-protocol path and
    does not require profile metadata.
    """
    if role == "control":
        return
    frozen_path = root / "protocol.json"
    if not frozen_path.is_file():
        raise AnalysisError(f"candidate parent protocol is missing: {frozen_path}")
    parent = load_json(frozen_path)
    if not isinstance(parent, dict):
        raise AnalysisError("candidate parent protocol must be a JSON object")
    if AUTHORIZED_PROFILE_FIELDS.intersection(parent):
        raise AnalysisError("candidate parent protocol already contains profile metadata")
    expected_keys = set(parent) | AUTHORIZED_PROFILE_FIELDS
    if set(protocol) != expected_keys:
        extras = sorted(set(protocol) - expected_keys)
        missing = sorted(expected_keys - set(protocol))
        raise AnalysisError(
            "candidate protocol fields differ from parent: "
            f"extra={extras!r} missing={missing!r}"
        )
    for key, value in parent.items():
        if protocol.get(key) != value:
            raise AnalysisError(f"candidate protocol changed frozen field: {key}")
    parent_hash = sha256_file(frozen_path)
    if protocol.get("profile_source_protocol_sha256") != parent_hash:
        raise AnalysisError("candidate protocol parent hash differs from frozen root protocol")
    if protocol.get("role") != "candidate" or protocol.get("profile_role") != "candidate":
        raise AnalysisError("candidate protocol role metadata is invalid")
    if protocol.get("common_flags") != parent.get("common_flags"):
        raise AnalysisError("candidate protocol common_flags differ from parent")
    if protocol.get("scope") != parent.get("scope"):
        raise AnalysisError("candidate protocol scope differs from parent")


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


def raw_trace_symbol_hits(path: Path, term: str, *, limit: int = 20) -> list[str]:
    """Find concrete symbol strings in interpreted records.

    ``heaptrack_print`` can omit a low-volume or inlined caller from its
    capped textual report even though the interpreted trace retains its
    symbol record.  Discovery from ``s`` records keeps the filter decision
    tied to the retained trace instead of silently calling a missing frame.
    """
    if not term:
        raise AnalysisError("raw trace symbol term is empty")
    found: list[str] = []
    seen: set[str] = set()
    for raw in trace_lines(path):
        if not raw.startswith(b"s "):
            continue
        try:
            fields = raw.rstrip(b"\n").decode("utf-8").split(None, 2)
        except UnicodeDecodeError as exc:
            raise AnalysisError("trace symbol record is not UTF-8") from exc
        if len(fields) == 3 and term in fields[2] and fields[2] not in seen:
            seen.add(fields[2])
            found.append(fields[2])
            if len(found) >= limit:
                break
    return found


def hex_token(token: str, label: str) -> int:
    try:
        return int(token, 16)
    except ValueError as exc:
        raise AnalysisError(f"invalid hexadecimal {label}: {token!r}") from exc


def raw_trace_aggregate(
    path: Path, symbol: str | tuple[str, ...] | None
) -> dict[str, Any]:
    """Aggregate + events by interpreted trace ancestry.

    Heaptrack v1.5.0 writes ``a size trace`` allocation-info records and ``+
    allocation-info-index`` events.  Summing sizes for + events reproduces
    the unfiltered -H histogram.  This parser deliberately does not inspect
    the binary or resolve addresses; all function names needed for the
    filter are already present in the interpreted trace.
    """
    if symbol is None:
        terms: tuple[str, ...] = ()
    else:
        terms = (symbol,) if isinstance(symbol, str) else tuple(symbol)
        if not terms or any(not isinstance(term, str) or not term for term in terms):
            raise AnalysisError("raw trace filter terms must be non-empty strings")
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
    phase_totals: dict[str, dict[str, int]] = {
        phase: {
            "allocation_events": 0,
            "deallocation_events": 0,
            "requested_bytes": 0,
            "deallocated_bytes": 0,
        }
        for phase in ("lifecycle", "corpus_setup", "correctness_gates", "other")
    }

    def phase_for_names(names: tuple[str, ...]) -> str:
        # The source-backed workload intentionally includes corpus generation,
        # the measured lifecycle, and refusal/correctness gates in one
        # process.  These labels are ancestry context only; they do not turn
        # Heaptrack's whole-command totals into operation-local measurements.
        if any("run_pptx_source_backed_cross_copy_lifecycle" in name for name in names):
            return "lifecycle"
        if any(
            token in name
            for name in names
            for token in (
                "build_pptx_source_backed_cross_copy_corpus",
                "build_pptx_cross_copy_corpus",
                "pptx_cross_copy_bytes",
            )
        ):
            return "corpus_setup"
        if any(
            token in name
            for name in names
            for token in (
                "verify_pptx_source_backed_cross_copy_lifecycle_gates",
                "verify_pptx_cross_copy_refusal_gates",
                "source_backed_cross_copy_revision_refusal",
                "source_backed_cross_copy_foreign_destination_refusal",
            )
        ):
            return "correctness_gates"
        return "other"

    def direct_writer_frame(name: str) -> bool:
        # This is the concrete Rust Write::write implementation in the
        # current trace, rather than generic Counted/Chunked callers that
        # merely carry BoundedVecWriter in a type parameter.
        return (
            terms == ("BoundedVecWriter",)
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
        matching = (
            tuple(name for name in names if any(term in name for term in terms))
            if terms
            else ()
        )
        if matching:
            matched_events += count
            matched_deallocation_events += deallocation_count
            matched_bytes += bytes_requested
            matched_deallocated_bytes += bytes_deallocated
            add_row(
                rows_by_trace, trace_id, matching, count, deallocation_count,
                bytes_requested, bytes_deallocated,
            )
            phase = phase_for_names(names)
            phase_row = phase_totals[phase]
            phase_row["allocation_events"] += count
            phase_row["deallocation_events"] += deallocation_count
            phase_row["requested_bytes"] += bytes_requested
            phase_row["deallocated_bytes"] += bytes_deallocated
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
    lifecycle_per_trace = [
        row
        for row in per_trace
        if any(
            "run_pptx_source_backed_cross_copy_lifecycle" in frame
            for frame in row["trace_frames"]
        )
    ]
    lifecycle_bounded = lifecycle_per_trace[:MAX_PER_TRACE_ROWS]
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
        "filter_terms": list(terms),
        "matched_events": matched_events,
        "matched_deallocation_events": matched_deallocation_events,
        "matched_requested_bytes": matched_bytes,
        "matched_deallocated_bytes": matched_deallocated_bytes,
        "matched_net_live_delta_bytes": matched_bytes - matched_deallocated_bytes,
        "matched_share_of_all_requested_percent": (
            matched_bytes * 100.0 / total_bytes if total_bytes else None
        ),
        "matched_phase_totals": phase_totals,
        "lifecycle_trace_count": len(lifecycle_per_trace),
        "lifecycle_trace_rows_retained": len(lifecycle_bounded),
        "lifecycle_trace_rows_truncated": len(lifecycle_per_trace) > len(lifecycle_bounded),
        "lifecycle_trace_receipt": {
            "predicate": "trace ancestry contains run_pptx_source_backed_cross_copy_lifecycle",
            "rows": lifecycle_bounded,
        },
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


def decorate_filter_raw(name: str, raw: dict[str, Any]) -> dict[str, Any]:
    """Add bounded, trace-derived receipts for the production clone stage.

    The regular per-trace table is capped at ``MAX_PER_TRACE_ROWS``.  The
    clone family is small enough to retain every lifecycle row, so keep an
    explicit receipt for allocations at least 1 MiB and classify the two
    current call-site stages from their actual ancestry.  This is derived
    solely from the interpreted trace and is deliberately recomputed during
    replay; it does not turn the whole-command trace into an operation-local
    measurement.
    """
    if name != "clone_bytes_checked":
        return raw
    receipt = raw.get("lifecycle_trace_receipt")
    if not isinstance(receipt, dict) or not isinstance(receipt.get("rows"), list):
        raise AnalysisError("clone trace is missing lifecycle receipt")

    def stage_for(row: dict[str, Any]) -> str:
        frames = row.get("trace_frames")
        if not isinstance(frames, list) or not all(isinstance(frame, str) for frame in frames):
            raise AnalysisError("clone lifecycle receipt has malformed trace frames")
        if any("publish_cross_slide_copy_to_stream" in frame for frame in frames):
            return "publish_cross_slide_copy_to_stream"
        if any("plan_cross_slide_copy" in frame for frame in frames):
            return "plan_cross_slide_copy"
        return "other_lifecycle_ancestry"

    large_rows: list[dict[str, Any]] = []
    for original in receipt["rows"]:
        if not isinstance(original, dict):
            raise AnalysisError("clone lifecycle receipt has malformed row")
        requested = original.get("requested_bytes")
        if not isinstance(requested, int) or requested < 0:
            raise AnalysisError("clone lifecycle receipt has invalid requested bytes")
        if requested < 1024 * 1024:
            continue
        row = dict(original)
        row["stage"] = stage_for(original)
        large_rows.append(row)

    by_stage: dict[str, dict[str, int]] = {}
    for row in large_rows:
        stage = row["stage"]
        stage_row = by_stage.setdefault(
            stage,
            {"trace_count": 0, "allocation_events": 0, "requested_bytes": 0, "deallocated_bytes": 0},
        )
        stage_row["trace_count"] += 1
        stage_row["allocation_events"] += row["allocation_events"]
        stage_row["requested_bytes"] += row["requested_bytes"]
        stage_row["deallocated_bytes"] += row["deallocated_bytes"]

    raw["lifecycle_large_allocation_receipt"] = {
        "predicate": (
            "clone_bytes_checked ancestry plus lifecycle runner and "
            "requested_bytes >= 1048576"
        ),
        "threshold_bytes": 1024 * 1024,
        "complete": not raw["lifecycle_trace_rows_truncated"],
        "trace_count": len(large_rows),
        "allocation_events": sum(row["allocation_events"] for row in large_rows),
        "requested_bytes": sum(row["requested_bytes"] for row in large_rows),
        "deallocated_bytes": sum(row["deallocated_bytes"] for row in large_rows),
        "by_stage": by_stage,
        "rows": large_rows,
    }
    return raw


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


def root_artifact(path: Path, root: Path) -> dict[str, Any]:
    """Hash a retained artifact with a portable path relative to the bundle."""
    if not path.is_file() or path.is_symlink():
        raise AnalysisError(f"retained artifact is not a regular file: {path}")
    try:
        relative = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError as exc:
        raise AnalysisError(f"retained artifact is outside bundle: {path}") from exc
    return {
        "path": relative,
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
    }


def capture_file_hash(capture: dict[str, Any], name: str) -> str:
    for item in capture.get("files", []):
        if isinstance(item, dict) and item.get("name") == name:
            value = item.get("sha256")
            if isinstance(value, str):
                return value
    raise AnalysisError(f"capture inventory has no {name!r}")


def expected_selector(corpus: str) -> str:
    return f"pptx_source_backed_cross_copy_{corpus}_lifecycle"


def validate_report_catalog(
    folder: Path,
    capture: dict[str, Any],
    protocol: dict[str, Any],
    build: dict[str, Any],
    corpus: str,
    role: str,
) -> dict[str, Any]:
    selector = expected_selector(corpus)
    report_path = folder / "report.json"
    catalog_path = folder / "catalog.json"
    if sha256_file(report_path) != capture_file_hash(capture, "report.json"):
        raise AnalysisError(f"{corpus}: report.json is not the captured artifact")
    if sha256_file(catalog_path) != capture_file_hash(capture, "catalog.json"):
        raise AnalysisError(f"{corpus}: catalog.json is not the captured artifact")
    report = load_json(report_path)
    catalog = load_json(catalog_path)
    if not isinstance(report, dict) or not isinstance(catalog, dict):
        raise AnalysisError(f"{corpus}: report/catalog must be objects")
    binary = build_binary(build, role)
    binary_hash = binary.get("binary_sha256", binary.get("sha256"))
    report_binary = report.get("binary_identity", {})
    if report_binary.get("binary_sha256") != binary_hash:
        raise AnalysisError(f"{corpus}: report binary hash differs from build identity")
    source = build_source(build, role)
    environment = report.get("environment", {})
    if environment.get("git_revision") != source.get("revision"):
        raise AnalysisError(f"{corpus}: report revision differs from build identity")
    if environment.get("git_worktree_dirty") is not False:
        raise AnalysisError(f"{corpus}: report source is not clean")
    if environment.get("cpu_affinity") != str(protocol.get("cpu")):
        raise AnalysisError(f"{corpus}: report CPU affinity is not protocol CPU2")
    configuration = report.get("configuration", {})
    if configuration.get("cases") != [selector]:
        raise AnalysisError(f"{corpus}: report case identity differs from capture")
    if configuration.get("samples_per_case") != protocol.get("samples"):
        raise AnalysisError(f"{corpus}: report sample count differs from protocol")
    if configuration.get("warmup_iterations_per_case") != protocol.get("warmups"):
        raise AnalysisError(f"{corpus}: report warmup count differs from protocol")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1 or results[0].get("case") != selector:
        raise AnalysisError(f"{corpus}: report result identity is invalid")
    if catalog.get("catalog_sha256") != report.get("corpus_catalog", {}).get("catalog_sha256"):
        raise AnalysisError(f"{corpus}: report/catalog catalog hash mismatch")
    bindings = catalog.get("case_bindings")
    if not isinstance(bindings, list) or len(bindings) != 1 or bindings[0].get("case") != selector:
        raise AnalysisError(f"{corpus}: catalog case binding is invalid")
    return {
        "report": root_artifact(report_path, folder.parent.parent),
        "catalog": root_artifact(catalog_path, folder.parent.parent),
        "report_binary_sha256": report_binary.get("binary_sha256"),
        "report_git_revision": environment.get("git_revision"),
        "report_result": {
            "case": results[0].get("case"),
            "output_sha256": results[0].get("output_sha256"),
            "elapsed_ns": results[0].get("elapsed_ns"),
        },
        "catalog_sha256": catalog.get("catalog_sha256"),
        "content_set_sha256": catalog.get("content_set_sha256"),
    }


def binding_for_capture(
    root: Path,
    folder: Path,
    capture: dict[str, Any],
    protocol: dict[str, Any],
    build: dict[str, Any],
    corpus: str,
    role: str,
    build_path: Path,
    protocol_path: Path,
) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    trace, capture_binding = capture_trace(folder, capture)
    if capture.get("protocol_sha256") != sha256_file(protocol_path):
        raise AnalysisError(f"{corpus}: protocol hash differs from capture binding")
    if capture.get("build_sha256") != sha256_file(build_path):
        raise AnalysisError(f"{corpus}: build hash differs from capture binding")
    if capture.get("corpus") != corpus or capture.get("selector") != expected_selector(corpus):
        raise AnalysisError(f"{corpus}: capture selector identity is invalid")
    validate_role_protocol(root, protocol, role)
    if protocol.get("change") != CHANGE:
        raise AnalysisError("0424 protocol/build revision binding is invalid")
    source_binding = build_source(build, role)
    if role == "control":
        if protocol.get("control_revision") != source_binding.get("revision"):
            raise AnalysisError("control protocol/build revision binding is invalid")
    else:
        require_candidate_baseline_ancestry(
            build, source_binding, protocol.get("control_revision")
        )
    if capture.get("verify_exit_code") != 0:
        raise AnalysisError(f"{corpus}: captured report verifier did not pass")
    source = capture.get("source", {})
    if source.get("revision") != source_binding.get("revision") or source.get("clean") is not True:
        raise AnalysisError(f"{corpus}: capture source identity differs from build")
    report_catalog = validate_report_catalog(folder, capture, protocol, build, corpus, role)
    binary = build_binary(build, role)
    parser_path = root.parent / "change-0419" / "analyze-heaptrack.py"
    parser_binding: dict[str, Any] = {
        "semantic": "copied change-0419 interpreted Heaptrack v3 raw_trace_aggregate",
        "path": "../change-0419/analyze-heaptrack.py",
    }
    if parser_path.is_file():
        parser_binding["sha256"] = sha256_file(parser_path)
    else:
        parser_binding["present_at_analysis"] = False
    binding = {
        **capture_binding,
        "protocol": root_artifact(protocol_path, root),
        "build": root_artifact(build_path, root),
        "report_catalog": report_catalog,
        "parser": parser_binding,
        "source_revision": source_binding.get("revision"),
        "source_identity_sha256": capture.get("source", {}).get("identity_sha256"),
        "binary": {
            "label": binary.get("label"),
            "sha256": binary.get("binary_sha256", binary.get("sha256")),
            "bytes": binary.get("binary_bytes", binary.get("bytes")),
            "path_at_capture": binary.get("path"),
        },
    }
    return trace, binding, report_catalog


def verify_binding_files(root: Path, manifest: dict[str, Any], role: str) -> None:
    for key in ("protocol", "build"):
        expected = manifest.get("capture", {}).get(key)
        if not isinstance(expected, dict) or not isinstance(expected.get("path"), str):
            raise AnalysisError(f"analysis binding {key} is malformed")
        path = root / expected["path"]
        if root_artifact(path, root) != expected:
            raise AnalysisError(f"analysis binding changed: {path}")
    report_catalog = manifest.get("capture", {}).get("report_catalog", {})
    if role == "control":
        report_root = root
    elif role == "candidate":
        # Existing candidate manifests were generated with report/catalog
        # paths relative to candidate-profile, while protocol/build/parser
        # bindings remain relative to the bundle root.
        report_root = root / "candidate-profile"
    else:
        raise AnalysisError(f"unknown profile role: {role}")
    for key in ("report", "catalog"):
        expected = report_catalog.get(key)
        if not isinstance(expected, dict) or not isinstance(expected.get("path"), str):
            raise AnalysisError(f"analysis binding {key} is malformed")
        path = report_root / expected["path"]
        if root_artifact(path, report_root) != expected:
            raise AnalysisError(f"analysis binding changed: {path}")


def verify_outputs(folder: Path, outputs: Any) -> None:
    if not isinstance(outputs, list) or not outputs:
        raise AnalysisError("analysis output inventory is missing")
    for expected in outputs:
        if not isinstance(expected, dict) or not isinstance(expected.get("path"), str):
            raise AnalysisError("malformed analysis output inventory")
        path = folder / expected["path"]
        if artifact(path, folder) != expected:
            raise AnalysisError(f"analysis output changed: {path}")


def verify_replay(root: Path, corpus: str, role: str) -> dict[str, Any]:
    runs_root, build_path, protocol_path = role_paths(root, role)
    folder = runs_root / corpus
    manifest_path = folder / "analysis.json"
    manifest = load_json(manifest_path)
    if not isinstance(manifest, dict) or manifest.get("status") != "pass":
        raise AnalysisError(f"analysis manifest is not successful: {manifest_path}")
    protocol = load_json(protocol_path)
    build = load_json(build_path)
    capture = load_json(folder / "capture.json")
    trace, observed, _ = binding_for_capture(
        root, folder, capture, protocol, build, corpus, role, build_path, protocol_path
    )
    if manifest.get("capture") != observed:
        raise AnalysisError(f"{corpus}: capture binding changed since analysis")
    verify_binding_files(root, manifest, role)
    verify_outputs(folder, manifest.get("outputs"))
    overall = manifest.get("overall", {})
    raw_overall = raw_trace_aggregate(trace, None)
    if raw_overall != overall.get("raw_trace"):
        raise AnalysisError(f"{corpus}: raw whole-trace aggregate changed")
    histogram_parse = overall.get("histogram_parse", {})
    if raw_overall["unfiltered_requested_bytes"] != histogram_parse.get("sum_size_times_count"):
        raise AnalysisError(f"{corpus}: raw bytes no longer match unfiltered histogram")
    if raw_overall["allocation_events"] != histogram_parse.get("sum_counts"):
        raise AnalysisError(f"{corpus}: raw events no longer match unfiltered histogram")
    filters = manifest.get("filters")
    if not isinstance(filters, list):
        raise AnalysisError(f"{corpus}: filter manifest is missing")
    for item in filters:
        if not isinstance(item, dict):
            raise AnalysisError(f"{corpus}: malformed filter manifest")
        resolved = item.get("resolved_term")
        if item.get("status") == "matched":
            observed_raw = decorate_filter_raw(
                item["name"], raw_trace_aggregate(trace, resolved)
            )
            if observed_raw != item.get("raw_trace"):
                raise AnalysisError(f"{corpus}: raw filtered aggregate changed for {item.get('name')}")
        elif item.get("status") != "not_found":
            raise AnalysisError(f"{corpus}: unknown filter status")
    return manifest


def resolve_filter_specs(report_path: Path, trace: Path) -> list[dict[str, Any]]:
    resolved: list[dict[str, Any]] = []
    for spec in FILTER_SPECS:
        chosen = None
        hits: list[str] = []
        discovery_source = None
        trace_symbols: list[str] = []
        for term in spec["terms"]:
            candidate_hits = symbol_lines([report_path], term)
            if candidate_hits:
                chosen = term
                hits = candidate_hits
                discovery_source = "heaptrack_print_unfiltered_report"
                break
            candidate_symbols = raw_trace_symbol_hits(trace, term)
            if candidate_symbols:
                chosen = term
                trace_symbols = candidate_symbols
                discovery_source = "interpreted_trace_symbol_records"
                break
        if chosen is None:
            resolved.append({
                "name": spec["name"],
                "purpose": spec["purpose"],
                "requested_terms": list(spec["terms"]),
                "status": "not_found",
                "resolved_term": None,
                "matched_report_lines": [],
                "matched_trace_symbols": [],
                "discovery_source": None,
            })
        else:
            resolved.append({
                "name": spec["name"],
                "purpose": spec["purpose"],
                "requested_terms": list(spec["terms"]),
                "status": "matched",
                "resolved_term": chosen,
                "matched_report_lines": hits,
                "matched_trace_symbols": trace_symbols,
                "discovery_source": discovery_source,
            })
    if len(resolved) > MAX_FILTERS:
        raise AnalysisError("filter specification exceeds bounded analysis limit")
    return resolved


def analyze_corpus(root: Path, corpus: str, args: argparse.Namespace) -> dict[str, Any]:
    runs_root, build_path, protocol_path = role_paths(root, args.role)
    folder = runs_root / corpus
    if not folder.is_dir():
        raise AnalysisError(f"missing capture directory: {folder}")
    manifest_path = folder / "analysis.json"
    if args.replay or manifest_path.exists():
        return verify_replay(root, corpus, args.role)
    protocol = load_json(protocol_path)
    build = load_json(build_path)
    capture = load_json(folder / "capture.json")
    if not isinstance(protocol, dict) or not isinstance(build, dict) or not isinstance(capture, dict):
        raise AnalysisError(f"{corpus}: protocol/build/capture JSON must be objects")
    trace, binding, report_catalog = binding_for_capture(
        root, folder, capture, protocol, build, corpus, args.role, build_path, protocol_path
    )
    tool_text = shutil.which(args.heaptrack_print)
    tool = Path(tool_text or args.heaptrack_print)
    if not tool.is_file() or not os.access(tool, os.X_OK):
        raise AnalysisError(f"heaptrack_print is unavailable: {args.heaptrack_print}")
    try:
        version = subprocess.run(
            [str(tool), "--version"], capture_output=True, text=True,
            check=False, timeout=30,
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
        tool, trace, folder, "heaptrack-print-overall",
        filter_symbol=None, timeout=args.timeout,
    )
    raw_overall = raw_trace_aggregate(trace, None)
    histogram_parse = overall["histogram_parse"]
    if raw_overall["unfiltered_requested_bytes"] != histogram_parse["sum_size_times_count"]:
        raise AnalysisError(f"{corpus}: raw trace does not reproduce unfiltered histogram bytes")
    if raw_overall["allocation_events"] != histogram_parse["sum_counts"]:
        raise AnalysisError(f"{corpus}: raw trace does not reproduce unfiltered histogram events")
    filters = resolve_filter_specs(
        folder / "heaptrack-print-overall.stdout.txt", trace
    )
    for item in filters:
        if item["status"] != "matched":
            item["raw_trace"] = None
            item["print"] = None
            continue
        stem = "heaptrack-print-" + item["name"]
        filtered = run_print(
            tool, trace, folder, stem,
            filter_symbol=item["resolved_term"], timeout=args.timeout,
        )
        filtered_hist = folder / f"{stem}.histogram.tsv"
        overall_hist = folder / "heaptrack-print-overall.histogram.tsv"
        item["print"] = filtered
        item["histogram_matches_overall"] = sha256_file(filtered_hist) == sha256_file(overall_hist)
        item["histogram_scope"] = "whole_trace_not_filter_specific"
        item["raw_trace"] = decorate_filter_raw(
            item["name"], raw_trace_aggregate(trace, item["resolved_term"])
        )
    outputs = [
        artifact(path, folder)
        for path in sorted(folder.iterdir())
        if path.name.startswith("heaptrack-print-")
    ]
    manifest: dict[str, Any] = {
        "status": "pass",
        "change": CHANGE,
        "corpus": corpus,
        "selector": expected_selector(corpus),
        "created_utc": utc_now(),
        "capture": binding,
        "report_catalog": report_catalog,
        "tool": tool_identity,
        "overall": {
            "scope": protocol["scope"],
            "histogram": overall["histogram"],
            "histogram_parse": histogram_parse,
            "raw_trace": raw_overall,
            "basis": "sum(size * count) from unfiltered Heaptrack -H allocation-info histogram",
            "interpretation": "whole-command allocation request events; not resident memory, output bytes, or physical copies",
        },
        "filters": filters,
        "semantics_source": {
            "version": "v1.5.0",
            "url": HEAPTRACK_SOURCE,
            "interpreted_format_url": (
                "https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/"
                "src/interpret/heaptrack_interpret.cpp"
            ),
            "filtered_histogram_supported": False,
            "reason": (
                "Heaptrack v1.5.0 builds its size histogram while reading allocation events; "
                "filtering is applied later, so filtered -H output remains whole-trace."
            ),
        },
        "claims": {
            "performance_claim": None,
            "operation_local_heaptrack_claim": None,
            "allowed": ["source-level stack attribution within this whole-command trace"],
        },
        "replay": {
            "requires": [
                "capture.json", "trace.zst", "analysis output files", "protocol.json",
                ROLE_BUILD_FILES[args.role],
            ],
            "does_not_require": ["profiled binary", "source worktree", "heaptrack_print executable"],
            "recomputes": ["unfiltered raw allocation aggregate", "each matched filtered ancestry aggregate"],
        },
        "outputs": outputs,
    }
    write_json(manifest_path, manifest)
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=DEFAULT_ROOT)
    parser.add_argument("--role", choices=("control", "candidate"), default="control")
    parser.add_argument("--corpus", action="append", choices=("plain", "media_rich"))
    parser.add_argument("--replay", action="store_true", help="recompute retained trace aggregates without heaptrack_print")
    parser.add_argument("--heaptrack-print", default="/usr/bin/heaptrack_print")
    parser.add_argument("--timeout", type=float, default=900.0)
    args = parser.parse_args()
    root = args.root.resolve()
    corpora = tuple(args.corpus or ("plain", "media_rich"))
    for corpus in corpora:
        try:
            manifest = analyze_corpus(root, corpus, args)
        except AnalysisError as exc:
            runs_root, _, _ = role_paths(root, args.role)
            folder = runs_root / corpus
            failure = folder / "analysis-failure.json"
            if not failure.exists():
                try:
                    write_json(failure, {
                        "status": "failed", "change": CHANGE, "corpus": corpus,
                        "error": str(exc), "created_utc": utc_now(),
                    })
                except OSError:
                    pass
            print(f"{corpus}: failed: {exc}", file=sys.stderr)
            return 1
        print(json.dumps({
            "corpus": corpus,
            "status": manifest.get("status"),
            "analysis": str(role_paths(root, args.role)[0] / corpus / "analysis.json"),
        }))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
