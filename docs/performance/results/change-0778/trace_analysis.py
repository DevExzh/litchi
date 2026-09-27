#!/usr/bin/env python3
"""Validate one 0778 ordinary-save publication trace.

The capture program runs a separate ``strace`` child for each corpus and
durability policy.  A child contains four setup saves, their cleanup, the
measured fifth save, a readback, and final cleanup.  This module attributes
only the fifth sibling-temporary publication to the measured transaction.

The parser is intentionally independent of the capture program.  It does not
read the filesystem, infer a path from a missing descriptor annotation, or
turn ``strace`` time into native save latency.  It binds the byte count and
SHA-256 digest to the JSON report emitted by the save harness and fails closed
when a path, descriptor, transaction boundary, or output identity is
ambiguous.
"""

from __future__ import annotations

import ast
from decimal import Decimal, InvalidOperation
import hashlib
import json
import os
from pathlib import Path
import re
from typing import Any, Mapping, Sequence


TRACE_OPTIONS = [
    "-ttt",
    "-T",
    "-yy",
    "-s",
    "256",
    "-e",
    "trace=%file,read,pread64,write,pwrite64,writev,lseek,close,fsync,fdatasync,fchmod",
]

SETUP_DESTINATION_NAMES = ("published", "published", "published-repeat", "published")
POLICIES = {"default", "full", "file-only", "no-sync"}
FORMAT_EXTENSIONS = {"docx": ".docx", "xlsx": ".xlsx", "pptx": ".pptx"}
TEMP_RE = re.compile(r"^\.litchi-[A-Za-z0-9_-]{6,128}\.tmp$")
QUOTED_RE = re.compile(r'"(?:\\.|[^"\\])*"')
ANNOTATED_FD_RE = re.compile(r"(?<![A-Za-z0-9_])(-?\d+)<([^>]*)>")
SYSCALL_RE = re.compile(
    r"^(?:(?P<pid>\d+)\s+)?(?P<timestamp>\d+\.\d+)\s+"
    r"(?P<call>[A-Za-z_][A-Za-z0-9_]*)\((?P<body>.*)\)\s+=\s+"
    r"(?P<result>.*?)\s+<(?P<duration>\d+(?:\.\d+)?)>\s*$"
)
EXIT_RE = re.compile(r"^(?:(?:\d+\s+)?\d+\.\d+\s+)?\+\+\+ exited with (?P<code>-?\d+) \+\+\+$")

OPEN_CALLS = {"open", "openat", "openat2", "creat"}
RENAME_CALLS = {"rename", "renameat", "renameat2"}
UNLINK_CALLS = {"unlink", "unlinkat", "remove"}
METADATA_CALLS = {
    "fstatat",
    "fstatat64",
    "lstat",
    "lstat64",
    "newfstatat",
    "stat",
    "stat64",
    "statx",
}
WRITE_CALLS = {"write", "pwrite", "pwrite64", "writev", "pwritev", "pwritev2"}
READ_CALLS = {"read", "pread", "pread64", "readv", "preadv", "preadv2"}
SYNC_CALLS = {"fsync", "fdatasync"}


class TraceError(AssertionError):
    """Raised when a trace does not prove the frozen publication contract."""


def _fail(message: str) -> None:
    raise TraceError(message)


def _require(condition: bool, message: str) -> None:
    if not condition:
        _fail(message)


class Event:
    def __init__(
        self,
        index: int,
        line_no: int,
        raw: str,
        timestamp_ns: int,
        call: str,
        body: str,
        result: str,
        result_code: int | None,
        duration_ns: int,
        pid: int | None,
    ) -> None:
        self.index = index
        self.line_no = line_no
        self.raw = raw
        self.timestamp_ns = timestamp_ns
        self.call = call
        self.body = body
        self.result = result
        self.result_code = result_code
        self.duration_ns = duration_ns
        self.pid = pid
        self.fd: int | None = None
        self.path: str | None = None
        self.open_path: str | None = None


class Candidate:
    def __init__(self, event: Event, source: str, destination: str) -> None:
        self.event = event
        self.source = source
        self.destination = destination


def _decimal_ns(value: str, label: str) -> int:
    try:
        decimal = Decimal(value)
    except InvalidOperation:
        _fail(f"{label} is not a decimal value: {value!r}")
    _require(decimal >= 0, f"{label} is negative")
    return int(decimal * Decimal(1_000_000_000))


def _result_code(result: str) -> int | None:
    match = re.match(r"\s*(-?\d+)(?:\s|<|$)", result)
    return int(match.group(1)) if match else None


def _result_fd(result: str) -> int | None:
    match = re.match(r"\s*(-?\d+)(?:<[^>]*>)?(?:\s|$)", result)
    if not match:
        return None
    value = int(match.group(1))
    return value if value >= 0 else None


def _result_annotation(result: str) -> str | None:
    match = re.search(r"(?:^|\s)-?\d+<([^>]*)>", result)
    return _norm_path(match.group(1)) if match else None


def _quoted_strings(body: str) -> list[str]:
    values: list[str] = []
    for token in QUOTED_RE.findall(body):
        try:
            value = ast.literal_eval(token)
        except (SyntaxError, ValueError) as error:
            _fail(f"unable to decode strace string {token!r}: {error}")
        _require(isinstance(value, str), f"decoded strace value is not a string: {token!r}")
        values.append(value)
    return values


def _norm_path(value: str) -> str:
    # ``-yy`` also annotates pipes, sockets, and anon-inodes.  They are kept
    # as opaque descriptor identities so unrelated stdio traffic does not
    # invalidate a trace; every descriptor used by the publication contract
    # is checked against an absolute filesystem path later.
    _require(value, "empty descriptor annotation")
    return os.path.normpath(value)


def _first_fd(body: str) -> int | None:
    match = re.match(r"\s*(-?\d+)(?:<[^>]*>)?(?:\s*,|\s*$)", body)
    return int(match.group(1)) if match else None


def _argument_annotation(body: str, fd: int | None) -> str | None:
    if fd is None:
        return None
    for match in ANNOTATED_FD_RE.finditer(body):
        if int(match.group(1)) == fd:
            return _norm_path(match.group(2))
    return None


def _open_path(event: Event) -> str | None:
    if event.call not in OPEN_CALLS or event.result_code is None or event.result_code < 0:
        return None
    values = _quoted_strings(event.body)
    if values:
        # openat/openat2 have the path after the dirfd.  Current harness paths
        # are the only quoted argument; keeping the first value also handles
        # open/openat flags without making a platform-specific flags parser.
        return _norm_path(values[0])
    return _result_annotation(event.result)


def _successful(event: Event) -> bool:
    return event.result_code == 0


def _parse_trace(path: Path) -> tuple[list[Event], dict[str, Any]]:
    _require(path.is_file() and not path.is_symlink(), f"missing or symlinked trace: {path}")
    data = path.read_bytes()
    digest = hashlib.sha256(data).hexdigest()
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        _fail(f"trace is not UTF-8: {error}")

    events: list[Event] = []
    exit_codes: list[int] = []
    ignored_lines = 0
    pids: set[int] = set()
    for line_no, raw_line in enumerate(text.splitlines(), 1):
        raw = raw_line.rstrip("\r")
        if not raw.strip():
            ignored_lines += 1
            continue
        if "<unfinished ...>" in raw or ("<..." in raw and "resumed>" in raw):
            _fail(f"trace contains a non-serial syscall line {line_no}: {raw}")
        exit_match = EXIT_RE.match(raw)
        if exit_match:
            exit_codes.append(int(exit_match.group("code")))
            continue
        if (
            raw.lstrip().startswith(("--- ", "+++ ", "strace:"))
            or re.match(r"^(?:\d+\s+)?\d+\.\d+\s+--- ", raw)
            or re.match(r"^(?:\d+\s+)?\d+\.\d+\s+\+\+\+ ", raw)
        ):
            ignored_lines += 1
            continue
        match = SYSCALL_RE.match(raw)
        if not match:
            if re.match(r"^(?:\d+\s+)?\d+\.\d+\s+", raw):
                _fail(f"unparsed timestamped trace line {line_no}: {raw}")
            ignored_lines += 1
            continue
        pid = int(match.group("pid")) if match.group("pid") else None
        if pid is not None:
            pids.add(pid)
        event = Event(
            index=len(events),
            line_no=line_no,
            raw=raw,
            timestamp_ns=_decimal_ns(match.group("timestamp"), f"line {line_no} timestamp"),
            call=match.group("call"),
            body=match.group("body"),
            result=match.group("result"),
            result_code=_result_code(match.group("result")),
            duration_ns=_decimal_ns(match.group("duration"), f"line {line_no} duration"),
            pid=pid,
        )
        event.fd = _first_fd(event.body)
        event.open_path = _open_path(event)
        event.path = event.open_path or _result_annotation(event.result)
        events.append(event)

    _require(exit_codes == [0], f"trace exit markers are not exactly [0]: {exit_codes!r}")
    _require(events, "trace has no completed syscall lines")
    _require(len(pids) <= 1, f"trace contains multiple traced processes: {sorted(pids)!r}")

    # -yy annotates most fd arguments.  The table is a conservative fallback
    # for a platform representation that annotates only the open result.
    fd_paths: dict[int, str] = {}
    for event in events:
        if event.call in OPEN_CALLS and event.result_code is not None and event.result_code >= 0:
            opened_fd = _result_fd(event.result)
            if opened_fd is not None and event.open_path is not None:
                fd_paths[opened_fd] = event.open_path
                event.fd = opened_fd
                event.path = event.open_path
        elif event.fd is not None:
            argument_path = _argument_annotation(event.body, event.fd)
            event.path = argument_path or event.path or fd_paths.get(event.fd)
        if event.call == "close" and event.result_code == 0 and event.fd is not None:
            fd_paths.pop(event.fd, None)

    return events, {
        "sha256": digest,
        "bytes": len(data),
        "lines": len(text.splitlines()),
        "ignored_lines": ignored_lines,
        "exit_codes": exit_codes,
        "pids": sorted(pids),
    }


def _event_paths(event: Event) -> list[str]:
    # ``statx(fd, "", AT_EMPTY_PATH, ...)`` is a normal readback probe, not a
    # pathname.  Ignore that empty string while retaining every real path for
    # exact transaction matching.
    return [_norm_path(value) for value in _quoted_strings(event.body) if value]


def _rename_paths(event: Event) -> tuple[str, str] | None:
    if event.call not in RENAME_CALLS or not _successful(event):
        return None
    paths = _event_paths(event)
    if len(paths) < 2:
        _fail(f"successful {event.call} has fewer than two path arguments: {event.raw}")
    # renameat and renameat2 have dirfd/path pairs; rename has only paths.
    return paths[-2], paths[-1]


def _path_argument(event: Event, expected: str) -> bool:
    return expected in _event_paths(event)


def _is_temp_path(value: str, parent: str) -> bool:
    return os.path.dirname(value) == parent and bool(TEMP_RE.fullmatch(os.path.basename(value)))


def _is_destination_path(value: str, parent: str, names: Sequence[str]) -> bool:
    return os.path.dirname(value) == parent and os.path.basename(value) in names


def _artifact(path: Path) -> dict[str, Any]:
    _require(path.is_file() and not path.is_symlink(), f"missing or symlinked report: {path}")
    data = path.read_bytes()
    return {"path": str(path), "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}


def _load_report(report: Path | Mapping[str, Any]) -> tuple[Mapping[str, Any], dict[str, Any] | None]:
    if isinstance(report, Mapping):
        return report, None
    report_path = Path(report)
    artifact = _artifact(report_path)
    try:
        value = json.loads(report_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        _fail(f"unable to read JSON report {report_path}: {error}")
    _require(isinstance(value, Mapping), "harness report is not a JSON object")
    return value, artifact


def _walk(value: Any, path: tuple[str, ...] = ()):
    if isinstance(value, Mapping):
        yield path, value
        for key, child in value.items():
            if isinstance(key, str):
                yield from _walk(child, path + (key,))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from _walk(child, path + (str(index),))


def _output_identity(report: Mapping[str, Any]) -> dict[str, Any]:
    byte_keys = ("published_bytes", "output_bytes")
    hash_keys = ("published_sha256", "output_sha256")
    pairs: dict[tuple[int, str], list[list[str]]] = {}
    related_hashes: list[tuple[str, str]] = []
    for path, value in _walk(report):
        byte_values = []
        for key in byte_keys:
            if key not in value:
                continue
            raw = value[key] if isinstance(value[key], list) else [value[key]]
            for item in raw:
                if isinstance(item, int) and not isinstance(item, bool):
                    byte_values.append((key, item))
        hash_values = []
        for key in hash_keys:
            if key not in value:
                continue
            raw = value[key] if isinstance(value[key], list) else [value[key]]
            for item in raw:
                if isinstance(item, str):
                    hash_values.append((key, item))
        for byte_key, byte_count in byte_values:
            if isinstance(byte_count, bool) or not isinstance(byte_count, int) or byte_count <= 0:
                continue
            for hash_key, digest in hash_values:
                if not isinstance(digest, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", digest):
                    continue
                pair = (byte_count, digest.lower())
                pairs.setdefault(pair, []).append(list(path + (byte_key, hash_key)))
        for key, raw in value.items():
            if not isinstance(key, str) or "sha" not in key.lower():
                continue
            lower = key.lower()
            if not any(token in lower for token in ("published", "output", "repeated")):
                continue
            values = raw if isinstance(raw, list) else [raw]
            for item in values:
                if isinstance(item, str) and re.fullmatch(r"[0-9a-fA-F]{64}", item):
                    related_hashes.append((".".join(path + (key,)), item.lower()))
        for key in ("published", "output"):
            nested = value.get(key)
            if not isinstance(nested, Mapping):
                continue
            byte_count = nested.get("bytes")
            digest = nested.get("sha256")
            if (
                isinstance(byte_count, int)
                and not isinstance(byte_count, bool)
                and byte_count > 0
                and isinstance(digest, str)
                and re.fullmatch(r"[0-9a-fA-F]{64}", digest)
            ):
                pair = (byte_count, digest.lower())
                pairs.setdefault(pair, []).append(list(path + (key, "bytes", "sha256")))
    _require(pairs, "harness report has no published/output byte-and-SHA identity")
    _require(len(pairs) == 1, f"harness report has ambiguous output identities: {sorted(pairs)!r}")
    (byte_count, digest), locations = next(iter(pairs.items()))
    _require(
        all(item == digest for _, item in related_hashes),
        f"harness report has non-identical published/output hashes: {sorted(set(item for _, item in related_hashes))!r}",
    )
    return {
        "bytes": byte_count,
        "sha256": digest,
        "report_paths": locations,
        "related_hash_count": len(related_hashes),
    }


def _format_extension(corpus: Any, report: Mapping[str, Any], explicit: str | None) -> str | None:
    if explicit is not None:
        value = explicit.lower()
        if not value.startswith("."):
            value = "." + value
        _require(value in set(FORMAT_EXTENSIONS.values()), f"unsupported destination extension: {explicit!r}")
        return value
    candidates: set[str] = set()
    if isinstance(corpus, str):
        value = corpus.lower().lstrip(".")
        if value in FORMAT_EXTENSIONS:
            candidates.add(FORMAT_EXTENSIONS[value])
    elif isinstance(corpus, Mapping):
        for key in ("format", "package_format", "extension"):
            value = corpus.get(key)
            if isinstance(value, str):
                lower = value.lower()
                for name, extension in FORMAT_EXTENSIONS.items():
                    if lower == name or lower.startswith(name) or lower.endswith(extension):
                        candidates.add(extension)
    for path, value in _walk(report):
        for key in ("format", "package_format", "extension"):
            raw = value.get(key)
            if isinstance(raw, str):
                lower = raw.lower()
                for name, extension in FORMAT_EXTENSIONS.items():
                    if lower == name or lower.startswith(name) or lower.endswith(extension):
                        candidates.add(extension)
    _require(len(candidates) <= 1, f"report/corpus format is ambiguous: {sorted(candidates)!r}")
    return next(iter(candidates)) if candidates else None


def _policy_counts(policy: str) -> tuple[int, int]:
    _require(policy in POLICIES, f"unknown durability policy: {policy!r}")
    if policy in {"default", "full"}:
        return 1, 1
    if policy == "file-only":
        return 1, 0
    return 0, 0


def _expected_destinations(extension: str | None) -> tuple[tuple[str, ...], tuple[str, str]]:
    _require(extension is not None, "corpus format or destination extension is required")
    suffix = extension
    setup = tuple(f"{name}{suffix}" for name in SETUP_DESTINATION_NAMES)
    return setup, (f"published{suffix}", f"published-repeat{suffix}")


def _destination_probe(
    events: list[Event], destination: str, lower: int, upper: int, number: int
) -> Event:
    probes = [
        event
        for event in events[lower:upper]
        if event.call in METADATA_CALLS and _path_argument(event, destination)
    ]
    _require(len(probes) == 1, f"transaction {number}: destination permission probe is ambiguous: {len(probes)}")
    return probes[0]


def _raw_event(event: Event) -> dict[str, Any]:
    return {"line": event.line_no, "call": event.call, "raw": event.raw}


def _transaction(
    events: list[Event],
    candidate: Candidate,
    parent: str,
    destination_names: tuple[str, str],
    number: int,
    lower: int,
    upper: int,
    expected_file_syncs: int,
    expected_parent_syncs: int,
    published: dict[str, Any],
    measured: bool,
) -> dict[str, Any]:
    rename = candidate.event
    source = candidate.source
    destination = candidate.destination

    probe = _destination_probe(events, destination, lower, rename.index, number)
    opens = [
        event
        for event in events[lower:rename.index]
        if event.call in OPEN_CALLS
        and event.open_path == source
        and event.result_code is not None
        and event.result_code >= 0
    ]
    _require(len(opens) == 1, f"transaction {number}: temporary creation is ambiguous: {len(opens)}")
    temp_open = opens[0]
    temp_fd = _result_fd(temp_open.result)
    _require(temp_fd is not None, f"transaction {number}: temporary open has no descriptor")
    _require(temp_open.path == source, f"transaction {number}: temporary open path is not source")

    writes = [
        event
        for event in events[temp_open.index + 1 : rename.index]
        if event.call in WRITE_CALLS and event.fd == temp_fd
    ]
    _require(writes, f"transaction {number}: temporary file has no writes")
    _require(all(event.path == source for event in writes), f"transaction {number}: temporary write path is ambiguous")
    _require(all(event.result_code is not None and event.result_code >= 0 for event in writes),
             f"transaction {number}: temporary write failed")
    written_bytes = sum(event.result_code or 0 for event in writes)
    _require(written_bytes == published["bytes"],
             f"transaction {number}: writes {written_bytes} != report bytes {published['bytes']}")

    pre_rename_closes = [
        event
        for event in events[temp_open.index + 1 : rename.index]
        if event.call == "close" and event.fd == temp_fd and _successful(event)
    ]
    _require(not pre_rename_closes, f"transaction {number}: temporary descriptor closed before rename")

    fchmods = [
        event
        for event in events[temp_open.index : rename.index + 1]
        if event.call == "fchmod" and event.fd == temp_fd
    ]
    _require(len(fchmods) <= 1, f"transaction {number}: multiple temporary fchmod calls")
    _require(all(event.path == source for event in fchmods), f"transaction {number}: fchmod path is ambiguous")
    if measured:
        _require(not fchmods, "measured transaction copied permissions despite an absent destination")
        _require(probe.result_code is not None and probe.result_code < 0 and "ENOENT" in probe.result,
                 "measured destination probe did not prove ENOENT")

    file_sync_attempts = [
        event
        for event in events[temp_open.index + 1 : rename.index]
        if event.call in SYNC_CALLS and event.fd == temp_fd
    ]
    _require(len(file_sync_attempts) == expected_file_syncs,
             f"transaction {number}: expected {expected_file_syncs} temporary sync calls, found {len(file_sync_attempts)}")
    _require(all(event.path == source for event in file_sync_attempts),
             f"transaction {number}: temporary sync descriptor path is ambiguous")
    _require(all(_successful(event) for event in file_sync_attempts),
             f"transaction {number}: temporary sync did not succeed")
    _require(all(max((event.index for event in writes), default=-1) < event.index for event in file_sync_attempts),
             f"transaction {number}: temporary sync precedes a write")

    _require(max(event.index for event in writes) < rename.index, f"transaction {number}: write/rename order is invalid")

    parent_opens = [
        event
        for event in events[rename.index + 1 : upper]
        if event.call in OPEN_CALLS
        and event.open_path == parent
        and event.result_code is not None
        and event.result_code >= 0
    ]
    _require(len(parent_opens) == expected_parent_syncs,
             f"transaction {number}: expected {expected_parent_syncs} parent opens, found {len(parent_opens)}")
    parent_open: Event | None = parent_opens[0] if parent_opens else None
    parent_fd: int | None = None
    parent_syncs: list[Event] = []
    parent_closes: list[Event] = []
    if parent_open is not None:
        parent_fd = _result_fd(parent_open.result)
        _require(parent_fd is not None and parent_fd != temp_fd,
                 f"transaction {number}: invalid parent directory descriptor")
        _require(parent_open.path == parent, f"transaction {number}: parent open path is ambiguous")
        parent_sync_attempts = [
            event
            for event in events[parent_open.index + 1 : upper]
            if event.call in SYNC_CALLS and event.fd == parent_fd
        ]
        _require(len(parent_sync_attempts) == 1,
                 f"transaction {number}: expected one parent sync call")
        parent_syncs = [event for event in parent_sync_attempts if _successful(event)]
        _require(len(parent_syncs) == 1, f"transaction {number}: parent sync did not succeed")
        _require(parent_syncs[0].path == parent, f"transaction {number}: parent sync path is ambiguous")
        _require(rename.index < parent_syncs[0].index, f"transaction {number}: parent sync precedes rename")
        parent_closes = [
            event
            for event in events[parent_syncs[0].index + 1 : upper]
            if event.call == "close"
            and event.fd == parent_fd
            and _successful(event)
            and event.path == parent
        ]
        _require(len(parent_closes) == 1, f"transaction {number}: parent descriptor close is ambiguous")
        _require(parent_closes[0].path == parent, f"transaction {number}: parent close path is ambiguous")
    else:
        stray_parent_syncs = [
            event
            for event in events[rename.index + 1 : upper]
            if event.call in SYNC_CALLS and event.path == parent
        ]
        _require(not stray_parent_syncs, f"transaction {number}: parent sync present for a weaker policy")

    temp_closes = [
        event
        for event in events[rename.index + 1 : upper]
        if event.call == "close" and event.fd == temp_fd and _successful(event)
    ]
    # NamedTempFile::persist returns the still-open file, so its first close
    # after rename is the publication boundary.  The harness may then reuse
    # the same numeric descriptor for readback and report I/O before the next
    # replacement (or after the measured one); those closes are deliberately
    # outside this boundary and are distinguished by their later index/path.
    _require(temp_closes, f"transaction {number}: missing persisted descriptor close")
    persisted_close = temp_closes[0]
    _require(persisted_close.path in {source, destination},
             f"transaction {number}: persisted close path is ambiguous")
    if parent_closes:
        _require(parent_closes[0].index < persisted_close.index,
                 f"transaction {number}: persisted descriptor closed before parent directory")

    start_event = probe
    end_event = persisted_close
    start_ns = start_event.timestamp_ns
    end_ns = end_event.timestamp_ns + end_event.duration_ns
    _require(end_ns >= start_ns, f"transaction {number}: negative trace window")
    in_window = events[start_event.index : end_event.index + 1]
    syscall_total_ns = sum(event.duration_ns for event in in_window)
    all_syncs = file_sync_attempts + parent_syncs
    return {
        "index": number,
        "destination": destination,
        "temp_path": source,
        "parent_path": parent,
        "temp_fd": temp_fd,
        "parent_fd": parent_fd,
        "permission_probe_enoent": probe.result_code is not None and probe.result_code < 0 and "ENOENT" in probe.result,
        "line_range": [start_event.line_no, end_event.line_no],
        "event_range": [start_event.index, end_event.index],
        "window_start": "destination_permission_probe",
        "window_end": "persisted_descriptor_close",
        "window_ns": end_ns - start_ns,
        "syscall_total_ns": syscall_total_ns,
        "written_bytes": written_bytes,
        "write_calls": len(writes),
        "file_sync_calls": len(file_sync_attempts),
        "parent_sync_calls": len(parent_syncs),
        "fchmod_calls": len(fchmods),
        "replacement_renames": 1,
        "readback_or_cleanup_excluded": measured,
        "key_events": {
            "permission_probe": _raw_event(probe),
            "temporary_open": _raw_event(temp_open),
            "writes": [_raw_event(event) for event in writes],
            "fchmod": [_raw_event(event) for event in fchmods],
            "file_sync": [_raw_event(event) for event in file_sync_attempts],
            "rename": _raw_event(rename),
            "parent_open": _raw_event(parent_open) if parent_open is not None else None,
            "parent_sync": [_raw_event(event) for event in parent_syncs],
            "parent_close": [_raw_event(event) for event in parent_closes],
            "persisted_close": _raw_event(persisted_close),
        },
        "contract": {
            "output_bytes_match_report": written_bytes == published["bytes"],
            "temporary_sync_before_rename": all(event.index < rename.index for event in file_sync_attempts),
            "parent_sync_after_rename": all(event.index > rename.index for event in parent_syncs),
            "one_atomic_replacement": True,
        },
    }


def _publication_candidates(
    events: list[Event], destination_names: tuple[str, str]
) -> list[Candidate]:
    candidates: list[Candidate] = []
    for event in events:
        paths = _rename_paths(event)
        if paths is None:
            continue
        source, destination = paths
        if os.path.basename(destination) not in destination_names:
            continue
        parent = os.path.dirname(destination)
        if not parent or not os.path.isabs(destination) or not _is_temp_path(source, parent):
            _fail(f"successful destination rename has an ambiguous path: {event.raw}")
        candidates.append(Candidate(event, source, destination))
    return candidates


def _cleanup_between(
    events: list[Event], parent: str, destination_names: tuple[str, str], lower: int, upper: int
) -> dict[str, Any]:
    selected: list[Event] = []
    paths = [os.path.join(parent, name) for name in destination_names]
    for path in paths:
        matches = [
            event
            for event in events[lower:upper]
            if event.call in UNLINK_CALLS and _successful(event) and _path_argument(event, path)
        ]
        _require(len(matches) == 1, f"setup cleanup for {path} is not exactly one successful unlink")
        selected.extend(matches)
    selected.sort(key=lambda event: event.index)
    _require(selected[0].index < selected[-1].index + 1, "cleanup event ordering is invalid")
    return {
        "paths": paths,
        "line_range": [selected[0].line_no, selected[-1].line_no],
        "event_range": [selected[0].index, selected[-1].index],
        "successful_unlinks": len(selected),
        "raw_events": [_raw_event(event) for event in selected],
    }


def _post_save_scope(
    events: list[Event], destination: str, measured: dict[str, Any], published: dict[str, Any]
) -> dict[str, Any]:
    lower = measured["event_range"][1] + 1
    opens = [
        event
        for event in events[lower:]
        if event.call in OPEN_CALLS
        and event.open_path == destination
        and event.result_code is not None
        and event.result_code >= 0
    ]
    _require(len(opens) == 1, f"measured destination readback open is ambiguous: {len(opens)}")
    readback_open = opens[0]
    readback_fd = _result_fd(readback_open.result)
    _require(readback_fd is not None, "measured destination readback open has no descriptor")
    readback_closes = [
        event
        for event in events[readback_open.index + 1 :]
        if event.call == "close"
        and event.fd == readback_fd
        and _successful(event)
        and event.path == destination
    ]
    _require(len(readback_closes) == 1, "measured destination readback close is ambiguous")
    readback_close = readback_closes[0]
    _require(readback_close.path == destination, "measured destination readback close path is ambiguous")
    reads = [
        event
        for event in events[readback_open.index + 1 : readback_close.index]
        if event.call in READ_CALLS and event.fd == readback_fd
    ]
    _require(reads, "measured destination readback has no reads")
    _require(all(event.path == destination for event in reads), "measured destination read path is ambiguous")
    _require(all(event.result_code is not None and event.result_code >= 0 for event in reads),
             "measured destination read failed")
    read_bytes = sum(event.result_code or 0 for event in reads)
    _require(read_bytes == published["bytes"],
             f"measured destination read {read_bytes} != report bytes {published['bytes']}")
    unlinks = [
        event
        for event in events[readback_close.index + 1 :]
        if event.call in UNLINK_CALLS and _successful(event) and _path_argument(event, destination)
    ]
    _require(len(unlinks) == 1, "measured destination cleanup unlink is ambiguous")
    _require(unlinks[0].index > readback_close.index, "measured destination unlinked before readback close")
    return {
        "outside_measured_window": True,
        "readback_open": _raw_event(readback_open),
        "readback_reads": [_raw_event(event) for event in reads],
        "readback_bytes": read_bytes,
        "readback_close": _raw_event(readback_close),
        "destination_unlink": _raw_event(unlinks[0]),
    }


def _plan_value(plan: Mapping[str, Any] | None, key: str, default: Any) -> Any:
    if plan is None:
        return default
    return plan.get(key, default)


def analyze(
    trace_path: str | Path,
    report: str | Path | Mapping[str, Any],
    policy: str,
    plan: Mapping[str, Any] | str | Path | None = None,
    corpus: Mapping[str, Any] | str | None = None,
    *,
    extension: str | None = None,
) -> dict[str, Any]:
    """Validate one trace and return exportable publication evidence.

    ``report`` may be the harness report path or its already-decoded JSON
    object.  ``plan`` is optional and can be a decoded plan or a JSON path;
    the default contract is the frozen 4-setup/1-measured sequence.  The
    convenience arguments are deliberately plain Python values so the root
    validator can bind them to its receipt rows without importing capture
    code.
    """

    _require(isinstance(policy, str), "policy must be a string")
    # Keep the fourth positional argument convenient for a validator that has
    # a corpus row but no decoded plan.  A real plan has one of the frozen
    # trace keys; a corpus row has an id/format instead.  Explicit keyword
    # arguments remain unambiguous.
    if (
        corpus is None
        and isinstance(plan, Mapping)
        and not any(key in plan for key in ("trace_transaction_count", "trace_setup_policy", "trace_setup_destination_names"))
        and any(key in plan for key in ("id", "format", "package_format"))
    ):
        corpus, plan = plan, None
    if plan is not None and not isinstance(plan, Mapping):
        plan_path = Path(plan)
        _require(plan_path.is_file() and not plan_path.is_symlink(), f"missing plan: {plan_path}")
        try:
            loaded_plan = json.loads(plan_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
            _fail(f"unable to read plan {plan_path}: {error}")
        _require(isinstance(loaded_plan, Mapping), "plan is not a JSON object")
        plan = loaded_plan
    if plan is not None:
        diagnostics = plan.get("diagnostics")
        if isinstance(diagnostics, Mapping) and isinstance(diagnostics.get("trace"), Mapping):
            trace_plan = diagnostics["trace"]
            _require(trace_plan.get("options") == TRACE_OPTIONS,
                     "trace options differ from the frozen diagnostics plan")
            _require(trace_plan.get("phase") == "atomic_publish",
                     "trace parser requires the atomic_publish diagnostics phase")
            _require(trace_plan.get("samples") == 1 and trace_plan.get("warmup") == 0,
                     "trace parser requires one sample and no warmup")

    report_value, report_artifact = _load_report(report)
    published = _output_identity(report_value)
    extension = _format_extension(corpus, report_value, extension)
    setup_names, destination_names = _expected_destinations(extension)
    events, trace_meta = _parse_trace(Path(trace_path))

    expected_setup_names = tuple(_plan_value(plan, "trace_setup_destination_names", setup_names))
    _require(expected_setup_names == setup_names,
             f"trace setup destination contract changed: {expected_setup_names!r} != {setup_names!r}")
    expected_transactions = int(_plan_value(plan, "trace_transaction_count", 5))
    _require(expected_transactions == 5, f"trace transaction count changed: {expected_transactions}")
    setup_policy = str(_plan_value(plan, "trace_setup_policy", "full"))
    setup_file_syncs, setup_parent_syncs = _policy_counts(setup_policy)
    measured_file_syncs, measured_parent_syncs = _policy_counts(policy)

    candidates = _publication_candidates(events, destination_names)
    _require(len(candidates) == expected_transactions,
             f"expected {expected_transactions} atomic destination replacements, found {len(candidates)}")
    destinations = [os.path.basename(candidate.destination) for candidate in candidates]
    _require(destinations == list(setup_names) + [destination_names[0]],
             f"publication destination sequence changed: {destinations!r}")
    parent = os.path.dirname(candidates[0].destination)
    _require(parent and os.path.isabs(parent), f"publication parent is not an absolute path: {parent!r}")
    _require(all(os.path.dirname(candidate.destination) == parent for candidate in candidates),
             "destination replacements do not share one parent")
    _require(all(os.path.dirname(candidate.source) == parent for candidate in candidates),
             "temporary replacements do not use the destination parent")
    _require(len({candidate.event.index for candidate in candidates}) == 5,
             "atomic replacement events are not distinct")

    transactions: list[dict[str, Any]] = []
    previous_end = -1
    for number, candidate in enumerate(candidates, 1):
        upper = candidates[number].event.index if number < len(candidates) else len(events)
        is_measured = number == 5
        file_syncs = measured_file_syncs if is_measured else setup_file_syncs
        parent_syncs = measured_parent_syncs if is_measured else setup_parent_syncs
        transaction = _transaction(
            events,
            candidate,
            parent,
            destination_names,
            number,
            previous_end + 1,
            upper,
            file_syncs,
            parent_syncs,
            published,
            is_measured,
        )
        transactions.append(transaction)
        previous_end = transaction["event_range"][1]

    measured = transactions[-1]
    setup_cleanup = _cleanup_between(
        events,
        parent,
        destination_names,
        transactions[3]["event_range"][1] + 1,
        measured["event_range"][0],
    )
    post_save = _post_save_scope(events, measured["destination"], measured, published)

    result: dict[str, Any] = {
        "schema_version": 1,
        "trace_options": TRACE_OPTIONS,
        "trace": trace_meta,
        "report": report_artifact,
        "policy": policy,
        "setup_policy": setup_policy,
        "destination_extension": extension,
        "workspace_parent": parent,
        "destination_sequence": destinations,
        "publication_count": len(candidates),
        "output": published,
        "transactions": transactions,
        "setup_transactions": transactions[:4],
        "measured_transaction": measured,
        "measured_trace_window": {
            "line_range": measured["line_range"],
            "event_range": measured["event_range"],
            "duration_ns": measured["window_ns"],
            "syscall_total_ns": measured["syscall_total_ns"],
            "start": measured["window_start"],
            "end": measured["window_end"],
            "diagnostic_only": True,
        },
        "setup_cleanup": setup_cleanup,
        "post_save_scope": post_save,
        "contract": {
            "all_atomic_one_replacement": all(item["replacement_renames"] == 1 for item in transactions),
            "published_output_same": all(item["written_bytes"] == published["bytes"] for item in transactions),
            "report_output_bound": True,
            "setup_sequence_exact": destinations[:4] == list(setup_names),
            "measured_destination_missing": measured["permission_probe_enoent"],
            "measured_permission_copies_absent": measured["fchmod_calls"] == 0,
            "default_or_full_two_syncs": policy in {"default", "full"}
            and measured["file_sync_calls"] == 1
            and measured["parent_sync_calls"] == 1,
            "file_only_one_file_sync": policy == "file-only"
            and measured["file_sync_calls"] == 1
            and measured["parent_sync_calls"] == 0,
            "no_sync_zero_syncs": policy == "no-sync"
            and measured["file_sync_calls"] == 0
            and measured["parent_sync_calls"] == 0,
        },
        "limits": [
            "Trace syscall and window durations include strace perturbation and are diagnostic only.",
            "The measured trace window is destination permission probe through persisted descriptor close.",
            "Readback, setup transactions, setup cleanup, and final cleanup are outside the measured window.",
            "The parser makes no native-time fraction, device, cache, or optimization claim.",
            "Output bytes and SHA-256 are bound to the harness report; strace cannot prove content identity by itself.",
        ],
    }
    return result


def analyze_trace(
    trace_path: str | Path,
    policy: str,
    report: str | Path | Mapping[str, Any],
    plan: Mapping[str, Any] | str | Path | None = None,
    corpus: Mapping[str, Any] | str | None = None,
    *,
    extension: str | None = None,
) -> dict[str, Any]:
    """Compatibility spelling for validators modeled after change 0714."""

    return analyze(trace_path, report, policy, plan, corpus, extension=extension)


__all__ = ["TRACE_OPTIONS", "TraceError", "analyze", "analyze_trace"]
