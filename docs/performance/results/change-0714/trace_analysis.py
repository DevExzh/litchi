#!/usr/bin/env python3
"""Parse and validate the whole-child ordinary-save strace boundary.

The trace is deliberately parsed as evidence of one publication transaction,
not as a source of native timing.  The child also contains corpus setup,
readback, JSON output, and cleanup, so publication transactions are identified
by their sibling tempfile and destination paths before any syscall metrics are
reported.
"""

from __future__ import annotations

import ast
from decimal import Decimal, InvalidOperation
import hashlib
import os
from pathlib import Path
import re
from typing import Any


TRACE_OPTIONS = [
    "-ttt",
    "-T",
    "-yy",
    "-s",
    "256",
    "-e",
    "trace=%file,read,pread64,write,pwrite64,writev,lseek,close,fsync,fdatasync,fchmod",
]
TEMP_RE = re.compile(r"^\.litchi-[A-Za-z0-9]{6}\.tmp$")
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
UNLINK_CALLS = {"unlink", "unlinkat"}
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
SYNC_CALLS = {"fsync", "fdatasync"}


def _fail(message: str) -> None:
    raise AssertionError(message)


def _require(condition: bool, message: str) -> None:
    if not condition:
        _fail(message)


class Event:
    def __init__(self, index: int, line_no: int, raw: str, timestamp_ns: int,
                 call: str, body: str, result: str, result_code: int | None,
                 duration_ns: int) -> None:
        self.index = index
        self.line_no = line_no
        self.raw = raw
        self.timestamp_ns = timestamp_ns
        self.call = call
        self.body = body
        self.result = result
        self.result_code = result_code
        self.duration_ns = duration_ns
        self.fd: int | None = None
        self.path: str | None = None
        self.open_path: str | None = None


def _decimal_ns(value: str, label: str) -> int:
    try:
        decimal = Decimal(value)
    except InvalidOperation:
        _fail(f"{label} is not a decimal timestamp: {value!r}")
    _require(decimal >= 0, f"{label} is negative")
    # -ttt/-T values in this packet have sub-nanosecond-free decimal output.
    return int(decimal * Decimal(1_000_000_000))


def _result_code(result: str) -> int | None:
    match = re.match(r"\s*(-?\d+)(?:\s|<|$)", result)
    return int(match.group(1)) if match else None


def _quoted_strings(body: str) -> list[str]:
    values: list[str] = []
    for token in QUOTED_RE.findall(body):
        try:
            value = ast.literal_eval(token)
        except (SyntaxError, ValueError):
            _fail(f"unable to decode strace string {token!r}")
        _require(isinstance(value, str), f"decoded strace value is not a string: {token!r}")
        values.append(value)
    return values


def _norm_path(value: str) -> str:
    # All harness paths are absolute.  normpath handles harmless // and /./
    # differences without consulting the filesystem (the workspace is removed
    # before the analyzer runs).
    return os.path.normpath(value)


def _first_fd(body: str) -> int | None:
    match = re.match(r"\s*(-?\d+)(?:<[^>]*>)?(?:\s*,|\s*$)", body)
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
    if not values:
        return _result_annotation(event.result)
    # openat/openat2 have the path after the dirfd; all quoted values in the
    # current trace are path arguments.  The first is sufficient and avoids
    # mistaking a quoted diagnostic string in a future flags object for a path.
    return _norm_path(values[0])


def _parse_trace(path: Path) -> tuple[list[Event], int, dict[str, Any]]:
    _require(path.is_file() and not path.is_symlink(), f"missing trace: {path}")
    data = path.read_bytes()
    digest = hashlib.sha256(data).hexdigest()
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        _fail(f"trace is not UTF-8: {error}")

    events: list[Event] = []
    exit_codes: list[int] = []
    ignored_lines = 0
    for line_no, raw_line in enumerate(text.splitlines(), 1):
        raw = raw_line.rstrip("\r")
        if not raw.strip():
            ignored_lines += 1
            continue
        if "<unfinished ...>" in raw or "<..." in raw and "resumed>" in raw:
            _fail(f"trace contains non-serial syscall line {line_no}: {raw}")
        exit_match = EXIT_RE.match(raw)
        if exit_match:
            exit_codes.append(int(exit_match.group("code")))
            continue
        # Signals and strace diagnostics are not syscalls.  A timestamped line
        # that is not a complete syscall is rejected so malformed evidence
        # cannot silently change transaction attribution.
        if (raw.lstrip().startswith(("--- ", "+++ ", "strace:"))
                or re.match(r"^(?:\d+\s+)?\d+\.\d+\s+--- ", raw)
                or re.match(r"^(?:\d+\s+)?\d+\.\d+\s+\+\+\+ ", raw)):
            ignored_lines += 1
            continue
        match = SYSCALL_RE.match(raw)
        if not match:
            if re.match(r"^(?:\d+\s+)?\d+\.\d+\s+", raw):
                _fail(f"unparsed timestamped trace line {line_no}: {raw}")
            ignored_lines += 1
            continue
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
        )
        event.fd = _first_fd(event.body)
        event.open_path = _open_path(event)
        event.path = event.open_path or _result_annotation(event.result)
        events.append(event)

    _require(exit_codes == [0], f"trace exit markers are not exactly [0]: {exit_codes!r}")
    _require(events, "trace has no completed syscall lines")

    # Annotations are authoritative when present.  Keep an fd table as a
    # fallback for fsync/close/write lines whose argument annotation is absent
    # on another strace/platform representation.
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
            known_path = fd_paths.get(event.fd)
            event.path = argument_path or event.path or known_path
        if event.call == "close" and event.result_code == 0 and event.fd is not None:
            # Preserve event.path for diagnostics, then release the descriptor.
            fd_paths.pop(event.fd, None)

    return events, len(text.splitlines()), {
        "sha256": digest,
        "bytes": len(data),
        "ignored_lines": ignored_lines,
        "exit_codes": exit_codes,
    }


def _successful(event: Event) -> bool:
    return event.result_code == 0


def _event_paths(event: Event) -> list[str]:
    return [_norm_path(value) for value in _quoted_strings(event.body)]


def _rename_paths(event: Event) -> tuple[str, str] | None:
    if event.call not in RENAME_CALLS or not _successful(event):
        return None
    paths = _event_paths(event)
    if len(paths) < 2:
        return None
    # renameat/renameat2 contain two dirfds and two quoted path arguments;
    # rename contains two quoted arguments.  Both therefore use the first pair.
    return paths[0], paths[1]


def _path_argument(event: Event, expected: str) -> bool:
    return expected in _event_paths(event)


def _is_temp_path(value: str, parent: str) -> bool:
    return _norm_path(os.path.dirname(value)) == parent and bool(TEMP_RE.fullmatch(os.path.basename(value)))


def _selection(plan: dict[str, Any], phase: str) -> tuple[list[str], str, int]:
    _require(isinstance(plan, dict), "plan is not an object")
    trace = plan.get("trace")
    _require(isinstance(trace, dict), "trace plan is missing")
    _require(trace.get("samples") == 1 and trace.get("warmup") == 0,
             "trace parser requires samples=1,warmup=0")
    _require(trace.get("options") == TRACE_OPTIONS, "trace options changed from frozen strace boundary")
    _require(phase in trace.get("phases", []), f"phase {phase!r} is not in trace plan")
    selection = plan.get("trace_selection")
    _require(isinstance(selection, dict), "trace selection is missing")
    setup = selection.get("setup_destinations")
    _require(setup == ["published.docx", "published.docx", "published-repeat.docx", "published.docx"],
             "setup destination sequence changed")
    destination = selection.get("atomic_measured_destination")
    _require(destination == "published.docx", "atomic measured destination changed")
    expected = selection.get("atomic_total_transactions" if phase == "atomic_publish"
                             else "counting_total_transactions")
    _require(expected == (5 if phase == "atomic_publish" else 4),
             f"{phase} transaction count changed")
    _require(selection.get("require_setup_cleanup_before_measured") is True,
             "setup cleanup boundary is not required")
    _require(selection.get("require_temporary_sync_before_rename") is True,
             "temporary sync boundary is not required")
    _require(selection.get("require_parent_sync_after_rename") is True,
             "parent sync boundary is not required")
    return setup, destination, expected


def _publication_candidates(events: list[Event]) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for event in events:
        paths = _rename_paths(event)
        if paths is None:
            continue
        source, destination = paths
        if (os.path.basename(destination) not in {"published.docx", "published-repeat.docx"}
                or not _is_temp_path(source, _norm_path(os.path.dirname(destination)))):
            continue
        candidates.append({"event": event, "source": source, "destination": destination})
    return candidates


def _probe_events(events: list[Event], destination: str, start: int, end: int) -> list[Event]:
    return [event for event in events[start:end]
            if event.call in METADATA_CALLS and _path_argument(event, destination)]


def _write_bytes(event: Event) -> int | None:
    if event.result_code is None or event.result_code < 0:
        return None
    return event.result_code


def _transaction(
    events: list[Event],
    candidate: dict[str, Any],
    parent: str,
    published_bytes: int,
    number: int,
    previous_end: int,
    upper_bound: int,
) -> dict[str, Any]:
    rename: Event = candidate["event"]
    source: str = candidate["source"]
    destination: str = candidate["destination"]
    opens = [event for event in events[previous_end + 1:rename.index]
             if event.call in OPEN_CALLS and event.open_path == source
             and event.result_code is not None and event.result_code >= 0]
    _require(len(opens) == 1, f"transaction {number}: expected one temp creation before rename")
    temp_open = opens[0]
    temp_fd = _result_fd(temp_open.result)
    _require(temp_fd is not None, f"transaction {number}: temp open has no fd")
    temp_return_path = _result_annotation(temp_open.result)
    _require(temp_return_path is None or temp_return_path == source,
             f"transaction {number}: temp open fd/path mismatch")

    writes = [event for event in events[temp_open.index + 1:rename.index]
              if event.call in WRITE_CALLS and event.fd == temp_fd]
    _require(writes, f"transaction {number}: no writes on temp fd")
    _require(all(event.path == source for event in writes),
             f"transaction {number}: a temp write is not on the temp path")
    _require(all(_write_bytes(event) is not None for event in writes),
             f"transaction {number}: temp write failed")
    written_bytes = sum(_write_bytes(event) or 0 for event in writes)
    _require(written_bytes == published_bytes,
             f"transaction {number}: temp writes {written_bytes} != published bytes {published_bytes}")

    premature_closes = [event for event in events[temp_open.index + 1:rename.index]
                         if event.call == "close" and event.fd == temp_fd and _successful(event)]
    _require(not premature_closes,
             f"transaction {number}: temp fd closed before rename")
    premature_reopens = [event for event in events[temp_open.index + 1:rename.index]
                         if event.call in OPEN_CALLS and _result_fd(event.result) == temp_fd
                         and event.result_code is not None and event.result_code >= 0]
    _require(not premature_reopens,
             f"transaction {number}: temp fd was reused before rename")

    file_sync_attempts = [event for event in events[temp_open.index + 1:rename.index]
                          if event.call in SYNC_CALLS and event.fd == temp_fd]
    _require(len(file_sync_attempts) == 1,
             f"transaction {number}: expected one temp fsync/fdatasync attempt")
    file_syncs = [event for event in file_sync_attempts if _successful(event)]
    _require(len(file_syncs) == 1,
             f"transaction {number}: temporary fsync/fdatasync did not succeed")
    _require(file_syncs[0].path == source,
             f"transaction {number}: temp sync fd/path mismatch")
    _require(file_syncs[0].index < rename.index,
             f"transaction {number}: temp sync is not before rename")
    _require(max(event.index for event in writes) < file_syncs[0].index,
             f"transaction {number}: a temp write occurs after file sync")

    chmods = [event for event in events[temp_open.index:rename.index + 1]
              if event.call == "fchmod"]
    _require(len(chmods) <= 1, f"transaction {number}: multiple temporary fchmod calls")

    parent_opens = [event for event in events[rename.index + 1:upper_bound]
                    if event.call in OPEN_CALLS and event.open_path == parent
                    and event.result_code is not None and event.result_code >= 0]
    _require(len(parent_opens) == 1,
             f"transaction {number}: expected one parent-directory open after rename")
    parent_open = parent_opens[0]
    parent_fd = _result_fd(parent_open.result)
    _require(parent_fd is not None and parent_fd != temp_fd,
             f"transaction {number}: invalid parent-directory fd")
    parent_return_path = _result_annotation(parent_open.result)
    _require(parent_return_path is None or parent_return_path == parent,
             f"transaction {number}: parent open fd/path mismatch")

    parent_sync_attempts = [event for event in events[parent_open.index + 1:upper_bound]
                            if event.call in SYNC_CALLS and event.fd == parent_fd]
    _require(len(parent_sync_attempts) == 1,
             f"transaction {number}: expected one parent-directory fsync/fdatasync attempt")
    parent_syncs = [event for event in parent_sync_attempts if _successful(event)]
    _require(len(parent_syncs) == 1,
             f"transaction {number}: parent-directory fsync/fdatasync did not succeed")
    parent_sync = parent_syncs[0]
    _require(parent_sync.index > rename.index, f"transaction {number}: parent sync precedes rename")
    _require(parent_sync.path == parent,
             f"transaction {number}: parent sync fd/path mismatch")

    parent_closes = [event for event in events[parent_sync.index + 1:upper_bound]
                     if event.call == "close" and event.fd == parent_fd and _successful(event)]
    _require(parent_closes,
             f"transaction {number}: no parent-directory close")
    parent_close = parent_closes[0]
    _require(parent_close.path == parent,
             f"transaction {number}: parent close fd/path mismatch")

    persisted_closes = [event for event in events[parent_sync.index + 1:upper_bound]
                        if event.call == "close" and event.fd == temp_fd and _successful(event)]
    _require(persisted_closes,
             f"transaction {number}: no persisted target fd close")
    persisted_close = persisted_closes[0]
    _require(parent_close.index < persisted_close.index,
             f"transaction {number}: persisted fd closed before parent directory")
    _require(persisted_close.path in {source, destination},
             f"transaction {number}: persisted close fd/path mismatch")

    # No event belonging to the next setup publication may be absorbed into
    # this transaction.  The close of the persisted target is the boundary.
    first_index = temp_open.index
    last_index = persisted_close.index
    in_window = events[first_index:last_index + 1]
    syscall_total_ns = sum(event.duration_ns for event in in_window)
    window_ns = (persisted_close.timestamp_ns + persisted_close.duration_ns
                 - temp_open.timestamp_ns)
    _require(window_ns >= 0, f"transaction {number}: negative trace window")

    probes = _probe_events(events, destination, previous_end + 1, temp_open.index)
    _require(probes, f"transaction {number}: missing destination permission probe")
    probe = probes[-1]

    summary: dict[str, Any] = {
        "index": number,
        "destination": destination,
        "temp_path": source,
        "parent_path": parent,
        "temp_fd": temp_fd,
        "parent_fd": parent_fd,
        "line_range": [temp_open.line_no, persisted_close.line_no],
        "event_range": [first_index, last_index],
        "raw_key_lines": {
            "permission_probe": [event.raw for event in probes],
            "temp_open": temp_open.raw,
            "writes": [event.raw for event in writes],
            "file_sync": [event.raw for event in file_syncs],
            "fchmod": [event.raw for event in chmods],
            "rename": rename.raw,
            "parent_open": parent_open.raw,
            "parent_sync": [event.raw for event in [parent_sync]],
            "parent_close": parent_close.raw,
            "persisted_close": persisted_close.raw,
        },
        "counts": {
            "writes": len(writes),
            "file_sync": len(file_syncs),
            "fchmod": len(chmods),
            "rename": 1,
            "parent_open": 1,
            "parent_sync": 1,
            "parent_close": 1,
            "persisted_close": 1,
            "syscalls_in_window": len(in_window),
        },
        "durations_ns": {
            "temp_open": temp_open.duration_ns,
            "writes": sum(event.duration_ns for event in writes),
            "file_sync": file_syncs[0].duration_ns,
            "fchmod": sum(event.duration_ns for event in chmods),
            "rename": rename.duration_ns,
            "parent_open": parent_open.duration_ns,
            "parent_sync": parent_sync.duration_ns,
            "parent_close": parent_close.duration_ns,
            "persisted_close": persisted_close.duration_ns,
            "syscall_total": syscall_total_ns,
            "window": window_ns,
        },
        "write_calls": len(writes),
        "written_bytes": written_bytes,
        "file_sync_ns": file_syncs[0].duration_ns,
        "parent_sync_ns": parent_sync.duration_ns,
        "syscall_total_ns": syscall_total_ns,
        "window_ns": window_ns,
        "permission_probe": {
            "line": probe.line_no,
            "raw": probe.raw,
            "result": probe.result,
            "enoent": "ENOENT" in probe.result,
        },
    }
    return summary


def _cleanup_summary(events: list[Event], destination: str, alternate: str,
                     after: int, before: int) -> dict[str, Any]:
    selected: list[Event] = []
    for path in (destination, alternate):
        matches = [event for event in events[after + 1:before]
                   if event.call in UNLINK_CALLS and _successful(event)
                   and _path_argument(event, path)]
        _require(len(matches) == 1, f"setup cleanup for {path} is not exactly one successful unlink")
        selected.extend(matches)
    selected.sort(key=lambda event: event.index)
    return {
        "line_range": [selected[0].line_no, selected[-1].line_no],
        "paths": [destination, alternate],
        "raw_lines": [event.raw for event in selected],
        "counts": {"successful_unlink": len(selected)},
    }


def _post_save_scope(events: list[Event], destination: str, measured: dict[str, Any]) -> dict[str, Any]:
    end = measured["event_range"][1]
    opens = [event for event in events[end + 1:]
             if event.call in OPEN_CALLS and event.open_path == destination
             and event.result_code is not None and event.result_code >= 0]
    unlinks = [event for event in events[end + 1:]
               if event.call in UNLINK_CALLS and _successful(event)
               and _path_argument(event, destination)]
    _require(opens, "measured destination readback open is not after atomic window")
    readback_fd = _result_fd(opens[0].result)
    _require(readback_fd is not None, "measured destination readback open has no fd")
    readback_closes = [event for event in events[opens[0].index + 1:]
                       if event.call == "close" and event.fd == readback_fd
                       and _successful(event)]
    _require(readback_closes, "measured destination readback fd is not closed")
    readback_close = readback_closes[0]
    readbacks = [event for event in events[opens[0].index + 1:]
                 if event.call in {"read", "pread64", "preadv", "preadv2"}
                 and event.fd == readback_fd
                 and event.index < readback_close.index
                 and event.result_code is not None and event.result_code >= 0]
    _require(readbacks, "measured destination readback has no successful read")
    _require(unlinks, "measured destination unlink is not after atomic window")
    return {
        "destination_readback_open": {"line": opens[0].line_no, "raw": opens[0].raw},
        "destination_readback_reads": [event.raw for event in readbacks],
        "destination_readback_close": {"line": readback_close.line_no, "raw": readback_close.raw},
        "destination_unlink": {"line": unlinks[0].line_no, "raw": unlinks[0].raw},
        "outside_measured_window": True,
    }


def analyze_trace(path: Path, phase: str, plan: dict, published_bytes: int) -> dict:
    """Return validated setup/measured publication evidence from one strace.

    ``window_ns`` and syscall durations are diagnostic values from strace's
    own clock.  They include tracer perturbation and must not be interpreted as
    native elapsed time or as a device/hardware cause.
    """

    _require(isinstance(path, Path), "trace path must be a pathlib.Path")
    _require(isinstance(published_bytes, int) and not isinstance(published_bytes, bool)
             and published_bytes > 0, "published_bytes must be a positive integer")
    setup_names, measured_name, expected_count = _selection(plan, phase)
    events, line_count, trace_meta = _parse_trace(path)
    candidates = _publication_candidates(events)
    _require(len(candidates) == expected_count,
             f"expected {expected_count} publication renames, found {len(candidates)}")
    destinations = [os.path.basename(candidate["destination"]) for candidate in candidates]
    _require(destinations == setup_names + ([measured_name] if phase == "atomic_publish" else []),
             f"publication destination sequence changed: {destinations!r}")
    parent = _norm_path(os.path.dirname(candidates[0]["destination"]))
    _require(parent, "publication parent directory is empty")
    _require(all(_norm_path(os.path.dirname(candidate["destination"])) == parent
                 and _norm_path(os.path.dirname(candidate["source"])) == parent
                 for candidate in candidates), "publication paths do not share one parent directory")
    frozen_root = plan.get("filesystem_root")
    _require(isinstance(frozen_root, str) and os.path.isabs(frozen_root),
             "frozen filesystem_root is missing or relative")
    frozen_root = _norm_path(frozen_root)
    try:
        beneath_root = os.path.commonpath([parent, frozen_root]) == frozen_root
    except ValueError:
        beneath_root = False
    _require(beneath_root and parent != frozen_root,
             f"workspace parent {parent!r} is outside frozen filesystem_root {frozen_root!r}")

    transactions: list[dict[str, Any]] = []
    previous_end = -1
    for number, candidate in enumerate(candidates, 1):
        upper_bound = (candidates[number]["event"].index
                       if number < len(candidates) else len(events))
        transaction = _transaction(events, candidate, parent, published_bytes, number,
                                   previous_end, upper_bound)
        transactions.append(transaction)
        previous_end = transaction["event_range"][1]

    cleanup = None
    measured = None
    post_save = None
    destination = _norm_path(os.path.join(parent, "published.docx"))
    alternate = _norm_path(os.path.join(parent, "published-repeat.docx"))
    if phase == "atomic_publish":
        measured = transactions[-1]
        setup_end = transactions[3]["event_range"][1]
        measured_start = measured["event_range"][0]
        cleanup = _cleanup_summary(
            events,
            destination,
            alternate,
            setup_end,
            measured_start,
        )
        _require(measured["permission_probe"]["enoent"],
                 "measured destination permission probe did not return ENOENT")
        _require(measured["counts"]["fchmod"] == 0,
                 "measured transaction unexpectedly changed destination permissions")
        post_save = _post_save_scope(
            events,
            destination,
            measured,
        )
    else:
        cleanup = _cleanup_summary(events, destination, alternate,
                                   transactions[-1]["event_range"][1], len(events))

    setup_summaries = transactions[:4]
    measured_fields = None
    if measured is not None:
        measured_fields = {
            key: measured[key]
            for key in ("write_calls", "written_bytes", "file_sync_ns", "parent_sync_ns",
                        "syscall_total_ns", "window_ns")
        }

    return {
        "schema_version": 1,
        "phase": phase,
        "trace_sha256": trace_meta["sha256"],
        "trace_bytes": trace_meta["bytes"],
        "trace_lines": line_count,
        "serial_completed": True,
        "exit_code": 0,
        "workspace_parent": parent,
        "destination_sequence": destinations,
        "setup_transaction_count": len(setup_summaries),
        "transaction_count": len(transactions),
        "setup_transactions": setup_summaries,
        "transactions": transactions,
        "measured": measured_fields,
        "measured_transaction": measured,
        "setup_cleanup": cleanup,
        "post_save_scope": post_save,
        "contract": {
            "published_bytes": published_bytes,
            "temporary_write_total_equals_published_bytes": all(
                transaction["written_bytes"] == published_bytes for transaction in transactions
            ),
            "temporary_sync_before_rename": True,
            "parent_sync_after_rename": True,
            "persisted_fd_close_after_parent_sync": True,
            "measured_permission_probe_enoent": phase == "atomic_publish",
            "measured_fchmod_absent": phase == "atomic_publish",
        },
        "limits": [
            "Trace syscall and window durations include strace perturbation; they are not native elapsed values.",
            "The trace does not identify a hardware or device cause and supports no optimization claim.",
            "The measured window is a diagnostic boundary from temp open through persisted target close.",
            "Global read/write totals are intentionally excluded from publication attribution.",
            "Durability semantics are treated as contract requirements: temp flush/sync, replacement rename, and parent-directory sync.",
        ],
    }
