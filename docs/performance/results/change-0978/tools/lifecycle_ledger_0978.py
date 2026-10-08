#!/usr/bin/env python3
"""Build the 0978 disjoint lifetime/attribution ledger.

The variable-layout sample decoder is deliberately kept in ``raw_stack_0978``
and is owned by the raw-parser author.  This module consumes its retained
sample rows and a ``perf script -D`` stream to bind each sample to the
COMM/MMAP2/EXIT timeline.  It does not reinterpret record bytes or repair a
malformed sample.

The output is intentionally a small, stable interface for the focused decoder
and reducer:

* ``samples`` is one row per retained PERF_RECORD_SAMPLE, with the raw row
  identity and a disjoint phase/unknown classification;
* ``lifecycle`` is the framed control-record census and target map timeline;
* ``ledger`` contains D/M/R, T and marker known/unknown conservation; and
* ``gates`` reports structural failures without applying the old
  all-samples-empty-callchain gate.

A physical owner supplied by the corrected perf-script decoder is accepted for
any validated DSO, including libc, the dynamic loader, Python, and other
DSOs.  ``target_source_eligible`` remains a separate claim and is never
inferred from marker ancestry alone.
"""
from __future__ import annotations

import collections
import hashlib
import json
import re
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


SCHEMA = "litchi.performance.0978.lifecycle_ledger.v1"
EVENT_SCHEMA = "litchi.performance.0978.lifecycle_event.v1"


class LedgerFailure(ValueError):
    """Fail-closed lifecycle or conservation error."""

    def __init__(self, code: str, message: str, details: Mapping[str, Any] | None = None):
        super().__init__(f"{code}: {message}")
        self.code = code
        self.message = message
        self.details = dict(details or {})

    def as_dict(self) -> dict[str, Any]:
        return {"status": "failed", "code": self.code, "message": self.message, "details": self.details}


def _fail(code: str, message: str, **details: Any) -> None:
    raise LedgerFailure(code, message, details)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def file_binding(path: Path) -> dict[str, Any]:
    path = Path(path).absolute()
    if not path.is_file():
        _fail("input_missing", "lifecycle input is missing", path=str(path))
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


# perf script -D puts the timestamp immediately after the CPU field.  The
# event descriptor itself can contain arbitrary hexadecimal offsets, so parse
# only the stable prefix and keep the complete descriptor for diagnostics.
EVENT_PREFIX_RE = re.compile(
    r"^(?P<cpu>-?\d+)\s+(?P<time>\d+)\s+(?P<offset>0x[0-9a-fA-F]+)\s+"
    r"\[(?P<size>0x[0-9a-fA-F]+)\]:\s+(?P<descriptor>PERF_RECORD_[^\n]+)$"
)
COMM_RE = re.compile(
    r"PERF_RECORD_COMM(?P<exec>\s+exec)?:\s*(?P<comm>.*?):(?P<pid>\d+)/(?P<tid>\d+)\s*$"
)
MMAP2_RE = re.compile(
    r"PERF_RECORD_MMAP2\s+(?P<pid>\d+)/(?:\s*)?(?P<tid>\d+):\s+"
    r"\[(?P<start>0x[0-9a-fA-F]+)\((?P<length>0x[0-9a-fA-F]+)\)\s+"
    r"@\s+(?P<offset>(?:0x)?[0-9a-fA-F]+)\s+(?P<meta>[^\]]*)\]:\s+"
    r"(?P<perm>\S+)\s+(?P<path>.*)$"
)
EXIT_RE = re.compile(
    r"PERF_RECORD_EXIT\((?P<pid>\d+):(?P<tid>\d+)\):\((?P<ppid>\d+):(?P<ptid>\d+)\)"
)
FORK_RE = re.compile(
    r"PERF_RECORD_(?P<kind>FORK|CLONE)\((?P<child_pid>\d+):(?P<child_tid>\d+)\):"
    r"\((?P<parent_pid>\d+):(?P<parent_tid>\d+)\)"
)
THREAD_MAP_RE = re.compile(r"PERF_RECORD_THREAD_MAP\s+nr:\s*(?P<nr>\d+)\s+thread:\s*(?P<threads>.*)$")
SAMPLE_DESCRIPTOR_RE = re.compile(
    r"PERF_RECORD_SAMPLE\(IP,\s*0x(?P<misc>[0-9a-fA-F]+)\):\s+"
    r"(?P<pid>\d+)/(?P<tid>\d+):\s+0x(?P<ip>[0-9a-fA-F]+)\s+period:\s+(?P<period>\d+)"
)
LOST_RE = re.compile(r"PERF_RECORD_LOST:\s+.*?\blost:(?P<lost>\d+)\b")
# perf versions print this either as ``...: 0`` or with a labelled count.
# Keep both forms bound to the raw type-13 payload below.
LOST_SAMPLES_RE = re.compile(
    r"PERF_RECORD_LOST_SAMPLES(?:\s*:\s*(?P<bare>\d+)|(?::|\s).*?\b(?:lost|count)\s*:\s*(?P<lost>\d+))",
    re.I,
)
RAW_EVENT_RE = re.compile(r"^\.\s+\.\.\. raw event: size (?P<size>\d+) bytes$")
RAW_HEX_RE = re.compile(r"^\.\s+(?P<offset>[0-9a-fA-F]+):\s+(?P<bytes>(?:[0-9a-fA-F]{2}(?:\s+|$))*)")

# The core perf record type is stable in linux/perf_event.h.  Unknown control
# records are still framed and hashed, but their descriptor is not silently
# assigned a guessed event-specific parser.
RECORD_TYPES = {
    "PERF_RECORD_MMAP": 1,
    "PERF_RECORD_LOST": 2,
    "PERF_RECORD_COMM": 3,
    "PERF_RECORD_EXIT": 4,
    "PERF_RECORD_THROTTLE": 5,
    "PERF_RECORD_UNTHROTTLE": 6,
    "PERF_RECORD_FORK": 7,
    "PERF_RECORD_READ": 8,
    "PERF_RECORD_SAMPLE": 9,
    "PERF_RECORD_MMAP2": 10,
    "PERF_RECORD_AUX": 11,
    "PERF_RECORD_ITRACE_START": 12,
    "PERF_RECORD_LOST_SAMPLES": 13,
    "PERF_RECORD_SWITCH": 14,
    "PERF_RECORD_SWITCH_CPU_WIDE": 15,
    "PERF_RECORD_NAMESPACES": 16,
    "PERF_RECORD_KSYMBOL": 17,
    "PERF_RECORD_BPF_EVENT": 18,
    "PERF_RECORD_CGROUP": 19,
    "PERF_RECORD_TEXT_POKE": 20,
    "PERF_RECORD_AUX_OUTPUT_HW_ID": 21,
    "PERF_RECORD_HEADER_ATTR": 64,
    "PERF_RECORD_HEADER_EVENT_TYPE": 65,
    "PERF_RECORD_HEADER_TRACING_DATA": 66,
    "PERF_RECORD_HEADER_BUILD_ID": 67,
    "PERF_RECORD_FINISHED_ROUND": 68,
    "PERF_RECORD_ID_INDEX": 69,
    "PERF_RECORD_AUXTRACE_INFO": 70,
    "PERF_RECORD_AUXTRACE": 71,
    "PERF_RECORD_AUXTRACE_ERROR": 72,
    "PERF_RECORD_THREAD_MAP": 73,
    "PERF_RECORD_CPU_MAP": 74,
    "PERF_RECORD_STAT_CONFIG": 75,
    "PERF_RECORD_STAT": 76,
    "PERF_RECORD_STAT_ROUND": 77,
    "PERF_RECORD_EVENT_UPDATE": 78,
    "PERF_RECORD_TIME_CONV": 79,
    "PERF_RECORD_HEADER_FEATURE": 80,
    "PERF_RECORD_COMPRESSED": 81,
    "PERF_RECORD_FINISHED_INIT": 82,
    "PERF_RECORD_CGROUP2": 83,
}


def _event_prefix(line: str) -> dict[str, Any] | None:
    match = EVENT_PREFIX_RE.match(line.rstrip("\n"))
    if match is None:
        return None
    descriptor = match.group("descriptor")
    kind_match = re.match(r"(PERF_RECORD_[A-Z0-9_]+)", descriptor)
    if kind_match is None:
        return None
    kind = kind_match.group(1)
    return {
        "schema": EVENT_SCHEMA,
        "line": None,
        "cpu": int(match.group("cpu")),
        "time_ns": int(match.group("time")),
        "record_offset": int(match.group("offset"), 16),
        "record_bytes": int(match.group("size"), 16),
        "kind": kind,
        "descriptor": descriptor,
    }


def _parse_event(prefix: Mapping[str, Any]) -> dict[str, Any]:
    row = dict(prefix)
    descriptor = str(row["descriptor"])
    kind = str(row["kind"])
    if kind == "PERF_RECORD_COMM":
        match = COMM_RE.search(descriptor)
        if match is None:
            _fail("lifecycle_record", "COMM record is malformed", descriptor=descriptor)
        row.update({
            "exec": bool(match.group("exec")),
            "comm": match.group("comm"),
            "pid": int(match.group("pid")),
            "tid": int(match.group("tid")),
        })
    elif kind == "PERF_RECORD_MMAP2":
        match = MMAP2_RE.search(descriptor)
        if match is None:
            _fail("lifecycle_record", "MMAP2 record is malformed", descriptor=descriptor)
        meta = match.group("meta").strip()
        build_ids = re.findall(r"<([0-9a-fA-F]{8,64})>", meta)
        row.update({
            "pid": int(match.group("pid")),
            "tid": int(match.group("tid")),
            "start": int(match.group("start"), 16),
            "length": int(match.group("length"), 16),
            "map_pgoff": int(match.group("offset"), 0) if match.group("offset").lower().startswith("0x") else int(match.group("offset"), 16),
            "permissions": match.group("perm"),
            "path": match.group("path"),
            "build_id": build_ids[-1].lower() if build_ids else None,
        })
        row["pf_x"] = "x" in str(row["permissions"])
        if row["length"] <= 0 or row["start"] + row["length"] <= row["start"]:
            _fail("lifecycle_record", "MMAP2 interval is invalid", line=row["line"], start=row["start"], length=row["length"])
    elif kind == "PERF_RECORD_EXIT":
        match = EXIT_RE.search(descriptor)
        if match is None:
            _fail("lifecycle_record", "EXIT record is malformed", descriptor=descriptor)
        row.update({key: int(value) for key, value in match.groupdict().items()})
    elif kind in {"PERF_RECORD_FORK", "PERF_RECORD_CLONE"}:
        match = FORK_RE.search(descriptor)
        if match is None:
            _fail("lifecycle_record", "task-creation record is malformed", descriptor=descriptor)
        row.update({key: (int(value) if key != "kind" else value) for key, value in match.groupdict().items()})
    elif kind == "PERF_RECORD_THREAD_MAP":
        match = THREAD_MAP_RE.search(descriptor)
        if match is None:
            _fail("lifecycle_record", "THREAD_MAP record is malformed", descriptor=descriptor)
        row.update({"nr": int(match.group("nr")), "threads": [int(value) for value in re.findall(r"\d+", match.group("threads"))]})
    elif kind == "PERF_RECORD_SAMPLE":
        match = SAMPLE_DESCRIPTOR_RE.search(descriptor)
        if match is None:
            _fail("lifecycle_record", "SAMPLE descriptor is malformed", descriptor=descriptor)
        row.update({
            "misc": int(match.group("misc"), 16),
            "pid": int(match.group("pid")),
            "tid": int(match.group("tid")),
            "ip": int(match.group("ip"), 16),
            "period": int(match.group("period")),
        })
    elif kind == "PERF_RECORD_LOST":
        match = LOST_RE.search(descriptor)
        if match is None:
            _fail("lifecycle_record", "LOST record is malformed", descriptor=descriptor)
        row["lost"] = int(match.group("lost"))
    elif kind == "PERF_RECORD_LOST_SAMPLES":
        match = LOST_SAMPLES_RE.search(descriptor)
        if match is not None:
            row["lost"] = int(match.group("lost") or match.group("bare"))
    return row


def _u32(raw: bytes, offset: int) -> int | None:
    if offset < 0 or offset + 4 > len(raw):
        return None
    return int.from_bytes(raw[offset : offset + 4], "little")


def _u64(raw: bytes, offset: int) -> int | None:
    if offset < 0 or offset + 8 > len(raw):
        return None
    return int.from_bytes(raw[offset : offset + 8], "little")


def _nul_string(raw: bytes, offset: int) -> str | None:
    if offset < 0 or offset >= len(raw):
        return None
    end = raw.find(b"\0", offset)
    if end < 0:
        return None
    try:
        return raw[offset:end].decode("utf-8", errors="strict")
    except UnicodeDecodeError:
        return None


def _raw_event_type_for(kind: str) -> int | None:
    if kind == "PERF_RECORD_CLONE":
        # perf may print the type-7 FORK payload as CLONE when the clone
        # record originated from clone(2).  Type 11 is PERF_RECORD_AUX, not
        # a CLONE record; the raw header remains the authority.
        return 7
    return RECORD_TYPES.get(kind)


def _raw_core_fields(kind: str, raw: bytes, *, raw_misc: int = 0) -> dict[str, Any]:
    """Decode identity fields needed to cross-check text descriptors.

    This is intentionally narrower than the variable SAMPLE decoder.  It
    proves that COMM/MMAP2/EXIT/FORK/LOST text was derived from the same raw
    bytes without taking ownership of sample layout parsing.
    """
    if len(raw) < 8:
        _fail("raw_record_header", "raw event is shorter than perf header", bytes=len(raw))
    fields: dict[str, Any] = {}
    if kind in {"PERF_RECORD_COMM", "PERF_RECORD_MMAP2"}:
        pid, tid = _u32(raw, 8), _u32(raw, 12)
        if pid is None or tid is None:
            _fail("raw_metadata", "raw record lacks pid/tid fields", kind=kind)
        fields.update({"pid": pid, "tid": tid})
    if kind == "PERF_RECORD_COMM":
        name = _nul_string(raw, 16)
        if name is None:
            _fail("raw_metadata", "COMM raw name is missing or invalid")
        fields["comm"] = name
    elif kind == "PERF_RECORD_MMAP2":
        values = {"start": _u64(raw, 16), "length": _u64(raw, 24), "map_pgoff": _u64(raw, 32)}
        if any(value is None for value in values.values()):
            _fail("raw_metadata", "MMAP2 raw interval fields are truncated")
        fields.update(values)
        # With PERF_RECORD_MISC_MMAP_BUILD_ID, the old maj/min/ino union is
        # replaced at offset 40 by a u32 byte count and the build ID bytes;
        # prot/flags and filename retain their offsets at 64/68/72.
        if raw_misc & (1 << 14):
            build_size = _u32(raw, 40)
            if build_size is None or build_size <= 0 or build_size > 64 or 44 + build_size > len(raw):
                _fail("raw_metadata", "MMAP2 build-id union is invalid", build_id_size=build_size)
            fields["build_id_size"] = build_size
            fields["build_id"] = raw[44 : 44 + build_size].hex()
        fields["prot"] = _u32(raw, 64)
        fields["flags"] = _u32(raw, 68)
        name = _nul_string(raw, 72)
        if name is not None:
            fields["path"] = name
    elif kind == "PERF_RECORD_EXIT":
        values = {"pid": _u32(raw, 8), "ppid": _u32(raw, 12), "tid": _u32(raw, 16), "ptid": _u32(raw, 20)}
        if any(value is None for value in values.values()):
            _fail("raw_metadata", "EXIT raw identity fields are truncated")
        fields.update(values)
    elif kind in {"PERF_RECORD_FORK", "PERF_RECORD_CLONE"}:
        values = {"child_pid": _u32(raw, 8), "parent_pid": _u32(raw, 12), "child_tid": _u32(raw, 16), "parent_tid": _u32(raw, 20)}
        if any(value is None for value in values.values()):
            _fail("raw_metadata", "task-creation raw identity fields are truncated")
        fields.update(values)
    elif kind == "PERF_RECORD_LOST":
        lost = _u64(raw, 16)
        if lost is None:
            _fail("raw_metadata", "LOST raw count is truncated")
        fields["lost"] = lost
    elif kind == "PERF_RECORD_LOST_SAMPLES":
        lost = _u64(raw, 8)
        if lost is None:
            _fail("raw_metadata", "LOST_SAMPLES raw count is truncated")
        fields["lost"] = lost
    elif kind == "PERF_RECORD_THREAD_MAP":
        nr = _u64(raw, 8)
        if nr is None:
            _fail("raw_metadata", "THREAD_MAP raw count is truncated")
        fields["nr"] = nr
        if nr > 4096 or 16 + nr * 24 > len(raw):
            _fail("raw_metadata", "THREAD_MAP entries are truncated", nr=nr, bytes=len(raw))
        fields["threads"] = []
        fields["thread_names"] = []
        for index in range(nr):
            entry = 16 + index * 24
            thread = _u64(raw, entry)
            name = _nul_string(raw, entry + 8)
            if thread is None or name is None:
                _fail("raw_metadata", "THREAD_MAP entry is malformed", index=index)
            fields["threads"].append(thread)
            fields["thread_names"].append(name)
    return fields


def _crosscheck_text_raw(event: Mapping[str, Any], raw: bytes, *, raw_offset: int) -> dict[str, Any]:
    kind = str(event["kind"])
    header_type = int.from_bytes(raw[0:4], "little")
    header_misc = int.from_bytes(raw[4:6], "little")
    header_size = int.from_bytes(raw[6:8], "little")
    expected_type = _raw_event_type_for(kind)
    if header_size != len(raw):
        _fail("raw_metadata", "raw header size differs from framed bytes", kind=kind, header_size=header_size, bytes=len(raw), offset=raw_offset)
    if expected_type is not None and header_type != expected_type:
        _fail("raw_metadata", "raw header type differs from descriptor", kind=kind, raw_type=header_type, expected_type=expected_type, offset=raw_offset)
    fields = _raw_core_fields(kind, raw, raw_misc=header_misc)
    checks: dict[str, Any] = {"header_type": header_type, "header_misc": header_misc, "header_size": header_size, "fields": fields}
    # Compare only fields represented in both layers.  Build IDs are carried
    # by perf's textual MMAP2 decoration and are independently checked by the
    # build-id export; they are not duplicated in struct perf_record_mmap2.
    for name in ("pid", "tid", "comm", "start", "length", "map_pgoff", "path", "ppid", "ptid", "child_pid", "child_tid", "parent_pid", "parent_tid", "lost", "nr"):
        if name in event and name in fields and event[name] != fields[name]:
            _fail("raw_metadata", "descriptor field differs from raw metadata", kind=kind, field=name, text=event[name], raw=fields[name], offset=raw_offset)
    if kind == "PERF_RECORD_MMAP2" and "path" in fields and str(event.get("path")) != str(fields["path"]):
        _fail("raw_metadata", "MMAP2 path differs from raw metadata", text=event.get("path"), raw=fields["path"], offset=raw_offset)
    if kind == "PERF_RECORD_MMAP2" and event.get("build_id") and fields.get("build_id") is not None and str(event.get("build_id")).lower() != str(fields["build_id"]).lower():
        _fail("raw_metadata", "MMAP2 build ID differs from raw metadata", text=event.get("build_id"), raw=fields.get("build_id"), offset=raw_offset)
    if kind == "PERF_RECORD_MMAP2":
        text_build = str(event.get("build_id") or "").lower()
        raw_build = str(fields.get("build_id") or "").lower()
        build_flag = bool(header_misc & (1 << 14))
        if text_build and (not build_flag or not raw_build or int(fields.get("build_id_size", 0)) != 20):
            _fail("raw_metadata", "MMAP2 textual BuildID lacks the raw build-id union", text=text_build, raw=raw_build, misc=hex(header_misc), offset=raw_offset)
        if build_flag and not raw_build:
            _fail("raw_metadata", "MMAP2 build-id misc bit has no raw build-id union", misc=hex(header_misc), offset=raw_offset)
    if kind == "PERF_RECORD_COMM":
        raw_exec = bool(header_misc & (1 << 13))
        if bool(event.get("exec")) != raw_exec:
            _fail("raw_metadata", "COMM exec flag differs from raw misc", text=bool(event.get("exec")), raw=raw_exec, offset=raw_offset)
    # sample_id_all appends TID, TIME, and CPU (plus reserved/other fields) to
    # the variable COMM/MMAP2/task records.  Check the fields that are present
    # in this exact evsel; this catches a descriptor-only timestamp/CPU edit.
    sample_id_start: int | None = None
    if kind == "PERF_RECORD_COMM":
        name_end = raw.find(b"\0", 16)
        if name_end >= 0:
            sample_id_start = (name_end + 8) & ~7
    elif kind == "PERF_RECORD_MMAP2":
        name_end = raw.find(b"\0", 72)
        if name_end >= 0:
            sample_id_start = (name_end + 8) & ~7
    elif kind in {"PERF_RECORD_EXIT", "PERF_RECORD_FORK", "PERF_RECORD_CLONE"}:
        sample_id_start = 32
    if sample_id_start is not None and sample_id_start + 24 <= len(raw):
        sample_id_bytes = raw[sample_id_start : sample_id_start + 24]
        if sample_id_bytes == b"\0" * 24:
            # The first synthetic ``perf-exec`` COMM record is emitted before
            # perf has populated sample_id_all.  It is the only accepted
            # all-zero trailer; its descriptor must remain the exact
            # zero-time/non-exec bootstrap identity.  Every later lifecycle
            # record must carry a real trailer so descriptor edits cannot be
            # accepted without a raw-byte mismatch.
            bootstrap = (
                kind == "PERF_RECORD_COMM"
                and event.get("exec") is False
                and event.get("comm") == "perf-exec"
                and int(event.get("time_ns", -1)) == 0
                and int(event.get("cpu", -1)) == 0
            )
            if not bootstrap:
                _fail("raw_metadata", "all-zero sample-id trailer is only valid for the initial perf-exec COMM", kind=kind, offset=raw_offset)
            checks["sample_id"] = {"available": False, "bootstrap": True, "start": sample_id_start}
        else:
            sid_tid = _u32(raw, sample_id_start)
            sid_time = _u64(raw, sample_id_start + 8)
            sid_cpu_raw = _u32(raw, sample_id_start + 16)
            sid_cpu = -1 if sid_cpu_raw == 0xFFFFFFFF else sid_cpu_raw
            if sid_tid is not None and sid_tid != int(event.get("tid", sid_tid)):
                _fail("raw_metadata", "sample-id TID differs from descriptor", text=event.get("tid"), raw=sid_tid, offset=raw_offset)
            if sid_time is not None and sid_time != int(event.get("time_ns", sid_time)):
                _fail("raw_metadata", "sample-id timestamp differs from descriptor", text=event.get("time_ns"), raw=sid_time, offset=raw_offset)
            if sid_cpu is not None and sid_cpu != int(event.get("cpu", sid_cpu)):
                _fail("raw_metadata", "sample-id CPU differs from descriptor", text=event.get("cpu"), raw=sid_cpu, offset=raw_offset)
            checks["sample_id"] = {"available": True, "start": sample_id_start, "tid": sid_tid, "time_ns": sid_time, "cpu": sid_cpu}
    if kind == "PERF_RECORD_LOST" and "lost" in event and event["lost"] != fields.get("lost"):
        _fail("raw_metadata", "LOST count differs from raw metadata", text=event.get("lost"), raw=fields.get("lost"), offset=raw_offset)
    if kind == "PERF_RECORD_THREAD_MAP" and event.get("threads") != fields.get("threads"):
        _fail("raw_metadata", "THREAD_MAP thread IDs differ from raw metadata", text=event.get("threads"), raw=fields.get("threads"), offset=raw_offset)
    checks["sha256"] = hashlib.sha256(raw).hexdigest()
    checks["raw_offset"] = raw_offset
    checks["passed"] = True
    return checks


def parse_lifecycle_dump(
    path: Path,
    *,
    expected_pid: int,
    expected_cpu: int | None = None,
    require_raw_metadata: bool = True,
) -> dict[str, Any]:
    """Frame and classify all textual ``perf script -D`` record descriptors.

    The raw scanner remains authoritative for sample byte layout.  This pass
    only binds the timestamped record descriptors needed for phase assignment,
    and fails on malformed lifecycle records instead of dropping them.
    """
    path = Path(path).absolute()
    if not path.is_file():
        _fail("input_missing", "lifecycle dump is missing", path=str(path))
    records: list[dict[str, Any]] = []
    counts: collections.Counter[str] = collections.Counter()
    raw_sample_descriptors = 0
    raw_lost_sum = 0
    terminal_lost_values: list[int] = []
    raw_event_count = 0
    raw_metadata_count = 0
    current_raw: dict[str, Any] | None = None
    seen_offsets: set[int] = set()

    def finish_raw(line_no: int) -> None:
        nonlocal current_raw
        if current_raw is None:
            return
        if current_raw["written_count"] != current_raw["size"]:
            _fail("raw_record_truncated", "raw event hex payload does not cover declared record", line=line_no, declared_bytes=current_raw["size"], written_bytes=current_raw["written_count"])
        if current_raw.get("descriptor_seen") is not True and require_raw_metadata:
            _fail("raw_metadata", "raw event has no matching perf-script descriptor", line=current_raw["line"])
        current_raw = None

    with path.open("r", encoding="utf-8", errors="strict") as stream:
        for line_no, raw_line in enumerate(stream, 1):
            line = raw_line.rstrip("\n")
            raw_header = RAW_EVENT_RE.match(line)
            if raw_header is not None:
                finish_raw(line_no)
                size = int(raw_header.group("size"))
                if size < 8:
                    _fail("raw_record_header", "raw event is shorter than perf header", line=line_no, bytes=size)
                current_raw = {"line": line_no, "size": size, "data": bytearray(size), "written": bytearray(size), "written_count": 0, "descriptor_seen": False}
                raw_event_count += 1
                continue
            hex_match = RAW_HEX_RE.match(line)
            if hex_match is not None and current_raw is not None:
                offset = int(hex_match.group("offset"), 16)
                expected_line_bytes = min(16, current_raw["size"] - offset)
                if expected_line_bytes < 0:
                    _fail("raw_hex_bounds", "raw hex offset exceeds declared record", line=line_no, offset=offset, record_bytes=current_raw["size"])
                # The perf dump has an ASCII gutter after the 16-byte hex
                # region.  Consume exactly the bytes implied by the offset
                # and record size; never mine arbitrary hex-looking pairs
                # from that gutter.
                tokens = line.split(":", 1)[1].strip().split()
                if len(tokens) < expected_line_bytes or any(not re.fullmatch(r"[0-9a-fA-F]{2}", token) for token in tokens[:expected_line_bytes]):
                    _fail("raw_hex", "raw hex line has an incomplete or malformed byte region", line=line_no, offset=offset, expected=expected_line_bytes)
                payload = bytes.fromhex("".join(tokens[:expected_line_bytes]))
                end = offset + len(payload)
                if offset < 0 or end > current_raw["size"]:
                    _fail("raw_hex_bounds", "raw hex payload exceeds declared record", line=line_no, offset=offset, payload_bytes=len(payload), record_bytes=current_raw["size"])
                if any(current_raw["written"][offset:end]):
                    _fail("raw_hex_overlap", "raw hex payload overlaps an earlier payload", line=line_no, offset=offset)
                current_raw["data"][offset:end] = payload
                current_raw["written"][offset:end] = b"\1" * len(payload)
                current_raw["written_count"] += len(payload)
                continue
            prefix = _event_prefix(raw_line)
            if prefix is None:
                continue
            prefix["line"] = line_no
            event = _parse_event(prefix)
            event_offset = int(event["record_offset"])
            if event_offset in seen_offsets:
                _fail("raw_metadata", "duplicate perf-data record offset", offset=event_offset, line=line_no)
            seen_offsets.add(event_offset)
            if current_raw is not None:
                finish_offset = int(event["record_offset"])
                if current_raw["written_count"] != current_raw["size"]:
                    _fail("raw_record_truncated", "descriptor precedes complete raw event bytes", line=line_no, declared_bytes=current_raw["size"], written_bytes=current_raw["written_count"])
                event["raw_metadata"] = _crosscheck_text_raw(event, bytes(current_raw["data"]), raw_offset=finish_offset)
                current_raw["descriptor_seen"] = True
                # The scanner retains the source line of each sample's
                # descriptor.  Keep the corresponding raw-event header line
                # on the lifecycle event as the stable bridge to the framed
                # physical record; the perf-data offset is retained by the
                # descriptor and is checked below when the two views join.
                event["raw_event_line"] = current_raw["line"]
                raw_metadata_count += 1
            elif require_raw_metadata:
                _fail("raw_metadata", "perf record descriptor has no raw event bytes", line=line_no)
            event["record_sequence"] = len(records) + 1
            records.append(event)
            kind = str(event["kind"])
            counts[kind] += 1
            if kind == "PERF_RECORD_SAMPLE":
                raw_sample_descriptors += 1
            if kind == "PERF_RECORD_LOST":
                raw_lost_sum += int(event.get("lost", 0))
            if kind == "PERF_RECORD_LOST_SAMPLES" and "lost" in event:
                terminal_lost_values.append(int(event["lost"]))
            if expected_cpu is not None and event["cpu"] not in {expected_cpu, 0, -1}:
                _fail("lifecycle_cpu", "lifecycle record CPU differs from expected", line=line_no, actual=event["cpu"], expected=expected_cpu)
    finish_raw(raw_event_count + len(records) + 1)
    if require_raw_metadata and raw_metadata_count != raw_event_count:
        _fail("raw_metadata", "framed raw event count differs from descriptor count", raw_events=raw_event_count, descriptors=raw_metadata_count)
    if not records:
        _fail("lifecycle_record", "no framed perf records were found", path=str(path))
    return {
        "schema": "litchi.performance.0978.lifecycle_stream.v1",
        "status": "complete",
        "dump": file_binding(path),
        "expected_pid": expected_pid,
        "expected_cpu": expected_cpu,
        "record_count": len(records),
        "raw_event_count": raw_event_count,
        "raw_metadata_count": raw_metadata_count,
        "raw_metadata_crosscheck": bool(require_raw_metadata and raw_metadata_count == raw_event_count),
        "unique_record_offsets": len(seen_offsets) == raw_event_count,
        "record_kind_counts": dict(sorted(counts.items())),
        "sample_descriptor_count": raw_sample_descriptors,
        "sample_records": [event for event in records if event.get("kind") == "PERF_RECORD_SAMPLE"],
        "raw_lost_sum": raw_lost_sum,
        "raw_lost_record_count": counts.get("PERF_RECORD_LOST", 0),
        "terminal_lost_values": terminal_lost_values,
        "records": records,
    }


def _row_int(row: Mapping[str, Any], name: str, *, index: int) -> int:
    value = row.get(name)
    if type(value) is not int:
        _fail("sample_row", "sample row field is not an integer", field=name, index=index, value=value)
    return value


def _interval_owner(maps: Sequence[Mapping[str, Any]], ip: int, *, record_sequence: int, comm_sequence: int | None) -> dict[str, Any] | None:
    active: list[Mapping[str, Any]] = []
    for item in maps:
        item_sequence = int(item.get("record_sequence", -1))
        if comm_sequence is not None and item_sequence < comm_sequence:
            continue
        if item_sequence > record_sequence:
            continue
        start = item.get("start")
        length = item.get("length")
        if type(start) is int and type(length) is int and start <= ip < start + length:
            active.append(item)
    if not active:
        return None
    # Latest map wins when a process remaps an interval.  This is deterministic
    # and keeps the physical owner separate from target-source eligibility.
    return max(active, key=lambda item: int(item.get("record_sequence", -1)))


def _target_maps(events: Sequence[Mapping[str, Any]], *, target_path: str, target_build_id: str, expected_pid: int) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for event in events:
        if event.get("kind") != "PERF_RECORD_MMAP2":
            continue
        if event.get("pid") != expected_pid or event.get("tid") != expected_pid:
            continue
        if event.get("path") != target_path or str(event.get("build_id") or "").lower() != target_build_id.lower():
            continue
        row = dict(event)
        row["target_path_match"] = True
        row["target_build_id_match"] = True
        row["pf_x"] = "x" in str(row.get("permissions", ""))
        result.append(row)
    return result


def _phase_for(record_sequence: int, *, comm_sequence: int | None, first_pf_x_sequence: int | None, exit_sequence: int | None) -> str:
    if comm_sequence is None or record_sequence < comm_sequence:
        return "startup"
    if exit_sequence is not None and record_sequence >= exit_sequence:
        return "post_target"
    if first_pf_x_sequence is None or record_sequence < first_pf_x_sequence:
        return "transition"
    return "target_execution"


def _as_address(value: Any, *, field: str) -> int:
    if type(value) is int and value >= 0:
        return value
    if isinstance(value, str):
        try:
            parsed = int(value, 0)
        except ValueError:
            parsed = -1
        if parsed >= 0:
            return parsed
    _fail("owner_binding", "physical leaf address is missing or invalid", field=field, value=value)


def _owner_from_attribution(
    attribution: Mapping[str, Any] | None,
    *,
    sample: Mapping[str, Any],
    active_maps: Sequence[Mapping[str, Any]],
    record_sequence: int,
    comm_sequence: int,
    target_path: str,
    target_build_id: str,
) -> tuple[str | None, bool, bool, dict[str, Any] | None]:
    """Validate and return a physical DSO owner against active MMAP2 state."""
    if not isinstance(attribution, Mapping):
        return None, False, False, None
    leaf = attribution.get("physical_leaf")
    if not isinstance(leaf, Mapping):
        # A marker ancestor can be resolved while the physical frame-0 leaf
        # is unknown (for example, a valid empty/variable-layout sample or a
        # frame that the offline decoder cannot resolve).  Preserve that
        # marker claim so build_ledger can account for U_marker; never invent
        # an owner or silently drop the row.  An explicit physical_unknown
        # flag makes this distinction auditable for non-marker rows.
        if attribution.get("physical_unknown") is True or isinstance(attribution.get("marker_frame"), Mapping) or isinstance(attribution.get("marker"), Mapping):
            declared_unknown_source = attribution.get("target_source_eligible")
            if declared_unknown_source is not None and declared_unknown_source is not False:
                _fail("owner_binding", "unknown physical leaf cannot claim target-source eligibility", sample_index=sample.get("index"), declared=declared_unknown_source)
            return None, False, False, None
        _fail("owner_binding", "attribution owner lacks a physical_leaf or explicit unknown binding", sample_index=sample.get("index"))
    dso_path = leaf.get("dso_path")
    build_id = str(leaf.get("build_id") or "").lower()
    symbol = leaf.get("symbol")
    if not isinstance(dso_path, str) or not dso_path or not re.fullmatch(r"[0-9a-fA-F]{8,64}", build_id) or not isinstance(symbol, str) or not symbol:
        _fail("owner_binding", "physical leaf lacks exact DSO path/build ID/symbol", sample_index=sample.get("index"), leaf=dict(leaf))
    leaf_ip = _as_address(leaf.get("ip"), field="physical_leaf.ip")
    # Maps before the target COMM_EXEC are valid for startup samples.  The
    # exec cutoff is only a constraint once the sample itself is at/after the
    # target exec; applying it to startup would reject a real pre-exec map.
    map_cutoff = comm_sequence if record_sequence >= comm_sequence else None
    map_row = _interval_owner(active_maps, leaf_ip, record_sequence=record_sequence, comm_sequence=map_cutoff)
    if map_row is None:
        _fail("owner_binding", "physical leaf IP is not covered by an active MMAP2 map", sample_index=sample.get("index"), ip=hex(leaf_ip), dso_path=dso_path)
    if map_row.get("path") != dso_path or str(map_row.get("build_id") or "").lower() != build_id:
        _fail("owner_binding", "physical leaf DSO differs from active map", sample_index=sample.get("index"), leaf_path=dso_path, map_path=map_row.get("path"), leaf_build_id=build_id, map_build_id=map_row.get("build_id"))
    declared_seq = leaf.get("map_record_sequence")
    if type(declared_seq) is not int or declared_seq != int(map_row.get("record_sequence")):
        _fail("owner_binding", "physical leaf map sequence is not the active map", sample_index=sample.get("index"), declared=declared_seq, active=map_row.get("record_sequence"))
    owner = f"{dso_path}::{symbol}"
    derived_target_source = dso_path == target_path and build_id == target_build_id.lower() and bool(map_row.get("pf_x", "x" in str(map_row.get("permissions", ""))))
    declared_target_source = attribution.get("target_source_eligible")
    if type(declared_target_source) is not bool or declared_target_source != derived_target_source:
        _fail("owner_binding", "target-source eligibility disagrees with physical DSO identity", sample_index=sample.get("index"), declared=declared_target_source, derived=derived_target_source)
    return owner, derived_target_source, True, dict(map_row)


def build_ledger(
    lifecycle: Mapping[str, Any],
    samples: Sequence[Mapping[str, Any]],
    *,
    expected_pid: int,
    expected_cpu: int,
    expected_period: int,
    target_path: str,
    target_build_id: str,
    target_comm: str,
    terminal_lost_samples: int | None,
    marker: str | None = None,
    marker_dso: str | None = None,
    attributions: Mapping[int, Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Join scanner rows and optional corrected-script attribution.

    ``attributions`` is keyed by the scanner's one-based sample index.  It is
    intentionally optional: a valid empty or unresolved sample enters the
    explicit unknown bucket rather than being omitted or rejected.
    """
    if lifecycle.get("status") != "complete":
        _fail("lifecycle_binding", "lifecycle stream is not complete")
    records = lifecycle.get("records")
    if not isinstance(records, list) or not records:
        _fail("lifecycle_binding", "lifecycle records are missing")
    if lifecycle.get("expected_pid") != expected_pid:
        _fail("lifecycle_binding", "lifecycle PID binding differs", actual=lifecycle.get("expected_pid"), expected=expected_pid)
    if type(terminal_lost_samples) is not int or terminal_lost_samples < 0:
        _fail("loss_binding", "terminal PERF_RECORD_LOST_SAMPLES value is missing or invalid", value=terminal_lost_samples)
    stream_terminal_losses = lifecycle.get("terminal_lost_values", [])
    if stream_terminal_losses:
        if any(type(value) is not int or value < 0 for value in stream_terminal_losses):
            _fail("loss_binding", "lifecycle terminal loss value is invalid", values=stream_terminal_losses)
        if any(value != terminal_lost_samples for value in stream_terminal_losses):
            _fail("loss_binding", "lifecycle and terminal LOST_SAMPLES values disagree", lifecycle=stream_terminal_losses, terminal=terminal_lost_samples)

    comm_exec_events = [
        event for event in records
        if event.get("kind") == "PERF_RECORD_COMM" and event.get("pid") == expected_pid
        and event.get("tid") == expected_pid and event.get("exec") is True
    ]
    target_comm_events = [event for event in comm_exec_events if str(event.get("comm", "")) == target_comm]
    if len(target_comm_events) != 1:
        _fail("lifecycle_identity", "target COMM_EXEC is absent or ambiguous", count=len(target_comm_events), target_comm=target_comm)
    target_comm_event = target_comm_events[0]
    comm_sequence = int(target_comm_event["record_sequence"])
    target_maps = [event for event in _target_maps(records, target_path=target_path, target_build_id=target_build_id, expected_pid=expected_pid) if int(event.get("record_sequence", -1)) >= comm_sequence]
    all_maps = [event for event in records if event.get("kind") == "PERF_RECORD_MMAP2" and event.get("pid") == expected_pid and event.get("tid") == expected_pid]
    pf_x_maps = [event for event in target_maps if event.get("pf_x") is True]
    exit_events = [event for event in records if event.get("kind") == "PERF_RECORD_EXIT" and event.get("pid") == expected_pid and event.get("tid") == expected_pid]
    if len(exit_events) != 1:
        _fail("lifecycle_identity", "target EXIT is absent or ambiguous", count=len(exit_events), expected_pid=expected_pid)
    exit_event = exit_events[0]
    exit_sequence = int(exit_event["record_sequence"])
    if exit_sequence <= comm_sequence:
        _fail("lifecycle_identity", "target EXIT precedes target COMM_EXEC", comm_sequence=comm_sequence, exit_sequence=exit_sequence)
    comm_exec_ns = int(target_comm_event["time_ns"])
    first_pf_x_event = min(pf_x_maps, key=lambda event: int(event.get("record_sequence", 0))) if pf_x_maps else None
    first_pf_x_ns = int(first_pf_x_event["time_ns"]) if first_pf_x_event is not None else None
    exit_ns = int(exit_event["time_ns"])
    if first_pf_x_event is not None and int(first_pf_x_event["record_sequence"]) <= comm_sequence:
        _fail("lifecycle_identity", "target PF_X map precedes target COMM_EXEC", comm_sequence=comm_sequence, map_sequence=first_pf_x_event.get("record_sequence"))
    if first_pf_x_event is not None and int(first_pf_x_event["record_sequence"]) >= exit_sequence:
        _fail("lifecycle_identity", "target PF_X map follows target EXIT", map_sequence=first_pf_x_event.get("record_sequence"), exit_sequence=exit_sequence)
    if first_pf_x_event is not None and int(first_pf_x_event["time_ns"]) < comm_exec_ns:
        _fail("lifecycle_identity", "target PF_X map timestamp precedes target COMM_EXEC", comm_time=comm_exec_ns, map_time=first_pf_x_event.get("time_ns"))
    if exit_ns < comm_exec_ns:
        _fail("lifecycle_identity", "target EXIT timestamp precedes target COMM_EXEC", comm_time=comm_exec_ns, exit_time=exit_ns)
    boundary_events = [target_comm_event, *sorted(target_maps, key=lambda event: int(event["record_sequence"])), exit_event]
    last_boundary_time = -1
    for event in boundary_events:
        event_time = int(event["time_ns"])
        if event_time < last_boundary_time:
            _fail("lifecycle_order", "target lifecycle timestamps regress in record sequence", sequence=event.get("record_sequence"), time_ns=event_time, previous_time_ns=last_boundary_time)
        last_boundary_time = event_time
    if not pf_x_maps:
        _fail("lifecycle_identity", "target executable PF_X MMAP2 is absent", target_path=target_path, target_build_id=target_build_id)

    # Match all scanner samples by retained one-based index.  A stream with a
    # missing or duplicated index is a structural failure, never a best effort.
    indexed: dict[int, Mapping[str, Any]] = {}
    for row in samples:
        index = _row_int(row, "index", index=-1)
        if index in indexed:
            _fail("sample_identity", "duplicate scanner sample index", index=index)
        indexed[index] = row
    expected_indices = list(range(1, len(samples) + 1))
    if sorted(indexed) != expected_indices:
        _fail("sample_identity", "scanner sample indices are not contiguous", first=sorted(indexed)[:4], count=len(indexed))
    if int(lifecycle.get("sample_descriptor_count", len(samples))) != len(samples):
        _fail("sample_conservation", "lifecycle sample count differs from scanner sample count", lifecycle=lifecycle.get("sample_descriptor_count"), scanner=len(samples))
    sample_records = lifecycle.get("sample_records")
    if not isinstance(sample_records, list) or len(sample_records) != len(samples):
        _fail("sample_conservation", "lifecycle sample identity rows are missing or incomplete", lifecycle_count=len(sample_records) if isinstance(sample_records, list) else None, scanner=len(samples))

    phase_periods: collections.Counter[str] = collections.Counter()
    phase_samples: collections.Counter[str] = collections.Counter()
    target_known: collections.Counter[str] = collections.Counter()
    marker_known: collections.Counter[str] = collections.Counter()
    target_source: collections.Counter[str] = collections.Counter()
    samples_out: list[dict[str, Any]] = []
    d_period = 0
    m_period = 0
    target_period = 0
    target_unknown_period = 0
    marker_unknown_period = 0
    marker_count = 0

    for index in expected_indices:
        row = indexed[index]
        descriptor = sample_records[index - 1]
        descriptor_sequence = int(descriptor.get("record_sequence", -1))
        if descriptor_sequence <= 0:
            _fail("sample_identity", "lifecycle sample lacks record sequence", index=index)
        for field in ("time_ns", "pid", "tid", "period"):
            if _row_int(row, field, index=index) != _row_int(descriptor, field, index=index):
                _fail("sample_identity", "raw scanner and lifecycle sample identities differ", index=index, field=field, scanner=row.get(field), lifecycle=descriptor.get(field))
        if _row_int(row, "cpu", index=index) != int(descriptor.get("cpu", -2)):
            _fail("sample_identity", "raw scanner and lifecycle sample CPU identities differ", index=index, scanner=row.get("cpu"), lifecycle=descriptor.get("cpu"))
        if row.get("ip") != descriptor.get("ip"):
            _fail("sample_identity", "raw scanner and lifecycle sample IP identities differ", index=index, scanner=row.get("ip"), lifecycle=descriptor.get("ip"))
        if row.get("raw_bytes") != descriptor.get("record_bytes"):
            _fail("sample_identity", "raw scanner and lifecycle sample record sizes differ", index=index, scanner=row.get("raw_bytes"), lifecycle=descriptor.get("record_bytes"))
        if row.get("dump_line") is not None and row.get("dump_line") != descriptor.get("line"):
            _fail("sample_identity", "raw scanner and lifecycle sample descriptor lines differ", index=index, scanner=row.get("dump_line"), lifecycle=descriptor.get("line"))
        if row.get("raw_event_line") is None or row.get("raw_event_line") != descriptor.get("raw_event_line"):
            _fail("sample_identity", "raw scanner and lifecycle sample raw-event lines differ", index=index, scanner=row.get("raw_event_line"), lifecycle=descriptor.get("raw_event_line"))
        if row.get("physical_offset") is None or row.get("physical_offset") != descriptor.get("record_offset"):
            _fail("sample_identity", "raw scanner and lifecycle sample physical offsets differ", index=index, scanner=row.get("physical_offset"), lifecycle=descriptor.get("record_offset"))
        if descriptor.get("record_offset") is None or descriptor.get("raw_metadata", {}).get("raw_offset") != descriptor.get("record_offset"):
            _fail("sample_identity", "lifecycle sample raw physical offset is missing or inconsistent", index=index, descriptor_offset=descriptor.get("record_offset"), metadata_offset=descriptor.get("raw_metadata", {}).get("raw_offset"))
        period = _row_int(row, "period", index=index)
        time_ns = _row_int(row, "time_ns", index=index)
        pid = _row_int(row, "pid", index=index)
        tid = _row_int(row, "tid", index=index)
        cpu = _row_int(row, "cpu", index=index)
        if period != expected_period or period <= 0:
            _fail("sample_contract", "sample period differs from bound event", index=index, period=period, expected=expected_period)
        if pid != expected_pid or tid != expected_pid or cpu != expected_cpu:
            _fail("sample_contract", "sample PID/TID/CPU differs from bound child", index=index, pid=pid, tid=tid, cpu=cpu, expected_pid=expected_pid, expected_cpu=expected_cpu)
        if descriptor_sequence >= comm_sequence and time_ns < comm_exec_ns:
            _fail("lifecycle_order", "sample timestamp precedes target COMM_EXEC after the exec sequence", index=index, sample_time_ns=time_ns, comm_time_ns=comm_exec_ns, record_sequence=descriptor_sequence)
        if descriptor_sequence >= int(first_pf_x_event["record_sequence"]) and time_ns < first_pf_x_ns:
            _fail("lifecycle_order", "sample timestamp precedes target PF_X map after the map sequence", index=index, sample_time_ns=time_ns, map_time_ns=first_pf_x_ns, record_sequence=descriptor_sequence)
        if descriptor_sequence >= exit_sequence and time_ns < exit_ns:
            _fail("lifecycle_order", "sample timestamp precedes target EXIT after the exit sequence", index=index, sample_time_ns=time_ns, exit_time_ns=exit_ns, record_sequence=descriptor_sequence)
        phase = _phase_for(descriptor_sequence, comm_sequence=comm_sequence, first_pf_x_sequence=int(first_pf_x_event["record_sequence"]), exit_sequence=exit_sequence)
        mode = str(row.get("mode", "unknown"))
        attribution = attributions.get(index) if isinstance(attributions, Mapping) else None
        owner, target_source_eligible, resolved_leaf, active_map = _owner_from_attribution(
            attribution,
            sample=row,
            active_maps=all_maps,
            record_sequence=descriptor_sequence,
            comm_sequence=comm_sequence,
            target_path=target_path,
            target_build_id=target_build_id,
        )
        # Non-user mode, missing user ancestry, an invalid/unknown map, or no
        # physical leaf is explicit unknown.  A valid zero-stack row is
        # therefore retained and classified, never filtered by a quality gate.
        unknown_reason: str | None = None
        if phase != "target_execution":
            unknown_reason = f"phase_{phase}"
        elif mode != "user":
            unknown_reason = f"mode_{mode}"
        elif not resolved_leaf:
            unknown_reason = "no_defensible_physical_leaf"
        elif attribution is not None and active_map is None:
            unknown_reason = "invalid_map_binding"
        if phase == "target_execution":
            target_period += period
            if owner is None or unknown_reason is not None:
                target_unknown_period += period
            else:
                target_known[owner] += period
                if target_source_eligible:
                    target_source[owner] += period
        is_marker = False
        if isinstance(attribution, Mapping):
            marker_info = attribution.get("marker_frame")
            if marker_info is None and isinstance(attribution.get("marker"), Mapping):
                marker_info = attribution.get("marker")
            marker_claim = bool(attribution.get("marker") is True) or isinstance(marker_info, Mapping)
            if marker_claim:
                if marker is None or marker_dso is None:
                    _fail("marker_binding", "marker claim has no exact marker symbol/DSO binding", sample_index=index)
                if phase != "target_execution":
                    _fail("marker_binding", "marker claim occurs outside target execution phase", sample_index=index, phase=phase)
                if not isinstance(marker_info, Mapping) or marker_info.get("present") is not True:
                    _fail("marker_binding", "marker claim lacks an explicit present marker frame", sample_index=index)
                if marker_info.get("symbol") != marker or marker_info.get("dso_path") != marker_dso or str(marker_info.get("build_id") or "").lower() != target_build_id.lower() or marker_info.get("resolved") is not True:
                    _fail("marker_binding", "marker frame identity differs from exact bound marker DSO/build ID/symbol", sample_index=index, marker_frame=dict(marker_info))
                marker_ip = _as_address(marker_info.get("ip"), field="marker_frame.ip")
                marker_cutoff = comm_sequence if descriptor_sequence >= comm_sequence else None
                marker_map = _interval_owner(all_maps, marker_ip, record_sequence=descriptor_sequence, comm_sequence=marker_cutoff)
                if marker_map is None:
                    _fail("marker_binding", "marker IP is not covered by an active MMAP2 map", sample_index=index, ip=hex(marker_ip))
                marker_map_exec = bool(marker_map.get("pf_x", "x" in str(marker_map.get("permissions", ""))))
                if marker_map.get("path") != marker_dso or str(marker_map.get("build_id") or "").lower() != target_build_id.lower() or not marker_map_exec:
                    _fail("marker_binding", "marker IP active map differs from exact bound marker DSO/build ID", sample_index=index, ip=hex(marker_ip), map_path=marker_map.get("path"), map_build_id=marker_map.get("build_id"))
                declared_marker_seq = marker_info.get("map_record_sequence")
                if declared_marker_seq is not None and (type(declared_marker_seq) is not int or declared_marker_seq != int(marker_map.get("record_sequence"))):
                    _fail("marker_binding", "marker IP map sequence is not the active map", sample_index=index, declared=declared_marker_seq, active=marker_map.get("record_sequence"))
                is_marker = True
        if is_marker:
            marker_count += 1
            m_period += period
            if owner is None or unknown_reason is not None:
                marker_unknown_period += period
            else:
                marker_known[owner] += period
        d_period += period
        phase_periods[phase] += period
        phase_samples[phase] += 1
        sample_out = {
            "index": index,
            "raw_identity": {
                "index": index,
                "dump_line": row.get("dump_line"),
                "raw_event_line": row.get("raw_event_line"),
                "physical_offset": row.get("physical_offset"),
                "record_offset": descriptor.get("record_offset"),
                "record_sequence": descriptor_sequence,
                "raw_bytes": row.get("raw_bytes"),
                "time_ns": time_ns,
                "pid": pid,
                "tid": tid,
                "cpu": cpu,
                "period": period,
            },
            "phase": phase,
            "mode": mode,
            "marker": is_marker,
            "physical_owner": owner,
            "target_source_eligible": bool(target_source_eligible and owner is not None and unknown_reason is None),
            "unknown": unknown_reason is not None,
            "unknown_reason": unknown_reason,
            "valid_empty_layout": row.get("callchain_nr") == 0 and not bool(row.get("regs_available", False)) and row.get("stack_size") == 0,
        }
        samples_out.append(sample_out)

    if d_period < m_period:
        _fail("conservation", "marker period exceeds lifetime period", D=d_period, M=m_period)
    r_period = d_period - m_period
    phase_sum = sum(phase_periods.values())
    if phase_sum != d_period:
        _fail("conservation", "phase periods do not conserve D", phase_sum=phase_sum, D=d_period)
    if sum(target_known.values()) + target_unknown_period != target_period:
        _fail("conservation", "target known/unknown periods do not conserve T", known=sum(target_known.values()), unknown=target_unknown_period, T=target_period)
    if sum(marker_known.values()) + marker_unknown_period != m_period:
        _fail("conservation", "marker known/unknown periods do not conserve M", known=sum(marker_known.values()), unknown=marker_unknown_period, M=m_period)

    return {
        "schema": SCHEMA,
        "status": "complete",
        "target": {"pid": expected_pid, "path": target_path, "build_id": target_build_id.lower(), "comm": target_comm, "comm_exec_time_ns": comm_exec_ns, "first_pf_x_time_ns": first_pf_x_ns, "exit_time_ns": exit_ns, "target_pf_x_maps": pf_x_maps},
        "lifecycle": {
            "record_count": lifecycle.get("record_count"),
            "record_kind_counts": lifecycle.get("record_kind_counts"),
            "raw_lost_sum": lifecycle.get("raw_lost_sum", 0),
            "raw_lost_record_count": lifecycle.get("raw_lost_record_count", 0),
            "terminal_lost_values": lifecycle.get("terminal_lost_values", []),
            "terminal_lost_samples": terminal_lost_samples,
            "comm_exec_count": len(comm_exec_events),
            "target_comm_exec_count": len(target_comm_events),
            "target_mmap2_count": len(target_maps),
            "target_pf_x_count": len(pf_x_maps),
            "target_exit_count": len(exit_events),
            "all_records_framed": True,
        },
        "ledger": {
            "D": d_period,
            "M": m_period,
            "R": r_period,
            "T": target_period,
            "U_target": target_unknown_period,
            "U_marker": marker_unknown_period,
            "marker_sample_count": marker_count,
            "phase_periods": dict(sorted(phase_periods.items())),
            "phase_sample_counts": dict(sorted(phase_samples.items())),
            "target_known_owner_period": dict(sorted(target_known.items())),
            "target_source_eligible_period": dict(sorted(target_source.items())),
            "marker_known_owner_period": dict(sorted(marker_known.items())),
            "conservation": {
                "D_equals_phase_sum": phase_sum == d_period,
                "R_equals_D_minus_M": r_period == d_period - m_period,
                "T_equals_known_plus_U_target": sum(target_known.values()) + target_unknown_period == target_period,
                "M_equals_known_plus_U_marker": sum(marker_known.values()) + marker_unknown_period == m_period,
            },
        },
        "gates": {
            "raw_loss_zero": lifecycle.get("raw_lost_sum", 0) == 0,
            "terminal_loss_zero": terminal_lost_samples == 0,
            "minimum_marker_samples": marker_count,
            "valid_empty_is_unknown": True,
            "all_sample_empty_gate": False,
            "passed_structural_conservation": True,
        },
        "samples": samples_out,
    }


__all__ = ["LedgerFailure", "SCHEMA", "EVENT_SCHEMA", "file_binding", "sha256", "parse_lifecycle_dump", "build_ledger"]
