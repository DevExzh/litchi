#!/usr/bin/env python3
"""Variable-layout offline ``perf script -D`` sample scanner for 0978.

The 0973 scanner assumed that every sample had a non-zero register dump and a
full user stack.  Linux perf's ABI permits ``PERF_SAMPLE_REGS_ABI_NONE`` and a
zero ``PERF_SAMPLE_STACK_USER`` size, so a valid sample can be only 72 bytes.
This scanner follows the sample-type order from the bound ``perf evlist -v``
record and consumes every raw record and every decoded sample line.

The field order and register ABI are from the Linux UAPI
``include/uapi/linux/perf_event.h``.  The conditional no-register/no-dynamic-
stack behavior is also documented by upstream perf's evsel implementation:
https://raw.githubusercontent.com/torvalds/linux/v6.12/tools/perf/util/evsel.c
(the scanner records the source binding supplied by the caller; it does not
download or execute source code).
"""
from __future__ import annotations

import collections
import hashlib
import json
import re
import struct
from pathlib import Path
from typing import Any, Callable, Mapping


SCHEMA = "litchi.performance.0978.raw_stack.v1"
LINUX_PERF_EVENT_HEADER = Path("/usr/include/linux/perf_event.h")
# The installed header is bound at runtime.  The upstream source is cited for
# the conditional PERF_SAMPLE_REGS_USER/PERF_SAMPLE_STACK_USER rules that are
# not expressed by the UAPI header alone.
UPSTREAM_EVSEL_SOURCE = "https://raw.githubusercontent.com/torvalds/linux/v6.12/tools/perf/util/evsel.c"

PERF_SAMPLE_IP = 1 << 0
PERF_SAMPLE_TID = 1 << 1
PERF_SAMPLE_TIME = 1 << 2
PERF_SAMPLE_CALLCHAIN = 1 << 5
PERF_SAMPLE_CPU = 1 << 7
PERF_SAMPLE_PERIOD = 1 << 8
PERF_SAMPLE_REGS_USER = 1 << 12
PERF_SAMPLE_STACK_USER = 1 << 13
SUPPORTED_SAMPLE_MASK = (
    PERF_SAMPLE_IP
    | PERF_SAMPLE_TID
    | PERF_SAMPLE_TIME
    | PERF_SAMPLE_CALLCHAIN
    | PERF_SAMPLE_CPU
    | PERF_SAMPLE_PERIOD
    | PERF_SAMPLE_REGS_USER
    | PERF_SAMPLE_STACK_USER
)
SAMPLE_NAME_TO_BIT = {
    "IP": PERF_SAMPLE_IP,
    "TID": PERF_SAMPLE_TID,
    "TIME": PERF_SAMPLE_TIME,
    "CALLCHAIN": PERF_SAMPLE_CALLCHAIN,
    "CPU": PERF_SAMPLE_CPU,
    "PERIOD": PERF_SAMPLE_PERIOD,
    "REGS_USER": PERF_SAMPLE_REGS_USER,
    "STACK_USER": PERF_SAMPLE_STACK_USER,
}
SAMPLE_BIT_TO_NAME = {value: key for key, value in SAMPLE_NAME_TO_BIT.items()}
REGS_ABI_NAMES = {0: "NONE", 1: "32", 2: "64"}
CPUMODE_NAMES = {
    0: "unknown",
    1: "kernel",
    2: "user",
    3: "hypervisor",
    4: "guest_kernel",
    5: "guest_user",
}

RAW_EVENT_RE = re.compile(r"^\.\s+\.\.\. raw event: size (?P<size>\d+) bytes$")
RAW_LOCATION_RE = re.compile(
    r"^0x(?P<offset>[0-9a-fA-F]+)@(?P<path>.+?)\s+\[0x(?P<size>[0-9a-fA-F]+)\]: event: (?P<event>\d+)$"
)
HEX_LINE_RE = re.compile(r"^\.\s+(?P<offset>[0-9a-fA-F]+):\s+(?P<body>.*)$")
SAMPLE_LINE_RE = re.compile(
    r"^(?P<cpu>\d+)\s+(?P<time>\d+)\s+\S+\s+\[0x[0-9a-fA-F]+\]:\s+"
    r"PERF_RECORD_SAMPLE\(IP,\s*0x(?P<misc>[0-9a-fA-F]+)\):\s+"
    r"(?P<pid>\d+)/(?P<tid>\d+):\s+0x(?P<ip>[0-9a-fA-F]+)\s+period:\s+(?P<period>\d+)"
)


class ScanFailure(ValueError):
    """Structural raw-layout failure; no sample is silently skipped."""

    def __init__(self, code: str, message: str, details: Mapping[str, Any] | None = None):
        super().__init__(f"{code}: {message}")
        self.code = code
        self.message = message
        self.details = dict(details or {})

    def as_dict(self) -> dict[str, Any]:
        return {"status": "failed", "code": self.code, "message": self.message, "details": self.details}


def _fail(code: str, message: str, **details: Any) -> None:
    raise ScanFailure(code, message, details)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def file_binding(path: Path) -> dict[str, Any]:
    path = Path(path).absolute()
    if not path.is_file():
        _fail("input_missing", "required scanner input is missing", path=str(path))
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def _exact_attr(line: str, label: str) -> str:
    values = re.findall(rf"(?:^|,\s*|:\s*){re.escape(label)}:\s*(.*?)(?=,\s*|$)", line)
    if not values:
        _fail("evlist_contract", "required evlist attribute is missing", attribute=label)
    if len(values) != 1:
        _fail("evlist_contract", "evlist attribute is duplicated", attribute=label, occurrences=len(values))
    return values[0].strip()


def make_event_contract(
    *,
    sample_type_names: list[str] | tuple[str, ...],
    regs_user_mask: int,
    period: int,
    stack_cap: int,
    source: str = "synthetic-test",
) -> dict[str, Any]:
    """Build a test or already-parsed event contract with explicit provenance."""
    if not sample_type_names:
        _fail("evlist_contract", "sample type cannot be empty")
    mask = 0
    names: list[str] = []
    for name in sample_type_names:
        if name not in SAMPLE_NAME_TO_BIT:
            _fail("evlist_contract", "unsupported sample type", sample_type=name)
        if name in names:
            _fail("evlist_contract", "duplicate sample type", sample_type=name)
        names.append(name)
        mask |= SAMPLE_NAME_TO_BIT[name]
    if mask & ~SUPPORTED_SAMPLE_MASK:
        _fail("evlist_contract", "unsupported sample-type bits", mask=hex(mask))
    if type(regs_user_mask) is not int or regs_user_mask < 0 or regs_user_mask >= 1 << 64:
        _fail("evlist_contract", "register mask is outside u64", mask=regs_user_mask)
    if type(period) is not int or period <= 0 or type(stack_cap) is not int or stack_cap < 0:
        _fail("evlist_contract", "period or stack cap is invalid", period=period, stack_cap=stack_cap)
    return {
        "schema": "litchi.performance.0978.evlist_contract.v1",
        "source": source,
        "sample_type": "|".join(names),
        "sample_type_names": names,
        "sample_type_mask": mask,
        "regs_user_mask": regs_user_mask,
        "period": period,
        "stack_cap": stack_cap,
        "supported_sample_mask": SUPPORTED_SAMPLE_MASK,
        "linux_uapi": str(LINUX_PERF_EVENT_HEADER),
        "linux_uapi_sha256": sha256(LINUX_PERF_EVENT_HEADER) if LINUX_PERF_EVENT_HEADER.is_file() else None,
        "upstream_evsel_source": UPSTREAM_EVSEL_SOURCE,
    }


def parse_evlist_contract(text: str, *, expected_period: int | None = None, expected_stack_cap: int | None = None) -> dict[str, Any]:
    """Parse one exact ``perf evlist -v`` event and bind its layout fields."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(lines) != 1:
        _fail("evlist_contract", "evlist must contain exactly one event line", nonempty_lines=len(lines))
    line = lines[0]
    if not line.startswith("instructions:u:"):
        _fail("evlist_contract", "event is not exact instructions:u", line=line[:300])
    if re.search(r"(?:^|,\s*)(?:freq|inherit):\s*", line):
        _fail("evlist_contract", "frequency or inheritance mode is not admitted", line=line[:300])
    period_text = _exact_attr(line, "{ sample_period, sample_freq }")
    period_match = re.fullmatch(r"\d+", period_text)
    if period_match is None:
        _fail("evlist_contract", "sample period is not an integer", value=period_text)
    period = int(period_text)
    if expected_period is not None and period != expected_period:
        _fail("evlist_contract", "sample period differs from expected", actual=period, expected=expected_period)
    stack_text = _exact_attr(line, "sample_stack_user")
    if not re.fullmatch(r"\d+", stack_text):
        _fail("evlist_contract", "sample stack cap is not an integer", value=stack_text)
    stack_cap = int(stack_text)
    if expected_stack_cap is not None and stack_cap != expected_stack_cap:
        _fail("evlist_contract", "sample stack cap differs from expected", actual=stack_cap, expected=expected_stack_cap)
    regs_text = _exact_attr(line, "sample_regs_user")
    if not re.fullmatch(r"0[xX][0-9a-fA-F]+", regs_text):
        _fail("evlist_contract", "register mask is not hexadecimal", value=regs_text)
    regs_mask = int(regs_text, 16)
    sample_text = _exact_attr(line, "sample_type")
    sample_names = sample_text.split("|")
    if any(not name for name in sample_names):
        _fail("evlist_contract", "sample type contains an empty component", sample_type=sample_text)
    contract = make_event_contract(
        sample_type_names=sample_names, regs_user_mask=regs_mask,
        period=period, stack_cap=stack_cap, source="perf evlist -v",
    )
    required = {
        "type: 4 (cpu)", "size: 144", "config: 0xc0 (instructions)",
        "read_format: ID|LOST", "disabled: 1", "exclude_kernel: 1", "exclude_hv: 1",
        "enable_on_exec: 1", "task: 1", "sample_id_all: 1",
    }
    for attr in required:
        if not re.search(rf"(?:^|,\s*|:\s*){re.escape(attr)}(?=,\s*|$)", line):
            _fail("evlist_contract", "required event attribute is missing", attribute=attr)
    contract.update({"status": "pass", "line": line, "authority": "perf evlist -v"})
    return contract


def _u32(data: bytes | bytearray, offset: int) -> int:
    if offset < 0 or offset + 4 > len(data):
        _fail("sample_truncated", "u32 field exceeds sample record", offset=offset, record_bytes=len(data))
    return struct.unpack_from("<I", data, offset)[0]


def _u64(data: bytes | bytearray, offset: int) -> int:
    if offset < 0 or offset + 8 > len(data):
        _fail("sample_truncated", "u64 field exceeds sample record", offset=offset, record_bytes=len(data))
    return struct.unpack_from("<Q", data, offset)[0]


def _parse_sample_payload(
    data: bytes,
    *,
    contract: Mapping[str, Any],
    event: Mapping[str, Any],
    stack_cap: int,
) -> dict[str, Any]:
    """Parse one PERF_RECORD_SAMPLE payload without fixed offsets."""
    if len(data) < 8:
        _fail("sample_truncated", "sample record is shorter than its header", record_bytes=len(data))
    record_type, record_misc, record_size = struct.unpack_from("<IHH", data, 0)
    if record_type != 9:
        _fail("sample_record_type", "decoded sample line paired with non-sample raw record", record_type=record_type)
    if record_size != len(data):
        _fail("record_size_mismatch", "sample header size differs from raw event size", header_size=record_size, bytes=len(data))
    offset = 8
    values: dict[str, Any] = {}
    mask = int(contract["sample_type_mask"])
    regs_mask = int(contract["regs_user_mask"])
    # PERF_SAMPLE fields are serialized in the kernel's fixed enum order, not
    # in the order printed by perf evlist.  In this supported subset that is
    # IP,TID,TIME,CPU,PERIOD,CALLCHAIN,REGS_USER,STACK_USER.  Keeping this
    # order explicit is what makes variable CALLCHAIN/REGS/STACK sizes safe.
    for name, bit in (
        ("IP", PERF_SAMPLE_IP), ("TID", PERF_SAMPLE_TID), ("TIME", PERF_SAMPLE_TIME),
        ("CPU", PERF_SAMPLE_CPU), ("PERIOD", PERF_SAMPLE_PERIOD),
        ("CALLCHAIN", PERF_SAMPLE_CALLCHAIN), ("REGS_USER", PERF_SAMPLE_REGS_USER),
        ("STACK_USER", PERF_SAMPLE_STACK_USER),
    ):
        if not mask & bit:
            continue
        if name == "IP":
            values["ip"] = _u64(data, offset); offset += 8
        elif name == "TID":
            values["pid"] = _u32(data, offset); values["tid"] = _u32(data, offset + 4); offset += 8
        elif name == "TIME":
            values["time_ns"] = _u64(data, offset); offset += 8
        elif name == "CALLCHAIN":
            nr = _u64(data, offset); offset += 8
            if nr > (len(data) - offset) // 8:
                _fail("callchain_truncated", "callchain count exceeds sample record", nr=nr, offset=offset, record_bytes=len(data))
            values["callchain_nr"] = nr
            offset += nr * 8
        elif name == "CPU":
            values["cpu"] = _u32(data, offset); values["cpu_reserved"] = _u32(data, offset + 4); offset += 8
        elif name == "PERIOD":
            values["period"] = _u64(data, offset); offset += 8
        elif name == "REGS_USER":
            abi = _u64(data, offset); offset += 8
            if abi not in REGS_ABI_NAMES:
                _fail("regs_abi", "unsupported PERF_SAMPLE_REGS_ABI value", abi=abi)
            words = regs_mask.bit_count() if abi else 0
            bytes_needed = words * 8
            if offset + bytes_needed > len(data):
                _fail("regs_truncated", "register payload exceeds sample record", abi=abi, words=words, offset=offset, record_bytes=len(data))
            values.update({"regs_abi": abi, "regs_abi_name": REGS_ABI_NAMES[abi], "regs_words": words, "regs_available": abi != 0})
            offset += bytes_needed
        elif name == "STACK_USER":
            stack_size = _u64(data, offset); offset += 8
            if stack_size > stack_cap:
                values["stack_bounds_valid"] = False
            else:
                values["stack_bounds_valid"] = True
            values["stack_size"] = stack_size
            if stack_size == 0:
                # Upstream perf does not emit a dyn_size trailer when the
                # captured user-stack size is zero.  This is the 72-byte
                # child-02 shape that motivated 0978.
                values.update({"stack_dyn_size": None, "stack_available": False, "stack_data_bytes": 0})
            else:
                if offset + stack_size + 8 > len(data):
                    _fail("stack_truncated", "stack data or dyn_size exceeds sample record", stack_size=stack_size, offset=offset, record_bytes=len(data))
                offset += stack_size
                dyn_size = _u64(data, offset); offset += 8
                values.update({"stack_dyn_size": dyn_size, "stack_available": True, "stack_data_bytes": stack_size})
                if dyn_size > stack_size:
                    values["stack_bounds_valid"] = False
    if offset != len(data):
        _fail("sample_trailing_bytes", "sample record has bytes after bound sample fields", parsed_bytes=offset, record_bytes=len(data), trailing_bytes=len(data) - offset)
    if "stack_bounds_valid" not in values:
        values["stack_bounds_valid"] = True
    values["record_type"] = record_type
    values["record_misc"] = record_misc
    values["record_size"] = record_size
    values["mode_code"] = record_misc & 7
    values["mode"] = CPUMODE_NAMES.get(values["mode_code"], "unknown")
    values["payload_presence"] = {
        "callchain": values.get("callchain_nr", 0) > 0,
        "user_stack": values.get("stack_available", False),
        "user_regs": values.get("regs_available", False),
    }
    values["unwindability"] = (
        "callchain_and_stack" if values.get("callchain_nr", 0) > 0 and values.get("stack_available") else
        "callchain_only" if values.get("callchain_nr", 0) > 0 else
        "stack_only" if values.get("stack_available") else "none"
    )
    values["unwindability_claim"] = "descriptive_payload_presence_only"
    values["missing_unwindability"] = []
    if values.get("callchain_nr", 0) == 0:
        values["missing_unwindability"].append("callchain")
    if "STACK_USER" in contract["sample_type_names"] and not values.get("stack_available"):
        values["missing_unwindability"].append("user_stack")
    if "REGS_USER" in contract["sample_type_names"] and not values.get("regs_available", False):
        values["missing_unwindability"].append("user_regs")
    return values


def _parse_hex_line(line: str) -> tuple[int, bytes] | None:
    match = HEX_LINE_RE.match(line)
    if match is None:
        return None
    # perf script appends an ASCII rendering after two spaces.  Restrict the
    # byte field before that separator; otherwise an ASCII pair such as ``d5``
    # can be mistaken for a seventeenth byte and create a false overlap.
    byte_text = re.split(r"\s{2,}", match.group("body").strip(), maxsplit=1)[0]
    byte_tokens = byte_text.split() if byte_text else []
    if len(byte_tokens) > 16:
        _fail("raw_hex_parse", "raw hex line contains more than the fixed 16-byte column", byte_count=len(byte_tokens), line=line[:240])
    if any(re.fullmatch(r"[0-9a-fA-F]{2}", token) is None for token in byte_tokens):
        _fail("raw_hex_parse", "raw hex line contains a non-byte token", line=line[:240])
    return int(match.group("offset"), 16), bytes.fromhex("".join(byte_tokens))


def scan_dump(
    dump_path: Path,
    *,
    perf_data: Path | None = None,
    binary: Path | None = None,
    expected_pid: int,
    expected_cpu: int | None = None,
    expected_period: int | None = None,
    stack_cap: int | None = None,
    event_contract: Mapping[str, Any],
    sample_callback: Callable[[dict[str, Any]], None] | None = None,
    retain_samples: bool = True,
) -> dict[str, Any]:
    """Scan all raw records and sample lines in one streaming pass.

    Structural truncation, missing pairing, unsupported ABI, and trailing
    bytes raise :class:`ScanFailure`.  Period/PID/CPU and stack-bound
    mismatches are represented in the returned summary so callers can report
    all samples and all mode domains without silently dropping records.
    """
    path = Path(dump_path).absolute()
    contract = dict(event_contract)
    if int(contract.get("sample_type_mask", 0)) & ~SUPPORTED_SAMPLE_MASK:
        _fail("evlist_contract", "event contract contains unsupported sample bits", mask=hex(int(contract.get("sample_type_mask", 0))))
    if expected_period is not None and int(contract.get("period", -1)) != expected_period:
        _fail("evlist_contract", "event contract period differs from expected", actual=contract.get("period"), expected=expected_period)
    if stack_cap is None:
        stack_cap = int(contract.get("stack_cap", 0))
    elif int(contract.get("stack_cap", -1)) != stack_cap:
        _fail("evlist_contract", "event contract stack cap differs from expected", actual=contract.get("stack_cap"), expected=stack_cap)
    if expected_period is None:
        expected_period = int(contract["period"])
    if not path.is_file():
        _fail("input_missing", "raw dump is missing", path=str(path))

    raw_event_count = 0
    sample_line_count = 0
    paired_sample_count = 0
    raw_unconsumed_count = 0
    raw_sample_unpaired_count = 0
    raw_type_counts: collections.Counter[int] = collections.Counter()
    raw_size_counts: collections.Counter[int] = collections.Counter()
    mode_counts: collections.Counter[str] = collections.Counter()
    mode_missing_counts: collections.Counter[str] = collections.Counter()
    regs_abi_counts: collections.Counter[str] = collections.Counter()
    stack_shape_counts: collections.Counter[str] = collections.Counter()
    violations: list[dict[str, Any]] = []
    period_valid = True
    pid_tid_cpu_valid = True
    bounds_valid = True
    time_valid = True
    physical_identity_valid = True
    samples: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    pending_location: dict[str, Any] | None = None
    seen_locations: set[tuple[str, int]] = set()

    def bind_raw_header(line_number: int) -> None:
        """Decode and count the fixed raw header after its bytes are present."""
        if current is None or current.get("header_type") is not None:
            return
        if current["written_count"] != current["size"]:
            return
        header_type, header_misc, header_size = struct.unpack_from("<IHH", current["data"], 0)
        if header_size != current["size"]:
            _fail(
                "record_size_mismatch", "raw header size differs from raw event declaration",
                line=line_number, header_size=header_size, declared_bytes=current["size"],
            )
        current["header_type"] = header_type
        current["header_misc"] = header_misc
        current["header_size"] = header_size
        if current.get("location") is not None and current["location"]["event"] != header_type:
            _fail(
                "raw_record_identity", "location event type differs from raw header",
                line=line_number, location_event=current["location"]["event"], header_type=header_type,
            )
        current["header_counted"] = True
        raw_type_counts[header_type] += 1
        raw_size_counts[current["size"]] += 1

    def finalize_current(line_number: int) -> None:
        nonlocal current, raw_unconsumed_count, raw_sample_unpaired_count
        if current is None:
            return
        if current["written_count"] != current["size"]:
            _fail("raw_record_truncated", "raw event hex payload does not cover declared record", line=line_number, declared_bytes=current["size"], written_bytes=current["written_count"], record_type=current.get("header_type"))
        bind_raw_header(line_number)
        if not current["consumed"]:
            raw_unconsumed_count += 1
            if current.get("header_type") == 9:
                raw_sample_unpaired_count += 1
        current = None

    def emit(sample: dict[str, Any]) -> None:
        if sample_callback is not None:
            sample_callback(sample)
        if retain_samples:
            samples.append(sample)

    with path.open("r", encoding="utf-8", errors="strict") as stream:
        for line_number, raw_line in enumerate(stream, 1):
            line = raw_line.rstrip("\n")
            raw_match = RAW_EVENT_RE.match(line)
            if raw_match:
                finalize_current(line_number)
                size = int(raw_match.group("size"))
                if size < 8:
                    _fail("raw_record_header", "raw event is shorter than perf header", line=line_number, bytes=size)
                if size > 0xFFFF:
                    _fail("raw_record_header", "raw event exceeds perf_event_header.size", line=line_number, bytes=size)
                location = pending_location
                pending_location = None
                if location is not None:
                    if location["size"] != size:
                        _fail("raw_record_identity", "location and raw event sizes differ", line=line_number, location_size=location["size"], raw_size=size)
                    key = (location["path"], location["offset"])
                    if key in seen_locations:
                        physical_identity_valid = False
                        _fail("raw_record_identity", "physical raw-event offset is duplicated", line=line_number, path=location["path"], offset=hex(location["offset"]))
                    seen_locations.add(key)
                current = {
                    "line": line_number, "size": size, "data": bytearray(size),
                    "written": bytearray(size), "written_count": 0, "consumed": False,
                    "header_type": None,
                    "location": location,
                }
                raw_event_count += 1
                continue
            location_match = RAW_LOCATION_RE.match(line)
            if location_match:
                pending_location = {
                    "offset": int(location_match.group("offset"), 16),
                    "path": location_match.group("path"),
                    "size": int(location_match.group("size"), 16),
                    "event": int(location_match.group("event")),
                }
                continue
            hex_payload = _parse_hex_line(line)
            if hex_payload is not None and current is not None:
                offset, payload = hex_payload
                end = offset + len(payload)
                if offset < 0 or end > current["size"]:
                    _fail("raw_hex_bounds", "raw hex payload exceeds declared record", line=line_number, offset=offset, payload_bytes=len(payload), record_bytes=current["size"])
                written = current["written"]
                duplicate = next((index for index in range(offset, end) if written[index]), None)
                if duplicate is not None:
                    _fail("raw_hex_overlap", "raw hex payload overlaps an earlier payload", line=line_number, offset=duplicate)
                current["data"][offset:end] = payload
                written[offset:end] = b"\1" * len(payload)
                current["written_count"] += len(payload)
                continue
            sample_match = SAMPLE_LINE_RE.match(line)
            if sample_match is None and "PERF_RECORD_SAMPLE" in line:
                _fail("sample_line_parse", "sample line does not match expected perf script format", line=line_number, text=line[:300])
            if sample_match is None:
                continue
            sample_line_count += 1
            if current is None:
                _fail("sample_without_raw", "decoded sample has no preceding raw event", line=line_number)
            if current["consumed"]:
                _fail("sample_double_pair", "two decoded samples consumed one raw event", line=line_number)
            finalize_data = current["data"]
            if current["written_count"] != current["size"]:
                _fail("raw_record_truncated", "sample raw event hex payload is incomplete", line=line_number, declared_bytes=current["size"], written_bytes=current["written_count"])
            bind_raw_header(line_number)
            record_type = int(current["header_type"])
            record_misc = int(current["header_misc"])
            record_size = int(current["header_size"])
            if record_type != 9:
                _fail("sample_record_type", "sample line paired with non-sample raw event", line=line_number, record_type=record_type)
            event = {
                "cpu_text": int(sample_match.group("cpu")), "time_text": int(sample_match.group("time")),
                "misc_text": int(sample_match.group("misc"), 16), "pid_text": int(sample_match.group("pid")),
                "tid_text": int(sample_match.group("tid")), "ip_text": int(sample_match.group("ip"), 16),
                "period_text": int(sample_match.group("period")),
            }
            parsed = _parse_sample_payload(bytes(finalize_data), contract=contract, event=event, stack_cap=stack_cap)
            current["consumed"] = True
            paired_sample_count += 1
            sample = {
                "index": paired_sample_count,
                "dump_line": line_number,
                "raw_event_line": current["line"],
                "summary_cpu": event["cpu_text"],
                "time_ns": parsed.get("time_ns", event["time_text"]),
                "ip": parsed.get("ip"),
                "period": parsed.get("period"),
                "pid": parsed.get("pid"),
                "tid": parsed.get("tid"),
                "summary_misc": event["misc_text"],
                "raw_bytes": current["size"],
                "physical_offset": current.get("location", {}).get("offset") if current.get("location") else None,
                **parsed,
            }
            if event["misc_text"] != parsed["record_misc"]:
                violations.append({"reason": "summary_misc_mismatch", "index": paired_sample_count, "summary": event["misc_text"], "raw": parsed["record_misc"]})
            if event["pid_text"] != parsed.get("pid") or event["tid_text"] != parsed.get("tid"):
                pid_tid_cpu_valid = False
                violations.append({"reason": "summary_pid_tid_mismatch", "index": paired_sample_count})
            if event["ip_text"] != parsed.get("ip"):
                violations.append({"reason": "summary_ip_mismatch", "index": paired_sample_count})
            if event["cpu_text"] != parsed.get("cpu"):
                pid_tid_cpu_valid = False
                violations.append({"reason": "summary_cpu_mismatch", "index": paired_sample_count, "summary": event["cpu_text"], "raw": parsed.get("cpu")})
            if event["time_text"] != parsed.get("time_ns"):
                time_valid = False
                violations.append({"reason": "summary_time_mismatch", "index": paired_sample_count, "summary": event["time_text"], "raw": parsed.get("time_ns")})
            if event["period_text"] != parsed.get("period"):
                period_valid = False
                violations.append({"reason": "summary_period_mismatch", "index": paired_sample_count})
            if parsed.get("period") != expected_period:
                period_valid = False
                violations.append({"reason": "sample_period", "index": paired_sample_count, "actual": parsed.get("period"), "expected": expected_period})
            if parsed.get("pid") != expected_pid or parsed.get("tid") != expected_pid:
                pid_tid_cpu_valid = False
                violations.append({"reason": "sample_pid_tid", "index": paired_sample_count, "pid": parsed.get("pid"), "tid": parsed.get("tid"), "expected": expected_pid})
            if expected_cpu is not None and "CPU" in contract["sample_type_names"] and parsed.get("cpu") != expected_cpu:
                pid_tid_cpu_valid = False
                violations.append({"reason": "sample_cpu", "index": paired_sample_count, "actual": parsed.get("cpu"), "expected": expected_cpu})
            if not parsed.get("stack_bounds_valid", True):
                bounds_valid = False
                violations.append({"reason": "stack_bound", "index": paired_sample_count, "stack_size": parsed.get("stack_size"), "stack_dyn_size": parsed.get("stack_dyn_size"), "stack_cap": stack_cap})
            mode_counts[parsed["mode"]] += 1
            if parsed["missing_unwindability"]:
                mode_missing_counts[parsed["mode"]] += 1
            if "regs_abi_name" in parsed:
                regs_abi_counts[parsed["regs_abi_name"]] += 1
            if "stack_size" in parsed:
                stack_shape_counts["zero" if parsed["stack_size"] == 0 else "nonzero"] += 1
            emit(sample)
    finalize_current(sample_line_count + raw_event_count)
    if sample_line_count != paired_sample_count:
        _fail("sample_conservation", "decoded sample lines and paired raw samples differ", sample_lines=sample_line_count, paired=paired_sample_count)
    if raw_sample_unpaired_count:
        _fail("sample_conservation", "raw sample records were not paired to decoded samples", unpaired=raw_sample_unpaired_count)

    raw_sample_record_count = sum(count for record_type, count in raw_type_counts.items() if record_type == 9)
    result = {
        "schema": SCHEMA,
        "status": "complete",
        "event_contract": contract,
        "perf_data": file_binding(perf_data) if perf_data is not None else None,
        "binary": file_binding(binary) if binary is not None else None,
        "dump": file_binding(path),
        "expected_pid": expected_pid,
        "expected_cpu": expected_cpu,
        "expected_period": expected_period,
        "stack_cap": stack_cap,
        "raw_event_count": raw_event_count,
        "physical_location_count": len(seen_locations),
        "raw_sample_record_count": sum(count for record_type, count in raw_type_counts.items() if record_type == 9),
        "raw_unconsumed_count": raw_unconsumed_count,
        "sample_line_count": sample_line_count,
        "sample_count": paired_sample_count,
        "sample_conservation": raw_sample_record_count == paired_sample_count,
        "period_valid": period_valid,
        "pid_tid_cpu_valid": pid_tid_cpu_valid,
        "time_valid": time_valid,
        "bounds_valid": bounds_valid,
        "physical_identity_valid": physical_identity_valid,
        "raw_record_type_counts": {str(key): value for key, value in sorted(raw_type_counts.items())},
        "raw_record_size_counts": {str(key): value for key, value in sorted(raw_size_counts.items())},
        "mode_counts": dict(sorted(mode_counts.items())),
        "mode_missing_unwindability_counts": dict(sorted(mode_missing_counts.items())),
        "regs_abi_counts": dict(sorted(regs_abi_counts.items())),
        "stack_shape_counts": dict(sorted(stack_shape_counts.items())),
        "violations": violations,
        "scope": {
            "all_perf_record_sample_lines": True,
            "all_raw_records_counted": True,
            "variable_layout": True,
            "no_unwind_or_owner_claim": True,
        },
    }
    if retain_samples:
        result["sample_ledger"] = samples
    else:
        result["sample_ledger_retained"] = False
    return result


__all__ = [
    "ScanFailure", "SCHEMA", "SUPPORTED_SAMPLE_MASK", "make_event_contract",
    "parse_evlist_contract", "scan_dump", "file_binding", "sha256",
]
