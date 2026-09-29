#!/usr/bin/env python3
"""Offline analysis of the 0836 perf and strace captures.

The capture is deliberately interpreted as a diagnostic.  A perf sample is
assigned to the owner population only when its PID belongs to the measured
children in the workload report, its DSO is the exact admitted executable,
and its instruction address falls inside an admitted owner symbol range.
Everything else is retained in a separate population.  The parser keeps
unknown, empty, malformed, and lost-event evidence visible and never turns a
partial stack into a CPU fraction.

This program never invokes Cargo, the workload, perf, strace, or Git.  It
consumes only the retained packet artifacts and writes one deterministic JSON
report.  ``--check`` replays that report byte-for-byte.
"""

from __future__ import annotations

import argparse
from collections import Counter
import gzip
import hashlib
import json
import math
import posixpath
import re
import sys
from pathlib import Path
from typing import Any, Iterable, NoReturn


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
PLAN_SCHEMA = "litchi.0836.opc-filesystem-profile-plan.v1"
ANALYSIS_SCHEMA = "litchi.0836.opc-filesystem-profile-analysis.v1"
STAGES = ("baseline", "fp")
STATES = ("warm", "cold-verified")
CASE = "opc_file_source_one_part_atomic_save"
PERF_EVENT = "cycles:u"
PERF_FREQUENCY = 997
MIN_OWNER_SAMPLES = 500
PAGE_SIZE = 4096
TRACE_SYSCALLS = ("fsync", "fdatasync", "rename", "renameat", "renameat2")
TRACE_RENAMES = ("rename", "renameat", "renameat2")

HEADER_RE = re.compile(
    r"^\s*(?P<command>.+?)\s+(?P<pid>[0-9]+)(?:/[0-9]+)?\s+"
    r"(?:\[[0-9]+\]\s+)?(?P<timestamp>[0-9]+(?:\.[0-9]+)?):\s+"
    # ``cycles:u`` contains an internal colon; the final colon is perf's
    # event-line terminator.  Capture the complete non-space event token and
    # keep the exact ``cycles:u`` check below as the admission gate.
    r"(?P<period>[0-9]+)\s+(?P<event>\S+):\s*$"
)
FRAME_RE = re.compile(
    r"^\s*(?:0x)?(?P<address>[0-9a-fA-F]+)\s+(?P<body>.+?)\s+"
    r"\((?P<dso>.*)\)\s*$"
)
TRACE_PREFIX_RE = re.compile(
    r"^\s*(?P<pid>[0-9]+)\s+(?P<timestamp>[0-9]+(?:\.[0-9]+)?)\s+"
    r"(?P<body>.*)$"
)
TRACE_UNFINISHED_RE = re.compile(
    r"^(?P<call>[A-Za-z_][A-Za-z0-9_]*)\((?P<args>.*)<unfinished \.\.\.>\s*$"
)
TRACE_RESUMED_RE = re.compile(
    r"^<\.\.\.\s+(?P<call>[A-Za-z_][A-Za-z0-9_]*) resumed>(?P<body>.*)$"
)
TRACE_NORMAL_RE = re.compile(
    r"^(?P<call>[A-Za-z_][A-Za-z0-9_]*)\((?P<args>.*)$"
)
TRACE_DURATION_RE = re.compile(r"<(?P<seconds>[0-9]+(?:\.[0-9]+)?)>\s*$")
TRACE_RETURN_RE = re.compile(r"=\s*(?P<return>-?[0-9]+)(?:\s+[A-Za-z_][A-Za-z0-9_]*(?:\s+\([^)]*\))?)?\s*(?:<|$)")
HEX_OFFSET_RE = re.compile(r"\+0x[0-9a-fA-F]+$")
SHA_RE = re.compile(r"[0-9a-f]{64}\Z")


class EvidenceError(RuntimeError):
    """A retained artifact is missing, malformed, or contradictory."""


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1 << 20), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")
    return value


def finite(value: Any, label: str = "json") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{label}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            require(isinstance(key, str), f"{label}: non-string key")
            finite(child, f"{label}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            finite(child, f"{label}[{index}]")


def valid_sha(value: Any) -> bool:
    return isinstance(value, str) and SHA_RE.fullmatch(value) is not None


def resolve_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}.path: missing")
    path = Path(raw)
    if not path.is_absolute():
        path = HERE / path
    return path.resolve(strict=False)


def _contains_descriptor(value: Any, expected: dict[str, Any]) -> bool:
    if isinstance(value, dict):
        if all(value.get(key) == expected.get(key)
               for key in ("path", "bytes", "sha256")):
            return True
        return any(_contains_descriptor(child, expected) for child in value.values())
    if isinstance(value, list):
        return any(_contains_descriptor(child, expected) for child in value)
    return False


def cleanup_contains(expected: dict[str, Any]) -> bool:
    path = HERE / "cleanup.json"
    if not path.is_file() or path.is_symlink():
        return False
    try:
        return _contains_descriptor(json.loads(path.read_text(encoding="utf-8")), expected)
    except (OSError, UnicodeError, json.JSONDecodeError):
        return False


def descriptor(value: Any, label: str, *, allow_missing: bool = False,
               nonempty: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: descriptor is malformed")
    raw_path = value.get("path")
    path = resolve_path(raw_path, label)
    size = value.get("bytes")
    require(type(size) is int and size >= (1 if nonempty else 0),
            f"{label}.bytes: invalid")
    digest = value.get("sha256")
    require(valid_sha(digest), f"{label}.sha256: invalid")
    require(not path.is_symlink(), f"{label}: descriptor is a symlink")
    if path.is_file():
        require(path.stat().st_size == size, f"{label}: byte count changed")
        require(sha256(path) == digest, f"{label}: hash changed")
    else:
        require(allow_missing, f"{label}: missing artifact {path}")
        require(cleanup_contains({"path": str(raw_path), "bytes": size,
                                  "sha256": digest}),
                f"{label}: missing artifact lacks cleanup witness")
    return {"path": str(raw_path), "bytes": size, "sha256": digest}


def descriptor_bytes(value: dict[str, Any], label: str, *, allow_missing: bool = False) -> bytes | None:
    checked = descriptor(value, label, allow_missing=allow_missing)
    path = resolve_path(checked["path"], label)
    if not path.is_file():
        return None
    try:
        return path.read_bytes()
    except OSError as error:
        fail(f"cannot read {path}: {error}")
    return None


def verify_bytes(data: bytes, value: dict[str, Any], label: str,
                 *, allow_missing: bool = False) -> None:
    require(len(data) == value.get("bytes"), f"{label}: decompressed byte count differs")
    require(hashlib.sha256(data).hexdigest() == value.get("sha256"),
            f"{label}: decompressed hash differs")
    descriptor(value, label, allow_missing=allow_missing)


def gzip_payload(value: dict[str, Any], label: str) -> bytes:
    stored = descriptor_bytes(value, label)
    require(stored is not None, f"{label}: gzip artifact is missing")
    try:
        data = gzip.decompress(stored)
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        fail(f"{label}: invalid gzip: {error}")
    return data


def write_json(path: Path, value: dict[str, Any], check: bool) -> None:
    encoded = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if check:
        require(path.is_file() and not path.is_symlink(), f"missing analysis output: {path}")
        require(path.read_bytes() == encoded, f"{path}: deterministic replay differs")
        return
    require(not path.exists(), f"refusing to overwrite {path}")
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
    except FileExistsError:
        fail(f"refusing to overwrite {path}")


def load_plan() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    finite(plan, "plan")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == PLAN_SCHEMA, "profile plan schema changed")
    require(plan.get("case") == CASE and plan.get("cpu") == 12,
            "profile case or CPU changed")
    contract = plan.get("perf_contract")
    require(isinstance(contract, dict), "perf contract is missing")
    require(contract.get("event") == PERF_EVENT
            and contract.get("frequency_hz") == PERF_FREQUENCY
            and contract.get("call_graph") == "fp"
            and contract.get("minimum_owner_samples_per_capture") == MIN_OWNER_SAMPLES,
            "perf contract changed")
    require(contract.get("decode") == ["perf", "script", "--no-inline", "--ns",
                                        "--show-lost-events"],
            "perf decode command changed")
    trace = plan.get("trace_contract")
    require(isinstance(trace, dict)
            and trace.get("options") == ["-f", "-qq", "-ttt", "-T", "-yy"]
            and trace.get("syscalls") == list(TRACE_SYSCALLS),
            "trace contract changed")
    return plan


def _range_row(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: range is malformed")
    address = value.get("address")
    size = value.get("size")
    end = value.get("end")
    require(type(address) is int and address >= 0, f"{label}.address: invalid")
    require(type(size) is int and size > 0, f"{label}.size: invalid")
    require(end == address + size, f"{label}.end: range equation differs")
    symbol = value.get("symbol")
    raw_symbol = value.get("raw_symbol", symbol)
    require(isinstance(symbol, str) and symbol, f"{label}.symbol: missing")
    require(isinstance(raw_symbol, str) and raw_symbol, f"{label}.raw_symbol: missing")
    return {"address": address, "size": size, "end": end,
            "symbol": symbol, "raw_symbol": raw_symbol}


def load_symbols() -> dict[str, Any]:
    value = read_json(HERE / "symbols.json")
    finite(value, "symbols")
    require(isinstance(value, dict) and value.get("status") == "pass",
            "symbols admission is not pass")
    owner = value.get("owner")
    require(isinstance(owner, str) and owner, "admitted owner is missing")
    stages = value.get("stages")
    require(isinstance(stages, dict), "admitted symbol stages are missing")
    checked: dict[str, Any] = {"status": "pass", "owner": owner,
                               "scope": value.get("scope"), "stages": {}}
    for stage in STAGES:
        item = stages.get(stage)
        require(isinstance(item, dict), f"symbols.{stage}: missing")
        binary = descriptor(item.get("binary"), f"symbols.{stage}.binary", allow_missing=True,
                            nonempty=True)
        ranges = item.get("ranges")
        require(isinstance(ranges, list) and ranges, f"symbols.{stage}.ranges: missing")
        checked_ranges = [_range_row(row, f"symbols.{stage}.ranges[{index}]")
                          for index, row in enumerate(ranges)]
        ordered = sorted(checked_ranges, key=lambda row: row["address"])
        for previous, current in zip(ordered, ordered[1:]):
            require(previous["end"] <= current["address"],
                    f"symbols.{stage}.ranges overlap")
        checked["stages"][stage] = {"binary": binary, "ranges": checked_ranges,
                                    "build_id": item.get("build_id")}
    require(checked["stages"]["baseline"]["ranges"]
            and checked["stages"]["fp"]["ranges"], "owner ranges are empty")
    return checked


def _report_pids(value: Any, found: set[int] | None = None) -> set[int]:
    if found is None:
        found = set()
    if isinstance(value, dict):
        for key, child in value.items():
            if key in {"child_process_id", "child_pid", "measured_pid"}:
                require(type(child) is int and child > 0,
                        f"report PID field {key} is invalid")
                found.add(child)
            elif key in {"pids", "child_pids", "measured_pids", "report_pids"}:
                if isinstance(child, list):
                    for pid in child:
                        require(type(pid) is int and pid > 0,
                                f"report PID list {key} is invalid")
                        found.add(pid)
            _report_pids(child, found)
    elif isinstance(value, list):
        for child in value:
            _report_pids(child, found)
    return found


def report_pids(row: dict[str, Any], label: str) -> tuple[dict[str, Any], list[int]]:
    report = descriptor(row.get("report"), f"{label}.report")
    path = resolve_path(report["path"], f"{label}.report")
    require(path.is_file(), f"{label}.report: missing")
    value = read_json(path)
    finite(value, str(path))
    pids = sorted(_report_pids(value))
    require(pids, f"{label}: report exposes no measured child PID")
    for key in ("pids", "child_pids", "measured_pids", "report_pids"):
        if key in row:
            raw = row[key]
            require(isinstance(raw, list) and sorted(raw) == pids,
                    f"{label}.{key}: PID binding differs from report")
    return report, pids


def receipt_from_row(row: dict[str, Any], label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    receipt = descriptor(row.get("receipt"), f"{label}.receipt")
    path = resolve_path(receipt["path"], f"{label}.receipt")
    value = read_json(path)
    require(value.get("exit_code") == 0 and value.get("error") is None,
            f"{label}.receipt: command failed")
    return receipt, value


def row_identity(row: dict[str, Any], expected: dict[str, Any], label: str) -> None:
    require(isinstance(row, dict), f"{label}: row is malformed")
    for key in ("stage", "state", "samples", "warmup"):
        require(row.get(key) == expected[key], f"{label}.{key}: plan differs")
    if "repeat" in expected:
        require(row.get("repeat") == expected["repeat"], f"{label}.repeat: plan differs")
    require(row.get("case") == CASE and row.get("exit_code") == 0,
            f"{label}: case or exit status differs")


def _frame_parts(body: str) -> tuple[str, int | None]:
    """Return the symbol and optional perf ``+0x`` instruction offset."""

    body = body.strip()
    match = HEX_OFFSET_RE.search(body)
    if match is None:
        return body, None
    return body[:match.start()].rstrip(), int(match.group()[3:], 16)


def _unknown_symbol(symbol: str, dso: str) -> bool:
    lowered = f"{symbol} {dso}".lower()
    return any(token in lowered for token in ("[unknown]", "<unknown>", "??"))


def parse_perf_text(data: bytes, measured_pids: Iterable[int], stage_info: dict[str, Any],
                    *, label: str = "perf") -> dict[str, Any]:
    """Parse one ``perf script`` stream and retain every sample population.

    ``stage_info`` is the checked entry from :func:`load_symbols`.  The
    Runtime addresses are accepted only when they directly fall inside the
    admitted static ELF ranges.  This reader never infers a PIE load bias from
    the samples themselves: an independently decoded mmap timeline must be
    supplied before relocated runtime addresses can be admitted.
    """

    require(isinstance(data, bytes) and data, f"{label}: empty perf stream")
    measured = set(measured_pids)
    require(measured and all(type(pid) is int and pid > 0 for pid in measured),
            f"{label}: measured PID set is empty or malformed")
    text = data.decode("utf-8", errors="replace")
    lines = text.splitlines()
    samples: list[dict[str, Any]] = []
    statuses: list[str] = []
    malformed_outside: list[str] = []
    current: dict[str, Any] | None = None

    def finish() -> None:
        nonlocal current
        if current is not None:
            samples.append(current)
            current = None

    for line_number, line in enumerate(lines, 1):
        match = HEADER_RE.fullmatch(line)
        if match:
            finish()
            event = match.group("event").strip()
            period = int(match.group("period"))
            require(period > 0, f"{label}:{line_number}: non-positive sample period")
            current = {
                "index": len(samples), "line": line_number,
                "command": match.group("command").strip(),
                "pid": int(match.group("pid")),
                "timestamp": match.group("timestamp"),
                "period": period, "event": event, "frames": [],
                "malformed_lines": [], "raw_lines": [line],
            }
            continue
        if current is None:
            if line.strip():
                statuses.append(line)
            continue
        if not line.strip():
            finish()
            continue
        current["raw_lines"].append(line)
        frame = FRAME_RE.fullmatch(line)
        if frame is None:
            current["malformed_lines"].append({"line": line_number, "text": line})
            continue
        symbol, offset = _frame_parts(frame.group("body"))
        current["frames"].append({
            "address": int(frame.group("address"), 16),
            "symbol": symbol,
            "offset": offset,
            "dso": frame.group("dso").strip(),
            "text": line,
        })
    finish()
    require(samples, f"{label}: perf stream has no sample headers")
    timestamps = [sample["timestamp"] for sample in samples]
    duplicate_timestamps = len(timestamps) - len(set(timestamps))
    lost_lines = [line for line in lines if "PERF_RECORD_LOST" in line
                  or ("lost" in line.lower() and
                      ("sample" in line.lower() or "event" in line.lower()))]
    require(all(sample["event"] == PERF_EVENT for sample in samples),
            f"{label}: sample event is not exactly {PERF_EVENT}")

    ranges = stage_info["ranges"]
    binary_path = stage_info["binary"]["path"]
    owner = stage_info.get("owner")
    if not owner:
        owner = ranges[0]["symbol"]
    exact_names = {row["symbol"] for row in ranges} | {row["raw_symbol"] for row in ranges}

    def candidates(frame: dict[str, Any]) -> list[dict[str, Any]]:
        if frame["symbol"] not in exact_names and frame["symbol"] != owner:
            return []
        return [row for row in ranges
                if frame["symbol"] in {row["symbol"], row["raw_symbol"]}
                or frame["symbol"] == owner]

    def in_range(frame: dict[str, Any], rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        matches: list[dict[str, Any]] = []
        for row in rows:
            if row["address"] <= frame["address"] < row["end"]:
                matches.append({"range": row, "load_bias": 0,
                                "adjusted_address": frame["address"],
                                "range_basis": "static_address"})
        return matches

    counters: Counter[str] = Counter()
    periods: Counter[str] = Counter()
    leaf_counts: Counter[tuple[str, str]] = Counter()
    leaf_periods: Counter[tuple[str, str]] = Counter()
    groups: Counter[str] = Counter()
    group_periods: Counter[str] = Counter()
    owner_rows: list[dict[str, Any]] = []
    sample_rows: list[dict[str, Any]] = []
    owner_period = 0
    measured_period = 0
    measured_samples = 0
    owner_other_dso = 0
    owner_range_mismatch = 0
    repeated_owner = 0

    for sample in samples:
        frames = sample["frames"]
        pid_measured = sample["pid"] in measured
        malformed = bool(sample["malformed_lines"])
        empty = not frames
        unknown = any(_unknown_symbol(frame["symbol"], frame["dso"]) for frame in frames)
        owner_named = [frame for frame in frames if frame["symbol"] in exact_names
                       or frame["symbol"] == owner]
        exact_owner: list[tuple[dict[str, Any], dict[str, Any]]] = []
        wrong_dso_owner = 0
        mismatch_owner = 0
        for frame in owner_named:
            rows = candidates(frame)
            if frame["dso"] != binary_path:
                wrong_dso_owner += 1
                continue
            matches = in_range(frame, rows)
            if matches:
                exact_owner.extend((frame, match) for match in matches)
            else:
                mismatch_owner += 1
        if wrong_dso_owner:
            owner_other_dso += wrong_dso_owner
        if mismatch_owner:
            owner_range_mismatch += mismatch_owner
        if len(exact_owner) > 1:
            repeated_owner += 1
        qualified = (pid_measured and not malformed and len(exact_owner) == 1)
        if pid_measured:
            measured_samples += 1
            measured_period += sample["period"]
        if sample["pid"] not in measured:
            partition = "outside_pid"
        elif malformed:
            partition = "malformed"
        elif empty:
            partition = "empty_stack"
        elif qualified:
            partition = "owner_qualified"
        else:
            partition = "outside_owner"
        counters[partition] += 1
        periods[partition] += sample["period"]
        if unknown:
            counters["unknown_stack"] += 1
            periods["unknown_stack"] += sample["period"]
        if qualified:
            frame, match = exact_owner[0]
            owner_period += sample["period"]
            interior = frames[:frames.index(frame)]
            if interior:
                leaf = (interior[0]["symbol"], interior[0]["dso"])
            else:
                leaf = (frame["symbol"], frame["dso"])
            leaf_counts[leaf] += 1
            leaf_periods[leaf] += sample["period"]
            group, reason = classify_group(interior)
            groups[group] += 1
            group_periods[group] += sample["period"]
            owner_row = {
                "sample_index": sample["index"], "pid": sample["pid"],
                "period": sample["period"], "timestamp": sample["timestamp"],
                "owner_frame_index": frames.index(frame),
                "owner_symbol": frame["symbol"], "owner_dso": frame["dso"],
                "owner_address": frame["address"],
                "owner_offset": frame["offset"],
                "owner_range": {key: match["range"][key]
                                for key in ("address", "size", "end", "symbol", "raw_symbol")},
                "load_bias": match["load_bias"],
                "adjusted_address": match["adjusted_address"],
                "range_basis": match["range_basis"],
                "self_leaf": {"symbol": leaf[0], "dso": leaf[1]},
                "group": group, "group_reason": reason,
                "unknown_interior": any(_unknown_symbol(item["symbol"], item["dso"])
                                         for item in interior),
            }
            owner_rows.append(owner_row)
        sample_rows.append({
            "sample_index": sample["index"], "pid": sample["pid"],
            "period": sample["period"], "timestamp": sample["timestamp"],
            "partition": partition, "frame_count": len(frames),
            "unknown": unknown, "malformed": malformed,
            "owner_named_frames": len(owner_named),
            "exact_owner_ranges": len(exact_owner),
        })

    diagnostics = {
        "raw_text_bytes": len(data), "raw_text_lines": len(lines),
        "parsed_samples": len(samples), "measured_pid_count": len(measured),
        "measured_sample_count": measured_samples,
        "outside_pid_samples": counters["outside_pid"],
        "outside_pid_period": periods["outside_pid"],
        "outside_owner_samples": sum(1 for sample in samples
                                      if sample["pid"] in measured
                                      and sample["index"] not in {row["sample_index"] for row in owner_rows}),
        "outside_owner_period": sum(sample["period"] for sample in samples
                                     if sample["pid"] in measured
                                     and sample["index"] not in {row["sample_index"] for row in owner_rows}),
        "empty_stack_samples": sum(not sample["frames"] for sample in samples),
        "unknown_stack_samples": sum(any(_unknown_symbol(frame["symbol"], frame["dso"])
                                          for frame in sample["frames"])
                                     for sample in samples),
        "malformed_samples": sum(bool(sample["malformed_lines"]) for sample in samples),
        "malformed_frame_lines": sum(len(sample["malformed_lines"]) for sample in samples),
        "malformed_outside_sample_lines": malformed_outside,
        "status_lines": statuses,
        "status_line_count": len(statuses),
        "lost_event_lines": lost_lines,
        "lost_event_count": len(lost_lines),
        "duplicate_timestamp_count": duplicate_timestamps,
    }
    partition = {key: counters[key] for key in
                 ("owner_qualified", "outside_owner", "outside_pid", "empty_stack", "malformed")}
    # Empty and malformed are diagnostics as well as exclusive labels.  The
    # equation below is intentionally over the exclusive labels, while the
    # independent unknown counter may overlap any of them.
    require(sum(partition.values()) == len(samples),
            f"{label}: exclusive sample partition does not reconstruct stream")
    require(sum(groups.values()) == len(owner_rows)
            and sum(group_periods.values()) == owner_period,
            f"{label}: owner group partition does not reconstruct owner population")
    require(sum(leaf_counts.values()) == len(owner_rows)
            and sum(leaf_periods.values()) == owner_period,
            f"{label}: self-leaf partition does not reconstruct owner population")

    cpu_reasons: list[str] = []
    if len(owner_rows) <= MIN_OWNER_SAMPLES:
        cpu_reasons.append(f"owner sample count {len(owner_rows)} is not greater than {MIN_OWNER_SAMPLES}")
    if diagnostics["lost_event_count"]:
        cpu_reasons.append("explicit perf lost-event evidence is present")
    if diagnostics["malformed_samples"] or diagnostics["malformed_frame_lines"]:
        cpu_reasons.append("malformed perf frame evidence is present")
    if diagnostics["unknown_stack_samples"] or diagnostics["empty_stack_samples"]:
        cpu_reasons.append("unknown or empty stack evidence is present")
    if diagnostics["duplicate_timestamp_count"]:
        cpu_reasons.append("duplicate perf timestamps are present")
    cpu_authorized = not cpu_reasons
    cpu_fraction = (owner_period / measured_period if cpu_authorized and measured_period else None)

    def census(counter: Counter[tuple[str, str]], periods_: Counter[tuple[str, str]]) -> list[dict[str, Any]]:
        return [{"symbol": symbol, "dso": dso, "samples": count,
                 "period": periods_[(symbol, dso)]}
                for (symbol, dso), count in sorted(counter.items(),
                                                   key=lambda item: (-item[1], item[0]))]

    group_rows = [{"group": group, "samples": groups[group], "period": group_periods[group]}
                  for group in ("deflate", "inflate", "copy-crc", "other")]
    return {
        "label": label, "owner": owner, "binary": stage_info["binary"],
        "measured_pids": sorted(measured), "binary_dso_exact": binary_path,
        "whole_process_samples": len(samples),
        "whole_process_period": sum(sample["period"] for sample in samples),
        "measured_pid_samples": measured_samples,
        "measured_pid_period": measured_period,
        "owner_qualified_samples": len(owner_rows),
        "owner_qualified_period": owner_period,
        "partition": partition,
        "partition_period": {key: periods[key] for key in partition},
        "self_leaf_census": census(leaf_counts, leaf_periods),
        "disjoint_groups": group_rows,
        "disjoint_group_equations": {
            "samples": f"sum(groups.samples) == {len(owner_rows)}",
            "period": f"sum(groups.period) == {owner_period}",
            "verified": True,
        },
        "owner_diagnostics": {
            "owner_symbol_other_dso_frames": owner_other_dso,
            "owner_address_or_symbol_range_mismatch_frames": owner_range_mismatch,
            "repeated_owner_samples": repeated_owner,
            "exact_owner_range_intersection": True,
            "runtime_address_policy": "direct static range only; no sample-derived PIE bias",
            "runtime_mapping_required_for_relocated_pie": True,
        },
        "decode_diagnostics": diagnostics,
        "sample_rows": sample_rows,
        "owner_sample_rows": owner_rows,
        "cpu_fraction_authorized": cpu_authorized,
        "cpu_fraction": cpu_fraction,
        "cpu_fraction_denominator": "measured PID sample period; diagnostic only" if cpu_authorized else None,
        "cpu_fraction_refusal_reasons": cpu_reasons,
        "claim_scope": "descriptive sampled cycles; static-range owner binding only; no relocated runtime-range claim",
    }


def classify_group(interior: list[dict[str, Any]]) -> tuple[str, str]:
    """Choose one disjoint codec/copy group from a leaf-to-owner stack."""

    for frame in interior:
        lowered = f"{frame['symbol']} {frame['dso']}".lower()
        if "copy" in lowered and "crc" in lowered:
            return "copy-crc", "nearest stack frame containing copy and crc"
        if "deflate" in lowered or "deflater" in lowered:
            return "deflate", "nearest stack frame containing deflate"
        if "inflate" in lowered or "inflater" in lowered:
            return "inflate", "nearest stack frame containing inflate"
    return "other", "no deflate, inflate, or copy-crc marker in interior"


def check_perf_receipt(row: dict[str, Any], raw: dict[str, Any], label: str) -> None:
    _, receipt = receipt_from_row(row, label)
    argv = receipt.get("argv")
    require(isinstance(argv, list), f"{label}.receipt.argv: missing")
    require(argv[:5] == ["perf", "record", "--no-buildid-cache", "-e", PERF_EVENT],
            f"{label}: perf record event/options changed")
    require("-F" in argv and argv[argv.index("-F") + 1] == str(PERF_FREQUENCY),
            f"{label}: perf frequency changed")
    require("--call-graph" in argv and argv[argv.index("--call-graph") + 1] == "fp",
            f"{label}: perf call graph changed")
    require("-o" in argv and argv[argv.index("-o") + 1] == raw["path"],
            f"{label}: perf raw output binding changed")


def check_decode_receipt(row: dict[str, Any], raw: dict[str, Any], label: str) -> None:
    decode = descriptor(row.get("decode_receipt"), f"{label}.decode_receipt")
    value = read_json(resolve_path(decode["path"], f"{label}.decode_receipt"))
    require(value.get("exit_code") == 0 and value.get("error") is None,
            f"{label}: perf decoder failed")
    argv = value.get("argv")
    require(argv == ["perf", "script", "--no-inline", "--ns",
                     "--show-lost-events", "-i", raw["path"]],
            f"{label}: strict perf script argv changed")


def profile_row(row: dict[str, Any], expected: dict[str, Any], symbols: dict[str, Any],
                index: int) -> dict[str, Any]:
    label = row.get("label", f"profile-{index:02}")
    row_identity(row, expected, label)
    report, pids = report_pids(row, label)
    receipt_from_row(row, label)
    raw = descriptor(row.get("raw"), f"{label}.raw", allow_missing=True, nonempty=True)
    raw_gzip = descriptor(row.get("raw_gzip"), f"{label}.raw_gzip", nonempty=True)
    raw_data = gzip_payload(raw_gzip, f"{label}.raw_gzip")
    verify_bytes(raw_data, raw, f"{label}.raw", allow_missing=True)
    decoded = descriptor(row.get("decoded"), f"{label}.decoded", nonempty=True)
    decoded_plain = descriptor(row.get("decoded_plain"), f"{label}.decoded_plain",
                               allow_missing=True, nonempty=True)
    decoded_data = gzip_payload(decoded, f"{label}.decoded")
    verify_bytes(decoded_data, decoded_plain, f"{label}.decoded_plain", allow_missing=True)
    check_perf_receipt(row, raw, label)
    check_decode_receipt(row, raw, label)
    stage_info = dict(symbols["stages"][row["stage"]])
    stage_info["owner"] = symbols["owner"]
    parsed = parse_perf_text(decoded_data, pids, stage_info, label=label)
    parsed["repeat"] = row["repeat"]
    parsed["stage"] = row["stage"]
    parsed["state"] = row["state"]
    parsed["samples_requested"] = row["samples"]
    parsed["warmup"] = row["warmup"]
    parsed["report"] = report
    parsed["raw"] = raw
    parsed["raw_gzip"] = raw_gzip
    parsed["decoded"] = decoded
    parsed["decoded_plain"] = decoded_plain
    parsed["decode_receipt"] = descriptor(row["decode_receipt"], f"{label}.decode_receipt")
    return parsed


def _trace_event(pid: int, timestamp: str, call: str, body: str,
                 raw: str, sequence: int, resumed: bool = False) -> dict[str, Any]:
    duration = TRACE_DURATION_RE.search(body)
    require(duration is not None, f"trace line has no -T duration: {raw}")
    seconds = float(duration.group("seconds"))
    require(math.isfinite(seconds) and seconds >= 0, f"trace duration is invalid: {raw}")
    return_value = None
    returned = TRACE_RETURN_RE.search(body)
    if returned:
        return_value = int(returned.group("return"))
    # ``-T`` also appends ``<seconds>``.  It is not -yy provenance; only
    # angle-bracket fields carrying a pathname/descriptor annotation count.
    annotations = [item for item in re.findall(r"<([^>]+)>", body)
                   if "/" in item or item.startswith(("AT_", "anon_inode:"))]
    # The angle brackets are removed by ``findall``; an item such as
    # ``4</tmp/file>`` is returned as ``/tmp/file``.  Keep only absolute
    # pathname annotations here; ``-T`` duration and anon-inode annotations
    # are provenance/status fields, not filesystem paths.
    fd_paths = [item for item in annotations if item.startswith("/")]
    quoted_paths = re.findall(r'"((?:\\.|[^"\\])*)"', body)
    return {"pid": pid, "timestamp": timestamp, "syscall": call,
            "duration_ns": int(round(seconds * 1_000_000_000)),
            "return": return_value, "resumed": resumed,
            "provenance": annotations, "provenance_available": bool(annotations),
            "fd_paths": fd_paths, "quoted_paths": quoted_paths,
            "pathnames": fd_paths + quoted_paths,
            "raw": raw, "sequence": sequence}


def _join_trace_path(path: str, base: str | None = None) -> str:
    """Normalize a strace pathname, resolving relative renameat arguments."""

    require(isinstance(path, str) and path, "trace pathname is missing")
    if path.startswith("/") or base is None:
        return posixpath.normpath(path)
    return posixpath.normpath(posixpath.join(base, path))


def _rename_paths(event: dict[str, Any], label: str) -> tuple[str, str]:
    """Recover old/new paths from rename-family string and dirfd evidence."""

    quoted = event.get("quoted_paths")
    fd_paths = event.get("fd_paths")
    require(isinstance(quoted, list) and len(quoted) >= 2,
            f"{label}: rename lacks two quoted path arguments")
    require(isinstance(fd_paths, list), f"{label}: rename fd provenance malformed")
    source_base = fd_paths[0] if fd_paths else None
    destination_base = fd_paths[1] if len(fd_paths) > 1 else source_base
    source = _join_trace_path(quoted[0], source_base)
    destination = _join_trace_path(quoted[1], destination_base)
    require(source != destination, f"{label}: rename source and destination coincide")
    return source, destination


def parse_trace_text(data: bytes, measured_pids: Iterable[int], *, label: str = "trace") -> dict[str, Any]:
    """Parse strace with explicit unfinished/resumed pairing.

    A missing pair is an error instead of a silently shortened timing row.
    Syscalls from wrapper/primer PIDs are retained as outside-PID evidence.
    """

    require(isinstance(data, bytes) and data, f"{label}: empty trace stream")
    measured = set(measured_pids)
    require(measured and all(type(pid) is int and pid > 0 for pid in measured),
            f"{label}: measured PID set is empty or malformed")
    text = data.decode("utf-8", errors="replace")
    events: list[dict[str, Any]] = []
    statuses: list[str] = []
    malformed: list[str] = []
    pending: dict[tuple[int, str], dict[str, Any]] = {}
    sequence = 0
    for line_number, line in enumerate(text.splitlines(), 1):
        if not line.strip():
            continue
        prefix = TRACE_PREFIX_RE.fullmatch(line)
        if prefix is None:
            # Signals and process termination are normal strace status lines.
            if line.lstrip().startswith(("+++", "---")):
                statuses.append(line)
                continue
            malformed.append(f"{line_number}: {line}")
            continue
        pid = int(prefix.group("pid"))
        timestamp = prefix.group("timestamp")
        body = prefix.group("body")
        unfinished = TRACE_UNFINISHED_RE.fullmatch(body)
        if unfinished:
            call = unfinished.group("call")
            require(call in TRACE_SYSCALLS,
                    f"{label}:{line_number}: unfinished non-contract syscall {call}")
            key = (pid, call)
            require(key not in pending, f"{label}:{line_number}: duplicate unfinished {key}")
            pending[key] = {"pid": pid, "call": call, "timestamp": timestamp,
                            "line": line, "line_number": line_number,
                            "args": unfinished.group("args")}
            continue
        resumed = TRACE_RESUMED_RE.fullmatch(body)
        if resumed:
            call = resumed.group("call")
            key = (pid, call)
            require(key in pending, f"{label}:{line_number}: unmatched resumed {key}")
            prior = pending.pop(key)
            combined = prior["args"] + resumed.group("body")
            events.append(_trace_event(pid, timestamp, call, combined,
                                       prior["line"] + "\n" + line, sequence, True))
            sequence += 1
            continue
        normal = TRACE_NORMAL_RE.fullmatch(body)
        if normal is None:
            # A non-contract syscall/status is outside this packet's syscall
            # filter.  A line beginning with a contract name is malformed.
            if any(body.startswith(call + "(") for call in TRACE_SYSCALLS):
                malformed.append(f"{line_number}: {line}")
            else:
                statuses.append(line)
            continue
        call = normal.group("call")
        if call not in TRACE_SYSCALLS:
            statuses.append(line)
            continue
        events.append(_trace_event(pid, timestamp, call, normal.group("args"),
                                   line, sequence, False))
        sequence += 1
    require(not pending, f"{label}: unfinished syscall has no resumed line: {sorted(pending)}")
    require(not malformed, f"{label}: malformed trace lines: {malformed[:3]}")
    require(events, f"{label}: no contract syscall events")
    measured_events = [event for event in events if event["pid"] in measured]
    outside_events = [event for event in events if event["pid"] not in measured]
    by_pid: dict[int, list[dict[str, Any]]] = {pid: [] for pid in sorted(measured)}
    for event in measured_events:
        by_pid[event["pid"]].append(event)
    pid_rows: list[dict[str, Any]] = []
    scoped_events: list[dict[str, Any]] = []
    excluded_measured_events: list[dict[str, Any]] = []
    for pid in sorted(measured):
        rows = by_pid[pid]
        counts = Counter(row["syscall"] for row in rows)
        rename_rows = [row for row in rows if row["syscall"] in TRACE_RENAMES]
        require(len(rename_rows) == 1,
                f"{label}: measured PID {pid} has {len(rename_rows)} rename-family calls, expected 1")
        rename_row = rename_rows[0]
        rename_position = rows.index(rename_row)
        source_path, destination_path = _rename_paths(rename_row, f"{label}: PID {pid}")
        parent_path = posixpath.dirname(destination_path)
        require(parent_path and parent_path != destination_path,
                f"{label}: PID {pid} rename destination has no parent directory")

        # Cold preparation may fsync the source before the operation.  Scope
        # exactly the temp-file fsync immediately before rename and the
        # destination-parent fsync immediately after it; retain every other
        # durability event as excluded evidence.
        temp_candidates = [
            index for index, row in enumerate(rows[:rename_position])
            if row["syscall"] == "fsync" and source_path in row["pathnames"]
        ]
        parent_candidates = [
            index for index, row in enumerate(rows[rename_position + 1:], rename_position + 1)
            if row["syscall"] == "fsync" and parent_path in row["pathnames"]
        ]
        require(temp_candidates,
                f"{label}: PID {pid} lacks temp-file fsync before rename")
        require(parent_candidates,
                f"{label}: PID {pid} lacks destination-parent fsync after rename")
        temp_position = temp_candidates[-1]
        parent_position = parent_candidates[0]
        require(temp_position < rename_position < parent_position,
                f"{label}: PID {pid} scoped fsync/rename order is invalid")
        scoped_positions = (temp_position, rename_position, parent_position)
        scoped = [rows[index] for index in scoped_positions]
        scoped_counts = Counter(row["syscall"] for row in scoped)
        require(scoped_counts["fsync"] == 2
                and sum(scoped_counts[name] for name in TRACE_RENAMES) == 1,
                f"{label}: PID {pid} scoped durability counts are not 2 fsync plus 1 rename")
        provenance = all(row["provenance_available"] for row in scoped)
        require(provenance, f"{label}: PID {pid} lacks -yy provenance on scoped durability calls")
        extra = [row for index, row in enumerate(rows) if index not in scoped_positions]
        scoped_events.extend(scoped)
        excluded_measured_events.extend(extra)
        relevant = [row["syscall"] for row in scoped]
        pid_rows.append({
            "pid": pid, "events": rows,
            "scope_events": scoped, "excluded_events": extra,
            "counts": {name: scoped_counts[name] for name in TRACE_SYSCALLS},
            "all_counts": {name: counts[name] for name in TRACE_SYSCALLS},
            "relevant_order": relevant, "provenance_complete": provenance,
            "fsync_count": scoped_counts["fsync"],
            "rename_count": sum(scoped_counts[name] for name in TRACE_RENAMES),
            "excluded_event_count": len(extra),
            "scope_paths": {"temp": source_path, "destination": destination_path,
                            "parent_directory": parent_path},
            "order_contract": "fsync,rename-family,fsync",
        })
    return {
        "label": label, "measured_pids": sorted(measured),
        "whole_contract_events": len(events),
        "measured_contract_events": len(measured_events),
        "scoped_contract_events": len(scoped_events),
        "excluded_measured_events": len(excluded_measured_events),
        "scope_contract": "temp-file fsync, atomic rename-family, destination-parent fsync",
        "outside_pid_events": len(outside_events),
        "outside_pid_rows": outside_events,
        "status_lines": statuses, "status_line_count": len(statuses),
        "malformed_lines": malformed, "malformed_line_count": len(malformed),
        "unfinished_resumed_pairs": sum(row["resumed"] for row in events),
        "pid_rows": pid_rows,
        "order_provenance_verified": True,
        "claim_scope": "ptrace-observed syscall order and durations; no causal wall-time fraction",
    }


def check_trace_receipt(row: dict[str, Any], raw: dict[str, Any], label: str) -> None:
    _, receipt = receipt_from_row(row, label)
    argv = receipt.get("argv")
    require(isinstance(argv, list), f"{label}.receipt.argv: missing")
    prefix = ["strace", "-f", "-qq", "-ttt", "-T", "-yy", "-e",
              "trace=fsync,fdatasync,rename,renameat,renameat2"]
    require(argv[:len(prefix)] == prefix, f"{label}: strace options changed")
    require("-o" in argv and argv[argv.index("-o") + 1] == raw["path"],
            f"{label}: strace raw output binding changed")


def trace_row(row: dict[str, Any], expected: dict[str, Any], index: int) -> dict[str, Any]:
    label = row.get("label", f"trace-{index:02}")
    row_identity(row, expected, label)
    report, pids = report_pids(row, label)
    receipt_from_row(row, label)
    raw = descriptor(row.get("raw"), f"{label}.raw", allow_missing=True, nonempty=True)
    raw_gzip = descriptor(row.get("raw_gzip"), f"{label}.raw_gzip", nonempty=True)
    data = gzip_payload(raw_gzip, f"{label}.raw_gzip")
    verify_bytes(data, raw, f"{label}.raw", allow_missing=True)
    check_trace_receipt(row, raw, label)
    parsed = parse_trace_text(data, pids, label=label)
    parsed.update({"repeat": row["repeat"], "stage": row["stage"],
                   "state": row["state"], "samples_requested": row["samples"],
                   "warmup": row["warmup"], "report": report,
                   "raw": raw, "raw_gzip": raw_gzip})
    return parsed


def load_rows(name: str, plan_rows: list[dict[str, Any]], symbols: dict[str, Any],
              *, profile: bool) -> list[dict[str, Any]]:
    path = HERE / name
    value = read_json(path)
    finite(value, str(path))
    require(isinstance(value, dict) and value.get("status") == "commands_pass",
            f"{name}: manifest is not commands_pass")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == len(plan_rows),
            f"{name}: row count differs from plan")
    expected_samples = sum(row["samples"] for row in plan_rows)
    require(value.get("report_count") == len(rows)
            and value.get("sample_count") == expected_samples,
            f"{name}: manifest cardinality differs")
    result: list[dict[str, Any]] = []
    for index, (row, expected) in enumerate(zip(rows, plan_rows)):
        if profile:
            result.append(profile_row(row, expected, symbols, index))
        else:
            result.append(trace_row(row, expected, index))
    return result


def analyze() -> dict[str, Any]:
    plan = load_plan()
    symbols = load_symbols()
    profiles = load_rows("profiles.json", plan["perf"], symbols, profile=True)
    traces = load_rows("traces.json", plan["trace"], symbols, profile=False)
    require(len(profiles) == 4 and sum(row["samples_requested"] for row in profiles) == 120,
            "profile cardinality changed")
    require(len(traces) == 4 and sum(row["samples_requested"] for row in traces) == 32,
            "trace cardinality changed")
    plan_desc = {"path": str(HERE / "plan.json"), "bytes": (HERE / "plan.json").stat().st_size,
                 "sha256": sha256(HERE / "plan.json"), "schema": plan["schema"]}
    symbol_path = HERE / "symbols.json"
    symbol_desc = {"path": str(symbol_path), "bytes": symbol_path.stat().st_size,
                   "sha256": sha256(symbol_path), "owner": symbols["owner"]}
    profile_owner_samples = sum(row["owner_qualified_samples"] for row in profiles)
    profile_lost = sum(row["decode_diagnostics"]["lost_event_count"] for row in profiles)
    # The packet treats CPU and syscall captures as one shared analysis: four
    # CPU reports/120 samples plus four trace reports/32 samples.  Keep the
    # component aliases explicit so independent custody can check either
    # vocabulary without mistaking a component count for the shared total.
    return {
        "schema": ANALYSIS_SCHEMA, "status": "pass", "base": plan["base"],
        "case": CASE, "reports": 8, "samples": 152,
        "report_count": 8, "sample_count": 152,
        "cpu_reports": len(profiles), "cpu_samples": 120,
        "profile_reports": len(profiles), "profile_samples": 120,
        "trace_reports": len(traces), "trace_samples": 32,
        "plan": plan_desc, "symbols": symbol_desc,
        "owner": symbols["owner"],
        "profiles": {"reports": len(profiles), "samples": 120,
                     "captures": profiles,
                     "total_owner_qualified_samples": profile_owner_samples,
                     "total_lost_event_lines": profile_lost,
                     "cpu_fraction_policy": {
                         "minimum_owner_samples": MIN_OWNER_SAMPLES,
                         "comparison": "strictly greater than threshold per capture",
                         "lost_event_refusal": True,
                         "malformed_unknown_empty_refusal": True,
                         "claim_authorized": all(row["cpu_fraction_authorized"]
                                                  for row in profiles),
                     }},
        "traces": {"reports": len(traces), "samples": 32, "captures": traces,
                   "order_provenance_verified": all(
                       row["order_provenance_verified"] for row in traces)},
        "claims": {
            "performance_claim": "none",
            "cpu": "descriptive sampled cycles only; exact owner/DSO/PID intersection",
            "trace": "descriptive ptrace syscall order and durations; tracing perturbs timing",
            "iwork": "excluded",
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args(argv)
    try:
        value = analyze()
        write_json(HERE / "profile-analysis.json", value, args.check)
        # The custody audit uses the unambiguous shared-population filename;
        # retain the user-facing profile-analysis spelling as well.  Both are
        # generated from the same in-memory value and therefore replay
        # byte-for-byte.
        write_json(HERE / "sharedprofile-analysis.json", value, args.check)
    except (EvidenceError, AssertionError, OSError, UnicodeError,
            TypeError, ValueError, KeyError) as error:
        print(f"0836 profile analysis failed: {error}", file=sys.stderr)
        return 1
    print("0836 profile/trace analysis PASS: exact PID/DSO/range census and fail-closed syscall order")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
