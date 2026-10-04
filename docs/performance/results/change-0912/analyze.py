#!/usr/bin/env python3
"""Offline, fail-closed reduction of the 0912 native perf pilot.

This reader never invokes perf, the workload, a compiler, or Git.  It consumes
the retained first-child packet and writes a structured failure receipt plus a
short Markdown record.  ``--check`` replays both outputs byte-for-byte.
"""

from __future__ import annotations

from collections import Counter
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any, NoReturn


PACKET = Path(__file__).resolve().parent
LABEL = "materialized-a"
ROUTE = "materialized"
OWNER = "litchi_pptx::package::model::Package::opened_presentation_transaction"
PARENT = "litchi_perf_baseline::pptx_owned_benchmark::timed_iteration"
SETUP_PARENT = "litchi_perf_baseline::build_pptx_source_edit_corpus"
BACKENDS = (
    "sha2::sha256::x86_sha::compress",
    "sha2::sha256::soft::unroll::compress",
)
ANALYSIS_SCHEMA = "litchi.native-snapshot-profile-0912-first-child-failure.v1"
HEADER_RE = re.compile(
    r"^\s*(?P<comm>.+?)\s+(?P<pid>[0-9]+)(?:/(?P<tid>[0-9]+))?\s+"
    r"\[(?P<cpu>[0-9]+)\]\s+(?P<time>[0-9]+(?:\.[0-9]+)?):\s+"
    r"(?P<period>[0-9]+)\s+(?P<event>\S+):\s*$"
)
LOST_RE = re.compile(
    r"^\s*(?P<comm>.+?)\s+(?P<pid>[0-9]+)(?:/(?P<tid>[0-9]+))?\s+"
    r"\[(?P<cpu>[0-9]+)\]\s+(?P<time>[0-9]+(?:\.[0-9]+)?):\s+"
    r"PERF_RECORD_LOST\s+lost\s+(?P<count>[0-9]+)\s*$"
)
FRAME_RE = re.compile(
    r"^\s*(?P<address>(?:0x)?[0-9a-fA-F]+)\s+(?P<body>.*?)\s+"
    r"\((?P<dso>.*)\)\s*$"
)
MAPPING_RE = re.compile(
    r"^\s*(?P<comm>.+?)\s+(?P<pid>[0-9]+)(?:/(?P<tid>[0-9]+))?\s+"
    r"\[(?P<cpu>[0-9]+)\]\s+(?P<time>[0-9]+(?:\.[0-9]+)?):\s+"
    r"PERF_RECORD_MMAP2\s+(?P<map_pid>[0-9]+)(?:/(?P<map_tid>[0-9]+))?:\s+"
    r"\[(?P<start>0x[0-9a-fA-F]+)\((?P<size>0x[0-9a-fA-F]+)\)\s+@\s+"
    r"(?P<offset>(?:0x)?[0-9a-fA-F]+)(?P<meta>.*?)\]:\s+"
    r"(?P<perms>\S+)\s+(?P<path>.*)$"
)


class EvidenceError(RuntimeError):
    pass


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def actual(path: Path) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    data = path.read_bytes()
    return {"path": str(path), "bytes": len(data), "sha256": digest_bytes(data)}


def packet_descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    try:
        relative = path.relative_to(PACKET)
    except ValueError:
        relative = path
    row = actual(path)
    row["path"] = str(relative)
    return row


def check_descriptor(row: Any, label: str) -> dict[str, Any]:
    require(isinstance(row, dict), f"{label}: descriptor is malformed")
    raw_path = row.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}: path missing")
    path = Path(raw_path)
    if not path.is_absolute():
        path = PACKET / path
    require(type(row.get("bytes")) is int and row["bytes"] >= 0,
            f"{label}: byte count malformed")
    require(isinstance(row.get("sha256"), str) and re.fullmatch(r"[0-9a-f]{64}", row["sha256"]),
            f"{label}: digest malformed")
    current = actual(path)
    require(current["bytes"] == row["bytes"] and current["sha256"] == row["sha256"],
            f"{label}: descriptor changed")
    return row


def gzip_bytes(row: Any, label: str) -> bytes:
    checked = check_descriptor(row, label)
    path = Path(checked["path"])
    if not path.is_absolute():
        path = PACKET / path
    try:
        return gzip.decompress(path.read_bytes())
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        fail(f"{label}: invalid gzip: {error}")


def desc_equal(left: Any, right: Any) -> bool:
    return isinstance(left, dict) and isinstance(right, dict) and all(
        left.get(key) == right.get(key) for key in ("path", "bytes", "sha256")
    )


def parse_script(data: bytes) -> tuple[list[dict[str, Any]], list[str]]:
    require(data, "perf script is empty")
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"perf script is not UTF-8: {error}")
    samples: list[dict[str, Any]] = []
    status: list[str] = []
    current: dict[str, Any] | None = None

    def finish() -> None:
        nonlocal current
        if current is not None:
            samples.append(current)
            current = None

    for line_number, line in enumerate(text.splitlines(), 1):
        header = HEADER_RE.fullmatch(line)
        if header:
            finish()
            period = int(header["period"])
            require(period > 0, f"script line {line_number}: non-positive period")
            current = {
                "index": len(samples), "line": line_number,
                "command": header["comm"].strip(), "pid": int(header["pid"]),
                "tid": int(header["tid"]) if header["tid"] else None,
                "cpu": int(header["cpu"]), "timestamp": header["time"],
                "period": period, "event": header["event"], "frames": [],
                "malformed_lines": [], "header": line,
            }
            continue
        if current is None:
            if line.strip():
                status.append(line)
            continue
        if not line.strip():
            finish()
            continue
        frame = FRAME_RE.fullmatch(line)
        if frame is None:
            current["malformed_lines"].append({"line": line_number, "text": line})
            continue
        body = frame["body"].strip()
        offset_match = re.search(r"\+0x([0-9a-fA-F]+)$", body)
        offset = int(offset_match.group(1), 16) if offset_match else None
        if offset_match:
            body = body[:offset_match.start()].rstrip()
        current["frames"].append({
            "address": int(frame["address"], 16), "symbol": body,
            "offset": offset, "dso": frame["dso"].strip(), "text": line,
        })
    finish()
    require(samples, "perf script has no sample records")
    require(all(sample["event"] == "cycles:u" for sample in samples),
            "perf script contains an event other than cycles:u")
    timestamps = [sample["timestamp"] for sample in samples]
    require(len(timestamps) == len(set(timestamps)), "perf sample timestamps repeat")
    return samples, status


def parse_lost(data: bytes) -> list[dict[str, Any]]:
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"lost-event export is not UTF-8: {error}")
    rows: list[dict[str, Any]] = []
    for line_number, line in enumerate(text.splitlines(), 1):
        match = LOST_RE.fullmatch(line)
        require(match is not None, f"lost-event line {line_number}: malformed")
        rows.append({
            "line": line_number, "pid": int(match["pid"]),
            "cpu": int(match["cpu"]), "timestamp": match["time"],
            "count": int(match["count"]), "text": line,
        })
    require(rows, "lost-event export is empty")
    return rows


def parse_mappings(data: bytes) -> list[dict[str, Any]]:
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"mapping export is not UTF-8: {error}")
    rows: list[dict[str, Any]] = []
    for line_number, line in enumerate(text.splitlines(), 1):
        # The mapping export intentionally also contains COMM, sample, and
        # task records because capture.py asked perf to show task events.  The
        # mmap timeline is the only part used for IP admission.
        if "PERF_RECORD_MMAP2" not in line:
            continue
        match = MAPPING_RE.fullmatch(line)
        require(match is not None, f"mapping line {line_number}: malformed")
        buildid_match = re.search(r"<([0-9a-fA-F]+)>", match["meta"])
        start = int(match["start"], 16)
        size = int(match["size"], 16)
        rows.append({
            "line": line_number, "pid": int(match["map_pid"]),
            "timestamp": float(match["time"]), "start": start,
            "end": start + size, "size": size,
            "offset": int(match["offset"], 16),
            "buildid": buildid_match.group(1).lower() if buildid_match else None,
            "perms": match["perms"], "path": match["path"].strip(),
            "text": line,
        })
    require(rows, "mapping export is empty")
    return rows


def parse_load_segments(path: Path) -> list[dict[str, int]]:
    """Read the retained readelf PT_LOAD table for PIE file-offset mapping."""

    rows: list[dict[str, int]] = []
    load_re = re.compile(
        r"^\s*LOAD\s+(0x[0-9a-fA-F]+)\s+(0x[0-9a-fA-F]+)\s+"
        r"(0x[0-9a-fA-F]+)\s+(0x[0-9a-fA-F]+)\s+(0x[0-9a-fA-F]+)\s+"
        r"[RWE](?:\s*[RWE])*\s+"
        r"(0x[0-9a-fA-F]+)\s*$"
    )
    for line in path.read_text(encoding="utf-8").splitlines():
        match = load_re.fullmatch(line)
        if match:
            rows.append({
                "offset": int(match[1], 16), "vaddr": int(match[2], 16),
                "filesz": int(match[4], 16), "memsz": int(match[5], 16),
                "align": int(match[6], 16),
            })
    require(rows, "retained ELF has no PT_LOAD rows")
    return rows


def static_address(frame_address: int, mapping: dict[str, Any],
                   segments: list[dict[str, int]]) -> int | None:
    """Translate a runtime IP through the PIE mmap and the PT_LOAD table."""

    page = 0x1000
    map_offset_page = mapping["offset"] & ~(page - 1)
    matches: list[dict[str, int]] = []
    for segment in segments:
        segment_offset_page = segment["offset"] & ~(page - 1)
        segment_end_page = (segment["offset"] + segment["filesz"] + page - 1) & ~(page - 1)
        if segment_offset_page <= map_offset_page < segment_end_page:
            matches.append(segment)
    if matches:
        # Adjacent PT_LOAD pages can overlap at a page boundary after
        # alignment.  The segment with the greatest file offset owns that
        # boundary in the recorded mmap.
        segment = max(matches, key=lambda row: row["offset"])
        vaddr_page = segment["vaddr"] & ~(page - 1)
        return frame_address - mapping["start"] + vaddr_page
    return None


def parse_ranges(path: Path, binary: dict[str, Any]) -> dict[str, dict[str, Any]]:
    value = read_json(path)
    require(desc_equal(value.get("binary"), binary), "symbol range binary differs")
    result: dict[str, dict[str, Any]] = {}
    for row in value.get("ranges", []):
        require(isinstance(row, dict) and isinstance(row.get("name"), str),
                "symbol range row malformed")
        require(row["name"] not in result, "duplicate bound symbol range")
        require(type(row.get("address")) is int and type(row.get("size")) is int
                and type(row.get("end")) is int and row["size"] > 0
                and row["end"] == row["address"] + row["size"],
                f"invalid range {row.get('name')}")
        result[row["name"]] = row
    expected = {OWNER, PARENT, SETUP_PARENT, "sha2::sha256::compress256", *BACKENDS}
    require(set(result) == expected, "bound symbol set changed")
    return result


def resolve_frame(frame: dict[str, Any], sample: dict[str, Any], mappings: list[dict[str, Any]],
                  ranges: dict[str, dict[str, Any]], binary_path: str,
                  segments: list[dict[str, int]]) -> dict[str, Any]:
    candidates = [row for row in mappings
                  if row["pid"] == sample["pid"] and row["path"] == binary_path
                  and row["start"] <= frame["address"] < row["end"]
                  and row["timestamp"] <= float(sample["timestamp"])]
    if not candidates:
        return {"valid": False, "reason": "no matching binary mmap"}
    mapping = max(candidates, key=lambda row: row["timestamp"])
    static_address_value = static_address(frame["address"], mapping, segments)
    if static_address_value is None:
        return {"valid": False, "reason": "mmap file offset has no PT_LOAD translation"}
    translated = static_address_value
    wanted = ranges[frame["symbol"]]
    valid = wanted["address"] <= translated < wanted["end"]
    return {
        "valid": valid,
        "reason": "ok" if valid else "address outside bound ELF range",
        "static_address": translated,
        "mapping": {key: mapping[key] for key in ("line", "start", "end", "offset", "buildid", "path")},
        "range": wanted,
    }


def frame_row(frame: dict[str, Any], resolution: dict[str, Any] | None = None) -> dict[str, Any]:
    row = {"address": f"0x{frame['address']:x}", "symbol": frame["symbol"],
           "offset": frame["offset"], "dso": frame["dso"]}
    if resolution is not None:
        row["resolution"] = resolution
    return row


def is_unknown(frame: dict[str, Any]) -> bool:
    value = f"{frame['symbol']} {frame['dso']}".lower()
    return any(token in value for token in ("[unknown]", "<unknown>", "??"))


def reduce_samples(samples: list[dict[str, Any]], mappings: list[dict[str, Any]],
                   ranges: dict[str, dict[str, Any]], binary_path: str,
                   segments: list[dict[str, int]]) -> dict[str, Any]:
    counts: Counter[str] = Counter()
    periods: Counter[str] = Counter()
    depths: Counter[int] = Counter()
    symbols: Counter[str] = Counter()
    symbol_periods: Counter[str] = Counter()
    backend_counts = {
        backend: {scope: {"samples": 0, "period": 0} for scope in
                  ("measured_owner", "setup_owner", "unclassified_owner", "outside_owner")}
        for backend in BACKENDS
    }
    backend_hits: list[dict[str, Any]] = []
    named: Counter[str] = Counter()
    invalid: Counter[str] = Counter()
    owner_named = 0
    owner_invalid = 0
    owner_missing = 0
    unknown_samples = 0
    unknown_frames = 0
    malformed_samples = 0
    depth_limit = 0

    for sample in samples:
        frames = sample["frames"]
        period = sample["period"]
        depths[len(frames)] += 1
        if frames:
            counts["callchain_present"] += 1
        else:
            counts["empty_stack"] += 1
        if len(frames) >= 256:
            depth_limit += 1
        malformed_samples += int(bool(sample["malformed_lines"]))
        unknown_here = sum(is_unknown(frame) for frame in frames)
        unknown_frames += unknown_here
        unknown_samples += int(unknown_here > 0)
        for frame in frames:
            symbols[frame["symbol"]] += 1
            symbol_periods[frame["symbol"]] += period

        exact: dict[str, list[tuple[dict[str, Any], dict[str, Any]]]] = {}
        targets = (OWNER, PARENT, SETUP_PARENT, "sha2::sha256::compress256", *BACKENDS)
        for target in targets:
            exact[target] = []
            for frame in frames:
                if frame["symbol"] != target:
                    continue
                named[target] += 1
                resolution = resolve_frame(frame, sample, mappings, ranges, binary_path, segments)
                if resolution["valid"]:
                    exact[target].append((frame, resolution))
                else:
                    invalid[target] += 1

        owners = exact[OWNER]
        if len(owners) == 1:
            owner_named += 1
            owner_index = next(index for index, frame in enumerate(frames)
                               if frame is owners[0][0])
            after = [frame["symbol"] for frame in frames[owner_index + 1:]]
            has_parent = PARENT in after
            has_setup = SETUP_PARENT in after
            if has_parent and not has_setup:
                scope = "measured_owner"
            elif has_setup and not has_parent:
                scope = "setup_owner"
            else:
                scope = "unclassified_owner"
        elif len(owners) == 0:
            scope = "outside_owner"
            owner_missing += 1
            owner_named_frames = sum(frame["symbol"] == OWNER for frame in frames)
            if owner_named_frames:
                owner_invalid += 1
        else:
            scope = "unclassified_owner"
        counts[scope] += 1
        periods[scope] += period

        for backend in BACKENDS:
            hits = exact[backend]
            require(len(hits) <= 1, f"sample {sample['index']}: repeated backend {backend}")
            if not hits:
                continue
            backend_counts[backend][scope]["samples"] += 1
            backend_counts[backend][scope]["period"] += period
            owner_after = []
            setup_after = []
            if owners:
                owner_index = next(index for index, frame in enumerate(frames)
                                   if frame is owners[0][0])
                owner_after = [frame["symbol"] for frame in frames[owner_index + 1:]]
                setup_after = owner_after
            backend_hits.append({
                "sample_index": sample["index"], "timestamp": sample["timestamp"],
                "period": period, "backend": backend, "scope": scope,
                "owner_present": bool(owners),
                "parent_present_after_owner": PARENT in owner_after,
                "setup_parent_present_after_owner": SETUP_PARENT in setup_after,
                "reason": "exact backend frame lacks owner and measured-parent scope frames"
                if scope == "outside_owner" else "backend frame is inside classified owner scope",
                "frames": [
                    frame_row(frame, next((resolution for candidate, resolution in
                                           exact.get(frame["symbol"], []) if candidate is frame), None))
                    for frame in frames
                ],
            })

    sample_period = sum(sample["period"] for sample in samples)
    return {
        "sample_records": len(samples),
        "user_samples": sum(sample["event"] == "cycles:u" for sample in samples),
        "sample_period": sample_period,
        "user_period": sum(sample["period"] for sample in samples
                            if sample["event"] == "cycles:u"),
        "callchains_present": counts["callchain_present"],
        "empty_stack": counts["empty_stack"],
        "malformed_frame_samples": malformed_samples,
        "unknown_frame_samples": unknown_samples,
        "unknown_frame_count": unknown_frames,
        "depth_histogram": {str(depth): depths[depth] for depth in sorted(depths)},
        "max_stack_depth": max(depths, default=0),
        "depth_limit_samples": depth_limit,
        "owner_scope": {
            scope: {"samples": counts[scope], "period": periods[scope]}
            for scope in ("measured_owner", "setup_owner", "unclassified_owner", "outside_owner")
        } | {
            "exact_owner_named_samples": owner_named,
            "exact_owner_invalid_range_samples": owner_invalid,
            "owner_missing_samples": owner_missing,
        },
        "target_frame_diagnostics": {
            "named": dict(sorted(named.items())),
            "range_invalid": dict(sorted(invalid.items())),
            "exact_measured_parent_samples": sum(
                any(frame["symbol"] == PARENT for frame in sample["frames"])
                for sample in samples
            ),
            "exact_setup_parent_samples": sum(
                any(frame["symbol"] == SETUP_PARENT for frame in sample["frames"])
                for sample in samples
            ),
        },
        "backend_observations": backend_counts,
        "backend_hits": backend_hits,
        "top_symbols_by_period": [
            {"symbol": symbol, "samples": symbols[symbol], "period": symbol_periods[symbol]}
            for symbol in sorted(symbol_periods, key=lambda name: (-symbol_periods[name], name))[:25]
        ],
        "scope_diagnostics": {
            "backend_hits_missing_owner_or_parent": sum(
                row["scope"] == "outside_owner" for row in backend_hits
            ),
            "explicit_depth_limit_marker": False,
            "depth_limit_samples": depth_limit,
            "interpretation": "Missing owner or parent frames are retained as scope evidence; no cause is inferred from the sampled stack alone.",
        },
    }


def verify_capture() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    plan = read_json(PACKET / "plan.json")
    require(plan.get("schema") == "litchi.native-snapshot-profile-0912.v1", "plan schema changed")
    require(plan.get("cpu") == 12 and plan.get("samples") == 50 and plan.get("warmup") == 0,
            "pilot plan dimensions changed")
    require(plan.get("event") == "cycles:u" and plan.get("frequency_hz") == 997,
            "pilot perf event changed")
    require(plan.get("owner") == OWNER and plan.get("parent") == PARENT
            and plan.get("setup_parent") == SETUP_PARENT, "scope symbols changed")
    binary = check_descriptor(plan["binary"], "plan.binary")
    check_descriptor(plan["helper"], "plan.helper")
    for index, row in enumerate(plan.get("inputs", [])):
        check_descriptor(row, f"plan.inputs[{index}]")
    for index, row in enumerate(plan.get("tools", [])):
        check_descriptor(row, f"plan.tools[{index}]")
    ranges = parse_ranges(PACKET / "symbol-ranges.json", binary)
    folder = PACKET / "capture" / LABEL
    command_path = folder / "command.json"
    command = read_json(command_path)
    command_desc = packet_descriptor(command_path)
    require(command.get("exit_code") == 0, "pilot command failed")
    require(command.get("row") == plan["schedule"][0], "pilot command row differs")
    require(command.get("cwd") == plan["cwd"] and command.get("environment") == plan["environment"],
            "pilot command environment differs")
    require(desc_equal(command.get("binary"), plan["binary"]), "pilot binary binding differs")
    require(desc_equal(command.get("helper"), plan["helper"]), "pilot helper binding differs")
    plan_descriptor = packet_descriptor(PACKET / "plan.json")
    require(isinstance(command.get("plan"), dict)
            and command["plan"].get("bytes") == plan_descriptor["bytes"]
            and command["plan"].get("sha256") == plan_descriptor["sha256"],
            "pilot plan receipt differs")
    raw = folder / "perf.data"
    expected_argv = [
        "/usr/bin/taskset", "-c", "12", "/usr/bin/perf", *plan["options"],
        "-o", str(raw), "--", plan["binary"]["path"], ROUTE, "0", "50",
        str(folder / "artifacts"),
    ]
    require(command.get("argv") == expected_argv, "pilot argv differs from frozen plan")
    for index, row in enumerate(command.get("artifacts", [])):
        check_descriptor(row, f"command.artifacts[{index}]")
    require(any(row.get("path") == str(raw) for row in command["artifacts"]),
            "raw perf descriptor missing from command receipt")
    start_path = folder / "start.json"
    start = read_json(start_path)
    require(start.get("argv") == expected_argv and start.get("cwd") == plan["cwd"]
            and start.get("environment") == plan["environment"], "start receipt differs")
    check_descriptor(packet_descriptor(start_path), "start receipt")
    semantics_path = folder / "semantics.json"
    semantics = read_json(semantics_path)
    require(semantics.get("status") == "pass" and semantics.get("route") == ROUTE
            and semantics.get("operations") == 50, "semantic receipt did not pass")
    check_descriptor(packet_descriptor(semantics_path), "semantics receipt")
    report_path = folder / "artifacts" / "report.json"
    report = read_json(report_path)
    require(report.get("status") == "completed" and report.get("route") == ROUTE
            and report.get("warmup_iterations") == 0 and report.get("measured_samples") == 50,
            "workload report dimensions changed")
    require(all(report.get("gates", {}).values()), "workload gate failed")
    check_descriptor(packet_descriptor(report_path), "workload report")
    for name in ("source", "output"):
        path = folder / "artifacts" / report[name]["file"]
        identity = actual(path)
        require(identity["bytes"] == report[name]["bytes"]
                and identity["sha256"] == report[name]["sha256"],
                f"workload {name} identity changed")
    decode = read_json(folder / "decode.json")
    require(decode.get("exit_code") == 0 and decode.get("argv") == [
        "/usr/bin/perf", *plan["decode_options"], "-i", str(raw)
    ], "decode receipt differs")
    check_descriptor(decode["raw"], "decode.raw")
    check_descriptor(decode["script"], "decode.script")
    check_descriptor(decode["stderr"], "decode.stderr")
    script_data = gzip_bytes(decode["script"], "decoded script")
    require(len(script_data) == decode["uncompressed"]["bytes"]
            and digest_bytes(script_data) == decode["uncompressed"]["sha256"],
            "decoded script identity changed")
    debug = read_json(folder / "sample-debug.json")
    require(debug.get("exit_code") == 0, "sample-debug export failed")
    check_descriptor(debug["stdout"], "sample-debug.stdout")
    check_descriptor(debug["stderr"], "sample-debug.stderr")
    check_descriptor(debug["raw"], "sample-debug.raw")
    lost_data = gzip_bytes(debug["stdout"], "sample-debug stdout")
    mappings = read_json(folder / "mappings.json")
    require(mappings.get("exit_code") == 0, "mapping export failed")
    check_descriptor(mappings["stdout"], "mappings.stdout")
    check_descriptor(mappings["stderr"], "mappings.stderr")
    check_descriptor(mappings["raw"], "mappings.raw")
    mapping_data = gzip_bytes(mappings["stdout"], "mappings stdout")
    buildids = read_json(folder / "buildids.json")
    require(buildids.get("exit_code") == 0, "build-id export failed")
    check_descriptor(buildids["stdout"], "buildids.stdout")
    check_descriptor(buildids["stderr"], "buildids.stderr")
    check_descriptor(buildids["raw"], "buildids.raw")
    buildid_data = gzip_bytes(buildids["stdout"], "buildids stdout")
    elf_segments_path = PACKET / "elf-segments.txt"
    elf_segments = parse_load_segments(elf_segments_path)
    check_descriptor(packet_descriptor(elf_segments_path), "ELF segment table")
    return plan, report, {
        "binary": binary, "ranges": ranges, "segments": elf_segments, "command": command,
        "command_descriptor": command_desc, "semantics": semantics,
        "semantics_path": semantics_path, "report_path": report_path,
        "decode": decode, "debug": debug, "mappings": mappings, "buildids": buildids,
        "script_data": script_data, "lost_data": lost_data,
        "mapping_data": mapping_data, "buildid_data": buildid_data,
        "raw": packet_descriptor(raw), "script": packet_descriptor(folder / "script.txt.gz"),
    }


def build_result() -> dict[str, Any]:
    plan, report, files = verify_capture()
    samples, status_lines = parse_script(files["script_data"])
    lost = parse_lost(files["lost_data"])
    mappings = parse_mappings(files["mapping_data"])
    binary_path = plan["binary"]["path"]
    buildid_lines = files["buildid_data"].decode("utf-8").splitlines()
    binary_buildids = [line.split()[0].lower() for line in buildid_lines
                       if len(line.split()) == 2 and line.split()[1] == binary_path]
    require(len(binary_buildids) == 1, "binary build-id binding is missing or ambiguous")
    map_buildids = sorted({row["buildid"] for row in mappings
                           if row["path"] == binary_path and row["buildid"]})
    require(map_buildids == binary_buildids, "binary mmap build-id differs from buildid-list")
    reduced = reduce_samples(samples, mappings, files["ranges"], binary_path, files["segments"])
    lost_count = sum(row["count"] for row in lost)
    owner = reduced["owner_scope"]
    return {
        "schema": ANALYSIS_SCHEMA,
        "status": "rejected",
        "attempt": "perf-pilot",
        "label": LABEL,
        "route": ROUTE,
        "reason_code": "lost_perf_records_and_missing_snapshot_scope",
        "reason": (
            "The pilot reports six lost perf records totaling 304 lost events. "
            "No exact opened_presentation_transaction owner frame appears in the "
            "decoded samples; the one exact x86 SHA backend sample is outside the "
            "owner and timed_iteration scope. The child is retained but cannot "
            "admit a native snapshot attribution."
        ),
        "protocol": {
            "schema": plan["schema"], "cpu": plan["cpu"], "event": plan["event"],
            "frequency_hz": plan["frequency_hz"], "stack_bytes": plan["stack_bytes"],
            "script_max_stack": plan["script_max_stack"], "options": plan["options"],
            "decode_options": plan["decode_options"], "owner": OWNER,
            "required_parent": PARENT, "setup_parent": SETUP_PARENT,
            "warmup": plan["warmup"], "samples": plan["samples"],
        },
        "bindings": {
            "evidence_parent": plan["evidence_parent"],
            "binary": plan["binary"],
            "plan": packet_descriptor(PACKET / "plan.json"),
            "capture_helper": plan["helper"],
            "protocol_review": packet_descriptor(PACKET / "protocol-review.md"),
            "symbols": packet_descriptor(PACKET / "symbols.txt"),
            "symbol_ranges": packet_descriptor(PACKET / "symbol-ranges.json"),
            "elf_segments": packet_descriptor(PACKET / "elf-segments.txt"),
            "perf": next(row for row in plan["tools"] if row["path"] == "/usr/bin/perf"),
        },
        "command": {
            "path": files["command_descriptor"]["path"],
            "bytes": files["command_descriptor"]["bytes"],
            "sha256": files["command_descriptor"]["sha256"],
            "exit_code": files["command"]["exit_code"],
            "argv_matches_plan": True, "binary_matches_plan": True,
            "environment_matches_plan": True, "cpu_affinity_matches_plan": True,
            "raw": files["raw"], "script": files["script"],
            "decode": packet_descriptor(PACKET / "capture" / LABEL / "decode.json"),
            "sample_debug": packet_descriptor(PACKET / "capture" / LABEL / "sample-debug.json"),
            "mappings": packet_descriptor(PACKET / "capture" / LABEL / "mappings.json"),
            "buildids": packet_descriptor(PACKET / "capture" / LABEL / "buildids.json"),
        },
        "workload": {
            "semantics": packet_descriptor(files["semantics_path"]),
            "report": packet_descriptor(files["report_path"]), "route": ROUTE,
            "operations": 50, "source": report["source"], "output": report["output"],
            "semantic_and_corpus_gates_pass": True,
        },
        "native": {
            "raw_script": {"samples": reduced["sample_records"],
                           "period": reduced["sample_period"],
                           "status_lines": status_lines},
            "lost_records": len(lost), "lost_events": lost_count,
            "lost_rows": lost, "reduction": reduced,
            "binary_build_id": binary_buildids[0], "mmap_binary_build_ids": map_buildids,
            "mapping_records": len(mappings),
            "mapping_binary_records": sum(row["path"] == binary_path for row in mappings),
        },
        "checks": {
            "binary_hash_bound": True, "symbol_ranges_bound": True,
            "command_and_decode_custody": True, "workload_semantic_gates": True,
            "script_event_is_exactly_cycles_u": True, "sample_periods_positive": True,
            "mapping_build_id_bound": True, "lost_records_zero": False,
            "exact_owner_frame_seen": owner["exact_owner_named_samples"] > 0,
            "exact_measured_parent_frame_seen": reduced["target_frame_diagnostics"]["exact_measured_parent_samples"] > 0,
            "measured_backend_scope_proven": owner["measured_owner"]["samples"] > 0,
            "admission": False,
        },
        "policy": {
            "lost_records_block_scoped_attribution": True,
            "outside_owner_backend_is_not_a_dispatch_claim": True,
            "no_native_elapsed_or_guest_instruction_claim": True,
            "no_route_comparison": True,
            "remaining_perf_legs": "stopped_after_first_child_failure",
            "gdb_fallback": "separate reachability-only lane; not a replacement perf result",
        },
    }


def render_markdown(result: dict[str, Any]) -> str:
    native = result["native"]
    reduction = native["reduction"]
    owner = reduction["owner_scope"]
    backend = reduction["backend_observations"]
    backend_rows = []
    for name in BACKENDS:
        backend_rows.append(
            f"| `{name}` | {backend[name]['measured_owner']['samples']} | "
            f"{backend[name]['setup_owner']['samples']} | "
            f"{backend[name]['unclassified_owner']['samples']} | "
            f"{backend[name]['outside_owner']['samples']} | "
            f"{backend[name]['outside_owner']['period']} |"
        )
    lost_rows = "\n".join(
        f"| {row['timestamp']} | {row['count']} |" for row in native["lost_rows"]
    )
    return (
        f"# 0912 first-child native perf failure\n\n"
        f"The first native pilot is retained and rejected. Its workload completed "
        f"the same 50-operation materialized route guards, but perf reported "
        f"**{native['lost_records']} lost records** totaling **{native['lost_events']} "
        f"lost events**. The decoded sample stream contains no exact "
        f"`{OWNER}` frame, so it cannot establish the measured snapshot interval.\n\n"
        f"The pilot recorded `{reduction['sample_records']}` `cycles:u` sample records "
        f"with **{reduction['callchains_present']}** nonempty callchains and "
        f"**{reduction['empty_stack']}** empty stacks. Their sampled periods sum to "
        f"`{reduction['sample_period']}`. These are sampled event periods, not "
        f"elapsed time, native cycle totals, guest instructions, or a route comparison.\n\n"
        f"| owner scope | samples | period |\n| --- | ---: | ---: |\n"
        f"| measured owner | {owner['measured_owner']['samples']} | {owner['measured_owner']['period']} |\n"
        f"| setup owner | {owner['setup_owner']['samples']} | {owner['setup_owner']['period']} |\n"
        f"| unclassified owner | {owner['unclassified_owner']['samples']} | {owner['unclassified_owner']['period']} |\n"
        f"| outside owner | {owner['outside_owner']['samples']} | {owner['outside_owner']['period']} |\n\n"
        f"The one exact `x86_sha::compress` frame has period **4,081,079**, but its "
        f"stack contains neither the snapshot owner nor `timed_iteration`; it is "
        f"retained as an outside-owner diagnostic. No software-backend frame was observed.\n\n"
        f"| backend | measured samples | setup samples | unclassified samples | outside samples | outside period |\n"
        f"| --- | ---: | ---: | ---: | ---: | ---: |\n"
        + "\n".join(backend_rows) + "\n\n"
        f"The stream has **{reduction['unknown_frame_count']}** unknown frames across "
        f"**{reduction['unknown_frame_samples']}** samples and no explicit depth-limit "
        f"marker. The backend candidate itself is an incomplete scope chain; the reader "
        f"does not infer whether missing frames arose from sampling or DWARF unwinding.\n\n"
        f"| timestamp | lost events |\n| --- | ---: |\n{lost_rows}\n\n"
        f"The exact raw `perf.data`, decoded script, mmap/build-id exports, command "
        f"receipt, and semantic artifacts remain bound in "
        f"[`first-child-failure.json`](first-child-failure.json). The native attribution "
        f"lane stops here; any GDB follow-up can establish reachability only and cannot "
        f"replace this rejected perf result or create a timing claim.\n"
    )


def main() -> None:
    output = PACKET / "first-child-failure.json"
    markdown = PACKET / "first-child-failure.md"
    result = build_result()
    encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
    rendered = render_markdown(result)
    if "--check" in sys.argv:
        require(output.read_text(encoding="utf-8") == encoded, "failure JSON is not reproducible")
        require(markdown.read_text(encoding="utf-8") == rendered, "failure Markdown is not reproducible")
    else:
        require(not output.exists() and not markdown.exists(), "failure outputs already exist")
        output.write_text(encoded, encoding="utf-8")
        markdown.write_text(rendered, encoding="utf-8")
    print("0912 first-child native perf failure reduction PASS", flush=True)


if __name__ == "__main__":
    try:
        main()
    except EvidenceError as error:
        print(f"analyze.py: {error}", file=sys.stderr)
        raise SystemExit(1)
