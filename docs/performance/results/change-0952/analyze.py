#!/usr/bin/env python3
"""Offline native admission for the fixed 0952 packet.

This reader consumes only sealed child artifacts plus the separate sample and
mmap exports.  It never starts the benchmark or perf.  Runtime addresses are
bound through recorded MMAP2 timestamps, the executable build-id, and ELF
PT_LOAD segments before static owner/parent ranges are tested.  All samples,
periods, lost records, unresolved frames, and non-owner observations remain in
the denominator.  A failed owner gate is emitted as a durable result and is
never converted into an admission.
This v3 reader is deliberately bound to the packet constants below; its
``--packet`` spelling is retained only for command compatibility and is not a
portable packet selector.
"""
from __future__ import annotations

import argparse
import collections
import decimal
import hashlib
import importlib.util
import json
import os
import re
import stat
import struct
import sys
from pathlib import Path
from typing import Any, Iterable

PACKET = Path("/home/zhuhe/litchi-performance-0952")
PLAN = PACKET / "plan-0952-v1.json"
PLAN_SHA = "7846288a54c6c5eebb97f958631d9e5fe4345ab08dd1c7527e94e4a0d751beed"
RUNNER = PACKET / "tools/capture_0952_v1.py"
RUNNER_SHA = "c517215f41ea0b7097d6b8fce88cc93317755681df9fc90e4b37475f9eae50da"
SEMANTIC = PACKET / "tools/semantic_0952_v1.py"
SEMANTIC_SHA = "ba7f5fc5e626d0c5e48655f18572765c3a91fab9fcce0c455b2ef7f97b04d325"
SAMPLE_EXPORTER = PACKET / "tools/export_samples_0952_v1.py"
SAMPLE_EXPORTER_SHA = "7a077649af0e847117bf35177d9e9d71e7a7a1d5fa20af0b6a0536eb0e2a0214"
MMAP_EXPORTER = PACKET / "tools/export_mmaps_0952_v2.py"
MMAP_EXPORTER_SHA = "0942acfb4ae28c2ddb29fe8f124aa2f2917ff32cabc2b723a375d192e8b122ca"
CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
EVENT = "cycles:u"
PAGE = 4096
OWNER_MIN_SAMPLES = 100
OWNER_MIN_PERIOD = 100_000
DECIMAL = decimal.Decimal
HEX = re.compile(r"^(?:0x)?[0-9a-fA-F]+$")
SAMPLE_HEADER = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>\d+)/(?P<tid>\d+)\s+"
    r"(?P<time>[0-9]+(?:\.[0-9]+)?)\s*:\s*(?P<period>\d+)\s+(?P<event>\S+)"
)
FRAME = re.compile(
    r"^\s*(?P<ip>(?:0x)?[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+"
    r"\((?P<dso>[^)]*)\)(?:\s+\d+)?\s*$"
)
MMAP = re.compile(
    r"^\s*\S+\s+\d+/\d+\s+(?P<header_time>[0-9]+(?:\.[0-9]+)?)\s*:\s+"
    r"PERF_RECORD_MMAP2\s+(?P<pid>\d+)/(?P<tid>\d+)\s*:\s+"
    r"\[(?P<start>0x[0-9a-fA-F]+)\((?P<size>0x[0-9a-fA-F]+)\)\s+@\s+"
    r"(?P<offset>(?:0x)?[0-9a-fA-F]+)\s+(?P<meta>.*?)\]:\s+"
    r"(?P<perm>\S+)\s+(?P<dso>.*)$"
)
UNKNOWN = {"", "??", "[unknown]", "unknown"}


class AdmissionError(RuntimeError):
    pass


def require(value: bool, message: str) -> None:
    if not value:
        raise AdmissionError(message)


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    st = path.stat()
    return {"path": str(path), "bytes": st.st_size, "sha256": sha(path),
            "mode": stat.S_IMODE(st.st_mode), "executable": bool(st.st_mode & stat.S_IXUSR)}


def unique_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in out, f"duplicate JSON key: {key}")
        out[key] = value
    return out


def load_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=unique_pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise AdmissionError(f"invalid JSON {path}: {exc}") from exc


def verify_descriptor(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label} descriptor is missing")
    raw = value.get("path")
    require(isinstance(raw, str) and raw and not "\x00" in raw, f"{label} path is invalid")
    path = Path(raw)
    actual = descriptor(path)
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        require(actual[key] == value.get(key), f"{label} descriptor {key} differs")
    return path.resolve()


def import_semantic() -> Any:
    require(sha(SEMANTIC) == SEMANTIC_SHA, "semantic adapter digest differs")
    spec = importlib.util.spec_from_file_location("semantic_0952_bound", SEMANTIC)
    require(spec is not None and spec.loader is not None, "semantic adapter cannot load")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def verify_bindings() -> tuple[dict[str, Any], Any]:
    require(sha(PLAN) == PLAN_SHA, "plan digest differs")
    require(sha(RUNNER) == RUNNER_SHA, "capture runner digest differs")
    require(sha(SAMPLE_EXPORTER) == SAMPLE_EXPORTER_SHA, "sample exporter digest differs")
    require(sha(MMAP_EXPORTER) == MMAP_EXPORTER_SHA, "mmap exporter digest differs")
    plan = load_json(PLAN)
    require(plan.get("schema") == "litchi.performance.0952.native-publish-sampling-plan.v1", "plan schema differs")
    binary_path = verify_descriptor(plan.get("binary"), "plan binary")
    require(plan.get("samples") == 200 and plan.get("warmup") == 0, "plan sample contract differs")
    require(plan.get("case") == CASE and plan.get("cpu") == 12, "plan workload contract differs")
    sem = import_semantic()
    sem_bindings = sem.verify_bindings()
    require(sem_bindings["plan"]["sha256"] == PLAN_SHA, "semantic plan binding differs")
    require(sem_bindings["runner"]["sha256"] == RUNNER_SHA, "semantic runner binding differs")
    return {"plan": descriptor(PLAN), "runner": descriptor(RUNNER),
            "sample_exporter": descriptor(SAMPLE_EXPORTER),
            "mmap_exporter": descriptor(MMAP_EXPORTER),
            "binary": descriptor(binary_path), "semantic": descriptor(SEMANTIC)}, sem


def terminal_pass(path: Path, label: str) -> dict[str, Any]:
    value = load_json(path)
    require(value.get("returncode") == 0 and value.get("timed_out") is False,
            f"{label} terminal did not pass")
    require(value.get("launch_error") is None and value.get("process_group_clear") is True
            and value.get("terminal_before_parse") is True, f"{label} terminal custody failed")
    return value


def equal_desc(left: Any, right: Any, label: str) -> None:
    require(isinstance(left, dict) and isinstance(right, dict), f"{label} descriptor missing")
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        require(left.get(key) == right.get(key), f"{label} descriptor {key} differs")


def child_context(plan: dict[str, Any], index: int) -> tuple[dict[str, Any], Path, dict[str, Any]]:
    rows = {row["index"]: row for row in plan.get("schedule", [])}
    require(index in rows, f"schedule index {index} missing")
    row = rows[index]
    child = PACKET / "capture" / f"{index:02d}-{row['shape']}"
    receipt = load_json(child / "receipt.json")
    require(receipt.get("status") == "raw_captured" and receipt.get("retry") is False,
            "child raw receipt is not admitted")
    require(receipt.get("row") == row, "child schedule row differs")
    require(receipt.get("binary", {}).get("sha256") == plan["binary"]["sha256"], "child binary differs")
    terminal_record = terminal_pass(child / "record-terminal.json", "record")
    terminal_script = terminal_pass(child / "script-terminal.json", "original perf script")
    # The original primary export is retained even if debug-mode made it empty;
    # its terminal remains a historical custody record.
    for key, name in (("report", "report.json"), ("perf_data", "perf.data"), ("perf_script", "perf-script.txt"),
                      ("record_terminal", "record-terminal.json"), ("script_terminal", "script-terminal.json")):
        actual = descriptor(child / name)
        equal_desc(receipt.get(key), actual, f"child {key}")
    return row, child, {"receipt": receipt, "record_terminal": terminal_record, "script_terminal": terminal_script}


def export_context(root: Path, child: Path, receipt: dict[str, Any], row: dict[str, Any], kind: str) -> dict[str, Any]:
    value = load_json(root / "receipt.json")
    require(value.get("schema") == f"litchi.performance.0952.{kind}-export-receipt.v1", f"{kind} export schema differs")
    require(value.get("status") == "pass" and value.get("terminal_before_parse") is True
            and value.get("workload_launched") is False, f"{kind} export did not pass")
    require(value.get("row") == row, f"{kind} export schedule row differs")
    require(not root.resolve().is_relative_to(PACKET / "capture"),
            f"{kind} export is inside sealed capture root")
    equal_desc(value.get("perf_data"), receipt.get("perf_data"), f"{kind} perf.data")
    for key, name in (("start", "start.json"), ("terminal", "terminal.json"),
                      ("events", "events.txt"), ("stderr", "stderr.txt")):
        actual = descriptor(root / name)
        equal_desc(value.get(key), actual, f"{kind} {key}")
    terminal = terminal_pass(root / "terminal.json", f"{kind} export")
    start = load_json(root / "start.json")
    argv = start.get("argv")
    require(isinstance(argv, list), f"{kind} export argv missing")
    require(start.get("child_receipt", {}).get("sha256") == sha(child / "receipt.json"),
            f"{kind} child receipt binding differs")
    helper_expected = SAMPLE_EXPORTER_SHA if kind == "sample" else MMAP_EXPORTER_SHA
    require(start.get("helper", {}).get("sha256") == helper_expected, f"{kind} helper binding differs")
    require(start.get("perf_data", {}).get("sha256") == receipt["perf_data"]["sha256"],
            f"{kind} start perf.data binding differs")
    require("--debug-mode" in argv if kind == "mmap" else "--debug-mode" not in argv,
            f"{kind} debug-mode contract differs")
    require("--show-mmap-events" in argv if kind == "mmap" else "--show-mmap-events" not in argv,
            f"{kind} mmap-event contract differs")
    require("--show-task-events" in argv if kind == "mmap" else "--show-task-events" not in argv,
            f"{kind} task-event contract differs")
    require("--hide-call-graph" in argv if kind == "mmap" else "--hide-call-graph" not in argv,
            f"{kind} call-graph contract differs")
    require("--show-lost-events" in argv and "--ns" in argv and "--no-inline" in argv,
            f"{kind} event-field contract differs")
    require(terminal.get("argv") == argv, f"{kind} terminal argv differs")
    events = root / "events.txt"
    require(events.stat().st_size > 0, f"{kind} event export is empty")
    return {"receipt": descriptor(root / "receipt.json"), "start": descriptor(root / "start.json"),
            "terminal": descriptor(root / "terminal.json"), "events": descriptor(events),
            "stderr": descriptor(root / "stderr.txt"), "terminal_value": terminal, "argv": argv}


def read_elf(binary: Path) -> tuple[list[dict[str, int]], str]:
    with binary.open("rb") as stream:
        ident = stream.read(64)
        require(len(ident) == 64 and ident[:4] == b"\x7fELF" and ident[4] == 2 and ident[5] == 1,
                "bound binary is not ELF64 little-endian")
        fields = struct.unpack_from("<HHIQQQIHHHHHH", ident, 16)
        _, _, _, _, phoff, _, _, _, phentsize, phnum, _, _, _ = fields
        require(phentsize >= 56 and phnum > 0, "ELF program header table is invalid")
        stream.seek(phoff)
        headers = [stream.read(phentsize) for _ in range(phnum)]
        loads: list[dict[str, int]] = []
        notes: list[tuple[int, int]] = []
        for raw in headers:
            require(len(raw) >= 56, "truncated ELF program header")
            p_type, p_flags, p_offset, p_vaddr, _, p_filesz, p_memsz, p_align = struct.unpack_from("<IIQQQQQQ", raw, 0)
            if p_type == 1:
                require(p_filesz <= p_memsz and p_offset % PAGE == p_vaddr % PAGE,
                        "PT_LOAD alignment/size is invalid")
                loads.append({"flags": p_flags, "offset": p_offset, "vaddr": p_vaddr,
                              "filesz": p_filesz, "memsz": p_memsz, "align": p_align})
            elif p_type == 4 and p_filesz:
                notes.append((p_offset, p_filesz))
        require(loads, "binary has no PT_LOAD segments")
        build_id = None
        for offset, size in notes:
            stream.seek(offset)
            data = stream.read(size)
            pos = 0
            while pos + 12 <= len(data):
                namesz, descsz, ntype = struct.unpack_from("<III", data, pos)
                pos += 12
                nend = pos + ((namesz + 3) // 4) * 4
                dend = nend + ((descsz + 3) // 4) * 4
                require(dend <= len(data), "truncated ELF note")
                name = data[pos:pos + namesz].rstrip(b"\0")
                desc = data[nend:nend + descsz]
                if name == b"GNU" and ntype == 3:
                    build_id = desc.hex()
                pos = dend
        require(build_id is not None, "GNU build-id is missing")
    return loads, build_id


def align_down(value: int) -> int:
    return value & ~(PAGE - 1)


def align_up(value: int) -> int:
    return (value + PAGE - 1) & ~(PAGE - 1)


def load_bias_for(offset: int, start: int, size: int, perm: str,
                  loads: list[dict[str, int]]) -> tuple[int | None, int | None, bool]:
    page_offset = align_down(offset)
    candidates = []
    for index, seg in enumerate(loads):
        lo = align_down(seg["offset"])
        hi = align_up(seg["offset"] + max(seg["filesz"], seg["memsz"]))
        if lo <= page_offset < hi and page_offset + size > lo:
            mapped_vaddr = seg["vaddr"] + (page_offset - seg["offset"])
            candidates.append((index, start - mapped_vaddr, seg))
    require(candidates, f"mapping file offset {offset:#x} does not bind to PT_LOAD")
    # ELF LOAD segments can share a page at a permission boundary, while the
    # loader may also expose a broad read-only file mapping.  Use mmap
    # permissions to select the compatible segment.  Never fall back to an
    # incompatible candidate.  Executable mappings must bind uniquely because
    # they can participate in static-IP ownership; ambiguous non-executable
    # rows remain present but cannot resolve a code frame.
    required_flags = ((4 if "r" in perm else 0)
                      | (2 if "w" in perm else 0)
                      | (1 if "x" in perm else 0))
    filtered = [item for item in candidates
                if required_flags & ~item[2]["flags"] == 0]
    executable = "x" in perm
    if not filtered:
        # Keep a non-executable loader overlay visible, but it cannot resolve
        # code and therefore must never be selected for static ownership.
        if not executable:
            return None, None, True
        raise AdmissionError(f"mapping permissions do not bind PT_LOAD at {offset:#x}")
    if executable:
        require(len(filtered) == 1,
                f"executable mapping file offset {offset:#x} binds ambiguously to PT_LOAD")
    if len(filtered) != 1:
        return None, None, True
    index, bias, _ = filtered[0]
    return index, bias, False


def parse_mmap_events(path: Path, binary: Path, loads: list[dict[str, int]], build_id: str) -> dict[str, Any]:
    target = str(binary.resolve())
    rows: list[dict[str, Any]] = []
    unsupported = 0
    for number, line in enumerate(path.read_text(encoding="utf-8", errors="strict").splitlines(), 1):
        if not line.strip():
            continue
        match = MMAP.match(line)
        if match is None:
            if "PERF_RECORD_LOST" in line or "PERF_RECORD_FORK" in line or "PERF_RECORD_EXIT" in line or "PERF_RECORD_COMM" in line:
                continue
            unsupported += 1
            continue
        meta = match.group("meta")
        ids = re.findall(r"<([0-9a-fA-F]+)>", meta)
        dso = match.group("dso").strip()
        row = {"line": number, "time": DECIMAL(match.group("header_time")),
               "pid": int(match.group("pid")), "tid": int(match.group("tid")),
               "start": int(match.group("start"), 16), "size": int(match.group("size"), 16),
               "offset": int(match.group("offset"), 16), "perm": match.group("perm"),
               "dso": dso, "build_id": ids[0] if ids else None}
        require(row["size"] > 0 and row["start"] + row["size"] > row["start"], "invalid MMAP2 range")
        row["is_bound_dso"] = dso == target or Path(dso).name == binary.name
        if row["is_bound_dso"]:
            require(row["build_id"] == build_id, f"bound DSO build-id differs at mmap line {number}")
            row["pt_load_index"], row["load_bias"], row["mapping_ambiguous"] = load_bias_for(row["offset"], row["start"], row["size"], row["perm"], loads)
        rows.append(row)
    require(unsupported == 0, f"unsupported mmap records: {unsupported}")
    require(rows, "mmap export contains no MMAP2 rows")
    bound = [row for row in rows if row["is_bound_dso"]]
    require(bound and any("x" in row["perm"] for row in bound), "no executable bound DSO mapping")
    return {"rows": rows, "bound_rows": bound, "line_count": len(rows), "unsupported": unsupported,
            "target_path": target, "build_id": build_id,
            "bound_pids": sorted({row["pid"] for row in bound}),
            "bound_build_ids": sorted({row["build_id"] for row in bound}),
            "ambiguous_bound_mappings": sum(1 for row in bound if row.get("mapping_ambiguous"))}


def map_for_ip(maps: list[dict[str, Any]], pid: int, when: decimal.Decimal, ip: int, dso: str) -> dict[str, Any] | None:
    candidates = [row for row in maps if row["pid"] == pid and row["time"] <= when
                  and row["is_bound_dso"] and row["start"] <= ip < row["start"] + row["size"]
                  and (dso == row["dso"] or Path(dso).name == Path(row["dso"]).name)]
    return max(candidates, key=lambda row: (row["time"], row["line"])) if candidates else None


def parse_frame(line: str) -> dict[str, Any] | None:
    match = FRAME.match(line)
    if match is None:
        stripped = line.strip()
        tokens = stripped.split()
        if not tokens or not HEX.fullmatch(tokens[0]):
            return None
        return {"ip": int(tokens[0], 16), "symbol": tokens[1] if len(tokens) > 1 else "[unknown]",
                "dso": "[unknown]", "unresolved": True}
    symbol = match.group("symbol").strip()
    dso = match.group("dso").strip() or "[unknown]"
    return {"ip": int(match.group("ip"), 16), "symbol": symbol, "dso": dso,
            "unresolved": symbol in UNKNOWN or dso.lower() in UNKNOWN}


def static_ip(frame: dict[str, Any], pid: int, when: decimal.Decimal, maps: list[dict[str, Any]], binary: Path) -> tuple[int | None, dict[str, Any] | None]:
    dso = frame["dso"]
    if dso not in {str(binary), binary.name} and Path(dso).name != binary.name:
        return None, None
    mapping = map_for_ip(maps, pid, when, frame["ip"], dso)
    if (mapping is None or mapping.get("mapping_ambiguous")
            or "x" not in mapping.get("perm", "")
            or mapping.get("load_bias") is None):
        return None, None
    return frame["ip"] - mapping["load_bias"], mapping


def parse_samples(path: Path, binary: Path, mmap: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    symbols = plan["static_symbols"]
    owners = [(int(item["address"], 16), int(item["end_exclusive"], 16)) for item in symbols["owner_ranges"]]
    parent = (int(symbols["parent_range"]["address"], 16), int(symbols["parent_range"]["end_exclusive"], 16))
    maps = mmap["rows"]
    all_samples = all_period = owner_samples = owner_period = 0
    lost_records = 0
    unresolved_period = 0
    unknown_period = 0
    unknown_samples = 0
    reasons: collections.Counter[str] = collections.Counter()
    frame_lengths: collections.Counter[int] = collections.Counter()
    leaves: collections.Counter[str] = collections.Counter()
    nested: collections.Counter[str] = collections.Counter()
    current: dict[str, Any] | None = None

    def finish(record: dict[str, Any]) -> None:
        nonlocal owner_samples, owner_period, unknown_samples, unknown_period, unresolved_period
        frames = record["frames"]
        frame_lengths[len(frames)] += 1
        resolved: list[tuple[int | None, dict[str, Any] | None]] = []
        for frame in frames:
            static, mapping = static_ip(frame, record["pid"], record["time"], maps, binary)
            resolved.append((static, mapping))
        owner_positions = [i for i, (static, mapping) in enumerate(resolved)
                           if static is not None and any(lo <= static < hi for lo, hi in owners)]
        reason = "owner_not_observed"
        qualified = False
        if any(frame.get("unresolved") for frame in frames) or record.get("truncated"):
            unresolved_period += record["period"]
            reason = "unresolved_or_truncated"
        if owner_positions:
            oi = owner_positions[0]
            if oi + 1 >= len(frames):
                reason = "owner_without_parent"
            else:
                parent_static, parent_mapping = resolved[oi + 1]
                if parent_static is None or not (parent[0] <= parent_static < parent[1]):
                    reason = "wrong_or_missing_parent"
                elif not frames[oi + 1].get("unresolved"):
                    qualified = True
                    reason = "owner_parent_bound"
        reasons[reason] += 1
        if qualified and not any(frame.get("unresolved") for frame in frames) and not record.get("truncated"):
            owner_samples += 1
            owner_period += record["period"]
            leaves[frames[0]["symbol"]] += record["period"]
            for frame in frames:
                symbol = frame["symbol"]
                if any(token in symbol for token in ("zlib", "deflate", "xml_minifier", "quick_xml", "soapberry_zip", "litchi_opc")):
                    nested[symbol] += record["period"]
        else:
            unknown_samples += 1
            unknown_period += record["period"]

    with path.open("r", encoding="utf-8", errors="strict") as stream:
        for line_number, line in enumerate(stream, 1):
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "PERF_RECORD_LOST" in line:
                lost_records += 1
                continue
            if stripped.startswith("..."):
                require(current is not None, f"sample truncation marker before sample at line {line_number}")
                current["truncated"] = True
                continue
            header = SAMPLE_HEADER.match(line)
            if header:
                if current is not None:
                    finish(current)
                event = header.group("event").rstrip(":")
                require(event == EVENT, f"unexpected perf event {event} at line {line_number}")
                period = int(header.group("period"))
                require(period > 0, f"nonpositive sample period at line {line_number}")
                all_samples += 1
                all_period += period
                current = {"pid": int(header.group("pid")), "tid": int(header.group("tid")),
                           "time": DECIMAL(header.group("time")), "period": period,
                           "frames": [], "truncated": False}
                continue
            require(current is not None, f"unparsed line outside sample at {line_number}: {stripped[:120]}")
            frame = parse_frame(line)
            require(frame is not None, f"unparsed frame at line {line_number}: {stripped[:120]}")
            current["frames"].append(frame)
    if current is not None:
        finish(current)
    require(all_samples > 0, "corrected sample export contains no samples")
    require(all_period == owner_period + unknown_period, "sample period partition is not disjoint")
    gate = {"lost_records_zero": lost_records == 0,
            "fully_resolved_owner_samples_min": owner_samples >= OWNER_MIN_SAMPLES,
            "owner_period_min": owner_period >= OWNER_MIN_PERIOD}
    return {"all": {"samples": all_samples, "period": all_period},
            "owner_qualified": {"samples": owner_samples, "period": owner_period},
            "unknown_or_unqualified": {"samples": unknown_samples, "period": unknown_period},
            "unresolved_period": unresolved_period, "lost_records": lost_records,
            "frame_length_counts": dict(sorted(frame_lengths.items())),
            "classification_reasons": dict(sorted(reasons.items())),
            "exclusive_leaf_ranking": [{"symbol": name, "period": value}
                                       for name, value in leaves.most_common()],
            "nested_compression_xml_rows": [{"symbol": name, "period": value}
                                             for name, value in nested.most_common()],
            "nested_rows_overlap": True, "gate": gate,
            "admitted": all(gate.values())}


def write_exclusive(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8", newline="") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")
        stream.flush()
        os.fsync(stream.fileno())


def run(args: argparse.Namespace) -> tuple[dict[str, Any], int]:
    bindings, semantic = verify_bindings()
    plan = load_json(PLAN)
    row, child, custody = child_context(plan, args.index)
    sample = export_context(args.sample_export.resolve(), child, custody["receipt"], row, "sample")
    mmap_export = export_context(args.mmap_export.resolve(), child, custody["receipt"], row, "mmap")
    binary = Path(plan["binary"]["path"])
    loads, build_id = read_elf(binary)
    mmap = parse_mmap_events(args.mmap_export / "events.txt", binary, loads, build_id)
    sem_result = semantic.validate_report(child / "report.json", row["shape"], receipt_path=child / "receipt.json")
    profile = parse_samples(args.sample_export / "events.txt", binary, mmap, plan)
    status = "pass" if profile["admitted"] else "fail"
    result = {"schema": "litchi.performance.0952.native-admission.v3", "status": status,
              "plan": bindings["plan"], "runner": bindings["runner"], "binary": bindings["binary"],
              "semantic_adapter": bindings["semantic"], "sample_exporter": bindings["sample_exporter"],
              "mmap_exporter": bindings["mmap_exporter"], "row": row, "child": descriptor(child / "receipt.json"),
              "custody": custody, "sample_export": sample, "mmap_export": mmap_export,
              "elf": {"build_id": build_id, "pt_loads": loads},
              "mmap_binding": {key: value for key, value in mmap.items() if key != "rows" and key != "bound_rows"},
              "semantic": sem_result, "profile": profile,
              "scope": "Descriptive cycles:u owner attribution only; no native timing, speedup, allocation, RSS, or adoption claim."}
    return result, 0 if status == "pass" else 2


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=PACKET)
    parser.add_argument("--index", type=int, required=True)
    parser.add_argument("--sample-export", type=Path, required=True)
    parser.add_argument("--mmap-export", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        result, code = run(args)
        write_exclusive(args.output.resolve(), result)
        print(json.dumps({"status": result["status"], "output": str(args.output.resolve()),
                          "sha256": sha(args.output.resolve())}, sort_keys=True))
        return code
    except (AdmissionError, OSError, ValueError, KeyError, ImportError, struct.error) as exc:
        failure = {"schema": "litchi.performance.0952.native-admission.v3", "status": "error",
                   "error": str(exc), "scope": "No admission; failed closed before native attribution."}
        try:
            write_exclusive(args.output.resolve(), failure)
        except Exception:
            pass
        print(f"analyze_0952_v3: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
