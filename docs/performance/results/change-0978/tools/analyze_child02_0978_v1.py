#!/usr/bin/env python3
"""Offline forensic ledger for retained 0977 child-02.

This is an external analysis helper. It never records or launches a workload;
it parses one already-produced corrected perf-script stream and its raw
perf-script -D ledger, retaining all sample periods and unknown/empty mass.
"""
from __future__ import annotations

import bisect
import collections
import datetime as dt
import hashlib
import importlib.util
import json
import re
from pathlib import Path
from typing import Any

OUT = Path("/home/zhuhe/litchi-performance-0978/cpu-profile-child02-v1")
CAP = OUT / "capture/child-02"
SCRIPT = CAP / "child-02.corrected.perf-script-no-inline.txt"
RAW = CAP / "child-02.perf-mmap-dump.txt"
SHIM_LOG = CAP / "child-02.corrected.dwfl.jsonl"
TERMINAL_REPORT = CAP / "child-02.terminal-report.txt"
PLAN = Path("/home/zhuhe/litchi-performance-0977/plans/plan-0977-v2.json")
DECODER = Path("/home/zhuhe/litchi-performance-0977/tools/focuseddecode_0977_v2.py")
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
RESULT = OUT / "analysis/child-02-forensic-summary-v1.json"

HEADER_RE = re.compile(
    r"^\s*(?:(?P<comm>\S+)\s+)?(?P<pid>\d+)(?:/(?P<tid>\d+))?\s+"
    r"\[(?P<cpu>\d+)\]\s+(?P<time>\d+(?:\.\d+)?):\s+(?P<period>\d+)\s+"
    r"(?P<event>[^\n]+):\s*$"
)
FRAME_RE = re.compile(r"^\s*(?P<ip>[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+\((?P<dso>.*)\)\s*$")
RAW_SAMPLE_RE = re.compile(
    r"^\s*\d+\s+(?P<ns>\d+)\s+(?P<offset>\S+)\s+\[(?P<size>\S+)\]:\s+"
    r"PERF_RECORD_SAMPLE\((?P<mode>[^)]*)\):\s+(?P<pid>\d+)/(?P<tid>\d+):\s+"
    r"(?P<ip>\S+)\s+period:\s+(?P<period>\d+)"
)
COMM_RE = re.compile(
    r"^\s*(?P<cpu>\d+)\s+(?P<ns>\d+)\s+\S+\s+\[\S+\]:\s+"
    r"PERF_RECORD_COMM(?: exec)?:\s*(?P<comm>[^:]+):(?P<pid>\d+)/(?P<tid>\d+)"
)
LOST_RE = re.compile(r"PERF_RECORD_LOST(?:\s|:).*?\b(?:lost|events?)\s*[:=]?\s*(\d+)", re.I)
TERMINAL_RE = re.compile(r"^#\s*Total Lost Samples:\s*(\d+)\s*$")


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def bind(path: Path) -> dict[str, Any]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def load_decoder() -> Any:
    spec = importlib.util.spec_from_file_location("focuseddecode_0977_v2_forensic", DECODER)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot import decoder: {DECODER}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def parse_script() -> list[dict[str, Any]]:
    samples: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None

    def finish() -> None:
        nonlocal current
        if current is not None:
            samples.append(current)
            current = None

    for line_number, raw in enumerate(SCRIPT.read_text(encoding="utf-8", errors="strict").splitlines(), 1):
        line = raw.rstrip("\n")
        header = HEADER_RE.match(line)
        if header is not None:
            finish()
            current = {
                "line": line_number,
                "time": float(header.group("time")),
                "period": int(header.group("period")),
                "event": header.group("event").strip(),
                "pid": int(header.group("pid")),
                "tid": int(header.group("tid") or header.group("pid")),
                "cpu": int(header.group("cpu")),
                "frames": [],
            }
            continue
        if not line.strip():
            continue
        if current is None:
            if not line.startswith("#"):
                raise RuntimeError(f"non-sample line {line_number}: {line[:200]}")
            continue
        frame = FRAME_RE.match(line)
        if frame is None:
            if line.startswith("#"):
                continue
            raise RuntimeError(f"malformed frame line {line_number}: {line[:300]}")
        symbol = frame.group("symbol").strip()
        base, _, offset = symbol.partition("+")
        dso = frame.group("dso").strip()
        current["frames"].append({
            "ip": frame.group("ip").lower(),
            "symbol": symbol,
            "symbol_base": base.strip(),
            "offset": offset or None,
            "dso": dso,
            "resolved": base.strip() not in {"[unknown]", "unknown", ""} and dso not in {"[unknown]", ""},
        })
    finish()
    return samples


def parse_raw() -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    samples: list[dict[str, Any]] = []
    comms: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None

    def finish() -> None:
        nonlocal current
        if current is not None:
            samples.append(current)
            current = None

    for line_number, raw in enumerate(RAW.read_text(encoding="utf-8", errors="strict").splitlines(), 1):
        line = raw.rstrip("\n")
        cm = COMM_RE.match(line)
        if cm is not None:
            row = cm.groupdict()
            row.update({"line": line_number, "ns": int(row["ns"]), "pid": int(row["pid"]), "tid": int(row["tid"]), "raw": line})
            comms.append(row)
        rm = RAW_SAMPLE_RE.match(line)
        if rm is not None:
            finish()
            row = rm.groupdict()
            row.update({
                "line": line_number,
                "ns": int(row["ns"]),
                "record_size": int(row["size"].strip("[]"), 16),
                "pid": int(row["pid"]),
                "tid": int(row["tid"]),
                "period": int(row["period"]),
                "raw": line,
                "fp_chain_nr": None,
                "ustack_size": None,
                "thread": None,
                "dso": None,
            })
            current = row
            continue
        if current is None:
            continue
        fp = re.search(r"\.\.\. FP chain: nr:(\d+)", line)
        if fp is not None:
            current["fp_chain_nr"] = int(fp.group(1))
        us = re.search(r"\.\.\. ustack: size (\d+),", line)
        if us is not None:
            current["ustack_size"] = int(us.group(1))
        th = re.search(r"\.\.\.\.\. thread: ([^:]+):(\d+)", line)
        if th is not None:
            current["thread"] = th.group(1)
        ds = re.search(r"\.\.\.\.\.\. dso: (.*)$", line)
        if ds is not None:
            current["dso"] = ds.group(1)
    finish()
    lost_values = [int(m.group(1)) for line in RAW.read_text(encoding="utf-8", errors="strict").splitlines() if (m := LOST_RE.search(line))]
    return samples, comms, [{"raw_lost_values": lost_values, "raw_lost_sum": sum(lost_values)}]


def comm_at(ns: int, comms: list[dict[str, Any]]) -> str:
    ordered = [row["ns"] for row in comms]
    index = bisect.bisect_right(ordered, ns) - 1
    return comms[index]["comm"] if index >= 0 else "before_first_COMM"


def summarize_samples(script: list[dict[str, Any]], raw: list[dict[str, Any]], comms: list[dict[str, Any]], decoder: Any, plan: dict[str, Any]) -> dict[str, Any]:
    if len(script) != len(raw):
        raise RuntimeError(f"raw/script sample count mismatch: {len(raw)} vs {len(script)}")
    marker = str(plan["marker"]["symbol"])
    marker_dso = str(plan["marker"]["dso_path"])
    expected_pid = int(json.loads((CAP / "child-02.launch.json").read_text())["exec_pid"])
    parsed = decoder._stream_perf_script(SCRIPT, marker=marker, marker_dso=marker_dso, expected_pid=expected_pid, expected_cpu=31, expected_period=1_000_000)
    combined: list[dict[str, Any]] = []
    mode_counts: dict[str, dict[str, Any]] = {}
    comm_counts: dict[str, dict[str, Any]] = {}
    comm_mode_counts: dict[str, dict[str, Any]] = {}
    empty_mode = collections.Counter()
    empty_size = collections.Counter()
    empty_comm = collections.Counter()
    all_unknown = collections.Counter()
    for index, (s, r) in enumerate(zip(script, raw)):
        if s["period"] != r["period"]:
            raise RuntimeError(f"period mismatch at sample {index}")
        comm = comm_at(r["ns"], comms)
        resolved = [f for f in s["frames"] if f["resolved"]]
        unresolved_frame_count = len(s["frames"]) - len(resolved)
        is_marker = any(f["symbol_base"] == marker and f["dso"] == marker_dso for f in s["frames"])
        row = {
            "index": index,
            "time": s["time"],
            "raw_time_ns": r["ns"],
            "period": s["period"],
            "pid": s["pid"],
            "tid": s["tid"],
            "cpu": s["cpu"],
            "raw_mode": r["mode"],
            "raw_ip": r["ip"],
            "raw_record_size": r["record_size"],
            "raw_fp_chain_nr": r["fp_chain_nr"],
            "raw_ustack_size": r["ustack_size"],
            "raw_thread": r["thread"],
            "raw_dso": r["dso"],
            "comm_exec_image": comm,
            "frame_count": len(s["frames"]),
            "resolved_frame_count": len(resolved),
            "unresolved_frame_count": unresolved_frame_count,
            "empty_callchain": not s["frames"],
            "no_resolved_frame": not resolved,
            "marker": is_marker,
        }
        combined.append(row)
        mode = mode_counts.setdefault(r["mode"], {"samples": 0, "period": 0, "empty": 0, "no_resolved": 0, "kernel_ip_like": 0})
        mode["samples"] += 1; mode["period"] += s["period"]; mode["empty"] += int(not s["frames"]); mode["no_resolved"] += int(not resolved); mode["kernel_ip_like"] += int(r["ip"].startswith("0xffffffff"))
        commrow = comm_counts.setdefault(comm, {"samples": 0, "period": 0, "empty": 0, "no_resolved": 0, "marker": 0})
        commrow["samples"] += 1; commrow["period"] += s["period"]; commrow["empty"] += int(not s["frames"]); commrow["no_resolved"] += int(not resolved); commrow["marker"] += int(is_marker)
        key = f"{comm}|{r['mode']}"
        cmrow = comm_mode_counts.setdefault(key, {"comm_exec_image": comm, "raw_mode": r["mode"], "samples": 0, "period": 0, "empty": 0, "marker": 0})
        cmrow["samples"] += 1; cmrow["period"] += s["period"]; cmrow["empty"] += int(not s["frames"]); cmrow["marker"] += int(is_marker)
        if not s["frames"]:
            empty_mode[r["mode"]] += 1; empty_size[r["record_size"]] += 1; empty_comm[comm] += 1
        all_unknown["no_resolved_samples"] += int(not resolved)
        all_unknown["no_resolved_period"] += s["period"] if not resolved else 0
        all_unknown["unresolved_frame_count"] += unresolved_frame_count
        all_unknown["empty_samples"] += int(not s["frames"])
        all_unknown["empty_period"] += s["period"] if not s["frames"] else 0
    if sum(x["period"] for x in combined) != parsed["total_period"]:
        raise RuntimeError("combined period does not match decoder parser")
    if len(combined) != parsed["sample_count"] or sum(x["marker"] for x in combined) != parsed["marker_sample_count"]:
        raise RuntimeError("combined sample/marker counts do not match decoder parser")
    terminal_text = TERMINAL_REPORT.read_text(encoding="utf-8", errors="strict")
    terminal_values = [int(m.group(1)) for line in terminal_text.splitlines() if (m := TERMINAL_RE.match(line))]
    if terminal_values != [0]:
        raise RuntimeError(f"terminal loss mismatch: {terminal_values}")
    raw_lost_values = [int(m.group(1)) for line in RAW.read_text(encoding="utf-8", errors="strict").splitlines() if (m := LOST_RE.search(line))]
    return {
        "decoder_parser_summary": parsed,
        "sample_count": len(combined),
        "total_period": sum(x["period"] for x in combined),
        "period_conservation": {"expected_samples": 7201, "expected_period": 7_201_000_000, "observed_samples": len(combined), "observed_period": sum(x["period"] for x in combined), "passed": len(combined) == 7201 and sum(x["period"] for x in combined) == 7_201_000_000},
        "sample_window": {"first_time": combined[0]["time"], "last_time": combined[-1]["time"], "duration_seconds": combined[-1]["time"] - combined[0]["time"]},
        "pid_tid_cpu": {"pids": sorted({x["pid"] for x in combined}), "tids": sorted({x["tid"] for x in combined}), "cpus": sorted({x["cpu"] for x in combined})},
        "periods": {"values": sorted({x["period"] for x in combined}), "counts": dict(collections.Counter(x["period"] for x in combined))},
        "raw_sample_modes": mode_counts,
        "raw_record_sizes": {str(k): v for k, v in collections.Counter(x["record_size"] for x in raw).items()},
        "raw_kernel_ip_like": sum(x["ip"].startswith("0xffffffff") for x in raw),
        "empty_callchain": {"samples": sum(x["empty_callchain"] for x in combined), "period": sum(x["period"] for x in combined if x["empty_callchain"]), "by_raw_mode": dict(empty_mode), "by_record_size": {str(k): v for k, v in empty_size.items()}, "by_comm_exec_image": dict(empty_comm), "first_time": min(x["time"] for x in combined if x["empty_callchain"]), "last_time": max(x["time"] for x in combined if x["empty_callchain"])},
        "unknown_mass_all_samples": dict(all_unknown),
        "comm_exec_timeline": [{k: row[k] for k in ("ns", "comm", "pid", "tid", "raw")} for row in comms],
        "sample_by_comm_exec_image": comm_counts,
        "sample_by_comm_and_raw_mode": comm_mode_counts,
        "marker": {"symbol": marker, "dso": marker_dso, "samples": parsed["marker_sample_count"], "period": parsed["marker_period"], "share_of_all_period": parsed["marker_period"] / parsed["total_period"], "first_time": min(x["time"] for x in combined if x["marker"]), "last_time": max(x["time"] for x in combined if x["marker"])},
        "loss": {"raw_lost_values": raw_lost_values, "raw_lost_sum": sum(raw_lost_values), "terminal_lost_values": terminal_values, "terminal_lost_samples": terminal_values[0]},
        "shim_log": decoder._stream_shim_log(SHIM_LOG, json.loads((CAP / "child-02.target-layout.json").read_text()), BINARY, expected_build_id="e5c457384eeebd057932c444efda77618f137399"),
        "forensic_sample_ledger": combined,
    }


def main() -> None:
    plan = json.loads(PLAN.read_text(encoding="utf-8"))
    decoder = load_decoder()
    raw, comms, raw_meta = parse_raw()
    script = parse_script()
    result = summarize_samples(script, raw, comms, decoder, plan)
    result["schema"] = "litchi.performance.0978.child02_forensic_summary.v1"
    result["status"] = "pass"
    result["created_utc"] = dt.datetime.now(dt.timezone.utc).isoformat()
    result["scope"] = "one retained child-02 offline forensic decode; no 0977 cohort admission, no native workload, no causal or speedup claim"
    result["raw_meta"] = raw_meta
    result["bindings"] = {"script": bind(SCRIPT), "raw_mmap": bind(RAW), "shim_log": bind(SHIM_LOG), "terminal_report": bind(TERMINAL_REPORT), "target_layout": bind(CAP / "child-02.target-layout.json"), "perf_data": bind(CAP / "child-02.perf.data"), "plan": bind(PLAN), "decoder": bind(DECODER), "binary": bind(BINARY)}
    result["interpretation"] = {
        "focused_parser_boundary": "The retained raw STACK_USER scanner can reject a child when sample-line accounting requires one decoded frame-bearing line per sample. Direct corrected perf-script output conserves all 7201 samples and exposes 76 empty callchains, so the focused decoder rejection is a parser/admission boundary rather than evidence of capture loss.",
        "72_byte_empty_kernel_sample": "One 0x48 (72-byte) PERF_RECORD_SAMPLE has a kernel-like IP 0xffffffff920000a7, raw mode IP, 0x1, and no user-stack/callchain output. It occurs before the doc_retention_p COMM exec image.",
        "remaining_empty_mass": "The other 75 empty callchains are 0x80f0 (33008-byte) sample records; 74 are raw mode IP, 0x2 and one is IP, 0x1. Their retained user-stack payloads do not yield decoded callchain frames in perf script. Keep their full period as unknown/empty mass.",
        "source_attribution_limit": "Only marker-containing samples support marker-owner views. Keep D=all period, M=marker period, R=D-M labeled marker ancestry not observed; do not assign R or empty/non-resolved periods to a source owner. For a future admitted 12-child cohort, retain physical frame-0 owner plus explicit unknown, deepest-resolved proxy, and inclusive owners with separate denominators; require all-child conservation and never infer causality or speedup.",
        "cohort_admission": "This child is diagnostic only; the 0977 all-12 cohort remains failed at child 2 and is not admitted.",
    }
    RESULT.parent.mkdir(parents=True, exist_ok=True)
    if RESULT.exists():
        raise RuntimeError(f"refusing overwrite {RESULT}")
    RESULT.write_text(json.dumps(result, indent=2, sort_keys=True, allow_nan=False) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(RESULT), "bytes": RESULT.stat().st_size, "sha256": sha(RESULT), "samples": result["sample_count"], "period": result["total_period"], "empty": result["empty_callchain"]["samples"], "marker": result["marker"]["samples"]}, sort_keys=True))

if __name__ == "__main__":
    main()
