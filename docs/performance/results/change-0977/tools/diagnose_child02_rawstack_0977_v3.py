#!/usr/bin/env python3
"""Bounded offline diagnosis of the retained 0977 child-02 raw-stack fail."""
from __future__ import annotations

import hashlib
import importlib.util
import json
import re
import struct
from pathlib import Path


ROOT = Path("/home/zhuhe/litchi-performance-0977")
CHILD = ROOT / "capture/instruction-profile-v2/child-02"
RAW_DUMP = CHILD / "child-02.perf-mmap-dump.txt"
PERF_DATA = CHILD / "child-02.perf.data"
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
LAYOUT = CHILD / "child-02.target-layout.json"
LAUNCH = CHILD / "child-02.launch.json"
TERMINAL = CHILD / "child-02.terminal.json"
STATUS = CHILD / "child-02.status.json"
EVLIST = CHILD / "child-02.perf-evlist-v.txt"
TERMINAL_REPORT = CHILD / "child-02.terminal-report.txt"
DECODER = ROOT / "tools/focuseddecode_0977_v2.py"
DECODER_SHA256 = "31184f1dc044022f262def181af909bf987c04936e1f926acd73f2c10e46e23c"
V2_RECEIPT = ROOT / "receipts/child02-rawstack-diagnostic-0977-v2.json"
V2_RECEIPT_SHA256 = "53d63ac212d4e4fdc8a082c8303788b860103f20213a7b102b84b52e7eade668"
RAW_SCANNER_SHA256 = "3a8e36ebc5f21f4f0b8293cafa675d722f2264fb1cfd6c7f48271b4f0a5056d2"
SAMPLE_RE = re.compile(
    r"PERF_RECORD_SAMPLE\(IP,\s*0x([0-9a-fA-F]+)\):\s+"
    r"(\d+)/(\d+):\s+0x([0-9a-fA-F]+)\s+period:\s+(\d+)"
)
RAW_RE = re.compile(r"^\.\s+\.\.\. raw event: size (\d+) bytes$")
HEX_RE = re.compile(r"^\.\s+([0-9a-fA-F]+):\s+((?:[0-9a-fA-F]{2}\s+)+)")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, object]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_decoder():
    if sha256(DECODER) != DECODER_SHA256:
        raise RuntimeError("frozen decoder v2 hash changed")
    spec = importlib.util.spec_from_file_location("focuseddecode_0977_child02_diag", DECODER)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot import frozen decoder v2")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def find_short_sample() -> dict[str, object]:
    """Locate every size-72 sample and decode its fixed raw header."""
    pending: dict[str, object] | None = None
    last_short: dict[str, object] | None = None
    sample_index = 0
    matches: list[dict[str, object]] = []
    with RAW_DUMP.open("r", encoding="utf-8", errors="strict") as stream:
        for line_number, line in enumerate(stream, 1):
            line = line.rstrip("\n")
            raw_match = RAW_RE.match(line)
            if raw_match:
                pending = {"raw_line": line_number, "bytes": int(raw_match.group(1)), "payload": bytearray(int(raw_match.group(1)))}
                continue
            hex_match = HEX_RE.match(line)
            if hex_match and pending is not None:
                offset = int(hex_match.group(1), 16)
                payload = bytes.fromhex(hex_match.group(2))
                pending_payload = pending["payload"]
                assert isinstance(pending_payload, bytearray)
                end = min(offset + len(payload), len(pending_payload))
                if offset < len(pending_payload):
                    pending_payload[offset:end] = payload[: end - offset]
                continue
            if last_short is not None:
                stack_match = re.search(r"\.\.\. ustack: size (\d+), offset (0x[0-9a-fA-F]+)", line)
                if stack_match is not None:
                    last_short["perf_text_stack_size"] = int(stack_match.group(1))
                    last_short["perf_text_stack_offset"] = stack_match.group(2).lower()
                    last_short = None
            sample_match = SAMPLE_RE.search(line)
            if sample_match is None:
                continue
            sample_index += 1
            if pending is None:
                continue
            raw_bytes = int(pending["bytes"])
            payload = bytes(pending["payload"])
            if raw_bytes == 72:
                if len(payload) < raw_bytes:
                    payload += b"\0" * (raw_bytes - len(payload))
                record_type, record_misc, record_size = struct.unpack_from("<IHH", payload, 0)
                matches.append({
                    "sample_index": sample_index,
                    "raw_event_line": pending["raw_line"],
                    "sample_line": line_number,
                    "raw_bytes": raw_bytes,
                    "record_type": record_type,
                    "record_misc": record_misc,
                    "record_size": record_size,
                    "event_ip": "0x" + sample_match.group(4).lower(),
                    "event_pid": int(sample_match.group(2)),
                    "event_tid": int(sample_match.group(3)),
                    "event_period": int(sample_match.group(5)),
                    "event_time_ns": int(line.split()[1]),
                    "raw_ip": hex(struct.unpack_from("<Q", payload, 8)[0]),
                    "raw_pid": struct.unpack_from("<I", payload, 16)[0],
                    "raw_tid": struct.unpack_from("<I", payload, 20)[0],
                    "raw_cpu": struct.unpack_from("<I", payload, 32)[0],
                    "raw_period": struct.unpack_from("<Q", payload, 40)[0],
                    "callchain_nr": struct.unpack_from("<Q", payload, 48)[0],
                    "callchain_nr_at_offset_48": struct.unpack_from("<Q", payload, 48)[0],
                    "regs_user_abi_at_offset_56": struct.unpack_from("<Q", payload, 56)[0],
                    "user_stack_size_at_offset_64": struct.unpack_from("<Q", payload, 64)[0],
                    "raw_event_hex": payload.hex(),
                })
                last_short = matches[-1]
            pending = None
    if len(matches) != 1:
        raise RuntimeError(f"expected exactly one 72-byte sample, found {len(matches)}")
    return matches[0]


def main() -> int:
    decoder = load_decoder()
    if sha256(V2_RECEIPT) != V2_RECEIPT_SHA256:
        raise RuntimeError("prior diagnostic v2 receipt changed")
    prior = json.loads(V2_RECEIPT.read_text(encoding="utf-8"))
    raw_summary = prior["raw_scanner_summary"]
    short = find_short_sample()
    terminal_report = TERMINAL_REPORT.read_text(encoding="utf-8", errors="strict")
    lost = re.findall(r"^#\s*Total Lost Samples:\s*(\d+)\s*$", terminal_report, re.MULTILINE)
    status = json.loads(STATUS.read_text(encoding="utf-8"))
    terminal = json.loads(TERMINAL.read_text(encoding="utf-8"))
    evlist_text = EVLIST.read_text(encoding="utf-8", errors="strict").strip()
    receipt = {
        "schema": "litchi.performance.0977.child02_rawstack_diagnostic.v1",
        "status": "decoder_rejection_confirmed",
        "scope": "offline raw scanner diagnosis only; no native rerun, workload, build, or decoder mutation",
        "decoder": binding(DECODER),
        "prior_diagnostic": binding(V2_RECEIPT),
        "sealed_raw_scanner_sha256": RAW_SCANNER_SHA256,
        "capture": {
            "perf_data": binding(PERF_DATA),
            "perf_mmap_dump": binding(RAW_DUMP),
            "target_layout": binding(LAYOUT),
            "binary": binding(BINARY),
            "launch": binding(LAUNCH),
            "terminal": binding(TERMINAL),
            "status": binding(STATUS),
            "evlist": binding(EVLIST),
            "terminal_report": binding(TERMINAL_REPORT),
        },
        "native_terminal": {
            "returncode": terminal.get("returncode"),
            "timed_out": terminal.get("timed_out"),
            "launch_error": terminal.get("launch_error"),
            "process_group_cleanup_passed": terminal.get("process_group_cleanup", {}).get("passed"),
            "terminal_lost_samples_lines": lost,
            "decoder_status": status,
        },
        "raw_scanner_summary": {
            "status": raw_summary.get("status"),
            "sample_count": raw_summary.get("sample_count"),
            "sample_line_count": raw_summary.get("sample_line_count"),
            "raw_event_count": raw_summary.get("raw_event_count"),
            "unconsumed_raw_count": raw_summary.get("unconsumed_raw_count"),
            "period_valid": raw_summary.get("period_valid"),
            "pid_tid_cpu_valid": raw_summary.get("pid_tid_cpu_valid"),
            "bounds_valid": raw_summary.get("bounds_valid"),
            "missing": raw_summary.get("missing"),
            "violations": raw_summary.get("violations"),
            "record_type_counts": raw_summary.get("record_type_counts"),
            "record_size_counts": raw_summary.get("record_size_counts"),
            "stack_size_counts": raw_summary.get("stack_size_counts"),
            "pf_x": raw_summary.get("pf_x"),
        },
        "evlist_line": evlist_text,
        "offending_sample": short,
        "scanner_gate_explanation": {
            "source_constants": {"MIN_HEADER": 240, "STACK_SIZE_OFFSET": 224, "STACK_BYTES_OFFSET": 232},
            "observed": "one PERF_RECORD_SAMPLE has a valid 72-byte record; callchain_nr at offset 48 is 0, regs-user ABI at offset 56 is 0, and stack size at offset 64 is 0",
            "perf_text_evidence": "the retained dump prints ustack: size 0, offset 0x40 for this exact record",
            "mode_evidence": "record misc low bits are 0x1 (PERF_RECORD_MISC_KERNEL per linux/perf_event.h); sample IP is kernel-address-shaped and no user regs/stack payload is present",
            "classification": "unsupported variable sample shape for this fixed-offset STACK_USER scanner; no stack_bound violation was observed",
            "admission_effect": "frozen decoder correctly rejects the child and does not produce corrected attribution",
        },
    }
    out = ROOT / "receipts/child02-rawstack-diagnostic-0977-v3.json"
    with out.open("x", encoding="utf-8") as handle:
        json.dump(receipt, handle, indent=2, sort_keys=True, allow_nan=False)
        handle.write("\n")
    print(json.dumps({"receipt": str(out), "receipt_sha256": sha256(out), "status": receipt["status"], "offending_sample": short, "violations": raw_summary.get("violations")}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
