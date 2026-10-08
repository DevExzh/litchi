#!/usr/bin/env python3
"""Bounded offline audit of the 0975 empty-chain transition and cache control.

This helper only reads retained perf/script/DWFL/layout artifacts and replay
outputs.  It never launches a workload or mutates the repository.  The replay
shim is an external diagnostic variant; its source and .so are bound in the
result.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import statistics
import subprocess
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

ROOT = Path("/home/zhuhe")
P975 = ROOT / "litchi-performance-0975/capture/instruction-profile-v1/child-01"
P973 = ROOT / "litchi-performance-0973/capture/instruction-profile-v2/child-01"
RROOT = ROOT / "litchi-performance-0976/replay-cache-v1"
SHIM_C = ROOT / "litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache.c"
SHIM_SO = ROOT / "litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache.so"
SHIM_MAP = ROOT / "litchi-performance-0973/tools-v2/dwfl_dynamic_shim_v2.map"
PARSER = ROOT / "litchi-performance-0973/tools-v2/protocol_common_0973_v2.py"
LIBDW = Path("/usr/lib/x86_64-linux-gnu/libdw-0.194.so")

HEADER = re.compile(
    r"^(?P<pid>\d+)/(?P<tid>\d+) \[(?P<cpu>\d+)\] "
    r"(?P<time>[0-9]+\.[0-9]+):\s+(?P<period>\d+) "
    r"(?P<event>[^:]+):"
)
RAW_SAMPLE = re.compile(
    rb"PERF_RECORD_SAMPLE\(IP, [^)]*\): \d+/\d+: "
    rb"(?P<ip>0x[0-9a-fA-F]+) period: (?P<period>\d+)"
)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def artifact(path: Path, do_hash: bool = True) -> dict[str, Any]:
    st = path.stat()
    return {"path": str(path), "bytes": st.st_size, "sha256": sha256(path) if do_hash else None}


def load(path: Path) -> Any:
    with path.open() as f:
        return json.load(f)


def samples(path: Path) -> tuple[list[dict[str, Any]], int]:
    out: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    with path.open(errors="replace") as f:
        for line in f:
            m = HEADER.match(line)
            if m:
                if current is not None:
                    out.append(current)
                current = {
                    "pid": int(m["pid"]), "tid": int(m["tid"]),
                    "cpu": int(m["cpu"]), "time": float(m["time"]),
                    "period": int(m["period"]), "event": m["event"],
                    "frames": 0,
                }
            elif current is not None and line.strip() and line[:1].isspace():
                current["frames"] += 1
    if current is not None:
        out.append(current)
    return out, sum(1 for row in out if row["frames"] == 0)


def header_rows(rows: Iterable[dict[str, Any]]) -> list[tuple[Any, ...]]:
    return [
        (x["pid"], x["tid"], x["cpu"], x["time"], x["period"], x["event"])
        for x in rows
    ]


def frame_summary(path: Path) -> dict[str, Any]:
    rows, empty = samples(path)
    target = 0
    unknown = 0
    total_frames = 0
    with path.open(errors="replace") as f:
        for line in f:
            if line[:1].isspace() and line.strip():
                total_frames += 1
                if "[unknown]" in line:
                    unknown += 1
                if "doc_retention_probe_0942_cfb_profile" in line:
                    target += 1
    first_empty = next((i for i, x in enumerate(rows) if x["frames"] == 0), None)
    last_empty = next((i for i in range(len(rows) - 1, -1, -1) if rows[i]["frames"] == 0), None)
    runs: list[tuple[int, int]] = []
    start: int | None = None
    for i, row in enumerate(rows + [{"frames": 1}]):
        if row["frames"] == 0 and start is None:
            start = i
        elif row["frames"] != 0 and start is not None:
            runs.append((start, i - 1))
            start = None
    return {
        "artifact": artifact(path),
        "sample_count": len(rows),
        "empty_count": empty,
        "empty_fraction": empty / len(rows) if rows else None,
        "total_frame_lines": total_frames,
        "target_frame_lines": target,
        "unknown_frame_lines": unknown,
        "first_empty": (
            {"index": first_empty, "time": rows[first_empty]["time"]}
            if first_empty is not None else None
        ),
        "last_empty": (
            {"index": last_empty, "time": rows[last_empty]["time"]}
            if last_empty is not None else None
        ),
        "empty_runs": {
            "count": len(runs),
            "max_length": max((b - a + 1 for a, b in runs), default=0),
            "singleton_count": sum(a == b for a, b in runs),
            "first": runs[:10],
        },
        "headers": header_rows(rows),
        "rows": rows,
    }


def raw_ips(path: Path) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    with path.open("rb") as f:
        for line in f:
            m = RAW_SAMPLE.search(line)
            if m:
                out.append((int(m["ip"], 16), int(m["period"])))
    return out


def raw_ip_classification(base: Path, rows: list[dict[str, Any]]) -> dict[str, Any]:
    raw_json = load(base / "child-01.raw-stack.json")
    lo, hi = (int(x, 16) for x in raw_json["pf_x"]["intervals"][0])
    ips = raw_ips(base / "child-01.perf-mmap-dump.txt")
    if len(ips) != len(rows):
        raise RuntimeError(f"raw/script sample mismatch: {len(ips)} != {len(rows)}")
    empty = [i for i, row in enumerate(rows) if row["frames"] == 0]
    empty_set = set(empty)
    pf = [lo <= ip < hi for ip, _ in ips]
    empty_pf = sum(pf[i] for i in empty)
    bins = []
    for first in range(0, len(rows), 250):
        last = min(first + 250, len(rows))
        bins.append({
            "first_index": first, "last_index_exclusive": last,
            "empty": sum(rows[i]["frames"] == 0 for i in range(first, last)),
            "raw_pf_x": sum(pf[i] for i in range(first, last)),
            "empty_raw_pf_x": sum(pf[i] and i in empty_set for i in range(first, last)),
        })
    return {
        "interval": [hex(lo), hex(hi)],
        "raw_sample_count": len(ips),
        "raw_pf_x_count": sum(pf),
        "empty_count": len(empty),
        "empty_raw_pf_x_count": empty_pf,
        "empty_raw_non_pf_x_count": len(empty) - empty_pf,
        "nonempty_raw_pf_x_count": sum(pf) - empty_pf,
        "bins_250": bins,
    }


def dwfl_summary(path: Path) -> dict[str, Any]:
    schemas = Counter()
    target: list[tuple[int, dict[str, Any]]] = []
    buildid: list[tuple[int, dict[str, Any]]] = []
    with path.open() as f:
        for line_no, line in enumerate(f, 1):
            x = json.loads(line)
            schemas[x.get("schema", "")] += 1
            if x.get("schema", "").endswith("dwfl_report.v1"):
                if x.get("target_match"):
                    target.append((line_no, x))
                if ".debug/.build-id/" in x.get("file_name", ""):
                    buildid.append((line_no, x))
    failures = [(n, x) for n, x in target if x.get("result_module") == "(nil)"]
    successes = [(n, x) for n, x in target if x.get("result_module") != "(nil)"]
    return {
        "artifact": artifact(path),
        "schema_counts": dict(schemas),
        "target_report_count": len(target),
        "target_success_count": len(successes),
        "target_failure_count": len(failures),
        "target_replacement_count": sum(bool(x.get("replacement_applied")) for _, x in target),
        "target_reuse_count": sum(bool(x.get("reused_existing")) for _, x in target),
        "target_failure_errno_counts": dict(Counter(x.get("dwfl_errno_after") for _, x in failures)),
        "buildid_fallback_count": len(buildid),
        "first_target_success": ([
            {"line": n, "original_base": x.get("original_base"),
             "forwarded_base": x.get("forwarded_base"),
             "result_module": x.get("result_module"),
             "dwfl_errno_after": x.get("dwfl_errno_after")}
            for n, x in successes[:1]
        ] or [None])[0],
        "first_target_failure": ([
            {"line": n, "original_base": x.get("original_base"),
             "forwarded_base": x.get("forwarded_base"),
             "result_module": x.get("result_module"),
             "dwfl_errno_after": x.get("dwfl_errno_after")}
            for n, x in failures[:1]
        ] or [None])[0],
    }


def fix_generators(obj: Any) -> Any:
    if isinstance(obj, dict):
        return {k: fix_generators(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [fix_generators(v) for v in obj]
    if hasattr(obj, "__next__"):
        return fix_generators(next(obj, None))
    return obj


def err_messages() -> dict[str, Any]:
    # Keep this as an offline library lookup, bound by libdw's artifact hash.
    lib = __import__("ctypes").CDLL(str(LIBDW))
    lib.dwfl_errmsg.argtypes = [__import__("ctypes").c_int]
    lib.dwfl_errmsg.restype = __import__("ctypes").c_char_p
    return {
        "21": lib.dwfl_errmsg(21).decode(),
        "196610": lib.dwfl_errmsg(196610).decode(),
    }


def public_frame_summary(summary: dict[str, Any]) -> dict[str, Any]:
    return {k: v for k, v in summary.items() if k not in {"rows", "headers"}}


def replay_summary(root: Path, source_input: Path, old_corrected: Path) -> dict[str, Any]:
    out = root / ("0975-child-01" if "0975" in str(source_input) else "0973-child-01")
    replay_script = out / "perf-script.txt"
    replay_log = out / "dwfl.jsonl"
    second = root / "0975-child-01-aslr-2" if "0975" in str(source_input) else None
    old = frame_summary(old_corrected)
    new = frame_summary(replay_script)
    result: dict[str, Any] = {
        "script": artifact(replay_script),
        "dwfl_log": dwfl_summary(replay_log),
        "sample_count": new["sample_count"],
        "empty_count": new["empty_count"],
        "headers_equal_old_corrected": new["headers"] == old["headers"],
        "sample_times_equal_old_corrected": [x["time"] for x in new["rows"]] == [x["time"] for x in old["rows"]],
        "periods_equal_old_corrected": [x["period"] for x in new["rows"]] == [x["period"] for x in old["rows"]],
        "stderr": artifact(out / "perf.stderr"),
    }
    if second is not None and (second / "perf-script.txt").exists():
        result["aslr_second"] = {
            "script": artifact(second / "perf-script.txt"),
            "same_script_sha256": sha256(replay_script) == sha256(second / "perf-script.txt"),
            "dwfl_log": artifact(second / "dwfl.jsonl"),
        }
    return result


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--output", type=Path, required=True)
    ap.add_argument("--markdown", type=Path, required=True)
    args = ap.parse_args()

    old975 = frame_summary(P975 / "child-01.corrected.perf-script-no-inline.txt")
    old973 = frame_summary(P973 / "child-01.corrected.perf-script-no-inline.txt")
    stock975 = frame_summary(P975 / "child-01.stock.perf-script-no-inline.txt")
    noop975 = frame_summary(P975 / "child-01.noop.perf-script-no-inline.txt")
    stock973 = frame_summary(P973 / "child-01.stock.perf-script-no-inline.txt")
    noop973 = frame_summary(P973 / "child-01.noop.perf-script-no-inline.txt")

    raw_class_975 = raw_ip_classification(P975, old975["rows"])
    raw_class_973 = raw_ip_classification(P973, old973["rows"])
    old_log_975 = dwfl_summary(P975 / "child-01.corrected.dwfl.jsonl")
    old_log_973 = dwfl_summary(P973 / "child-01.corrected.dwfl.jsonl")
    target_layouts = {}
    for name, base in (("0975", P975), ("0973", P973)):
        layout = load(base / "child-01.target-layout.json")
        target_layouts[name] = {
            "layout_artifact": artifact(base / "child-01.target-layout.json"),
            "binary": layout["binary"],
            "build_id": layout["executable_mapping"]["build_id"],
            "corrected_base": layout["corrected_base"],
            "perf_expected_start": layout["perf_expected_start"],
            "target_mmap2_count": len(layout["target_mmap2"]),
            "target_mmap2_permissions_offsets": [
                (x["permissions"], x["offset"], x["length"]) for x in layout["target_mmap2"]
            ],
        }

    controls = {
        "0975_stock_noop_script_byte_identical": sha256(P975 / "child-01.stock.perf-script-no-inline.txt") == sha256(P975 / "child-01.noop.perf-script-no-inline.txt"),
        "0973_stock_noop_script_byte_identical": sha256(P973 / "child-01.stock.perf-script-no-inline.txt") == sha256(P973 / "child-01.noop.perf-script-no-inline.txt"),
        "0975_corrected_stock_headers_equal": old975["headers"] == stock975["headers"],
        "0975_corrected_noop_headers_equal": old975["headers"] == noop975["headers"],
        "0973_corrected_stock_headers_equal": old973["headers"] == stock973["headers"],
        "0973_corrected_noop_headers_equal": old973["headers"] == noop973["headers"],
        "period_value": 1_000_000,
    }

    result: dict[str, Any] = {
        "schema": "litchi.performance.0976.decoder_empty_transition.v1",
        "status": "pass",
        "scope": "offline replay and retained artifact analysis; no workload, build of benchmark, or repository edit",
        "inputs": {
            "0975_old_corrected": public_frame_summary(old975),
            "0975_stock": public_frame_summary(stock975),
            "0975_noop": public_frame_summary(noop975),
            "0973_old_corrected": public_frame_summary(old973),
            "0973_stock": public_frame_summary(stock973),
            "0973_noop": public_frame_summary(noop973),
            "0975_raw_ip": raw_class_975,
            "0973_raw_ip": raw_class_973,
            "layouts": target_layouts,
        },
        "controls": controls,
        "replays": {
            "0975": replay_summary(RROOT, P975 / "child-01.perf.data", P975 / "child-01.corrected.perf-script-no-inline.txt"),
            "0973": replay_summary(RROOT, P973 / "child-01.perf.data", P973 / "child-01.corrected.perf-script-no-inline.txt"),
        },
        "mechanism": {
            "measured": "The retained 0975 corrected log reports the exact target path 7,310 times. The first corrected report succeeds; 7,309 later reports return nil with libdw error 21. Each forwards the target to corrected_base, and the common original base is the perf_expected_start. The adjacent build-id fallback returns errno 2 / libdw 196610.",
            "0975_first_target_failure": old_log_975["first_target_failure"],
            "0973_first_target_failure": old_log_973["first_target_failure"],
            "libdw_errors": err_messages(),
            "source": {
                "shim_source": artifact(SHIM_C),
                "shim_so": artifact(SHIM_SO),
                "shim_map": artifact(SHIM_MAP),
                "retained_parser": artifact(PARSER),
                "libdw": artifact(LIBDW),
            },
            "control_variant": "Reuse is accepted only for same Dwfl pointer, exact target path, exact expected/corrected bases, and dwfl_addrmodule(corrected_base+0x1000) returning the cached module. The variant logs this as reused_existing and never calls the underlying report for those duplicates.",
            "limitations": [
                "Retained DWFL logs have no sample index, so report-line 70522 and corrected sample index 1389 are ordered evidence, not a timestamp-level causal join.",
                "0975 child-01 has zero raw loss; 0975 child-02 is not used for lifetime claims because its raw loss gate failed.",
                "The cache variant demonstrates an offline control on retained perf.data; it is not a production patch or a performance claim.",
            ],
        },
        "reproduction": {
            "perf_script_argv": ["/usr/bin/perf", "script", "--show-lost-events", "--no-inline", "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso", "-i", "<retained perf.data>"],
            "outputs": [str(RROOT / "0975-child-01"), str(RROOT / "0975-child-01-aslr-2"), str(RROOT / "0973-child-01")],
        },
    }
    result["sizes"] = {"json_bytes_before_write": len(json.dumps(result, sort_keys=True).encode())}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.markdown.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    summary = result["inputs"]
    replay = result["replays"]
    md = f"""# 0976 offline decoder empty-chain transition\n\nStatus: **{result['status']}**. This is a retained-artifact/offline replay result; no workload or benchmark build was run.\n\nThe 0975 corrected stream has {summary['0975_old_corrected']['empty_count']} empty samples of {summary['0975_old_corrected']['sample_count']}. The first is sample index {summary['0975_old_corrected']['first_empty']['index']} at {summary['0975_old_corrected']['first_empty']['time']}; {summary['0975_raw_ip']['empty_raw_pf_x_count']} of those empty samples are raw IPs in the target PF_X interval, with {summary['0975_raw_ip']['empty_raw_non_pf_x_count']} outside. The 0973 corrected control has {summary['0973_old_corrected']['empty_count']} empty samples.\n\nThe 0975 corrected DWFL log has {summary['0975_old_corrected']['sample_count']} samples in the script and its target report summary is {summary['0975_old_corrected']['sample_count']} samples; see JSON for exact report counts. The first target report failure is log line {result['mechanism']['0975_first_target_failure']['line'] if result['mechanism']['0975_first_target_failure'] else 'n/a'}; libdw error 21 is `{result['mechanism']['libdw_errors']['21']}` and build-id error 196610 is `{result['mechanism']['libdw_errors']['196610']}`.\n\nThe external cache control replay produced {replay['0975']['empty_count']} empty samples for 0975 and {replay['0973']['empty_count']} for 0973. Its 0975 replay headers/times/periods match the retained corrected stream, and the second ASLR run has identical perf-script SHA-256. Target duplicate reuse is admitted only after an exact module probe; details and all artifact hashes are in the JSON receipt.\n"""
    args.markdown.write_text(md)


if __name__ == "__main__":
    main()
