#!/usr/bin/env python3
"""Classify direct sampled leaf symbols from a perf script leaf-period export.

This is a descriptive whole-process self-IP grouping. It does not attribute
samples to query/publication phases and does not infer wall-time or speedup.
"""
import argparse
import collections
import hashlib
import json
import re
from pathlib import Path

HEADER = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>\d+)/(?P<tid>\d+)\s+[^:]+:\s+"
    r"(?P<period>\d+)\s+cycles:u:\s+(?P<leaf>.*)$"
)

def family(leaf: str) -> str:
    low = leaf.lower()
    if "[unknown]" in low:
        return "unknown"
    if "sha2::sha256::x86_sha::compress" in low:
        return "sha256"
    if "litchi_perf_baseline::" in low:
        return "harness"
    if "litchi_opc::source_backed::" in low:
        return "source_provider"
    # Keep allocator/runtime wrapper entries in the residual bucket. This
    # makes the memory bucket mean the explicit libc/string primitives used in
    # the report, rather than every allocation-adjacent symbol.
    excluded_runtime = (
        "unlink_chunk", "__cxa_atexit", "__rustc::__rdl_dealloc",
        "__rustc::__rust_dealloc", " realloc+",
    )
    memory_names = (
        "memmove", "memcpy", "memcmp", "malloc", "calloc", "free",
        "memset", "memchr",
    )
    if (any(name in low for name in memory_names) or "libc.so" in low) and not any(
        name in low for name in excluded_runtime
    ):
        return "memory_libc"
    if any(name in low for name in (
        "quick_xml", "litchi_docx", "litchi_ooxml_common", "xml_minifier",
    )):
        return "xml_layout"
    if any(name in low for name in (
        "zlib_rs", "crc32fast", "zip", "deflate", "inflate", "adler",
        "flate2::mem::decompress",
    )):
        return "zip_deflate_crc"
    if "ld-linux" in low:
        return "loader"
    return "other"

def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input", type=Path)
    ap.add_argument("--output", type=Path)
    args = ap.parse_args()
    raw = args.input.read_bytes()
    rows = []
    for line in raw.decode("utf-8").splitlines():
        m = HEADER.match(line)
        if m:
            rows.append({
                "comm": m.group("comm"),
                "pid": int(m.group("pid")),
                "tid": int(m.group("tid")),
                "period": int(m.group("period")),
                "leaf": m.group("leaf"),
            })
    by_family = collections.defaultdict(lambda: {"samples": 0, "period": 0})
    by_comm = collections.defaultdict(lambda: {"samples": 0, "period": 0})
    for row in rows:
        f = family(row["leaf"])
        by_family[f]["samples"] += 1
        by_family[f]["period"] += row["period"]
        by_comm[row["comm"]]["samples"] += 1
        by_comm[row["comm"]]["period"] += row["period"]
    total = sum(row["period"] for row in rows)
    result = {
        "schema": "litchi.performance.0928.leaf-period-classification.v1",
        "scope": "whole_process_direct_sampled_ip_descriptive_only",
        "input": str(args.input),
        "input_sha256": hashlib.sha256(raw).hexdigest(),
        "sample_count": len(rows),
        "period_total": total,
        "families": {
            k: {**v, "period_share": v["period"] / total if total else 0.0}
            for k, v in sorted(by_family.items(), key=lambda kv: -kv[1]["period"])
        },
        "commands": {
            "export": "perf script --hide-call-graph -F comm,pid,tid,time,event,period,ip,sym,symoff,dso --demangle --input=perf.data",
            "classify": "python3 -B classify_leaf_period_v1.py perf-script-leaf-period.txt --output leaf-period-classification-v1.json",
        },
        "comm_counts": {
            k: v for k, v in sorted(by_comm.items(), key=lambda kv: -kv[1]["period"])
        },
    }
    encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.write_text(encoded)
    else:
        print(encoded, end="")

if __name__ == "__main__":
    main()
