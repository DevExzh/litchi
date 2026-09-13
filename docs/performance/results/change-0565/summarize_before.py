#!/usr/bin/env python3
"""Summarize one capture directory produced by capture_before.sh.

Reuses `summarize()` and `counted()` from the retained change-0564 summarizer
(docs/performance/results/change-0564/summarize_open_reads.py) rather than
re-implementing them: per-operation pread64 and statx are isolated by
differencing the `strace -f -c` pair at 1 and 11 samples and dividing by ten;
the size histogram comes from the full `strace -f -e trace=pread64` capture of
one child with one warmup and one sample (two operations).

Usage: summarize_before.py --capture DIR [--summarizer PATH] [--output DIR/summary.json]
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import pathlib

DEFAULT_SUMMARIZER = "/home/zhuhe/code/litchi/docs/performance/results/change-0564/summarize_open_reads.py"
MODES = (("file-source", "xls_file_source"), ("owned-readat", "xls_owned_readat"))
OPERATIONS = ("open", "list", "one-cell")
LOGICAL = ("read_calls", "read_bytes", "version_calls", "len_calls", "seek_calls")


def load_summarizer(path: str):
    spec = importlib.util.spec_from_file_location("summarize_open_reads", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--capture", required=True)
    parser.add_argument("--summarizer", default=DEFAULT_SUMMARIZER)
    parser.add_argument("--output")
    args = parser.parse_args()
    root = pathlib.Path(args.capture)
    s = load_summarizer(args.summarizer)

    operations = {}
    for mode, stem in MODES:
        for op in OPERATIONS:
            key = f"{mode}/{op}"
            entry = {"mode": mode, "operation": op}
            attribution = root / "attribution" / f"{mode}-{op}.json"
            if attribution.exists():
                report = json.loads(attribution.read_text())
                metrics = [r["metrics"] for r in report["records"]]
                first = metrics[0]
                for other in metrics[1:]:
                    for name in LOGICAL:
                        if other[name] != first[name]:
                            raise SystemExit(f"{attribution.name}: {name} varies across samples")
                entry["logical"] = {name: first[name] for name in LOGICAL}
                entry["logical"]["elapsed_ns_samples"] = report["elapsed_samples_ns"]
                entry["logical"]["source_version_stable"] = all(r["source_version_stable"] for r in report["records"])
                entry["logical"]["warmups"] = report["warmups"]
                entry["logical"]["samples"] = report["samples"]
                entry["logical"]["binary_sha256"] = report["binary"]["sha256"]
                entry["logical"]["input_sha256"] = report["input"]["sha256"]
                entry["logical"]["observation"] = report["records"][0]["observation"]
            low = root / "syscalls" / f"{stem}_{op}.samples-1.strace.txt"
            high = root / "syscalls" / f"{stem}_{op}.samples-11.strace.txt"
            if low.exists() and high.exists():
                entry["isolation"] = {
                    "pread64_per_op": (s.counted(high, "pread64") - s.counted(low, "pread64")) / 10,
                    "statx_per_op": (s.counted(high, "statx") - s.counted(low, "statx")) / 10,
                    "whole_child_samples_1": {"pread64": s.counted(low, "pread64"), "statx": s.counted(low, "statx")},
                    "whole_child_samples_11": {"pread64": s.counted(high, "pread64"), "statx": s.counted(high, "statx")},
                    "method": "one warmup plus 1 and 11 samples; the difference divided by ten isolates one operation",
                }
            trace = root / "traces" / f"{stem}_{op}.pread.txt"
            if trace.exists():
                entry["trace"] = s.summarize(trace, opens=2)
            operations[key] = entry

    result = {
        "schema_version": 1,
        "summary_kind": "litchi-perf-change-0565-capture-summary",
        "capture": str(root),
        "summarizer": args.summarizer,
        "operations": operations,
    }
    output = pathlib.Path(args.output or (root / "summary.json"))
    output.write_text(json.dumps(result, indent=2) + "\n")
    print(f"{'operation':22s} {'pread64/op':>10s} {'statx/op':>9s} {'read_calls':>10s} {'read_bytes':>10s} {'version':>8s} {'len':>4s}  4-byte share  histogram")
    for key, e in operations.items():
        iso, log, tr = e.get("isolation") or {}, e.get("logical") or {}, e.get("trace") or {}
        print(f"{key:22s} {iso.get('pread64_per_op', float('nan')):10.1f} {iso.get('statx_per_op', float('nan')):9.1f} "
              f"{log.get('read_calls', 0):10d} {log.get('read_bytes', 0):10d} {log.get('version_calls', 0):8d} {log.get('len_calls', 0):4d}  "
              f"{tr.get('four_byte_share_percent', float('nan')):6.1f}%      {tr.get('size_histogram')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
