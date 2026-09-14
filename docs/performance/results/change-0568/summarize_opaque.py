#!/usr/bin/env python3
"""Extract change-0568 before numbers on the comments-opaque-heavy corpus.

litchi-perf-baseline's xls_source_backed_* / xls_owned_source_* selectors run on
the in-memory `xls-comments-opaque-heavy` corpus and publish per-sample
deterministic logical counters under results[].source, including the exact
quantity change 0568 moves: selected_worksheet_read_calls / _read_bytes.
Every counter is asserted constant across all samples; a varying one is fatal.
"""
from __future__ import annotations

import argparse
import json
import pathlib
import statistics
import sys

FLAT = ("read_calls", "read_bytes", "ordinary_payload_read_calls",
        "ordinary_payload_read_bytes", "max_in_flight_reads")
XLS = ("source_retained_bytes", "complete_archive_materialized_bytes",
       "parsed_sheet_counts", "parsed_cell_counts", "source_version_checks",
       "cfb_structural_read_calls", "cfb_structural_read_bytes",
       "workbook_global_read_calls", "workbook_global_read_bytes",
       "selected_worksheet_read_calls", "selected_worksheet_read_bytes",
       "unselected_worksheet_read_calls", "unselected_worksheet_read_bytes",
       "opaque_payload_read_calls", "opaque_payload_read_bytes")


def constant(name: str, values: list, where: str):
    uniq = set(values)
    if len(uniq) != 1:
        raise SystemExit(f"{where}: {name} varies across samples: {sorted(uniq)[:8]}")
    return uniq.pop()


def quantile(values, fraction):
    ordered = sorted(values)
    if len(ordered) == 1:
        return float(ordered[0])
    position = fraction * (len(ordered) - 1)
    low = int(position)
    high = min(low + 1, len(ordered) - 1)
    return ordered[low] + (ordered[high] - ordered[low]) * (position - low)


def counted(path: pathlib.Path, name: str) -> int:
    if not path.exists():
        return 0
    for line in path.read_text().splitlines():
        fields = line.split()
        if len(fields) >= 4 and fields[-1] == name:
            return int(fields[3])
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--capture", required=True)
    ap.add_argument("--output")
    args = ap.parse_args()
    root = pathlib.Path(args.capture)

    selectors = {}
    for report in sorted((root / "harness").glob("*.json")):
        if report.name.endswith(".receipt.json"):
            continue
        sel = report.stem
        doc = json.loads(report.read_text())
        result = doc["results"][0]
        src = result.get("source") or {}
        xls = src.get("xls") or {}
        elapsed = result["elapsed_ns"]["samples"]
        entry = {
            "selector": sel,
            "corpus": {k: result["corpus"][k] for k in
                       ("name", "generator", "shape", "entry_count", "archive_member_count",
                        "archive_bytes", "target_entry", "target_payload_bytes",
                        "archive_sha256", "target_payload_sha256")},
            "instrumented_source_block_present": bool(src),
            "implementation": xls.get("implementation"),
            "operation": xls.get("operation"),
            "samples": len(elapsed),
            "warmup": doc["configuration"]["warmup_iterations_per_case"],
            "logical": {name: constant(name, src[name], sel) for name in FLAT if name in src},
            "xls_logical": {name: constant(name, xls[name], sel) for name in XLS if name in xls},
            "latency_ns": {"p50": quantile(elapsed, 0.50), "p95": quantile(elapsed, 0.95),
                           "p99": quantile(elapsed, 0.99), "mean": statistics.fmean(elapsed),
                           "min": min(elapsed), "max": max(elapsed)},
            "environment": doc["environment"],
        }
        low = root / "syscalls" / f"{sel}.samples-1.strace.txt"
        high = root / "syscalls" / f"{sel}.samples-11.strace.txt"
        if low.exists() and high.exists():
            entry["isolation"] = {
                "pread64_per_op": (counted(high, "pread64") - counted(low, "pread64")) / 10,
                "statx_per_op": (counted(high, "statx") - counted(low, "statx")) / 10,
                "whole_child_samples_1": {"pread64": counted(low, "pread64"), "statx": counted(low, "statx")},
                "whole_child_samples_11": {"pread64": counted(high, "pread64"), "statx": counted(high, "statx")},
                "method": "one warmup plus 1 and 11 samples; the difference divided by ten isolates one operation",
                "note": ("this corpus is an in-memory Vec<u8> ReadAt, not a file, so a zero here is "
                         "expected and says the selector performs no filesystem I/O at all"),
            }
        selectors[sel] = entry

    out = {
        "schema_version": 1,
        "summary_kind": "litchi-perf-change-0568-opaque-corpus-before",
        "capture": str(root),
        "corpus_note": ("the opaque bulk of xls-comments-opaque-heavy is 8 sibling CFB streams "
                        "(OpaquePayloads/Payload000..007, 2 MiB each) plus OpaqueMetadata; the "
                        "Workbook stream's worksheet substreams are made of small MSODRAWING/OBJ/"
                        "TXO/CONTINUE/NOTE records, so this corpus does NOT exercise change 0568's "
                        "1 KiB mean-framed-bytes-per-record density gate"),
        "selectors": selectors,
    }
    if args.output:
        pathlib.Path(args.output).write_text(json.dumps(out, indent=2) + "\n")
    header = (f"{'selector':42s} {'reads':>6s} {'bytes':>8s} {'cfb':>5s} {'glob':>5s} "
              f"{'selWS':>6s} {'selWSb':>7s} {'unsel':>6s} {'opaque':>7s} {'ver':>4s} "
              f"{'pread':>6s} {'p50 ns':>9s}")
    print(header)
    for sel, e in selectors.items():
        lg, x, iso = e["logical"], e["xls_logical"], e.get("isolation") or {}
        if not x:
            print(f"{sel:42s} {'-':>6s} {'-':>8s} {'-':>5s} {'-':>5s} {'-':>6s} {'-':>7s} "
                  f"{'-':>6s} {'-':>7s} {'-':>4s} {iso.get('pread64_per_op', float('nan')):6.1f} "
                  f"{e['latency_ns']['p50']:9.1f}   (no instrumented source block)")
            continue
        print(f"{sel:42s} {lg['read_calls']:6d} {lg['read_bytes']:8d} "
              f"{x['cfb_structural_read_calls']:5d} {x['workbook_global_read_calls']:5d} "
              f"{x['selected_worksheet_read_calls']:6d} {x['selected_worksheet_read_bytes']:7d} "
              f"{x['unselected_worksheet_read_calls']:6d} {x['opaque_payload_read_calls']:7d} "
              f"{x['source_version_checks']:4d} {iso.get('pread64_per_op', float('nan')):6.1f} "
              f"{e['latency_ns']['p50']:9.1f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
