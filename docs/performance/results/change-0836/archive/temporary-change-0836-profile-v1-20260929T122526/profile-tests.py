#!/usr/bin/env python3
"""Synthetic mutation tests for the 0836 perf/strace reader.

These tests exercise the parser in memory.  They deliberately mutate PIDs,
DSOs, symbol ranges, stack records, lost-event lines, and unfinished strace
pairs so a future reader cannot silently broaden the measured population.
When the retained analysis exists, its cardinality and shared-population
contract are checked as an additional binding.  No capture, build, decoder,
or Git command is executed.
"""

from __future__ import annotations

from copy import deepcopy
import hashlib
import json
from pathlib import Path
import sys
from typing import Any, Callable

import profile_analysis as analysis


HERE = Path(__file__).resolve().parent


def expect_reject(name: str, operation: Callable[[], Any], checks: list[dict[str, Any]]) -> None:
    try:
        operation()
    except (analysis.EvidenceError, AssertionError, KeyError, TypeError, ValueError):
        checks.append({"name": name, "rejected": True})
    else:
        raise AssertionError(f"mutation was accepted: {name}")


def frame(address: int, symbol: str, dso: str = "/tmp/0836-binary") -> str:
    return f"  {address:x} {symbol}+0x0 ({dso})"


def header(pid: int, timestamp: str, period: int = 100) -> str:
    return f"fp {pid} {timestamp}: {period} cycles:u:"


def perf_stream(rows: list[tuple[int, str, list[str]]], *, lost: bool = False) -> bytes:
    blocks = []
    for pid, timestamp, frames in rows:
        blocks.append("\n".join([header(pid, timestamp), *frames]))
    text = "\n\n".join(blocks) + "\n\n"
    if lost:
        text += "PERF_RECORD_LOST: 1 events\n"
    return text.encode()


def stage_info() -> dict[str, Any]:
    owner = "owner::write_single_part_overlay_to_stream"
    return {
        "owner": owner,
        "binary": {"path": "/tmp/0836-binary", "bytes": 1, "sha256": "0" * 64},
        "ranges": [{"address": 0x1000, "size": 0x100, "end": 0x1100,
                    "symbol": owner, "raw_symbol": owner}],
    }


def perf_synthetic_checks(checks: list[dict[str, Any]]) -> None:
    info = stage_info()
    owner = info["owner"]
    rows = [
        (42, "1.000000", [frame(0x2000, "zlib::deflate"), frame(0x1001, owner)]),
        (42, "1.000001", [frame(0x2100, "zlib::inflate"), frame(0x1002, owner)]),
        (42, "1.000002", [frame(0x2200, "copy_crc"), frame(0x1003, owner)]),
        (42, "1.000003", [frame(0x2300, "ordinary_leaf"), frame(0x1004, owner)]),
    ]
    valid = analysis.parse_perf_text(perf_stream(rows), {42}, info, label="synthetic.valid")
    if valid["owner_qualified_samples"] != 4:
        raise AssertionError("valid synthetic owner population changed")
    if [row["group"] for row in valid["disjoint_groups"]] != [
        "deflate", "inflate", "copy-crc", "other"
    ]:
        raise AssertionError("codec/copy group order changed")
    if sum(row["samples"] for row in valid["disjoint_groups"]) != 4:
        raise AssertionError("disjoint groups do not reconstruct owner samples")
    checks.append({"name": "valid-owner-range-and-disjoint-groups", "rejected": False,
                   "owner_samples": valid["owner_qualified_samples"]})

    expect_reject(
        "wrong-perf-event",
        lambda: analysis.parse_perf_text(
            perf_stream(rows).replace(b"cycles:u", b"instructions:u"), {42}, info,
            label="synthetic.wrong-event"),
        checks,
    )
    wrong_pid = analysis.parse_perf_text(perf_stream(rows), {99}, info, label="synthetic.pid")
    if wrong_pid["partition"]["outside_pid"] != 4 or wrong_pid["owner_qualified_samples"]:
        raise AssertionError("outside-PID mutation was not isolated")
    checks.append({"name": "outside-pid-isolation", "rejected": False,
                   "outside_pid": wrong_pid["partition"]["outside_pid"]})

    wrong_dso = perf_stream([(42, "2.000000", [frame(0x1001, owner, "/tmp/other")])])
    wrong_dso_result = analysis.parse_perf_text(wrong_dso, {42}, info, label="synthetic.dso")
    if wrong_dso_result["owner_qualified_samples"] != 0 \
            or wrong_dso_result["owner_diagnostics"]["owner_symbol_other_dso_frames"] != 1:
        raise AssertionError("wrong-DSO owner mutation was not isolated")
    checks.append({"name": "wrong-owner-dso-isolation", "rejected": False,
                   "outside_owner": wrong_dso_result["partition"]["outside_owner"]})

    range_miss = analysis.parse_perf_text(
        perf_stream([(42, "3.000000", [frame(0x3001, owner)])]), {42}, info,
        label="synthetic.range")
    if range_miss["owner_qualified_samples"] != 0 \
            or range_miss["owner_diagnostics"]["owner_address_or_symbol_range_mismatch_frames"] != 1:
        raise AssertionError("owner range mutation was not isolated")
    checks.append({"name": "owner-address-outside-admitted-range", "rejected": False})


def perf_diagnostic_checks(checks: list[dict[str, Any]]) -> None:
    info = stage_info()
    owner = info["owner"]
    malformed = perf_stream([(42, "4.000000", [frame(0x1001, owner), "  not-a-frame"] )])
    parsed = analysis.parse_perf_text(malformed, {42}, info, label="synthetic.malformed")
    if parsed["decode_diagnostics"]["malformed_samples"] != 1 \
            or parsed["cpu_fraction_authorized"]:
        raise AssertionError("malformed sample was not retained/refused")
    checks.append({"name": "malformed-frame-refuses-cpu-fraction", "rejected": False})

    lost = analysis.parse_perf_text(
        perf_stream([(42, "5.000000", [frame(0x1001, owner)])], lost=True),
        {42}, info, label="synthetic.lost")
    if lost["decode_diagnostics"]["lost_event_count"] != 1 \
            or lost["cpu_fraction_authorized"]:
        raise AssertionError("lost event was not retained/refused")
    checks.append({"name": "lost-event-refuses-cpu-fraction", "rejected": False})

    unknown = perf_stream([(42, "6.000000", [frame(0x1001, owner),
                                               "  2000 [unknown] ([unknown])"])])
    unknown_result = analysis.parse_perf_text(unknown, {42}, info, label="synthetic.unknown")
    if unknown_result["decode_diagnostics"]["unknown_stack_samples"] != 1 \
            or unknown_result["cpu_fraction_authorized"]:
        raise AssertionError("unknown stack was not retained/refused")
    checks.append({"name": "unknown-stack-refuses-cpu-fraction", "rejected": False})


def trace_stream(*, resumed: bool = True, outside: bool = False) -> bytes:
    lines = [
        "42 100.000000 fsync(3</tmp/source>) = 0 <0.000001>",
    ]
    if resumed:
        lines.extend([
            "42 100.000001 renameat(AT_FDCWD</tmp>, \"source\", AT_FDCWD</tmp>, \"dest\", 0 <unfinished ...>",
            "42 100.000002 <... renameat resumed> ) = 0 <0.000002>",
        ])
    else:
        lines.append("42 100.000001 <... renameat resumed> ) = 0 <0.000002>")
    lines.append("42 100.000003 fsync(3</tmp/dest>) = 0 <0.000001>")
    if outside:
        lines.append("41 100.000004 fsync(3</tmp/other>) = 0 <0.000001>")
    return ("\n".join(lines) + "\n").encode()


def trace_synthetic_checks(checks: list[dict[str, Any]]) -> None:
    valid = analysis.parse_trace_text(trace_stream(), {42}, label="synthetic.trace")
    if not valid["order_provenance_verified"] or valid["pid_rows"][0]["relevant_order"] != [
        "fsync", "renameat", "fsync"
    ]:
        raise AssertionError("valid trace order/provenance changed")
    checks.append({"name": "valid-fsync-rename-fsync-provenance", "rejected": False})

    outside = analysis.parse_trace_text(trace_stream(outside=True), {42}, label="synthetic.trace-outside")
    if outside["outside_pid_events"] != 1:
        raise AssertionError("outside trace PID was not retained")
    checks.append({"name": "outside-trace-pid-retained", "rejected": False,
                   "outside_events": outside["outside_pid_events"]})

    expect_reject(
        "unmatched-resumed-syscall",
        lambda: analysis.parse_trace_text(trace_stream(resumed=False), {42},
                                          label="synthetic.trace-unmatched"),
        checks,
    )

    wrong_order = trace_stream().replace(
        b"42 100.000003 fsync(3</tmp/dest>) = 0 <0.000001>",
        b"42 100.000003 rename(\"/tmp/dest\", \"/tmp/dest2\") = 0 <0.000001>",
    )
    expect_reject(
        "wrong-durability-order",
        lambda: analysis.parse_trace_text(wrong_order, {42}, label="synthetic.trace-order"),
        checks,
    )


def retained_binding(checks: list[dict[str, Any]]) -> str | None:
    path = HERE / "profile-analysis.json"
    if not path.is_file():
        raise AssertionError("normal mutation run requires retained profile analysis")
    value = analysis.read_json(path)
    if value.get("schema") != analysis.ANALYSIS_SCHEMA or value.get("status") != "pass":
        raise AssertionError("retained profile analysis schema/status changed")
    if (value.get("reports"), value.get("samples")) != (8, 152):
        raise AssertionError("retained shared analysis cardinality changed")
    if value.get("cpu_reports") != 4 or value.get("cpu_samples") != 120 \
            or value.get("trace_reports") != 4 or value.get("trace_samples") != 32:
        raise AssertionError("retained shared component cardinality changed")
    profiles = value.get("profiles", {}).get("captures", [])
    traces = value.get("traces", {}).get("captures", [])
    if len(profiles) != 4 or len(traces) != 4:
        raise AssertionError("retained capture matrix changed")
    checks.append({"name": "retained-shared-analysis-binding", "rejected": False,
                   "reports": 8, "samples": 152})
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    allowed = ([], ["--check"], ["--preflight"], ["--preflight", "--check"])
    if sys.argv[1:] not in allowed:
        print("usage: profile-tests.py [--preflight] [--check]", file=sys.stderr)
        return 2
    preflight = "--preflight" in sys.argv
    check = "--check" in sys.argv
    checks: list[dict[str, Any]] = []
    try:
        perf_synthetic_checks(checks)
        perf_diagnostic_checks(checks)
        trace_synthetic_checks(checks)
        retained = None if preflight else retained_binding(checks)
        result = {
            "schema": "litchi.0836.profile-mutation-tests.v1",
            "status": "pass", "mode": "preflight" if preflight else "normal",
            "checks": checks,
            "rejected_mutation_count": sum(item["rejected"] for item in checks),
            "retained_analysis_sha256": retained,
            "scope": "synthetic parser/census mutations; no workload or production claim",
        }
        destination = HERE / ("profile-preflight-tests.json" if preflight else "profile-tests.json")
        encoded = (json.dumps(result, indent=2, sort_keys=True) + "\n").encode("utf-8")
        if check:
            if not destination.is_file() or destination.read_bytes() != encoded:
                raise AssertionError(f"{destination.name} differs from deterministic replay")
        else:
            if destination.exists():
                raise AssertionError(f"refusing to overwrite {destination.name}")
            with destination.open("xb") as stream:
                stream.write(encoded)
    except (analysis.EvidenceError, AssertionError, OSError, UnicodeError,
            TypeError, ValueError, KeyError) as error:
        print(f"0836 profile mutation tests failed: {error}", file=sys.stderr)
        return 1
    print(f"0836 profile mutation tests PASS: {len(checks)} checks")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
