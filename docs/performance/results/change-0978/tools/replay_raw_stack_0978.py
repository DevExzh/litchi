#!/usr/bin/env python3
"""Offline replay of the variable-layout scanner on retained 0977 raw dumps."""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import raw_stack_0978 as scanner  # noqa: E402


ROOT = Path("/home/zhuhe/litchi-performance-0977/capture/instruction-profile-v2")
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
PERIOD = 1_000_000
STACK = 32_768


def child_inputs(number: int) -> tuple[Path, Path, Path, Path, Path, Path]:
    child = ROOT / f"child-{number:02d}"
    stem = f"child-{number:02d}"
    return (
        child / f"{stem}.perf-mmap-dump.txt",
        child / f"{stem}.perf.data",
        child / f"{stem}.perf-evlist-v.txt",
        child / f"{stem}.launch.json",
        child / f"{stem}.report.json",
        child / f"{stem}.terminal.json",
    )


def binding(path: Path) -> dict:
    return scanner.file_binding(path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--child", type=int, action="append", default=None)
    parser.add_argument("--output", type=Path, default=Path("/home/zhuhe/litchi-performance-0978/receipts/replay-0977-child01-02-v1.json"))
    args = parser.parse_args()
    children = args.child if args.child else [1, 2]
    output: dict = {
        "schema": "litchi.performance.0978.raw_stack_replay.v1",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "tool": binding(Path(__file__).resolve()),
        "scanner": binding(Path(scanner.__file__).resolve()),
        "uapi": {
            "path": str(scanner.LINUX_PERF_EVENT_HEADER),
            "sha256": scanner.sha256(scanner.LINUX_PERF_EVENT_HEADER),
            "upstream_evsel_source": scanner.UPSTREAM_EVSEL_SOURCE,
        },
        "binary": binding(BINARY),
        "children": [],
        "scope": {
            "offline_only": True,
            "workload_or_build": False,
            "sealed_inputs_0977_modified": False,
            "sample_ledger_retained": False,
            "payload_presence_is_not_unwind_proof": True,
        },
    }
    for number in children:
        dump, perf_data, evlist, launch_path, report, terminal = child_inputs(number)
        launch = json.loads(launch_path.read_text())
        event_contract = scanner.parse_evlist_contract(
            evlist.read_text(), expected_period=PERIOD, expected_stack_cap=STACK,
        )
        summary = scanner.scan_dump(
            dump, perf_data=perf_data, binary=BINARY,
            expected_pid=int(launch["exec_pid"]), expected_cpu=31,
            expected_period=PERIOD, stack_cap=STACK,
            event_contract=event_contract, retain_samples=False,
        )
        output["children"].append({
            "child": number,
            "inputs": {
                "dump": binding(dump), "perf_data": binding(perf_data),
                "evlist": binding(evlist), "launch": binding(launch_path),
                "report": binding(report), "terminal": binding(terminal),
            },
            "expected_pid": int(launch["exec_pid"]),
            "summary": summary,
        })
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(output, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"output": str(args.output), "children": [
        {"child": row["child"], "status": row["summary"]["status"], "samples": row["summary"]["sample_count"], "raw": row["summary"]["raw_event_count"], "violations": len(row["summary"]["violations"])}
        for row in output["children"]
    ]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
