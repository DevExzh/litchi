#!/usr/bin/env python3
"""Parse one retained perf -D dump with the frozen 0978 lifecycle helper.

Offline only: no perf invocation, native workload, build, or input mutation.
The helper and raw dump are hashed before and after the single parse.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import sys
from pathlib import Path
from typing import Any

TOOLS = Path(__file__).resolve().parent
HELPER = TOOLS / "lifecycle_ledger_0978.py"
EXPECTED_PID = 3849325
EXPECTED_CPU = 31


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if not path.is_file():
        raise RuntimeError(f"missing input: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_helper(path: Path) -> Any:
    spec = importlib.util.spec_from_file_location("lifecycle_ledger_0978_frozen_reparse", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def write_exclusive(path: Path, value: Any) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True, allow_nan=False)
        handle.write("\n")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dump", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--receipt", type=Path, required=True)
    parser.add_argument("--pid", type=int, default=EXPECTED_PID)
    parser.add_argument("--cpu", type=int, default=EXPECTED_CPU)
    args = parser.parse_args(argv)
    dump = args.dump.resolve()
    output = args.output.resolve()
    receipt_path = args.receipt.resolve()
    script = Path(__file__).resolve()
    before = {"script": binding(script), "helper": binding(HELPER), "dump": binding(dump)}
    lifecycle = load_helper(HELPER)
    stream = lifecycle.parse_lifecycle_dump(dump, expected_pid=args.pid, expected_cpu=args.cpu, require_raw_metadata=True)
    after = {"script": binding(script), "helper": binding(HELPER), "dump": binding(dump)}
    if before != after:
        raise RuntimeError("bound script/helper/dump changed during parse")
    if stream.get("status") != "complete" or stream.get("raw_metadata_crosscheck") is not True or stream.get("unique_record_offsets") is not True:
        raise RuntimeError("frozen parser did not prove complete metadata custody")
    if stream.get("raw_lost_sum") != 0 or stream.get("terminal_lost_values") not in ([], [0]):
        raise RuntimeError("frozen parser found lifecycle loss")
    write_exclusive(output, stream)
    output_binding = binding(output)
    receipt = {
        "schema": "litchi.performance.0978.lifecycle_stream_reparse.v1",
        "status": "pass",
        "scope": "one offline parse of retained child-02 perf -D dump with frozen helper",
        "native_actions": {"build": False, "recording": False, "workload": False},
        "command": {"argv": [sys.executable, "-B", str(script), "--dump", str(dump), "--output", str(output), "--receipt", str(receipt_path), "--pid", str(args.pid), "--cpu", str(args.cpu)], "cwd": os.getcwd(), "offline_only": True, "perf_subprocess": False},
        "source_before": before,
        "source_after": after,
        "result": {"record_count": stream.get("record_count"), "sample_descriptor_count": stream.get("sample_descriptor_count"), "raw_lost_sum": stream.get("raw_lost_sum"), "terminal_lost_values": stream.get("terminal_lost_values"), "raw_metadata_crosscheck": stream.get("raw_metadata_crosscheck"), "unique_record_offsets": stream.get("unique_record_offsets")},
        "output": output_binding,
    }
    write_exclusive(receipt_path, receipt)
    print(json.dumps({"status": "pass", "output": output_binding, "receipt": str(receipt_path)} , sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"status": "failed", "error": str(exc)}, sort_keys=True), file=sys.stderr)
        raise SystemExit(1)
