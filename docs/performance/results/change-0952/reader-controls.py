#!/usr/bin/env python3
"""Standalone real-map positive and fail-closed controls for analyze_0952_v3.

The controls use the retained 0952 plan, ELF and mmap export.  Only synthetic
one-sample text files are made under a temporary directory; no workload,
benchmark or perf command is started.  The real-map positive checks the owner
and parent ranges, while each negative checks a separate binding or parser
failure.  This is intentionally packet-bound, like the v3 reader.
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PACKET = Path("/home/zhuhe/litchi-performance-0952")
READER = PACKET / "tools/analyze_0952_v3.py"
SAMPLE_EVENTS = PACKET / "analysis/samples/00-medium/events.txt"
MMAP_EVENTS = PACKET / "analysis/mmap/00-medium/events.txt"


def load_reader() -> Any:
    spec = importlib.util.spec_from_file_location("analyze_0952_v3_controls_reader", READER)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load v3 reader")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def write_temp(root: Path, name: str, contents: str) -> Path:
    path = root / name
    path.write_text(contents, encoding="utf-8", newline="")
    return path


def sample_text(pid: int, when: str, owner_ip: int, parent_ip: int,
                binary: str, *, parent_dso: str | None = None,
                owner_ip_override: int | None = None,
                event: str = "cycles:u") -> str:
    owner = owner_ip if owner_ip_override is None else owner_ip_override
    parent = parent_dso if parent_dso is not None else binary
    return (
        f"litchi-perf-bas {pid}/{pid} {when}: 7 {event}:\n"
        f"    {owner:#x} owner ({binary})\n"
        f"    {parent_ip:#x} parent ({parent})\n"
    )


def expect_error(label: str, action: Any, needle: str) -> dict[str, Any]:
    try:
        action()
    except Exception as exc:  # the v3 reader exposes typed AdmissionError
        text = str(exc)
        if needle not in text:
            raise AssertionError(f"{label}: error {text!r} lacks {needle!r}") from exc
        return {"status": "rejected", "error_type": type(exc).__name__, "error": text}
    raise AssertionError(f"{label}: malformed input was accepted")


def run(receipt_path: Path) -> dict[str, Any]:
    reader = load_reader()
    plan = reader.load_json(reader.PLAN)
    binary = Path(plan["binary"]["path"])
    loads, build_id = reader.read_elf(binary)
    mmap = reader.parse_mmap_events(MMAP_EVENTS, binary, loads, build_id)
    maps = mmap["rows"]
    executable = [row for row in mmap["bound_rows"]
                  if "x" in row["perm"] and not row.get("mapping_ambiguous")]
    if len(executable) != 1:
        raise AssertionError(f"expected one unambiguous executable map, got {len(executable)}")
    exec_map = executable[0]
    owner = plan["static_symbols"]["owner_ranges"][0]
    parent = plan["static_symbols"]["parent_range"]
    owner_static = int(owner["address"], 16)
    parent_static = int(parent["address"], 16)
    bias = exec_map["load_bias"]
    owner_ip = bias + owner_static
    parent_ip = bias + parent_static
    if not (exec_map["start"] <= owner_ip < exec_map["start"] + exec_map["size"]):
        raise AssertionError("owner synthetic IP is outside real executable mapping")
    if not (exec_map["start"] <= parent_ip < exec_map["start"] + exec_map["size"]):
        raise AssertionError("parent synthetic IP is outside real executable mapping")
    pid = int(exec_map["pid"])
    when = max(row["time"] for row in mmap["bound_rows"]) + reader.DECIMAL("1")
    when_text = format(when, "f")
    binary_text = str(binary)
    results: dict[str, Any] = {}
    with tempfile.TemporaryDirectory(prefix="analyze-0952-v3-controls-") as raw:
        root = Path(raw)
        positive = write_temp(root, "positive-owner-parent.txt",
                              sample_text(pid, when_text, owner_ip, parent_ip, binary_text))
        positive_profile = reader.parse_samples(positive, binary, mmap, plan)
        if positive_profile["all"] != {"samples": 1, "period": 7}:
            raise AssertionError(f"positive all profile differs: {positive_profile['all']}")
        if positive_profile["owner_qualified"] != {"samples": 1, "period": 7}:
            raise AssertionError(f"positive owner profile differs: {positive_profile['owner_qualified']}")
        results["positive_owner_parent_real_map"] = {
            "status": "qualified",
            "profile": positive_profile,
            "owner_static": hex(owner_static),
            "parent_static": hex(parent_static),
            "load_bias": hex(bias),
            "runtime_owner": hex(owner_ip),
            "runtime_parent": hex(parent_ip),
            "pid": pid,
        }

        wrong_parent = write_temp(root, "wrong-parent-dso.txt",
                                  sample_text(pid, when_text, owner_ip, parent_ip, binary_text,
                                              parent_dso="/lib/x86_64-linux-gnu/libc.so.6"))
        wrong_parent_profile = reader.parse_samples(wrong_parent, binary, mmap, plan)
        if wrong_parent_profile["owner_qualified"]["samples"] != 0:
            raise AssertionError("wrong-parent DSO unexpectedly qualified owner")
        results["wrong_parent_dso"] = {"status": "rejected_for_qualification",
                                        "profile": wrong_parent_profile}

        outside = write_temp(root, "outside-owner-ip.txt",
                             sample_text(pid, when_text, owner_ip, parent_ip, binary_text,
                                         owner_ip_override=0x7FFF00000000))
        outside_profile = reader.parse_samples(outside, binary, mmap, plan)
        if outside_profile["owner_qualified"]["samples"] != 0:
            raise AssertionError("outside owner IP unexpectedly qualified owner")
        results["outside_executable_ip"] = {"status": "rejected_for_qualification",
                                              "profile": outside_profile}

        wrong_pid = write_temp(root, "wrong-pid.txt",
                               sample_text(pid + 1, when_text, owner_ip, parent_ip, binary_text))
        wrong_pid_profile = reader.parse_samples(wrong_pid, binary, mmap, plan)
        if wrong_pid_profile["owner_qualified"]["samples"] != 0:
            raise AssertionError("wrong PID unexpectedly qualified owner")
        results["wrong_pid"] = {"status": "rejected_for_qualification",
                                 "profile": wrong_pid_profile}

        mmap_text = MMAP_EVENTS.read_text(encoding="utf-8")
        wrong_build_id = mmap_text.replace(build_id, "0" * len(build_id))
        wrong_build_id_path = write_temp(root, "wrong-build-id-mmap.txt", wrong_build_id)
        results["wrong_build_id"] = expect_error(
            "wrong build-id", lambda: reader.parse_mmap_events(wrong_build_id_path, binary, loads, build_id),
            "bound DSO build-id differs")

        wrong_event = write_temp(root, "wrong-event.txt",
                                 sample_text(pid, when_text, owner_ip, parent_ip, binary_text,
                                             event="instructions:u"))
        results["wrong_event"] = expect_error(
            "wrong event", lambda: reader.parse_samples(wrong_event, binary, mmap, plan),
            "unexpected perf event")

    result = {
        "schema": "litchi.performance.0952.native-admission.v3-controls.v1",
        "status": "pass",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "scope": "Standalone parser and map-binding controls only; no workload or perf capture.",
        "reader": {"path": str(READER.resolve()), "sha256": reader.sha(READER),
                   "schema": "litchi.performance.0952.native-admission.v3"},
        "plan": {"path": str(reader.PLAN.resolve()), "sha256": reader.sha(reader.PLAN)},
        "binary": {"path": str(binary.resolve()), "sha256": reader.sha(binary), "build_id": build_id},
        "mmap_events": {"path": str(MMAP_EVENTS.resolve()), "sha256": reader.sha(MMAP_EVENTS),
                        "bound_pid": int(exec_map["pid"]), "executable_maps": len(executable),
                        "ambiguous_bound_mappings": mmap["ambiguous_bound_mappings"]},
        "controls": results,
        "temporary_inputs_removed": True,
    }
    if receipt_path.exists():
        raise RuntimeError(f"refusing to overwrite {receipt_path}")
    receipt_path.parent.mkdir(parents=True, exist_ok=True)
    with receipt_path.open("x", encoding="utf-8", newline="") as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write("\n")
        stream.flush()
        os.fsync(stream.fileno())
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--receipt", type=Path, required=True)
    args = parser.parse_args()
    result = run(args.receipt.resolve())
    print(json.dumps({"status": result["status"], "receipt": str(args.receipt.resolve())}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
