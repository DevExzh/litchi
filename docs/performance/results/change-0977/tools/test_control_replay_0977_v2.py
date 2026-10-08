#!/usr/bin/env python3
"""Replay the retained harmless-control packet through the full v2 path.

This is an offline integration control.  It hard-links the retained perf.data
and adapts only the older 0976 receipt filenames to the 0977 child filenames;
the command argv and target receipt are copied from the retained receipts
without editing.  The decoder's own current-perf evlist/header/report/script
subcommands are exercised; no capture, workload, or build is started.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import os
from pathlib import Path


ROOT = Path("/home/zhuhe/litchi-performance-0977")
SOURCE = Path("/home/zhuhe/litchi-performance-0976/control/affinity-control-v6")
OUT = ROOT / "control-replay-v2d"
CHILD = OUT / "child-01"
PERF = Path("/usr/bin/perf")
PYTHON = Path("/usr/bin/python3.14")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, object]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_decoder():
    path = ROOT / "tools/focuseddecode_0977_v2.py"
    spec = importlib.util.spec_from_file_location("focuseddecode_0977_control_replay", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot import v2 decoder")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def stage() -> tuple[dict[str, object], dict[str, object], dict[str, object]]:
    if OUT.exists():
        raise RuntimeError(f"refusing to reuse replay output: {OUT}")
    CHILD.mkdir(parents=True)
    # Hard-link the immutable retained capture; this does not duplicate or
    # mutate the raw evidence.
    os.link(SOURCE / "control.perf.data", CHILD / "child-01.perf.data")
    command = json.loads((SOURCE / "command-receipt.json").read_text(encoding="utf-8"))
    terminal = json.loads((SOURCE / "terminal-receipt.json").read_text(encoding="utf-8"))
    target = json.loads((SOURCE / "target-receipt.json").read_text(encoding="utf-8"))
    # 0976 stored command separately; 0977's runner records it in the
    # terminal receipt.  The value is copied byte-for-byte from command.argv.
    terminal["command"] = command["argv"]
    terminal["native_ld_preload"] = "absent"
    (CHILD / "child-01.command.json").write_text(json.dumps(command, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    (CHILD / "child-01.terminal.json").write_text(json.dumps(terminal, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    (CHILD / "child-01.control-target.json").write_text(json.dumps(target, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    plan = {
        "python": {"path": str(PYTHON), "sha256": sha256(PYTHON)},
        "perf": {"path": str(PERF), "sha256": sha256(PERF), "period": 1_000_000, "stack_cap": 32_768},
    }
    (OUT / "plan.json").write_text(json.dumps(plan, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return plan, terminal, target


def main() -> int:
    plan, terminal, target = stage()
    decoder = load_decoder()
    packet = decoder.validate_control_child(plan, CHILD, terminal, target)
    summary = packet.get("summary")
    if packet.get("schema") != "litchi.performance.0977.control_packet.v1" or packet.get("status") != "complete":
        raise AssertionError("control packet status/schema failed")
    if packet.get("control_only") is not True or packet.get("native_workload_started") is not False:
        raise AssertionError("control packet native/workload provenance failed")
    if not isinstance(summary, dict) or summary.get("sample_count", 0) <= 0:
        raise AssertionError("control packet has no raw samples")
    if summary.get("raw_lost_sum") != 0 or summary.get("terminal_lost_samples") != 0:
        raise AssertionError("control loss gates are nonzero")
    if summary.get("cpu_gate") is not True or summary.get("task_gate") is not True:
        raise AssertionError("control CPU/task gates failed")
    if packet.get("evlist_contract", {}).get("status") != "pass":
        raise AssertionError("control evlist authority did not pass")
    if packet.get("header_contract", {}).get("status") != "unavailable":
        raise AssertionError("current header warning was not retained descriptively")
    if packet.get("header_metadata", {}).get("known_invalid_header_event_desc_warning") is not True:
        raise AssertionError("current header warning evidence is missing")
    if "terminal_report_stderr" not in packet.get("artifacts", {}):
        raise AssertionError("terminal report stderr custody is missing")
    output = OUT / "control-packet.json"
    output.write_text(json.dumps(packet, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({
        "status": "pass",
        "native_workload_started": False,
        "sample_count": summary["sample_count"],
        "raw_lost_sum": summary["raw_lost_sum"],
        "terminal_lost_samples": summary["terminal_lost_samples"],
        "cpu_gate": summary["cpu_gate"],
        "task_gate": summary["task_gate"],
        "header_warning": True,
        "evlist": "pass",
        "decoder": binding(ROOT / "tools/focuseddecode_0977_v2.py"),
        "retained_raw": binding(SOURCE / "control.perf.data"),
        "output": binding(output),
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
