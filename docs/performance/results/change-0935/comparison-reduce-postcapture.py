#!/usr/bin/env python3
"""Recompute the post-capture allocation gate from the immutable 0935 packet.

The frozen runner's descriptive field name inherited a ``ge_20_percent`` label
from 0933.  This offline supplement replays the packet through that frozen
runner and adds the 0935 manifest threshold without changing the packet or the
frozen runner/reduction output.
"""
from __future__ import annotations
import argparse
import copy
import importlib.util
import json
from pathlib import Path


def load_runner(path: Path):
    spec = importlib.util.spec_from_file_location("litchi_perf_0935_frozen_runner", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load frozen runner: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--runner", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    runner = load_runner(args.runner)
    packet = args.packet.resolve()
    summary = runner.replay_packet(packet)
    manifest = json.loads((packet / "manifest.json").read_text(encoding="utf-8"))
    threshold = float(manifest["formal_gate"]["allocator_min_requested_bytes_reduction_fraction"])
    corrected = copy.deepcopy(summary)
    for cell in corrected["allocator"]["cells"].values():
        review = cell["gate_review"]
        fractions = review["requested_bytes_reduction_fractions"]
        review["frozen_runner_field_note"] = (
            "requested_bytes_reduction_all_ge_20_percent is a stale descriptive field "
            "from the frozen reducer; the 0935 manifest gate is evaluated below."
        )
        review["0935_requested_bytes_reduction_threshold"] = threshold
        review["requested_bytes_reduction_all_ge_0935_threshold"] = all(
            value is not None and value >= threshold for value in fractions
        )
    corrected["schema"] = "litchi.performance.0935.formal-replay-corrected.v1"
    corrected["frozen_replay_schema"] = summary["schema"]
    corrected["correction_scope"] = "offline gate-label correction; packet, runner, and frozen reduction are unchanged"
    args.output.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(corrected, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if len(data) >= 1 << 20:
        raise RuntimeError(f"corrected reduction exceeds 1 MiB: {len(data)}")
    with args.output.open("xb") as stream:
        stream.write(data)
    print(json.dumps({"status": corrected["status"], "output": str(args.output), "bytes": len(data)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
