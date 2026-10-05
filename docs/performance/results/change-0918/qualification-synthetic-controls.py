#!/usr/bin/env python3
"""Run retained in-memory controls for every 0918 reduction path.

This control does not execute a probe and does not write capture artifacts.  It
monkeypatches only the frozen comparison runner's receipt loader with a
deterministic 192-native/96-allocator synthetic schedule, then exercises the
passing reduction and the latency, allocation, peak-live, missing-pair, and
semantic-corruption negative controls.  The report is a small review witness;
the real qualification custody remains in the external root.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
RUNNER_PATH = HERE / "comparison.py"
PROTOCOL_PATH = HERE / "comparison-protocol.json"
DEFAULT_OUTPUT = HERE / "qualification-synthetic-controls.json"


class ControlError(RuntimeError):
    pass


def fail(message: str) -> None:
    raise ControlError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, object]:
    path = path.resolve()
    return {"path": str(path.relative_to(HERE.parents[3])),
            "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_runner():
    spec = importlib.util.spec_from_file_location("litchi_perf_0918_synthetic_runner",
                                                  RUNNER_PATH)
    require(spec is not None and spec.loader is not None,
            f"cannot load comparison runner: {RUNNER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def stable_report(spec: dict[str, object], *, latency_factor: float = 1.0,
                  allocation_factor: float = 1.0, peak_factor: float = 1.0,
                  semantic_marker: str = "stable") -> dict[str, object]:
    if spec["lane"] == "native":
        # Fifty deterministic samples mirror the frozen native lane's shape.
        base = 1_000_000 + int(spec["round"]) * 1_000 + int(spec["edits"]) * 10
        effective_factor = latency_factor if spec["version"] == "candidate" else 1.0
        values = [int((base + index * 13) * effective_factor) for index in range(50)]
        return {
            "elapsed_ns": values,
            "samples": [{"elapsed_ns": value} for value in values],
            "synthetic_semantic_marker": semantic_marker,
        }
    base = 1_000_000
    effective_allocation_factor = allocation_factor if spec["version"] == "candidate" else 1.0
    if spec["version"] == "candidate" and spec["case"] == "docnohf" \
            and spec["edits"] == 32 and spec["retention"] == "default":
        # Satisfy the positive admission target in the passing control.
        base = 1_000_000
        effective_allocation_factor *= 0.90
    allocated = int(base * effective_allocation_factor)
    effective_peak_factor = peak_factor if spec["version"] == "candidate" else 1.0
    peak = int(100_000 * effective_peak_factor)
    allocation = {
        "allocated_bytes": allocated,
        "deallocated_bytes": allocated,
        "allocation_calls": 100,
        "peak_live_bytes": peak,
        "retained_bytes": 10_000,
    }
    return {
        "samples": [{"allocation": copy.deepcopy(allocation)}],
        "allocation": allocation,
        "synthetic_semantic_marker": semantic_marker,
    }


def synthetic_items(runner: Any, *, latency_factor: float = 1.0,
                    allocation_factor: float = 1.0, peak_factor: float = 1.0,
                    corrupt_key: tuple[Any, ...] | None = None,
                    missing_label: str | None = None) -> dict[str, dict[str, Any]]:
    protocol = runner.protocol()
    items: dict[str, dict[str, Any]] = {}
    for lane in runner.LANES:
        for spec in runner.schedule(protocol, lane):
            report = stable_report(spec, latency_factor=latency_factor,
                                   allocation_factor=allocation_factor,
                                   peak_factor=peak_factor,
                                   semantic_marker=("corrupt" if corrupt_key == (
                                       spec["lane"], spec.get("round", spec.get("repetition")),
                                       spec["case"], spec["edits"], spec["retention"],
                                       spec["version"]) else "stable"))
            if missing_label == spec["label"]:
                continue
            items[str(spec["label"])] = {
                "spec": spec,
                "receipt": {"process_metrics": {"max_rss_kib": 100}},
                "report": report,
            }
    return items


def run_reduction(runner: Any, items: dict[str, dict[str, Any]]) -> dict[str, object]:
    original = runner.read_receipt
    original_manifest_descriptor = runner.manifest_descriptor

    def load(_manifest: dict[str, Any], spec: dict[str, object],
             _schedule_rows: list[dict[str, object]]) -> dict[str, Any]:
        label = str(spec["label"])
        if label not in items:
            raise runner.EvidenceError(f"synthetic receipt missing: {label}")
        return items[label]

    runner.read_receipt = load
    # The reducer returns manifest provenance, but synthetic controls do not
    # have a frozen capture manifest.  Keep that provenance explicit without
    # creating a fake packet file.
    runner.manifest_descriptor = lambda: {
        "path": "<synthetic-manifest>", "bytes": 0,
        "sha256": "0" * 64,
    }
    try:
        return runner.reduce_data({"synthetic": True})
    finally:
        runner.read_receipt = original
        runner.manifest_descriptor = original_manifest_descriptor


def summarize_result(result: dict[str, object]) -> dict[str, object]:
    native = result["native"]
    allocator = result["allocator"]
    gates = result["gates"]
    require(native["children"] == 192 and native["paired_cells"] == 96
            and native["scenario_cells"] == 8,
            "native reduction did not traverse the full 192-child schedule")
    require(allocator["children"] == 96 and len(allocator["cells"]) == 48,
            "allocator reduction did not traverse the full 96-child schedule")
    return {
        "native_children": native["children"],
        "native_paired_cells": native["paired_cells"],
        "native_scenario_cells": native["scenario_cells"],
        "allocator_children": allocator["children"],
        "allocator_cells": len(allocator["cells"]),
        "gates": {
            "latency": gates["all_native_default_latency_lower_bounds"],
            "allocation": gates["no_allocator_allocated_bytes_increase_over_1_percent"],
            "peak_live": gates["no_allocator_peak_live_increase_over_5_percent"],
            "adoption": gates["adoption"],
        },
    }


def expect_failure(runner: Any, items: dict[str, dict[str, Any]], needle: str) -> str:
    try:
        run_reduction(runner, items)
    except Exception as error:
        text = f"{type(error).__name__}: {error}"
        require(needle in text, f"negative control failed with unexpected error: {text}")
        return text
    fail(f"negative control unexpectedly passed: {needle}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    try:
        runner = load_runner()
        passing = summarize_result(run_reduction(runner, synthetic_items(runner)))
        require(passing["gates"] == {
            "latency": True, "allocation": True, "peak_live": True, "adoption": True,
        }, "passing synthetic control did not satisfy all gates")

        latency = summarize_result(run_reduction(
            runner, synthetic_items(runner, latency_factor=1.20)))
        require(latency["gates"]["latency"] is False and latency["gates"]["adoption"] is False,
                "latency negative control did not reject adoption")

        allocation = summarize_result(run_reduction(
            runner, synthetic_items(runner, allocation_factor=1.02)))
        require(allocation["gates"]["allocation"] is False
                and allocation["gates"]["adoption"] is False,
                "allocation negative control did not reject adoption")

        peak_live = summarize_result(run_reduction(
            runner, synthetic_items(runner, peak_factor=1.06)))
        require(peak_live["gates"]["peak_live"] is False
                and peak_live["gates"]["adoption"] is False,
                "peak-live negative control did not reject adoption")

        protocol = runner.protocol()
        first_native = runner.schedule(protocol, "native")[0]
        missing_items = synthetic_items(runner, missing_label=str(first_native["label"]))
        missing_error = expect_failure(runner, missing_items, "synthetic receipt missing")

        corruption_key = ("native", first_native["round"], first_native["case"],
                          first_native["edits"], first_native["retention"], "candidate")
        corrupt_items = synthetic_items(runner, corrupt_key=corruption_key)
        corrupt_error = expect_failure(runner, corrupt_items, "native semantic identity differs")

        report = {
            "schema": "litchi.performance.0918.qualification-synthetic-controls.v1",
            "runner": descriptor(RUNNER_PATH),
            "protocol": descriptor(PROTOCOL_PATH),
            "command": "PYTHONDONTWRITEBYTECODE=1 python3 -B "
                       "docs/performance/results/change-0918/qualification-synthetic-controls.py",
            "execution": "in-memory synthetic receipts; no binaries, probes, builds, or real raw captures",
            "controls": {
                "passing_full_schedule": passing,
                "latency_gate_failure": {**latency, "expected_adoption": False},
                "allocation_gate_failure": {**allocation, "expected_adoption": False},
                "peak_live_gate_failure": {**peak_live, "expected_adoption": False},
                "missing_pair_failure": {"expected_failure": missing_error},
                "semantic_corruption_failure": {"expected_failure": corrupt_error},
            },
            "coverage": {
                "native_children": 192, "native_paired_cells": 96,
                "native_scenario_cells": 8, "allocator_children": 96,
                "allocator_cells": 48,
            },
        }
        output = args.output.resolve()
        require(not output.exists(), f"refusing to replace retained control report: {output}")
        data = (json.dumps(report, indent=2, sort_keys=True) + "\n").encode()
        require(len(data) < (1 << 20), "synthetic control report is at least 1 MiB")
        output.write_bytes(data)
        print(json.dumps({"status": "verified", "report": str(output),
                          "sha256": sha256(output)}, sort_keys=True))
    except (ControlError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"0918 synthetic controls failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
