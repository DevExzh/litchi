#!/usr/bin/env python3
"""Exercise the 0920 capacity comparator without running performance children.

The input is one already-qualified 0917 report.  The controls deliberately
mutate an in-memory copy and retain only the small pass/fail receipt.  They
therefore test comparison policy without manufacturing performance evidence.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
RUNNER_PATH = HERE / "comparison.py"
PROTOCOL_PATH = HERE / "comparison-protocol.json"
VALIDATOR_PATH = HERE.parents[3] / "docs/performance/results/change-0917/capture.py"


def load_runner():
    spec = importlib.util.spec_from_file_location("litchi_perf_0920_controls_runner", RUNNER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load runner: {RUNNER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def report_from_stdout(path: Path) -> dict[str, Any]:
    for line in reversed([line.strip() for line in path.read_text(encoding="utf-8").splitlines()
                          if line.strip()]):
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    raise RuntimeError(f"no JSON report in {path}")


def spec_for(report: dict[str, Any]) -> dict[str, object]:
    route = {"zero": "zero", "default_8mib": "default", "release_8mib": "release"}[str(report["retention"])]
    return {"lane": "allocator", "case": report["case"], "edits": report["edits"],
            "retention": route, "label": "capacity-controls"}


def validate(module: Any, report: dict[str, Any], spec: dict[str, object]) -> None:
    module.validate_probe_report(report, spec, warmups=0, samples=1)


def inflate_capacities(report: dict[str, Any], delta: int) -> None:
    """Inflate only admitted capacity values, preserving the real report."""
    report["output_capacity"] += delta
    for route in report["route_outputs"]:
        route["output_capacity"] += delta
        if route["final_retained_before_commit"] is not None:
            route["final_retained_before_commit"] += delta
        for stage in route["stages"]:
            for key in ("before", "after", "after_release"):
                if stage[key] is not None:
                    stage[key] += delta
    for sample in report["samples"]:
        sample["output_capacity"] += delta
        for stage in sample["stages"]:
            for key in ("before", "after", "after_release"):
                if stage[key] is not None:
                    stage[key] += delta
        allocation = sample.get("allocation")
        if isinstance(allocation, dict):
            allocation["retained_bytes"] += delta
    top = report.get("allocation")
    if isinstance(top, dict):
        top["retained_bytes"] += delta


def valid_smaller_pair(source: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    baseline = copy.deepcopy(source)
    candidate = copy.deepcopy(source)
    inflate_capacities(baseline, 64)
    return baseline, candidate


def planned_population_size(module: Any, source: dict[str, Any], spec: dict[str, object]) -> int:
    """Exercise the compact summary against the planned 192/96 population."""
    native_path = Path(os.environ.get(
        "LITCHI_PERF_0920_NATIVE_CONTROL",
        str(Path("/tmp/litchi-performance-0920-baseline/qualification")
            / f"native-{source['case']}-e{int(source['edits']):02d}-default" / "stdout"),
    ))
    native = report_from_stdout(native_path)
    native_spec = {"lane": "native", "case": native["case"], "edits": native["edits"],
                   "retention": "default", "label": "capacity-population-native"}
    validate(module, native, native_spec)
    native_pair = copy.deepcopy(native)
    native_pair["samples"] = [copy.deepcopy(native["samples"][0]) for _ in range(50)]
    for index, sample in enumerate(native_pair["samples"]):
        sample["index"] = index
    native_summary = module.compare_reports(native_pair, native_pair, "population-native")
    allocator_summary = module.compare_reports(source, source, "population-allocator")
    rows = [native_summary for _ in range(96)] + [allocator_summary for _ in range(48)]
    encoded = json.dumps(rows, sort_keys=True, separators=(",", ":")).encode("utf-8")
    if len(encoded) >= 1 << 20:
        raise RuntimeError(f"planned capacity summary is at least 1 MiB: {len(encoded)}")
    return len(encoded)


def mutations(source: dict[str, Any]) -> dict[str, dict[str, Any]]:
    baseline, valid_candidate = valid_smaller_pair(source)
    result = {
        "output": copy.deepcopy(valid_candidate),
        "identity": copy.deepcopy(valid_candidate),
        "refusal": copy.deepcopy(valid_candidate),
        "ceiling": copy.deepcopy(valid_candidate),
        "retainedtransition": copy.deepcopy(valid_candidate),
        "mismatch": copy.deepcopy(valid_candidate),
        "capacity_growth": copy.deepcopy(valid_candidate),
    }
    result["output"]["output_sha256"] = "f" * 64
    result["identity"]["source_sha256"] = "e" * 64
    result["refusal"]["correctness"]["foreign_patch_refused"] = False
    result["ceiling"]["route_outputs"][1]["retention_ceiling_bytes"] = 1
    result["retainedtransition"]["route_outputs"][1]["stages"][1]["before"] = None
    result["mismatch"]["samples"][0]["allocation"]["retained_bytes"] += 1
    inflate_capacities(result["capacity_growth"], 128)
    # Keep the baseline available to make the intended relation explicit.
    result["_baseline"] = baseline
    return result


def run(source_path: Path, output: Path | None) -> dict[str, object]:
    module = load_runner()
    source = report_from_stdout(source_path)
    spec = spec_for(source)
    validate(module, source, spec)
    baseline, candidate = valid_smaller_pair(source)
    validate(module, baseline, spec)
    validate(module, candidate, spec)
    positive = module.compare_reports(baseline, candidate, "positive smaller capacity")
    rows: list[dict[str, object]] = []
    validator_error = module.validator_module(module.VALIDATOR_PATH).EvidenceError
    for name, mutated in mutations(source).items():
        if name == "_baseline":
            continue
        rejected = False
        error = None
        try:
            validate(module, mutated, spec)
            module.compare_reports(baseline, mutated, name)
        except (module.EvidenceError, validator_error) as exc:  # Each control must fail closed.
            rejected = True
            error = f"{type(exc).__name__}: {exc}"
        if not rejected:
            raise RuntimeError(f"control unexpectedly passed: {name}")
        rows.append({"name": name, "status": "rejected", "error": error})
    result: dict[str, object] = {
        "schema": "litchi.performance.0920.capacity-controls.v1",
        "runner": {"bytes": RUNNER_PATH.stat().st_size,
                   "sha256": hashlib.sha256(RUNNER_PATH.read_bytes()).hexdigest()},
        "protocol": {"bytes": PROTOCOL_PATH.stat().st_size,
                      "sha256": hashlib.sha256(PROTOCOL_PATH.read_bytes()).hexdigest()},
        "controls": {"bytes": Path(__file__).stat().st_size,
                      "sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest()},
        "validator": {"bytes": VALIDATOR_PATH.stat().st_size,
                      "sha256": hashlib.sha256(VALIDATOR_PATH.read_bytes()).hexdigest()},
        "source_report": {"path": str(source_path.resolve()), "bytes": source_path.stat().st_size,
                          "sha256": hashlib.sha256(source_path.read_bytes()).hexdigest()},
        "positive_smaller_capacity": {"status": "accepted", "capacity": positive},
        "negative_controls": rows,
        "planned_population_capacity_summary_bytes": planned_population_size(module, source, spec),
        "performance_children_run": 0,
    }
    if output is not None:
        data = (json.dumps(result, indent=2, sort_keys=True) + "\n").encode("utf-8")
        if len(data) >= 1 << 20:
            raise RuntimeError("control receipt is at least 1 MiB")
        if output.exists():
            raise RuntimeError(f"refusing to replace retained output: {output}")
        output.write_bytes(data)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        print(json.dumps(run(args.report.resolve(), args.output.resolve() if args.output else None),
                         sort_keys=True))
    except Exception as error:
        print(f"0920 controls failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
