#!/usr/bin/env python3
"""Run 0923 report and reducer controls without native benchmark children."""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
RUNNER_PATH = HERE / "comparison.py"
PROTOCOL_PATH = HERE / "comparison-protocol.json"


def load_runner() -> Any:
    spec = importlib.util.spec_from_file_location("litchi_perf_0923_comparison", RUNNER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load comparison runner: {RUNNER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def read_report(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError(f"report is not an object: {path}")
    return value


def route(report: dict[str, Any]) -> str:
    try:
        return str(report["results"][0]["source"]["docx_section_layout"]["ingress_route"])
    except (KeyError, IndexError, TypeError) as error:
        raise RuntimeError(f"report route is unavailable: {error}") from error


def mutate(report: dict[str, Any], name: str) -> dict[str, Any]:
    value = copy.deepcopy(report)
    result = value["results"][0]
    section = result["source"]["docx_section_layout"]
    source = result["source"]
    if name == "route_identity":
        section["ingress_route"] = "owned_vec" if section["ingress_route"] == "generic_read_at" else "generic_read_at"
    elif name == "phase_sum":
        section["query_ns"][0] += 1
    elif name == "owned_positional_counter":
        source["read_calls"] = [1] * len(result["elapsed_ns"]["samples"])
    elif name == "owned_ingress_cardinality":
        if section["ingress_route"] == "generic_read_at":
            section["source_read_calls"].pop()
        else:
            section["owned_ingress_owned_bytes"].pop()
    elif name == "gate":
        section["patch_forward_inverse_verified"] = False
    elif name == "output":
        section["output_sha256"][0] = "f" * 64
    elif name == "sink":
        result["sink"]["accepted_bytes"] += 1
    elif name == "source":
        section["source_sha256"] = "e" * 64
    else:
        raise ValueError(f"unknown mutation {name}")
    return value


def expect_rejection(module: Any, report: dict[str, Any], expected_route: str, name: str,
                     *, warmups: int, samples: int) -> dict[str, str]:
    try:
        module.validate_report(report, expected_route, warmups=warmups, samples=samples, label=name)
    except module.EvidenceError as error:
        return {"name": name, "status": "rejected", "error": f"{type(error).__name__}: {error}"}
    raise RuntimeError(f"negative control unexpectedly passed: {name}")


def run(report_paths: list[Path], output: Path | None) -> dict[str, Any]:
    module = load_runner()
    if not report_paths:
        raise RuntimeError("at least one --report is required")
    protocol = json.loads(PROTOCOL_PATH.read_text(encoding="utf-8"))
    native = protocol.get("native", {})
    expected_schedule = {
        "routes": ["generic_read_at", "owned_vec", "owned_reader"],
        "pairs": 8,
        "children_per_route": 16,
        "children_per_version": 24,
        "children_total": 48,
    }
    actual_schedule = {key: native.get(key) for key in expected_schedule}
    if actual_schedule != expected_schedule:
        raise RuntimeError(f"prospective native schedule changed: {actual_schedule}")
    reports = [(path.resolve(), read_report(path.resolve())) for path in report_paths]
    validated: list[dict[str, Any]] = []
    for path, report in reports:
        current_route = route(report)
        samples = int(report["configuration"]["samples_per_case"])
        warmups = int(report["configuration"]["warmup_iterations_per_case"])
        validated.append(module.validate_report(report, current_route, warmups=warmups,
                                                samples=samples, label=current_route))

    controls: list[dict[str, str]] = []
    source_path, source_report = reports[0]
    source_route = route(source_report)
    samples = int(source_report["configuration"]["samples_per_case"])
    warmups = int(source_report["configuration"]["warmup_iterations_per_case"])
    for name in ("route_identity", "phase_sum", "owned_positional_counter", "owned_ingress_cardinality",
                 "gate", "output", "sink", "source"):
        controls.append(expect_rejection(module, mutate(source_report, name), source_route, name,
                                         warmups=warmups, samples=samples))

    # Exercise both the generic and owned-specific fail-closed branches when
    # their reports are present in the same receipt.
    route_names = {route(report) for _, report in reports}
    if "generic_read_at" in route_names:
        generic = next(report for _, report in reports if route(report) == "generic_read_at")
        mutated = copy.deepcopy(generic)
        mutated["results"][0]["source"]["read_calls"][0] = 0
        controls.append(expect_rejection(module, mutated, "generic_read_at", "generic_zero_read",
                                          warmups=int(generic["configuration"]["warmup_iterations_per_case"]),
                                          samples=int(generic["configuration"]["samples_per_case"])))
    if "owned_vec" in route_names:
        owned = next(report for _, report in reports if route(report) == "owned_vec")
        mutated = copy.deepcopy(owned)
        mutated["results"][0]["source"]["docx_section_layout"]["owned_ingress_read_calls"][0] = 1
        controls.append(expect_rejection(module, mutated, "owned_vec", "owned_vec_fabricated_reader",
                                          warmups=int(owned["configuration"]["warmup_iterations_per_case"]),
                                          samples=int(owned["configuration"]["samples_per_case"])))

    # A same-report pair proves the comparator can consume a valid route while
    # leaving the timing gate unset.  This is a policy smoke check only.
    positive = []
    for path, report in reports:
        positive.append(module.compare_reports(report, report, f"positive-{route(report)}",
                                                route=route(report), gate=None,
                                                warmups=int(report["configuration"]["warmup_iterations_per_case"]),
                                                samples=int(report["configuration"]["samples_per_case"])))
    planned_bytes, planned_sha256 = module.planned_reduction_bytes()
    if planned_bytes >= 1 << 20:
        raise RuntimeError(f"planned compact reduction is at least 1 MiB: {planned_bytes}")

    def descriptor(path: Path) -> dict[str, Any]:
        data = path.read_bytes()
        return {"path": str(path), "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}

    receipt: dict[str, Any] = {
        "schema": "litchi.performance.0923.comparison-controls.v1",
        "runner": descriptor(RUNNER_PATH),
        "protocol": descriptor(PROTOCOL_PATH),
        "controls": descriptor(Path(__file__).resolve()),
        "reports": [
            {"route": route(report), "descriptor": descriptor(path), "validation": summary}
            for (path, report), summary in zip(reports, validated)
        ],
        "positive_same_report_pairs": positive,
        "negative_controls": controls,
        "planned_reduction": {
            "native_rows": 48,
            "allocator_rows": 0,
            "rounds_per_row": 1,
            "bytes": planned_bytes,
            "sha256": planned_sha256,
            "under_one_mib": True,
        },
        "schedule_assertion": {
            "status": "passed",
            "routes": expected_schedule["routes"],
            "pairs": expected_schedule["pairs"],
            "children_per_route": expected_schedule["children_per_route"],
            "children_per_version": expected_schedule["children_per_version"],
            "children_total": expected_schedule["children_total"],
        },
        "performance_children_run": 0,
    }
    data = json.dumps(receipt, indent=2, sort_keys=True).encode("utf-8") + b"\n"
    if len(data) >= 1 << 20:
        raise RuntimeError("control receipt is at least 1 MiB")
    if output is not None:
        if output.exists():
            raise RuntimeError(f"refusing to replace retained output: {output}")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(data)
    return receipt


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", action="append", type=Path, required=True,
                        help="one or more external harness reports")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        value = run(args.report, args.output)
        print(json.dumps(value, sort_keys=True))
    except Exception as error:
        print(f"0923 comparison controls failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
