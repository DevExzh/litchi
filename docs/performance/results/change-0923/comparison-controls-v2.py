#!/usr/bin/env python3
"""Fail-closed controls for the prospective 0923 v2 comparison tooling.

This receipt validator uses retained baseline reports only.  It never starts a
benchmark child and keeps all report inputs outside the repository.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
RUNNER_PATH = HERE / "comparison-v2.py"
PROTOCOL_PATH = HERE / "comparison-protocol-v2.json"


def load_runner() -> Any:
    spec = importlib.util.spec_from_file_location("litchi_perf_0923_comparison_v2", RUNNER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load runner: {RUNNER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def descriptor(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    return {"path": str(path.resolve()), "bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest()}


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


def result_parts(report: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    result = report["results"][0]
    source = result["source"]
    section = source["docx_section_layout"]
    return result, source, section


def mutate(report: dict[str, Any], name: str) -> dict[str, Any]:
    value = copy.deepcopy(report)
    result, source, section = result_parts(value)
    samples = len(result["elapsed_ns"]["samples"])
    if name == "generic_read_calls":
        source["read_calls"][0] += 1
    elif name == "generic_read_bytes":
        source["read_bytes"][0] += 1
    elif name == "generic_ordinary_calls":
        source["ordinary_payload_read_calls"][0] += 1
    elif name == "generic_ordinary_bytes":
        source["ordinary_payload_read_bytes"][0] += 1
    elif name == "generic_max_in_flight":
        source["max_in_flight_reads"][0] += 1
    elif name == "generic_materializations":
        source["ordinary_payload_materializations"][0] += 1
    elif name == "generic_cache_hits":
        section["source_cache_before_publication_after_commit_hits"][0] += 1
    elif name == "generic_cache_cold_loads":
        section["source_cache_before_publication_after_commit_cold_loads"][0] += 1
    elif name == "generic_cache_successful_loads":
        section["source_cache_before_publication_after_commit_successful_loads"][0] += 1
    elif name == "generic_cache_retained_entries":
        section["source_cache_before_publication_after_commit_retained_entries"][0] += 1
    elif name == "generic_cache_retained_bytes":
        section["source_cache_before_publication_after_commit_retained_bytes"][0] += 1
    elif name == "owned_reader_calls":
        section["owned_ingress_read_calls"][0] += 1
    elif name == "owned_reader_bytes":
        section["owned_ingress_read_bytes"][0] += 1
    elif name == "owned_vec_calls":
        section["owned_ingress_read_calls"][0] = 1
    elif name == "owned_vec_bytes":
        section["owned_ingress_read_bytes"][0] = 1
    elif name == "owned_storage_bytes":
        section["owned_ingress_owned_bytes"][0] += 1
    elif name == "binary_sha":
        value["binary_identity"]["binary_sha256"] = "0" * 64
    elif name == "binary_bytes":
        value["binary_identity"]["binary_bytes"] += 1
    elif name == "sample_count":
        value["configuration"]["samples_per_case"] = samples - 1
        result["elapsed_ns"]["samples"].pop()
    elif name == "warmup_count":
        value["configuration"]["warmup_iterations_per_case"] += 1
    elif name == "sample_order":
        result["elapsed_ns"]["sample_order"][0] = result["elapsed_ns"]["sample_order"][1]
        section["sample_order"] = result["elapsed_ns"]["sample_order"]
    elif name == "route_identity":
        section["ingress_route"] = "owned_vec" if section["ingress_route"] == "generic_read_at" else "generic_read_at"
    else:
        raise ValueError(f"unknown mutation {name}")
    return value


def expect_rejection(module: Any, report: dict[str, Any], expected_route: str, name: str,
                     binary: dict[str, Any], *, warmups: int, samples: int) -> dict[str, str]:
    try:
        module.validate_report(report, expected_route, warmups=warmups, samples=samples,
                               label=name, binary=binary)
    except module.EvidenceError as error:
        return {"name": name, "status": "rejected", "error": f"{type(error).__name__}: {error}"}
    raise RuntimeError(f"negative control unexpectedly passed: {name}")


def run(report_paths: list[Path], binary_path: Path, output: Path | None) -> dict[str, Any]:
    module = load_runner()
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
    if protocol.get("allocation_evidence", {}).get("measured") is not False:
        raise RuntimeError("v2 protocol must explicitly say allocator evidence is unmeasured")
    if protocol.get("allocation_evidence", {}).get("lane") is not None:
        raise RuntimeError("v2 native protocol must not claim an allocator lane")
    if len(report_paths) != 3:
        raise RuntimeError("expected exactly one retained baseline report for each route")
    binary = descriptor(binary_path.resolve())
    reports = [(path.resolve(), read_report(path.resolve())) for path in report_paths]
    by_route = {route(report): (path, report) for path, report in reports}
    if set(by_route) != set(expected_schedule["routes"]):
        raise RuntimeError(f"baseline route set changed: {sorted(by_route)}")
    validated: list[dict[str, Any]] = []
    for current_route in expected_schedule["routes"]:
        path, report = by_route[current_route]
        samples = int(report["configuration"]["samples_per_case"])
        warmups = int(report["configuration"]["warmup_iterations_per_case"])
        validated.append(module.validate_report(report, current_route, warmups=warmups,
                                                samples=samples, label=current_route,
                                                binary=binary))

    controls: list[dict[str, str]] = []
    generic = by_route["generic_read_at"][1]
    generic_samples = int(generic["configuration"]["samples_per_case"])
    generic_warmups = int(generic["configuration"]["warmup_iterations_per_case"])
    for name in (
        "generic_read_calls", "generic_read_bytes", "generic_ordinary_calls",
        "generic_ordinary_bytes", "generic_max_in_flight", "generic_materializations",
        "generic_cache_hits", "generic_cache_cold_loads", "generic_cache_successful_loads",
        "generic_cache_retained_entries", "generic_cache_retained_bytes",
    ):
        controls.append(expect_rejection(module, mutate(generic, name), "generic_read_at", name,
                                         binary, warmups=generic_warmups, samples=generic_samples))
    for route_name, names in (("owned_vec", ("owned_vec_calls", "owned_vec_bytes", "owned_storage_bytes")),
                              ("owned_reader", ("owned_reader_calls", "owned_reader_bytes", "owned_storage_bytes"))):
        report = by_route[route_name][1]
        samples = int(report["configuration"]["samples_per_case"])
        warmups = int(report["configuration"]["warmup_iterations_per_case"])
        for name in names:
            controls.append(expect_rejection(module, mutate(report, name), route_name, name, binary,
                                             warmups=warmups, samples=samples))
    for name in ("binary_sha", "binary_bytes", "sample_count", "warmup_count", "sample_order", "route_identity"):
        controls.append(expect_rejection(module, mutate(generic, name), "generic_read_at", name,
                                         binary, warmups=generic_warmups, samples=generic_samples))

    positives = []
    for current_route in expected_schedule["routes"]:
        report = by_route[current_route][1]
        samples = int(report["configuration"]["samples_per_case"])
        warmups = int(report["configuration"]["warmup_iterations_per_case"])
        positives.append(module.compare_reports(report, report, f"positive-{current_route}",
                                                route=current_route, gate=None,
                                                warmups=warmups, samples=samples,
                                                baseline_binary=binary, candidate_binary=binary))
    planned_bytes, planned_sha256 = module.planned_reduction_bytes(native_rows=48, allocator_rows=0, rounds=24)
    if planned_bytes >= 1 << 20:
        raise RuntimeError(f"full planned compact reduction is at least 1 MiB: {planned_bytes}")
    receipt: dict[str, Any] = {
        "schema": "litchi.performance.0923.comparison-controls.v2",
        "runner": descriptor(RUNNER_PATH),
        "protocol": descriptor(PROTOCOL_PATH),
        "controls": descriptor(Path(__file__).resolve()),
        "binary": binary,
        "reports": [
            {"route": current_route, "descriptor": descriptor(by_route[current_route][0]),
             "validation": validated[index]}
            for index, current_route in enumerate(expected_schedule["routes"])
        ],
        "positive_same_report_pairs": positives,
        "negative_controls": controls,
        "planned_reduction": {
            "native_rows": 48, "allocator_rows": 0, "rounds_per_row": 24,
            "bytes": planned_bytes, "sha256": planned_sha256, "under_one_mib": True,
            "meaning": "full 48-child native schedule with compact phase/counter/cache/sink/gate rounds; no raw reports embedded",
        },
        "schedule_assertion": {"status": "passed", **expected_schedule},
        "performance_children_run": 0,
    }
    data = json.dumps(receipt, indent=2, sort_keys=True).encode("utf-8") + b"\n"
    if len(data) >= 1 << 20:
        raise RuntimeError("control receipt is at least 1 MiB")
    if output is not None:
        output = output.resolve()
        if output.exists():
            raise RuntimeError(f"refusing to replace retained output: {output}")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(data)
    return receipt


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", action="append", type=Path, required=True)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        value = run(args.report, args.binary, args.output)
        print(json.dumps(value, sort_keys=True))
    except Exception as error:
        print(f"0923 v2 comparison controls failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
