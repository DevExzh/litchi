#!/usr/bin/env python3
"""Verify semantic/report qualification without treating debug timings as a baseline."""
import argparse
import hashlib
import json
from pathlib import Path

if not __debug__:
    raise RuntimeError("Qualification checks require Python without -O")

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--directory", type=Path, default=Path(__file__).resolve().parent / "qualification")
parser.add_argument("--output", type=Path)
args = parser.parse_args()
case = "xlsx_producer_medium_source_full_cell_scan"
archive = "55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8"
digest = "8d29b40718aa9d5b08c746f841195bba17ab8e01608b12e6093daf160c00764b"
reports = {}
sidecars = []
corpora = []
for mode in ("native", "allocator"):
    path = args.directory / (mode + ".json")
    report = json.loads(path.read_text())
    tool = report["tool"]
    assert tool["profile"] == "debug" and tool["version"] == "0.1.0"
    assert tool["binary"] == ("litchi-perf-baseline" if mode == "native" else "litchi-perf-baseline-alloc")
    assert tool["instrumentation"] == ("none" if mode == "native" else "system_allocator_operation_scoped")
    assert report["configuration"]["cases"] == [case]
    assert report["configuration"]["samples_per_case"] == 3
    assert report["configuration"]["warmup_iterations_per_case"] == 1
    assert len(report["results"]) == 1
    row = report["results"][0]
    assert row["case"] == case
    corpus = row["corpus"]
    assert corpus["archive_sha256"] == archive and corpus["archive_bytes"] == 22739
    assert corpus["entry_count"] == 4096 and corpus["archive_member_count"] == 12
    corpora.append(corpus)
    elapsed = row["elapsed_ns"]
    order = elapsed["sample_order"]
    assert sorted(order) == [0, 1, 2]
    assert len(elapsed["samples"]) == 3 and all(type(x) is int and x > 0 for x in elapsed["samples"])
    scan = row["source"]["producer_full_cell_scan"]
    assert scan["sample_indices"] == order
    for key, value in (("cell", 4096), ("numeric", 3686), ("text", 410)):
        assert scan["expected_" + key + "_count"] == value
        assert scan["observed_" + key + "_counts"] == [value] * 3
    assert scan["oracle_digest_sha256"] == digest
    assert scan["observed_digest_sha256"] == [digest] * 3
    assert scan["archive_sha256"] == archive and scan["archive_bytes"] == 22739
    assert scan["source_version_stable"] == [True] * 3
    assert scan["physical_cold_eligible"] is False and scan["runtime_route_observed"] is False
    routes = scan["worksheet_route_eligibility"]
    assert [r["worksheet"] for r in routes] == [f"xl/worksheets/sheet{i}.xml" for i in range(1, 5)]
    assert all(r["source_stream_eligible"] is False for r in routes)
    metrics = row["operation_metrics"]
    assert metrics["sample_count"] == 3 and metrics["sample_indices"] == order
    if mode == "native":
        assert "allocation" not in metrics
    else:
        assert tool["allocator_counter_revision"] == "serialized_region_peak_v3"
        allocation = metrics["allocation"]
        assert allocation["status"] == "measured" and allocation["scope"] == "operation_global_system_allocator"
        for key, value in allocation.items():
            if isinstance(value, dict):
                assert value["status"] == "measured" and value["scope"] == "operation_global_system_allocator"
                assert len(value["values"]) == 3 and all(type(x) is int and x >= 0 for x in value["values"])
        assert allocation["failed_allocation_calls"]["values"] == [0] * 3
        assert all(x > 0 for key in ("allocation_calls", "allocated_bytes", "region_peak_live_bytes") for x in allocation[key]["values"])
    sidecars.append(json.loads((args.directory / (mode + "-producer.json")).read_text()))
    reports[mode] = {"path": str(path), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
assert sidecars[0] == sidecars[1] and corpora[0] == corpora[1]
result = {"status": "pass", "profile": "debug", "measured_samples": 6, "warmups": 2, "cells_per_scan": 4096, "reports": reports, "performance_claim": None}
if args.output:
    with args.output.open("x") as handle:
        json.dump(result, handle, indent=2)
        handle.write("\n")
print(json.dumps(result, indent=2))
