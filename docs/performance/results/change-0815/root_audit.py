"""Independent raw numerical audit for the 0815 PPTX workflow packet.

This reader deliberately does not import ``analysis.py`` and never invokes
Cargo, a probe, a workload, or a profiler.  It rereads the retained JSON
reports, validates their semantic and allocation fields, and recomputes the
paired p50 ratios and frozen bootstrap directly from raw samples.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import statistics
import sys
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
FIXTURE_PACKET = ROOT / "docs" / "performance" / "results" / "change-0813"
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
MODES = ("capture", "commit", "lifecycle")
CASES = tuple((shape, mode) for shape in SHAPES for mode in MODES)
DIMENSIONS = {
    "tiny": (3, 4),
    "medium": (12, 8),
    "large": (100, 100),
    "vendor": (12, 8),
    "unicode-vendor": (12, 8),
    "valid-4attr": (12, 8),
}
PROBE_SCHEMA = "litchi.pptx.public-workflow-probe-0806.v1"
PROBE_TOOL = "public-pptx-probe-0806"
PROBE_MARKER = "litchi-perf-0780-static-mce-capabilities"
TIMING_SCOPES = {
    "capture": "Package::opened_presentation only",
    "commit": "Transaction::commit only; package capture and one set_shape_text staging are outside the clock",
    "lifecycle": "Package::opened_presentation, edit, set_shape_text, commit, apply_opened_presentation_commit, and Package::to_bytes",
}
VALID_URIS = [
    "urn:litchi:perf:0806:extension:one",
    "urn:litchi:perf:0806:extension:two",
    "urn:litchi:perf:0806:extension:three",
    "urn:litchi:perf:0806:extension:four",
]
VALID_NAMES = ["lx1:probeOne", "lx2:probeTwo", "lx3:probeThree", "lx4:probeFour"]
VALID_VALUES = [
    "litchi-perf-0806-valid-4attr-one",
    "litchi-perf-0806-valid-4attr-two",
    "litchi-perf-0806-valid-4attr-three",
    "litchi-perf-0806-valid-4attr-four",
]
RAW_ALLOC = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
RESOURCE_METRICS = ("allocation_calls", "allocated_bytes", "net_live", "peak_above_entry")
BOOTSTRAP_SEED = 815815
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW = 250
BOOTSTRAP_HIGH = 9749
REPORT_KEYS = {
    "schema", "tool", "mode", "shape", "slides", "shapes_per_slide",
    "timing_scope", "marker", "source", "fixture", "warmup",
    "samples_requested", "samples", "allocator",
}
COMMON_VERIFICATION_KEYS = {
    "semantic_check", "reopened", "expected_text", "actual_text",
    "semantic_text_bytes", "semantic_text_sha256", "readback_bytes",
    "readback_sha256", "marker_matches", "unknown_namespace_check",
    "unknown_namespace_occurrences",
}
EXTENSION_VERIFICATION_KEYS = {
    "extension_preservation_check", "extension_text_tags",
    "extension_attributes_per_text_tag", "extension_attribute_occurrences",
    "extension_value_occurrences", "extension_namespace_declarations_per_slide",
    "extension_namespace_uris", "extension_attribute_names",
    "extension_attribute_values",
}
RAW_ALLOCATION_KEYS = {"status", "scope", *RAW_ALLOC}


def read(path: Path) -> Any:
    assert path.is_file() and not path.is_symlink(), f"missing evidence: {path}"
    return json.loads(path.read_text(encoding="utf-8"))


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(
        character in "0123456789abcdef" for character in value
    )


def integer(value: Any, label: str, *, positive: bool = False) -> None:
    assert isinstance(value, int) and not isinstance(value, bool), f"{label}: integer required"
    assert value > 0 if positive else value >= 0, f"{label}: invalid value"


def finite_number(value: Any, label: str) -> None:
    assert isinstance(value, (int, float)) and not isinstance(value, bool)
    assert math.isfinite(float(value)) and value >= 0, f"{label}: invalid number"


def nearest(values: Iterable[float], percentile: float) -> float:
    ordered = sorted(values)
    assert ordered
    return ordered[max(1, math.ceil(len(ordered) * percentile)) - 1]


def bootstrap(values: list[float]) -> dict[str, Any]:
    assert values
    rng = random.Random(BOOTSTRAP_SEED)
    draws = sorted(
        statistics.median(values[rng.randrange(len(values))] for _ in values)
        for _ in range(BOOTSTRAP_RESAMPLES)
    )
    return {
        "seed": BOOTSTRAP_SEED,
        "resamples": BOOTSTRAP_RESAMPLES,
        "low": draws[BOOTSTRAP_LOW],
        "high": draws[BOOTSTRAP_HIGH],
    }


def frozen_plan() -> dict[str, Any]:
    plan = read(PACKET / "plan.json")
    assert plan["schema"] == "litchi.performance.0815.v1"
    assert plan["cpu"] == 12
    assert plan["source_allowlist"] == ["crates/litchi-pptx/src/notes/codec.rs"]
    assert plan["cases"] == [{"mode": mode, "shape": shape} for shape in SHAPES for mode in MODES]
    assert plan["qualification"] == {
        "leg": "before",
        "blocks": 1,
        "orders": [["before"]],
        "samples": 1,
        "warmup": 0,
        "reports": 18,
        "samples_total": 18,
        "binary": "allocation",
    }
    assert plan["native"]["blocks"] == 6
    assert plan["native"]["orders"] == [
        ["before", "after"],
        ["after", "before"],
        ["before", "after"],
        ["after", "before"],
        ["after", "before"],
        ["before", "after"],
    ]
    assert plan["native"]["samples"] == 30 and plan["native"]["warmup"] == 3
    assert plan["native"]["reports"] == 216 and plan["native"]["samples_total"] == 6480
    assert plan["allocation"] == {
        "blocks": 2,
        "orders": [["before", "after"], ["after", "before"]],
        "samples": 3,
        "warmup": 0,
        "reports": 72,
        "samples_total": 216,
    }
    profile = plan["profile"]
    assert profile["owner"] == "namespace_uri_probe::capture_region_0793"
    assert profile["binary"] == "profile" and profile["reports"] == 4
    assert profile["samples_total"] == 4
    assert plan["totals"] == {"reports": 310, "samples": 6718}
    assert plan["bootstrap"] == {
        "resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED,
        "statistic": "median",
        "sorted_zero_based_endpoints": [BOOTSTRAP_LOW, BOOTSTRAP_HIGH],
    }
    policy = read(PACKET / "adoption-policy.json")
    assert policy["schema"] == "litchi.performance.0815.adoption-policy.v1"
    assert policy["latency"]["seed"] == BOOTSTRAP_SEED
    assert policy["latency"]["resamples"] == BOOTSTRAP_RESAMPLES
    assert policy["latency"]["maximum_ratio"] == 1.05
    assert policy["benefit"] == {
        "eligible_modes": ["capture", "lifecycle"],
        "at_least_one_case_required": True,
        "minimum_improvement_percent": 3.0,
        "bootstrap95_high_below": 1.0,
    }
    assert policy["memory"]["comparison"] == "each paired allocation block median"
    return plan


def check_fixture(report: dict[str, Any], shape: str, label: str) -> None:
    assert (report["slides"], report["shapes_per_slide"]) == DIMENSIONS[shape], label
    fixture = report["fixture"]
    assert isinstance(fixture, dict)
    if shape == "valid-4attr":
        assert fixture == {
            "injection": "valid-four-distinct-namespaced-extension-attributes",
            "slide_parts": 12,
            "replaced_text_tags": 96,
            "namespace_declarations": 4,
            "namespaced_attributes": 4,
            "namespace_uris": VALID_URIS,
            "attribute_names": VALID_NAMES,
        }
    elif shape in {"vendor", "unicode-vendor"}:
        expected = (
            "same-length-known-uri-near-misses"
            if shape == "vendor"
            else "same-length-valid-utf8-unknown-uris"
        )
        assert fixture["injection"] == expected
        assert fixture["slide_parts"] == 12 and fixture["replaced_text_tags"] == 96
        assert fixture["namespace_declarations"] == 6 and fixture["namespaced_attributes"] == 6
        for key in ("namespace_uris", "attribute_names"):
            values = fixture[key]
            assert isinstance(values, list) and len(values) == 6 and len(set(values)) == 6
            assert all(isinstance(value, str) and value for value in values)
    else:
        assert fixture == {
            "injection": "none",
            "slide_parts": 0,
            "replaced_text_tags": 0,
            "namespace_declarations": 0,
            "namespaced_attributes": 0,
            "namespace_uris": [],
            "attribute_names": [],
        }


def check_extension(report: dict[str, Any], verification: dict[str, Any], label: str) -> None:
    assert verification["extension_preservation_check"] is True, label
    assert verification["extension_text_tags"] == 96
    assert verification["extension_attributes_per_text_tag"] == 4
    assert verification["extension_attribute_occurrences"] == 384
    assert verification["extension_value_occurrences"] == 384
    assert verification["extension_namespace_declarations_per_slide"] == 4
    assert verification["extension_namespace_uris"] == VALID_URIS
    assert verification["extension_attribute_names"] == VALID_NAMES
    assert verification["extension_attribute_values"] == VALID_VALUES


@lru_cache(maxsize=None)
def sealed_oracle(shape: str, mode: str) -> dict[str, Any]:
    """Return the exact 0813 source/fixture/output/verification oracle."""

    relative = f"qualification/0-{shape}-{mode}-before.json"
    path = FIXTURE_PACKET / relative
    seal = read(FIXTURE_PACKET / "seal.json")
    assert isinstance(seal.get("files"), dict)
    assert seal["files"].get(str(path.relative_to(ROOT))) == sha(path)
    report = read(path)
    assert isinstance(report.get("samples"), list) and len(report["samples"]) == 1
    sample = report["samples"][0]
    return {
        "source": report["source"],
        "fixture": report["fixture"],
        "source_sha256": sample["source_sha256"],
        "output": sample["output"],
        "verification": sample["verification"],
    }


def allocation_values(sample: dict[str, Any], label: str) -> dict[str, int]:
    value = sample["allocation"]
    assert value["status"] == "measured" and value["scope"] == "operation_global_system_allocator"
    result = {}
    for field in RAW_ALLOC:
        integer(value[field], f"{label}.{field}")
        result[field] = value[field]
    assert result["failed_allocation_calls"] == 0
    assert result["live_bytes_after"] == (
        result["live_bytes_before"] + result["allocated_bytes"] - result["deallocated_bytes"]
    )
    assert result["region_peak_live_bytes"] >= max(
        result["live_bytes_before"], result["live_bytes_after"]
    )
    assert result["peak_live_bytes_after"] >= result["region_peak_live_bytes"]
    result["net_live"] = result["live_bytes_after"] - result["live_bytes_before"]
    result["peak_above_entry"] = result["region_peak_live_bytes"] - result["live_bytes_before"]
    return result


def read_report(
    path: Path,
    shape: str,
    mode: str,
    samples: int,
    warmup: int,
    *,
    allocation: bool,
) -> dict[str, Any]:
    label = str(path.relative_to(PACKET))
    report = read(path)
    assert set(report) == REPORT_KEYS, f"{label}: report fields changed"
    assert report["schema"] == PROBE_SCHEMA and report["tool"] == PROBE_TOOL
    assert report["marker"] == PROBE_MARKER
    assert report["shape"] == shape and report["mode"] == mode
    assert report["timing_scope"] == TIMING_SCOPES[mode]
    assert report["warmup"] == warmup
    assert report["samples_requested"] == samples
    check_fixture(report, shape, label)
    oracle = sealed_oracle(shape, mode)
    assert report["source"] == oracle["source"], f"{label}: source oracle changed"
    assert report["fixture"] == oracle["fixture"], f"{label}: fixture oracle changed"
    source = report["source"]
    assert is_sha(source["sha256"]) and isinstance(source["bytes"], int) and source["bytes"] > 0
    allocator = report["allocator"]
    assert allocator["binary"]
    if allocation:
        assert allocator["instrumentation"] == "system_allocator_operation_scoped"
        assert allocator["allocator"] == "CountingSystemAllocator(std::alloc::System)"
        assert allocator["counter_revision"] == "serialized_region_peak_v3"
    else:
        assert allocator["instrumentation"] == "none"
        assert allocator["allocator"] == "Rust system allocator"
        assert allocator["counter_revision"] is None
    rows = report["samples"]
    assert isinstance(rows, list) and len(rows) == samples
    elapsed = []
    allocations = {field: [] for field in (*RAW_ALLOC, "net_live", "peak_above_entry")}
    outputs = set()
    for index, sample in enumerate(rows):
        expected_sample_keys = {
            "index", "elapsed_ns", "metrics", "source_sha256", "output", "verification",
        }
        if allocation:
            expected_sample_keys.add("allocation")
        assert set(sample) == expected_sample_keys, f"{label}:{index}: sample fields changed"
        assert sample["index"] == index
        integer(sample["elapsed_ns"], f"{label}:{index}:elapsed", positive=True)
        elapsed.append(sample["elapsed_ns"])
        assert sample["source_sha256"] == source["sha256"]
        assert sample["source_sha256"] == oracle["source_sha256"], f"{label}:{index}: source oracle changed"
        metrics = sample["metrics"]
        assert metrics["elapsed_ns"] == sample["elapsed_ns"]
        assert metrics["slides"] == report["slides"]
        assert metrics["shapes_per_slide"] == report["shapes_per_slide"]
        verification = sample["verification"]
        expected_verification_keys = set(COMMON_VERIFICATION_KEYS)
        if shape == "valid-4attr":
            expected_verification_keys.update(EXTENSION_VERIFICATION_KEYS)
        assert set(verification) == expected_verification_keys, f"{label}:{index}: verification fields changed"
        assert sample["output"] == oracle["output"], f"{label}:{index}: output oracle changed"
        assert verification == oracle["verification"], f"{label}:{index}: verification oracle changed"
        assert verification["semantic_check"] is True and verification["reopened"] is True
        assert verification["expected_text"] == verification["actual_text"]
        assert is_sha(verification["semantic_text_sha256"])
        output = sample["output"]
        assert is_sha(output["sha256"]) and isinstance(output["bytes"], int) and output["bytes"] >= 0
        assert verification["readback_bytes"] == output["bytes"]
        assert verification["readback_sha256"] == output["sha256"]
        expected_marker = mode in {"commit", "lifecycle"}
        assert verification["marker_matches"] is (True if expected_marker else None)
        vendor = shape in {"vendor", "unicode-vendor"}
        assert verification["unknown_namespace_check"] is (True if vendor else None)
        assert verification["unknown_namespace_occurrences"] == (
            report["slides"] * report["shapes_per_slide"] if vendor else None
        )
        if shape == "valid-4attr":
            check_extension(report, verification, f"{label}:{index}")
        else:
            for field in (
                "extension_preservation_check", "extension_text_tags",
                "extension_attributes_per_text_tag", "extension_attribute_occurrences",
                "extension_value_occurrences", "extension_namespace_declarations_per_slide",
                "extension_namespace_uris", "extension_attribute_names", "extension_attribute_values",
            ):
                assert verification.get(field) is None
        outputs.add((output["bytes"], output["sha256"]))
        if allocation:
            assert set(sample["allocation"]) == RAW_ALLOCATION_KEYS
            values = allocation_values(sample, f"{label}:{index}")
            for field, value in values.items():
                allocations[field].append(value)
        else:
            assert sample.get("allocation") is None
    assert len(outputs) == 1
    return {
        "elapsed": elapsed,
        "p50": nearest(elapsed, 0.5),
        "allocation": allocations if allocation else None,
        "output": next(iter(outputs)),
    }


def report_path(lane: str, block: int, shape: str, mode: str, leg: str) -> Path:
    return PACKET / lane / f"{block}-{shape}-{mode}-{leg}.json"


def collect(plan: dict[str, Any], lane: str) -> dict[tuple[str, str, int, str], dict[str, Any]]:
    lane_plan = plan[lane]
    orders = lane_plan["orders"]
    result = {}
    for block, order in enumerate(orders):
        for shape, mode in CASES:
            for leg in order:
                path = report_path(lane, block, shape, mode, leg)
                result[(shape, mode, block, leg)] = read_report(
                    path,
                    shape,
                    mode,
                    lane_plan["samples"],
                    lane_plan["warmup"],
                    allocation=lane == "allocation",
                )
    assert len(result) == lane_plan["reports"]
    return result


def collect_qualification(plan: dict[str, Any]) -> dict[tuple[str, str], dict[str, Any]]:
    result = {}
    for shape, mode in CASES:
        path = report_path("qualification", 0, shape, mode, "before")
        result[(shape, mode)] = read_report(
            path,
            shape,
            mode,
            1,
            plan["qualification"]["warmup"],
            allocation=True,
        )
    assert len(result) == plan["qualification"]["reports"]
    return result


def main() -> dict[str, Any]:
    plan = frozen_plan()
    native = collect(plan, "native")
    allocation = collect(plan, "allocation")
    qualification = collect_qualification(plan)
    latency_rows = []
    latency_violations = []
    benefits = []
    resource_rows = []
    resource_violations = []
    for shape, mode in CASES:
        ratios = []
        native_blocks = []
        for block in range(plan["native"]["blocks"]):
            before = native[(shape, mode, block, "before")]["p50"]
            after = native[(shape, mode, block, "after")]["p50"]
            ratio = after / before if before else (1.0 if after == 0 else None)
            assert ratio is not None
            ratios.append(ratio)
            native_blocks.append({"block": block, "before": before, "after": after, "ratio": ratio})
        ratio_median = statistics.median(ratios)
        ci = bootstrap(ratios)
        row = {
            "shape": shape,
            "mode": mode,
            "before_p50_ns": statistics.median(
                native[(shape, mode, block, "before")]["p50"]
                for block in range(plan["native"]["blocks"])
            ),
            "after_p50_ns": statistics.median(
                native[(shape, mode, block, "after")]["p50"]
                for block in range(plan["native"]["blocks"])
            ),
            "paired_ratios": ratios,
            "ratio": ratio_median,
            "ci95": ci,
            "blocks": native_blocks,
        }
        latency_rows.append(row)
        if ratio_median > 1.05 and ci["low"] > 1.0:
            latency_violations.append(row)
        if mode in {"capture", "lifecycle"} and (1 - ratio_median) * 100 >= 3.0 and ci["high"] < 1.0:
            benefits.append(row)
        for block in range(plan["allocation"]["blocks"]):
            before = allocation[(shape, mode, block, "before")]["allocation"]
            after = allocation[(shape, mode, block, "after")]["allocation"]
            assert before is not None and after is not None
            metrics = {}
            for metric in RESOURCE_METRICS:
                left = statistics.median(before[metric])
                right = statistics.median(after[metric])
                item = {"before": left, "after": right, "increase": right > left}
                metrics[metric] = item
                resource_rows.append({
                    "shape": shape, "mode": mode, "block": block,
                    "metric": metric, **item,
                })
                if item["increase"]:
                    resource_violations.append({
                        "shape": shape, "mode": mode, "block": block,
                        "metric": metric, **item,
                    })
    assert len(latency_rows) == 18 and len(qualification) == 18
    assert len(resource_rows) == 18 * 2 * len(RESOURCE_METRICS)
    adoption_eligible = not latency_violations and not resource_violations and bool(benefits)
    return {
        "schema": "litchi.performance.0815.root-audit.v1",
        "reports": 306,
        "samples": 6714,
        "native": latency_rows,
        "qualification": [
            {"shape": shape, "mode": mode, "output": list(value["output"])}
            for (shape, mode), value in qualification.items()
        ],
        "allocation": resource_rows,
        "latency_violations": latency_violations,
        "benefits": benefits,
        "resource_violations": resource_violations,
        "resource_guard": not resource_violations,
        "adoption_eligible": adoption_eligible,
        "bootstrap": {
            "seed": BOOTSTRAP_SEED,
            "resamples": BOOTSTRAP_RESAMPLES,
            "low_rank": BOOTSTRAP_LOW,
            "high_rank": BOOTSTRAP_HIGH,
        },
        "scope": "Independent raw numerical policy audit only; adoption_eligible is not a retention decision. Custody, profile, and final disposition remain separate readers.",
    }


if __name__ == "__main__":
    value = main()
    output = PACKET / "root-audit.json"
    if "--check" in sys.argv:
        assert json.loads(output.read_text(encoding="utf-8")) == value
        print("Independent 0815 raw numerical audit PASS")
    else:
        assert not output.exists()
        output.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({
            "latency_violations": len(value["latency_violations"]),
            "benefits": len(value["benefits"]),
            "resource_violations": len(value["resource_violations"]),
            "resource_guard": value["resource_guard"],
        }))
