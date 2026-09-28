"""Independent 0806 raw numerical audit.

The audit intentionally does not import the main analyzer and never executes a
probe. It rereads native and allocation JSON reports, checks the semantic
sample identity that is cheap to verify from each report, and recomputes all
paired medians and the frozen bootstrap from raw values.
"""

from __future__ import annotations

import json
import math
import random
import statistics
import sys
from pathlib import Path


P = Path(__file__).resolve().parent
CASES = tuple(json.loads((P / "plan.json").read_text())["cases"])
POLICY = json.loads((P / "adoption-policy.json").read_text())
SEED = POLICY["latency"]["seed"]
RESAMPLES = POLICY["latency"]["resamples"]
LOW_RANK, HIGH_RANK = 250, 9749
LEGS = ("before", "after")
RESOURCE_FIELDS = {
    "allocation_calls": "calls",
    "allocated_bytes": "bytes",
    "net_live": "net",
    "peak_above_entry": "peak",
}
DIMENSIONS = {"tiny": (3, 4), "medium": (12, 8), "large": (100, 100),
              "vendor": (12, 8), "unicode-vendor": (12, 8),
              "valid-4attr": (12, 8)}
PROBE_SCHEMA = "litchi.pptx.public-workflow-probe-0806.v1"
PROBE_TOOL = "public-pptx-probe-0806"
PROBE_MARKER = "litchi-perf-0780-static-mce-capabilities"
VALID_URIS = [
    "urn:litchi:perf:0806:extension:one",
    "urn:litchi:perf:0806:extension:two",
    "urn:litchi:perf:0806:extension:three",
    "urn:litchi:perf:0806:extension:four",
]
VALID_NAMES = [
    "lx1:probeOne", "lx2:probeTwo", "lx3:probeThree", "lx4:probeFour",
]
VALID_VALUES = [
    "litchi-perf-0806-valid-4attr-one",
    "litchi-perf-0806-valid-4attr-two",
    "litchi-perf-0806-valid-4attr-three",
    "litchi-perf-0806-valid-4attr-four",
]


def read(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def nearest(values, percentile: float):
    ordered = sorted(values)
    return ordered[max(1, math.ceil(len(ordered) * percentile)) - 1]


def median(values):
    values = list(values)
    assert values
    return statistics.median(values)


def bootstrap(values):
    assert values
    rng = random.Random(SEED)
    draws = sorted(
        median(values[rng.randrange(len(values))] for _ in values)
        for _ in range(RESAMPLES)
    )
    return {"seed": SEED, "resamples": RESAMPLES, "low": draws[LOW_RANK],
            "high": draws[HIGH_RANK]}


def check_extension_oracle(value: dict, sample: dict, case: dict) -> None:
    verification = sample["verification"]
    fixture = value["fixture"]
    slides, shapes = DIMENSIONS["valid-4attr"]
    if case["shape"] == "valid-4attr":
        assert fixture == {
            "injection": "valid-four-distinct-namespaced-extension-attributes",
            "slide_parts": slides,
            "replaced_text_tags": slides * shapes,
            "namespace_declarations": 4,
            "namespaced_attributes": 4,
            "namespace_uris": VALID_URIS,
            "attribute_names": VALID_NAMES,
        }
        assert verification["extension_preservation_check"] is True
        assert verification["extension_text_tags"] == slides * shapes
        assert verification["extension_attributes_per_text_tag"] == 4
        assert verification["extension_attribute_occurrences"] == slides * shapes * 4
        assert verification["extension_value_occurrences"] == slides * shapes * 4
        assert verification["extension_namespace_declarations_per_slide"] == 4
        assert verification["extension_namespace_uris"] == VALID_URIS
        assert verification["extension_attribute_names"] == VALID_NAMES
        assert verification["extension_attribute_values"] == VALID_VALUES
    else:
        assert all(verification.get(name) is None for name in (
            "extension_preservation_check", "extension_text_tags",
            "extension_attributes_per_text_tag", "extension_attribute_occurrences",
            "extension_value_occurrences",
            "extension_namespace_declarations_per_slide", "extension_namespace_uris",
            "extension_attribute_names", "extension_attribute_values"))


def report(path: Path, case: dict, samples: int, lane: str) -> dict:
    value = read(path)
    assert value["schema"] == PROBE_SCHEMA
    assert value["tool"] == PROBE_TOOL
    assert value["marker"] == PROBE_MARKER
    assert value["mode"] == case["mode"] and value["shape"] == case["shape"]
    slides, shapes = DIMENSIONS[case["shape"]]
    assert value["slides"] == slides and value["shapes_per_slide"] == shapes
    fixture = value["fixture"]
    assert isinstance(fixture, dict)
    assert value["samples_requested"] == samples
    rows = value["samples"]
    assert len(rows) == samples
    elapsed = []
    allocations = []
    outputs = set()
    for index, sample in enumerate(rows):
        assert sample["index"] == index
        elapsed_value = sample["elapsed_ns"]
        assert isinstance(elapsed_value, int) and elapsed_value >= 0
        elapsed.append(elapsed_value)
        verification = sample["verification"]
        assert verification["semantic_check"] is True
        assert verification["reopened"] is True
        assert verification["expected_text"] == verification["actual_text"]
        output = sample["output"]
        check_extension_oracle(value, sample, case)
        outputs.add((output["bytes"], output["sha256"]))
        if lane == "allocation":
            allocation = sample["allocation"]
            assert allocation["status"] == "measured"
            assert allocation["scope"] == "operation_global_system_allocator"
            assert allocation["failed_allocation_calls"] == 0
            assert allocation["live_bytes_after"] == (
                allocation["live_bytes_before"] + allocation["allocated_bytes"]
                - allocation["deallocated_bytes"]
            )
            assert allocation["region_peak_live_bytes"] >= max(
                allocation["live_bytes_before"], allocation["live_bytes_after"]
            )
            allocations.append({
                **allocation,
                "net_live": allocation["live_bytes_after"]
                - allocation["live_bytes_before"],
                "peak_above_entry": allocation["region_peak_live_bytes"]
                - allocation["live_bytes_before"],
            })
        else:
            assert sample.get("allocation") is None
    assert len(outputs) == 1
    return {"path": str(path.relative_to(P)), "elapsed": elapsed,
            "p50": nearest(elapsed, 0.5), "allocation": allocations,
            "output": next(iter(outputs))}


def main() -> dict:
    assert len(CASES) == 18
    native_rows = []
    allocation_rows = []
    qualifications = []
    for case in CASES:
        native = {leg: [] for leg in LEGS}
        allocation = {leg: [] for leg in LEGS}
        for block in range(6):
            for leg in LEGS:
                path = P / "native" / f"{block}-{case['shape']}-{case['mode']}-{leg}.json"
                native[leg].append(report(path, case, 30, "native"))
        ratios = [after["p50"] / before["p50"]
                  for before, after in zip(native["before"], native["after"])]
        ci = bootstrap(ratios)
        ratio = median(ratios)
        native_rows.append({**case, "before_p50_ns": median(x["p50"] for x in native["before"]),
                            "after_p50_ns": median(x["p50"] for x in native["after"]),
                            "paired_ratios": ratios, "ratio": ratio, "ci95": ci,
                            "regression": ratio > 1.05 and ci["low"] > 1.0,
                            "benefit": case["mode"] in ("capture", "lifecycle")
                            and ratio <= 0.97 and ci["high"] < 1.0})
        for block in range(2):
            for leg in LEGS:
                path = P / "allocation" / f"{block}-{case['shape']}-{case['mode']}-{leg}.json"
                allocation[leg].append(report(path, case, 3, "allocation"))
            before, after = allocation["before"][-1], allocation["after"][-1]
            values = {}
            for field, label in RESOURCE_FIELDS.items():
                left = median(sample[field] for sample in before["allocation"])
                right = median(sample[field] for sample in after["allocation"])
                values[label] = {"before": left, "after": right,
                                 "increase": right > left}
            allocation_rows.append({**case, "block": block, "metrics": values,
                                    "resource_regression": any(
                                        value["increase"] for value in values.values())})
        path = P / "qualification" / f"0-{case['shape']}-{case['mode']}-before.json"
        value = read(path)
        assert value["shape"] == case["shape"] and value["mode"] == case["mode"]
        assert value["schema"] == PROBE_SCHEMA and value["tool"] == PROBE_TOOL
        assert value["marker"] == PROBE_MARKER
        assert len(value["samples"]) == 1
        assert value["samples"][0]["verification"]["semantic_check"] is True
        check_extension_oracle(value, value["samples"][0], case)
        qualifications.append({"case": f"{case['shape']}/{case['mode']}",
                               "report": str(path.relative_to(P)),
                               "source": value["source"],
                               "output": value["samples"][0]["output"]})
    violations = [row for row in native_rows if row["regression"]]
    benefits = [row for row in native_rows if row["benefit"]]
    resource_violations = [row for row in allocation_rows if row["resource_regression"]]
    return {
        "schema": "litchi.performance.0806.root-audit.v1",
        "reports": 306,
        "samples": 6714,
        "native": native_rows,
        "allocation": allocation_rows,
        "qualification": qualifications,
        "latency_violations": violations,
        "benefits": benefits,
        "resource_violations": resource_violations,
        "resource_guard": not resource_violations,
        "production_adoption": False,
        "scope": "Independent raw numerical audit only; main analyzer checks custody, source identity, and semantic oracles.",
    }


if __name__ == "__main__":
    result = main()
    output = P / "root-audit.json"
    if "--check" in sys.argv:
        assert json.loads(output.read_text()) == result
        print("Independent 0806 raw numerical audit PASS")
    else:
        assert not output.exists()
        output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
        print(json.dumps({"latency_violations": len(result["latency_violations"]),
                          "benefits": len(result["benefits"]),
                          "resource_violations": len(result["resource_violations"]),
                          "resource_guard": result["resource_guard"]}))
