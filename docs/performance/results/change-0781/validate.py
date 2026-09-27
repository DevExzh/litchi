"""Fail-closed offline replay for the 0781 borrowed-text evidence packet."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import analyze


PACKET = analyze.PACKET


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def validate(*, require_final_seal: bool = False) -> dict[str, Any]:
    result = analyze.analyze()
    seal_paths = [PACKET / name for name in ("seal.json", "final-seal.json")]
    seal_present = any(path.is_file() and not path.is_symlink() for path in seal_paths)
    if require_final_seal:
        require(seal_present, "final seal is missing")
    retained_path = PACKET / "analysis.json"
    require(retained_path.is_file(), "analysis.json is missing; run analyze.py after capture")
    retained = analyze.read_json(retained_path)
    require(retained == result, "analysis.json does not replay byte-for-byte")

    require(result["native"]["children"] == 120, "native child cardinality changed")
    require(result["allocation"]["children"] == 40, "allocation child cardinality changed")
    require(result["qualification"]["children"] == 10,
            "qualification child cardinality changed")
    require(result["quality"]["gates"] == 6, "quality gate cardinality changed")
    require(result["test_summary"]["passed"] >= 0
            and result["test_summary"]["failed"] == 0
            and result["test_summary"]["ignored"] >= 0
            and result["test_summary"]["suites"] > 0,
            "quality test summary changed")
    probe_tests = result["probe_tests"]
    require(result["verification"]["probe_tests_receipt_and_named_oracles_checked"] is True,
            "probe test receipt was not checked")
    require(isinstance(probe_tests.get("rows"), list) and len(probe_tests["rows"]) == 2,
            "probe test row cardinality changed")
    for row in probe_tests["rows"]:
        tests = row.get("tests", {})
        require(tests.get("failed") == 0 and tests.get("ignored") == 0
                and tests.get("passed", 0) >= len(analyze.PROBE_TEST_NAMES)
                and set(analyze.PROBE_TEST_NAMES).issubset(set(tests.get("names", []))),
                "probe oracle test names or count changed")
    initial_relocations = result["initial_binary_relocations"]
    require(result["verification"]["initial_binary_relocations_replayed"] is True
            and isinstance(initial_relocations.get("binaries"), list)
            and {row.get("kind") for row in initial_relocations["binaries"]}
            == {"native", "allocation"}
            and initial_relocations.get("relocations") == {
                "build-before": "build-before-initial",
                "qualification": "qualification-initial",
            }, "initial binary relocation custody changed")
    parity = result["qualification_parity"]
    require(result["verification"]["qualification_parity_replayed"] is True
            and parity.get("source_and_output_unchanged") is True
            and isinstance(parity.get("rows"), list)
            and len(parity["rows"]) == len(analyze.QUALIFICATION_PARITY_CASES),
            "qualification correction parity changed")
    baseline = result["baseline_attribution"]
    require(result["verification"]["baseline_attribution_replayed"] is True
            and baseline.get("crosscheck") is True
            and baseline.get("conversion", {}).get("symbol") ==
            "convert_shape_to_escher_with_sound_mapping",
            "baseline attribution replay changed")
    require(result["source"]["changed_files"] == sorted(analyze.load_plan()["source_allowlist"]),
            "source census allowlist changed")
    require(result["verification"]["native_has_no_allocation_metrics"] is True,
            "native allocation exclusion was not recorded")
    require(result["verification"]["qualification_preflight_outside_paired_matrix"] is True,
            "qualification preflight was not separated")
    require(result["verification"]["fixed_release_profile_checked"] is True,
            "fixed release profile was not checked")
    require(result["verification"]["quality_commands_checked_exactly"] is True,
            "quality commands were not checked exactly")
    require(result["verification"]["observer_analysis_replayed"] is True,
            "observer analysis was not replayed")
    require(result["observer"]["analysis"]["schema"] == "litchi-0781-observer-analysis-v1",
            "observer analysis schema changed")
    require(result["verification"]["ci_analysis_replayed"] is True,
            "CI analysis was not replayed")
    require(result["ci"]["analysis"].get("schema") == analyze.CI_ANALYSIS_SCHEMA,
            "CI analysis schema changed")

    native_analysis = result["native"]["analysis"]
    allocation_analysis = result["allocation"]["analysis"]
    require(len(native_analysis["groups"]) == 10 * 2,
            "native process group cardinality changed")
    require(len(allocation_analysis["groups"]) == 10 * 2,
            "allocation process group cardinality changed")
    require(set(allocation_analysis["fields"]) == set(analyze.ALLOCATION_FIELDS),
            "allocator metric field set changed")
    require(result["bootstrap"] == {
        "seed": analyze.BOOTSTRAP_SEED,
        "resamples": analyze.BOOTSTRAP_RESAMPLES,
        "confidence": analyze.BOOTSTRAP_CONFIDENCE,
        "statistic": "median",
    }, "bootstrap contract changed")

    for family in (native_analysis, allocation_analysis):
        for paired in family["paired_by_block_before_after"].values():
            for metric in paired["metrics"].values():
                require(isinstance(metric.get("bootstrap"), dict),
                        "paired bootstrap confidence interval is missing")
                require(metric["bootstrap"]["seed"] == analyze.BOOTSTRAP_SEED,
                        "paired bootstrap seed changed")
                require(len(metric["by_block"]) == paired["blocks"],
                        "paired block ratio cardinality changed")
                require(metric["defined_ratio_blocks"] + metric["undefined_ratio_blocks"]
                        == paired["blocks"], "paired zero-baseline rows were dropped")

    require(result["disposition"]["status"] in {"retained", "rejected"},
            "candidate disposition is not recorded")
    require(result["verification"]["cleanup_binary_witness_required_when_missing"] is True,
            "cleanup binary custody contract is unstable")
    require(result["verification"]["optional_file_inventory_seal_checked"] is True,
            "final seal was not checked")
    return {
        "native_children": result["native"]["children"],
        "allocation_children": result["allocation"]["children"],
        "qualification_children": result["qualification"]["children"],
        "quality_gates": result["quality"]["gates"],
        "native_spread_flags": len(native_analysis["spread_flags_over_5_percent"]),
        "native_regression_flags": len(native_analysis["regression_flags_over_5_percent"]),
        "allocation_spread_flags": len(allocation_analysis["spread_flags_over_5_percent"]),
        "allocation_regression_flags": len(
            allocation_analysis["regression_flags_over_5_percent"]
        ),
        "seal_checked": seal_present,
    }


if __name__ == "__main__":
    print(json.dumps(validate(require_final_seal=True), indent=2, sort_keys=True))
