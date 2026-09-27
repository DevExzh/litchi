"""Fail-closed offline validation for the 0779 XLSX evidence packet.

Validation replays analyze.py and compares its deterministic result with the
retained analysis.json.  It never starts Cargo, the probe, or a profiler.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import analyze


PACKET = analyze.PACKET


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def validate() -> dict[str, Any]:
    plan = analyze.load_plan()
    builds, _cleanup, _cleanup_verified = analyze.load_builds(plan)
    quality = analyze.check_quality(builds["after"]["source"])
    require(quality is not None, "quality.json is missing")
    result = analyze.analyze()

    analysis_path = PACKET / "analysis.json"
    require(analysis_path.is_file(), "analysis.json is missing; run analyze.py after capture")
    retained = analyze.read_json(analysis_path)
    require(retained == result, "analysis.json does not replay byte-for-byte")

    require(result["native"]["children"] == 96, "native child cardinality changed")
    require(result["allocation"]["children"] == 32, "allocation child cardinality changed")
    require(result["qualification"]["children"] == 8,
            "qualification child cardinality changed")
    small = result["small_controls"]
    require(small["native"]["children"] == 24,
            "small native child cardinality changed")
    require(small["allocation"]["children"] == 8,
            "small allocation child cardinality changed")
    require(result["source"]["changed_files"] == [analyze.PHYS_PKG],
            "source census allowlist changed")
    disposition = result["disposition"]
    require(disposition["status"] == "rejected"
            and disposition["production_change_retained"] is False,
            "candidate disposition is not the recorded rejection")
    require(disposition["source_file"] == analyze.PHYS_PKG,
            "disposition source file changed")
    require(disposition["live_source_files_match_before"] is True,
            "live source census was not restored to before")
    require(
        disposition["before_source"]["sha256"]
        == result["source"]["before"]["files"][analyze.PHYS_PKG],
        "archived before source does not match before census",
    )
    require(
        disposition["candidate_source"]["sha256"]
        == result["source"]["after"]["files"][analyze.PHYS_PKG],
        "archived candidate source does not match after census",
    )
    verification = result["verification"]
    require(verification["source_binary_probe_lock_fixture_receipts_checked"] is True,
            "custody checks were not recorded")
    require(verification["native_has_no_allocation_metrics"] is True,
            "native allocation exclusion was not recorded")
    require(verification["published_output_digest_before_after_checked"] is True,
            "published before/after output check was not recorded")
    require(verification["published_output_digest_matches_0778_no_sync"] is True,
            "0778 no-sync output check was not recorded")
    require(verification["marker_verification_checked"] is True,
            "marker verification was not recorded")
    require(verification["cleanup_binary_witness_required_when_missing"] is True,
            "cleanup custody contract is unstable")
    require(verification["optional_file_inventory_seal_checked"] is True,
            "optional seal was not checked")

    # Recompute the retained receipt cardinalities independently of summary
    # group counts, so a dropped row cannot hide inside an aggregate.
    require(len(result["native"]["receipts"]) == 96, "native receipt list is incomplete")
    require(len(result["allocation"]["receipts"]) == 32,
            "allocation receipt list is incomplete")
    require(len(result["qualification"]["rows"]) == 8,
            "qualification receipt list is incomplete")
    require(len(small["native"]["receipts"]) == 24,
            "small native receipt list is incomplete")
    require(len(small["allocation"]["receipts"]) == 8,
            "small allocation receipt list is incomplete")
    native_analysis = result["native"]["analysis"]
    allocation_analysis = result["allocation"]["analysis"]
    small_native_analysis = small["native"]["analysis"]
    small_allocation_analysis = small["allocation"]["analysis"]
    require("paired_by_block_before_after" in native_analysis
            and "regression_flags_over_5_percent" in native_analysis,
            "native paired regression analysis is missing")
    require("paired_by_block_before_after" in allocation_analysis
            and "regression_flags_over_5_percent" in allocation_analysis,
            "allocation paired regression analysis is missing")
    require(set(small_native_analysis["groups"]) == {
        "small-xlsx/open/before", "small-xlsx/open/after",
        "boundary-xlsx/open/before", "boundary-xlsx/open/after",
    }, "small native group keys are not separate")
    require(set(small_allocation_analysis["groups"]) == {
        "small-xlsx/open/before", "small-xlsx/open/after",
        "boundary-xlsx/open/before", "boundary-xlsx/open/after",
    }, "small allocation group keys are not separate")
    return {
        "quality_gates": quality["gates"],
        "native_children": result["native"]["children"],
        "allocation_children": result["allocation"]["children"],
        "qualification_children": result["qualification"]["children"],
        "small_native_children": small["native"]["children"],
        "small_allocation_children": small["allocation"]["children"],
        "native_spread_flags": len(native_analysis["spread_flags_over_5_percent"]),
        "native_regression_flags": len(native_analysis["regression_flags_over_5_percent"]),
        "allocation_spread_flags": len(
            allocation_analysis["spread_flags_over_5_percent"]
        ),
        "allocation_regression_flags": len(
            allocation_analysis["regression_flags_over_5_percent"]
        ),
        "small_native_spread_flags": len(
            small_native_analysis["spread_flags_over_5_percent"]
        ),
        "small_allocation_spread_flags": len(
            small_allocation_analysis["spread_flags_over_5_percent"]
        ),
        "optional_file_inventory_seal": any(
            (PACKET / name).is_file() for name in ("seal.json", "final-seal.json")
        ),
    }


if __name__ == "__main__":
    print(json.dumps(validate(), indent=2, sort_keys=True))
