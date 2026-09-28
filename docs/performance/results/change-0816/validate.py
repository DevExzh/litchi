"""Fail-closed offline validator for the 0816 delayed-source packet."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import analyze


PACKET = analyze.PACKET


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def validate(*, require_cleanup: bool = False,
             require_final_seal: bool = False) -> dict[str, Any]:
    result = analyze.analyze(check=True)
    require(result["schema"] == analyze.ANALYSIS_SCHEMA, "analysis schema changed")
    require(result["plan_schema"] == "litchi.performance.0816.plan.v1", "plan schema changed")
    require(result["report_schema"] == analyze.REPORT_SCHEMA
            and result["range_report_schema"] == analyze.RANGE_REPORT_SCHEMA,
            "probe schemas changed")
    require(result["counts"] == {
        "reports": 648,
        "samples": 13320,
        "native_reports": 432,
        "native_samples": 12960,
        "observer_reports": 144,
        "observer_samples": 288,
        "qualification_reports": 72,
        "qualification_samples": 72,
    }, "aggregate report/sample cardinality changed")
    require(result["quality"]["gates"] == 6, "quality gate cardinality changed")
    require(result["source"]["production_file_count"] == 9196
            and result["source"]["production_byte_identical_to_base"] is True
            and result["source"]["tool_source_frozen"] is True,
            "source custody changed")
    architecture = result["architecture_inputs"]
    require(architecture["count"] == 35
            and architecture["revision"] == analyze.origin()["base"]
            and architecture["live_files_match"] is True
            and architecture["origin_blob_hashes_match"] is True,
            "architecture input custody changed")
    require(result["bootstrap"] == {
        "seed": analyze.BOOTSTRAP_SEED,
        "resamples": analyze.BOOTSTRAP_RESAMPLES,
        "confidence": analyze.BOOTSTRAP_CONFIDENCE,
        "statistic": "median of six paired block ratios",
        "low_rank": 250,
        "high_rank": 9749,
    }, "bootstrap contract changed")
    require(len(result["scaling"]) == 72, "scaling row cardinality changed")
    amdahl = result.get("amdahl")
    require(isinstance(amdahl, dict) and len(amdahl) == 18,
            "Amdahl family cardinality changed")
    for family, fit in amdahl.items():
        require(isinstance(fit, dict), f"Amdahl fit is malformed: {family}")
        unconstrained = fit.get("unconstrained_serial_fraction")
        if isinstance(unconstrained, (int, float)):
            admissible = unconstrained == unconstrained and unconstrained not in (
                float("inf"), float("-inf")) and 0.0 <= unconstrained <= 1.0
            if not admissible:
                reasons = fit.get("invalid_reasons")
                require(fit.get("valid") is False and fit.get("invalid") is True
                        and isinstance(reasons, list)
                        and any("outside model-admissible [0,1]" in str(reason)
                                for reason in reasons),
                        f"out-of-range Amdahl fit was not invalidated: {family}")
        if fit.get("valid") is True:
            require(fit.get("invalid") is False
                    and fit.get("invalid_reasons") == [],
                    f"valid Amdahl fit carries invalid diagnostics: {family}")
    raw_audit = result["raw_audit"]
    require(raw_audit["reports"] == 648 and raw_audit["samples"] == 13320
            and raw_audit["independently_reconstructed_payloads"] == 64
            and raw_audit["rows"] == 72
            and raw_audit["source_control_rows"] == 24
            and raw_audit["paired_curves_match"] is True,
            "independent raw audit changed")
    require(result["observer"]["timings_pooled_with_native"] is False
            and result["observer"]["reports"] == 216,
            "observer separation changed")
    require(result["qualification_audit"]["accepted_before_native"] is True
            and result["qualification_audit"]["timings_imported"] is False
            and result["qualification_audit"]["reports"] == 72,
            "qualification acceptance gate changed")
    verification = result["verification"]
    for key in (
        "plan_and_order_checked", "receipts_checked", "report_schema_checked",
        "output_parity_checked", "resource_limits_checked",
        "final_permits_released_checked", "production_base_git_blobs_checked",
        "architecture_inputs_checked", "tool_source_checked",
        "quality_commands_checked_exactly", "observer_timing_separated",
        "amdahl_descriptive_only", "raw_audit_checked", "qualification_audit_checked",
    ):
        require(verification.get(key) is True, f"verification marker missing: {key}")
    for row in result["reports"]:
        require(row["verification_ok"] is True and row["resource_ok"] is True,
                "report correctness witness changed")
        require(row["report"]["sha256"] == analyze.sha256(
            analyze.resolve_path(row["report"]["path"])),
                "retained report digest changed")
    cleanup_checked = False
    if require_cleanup:
        cleanup = PACKET / "cleanup.json"
        require(cleanup.is_file() and not cleanup.is_symlink(),
                "cleanup witness is missing")
        cleanup_value = analyze.read_json(cleanup)
        require(cleanup_value.get("schema") == "litchi.performance.0816.cleanup.v1"
                and cleanup_value.get("target_removed") is True
                and cleanup_value.get("binaries_verified_before_removal") is True,
                "cleanup witness changed")
        binaries = cleanup_value.get("removed_binaries")
        require(isinstance(binaries, list) and len(binaries) == 2,
                "cleanup binary witness cardinality changed")
        build_binaries = result["build"]["binaries"]
        require(all(any(
            item == {key: expected[key] for key in ("path", "bytes", "sha256")}
            for item in binaries
        ) for expected in build_binaries.values()),
                "cleanup binary witness does not exactly match build")
        cleanup_source = cleanup_value.get("source")
        require(isinstance(cleanup_source, dict), "cleanup source witness is missing")
        source_path = analyze.resolve_path(result["build"]["source"]["path"])
        require(cleanup_source.get("bytes") == source_path.stat().st_size
                and cleanup_source.get("sha256") == analyze.sha256(source_path),
                "cleanup source witness is not bound to build")
        cleanup_checked = True
    if require_final_seal:
        seal = PACKET / "seal.json"
        require(seal.is_file() and not seal.is_symlink(), "final seal is missing")
        value = analyze.read_json(seal)
        require(value.get("schema") == "litchi.performance.0816.seal.v1",
                "final seal schema changed")
        files = value.get("files")
        require(isinstance(files, dict) and files, "final seal files are missing")
        for name, digest in files.items():
            path = analyze.ROOT / name
            require(path.is_file() and not path.is_symlink()
                    and analyze.sha256(path) == digest,
                    f"final seal payload changed: {name}")
    return {
        "reports": result["counts"]["reports"],
        "samples": result["counts"]["samples"],
        "native_reports": result["counts"]["native_reports"],
        "observer_reports": result["counts"]["observer_reports"],
        "qualification_reports": result["counts"]["qualification_reports"],
        "scaling_rows": len(result["scaling"]),
        "raw_audit_rows": raw_audit["rows"],
        "quality_gates": result["quality"]["gates"],
        "cleanup_checked": cleanup_checked,
        "seal_checked": require_final_seal,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--final", dest="require_cleanup", action="store_true",
                        help="require the post-cleanup witness")
    parser.add_argument("--require-final-seal", dest="require_final_seal",
                        action="store_true", help="also check seal.json payloads")
    args = parser.parse_args(argv)
    try:
        print(json.dumps(validate(require_cleanup=args.require_cleanup,
                                  require_final_seal=args.require_final_seal),
                         indent=2, sort_keys=True))
    except analyze.ReplayError as error:
        print(f"0816 validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
