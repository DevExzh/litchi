"""Fail-closed offline validator for the 0817 ordinary-save packet.

This module replays retained evidence through :mod:`analyze`; it never starts
Cargo, a benchmark child, an exporter, or a filesystem workload.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import analyze


PACKET = analyze.PACKET
ROOT = analyze.ROOT


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def check_quality_attempt_acknowledgement(result: dict[str, Any]) -> None:
    quality = result["quality"]
    require(quality.get("attempt", 0) >= 2,
            "final quality receipt does not acknowledge the failed recovery attempt")
    recovery = quality.get("recovery")
    require(isinstance(recovery, dict)
            and recovery.get("inherited_passed") == 640
            and recovery.get("inherited_ignored") == 1
            and recovery.get("inherited_successful_suites") == 26
            and recovery.get("changed_test") == analyze.ALLOWLISTED_TEST_SOURCE,
            "final quality receipt does not retain the bounded test recovery")
    driver_recovery = quality.get("packet_driver_recovery")
    require(isinstance(driver_recovery, dict)
            and analyze.is_sha(driver_recovery.get("historical_build_sha256"))
            and analyze.is_sha(driver_recovery.get("current_build_sha256"))
            and driver_recovery["historical_build_sha256"] !=
            driver_recovery["current_build_sha256"],
            "final quality receipt does not retain the build-driver correction")
    expected_packet = dict(result["build"]["packet"])
    expected_drivers = dict(result["build"]["drivers"])
    expected_packet["build.py"] = driver_recovery["historical_build_sha256"]
    expected_drivers["build.py"] = driver_recovery["historical_build_sha256"]
    require(quality.get("packet") == expected_packet
            and quality.get("drivers") == expected_drivers,
            "quality receipt packet/driver provenance changed")
    require(quality.get("provenance") == result["build"]["provenance"],
            "quality receipt provenance changed")
    failed = PACKET / "quality-0" / "guard-failure.json"
    require(failed.is_file() and not failed.is_symlink(),
            "quality attempt-0 guard-failure witness is missing")
    value = analyze.read_json(failed)
    require(value.get("schema") == "litchi.performance.0817.guard-failure.v1"
            and value.get("exit_code") == 1
            and value.get("successful_cargo_gates") == 2,
            "quality attempt-0 guard-failure witness changed")
    require("in-memory packet hash" in value.get("evidence_limit", "")
            and "No build or workload has run" in value.get("recovery", ""),
            "quality attempt-0 evidence-limit acknowledgement changed")


def check_counts(result: dict[str, Any]) -> None:
    require(result["counts"] == {
        "reports": 0,
        "samples": 0,
        "native_reports": 0,
        "native_samples": 0,
        "observer_reports": 0,
        "observer_samples": 0,
        "qualification_reports": 0,
        "qualification_samples": 0,
        "artifact_cases": 6,
        "artifact_policy_outputs": 30,
    }, "aggregate report/sample cardinality changed")
    artifacts = result["artifacts"]
    require(artifacts["cases"] == 6
            and artifacts["policy_outputs"] == 30
            and artifacts["audit_cases"] == 6
            and artifacts["audit_ok"] is False
            and len(artifacts["audit_case_status"]) == 6,
            "artifact cardinality or failed-audit status changed")


def check_source_custody(result: dict[str, Any]) -> None:
    custody = result["custody"]
    require(custody.get("production_matches_base") is True
            and custody.get("harness_matches_base") is False
            and custody.get("runtime_harness_matches_base") is True
            and custody.get("tool_allowlist") == [analyze.ALLOWLISTED_TEST_SOURCE]
            and custody.get("allowlisted_test_change") == analyze.ALLOWLISTED_TEST_SOURCE,
            "source allowlist custody changed")


def check_build_recovery(result: dict[str, Any]) -> None:
    recovery = result["build"].get("driver_recovery")
    require(isinstance(recovery, dict)
            and recovery.get("schema") ==
            "litchi.performance.0817.build-driver-recovery.v1"
            and recovery.get("failed_attempt") == 0
            and recovery.get("cargo_started") is False
            and recovery.get("exit_code") == 1
            and recovery.get("packet_changes") == ["build.py"]
            and recovery.get("driver_changes") == ["build.py"],
            "build-driver recovery witness changed")


def check_admission_failure(result: dict[str, Any]) -> None:
    require(result.get("status") == "admission_failed"
            and result.get("timing_status", {}).get("accepted") is False,
            "admission failure status changed")
    admission = result.get("admissions")
    require(isinstance(admission, dict)
            and admission.get("schema") == analyze.ADMISSION_SUMMARY_SCHEMA
            and admission.get("accepted") is False
            and admission.get("stage") == "artifacts"
            and admission.get("attempt") == "admission-0"
            and admission.get("exit_code") == 1
            and admission.get("audit_ok") is False
            and admission.get("reason") == analyze.ADMISSION_FAILURE_REASON,
            "admission summary changed")
    require(admission.get("case_count") == 6
            and admission.get("policy_output_count") == 30
            and admission.get("audit_error_count") == 22
            and len(admission.get("audit_errors", [])) == 22
            and len(admission.get("cases", [])) == 6,
            "failed admission evidence cardinality changed")
    require(result["artifacts"].get("audit_errors") == admission.get("audit_errors"),
            "artifact audit errors are not retained in admission summary")
    require(result["observer"].get("reports") == []
            and result["native"] == []
            and result["qualification"] == []
            and result["raw_reports"] == {
                "native": [], "observer": [], "qualification": []
            }, "timing rows were synthesized after admission failure")
    require(result["timing_status"] == {
        "accepted": False,
        "reason": analyze.ADMISSION_FAILURE_REASON,
        "reports": 0,
        "samples": 0,
        "lanes": {
            "qualification": {"reports": 0, "samples": 0},
            "native": {"reports": 0, "samples": 0},
            "observer": {"reports": 0, "samples": 0},
        },
    }, "timing absence witness changed")


def check_no_timing(result: dict[str, Any]) -> None:
    require(result["bootstrap"].get("status") == "not_run",
            "bootstrap was unexpectedly computed without timing samples")
    require(result["bootstrap"].get("seed") == analyze.BOOTSTRAP_SEED
            and result["bootstrap"].get("resamples") == analyze.BOOTSTRAP_RESAMPLES
            and result["bootstrap"].get("low_rank") == analyze.BOOTSTRAP_LOW_RANK
            and result["bootstrap"].get("high_rank") == analyze.BOOTSTRAP_HIGH_RANK,
            "unrun bootstrap contract changed")
    require(result["raw_reports"] == {
        "native": [], "observer": [], "qualification": []
    }, "raw timing reports were retained despite failed admission")


def check_cleanup(result: dict[str, Any]) -> None:
    path = PACKET / "cleanup.json"
    require(path.is_file() and not path.is_symlink(), "cleanup witness is missing")
    value = analyze.read_json(path)
    require(set(value) == {
        "schema", "target_removed", "scratch_removed",
        "binaries_verified_before_removal", "removed_binaries", "source",
        "removed", "started", "ended",
    }, "cleanup schema changed")
    require(value.get("schema") == "litchi.performance.0817.cleanup.v1"
            and value.get("target_removed") is True
            and value.get("scratch_removed") is True
            and value.get("binaries_verified_before_removal") is True,
            "cleanup removal witness changed")
    removed_binaries = value.get("removed_binaries")
    require(isinstance(removed_binaries, list) and len(removed_binaries) == 3,
            "cleanup binary witness cardinality changed")
    build_json = analyze.read_json(PACKET / "build.json")
    expected_binaries = [
        {key: entry["artifact"].get(key) for key in ("path", "bytes", "sha256")}
        for entry in build_json["binaries"].values()
    ]
    require(all(isinstance(item, dict) and set(item) == {"path", "bytes", "sha256"}
                for item in removed_binaries),
            "cleanup binary descriptor shape changed")
    require(all(item in removed_binaries for item in expected_binaries)
            and all(item in expected_binaries for item in removed_binaries),
            "cleanup binary witness does not exactly match build")
    source = value.get("source")
    require(source == build_json.get("source"),
            "cleanup source witness differs from build source")
    removed = value.get("removed")
    require(isinstance(removed, list) and len(removed) == 2,
            "cleanup directory witness cardinality changed")
    removed_paths = {row.get("path") for row in removed if isinstance(row, dict)}
    require(removed_paths == {str(analyze.custody.TARGET), str(analyze.custody.SCRATCH)},
            "cleanup directory witness changed")
    require(not analyze.custody.TARGET.exists() and not analyze.custody.SCRATCH.exists(),
            "owned build/scratch directories remain after cleanup")
    analyze.finite_number(value.get("started"), "cleanup start time")
    analyze.finite_number(value.get("ended"), "cleanup end time")
    require(value["started"] <= value["ended"], "cleanup timestamps changed")


def check_seal() -> None:
    path = PACKET / "seal.json"
    require(path.is_file() and not path.is_symlink(), "final seal is missing")
    value = analyze.read_json(path)
    require(value.get("schema") == "litchi.performance.0817.seal.v1",
            "final seal schema changed")
    files = value.get("files")
    require(isinstance(files, dict) and files, "final seal files are missing")
    for name, digest in files.items():
        require(isinstance(name, str) and not Path(name).is_absolute()
                and ".." not in Path(name).parts and analyze.is_sha(digest),
                f"final seal path/digest is malformed: {name!r}")
        target = (ROOT / name).resolve(strict=False)
        require(target.is_relative_to(ROOT.resolve())
                and target.is_file() and not target.is_symlink()
                and analyze.sha256(target) == digest,
                f"final seal payload changed: {name}")


def validate(*, require_cleanup: bool = False,
             require_final_seal: bool = False) -> dict[str, Any]:
    require(not require_final_seal or require_cleanup,
            "final seal validation requires the final cleanup witness")
    result = analyze.analyze(check=True)
    require(result["schema"] == analyze.ANALYSIS_SCHEMA, "analysis schema changed")
    require(result["plan_schema"] == "litchi.performance.0817.plan.v1",
            "plan schema changed")
    require(result.get("historical_timing_comparison") is False
            and result.get("regression_policy") ==
            "baseline only; no candidate adoption or historical timing comparisons",
            "historical timing comparison policy changed")
    require(result["report_schema_version"] == analyze.REPORT_SCHEMA_VERSION,
            "report schema version changed")
    require(result["bootstrap"].get("status") == "not_run"
            and result["bootstrap"].get("seed") == analyze.BOOTSTRAP_SEED
            and result["bootstrap"].get("resamples") == analyze.BOOTSTRAP_RESAMPLES,
            "bootstrap status changed")
    check_counts(result)
    check_source_custody(result)
    check_build_recovery(result)
    check_quality_attempt_acknowledgement(result)
    check_admission_failure(result)
    check_no_timing(result)
    verification = result["verification"]
    for key in (
        "packet_custody_checked", "production_source_checked", "harness_source_checked",
        "runtime_harness_source_checked", "allowlisted_test_source_checked",
        "quality_checked", "quality_recovery_checked", "build_checked",
        "build_driver_recovery_checked", "artifact_export_checked",
        "independent_artifact_oracle_checked", "artifact_admission_checked",
        "qualification_admission_checked", "report_schema_checked",
        "binary_identity_checked", "corpus_identity_checked", "outcome_identity_checked",
        "sample_cardinality_checked", "timing_absence_checked",
        "native_timing_separated", "historical_comparison_omitted",
    ):
        require(verification.get(key) is True, f"verification marker missing: {key}")
    require(verification.get("observer_diagnostics_retained") is False
            and verification.get("bootstrap_checked") is False
            and verification.get("logical_workload_descriptors_checked") is False,
            "unrun timing evidence was marked as retained or computed")
    cleanup_checked = False
    if require_cleanup:
        check_cleanup(result)
        cleanup_checked = True
    seal_checked = False
    seal_path = PACKET / "seal.json"
    if require_final_seal or seal_path.exists():
        check_seal()
        seal_checked = True
    return {
        "reports": result["counts"]["reports"],
        "samples": result["counts"]["samples"],
        "native_reports": result["counts"]["native_reports"],
        "observer_reports": result["counts"]["observer_reports"],
        "qualification_reports": result["counts"]["qualification_reports"],
        "quality_attempt": result["quality"]["attempt"],
        "native_rows": len(result["native"]),
        "observer_rows": len(result["observer"]["reports"]),
        "admission_status": result["admissions"]["accepted"],
        "cleanup_checked": cleanup_checked,
        "seal_checked": seal_checked,
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
    except (analyze.ReplayError, AssertionError, OSError, ValueError, KeyError,
            TypeError, IndexError) as error:
        print(f"0817 validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
