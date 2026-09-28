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
        "reports": 108,
        "samples": 2244,
        "native_reports": 72,
        "native_samples": 2160,
        "observer_reports": 24,
        "observer_samples": 72,
        "qualification_reports": 12,
        "qualification_samples": 12,
    }, "aggregate report/sample cardinality changed")
    require(result["artifacts"] == {
        "complete": result["artifacts"]["complete"],
        "manifest": result["artifacts"]["manifest"],
        "receipt": result["artifacts"]["receipt"],
        "audit": result["artifacts"]["audit"],
        "cases": 6,
        "policy_outputs": 30,
        "audit_cases": 6,
    }, "artifact cardinality changed")


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


def check_native(result: dict[str, Any]) -> None:
    native = result["native"]
    require(isinstance(native, list) and len(native) == 12,
            "native summary cardinality changed")
    require([row["case"] for row in native] ==
            [case["case"] for case in analyze.plan()["cases"]],
            "native summary order changed")
    expected_cases = {case["case"] for case in result["raw_reports"]["native"]}
    require({row["case"] for row in native} == expected_cases,
            "native summary selector set changed")
    require(result.get("quantile_definitions") == {
        "analysis": "nearest rank within each process; median over six process blocks",
        "raw_report": "harness p50 integer midpoint; p95/p99 nearest rank",
    }, "quantile definitions changed")
    for row in native:
        require(row["blocks"] == 6 and row["samples_per_report"] == 30,
                f"native lane dimensions changed: {row['case']}")
        require(row["historical_timing_comparison"] is False,
                f"native historical comparison appeared: {row['case']}")
        require(len(row["block_raw_quantiles"]) == 6,
                f"native raw block quantiles changed: {row['case']}")
        require(row["quantile_definition"] ==
                "nearest rank within each process; median over six process blocks"
                and row["raw_report_quantile_definition"] ==
                "harness p50 integer midpoint; p95/p99 nearest rank",
                f"native quantile definitions changed: {row['case']}")
        for block in row["block_raw_quantiles"]:
            require(block["quantile_definition"] ==
                    "nearest rank within this process block"
                    and block["raw_report_quantile_definition"] ==
                    "harness p50 integer midpoint; p95/p99 nearest rank"
                    and isinstance(block.get("raw_reported"), dict),
                    f"native raw quantile witness changed: {row['case']}")
            samples = block["samples"]
            raw = block["raw_reported"]
            require(raw.get("p50") == analyze.integer_midpoint(samples[14], samples[15])
                    and raw.get("p95") == analyze.nearest_rank(samples, 0.95)
                    and raw.get("p99") == analyze.nearest_rank(samples, 0.99)
                    and raw.get("mean") == analyze.statistics.fmean(samples),
                    f"native raw report aggregate changed: {row['case']}")
        require(row["spread_flag"] == bool(row["spread_flags"]),
                f"native spread flag changed: {row['case']}")
        require(row["spread_flags"] == [name for name in ("p50", "p95", "p99", "mean")
                                         if row["spread_ratios"][name] > 0.05],
                f"native spread flag list changed: {row['case']}")
        require(row["tail_flag"] == (row["p99_to_p50_ratio"] > 1.05),
                f"native tail flag changed: {row['case']}")
        require(row["p50_ns"] == analyze.statistics.median(row["p50_block_values_ns"])
                and row["p95_ns"] == analyze.statistics.median(row["p95_block_values_ns"])
                and row["p99_ns"] == analyze.statistics.median(row["p99_block_values_ns"])
                and row["mean_ns"] == analyze.statistics.median(row["mean_block_values_ns"]),
                f"native block medians changed: {row['case']}")
        for field in ("p50", "p95", "p99", "mean"):
            values = row[f"{field}_block_values_ns"]
            require(row["spread_ratios"][field] ==
                    max(values) / min(values) - 1.0,
                    f"native spread ratio changed: {row['case']}/{field}")
        require(row["p99_to_p50_ratio"] ==
                analyze.statistics.median(row["p99_block_values_ns"]) /
                analyze.statistics.median(row["p50_block_values_ns"]),
                f"native tail ratio changed: {row['case']}")
        require(row["p50_bootstrap_ci95_ns"] ==
                analyze.bootstrap_absolute(row["p50_block_values_ns"]),
                f"native bootstrap changed: {row['case']}")
        workload = row["workload"]
        require(workload["claim"].startswith("descriptive logical workload rate only"),
                f"native workload claim changed: {row['case']}")
        for key in ("source_logical_bytes_per_second", "published_logical_bytes_per_second"):
            analyze.finite_number(workload[key], f"{row['case']}.{key}")
            require(workload[key] > 0, f"{row['case']}.{key} is not positive")


def check_observer(result: dict[str, Any]) -> None:
    observer = result["observer"]
    require(observer.get("timings_pooled_with_native") is False
            and observer.get("operation_metrics_subtracted") is False
            and observer.get("procfs_controls_subtracted") is False
            and observer.get("latency_claim") ==
            "diagnostic_only; observer elapsed values are not latency evidence",
            "observer diagnostic separation changed")
    require(isinstance(observer.get("reports"), list)
            and len(observer["reports"]) == 36,
            "observer/qualification raw report cardinality changed")
    for row in observer["reports"]:
        require(row["lane"] in {"observer", "qualification"}
                and row["process_probe"] is not None,
                "observer diagnostic report is missing procfs evidence")
        require(row["operation_metrics"]["allocation"]["status"] == "measured"
                and row["operation_metrics"]["process"]["status"] == "measured",
                "observer operation metrics status changed")
        workload = row["workload"]
        require(workload["claim"].startswith("descriptive logical workload rate only"),
                "observer workload claim changed")


def check_raw_reports(result: dict[str, Any]) -> None:
    raw = result["raw_reports"]
    require(len(raw["native"]) == 72 and len(raw["observer"]) == 24
            and len(raw["qualification"]) == 12,
            "raw report lane cardinality changed")
    for lane, rows in raw.items():
        for row in rows:
            report = row["report"]
            path = analyze.resolve_path(report["path"], packet_bound=True)
            require(report["sha256"] == analyze.sha256(path),
                    f"retained {lane} report digest changed")
            require(len(row["samples"]) == {
                "native": 30, "observer": 3, "qualification": 1,
            }[lane], f"retained {lane} sample cardinality changed")
            require(row["workload"]["source_logical_bytes"] > 0
                    and row["workload"]["published_logical_bytes"] > 0,
                    f"retained {lane} workload bytes changed")


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
    require(result["bootstrap"] == {
        "seed": analyze.BOOTSTRAP_SEED,
        "resamples": analyze.BOOTSTRAP_RESAMPLES,
        "confidence": analyze.BOOTSTRAP_CONFIDENCE,
        "low_rank": analyze.BOOTSTRAP_LOW_RANK,
        "high_rank": analyze.BOOTSTRAP_HIGH_RANK,
        "statistic": "median of six process p50 values",
        "units": "ns",
    }, "bootstrap contract changed")
    check_counts(result)
    check_source_custody(result)
    check_build_recovery(result)
    check_quality_attempt_acknowledgement(result)
    check_native(result)
    check_observer(result)
    check_raw_reports(result)
    verification = result["verification"]
    for key in (
        "packet_custody_checked", "production_source_checked", "harness_source_checked",
        "runtime_harness_source_checked", "allowlisted_test_source_checked",
        "quality_checked", "quality_recovery_checked", "build_checked",
        "build_driver_recovery_checked", "artifact_export_checked",
        "independent_artifact_oracle_checked", "artifact_admission_checked",
        "qualification_admission_checked", "report_schema_checked",
        "binary_identity_checked", "corpus_identity_checked", "outcome_identity_checked",
        "sample_cardinality_checked", "observer_diagnostics_retained",
        "native_timing_separated", "historical_comparison_omitted", "bootstrap_checked",
        "logical_workload_descriptors_checked",
    ):
        require(verification.get(key) is True, f"verification marker missing: {key}")
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
