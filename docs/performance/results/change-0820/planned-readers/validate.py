"""Fail-closed offline validator for the 0820 save-durability packet."""

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


def check_counts(value: dict[str, Any]) -> None:
    require(value.get("counts") == {
        "reports": 216, "samples": 4488,
        "native_reports": 144, "native_samples": 4320,
        "observer_reports": 48, "observer_samples": 144,
        "qualification_reports": 24, "qualification_samples": 24,
        "artifact_cases": 6, "artifact_policy_outputs": 30,
    }, "aggregate counts changed")
    require(value.get("status") == "accepted"
            and value.get("timing_status", {}).get("accepted") is True,
            "timing admission status changed")


def check_policy(value: dict[str, Any]) -> None:
    require(value.get("historical_timing_comparison") is False
            and value.get("regression_policy") ==
            "configuration attribution only; no production optimization, historical timing comparisons, or weakened default adoption",
            "historical comparison policy changed")
    require(value.get("comparison") ==
            "within-block policy/default ratios for the same format and phase; full/default controls explicit API route; weaker-policy differences change durability semantics"
            and value.get("policies") == list(analyze.POLICIES)
            and value.get("policy_semantics") == analyze.POLICY_SEMANTICS,
            "durability comparison policy changed")
    bootstrap = value.get("bootstrap")
    require(isinstance(bootstrap, dict)
            and bootstrap.get("status") == "computed"
            and bootstrap.get("seed") == analyze.BOOTSTRAP_SEED
            and bootstrap.get("resamples") == analyze.BOOTSTRAP_RESAMPLES
            and bootstrap.get("low_rank") == analyze.BOOTSTRAP_LOW_RANK
            and bootstrap.get("high_rank") == analyze.BOOTSTRAP_HIGH_RANK
            and bootstrap.get("paired_statistic") ==
            "median of six matched process-block policy/default p50 ratios",
            "bootstrap contract changed")
    observer = value.get("observer")
    require(isinstance(observer, dict)
            and observer.get("timings_pooled_with_native") is False
            and observer.get("operation_metrics_subtracted") is False
            and observer.get("procfs_controls_subtracted") is False,
            "observer separation changed")


def check_source(value: dict[str, Any]) -> None:
    custody = value.get("custody")
    require(isinstance(custody, dict)
            and custody.get("production_file_count") == analyze.PRODUCTION_FILES
            and custody.get("production_matches_base") is True
            and custody.get("harness_matches_base") is True
            and custody.get("runtime_harness_matches_base") is True
            and custody.get("tool_allowlist") == [],
            "source custody changed")


def check_native(value: dict[str, Any]) -> None:
    rows = value.get("native")
    require(isinstance(rows, list) and len(rows) == 24, "native summary cardinality changed")
    seen = set()
    for row in rows:
        require(row.get("id") not in seen and row.get("policy") in analyze.POLICIES,
                f"native selector identity changed: {row.get('id')}")
        seen.add(row.get("id"))
        require(row.get("blocks") == 6 and row.get("samples_per_report") == 30
                and row.get("historical_timing_comparison") is False
                and row.get("quantile_definition") ==
                "nearest rank within each process; median over six process blocks"
                and row.get("raw_report_quantile_definition") ==
                "harness p50 integer midpoint; p95/p99 nearest rank",
                f"native summary metadata changed: {row.get('id')}")
        require(len(row.get("block_raw_quantiles", [])) == 6,
                f"native block witnesses changed: {row.get('id')}")
        for block in row["block_raw_quantiles"]:
            samples = block["samples"]
            reported = block["raw_reported"]
            require(reported["p50"] == analyze.integer_midpoint(samples[14], samples[15])
                    and reported["p95"] == analyze.nearest_rank(samples, .95)
                    and reported["p99"] == analyze.nearest_rank(samples, .99)
                    and abs(float(reported["mean"]) - analyze.rust_welford_mean(samples)) < 1e-12,
                    f"native raw quantiles changed: {row.get('id')}")
        spread = row["spread_ratios"]
        require(row["spread_flag"] == bool(row["spread_flags"])
                and row["spread_flags"] == [key for key in ("p50", "p95", "p99", "mean")
                                             if spread[key] > 1.05]
                and row["tail_flag"] == (row["p99_to_p50_ratio"] > 1.05),
                f"native diagnostic flags changed: {row.get('id')}")
        require(row["p50_bootstrap_ci95_ns"] == analyze.bootstrap(row["p50_block_values_ns"]),
                f"native bootstrap changed: {row.get('id')}")
        paired = row.get("paired_ratio_to_default")
        require(isinstance(paired, dict)
                and paired.get("control_policy") == "default"
                and paired.get("control_id") == f"{row['case']}__default"
                and paired.get("definition") ==
                "matched process-block policy p50 / default p50 for the same format and phase"
                and paired.get("configuration_attribution_only") is True
                and paired.get("default_full_control") == (row["policy"] in {"default", "full"})
                and paired.get("weaker_policy_semantics_changed") ==
                (row["policy"] in {"file-only", "no-sync"}),
                f"native paired comparison metadata changed: {row.get('id')}")
        ratios = paired.get("block_p50_ratios")
        require(isinstance(ratios, list) and len(ratios) == 6
                and all(isinstance(x, (int, float)) and x > 0 for x in ratios),
                f"native paired ratios changed: {row.get('id')}")
        expected_ratio = analyze.paired_bootstrap(ratios)
        require(paired.get("bootstrap_ci95") == expected_ratio
                and paired.get("ratio_median") == expected_ratio["estimate"],
                f"native paired bootstrap changed: {row.get('id')}")
        workload = row.get("workload")
        require(isinstance(workload, dict)
                and workload.get("claim", "").startswith("descriptive logical workload rate only"),
                f"native workload scope changed: {row.get('id')}")
    require(seen == {case["id"] for case in analyze.plan()["cases"]},
            "native selector set changed")


def check_observer(value: dict[str, Any]) -> None:
    reports = value.get("observer", {}).get("reports")
    require(isinstance(reports, list) and len(reports) == 72,
            "observer/qualification report cardinality changed")
    lane_counts = {"observer": 0, "qualification": 0}
    for row in reports:
        require(row.get("lane") in {"observer", "qualification"}
                and row.get("policy") in analyze.POLICIES
                and isinstance(row.get("process_probe"), dict),
                "observer probe evidence missing")
        lane_counts[row["lane"]] += 1
        metrics = row.get("operation_metrics", {})
        require(metrics.get("allocation", {}).get("status") == "measured"
                and metrics.get("process", {}).get("status") == "measured",
                "observer operation metrics changed")
    require(lane_counts == {"observer": 48, "qualification": 24},
            "observer/qualification lane cardinality changed")


def check_cleanup(value: dict[str, Any]) -> None:
    path = PACKET / "cleanup.json"
    require(path.is_file() and not path.is_symlink(), "cleanup witness is missing")
    cleanup = analyze.read_json(path)
    require(cleanup.get("schema") == "litchi.performance.0820.cleanup.v1"
            and cleanup.get("target_removed") is True
            and cleanup.get("scratch_removed") is True
            and cleanup.get("binaries_verified_before_removal") is True,
            "cleanup witness changed")
    build = analyze.read_json(PACKET / "build.json")
    expected = [{key: row["artifact"][key] for key in ("path", "bytes", "sha256")}
                for row in build["binaries"].values()]
    observed = cleanup.get("removed_binaries")
    require(isinstance(observed, list)
            and observed == sorted(observed, key=lambda row: row["path"])
            and observed == sorted(expected, key=lambda row: row["path"]),
            "cleanup binary descriptors do not match build")
    removed = cleanup.get("removed")
    require(isinstance(removed, list) and len(removed) == 2
            and all(isinstance(row, dict)
                    and set(row) == {"path", "files", "logical_bytes"}
                    and isinstance(row.get("path"), str)
                    and analyze.nonnegative_int(row.get("files"), "cleanup file count") is None
                    and analyze.nonnegative_int(row.get("logical_bytes"), "cleanup logical bytes") is None
                    for row in removed)
            and {row["path"] for row in removed} ==
            {str(analyze.custody.TARGET), str(analyze.custody.SCRATCH)},
            "cleanup directory rows changed")
    source = cleanup.get("source")
    build_source = build.get("source")
    analyze.same_descriptor(source, build_source, "cleanup source witness")
    require(not analyze.custody.TARGET.exists() and not analyze.custody.SCRATCH.exists(),
            "owned build/scratch directories remain")


def check_seal() -> None:
    path = PACKET / "seal.json"
    require(path.is_file() and not path.is_symlink(), "seal witness is missing")
    value = analyze.read_json(path)
    require(value.get("schema") == "litchi.performance.0820.seal.v1",
            "seal schema changed")
    files = value.get("files")
    require(isinstance(files, dict) and files, "seal file map is missing")
    for name, digest in files.items():
        relative = Path(name)
        require(not relative.is_absolute() and ".." not in relative.parts
                and analyze.is_sha(digest), f"malformed seal entry: {name}")
        path = (ROOT / relative).resolve(strict=False)
        require(path.is_relative_to(ROOT.resolve()) and path.is_file()
                and not path.is_symlink() and analyze.sha256(path) == digest,
                f"sealed payload changed: {name}")


def validate(*, final: bool = False) -> dict[str, Any]:
    value = analyze.analyze(check=True)
    require(value.get("schema") == analyze.ANALYSIS_SCHEMA
            and value.get("plan_schema") == analyze.PLAN_SCHEMA
            and value.get("report_schema_version") == analyze.REPORT_SCHEMA_VERSION,
            "analysis schema changed")
    check_counts(value)
    check_policy(value)
    check_source(value)
    check_native(value)
    check_observer(value)
    verification = value.get("verification", {})
    for marker in ("packet_custody_checked", "production_source_checked",
                   "harness_source_checked", "runtime_harness_source_checked",
                   "quality_checked", "build_checked", "artifact_export_checked",
                   "independent_artifact_oracle_checked", "zip_preservation_checked",
                   "artifact_admission_checked", "qualification_admission_checked",
                   "report_schema_checked", "binary_identity_checked",
                   "corpus_identity_checked", "outcome_identity_checked",
                   "sample_cardinality_checked", "observer_diagnostics_retained",
                   "native_timing_separated", "historical_comparison_omitted",
                   "bootstrap_checked", "logical_workload_descriptors_checked",
                   "policy_route_checked", "paired_comparisons_checked",
                   "policy_semantics_retained"):
        require(verification.get(marker) is True, f"verification marker missing: {marker}")
    cleanup_checked = False
    if final:
        check_cleanup(value)
        cleanup_checked = True
    seal_checked = False
    # seal.py --write calls this validator before it can create seal.json.
    # Once the seal exists, replay it on every validation invocation.
    if (PACKET / "seal.json").is_file():
        check_seal()
        seal_checked = True
    return {"reports": value["counts"]["reports"], "samples": value["counts"]["samples"],
            "native_reports": value["counts"]["native_reports"],
            "observer_reports": value["counts"]["observer_reports"],
            "qualification_reports": value["counts"]["qualification_reports"],
            "cleanup_checked": cleanup_checked, "seal_checked": seal_checked}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--final", action="store_true",
                        help="require cleanup; replay seal.json when present")
    args = parser.parse_args(argv)
    try:
        print(json.dumps(validate(final=args.final), indent=2, sort_keys=True))
    except (analyze.ReplayError, AssertionError, OSError, ValueError, KeyError,
            TypeError, IndexError) as error:
        print(f"0820 validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
