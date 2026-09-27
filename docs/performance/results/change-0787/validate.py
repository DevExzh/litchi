"""Fail-closed offline validator for the 0787 paired replay packet."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from typing import Any

import analyze


PACKET = analyze.PACKET
EXPECTED_COUNTS = {
    "reports": 1080,
    "samples": 22200,
    "native_reports": 720,
    "native_samples": 21600,
    "observer_reports": 240,
    "observer_samples": 480,
    "qualification_reports": 120,
    "qualification_samples": 120,
}
PAIR_METRICS = ("p50", "p95", "p99", "rss", "cpu_wall_ratio", "throughput_bytes_s")


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def _expected_keys() -> list[tuple[Any, ...]]:
    return [analyze._case_key(case) for case in analyze.expected_cases()]


def _validate_bootstrap(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label} bootstrap is missing")
    require(value.get("seed") == analyze.BOOTSTRAP_SEED
            and value.get("resamples") == analyze.BOOTSTRAP_RESAMPLES
            and value.get("confidence") == analyze.BOOTSTRAP_CONFIDENCE
            and value.get("statistic") == "median"
            and value.get("endpoint_indexes") == [249, 9749],
            f"{label} bootstrap contract changed")


def _validate_metric(metric: Any, label: str) -> None:
    require(isinstance(metric, dict), f"{label} metric is missing")
    for field in ("estimate", "ci95_low", "ci95_high", "block_spread_ratio"):
        analyze.finite_number(metric.get(field), f"{label}.{field}")
    before = metric.get("before_block_values")
    after = metric.get("after_block_values")
    ratios = metric.get("block_ratios")
    flags = metric.get("block_flags")
    require(isinstance(before, list) and len(before) == analyze.NATIVE_BLOCKS,
            f"{label} before block values changed")
    require(isinstance(after, list) and len(after) == analyze.NATIVE_BLOCKS,
            f"{label} after block values changed")
    require(isinstance(ratios, list) and len(ratios) == analyze.NATIVE_BLOCKS,
            f"{label} block ratios changed")
    require(isinstance(flags, list) and len(flags) == analyze.NATIVE_BLOCKS,
            f"{label} block flags changed")
    expected_flags = []
    for index, (old, new, ratio, flag) in enumerate(zip(before, after, ratios, flags)):
        analyze.finite_number(old, f"{label}.before[{index}]")
        analyze.finite_number(new, f"{label}.after[{index}]")
        analyze.finite_number(ratio, f"{label}.ratio[{index}]")
        require(ratio == new / old, f"{label} ratio does not replay at block {index}")
        require(isinstance(flag, dict) and flag.get("block") == index
                and flag.get("ratio") == ratio
                and flag.get("delta") == ratio - 1.0
                and flag.get("absolute_over_5_percent") is (abs(ratio - 1.0) > 0.05)
                and flag.get("regression_over_5_percent") is (ratio - 1.0 > 0.05),
                f"{label} block flag changed at {index}")
        expected_flags.append(flag)
    expected_spread = max(ratios) / min(ratios) - 1.0 if min(ratios) > 0 else None
    require(metric.get("block_spread_ratio") == expected_spread,
            f"{label} block spread changed")
    require(metric.get("over_5_percent_blocks") == [f["block"] for f in expected_flags
                                                     if f["absolute_over_5_percent"]]
            and metric.get("regression_over_5_percent_blocks") ==
            [f["block"] for f in expected_flags if f["regression_over_5_percent"]],
            f"{label} >5% flags changed")
    _validate_bootstrap(metric.get("bootstrap"), label)
    boot = metric["bootstrap"]
    require(metric["estimate"] == boot["estimate"]
            and metric["ci95_low"] == boot["lower"]
            and metric["ci95_high"] == boot["upper"],
            f"{label} bootstrap projection changed")


def _validate_raw_audit(result: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "raw-audit.json"
    require(path.is_file() and not path.is_symlink(), "independent raw audit is missing")
    raw = analyze.read_json(path)
    require(isinstance(raw, dict)
            and raw.get("reports") == EXPECTED_COUNTS["reports"]
            and raw.get("samples") == EXPECTED_COUNTS["samples"]
            and raw.get("independently_reconstructed_payloads") == 96,
            "independent raw audit counts changed")
    rows = raw.get("rows")
    require(isinstance(rows, list) and len(rows) == 60, "independent raw audit row count changed")
    paired = {tuple(row[key] for key in ("route", "shape", "state", "task_floor", "workers")): row
              for row in result["paired"]}
    require(len(paired) == 60, "paired case keys changed before raw audit comparison")
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and isinstance(row.get("case"), list)
                and len(row["case"]) == 5, f"raw audit row {index} is malformed")
        key = tuple(row["case"])
        require(key in paired, f"raw audit row {index} has an unknown case")
        metrics = row.get("metrics")
        require(isinstance(metrics, dict), f"raw audit row {index} metrics are missing")
        for name in ("p50", "p95", "p99", "rss"):
            raw_metric = metrics.get(name)
            derived = paired[key]["metrics"][name]
            require(isinstance(raw_metric, dict)
                    and raw_metric.get("paired_ratios") == derived["block_ratios"]
                    and raw_metric.get("median") == derived["estimate"]
                    and raw_metric.get("ci95") == [derived["ci95_low"], derived["ci95_high"]],
                    f"raw audit {key} {name} mismatch")
    raw_rejections = {(tuple(row["case"]), row["metric"])
                      for row in raw.get("rejections", [])}
    derived_rejections = {(tuple(row[key] for key in
                                  ("route", "shape", "state", "task_floor", "workers")),
                           row["kind"])
                          for row in result["adoption"]["latency_or_rss_rejections"]}
    require(raw_rejections == derived_rejections,
            "raw audit rejection decision changed")
    benefit_cases = {
        tuple(row[key] for key in ("route", "shape", "state", "task_floor", "workers"))
        for row in result["adoption"]["material_benefit_rows"]
    }
    require({tuple(case) for case in raw.get("benefits", [])} == benefit_cases,
            "raw audit benefit decision changed")
    require(raw.get("candidate_eligible_for_retention") ==
            result["adoption"]["candidate_eligible_for_retention"],
            "raw audit adoption decision changed")
    corpora = raw.get("corpora")
    identities = raw.get("container_identities")
    require(isinstance(corpora, dict) and set(corpora) == set(analyze.SHAPES)
            and isinstance(identities, dict) and set(identities) == set(analyze.SHAPES),
            "raw audit corpus identity set changed")
    return {"path": analyze.rel(path), "sha256": analyze.sha256(path),
            "reports": raw["reports"], "samples": raw["samples"],
            "independently_reconstructed_payloads": raw["independently_reconstructed_payloads"],
            "rows": len(rows), "paired_curves_match": True}


def _validate_paired(result: dict[str, Any]) -> None:
    paired = result.get("paired")
    require(isinstance(paired, list) and len(paired) == 60, "paired row cardinality changed")
    keys = [(row.get("route"), row.get("shape"), row.get("state"),
             row.get("task_floor"), row.get("workers")) for row in paired]
    require(keys == _expected_keys(), "paired case order or matrix changed")
    for row in paired:
        require(row.get("blocks") == analyze.NATIVE_BLOCKS
                and row.get("samples_per_report") == analyze.NATIVE_SAMPLES,
                "paired block/sample contract changed")
        metrics = row.get("metrics")
        require(isinstance(metrics, dict), "paired metric map is missing")
        for name in PAIR_METRICS:
            _validate_metric(metrics.get(name),
                             f"{row['shape']}/{row['state']}/{row['task_floor']}/{row['workers']} {name}")
        require(row.get("p50") == metrics["p50"]
                and row.get("p95") == metrics["p95"]
                and row.get("p99") == metrics["p99"]
                and row.get("rss") == metrics["rss"],
                "paired metric projection changed")


def _validate_reports(result: dict[str, Any]) -> None:
    require(isinstance(result.get("reports"), list)
            and len(result["reports"]) == EXPECTED_COUNTS["reports"],
            "retained report cardinality changed")
    for row in result["reports"]:
        require(row.get("verification_ok") is True and row.get("resource_ok") is True,
                "report correctness witness changed")
        report = row.get("report")
        require(isinstance(report, dict) and isinstance(report.get("path"), str)
                and report.get("sha256") == analyze.sha256(analyze.resolve_path(report["path"])),
                "retained report digest changed")
        require(row.get("samples") == len(row.get("wall_ns", [])),
                "report sample timing cardinality changed")
    observer = result["observer"]
    require(observer.get("timings_pooled_with_native") is False
            and observer.get("reports") == 360
            and len(observer.get("rows", [])) == 360
            and observer.get("max_active_parity") ==
            "scheduler-dependent; retained per leg and bounded by workers",
            "observer separation/cardinality changed")
    for row in observer["rows"]:
        require(row.get("available") is True
                and row.get("calls") == (64 if row.get("state") == "fresh" else 0)
                and isinstance(row.get("max_active"), (int, float))
                and row["max_active"] <= row["workers"],
                "observer source call-count contract changed")


def _validate_seal(*, required: bool) -> bool:
    if not required:
        return False
    seal = PACKET / "seal.json"
    require(seal.is_file() and not seal.is_symlink(), "final seal is missing")
    require(not (PACKET / "final-seal.json").exists(), "ambiguous second final seal")
    value = analyze.read_json(seal)
    require(isinstance(value, dict)
            and value.get("schema") == "litchi.execution-scaling-seal.v1",
            "final seal schema changed")
    files = value.get("files")
    require(isinstance(files, dict) and files, "final seal file set is missing")
    actual: dict[str, str] = {}
    for path in PACKET.rglob("*"):
        if path == seal or "__pycache__" in path.parts:
            continue
        require(not path.is_symlink(), f"symlink in sealed packet: {path}")
        if path.is_file():
            actual[str(path.relative_to(PACKET))] = analyze.sha256(path)
    require(actual == files, "final seal file set or payload hash mismatch")
    require(value.get("payload_count") == len(actual), "final seal payload count changed")
    return True


def _current_tracked_production_files() -> dict[str, str]:
    """Hash the production files tracked by the current checkout.

    The final disposition must remain valid if the packet is moved to another
    worktree or the checkout receives a new commit.  Consequently this uses
    the current tracked path set and compares bytes, rather than comparing a
    revision identifier with the historical restoration witness.
    """
    try:
        completed = subprocess.run(
            ["git", "ls-files", "-z", "--", *analyze.PRODUCTION_PATHSPEC],
            cwd=analyze.ROOT,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        analyze.fail(f"cannot read current tracked production files: {error}")
    names = [item.decode() for item in completed.stdout.split(b"\0") if item]
    require(len(names) == analyze.EXPECTED_PRODUCTION_FILES,
            f"current tracked production file count changed: {len(names)}")
    require(len(set(names)) == len(names),
            "current tracked production files contain duplicates")
    result: dict[str, str] = {}
    for name in names:
        path = analyze.ROOT / name
        require(path.is_file() and not path.is_symlink(),
                f"current tracked production file is missing or symlinked: {name}")
        result[name] = analyze.sha256(path)
    return result


def _validate_final_disposition(result: dict[str, Any]) -> dict[str, Any]:
    """Validate the sealed reject/retain decision and restored source witness."""
    disposition_path = analyze.PACKET / "disposition.json"
    require(disposition_path.is_file() and not disposition_path.is_symlink(),
            "final disposition is missing")
    disposition = analyze.read_json(disposition_path)
    require(isinstance(disposition, dict)
            and disposition.get("schema") == "litchi.cached-part-disposition.0787.v1",
            "final disposition schema changed")

    eligible = result["adoption"].get("candidate_eligible_for_retention")
    require(isinstance(eligible, bool), "analysis retention decision is malformed")
    expected_decision = "retain" if eligible else "reject"
    require(disposition.get("decision") == expected_decision,
            "final disposition decision disagrees with analysis")
    require(disposition.get("guard_failures") ==
            result["adoption"].get("latency_or_rss_rejections"),
            "final disposition guard failures disagree with analysis")
    require(disposition.get("production_byte_identical_to_base") is True
            and disposition.get("production_files") == analyze.EXPECTED_PRODUCTION_FILES,
            "final disposition production restoration witness changed")
    require(disposition.get("candidate_retained_only_in_archive") is True,
            "final disposition candidate archive custody changed")

    artifact_specs = {
        "analysis": (analyze.PACKET / "analysis.json", "disposition analysis artifact"),
        "raw_audit": (analyze.PACKET / "raw-audit.json", "disposition raw audit artifact"),
        "restored_source": (analyze.PACKET / "restored-source.json",
                             "disposition restored source artifact"),
    }
    checked: dict[str, dict[str, Any]] = {}
    for key, (expected_path, label) in artifact_specs.items():
        actual_path = analyze.artifact_path(disposition.get(key), label)
        require(actual_path.resolve() == expected_path.resolve(),
                f"{label} points at the wrong packet file")
        checked[key] = analyze.file_identity(actual_path)

    restored = analyze.read_json(artifact_specs["restored_source"][0])
    restored_manifest = analyze.load_source_manifest(restored, "restored source")
    require(len(restored_manifest["files"]) == analyze.EXPECTED_PRODUCTION_FILES,
            "restored source file count changed")
    baseline = result["source"]["before"]["production"]["files"]
    require(restored_manifest["files"] == baseline,
            "restored source differs from the baseline source hashes")

    current = _current_tracked_production_files()
    require(current == restored_manifest["files"],
            "current tracked production differs from restored source hashes")
    return {
        "path": analyze.rel(disposition_path),
        "sha256": analyze.sha256(disposition_path),
        "decision": disposition["decision"],
        "restored_production_files": len(current),
        "artifacts": checked,
        "baseline_hashes_match": True,
        "current_tracked_hashes_match": True,
    }


def validate(*, require_final_seal: bool = False) -> dict[str, Any]:
    result = analyze.analyze_0787(check=True)
    require(result.get("schema") == analyze.ANALYSIS_SCHEMA, "analysis schema changed")
    require(result.get("plan_schema") == "litchi.cached-part-scheduling.0787.v1",
            "plan schema changed")
    require(result.get("report_schema") == analyze.REPORT_SCHEMA, "report schema changed")
    require(result.get("counts") == EXPECTED_COUNTS, "aggregate report/sample cardinality changed")
    require(result["quality"].get("gates") == 6, "quality gate cardinality changed")
    source = result["source"]
    require(source.get("production_file_count") == 9196
            and source.get("before_production_byte_identical_to_base") is True
            and source.get("after_candidate_diff_bound_to_archive") is True
            and source.get("tool_source_frozen") is True,
            "source custody changed")
    archive = source.get("candidate_archive")
    require(isinstance(archive, dict) and archive.get("changed_file_count") == 3,
            "candidate archive custody changed")
    architecture = result["architecture_inputs"]
    require(architecture.get("count") == 35
            and architecture.get("revision") == analyze.origin()["base"]
            and architecture.get("base_git_blob_hashes_match") is True,
            "architecture input custody changed")
    _validate_bootstrap(result.get("bootstrap"), "analysis")
    _validate_paired(result)
    _validate_reports(result)
    policy = result["adoption"]
    require(policy.get("cpu_is_descriptive_only") is True,
            "CPU adoption interpretation changed")
    interpretation = result["policy_interpretation"]
    require(interpretation.get("rejection_is_conjunctive") is True
            and interpretation.get("benefit_is_conjunctive") is True,
            "frozen policy interpretation changed")
    raw = _validate_raw_audit(result)
    seal_checked = _validate_seal(required=require_final_seal)
    disposition = _validate_final_disposition(result) if seal_checked else None
    return {
        "reports": result["counts"]["reports"],
        "samples": result["counts"]["samples"],
        "paired_cases": len(result["paired"]),
        "observer_reports": result["counts"]["observer_reports"],
        "qualification_reports": result["counts"]["qualification_reports"],
        "raw_audit_rows": raw["rows"],
        "quality_gates": result["quality"]["gates"],
        "candidate_eligible_for_retention": policy["candidate_eligible_for_retention"],
        "seal_checked": seal_checked,
        "disposition_checked": disposition is not None,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--require-final-seal", action="store_true")
    args = parser.parse_args(argv)
    try:
        print(json.dumps(validate(require_final_seal=args.require_final_seal),
                         indent=2, sort_keys=True))
    except analyze.ReplayError as error:
        print(f"0787 validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
