#!/usr/bin/env python3
"""Qualify the strict 0467 primary XLSX ABBA follow-up.

The normal 0467 analyzer intentionally remains descriptive because its six
rows have 100 samples.  This module is the separate strict follow-up for the
single primary cell: ``xlsx_one_percent_commit_save`` on the deterministic
``dense-wide`` corpus.  It accepts exactly four clean 500-sample legs in
``A1-500-clean, B1-500-clean, B2-500-clean, A2-500-clean`` order and delegates
raw report validation, statistic recomputation, identity checks, and ABBA
decision logic to :mod:`tools.perf_abba_summary`.

The helper is read-only.  It emits a candidate claim-registry entry for human
review; it never edits the registry or creates a publication package.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
from typing import Any, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
REPO_ROOT = next(
    (candidate for candidate in ROOT.parents if (candidate / "tools").is_dir()),
    ROOT,
)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import analyze as normal  # noqa: E402
from tools import perf_abba_summary  # noqa: E402


SCHEMA = "litchi-0467-xlsx-primary-abba-qualification-v1"
STRICT_LANES = ("A1-500-clean", "B1-500-clean", "B2-500-clean", "A2-500-clean")
STRICT_LEGS = ("a1", "b1", "b2", "a2")
STRICT_ROLES = {
    "A1-500-clean": "control",
    "A2-500-clean": "control",
    "B1-500-clean": "candidate",
    "B2-500-clean": "candidate",
}
PRIMARY_CASE = "xlsx_one_percent_commit_save"
PRIMARY_SHAPE = "dense-wide"
SAMPLES = 500
WARMUPS = 5
CPU = "2"
WORKERS = 1
DRIFT_CEILINGS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}
STATISTICS = ("p50", "mean", "p95", "p99")
PACKAGE_CHANGE_ID = "0467-xlsx-cell-attributes-abba"
CHANGE_ID = PACKAGE_CHANGE_ID
CLAIM_ID = "claim-0467-xlsx-cell-attributes"
EVIDENCE_ID = "abba-0467-xlsx-cell-attributes"
DOCUMENTATION = "docs/performance/changes/0467-xlsx-cell-attributes.md"
PACKAGE_MANIFEST_NAME = f"{PACKAGE_CHANGE_ID}-manifest.json"
PACKAGE_SUMMARY_NAME = "summary.json"
PROTOCOL_CANDIDATES = ("protocol-r2.json", "protocol-500.json", "protocol-strict500.json")


class QualificationError(normal.AnalysisError):
    """Raised when strict primary evidence does not meet its frozen contract."""


def _object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise QualificationError(f"{location} must be an object")
    return value


def _exact_list(value: Any, expected: Sequence[Any], location: str) -> None:
    if value != list(expected):
        raise QualificationError(f"{location} does not match {list(expected)!r}: {value!r}")


def _protocol_path(root: Path, requested: Path | None) -> Path:
    if requested is not None:
        path = requested.resolve()
        if not path.is_file():
            raise QualificationError(f"missing strict protocol {path}")
        return path
    for name in PROTOCOL_CANDIDATES:
        path = root / name
        if path.is_file():
            return path
    raise QualificationError(
        "missing strict protocol; pass --protocol (expected a frozen 500-sample protocol)"
    )


def _validate_protocol(path: Path) -> dict[str, Any]:
    protocol = _object(normal.load_json(path), str(path))
    if protocol.get("schema") != "litchi-0467-qualification-protocol-v1":
        raise QualificationError(f"{path}.schema is not the frozen 0467 qualification protocol")
    if protocol.get("status") != "frozen":
        raise QualificationError(f"{path}.status must be frozen")
    _exact_list(protocol.get("roles"), ("control", "candidate"), f"{path}.roles")
    _exact_list(protocol.get("order"), STRICT_LANES, f"{path}.order")
    if protocol.get("samples") != SAMPLES:
        raise QualificationError(f"{path}.samples must be {SAMPLES}")
    if protocol.get("warmups") != WARMUPS:
        raise QualificationError(f"{path}.warmups must be {WARMUPS}")
    if protocol.get("cpu") != 2 or protocol.get("workers") != WORKERS:
        raise QualificationError(f"{path} CPU/worker settings must be CPU 2 and one worker")
    if "primary_case" in protocol and protocol["primary_case"] != PRIMARY_CASE:
        raise QualificationError(f"{path}.primary_case is not the 0467 primary case")
    if "primary_corpus" in protocol and protocol["primary_corpus"] != "xlsx-dense-wide":
        raise QualificationError(f"{path}.primary_corpus is not xlsx-dense-wide")
    _exact_list(protocol.get("cases"), (PRIMARY_CASE,), f"{path}.cases")
    _exact_list(protocol.get("shapes"), (PRIMARY_SHAPE,), f"{path}.shapes")
    triggers = protocol.get("review_triggers")
    if triggers is not None:
        triggers = _object(triggers, f"{path}.review_triggers")
        if triggers.get("latency_percent") != 5 or triggers.get("process_rss_percent") != 5:
            raise QualificationError(f"{path}.review_triggers must retain the 5% thresholds")
    return protocol


def _report_configuration(report: Mapping[str, Any], label: str) -> dict[str, Any]:
    configuration = _object(report.get("configuration"), f"{label}.configuration")
    if configuration.get("samples_per_case") != SAMPLES:
        raise QualificationError(f"{label}.configuration.samples_per_case must be {SAMPLES}")
    if configuration.get("warmup_iterations_per_case") != WARMUPS:
        raise QualificationError(f"{label}.configuration.warmup_iterations_per_case must be {WARMUPS}")
    _exact_list(configuration.get("cases"), (PRIMARY_CASE,), f"{label}.configuration.cases")
    _exact_list(configuration.get("xlsx_shapes"), (PRIMARY_SHAPE,), f"{label}.configuration.xlsx_shapes")
    if configuration.get("execution_workers") != [WORKERS]:
        raise QualificationError(f"{label}.configuration.execution_workers must be [1]")
    return configuration


def _catalog_identity(root: Path, lane_dir: Path, report: Mapping[str, Any], label: str) -> dict[str, Any]:
    try:
        return normal._catalog_file_identity(root, lane_dir, report, label)
    except normal.AnalysisError as error:
        raise QualificationError(str(error)) from error


def _validate_report(
    report: Mapping[str, Any],
    label: str,
    binding: Mapping[str, Any],
    expected_configuration: Mapping[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    normal._report_profile(report, label)
    environment = _object(report.get("environment"), f"{label}.environment")
    if environment.get("git_worktree_dirty") is not False:
        raise QualificationError(f"{label}.environment.git_worktree_dirty must be false")
    if environment.get("cpu_affinity") != CPU:
        raise QualificationError(f"{label}.environment.cpu_affinity must be CPU 2")
    revision = normal._revision(environment.get("git_revision"), f"{label}.environment.git_revision")
    if revision != binding["revision"]:
        raise QualificationError(f"{label}.environment.git_revision does not match its binding")
    configuration = _report_configuration(report, label)
    if configuration != dict(expected_configuration):
        raise QualificationError(f"{label}.configuration differs from the clean control configuration")
    binary = _object(report.get("binary_identity"), f"{label}.binary_identity")
    if normal._sha(binary.get("binary_sha256"), f"{label}.binary_identity.binary_sha256") != binding["binary_sha256"]:
        raise QualificationError(f"{label}.binary_identity.binary_sha256 does not match its binding")
    if binary.get("binary_bytes") != binding["bytes"]:
        raise QualificationError(f"{label}.binary_identity.binary_bytes does not match its binding")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1:
        raise QualificationError(f"{label}.results must contain exactly the primary row")
    row = _object(results[0], f"{label}.results[0]")
    if row.get("case") != PRIMARY_CASE:
        raise QualificationError(f"{label}.results[0].case is not the primary case")
    corpus = _object(row.get("corpus"), f"{label}.results[0].corpus")
    if corpus.get("shape") != PRIMARY_SHAPE or corpus.get("name") != "xlsx-dense-wide":
        raise QualificationError(f"{label}.results[0].corpus is not xlsx-dense-wide")
    elapsed = _object(row.get("elapsed_ns"), f"{label}.results[0].elapsed_ns")
    samples = elapsed.get("samples")
    if not isinstance(samples, list) or len(samples) != SAMPLES:
        raise QualificationError(f"{label}.results[0].elapsed_ns must contain {SAMPLES} samples")
    try:
        statistics = perf_abba_summary.recompute_statistics(
            elapsed, f"{label}.results[0].elapsed_ns"
        )
    except Exception as error:
        raise QualificationError(f"{label} statistic validation failed: {error}") from error
    return row, statistics


def _primary_reference(root: Path) -> str | None:
    """Return the normal A1 primary corpus identity when it is retained."""

    report_path = root / "A1-clean" / "report.json"
    if not report_path.is_file():
        return None
    report = _object(normal.load_json(report_path), str(report_path))
    for raw in report.get("results", []):
        row = _object(raw, f"{report_path}.results")
        if row.get("case") == PRIMARY_CASE and _object(row.get("corpus"), "A1 corpus").get("shape") == PRIMARY_SHAPE:
            return normal.canonical_sha256(row["corpus"])
    raise QualificationError("A1-clean is present but has no retained primary corpus row")


def _cell_list(summary_result: Mapping[str, Any], field: str) -> list[dict[str, str]]:
    elapsed = _object(summary_result.get("elapsed_ns"), "summary result.elapsed_ns")
    values = elapsed.get(field, [])
    if not isinstance(values, list):
        raise QualificationError(f"summary result.elapsed_ns.{field} must be a list")
    corpus = _object(summary_result.get("corpus"), "summary result.corpus")
    corpus_name = normal._string(corpus.get("name"), "summary result.corpus.name")
    for statistic in values:
        if statistic not in STATISTICS:
            raise QualificationError(f"summary result.elapsed_ns.{field} contains {statistic!r}")
    return [
        {
            "case": PRIMARY_CASE,
            "corpus": corpus_name,
            "statistic": statistic,
        }
        for statistic in values
    ]


def _proposal(
    summary_result: Mapping[str, Any],
    *,
    strict_summary_path: str,
) -> dict[str, Any]:
    accepted = _cell_list(summary_result, "accepted_statistics")
    adverse = _cell_list(summary_result, "adverse_both_statistics")
    status = "landed" if accepted else "rejected"
    code_state = "landed" if accepted else "not_landed"
    reasons = ["accepted_latency_only"] if accepted else ["no_accepted_latency_cells"]
    if adverse:
        reasons.append("explicit_adverse_both_cells")
    corpus = _object(summary_result.get("corpus"), "summary result.corpus")
    scope_corpus = {
        key: corpus[key]
        for key in ("name", "archive_sha256", "generator", "shape", "package_format")
        if key in corpus
    }
    latency: dict[str, Any] = {
        "evidence_id": EVIDENCE_ID,
        "allowed_statistics": list(STATISTICS),
        "accepted_cells": len(accepted),
        "adverse_both_cells": len(adverse),
        "accepted_statistics": accepted,
    }
    if adverse:
        latency["adverse_both_statistics"] = adverse
    entry: dict[str, Any] = {
        "id": CLAIM_ID,
        "change_id": CHANGE_ID,
        "claim_class": "optimization",
        "status": status,
        "code_state": code_state,
        "reason_codes": reasons,
        "scope": {
            "format": "xlsx",
            "selectors": [PRIMARY_CASE],
            "corpora": [scope_corpus],
        },
        "latency_evidence": latency,
        "documentation": [DOCUMENTATION],
    }
    return {
        "proposal_only": True,
        "strict_summary_path": strict_summary_path,
        "package_manifest_name": PACKAGE_MANIFEST_NAME,
        "entry": entry,
    }


def _package_record(
    root: Path,
    summary: Mapping[str, Any],
    *,
    package_dir: Path | None,
) -> dict[str, Any]:
    """Describe and, when retained, verify the standard strict ABBA package.

    Packaging itself remains an explicit operation through
    ``tools/perf_abba_package.py``.  This read-only check accepts either the
    result bundle root or a ``strict`` subdirectory and binds the manifest to
    the canonical summary and its four report identities.
    """

    candidate_dirs: list[Path]
    if package_dir is not None:
        candidate_dirs = [package_dir.resolve()]
    else:
        candidate_dirs = [root / "strict", root]
    manifest_path: Path | None = None
    for candidate in candidate_dirs:
        item = candidate / PACKAGE_MANIFEST_NAME
        if item.is_file():
            manifest_path = item
            break
    package = {
        "required_before_registry_registration": True,
        "change_id": PACKAGE_CHANGE_ID,
        "manifest_name": PACKAGE_MANIFEST_NAME,
        "summary_name": PACKAGE_SUMMARY_NAME,
        "present": manifest_path is not None,
        "verified": False,
    }
    if manifest_path is None:
        if package_dir is not None and not package_dir.is_dir():
            raise QualificationError(f"missing requested package directory {package_dir}")
        return package
    manifest = _object(normal.load_json(manifest_path), str(manifest_path))
    if manifest.get("schema_version") != 1 or manifest.get("manifest_kind") != "litchi-perf-abba-artifacts":
        raise QualificationError(f"{manifest_path} is not a standard ABBA package manifest")
    if manifest.get("change_id") != PACKAGE_CHANGE_ID or manifest.get("change") != PACKAGE_CHANGE_ID:
        raise QualificationError(f"{manifest_path} change identity does not match {PACKAGE_CHANGE_ID}")
    if manifest.get("manifest_path") != PACKAGE_MANIFEST_NAME or manifest.get("self_excluded") is not True:
        raise QualificationError(f"{manifest_path} path/self-exclusion metadata is invalid")
    compression = _object(manifest.get("compression"), f"{manifest_path}.compression")
    if compression.get("format") != "zstd" or compression.get("threads") != 1:
        raise QualificationError(f"{manifest_path}.compression must be single-threaded zstd")
    artifacts = manifest.get("artifacts")
    if not isinstance(artifacts, list) or [item.get("role") for item in artifacts if isinstance(item, dict)] != list(STRICT_LEGS):
        raise QualificationError(f"{manifest_path}.artifacts must contain a1,b1,b2,a2")
    for index, item in enumerate(artifacts):
        item = _object(item, f"{manifest_path}.artifacts[{index}]")
        artifact_name = item.get("path")
        if not isinstance(artifact_name, str) or not artifact_name or Path(artifact_name).name != artifact_name:
            raise QualificationError(f"{manifest_path}.artifacts[{index}].path must be a flat file name")
        if not (manifest_path.parent / artifact_name).is_file():
            raise QualificationError(f"package is missing {manifest_path.parent / artifact_name}")
    summary_meta = _object(manifest.get("summary"), f"{manifest_path}.summary")
    if summary_meta.get("path") != PACKAGE_SUMMARY_NAME:
        raise QualificationError(f"{manifest_path}.summary.path must be {PACKAGE_SUMMARY_NAME}")
    summary_path = manifest_path.parent / PACKAGE_SUMMARY_NAME
    if not summary_path.is_file():
        raise QualificationError(f"package is missing {summary_path}")
    package_summary = _object(normal.load_json(summary_path), str(summary_path))
    if package_summary != dict(summary):
        raise QualificationError("retained package summary differs from strict recomputation")
    if summary_meta.get("sha256") != normal.raw_sha256(summary_path):
        raise QualificationError(f"{manifest_path}.summary.sha256 does not match summary.json")
    if summary_meta.get("canonical_sha256") != normal.canonical_sha256(package_summary):
        raise QualificationError(f"{manifest_path}.summary.canonical_sha256 does not match summary.json")
    if summary_meta.get("report_identity") != package_summary.get("report_identity"):
        raise QualificationError(f"{manifest_path}.summary.report_identity does not match summary.json")
    if manifest.get("summary_identity") != summary_meta:
        raise QualificationError(f"{manifest_path}.summary_identity differs from summary metadata")
    package.update(
        {
            "manifest_path": str(manifest_path.relative_to(root))
            if manifest_path.is_relative_to(root)
            else str(manifest_path),
            "manifest_sha256": normal.raw_sha256(manifest_path),
            "summary_path": str(summary_path.relative_to(root))
            if summary_path.is_relative_to(root)
            else str(summary_path),
            "summary_sha256": normal.raw_sha256(summary_path),
            "summary_canonical_sha256": normal.canonical_sha256(package_summary),
            "verified": True,
        }
    )
    return package


def qualify(
    root: Path = ROOT,
    *,
    protocol_path: Path | None = None,
    lane_paths: Mapping[str, Path] | None = None,
    package_dir: Path | None = None,
) -> dict[str, Any]:
    root = root.resolve()
    protocol_file = _protocol_path(root, protocol_path)
    protocol = _validate_protocol(protocol_file)
    lanes = dict(lane_paths or {name: root / name for name in STRICT_LANES})
    if set(lanes) != set(STRICT_LANES):
        raise QualificationError(f"lane set must be {list(STRICT_LANES)!r}")
    if any(not path.is_dir() for path in lanes.values()):
        missing = [str(path) for path in lanes.values() if not path.is_dir()]
        raise QualificationError(f"strict primary lanes are incomplete: {missing!r}")

    bindings = {role: normal._binding(root, role) for role in ("control", "candidate")}
    binding_values = {role: value for role, (_, value) in bindings.items()}
    if binding_values["control"]["revision"] == binding_values["candidate"]["revision"]:
        raise QualificationError("control and candidate revisions must be distinct")
    if binding_values["control"]["binary_sha256"] == binding_values["candidate"]["binary_sha256"]:
        raise QualificationError("control and candidate binaries must be distinct")

    reports: dict[str, dict[str, Any]] = {}
    rows: dict[str, dict[str, Any]] = {}
    statistics: dict[str, dict[str, Any]] = {}
    receipts: dict[str, dict[str, Any]] = {}
    catalogs: dict[str, dict[str, Any]] = {}
    resources: dict[str, dict[str, Any]] = {}
    expected_configuration: dict[str, Any] | None = None
    for lane, leg in zip(STRICT_LANES, STRICT_LEGS):
        lane_dir = lanes[lane]
        role = STRICT_ROLES[lane]
        binding_path, binding = bindings[role]
        receipt_path = lane_dir / "receipt.json"
        receipt = normal._validate_receipt(
            receipt_path,
            lane,
            binding_path,
            binding,
            role=role,
            samples=SAMPLES,
            warmups=WARMUPS,
            heaptrack=False,
        )
        report_path = lane_dir / "report.json"
        report = _object(normal.load_json(report_path), str(report_path))
        configuration = _report_configuration(report, lane)
        if expected_configuration is None:
            expected_configuration = configuration
        row, checked = _validate_report(
            report, lane, binding, expected_configuration
        )
        reports[leg] = report
        rows[leg] = row
        statistics[leg] = checked
        receipts[leg] = receipt
        catalogs[leg] = _catalog_identity(root, lane_dir, report, lane)
        resources[leg] = normal._resource_rss(lane_dir / "resource.log")

    assert expected_configuration is not None
    corpus_identity = normal.canonical_sha256(rows["a1"]["corpus"])
    if any(normal.canonical_sha256(row["corpus"]) != corpus_identity for row in rows.values()):
        raise QualificationError("strict ABBA lanes do not retain the same corpus identity")
    if len({catalog["corpus_identity_sha256"] for catalog in catalogs.values()}) != 1:
        raise QualificationError("strict ABBA lanes do not retain the same corpus catalog")
    normal_reference = _primary_reference(root)
    if normal_reference is not None and normal_reference != corpus_identity:
        raise QualificationError("strict primary corpus differs from retained A1-clean corpus")

    try:
        summary = perf_abba_summary.summarize_reports(
            [reports[leg] for leg in STRICT_LEGS],
            drift_ceilings=DRIFT_CEILINGS,
            cases=(PRIMARY_CASE,),
            shapes=(PRIMARY_SHAPE,),
        )
    except Exception as error:
        raise QualificationError(f"strict canonical ABBA summary validation failed: {error}") from error
    summary_results = summary.get("results")
    if not isinstance(summary_results, list) or len(summary_results) != 1:
        raise QualificationError("strict canonical summary must contain one primary result")
    summary_result = _object(summary_results[0], "strict summary.results[0]")
    if summary_result.get("case") != PRIMARY_CASE or summary_result.get("shape") != PRIMARY_SHAPE:
        raise QualificationError("strict canonical summary selected the wrong primary result")

    lane_records: dict[str, Any] = {}
    for lane, leg in zip(STRICT_LANES, STRICT_LEGS):
        lane_dir = lanes[lane]
        binding_path, binding = bindings[STRICT_ROLES[lane]]
        report_path = lane_dir / "report.json"
        receipt_path = lane_dir / "receipt.json"
        lane_records[lane] = {
            "leg": leg,
            "role": STRICT_ROLES[lane],
            "path": str(lane_dir.relative_to(root)),
            "report": {
                "path": str(report_path.relative_to(root)),
                "sha256": normal.raw_sha256(report_path),
                "canonical_sha256": normal.canonical_sha256(reports[leg]),
                "bytes": report_path.stat().st_size,
            },
            "receipt": {
                "path": str(receipt_path.relative_to(root)),
                "sha256": normal.raw_sha256(receipt_path),
                "bytes": receipt_path.stat().st_size,
            },
            "receipt_record": receipts[leg],
            "binding_path": str(binding_path.relative_to(root)),
            "binding": binding,
            "catalog": catalogs[leg],
            "statistics_ns": statistics[leg],
        }

    package = _package_record(root, summary, package_dir=package_dir)
    strict_summary_path = PACKAGE_SUMMARY_NAME
    accepted = _cell_list(summary_result, "accepted_statistics")
    adverse = _cell_list(summary_result, "adverse_both_statistics")
    rss_pairings = {
        "a1_control_to_b1_candidate": normal._ratio_percent(
            resources["a1"]["max_rss_kib"],
            resources["b1"]["max_rss_kib"],
            "strict_primary.process_rss.a1-b1",
        ),
        "a2_control_to_b2_candidate": normal._ratio_percent(
            resources["a2"]["max_rss_kib"],
            resources["b2"]["max_rss_kib"],
            "strict_primary.process_rss.a2-b2",
        ),
    }
    rss_drift = {
        "control": normal._drift(
            resources["a1"]["max_rss_kib"],
            resources["a2"]["max_rss_kib"],
            "strict_primary.process_rss.control-drift",
        ),
        "candidate": normal._drift(
            resources["b1"]["max_rss_kib"],
            resources["b2"]["max_rss_kib"],
            "strict_primary.process_rss.candidate-drift",
        ),
    }
    return {
        "schema": SCHEMA,
        "purpose": "strict scoped latency qualification for the 0467 primary XLSX cell",
        "protocol": {
            "path": str(protocol_file.relative_to(root))
            if protocol_file.is_relative_to(root)
            else str(protocol_file),
            "sha256": normal.raw_sha256(protocol_file),
            "canonical_sha256": normal.canonical_sha256(protocol),
            "schema": protocol.get("schema"),
            "order": list(STRICT_LANES),
            "samples": SAMPLES,
            "warmups": WARMUPS,
            "cpu": 2,
            "workers": WORKERS,
        },
        "scope": {
            "case": PRIMARY_CASE,
            "shape": PRIMARY_SHAPE,
            "corpus": rows["a1"]["corpus"],
            "configuration": expected_configuration,
        },
        "bindings": {
            role: {
                **binding,
                "path": str(path.relative_to(root)),
            }
            for role, (path, binding) in bindings.items()
        },
        "lanes": lane_records,
        "process_rss": {
            "scope": "whole process per strict lane; GNU time maximum resident set size",
            "kib": {leg: resources[leg]["max_rss_kib"] for leg in STRICT_LEGS},
            "pairings": rss_pairings,
            "same_implementation_drift": rss_drift,
            "review_threshold_percent": 5.0,
            "claim_role": "descriptive guard context; no resource guardrail registered",
        },
        "abba_summary": summary,
        "package": package,
        "claim_registration": {
            "registry_path": "docs/performance/claim-registry-v1.json",
            "latency_policy": "latency-abba-v1",
            "minimum_samples_per_case": 500,
            "drift_ceiling_percent": dict(DRIFT_CEILINGS),
            "accepted_cells": accepted,
            "adverse_both_cells": adverse,
            "strict_latency_claim_eligible": bool(accepted),
            "registry_entry_ready": bool(accepted) and package["verified"],
            "resource_guardrail_required": False,
            "retain_normal_six_descriptive_rows": True,
            "retain_full_default_and_heaptrack_guards_separately": True,
            "registry_entry_proposal": _proposal(
                summary_result, strict_summary_path=strict_summary_path
            ),
        },
        "verification": {
            "strict_protocol_verified": True,
            "clean_worktrees_verified": True,
            "report_clean_metadata_verified": True,
            "distinct_revisions_verified": True,
            "distinct_binaries_verified": True,
            "build_receipt_bindings_verified": True,
            "source_binding_verified_via_build_receipts": True,
            "configuration_identity_verified": True,
            "corpus_identity_verified": True,
            "catalog_identity_verified": True,
            "statistics_recomputed_from_500_samples": True,
            "canonical_abba_summary_verified": True,
            "drift_policy_verified": True,
            "process_rss_receipts_verified": True,
            "process_rss_review_trigger_is_not_claim": True,
            "package_manifest_verified": package["verified"],
            "instrumented_latency_not_used": True,
            "resource_guardrail_not_registered": True,
        },
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--protocol", type=Path)
    parser.add_argument("--json-out", type=Path)
    parser.add_argument(
        "--summary-out",
        type=Path,
        help="write the canonical perf_abba_summary object for packaging",
    )
    parser.add_argument(
        "--package-dir",
        type=Path,
        help="optional standard package directory to verify (root or strict/)",
    )
    for lane in STRICT_LANES:
        parser.add_argument(
            f"--{lane.lower().replace('-', '_')}",
            dest=lane.lower().replace('-', '_'),
            type=Path,
            help=f"{lane} formal lane directory",
        )
    return parser


def _write_json(path: Path, value: Mapping[str, Any]) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError as error:
        raise QualificationError(f"output already exists: {path}") from error
    except OSError as error:
        raise QualificationError(f"cannot write {path}: {error}") from error


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(list(argv) if argv is not None else None)
    try:
        root = args.root.resolve()
        supplied = {
            lane: getattr(args, lane.lower().replace('-', '_'))
            for lane in STRICT_LANES
            if getattr(args, lane.lower().replace('-', '_')) is not None
        }
        if supplied and len(supplied) != len(STRICT_LANES):
            raise QualificationError("supply all four strict lane paths or none")
        result = qualify(
            root,
            protocol_path=args.protocol.resolve() if args.protocol else None,
            lane_paths={lane: path.resolve() for lane, path in supplied.items()} or None,
            package_dir=args.package_dir.resolve() if args.package_dir else None,
        )
        if args.json_out is not None:
            _write_json(args.json_out, result)
        if args.summary_out is not None:
            _write_json(args.summary_out, result["abba_summary"])
        json.dump(result, sys.stdout, indent=2, sort_keys=True, allow_nan=False)
        sys.stdout.write("\n")
        return 0
    except (QualificationError, OSError, ValueError) as error:
        print(f"{SCHEMA}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
