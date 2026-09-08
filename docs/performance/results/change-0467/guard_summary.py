#!/usr/bin/env python3
"""Recompute the 0467 supplemental six-case ABBA guard.

This is descriptive review evidence.  It validates the retained clean capture
receipts and corpus identities, recomputes every mean/p50/p95/p99 cell, and
records both ABBA directions plus same-role drift.  It deliberately calls the
canonical ABBA summary without a global shape filter because the selected
matrix contains CFB, DOC, and XLSX rows.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import sys
import tempfile
from typing import Any, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0467-supplemental-guard-summary-v1"
PROTOCOL_SCHEMA = "litchi-0467-supplemental-guard-protocol-v1"
CAPTURE_SCHEMA = "litchi-0467-capture-v1"
LANES = ("A1-guard-clean", "B1-guard-clean", "B2-guard-clean", "A2-guard-clean")
ROLES = {
    "A1-guard-clean": "control",
    "A2-guard-clean": "control",
    "B1-guard-clean": "candidate",
    "B2-guard-clean": "candidate",
}
CASES = (
    "cfb_list_streams",
    "cfb_read_one",
    "cfb_shared_concurrent_reads",
    "doc_fresh_write_to",
    "xlsx_source_first_cell",
    "xlsx_source_narrow_column_range_scan",
)
# The three CFB cases run against four shapes with both payload kinds (8 rows
# each), the DOC case uses its three writer shapes, and the two XLSX selectors
# each use the one declared medium corpus.  Keep this cardinality explicit so
# a reduced report cannot masquerade as a complete guard.
EXPECTED_CASE_COUNTS = {
    "cfb_list_streams": 8,
    "cfb_read_one": 8,
    "cfb_shared_concurrent_reads": 8,
    "doc_fresh_write_to": 3,
    "xlsx_source_first_cell": 1,
    "xlsx_source_narrow_column_range_scan": 1,
}
EXPECTED_RESULT_COUNT = sum(EXPECTED_CASE_COUNTS.values())
SHAPES = ("medium",)
STATISTICS = ("mean", "p50", "p95", "p99")
SAMPLES = 100
WARMUPS = 5
CPU = "2"
WORKERS = 1
THRESHOLD = 5.0
EXPECTED_CAPTURE_ENVIRONMENT = {
    "RUSTUP_TOOLCHAIN": "1.98.1",
    "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "DEBUGINFOD_URLS": "",
    "LC_ALL": "C",
}


class GuardSummaryError(ValueError):
    """Raised when the supplemental guard cannot be authenticated."""


def fail(message: str) -> None:
    raise GuardSummaryError(message)


def _load_verify() -> Any:
    path = ROOT / "verify.py"
    if not path.is_file():
        fail("verify.py is missing")
    spec = importlib.util.spec_from_file_location("litchi_0467_guard_verify", path)
    if spec is None or spec.loader is None:
        fail("cannot load verify.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except (OSError, ImportError, SyntaxError, TypeError, ValueError) as error:
        fail(f"cannot load verify.py ({error})")
    module.ROOT = ROOT
    return module


def _protocol(verify: Any) -> dict[str, Any]:
    path = ROOT / "protocol-guard.json"
    protocol = verify.obj(verify.load_json(path, path.name), path.name)
    if protocol.get("schema") != PROTOCOL_SCHEMA or protocol.get("status") != "frozen":
        fail("guard protocol schema/status differs")
    if protocol.get("order") != list(LANES):
        fail("guard protocol lane order differs")
    if protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS:
        fail("guard protocol sample/warmup counts differ")
    if protocol.get("cpu") != 2 or protocol.get("workers") != WORKERS:
        fail("guard protocol CPU/worker configuration differs")
    if protocol.get("cases") != list(CASES) or protocol.get("shapes") != list(SHAPES):
        fail("guard protocol case/shape matrix differs")
    triggers = verify.obj(protocol.get("review_triggers"), "protocol-guard.review_triggers")
    if triggers.get("latency_percent") != THRESHOLD or triggers.get("process_rss_percent") != THRESHOLD:
        fail("guard protocol review thresholds differ")
    driver_hash = verify.digest(protocol.get("capture_driver_sha256"), "protocol-guard.capture_driver_sha256")
    if driver_hash != verify.sha256_file(ROOT / "capture-guard.py"):
        fail("guard protocol driver hash differs")
    selection = verify.obj(protocol.get("selection_threshold"), "protocol-guard.selection_threshold")
    if selection.get("latency_increase_percent") != THRESHOLD or selection.get("baseline_latency_ns") != 100000 or selection.get("statistics") != ["mean", "p50"]:
        fail("guard protocol selection threshold differs")
    selection_inputs = verify.obj(protocol.get("selection_inputs"), "protocol-guard.selection_inputs")
    if set(selection_inputs) != {"A-full-clean", "B-full-clean"}:
        fail("guard protocol selection inputs differ")
    for lane, expected_hash in selection_inputs.items():
        if verify.digest(expected_hash, f"protocol-guard.selection_inputs.{lane}") != verify.sha256_file(
            ROOT / lane / "report.json"
        ):
            fail(f"guard protocol selection input hash differs for {lane}")
    return protocol


def _argv_guard(argv: Any, label: str, verify: Any) -> None:
    if not isinstance(argv, list) or argv[:3] != ["taskset", "-c", CPU] or "heaptrack" in argv:
        fail(f"{label}.argv is not the clean CPU-pinned run")
    if "--workers" not in argv or argv[argv.index("--workers") + 1] != str(WORKERS):
        fail(f"{label}.argv worker count differs")
    if "--warmup" not in argv or argv[argv.index("--warmup") + 1] != str(WARMUPS):
        fail(f"{label}.argv warmup count differs")
    if "--samples" not in argv or argv[argv.index("--samples") + 1] != str(SAMPLES):
        fail(f"{label}.argv sample count differs")
    if "--case" not in argv or argv[argv.index("--case") + 1] != ",".join(CASES):
        fail(f"{label}.argv case matrix differs")
    if "--xlsx-shape" not in argv or argv[argv.index("--xlsx-shape") + 1] != "medium":
        fail(f"{label}.argv XLSX shape differs")


def _receipt(
    lane: str,
    binding: Mapping[str, Any],
    verify: Any,
) -> tuple[dict[str, Any], dict[str, Path]]:
    lane_dir = ROOT / lane
    path = lane_dir / "receipt.json"
    receipt = verify.obj(verify.load_json(path, str(path)), str(path))
    role = ROLES[lane]
    if receipt.get("schema") != CAPTURE_SCHEMA or receipt.get("lane") != lane or receipt.get("role") != role:
        fail(f"{lane}: receipt schema/lane/role differs")
    if receipt.get("revision") != binding["revision"] or receipt.get("binary_sha256") != binding["binary_sha256"]:
        fail(f"{lane}: receipt binding identity differs")
    binding_path = ROOT / f"{role}-binding.json"
    if receipt.get("binding_sha256") != verify.sha256_file(binding_path):
        fail(f"{lane}: receipt binding hash differs")
    if verify.digest(receipt.get("driver_sha256"), f"{lane}.driver_sha256") != verify.sha256_file(ROOT / "capture-guard.py"):
        fail(f"{lane}: receipt driver hash differs")
    if receipt.get("samples") != SAMPLES or receipt.get("warmups") != WARMUPS:
        fail(f"{lane}: receipt sample/warmup counts differ")
    if receipt.get("environment") != EXPECTED_CAPTURE_ENVIRONMENT:
        fail(f"{lane}: receipt environment differs")
    if receipt.get("exit_code") != 0 or receipt.get("clean_before") is not True or receipt.get("clean_after") is not True or receipt.get("binary_unchanged") is not True or receipt.get("report_metadata_matches_clean_role") is not True:
        fail(f"{lane}: receipt does not prove successful clean custody")
    _argv_guard(receipt.get("argv"), lane, verify)
    artifacts = verify._artifact_hashes(lane, receipt)
    started = verify.obj(verify.load_json(artifacts["started.json"], f"{lane}.started.json"), f"{lane}.started.json")
    for key in ("schema", "lane", "role", "revision", "binary_sha256", "binding_sha256", "driver_sha256", "samples", "warmups", "cwd", "argv", "environment"):
        if started.get(key) != receipt.get(key):
            fail(f"{lane}.started.json.{key} differs from receipt")
    return receipt, artifacts


def _report(
    lane: str,
    binding: Mapping[str, Any],
    artifacts: Mapping[str, Path],
    verify: Any,
) -> tuple[dict[str, Any], dict[tuple[str, str], Mapping[str, Any]], str]:
    report_path = artifacts["report.json"]
    report = verify.obj(verify.load_json(report_path, f"{lane}.report.json"), f"{lane}.report.json")
    verify._report_identity(report, lane, binding)
    verify._configuration(report, lane, SAMPLES, WARMUPS, CASES, SHAPES)
    _catalog, catalog_identity = verify._catalog(
        report, artifacts["corpus-catalog.json"], lane, binding
    )
    indexed = verify._rows(report, SAMPLES, lane)
    if len(indexed) != EXPECTED_RESULT_COUNT:
        fail(f"{lane}: guard report must contain {EXPECTED_RESULT_COUNT} rows")
    counts = {case: 0 for case in CASES}
    for case, _identity in indexed:
        if case not in counts:
            fail(f"{lane}: unexpected guard case {case!r}")
        counts[case] += 1
    if counts != EXPECTED_CASE_COUNTS:
        fail(f"{lane}: guard case set differs")
    for case, identity in indexed:
        # ``--xlsx-shape medium`` selects the XLSX rows.  CFB and DOC corpora
        # intentionally retain their default multi-shape matrices.
        corpus = indexed[(case, identity)].get("corpus")
        if not isinstance(corpus, Mapping):
            fail(f"{lane}.{case}: corpus identity is not an object")
        if case.startswith("xlsx_") and corpus.get("shape") != "medium":
            fail(f"{lane}.{case}: corpus shape differs")
    return report, indexed, catalog_identity


def _ratio(baseline: float, current: float, label: str) -> dict[str, float | bool]:
    if not math.isfinite(baseline) or not math.isfinite(current) or baseline <= 0 or current <= 0:
        fail(f"{label}: requires positive finite values")
    delta = (current / baseline - 1.0) * 100.0
    return {
        "baseline": baseline,
        "current": current,
        "ratio_current_over_baseline": current / baseline,
        "delta_percent": delta,
        "reduction_percent": -delta,
        "regression_over_5_percent": delta > THRESHOLD,
        "improvement_over_5_percent": -delta > THRESHOLD,
    }


def _drift(first: float, second: float, label: str) -> dict[str, float | bool]:
    value = _ratio(first, second, label)
    return {
        "first": first,
        "second": second,
        "delta_percent": value["delta_percent"],
        "absolute_delta_percent": abs(value["delta_percent"]),
        "review_triggered": abs(value["delta_percent"]) > THRESHOLD,
    }


def _canonical_json(value: Any) -> str:
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        )
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot canonicalize identity value ({error})")
    raise AssertionError("unreachable")


def _flatten(value: Any, prefix: str = "") -> dict[str, Any]:
    """Flatten identity objects while retaining vectors as atomic values."""

    if isinstance(value, Mapping):
        result: dict[str, Any] = {}
        for key in sorted(value):
            child = f"{prefix}.{key}" if prefix else str(key)
            result.update(_flatten(value[key], child))
        return result
    return {prefix or "<root>": value}


def _identity_review(
    key: tuple[str, str],
    indexed: Mapping[str, Mapping[tuple[str, str], Mapping[str, Any]]],
    perf_abba_summary: Any,
) -> dict[str, Any]:
    """Apply the canonical tool's identity projection without changing rows."""

    case, _identity = key
    result: dict[str, Any] = {}
    for field in ("source", "sink"):
        projections: dict[str, Any] = {}
        details: dict[str, Any] = {}
        for lane in LANES:
            row = indexed[lane][key]
            present = field in row and row[field] is not None
            detail: dict[str, Any] = {"present": present}
            if present:
                value = row[field]
                if field == "source":
                    project = getattr(perf_abba_summary, "_source_identity_projection", None)
                    if not callable(project):
                        fail("canonical ABBA tool has no source identity projection")
                    projected = project(value)
                else:
                    project = getattr(perf_abba_summary, "_sink_identity_projection", None)
                    if not callable(project):
                        fail("canonical ABBA tool has no sink identity projection")
                    projected = project(value, case)
                canonical = _canonical_json(projected)
                projections[lane] = projected
                detail["identity_sha256"] = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
            details[lane] = detail
        present_values = {lane: details[lane]["present"] for lane in LANES}
        if len(set(present_values.values())) != 1:
            status = "presence_differs"
            changed_paths: list[dict[str, Any]] = []
        elif not next(iter(present_values.values())):
            status = "consistently_absent"
            changed_paths = []
        else:
            canonical_values = {
                lane: _canonical_json(projections[lane]) for lane in LANES
            }
            if len(set(canonical_values.values())) == 1:
                status = "verified_equal"
                changed_paths = []
            else:
                status = "identity_differs"
                flattened = {
                    lane: _flatten(projections[lane]) for lane in LANES
                }
                paths = sorted({path for value in flattened.values() for path in value})
                changed_paths = [
                    {
                        "path": path,
                        "values": {
                            lane: flattened[lane].get(path, "<missing>")
                            for lane in LANES
                        },
                    }
                    for path in paths
                    if len(
                        {
                            _canonical_json(flattened[lane].get(path, "<missing>"))
                            for lane in LANES
                        }
                    )
                    > 1
                ]
        claim_eligible = status in {"verified_equal", "consistently_absent"}
        result[field] = {
            "status": status,
            "claim_eligible": claim_eligible,
            "lanes": details,
            "changed_paths": changed_paths,
        }
    result["claim_eligible"] = all(
        result[field]["claim_eligible"] for field in ("source", "sink")
    )
    return result


def _canonical_validation(
    reports: Sequence[Mapping[str, Any]],
    perf_abba_summary: Any,
) -> tuple[Any, dict[str, Any]]:
    """Retain full canonical rejection and salvage case-scoped evidence."""

    try:
        full = perf_abba_summary.summarize_reports(reports=reports)
    except Exception as error:
        scoped: dict[str, Any] = {}
        for case in CASES:
            try:
                scoped[case] = {
                    "status": "passed",
                    "summary": perf_abba_summary.summarize_reports(
                        reports=reports, cases=[case]
                    ),
                }
            except Exception as scoped_error:
                scoped[case] = {
                    "status": "rejected",
                    "error": str(scoped_error),
                }
        return None, {
            "status": "rejected",
            "scope": "all 29 retained rows",
            "error": str(error),
            "scoped_cases": scoped,
        }
    return full, {
        "status": "passed",
        "scope": "all 29 retained rows",
        "scoped_cases": {
            case: {"status": "passed", "derived_from": "full_matrix"}
            for case in CASES
        },
    }


def _write_json_atomic(path: Path, value: Mapping[str, Any]) -> None:
    """Replace an explicitly requested output only after serialization succeeds."""

    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    except Exception:
        try:
            temporary.unlink()
        except FileNotFoundError:
            pass
        raise


def summarize(root: Path = ROOT) -> dict[str, Any]:
    global ROOT
    ROOT = root.resolve()
    verify = _load_verify()
    protocol = _protocol(verify)
    bindings = verify._bindings()
    reports: list[Mapping[str, Any]] = []
    indexed: dict[str, dict[tuple[str, str], Mapping[str, Any]]] = {}
    catalog_identities: dict[str, str] = {}
    resources: dict[str, int] = {}
    lane_records: dict[str, Any] = {}
    for lane in LANES:
        role = ROLES[lane]
        receipt, artifacts = _receipt(lane, bindings[role], verify)
        report, rows, catalog_identity = _report(lane, bindings[role], artifacts, verify)
        _argv_guard(receipt["argv"], lane, verify)
        resources[lane] = verify._resource(artifacts["resource.log"], f"{lane}.resource.log")
        reports.append(report)
        indexed[lane] = rows
        catalog_identities[lane] = catalog_identity
        lane_records[lane] = {
            "role": role,
            "revision": bindings[role]["revision"],
            "binary_sha256": bindings[role]["binary_sha256"],
            "report_sha256": verify.sha256_file(artifacts["report.json"]),
            "catalog_sha256": verify.sha256_file(artifacts["corpus-catalog.json"]),
            "receipt_sha256": verify.sha256_file(ROOT / lane / "receipt.json"),
            "resource": {"path": "resource.log", "max_rss_kib": resources[lane]},
            "samples": SAMPLES,
            "warmups": WARMUPS,
        }
    if len(set(catalog_identities.values())) != 1:
        fail("guard corpus catalogs differ")
    keys = set(indexed[LANES[0]])
    if any(set(indexed[lane]) != keys for lane in LANES[1:]):
        fail("guard lanes do not share exact case/corpus identities")
    if len(keys) != EXPECTED_RESULT_COUNT:
        fail("guard result key cardinality differs")
    try:
        from tools import perf_abba_summary
    except (ImportError, ModuleNotFoundError) as error:
        fail(f"canonical supplemental ABBA tool is unavailable ({error})")
    canonical, canonical_validation = _canonical_validation(reports, perf_abba_summary)
    identity_reviews = {
        key: _identity_review(key, indexed, perf_abba_summary) for key in keys
    }

    rows: list[dict[str, Any]] = []
    for key in sorted(
        keys,
        key=lambda value: (
            value[0],
            str(indexed[LANES[0]][value].get("corpus", {}).get("name", "")),
        ),
    ):
        case, identity = key
        corpus = indexed[LANES[0]][key]["corpus"]
        identity_review = identity_reviews[key]
        case_scope = canonical_validation["scoped_cases"][case]
        row_claim_eligible = bool(identity_review["claim_eligible"])
        values: dict[str, dict[str, float]] = {}
        for lane in LANES:
            elapsed = indexed[lane][key]["elapsed_ns"]
            checked = perf_abba_summary.recompute_statistics(elapsed, f"{lane}.{case}.elapsed_ns")
            values[lane] = {name: checked[name] for name in STATISTICS}
        pairings = {
            "a1_control_to_b1_candidate": {
                name: _ratio(values["A1-guard-clean"][name], values["B1-guard-clean"][name], f"{case}.{name}.a1-b1")
                for name in STATISTICS
            },
            "a2_control_to_b2_candidate": {
                name: _ratio(values["A2-guard-clean"][name], values["B2-guard-clean"][name], f"{case}.{name}.a2-b2")
                for name in STATISTICS
            },
        }
        drift = {
            "control_a1_to_a2": {
                name: _drift(values["A1-guard-clean"][name], values["A2-guard-clean"][name], f"{case}.{name}.a1-a2")
                for name in STATISTICS
            },
            "candidate_b1_to_b2": {
                name: _drift(values["B1-guard-clean"][name], values["B2-guard-clean"][name], f"{case}.{name}.b1-b2")
                for name in STATISTICS
            },
        }
        rss_pairings = {
            "a1_control_to_b1_candidate": _ratio(resources["A1-guard-clean"], resources["B1-guard-clean"], f"{case}.rss.a1-b1"),
            "a2_control_to_b2_candidate": _ratio(resources["A2-guard-clean"], resources["B2-guard-clean"], f"{case}.rss.a2-b2"),
        }
        rss_drift = {
            "control_a1_to_a2": _drift(resources["A1-guard-clean"], resources["A2-guard-clean"], f"{case}.rss.a1-a2"),
            "candidate_b1_to_b2": _drift(resources["B1-guard-clean"], resources["B2-guard-clean"], f"{case}.rss.b1-b2"),
        }
        rows.append(
            {
                "case": case,
                "shape": corpus.get("shape"),
                "corpus": corpus,
                "identity_review": identity_review,
                "claim_eligible": row_claim_eligible,
                "claim_eligibility_scope": "source/sink identity only; strict latency claims remain disabled globally",
                "canonical_case_scope": {
                    "status": case_scope["status"],
                    "error": case_scope.get("error"),
                },
                "descriptive_timing_review": True,
                "statistics_ns": values,
                "latency_pairings": pairings,
                "latency_same_role_drift": drift,
                "process_rss_kib": {lane: resources[lane] for lane in LANES},
                "rss_scope": "whole process per lane; repeated for row context, not row-local RSS",
                "rss_pairings": rss_pairings,
                "rss_same_role_drift": rss_drift,
                "review_triggers": {
                    "threshold_percent": THRESHOLD,
                    "latency": {
                        "pair_regression_triggered": any(
                            value["regression_over_5_percent"]
                            for pair in pairings.values()
                            for value in pair.values()
                        ),
                        "same_role_drift_triggered": any(
                            value["review_triggered"]
                            for pair in drift.values()
                            for value in pair.values()
                        ),
                    },
                    "process_rss": {
                        "pair_regression_triggered": any(
                            value["regression_over_5_percent"] for value in rss_pairings.values()
                        ),
                        "same_role_drift_triggered": any(
                            value["review_triggered"] for value in rss_drift.values()
                        ),
                    },
                },
            }
        )
    return {
        "schema": SCHEMA,
        "purpose": "descriptive six-case supplemental ABBA guard; no strict 500-sample claim",
        "protocol": {
            "path": "protocol-guard.json",
            "sha256": verify.sha256_file(ROOT / "protocol-guard.json"),
            "schema": protocol["schema"],
            "order": list(LANES),
            "samples": SAMPLES,
            "warmups": WARMUPS,
            "cpu": 2,
            "workers": WORKERS,
            "cases": list(CASES),
            "shapes": list(SHAPES),
            "review_threshold_percent": THRESHOLD,
            "selection_threshold": protocol["selection_threshold"],
            "selection_inputs": protocol["selection_inputs"],
        },
        "bindings": bindings,
        "lanes": lane_records,
        "results": rows,
        "abba_summary": canonical,
        "canonical_validation": canonical_validation,
        "correction": (
            {
                "canonical_full_matrix": {
                    "status": "rejected",
                    "scope": canonical_validation["scope"],
                    "error": canonical_validation["error"],
                },
                "affected_rows": [
                    {
                        "case": key[0],
                        "corpus": indexed[LANES[0]][key]["corpus"],
                        "identity_review": identity_reviews[key],
                    }
                    for key, review in identity_reviews.items()
                    if not review["claim_eligible"]
                ],
                "method": "retain all raw rows and report descriptive timings; do not normalize or drop source counters",
            }
            if canonical_validation["status"] == "rejected"
            else None
        ),
        "claim_registration": {
            "strict_latency_claim_eligible": False,
            "row_identity_claim_eligible_count": sum(
                1 for review in identity_reviews.values() if review["claim_eligible"]
            ),
            "reason": (
                "supplemental guard has 100 samples per selected case and is review evidence only; canonical source-identity rejection remains recorded"
                if canonical_validation["status"] == "rejected"
                else "supplemental guard has 100 samples per selected case and is review evidence only"
            ),
        },
        "verification": {
            "clean_worktrees_verified": True,
            "distinct_revisions_verified": bindings["control"]["revision"] != bindings["candidate"]["revision"],
            "distinct_binaries_verified": bindings["control"]["binary_sha256"] != bindings["candidate"]["binary_sha256"],
            "receipt_artifacts_verified": True,
            "configuration_identity_verified": True,
            "corpus_identity_verified": True,
            "statistics_recomputed_from_100_samples": True,
            "canonical_abba_summary_without_shape_filter": canonical_validation["status"] == "passed",
            "canonical_abba_summary_rejection_recorded": canonical_validation["status"] == "rejected",
            "individual_mean_p50_p95_p99_reviewed": True,
            "same_role_drift_reviewed": True,
            "review_trigger_is_not_claim": True,
        },
    }


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--json-out", type=Path)
    args = parser.parse_args(list(argv) if argv is not None else None)
    try:
        result = summarize(args.root)
        if args.json_out is not None:
            _write_json_atomic(args.json_out, result)
            print(f"{SCHEMA}: wrote {args.json_out}")
        else:
            print(json.dumps(result, indent=2, sort_keys=True, allow_nan=False))
        return 0
    except (GuardSummaryError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"{SCHEMA}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
