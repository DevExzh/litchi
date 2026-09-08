#!/usr/bin/env python3
"""Review the frozen 0470 targeted guard without changing its raw reports.

The canonical ABBA call receives all four raw reports and all retained rows.
There is no comparison-copy projection and no deletion of measured counters.
If the canonical checker rejects a row identity, this review records the
rejection and emits per-row raw-sample descriptive statistics so the failure
remains inspectable without becoming a latency claim.
"""

from __future__ import annotations

import copy
import hashlib
import json
import math
from pathlib import Path
import sys
from typing import Any, Mapping, Sequence

import analyze
import verify


ROOT = Path(__file__).resolve().parent
# Historical command destinations stay fixed when the evidence is relocated.
CAPTURE_ROOT = Path("/home/zhuhe/code/litchi/docs/performance/results/change-0470")
SCHEMA = "litchi-0470-targeted-guard-review-v1"
LANES = ("guard-A1", "guard-B1", "guard-B2", "guard-A2")
ROLES = {
    "guard-A1": "control",
    "guard-A2": "control",
    "guard-B1": "candidate",
    "guard-B2": "candidate",
}
CASES = (
    "cfb_shared_read_one",
    "opc_noop_save",
    "cfb_read_one",
    "opc_source_open",
    "ppt_fresh_write_to",
)
SYNTHETIC_CASES = frozenset(CASES[:-1])
SHAPES = ("many-small", "few-large", "wide-root")
PAYLOADS = ("compressible", "incompressible")
WRITER_SHAPES = ("payload-heavy",)
SAMPLES = 100
WARMUPS = 5
EXPECTED_ROWS = 25
THRESHOLD = 5.0


def require(condition: bool, message: str) -> None:
    if not condition:
        raise verify.VerificationError(message)


def _sha(path: Path) -> tuple[str, int]:
    return verify.sha256_file(path)


def _canonical_digest(value: Any) -> str:
    return hashlib.sha256(verify.canonical(value)).hexdigest()


def _protocol(root: Path) -> dict[str, Any]:
    protocol = verify.read_json(root / "guard-protocol.json", "guard protocol")
    require(protocol.get("schema") == "litchi-0470-targeted-guard-protocol-v1", "guard protocol schema differs")
    require(protocol.get("order") == list(LANES), "guard protocol order differs")
    require(protocol.get("roles") == ROLES, "guard protocol role mapping differs")
    require(protocol.get("cases") == list(CASES), "guard protocol case list differs")
    require(protocol.get("synthetic_cases") == list(CASES[:-1]), "guard protocol synthetic cases differ")
    require(protocol.get("shapes") == list(SHAPES), "guard protocol shapes differ")
    require(protocol.get("payload_kinds") == list(PAYLOADS), "guard protocol payload list differs")
    require(protocol.get("writer_shapes") == list(WRITER_SHAPES), "guard protocol writer shapes differ")
    require((protocol.get("samples"), protocol.get("warmups")) == (SAMPLES, WARMUPS), "guard protocol sample configuration differs")
    require((protocol.get("cpu"), protocol.get("workers")) == (2, 1), "guard protocol CPU/worker configuration differs")
    require(protocol.get("expected_rows") == EXPECTED_ROWS, "guard protocol row cardinality differs")
    require(protocol.get("performance_claim") == "none", "guard protocol permits a performance claim")
    require(protocol.get("strict_abba") == {"profile": "current-v1", "all_rows": True, "report_projection": False, "counter_erasure": False}, "guard protocol strictness differs")
    require(verify.sha256_file(root / "guard_capture.py")[0] == protocol.get("capture_driver_sha256"), "guard capture driver hash differs")
    require(protocol.get("original_full_guard", {}).get("review_summary") == "review-summary.json", "original full guard review reference differs")
    require(protocol.get("original_full_guard", {}).get("canonical_summary") == "summary.json", "original full guard canonical reference differs")
    return protocol


def _expected_row_keys() -> set[tuple[str, bytes]]:
    # The four synthetic cases use every selected corpus shape and both default
    # payload kinds.  Fresh writer corpora are independent of those selectors.
    expected: set[tuple[str, bytes]] = set()
    for case in SYNTHETIC_CASES:
        for shape in SHAPES:
            for payload in PAYLOADS:
                expected.add((case, verify.canonical({"shape": shape, "payload_kind": payload})))
    return expected


def _row_identity(report: Mapping[str, Any], label: str) -> dict[tuple[str, bytes], dict[str, Any]]:
    rows = verify._report_rows(report, label)
    require(len(rows) == EXPECTED_ROWS, f"{label}: expected {EXPECTED_ROWS} rows, found {len(rows)}")
    seen_synthetic: set[tuple[str, str, str]] = set()
    seen_writer = False
    for (case, _), row in rows.items():
        require(case in CASES, f"{label}.{case}: unexpected case")
        corpus = verify.obj(row.get("corpus"), f"{label}.{case}.corpus")
        shape = corpus.get("shape")
        payload = corpus.get("payload_kind")
        if case in SYNTHETIC_CASES:
            require(shape in SHAPES, f"{label}.{case}: unexpected synthetic shape")
            require(payload in PAYLOADS, f"{label}.{case}: payload selector is not both defaults")
            identity = (case, str(shape), str(payload))
            require(identity not in seen_synthetic, f"{label}.{case}: duplicate synthetic row")
            seen_synthetic.add(identity)
        else:
            require(shape == "payload-heavy", f"{label}.{case}: writer shape differs")
            require(payload == "not-applicable", f"{label}.{case}: writer payload differs")
            require(not seen_writer, f"{label}.{case}: duplicate writer row")
            seen_writer = True
    require(len(seen_synthetic) == 24, f"{label}: synthetic row set is incomplete")
    require(seen_writer, f"{label}: payload-heavy PPT writer row is missing")
    return rows


def _verify_original_full_guard(root: Path, protocol: Mapping[str, Any]) -> dict[str, Any]:
    specification = verify.obj(protocol.get("original_full_guard"), "guard protocol.original_full_guard")
    review_path = root / verify.relative(specification.get("review_summary"), "original full guard review summary")
    canonical_path = root / verify.relative(specification.get("canonical_summary"), "original full guard canonical summary")
    review_sha, review_bytes = _sha(review_path)
    canonical_sha, canonical_bytes = _sha(canonical_path)
    expected_review_sha = protocol.get("original_full_guard_review_sha256")
    expected_canonical_sha = protocol.get("original_full_guard_summary_sha256")
    require(expected_review_sha == review_sha, "original full guard review summary hash differs")
    require(expected_canonical_sha == canonical_sha, "original full guard canonical summary hash differs")
    review = verify.read_json(review_path, "original full guard review summary")
    canonical = verify.read_json(canonical_path, "original full guard canonical summary")
    full_review = verify.obj(review.get("full_guard"), "original full guard review.full_guard")
    raw_flags = full_review.get("flags_above_threshold")
    require(isinstance(raw_flags, list), "original full guard raw flags are missing")
    require(len(raw_flags) == specification.get("raw_flags_above_five_percent"), "original full guard raw flag count differs")
    require(full_review.get("rows") == specification.get("rows"), "original full guard row count differs")
    require(full_review.get("samples") == 15 and full_review.get("warmups") == 3, "original full guard sample configuration differs")
    canonical_summary = verify.obj(full_review.get("canonical_summary"), "original full guard canonical summary")
    require(canonical_summary.get("regressions") == specification.get("canonical_regressions"), "original full guard canonical regression count differs")
    canonical_full = verify.obj(canonical.get("full_guard"), "original canonical summary.full_guard")
    comparison = verify.obj(canonical_full.get("comparison"), "original canonical full guard comparison")
    require(verify.obj(comparison.get("summary"), "original canonical comparison.summary").get("regressions") == specification.get("canonical_regressions"), "original canonical regression count differs")
    return {
        "review_summary": {"path": review_path.name, "sha256": review_sha, "bytes": review_bytes},
        "canonical_summary": {"path": canonical_path.name, "sha256": canonical_sha, "bytes": canonical_bytes},
        "rows": full_review["rows"],
        "samples": full_review["samples"],
        "warmups": full_review["warmups"],
        "threshold_percent": specification["threshold_percent"],
        "canonical_status": full_review.get("canonical_status"),
        "canonical_summary_counts": canonical_summary,
        # Keep every original flag in the targeted evidence so later review
        # cannot silently reduce the original guard to a single count.
        "flags_above_threshold": copy.deepcopy(raw_flags),
    }


def _verify_artifacts(root: Path, lane: str, receipt: Mapping[str, Any]) -> None:
    expected = {"report.json", "corpus-catalog.json", "resource.log", "stdout.log", "stderr.log", "started.json"}
    artifacts = verify.obj(receipt.get("artifacts"), f"{lane}.artifacts")
    require(set(artifacts) == expected, f"{lane}: artifact set differs")
    for name, metadata in artifacts.items():
        path = verify.bundle_file(root, f"{lane}/{name}", f"{lane}.{name}")
        metadata = verify.obj(metadata, f"{lane}.{name}.artifact")
        require(verify.digest(metadata.get("sha256"), f"{lane}.{name}.sha256") == _sha(path)[0], f"{lane}.{name}: artifact hash differs")
        require(verify.integer(metadata.get("bytes"), f"{lane}.{name}.bytes") == path.stat().st_size, f"{lane}.{name}: artifact size differs")


def _verify_lane(root: Path, lane: str, protocol: Mapping[str, Any], bindings: Mapping[str, Mapping[str, Any]]) -> tuple[dict[str, Any], dict[str, Any], dict[tuple[str, bytes], dict[str, Any]]]:
    role = ROLES[lane]
    binding = bindings[role]
    started = verify.read_json(root / lane / "started.json", f"{lane}.started.json")
    receipt = verify.read_json(root / lane / "receipt.json", f"{lane}.receipt.json")
    for item, label in ((started, "started"), (receipt, "receipt")):
        require(item.get("schema") == "litchi-0470-targeted-guard-capture-v1", f"{lane}.{label}: schema differs")
        require(item.get("lane") == lane and item.get("role") == role, f"{lane}.{label}: lane/role differs")
        require(item.get("revision") == binding.get("revision"), f"{lane}.{label}: revision differs")
        require(item.get("binary_sha256") == binding.get("binary_sha256"), f"{lane}.{label}: binary hash differs")
        require(item.get("binary_bytes") == binding.get("bytes"), f"{lane}.{label}: binary size differs")
        require(item.get("binary_path") == binding.get("binary_path"), f"{lane}.{label}: binary path differs")
        require(item.get("binding_sha256") == _sha(root / f"{role}-binding.json")[0], f"{lane}.{label}: binding hash differs")
        require(item.get("source_manifest_sha256") == binding.get("source_manifest_sha256"), f"{lane}.{label}: source manifest hash differs")
        require(item.get("source_inventory_sha256") == binding.get("source_manifest_sha256"), f"{lane}.{label}: source inventory hash differs")
        require(item.get("source_inventory_count") == len(verify.read_json(verify.bundle_file(root, binding["source_manifest"], f"{role}.source_manifest"), f"{role}.source_manifest")), f"{lane}.{label}: source inventory count differs")
        require(item.get("fixture_inventory") == binding.get("included_fixtures"), f"{lane}.{label}: fixture inventory differs")
        require(item.get("capture_driver_sha256") == protocol.get("capture_driver_sha256"), f"{lane}.{label}: capture driver hash differs")
        require(item.get("protocol_sha256") == _sha(root / "guard-protocol.json")[0], f"{lane}.{label}: protocol hash differs")
        require(item.get("cwd") == binding.get("build_path"), f"{lane}.{label}: build path differs")
        require(item.get("environment") == protocol.get("environment"), f"{lane}.{label}: environment differs")
    for key, value in started.items():
        require(receipt.get(key) == value, f"{lane}: started receipt field {key!r} changed")
    require(receipt.get("exit_code") == 0, f"{lane}: capture command failed")
    require(receipt.get("clean_before") is True and receipt.get("clean_after") is True, f"{lane}: clean checkout attestation missing")
    require(receipt.get("source_inventory_verified_before") is True, f"{lane}: source inventory attestation missing")
    require(receipt.get("binary_unchanged") is True and receipt.get("report_metadata_matches_clean_role") is True, f"{lane}: identity attestation missing")
    require(receipt.get("samples") == SAMPLES and receipt.get("warmups") == WARMUPS, f"{lane}: sample configuration differs")
    argv = receipt.get("argv")
    require(isinstance(argv, list) and all(isinstance(value, str) for value in argv), f"{lane}: argv is malformed")
    expected_tokens = [
        "taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(CAPTURE_ROOT / lane / "resource.log"),
        str(binding["binary_path"]), "--workers", "1", "--warmup", "5", "--samples", "100",
        "--case", ",".join(CASES), "--shape", ",".join(SHAPES), "--writer-shape", ",".join(WRITER_SHAPES),
        "--json", str(CAPTURE_ROOT / lane / "report.json"), "--corpus-manifest", str(CAPTURE_ROOT / lane / "corpus-catalog.json"),
    ]
    require(argv == expected_tokens, f"{lane}: frozen command differs")
    require("--payload" not in argv, f"{lane}: explicit payload selector weakens default-payload coverage")
    _verify_artifacts(root, lane, receipt)
    report = verify.read_json(root / lane / "report.json", f"{lane}.report.json")
    verify.verify_report_metadata(report, f"{lane}.report", binding, SAMPLES, WARMUPS, cases=CASES)
    catalog = verify.read_json(root / lane / "corpus-catalog.json", f"{lane}.corpus-catalog.json")
    require(catalog.get("catalog_id") == "litchi-perf-corpus-v2" and catalog.get("manifest_version") == 2, f"{lane}: corpus catalog identity differs")
    require(isinstance(catalog.get("corpora"), list) and len(catalog["corpora"]) == 13, f"{lane}: corpus catalog corpus cardinality differs")
    require(isinstance(catalog.get("case_bindings"), list) and len(catalog["case_bindings"]) == EXPECTED_ROWS, f"{lane}: corpus catalog case cardinality differs")
    report_catalog = verify.obj(report.get("corpus_catalog"), f"{lane}.report.corpus_catalog")
    require(report_catalog.get("catalog_sha256") == catalog.get("catalog_sha256"), f"{lane}: report/catalog hash differs")
    configuration = verify.obj(report.get("configuration"), f"{lane}.configuration")
    require(configuration.get("corpus_shapes") == list(SHAPES), f"{lane}: corpus shape configuration differs")
    require(configuration.get("payload_kinds") == list(PAYLOADS), f"{lane}: payload configuration differs")
    require(configuration.get("writer_shapes") == list(WRITER_SHAPES), f"{lane}: writer configuration differs")
    rows = _row_identity(report, f"{lane}.report")
    return report, receipt, rows


def _timestamp(value: Any, label: str):
    return verify._timestamp(value, label)


def _chronology(root: Path, protocol: Mapping[str, Any], receipts: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
    intervals = []
    for lane in LANES:
        start = _timestamp(receipts[lane].get("started_utc"), f"{lane}.started_utc")
        finish = _timestamp(receipts[lane].get("finished_utc"), f"{lane}.finished_utc")
        require(finish >= start, f"{lane}: capture finished before it started")
        intervals.append((lane, start, finish))
    for previous, current in zip(intervals, intervals[1:]):
        require(current[1] >= previous[2], f"guard capture chronology overlaps: {previous[0]} and {current[0]}")
    prior_intervals = []
    for lane in tuple(protocol.get("prior_capture_lanes", ())) + tuple(protocol.get("prior_rss_lanes", ())):
        receipt_path = root / lane / "receipt.json"
        require(receipt_path.is_file(), f"prior serialized lane is missing: {lane}")
        prior = verify.read_json(receipt_path, f"{lane}.receipt.json")
        start = _timestamp(prior.get("started_utc"), f"{lane}.started_utc")
        finish = _timestamp(prior.get("finished_utc"), f"{lane}.finished_utc")
        require(finish >= start, f"{lane}: prior capture finished before it started")
        prior_intervals.append((lane, start, finish))
    ordered_prior = sorted(prior_intervals, key=lambda item: item[1])
    for previous, current in zip(ordered_prior, ordered_prior[1:]):
        require(current[1] >= previous[2], f"serialized prior captures overlap: {previous[0]} and {current[0]}")
    if prior_intervals:
        latest_prior = max(prior_intervals, key=lambda item: item[2])
        require(intervals[0][1] >= latest_prior[2], f"guard capture starts before serialized prior lane {latest_prior[0]} finishes")
    return {
        "guard_order": [item[0] for item in intervals],
        "guard_intervals_non_overlapping": True,
        "prior_lanes": [item[0] for item in prior_intervals],
        "prior_intervals_non_overlapping": True,
        "prior_intervals_non_overlapping_with_guard": True,
        "guard_starts_after_prior_captures": True,
    }


def _descriptive_rows(reports: Mapping[str, Mapping[str, Any]], row_maps: Mapping[str, Mapping[tuple[str, bytes], Mapping[str, Any]]], comparator: Any) -> list[dict[str, Any]]:
    keys = sorted(row_maps[LANES[0]], key=lambda item: (item[0], item[1]))
    output = []
    for key in keys:
        case, corpus_bytes = key
        first_row = row_maps[LANES[0]][key]
        corpus = first_row["corpus"]
        legs = {}
        for lane in LANES:
            row = row_maps[lane][key]
            require(row["case"] == case and verify.canonical(row["corpus"]) == corpus_bytes, f"{lane}.{case}: row identity differs")
            samples = row["elapsed_ns"]["samples"]
            percentiles = comparator._latencies(dict(row), f"{lane}.{case}", SAMPLES)
            legs[lane] = {
                "sample_count": len(samples),
                "min": min(samples),
                "max": max(samples),
                "mean": math.fsum(samples) / len(samples),
                **percentiles,
            }
        output.append({"case": case, "corpus": copy.deepcopy(corpus), "legs_ns": legs})
    return output


def _verify_oracles(row_maps: Mapping[str, Mapping[tuple[str, bytes], Mapping[str, Any]]], abba: Any) -> dict[str, Any]:
    source_variations = []
    keys = sorted(row_maps[LANES[0]], key=lambda item: (item[0], item[1]))
    for key in keys:
        case, corpus_bytes = key
        rows = [row_maps[lane][key] for lane in LANES]
        for field in ("sink", "output_sha256"):
            present = [field in row and row[field] is not None for row in rows]
            require(len(set(present)) == 1, f"{case}: {field} oracle presence differs")
            if present[0]:
                expected = verify.canonical(rows[0][field])
                require(all(verify.canonical(row[field]) == expected for row in rows[1:]), f"{case}: {field} oracle differs")
        source_present = ["source" in row and row["source"] is not None for row in rows]
        require(len(set(source_present)) == 1, f"{case}: source oracle presence differs")
        if source_present[0]:
            expected = _canonical_digest(rows[0]["source"])
            for lane, row in zip(LANES[1:], rows[1:]):
                observed = _canonical_digest(row["source"])
                if observed != expected:
                    source_variations.append({"case": case, "corpus_sha256": hashlib.sha256(corpus_bytes).hexdigest(), "lane": lane, "expected_source_sha256": expected, "observed_source_sha256": observed})
        operation_present = ["operation_metrics" in row and row["operation_metrics"] is not None for row in rows]
        require(len(set(operation_present)) == 1, f"{case}: operation metric oracle presence differs")
        if operation_present[0]:
            identities = []
            for lane, row in zip(LANES, rows):
                try:
                    identities.append(abba._operation_metrics_identity(row, f"{lane}.{case}", 1))
                except (AttributeError, KeyError, TypeError, ValueError) as error:
                    raise verify.VerificationError(f"{case}: operation metric oracle invalid ({error})") from error
            require(len(set(identities)) == 1, f"{case}: operation metric oracle differs")
    return {
        "correctness_oracles_verified": True,
        "sink_and_output_identity_verified": True,
        "operation_metric_shape_verified": True,
        "source_identity_variations_retained": source_variations,
        "source_identity_variation_count": len(source_variations),
        "source_vectors_are_not_projected": True,
    }


def evaluate(root: Path = ROOT) -> dict[str, Any]:
    root = root.resolve()
    protocol = _protocol(root)
    bindings = {role: verify.verify_role_binding(root, role) for role in ("control", "candidate")}
    require(bindings["control"].get("build_path") == protocol.get("build_path"), "control build path differs from guard protocol")
    require(bindings["candidate"].get("build_path") == protocol.get("build_path"), "candidate build path differs from guard protocol")
    require(bindings["control"].get("binary_path") == protocol.get("binary_paths", {}).get("control"), "control binary path differs from guard protocol")
    require(bindings["candidate"].get("binary_path") == protocol.get("binary_paths", {}).get("candidate"), "candidate binary path differs from guard protocol")
    verify.verify_source_manifest_pair(root, bindings)
    if bindings["control"].get("reused_binding_path") is not None:
        verify.verify_reused_build_receipt(root, bindings["control"])
    else:
        verify.verify_build_receipt(root, bindings["control"], "control")
    verify.verify_build_receipt(root, bindings["candidate"], "candidate")

    reports = {}
    receipts = {}
    row_maps = {}
    for lane in LANES:
        report, receipt, rows = _verify_lane(root, lane, protocol, bindings)
        reports[lane] = report
        receipts[lane] = receipt
        row_maps[lane] = rows
    for lane in LANES[1:]:
        require(set(row_maps[lane]) == set(row_maps[LANES[0]]), f"{lane}: row identity set differs")

    abba, abba_path = analyze.load_abba(root)
    comparator, comparator_path = analyze.load_comparator(root)
    strict_summary = None
    strict_status = "valid"
    strict_error = None
    try:
        # Pass the untouched reports.  In particular, do not use the full
        # guard's empty-vector comparison projection and do not erase source
        # counters to manufacture acceptance.
        strict_summary = abba.summarize_reports(reports=[reports[lane] for lane in LANES], profile="current-v1")
        require(len(strict_summary.get("results", [])) == EXPECTED_ROWS, "strict ABBA summary row count differs")
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        strict_status = "rejected"
        strict_error = str(error)

    descriptive = _descriptive_rows(reports, row_maps, comparator)
    oracles = _verify_oracles(row_maps, abba)
    original = _verify_original_full_guard(root, protocol)
    chronology = _chronology(root, protocol, receipts)
    return {
        "schema": SCHEMA,
        "scope": "Targeted 25-row non-XLSX/legacy-writer guard; 100 samples/five warmups; strict ABBA rejection and raw per-row descriptive statistics retained; no registered latency claim",
        "performance_claim": "none",
        "protocol": analyze.file_binding(root / "guard-protocol.json", root),
        "drivers": {
            "capture": analyze.file_binding(root / "guard_capture.py", root),
            "review": analyze.file_binding(root / "guard_review.py", root),
            "abba": analyze.file_binding(abba_path, root),
            "comparator": analyze.file_binding(comparator_path, root),
        },
        "bindings": {role: analyze.file_binding(root / f"{role}-binding.json", root) for role in ("control", "candidate")},
        "source_manifest_pair_verified": True,
        "chronology": chronology,
        "strict_abba_status": strict_status,
        "strict_abba_error": strict_error,
        "strict_abba": strict_summary,
        "descriptive_rows": descriptive,
        "oracles": oracles,
        "original_full_guard": original,
        "inputs": {
            "reports": [analyze.file_binding(root / lane / "report.json", root) for lane in LANES],
            "receipts": [analyze.file_binding(root / lane / "receipt.json", root) for lane in LANES],
        },
    }


def verify_summary(root: Path = ROOT) -> dict[str, Any]:
    """Public integration hook for the bundle verifier/evidence agent."""

    return evaluate(root)


def main() -> int:
    result = evaluate()
    (ROOT / "guard-review-summary.json").write_text(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"strict_abba_status": result["strict_abba_status"], "rows": len(result["descriptive_rows"])}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, KeyError, TypeError, ValueError, verify.VerificationError) as error:
        print(f"targeted guard review failed: {error}", file=sys.stderr)
        raise SystemExit(1)
