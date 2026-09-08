#!/usr/bin/env python3
"""Derive the retained paired PPTX metadata-spool comparison.

The analyzer is deliberately data-only.  It reads the frozen protocol and
the retained producer reports, checks their identity/oracle fields, and emits
the descriptive summary used by the portable verifier.  Normal timings and
allocator observations are kept in separate rows; neither is promoted to a
registered latency or whole-process memory claim.
"""

from __future__ import annotations

import json
import math
import re
import statistics
from pathlib import Path
from typing import Any, Mapping


ROOT = Path(__file__).resolve().parent
PROTOCOL_SCHEMA = "pptx-metadata-spool-capture-v1"
REPORT_SCHEMA = "pptx-metadata-spool-v1"
SUMMARY_SCHEMA = "pptx-metadata-spool-summary-v1"
COUNTS = (8, 256, 8192)
POLICIES = ("control", "spool")
INSTRUMENTATIONS = ("normal", "allocator")
SAMPLES = 30
WARMUPS = 3
CAPTURE_REPEATS = (1, 2)
MAX_SPOOL_BYTES = 64 * 1024 * 1024
SPOOL_BUFFER_BYTES = 16 * 1024
SHA256 = re.compile(r"^[0-9a-f]{64}$")
# These are the retained 0474/0476 corpus identities recorded in
# measurement-review.md.  Reports must bind to these bytes and semantic/text
# digests; a same-run self-consistent but different corpus is insufficient.
PRECEDING_EVIDENCE = "change-0476/summary.json"
PRECEDING_EVIDENCE_SHA256 = (
    "8dde86118498e8b00405e9db155f3e54d2704dc8ed8cc9e0e78c5dbc8dd81860"
)
HISTORICAL_CORPUS_IDENTITIES = {
    8: {
        "entry_count": 53,
        "source_archive_bytes": 36259,
        "source_archive_sha256": "951505889af106f032241c30f07b5d237e54822dade768c911aca2d0f68c22c5",
        "semantic_sha256": "f3444404ef5130757c79ae624a161a38bc7c2f7ed053b47315ccf31067573a4e",
        "full_text_sha256": "bbc9f0e6cf3b3c48cd991dca9dded766cec7c93f762dd74b1559a34cdc05a966",
    },
    256: {
        "entry_count": 549,
        "source_archive_bytes": 274398,
        "source_archive_sha256": "1f33f8b2c36a2a51abc62d827e4c915dd3b4e52859b300323f250ab561942cf2",
        "semantic_sha256": "147697c24b54b91e37e9906330802f40a2c4f9da1af6b564ca92eca6faad4f3b",
        "full_text_sha256": "4c4a1a185cd33c9a3362bed00bd9e77cd64f898976a57b56220ea62bb3a614ee",
    },
    8192: {
        "entry_count": 16421,
        "source_archive_bytes": 7940406,
        "source_archive_sha256": "c7b08da644e651046d368b1baaff9a12c6d7218c4f96914dacb1033722e4b527",
        "semantic_sha256": "1bf460386d6f8962d4a04444a5e2b3971e6548fc1273cc933a6beefff9bf1417",
        "full_text_sha256": "521f638a72f55371d0d40ddf01dac1577f1e584f58ae535bb0cc01f90b76e1f9",
    },
}
MEMORY_GATE_TEXT = (
    "Across 8/256/8192 slides, generated-route operation peak range must be within 1% "
    "of its smallest-count peak, with zero failed allocations and zero live exit delta; "
    "larger growth requires source/heap investigation before any bounded-window claim. "
    "Descriptor, active-name and integer-width effects must be accounted for."
)


class AnalysisError(ValueError):
    """The retained data cannot support the declared comparison."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def reject_constant(value: str) -> Any:
    raise AnalysisError(f"non-finite JSON value {value}")


def reject_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            fail(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=reject_pairs,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, AnalysisError) as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
        stream.write("\n")


def integer(value: Any, label: str, expected: int | None = None) -> int:
    if not isinstance(value, int) or isinstance(value, bool) or value < 0:
        fail(f"{label}: expected a non-negative integer")
    if expected is not None and value != expected:
        fail(f"{label}: expected {expected}, got {value}")
    return value


def digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or SHA256.fullmatch(value) is None:
        fail(f"{label}: expected a lower-case SHA-256 digest")
    return value


def expected_captures() -> list[dict[str, Any]]:
    forward = [
        (instrumentation, count, policy)
        for instrumentation in INSTRUMENTATIONS
        for count in COUNTS
        for policy in POLICIES
    ]
    result: list[dict[str, Any]] = []
    for repeat, sequence in ((1, forward), (2, list(reversed(forward)))):
        for instrumentation, count, policy in sequence:
            result.append({
                "label": f"r{repeat}-{instrumentation}-{count}-{policy}",
                "instrumentation": instrumentation,
                "count": count,
                "policy": policy,
                "repeat": repeat,
            })
    return result


def expected_spool_scratch_bytes(count: int) -> int:
    """Return the independently derived extent for the fixed ZIP-name set."""
    integer(count, "slide count")
    return 2914 + 143 * count + 2 * sum(
        len(str(index)) for index in range(1, count + 1)
    )


def protocol_rows(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    if protocol.get("schema") != PROTOCOL_SCHEMA:
        fail("protocol schema differs")
    integer(protocol.get("samples"), "protocol.samples", SAMPLES)
    integer(protocol.get("warmups"), "protocol.warmups", WARMUPS)
    integer(protocol.get("cpu"), "protocol.cpu", 2)
    if protocol.get("preceding_evidence") != PRECEDING_EVIDENCE:
        fail("protocol.preceding_evidence differs from the frozen corpus evidence")
    if protocol.get("preceding_evidence_sha256") != PRECEDING_EVIDENCE_SHA256:
        fail("protocol.preceding_evidence_sha256 differs from the frozen corpus evidence")
    if protocol.get("memory_gate") != MEMORY_GATE_TEXT:
        fail("protocol.memory_gate differs from the frozen acceptance rule")
    if protocol.get("normal_and_allocator_timings_separate") is not True:
        fail("protocol must separate normal and allocator timings")
    captures = protocol.get("captures")
    expected = expected_captures()
    if not isinstance(captures, list) or len(captures) != len(expected):
        fail(f"protocol must contain exactly {len(expected)} captures")
    rows: list[dict[str, Any]] = []
    for index, (actual, wanted) in enumerate(zip(captures, expected)):
        if not isinstance(actual, dict):
            fail(f"protocol.captures[{index}] must be an object")
        for key, value in wanted.items():
            if actual.get(key) != value:
                fail(f"protocol.captures[{index}].{key} differs")
        rows.append(dict(actual))
    if len({row["label"] for row in rows}) != len(rows):
        fail("protocol capture labels are not unique")
    return rows


def _text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected a non-empty string")
    return value


def _sample_stats(values: list[int]) -> dict[str, Any]:
    if len(values) != SAMPLES or any(type(value) is not int or value < 0 for value in values):
        fail("sample statistic input is invalid")
    ordered = sorted(values)
    mean = statistics.mean(values)
    half = 2.045 * statistics.stdev(values) / math.sqrt(len(values))
    return {
        "n": len(values),
        "mean": mean,
        "minimum": min(values),
        "maximum": max(values),
        "p50": ordered[math.ceil(len(values) * 0.50) - 1],
        "p95": ordered[math.ceil(len(values) * 0.95) - 1],
        "p99": ordered[math.ceil(len(values) * 0.99) - 1],
        "mean_t95_interval": [mean - half, mean + half],
    }


def _percent(before: float, after: float) -> float | None:
    return 100 * (after - before) / before if before else None


def _check_process(value: Any, label: str) -> None:
    if not isinstance(value, dict):
        fail(f"{label}: process observer is missing")
    for key in (
        "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
        "syscr", "syscw", "minor_faults", "major_faults", "user_cpu_ticks",
        "system_cpu_ticks", "clock_ticks_per_second", "voluntary_context_switches",
        "nonvoluntary_context_switches", "rss_bytes", "peak_rss_bytes",
    ):
        integer(value.get(key), f"{label}.{key}")


def _check_allocation(value: Any, instrumentation: str, label: str) -> None:
    if instrumentation == "normal":
        if value is not None:
            fail(f"{label}: normal samples must omit allocation metrics")
        return
    if not isinstance(value, dict):
        fail(f"{label}: allocator sample is missing")
    if value.get("status") != "measured":
        fail(f"{label}: allocator sample is not measured")
    if value.get("scope") != "operation_global_system_allocator":
        fail(f"{label}: allocator scope differs")
    fields = (
        "allocation_calls", "deallocation_calls", "reallocation_calls",
        "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
        "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
        "peak_live_bytes_after", "region_peak_live_bytes",
    )
    for field in fields:
        integer(value.get(field), f"{label}.{field}")
    if value["failed_allocation_calls"] != 0:
        fail(f"{label}: failed allocation count is non-zero")
    if value["live_bytes_before"] != value["live_bytes_after"]:
        fail(f"{label}: operation did not return to the same live-byte level")
    if value["region_peak_live_bytes"] < value["live_bytes_before"]:
        fail(f"{label}: region peak is below its starting live bytes")


def _check_case(case: Any, spec: Mapping[str, Any], label: str) -> tuple[int, str]:
    if not isinstance(case, dict):
        fail(f"{label}: expected an object")
    if case.get("mode") not in POLICIES:
        fail(f"{label}.mode differs")
    if case.get("slide_count") != spec["count"]:
        fail(f"{label}.slide_count differs")
    integer(case.get("entry_count"), f"{label}.entry_count")
    integer(case.get("source_archive_bytes"), f"{label}.source_archive_bytes")
    digest(case.get("source_archive_sha256"), f"{label}.source_archive_sha256")
    output_bytes = integer(case.get("output_bytes"), f"{label}.output_bytes")
    output_hash = digest(case.get("output_sha256"), f"{label}.output_sha256")
    for field in (
        "byte_exact_control_match", "every_physical_member_verified",
        "every_slide_semantic_verified", "presentation_graph_verified",
        "slide_geometry_verified", "text_digest_verified",
    ):
        if case.get(field) is not True:
            fail(f"{label}.{field}: oracle proof is missing")
    scratch = case.get("scratch_bytes")
    if case["mode"] == "control":
        if scratch is not None:
            fail(f"{label}: control oracle has scratch bytes")
    else:
        integer(scratch, f"{label}.scratch_bytes")
        expected_scratch = expected_spool_scratch_bytes(int(spec["count"]))
        if scratch != expected_scratch:
            fail(
                f"{label}.scratch_bytes: expected independently derived extent "
                f"{expected_scratch}, got {scratch}"
            )
    return output_bytes, output_hash


def validate_report(report: Mapping[str, Any], spec: Mapping[str, Any], path: str = "report") -> dict[str, Any]:
    if report.get("schema") != REPORT_SCHEMA:
        fail(f"{path}: report schema differs")
    expected_instrumentation = (
        "none" if spec["instrumentation"] == "normal" else "system_allocator_operation_scoped"
    )
    if report.get("instrumentation") != expected_instrumentation:
        fail(f"{path}: instrumentation identity differs")
    expected_allocator = (
        "Rust system allocator" if spec["instrumentation"] == "normal"
        else "CountingSystemAllocator(std::alloc::System)"
    )
    if report.get("allocator") != expected_allocator:
        fail(f"{path}: allocator identity differs")
    if (report.get("samples"), report.get("warmups"), report.get("repeats")) != (SAMPLES, WARMUPS, 1):
        fail(f"{path}: sample protocol differs")
    if report.get("counts") != [spec["count"]] or report.get("modes") != [spec["policy"]]:
        fail(f"{path}: selected count or mode differs")
    if report.get("spool_max_bytes") != MAX_SPOOL_BYTES:
        fail(f"{path}: spool maximum differs")
    if report.get("spool_buffer_bytes") != SPOOL_BUFFER_BYTES:
        fail(f"{path}: spool buffer differs")
    for field in (
        "timing_scope", "oracle_scope", "cleanup_scope",
        "control_storage_policy", "spool_storage_policy",
    ):
        _text(report.get(field), f"{path}.{field}")
    timing = report["timing_scope"]
    for phrase in ("timed operation", "public writer", "File", "finish", "flush", "close"):
        if phrase.lower() not in timing.lower():
            fail(f"{path}.timing_scope omits {phrase!r}")
    for phrase in (
        "oracle", "byte-compared", "physical-member", "per-slide",
        "text/geometry", "relationship-graph",
    ):
        if phrase.lower() not in report["oracle_scope"].lower():
            fail(f"{path}.oracle_scope omits {phrase!r}")
    cleanup = report["cleanup_scope"].lower()
    if "unlink" not in cleanup and "removed" not in cleanup:
        fail(f"{path}.cleanup_scope omits spool removal boundary")

    corpora = report.get("corpora")
    if not isinstance(corpora, list) or len(corpora) != 1:
        fail(f"{path}: expected one selected corpus")
    corpus = corpora[0]
    if not isinstance(corpus, dict) or corpus.get("slide_count") != spec["count"]:
        fail(f"{path}.corpora: corpus identity differs")
    historical = HISTORICAL_CORPUS_IDENTITIES.get(spec["count"])
    if historical is None:
        fail(f"{path}.corpora: count has no retained historical identity")
    if corpus.get("entry_count") != historical["entry_count"]:
        fail(f"{path}.corpora.entry_count differs")
    for field in ("input_text_bytes", "source_archive_bytes"):
        integer(corpus.get(field), f"{path}.corpora[0].{field}")
    for field in ("source_archive_sha256", "semantic_sha256", "full_text_sha256"):
        digest(corpus.get(field), f"{path}.corpora[0].{field}")
    for field in (
        "entry_count", "source_archive_bytes", "source_archive_sha256",
        "semantic_sha256", "full_text_sha256",
    ):
        if corpus[field] != historical[field]:
            fail(f"{path}.corpora.{field}: historical identity differs")

    cases = report.get("cases")
    if not isinstance(cases, list) or len(cases) != 2:
        fail(f"{path}: expected control and spool oracle cases")
    projections: set[tuple[int, str]] = set()
    case_modes: set[str] = set()
    case_scratch: dict[str, int | None] = {}
    for index, case in enumerate(cases):
        projection = _check_case(case, spec, f"{path}.cases[{index}]")
        projections.add(projection)
        case_modes.add(case["mode"])
        case_scratch[case["mode"]] = case.get("scratch_bytes")
        if case.get("source_archive_bytes") != corpus["source_archive_bytes"]:
            fail(f"{path}.cases[{index}].source_archive_bytes differs from corpus")
        if case.get("source_archive_sha256") != corpus["source_archive_sha256"]:
            fail(f"{path}.cases[{index}].source_archive_sha256 differs from corpus")
        if case.get("entry_count") != corpus["entry_count"]:
            fail(f"{path}.cases[{index}].entry_count differs from corpus")
    if case_modes != set(POLICIES) or len(projections) != 1:
        fail(f"{path}: control and spool oracle projections differ")

    operations = report.get("operations")
    if not isinstance(operations, list) or len(operations) != SAMPLES:
        fail(f"{path}: expected {SAMPLES} measured operations")
    if {row.get("sample") for row in operations if isinstance(row, dict)} != set(range(SAMPLES)):
        fail(f"{path}: measured sample indices are incomplete")
    projection = next(iter(projections))
    for index, row in enumerate(operations):
        label = f"{path}.operations[{index}]"
        if not isinstance(row, dict):
            fail(f"{label}: expected an object")
        if (row.get("mode"), row.get("slide_count"), row.get("entry_count"), row.get("repeat")) != (
            spec["policy"], spec["count"], 37 + 2 * spec["count"], 0
        ):
            fail(f"{label}: operation identity differs")
        integer(row.get("sample"), f"{label}.sample")
        integer(row.get("elapsed_ns"), f"{label}.elapsed_ns")
        integer(row.get("output_bytes"), f"{label}.output_bytes")
        integer(row.get("output_write_calls"), f"{label}.output_write_calls")
        output_hash = digest(row.get("output_sha256"), f"{label}.output_sha256")
        if row.get("output_matches_oracle") is not True or (row["output_bytes"], output_hash) != projection:
            fail(f"{label}: output oracle failed")
        if row.get("source_archive_bytes") != corpus["source_archive_bytes"]:
            fail(f"{label}: source archive bytes differ")
        if row.get("source_archive_sha256") != corpus["source_archive_sha256"]:
            fail(f"{label}: source archive digest differs")
        integer(row.get("input_text_bytes"), f"{label}.input_text_bytes")
        scratch = row.get("scratch_bytes")
        if spec["policy"] == "control":
            if scratch is not None:
                fail(f"{label}: control operation has scratch bytes")
        else:
            integer(scratch, f"{label}.scratch_bytes")
            expected_scratch = expected_spool_scratch_bytes(int(spec["count"]))
            if scratch != expected_scratch:
                fail(
                    f"{label}: scratch extent differs from independently derived "
                    f"extent {expected_scratch}"
                )
            if scratch != case_scratch["spool"]:
                fail(f"{label}: scratch extent differs from spool oracle")
        _check_allocation(row.get("allocation"), spec["instrumentation"], f"{label}.allocation")
        _check_process(row.get("process"), f"{label}.process")
    return dict(report)


def _capture_report(spec: Mapping[str, Any]) -> dict[str, Any]:
    path = ROOT / "captures" / f"{spec['label']}.report.json"
    if not path.is_file():
        fail(f"{spec['label']}: report is missing")
    value = read_json(path)
    if not isinstance(value, dict):
        fail(f"{spec['label']}: report must be an object")
    return validate_report(value, spec, str(path))


def _resource_rss(spec: Mapping[str, Any]) -> int:
    path = ROOT / "captures" / f"{spec['label']}.resource"
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as error:
        fail(f"{spec['label']}: resource file cannot be read: {error}")
    match = re.search(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", text, re.MULTILINE)
    if match is None:
        fail(f"{spec['label']}: resource RSS is missing")
    return int(match.group(1))


def derive() -> dict[str, Any]:
    protocol = read_json(ROOT / "protocol.json")
    if not isinstance(protocol, dict):
        fail("protocol must be an object")
    rows = protocol_rows(protocol)
    reports: dict[str, dict[str, Any]] = {}
    rss: dict[str, int] = {}
    identities: dict[str, tuple[Any, ...]] = {}
    derived_rows: dict[str, dict[str, Any]] = {}
    for spec in rows:
        label = spec["label"]
        report = _capture_report(spec)
        reports[label] = report
        rss[label] = _resource_rss(spec)
        corpus = report["corpora"][0]
        projection = (report["cases"][0]["output_bytes"], report["cases"][0]["output_sha256"])
        identity = (
            corpus["source_archive_sha256"], corpus["source_archive_bytes"],
            corpus["input_text_bytes"], corpus["semantic_sha256"],
            corpus["full_text_sha256"], corpus["entry_count"], *projection,
        )
        key = str(spec["count"])
        if key in identities and identities[key] != identity:
            fail(f"{key}: source/output identity changed across captures")
        identities[key] = identity
        operations = report["operations"]
        row: dict[str, Any] = {
            "capture": spec,
            "normal_timing": spec["instrumentation"] == "normal",
            "elapsed_ns": _sample_stats([item["elapsed_ns"] for item in operations]),
            "output_bytes": projection[0],
            "output_sha256": projection[1],
            "output_write_calls": _sample_stats([item["output_write_calls"] for item in operations]),
            "process_max_rss_kib": rss[label],
        }
        scratch_values = [item["scratch_bytes"] for item in operations]
        if spec["policy"] == "spool":
            if any(type(value) is not int or value < 0 for value in scratch_values):
                fail(f"{label}: spool scratch sample is invalid")
            row["scratch_bytes"] = _sample_stats(scratch_values)
        else:
            if any(value is not None for value in scratch_values):
                fail(f"{label}: control scratch sample is present")
            row["scratch_bytes"] = None
        processes = [item["process"] for item in operations]
        row["process_observer_deltas"] = {
            field: _sample_stats([process[field] for process in processes])
            for field in (
                "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
                "syscr", "syscw", "minor_faults", "major_faults", "user_cpu_ticks",
                "system_cpu_ticks", "voluntary_context_switches", "nonvoluntary_context_switches",
                "rss_bytes", "peak_rss_bytes",
            )
        }
        if spec["instrumentation"] == "allocator":
            samples = [item["allocation"] for item in operations]
            row["allocation_calls"] = _sample_stats([item["allocation_calls"] for item in samples])
            row["requested_bytes"] = _sample_stats([item["allocated_bytes"] for item in samples])
            row["deallocated_bytes"] = _sample_stats([item["deallocated_bytes"] for item in samples])
            row["incremental_peak_live_bytes"] = _sample_stats([
                item["region_peak_live_bytes"] - item["live_bytes_before"]
                for item in samples
            ])
        derived_rows[label] = row

    pairs: list[dict[str, Any]] = []
    drifts: list[dict[str, Any]] = []
    for instrumentation in INSTRUMENTATIONS:
        for count in COUNTS:
            for repeat in CAPTURE_REPEATS:
                before = derived_rows[f"r{repeat}-{instrumentation}-{count}-control"]
                after = derived_rows[f"r{repeat}-{instrumentation}-{count}-spool"]
                changes: dict[str, float | None] = {
                    "process_max_rss_kib": _percent(before["process_max_rss_kib"], after["process_max_rss_kib"]),
                }
                if instrumentation == "normal":
                    changes.update({
                        f"elapsed_{metric}": _percent(before["elapsed_ns"][metric], after["elapsed_ns"][metric])
                        for metric in ("mean", "p50", "p95", "p99")
                    })
                else:
                    changes.update({
                        metric: _percent(before[metric]["mean"], after[metric]["mean"])
                        for metric in ("allocation_calls", "requested_bytes", "incremental_peak_live_bytes")
                    })
                pairs.append({
                    "instrumentation": instrumentation,
                    "count": count,
                    "repeat": repeat,
                    "percent_changes": changes,
                    "positive_review_flags": [
                        key for key, value in changes.items()
                        if value is not None and value > 5
                    ],
                })
            for policy in POLICIES:
                first = derived_rows[f"r1-{instrumentation}-{count}-{policy}"]
                second = derived_rows[f"r2-{instrumentation}-{count}-{policy}"]
                changes = {
                    "process_max_rss_kib": _percent(first["process_max_rss_kib"], second["process_max_rss_kib"]),
                }
                if instrumentation == "normal":
                    changes.update({
                        f"elapsed_{metric}": _percent(first["elapsed_ns"][metric], second["elapsed_ns"][metric])
                        for metric in ("mean", "p50", "p95", "p99")
                    })
                drifts.append({
                    "instrumentation": instrumentation,
                    "count": count,
                    "policy": policy,
                    "percent_changes": changes,
                    "absolute_review_flags": [
                        key for key, value in changes.items()
                        if value is not None and abs(value) > 5
                    ],
                })
    lane_maximum_peaks = {
        f"r{repeat}-allocator-{count}-spool": derived_rows[
            f"r{repeat}-allocator-{count}-spool"
        ]["incremental_peak_live_bytes"]["maximum"]
        for repeat in CAPTURE_REPEATS
        for count in COUNTS
    }
    baseline_minimum_peaks = [
        derived_rows[f"r{repeat}-allocator-{COUNTS[0]}-spool"][
            "incremental_peak_live_bytes"
        ]["minimum"]
        for repeat in CAPTURE_REPEATS
    ]
    global_minimum = min(
        derived_rows[f"r{repeat}-allocator-{count}-spool"][
            "incremental_peak_live_bytes"
        ]["minimum"]
        for repeat in CAPTURE_REPEATS
        for count in COUNTS
    )
    smallest_peak = min(baseline_minimum_peaks)
    largest_peak = max(lane_maximum_peaks.values())
    range_percent = (
        100 * (largest_peak - global_minimum) / smallest_peak
        if smallest_peak else None
    )
    within_threshold = range_percent is not None and range_percent <= 1
    memory_gate = {
        "metric": "allocator operation incremental peak live bytes",
        "instrumentation": "allocator",
        "policy": "spool",
        "counts": list(COUNTS),
        "repeats": list(CAPTURE_REPEATS),
        "lane_maximum_peaks": lane_maximum_peaks,
        "baseline_count": COUNTS[0],
        "baseline_minimum_peaks": baseline_minimum_peaks,
        "smallest_count_baseline_minimum": smallest_peak,
        "global_minimum": global_minimum,
        "global_maximum": largest_peak,
        "range_percent": range_percent,
        "threshold_percent": 1,
        "within_threshold": within_threshold,
        "zero_failed_allocations": True,
        "zero_live_exit_delta": True,
        "status": "pass" if within_threshold else "requires_investigation",
    }
    return {
        "schema": SUMMARY_SCHEMA,
        "captures": len(derived_rows),
        "samples": len(derived_rows) * SAMPLES,
        "rows": derived_rows,
        "identities": {key: list(value) for key, value in identities.items()},
        "pairs": pairs,
        "repeat_drifts": drifts,
        "memory_gate": memory_gate,
        "uncertainty": "nearest-rank percentiles; mean interval uses t(29)=2.045; independent process repeats retained, no registered latency claim",
    }


def main() -> int:
    try:
        write_json(ROOT / "summary.json", derive())
    except (AnalysisError, OSError) as error:
        print(f"analyze.py: FAIL: {error}")
        return 1
    print("analyze.py: wrote summary.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
