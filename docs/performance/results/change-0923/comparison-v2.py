#!/usr/bin/env python3
"""Validate and compare the prospective 0923 DOCX ingress routes.

The Rust harness is kept in external custody.  This module owns the small
report validator, route-specific counter contract, paired median/bootstrap
reduction, and planned-size check.  It does not build binaries or run native
children, and it never turns a validation result into an adoption decision.
"""

from __future__ import annotations

import json
import math
import random
import statistics
from pathlib import Path
from typing import Any, Iterable


CASE = "docx_source_backed_existing_section_layout_edit_save"
ROUTES = ("generic_read_at", "owned_vec", "owned_reader")
PHASES = ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")
GATES = (
    "phase_sum_verified",
    "semantic_reopen_all_sections_verified",
    "table_cell_sectpr_excluded_verified",
    "raw_untouched_members_verified",
    "raw_member_order_and_comment_verified",
    "header_footer_relationships_verified",
    "source_immutability_verified",
    "patch_forward_inverse_verified",
    "publication_inverse_exact_source_verified",
    "stale_foreign_signed_noop_limits_partial_sink_verified",
    "cache_counters_verified",
    "sink_counters_verified",
)
SOURCE_BYTES = 2_101_427
SOURCE_SHA256 = "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8"
OUTPUT_BYTES = 2_101_446
OUTPUT_SHA256 = "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
# These are the observed contracts of the unchanged generic source-backed
# route.  The prospective runner treats them as identity/configuration guards,
# rather than accepting any positive counter vector.
GENERIC_COUNTER_CONTRACT = {
    "read_calls": 202,
    "read_bytes": 8_409_488,
    "ordinary_payload_read_calls": 138,
    "ordinary_payload_read_bytes": 8_397_632,
    "max_in_flight_reads": 1,
    "ordinary_payload_materializations": 1,
}
GENERIC_CACHE_CONTRACT = {
    "source_cache_before_publication_after_commit_hits": 0,
    "source_cache_before_publication_after_commit_cold_loads": 1,
    "source_cache_before_publication_after_commit_successful_loads": 1,
    "source_cache_before_publication_after_commit_retained_entries": 1,
    "source_cache_before_publication_after_commit_retained_bytes": 18_855,
}
OWNED_READER_CALLS = 258
ROUTE_SCOPES = {
    "generic_read_at": "instrumented positional ReadAt calls and bytes; ordinary payload counters are source-observed",
    "owned_vec": "moved immutable Vec owner; no positional ReadAt or sequential reader observer is available",
    "owned_reader": "instrumented sequential reader ingress; no positional ReadAt observer is available after ownership transfer",
}
SINK_BUCKETS = (
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
)


class EvidenceError(RuntimeError):
    """A fail-closed report or comparison error."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def integer(value: object, label: str, *, positive: bool = False) -> int:
    require(type(value) is int, f"{label}: expected integer")
    require(value > 0 if positive else value >= 0, f"{label}: out of range")
    return int(value)


def _result(report: dict[str, Any], label: str) -> dict[str, Any]:
    require(report.get("schema_version") == 1, f"{label}: report schema changed")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}: configuration missing")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1, f"{label}: expected one result")
    result = results[0]
    require(isinstance(result, dict) and result.get("case") == CASE, f"{label}: case changed")
    return result


def _vector(value: object, length: int, label: str, *, positive: bool = False) -> list[int]:
    require(isinstance(value, list) and len(value) == length, f"{label}: cardinality changed")
    return [integer(item, f"{label}[{index}]", positive=positive) for index, item in enumerate(value)]


def _empty_vector(value: object, label: str) -> None:
    require(value in (None, []), f"{label}: unavailable counter was populated")


def _exact_vector(value: object, expected: int, samples: int, label: str) -> list[int]:
    values = _vector(value, samples, label)
    require(values == [expected] * samples,
            f"{label}: expected the fixed contract value {expected}, got {values[:3]}…")
    return values


def _validate_source_vectors(source: dict[str, Any], section: dict[str, Any], route: str,
                             samples: int, label: str) -> None:
    generic_keys = (
        "read_calls",
        "read_bytes",
        "ordinary_payload_read_calls",
        "ordinary_payload_read_bytes",
        "max_in_flight_reads",
    )
    if route == "generic_read_at":
        for key in generic_keys:
            values = _exact_vector(source.get(key), GENERIC_COUNTER_CONTRACT[key], samples,
                                   f"{label}.source.{key}")
            require(values == _vector(section.get("source_" + key.removeprefix("ordinary_payload_")), samples,
                                     f"{label}.section source {key}")
                    if key in ("read_calls", "read_bytes") else True,
                    f"{label}: source/section counter mismatch for {key}")
        materializations = _exact_vector(source.get("ordinary_payload_materializations"),
                                         GENERIC_COUNTER_CONTRACT["ordinary_payload_materializations"],
                                         samples, f"{label}.source.ordinary_payload_materializations")
        require(section.get("source_read_calls") == source.get("read_calls"),
                f"{label}: source read-call vectors differ")
        require(section.get("source_read_bytes") == source.get("read_bytes"),
                f"{label}: source read-byte vectors differ")
        require(section.get("ordinary_payload_materializations") == materializations,
                f"{label}: source materialization vectors differ")
        for key in ("owned_ingress_read_calls", "owned_ingress_read_bytes", "owned_ingress_owned_bytes"):
            _empty_vector(section.get(key), f"{label}.section.{key}")
        return

    for key in generic_keys:
        _empty_vector(source.get(key), f"{label}.source.{key}")
    _empty_vector(source.get("ordinary_payload_materializations"),
                  f"{label}.source.ordinary_payload_materializations")
    for key in ("source_read_calls", "source_read_bytes"):
        _empty_vector(section.get(key), f"{label}.section.{key}")
    _vector(section.get("ordinary_payload_materializations"), samples,
            f"{label}.section.ordinary_payload_materializations")
    require(section["ordinary_payload_materializations"] == [1] * samples,
            f"{label}: owned materialization contract changed")
    owned_calls = _vector(section.get("owned_ingress_read_calls"), samples,
                          f"{label}.section.owned_ingress_read_calls")
    owned_bytes = _vector(section.get("owned_ingress_read_bytes"), samples,
                          f"{label}.section.owned_ingress_read_bytes")
    owned_storage = _vector(section.get("owned_ingress_owned_bytes"), samples,
                            f"{label}.section.owned_ingress_owned_bytes", positive=True)
    require(owned_storage == [SOURCE_BYTES] * samples,
            f"{label}: owned storage byte identity changed")
    if route == "owned_vec":
        require(owned_calls == [0] * samples and owned_bytes == [0] * samples,
                f"{label}: owned-vector route fabricated reader observations")
    else:
        require(owned_calls == [OWNED_READER_CALLS] * samples,
                f"{label}: owned-reader call contract changed")
        require(owned_bytes == [SOURCE_BYTES] * samples,
                f"{label}: owned-reader byte count changed")


def validate_report(report: dict[str, Any], route: str, *, warmups: int, samples: int,
                    label: str = "report", binary: dict[str, Any] | None = None) -> dict[str, Any]:
    """Validate one complete harness report for one ingress route."""
    require(route in ROUTES, f"{label}: unknown route {route!r}")
    result = _result(report, label)
    configuration = report["configuration"]
    require(configuration.get("samples_per_case") == samples,
            f"{label}: sample count changed")
    require(configuration.get("warmup_iterations_per_case") == warmups,
            f"{label}: warmup count changed")
    require(configuration.get("cases") == [CASE], f"{label}: selected case changed")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{label}: binary identity missing")
    if binary is not None:
        require(identity.get("binary_sha256") == binary.get("sha256")
                and identity.get("binary_bytes") == binary.get("bytes"),
                f"{label}: binary identity differs from receipt")
    corpus = result.get("corpus")
    require(isinstance(corpus, dict), f"{label}: corpus manifest missing")
    expected_corpus = {
        "generator": "litchi-docx-section-layout-source-v1",
        "archive_bytes": SOURCE_BYTES,
        "archive_sha256": SOURCE_SHA256,
        "entry_count": 7,
        "archive_member_count": 10,
        "entry_bytes": 524288,
        "target_entry": "word/document.xml",
    }
    for key, expected in expected_corpus.items():
        require(corpus.get(key) == expected, f"{label}: corpus {key} changed")
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict), f"{label}: elapsed statistics missing")
    totals = _vector(elapsed.get("samples"), samples, f"{label}.elapsed_ns.samples", positive=True)
    order = elapsed.get("sample_order")
    require(isinstance(order, list) and sorted(order) == list(range(samples)),
            f"{label}: elapsed sample order changed")
    source = result.get("source")
    require(isinstance(source, dict), f"{label}: source summary missing")
    section = source.get("docx_section_layout")
    require(isinstance(section, dict), f"{label}: section summary missing")
    require(section.get("ingress_route") == route, f"{label}: ingress route changed")
    require(section.get("source_observation_scope") == ROUTE_SCOPES[route],
            f"{label}: source observation scope changed")
    identity_fields = {
        "implementation": "litchi-docx-source-backed-section-layout",
        "performance_claim": "none",
        "source_bytes": SOURCE_BYTES,
        "source_sha256": SOURCE_SHA256,
        "expected_output_sha256": OUTPUT_SHA256,
        "total_main_story_paragraph_count": 257,
        "section_count": 3,
        "paragraph_owned_section_positions": [64, 129],
        "selected_paragraph": 64,
        "columns_before": 2,
        "columns_after": 3,
        "media_count": 4,
        "media_bytes": 524288,
    }
    for key, expected in identity_fields.items():
        require(section.get(key) == expected, f"{label}: section field {key} changed")
    require(section.get("sample_order") == order, f"{label}: section sample order changed")
    phases: dict[str, list[int]] = {}
    for key in PHASES:
        phases[key] = _vector(section.get(key), samples, f"{label}.section.{key}", positive=True)
    for index, total in enumerate(totals):
        require(sum(phases[key][index] for key in PHASES) == total,
                f"{label}: six-phase sum mismatch at sample {index}")
    for key in GATES:
        require(section.get(key) is True, f"{label}: gate {key} failed")
    _validate_source_vectors(source, section, route, samples, label)
    cache_keys = (
        "source_cache_before_publication_after_commit_hits",
        "source_cache_before_publication_after_commit_cold_loads",
        "source_cache_before_publication_after_commit_successful_loads",
        "source_cache_before_publication_after_commit_retained_entries",
        "source_cache_before_publication_after_commit_retained_bytes",
    )
    cache: dict[str, list[int]] = {}
    for key in cache_keys:
        expected = GENERIC_CACHE_CONTRACT[key]
        cache[key] = _exact_vector(section.get(key), expected, samples, f"{label}.section.{key}")
    sink = result.get("sink")
    require(isinstance(sink, dict), f"{label}: sink summary missing")
    require(sink.get("accepted_bytes") == OUTPUT_BYTES, f"{label}: sink byte count changed")
    integer(sink.get("write_calls"), f"{label}.sink.write_calls", positive=True)
    require(integer(sink.get("largest_write"), f"{label}.sink.largest_write", positive=True) <= 65536,
            f"{label}: sink write bound changed")
    buckets = sink.get("write_size_buckets")
    require(isinstance(buckets, dict), f"{label}: sink bucket summary missing")
    bucket_values = {key: integer(buckets.get(key), f"{label}.sink.{key}") for key in SINK_BUCKETS}
    require(sum(bucket_values.values()) == sink["write_calls"], f"{label}: sink buckets do not sum")
    require(bucket_values["bytes_over_65536"] == 0, f"{label}: oversized sink write observed")
    outputs = section.get("output_sha256")
    require(outputs == [OUTPUT_SHA256] * samples, f"{label}: output digest changed")
    require(result.get("output_sha256") == OUTPUT_SHA256, f"{label}: top-level output digest changed")
    return {
        "label": label,
        "route": route,
        "samples": samples,
        "elapsed_ns": totals,
        "phase_ns": phases,
        "source": source,
        "section": section,
        "cache": cache,
        "sink": {"accepted_bytes": sink["accepted_bytes"], "write_calls": sink["write_calls"],
                  "largest_write": sink["largest_write"], "write_size_buckets": bucket_values},
        "gates": {key: section[key] for key in GATES},
    }


def nearest_rank(values: list[float], fraction: float) -> float:
    require(values, "empty statistic vector")
    ordered = sorted(values)
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def bootstrap(ratios: list[float], *, draws: int = 10000, seed: int = 923) -> dict[str, Any]:
    require(ratios, "empty paired ratio vector")
    rng = random.Random(seed)
    medians = [statistics.median(ratios[rng.randrange(len(ratios))] for _ in ratios)
               for _ in range(draws)]
    return {
        "draws": draws,
        "seed": seed,
        "statistic": "median of paired baseline_p50/candidate_p50 ratios",
        "ratios": ratios,
        "median_ratio": statistics.median(ratios),
        "confidence_interval_95": {
            "method": "nearest-rank 2.5th/97.5th percentiles",
            "lower": nearest_rank(medians, 0.025),
            "upper": nearest_rank(medians, 0.975),
        },
    }


def _p50(values: Iterable[int]) -> float:
    values = list(values)
    require(values, "empty elapsed vector")
    return float(statistics.median(values))


def _gate(summary: dict[str, Any], gate: str | None) -> dict[str, Any]:
    lower = float(summary["paired_bootstrap"]["confidence_interval_95"]["lower"])
    median = float(summary["paired_bootstrap"]["median_ratio"])
    if gate is None:
        eligible = None
    elif gate == "owned_benefit":
        eligible = median >= 1.05 and lower > 1.0
    elif gate == "generic_control":
        eligible = lower >= 1.0 / 1.05
    else:
        raise EvidenceError(f"unknown timing gate {gate!r}")
    summary["timing_gate"] = {"name": gate, "eligible": eligible, "median_ratio": median,
                               "lower_95": lower}
    return summary


def compare_reports(baseline: dict[str, Any], candidate: dict[str, Any], label: str,
                    *, route: str | None = None, gate: str | None = None,
                    warmups: int = 0, samples: int | None = None,
                    baseline_binary: dict[str, Any] | None = None,
                    candidate_binary: dict[str, Any] | None = None) -> dict[str, Any]:
    """Compare one baseline/candidate child pair.

    The ratio is baseline p50 divided by candidate p50, so values above one
    mean a candidate speedup.  ``gate`` is optional for controls and smoke
    reports; a validator pass never implies adoption.
    """
    if route is None:
        route = str(baseline.get("results", [{}])[0].get("source", {})
                   .get("docx_section_layout", {}).get("ingress_route", ""))
    require(route in ROUTES, f"{label}: route is missing")
    if samples is None:
        configuration = baseline.get("configuration", {})
        samples = int(configuration.get("samples_per_case", 0))
    base = validate_report(baseline, route, warmups=warmups, samples=samples, label=f"{label}.baseline", binary=baseline_binary)
    cand = validate_report(candidate, route, warmups=warmups, samples=samples, label=f"{label}.candidate", binary=candidate_binary)
    require(base["section"]["source_sha256"] == cand["section"]["source_sha256"],
            f"{label}: source corpus differs")
    require(base["section"]["expected_output_sha256"] == cand["section"]["expected_output_sha256"],
            f"{label}: expected output differs")
    baseline_p50 = _p50(base["elapsed_ns"])
    candidate_p50 = _p50(cand["elapsed_ns"])
    require(candidate_p50 > 0, f"{label}: candidate p50 is zero")
    summary = {
        "label": label,
        "route": route,
        "baseline_p50_ns": baseline_p50,
        "candidate_p50_ns": candidate_p50,
        "baseline_over_candidate_p50_ratio": baseline_p50 / candidate_p50,
        "validation": {"baseline": base["gates"], "candidate": cand["gates"]},
    }
    summary["paired_bootstrap"] = bootstrap([summary["baseline_over_candidate_p50_ratio"]])
    return _gate(summary, gate)


def compare_paired_cells(cells: Iterable[tuple[dict[str, Any], dict[str, Any]]], route: str,
                         label: str, *, gate: str | None = None, warmups: int = 0,
                         samples: int | None = None,
                         baseline_binary: dict[str, Any] | None = None,
                         candidate_binary: dict[str, Any] | None = None) -> dict[str, Any]:
    """Reduce already paired child reports for one route without recapture."""
    pairs = list(cells)
    require(pairs, f"{label}: no paired cells")
    details = [compare_reports(base, cand, f"{label}-cell-{index:03d}", route=route,
                               warmups=warmups, samples=samples,
                               baseline_binary=baseline_binary, candidate_binary=candidate_binary)
               for index, (base, cand) in enumerate(pairs, start=1)]
    ratios = [float(detail["baseline_over_candidate_p50_ratio"]) for detail in details]
    result: dict[str, Any] = {
        "schema": "litchi.performance.0923.paired-reduction.v2",
        "label": label,
        "route": route,
        "pairs": len(details),
        "pair_details": details,
        "paired_bootstrap": bootstrap(ratios),
        "performance_claim": "timing eligibility only; no production adoption",
    }
    return _gate(result, gate)


def planned_reduction_bytes(*, native_rows: int = 48, allocator_rows: int = 0,
                            rounds: int = 1) -> tuple[int, str]:
    """Serialize the full planned compact population and return size/hash.

    This is deliberately a synthetic schema exercise.  It contains phase,
    route-counter, cache, sink, and gate fields for every planned row, while
    retaining no raw report or per-sample payload.
    """
    rows: list[dict[str, Any]] = []
    total = native_rows + allocator_rows
    for index in range(total):
        route = ROUTES[index % len(ROUTES)]
        rows.append({
            "index": index,
            "kind": "native" if index < native_rows else "allocator",
            "route": route,
            "pair": (index % 8) + 1,
            "rounds": [
                {
                    "elapsed_ns": 4_000_000 + index * 17 + round_index,
                    "phase_ns": [37521, 45910, 1337885, 55891, 355431, 2693711],
                    "source_read_calls": 202 if route == "generic_read_at" else None,
                    "source_read_bytes": 8409488 if route == "generic_read_at" else None,
                    "owned_ingress_read_calls": (258 if route == "owned_reader" else 0),
                    "owned_ingress_read_bytes": (SOURCE_BYTES if route == "owned_reader" else 0),
                    "owned_ingress_owned_bytes": (SOURCE_BYTES if route != "generic_read_at" else 0),
                    "ordinary_payload_materializations": 1,
                    "cache": {"cold_loads": 1, "successful_loads": 1, "retained_entries": 1,
                              "retained_bytes": 18855},
                    "sink": {"accepted_bytes": OUTPUT_BYTES, "largest_write": 65536},
                    "gates": True,
                }
                for round_index in range(rounds)
            ],
        })
    data = json.dumps({"schema": "litchi.performance.0923.compact-reduction.v1", "rows": rows},
                      sort_keys=True, separators=(",", ":")).encode("utf-8")
    import hashlib
    return len(data), hashlib.sha256(data).hexdigest()
