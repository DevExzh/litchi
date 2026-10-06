#!/usr/bin/env python3
"""Validate and reduce the 0925 descriptive DOCX provider reports.

This module never starts a benchmark process.  It validates the existing
``litchi-perf-baseline`` report envelope, the fixed section-layout corpus, the
six timed phases, and route-specific provider observations.  A later capture
runner may use these functions, but a validation pass never authorizes a
capture or a provider ranking.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
from pathlib import Path
from typing import Any, Iterable


CASE = "docx_source_backed_existing_section_layout_edit_save"
ROUTES = (
    "generic_read_at",
    "owned_vec",
    "owned_reader",
    "file_warm",
    "range_short",
    "range_zero",
    "range_delayed",
)
CONTROL_ROUTES = ROUTES[:3]
PROVIDER_ROUTES = ROUTES[3:]
PHASES = (
    "preparation_ns",
    "open_ns",
    "query_ns",
    "edit_ns",
    "commit_ns",
    "publication_ns",
)
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
REQUEST_BUCKETS = 18
MAX_ARTIFACT_BYTES = 1 << 20
CPU = 12

ROUTE_SCOPES = {
    "generic_read_at": "instrumented positional ReadAt calls and bytes; ordinary payload counters are source-observed",
    "owned_vec": "moved immutable Vec owner; no positional ReadAt or sequential reader observer is available",
    "owned_reader": "instrumented sequential reader ingress; no positional ReadAt observer is available after ownership transfer",
    "file_warm": "warm recently-written FileSource positional ReadAt wrapper; counters are logical reads, not physical filesystem I/O",
    "range_short": "PptxRangeSource logical requests over an immutable in-memory backing source; delegated counters are wrapper observations, not physical or network I/O",
    "range_zero": "PptxRangeSource logical requests over an immutable in-memory backing source with zero fixed delay; counters are not physical or network I/O",
    "range_delayed": "PptxRangeSource logical requests over an immutable in-memory backing source with configured service delay; counters are not physical or network I/O",
}

PROVIDER_SOURCE_SCOPES = {
    "file_warm": (
        "staged regular file written and synced before the run; each FileSource is "
        "opened during sample preparation; warm recently-written observation, not a "
        "filesystem-cold result"
    ),
    "range_short": (
        "immutable OwnedSource memory backing beneath the benchmark-only PptxRangeSource; "
        "simulated range/service behavior, with no network or physical-storage claim"
    ),
    "range_zero": (
        "immutable OwnedSource memory backing beneath the benchmark-only PptxRangeSource; "
        "simulated range/service behavior, with no network or physical-storage claim"
    ),
    "range_delayed": (
        "immutable OwnedSource memory backing beneath the benchmark-only PptxRangeSource; "
        "simulated range/service behavior, with no network or physical-storage claim"
    ),
}

GENERIC_COUNTER_CONTRACT = {
    "read_calls": 202,
    "read_bytes": 8_409_488,
    "ordinary_payload_read_calls": 138,
    "ordinary_payload_read_bytes": 8_397_632,
    "max_in_flight_reads": 1,
    "ordinary_payload_materializations": 1,
}
GENERIC_OPERATION_END_COUNTER_CONTRACT = {
    "operation_end_source_read_calls": 169,
    "operation_end_source_read_bytes": 6_308_061,
}
GENERIC_CACHE_CONTRACT = {
    "source_cache_before_publication_after_commit_hits": 0,
    "source_cache_before_publication_after_commit_cold_loads": 1,
    "source_cache_before_publication_after_commit_successful_loads": 1,
    "source_cache_before_publication_after_commit_retained_entries": 1,
    "source_cache_before_publication_after_commit_retained_bytes": 18_855,
}
OWNED_READER_CALLS = 258


class EvidenceError(RuntimeError):
    """A report or reduced-evidence contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    size = path.stat().st_size
    return {"path": str(path), "bytes": size, "sha256": sha256_path(path)}


def read_json(path: Path, *, max_bytes: int | None = MAX_ARTIFACT_BYTES) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing report: {path}")
    size = path.stat().st_size
    if max_bytes is not None:
        require(size < max_bytes, f"report is at least 1 MiB: {path} ({size} bytes)")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise EvidenceError(f"cannot parse JSON report {path}: {error}") from error
    require(isinstance(value, dict), f"report is not an object: {path}")
    return value


def _integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(type(value) is int, f"{label}: expected integer")
    require(value > 0 if positive else value >= 0, f"{label}: out of range")
    return int(value)


def _vector(value: Any, length: int, label: str, *, positive: bool = False) -> list[int]:
    require(isinstance(value, list) and len(value) == length, f"{label}: cardinality changed")
    return [_integer(item, f"{label}[{index}]", positive=positive)
            for index, item in enumerate(value)]


def _empty(value: Any, label: str) -> None:
    require(value is None or value == [], f"{label}: unavailable counter was populated")


def _exact_vector(value: Any, expected: int, length: int, label: str) -> list[int]:
    values = _vector(value, length, label)
    require(values == [expected] * length,
            f"{label}: expected fixed value {expected}, got {values[:3]}")
    return values


def _percentile(values: list[int], fraction: float) -> float:
    ordered = sorted(values)
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return float(ordered[index])


def _statistics(values: list[int]) -> dict[str, float | int]:
    require(values, "empty timing vector")
    return {
        "min": min(values),
        "p50": float(statistics.median(values)),
        "p95": _percentile(values, 0.95),
        "p99": _percentile(values, 0.99),
        "max": max(values),
        "mean": statistics.fmean(values),
    }


def _result(report: dict[str, Any], label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    require(report.get("schema_version") == 1, f"{label}: schema_version changed")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}: configuration missing")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label}: expected one result")
    result = results[0]
    require(isinstance(result, dict) and result.get("case") == CASE,
            f"{label}: case changed")
    return configuration, result


def _corpus(result: dict[str, Any], label: str) -> None:
    corpus = result.get("corpus")
    require(isinstance(corpus, dict), f"{label}: corpus missing")
    expected = {
        "name": "docx-source-backed-existing-section-layout",
        "generator": "litchi-docx-section-layout-source-v1",
        "package_format": "DOCX/OPC/ZIP",
        "shape": "existing-section-layout",
        "payload_kind": "deterministic-incompressible-media",
        "compression": "deflate",
        "entry_count": 7,
        "archive_member_count": 10,
        "entry_bytes": 524_288,
        "uncompressed_payload_bytes": 2_116_441,
        "archive_bytes": SOURCE_BYTES,
        "archive_sha256": SOURCE_SHA256,
        "target_entry": "word/document.xml",
        "target_payload_bytes": 18_855,
        "target_payload_sha256": "67cb9066f112bfee6905a23f3eabdc76b65c86cd0c6613dd60a5a18895ceac81",
        "xlsx": None,
    }
    for key, value in expected.items():
        require(corpus.get(key) == value, f"{label}: corpus {key} changed")


def _validate_observation(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: observation missing")
    fields = (
        "read_calls", "requested_bytes", "returned_bytes", "min_request_bytes",
        "max_request_bytes", "short_reads", "delayed_calls", "transfer_paced_calls",
        "transfer_delay_ns", "request_size_counts",
    )
    for field in fields:
        require(field in value, f"{label}: field {field} missing")
    calls = _integer(value["read_calls"], f"{label}.read_calls")
    requested = _integer(value["requested_bytes"], f"{label}.requested_bytes")
    returned = _integer(value["returned_bytes"], f"{label}.returned_bytes")
    require(returned <= requested, f"{label}: returned bytes exceed requested bytes")
    for field in ("short_reads", "delayed_calls", "transfer_paced_calls", "transfer_delay_ns"):
        _integer(value[field], f"{label}.{field}")
    require(value["short_reads"] <= calls
            and value["delayed_calls"] <= calls
            and value["transfer_paced_calls"] <= calls,
            f"{label}: event counters exceed read calls")
    counts = value["request_size_counts"]
    require(isinstance(counts, list) and len(counts) == REQUEST_BUCKETS,
            f"{label}.request_size_counts: bucket count changed")
    counts = [_integer(item, f"{label}.request_size_counts[{index}]")
              for index, item in enumerate(counts)]
    require(sum(counts) == calls, f"{label}: request histogram does not conserve calls")
    for field in ("min_request_bytes", "max_request_bytes"):
        if value[field] is not None:
            _integer(value[field], f"{label}.{field}")
    if calls == 0:
        require(value["min_request_bytes"] is None and value["max_request_bytes"] is None,
                f"{label}: empty observation has request bounds")
    else:
        require(value["min_request_bytes"] is not None and value["max_request_bytes"] is not None,
                f"{label}: nonempty observation lacks request bounds")
    return {**value, "request_size_counts": counts}


def _provider_parameters(section: dict[str, Any], route: str, label: str) -> None:
    provider = section.get("provider")
    if route in CONTROL_ROUTES:
        _empty(provider, f"{label}.provider")
        return
    require(isinstance(provider, dict), f"{label}.provider: summary missing")
    require(provider.get("provider") == route, f"{label}.provider identity changed")
    require(provider.get("source_scope") == PROVIDER_SOURCE_SCOPES[route],
            f"{label}.provider source scope changed")
    require(isinstance(provider.get("staging_scope"), str),
            f"{label}.provider staging scope missing")
    require(isinstance(provider.get("observation_boundary"), str),
            f"{label}.provider observation boundary missing")
    require(isinstance(provider.get("logical_observation_scope"), str),
            f"{label}.provider logical scope missing")
    require(isinstance(provider.get("underlying_observation_scope"), str),
            f"{label}.provider underlying scope missing")
    params = {
        "file_warm": (None, None, None, None),
        "range_short": (4096, None, None, None),
        "range_zero": (65536, 0, 104_857_600, "minimum-service"),
        "range_delayed": (65536, 1_000_000, 104_857_600, "minimum-service"),
    }[route]
    actual = tuple(provider.get(key) for key in (
        "max_returned_bytes", "fixed_delay_ns", "transfer_bytes_per_second",
        "transfer_delay_policy",
    ))
    require(actual == params, f"{label}.provider parameters changed: {actual}")


def _provider_vectors(section: dict[str, Any], route: str, samples: int, label: str) -> None:
    if route in CONTROL_ROUTES:
        return
    provider = section["provider"]
    # The reviewed coder patch may wrap these vectors under operation_end and
    # after_checks.  Until that patch is present, the current provider summary
    # is accepted only as a preflight shape and is never a capture receipt.
    for name in ("logical", "underlying"):
        values = provider.get(name)
        require(isinstance(values, list) and len(values) == samples,
                f"{label}.provider.{name}: vector cardinality changed")
        for index, value in enumerate(values):
            if value is None:
                require(route == "file_warm" and name == "underlying",
                        f"{label}.provider.{name}[{index}]: unexpected null observation")
                continue
            _validate_observation(value, f"{label}.provider.{name}[{index}]")
    delays = _vector(provider.get("nominal_service_delay_ns"), samples,
                     f"{label}.provider.nominal_service_delay_ns")
    fixed_delay = provider.get("fixed_delay_ns") or 0
    _integer(fixed_delay, f"{label}.provider.fixed_delay_ns")
    logical_values = provider.get("logical")
    require(isinstance(logical_values, list), f"{label}.provider.logical: vector missing")
    for index, (observation, nominal) in enumerate(zip(logical_values, delays)):
        require(observation is not None,
                f"{label}.provider.logical[{index}]: null observation")
        expected_nominal = fixed_delay * observation["delayed_calls"] + observation["transfer_delay_ns"]
        require(nominal == expected_nominal,
                f"{label}.provider.nominal_service_delay_ns[{index}]: expected "
                f"fixed*delayed+transfer={expected_nominal}, got {nominal}")
    if route == "range_short":
        require(delays == [0] * samples, f"{label}: short-range service delay changed")
    if route in ("range_zero", "range_delayed"):
        require(all(delay > 0 for delay in delays),
                f"{label}: transfer-paced range service delay disappeared")


def _source_contract(source: dict[str, Any], section: dict[str, Any], route: str,
                     samples: int, label: str) -> None:
    if route == "generic_read_at":
        for field, expected in GENERIC_COUNTER_CONTRACT.items():
            _exact_vector(source.get(field), expected, samples, f"{label}.source.{field}")
        for field, expected in GENERIC_OPERATION_END_COUNTER_CONTRACT.items():
            _exact_vector(section.get(field), expected, samples,
                          f"{label}.section.{field}")
        require(section.get("source_read_calls") == source.get("read_calls"),
                f"{label}: section/source read-call vectors differ")
        require(section.get("source_read_bytes") == source.get("read_bytes"),
                f"{label}: section/source read-byte vectors differ")
        for field in ("owned_ingress_read_calls", "owned_ingress_read_bytes",
                      "owned_ingress_owned_bytes"):
            _empty(section.get(field), f"{label}.section.{field}")
    else:
        for field in ("read_calls", "read_bytes", "ordinary_payload_read_calls",
                      "ordinary_payload_read_bytes", "max_in_flight_reads",
                      "ordinary_payload_materializations"):
            _empty(source.get(field), f"{label}.source.{field}")
        for field in ("source_read_calls", "source_read_bytes"):
            _empty(section.get(field), f"{label}.section.{field}")
        for field in ("operation_end_source_read_calls", "operation_end_source_read_bytes"):
            _empty(section.get(field), f"{label}.section.{field}")
        _vector(section.get("ordinary_payload_materializations"), samples,
                f"{label}.section.ordinary_payload_materializations")
        require(section["ordinary_payload_materializations"] == [1] * samples,
                f"{label}: owned materialization contract changed")
        if route == "owned_vec":
            _exact_vector(section.get("owned_ingress_read_calls"), 0, samples,
                          f"{label}.section.owned_ingress_read_calls")
            _exact_vector(section.get("owned_ingress_read_bytes"), 0, samples,
                          f"{label}.section.owned_ingress_read_bytes")
        elif route == "owned_reader":
            _exact_vector(section.get("owned_ingress_read_calls"), OWNED_READER_CALLS,
                          samples, f"{label}.section.owned_ingress_read_calls")
            _exact_vector(section.get("owned_ingress_read_bytes"), SOURCE_BYTES,
                          samples, f"{label}.section.owned_ingress_read_bytes")
        else:
            _empty(section.get("owned_ingress_read_calls"),
                   f"{label}.section.owned_ingress_read_calls")
            _empty(section.get("owned_ingress_read_bytes"),
                   f"{label}.section.owned_ingress_read_bytes")
        if route in ("owned_vec", "owned_reader"):
            _exact_vector(section.get("owned_ingress_owned_bytes"), SOURCE_BYTES, samples,
                          f"{label}.section.owned_ingress_owned_bytes")
        else:
            _empty(section.get("owned_ingress_owned_bytes"),
                   f"{label}.section.owned_ingress_owned_bytes")
    for field, expected in GENERIC_CACHE_CONTRACT.items():
        _exact_vector(section.get(field), expected, samples, f"{label}.section.{field}")


def validate_report(report: dict[str, Any], route: str, *, warmups: int = 5,
                    samples: int = 50, label: str = "report",
                    binary: dict[str, Any] | None = None) -> dict[str, Any]:
    """Validate one complete 0925 report and return raw vectors for reduction."""
    require(route in ROUTES, f"{label}: unknown route {route!r}")
    configuration, result = _result(report, label)
    require(configuration.get("samples_per_case") == samples, f"{label}: sample count changed")
    require(configuration.get("warmup_iterations_per_case") == warmups, f"{label}: warmup changed")
    require(configuration.get("cases") == [CASE], f"{label}: selected case changed")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{label}: binary identity missing")
    if binary is not None:
        require(identity.get("binary_sha256") == binary.get("sha256")
                and identity.get("binary_bytes") == binary.get("bytes"),
                f"{label}: binary identity differs from receipt")
    _corpus(result, label)
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict), f"{label}: elapsed statistics missing")
    totals = _vector(elapsed.get("samples"), samples, f"{label}.elapsed_ns.samples", positive=True)
    order = elapsed.get("sample_order")
    require(isinstance(order, list) and sorted(order) == list(range(samples)),
            f"{label}: elapsed sample order is not a permutation")
    section = (result.get("source") or {}).get("docx_section_layout")
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
        "media_bytes": 524_288,
    }
    for field, expected in identity_fields.items():
        require(section.get(field) == expected, f"{label}: section field {field} changed")
    require(section.get("sample_order") == order, f"{label}: section order changed")
    phases = {field: _vector(section.get(field), samples, f"{label}.section.{field}")
              for field in PHASES}
    for index, total in enumerate(totals):
        require(sum(phases[field][index] for field in PHASES) == total,
                f"{label}: six-phase sum mismatch at sample {index}")
    for field in GATES:
        require(section.get(field) is True, f"{label}: gate {field} failed")
    _source_contract(result.get("source") or {}, section, route, samples, label)
    _provider_parameters(section, route, label)
    _provider_vectors(section, route, samples, label)
    sink = result.get("sink")
    require(isinstance(sink, dict), f"{label}: sink summary missing")
    require(sink.get("accepted_bytes") == OUTPUT_BYTES, f"{label}: sink bytes changed")
    writes = _integer(sink.get("write_calls"), f"{label}.sink.write_calls", positive=True)
    largest = _integer(sink.get("largest_write"), f"{label}.sink.largest_write", positive=True)
    require(largest <= 65_536, f"{label}: sink write bound changed")
    buckets = sink.get("write_size_buckets")
    require(isinstance(buckets, dict), f"{label}: sink bucket summary missing")
    bucket_values = [_integer(buckets.get(name), f"{label}.sink.{name}") for name in (
        "bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384",
        "bytes_16385_to_65536", "bytes_over_65536")]
    require(sum(bucket_values) == writes and bucket_values[-1] == 0,
            f"{label}: sink buckets do not conserve writes")
    outputs = section.get("output_sha256")
    require(outputs == [OUTPUT_SHA256] * samples, f"{label}: output digest changed")
    require(result.get("output_sha256") == OUTPUT_SHA256, f"{label}: top-level output changed")
    return {
        "route": route,
        "samples": samples,
        "warmups": warmups,
        "elapsed_ns": totals,
        "phase_ns": phases,
        "sample_order": order,
        "provider": section.get("provider"),
    }


def reduce_reports(reports: dict[str, list[dict[str, Any]]]) -> dict[str, Any]:
    """Return descriptive per-route raw pooled statistics, with no ranking."""
    def observation_reduction(values: list[dict[str, Any]]) -> dict[str, Any]:
        require(values, "provider observation reduction has no observations")
        numeric_fields = (
            "read_calls", "requested_bytes", "returned_bytes", "short_reads",
            "delayed_calls", "transfer_paced_calls", "transfer_delay_ns",
        )
        result: dict[str, Any] = {"samples": len(values)}
        for field in numeric_fields:
            vector = [int(value[field]) for value in values]
            result[field] = {"total": sum(vector), "statistics": _statistics(vector)}
        result["request_size_counts"] = [
            sum(value["request_size_counts"][index] for value in values)
            for index in range(REQUEST_BUCKETS)
        ]
        result["request_size_count_total"] = sum(result["request_size_counts"])
        result["min_request_bytes"] = min(
            int(value["min_request_bytes"]) for value in values
            if value["min_request_bytes"] is not None
        )
        result["max_request_bytes"] = max(
            int(value["max_request_bytes"]) for value in values
            if value["max_request_bytes"] is not None
        )
        require(result["request_size_count_total"] == result["read_calls"]["total"],
                "provider reduction request histogram does not conserve calls")
        return result

    def provider_reduction(rows: list[dict[str, Any]]) -> dict[str, Any] | None:
        providers = [row.get("provider") for row in rows]
        if not any(provider is not None for provider in providers):
            require(all(provider is None for provider in providers),
                    "provider summary is inconsistently present across children")
            return None
        require(all(isinstance(provider, dict) for provider in providers),
                "provider summary is inconsistently present across children")
        first = providers[0]
        require(isinstance(first, dict), "provider summary missing")
        invariant_fields = (
            "provider", "source_scope", "staging_scope", "observation_boundary",
            "logical_observation_scope", "underlying_observation_scope",
            "max_returned_bytes", "fixed_delay_ns", "transfer_bytes_per_second",
            "transfer_delay_policy",
        )
        for provider in providers[1:]:
            require(isinstance(provider, dict), "provider summary missing")
            for field in invariant_fields:
                require(provider.get(field) == first.get(field),
                        f"provider {field} changed across children")
        logical = [observation for provider in providers for observation in provider["logical"]]
        underlying = [
            observation
            for provider in providers
            for observation in provider["underlying"]
            if observation is not None
        ]
        nominal = [
            value for provider in providers for value in provider["nominal_service_delay_ns"]
        ]
        return {
            **{field: first[field] for field in invariant_fields},
            "logical": observation_reduction(logical),
            "underlying": observation_reduction(underlying) if underlying else None,
            "nominal_service_delay_ns": _statistics(nominal),
        }

    reduced: dict[str, Any] = {}
    for route in ROUTES:
        rows = reports.get(route)
        require(isinstance(rows, list) and rows, f"{route}: no validated children")
        elapsed = [value for row in rows for value in row["elapsed_ns"]]
        phases = {
            phase: [value for row in rows for value in row["phase_ns"][phase]]
            for phase in PHASES
        }
        total = sum(elapsed)
        child_statistics = [_statistics(row["elapsed_ns"]) for row in rows]
        reduced[route] = {
            "children": len(rows),
            "samples": len(elapsed),
            "elapsed_ns": _statistics(elapsed),
            "phase_statistics_ns": {phase: _statistics(values)
                                     for phase, values in phases.items()},
            "phase_shares": {phase: sum(values) / total for phase, values in phases.items()},
            "child_latency_ns": [
                {
                    "p50": stats["p50"], "p95": stats["p95"], "p99": stats["p99"],
                    "min": stats["min"], "max": stats["max"],
                    "spread": stats["max"] - stats["min"],
                }
                for stats in child_statistics
            ],
            "child_p50_ns": [stats["p50"] for stats in child_statistics],
            "child_p95_ns": [stats["p95"] for stats in child_statistics],
            "child_p99_ns": [stats["p99"] for stats in child_statistics],
            "child_spread_ns": [stats["max"] - stats["min"] for stats in child_statistics],
            "provider_accounting": provider_reduction(rows),
            "claim": "descriptive provider baseline; no provider ranking or optimization comparison",
        }
    return {"schema": "litchi.performance.0925.descriptive-reduction.v1", "routes": reduced}


def planned_reduction_bytes(*, children: int = 42, samples: int = 50) -> tuple[int, str]:
    """Measure the full planned compact raw reduction before capture."""
    rows = []
    for child in range(children):
        rows.append({
            "child": child,
            "route": ROUTES[child % len(ROUTES)],
            "samples": [{
                "elapsed_ns": 210 + sample,
                "phase_ns": [10 + sample, 20, 30, 40, 50, 60],
                "logical": [0] * REQUEST_BUCKETS,
                "underlying": [0] * REQUEST_BUCKETS,
                "nominal_service_delay_ns": 0,
            } for sample in range(samples)],
            "child_latency_ns": {"p50": 250, "p95": 299, "p99": 299, "spread": 49},
            "provider_accounting": {
                "logical_read_calls": samples,
                "logical_requested_bytes": samples * 64,
                "logical_returned_bytes": samples * 64,
                "underlying_read_calls": samples,
                "underlying_requested_bytes": samples * 64,
                "underlying_returned_bytes": samples * 64,
                "request_size_counts": [samples] + [0] * (REQUEST_BUCKETS - 1),
            },
        })
    data = json.dumps({"schema": "litchi.performance.0925.compact-reduction.v1", "children": rows},
                      separators=(",", ":"), sort_keys=True).encode()
    return len(data), sha256_bytes(data)


def schedule_rows() -> list[dict[str, Any]]:
    """Return the fixed six-block route schedule used by the capture runner."""
    rows: list[dict[str, Any]] = []
    for block in range(1, 7):
        order = list(ROUTES if block % 2 else tuple(reversed(ROUTES)))
        direction = "forward" if block % 2 else "reverse"
        for position, route in enumerate(order, start=1):
            rows.append({
                "id": f"provider-block-{block:02d}-{position:02d}-{route}",
                "block": block,
                "order": direction,
                "position": position,
                "route": route,
                "cpu": CPU,
                "warmups": 5,
                "samples": 50,
                "failure_policy": "retain_child_and_invalidate_cohort; never_rerun_or_recapture",
            })
    require(len(rows) == 42, "provider schedule cardinality changed")
    require({row["route"] for row in rows} == set(ROUTES), "provider schedule route set changed")
    require(all(sum(row["route"] == route for row in rows) == 6 for route in ROUTES),
            "provider schedule route counts changed")
    return rows


def load_and_validate(path: Path, route: str, *, binary: dict[str, Any] | None = None,
                      warmups: int = 5, samples: int = 50) -> dict[str, Any]:
    return validate_report(read_json(path), route, warmups=warmups, samples=samples,
                           label=str(path), binary=binary)


def _synthetic_observation(calls: int = 1, *, delayed_calls: int = 0,
                           transfer_delay_ns: int = 0) -> dict[str, Any]:
    histogram = [0] * REQUEST_BUCKETS
    histogram[0] = calls
    return {
        "read_calls": calls,
        "requested_bytes": 64 * calls,
        "returned_bytes": 64 * calls,
        "min_request_bytes": 64 if calls else None,
        "max_request_bytes": 64 if calls else None,
        "short_reads": 0,
        "delayed_calls": delayed_calls,
        "transfer_paced_calls": 0,
        "transfer_delay_ns": transfer_delay_ns,
        "request_size_counts": histogram,
    }


def synthetic_report(route: str, *, warmups: int = 5, samples: int = 50,
                     binary: dict[str, Any] | None = None) -> dict[str, Any]:
    """Build a deterministic in-memory report for validator preflight tests."""
    require(route in ROUTES, f"unknown synthetic route: {route}")
    if binary is None:
        binary = {"path": "/tmp/0925-synthetic-binary", "sha256": "a" * 64,
                  "bytes": 123_456, "mode_bits": 0o755, "executable": True,
                  "profile": "release"}
    elapsed: list[int] = []
    phase_vectors = {phase: [] for phase in PHASES}
    for index in range(samples):
        values = (10 + index, 20, 30, 40, 50, 60)
        for phase, value in zip(PHASES, values):
            phase_vectors[phase].append(value)
        elapsed.append(sum(values))
    elapsed_statistics = {"unit": "ns", "samples": elapsed, "sample_order": list(range(samples)),
                          **_statistics(elapsed), "standard_deviation": statistics.pstdev(elapsed),
                          "confidence_interval_95": None}
    corpus = {
        "name": "docx-source-backed-existing-section-layout",
        "generator": "litchi-docx-section-layout-source-v1",
        "package_format": "DOCX/OPC/ZIP",
        "shape": "existing-section-layout",
        "payload_kind": "deterministic-incompressible-media",
        "compression": "deflate",
        "entry_count": 7,
        "archive_member_count": 10,
        "entry_bytes": 524_288,
        "uncompressed_payload_bytes": 2_116_441,
        "archive_bytes": SOURCE_BYTES,
        "archive_sha256": SOURCE_SHA256,
        "target_entry": "word/document.xml",
        "target_payload_bytes": 18_855,
        "target_payload_sha256": "67cb9066f112bfee6905a23f3eabdc76b65c86cd0c6613dd60a5a18895ceac81",
        "xlsx": None,
    }
    section: dict[str, Any] = {
        "implementation": "litchi-docx-source-backed-section-layout",
        "timing_scope": "preparation, source-backed open, section query, columns edit, commit, and sequential publication; correctness, reopen, patch, refusal, cache, source, sink, and ZIP gates are excluded",
        "performance_claim": "none",
        "ingress_route": route,
        "source_observation_scope": ROUTE_SCOPES[route],
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
        "media_bytes": 524_288,
        "sample_order": list(range(samples)),
        **phase_vectors,
        "source_read_calls": [202] * samples if route == "generic_read_at" else [],
        "source_read_bytes": [8_409_488] * samples if route == "generic_read_at" else [],
        "operation_end_source_read_calls": (
            [169] * samples if route == "generic_read_at" else []
        ),
        "operation_end_source_read_bytes": (
            [6_308_061] * samples if route == "generic_read_at" else []
        ),
        "owned_ingress_read_calls": (
            [0] * samples if route == "owned_vec" else
            [OWNED_READER_CALLS] * samples if route == "owned_reader" else []
        ),
        "owned_ingress_read_bytes": (
            [0] * samples if route == "owned_vec" else
            [SOURCE_BYTES] * samples if route == "owned_reader" else []
        ),
        "owned_ingress_owned_bytes": (
            [SOURCE_BYTES] * samples
            if route in ("owned_vec", "owned_reader") else []
        ),
        "ordinary_payload_materializations": [1] * samples,
        "source_cache_before_publication_after_commit_hits": [0] * samples,
        "source_cache_before_publication_after_commit_cold_loads": [1] * samples,
        "source_cache_before_publication_after_commit_successful_loads": [1] * samples,
        "source_cache_before_publication_after_commit_retained_entries": [1] * samples,
        "source_cache_before_publication_after_commit_retained_bytes": [18_855] * samples,
        "output_sha256": [OUTPUT_SHA256] * samples,
        **{gate: True for gate in GATES},
        "provider": None,
    }
    source: dict[str, Any] = {
        "read_calls": [202] * samples if route == "generic_read_at" else [],
        "read_bytes": [8_409_488] * samples if route == "generic_read_at" else [],
        "ordinary_payload_read_calls": [138] * samples if route == "generic_read_at" else [],
        "ordinary_payload_read_bytes": [8_397_632] * samples if route == "generic_read_at" else [],
        "max_in_flight_reads": [1] * samples if route == "generic_read_at" else [],
        "ordinary_payload_materializations": None if route != "generic_read_at" else [1] * samples,
        "docx_section_layout": section,
    }
    if route in PROVIDER_ROUTES:
        params = {
            "file_warm": (None, None, None, None),
            "range_short": (4096, None, None, None),
            "range_zero": (65536, 0, 104_857_600, "minimum-service"),
            "range_delayed": (65536, 1_000_000, 104_857_600, "minimum-service"),
        }[route]
        section["provider"] = {
            "provider": route,
            "source_scope": PROVIDER_SOURCE_SCOPES[route],
            "staging_scope": "synthetic preflight",
            "observation_boundary": "operation end before checks",
            "logical_observation_scope": "logical wrapper",
            "underlying_observation_scope": "delegated wrapper",
            "max_returned_bytes": params[0],
            "fixed_delay_ns": params[1],
            "transfer_bytes_per_second": params[2],
            "transfer_delay_policy": params[3],
            "logical": [
                _synthetic_observation(
                    delayed_calls=1 if route == "range_delayed" else 0,
                    transfer_delay_ns=(1000 if route in ("range_zero", "range_delayed") else 0),
                )
                for _ in range(samples)
            ],
            "underlying": (
                [None] * samples if route == "file_warm"
                else [_synthetic_observation() for _ in range(samples)]
            ),
            "nominal_service_delay_ns": (
                [1_001_000] * samples if route == "range_delayed" else
                [1000] * samples if route == "range_zero" else
                [0] * samples
            ),
        }
    sink = {
        "accepted_bytes": OUTPUT_BYTES,
        "write_calls": 1,
        "largest_write": OUTPUT_BYTES if OUTPUT_BYTES <= 65_536 else 65_536,
        "write_size_buckets": {
            "bytes_0": 0, "bytes_1_to_512": 0, "bytes_513_to_4096": 0,
            "bytes_4097_to_16384": 0, "bytes_16385_to_65536": 1,
            "bytes_over_65536": 0,
        },
    }
    return {
        "schema_version": 1,
        "tool": {"name": "litchi-perf-baseline", "version": "synthetic",
                  "binary": "litchi-perf-baseline", "profile": "release",
                  "target_os": "linux", "target_arch": "x86_64",
                  "instrumentation": "normal"},
        "binary_identity": {
            "path": binary.get("path"),
            "binary_sha256": binary.get("sha256"),
            "binary_bytes": binary.get("bytes"),
            "mode_bits": binary.get("mode_bits"),
            "executable": binary.get("executable", True),
            "profile": binary.get("profile", "release"),
        },
        "environment": {"rustc_version": "synthetic", "git_revision": "aa077c030819c3019288e350f34959666ab7204f"},
        "configuration": {
            "samples_per_case": samples, "warmup_iterations_per_case": warmups,
            "cases": [CASE], "filesystem_cache_states": [],
            "filesystem_fresh_child_per_sample": True, "filesystem_process_isolated": True,
            "filesystem_root_selected": False, "range_simulation": {},
            "execution_workers": [1], "opc_cache_lock_diagnostics": False,
        },
        "parallel_metrics": {},
        "results": [{"case": CASE, "corpus": corpus, "elapsed_ns": elapsed_statistics,
                      "sink": sink, "source": source, "output_sha256": OUTPUT_SHA256,
                      "operation_metrics": None}],
    }
