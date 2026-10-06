#!/usr/bin/env python3
"""Independently reconcile the retained 0925 provider report cohort.

The command only reads the committed raw reports and their terminal inventory.
It does not import the 0925 validator, run a child, read timing data, or
authorize a capture.  It checks report hashes, source/binary identities,
per-sample provider equality, histogram conservation, and aggregate counters.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


MAX_ARTIFACT_BYTES = 1 << 20
SAMPLES = 50
WARMUPS = 5
ROUTES = (
    "generic_read_at",
    "owned_vec",
    "owned_reader",
    "file_warm",
    "range_short",
    "range_zero",
    "range_delayed",
)
PROVIDER_ROUTES = ("file_warm", "range_short", "range_zero", "range_delayed")
CASE = "docx_source_backed_existing_section_layout_edit_save"
SOURCE_BYTES = 2_101_427
SOURCE_SHA256 = "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8"
OUTPUT_SHA256 = "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
REQUEST_BUCKETS = 18
OBSERVATION_FIELDS = (
    "read_calls", "requested_bytes", "returned_bytes", "short_reads",
    "delayed_calls", "transfer_paced_calls", "transfer_delay_ns",
)
BUCKET_DEFINITIONS = [
    {"index": 0, "lower_inclusive": 0, "upper_inclusive": 1},
    {"index": 1, "lower_inclusive": 2, "upper_inclusive": 2},
    *[
        {"index": index, "lower_inclusive": (1 << (index - 1)) + 1,
         "upper_inclusive": 1 << index}
        for index in range(2, 17)
    ],
    {"index": 17, "lower_inclusive": 65_537, "upper_inclusive": None},
]


class ReconciliationError(RuntimeError):
    """A retained report or reconciliation contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReconciliationError(message)


def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, root: Path) -> dict[str, Any]:
    path = path.resolve()
    root = root.resolve()
    try:
        relative = path.relative_to(root)
    except ValueError as error:
        raise ReconciliationError(f"file is outside packet: {path}") from error
    return {
        "path": str(relative),
        "bytes": path.stat().st_size,
        "sha256": sha256_path(path),
    }


def read_json(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise ReconciliationError(f"cannot read {path}: {error}") from error
    require(isinstance(value, dict), f"expected JSON object: {path}")
    return value


def expected_schedule() -> list[dict[str, Any]]:
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
                "cpu": 12,
                "warmups": WARMUPS,
                "samples": SAMPLES,
                "failure_policy": "retain_child_and_invalidate_cohort; never_rerun_or_recapture",
            })
    return rows


def observation(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: observation is missing")
    for field in OBSERVATION_FIELDS:
        require(type(value.get(field)) is int and value[field] >= 0,
                f"{label}: {field} is not a nonnegative integer")
    require(value["returned_bytes"] <= value["requested_bytes"],
            f"{label}: returned bytes exceed requested bytes")
    require(value["short_reads"] <= value["read_calls"]
            and value["delayed_calls"] <= value["read_calls"]
            and value["transfer_paced_calls"] <= value["read_calls"],
            f"{label}: event counter exceeds calls")
    histogram = value.get("request_size_counts")
    require(isinstance(histogram, list) and len(histogram) == REQUEST_BUCKETS,
            f"{label}: request histogram cardinality changed")
    require(all(type(item) is int and item >= 0 for item in histogram),
            f"{label}: request histogram contains an invalid count")
    require(sum(histogram) == value["read_calls"],
            f"{label}: request histogram does not conserve calls")
    calls = value["read_calls"]
    if calls == 0:
        require(value.get("min_request_bytes") is None
                and value.get("max_request_bytes") is None,
                f"{label}: empty observation has request bounds")
    else:
        require(type(value.get("min_request_bytes")) is int
                and type(value.get("max_request_bytes")) is int
                and value["min_request_bytes"] >= 0
                and value["min_request_bytes"] <= value["max_request_bytes"],
                f"{label}: request bounds are invalid")
    return value


def aggregate_observations(values: list[dict[str, Any]], label: str) -> dict[str, Any]:
    require(values, f"{label}: no observations")
    for index, value in enumerate(values):
        observation(value, f"{label}[{index}]")
    totals = {field: sum(value[field] for value in values) for field in OBSERVATION_FIELDS}
    histogram = [sum(value["request_size_counts"][index] for value in values)
                 for index in range(REQUEST_BUCKETS)]
    require(sum(histogram) == totals["read_calls"],
            f"{label}: aggregate histogram does not conserve calls")
    bounds = [value["min_request_bytes"] for value in values
              if value["min_request_bytes"] is not None]
    upper_bounds = [value["max_request_bytes"] for value in values
                    if value["max_request_bytes"] is not None]
    return {
        "sample_observations": len(values),
        "totals": totals,
        "request_size_counts": histogram,
        "request_size_count_total": sum(histogram),
        "min_request_bytes": min(bounds) if bounds else None,
        "max_request_bytes": max(upper_bounds) if upper_bounds else None,
    }


def provider_parameters(provider: dict[str, Any], route: str) -> dict[str, Any]:
    expected = {
        "file_warm": {
            "max_returned_bytes": None, "fixed_delay_ns": None,
            "transfer_bytes_per_second": None, "transfer_delay_policy": None,
        },
        "range_short": {
            "max_returned_bytes": 4096, "fixed_delay_ns": None,
            "transfer_bytes_per_second": None, "transfer_delay_policy": None,
        },
        "range_zero": {
            "max_returned_bytes": 65536, "fixed_delay_ns": 0,
            "transfer_bytes_per_second": 104_857_600,
            "transfer_delay_policy": "minimum-service",
        },
        "range_delayed": {
            "max_returned_bytes": 65536, "fixed_delay_ns": 1_000_000,
            "transfer_bytes_per_second": 104_857_600,
            "transfer_delay_policy": "minimum-service",
        },
    }[route]
    actual = {key: provider.get(key) for key in expected}
    require(actual == expected, f"{route}: provider parameters changed: {actual}")
    return actual


def reconcile(packet: Path) -> dict[str, Any]:
    packet = packet.resolve()
    require(packet.is_dir(), f"packet is not a directory: {packet}")
    manifest_path = packet / "manifest.json"
    failure_path = packet / "capture-failure.json"
    manifest = read_json(manifest_path)
    failure = read_json(failure_path)
    require(manifest.get("schema") == "litchi.performance.0925.provider-capture-manifest.v1",
            "manifest schema changed")
    manifest_binary = manifest.get("binary")
    require(isinstance(manifest_binary, dict)
            and type(manifest_binary.get("bytes")) is int
            and isinstance(manifest_binary.get("sha256"), str),
            "manifest binary binding is missing")
    require(failure.get("schema") == "litchi.performance.0925.capture-terminal-failure.v1",
            "failure receipt schema changed")
    require(failure.get("status") == "invalid"
            and failure.get("capture_state_present") is False,
            "retained terminal invalidity changed")
    rows = manifest.get("schedule", {}).get("rows")
    require(rows == expected_schedule(), "committed 0925 schedule changed")
    require(failure.get("reports_present") == len(rows) == 42,
            "retained report count changed")
    inventory = failure.get("reports")
    require(isinstance(inventory, list) and len(inventory) == 42,
            "failure report inventory is incomplete")
    inventory_by_path = {row.get("path"): row for row in inventory}
    require(len(inventory_by_path) == 42, "failure report inventory has duplicate paths")

    report_records: list[dict[str, Any]] = []
    reports_by_route: dict[str, list[dict[str, Any]]] = {route: [] for route in ROUTES}
    identity: dict[str, Any] | None = None
    binary_identity: dict[str, Any] | None = None
    for row in rows:
        child_id = row["id"]
        report_path = packet / "children" / child_id / "report.json"
        relative = f"children/{child_id}/report.json"
        expected_descriptor = inventory_by_path.get(relative)
        require(isinstance(expected_descriptor, dict), f"{relative}: missing failure binding")
        actual_descriptor = descriptor(report_path, packet)
        require(actual_descriptor["bytes"] == expected_descriptor.get("bytes")
                and actual_descriptor["sha256"] == expected_descriptor.get("sha256"),
                f"{relative}: report hash or size differs from failure receipt")
        report = read_json(report_path)
        require(report.get("schema_version") == 1, f"{relative}: schema changed")
        config = report.get("configuration")
        require(isinstance(config, dict)
                and config.get("samples_per_case") == SAMPLES
                and config.get("warmup_iterations_per_case") == WARMUPS
                and config.get("cases") == [CASE],
                f"{relative}: sample/warmup/case contract changed")
        results = report.get("results")
        require(isinstance(results, list) and len(results) == 1,
                f"{relative}: result cardinality changed")
        result = results[0]
        require(isinstance(result, dict) and result.get("case") == CASE,
                f"{relative}: case changed")
        section = (result.get("source") or {}).get("docx_section_layout")
        corpus = result.get("corpus")
        require(isinstance(section, dict) and isinstance(corpus, dict),
                f"{relative}: source/corpus summary missing")
        require(section.get("ingress_route") == row["route"],
                f"{relative}: route identity changed")
        current_identity = {
            "source_bytes": section.get("source_bytes"),
            "source_sha256": section.get("source_sha256"),
            "corpus_archive_bytes": corpus.get("archive_bytes"),
            "corpus_archive_sha256": corpus.get("archive_sha256"),
            "corpus_entry_bytes": corpus.get("entry_bytes"),
            "corpus_uncompressed_payload_bytes": corpus.get("uncompressed_payload_bytes"),
            "target_payload_bytes": corpus.get("target_payload_bytes"),
            "target_payload_sha256": corpus.get("target_payload_sha256"),
            "expected_output_sha256": section.get("expected_output_sha256"),
            "output_sha256": result.get("output_sha256"),
        }
        require(current_identity["source_bytes"] == SOURCE_BYTES
                and current_identity["source_sha256"] == SOURCE_SHA256
                and current_identity["corpus_archive_bytes"] == SOURCE_BYTES
                and current_identity["corpus_archive_sha256"] == SOURCE_SHA256
                and current_identity["expected_output_sha256"] == OUTPUT_SHA256
                and current_identity["output_sha256"] == OUTPUT_SHA256,
                f"{relative}: source or output identity changed")
        if identity is None:
            identity = current_identity
        else:
            require(current_identity == identity,
                    f"{relative}: source/corpus/output identity differs across reports")
        current_binary = report.get("binary_identity")
        require(isinstance(current_binary, dict), f"{relative}: binary identity missing")
        normalized_binary = {
            "binary_sha256": current_binary.get("binary_sha256"),
            "binary_bytes": current_binary.get("binary_bytes"),
            "profile": current_binary.get("profile"),
        }
        require(normalized_binary["binary_sha256"] == manifest_binary["sha256"]
                and normalized_binary["binary_bytes"] == manifest_binary["bytes"],
                f"{relative}: binary identity differs from manifest binding")
        if binary_identity is None:
            binary_identity = normalized_binary
        else:
            require(normalized_binary == binary_identity,
                    f"{relative}: binary identity differs across reports")
        provider = section.get("provider")
        if row["route"] in PROVIDER_ROUTES:
            require(isinstance(provider, dict), f"{relative}: provider summary missing")
            logical = provider.get("logical")
            underlying = provider.get("underlying")
            nominal = provider.get("nominal_service_delay_ns")
            require(isinstance(logical, list) and len(logical) == SAMPLES,
                    f"{relative}: logical sample cardinality changed")
            require(isinstance(underlying, list) and len(underlying) == SAMPLES,
                    f"{relative}: underlying sample cardinality changed")
            require(isinstance(nominal, list) and len(nominal) == SAMPLES,
                    f"{relative}: nominal sample cardinality changed")
            provider_parameters(provider, row["route"])
            fixed_delay = provider.get("fixed_delay_ns") or 0
            for sample_index, (logical_observation, nominal_delay) in enumerate(
                    zip(logical, nominal)):
                observation(logical_observation, f"{relative}: logical[{sample_index}]")
                require(type(nominal_delay) is int and nominal_delay >= 0,
                        f"{relative}: nominal delay is invalid at sample {sample_index}")
                expected_nominal = (
                    fixed_delay * logical_observation["delayed_calls"]
                    + logical_observation["transfer_delay_ns"]
                )
                require(nominal_delay == expected_nominal,
                        f"{relative}: nominal service formula differs at sample {sample_index}")
            base = {key: value for key, value in provider.items()
                    if key not in ("logical", "underlying", "nominal_service_delay_ns")}
            record = {
                "id": child_id,
                "report": actual_descriptor,
                "provider_base": base,
                "logical": logical,
                "underlying": underlying,
                "nominal_service_delay_ns": nominal,
            }
            reports_by_route[row["route"]].append(record)
        else:
            require(provider is None, f"{relative}: control route has provider counters")
        report_records.append({
            "id": child_id,
            "route": row["route"],
            "path": relative,
            "bytes": actual_descriptor["bytes"],
            "sha256": actual_descriptor["sha256"],
        })

    require(identity is not None and binary_identity is not None, "report identity inventory is empty")
    route_summaries: dict[str, Any] = {}
    for route in PROVIDER_ROUTES:
        records = reports_by_route[route]
        require(len(records) == 6, f"{route}: expected six retained reports")
        first = records[0]
        for record in records[1:]:
            require(record["provider_base"] == first["provider_base"],
                    f"{route}: provider parameters/scope differ across reports")
        logical_values = [value for record in records for value in record["logical"]]
        underlying_values = [value for record in records for value in record["underlying"]]
        nominal_values = [value for record in records for value in record["nominal_service_delay_ns"]]
        logical_equal = all(record["logical"] == first["logical"] for record in records[1:])
        underlying_equal = all(record["underlying"] == first["underlying"] for record in records[1:])
        nominal_equal = all(record["nominal_service_delay_ns"]
                            == first["nominal_service_delay_ns"] for record in records[1:])
        require(logical_equal and underlying_equal and nominal_equal,
                f"{route}: per-sample counters or histograms differ")
        logical_aggregate = aggregate_observations(logical_values, f"{route}.logical")
        if all(value is None for value in underlying_values):
            underlying_aggregate: dict[str, Any] | None = None
            underlying_nulls = len(underlying_values)
        else:
            require(all(isinstance(value, dict) for value in underlying_values),
                    f"{route}: underlying observations mix null and object values")
            underlying_aggregate = aggregate_observations(
                [value for value in underlying_values if isinstance(value, dict)],
                f"{route}.underlying"
            )
            underlying_nulls = 0
        route_summaries[route] = {
            "children": len(records),
            "samples": len(logical_values),
            "provider": first["provider_base"],
            "per_sample_equality": {
                "logical": logical_equal,
                "underlying": underlying_equal,
                "nominal_service_delay_ns": nominal_equal,
                "comparisons": (len(records) - 1) * SAMPLES,
            },
            "logical": logical_aggregate,
            "underlying": underlying_aggregate,
            "underlying_null_samples": underlying_nulls,
            "nominal_service_delay_ns": {
                "total": sum(nominal_values),
                "min": min(nominal_values),
                "max": max(nominal_values),
                "samples": len(nominal_values),
            },
        }
    return {
        "schema": "litchi.performance.0926.provider-reconciliation.v1",
        "status": "passed",
        "claim_scope": "offline raw-report reconciliation only; no timing, capture, provider ranking, or optimization claim",
        "performance_children_run": 0,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "input": {
            "packet": str(packet),
            "manifest": descriptor(packet / "manifest.json", packet),
            "failure_receipt": descriptor(packet / "capture-failure.json", packet),
            "reports": len(report_records),
            "report_bytes": sum(record["bytes"] for record in report_records),
            "report_bindings": report_records,
        },
        "source_identity": identity,
        "binary_identity": binary_identity,
        "request_size_bucket_scope": (
            "fixed log2 bucket counts from the retained provider counters; this is not an exact per-request size trace"
        ),
        "request_size_buckets": BUCKET_DEFINITIONS,
        "provider_routes": route_summaries,
        "checks": {
            "report_sha256_and_size_bound_to_failure_inventory": True,
            "source_lengths_and_hashes_equal_across_42_reports": True,
            "binary_identities_equal_across_42_reports": True,
            "per_sample_provider_counters_and_histograms_equal": True,
            "nominal_service_formula_conserved": True,
            "aggregate_histograms_conserve_read_calls": True,
        },
    }


def main() -> int:
    default_packet = Path(__file__).resolve().parent.parent / "change-0925" / "capture-v4"
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=default_packet)
    parser.add_argument("--output", type=Path,
                        default=Path(__file__).resolve().parent / "provider-reconciliation-final.json")
    args = parser.parse_args()
    try:
        result = reconcile(args.packet)
        data = (json.dumps(result, indent=2, sort_keys=True) + "\n").encode()
        require(len(data) < MAX_ARTIFACT_BYTES, "reconciliation receipt is at least 1 MiB")
        output = args.output.resolve()
        require(not output.exists(), f"refusing to replace retained receipt: {output}")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(data)
        print(json.dumps({
            "status": result["status"],
            "output": str(output),
            "bytes": len(data),
            "reports": result["input"]["reports"],
            "raw_report_bytes": result["input"]["report_bytes"],
            "performance_children_run": result["performance_children_run"],
        }, sort_keys=True))
    except Exception as error:
        print(f"0926 reconciliation failed: {type(error).__name__}: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
