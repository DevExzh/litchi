#!/usr/bin/env python3
"""Validate one DrawingML paired raw-probe output directory.

This validator is intentionally stricter than a JSON-shape check.  It binds
every row to the lane, source role, source revision, fixture, and allocator
mode, checks the fixed ABBA order, and refuses to turn an unavailable adapter
into a performance claim.  Semantic fidelity and replay checks belong in the
lane adapter; when that adapter is absent, its rows must say so explicitly.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any, Mapping


SCHEMA = "litchi.drawingml.raw.probe.v1"
SUMMARY_SCHEMA = "litchi.drawingml.raw.validation.v1"
LANES = (
    "drawingml-chart-raw",
    "drawingml-diagram-raw",
    "drawingml-text-media-raw",
    "drawingml-theme-ink",
    "drawingml-geometry-color-formatcode2-transform",
)
RAW_OPERATIONS = (
    "entry_read",
    "read_shared_when_supported",
    "vec_write",
    "write_to_bounded_sink",
    "preflight_refusal",
    "malformed_refusal",
)
OWNER_OPERATIONS = (
    "noop_commit",
    "changed_commit",
    "forward_replay",
    "inverse_replay",
    "stale_replay",
)
ABBA = ("baseline", "candidate", "candidate", "baseline")
MINIMUM_MS = 150
TARGET_MS = 500
WARMUPS = 3
SAMPLES = 20
BLOCKS = 5
PRODUCTION_SINK_PARTS = {
    "chart",
    "formatcode2",
    "svg-blip",
    "model3d",
    "theme-family",
    "ink-actions",
    "transform",
}
STRUCTURAL_BOUNDARY_PARTS = {"geometry"}


class ValidationError(ValueError):
    """A fail-closed probe output error."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise ValidationError(message)


def _sha256(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink() and path.is_file(), f"not a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _load(path: Path) -> Mapping[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read probe output {path}: {error}") from error
    _expect(isinstance(value, dict), f"probe output root is not an object: {path}")
    return value


def _ordered_outputs(output_dir: Path, prefix: str) -> list[Path]:
    values = sorted(
        output_dir.glob(f"{prefix}-*.json"),
        key=lambda path: int(re.search(r"-(\d+)\.json$", path.name).group(1))
        if re.search(r"-(\d+)\.json$", path.name)
        else -1,
    )
    _expect(len(values) == 4, f"expected four fresh {prefix} reports, found {len(values)}")
    _expect(
        [int(re.search(r"-(\d+)\.json$", path.name).group(1)) for path in values]
        == [1, 2, 3, 4],
        f"{prefix} report indices are not the four ABBA legs",
    )
    return values


def _is_hash(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(
        character in "0123456789abcdef" for character in value
    )


def _fixture_lane_identity(lane: str, fixture: Mapping[str, Any]) -> str:
    lanes = fixture.get("lanes")
    _expect(isinstance(lanes, dict), "fixture lanes are missing")
    value = lanes.get(lane)
    _expect(isinstance(value, dict), f"fixture lane is missing: {lane}")
    for key in ("root", "mutation", "malformed"):
        _expect(isinstance(value.get(key), str), f"fixture {lane} field is missing: {key}")
    return hashlib.sha256(
        f"drawingml-fixture-v1\0{lane}\0{value['root']}\0{value['mutation']}\0{value['malformed']}".encode()
    ).hexdigest()


def _fixture_part_identity(
    lane_identity: str,
    part_id: str,
    source_sha256: str,
    changed_sha256: str,
    malformed_sha256: str,
) -> str:
    return hashlib.sha256(
        f"drawingml-fixture-part-v1\0{lane_identity}\0{part_id}\0{source_sha256}\0{changed_sha256}\0{malformed_sha256}".encode()
    ).hexdigest()


def _owner_fixture_identity(
    lane_identity: str,
    route: str,
    source_sha256: str,
    stale_sha256: str,
) -> str:
    return hashlib.sha256(
        f"drawingml-owner-fixture-v1\0{lane_identity}\0{route}\0{source_sha256}\0{stale_sha256}".encode()
    ).hexdigest()


def _refusal(value: Any, kind: str, path: Path, label: str) -> None:
    _expect(isinstance(value, dict), f"{path}: {label} refusal is missing")
    _expect(value.get("kind") == kind, f"{path}: {label} refusal kind differs")
    _expect(
        isinstance(value.get("observed"), int) and value["observed"] > 0,
        f"{path}: {label} observed bytes are missing",
    )
    if kind == "limit":
        _expect(
            isinstance(value.get("resource"), str)
            and value["resource"]
            and isinstance(value.get("limit_kind"), str)
            and value["limit_kind"]
            and isinstance(value.get("limit"), int)
            and value["limit"] >= 0
            and value["observed"] > value["limit"],
            f"{path}: {label} limit receipt is incomplete",
        )
    elif kind == "patch_conflict":
        _expect(
            value.get("source_kind") in {"patch_conflict", "invalid_input"},
            f"{path}: {label} patch conflict source kind is missing",
        )
        _expect(value.get("resource") is None, f"{path}: {label} has a resource")
    else:
        _expect(value.get("resource") is None, f"{path}: {label} has a resource")


def _validate_rows(
    report: Mapping[str, Any],
    path: Path,
    expected_fixture_lane_identity: str,
) -> str:
    rows = report.get("rows")
    _expect(isinstance(rows, list), f"{path}: rows are missing")
    _expect(all(isinstance(row, dict) for row in rows), f"{path}: rows contain a non-object")
    status = report.get("status")
    _expect(status in {"unavailable", "qualified"}, f"{path}: unsupported status {status!r}")
    if status == "unavailable":
        _expect(
            report.get("qualification_claim") is False,
            f"{path}: unavailable report claims qualification",
        )
        _expect(
            isinstance(report.get("availability_reason"), str)
            and report["availability_reason"],
            f"{path}: unavailable reason is missing",
        )
        _expect(
            all(row.get("status") == "unavailable" for row in rows),
            f"{path}: unavailable report has a non-unavailable row",
        )
        _expect(
            all(row.get("elapsed_ns") is None for row in rows),
            f"{path}: unavailable report carries timing",
        )
        return status

    _expect(report.get("qualification_claim") is True, f"{path}: qualified report has no claim marker")
    endpoint = report.get("endpoint", "raw")
    _expect(endpoint in {"raw", "owner"}, f"{path}: unsupported endpoint {endpoint!r}")
    if endpoint == "owner":
        _expect(
            isinstance(report.get("owner_route"), str) and report["owner_route"],
            f"{path}: owner endpoint route is missing",
        )
        _expect(report.get("raw_scope") is None, f"{path}: owner endpoint carries a raw scope")
    else:
        _expect(report.get("owner_route") is None, f"{path}: raw endpoint carries an owner route")
        _expect(
            report.get("raw_scope") in {None, "bulk", "family"},
            f"{path}: unsupported raw scope",
        )
    early_protocol = report.get("protocol")
    _expect(isinstance(early_protocol, dict), f"{path}: protocol receipt is missing")
    _expect(
        early_protocol.get("minimum_ms") == MINIMUM_MS,
        f"{path}: minimum duration was weakened",
    )
    for row in rows:
        _expect(
            isinstance(row.get("elapsed_ns"), int)
            and row["elapsed_ns"] >= MINIMUM_MS * 1_000_000,
            f"{path}: accepted leg is shorter than {MINIMUM_MS} ms",
        )
    _expect(
        report.get("fixture_lane_identity_sha256") == expected_fixture_lane_identity,
        f"{path}: fixture lane identity differs",
    )
    capabilities = report.get("capabilities")
    _expect(
        isinstance(capabilities, list) and capabilities,
        f"{path}: capability matrix is missing",
    )
    raw_routes: set[str] = set()
    raw_capabilities: dict[str, Mapping[str, Any]] = {}
    shared_reader_supported = False
    owner_operations: dict[str, list[str]] = {operation: [] for operation in OWNER_OPERATIONS}
    owner_route_names: set[str] = set()
    for capability in capabilities:
        _expect(isinstance(capability, dict), f"{path}: capability row is not an object")
        route = capability.get("route")
        kind = capability.get("kind")
        status_value = capability.get("status")
        evidence = capability.get("evidence")
        _expect(
            isinstance(route, str) and route,
            f"{path}: capability route is missing",
        )
        _expect(
            isinstance(kind, str) and kind,
            f"{path}: capability owner identity is missing",
        )
        _expect(
            isinstance(evidence, str) and evidence,
            f"{path}: capability source evidence is missing",
        )
        _expect(
            isinstance(capability.get("read_shared"), bool),
            f"{path}: capability shared-reader receipt is missing",
        )
        _expect(
            isinstance(capability.get("production_sink"), bool),
            f"{path}: capability production-sink receipt is missing",
        )
        _expect(
            capability.get("production_boundary") in {"none", "input_limit", "structural"},
            f"{path}: capability production-boundary receipt is invalid",
        )
        if kind == "raw":
            _expect(endpoint == "raw", f"{path}: raw capability leaked into owner endpoint")
            _expect(status_value == "qualified", f"{path}: raw capability is not qualified")
            _expect(route not in raw_routes, f"{path}: duplicate raw capability {route}")
            raw_routes.add(route)
            raw_capabilities[route] = capability
            expected_sink = route in PRODUCTION_SINK_PARTS
            expected_boundary = (
                "structural"
                if route in STRUCTURAL_BOUNDARY_PARTS
                else "none"
                if route in {"image-blip"}
                else "input_limit"
            )
            _expect(
                capability["production_sink"] is expected_sink,
                f"{path}: {route} production-sink capability is not source-backed",
            )
            _expect(
                capability["production_boundary"] == expected_boundary,
                f"{path}: {route} production boundary is not source-backed",
            )
            shared_reader_supported |= capability["read_shared"]
        else:
            _expect(endpoint == "owner", f"{path}: owner capability leaked into raw endpoint")
            _expect(
                capability["production_sink"] is False
                and capability["production_boundary"] == "none",
                f"{path}: owner capability carries raw boundary claims",
            )
            _expect(
                status_value in {"qualified", "not_applicable"},
                f"{path}: unsupported owner capability status",
            )
            _expect(
                route in OWNER_OPERATIONS,
                f"{path}: owner capability names an unknown operation",
            )
            owner_operations[route].append(status_value)
            if status_value == "qualified":
                owner_route_names.add(kind)
    if endpoint == "raw":
        _expect(raw_routes, f"{path}: no raw capability was supplied")
    else:
        _expect(not raw_routes, f"{path}: owner endpoint carries raw capabilities")
    raw_fidelity = report.get("raw_fidelity")
    _expect(isinstance(raw_fidelity, list), f"{path}: raw fidelity receipts are missing")
    fidelity_routes: set[str] = set()
    for receipt in raw_fidelity:
        _expect(isinstance(receipt, dict), f"{path}: raw fidelity receipt is not an object")
        part_id = receipt.get("part_id")
        _expect(
            isinstance(part_id, str) and part_id in raw_routes and part_id not in fidelity_routes,
            f"{path}: raw fidelity part identity is missing or duplicated",
        )
        fidelity_routes.add(part_id)
        for key in (
            "source_sha256",
            "source_output_sha256",
            "changed_sha256",
            "changed_output_sha256",
            "malformed_sha256",
            "source_readback_sha256",
            "changed_readback_sha256",
        ):
            _expect(_is_hash(receipt.get(key)), f"{path}: raw fidelity {key} is missing")
        _expect(
            isinstance(receipt.get("source_preserving_expected"), bool)
            and isinstance(receipt.get("source_output_exact"), bool),
            f"{path}: raw fidelity source-preservation receipt is missing",
        )
        _expect(
            receipt["source_readback_sha256"] == receipt["source_output_sha256"]
            and receipt["changed_readback_sha256"] == receipt["changed_output_sha256"],
            f"{path}: raw fidelity readback is not canonical",
        )
        if receipt["source_preserving_expected"]:
            _expect(
                receipt["source_output_exact"] is True
                and receipt["source_output_sha256"] == receipt["source_sha256"],
                f"{path}: promised source-preserving route changed bytes",
            )
        _expect(
            receipt.get("fixture_identity_sha256")
            == _fixture_part_identity(
                expected_fixture_lane_identity,
                part_id,
                receipt["source_sha256"],
                receipt["changed_sha256"],
                receipt["malformed_sha256"],
            ),
            f"{path}: raw fidelity fixture identity is not bound to the part hashes",
        )
        _expect(
            receipt.get("changed_output_differs") is True
            and receipt["source_sha256"] != receipt["changed_sha256"]
            and receipt["source_output_sha256"] != receipt["changed_output_sha256"],
            f"{path}: raw fidelity does not prove the changed readback",
        )
    if endpoint == "raw":
        _expect(fidelity_routes == raw_routes, f"{path}: raw fidelity omits a part")
    else:
        _expect(not fidelity_routes, f"{path}: owner endpoint carries raw fidelity")
    family_mutations = report.get("theme_family_mutations")
    _expect(isinstance(family_mutations, list), f"{path}: theme-family mutation receipts are missing")
    if endpoint == "raw" and report.get("raw_scope") == "family":
        _expect(
            {item.get("operation") for item in family_mutations if isinstance(item, dict)}
            == {"replace_family", "remove_family", "add_family"},
            f"{path}: theme-family add/replace/remove matrix differs",
        )
        mutation_by_operation: dict[str, Mapping[str, Any]] = {}
        for item in family_mutations:
            _expect(isinstance(item, dict), f"{path}: theme-family mutation is not an object")
            operation = item.get("operation")
            _expect(
                isinstance(operation, str) and operation not in mutation_by_operation,
                f"{path}: duplicate theme-family mutation operation",
            )
            mutation_by_operation[operation] = item
            _expect(_is_hash(item.get("source_sha256")), f"{path}: theme-family mutation source hash is missing")
            _expect(_is_hash(item.get("output_sha256")), f"{path}: theme-family mutation output hash is missing")
            _expect(
                isinstance(item.get("output_bytes"), int) and item["output_bytes"] > 0,
                f"{path}: theme-family mutation output size is missing",
            )
            _expect(item.get("output_differs") is True, f"{path}: theme-family mutation was a no-op")
        replace = mutation_by_operation["replace_family"]
        remove = mutation_by_operation["remove_family"]
        add = mutation_by_operation["add_family"]
        _expect(
            replace["source_sha256"] == remove["source_sha256"]
            and add["source_sha256"] == remove["output_sha256"],
            f"{path}: theme-family mutation source lineage is incomplete",
        )
    else:
        _expect(not family_mutations, f"{path}: theme-family mutation receipts leaked outside family scope")
    owner_applicable = endpoint == "owner"
    if owner_applicable:
        _expect(owner_route_names, f"{path}: owner endpoint has no qualified owner route")
    for operation, statuses in owner_operations.items():
        if endpoint == "raw":
            _expect(not statuses, f"{path}: raw endpoint carries owner capability {operation}")
            continue
        _expect(statuses, f"{path}: owner capability {operation} is missing")
        if owner_applicable:
            _expect(
                all(status_value == "qualified" for status_value in statuses),
                f"{path}: owner capability {operation} is only partially qualified",
            )

    expected_operations = list(RAW_OPERATIONS) if endpoint == "raw" else list(OWNER_OPERATIONS)
    _expect(
        [row.get("operation") for row in rows] == expected_operations,
        f"{path}: operation matrix differs",
    )
    protocol = report.get("protocol")
    _expect(isinstance(protocol, dict), f"{path}: protocol receipt is missing")
    _expect(protocol.get("minimum_ms") == MINIMUM_MS, f"{path}: minimum duration was weakened")
    _expect(protocol.get("target_ms") == TARGET_MS, f"{path}: target duration differs")
    _expect(
        protocol.get("warmups") == WARMUPS
        and protocol.get("samples_per_leg") == SAMPLES
        and protocol.get("blocks") == BLOCKS,
        f"{path}: sample shape differs",
    )
    _expect(protocol.get("fresh_process") is True, f"{path}: fresh-process provenance is missing")
    _expect(protocol.get("abba_order") == list(ABBA), f"{path}: ABBA order differs")
    expected_samples = BLOCKS * SAMPLES
    sink_operations = {
        "entry_read",
        "read_shared_when_supported",
        "vec_write",
        "write_to_bounded_sink",
        "noop_commit",
        "changed_commit",
        "forward_replay",
        "inverse_replay",
    }
    for row in rows:
        _expect(row.get("status") == "qualified", f"{path}: qualified report has an unavailable row")
        elapsed = row.get("elapsed_ns")
        _expect(
            isinstance(elapsed, int) and elapsed >= MINIMUM_MS * 1_000_000,
            f"{path}: accepted leg is shorter than {MINIMUM_MS} ms",
        )
        samples = row.get("samples")
        _expect(
            isinstance(samples, list) and len(samples) == expected_samples,
            f"{path}: sample receipt shape differs",
        )
        sample_elapsed = 0
        for sample in samples:
            _expect(isinstance(sample, dict), f"{path}: sample receipt is not an object")
            _expect(
                isinstance(sample.get("elapsed_ns"), int) and sample["elapsed_ns"] > 0,
                f"{path}: sample timer is missing",
            )
            _expect(
                isinstance(sample.get("iterations"), int) and sample["iterations"] > 0,
                f"{path}: sample iteration receipt is missing",
            )
            sample_elapsed += sample["elapsed_ns"]
        _expect(sample_elapsed == elapsed, f"{path}: row timer is not the sum of measured samples")
        _expect(row.get("fidelity_checked") is True, f"{path}: row lacks outside-clock fidelity check")
        details = row.get("details")
        _expect(isinstance(details, dict), f"{path}: semantic details are missing")
        operation = row.get("operation")
        if operation == "read_shared_when_supported" and details.get("not_applicable") is True:
            _expect(
                not shared_reader_supported,
                f"{path}: shared-reader capability was supplied but the row is N/A",
            )
            _expect(
                row.get("sink") is None and details.get("reason"),
                f"{path}: shared-reader N/A receipt is incomplete",
            )
        elif operation == "write_to_bounded_sink":
            if details.get("not_applicable") is True:
                _expect(
                    row.get("sink") is None
                    and details.get("reason")
                    and details.get("production_sink_parts") == [],
                    f"{path}: bounded-sink N/A receipt is incomplete",
                )
            else:
                sink = row.get("sink")
                _expect(isinstance(sink, dict), f"{path}: bounded sink receipt is missing")
                _expect(_is_hash(sink.get("sha256")), f"{path}: bounded sink hash is missing")
                _expect(
                    isinstance(sink.get("bytes"), int) and sink["bytes"] > 0,
                    f"{path}: bounded sink byte receipt is missing",
                )
                _expect(
                    isinstance(sink.get("writes"), int)
                    and sink["writes"] > 0
                    and row.get("sink_writes") == sink["writes"],
                    f"{path}: bounded sink write receipt is incomplete",
                )
                _expect(
                    details.get("sink_hash") == sink["sha256"]
                    and details.get("sink_bytes") == sink["bytes"]
                    and details.get("sink_writes") == sink["writes"],
                    f"{path}: bounded sink detail receipt differs from row receipt",
                )
                expected_parts = sorted(
                    route
                    for route, capability in raw_capabilities.items()
                    if capability["production_sink"]
                )
                _expect(
                    sorted(details.get("route", {}).get("production_sink_parts", []))
                    == expected_parts,
                    f"{path}: bounded sink does not identify production routes",
                )
        elif operation in sink_operations - {"write_to_bounded_sink"}:
            sink = row.get("sink")
            _expect(isinstance(sink, dict), f"{path}: {operation} sink receipt is missing")
            _expect(_is_hash(sink.get("sha256")), f"{path}: {operation} sink hash is missing")
            _expect(
                isinstance(sink.get("bytes"), int) and sink["bytes"] > 0,
                f"{path}: {operation} sink byte receipt is missing",
            )
            _expect(
                isinstance(sink.get("writes"), int)
                and sink["writes"] > 0
                and row.get("sink_writes") == sink["writes"],
                f"{path}: {operation} sink write receipt is incomplete",
            )
            _expect(
                details.get("sink_hash") == sink["sha256"]
                and details.get("sink_bytes") == sink["bytes"]
                and details.get("sink_writes") == sink["writes"],
                f"{path}: {operation} sink detail receipt differs from row receipt",
            )
        elif operation == "stale_replay":
            _expect(
                row.get("typed_refusal_checked") is True,
                f"{path}: stale replay refusal was not checked",
            )
            refusal = details.get("refusal")
            _refusal(refusal, "patch_conflict", path, "stale replay")
            refusals = details.get("refusals")
            _expect(
                isinstance(refusals, list) and refusals,
                f"{path}: stale replay has no per-route refusal receipts",
            )
            for value in refusals:
                _refusal(value, "patch_conflict", path, "stale replay")
        elif operation == "preflight_refusal":
            _expect(
                row.get("typed_refusal_checked") is True
                and details.get("preflight_zero_writes") is True
                and details.get("sink_writes") == 0,
                f"{path}: preflight wrote before refusal",
            )
            parts = details.get("parts")
            _expect(isinstance(parts, list) and parts, f"{path}: preflight part receipts are missing")
            seen = set()
            for part in parts:
                _expect(isinstance(part, dict), f"{path}: preflight part is not an object")
                part_id = part.get("part_id")
                _expect(
                    isinstance(part_id, str) and part_id not in seen,
                    f"{path}: duplicate preflight part identity",
                )
                seen.add(part_id)
                capability = raw_capabilities.get(part_id)
                _expect(capability is not None, f"{path}: preflight part is not a capability route")
                boundary = part.get("boundary")
                _expect(
                    boundary == capability["production_boundary"],
                    f"{path}: preflight boundary differs from capability matrix",
                )
                sink_supported = capability["production_sink"]
                _expect(
                    part.get("sink_supported") is sink_supported,
                    f"{path}: preflight sink capability differs from capability matrix",
                )
                production_refusal = part.get("production_refusal")
                if boundary == "none":
                    _expect(production_refusal is None, f"{path}: unbounded part has a production refusal")
                else:
                    expected_kind = "limit" if boundary == "input_limit" else None
                    if expected_kind is None:
                        _expect(
                            isinstance(production_refusal, dict)
                            and production_refusal.get("kind") in {"limit", "invalid_input"},
                            f"{path}: structural production refusal is missing",
                        )
                        expected_kind = production_refusal["kind"]
                    _refusal(production_refusal, expected_kind, path, "production boundary")
                    if expected_kind == "limit":
                        _expect(
                            production_refusal.get("limit_kind")
                            and isinstance(production_refusal.get("limit"), int)
                            and production_refusal["observed"] > production_refusal["limit"],
                            f"{path}: production boundary lacks typed limit details",
                        )
                if sink_supported:
                    _expect(
                        isinstance(part.get("preflight_bytes"), int)
                        and part["preflight_bytes"] > 0
                        and part.get("sink_writes") == 0,
                        f"{path}: preflight part sink receipt is incomplete",
                    )
                    _expect(
                        part.get("sink_limit") == part["preflight_bytes"] - 1
                        and part.get("observed_output") == part["preflight_bytes"],
                        f"{path}: preflight sink bound is not bound to the measured output",
                    )
                    _refusal(part.get("sink_refusal"), "io_error", path, "preflight sink")
                else:
                    _expect(
                        part.get("sink_writes") == 0
                        and "preflight_bytes" not in part
                        and "sink_refusal" not in part,
                        f"{path}: non-sink route carries synthetic preflight evidence",
                    )
            _expect(seen == raw_routes, f"{path}: preflight does not cover every raw part")
        elif operation == "malformed_refusal":
            _expect(
                row.get("typed_refusal_checked") is True
                and details.get("malformed_refusal") is True,
                f"{path}: malformed refusal receipt is incomplete",
            )
            parts = details.get("parts")
            _expect(isinstance(parts, list) and parts, f"{path}: malformed part receipts are missing")
            seen = set()
            for part in parts:
                _expect(isinstance(part, dict), f"{path}: malformed part is not an object")
                part_id = part.get("part_id")
                _expect(
                    isinstance(part_id, str) and part_id not in seen,
                    f"{path}: duplicate malformed part identity",
                )
                seen.add(part_id)
                _refusal(part.get("refusal"), "invalid_input", path, "malformed")
            _expect(seen == raw_routes, f"{path}: malformed refusal does not cover every raw part")
            limit = details.get("limit")
            _expect(isinstance(limit, dict), f"{path}: malformed row lacks limit receipts")
            limit_parts = limit.get("parts")
            _expect(isinstance(limit_parts, list), f"{path}: limit part receipts are missing")
            limit_seen: set[str] = set()
            for part in limit_parts:
                _expect(isinstance(part, dict), f"{path}: limit part is not an object")
                part_id = part.get("part_id")
                _expect(
                    isinstance(part_id, str)
                    and part_id in raw_routes
                    and part_id not in limit_seen,
                    f"{path}: limit part identity is missing or duplicated",
                )
                limit_seen.add(part_id)
                boundary = part.get("boundary")
                capability = raw_capabilities[part_id]
                _expect(
                    boundary == capability["production_boundary"]
                    and boundary != "none",
                    f"{path}: limit boundary differs from capability matrix",
                )
                expected_kind = part.get("expected_kind")
                _expect(
                    expected_kind in {"limit", "invalid_input"},
                    f"{path}: limit expected kind is missing",
                )
                _refusal(part.get("refusal"), expected_kind, path, "limit")
                if expected_kind == "limit":
                    _expect(
                        isinstance(part["refusal"].get("limit_kind"), str)
                        and isinstance(part["refusal"].get("limit"), int)
                        and part["refusal"]["observed"] > part["refusal"]["limit"],
                        f"{path}: limit receipt lacks typed boundary details",
                    )
            _expect(limit_seen == {part_id for part_id, capability in raw_capabilities.items() if capability["production_boundary"] != "none"}, f"{path}: limit receipts omit a production boundary")
        else:
            _expect(
                operation in {"entry_read", "read_shared_when_supported", "vec_write", "write_to_bounded_sink"}
                or operation in OWNER_OPERATIONS,
                f"{path}: unknown qualified operation {operation}",
            )
    if endpoint == "owner":
        _expect(report.get("size_classes") == [], f"{path}: owner endpoint carries raw size classes")
        proofs = report.get("owner_proofs")
        _expect(isinstance(proofs, list) and proofs, f"{path}: owner proof list is missing")
        proof_routes = set()
        for proof in proofs:
            _expect(isinstance(proof, dict), f"{path}: owner proof is not an object")
            route = proof.get("route")
            _expect(
                isinstance(route, str) and route in owner_route_names and route not in proof_routes,
                f"{path}: owner proof route is missing or duplicated",
            )
            proof_routes.add(route)
            for key in (
                "source_sha256",
                "noop_sha256",
                "stale_sha256",
                "changed_sha256",
                "forward_sha256",
                "inverse_sha256",
            ):
                _expect(_is_hash(proof.get(key)), f"{path}: owner proof {key} is missing")
            _expect(
                proof.get("fixture_identity_sha256")
                == _owner_fixture_identity(
                    expected_fixture_lane_identity,
                    route,
                    proof["source_sha256"],
                    proof["stale_sha256"],
                ),
                f"{path}: owner proof fixture identity is not bound to its source",
            )
            _expect(
                proof["noop_sha256"] == proof["source_sha256"],
                f"{path}: owner no-op commit did not preserve source bytes",
            )
            _expect(
                proof["source_sha256"] != proof["changed_sha256"]
                and proof["forward_sha256"] == proof["changed_sha256"]
                and proof["inverse_sha256"] == proof["source_sha256"],
                f"{path}: owner proof identities are not reversible",
            )
            _expect(proof.get("stale_kind") == "patch_conflict", f"{path}: owner proof stale kind differs")
            _refusal(proof.get("stale_refusal"), "patch_conflict", path, "owner stale")
        _expect(proof_routes == owner_route_names, f"{path}: owner proofs do not cover capabilities")
        return status
    shared_row = next(row for row in rows if row.get("operation") == "read_shared_when_supported")
    if shared_reader_supported:
        _expect(
            isinstance(shared_row.get("sink"), dict),
            f"{path}: shared-reader capability has no sink receipt",
        )
    else:
        _expect(
            shared_row.get("details", {}).get("not_applicable") is True,
            f"{path}: unsupported shared-reader lane lacks an N/A receipt",
        )
    size_classes = report.get("size_classes")
    _expect(isinstance(size_classes, list) and size_classes, f"{path}: size-class matrix is missing")
    by_part: dict[str, dict[str, Mapping[str, Any]]] = {}
    for item in size_classes:
        _expect(isinstance(item, dict), f"{path}: size-class row is not an object")
        part_id = item.get("part_id")
        class_name = item.get("class")
        _expect(
            isinstance(part_id, str) and part_id in raw_routes,
            f"{path}: size-class part identity is unknown",
        )
        _expect(
            class_name in {"small", "near_limit", "at_limit", "over_limit", "malformed"},
            f"{path}: unsupported size class",
        )
        _expect(
            part_id not in by_part or class_name not in by_part[part_id],
            f"{path}: duplicate size-class receipt",
        )
        by_part.setdefault(part_id, {})[class_name] = item
        _expect(
            isinstance(item.get("bytes"), int) and item["bytes"] >= 0,
            f"{path}: size-class byte receipt is missing",
        )
        capability = raw_capabilities[part_id]
        _expect(
            item.get("boundary") == capability["production_boundary"],
            f"{path}: size-class boundary differs from capability matrix",
        )
        applicable = item.get("applicable")
        _expect(isinstance(applicable, bool), f"{path}: size-class applicability is missing")
        if not applicable:
            _expect(
                class_name == "over_limit"
                and item["bytes"] == 0
                and item.get("accepted") is False
                and item.get("refusal") is None,
                f"{path}: non-applicable size class is not explicit",
            )
            continue
        if class_name in {"small", "near_limit", "at_limit"}:
            _expect(
                item.get("accepted") is True
                and _is_hash(item.get("output_sha256"))
                and item.get("refusal") is None,
                f"{path}: accepted size class lacks readback identity",
            )
        elif class_name == "over_limit":
            _expect(item.get("accepted") is False, f"{path}: over-limit input was accepted")
            expected_kind = "limit" if item["boundary"] == "input_limit" else (
                item.get("refusal", {}).get("kind")
                if isinstance(item.get("refusal"), dict)
                else None
            )
            _expect(
                expected_kind in {"limit", "invalid_input"},
                f"{path}: structural over-limit refusal kind is missing",
            )
            _refusal(item.get("refusal"), expected_kind, path, "over-limit")
            if expected_kind == "limit":
                _expect(
                    item["bytes"] > item["refusal"]["limit"]
                    and item["refusal"]["observed"] == item["bytes"],
                    f"{path}: over-limit observed bytes are not bound to the typed limit",
                )
        else:
            _expect(item.get("accepted") is False, f"{path}: malformed size class was accepted")
            _refusal(item.get("refusal"), "invalid_input", path, "malformed size")
    _expect(set(by_part) == raw_routes, f"{path}: size classes omit a raw part")
    for part_id, classes in by_part.items():
        _expect(
            set(classes) == {"small", "near_limit", "at_limit", "over_limit", "malformed"},
            f"{path}: size classes omit a class for {part_id}",
        )

    proofs = report.get("owner_proofs")
    _expect(isinstance(proofs, list), f"{path}: owner proof list is missing")
    if owner_applicable:
        _expect(proofs, f"{path}: owner capability has no proof")
        proof_routes = set()
        for proof in proofs:
            _expect(isinstance(proof, dict), f"{path}: owner proof is not an object")
            route = proof.get("route")
            _expect(
                isinstance(route, str) and route in owner_route_names and route not in proof_routes,
                f"{path}: owner proof route is missing or duplicated",
            )
            proof_routes.add(route)
            for key in (
                "source_sha256",
                "noop_sha256",
                "stale_sha256",
                "changed_sha256",
                "forward_sha256",
                "inverse_sha256",
            ):
                _expect(_is_hash(proof.get(key)), f"{path}: owner proof {key} is missing")
            _expect(
                proof.get("fixture_identity_sha256")
                == _owner_fixture_identity(
                    expected_fixture_lane_identity,
                    route,
                    proof["source_sha256"],
                    proof["stale_sha256"],
                ),
                f"{path}: owner proof fixture identity is not bound to its source",
            )
            _expect(
                proof["noop_sha256"] == proof["source_sha256"],
                f"{path}: owner no-op commit did not preserve source bytes",
            )
            _expect(
                proof["source_sha256"] != proof["changed_sha256"]
                and proof["forward_sha256"] == proof["changed_sha256"]
                and proof["inverse_sha256"] == proof["source_sha256"],
                f"{path}: owner proof identities are not reversible",
            )
            _expect(proof.get("stale_kind") == "patch_conflict", f"{path}: owner proof stale kind differs")
            _refusal(proof.get("stale_refusal"), "patch_conflict", path, "owner stale")
        _expect(proof_routes == owner_route_names, f"{path}: owner proofs do not cover capabilities")
    else:
        _expect(proofs == [], f"{path}: raw-only lane carries owner proofs")
    return status



def validate(arguments: argparse.Namespace) -> dict[str, Any]:
    output_dir = arguments.output_dir.resolve(strict=False)
    _expect(output_dir.is_dir() and not output_dir.is_symlink(), f"output directory is invalid: {output_dir}")
    _expect(arguments.lane in LANES, f"unsupported lane: {arguments.lane}")
    endpoint = getattr(arguments, "endpoint", "raw")
    _expect(endpoint in {"raw", "owner"}, f"unsupported endpoint: {endpoint}")
    fixture = arguments.fixture.resolve(strict=False)
    fixture_hash, fixture_bytes = _sha256(fixture)
    try:
        fixture_data = json.loads(fixture.read_text(encoding="utf-8"))
        fixture_root = fixture_data["lanes"][arguments.lane]["root"]
    except (OSError, KeyError, TypeError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"fixture descriptor cannot provide lane root: {error}") from error
    fixture_lane_identity = _fixture_lane_identity(arguments.lane, fixture_data)
    reports: list[dict[str, Any]] = []
    unavailable = False
    for mode, prefix in (("normal", "dml-native"), ("allocator", "dml-allocator")):
        for index, path in enumerate(_ordered_outputs(output_dir, prefix)):
            report = _load(path)
            _expect(report.get("schema") == SCHEMA, f"{path}: schema differs")
            _expect(report.get("lane") == arguments.lane, f"{path}: lane differs")
            _expect(report.get("endpoint", "raw") == endpoint, f"{path}: endpoint differs")
            _expect(report.get("mode") == mode, f"{path}: mode differs")
            expected_role = ABBA[index]
            _expect(report.get("role") == expected_role, f"{path}: expected role {expected_role}")
            expected_revision = arguments.baseline_revision if expected_role == "baseline" else arguments.candidate_revision
            _expect(report.get("source_revision") == expected_revision, f"{path}: source revision differs")
            _expect(report.get("fixture_bytes") == fixture_bytes, f"{path}: fixture byte count differs")
            status = _validate_rows(report, path, fixture_lane_identity)
            if status == "qualified":
                _expect(report.get("fixture_sha256") == fixture_hash, f"{path}: fixture SHA-256 differs")
                _expect(report.get("fixture_root_sha256") == hashlib.sha256(fixture_root.encode()).hexdigest(), f"{path}: fixture root identity differs")
                _expect(isinstance(report.get("fixture_mutation"), str) and report["fixture_mutation"], f"{path}: fixture mutation receipt is missing")
            unavailable |= status == "unavailable"
            report_hash, report_bytes = _sha256(path)
            reports.append({
                "path": str(path),
                "sha256": report_hash,
                "bytes": report_bytes,
                "mode": mode,
                "role": expected_role,
                "status": status,
            })
            allocator_metrics = report.get("allocator_metrics")
            resource_metrics = report.get("resource_metrics")
            _expect(isinstance(resource_metrics, dict), f"{path}: process RSS receipt is missing")
            for key in ("process_rss_bytes", "peak_rss_bytes"):
                _expect(key in resource_metrics, f"{path}: resource metric {key} is missing")
                if status == "qualified":
                    _expect(isinstance(resource_metrics[key], int) and resource_metrics[key] > 0, f"{path}: qualified {key} is missing")
            if status == "qualified":
                _expect(
                    isinstance(resource_metrics.get("pid"), int)
                    and resource_metrics["pid"] > 0,
                    f"{path}: qualified resource PID receipt is missing",
                )
                _expect(
                    _is_hash(resource_metrics.get("command_sha256")),
                    f"{path}: qualified resource command identity is missing",
                )
                _expect(
                    resource_metrics.get("source_revision") == report.get("source_revision"),
                    f"{path}: RSS receipt is not bound to the source revision",
                )
                _expect(
                    isinstance(resource_metrics.get("whole_process_peak_rss_bytes"), int)
                    and resource_metrics["whole_process_peak_rss_bytes"] > 0,
                    f"{path}: qualified whole-process RSS high-water mark is missing",
                )
                _expect(
                    resource_metrics["whole_process_peak_rss_bytes"]
                    >= resource_metrics["peak_rss_bytes"],
                    f"{path}: whole-process RSS high-water mark is below sampled peak",
                )
            if mode == "normal":
                _expect(allocator_metrics is None, f"{path}: normal report exposes allocator metrics")
            else:
                _expect(isinstance(allocator_metrics, dict), f"{path}: allocator report has no allocator receipt")
                _expect(allocator_metrics.get("elapsed_claimable") is False, f"{path}: allocator elapsed was made claimable")
                metric_keys = ("direct_bytes", "realloc_bytes", "live_balance", "peak_live_delta", "process_rss_bytes", "peak_rss_bytes") if status == "unavailable" else ("direct_bytes", "deallocated_bytes", "realloc_old_bytes", "realloc_new_bytes", "live_before", "live_after", "live_balance", "peak_live_delta", "process_rss_bytes", "peak_rss_bytes", "whole_process_peak_rss_bytes", "allocation_calls", "deallocation_calls", "reallocation_calls", "failed_calls", "conservation_ok", "pid", "command_sha256", "source_revision")
                for key in metric_keys:
                    _expect(key in allocator_metrics, f"{path}: allocator metric {key} is missing")
                    if status == "qualified":
                        if key == "conservation_ok":
                            _expect(allocator_metrics[key] is True, f"{path}: allocator conservation is not proven")
                        elif key == "command_sha256":
                            _expect(_is_hash(allocator_metrics[key]), f"{path}: allocator command identity is invalid")
                        elif key == "source_revision":
                            _expect(allocator_metrics[key] == report.get("source_revision"), f"{path}: allocator source identity differs")
                        else:
                            _expect(isinstance(allocator_metrics[key], int) and allocator_metrics[key] >= 0, f"{path}: qualified allocator metric {key} is invalid")
                if status == "qualified":
                    _expect(allocator_metrics.get("allocation_calls", 0) > 0, f"{path}: allocator observed no allocation calls")
                    _expect(allocator_metrics["live_after"] - allocator_metrics["live_before"] == allocator_metrics["live_balance"], f"{path}: allocator live balance is inconsistent")
                    _expect(allocator_metrics["live_balance"] == allocator_metrics["direct_bytes"] + allocator_metrics["realloc_new_bytes"] - allocator_metrics["deallocated_bytes"] - allocator_metrics["realloc_old_bytes"], f"{path}: allocator byte conservation equation is not proven")
                    _expect(allocator_metrics["peak_live_delta"] >= max(0, allocator_metrics["live_balance"]), f"{path}: allocator peak-live receipt is inconsistent")
                    _expect(
                        allocator_metrics["process_rss_bytes"] == resource_metrics["process_rss_bytes"]
                        and allocator_metrics["peak_rss_bytes"] == resource_metrics["peak_rss_bytes"]
                        and allocator_metrics["whole_process_peak_rss_bytes"] == resource_metrics["whole_process_peak_rss_bytes"]
                        and allocator_metrics.get("pid") == resource_metrics.get("pid")
                        and allocator_metrics.get("command_sha256") == resource_metrics.get("command_sha256")
                        and allocator_metrics.get("source_revision") == resource_metrics.get("source_revision"),
                        f"{path}: allocator and process RSS receipts differ",
                    )
    return {
        "schema": SUMMARY_SCHEMA,
        "lane": arguments.lane,
        "endpoint": endpoint,
        "baseline_revision": arguments.baseline_revision,
        "candidate_revision": arguments.candidate_revision,
        "fixture": {"path": str(fixture), "sha256": fixture_hash, "bytes": fixture_bytes},
        "source_fence": arguments.source_fence,
        "abba_order": list(ABBA),
        "qualification_status": "unavailable" if unavailable else "qualified",
        "accepted": not unavailable,
        "claim_scope": "paired DrawingML raw and detached-owner endpoints only; no absolute-profile or whole-library claim",
        "reports": reports,
        "unavailable_reason": "an external source receipt is unavailable" if unavailable else None,
    }


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--lane", required=True)
    parser.add_argument("--endpoint", choices=("raw", "owner"), default="raw")
    parser.add_argument("--fixture", type=Path, required=True)
    parser.add_argument("--baseline-revision", required=True)
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument("--source-fence", required=True)
    arguments = parser.parse_args(argv)
    try:
        summary = validate(arguments)
        target = arguments.output_dir / "summary.json"
        _expect(not target.exists() and not target.is_symlink(), f"refusing to overwrite {target}")
        target.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    except (OSError, ValidationError) as error:
        print(f"drawingml raw validation refused: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
