#!/usr/bin/env python3
"""Validation rules for the signed native qualification protocol.

This module is intentionally independent of the Rust probe.  It validates
the probe's correctness receipts, endpoint provenance, and the before/after
capability boundary before a guarded timing run is allowed to start.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import math
import re
from typing import Any

SCHEMA = "litchi_signed_native_probe.v1"
ENDPOINT_SCHEMA = "litchi_signed_native_endpoint.v1"
FULL_COMMIT = re.compile(r"^[0-9a-f]{40}$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")

TARGET_NS = 500_000_000
MIN_SAMPLE_NS = 150_000_000
MAX_CALLS = 2_000_000
MIN_ABBA_ROUNDS = 5
BASELINE_SOURCE_COMMIT = "f2bbbc8d07ab5da26c180f1ff74813d39f632d54"
CANDIDATE_SOURCE_COMMIT = "472b6705475b073e7fe08f6e11ee17fb1e29f838"

ROUTES = (
    "signed_leaf/noop_commit",
    "signed_facade/noop_commit",
    "signed_leaf/noop_write",
    "signed_facade/noop_write",
    "signed_leaf/noop_replay",
    "signed_facade/noop_replay",
    "signed_leaf/refusal",
    "signed_facade/refusal",
    "signed_leaf/unsign_write",
    "signed_facade/unsign_write",
    "signed_leaf/unsigned_negative",
    "signed_facade/unsigned_negative",
    "signed_leaf/source_freshness",
    "signed_facade/source_freshness",
    "signed_leaf/budget",
    "signed_facade/budget",
)

CORRECTNESS_ONLY = frozenset(
    {
        "signed_leaf/refusal",
        "signed_facade/refusal",
        "signed_leaf/unsigned_negative",
        "signed_facade/unsigned_negative",
        "signed_leaf/source_freshness",
        "signed_facade/source_freshness",
        "signed_leaf/budget",
        "signed_facade/budget",
    }
)
TIMED = tuple(route for route in ROUTES if route not in CORRECTNESS_ONLY)
FINAL_ONLY = frozenset(
    {
        "signed_leaf/noop_commit",
        "signed_facade/noop_commit",
        "signed_leaf/noop_write",
        "signed_facade/noop_write",
        "signed_leaf/noop_replay",
        "signed_facade/noop_replay",
        "signed_leaf/unsign_write",
        "signed_facade/unsign_write",
    }
)
BASELINE_NOT_APPLICABLE = frozenset(
    {
        "signed_leaf/unsigned_negative",
        "signed_facade/unsigned_negative",
        "signed_leaf/source_freshness",
        "signed_facade/source_freshness",
        "signed_leaf/budget",
        "signed_facade/budget",
    }
)

ORIGINS = {"Opc", "Cfb"}
CRYPTO_STATUS = "not_performed"
NATIVE_FORMATS = frozenset({"DOCX", "PPTX", "XLSX", "XLSB", "DOC", "PPT", "XLS"})


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def artifact(path: Path) -> dict[str, object]:
    path = path.resolve()
    if not path.is_file():
        raise ValueError(f"artifact is not a file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def canonical(value: object) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        + "\n"
    ).encode()


def load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValueError(f"{label} is invalid JSON: {error}") from error
    if not isinstance(value, dict):
        raise ValueError(f"{label} is not a JSON object")
    return value


def checked_hash(value: object, label: str) -> str:
    if not isinstance(value, str) or SHA256.fullmatch(value) is None:
        raise ValueError(f"{label} is not a SHA-256")
    return value


def checked_artifact(row: object, label: str) -> dict[str, object]:
    if not isinstance(row, dict):
        raise ValueError(f"{label} is missing")
    path = row.get("path")
    if not isinstance(path, str) or not path:
        raise ValueError(f"{label}.path is missing")
    expected = {
        "path": str(Path(path).resolve()),
        "bytes": row.get("bytes"),
        "sha256": checked_hash(row.get("sha256"), f"{label}.sha256"),
    }
    actual = artifact(Path(path))
    if actual != expected:
        raise ValueError(f"{label} changed: expected {expected}, found {actual}")
    return actual


def checked_relative_path(value: object, label: str) -> str:
    """Validate a fixture path before either consumer joins it to its root."""
    if not isinstance(value, str) or not value:
        raise ValueError(f"{label} is missing")
    path = Path(value)
    if path.is_absolute() or "\\" in value:
        raise ValueError(f"{label} must be a portable relative path")
    if any(part in {"", ".", ".."} for part in path.parts):
        raise ValueError(f"{label} contains traversal or an unnormalised component")
    return value


def load_fixture_manifest(value: dict[str, Any], label: str = "fixtures") -> dict[str, Any]:
    if value.get("schema") != "litchi_signed_fixture_manifest.v1":
        raise ValueError(f"{label} schema is invalid")
    fixtures = value.get("fixtures")
    if not isinstance(fixtures, list) or not fixtures:
        raise ValueError(f"{label}.fixtures is empty")
    labels: set[str] = set()
    formats: set[str] = set()
    for index, fixture in enumerate(fixtures):
        if not isinstance(fixture, dict):
            raise ValueError(f"{label}.fixtures[{index}] is not an object")
        name = fixture.get("name")
        format_name = fixture.get("format")
        origin = fixture.get("origin")
        mode = fixture.get("fixture_mode")
        if not isinstance(name, str) or not name or name in labels:
            raise ValueError(f"{label}.fixtures[{index}] has a duplicate name")
        if not isinstance(format_name, str) or not format_name:
            raise ValueError(f"{label}.fixtures[{index}] has no format")
        checked_relative_path(fixture.get("relative_path"), f"{label}.fixtures[{index}].relative_path")
        if origin not in ORIGINS or mode not in {"producer_signed", "synthetic_presence"}:
            raise ValueError(f"{label}.fixtures[{index}] has an invalid origin/mode")
        if fixture.get("signature_presence") != "detected":
            raise ValueError(f"{label}.fixtures[{index}] does not prove detection")
        if fixture.get("cryptographic_verification") != CRYPTO_STATUS:
            raise ValueError(f"{label}.fixtures[{index}] makes a crypto claim")
        checked_hash(fixture.get("input_sha256"), f"{label}.fixtures[{index}].input_sha256")
        if mode == "producer_signed":
            checked_hash(
                fixture.get("signed_sha256"),
                f"{label}.fixtures[{index}].signed_sha256",
            )
        elif fixture.get("signed_sha256") is not None:
            checked_hash(
                fixture.get("signed_sha256"),
                f"{label}.fixtures[{index}].signed_sha256",
            )
        elif fixture.get("signed_hash_mode") != "probe_generated":
            raise ValueError(
                f"{label}.fixtures[{index}] lacks a deterministic signed hash policy"
            )
        target = fixture.get("target")
        property_name = fixture.get("property")
        if (
            not isinstance(target, str)
            or not target
            or not isinstance(property_name, str)
            or not property_name
        ):
            raise ValueError(f"{label}.fixtures[{index}] native target is incomplete")
        labels.add(name)
        formats.add(format_name)
    if len(formats) != 7:
        raise ValueError(f"{label} covers {len(formats)} formats, expected 7 signed formats")
    return value


def validate_endpoint_manifest(path: Path, label: str) -> dict[str, Any]:
    value = load_json(path, label)
    if value.get("schema") != ENDPOINT_SCHEMA:
        raise ValueError(f"{label} schema is invalid")
    commit = value.get("source_commit")
    if not isinstance(commit, str) or FULL_COMMIT.fullmatch(commit) is None:
        raise ValueError(f"{label}.source_commit is not a full commit")
    for name in (
        "binary",
        "source",
        "source_archive",
        "lock",
        "probe",
        "fixtures",
        "dependency_projection",
        "build_receipt",
    ):
        checked_artifact(value.get(name), f"{label}.{name}")
    identity = value.get("source_identity")
    if not isinstance(identity, dict):
        raise ValueError(f"{label}.source_identity is missing")
    if identity.get("commit") != commit:
        raise ValueError(f"{label}.source_identity.commit does not match source_commit")
    if not isinstance(identity.get("source_ref"), str) or not identity["source_ref"]:
        raise ValueError(f"{label}.source_identity.source_ref is missing")
    if not isinstance(identity.get("source_capture"), str) or not identity["source_capture"]:
        raise ValueError(f"{label}.source_identity.source_capture is missing")
    source_root = identity.get("source_root")
    if (
        not isinstance(source_root, str)
        or not Path(source_root).is_absolute()
        or not Path(source_root).is_dir()
    ):
        raise ValueError(f"{label}.source_identity.source_root must be an existing absolute directory")
    checked_hash(identity.get("tree_sha256"), f"{label}.source_identity.tree_sha256")
    archive = checked_artifact(identity.get("archive"), f"{label}.source_identity.archive")
    if archive != value["source_archive"]:
        raise ValueError(f"{label}.source_identity.archive is not the endpoint source archive")
    adapter = checked_artifact(identity.get("adapter"), f"{label}.source_identity.adapter")
    if not adapter["path"].endswith(".rs"):
        raise ValueError(f"{label}.source_identity.adapter is not a Rust consumer")
    source_record = load_json(Path(value["source"]["path"]), f"{label}.source")
    for field in ("commit", "source_ref", "source_capture", "source_root", "tree_sha256"):
        if source_record.get(field) != identity.get(field):
            raise ValueError(f"{label}.source does not bind source_identity.{field}")
    build_receipt = load_json(Path(value["build_receipt"]["path"]), f"{label}.build_receipt")
    if build_receipt.get("source_commit") != commit:
        raise ValueError(f"{label}.build_receipt source commit differs")
    if build_receipt.get("binary_sha256") != value["binary"]["sha256"]:
        raise ValueError(f"{label}.build_receipt binary differs")
    if build_receipt.get("source_archive_sha256") != value["source_archive"]["sha256"]:
        raise ValueError(f"{label}.build_receipt source archive differs")
    if build_receipt.get("adapter_sha256") != adapter["sha256"]:
        raise ValueError(f"{label}.build_receipt adapter differs")
    if build_receipt.get("lock_sha256") != value["lock"]["sha256"]:
        raise ValueError(f"{label}.build_receipt lock differs")
    fixtures_path = Path(value["fixtures"]["path"])
    fixture_value = load_json(fixtures_path, f"{label}.fixtures")
    if canonical(fixture_value) != fixtures_path.read_bytes():
        raise ValueError(f"{label}.fixtures is not canonical JSON")
    load_fixture_manifest(fixture_value, f"{label}.fixtures")
    capabilities = value.get("capabilities")
    if not isinstance(capabilities, dict):
        raise ValueError(f"{label}.capabilities is missing")
    if set(capabilities) != set(ROUTES):
        raise ValueError(f"{label}.capabilities does not enumerate every signed route")
    for route, status in capabilities.items():
        if status not in {
            "available",
            "capability_missing",
            "expected_refusal",
            "not_applicable",
        }:
            raise ValueError(f"{label}.capabilities[{route}] has invalid status")
        if status == "capability_missing" and route not in FINAL_ONLY:
            raise ValueError(f"{label}.capabilities[{route}] cannot be missing")
        if status == "not_applicable" and route not in BASELINE_NOT_APPLICABLE:
            raise ValueError(f"{label}.capabilities[{route}] cannot be not_applicable")
    root = value.get("fixture_root")
    if not isinstance(root, str) or not Path(root).is_absolute() or not Path(root).is_dir():
        raise ValueError(f"{label}.fixture_root must be an existing absolute directory")
    return value


def shared_identity(before: dict[str, Any], after: dict[str, Any]) -> dict[str, str]:
    fields = ("lock", "probe", "fixtures", "dependency_projection")
    for field in fields:
        left = before[field]["sha256"]
        right = after[field]["sha256"]
        if left != right:
            raise ValueError(f"endpoint {field} differs: {left} != {right}")
    return {f"{field}_sha256": before[field]["sha256"] for field in fields}


def validate_pair(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    if before["source_commit"] != BASELINE_SOURCE_COMMIT:
        raise ValueError(
            f"before endpoint is not the pinned f2 baseline: {before['source_commit']}"
        )
    if after["source_commit"] != CANDIDATE_SOURCE_COMMIT:
        raise ValueError(
            f"after endpoint is not the pinned 472b capture: {after['source_commit']}"
        )
    if before["binary"]["path"] == after["binary"]["path"]:
        raise ValueError("before and after endpoint binaries must have distinct paths")
    if before["binary"]["sha256"] == after["binary"]["sha256"]:
        raise ValueError("before and after endpoint binaries must differ")
    if before["source_archive"]["sha256"] == after["source_archive"]["sha256"]:
        raise ValueError("before and after source archives must differ")
    if before["source_identity"]["tree_sha256"] == after["source_identity"]["tree_sha256"]:
        raise ValueError("before and after source trees must differ")
    shared = shared_identity(before, after)
    for route in ROUTES:
        final_status = after["capabilities"][route]
        before_status = before["capabilities"][route]
        if final_status == "capability_missing":
            raise ValueError(f"final endpoint lacks required route {route}")
        if route in BASELINE_NOT_APPLICABLE and before_status != "not_applicable":
            raise ValueError(f"baseline route {route} must be explicitly not_applicable")
        if route.endswith("/refusal") and final_status != "expected_refusal":
            raise ValueError(f"final refusal route {route} is not marked expected_refusal")
        if route not in CORRECTNESS_ONLY and final_status != "available":
            raise ValueError(f"final timed route {route} is not available")
    return {
        "shared": shared,
        "baseline_capability_gaps": [
            route for route in ROUTES if before["capabilities"][route] != "available"
        ],
    }


def _fixture_rows(value: dict[str, Any], route: str) -> list[dict[str, Any]]:
    if value.get("schema") != SCHEMA or value.get("route") != route:
        raise ValueError(f"{route}: probe schema or route identity is invalid")
    rows = value.get("fixtures")
    if not isinstance(rows, list) or not rows:
        raise ValueError(f"{route}: fixture results are empty")
    return rows


def _common_fixture_checks(row: dict[str, Any], route: str) -> None:
    for key in ("format", "name", "origin"):
        if not isinstance(row.get(key), str) or not row[key]:
            raise ValueError(f"{route}: fixture {key} is missing")
    if row.get("origin") not in ORIGINS:
        raise ValueError(f"{route}: unknown signature origin")
    if row.get("signature_presence") != "detected":
        raise ValueError(f"{route}: signature presence was not detected")
    if row.get("cryptographic_verification") != CRYPTO_STATUS:
        raise ValueError(f"{route}: cryptographic status is not explicit")
    checked_hash(row.get("source_sha256"), f"{route}: source")


def validate_route(
    value: dict[str, Any], route: str, *, endpoint: str = "final"
) -> dict[str, Any]:
    rows = _fixture_rows(value, route)
    capability = value.get("capability")
    if capability not in {
        "available",
        "capability_missing",
        "expected_refusal",
        "not_applicable",
    }:
        raise ValueError(f"{route}: invalid capability status")
    if capability == "not_applicable":
        if endpoint != "before" or route not in BASELINE_NOT_APPLICABLE:
            raise ValueError(f"{route}: not_applicable is only a baseline disposition")
        for row in rows:
            _common_fixture_checks(row, route)
            if (
                row.get("status") != "not_applicable"
                or row.get("disposition") != "requires_signed_noop_admission"
                or row.get("reason")
                != "the historical endpoint cannot construct a signed no-op base"
                or row.get("signed_before") is not True
            ):
                raise ValueError(f"{route}: baseline not_applicable disposition is incomplete")
        return {"route": route, "status": "not_applicable", "endpoint": endpoint}
    if capability == "capability_missing":
        if endpoint == "final" or route not in FINAL_ONLY:
            raise ValueError(f"{route}: required capability missing on {endpoint}")
        for row in rows:
            _common_fixture_checks(row, route)
            if (
                row.get("status") != "capability_missing"
                or row.get("refusal_kind") != "ProtectedDocument"
                or row.get("refusal_code") != "document_protected"
                or row.get("signed_before") is not True
            ):
                raise ValueError(f"{route}: baseline gap lacks its typed refusal proof")
        return {"route": route, "status": "capability_missing", "endpoint": endpoint}
    native = value.get("native_preflight")
    if not isinstance(native, list) or not native:
        raise ValueError(f"{route}: direct native preflight is missing")
    native_formats: set[str] = set()
    for native_row in native:
        if not isinstance(native_row, dict):
            raise ValueError(f"{route}: native preflight row is not an object")
        format_name = native_row.get("format")
        status = native_row.get("status")
        if not isinstance(format_name, str) or status not in {
            "exact_noop",
            "unsupported_signed_source",
        }:
            raise ValueError(f"{route}: invalid native preflight disposition")
        if format_name in native_formats:
            raise ValueError(f"{route}: duplicate native preflight format {format_name}")
        native_formats.add(format_name)
        if format_name in {"DOC", "PPT"} and status != "unsupported_signed_source":
            raise ValueError(f"{route}: {format_name} signed source should refuse direct source edit")
        if format_name not in {"DOC", "PPT"} and status != "exact_noop":
            raise ValueError(f"{route}: {format_name} direct native no-op was not proven")
    if native_formats != NATIVE_FORMATS:
        raise ValueError(
            f"{route}: direct native preflight covers {sorted(native_formats)}, "
            f"expected {sorted(NATIVE_FORMATS)}"
        )
    for row in rows:
        if not isinstance(row, dict):
            raise ValueError(f"{route}: fixture row is not an object")
        _common_fixture_checks(row, route)
        if route.endswith("/refusal"):
            if row.get("status") != "expected_refusal":
                raise ValueError(f"{route}: changed signed edit did not refuse")
            if (
                row.get("refusal_kind") != "ProtectedDocument"
                or row.get("code") != "document_protected"
            ):
                raise ValueError(
                    f"{route}: refusal classification is not protected/document_protected"
                )
            if row.get("source_after_sha256") != row.get("source_sha256"):
                raise ValueError(f"{route}: refusal changed source bytes")
            if row.get("output_published") is not False:
                raise ValueError(f"{route}: refusal published output")
            continue
        if route.endswith("/unsigned_negative"):
            if (
                row.get("status") != "expected_refusal"
                or row.get("refusal_kind") != "PatchConflict"
                or row.get("code") != "invalid_input"
            ):
                raise ValueError(
                    f"{route}: unsigned negative is not a typed patch conflict"
                )
            if row.get("re_marked_signed") is not False:
                raise ValueError(f"{route}: unsigned negative re-marked a signature")
            if row.get("recipient_signed_after") is not True or row.get("output_published") is not False:
                raise ValueError(f"{route}: patch conflict did not preserve the signed recipient")
            continue
        if route.endswith("/source_freshness"):
            if (
                row.get("status") != "expected_refusal"
                or row.get("refusal_kind") != "SourceChanged"
                or row.get("code") != "io_error"
            ):
                raise ValueError(f"{route}: source fence is not SourceChanged/io_error")
            checked_hash(row.get("source_mutated_sha256"), f"{route}: mutated source")
            if row.get("source_mutated_sha256") == row.get("source_sha256"):
                raise ValueError(f"{route}: source fence did not mutate the source")
            if row.get("output_published") is not False:
                raise ValueError(f"{route}: source fence published output")
            continue
        if route.endswith("/budget"):
            if (
                row.get("status") != "expected_refusal"
                or row.get("refusal_kind") != "LimitExceeded"
                or row.get("code") != "limit_exceeded"
            ):
                raise ValueError(f"{route}: budget refusal is not typed")
            if row.get("source_after_sha256") != row.get("source_sha256"):
                raise ValueError(f"{route}: budget refusal changed source bytes")
            if row.get("limit_kind") != "operations":
                raise ValueError(f"{route}: budget refusal lacks its typed limit kind")
            if row.get("output_published") is not False:
                raise ValueError(f"{route}: budget refusal published output")
            continue
        if row.get("status") != "ok":
            raise ValueError(f"{route}: available route has non-ok fixture result")
        if row.get("signed_before") is not True:
            raise ValueError(f"{route}: input is not signed")
        if route.endswith("/noop_commit"):
            if row.get("diagnostics_changed") is not False or row.get("patch_noop") is not True:
                raise ValueError(f"{route}: no-op diagnostics/patch proof failed")
            if (
                row.get("signed_after") is not True
                or row.get("output_sha256") != row.get("source_sha256")
            ):
                raise ValueError(f"{route}: signed no-op changed bytes or signature state")
            if row.get("output_equals_source") is not True:
                raise ValueError(f"{route}: signed no-op lacks exact byte proof")
            if row.get("source_after_sha256") != row.get("source_sha256"):
                raise ValueError(f"{route}: signed no-op changed its source endpoint")
        elif route.endswith("/noop_write"):
            if (
                row.get("diagnostics_changed") is not False
                or row.get("patch_noop") is not True
                or row.get("signed_after") is not True
                or row.get("output_sha256") != row.get("source_sha256")
            ):
                raise ValueError(f"{route}: signed publication changed no-op/state proof")
            if (
                row.get("stream_bounded") is not True
                or not isinstance(row.get("max_write_bytes"), int)
                or row["max_write_bytes"] <= 0
            ):
                raise ValueError(f"{route}: publication is not bounded")
            if row.get("output_equals_source") is not True:
                raise ValueError(f"{route}: signed publication lacks exact byte proof")
            if row.get("source_after_sha256") != row.get("source_sha256"):
                raise ValueError(f"{route}: signed publication changed its source endpoint")
        elif route.endswith("/noop_replay"):
            if row.get("patch_noop") is not True or row.get("replay_base_match") is not True:
                raise ValueError(f"{route}: replay lacks content-bound no-op proof")
            if (
                row.get("signed_after") is not True
                or row.get("output_sha256") != row.get("source_sha256")
            ):
                raise ValueError(f"{route}: replay changed signed publication")
            if row.get("output_equals_source") is not True:
                raise ValueError(f"{route}: replay lacks exact byte proof")
        elif route.endswith("/unsign_write"):
            if row.get("signed_after") is not False or row.get("signature_parts_absent") is not True:
                raise ValueError(f"{route}: unsign output retains signature state/parts")
            if row.get("semantic_changed") is not True or row.get("output_equals_source") is not False:
                raise ValueError(f"{route}: unsign route did not publish a real change")
            if row.get("target_changed") is not True:
                raise ValueError(f"{route}: unsign route did not change the native target")
            checked_hash(row.get("output_sha256"), f"{route}: output")
            if row.get("stream_sha256") != row.get("output_sha256"):
                raise ValueError(f"{route}: bounded stream digest differs from output")
            if row.get("stream_bounded") is not True:
                raise ValueError(f"{route}: unsign publication is not bounded")
            if row.get("signed_clone_still_refuses") is not True:
                raise ValueError(f"{route}: signed clone lost protection")
    return {"route": route, "status": "ok", "endpoint": endpoint, "fixtures": len(rows)}


def validate_correctness_receipt(
    value: dict[str, Any], route: str, *, endpoint: str = "final"
) -> dict[str, Any]:
    """Public wrapper used by tests and the guarded runner."""
    return validate_route(value, route, endpoint=endpoint)


def validate_measure(
    value: dict[str, Any], route: str, calls: int, *, expected_route: str | None = None
) -> None:
    actual_route = expected_route or route
    if value.get("schema") != SCHEMA or value.get("route") != actual_route:
        raise ValueError(f"{route}: timed output identity is invalid")
    if value.get("mode") != "measure" or value.get("calls") != calls:
        raise ValueError(f"{route}: timed call count/mode mismatch")
    elapsed = value.get("elapsed_ns")
    if not isinstance(elapsed, int) or elapsed < MIN_SAMPLE_NS:
        raise ValueError(f"{route}: sample is below the {MIN_SAMPLE_NS} ns floor")
    if calls <= 0 or calls > MAX_CALLS:
        raise ValueError(f"{route}: calls exceed the explicit ceiling")
    if not isinstance(value.get("operations"), int) or value["operations"] < calls:
        raise ValueError(f"{route}: operation count is not reported")


def validate_calibration(value: dict[str, Any], route: str) -> int:
    if (
        value.get("schema") != SCHEMA
        or value.get("route") != route
        or value.get("mode") != "calibrate"
    ):
        raise ValueError(f"{route}: malformed calibration")
    calls = value.get("common_calls")
    elapsed = value.get("elapsed_ns")
    if not isinstance(calls, int) or calls <= 0 or calls > MAX_CALLS:
        raise ValueError(f"{route}: calibration call count is outside the ceiling")
    if not isinstance(elapsed, int) or elapsed < TARGET_NS:
        raise ValueError(f"{route}: calibration did not reach the {TARGET_NS} ns target")
    if not isinstance(value.get("operations"), int) or value["operations"] < calls:
        raise ValueError(f"{route}: calibration operation count is not reported")
    if value.get("target_ns") != TARGET_NS:
        raise ValueError(f"{route}: calibration target differs from protocol")
    return calls


def validate_abba(manifest: dict[str, Any]) -> None:
    if manifest.get("schema") != "litchi_signed_native_qualification.v1":
        raise ValueError("qualification manifest schema is invalid")
    rounds = manifest.get("rounds")
    if not isinstance(rounds, int) or rounds < MIN_ABBA_ROUNDS:
        raise ValueError("qualification needs at least five ABBA rounds")
    if manifest.get("target_ns") != TARGET_NS or manifest.get("minimum_sample_ns") != MIN_SAMPLE_NS:
        raise ValueError("qualification timing constants differ from the protocol")
    rows = manifest.get("routes")
    if not isinstance(rows, list):
        raise ValueError("qualification route rows are missing")
    if len(rows) != len(TIMED) or {row.get("route") for row in rows if isinstance(row, dict)} != set(TIMED):
        raise ValueError("qualification must contain each timed route exactly once")
    for row in rows:
        if not isinstance(row, dict) or row.get("route") not in TIMED:
            raise ValueError("qualification contains an unknown timed route")
        calls = row.get("calls")
        if not isinstance(calls, int) or calls <= 0 or calls > MAX_CALLS:
            raise ValueError(f"{row.get('route')}: invalid shared call count")
        blocks = row.get("blocks")
        if not isinstance(blocks, list) or len(blocks) < rounds:
            raise ValueError(f"{row.get('route')}: incomplete ABBA blocks")
        route = row["route"]
        if row.get("baseline") == "leaf_facade":
            if route not in FINAL_ONLY:
                raise ValueError(f"{route}: leaf/facade comparison is final-only")
            paired = row.get("paired_route")
            expected_pair = _paired_route(route)
            if paired != expected_pair:
                raise ValueError(f"{route}: paired route is incomplete or incorrect")
            if row.get("comparison") != "final_leaf_vs_facade":
                raise ValueError(f"{route}: final comparison label is missing")
            calibration = row.get("calibration")
            if not isinstance(calibration, dict):
                raise ValueError(f"{route}: leaf/facade calibration is missing")
            for side, side_route in (("leaf", _leaf_route(route)), ("facade", _facade_route(route))):
                value = calibration.get(side)
                if not isinstance(value, dict):
                    raise ValueError(f"{route}: {side} calibration is missing")
                calibrated_calls = validate_calibration(value, side_route)
                if calibrated_calls > calls:
                    raise ValueError(f"{route}: shared calls undercut {side} calibration")
            for block in blocks:
                legs = block.get("legs") if isinstance(block, dict) else None
                if (
                    not isinstance(legs, list)
                    or [leg.get("side") for leg in legs]
                    != ["leaf", "facade", "facade", "leaf"]
                ):
                    raise ValueError(f"{route}: leaf/facade ABBA order is invalid")
                for leg in legs:
                    side = leg.get("side")
                    expected = _leaf_route(route) if side == "leaf" else _facade_route(route)
                    validate_measure(leg.get("value"), route, calls, expected_route=expected)
            leaf = _median_per_call(blocks, "leaf")
            facade = _median_per_call(blocks, "facade")
            _finite_ratio(row.get("leaf_ns_per_call"), leaf, route)
            _finite_ratio(row.get("facade_ns_per_call"), facade, route)
            expected_ratio = facade / leaf
            _finite_ratio(row.get("facade_leaf_ratio"), expected_ratio, route)
            if "after_before_ratio" in row or "historical_ratio" in row:
                raise ValueError(f"{route}: final-only row contains a historical ratio")
            continue
        if row.get("baseline") != "matched":
            raise ValueError(f"{route}: timing baseline is not a truthful comparison")
        if route in FINAL_ONLY:
            raise ValueError(f"{route}: final-only route cannot claim matched baseline")
        for block in blocks:
            legs = block.get("legs") if isinstance(block, dict) else None
            if (
                not isinstance(legs, list)
                or [leg.get("side") for leg in legs]
                != ["before", "after", "after", "before"]
            ):
                raise ValueError(f"{route}: ABBA order is invalid")
            for leg in legs:
                validate_measure(leg.get("value"), route, calls)
        before = _median_per_call(blocks, "before")
        after = _median_per_call(blocks, "after")
        _finite_ratio(row.get("before_ns_per_call"), before, route)
        _finite_ratio(row.get("after_ns_per_call"), after, route)
        _finite_ratio(row.get("after_before_ratio"), after / before, route)


def _paired_route(route: str) -> str:
    if route.startswith("signed_leaf/"):
        return "signed_facade/" + route.split("/", 1)[1]
    if route.startswith("signed_facade/"):
        return "signed_leaf/" + route.split("/", 1)[1]
    raise ValueError(f"invalid signed route {route}")


def _leaf_route(route: str) -> str:
    return route if route.startswith("signed_leaf/") else _paired_route(route)


def _facade_route(route: str) -> str:
    return route if route.startswith("signed_facade/") else _paired_route(route)


def _median_per_call(blocks: list[dict[str, Any]], side: str) -> float:
    values = [
        leg["value"]["elapsed_ns"] / leg["value"]["calls"]
        for block in blocks
        for leg in block["legs"]
        if leg.get("side") == side
    ]
    if not values:
        raise ValueError(f"no {side} measurements")
    return float(__import__("statistics").median(values))


def _finite_ratio(value: object, expected: float, route: str) -> None:
    if not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        raise ValueError(f"{route}: ratio/median is not finite")
    if expected < 0 or not math.isclose(float(value), expected, rel_tol=1e-9, abs_tol=1e-9):
        raise ValueError(f"{route}: ratio/median does not match ABBA measurements")


def validate_memory_summary(
    value: dict[str, Any], endpoint: dict[str, Any], fixture_names: set[str], routes: tuple[str, ...]
) -> None:
    if value.get("schema") != "litchi_signed_native_memory.v1":
        raise ValueError("memory summary schema is invalid")
    if value.get("call_counts") != [1, 8, 128]:
        raise ValueError("memory summary must contain the paired 1/8/128 call counts")
    if value.get("scope") != "whole_process_including_fixture_open_and_setup; profiler overhead included":
        raise ValueError("memory summary scope is not whole-process")
    raw_root = value.get("raw_root")
    if not isinstance(raw_root, str) or not Path(raw_root).is_absolute():
        raise ValueError("memory raw_root must be absolute")
    repository = Path(__file__).resolve().parents[3]
    raw_path = Path(raw_root).resolve()
    if raw_path == repository or repository in raw_path.parents:
        raise ValueError("memory raw artifacts must be outside the repository")
    rows = value.get("rows")
    if not isinstance(rows, list):
        raise ValueError("memory rows are missing")
    expected = {(name, route, calls) for name in fixture_names for route in routes for calls in (1, 8, 128)}
    observed: set[tuple[str, str, int]] = set()
    for index, row in enumerate(rows):
        if not isinstance(row, dict):
            raise ValueError(f"memory row {index} is not an object")
        fixture = row.get("fixture")
        route = row.get("route")
        calls = row.get("calls")
        if not isinstance(fixture, str) or not isinstance(route, str) or not isinstance(calls, int):
            raise ValueError(f"memory row {index} has invalid identity fields")
        key = (fixture, route, calls)
        if fixture not in fixture_names or route not in routes or calls not in (1, 8, 128) or key in observed:
            raise ValueError(f"memory row {index} is not an exact fixture/route/call tuple")
        observed.add(key)
        if row.get("scope") != value["scope"]:
            raise ValueError(f"memory row {index} has a different scope")
        if row.get("binary") != endpoint["binary"]:
            raise ValueError(f"memory row {index} is bound to a different binary")
        for artifact_name in ("capture", "print", "histogram", "stdout", "stderr"):
            checked_artifact(row.get(artifact_name), f"memory row {index}.{artifact_name}")
        parsed = row.get("parsed")
        if not isinstance(parsed, dict) or parsed.get("status") != "ok":
            raise ValueError(f"memory row {index} has no parsed heap totals")
        for metric in ("allocation_calls", "temporary_allocations", "peak_heap_bytes", "peak_rss_bytes", "allocated_bytes"):
            metric_value = parsed.get(metric)
            if not isinstance(metric_value, (int, float)) or not math.isfinite(float(metric_value)) or metric_value < 0:
                raise ValueError(f"memory row {index}.{metric} is missing or non-finite")
    if observed != expected:
        raise ValueError("memory summary is missing a fixture/route/call tuple")
