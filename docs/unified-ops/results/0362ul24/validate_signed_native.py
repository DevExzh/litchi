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

ORIGINS = {"Opc", "Cfb"}
CRYPTO_STATUS = "not_performed"


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
        "lock",
        "probe",
        "fixtures",
        "dependency_projection",
    ):
        checked_artifact(value.get(name), f"{label}.{name}")
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
        if status not in {"available", "capability_missing", "expected_refusal"}:
            raise ValueError(f"{label}.capabilities[{route}] has invalid status")
        if status == "capability_missing" and route not in FINAL_ONLY and route not in CORRECTNESS_ONLY:
            raise ValueError(f"{label}.capabilities[{route}] cannot be missing")
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
    shared = shared_identity(before, after)
    for route in ROUTES:
        final_status = after["capabilities"][route]
        before_status = before["capabilities"][route]
        if final_status == "capability_missing":
            raise ValueError(f"final endpoint lacks required route {route}")
        if route in CORRECTNESS_ONLY and route.endswith("/refusal") and final_status != "expected_refusal":
            raise ValueError(f"final refusal route {route} is not marked expected_refusal")
        if route in FINAL_ONLY and before_status == "available":
            # Availability alone is not enough; the actual preflight result
            # must still prove the same semantic operation below.
            continue
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
    if capability not in {"available", "capability_missing", "expected_refusal"}:
        raise ValueError(f"{route}: invalid capability status")
    if capability == "capability_missing":
        if endpoint == "final" or route not in FINAL_ONLY:
            raise ValueError(f"{route}: required capability missing on {endpoint}")
        return {"route": route, "status": "capability_missing", "endpoint": endpoint}
    native = value.get("native_preflight")
    if not isinstance(native, list) or not native:
        raise ValueError(f"{route}: direct native preflight is missing")
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
        if format_name in {"DOC", "PPT"} and status != "unsupported_signed_source":
            raise ValueError(f"{route}: {format_name} signed source should refuse direct source edit")
        if format_name not in {"DOC", "PPT"} and status != "exact_noop":
            raise ValueError(f"{route}: {format_name} direct native no-op was not proven")
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
            continue
        if route.endswith("/source_freshness"):
            if (
                row.get("status") != "expected_refusal"
                or row.get("refusal_kind") != "SourceChanged"
                or row.get("code") != "io_error"
            ):
                raise ValueError(f"{route}: source fence is not SourceChanged/io_error")
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
        elif route.endswith("/noop_write"):
            if (
                row.get("signed_after") is not True
                or row.get("output_sha256") != row.get("source_sha256")
            ):
                raise ValueError(f"{route}: signed publication changed bytes/state")
            if (
                row.get("stream_bounded") is not True
                or not isinstance(row.get("max_write_bytes"), int)
                or row["max_write_bytes"] <= 0
            ):
                raise ValueError(f"{route}: publication is not bounded")
        elif route.endswith("/noop_replay"):
            if row.get("patch_noop") is not True or row.get("replay_base_match") is not True:
                raise ValueError(f"{route}: replay lacks content-bound no-op proof")
            if (
                row.get("signed_after") is not True
                or row.get("output_sha256") != row.get("source_sha256")
            ):
                raise ValueError(f"{route}: replay changed signed publication")
        elif route.endswith("/unsign_write"):
            if row.get("signed_after") is not False or row.get("signature_parts_absent") is not True:
                raise ValueError(f"{route}: unsign output retains signature state/parts")
            if row.get("semantic_changed") is not True or row.get("output_equals_source") is not False:
                raise ValueError(f"{route}: unsign route did not publish a real change")
            checked_hash(row.get("output_sha256"), f"{route}: output")
            if row.get("signed_clone_still_refuses") is not True:
                raise ValueError(f"{route}: signed clone lost protection")
    return {"route": route, "status": "ok", "endpoint": endpoint, "fixtures": len(rows)}


def validate_correctness_receipt(
    value: dict[str, Any], route: str, *, endpoint: str = "final"
) -> dict[str, Any]:
    """Public wrapper used by tests and the guarded runner."""
    return validate_route(value, route, endpoint=endpoint)


def validate_measure(value: dict[str, Any], route: str, calls: int) -> None:
    if value.get("schema") != SCHEMA or value.get("route") != route:
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
    if not isinstance(elapsed, int) or elapsed < MIN_SAMPLE_NS:
        raise ValueError(f"{route}: calibration did not reach the sample floor")
    if value.get("target_ns") != TARGET_NS:
        raise ValueError(f"{route}: calibration target differs from protocol")
    return calls


def validate_abba(manifest: dict[str, Any]) -> None:
    if manifest.get("schema") != "litchi_signed_native_qualification.v1":
        raise ValueError("qualification manifest schema is invalid")
    rounds = manifest.get("rounds")
    if not isinstance(rounds, int) or rounds < MIN_ABBA_ROUNDS:
        raise ValueError("qualification needs at least five ABBA rounds")
    rows = manifest.get("routes")
    if not isinstance(rows, list):
        raise ValueError("qualification route rows are missing")
    for row in rows:
        if not isinstance(row, dict) or row.get("route") not in TIMED:
            raise ValueError("qualification contains an unknown timed route")
        if row.get("baseline") == "capability_missing":
            if row["route"] not in FINAL_ONLY:
                raise ValueError(
                    f"{row['route']}: only final-only routes may lack a baseline"
                )
            continue
        calls = row.get("calls")
        if not isinstance(calls, int) or calls <= 0 or calls > MAX_CALLS:
            raise ValueError(f"{row.get('route')}: invalid shared call count")
        blocks = row.get("blocks")
        if not isinstance(blocks, list) or len(blocks) < rounds:
            raise ValueError(f"{row.get('route')}: incomplete ABBA blocks")
        for block in blocks:
            legs = block.get("legs") if isinstance(block, dict) else None
            if (
                not isinstance(legs, list)
                or [leg.get("side") for leg in legs]
                != ["before", "after", "after", "before"]
            ):
                raise ValueError(f"{row.get('route')}: ABBA order is invalid")
            for leg in legs:
                validate_measure(leg.get("value"), row["route"], calls)
