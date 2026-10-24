#!/usr/bin/env python3
"""Failure-case tests for the signed native qualification validator."""
from __future__ import annotations

import copy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest


HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location(
    "validate_signed_native", HERE / "validate_signed_native.py"
)
assert SPEC and SPEC.loader
VALIDATE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(VALIDATE)


def fixture_row(index: int = 0) -> dict:
    formats = ("DOCX", "PPTX", "XLSX", "XLSB", "DOC", "PPT", "XLS")
    return {
        "name": f"fixture-{index}",
        "format": formats[index % len(formats)],
        "origin": "Opc" if index < 4 else "Cfb",
        "fixture_mode": "producer_signed" if index < 3 else "synthetic_presence",
        "signature_presence": "detected",
        "cryptographic_verification": "not_performed",
        "input_sha256": "a" * 64,
        "signed_sha256": "b" * 64,
        "target": "/body/p[1]",
        "property": "text",
    }


def fixture_manifest() -> dict:
    return {
        "schema": "litchi_signed_fixture_manifest.v1",
        "fixtures": [fixture_row(index) for index in range(7)],
    }


def route_row(route: str, **overrides: object) -> dict:
    row = {
        "name": "fixture-0",
        "format": "DOCX",
        "origin": "Opc",
        "signature_presence": "detected",
        "cryptographic_verification": "not_performed",
        "source_sha256": "a" * 64,
        "status": "ok",
        "signed_before": True,
        "signed_after": True,
        "diagnostics_changed": False,
        "patch_noop": True,
        "output_sha256": "a" * 64,
        "output_equals_source": True,
        "stream_bounded": True,
        "max_write_bytes": 8192,
        "replay_base_match": True,
        "signature_parts_absent": False,
        "semantic_changed": False,
        "signed_clone_still_refuses": True,
        "source_after_sha256": "a" * 64,
        "refusal_kind": "ProtectedDocument",
        "code": "document_protected",
        "re_marked_signed": False,
    }
    if route.endswith("/unsign_write"):
        row.update(
            {
                "signed_after": False,
                "signature_parts_absent": True,
                "semantic_changed": True,
                "output_equals_source": False,
                "output_sha256": "c" * 64,
            }
        )
    if route.endswith("/unsigned_negative"):
        row.update(
            {
                "status": "expected_refusal",
                "refusal_kind": "PatchConflict",
                "code": "invalid_input",
            }
        )
    if route.endswith("/source_freshness"):
        row.update(
            {
                "status": "expected_refusal",
                "refusal_kind": "SourceChanged",
                "code": "io_error",
            }
        )
    if route.endswith("/budget"):
        row.update(
            {
                "status": "expected_refusal",
                "refusal_kind": "LimitExceeded",
                "code": "limit_exceeded",
            }
        )
    if route.endswith("/refusal"):
        row.update({"status": "expected_refusal"})
    row.update(overrides)
    return row


def receipt(route: str, **overrides: object) -> dict:
    value = {
        "schema": VALIDATE.SCHEMA,
        "route": route,
        "capability": "available",
        "fixtures": [route_row(route)],
        "native_preflight": [
            {"format": format_name, "status": (
                "unsupported_signed_source" if format_name in {"DOC", "PPT"} else "exact_noop"
            )}
            for format_name in ("DOCX", "PPTX", "XLSX", "XLSB", "DOC", "PPT", "XLS")
        ],
    }
    value.update(overrides)
    return value


class SignedNativeValidatorTests(unittest.TestCase):
    def test_fixture_manifest_requires_all_seven_signed_formats(self) -> None:
        VALIDATE.load_fixture_manifest(fixture_manifest())
        invalid = fixture_manifest()
        invalid["fixtures"][-1]["format"] = "DOC"
        with self.assertRaises(ValueError):
            VALIDATE.load_fixture_manifest(invalid)

    def test_fixture_manifest_rejects_a_cryptographic_claim(self) -> None:
        invalid = fixture_manifest()
        invalid["fixtures"][0]["cryptographic_verification"] = "verified"
        with self.assertRaises(ValueError):
            VALIDATE.load_fixture_manifest(invalid)

    def test_signed_noop_requires_exact_bytes_and_native_noop(self) -> None:
        route = "signed_leaf/noop_commit"
        VALIDATE.validate_correctness_receipt(receipt(route), route)
        changed = receipt(route)
        changed["fixtures"][0]["output_sha256"] = "c" * 64
        with self.assertRaises(ValueError):
            VALIDATE.validate_correctness_receipt(changed, route)

    def test_signed_refusal_is_not_a_timed_success(self) -> None:
        route = "signed_facade/refusal"
        VALIDATE.validate_correctness_receipt(receipt(route), route)
        timed = copy.deepcopy(receipt(route))
        timed["fixtures"][0]["status"] = "ok"
        with self.assertRaises(ValueError):
            VALIDATE.validate_correctness_receipt(timed, route)

    def test_unsign_requires_signature_absence_and_a_real_change(self) -> None:
        route = "signed_leaf/unsign_write"
        VALIDATE.validate_correctness_receipt(receipt(route), route)
        invalid = receipt(route)
        invalid["fixtures"][0]["signature_parts_absent"] = False
        with self.assertRaises(ValueError):
            VALIDATE.validate_correctness_receipt(invalid, route)

    def test_patch_conflict_source_fence_and_budget_keep_typed_codes(self) -> None:
        for route, kind, code in (
            ("signed_leaf/unsigned_negative", "PatchConflict", "invalid_input"),
            ("signed_leaf/source_freshness", "SourceChanged", "io_error"),
            ("signed_leaf/budget", "LimitExceeded", "limit_exceeded"),
        ):
            value = receipt(route)
            value["fixtures"][0]["refusal_kind"] = kind
            value["fixtures"][0]["code"] = code
            VALIDATE.validate_correctness_receipt(value, route)
            value["fixtures"][0]["code"] = "document_protected"
            with self.assertRaises(ValueError):
                VALIDATE.validate_correctness_receipt(value, route)

    def test_final_only_capability_gap_is_allowed_only_on_before(self) -> None:
        route = "signed_leaf/noop_replay"
        value = receipt(route, capability="capability_missing")
        VALIDATE.validate_correctness_receipt(value, route, endpoint="before")
        with self.assertRaises(ValueError):
            VALIDATE.validate_correctness_receipt(value, route, endpoint="final")

    def test_calibration_and_samples_fail_closed(self) -> None:
        route = "signed_leaf/noop_write"
        calibration = {
            "schema": VALIDATE.SCHEMA,
            "route": route,
            "mode": "calibrate",
            "common_calls": 10,
            "elapsed_ns": VALIDATE.MIN_SAMPLE_NS,
            "target_ns": VALIDATE.TARGET_NS,
        }
        self.assertEqual(VALIDATE.validate_calibration(calibration, route), 10)
        sample = {
            "schema": VALIDATE.SCHEMA,
            "route": route,
            "mode": "measure",
            "calls": 10,
            "elapsed_ns": VALIDATE.MIN_SAMPLE_NS,
            "operations": 10,
        }
        VALIDATE.validate_measure(sample, route, 10)
        sample["elapsed_ns"] -= 1
        with self.assertRaises(ValueError):
            VALIDATE.validate_measure(sample, route, 10)

    def test_manifest_rejects_an_incomplete_abba(self) -> None:
        manifest = {
            "schema": "litchi_signed_native_qualification.v1",
            "rounds": 5,
            "routes": [
                {
                    "route": "signed_leaf/noop_commit",
                    "baseline": "capability_missing",
                }
            ],
        }
        VALIDATE.validate_abba(manifest)
        manifest["rounds"] = 4
        with self.assertRaises(ValueError):
            VALIDATE.validate_abba(manifest)


if __name__ == "__main__":
    unittest.main()
