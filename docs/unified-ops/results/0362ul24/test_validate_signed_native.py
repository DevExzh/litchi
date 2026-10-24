#!/usr/bin/env python3
"""Failure-case tests for the signed native qualification validator."""
from __future__ import annotations

import copy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock


HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location(
    "validate_signed_native", HERE / "validate_signed_native.py"
)
assert SPEC and SPEC.loader
VALIDATE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(VALIDATE)
QUALIFICATION_SPEC = importlib.util.spec_from_file_location(
    "signed_native_qualification", HERE / "signed_native_qualification.py"
)
assert QUALIFICATION_SPEC and QUALIFICATION_SPEC.loader
QUALIFICATION = importlib.util.module_from_spec(QUALIFICATION_SPEC)
QUALIFICATION_SPEC.loader.exec_module(QUALIFICATION)


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
        "relative_path": f"test-data/fixture-{index}.bin",
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
        "target_changed": False,
        "signed_clone_still_refuses": True,
        "source_after_sha256": "a" * 64,
        "source_before_sha256": "a" * 64,
        "source_mutated_sha256": "d" * 64,
        "output_published": False,
        "publication_attempts": 0,
        "recipient_signed_after": True,
        "limit_kind": "operations",
        "stream_sha256": "c" * 64,
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
                "target_changed": True,
                "output_equals_source": False,
                "output_sha256": "c" * 64,
                "stream_sha256": "c" * 64,
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
                "source_after_sha256": "d" * 64,
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
    if value["capability"] == "capability_missing":
        value["fixtures"] = [
            route_row(
                route,
                status="capability_missing",
                refusal_kind="ProtectedDocument",
                refusal_code="document_protected",
            )
        ]
    return value


def final_only_manifest() -> dict:
    rows = []
    for route in VALIDATE.TIMED:
        leaf_route = route if route.startswith("signed_leaf/") else "signed_leaf/" + route.split("/", 1)[1]
        facade_route = route if route.startswith("signed_facade/") else "signed_facade/" + route.split("/", 1)[1]
        value = lambda actual: {
            "schema": VALIDATE.SCHEMA,
            "route": actual,
            "mode": "measure",
            "calls": 1,
            "elapsed_ns": VALIDATE.TARGET_NS,
            "fixture_count": 1,
            "operations": 1,
        }
        blocks = [
            {
                "legs": [
                    {"side": "leaf", "value": value(leaf_route)},
                    {"side": "facade", "value": value(facade_route)},
                    {"side": "facade", "value": value(facade_route)},
                    {"side": "leaf", "value": value(leaf_route)},
                ]
            }
            for _ in range(5)
        ]
        rows.append(
            {
                "route": route,
                "paired_route": facade_route if route.startswith("signed_leaf/") else leaf_route,
                "baseline": "leaf_facade",
                "comparison": "final_leaf_vs_facade",
                "calls": 1,
                "calibration": {
                    "leaf": {
                        "schema": VALIDATE.SCHEMA,
                        "route": leaf_route,
                        "mode": "calibrate",
                        "common_calls": 1,
                        "elapsed_ns": VALIDATE.TARGET_NS,
                        "fixture_count": 1,
                        "operations": 1,
                        "target_ns": VALIDATE.TARGET_NS,
                    },
                    "facade": {
                        "schema": VALIDATE.SCHEMA,
                        "route": facade_route,
                        "mode": "calibrate",
                        "common_calls": 1,
                        "elapsed_ns": VALIDATE.TARGET_NS,
                        "fixture_count": 1,
                        "operations": 1,
                        "target_ns": VALIDATE.TARGET_NS,
                    },
                },
                "blocks": blocks,
                "leaf_ns_per_call": float(VALIDATE.TARGET_NS),
                "facade_ns_per_call": float(VALIDATE.TARGET_NS),
                "facade_leaf_ratio": 1.0,
            }
        )
    return {
        "schema": "litchi_signed_native_qualification.v1",
        "rounds": 5,
        "target_ns": VALIDATE.TARGET_NS,
        "minimum_sample_ns": VALIDATE.MIN_SAMPLE_NS,
        "routes": rows,
    }


def validate_shape(manifest: dict) -> None:
    output = Path(tempfile.mkdtemp())
    before = {
        "binary": {"path": "before", "sha256": "a" * 64},
        "fixtures": {"path": "fixtures", "sha256": "b" * 64},
        "fixture_root": "/tmp/before-fixtures",
        "source_commit": VALIDATE.BASELINE_SOURCE_COMMIT,
    }
    after = {
        "binary": {"path": "after", "sha256": "c" * 64},
        "fixtures": {"path": "fixtures", "sha256": "b" * 64},
        "fixture_root": "/tmp/after-fixtures",
        "source_commit": VALIDATE.CANDIDATE_SOURCE_COMMIT,
    }
    pair = {"shared": {}, "baseline_capability_gaps": []}
    manifest.update(
        {
            "status": "complete",
            "output": str(output),
            "provenance": {
                "pair": pair,
                "before_manifest": {"path": "before.json"},
                "after_manifest": {"path": "after.json"},
                "before_binary": before["binary"],
                "after_binary": after["binary"],
                "before_source_commit": before["source_commit"],
                "after_source_commit": after["source_commit"],
                "cpu": 1,
            },
            "cpu": 1,
            "guard": "reviewed W8 full /proc guard before/during/after every child",
        }
    )
    with mock.patch.object(VALIDATE, "checked_artifact", return_value={"path": "x"}), \
        mock.patch.object(VALIDATE, "require_external"), \
        mock.patch.object(VALIDATE, "validate_endpoint_manifest", side_effect=[before, after]), \
        mock.patch.object(VALIDATE, "validate_pair", return_value=pair), \
        mock.patch.object(VALIDATE, "validate_child_receipt"):
        VALIDATE.validate_abba(manifest)


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

    def test_fixture_manifest_rejects_absolute_and_traversal_paths(self) -> None:
        for invalid_path in ("/tmp/signed.bin", "../signed.bin", "test-data/../signed.bin"):
            invalid = fixture_manifest()
            invalid["fixtures"][0]["relative_path"] = invalid_path
            with self.assertRaises(ValueError):
                VALIDATE.load_fixture_manifest(invalid)

    def test_fixture_manifest_rejects_3rdparty_paths(self) -> None:
        invalid = fixture_manifest()
        invalid["fixtures"][0]["relative_path"] = "test-data/3rdparty/signed.bin"
        with self.assertRaises(ValueError):
            VALIDATE.load_fixture_manifest(invalid)

    def test_shared_graph_uses_the_fixed_w8_policy(self) -> None:
        before, after, evidence = VALIDATE._reviewed_w8_projection_pair()
        self.assertEqual(before["schema"], "litchi_w8_dependency_projection.v1")
        self.assertEqual(after["schema"], "litchi_w8_dependency_projection.v1")
        self.assertEqual(evidence["registry_package_count"], 518)
        self.assertTrue(evidence["registry_identities_checksums_dependencies_equal"])

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

    def test_refusal_and_budget_must_not_publish_output(self) -> None:
        for route in ("signed_leaf/refusal", "signed_leaf/budget"):
            invalid = receipt(route)
            invalid["fixtures"][0]["output_published"] = True
            with self.assertRaises(ValueError):
                VALIDATE.validate_correctness_receipt(invalid, route)
            invalid = receipt(route)
            invalid["fixtures"][0].pop("publication_attempts")
            with self.assertRaises(ValueError):
                VALIDATE.validate_correctness_receipt(invalid, route)

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

    def test_not_applicable_is_explicit_and_route_scoped(self) -> None:
        route = "signed_leaf/budget"
        value = receipt(
            route,
            capability="not_applicable",
            fixtures=[
                route_row(
                    route,
                    status="not_applicable",
                    disposition="requires_signed_noop_admission",
                    reason="the historical endpoint cannot construct a signed no-op base",
                )
            ],
            native_preflight=[],
        )
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
            "elapsed_ns": VALIDATE.TARGET_NS,
            "target_ns": VALIDATE.TARGET_NS,
            "fixture_count": 1,
            "operations": 10,
        }
        self.assertEqual(VALIDATE.validate_calibration(calibration, route), 10)
        calibration["elapsed_ns"] = VALIDATE.MIN_SAMPLE_NS
        with self.assertRaises(ValueError):
            VALIDATE.validate_calibration(calibration, route)
        sample = {
            "schema": VALIDATE.SCHEMA,
            "route": route,
            "mode": "measure",
            "calls": 10,
            "elapsed_ns": VALIDATE.MIN_SAMPLE_NS,
            "fixture_count": 1,
            "operations": 10,
        }
        VALIDATE.validate_measure(sample, route, 10)
        sample["operations"] = 11
        with self.assertRaises(ValueError):
            VALIDATE.validate_measure(sample, route, 10)
        sample["operations"] = 10
        sample["elapsed_ns"] -= 1
        with self.assertRaises(ValueError):
            VALIDATE.validate_measure(sample, route, 10)

    def test_manifest_rejects_an_incomplete_abba(self) -> None:
        manifest = final_only_manifest()
        validate_shape(manifest)
        del manifest["routes"][-1]
        with self.assertRaises(ValueError):
            validate_shape(manifest)
        manifest = final_only_manifest()
        manifest["rounds"] = 4
        with self.assertRaises(ValueError):
            validate_shape(manifest)

    def test_final_only_requires_leaf_facade_abba_and_unique_routes(self) -> None:
        manifest = final_only_manifest()
        manifest["routes"][0]["blocks"][0]["legs"] = [
            {"side": "facade", "value": manifest["routes"][0]["blocks"][0]["legs"][0]["value"]}
        ] * 4
        with self.assertRaises(ValueError):
            validate_shape(manifest)

    def test_abba_requires_external_provenance_and_child_receipts(self) -> None:
        manifest = final_only_manifest()
        with self.assertRaises(ValueError):
            VALIDATE.validate_abba(manifest)
        manifest = final_only_manifest()
        manifest["routes"][1]["route"] = manifest["routes"][0]["route"]
        with self.assertRaises(ValueError):
            validate_shape(manifest)

    def test_pair_rejects_same_source_and_binary_identity(self) -> None:
        def endpoint(commit: str, binary_path: str, binary_hash: str, tree: str) -> dict:
            shared = {"sha256": "a" * 64}
            return {
                "source_commit": commit,
                "binary": {"path": binary_path, "sha256": binary_hash},
                "source_archive": {"sha256": "b" * 64},
                "source_identity": {"tree_sha256": tree},
                "lock": shared,
                "probe": shared,
                "fixtures": shared,
                "dependency_projection": shared,
                "capabilities": {route: "available" for route in VALIDATE.ROUTES},
            }
        before = endpoint(VALIDATE.BASELINE_SOURCE_COMMIT, "/tmp/before", "c" * 64, "d" * 64)
        after = endpoint(VALIDATE.CANDIDATE_SOURCE_COMMIT, "/tmp/after", "c" * 64, "e" * 64)
        with self.assertRaises(ValueError):
            VALIDATE.validate_pair(before, after)

    def test_runner_reserves_wildcard_for_route_listing(self) -> None:
        with self.assertRaises(ValueError):
            QUALIFICATION.child(
                Path(tempfile.mkdtemp()),
                "invalid",
                Path("/does/not/exist"),
                {"fixture_root": "/tmp"},
                "*",
                "measure",
                0,
            )

    def test_runner_has_process_caps_and_final_pairing(self) -> None:
        self.assertGreater(QUALIFICATION.CHILD_TIMEOUT_S, 0)
        self.assertGreater(QUALIFICATION.CHILD_OUTPUT_LIMIT, 0)
        self.assertGreater(QUALIFICATION.CHILD_RSS_LIMIT, 0)
        self.assertEqual(
            QUALIFICATION.paired_route("signed_leaf/noop_write"),
            "signed_facade/noop_write",
        )

    def test_memory_summary_rejects_unpaired_call_counts(self) -> None:
        with self.assertRaises(ValueError):
            VALIDATE.validate_memory_summary(
                {
                    "schema": "litchi_signed_native_memory.v1",
                    "call_counts": [1, 8],
                },
                {},
                set(),
                (),
            )

    def test_memory_summary_requires_endpoint_and_tool_provenance(self) -> None:
        with self.assertRaises(ValueError):
            VALIDATE.validate_memory_summary(
                {
                    "schema": "litchi_signed_native_memory.v1",
                    "status": "complete",
                    "call_counts": [1, 8, 128],
                    "scope": "whole_process_including_fixture_open_and_setup; profiler overhead included",
                    "raw_root": "/tmp/signed-memory",
                    "rows": [],
                },
                {},
                set(),
                (),
            )


if __name__ == "__main__":
    unittest.main()
