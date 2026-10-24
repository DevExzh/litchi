#!/usr/bin/env python3
"""Static checks for the source-fenced DrawingML raw protocol.

These tests never invoke Cargo, Rust, a native probe, or a timing workload.
"""

from __future__ import annotations

import json
from pathlib import Path
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from native_qualification import drawingml_raw_paired_runner as paired  # noqa: E402
from native_qualification import prepare_drawingml_raw_inputs as preparation  # noqa: E402
from native_qualification import drawingml_raw_validator as validator  # noqa: E402


LANES = (
    "drawingml-chart-raw",
    "drawingml-diagram-raw",
    "drawingml-text-media-raw",
    "drawingml-theme-ink",
    "drawingml-geometry-color-formatcode2-transform",
)
OPERATIONS = (
    "entry_read",
    "read_shared_when_supported",
    "noop_commit",
    "changed_commit",
    "forward_replay",
    "inverse_replay",
    "stale_replay",
    "vec_write",
    "write_to_bounded_sink",
    "preflight_refusal",
    "malformed_refusal",
)


class DrawingMlRawProtocolTests(unittest.TestCase):
    def test_manifest_declares_all_five_source_fences_and_fixed_pair(self) -> None:
        manifest = json.loads(
            (ROOT / "tools/native_qualification/manifests/drawingml-raw.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(tuple(profile["id"] for profile in manifest["profiles"]), LANES)
        self.assertEqual(manifest["paired"]["abba_order"], list(paired.ABBA))
        self.assertEqual(manifest["paired"]["minimum_leg_ms"], 150)
        self.assertEqual(manifest["paired"]["samples_per_leg"], 20)
        self.assertTrue(manifest["paired"]["normal_and_allocator_are_separate"])
        self.assertEqual(manifest["status"], "real_adapters_ready_external_receipts_pending")
        self.assertEqual(
            manifest["baseline_workspace_lock_sha256"],
            preparation.HISTORICAL_WORKSPACE_LOCK_SHA256,
        )
        self.assertEqual(
            manifest["baseline_workspace_lock_bytes"],
            preparation.HISTORICAL_WORKSPACE_LOCK_BYTES,
        )
        self.assertEqual(manifest["candidate_workspace_lock_bytes"], 139606)
        for profile in manifest["profiles"]:
            self.assertTrue(profile["source_fence"])
            self.assertTrue(profile["record"])
            self.assertTrue(profile["approved_baseline_revisions"])
        theme = manifest["profiles"][3]
        self.assertEqual(
            theme["approved_baseline_revisions"],
            ["996651b6b19717a0192dfd4144bccd367fb2ea44"],
        )
        self.assertEqual(theme["raw_scope"], "bulk")
        route_profiles = manifest["route_profiles"]
        self.assertEqual(route_profiles[0]["raw_scope"], "family")
        self.assertEqual(
            route_profiles[0]["approved_baseline_revisions"],
            ["11728db894181ffbcdbcea3ae0008d51fdd2d8af"],
        )
        self.assertEqual(route_profiles[0]["baseline_features"], ["drawingml-owner-current"])
        self.assertEqual(
            route_profiles[0]["candidate_features"],
            [
                "drawingml-owner-current",
                "drawingml-theme-family-current",
                "drawingml-ink-limits-current",
            ],
        )
        self.assertEqual(len(manifest["owner_profiles"]), 4)
        self.assertTrue(all(profile["endpoint"] == "owner" for profile in manifest["owner_profiles"]))

    def test_materialized_profile_is_checked_by_generic_runner(self) -> None:
        master = {
            "candidate_root": ROOT,
            "candidate_revision": "a" * 40,
            "baseline_workspace_lock_sha256": preparation.HISTORICAL_WORKSPACE_LOCK_SHA256,
            "baseline_workspace_lock_bytes": preparation.HISTORICAL_WORKSPACE_LOCK_BYTES,
            "candidate_workspace_lock_sha256": "b" * 64,
            "candidate_workspace_lock_bytes": 139606,
            "output_dir": Path("/tmp/dml-protocol-output"),
            "target_root": Path("/tmp/dml-protocol-target"),
        }
        profile = {
            "id": LANES[0],
            "baseline_root": Path("/tmp/dml-protocol-baseline"),
            "baseline_revision": "2f7db7ee7fa01893e9099217fc4243687d196f63",
            "preparation_receipt": Path("/tmp/dml-protocol-preparation.json"),
            "source_fence": "test fence",
            "record": "test record",
        }
        materialized = paired._profile_manifest(master, profile, master["output_dir"])
        paired.generic._validate_drawingml_raw_workload(materialized)
        self.assertEqual(materialized["abba_order"], list(paired.generic.ABBA_ORDER))
        self.assertEqual(materialized["legs"]["before"]["lock_paths"], ["Cargo.lock", "tools/perf-baseline/Cargo.lock"])

    def test_probe_is_real_source_backed_adapter(self) -> None:
        source = (ROOT / "tools/perf-baseline/src/drawingml_raw_probe.rs").read_text(encoding="utf-8")
        for endpoint in (
            "chart::reader::read",
            "diagram::edit::set_drawing_reference",
            "text::body::reader::read",
            "svg_blip::read",
            "model3d::codec::read",
            "theme::family::read_shared",
            "theme::family::part::read",
            "ink::actions::read_profile",
            "CustomGeometryBuilder",
            "color::read",
            "transform::codec::read",
            "theme_family_mutations",
            "Package::apply",
            "VmHWM",
        ):
            self.assertIn(endpoint, source)
        self.assertIn("size_class_receipts", source)
        self.assertIn("conservation_ok", source)
        self.assertIn("supports_production_sink", source)
        self.assertIn("production_preflight", source)
        self.assertIn("source_readback_sha256", source)
        self.assertNotIn("FormattingSink", source)
        self.assertNotIn("lane adapter is not supplied until the final source freeze", source)

    def test_reviewed_adapter_hashes_bind_checked_in_sources(self) -> None:
        import hashlib

        for relative, expected in preparation.STATIC_REVIEWED_ADAPTER_SHA256.items():
            path = ROOT / relative
            if path.is_file():
                self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), expected, relative)

    def _write_unavailable_reports(self, output: Path) -> None:
        fixture = ROOT / "tools/native_qualification/fixtures/drawingml_raw_fixtures.json"
        fixture_bytes = fixture.read_bytes()
        for mode, prefix in (("normal", "dml-native"), ("allocator", "dml-allocator")):
            for index, role in enumerate(validator.ABBA, 1):
                report = {
                    "schema": validator.SCHEMA,
                    "lane": LANES[0],
                    "role": role,
                    "mode": mode,
                    "source_revision": "a" * 40 if role == "baseline" else "b" * 40,
                    "status": "unavailable",
                    "qualification_claim": False,
                    "availability_reason": "test adapter pending",
                    "fixture_bytes": len(fixture_bytes),
                    "fixture_fingerprint": "fixture",
                    "protocol": {
                        "warmups": 3,
                        "samples_per_leg": 20,
                        "blocks": 5,
                        "target_ms": 500,
                        "minimum_ms": 150,
                        "fresh_process": True,
                        "abba_order": list(validator.ABBA),
                    },
                    "allocator_metrics": None
                    if mode == "normal"
                    else {
                        "status": "unavailable",
                        "elapsed_claimable": False,
                        "direct_bytes": None,
                        "realloc_bytes": None,
                        "live_balance": None,
                        "peak_live_delta": None,
                        "process_rss_bytes": None,
                        "peak_rss_bytes": None,
                    },
                    "resource_metrics": {"process_rss_bytes": None, "peak_rss_bytes": None},
                    "rows": [
                        {"operation": operation, "status": "unavailable", "elapsed_ns": None}
                        for operation in OPERATIONS
                    ],
                }
                path = output / f"{prefix}-{'before' if role == 'baseline' else 'after'}-{index}.json"
                path.write_text(json.dumps(report), encoding="utf-8")

    def test_unavailable_reports_are_preserved_without_acceptance_claim(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-raw-validator-") as directory:
            output = Path(directory)
            self._write_unavailable_reports(output)
            arguments = type(
                "Arguments",
                (),
                {
                    "output_dir": output,
                    "lane": LANES[0],
                    "fixture": ROOT / "tools/native_qualification/fixtures/drawingml_raw_fixtures.json",
                    "baseline_revision": "a" * 40,
                    "candidate_revision": "b" * 40,
                    "source_fence": "test fence",
                },
            )()
            summary = validator.validate(arguments)
            self.assertEqual(summary["qualification_status"], "unavailable")
            self.assertFalse(summary["accepted"])
            self.assertFalse(any(report["status"] == "qualified" for report in summary["reports"]))

    def test_validator_rejects_a_short_qualified_leg(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-raw-short-") as directory:
            output = Path(directory)
            self._write_unavailable_reports(output)
            path = output / "dml-native-before-1.json"
            report = json.loads(path.read_text(encoding="utf-8"))
            report["status"] = "qualified"
            report["qualification_claim"] = True
            report["resource_metrics"] = {"process_rss_bytes": 1, "peak_rss_bytes": 1}
            report["rows"] = [
                {
                    "operation": operation,
                    "status": "qualified",
                    "elapsed_ns": 1,
                    "fidelity_checked": True,
                    "typed_refusal_checked": True,
                    "sink_writes": 1,
                }
                for operation in OPERATIONS
            ]
            path.write_text(json.dumps(report), encoding="utf-8")
            arguments = type(
                "Arguments",
                (),
                {
                    "output_dir": output,
                    "lane": LANES[0],
                    "fixture": ROOT / "tools/native_qualification/fixtures/drawingml_raw_fixtures.json",
                    "baseline_revision": "a" * 40,
                    "candidate_revision": "b" * 40,
                    "source_fence": "test fence",
                },
            )()
            with self.assertRaisesRegex(ValueError, "shorter than 150"):
                validator.validate(arguments)


if __name__ == "__main__":
    unittest.main()
