#!/usr/bin/env python3
"""Static checks for the source-fenced DrawingML raw protocol.

These tests never invoke Cargo, Rust, a native probe, or a timing workload.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
import tempfile
import tomllib
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
        self.assertEqual(manifest["plan_path"], paired.PLAN_PATH)
        self.assertEqual(manifest["plan_sha256"], paired.PLAN_SHA256)
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
                "drawingml-theme-family-owner-limits-current",
                "drawingml-ink-limits-current",
                "drawingml-transform-limits-current",
                "drawingml-raw-current",
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
        self.assertEqual(
            materialized["drawingml_baseline_features"],
            [paired.DRAWINGML_PROBE_FEATURE],
        )
        for leg in ("before", "after"):
            argv = materialized["legs"][leg]["build"][0]["argv"]
            feature_index = argv.index("--features")
            self.assertEqual(argv[feature_index + 1], paired.DRAWINGML_PROBE_FEATURE)
        self.assertEqual(materialized["legs"]["before"]["lock_paths"], ["Cargo.lock", "tools/perf-baseline/Cargo.lock"])
        self.assertIn("fixture_bytes", materialized["controls"][0]["compare"])

    def test_drawingml_probe_isolated_from_no_default_feature_lanes(self) -> None:
        cargo = tomllib.loads((ROOT / "tools/perf-baseline/Cargo.toml").read_text(encoding="utf-8"))
        default = set(cargo["features"]["default"])
        self.assertIn(paired.DRAWINGML_PROBE_FEATURE, cargo["features"])
        self.assertIn(paired.DRAWINGML_PROBE_FEATURE, default)
        self.assertIn("rtf-current-api", default)
        for feature in (
            "drawingml-raw-current",
            "drawingml-owner-current",
            "drawingml-theme-family-current",
            "drawingml-theme-family-owner-limits-current",
            "drawingml-ink-limits-current",
            "drawingml-transform-limits-current",
        ):
            self.assertIn(feature, default)
        lib = (ROOT / "tools/perf-baseline/src/lib.rs").read_text(encoding="utf-8")
        self.assertIn('#[cfg(feature = "drawingml-probe")]\npub mod drawingml_raw_probe;', lib)
        manifest = (ROOT / "tools/perf-baseline/Cargo.toml").read_text(encoding="utf-8")
        self.assertIn('required-features = ["drawingml-probe"]', manifest)

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
        self.assertIn("PreparedValue", source)
        self.assertIn("semantic_changed_differs", source)
        self.assertIn("generated_fixture_parts", source)
        self.assertIn("source_bytes", source)
        self.assertIn("changed_bytes", source)
        self.assertIn("malformed_bytes", source)
        self.assertIn("--preflight", source)
        self.assertIn("output_limit_attempt", source)
        self.assertIn("partial_output_allowed", source)
        self.assertIn("changed_output_exact", source)
        self.assertIn("owner_native_commit", source)
        self.assertIn("owner_native_replay", source)
        self.assertIn("owner_native_commit_output", source)
        self.assertIn("owner_native_replay_output", source)
        self.assertIn("owner_native_stale_refusal", source)
        self.assertIn("owner_operation_path", source)
        self.assertIn("result_serialized_outside_clock", source)
        self.assertIn("Prepared::edit", source)
        self.assertIn("OwnerSemanticReceipt", source)
        self.assertIn("source_edit_serialize", source)
        self.assertIn('"string_output"', source)
        self.assertIn('"vec_output"', source)
        self.assertIn('"scanner_helper_parts"', source)
        self.assertIn('"scanner_helper_scope"', source)
        self.assertIn("read(xml,scope,limits)", source)
        self.assertIn("write_contextual_to", source)
        self.assertIn("read_svg_contextual_value", source)
        self.assertIn("read_svg_standalone_value", source)
        self.assertIn("standalone_svg_roundtrip", source)
        self.assertIn("contextual_source", source)
        self.assertIn('"contextual_read_write"', source)
        self.assertIn("standalone_route_checked", source)
        self.assertIn("source_raw_sha256", source)
        self.assertIn("splice_oracle", source)
        self.assertIn("source_identity_matches", source)
        self.assertIn("parse_geometry_native", source)
        self.assertIn("read_write_route", source)
        self.assertIn("requested_identity_matches", source)
        self.assertIn("shared_prepared", source)
        self.assertIn("prepared_operations", source)
        self.assertIn("timed_entry_read", source)
        self.assertNotIn("source_readback_sha256", source)
        self.assertNotIn("changed_readback_sha256", source)
        self.assertNotIn("FormattingSink", source)
        self.assertNotIn("lane adapter is not supplied until the final source freeze", source)

        paired_source = (
            ROOT / "tools/native_qualification/drawingml_raw_paired_runner.py"
        ).read_text(encoding="utf-8")
        self.assertIn("drawingml-fixture-preflight", paired_source)
        self.assertIn("generated_fixture_parts", paired_source)
        generic_source = (ROOT / "tools/native_qualification_runner.py").read_text(
            encoding="utf-8"
        )
        self.assertIn('"event_kind": event_kind', generic_source)
        self.assertIn('"measured": measured', generic_source)
        self.assertIn("self._run_controls(identities)", generic_source)
        record = (
            ROOT / "docs/unified-ops/0362dvib-drawingml-native-raw-repair.md"
        ).read_text(encoding="utf-8")
        self.assertIn("contextual", record)
        self.assertIn("standalone", record)

    def test_owner_semantic_proof_requires_mutation_and_remainder_identity(self) -> None:
        hashes = {
            "source_semantic_sha256": "a" * 64,
            "noop_semantic_sha256": "a" * 64,
            "changed_semantic_sha256": "b" * 64,
            "forward_semantic_sha256": "b" * 64,
            "inverse_semantic_sha256": "a" * 64,
            "source_rest_semantic_sha256": "c" * 64,
            "changed_rest_semantic_sha256": "c" * 64,
            "expected_mutation": "chart.title=DrawingML qualification mutation",
            "mutation_verified": True,
            "unchanged_remainder": True,
        }
        self.assertEqual(
            validator._validate_owner_semantic_proof(
                {"semantic": hashes}, Path("owner-proof.json")
            ),
            hashes,
        )
        invalid = dict(hashes, mutation_verified=False)
        with self.assertRaisesRegex(ValueError, "mutation/remainder proof"):
            validator._validate_owner_semantic_proof(
                {"semantic": invalid}, Path("owner-proof.json")
            )

    def test_generated_fixture_identity_binds_all_three_byte_roles(self) -> None:
        parts = [
            {
                "part_id": "chart",
                "source_sha256": "a" * 64,
                "source_bytes": 1,
                "changed_sha256": "b" * 64,
                "changed_bytes": 2,
                "malformed_sha256": "c" * 64,
                "malformed_bytes": 3,
            }
        ]
        identity = validator._generated_fixture_identity(parts)
        self.assertEqual(identity, validator._generated_fixture_identity(parts))
        changed = [dict(parts[0], malformed_sha256="d" * 64)]
        self.assertNotEqual(identity, validator._generated_fixture_identity(changed))
        resized = [dict(parts[0], malformed_bytes=4)]
        self.assertNotEqual(identity, validator._generated_fixture_identity(resized))
    def _write_external_resource_capture(self, output: Path) -> tuple[dict, dict]:
        event_id = validator._external_event_id(1, "baseline", "normal")
        argv = ["/external/drawingml-raw-native", "--lane", LANES[0]]
        source_root = "/external/drawingml-source"
        source_revision = "a" * 40
        source_identity = "c" * 64
        resource_path = output / validator.EXTERNAL_RESOURCE_FILENAME
        record = {
            "schema": validator.EXTERNAL_RESOURCE_SCHEMA,
            "event_id": event_id,
            "leg": "baseline",
            "status": "captured",
            "reason": None,
            "pid": 4242,
            "start_identity": {
                "pid": 4242,
                "boot_id": "boot-id",
                "starttime_ticks": 1234,
            },
            "argv_sha256": validator._argv_sha256(argv),
            "cwd": "/external/drawingml-output",
            "source_ref": {
                "source_root": source_root,
                "source_revision": source_revision,
                "source_identity_sha256": source_identity,
                "leg": "baseline",
            },
            "returncode": 0,
            "sample_count": 2,
            "process_rss_bytes": 2048,
            "peak_rss_bytes": 3072,
            "peak_vm_hwm_bytes": 4096,
        }
        record_line = json.dumps(record, sort_keys=True, separators=(",", ":")) + "\n"
        resource_path.write_text(record_line, encoding="utf-8")
        event = {
            "event_id": event_id,
            "leg": "baseline",
            "argv": argv,
            "cwd": record["cwd"],
            "source_root": source_root,
            "source_head": source_revision,
            "source_identity_sha256": source_identity,
            "returncode": 0,
            "runner_error": None,
            "process_observation": {
                "schema": validator.EXTERNAL_RESOURCE_SCHEMA,
                "path": str(resource_path),
                "event_id": event_id,
                "record_sha256": hashlib.sha256(record_line.encode()).hexdigest(),
                "record_bytes": len(record_line.encode()),
            },
        }
        (output / "events-abba.jsonl").write_text(
            json.dumps(event, sort_keys=True, separators=(",", ":")) + "\n",
            encoding="utf-8",
        )
        report = {
            "source_revision": source_revision,
            "resource_metrics": {
                "pid": 4242,
                "command_sha256": validator._argv_sha256(argv),
                "source_revision": source_revision,
                "process_rss_bytes": 2048,
                "peak_rss_bytes": 3072,
                "whole_process_peak_rss_bytes": 4096,
            },
        }
        return report, event

    def test_external_resource_capture_binds_mutations_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-raw-resource-") as directory:
            output = Path(directory)
            report, _ = self._write_external_resource_capture(output)
            path = output / "dml-native-before-1.json"
            capture = validator._load_external_resource_capture(output)
            observed = validator._validate_external_resource(
                capture,
                report,
                path,
                index=1,
                role="baseline",
                mode="normal",
            )
            self.assertEqual(observed["pid"], 4242)

            def mutate_argv(value: dict) -> None:
                value["records"][next(iter(value["records"]))]["argv_sha256"] = "0" * 64

            def mutate_source(value: dict) -> None:
                value["events"][next(iter(value["events"]))]["source_head"] = "b" * 40

            def mutate_pointer(value: dict) -> None:
                value["events"][next(iter(value["events"]))]["process_observation"][
                    "record_sha256"
                ] = "0" * 64

            mutations = (
                ("argv", mutate_argv, "external argv hash differs"),
                ("source", mutate_source, "external source/event binding differs"),
                ("pointer", mutate_pointer, "external observation pointer differs"),
            )
            for _, mutate, message in mutations:
                mutated = validator._load_external_resource_capture(output)
                mutate(mutated)
                with self.assertRaisesRegex(ValueError, message):
                    validator._validate_external_resource(
                        mutated,
                        report,
                        path,
                        index=1,
                        role="baseline",
                        mode="normal",
                    )

    def test_validator_control_handoff_does_not_change_sidecar_receipt(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-raw-handoff-") as directory:
            output = Path(directory)
            observer = paired.ExternalResourceObserver(output)

            class ProbeToken:
                def finish(self, returncode: int | None) -> dict:
                    return {
                        "schema": paired.EXTERNAL_RESOURCE_SCHEMA,
                        "event_id": "abba-01-baseline-raw-native",
                        "leg": "baseline",
                        "status": "captured",
                        "returncode": returncode,
                    }

            observer.process_finished(ProbeToken(), returncode=0)
            before = observer.receipt()
            for event_id in (
                "validator-drawingml-raw-identity",
                "control-drawingml-fixture-preflight-before",
                "control-drawingml-fixture-preflight-after",
            ):
                self.assertIsNone(observer.process_started(event_id=event_id))
            self.assertIsNone(
                observer.process_started(
                    event_id="unlisted-control",
                    event_kind="control",
                    measured=False,
                )
            )
            self.assertIsNone(
                observer.process_started(
                    event_id="untimed-measurement",
                    event_kind="measurement",
                    measured=False,
                )
            )
            after = observer.receipt()
            self.assertEqual(after, before)
            self.assertEqual(after["records"], 1)

    def test_disappeared_child_is_recorded_unavailable_without_claim(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-raw-disappeared-child-") as directory:
            output = Path(directory)
            observer = paired.ExternalResourceObserver(output)
            token = observer.process_started(
                event_id="abba-01-baseline-raw-native",
                leg="baseline",
                argv=["/external/drawingml-raw-native"],
                cwd=output,
                identity={
                    "source_root": "/external/drawingml-source",
                    "head": "a" * 40,
                    "identity_sha256": "b" * 64,
                },
                process=type("FastProcess", (), {"pid": 999999999})(),
            )
            self.assertIsNotNone(token)
            observer.process_finished(token, returncode=0)
            record = json.loads(
                (output / paired.EXTERNAL_RESOURCE_FILENAME).read_text(encoding="utf-8")
            )
            self.assertEqual(record["status"], "unavailable")
            self.assertEqual(record["sample_count"], 0)
            self.assertIn("disappeared", record["reason"])

    def test_reviewed_adapter_hashes_bind_checked_in_sources(self) -> None:
        import hashlib

        for relative, expected in preparation.STATIC_REVIEWED_ADAPTER_SHA256.items():
            if relative == preparation.CONSUMER_API_ADAPTER_PATH:
                source, frozen = preparation._frozen_consumer_api_source(ROOT)
                self.assertEqual(frozen["ref"], preparation.FROZEN_CONSUMER_API_SOURCE_REF)
                self.assertEqual(frozen["blob_sha1"], preparation.FROZEN_CONSUMER_API_SOURCE_BLOB_SHA1)
                self.assertEqual(hashlib.sha256(source).hexdigest(), expected, relative)
                continue
            path = ROOT / relative
            if path.is_file():
                self.assertEqual(
                    hashlib.sha256(path.read_bytes()).hexdigest(), expected, relative
                )

    def test_adapter_locks_bind_each_source_role(self) -> None:
        baseline = preparation._reviewed_adapter_hashes(
            "a" * 64,
            preparation.HISTORICAL_WORKSPACE_LOCK_SHA256,
            "baseline",
        )
        candidate = preparation._reviewed_adapter_hashes("b" * 64, "c" * 64)
        self.assertNotIn("Cargo.lock", preparation.STATIC_REVIEWED_ADAPTER_SHA256)
        self.assertEqual(baseline["Cargo.lock"], preparation.HISTORICAL_WORKSPACE_LOCK_SHA256)
        self.assertEqual(candidate["Cargo.lock"], "c" * 64)
        self.assertEqual(baseline["tools/perf-baseline/Cargo.lock"], "a" * 64)
        self.assertEqual(candidate["tools/perf-baseline/Cargo.lock"], "b" * 64)
        self.assertEqual(
            baseline[preparation.CONSUMER_API_ADAPTER_PATH],
            preparation.CONSUMER_API_HISTORICAL_SHA256,
        )
        self.assertEqual(
            candidate[preparation.CONSUMER_API_ADAPTER_PATH],
            preparation.CONSUMER_API_CURRENT_SHA256,
        )

    def test_historical_consumer_adapter_changes_only_six_spellings(self) -> None:
        source, frozen = preparation._frozen_consumer_api_source(ROOT)
        self.assertEqual(frozen["ref"], preparation.FROZEN_CONSUMER_API_SOURCE_REF)
        self.assertEqual(frozen["path"], preparation.FROZEN_CONSUMER_API_SOURCE_PATH)
        self.assertEqual(
            hashlib.sha256(source).hexdigest(), preparation.CONSUMER_API_CURRENT_SHA256
        )
        adapted = preparation._adapt_baseline_consumer_adapter(
            preparation.CONSUMER_API_ADAPTER_PATH, source
        )
        expected = source
        for current, historical in preparation.CONSUMER_API_REPLACEMENTS:
            self.assertEqual(source.count(current), 1)
            self.assertEqual(source.count(historical), 0)
            expected = expected.replace(current, historical)
        self.assertEqual(adapted, expected)
        self.assertEqual(
            hashlib.sha256(adapted).hexdigest(),
            preparation.CONSUMER_API_HISTORICAL_SHA256,
        )
        self.assertEqual(preparation.CONSUMER_API_TRANSFORM_COUNT, 6)
        self.assertEqual(preparation.CONSUMER_API_TRANSFORM_BYTES, 7)
        self.assertEqual(
            preparation._consumer_api_adapter_receipt(),
            {
                "path": preparation.CONSUMER_API_ADAPTER_PATH,
                "candidate_sha256": preparation.CONSUMER_API_CURRENT_SHA256,
                "baseline_sha256": preparation.CONSUMER_API_HISTORICAL_SHA256,
                "replacement_count": 6,
                "changed_bytes": 7,
                "replacements": [
                    {
                        "candidate": current.decode("ascii"),
                        "baseline": historical.decode("ascii"),
                        "count": 1,
                    }
                    for current, historical in preparation.CONSUMER_API_REPLACEMENTS
                ],
            },
        )
        with self.assertRaises(preparation.PreparationError):
            preparation._adapt_baseline_consumer_adapter(
                preparation.CONSUMER_API_ADAPTER_PATH,
                adapted,
            )

    def test_v1_rejects_a_nonfrozen_current_candidate_source(self) -> None:
        source, _ = preparation._frozen_consumer_api_source(ROOT)
        with self.assertRaises(preparation.PreparationError):
            preparation._adapt_baseline_consumer_adapter(
                preparation.CONSUMER_API_ADAPTER_PATH,
                source + b"\n",
            )

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
