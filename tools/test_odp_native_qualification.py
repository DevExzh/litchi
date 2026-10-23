#!/usr/bin/env python3
"""Contract tests for the ODP native qualification runner."""

from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import unittest


MODULE_PATH = Path(__file__).with_name("odp_native_qualification.py")
SPEC = importlib.util.spec_from_file_location("odp_native_qualification", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
RUNNER = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = RUNNER
SPEC.loader.exec_module(RUNNER)


def digest(value: bytes) -> dict[str, object]:
    return {"bytes": len(value), "sha256": hashlib.sha256(value).hexdigest()}


def correctness_report(fixture: bytes) -> dict[str, object]:
    source = digest(fixture)
    append = digest(fixture + b"-append")
    return {
        "schema": RUNNER.PROBE_SCHEMA,
        "mode": "verify",
        "fixture": source,
        "checks": {
            "no_op_exact_source": True,
            "read_semantics": True,
            "append_changed": True,
            "same_owner_replay_exact": True,
            "content_matching_forward_exact": True,
            "inverse_exact_source": True,
            "changed_recipient_refused": True,
        },
        "read": {
            "slide_count": 2,
            "first_slide_present": True,
            "text_bytes": len(fixture),
            "text_sha256": source["sha256"],
        },
        "patch": {
            "is_noop": False,
            "source_content_sha256": source["sha256"],
            "target_content_sha256": append["sha256"],
        },
        "outputs": {
            "source": source,
            "no_op": source,
            "append": append,
            "same_owner_replay": append,
            "content_matching_replay": append,
            "inverse": source,
        },
    }


class OdpNativeQualificationTests(unittest.TestCase):
    def test_manifest_duplicate_keys_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            manifest = Path(directory) / "input.json"
            manifest.write_text(
                '{"schema":"litchi.odp.native-qualification.input.v1",'
                '"schema":"litchi.odp.native-qualification.input.v1"}',
                encoding="utf-8",
            )
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.load_manifest(manifest)

    def test_manifest_is_parameterized_and_preserves_requested_refs(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            fixture_bytes = b"frozen ODP fixture"
            fixture = root / "fixture.odp"
            fixture.write_bytes(fixture_bytes)
            manifest = root / "input.json"
            source_a = root / "source-a"
            source_b = root / "source-b"
            target_a = root / "external-a"
            target_b = root / "external-b"
            value = {
                "schema": RUNNER.INPUT_SCHEMA,
                "fixture": {
                    "path": str(fixture),
                    "bytes": len(fixture_bytes),
                    "sha256": hashlib.sha256(fixture_bytes).hexdigest(),
                    "provenance": "controlled fixture manifest",
                },
                "baseline": {
                    "root": str(source_a),
                    "ref": "frozen-baseline-ref",
                    "binary": str(target_a / "probe"),
                    "target_dir": str(target_a),
                },
                "candidate": {
                    "root": str(source_b),
                    "ref": "frozen-candidate-ref",
                    "binary": str(target_b / "probe"),
                    "target_dir": str(target_b),
                },
                "protocol": {
                    "cpu": 3,
                    "blocks": 2,
                    "warmups": 1,
                    "target_ms": 100,
                    "minimum_ms": 20,
                    "max_load": 0.5,
                    "max_iterations": 1000,
                    "timeout_s": 30,
                },
            }
            manifest.write_text(json.dumps(value), encoding="utf-8")
            parsed = RUNNER.load_manifest(manifest)
            self.assertEqual(parsed.baseline.requested_ref, "frozen-baseline-ref")
            self.assertEqual(parsed.candidate.requested_ref, "frozen-candidate-ref")
            self.assertEqual(parsed.fixture.provenance, "controlled fixture manifest")
            self.assertEqual(parsed.protocol.target_ns, 100_000_000)

    def test_fixture_identity_is_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            fixture = Path(directory) / "fixture.odp"
            fixture.write_bytes(b"actual")
            expected = RUNNER.Fixture(
                path=fixture,
                bytes=fixture.stat().st_size,
                sha256=hashlib.sha256(b"other").hexdigest(),
                provenance="test",
            )
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.validate_fixture(expected)

    def test_correctness_report_requires_all_checks_and_exact_digests(self) -> None:
        fixture = b"fixture"
        report = correctness_report(fixture)
        outputs = RUNNER.validate_correctness_report(
            report,
            RUNNER.Fixture(
                path=Path("fixture.odp"),
                bytes=len(fixture),
                sha256=hashlib.sha256(fixture).hexdigest(),
                provenance="test",
            ),
            "candidate",
        )
        self.assertEqual(outputs["source"], outputs["inverse"])
        broken = json.loads(json.dumps(report))
        broken["checks"]["content_matching_forward_exact"] = False
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_correctness_report(
                broken,
                RUNNER.Fixture(
                    path=Path("fixture.odp"),
                    bytes=len(fixture),
                    sha256=hashlib.sha256(fixture).hexdigest(),
                    provenance="test",
                ),
                "candidate",
            )

    def test_calibration_and_timing_reject_short_or_unbounded_reports(self) -> None:
        protocol = RUNNER.Protocol(
            cpu=0,
            blocks=1,
            warmups=0,
            target_ns=100,
            minimum_ns=50,
            max_load=1.0,
            max_iterations=10,
            timeout_s=1.0,
        )
        calibration = {
            "mode": "calibrate",
            "route": "append",
            "target_ns": 100,
            "iterations": 2,
            "elapsed_ns": 100,
            "output_bytes": 4,
            "attempts": [
                {"iterations": 1, "elapsed_ns": 40, "output_bytes": 4},
                {"iterations": 2, "elapsed_ns": 100, "output_bytes": 4},
            ],
        }
        self.assertEqual(
            RUNNER.validate_calibration(calibration, "append", protocol, "calibration"),
            2,
        )
        timed = {"mode": "timed", "route": "append", "iterations": 2, "warmups": 0, "samples": [40], "output_bytes": 4}
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_timed(timed, "append", 2, protocol, "timed")
        calibration["iterations"] = 11
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_calibration(calibration, "append", protocol, "calibration")

    def test_process_scan_classification_is_narrow_and_explicit(self) -> None:
        self.assertTrue(RUNNER.is_heavy_process("cargo", "cargo check"))
        self.assertTrue(
            RUNNER.is_heavy_process("python3", "python3 /tmp/odp_native_qualification_probe.py")
        )
        self.assertFalse(RUNNER.is_heavy_process("python3", "python3 harmless.py"))

    def test_external_artifact_guard_rejects_source_subdirectories(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory).resolve()
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER._validate_external_target(root / "target", (root,), "target")

    def test_statistics_preserve_positive_samples(self) -> None:
        summary = RUNNER._stats([10.0, 30.0, 20.0])
        self.assertEqual(summary["count"], 3)
        self.assertEqual(summary["p50_ns"], 20.0)
        self.assertAlmostEqual(summary["mean_ns"], 20.0)
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER._stats([0.0])


if __name__ == "__main__":
    unittest.main()
