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
from unittest import mock
import zipfile


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
    recipient_archive = digest(fixture + b"-recipient")
    source_member = digest(b"source-member")
    recipient_member = digest(b"changed-member")
    return {
        "schema": RUNNER.PROBE_SCHEMA,
        "mode": "verify",
        "fixture": {**source, "content": source},
        "checks": {
            "no_op_exact_source": True,
            "read_semantics": True,
            "append_changed": True,
            "patch_source_matches_fixture": True,
            "append_output_changed": True,
            "scoped_recipient_content_matches": True,
            "scoped_recipient_member_changed": True,
            "scoped_forward_content_matches_patch": True,
            "scoped_forward_preserves_unrelated_member": True,
            "scoped_inverse_preserves_unrelated_member": True,
            "scoped_inverse_restores_content_scope": True,
            "same_owner_replay_exact": True,
            "content_matching_forward_exact": True,
            "whole_owner_inverse_exact_source": True,
            "inverse_content_exact_source": True,
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
        "scoped_replay": {
            "provenance": "deterministic generated recipient",
            "member_name": "Qualification/unrelated.bin",
            "recipient_archive": recipient_archive,
            "source_member": source_member,
            "recipient_member": recipient_member,
            "forward_member": recipient_member,
            "inverse_member": recipient_member,
            "recipient_content": source,
            "forward_content": append,
            "inverse_content": source,
        },
        "outputs": {
            "source": source,
            "no_op": source,
            "append": append,
            "same_owner_replay": append,
            "content_matching_replay": append,
            "whole_owner_inverse": source,
            "inverse": source,
            "inverse_content": source,
        },
    }


def legacy_correctness_report(fixture: bytes) -> dict[str, object]:
    source = digest(fixture)
    append = digest(fixture + b"-append")
    return {
        "schema": RUNNER.LEGACY_PROBE_SCHEMA,
        "mode": "verify",
        "capabilities": {
            "patch_replay": False,
            "inverse_replay": False,
            "scoped_replay": False,
            "timed_routes": list(RUNNER.ROUTES),
            "replay_scope": "independent_plan_not_patch",
        },
        "fixture": {**source, "content": source},
        "checks": {
            "no_op_exact_source": True,
            "read_semantics": True,
            "append_changed": True,
            "patch_source_matches_fixture": True,
            "append_output_changed": True,
            "independent_plan_repeat_exact": True,
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
            "independent_append": append,
        },
    }


class OdpNativeQualificationTests(unittest.TestCase):
    def test_correctness_only_mode_is_mutually_exclusive_with_plan(self) -> None:
        parsed = RUNNER.parse_args(
            ["--manifest", "input.json", "--out-dir", "run", "--correctness-only"]
        )
        self.assertTrue(parsed.correctness_only)
        with self.assertRaises(SystemExit):
            RUNNER.parse_args(
                [
                    "--manifest",
                    "input.json",
                    "--out-dir",
                    "run",
                    "--plan",
                    "--correctness-only",
                ]
            )

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
                    "profile": RUNNER.LEGACY_PROFILE,
                    "binary": str(target_a / "probe"),
                    "target_dir": str(target_a),
                },
                "candidate": {
                    "root": str(source_b),
                    "ref": "frozen-candidate-ref",
                    "profile": RUNNER.CURRENT_PROFILE,
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
            self.assertEqual(parsed.baseline.profile, RUNNER.LEGACY_PROFILE)
            self.assertEqual(parsed.candidate.profile, RUNNER.CURRENT_PROFILE)
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

    def test_fixture_identity_records_archive_and_content_hashes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            fixture = Path(directory) / "fixture.odp"
            content = b"<office:document-content/>"
            with zipfile.ZipFile(fixture, "w") as archive:
                archive.writestr("content.xml", content)
            expected = RUNNER.Fixture(
                path=fixture,
                bytes=fixture.stat().st_size,
                sha256=hashlib.sha256(fixture.read_bytes()).hexdigest(),
                provenance="test",
            )
            identity = RUNNER.validate_fixture(expected)
            self.assertEqual(identity["content"], digest(content))

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
            {"content": digest(fixture)},
            "candidate",
        )
        self.assertEqual(outputs["source"], outputs["inverse_content"])
        self.assertEqual(outputs["source"], outputs["whole_owner_inverse"])
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
                {"content": digest(fixture)},
                "candidate",
            )
        broken_whole_owner = json.loads(json.dumps(report))
        broken_whole_owner["checks"]["whole_owner_inverse_exact_source"] = False
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_correctness_report(
                broken_whole_owner,
                RUNNER.Fixture(
                    path=Path("fixture.odp"),
                    bytes=len(fixture),
                    sha256=hashlib.sha256(fixture).hexdigest(),
                    provenance="test",
                ),
                {"content": digest(fixture)},
                "candidate",
            )
        broken_scope = json.loads(json.dumps(report))
        broken_scope["scoped_replay"]["forward_member"] = broken_scope["scoped_replay"][
            "source_member"
        ]
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_correctness_report(
                broken_scope,
                RUNNER.Fixture(
                    path=Path("fixture.odp"),
                    bytes=len(fixture),
                    sha256=hashlib.sha256(fixture).hexdigest(),
                    provenance="test",
                ),
                {"content": digest(fixture)},
                "candidate",
            )

    def test_legacy_report_has_separate_capability_boundary_and_no_replay_timing(self) -> None:
        fixture = b"fixture"
        report = legacy_correctness_report(fixture)
        record = RUNNER.validate_legacy_correctness_report(
            report,
            RUNNER.Fixture(
                path=Path("fixture.odp"),
                bytes=len(fixture),
                sha256=hashlib.sha256(fixture).hexdigest(),
                provenance="test",
            ),
            {"content": digest(fixture)},
            "baseline",
        )
        self.assertEqual(record["profile"], RUNNER.LEGACY_PROFILE)
        self.assertFalse(record["capabilities"]["patch_replay"])
        self.assertEqual(record["capabilities"]["timed_routes"], list(RUNNER.ROUTES))
        self.assertNotIn("replay", RUNNER.ROUTES)
        broken = json.loads(json.dumps(report))
        broken["capabilities"]["patch_replay"] = True
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_legacy_correctness_report(
                broken,
                RUNNER.Fixture(
                    path=Path("fixture.odp"),
                    bytes=len(fixture),
                    sha256=hashlib.sha256(fixture).hexdigest(),
                    provenance="test",
                ),
                {"content": digest(fixture)},
                "baseline",
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
            "proof": {"kind": "package", "output": digest(b"1234")},
            "attempts": [
                {"iterations": 1, "elapsed_ns": 40, "output_bytes": 4},
                {"iterations": 2, "elapsed_ns": 100, "output_bytes": 4},
            ],
        }
        self.assertEqual(
            RUNNER.validate_calibration(
                calibration,
                "append",
                protocol,
                calibration["proof"],
                "calibration",
            ),
            2,
        )
        timed = {
            "mode": "timed",
            "route": "append",
            "iterations": 2,
            "warmups": 0,
            "samples": [40],
            "output_bytes": 4,
            "proof": calibration["proof"],
        }
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_timed(
                timed, "append", 2, protocol, calibration["proof"], "timed"
            )
        timed["samples"] = [100]
        timed["proof"] = {"kind": "package", "output": digest(b"5678")}
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_timed(
                timed, "append", 2, protocol, calibration["proof"], "timed"
            )
        calibration["iterations"] = 11
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_calibration(
                calibration,
                "append",
                protocol,
                calibration["proof"],
                "calibration",
            )

    def test_resolve_endpoint_rejects_clean_head_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "source"
            target = Path(directory) / "target"
            root.mkdir()
            target.mkdir()
            binary = target / "probe"
            binary.write_bytes(b"probe")
            binary.chmod(0o755)
            endpoint = RUNNER.Endpoint(
                name="baseline",
                root=root,
                requested_ref="frozen-ref",
                binary=binary,
                target_dir=target,
                profile=RUNNER.LEGACY_PROFILE,
            )
            with mock.patch.object(
                RUNNER,
                "_run_git",
                side_effect=["", "a" * 40, "b" * 40],
            ), self.assertRaises(RUNNER.RunFailure):
                RUNNER.resolve_endpoint(endpoint, (root.parent,))

    def test_postrun_binary_identity_change_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "source"
            target = Path(directory) / "target"
            root.mkdir()
            target.mkdir()
            binary = target / "probe"
            binary.write_bytes(b"before")
            binary.chmod(0o755)
            endpoint = RUNNER.Endpoint(
                "baseline", root, "frozen-ref", binary, target, RUNNER.LEGACY_PROFILE
            )
            resolved = RUNNER.ResolvedEndpoint(
                endpoint,
                "a" * 40,
                binary.stat().st_size,
                RUNNER.sha256_file(binary),
            )
            binary.write_bytes(b"after")
            with mock.patch.object(
                RUNNER, "_run_git", side_effect=["", "a" * 40, "a" * 40]
            ), self.assertRaises(RUNNER.RunFailure):
                RUNNER._assert_endpoint_unchanged(resolved)

    def test_process_scan_classification_is_narrow_and_explicit(self) -> None:
        self.assertTrue(RUNNER.is_heavy_process("cargo", "cargo check"))
        self.assertTrue(
            RUNNER.is_heavy_process("python3", "python3 /tmp/odp_native_qualification_probe.py")
        )
        self.assertFalse(RUNNER.is_heavy_process("python3", "python3 harmless.py"))

    def test_process_guard_rejects_busy_or_opaque_state(self) -> None:
        protocol = RUNNER.Protocol(0, 1, 0, 100, 50, 1.0, 10, 1.0)
        with mock.patch.object(RUNNER, "load_average", return_value=0.0), mock.patch.object(
            RUNNER,
            "process_scan",
            return_value=[{"pid": 7, "executable": "rustc", "command": ["rustc"]}],
        ), self.assertRaises(RUNNER.RunFailure):
            RUNNER.quiet_guard(protocol, "test-busy")
        with mock.patch.object(
            RUNNER.Path, "read_bytes", side_effect=PermissionError("opaque")
        ), self.assertRaises(RUNNER.RunFailure):
            RUNNER.process_scan()
        with mock.patch.object(
            RUNNER.Path, "read_bytes", side_effect=FileNotFoundError("vanished")
        ), mock.patch.object(
            RUNNER.Path, "stat", side_effect=FileNotFoundError("vanished")
        ):
            self.assertIsNone(RUNNER._read_process_cmdline(Path("/proc/999999"), 999999))

    def test_malformed_manifest_writes_structured_failure_receipt(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = root / "input.json"
            output = root / "run"
            output.mkdir()
            manifest.write_text("{", encoding="utf-8")
            self.assertEqual(
                RUNNER.main(["--manifest", str(manifest), "--out-dir", str(output)]),
                1,
            )
            receipt = json.loads((output / "failure.json").read_text(encoding="utf-8"))
            self.assertEqual(receipt["schema"], RUNNER.REPORT_SCHEMA)
            self.assertEqual(receipt["status"], "failed")
            self.assertIn("error_type", receipt)

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
