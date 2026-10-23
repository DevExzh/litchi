"""Unit tests for the fail-closed IWA native qualification runner.

These tests use a disposable executable that emits probe-shaped JSON.  They
do not build a Rust crate, run a benchmark, or create evidence under the
checkout.
"""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import stat
import tempfile
import unittest
from unittest import mock


HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("iwa_native_qualification", HERE / "iwa_native_qualification.py")
assert SPEC and SPEC.loader
RUNNER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(RUNNER)


def _fake_probe(path: Path, *, owner: bool, valid: bool = True) -> None:
    source = "a" * 64
    output = "b" * 64
    if owner:
        correctness = {
            "status": "ok",
            "replay_matches_changed": valid,
            "inverse_matches_source": valid,
            "noop_is_noop": valid,
            "noop_matches_expected": valid,
        }
        body = {
            "schema": "litchi.iwa.owner.probe.v2",
            "route": "numbers_headers",
            "iterations": 1,
            "ns": 500_000_000,
            "checksum": 11,
            "output_bytes": 12,
            "source_sha256": source,
            "output_sha256": output,
            "output_differs": True,
            "correctness": correctness,
        }
    else:
        body = {
            "schema": "litchi.iwa.constructor.probe.v3",
            "route": "field_path",
            "iterations": 1,
            "ns": 500_000_000,
            "checksum": 11,
            "output_bytes": 12,
            "source_sha256": source,
            "output_sha256": source,
            "output_matches_source": True,
            "correctness": {
                "status": "ok",
                "roundtrip": True,
                "replay": "not-applicable",
                "noop": "not-applicable",
            },
        }
    path.write_text(
        "#!/usr/bin/env python3\n"
        "import json, sys\n"
        f"print({json.dumps(json.dumps(body))})\n",
        encoding="utf-8",
    )
    path.chmod(path.stat().st_mode | stat.S_IXUSR)


class NativeQualificationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)

    def test_artifact_identity_is_content_bound(self) -> None:
        path = self.directory / "probe"
        path.write_bytes(b"one")
        identity = RUNNER.artifact(path)
        RUNNER.verify_artifact(identity)
        path.write_bytes(b"two")
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.verify_artifact(identity)

    def test_output_inside_checkout_is_rejected(self) -> None:
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.outside_checkout(RUNNER.ROOT / "docs/unified-ops/results/0361fda")

    def test_owner_fixture_manifest_is_content_bound(self) -> None:
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.fixture_identity(self.directory)
        fixture_dir = self.directory / "fixtures"
        RUNNER.prepare_fixtures(fixture_dir)
        identity = RUNNER.fixture_identity(fixture_dir)
        self.assertEqual(set(identity["files"]), set(RUNNER.OWNER_FIXTURES))
        (fixture_dir / "tone.wav").write_bytes(b"changed")
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.fixture_identity(fixture_dir)

    def test_fixed_protocol_rejects_weakening_without_diagnostic(self) -> None:
        args = type(
            "Args",
            (),
            {
                "start_count": RUNNER.START_COUNTS["constructor"],
                "blocks": RUNNER.FIXED_BLOCKS - 1,
                "warmups": RUNNER.FIXED_WARMUPS,
                "target_ns": RUNNER.TARGET_NS,
                "minimum_ns": RUNNER.MINIMUM_NS,
                "growth": RUNNER.FIXED_GROWTH,
                "max_iterations": RUNNER.MAX_ITERATIONS,
                "timeout_s": RUNNER.FIXED_TIMEOUT_S,
            },
        )()
        protocol = RUNNER._protocol(args, "constructor", RUNNER.ROUTES["constructor"])
        self.assertFalse(RUNNER._fixed_protocol(protocol, "constructor"))
        args.blocks = RUNNER.FIXED_BLOCKS
        self.assertTrue(RUNNER._fixed_protocol(protocol | {"blocks": RUNNER.FIXED_BLOCKS}, "constructor"))

    def test_unapproved_candidate_source_addition_is_rejected(self) -> None:
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._verify_source_sets(
                {"shared.rs": "a"},
                {"shared.rs": "a", "unapproved.txt": "b"},
                "constructor",
            )
        RUNNER._verify_source_sets(
            {"crates/litchi-iwa-core/old.rs": "a"},
            {
                "crates/litchi-iwa-core/old.rs": "b",
                "crates/litchi-iwa-core/new.rs": "c",
            },
            "constructor",
        )

    def test_duplicate_json_keys_are_rejected(self) -> None:
        path = self.directory / "duplicate.json"
        path.write_text('{"a": 1, "a": 2}', encoding="utf-8")
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER.load_json(path)

    def test_constructor_probe_requires_exact_roundtrip(self) -> None:
        value = {
            "route": "field_path",
            "iterations": 3,
            "ns": 10,
            "checksum": 1,
            "source_sha256": "a" * 64,
            "output_sha256": "a" * 64,
            "output_matches_source": True,
            "correctness": {"status": "ok", "roundtrip": True},
        }
        self.assertIs(RUNNER._strict_probe_value(value, "constructor", "field_path", 3), value)
        value["correctness"]["roundtrip"] = False
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._strict_probe_value(value, "constructor", "field_path", 3)

    def test_owner_probe_requires_replay_inverse_and_noop(self) -> None:
        value = {
            "route": "numbers_headers",
            "iterations": 3,
            "ns": 10,
            "checksum": 1,
            "source_sha256": "a" * 64,
            "output_sha256": "b" * 64,
            "output_differs": True,
            "correctness": {
                "status": "ok",
                "replay_matches_changed": True,
                "inverse_matches_source": True,
                "noop_is_noop": True,
                "noop_matches_expected": True,
            },
        }
        self.assertIs(RUNNER._strict_probe_value(value, "owner", "numbers_headers", 3), value)
        value["correctness"]["inverse_matches_source"] = False
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._strict_probe_value(value, "owner", "numbers_headers", 3)

    def test_invoke_records_receipt_and_rejects_wrong_identity(self) -> None:
        binary = self.directory / "probe.py"
        _fake_probe(binary, owner=False)
        expected = RUNNER.artifact(binary)
        with mock.patch.object(RUNNER, "host_snapshot"):
            value = RUNNER.invoke(
                self.directory,
                "constructor",
                "field_path",
                binary,
                "field_path",
                1,
                min(RUNNER.os.sched_getaffinity(0)),
                None,
                expected,
            )
        self.assertEqual(value["route"], "field_path")
        receipt = RUNNER.load_json(self.directory / "field_path.receipt.json")
        self.assertEqual(receipt["exit_code"], 0)
        self.assertEqual(receipt["binary"], expected)

    def test_invoke_fails_closed_on_owner_correctness(self) -> None:
        binary = self.directory / "probe.py"
        _fake_probe(binary, owner=True, valid=False)
        expected = RUNNER.artifact(binary)
        with mock.patch.object(RUNNER, "host_snapshot"):
            with self.assertRaises(RUNNER.QualificationError):
                RUNNER.invoke(
                    self.directory,
                    "owner",
                    "numbers_headers",
                    binary,
                    "numbers_headers",
                    1,
                    min(RUNNER.os.sched_getaffinity(0)),
                    self.directory,
                    expected,
                )

    def test_invoke_enforces_timed_child_bound(self) -> None:
        binary = self.directory / "slow_probe.py"
        binary.write_text(
            "#!/usr/bin/env python3\nimport time\ntime.sleep(1)\n",
            encoding="utf-8",
        )
        binary.chmod(binary.stat().st_mode | stat.S_IXUSR)
        expected = RUNNER.artifact(binary)
        with mock.patch.object(RUNNER, "host_snapshot"):
            with self.assertRaises(RUNNER.QualificationError):
                RUNNER.invoke(
                    self.directory,
                    "constructor",
                    "timeout",
                    binary,
                    "field_path",
                    1,
                    min(RUNNER.os.sched_getaffinity(0)),
                    None,
                    expected,
                    timeout_s=0.05,
                )
        receipt = RUNNER.load_json(self.directory / "timeout.receipt.json")
        self.assertTrue(receipt["timed_out"])
        self.assertEqual(receipt["timeout_s"], 0.05)

    def test_abba_identity_check_rejects_mismatched_outputs(self) -> None:
        values = [
            {"checksum": 1, "output_sha256": "a", "source_sha256": "s"},
            {"checksum": 1, "output_sha256": "b", "source_sha256": "s"},
            {"checksum": 1, "output_sha256": "b", "source_sha256": "s"},
            {"checksum": 1, "output_sha256": "a", "source_sha256": "s"},
        ]
        with self.assertRaises(RUNNER.QualificationError):
            RUNNER._ratios_equal(values)

    def test_host_guard_fails_closed_on_foreign_build(self) -> None:
        foreign = {"pid": 99999, "parent": 1, "exe": "/usr/bin/cargo", "args": "cargo build"}
        with mock.patch.object(RUNNER, "process_rows", return_value=[foreign]):
            with self.assertRaises(RUNNER.QualificationError):
                RUNNER.host_snapshot(self.directory, "blocked")
        snapshot = RUNNER.load_json(self.directory / "blocked.host.json")
        self.assertEqual(snapshot["blockers"], [foreign])


if __name__ == "__main__":
    unittest.main()
