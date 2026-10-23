#!/usr/bin/env python3
"""Build-free contract tests for the ODG joined qualification runner."""

from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest import mock
import sys


MODULE_PATH = Path(__file__).with_name("odg_joined_qualification.py")
SPEC = importlib.util.spec_from_file_location("odg_joined_qualification", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
RUNNER = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = RUNNER
SPEC.loader.exec_module(RUNNER)


def digest(value: bytes) -> dict[str, object]:
    return {"bytes": len(value), "sha256": hashlib.sha256(value).hexdigest()}


CONTENT = (
    b'<?xml version="1.0"?><office:document-content '
    b'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
    b'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">'
    b"<office:body><office:drawing><draw:page draw:name=\"Page 1\"/></office:drawing></office:body>"
    b"</office:document-content>"
)


def write_manifest(root: Path, *, limits: dict[str, int] | None = None, target_ms: int = 500) -> Path:
    fixture = root / "fixture.xml"
    fixture.write_bytes(CONTENT)
    common = root / "common.rs"
    baseline_probe = root / "baseline.rs"
    candidate_probe = root / "candidate.rs"
    for path in (common, baseline_probe, candidate_probe):
        path.write_bytes(path.name.encode("ascii"))
    baseline_root = root / "baseline"
    candidate_root = root / "candidate"
    baseline_root.mkdir()
    candidate_root.mkdir()
    for index, source in enumerate((baseline_root, candidate_root)):
        subprocess.run(["git", "-C", str(source), "init", "-q"], check=True)
        subprocess.run(["git", "-C", str(source), "config", "user.email", "test@example.invalid"], check=True)
        subprocess.run(["git", "-C", str(source), "config", "user.name", "Test"], check=True)
        (source / "tracked").write_text(f"source-{index}", encoding="utf-8")
        (source / "Cargo.lock").write_text(f"# lock-{index}\n", encoding="utf-8")
        (source / "rust-toolchain.toml").write_text(
            f"[toolchain]\nchannel = \"test-{index}\"\n", encoding="utf-8"
        )
        subprocess.run(["git", "-C", str(source), "add", "tracked", "Cargo.lock", "rust-toolchain.toml"], check=True)
        subprocess.run(["git", "-C", str(source), "commit", "-q", "-m", "source"], check=True)
    baseline_ref = subprocess.check_output(
        ["git", "-C", str(baseline_root), "rev-parse", "HEAD"], text=True
    ).strip()
    baseline_tree = subprocess.check_output(
        ["git", "-C", str(baseline_root), "rev-parse", "HEAD^{tree}"], text=True
    ).strip()
    candidate_ref = subprocess.check_output(
        ["git", "-C", str(candidate_root), "rev-parse", "HEAD"], text=True
    ).strip()
    candidate_tree = subprocess.check_output(
        ["git", "-C", str(candidate_root), "rev-parse", "HEAD^{tree}"], text=True
    ).strip()
    baseline_target = root / "baseline-target"
    candidate_target = root / "candidate-target"
    baseline_target.mkdir()
    candidate_target.mkdir()
    binary_name = "litchi-perf-odg-joined-qualification"
    for target in (baseline_target, candidate_target):
        binary = target / binary_name
        binary.write_text("#!/bin/sh\nexit 0\n", encoding="utf-8")
        binary.chmod(0o755)
    tmpdir = root / "tmp"
    tmpdir.mkdir()
    def identity(path: Path) -> dict[str, object]:
        return {"path": str(path), **digest(path.read_bytes())}

    values = {
        "schema": RUNNER.INPUT_SCHEMA,
        "fixture": {"path": str(fixture), **digest(CONTENT), "provenance": "generated test fixture"},
        "baseline": {"root": str(baseline_root), "ref": baseline_ref, "tree": baseline_tree,
                      "profile": RUNNER.BASELINE_PROFILE,
                      "binary": str(baseline_target / binary_name), "target_dir": str(baseline_target),
                      "lock": identity(baseline_root / "Cargo.lock"),
                      "toolchain": identity(baseline_root / "rust-toolchain.toml"), "features": ["odg"]},
        "candidate": {"root": str(candidate_root), "ref": candidate_ref, "tree": candidate_tree,
                       "profile": RUNNER.CANDIDATE_PROFILE,
                       "binary": str(candidate_target / binary_name), "target_dir": str(candidate_target),
                       "lock": identity(candidate_root / "Cargo.lock"),
                       "toolchain": identity(candidate_root / "rust-toolchain.toml"), "features": ["odg"]},
        "probe_sources": {
            name: {"path": str(path), **digest(path.read_bytes()),
                   "destination": RUNNER.EXPECTED_PROBE_DESTINATIONS[name]}
            for name, path in (("common", common), ("baseline", baseline_probe), ("candidate", candidate_probe))
        },
        "protocol": {"cpu": 0, "blocks": 1, "warmups": 0, "target_ms": target_ms,
                      "minimum_ms": 150, "max_load": 1000, "max_iterations": 1000, "timeout_s": 2},
        "runtime": {"tmpdir": str(tmpdir)},
        "composition_limits": limits or RUNNER.EXPECTED_LIMITS,
    }
    manifest = root / "manifest.json"
    manifest.write_text(json.dumps(values), encoding="utf-8")
    return manifest


class ODGJoinedQualificationTests(unittest.TestCase):
    def test_manifest_is_parameterized_and_profiles_are_narrow(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            parsed = RUNNER.load_manifest(write_manifest(Path(directory)))
            self.assertEqual(parsed.baseline.profile, RUNNER.BASELINE_PROFILE)
            self.assertEqual(parsed.candidate.profile, RUNNER.CANDIDATE_PROFILE)
            self.assertEqual(parsed.limits, RUNNER.EXPECTED_LIMITS)

    def test_duplicate_manifest_keys_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "manifest.json"
            path.write_text('{"schema":"x","schema":"y"}', encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.load_manifest(path)

    def test_limits_and_calibration_floor_are_fixed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.load_manifest(write_manifest(root, limits={**RUNNER.EXPECTED_LIMITS, "max_conflicts": 15}))
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.load_manifest(write_manifest(Path(directory), target_ms=499))

    def test_fixture_and_probe_tampering_is_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            parsed = RUNNER.load_manifest(manifest)
            parsed.fixture.path.write_bytes(CONTENT + b"tampered")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.validate_fixture(parsed.fixture)
            parsed = RUNNER.load_manifest(manifest)
            parsed.probe_sources[0].path.write_bytes(b"tampered")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.validate_probe_sources(parsed.probe_sources)

    def test_stale_requested_ref_and_binary_changes_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            parsed = RUNNER.load_manifest(write_manifest(root))
            out_dir = root / "out"
            out_dir.mkdir()
            identity = RUNNER.runtime_identity(parsed, out_dir)
            endpoint = identity["endpoints"]["baseline"]
            (endpoint.endpoint.root / "tracked").write_text("changed", encoding="utf-8")
            subprocess.run(["git", "-C", str(endpoint.endpoint.root), "add", "tracked"], check=True)
            subprocess.run(["git", "-C", str(endpoint.endpoint.root), "commit", "-q", "-m", "moved"], check=True)
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER._assert_unchanged(parsed, identity["endpoints"], "stale-ref")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            parsed = RUNNER.load_manifest(write_manifest(root))
            out_dir = root / "out"
            out_dir.mkdir()
            identity = RUNNER.runtime_identity(parsed, out_dir)
            binary = identity["endpoints"]["candidate"].endpoint.binary
            binary.write_text("#!/bin/sh\nexit 1\n", encoding="utf-8")
            binary.chmod(0o755)
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER._assert_unchanged(parsed, identity["endpoints"], "binary-change")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            parsed = RUNNER.load_manifest(write_manifest(root))
            out_dir = root / "out"
            out_dir.mkdir()
            identity = RUNNER.runtime_identity(parsed, out_dir)
            lock = identity["endpoints"]["baseline"].endpoint.lock.path
            lock.write_text(lock.read_text(encoding="utf-8") + "tampered\n", encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER._assert_unchanged(parsed, identity["endpoints"], "lock-change")

    def test_manifest_mutation_is_rejected_after_identity_capture(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            parsed = RUNNER.load_manifest(manifest)
            out_dir = root / "out"
            out_dir.mkdir()
            identity = RUNNER.runtime_identity(parsed, out_dir)
            manifest.write_text(manifest.read_text(encoding="utf-8") + "\n", encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER._assert_unchanged(parsed, identity["endpoints"], "manifest-change")

    def test_path_layout_rejects_target_inside_source(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            data = json.loads(manifest.read_text(encoding="utf-8"))
            data["candidate"]["target_dir"] = data["candidate"]["root"]
            manifest.write_text(json.dumps(data), encoding="utf-8")
            parsed = RUNNER.load_manifest(manifest)
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.runtime_identity(parsed, root / "out")

    def test_verify_route_proof_rejects_incomplete_legs_and_output_tamper(self) -> None:
        base = {
            "output": digest(b"out"), "published": digest(b"out"),
            "semantic": {"pages": [{"name": "Page 1", "shapes": []},
                                    {"name": "Page 2", "shapes": []},
                                    {"name": "Replacement Page", "shapes": []}]},
            "checks": {"source_unchanged": True, "publication_readback_exact": True,
                        "route_semantics": True, "changed_output": True, "exact_noop": True,
                        "deterministic_repeat": True},
            "baseline": {"apply_joined": True, "commit_triads": False, "patch_replay": False},
        }
        self.assertEqual(RUNNER._validate_route("package_replace_one", base, False)["output"], digest(b"out"))
        broken = dict(base)
        broken["checks"] = dict(base["checks"])
        broken["checks"]["route_semantics"] = False
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER._validate_route("package_replace_one", broken, False)
        altered = dict(base)
        altered["semantic"] = {"pages": [{"name": "wrong", "shapes": []}]}
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER._validate_route("package_replace_one", altered, False)
        self.assertEqual(list(RUNNER.ABBA_ORDER), ["a1", "b1", "b2", "a2"])

    def test_abba_validator_rejects_incomplete_or_mismatched_legs(self) -> None:
        row = {"role": "a1", "endpoint": "baseline", "iterations": 1,
               "sample_ns": 200_000_000, "per_operation_ns": 200_000_000.0,
               "guard": {}, "target": {}, "logs": {}, "proof": {}}
        rows = [dict(row), dict(row), dict(row), dict(row)]
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_abba_rows(rows[:3], "empty_join/join_apply")
        rows[1]["role"] = "a2"
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_abba_rows(rows, "empty_join/join_apply")

    def test_timed_validator_rejects_wrong_binary_report_identity_and_short_sample(self) -> None:
        protocol = RUNNER.Protocol(0, 1, 0, 500_000_000, 150_000_000, 1000, 100, 2)
        expected = {"output": digest(b"out"), "semantic": {"pages": []}}
        report = {"schema": RUNNER.BASELINE_SCHEMA, "endpoint": "wrong", "mode": "timed",
                  "route": "empty_join", "scope": "join_apply", "iterations": 1, "warmups": 0,
                  "samples": [200_000_000], "proof": expected}
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_timed(report, "baseline", "empty_join", "join_apply", 1, protocol, expected)

    def test_manifest_report_endpoint_and_extra_fields_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            manifest = write_manifest(Path(directory))
            data = json.loads(manifest.read_text(encoding="utf-8"))
            data["protocol"]["unexpected"] = True
            manifest.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.load_manifest(manifest)
        base = {
            "fixture": digest(CONTENT), "source_package": digest(b"source"),
            "routes": {}, "refusals": {}, "capabilities": {}, "checks": {},
            "schema": RUNNER.BASELINE_SCHEMA, "endpoint": "wrong", "mode": "verify",
        }
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_verify_report(base, "baseline", digest(CONTENT))

    def test_child_log_binding_rejects_binary_or_phase_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            parsed = RUNNER.load_manifest(write_manifest(root))
            out_dir = root / "out"
            out_dir.mkdir()
            identity = RUNNER.runtime_identity(parsed, out_dir)
            endpoint = identity["endpoints"]["baseline"]
            logs = {
                "command": [], "returncode": 0, "endpoint": "baseline",
                "profile": RUNNER.BASELINE_PROFILE, "resolved_ref": endpoint.resolved_ref,
                "binary": endpoint.binary,
                "binding": {"mode": "verify", "route": None, "scope": None},
                "stdout": {}, "stderr": {},
            }
            RUNNER._validate_log_binding(logs, endpoint, "verify", None, None, "verify")
            altered = dict(logs)
            altered["binary"] = dict(endpoint.binary)
            altered["binary"]["sha256"] = "0" * 64
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER._validate_log_binding(altered, endpoint, "verify", None, None, "verify")
            altered = dict(logs)
            altered["binding"] = {"mode": "timed", "route": None, "scope": None}
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER._validate_log_binding(altered, endpoint, "verify", None, None, "verify")

    def test_quiet_guard_failure_is_propagated(self) -> None:
        with mock.patch.object(RUNNER, "quiet_guard", side_effect=RUNNER.RunFailure("busy")):
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.quiet_guard(RUNNER.Protocol(0, 1, 0, 500, 150, 1000, 10, 1), "test")


if __name__ == "__main__":
    unittest.main()
