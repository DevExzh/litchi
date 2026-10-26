#!/usr/bin/env python3
"""Build-free contract tests for the ODG joined qualification runner."""

from __future__ import annotations

import hashlib
import importlib.util
import json
from dataclasses import replace
import os
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
    from tools import odg_joined_qualification_builder as builder

    fixture = RUNNER.REVIEWED_REPO_ROOT / RUNNER.REVIEWED_FIXTURE_RELATIVE
    reviewed = {
        name: RUNNER.REVIEWED_REPO_ROOT / relative
        for name, relative in RUNNER.REVIEWED_PROBE_RELATIVES.items()
    }
    common = reviewed["common"]
    allocator = reviewed["allocator"]
    baseline_probe = reviewed["baseline"]
    candidate_probe = reviewed["candidate"]
    baseline_root = root / "baseline"
    candidate_root = root / "candidate"
    baseline_root.mkdir()
    candidate_root.mkdir()
    for index, source in enumerate((baseline_root, candidate_root)):
        subprocess.run(["git", "-C", str(source), "init", "-q"], check=True)
        subprocess.run(["git", "-C", str(source), "config", "user.email", "test@example.invalid"], check=True)
        subprocess.run(["git", "-C", str(source), "config", "user.name", "Test"], check=True)
        (source / "tracked").write_text(f"source-{index}", encoding="utf-8")
        (source / "Cargo.toml").write_text(
            "[workspace]\nmembers = [\"crates/litchi-odg\"]\n", encoding="utf-8"
        )
        (source / "Cargo.lock").write_text(f"# lock-{index}\n", encoding="utf-8")
        (source / "rust-toolchain.toml").write_text(
            f"[toolchain]\nchannel = \"test-{index}\"\n", encoding="utf-8"
        )
        (source / ".cargo").mkdir()
        (source / ".cargo" / "config.toml").write_text("[build]\njobs = 1\n", encoding="utf-8")
        subprocess.run(["git", "-C", str(source), "add", "tracked", "Cargo.toml", "Cargo.lock", "rust-toolchain.toml", ".cargo/config.toml"], check=True)
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
    baseline_allocator_target = root / "baseline-allocator-target"
    candidate_allocator_target = root / "candidate-allocator-target"
    baseline_target.mkdir()
    candidate_target.mkdir()
    baseline_allocator_target.mkdir()
    candidate_allocator_target.mkdir()
    binary_name = "litchi-perf-odg-joined-qualification"
    for target in (baseline_target, candidate_target, baseline_allocator_target, candidate_allocator_target):
        binary = target / "release" / "examples" / binary_name
        binary.parent.mkdir(parents=True)
        binary.write_text("#!/bin/sh\nexit 0\n", encoding="utf-8")
        binary.chmod(0o755)
    tmpdir = root / "tmp"
    tmpdir.mkdir()
    def identity(path: Path) -> dict[str, object]:
        return {"path": str(path), **digest(path.read_bytes())}

    def tree_identity(source: Path) -> dict[str, object]:
        value = subprocess.check_output(
            ["git", "-C", str(source), "ls-tree", "-r", "--full-tree", "--long", "HEAD"]
        )
        return {"entries": len(value.splitlines()), "bytes": len(value),
                "sha256": hashlib.sha256(value).hexdigest()}

    def source_closure(source: Path, revision: str) -> dict[str, object]:
        return builder._capture_source_closure(source, revision)["binding"]

    values = {
        "schema": RUNNER.INPUT_SCHEMA,
        "fixture": {"path": str(fixture), **RUNNER.REVIEWED_FIXTURE_IDENTITY,
                    "provenance": RUNNER.REVIEWED_FIXTURE_PROVENANCE},
        "baseline": {"root": str(baseline_root), "ref": baseline_ref, "tree": baseline_tree,
                      "profile": RUNNER.BASELINE_PROFILE,
                      "binary": str(baseline_target / "release" / "examples" / binary_name), "target_dir": str(baseline_target),
                      "lock": identity(baseline_root / "Cargo.lock"),
                      "toolchain": identity(baseline_root / "rust-toolchain.toml"),
                      "cargo_config": identity(baseline_root / ".cargo/config.toml"),
                      "cargo_manifest": identity(baseline_root / "Cargo.toml"),
                      "source_manifest": tree_identity(baseline_root),
                      "source_closure": source_closure(baseline_root, baseline_ref),
                      "allocator_binary": str(baseline_allocator_target / "release" / "examples" / binary_name),
                      "allocator_target_dir": str(baseline_allocator_target), "features": ["odg"],
                      "binary_identity": digest((baseline_target / "release" / "examples" / binary_name).read_bytes()),
                      "allocator_binary_identity": digest((baseline_allocator_target / "release" / "examples" / binary_name).read_bytes())},
        "candidate": {"root": str(candidate_root), "ref": candidate_ref, "tree": candidate_tree,
                       "profile": RUNNER.CANDIDATE_PROFILE,
                       "binary": str(candidate_target / "release" / "examples" / binary_name), "target_dir": str(candidate_target),
                       "lock": identity(candidate_root / "Cargo.lock"),
                       "toolchain": identity(candidate_root / "rust-toolchain.toml"),
                       "cargo_config": identity(candidate_root / ".cargo/config.toml"),
                       "cargo_manifest": identity(candidate_root / "Cargo.toml"),
                       "source_manifest": tree_identity(candidate_root),
                       "source_closure": source_closure(candidate_root, candidate_ref),
                       "allocator_binary": str(candidate_allocator_target / "release" / "examples" / binary_name),
                       "allocator_target_dir": str(candidate_allocator_target), "features": ["odg"],
                       "binary_identity": digest((candidate_target / "release" / "examples" / binary_name).read_bytes()),
                       "allocator_binary_identity": digest((candidate_allocator_target / "release" / "examples" / binary_name).read_bytes())},
        "probe_sources": {
            name: {"path": str(path), **digest(path.read_bytes()),
                   "destination": RUNNER.EXPECTED_PROBE_DESTINATIONS[name]}
            for name, path in (("common", common), ("allocator", allocator),
                               ("baseline", baseline_probe), ("candidate", candidate_probe))
        },
        "protocol": {"cpu": 0, "blocks": 2, "warmups": 20, "target_ms": target_ms,
                      "minimum_ms": 150, "max_load": 1000, "max_iterations": 1000, "timeout_s": 2},
        "runtime": {"tmpdir": str(tmpdir)},
        "build": {"prepare_receipt": str(root / "prepare.json"), "receipt": str(root / "build.json"),
                  "receipt_identity": {"path": str(root / "build.json"), "bytes": 1, "sha256": "0" * 64},
                  "final_candidate_ref": candidate_ref, "cpu": 0},
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

    def test_protocol_counts_are_fixed_for_qualification(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            data = json.loads(manifest.read_text(encoding="utf-8"))
            data["protocol"]["blocks"] = RUNNER.REQUIRED_BLOCKS - 1
            manifest.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.load_manifest(manifest)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            data = json.loads(manifest.read_text(encoding="utf-8"))
            data["protocol"]["warmups"] = RUNNER.REQUIRED_WARMUPS - 1
            manifest.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.load_manifest(manifest)

    def test_fixture_and_probe_tampering_is_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            parsed = RUNNER.load_manifest(manifest)
            tampered_fixture = replace(parsed.fixture, sha256="0" * 64)
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.validate_fixture(tampered_fixture)
            parsed = RUNNER.load_manifest(manifest)
            tampered_probe = replace(parsed.probe_sources[0], sha256="0" * 64)
            tampered_sources = (tampered_probe, *parsed.probe_sources[1:])
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.validate_probe_sources(tampered_sources)

    def test_stale_requested_ref_and_binary_changes_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            parsed = RUNNER.load_manifest(write_manifest(root))
            out_dir = root / "out"
            out_dir.mkdir()
            identity = RUNNER.runtime_identity(parsed, out_dir, require_build_receipt=False)
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
            identity = RUNNER.runtime_identity(parsed, out_dir, require_build_receipt=False)
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
            identity = RUNNER.runtime_identity(parsed, out_dir, require_build_receipt=False)
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
            identity = RUNNER.runtime_identity(parsed, out_dir, require_build_receipt=False)
            manifest.write_text(manifest.read_text(encoding="utf-8") + "\n", encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER._assert_unchanged(parsed, identity["endpoints"], "manifest-change")

    def test_forced_source_closure_catches_bytes_when_stat_probe_is_clean(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            parsed = RUNNER.load_manifest(write_manifest(root))
            out_dir = root / "out"
            out_dir.mkdir()
            identity = RUNNER.runtime_identity(parsed, out_dir, require_build_receipt=False)
            endpoint = identity["endpoints"]["baseline"]
            tracked = endpoint.endpoint.root / "tracked"
            original = tracked.stat()
            tracked.write_bytes(b"tamper-0")
            os.utime(tracked, ns=(original.st_atime_ns, original.st_mtime_ns))
            from tools import odg_joined_qualification_builder as builder

            with mock.patch.object(builder, "_cheap_source_closure_probe", return_value=True):
                with self.assertRaises(RUNNER.RunFailure):
                    builder._validate_source_closure(
                        endpoint.endpoint.root,
                        endpoint.source_closure,
                        endpoint.endpoint.source_closure,
                        force_full=True,
                    )

    def test_path_layout_rejects_target_inside_source(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            data = json.loads(manifest.read_text(encoding="utf-8"))
            data["candidate"]["target_dir"] = data["candidate"]["root"]
            manifest.write_text(json.dumps(data), encoding="utf-8")
            parsed = RUNNER.load_manifest(manifest)
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.runtime_identity(parsed, root / "out", require_build_receipt=False)

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

    def test_baseline_shaped_report_preserves_candidate_only_refusal_scope(self) -> None:
        source = digest(b"source")
        pages = [
            {"name": "Page 1", "shapes": [
                {"name": "Shape 1", "text": "Joined page one"},
                {"name": "Joined page one name", "text": ""},
            ]},
            {"name": "Page 2", "shapes": [
                {"name": "Shape 3", "text": "Joined page two"},
                {"name": "Joined page two name", "text": ""},
            ]},
            {"name": "Replacement Page", "shapes": []},
        ]
        pages_two = json.loads(json.dumps(pages))
        pages_two[1]["shapes"][1]["name"] = "Joined page two"
        route = lambda output, semantic: {
            "output": digest(output), "published": digest(output), "semantic": semantic,
            "checks": {"source_unchanged": True, "publication_readback_exact": True,
                        "route_semantics": True, "changed_output": True, "exact_noop": True,
                        "deterministic_repeat": True},
            "baseline": {"apply_joined": True, "commit_triads": False, "patch_replay": False},
        }
        routes = {
            "package_replace_one": route(b"replace", {"pages": pages}),
            "semantic_two_disjoint": route(b"two", {"pages": pages_two}),
            "semantic_four_disjoint": route(b"four", {"pages": pages}),
            "empty_join": route(b"source", {"pages": pages}),
        }
        routes["empty_join"]["checks"]["changed_output"] = True
        routes["empty_join"]["checks"]["exact_noop"] = True
        report = {
            "schema": RUNNER.BASELINE_SCHEMA, "endpoint": "baseline-apply-joined", "mode": "verify",
            "fixture": digest(CONTENT), "source_package": source, "routes": routes,
            "refusals": {"stale_lineage_refused": True, "conflicting_effect_refused": True,
                          "changed_recipient_refused": False, "changed_recipient_scope": "candidate-only"},
            "capabilities": {"apply_joined": True, "edit_join_commit": False, "commit_triads": False,
                              "patch_replay": False, "inverse_replay": False,
                              "content_matching_replay": False, "changed_recipient_refusal": False,
                              "changed_recipient_refusal_scope": "candidate-only"},
            "checks": {"source_fixture_deterministic": True, "all_routes_checked": True,
                        "all_refusal_proofs": True, "all_baseline_apply_proofs": True},
        }
        self.assertEqual(RUNNER.validate_verify_report(report, "baseline", digest(CONTENT))["source"], source)
        altered = json.loads(json.dumps(report))
        altered["routes"]["empty_join"]["output"] = digest(b"not-source")
        altered["routes"]["empty_join"]["published"] = digest(b"not-source")
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_verify_report(altered, "baseline", digest(CONTENT))

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

    def test_calibration_binds_reviewed_adaptive_iteration_contract(self) -> None:
        protocol = RUNNER.Protocol(0, 1, RUNNER.REQUIRED_WARMUPS,
                                   500_000_000, 150_000_000, 1000, 8, 2)
        expected = {"output": digest(b"out"), "semantic": {"pages": []}}
        report = {
            "schema": RUNNER.BASELINE_SCHEMA, "endpoint": "baseline-apply-joined",
            "mode": "calibrate", "route": "empty_join", "scope": "join_apply",
            "target_ns": protocol.target_ns, "iterations": 4, "elapsed_ns": protocol.target_ns,
            "iteration_policy": RUNNER.CALIBRATION_POLICY, "initial_iterations": 1,
            "max_iterations": protocol.max_iterations,
            "attempts": [
                {"iterations": 1, "elapsed_ns": 100, "proof": expected},
                {"iterations": 2, "elapsed_ns": 200, "proof": expected},
                {"iterations": 4, "elapsed_ns": protocol.target_ns, "proof": expected},
            ],
        }
        self.assertEqual(
            RUNNER.validate_calibration(report, "baseline", "empty_join", "join_apply", protocol, expected),
            4,
        )
        altered = json.loads(json.dumps(report))
        altered["attempts"][1]["iterations"] = 3
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_calibration(altered, "baseline", "empty_join", "join_apply", protocol, expected)

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

    def test_manifest_requires_external_receipt_and_binary_bindings(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            data = json.loads(manifest.read_text(encoding="utf-8"))
            del data["baseline"]["binary_identity"]
            manifest.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.load_manifest(manifest)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            data = json.loads(manifest.read_text(encoding="utf-8"))
            del data["build"]["receipt_identity"]
            manifest.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER.load_manifest(manifest)

    def test_child_log_binding_rejects_binary_or_phase_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            parsed = RUNNER.load_manifest(write_manifest(root))
            out_dir = root / "out"
            out_dir.mkdir()
            identity = RUNNER.runtime_identity(parsed, out_dir, require_build_receipt=False)
            endpoint = identity["endpoints"]["baseline"]
            stdout_path = out_dir / "stdout.log"
            stderr_path = out_dir / "stderr.log"
            stdout_path.write_bytes(b"{}\n")
            stderr_path.write_bytes(b"x")
            quiet = {"phase": "guard", "load_1m": 0.0, "max_load": 1000.0,
                     "heavy_processes": [], "checked_unix_ns": 1}
            logs = {
                "command": [], "returncode": 0, "endpoint": "baseline",
                "profile": RUNNER.BASELINE_PROFILE, "resolved_ref": endpoint.resolved_ref,
                "binary": endpoint.binary,
                "environment": endpoint.environment,
                "binding": {"mode": "verify", "route": None, "scope": None},
                "stdout": {"path": str(stdout_path), "bytes": 3,
                            "sha256": hashlib.sha256(b"{}\n").hexdigest()},
                "stderr": {"path": str(stderr_path), "bytes": 1,
                            "sha256": hashlib.sha256(b"x").hexdigest()},
                "guard": {"before": quiet, "live": [{"checked_unix_ns": 1, "output_bytes": 4,
                                                         "rss_bytes": 0, "load_1m": 0.0,
                                                         "outside_processes": [], "owned_processes": []}], "after": quiet,
                           "memory_limit_bytes": 20_000_000_000,
                           "max_output_bytes": RUNNER.MAX_CHILD_OUTPUT,
                           "max_load": 1000.0},
                "target": {"path": str(endpoint.endpoint.target_dir), "bytes_before": 0,
                           "bytes_after": 0, "hard_limit_bytes": RUNNER.TARGET_HARD_LIMIT_BYTES},
            }
            RUNNER._validate_log_binding(logs, endpoint, "verify", None, None, "verify")
            altered = dict(logs)
            altered["binary"] = dict(endpoint.binary)
            altered["binary"]["sha256"] = "0" * 64
            with self.assertRaises(RUNNER.RunFailure):
                RUNNER._validate_log_binding(altered, endpoint, "verify", None, None, "verify")

    def test_builder_refuses_upstream_and_overwrite_paths(self) -> None:
        from tools import odg_joined_qualification_builder as builder

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            upstream_component = "3rd" + "party"
            with self.assertRaises(RUNNER.RunFailure):
                builder._safe_path(root / upstream_component / "fixture.xml", "fixture")
            receipt = root / "receipt.json"
            builder._write_fresh_json(receipt, {"ok": True}, "receipt")
            with self.assertRaises(RUNNER.RunFailure):
                builder._write_fresh_json(receipt, {"ok": False}, "receipt")
            logs = root / "logs"
            logs.mkdir()
            empty = logs / "stderr"
            empty.write_bytes(b"")
            self.assertEqual(
                builder._validate_log_identity(
                    {"path": str(empty), "bytes": 0, "sha256": hashlib.sha256(b"").hexdigest()},
                    "empty.stderr", logs,
                )["bytes"],
                0,
            )

    def test_builder_binds_release_package_command_and_reviewed_inputs(self) -> None:
        from tools import odg_joined_qualification_builder as builder

        command = builder._command(
            {
                "root": "/source",
                "target_dir": "/target",
                "allocator_target_dir": "/allocator-target",
                "features": [],
            },
            allocator=False,
            cpu=3,
        )
        self.assertIn("--package", command["argv"])
        self.assertIn("litchi-odg", command["argv"])
        self.assertIn("--release", command["argv"])
        self.assertEqual(
            command["argv"][-2:],
            ["--example", "litchi-perf-odg-joined-qualification"],
        )
        with tempfile.TemporaryDirectory() as directory:
            fake = Path(directory) / "odg_joined_qualification_common.rs"
            fake.write_bytes(b"not reviewed")
            with self.assertRaises(RUNNER.RunFailure):
                builder._reviewed_file(
                    fake,
                    RUNNER.REVIEWED_PROBE_RELATIVES["common"],
                    RUNNER.REVIEWED_PROBE_IDENTITIES["common"],
                    "fake probe",
                )
            self.assertEqual(
                builder._reviewed_file(
                    RUNNER.REVIEWED_REPO_ROOT / RUNNER.REVIEWED_PROBE_RELATIVES["common"],
                    RUNNER.REVIEWED_PROBE_RELATIVES["common"],
                    RUNNER.REVIEWED_PROBE_IDENTITIES["common"],
                    "reviewed probe",
                )["sha256"],
                RUNNER.REVIEWED_PROBE_IDENTITIES["common"]["sha256"],
            )

    def test_source_binder_accepts_only_pending_binary_slots(self) -> None:
        from tools import odg_joined_qualification_builder as builder

        def endpoint(name: str) -> dict[str, object]:
            root = "/external/" + name
            digest_value = "a" * 64
            files = {
                "path": "Cargo.lock", "bytes": 1, "sha256": digest_value,
            }
            return {
                "root": root,
                "ref": "1" * 40,
                "tree": "2" * 40,
                "profile": "legacy-apply-joined" if name == "baseline" else "current-edit-join-commit",
                "features": ["odg"],
                "source_manifest": {"entries": 1, "bytes": 1, "sha256": digest_value},
                "source_closure": {"entries": 1, "bytes": 1, "sha256": digest_value},
                "lock": files,
                "toolchain": {**files, "path": "rust-toolchain.toml"},
                "cargo_config": {**files, "path": ".cargo/config.toml"},
                "cargo_manifest": {**files, "path": "Cargo.toml"},
                "binary": {
                    "path": root + "/target/release/examples/litchi-perf-odg-joined-qualification",
                    "bytes": None, "sha256": None,
                    "status": "pending " + ("baseline" if name == "baseline" else "C2a") + " normal build",
                },
                "allocator_binary": {
                    "path": root + "/target/release/examples/litchi-perf-odg-joined-qualification",
                    "bytes": None, "sha256": None,
                    "status": "pending " + ("baseline" if name == "baseline" else "C2a") + " allocator build",
                },
            }

        accepted = builder._source_binder_endpoint({"candidate": endpoint("candidate")}, "candidate")
        self.assertIsNone(accepted["binary"]["sha256"])
        rejected = endpoint("candidate")
        rejected["binary"] = dict(rejected["binary"])
        rejected["binary"]["bytes"] = 1
        with self.assertRaises(RUNNER.RunFailure):
            builder._source_binder_endpoint({"candidate": rejected}, "candidate")

    def test_builder_rejects_assume_unchanged_and_skip_worktree(self) -> None:
        from tools import odg_joined_qualification_builder as builder

        for flag in ("--assume-unchanged", "--skip-worktree"):
            with tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                subprocess.run(["git", "-C", str(root), "init", "-q"], check=True)
                subprocess.run(["git", "-C", str(root), "config", "user.email", "test@example.invalid"], check=True)
                subprocess.run(["git", "-C", str(root), "config", "user.name", "Test"], check=True)
                tracked = root / "tracked"
                tracked.write_text("tracked", encoding="utf-8")
                subprocess.run(["git", "-C", str(root), "add", "tracked"], check=True)
                subprocess.run(["git", "-C", str(root), "commit", "-q", "-m", "tracked"], check=True)
                subprocess.run(["git", "-C", str(root), "update-index", flag, "tracked"], check=True)
                with self.assertRaises(RUNNER.RunFailure):
                    builder._check_worktree_shape(root, flag)

    def test_builder_binds_effective_environment_and_final_target_sample(self) -> None:
        from tools import odg_joined_qualification_builder as builder

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            parsed = RUNNER.load_manifest(write_manifest(root))
            endpoint = parsed.baseline
            for variable in ("RUSTC", "CARGO_BUILD_RUSTC", "RUSTC_WRAPPER"):
                with mock.patch.dict(os.environ, {variable: "/unreviewed/compiler"}, clear=False):
                    with self.assertRaises(RUNNER.RunFailure):
                        builder._environment_identity(
                            endpoint.root,
                            ref=endpoint.requested_ref,
                            tree=endpoint.expected_tree,
                            source_closure=endpoint.source_closure,
                        )
            tool_dir = root / "fake-tools"
            tool_dir.mkdir()
            for name in ("cargo", "rustc"):
                tool = tool_dir / name
                tool.write_text(f"#!/bin/sh\n# {name}\n", encoding="utf-8")
                tool.chmod(0o755)
            first = builder._environment_identity(
                endpoint.root,
                ref=endpoint.requested_ref,
                tree=endpoint.expected_tree,
                source_closure=endpoint.source_closure,
                overrides={"PATH": str(tool_dir)},
            )
            (tool_dir / "rustc").write_text("#!/bin/sh\n# changed\n", encoding="utf-8")
            second = builder._environment_identity(
                endpoint.root,
                ref=endpoint.requested_ref,
                tree=endpoint.expected_tree,
                source_closure=endpoint.source_closure,
                overrides={"PATH": str(tool_dir)},
            )
            self.assertNotEqual(first["tools"]["rustc"], second["tools"]["rustc"])
            with self.assertRaises(RUNNER.RunFailure):
                builder._validate_target_accounting(
                    {"path": str(root / "target"), "bytes_before": 0,
                     "bytes_after": builder.MEMORY_LIMIT_BYTES + 1,
                     "hard_limit_bytes": builder.MEMORY_LIMIT_BYTES},
                    "final-write.target",
                    root / "target",
                )

    def test_directory_guard_rejects_missing_and_symlinked_targets(self) -> None:
        from tools import odg_joined_qualification_builder as builder

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target = root / "target"
            target.mkdir()
            with self.assertRaises(RUNNER.RunFailure):
                builder._directory_bytes(root / "missing")
            (target / "link").symlink_to(root)
            with self.assertRaises(RUNNER.RunFailure):
                builder._directory_bytes(target)

    def test_stream_guard_catches_detached_owned_descendant(self) -> None:
        from tools import odg_joined_qualification_builder as builder

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target = root / "target"
            target.mkdir()
            logs = root / "logs"
            logs.mkdir()
            command = [
                sys.executable,
                "-c",
                (
                    "import ctypes, subprocess\n"
                    "def clear():\n"
                    "    p=ctypes.CDLL(None, use_errno=True).prctl\n"
                    "    p(1, 0, 0, 0, 0)\n"
                    "subprocess.Popen(['sleep', '5'], start_new_session=True, preexec_fn=clear)\n"
                ),
            ]
            with mock.patch.object(builder, "process_scan", return_value=[]), \
                    mock.patch.object(builder, "load_average", return_value=0.0):
                with self.assertRaises(RUNNER.RunFailure):
                    builder._stream_command(command, root, logs, "detached", target, 0, timeout_s=5)

    def test_builder_requires_pristine_baseline_and_final_candidate_binding(self) -> None:
        from tools import odg_joined_qualification_builder as builder

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            with self.assertRaises(RUNNER.RunFailure) as failure:
                builder.prepare_plan(manifest)
            self.assertIn("baseline", str(failure.exception))
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = write_manifest(root)
            data = json.loads(manifest.read_text(encoding="utf-8"))
            data["baseline"]["ref"] = RUNNER.REVISION.pattern.strip("\\Z")
            manifest.write_text(json.dumps(data), encoding="utf-8")
            with self.assertRaises(RUNNER.RunFailure):
                builder.prepare_plan(manifest)

    def test_allocation_validator_requires_operation_scope_and_conservation(self) -> None:
        protocol = RUNNER.Protocol(0, RUNNER.REQUIRED_BLOCKS, RUNNER.REQUIRED_WARMUPS,
                                   500_000_000, 150_000_000, 1000, 100, 2)
        expected = {"output": digest(b"out"), "semantic": {"pages": []}}
        allocation = {
            "status": "measured", "scope": "operation_global_system_allocator",
            "allocation_calls": 2, "deallocation_calls": 1, "reallocation_calls": 0,
            "failed_allocation_calls": 0, "allocated_bytes": 16, "deallocated_bytes": 8,
            "live_bytes_before": 4, "live_bytes_after": 12,
            "peak_live_bytes_before": 4, "peak_live_bytes_after": 12,
            "region_peak_live_bytes": 12, "live_underflow": False,
        }
        report = {"schema": RUNNER.BASELINE_SCHEMA, "endpoint": "baseline-apply-joined",
                  "mode": "allocation", "route": "empty_join", "scope": "join_apply",
                  "iterations": 1, "warmups": RUNNER.REQUIRED_WARMUPS,
                  "samples": [{"metrics": {"tool": {
                      "binary": "litchi-perf-odg-joined-qualification-alloc",
                      "instrumentation": "system_allocator_operation_scoped",
                      "allocator_counter_revision": "serialized_region_peak_v3"},
                      "allocation": allocation,
                      "latency_claim": "allocator_instrumented_elapsed_not_latency_claim"},
                      "proof": expected}]}
        self.assertEqual(RUNNER.validate_allocation(report, "baseline", "empty_join", "join_apply",
                                                     protocol, expected)["proof"], expected)
        broken = json.loads(json.dumps(report))
        broken["samples"][0]["metrics"]["allocation"]["deallocated_bytes"] = 0
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_allocation(broken, "baseline", "empty_join", "join_apply", protocol, expected)
        broken = json.loads(json.dumps(report))
        broken["samples"][0]["metrics"]["allocation"]["live_underflow"] = True
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_allocation(broken, "baseline", "empty_join", "join_apply", protocol, expected)
        broken = json.loads(json.dumps(report))
        broken["samples"][0]["metrics"]["allocation"]["peak_live_bytes_after"] = 3
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_allocation(broken, "baseline", "empty_join", "join_apply", protocol, expected)
        broken = json.loads(json.dumps(report))
        broken["samples"][0]["metrics"]["allocation"]["reallocation_calls"] = 3
        with self.assertRaises(RUNNER.RunFailure):
            RUNNER.validate_allocation(broken, "baseline", "empty_join", "join_apply", protocol, expected)
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
