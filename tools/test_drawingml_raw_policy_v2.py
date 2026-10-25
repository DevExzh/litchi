#!/usr/bin/env python3
"""Source-only tests for the fail-closed DrawingML qualification policy v2."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest import mock

import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from native_qualification import prepare_drawingml_raw_inputs_v2 as policy  # noqa: E402


class DrawingMlRawPolicyV2Tests(unittest.TestCase):
    def _git_repo(self, root: Path, files: dict[str, bytes]) -> None:
        subprocess.run(["git", "init", "-q", str(root)], check=True)
        subprocess.run(["git", "-C", str(root), "config", "user.email", "test@example.invalid"], check=True)
        subprocess.run(["git", "-C", str(root), "config", "user.name", "policy test"], check=True)
        for relative, payload in files.items():
            path = root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(payload)
        subprocess.run(["git", "-C", str(root), "add", "."], check=True)
        subprocess.run(["git", "-C", str(root), "commit", "-qm", "fixture"], check=True)

    def test_fixed_historical_blob_identity(self) -> None:
        payload, receipt = policy._trusted_historical_payload(ROOT)
        self.assertEqual(receipt["ref"], policy.HISTORICAL_SOURCE_REF)
        self.assertEqual(receipt["path"], policy.HISTORICAL_SOURCE_PATH)
        self.assertEqual(receipt["blob_sha1"], policy.HISTORICAL_SOURCE_BLOB_SHA1)
        self.assertEqual(receipt["sha256"], policy.HISTORICAL_SOURCE_SHA256)
        self.assertEqual(receipt["bytes"], policy.HISTORICAL_SOURCE_BYTES)
        self.assertEqual(hashlib.sha256(payload).hexdigest(), policy.HISTORICAL_SOURCE_SHA256)

    def test_materialization_uses_trusted_blob_not_candidate_bytes(self) -> None:
        with tempfile.TemporaryDirectory(prefix="dml-policy-v2-baseline-") as name:
            baseline = Path(name)
            self._git_repo(baseline, {"README": b"baseline\n"})
            receipt = policy.materialize_historical_library(ROOT, baseline)
            destination = baseline / policy.HISTORICAL_SOURCE_PATH
            self.assertEqual(receipt["destination"]["sha256"], policy.HISTORICAL_SOURCE_SHA256)
            self.assertEqual(destination.stat().st_size, policy.HISTORICAL_SOURCE_BYTES)
            self.assertEqual(
                hashlib.sha256(destination.read_bytes()).hexdigest(),
                policy.HISTORICAL_SOURCE_SHA256,
            )

    def test_prepare_validates_trusted_blob_before_copying_routes(self) -> None:
        with tempfile.TemporaryDirectory(prefix="dml-policy-v2-order-") as name:
            root = Path(name)
            baseline = root / "baseline"
            candidate = root / "candidate"
            untrusted = root / "untrusted"
            self._git_repo(baseline, {"README": b"baseline\n"})
            self._git_repo(candidate, {"README": b"candidate\n"})
            self._git_repo(untrusted, {"README": b"not the trusted source\n"})
            candidate_revision = subprocess.check_output(
                ["git", "-C", str(candidate), "rev-parse", "HEAD"], text=True
            ).strip()
            fake_binding = {
                "status": "bound",
                "revision": candidate_revision,
                "manifests": dict.fromkeys(policy.CANDIDATE_MANIFEST_PATHS, {}),
                "vendor_tree": {"sha256": "a" * 64, "files": 1, "bytes": 1},
            }
            with (
                mock.patch.object(policy, "CANDIDATE_BINDING", fake_binding),
                mock.patch.object(policy, "_candidate_identity", return_value={}),
                mock.patch.object(policy, "_baseline_identity", return_value={}),
                mock.patch.object(policy, "_copy_probe_paths") as copy_routes,
            ):
                with self.assertRaises(policy.PreparationError):
                    policy.prepare(
                        baseline,
                        candidate,
                        next(iter(policy.APPROVED_BASELINE_REVISIONS)),
                        candidate_revision,
                        untrusted,
                        root / "receipt.json",
                    )
            copy_routes.assert_not_called()
            self.assertFalse((baseline / policy.HISTORICAL_SOURCE_PATH).exists())

    def test_tampered_historical_library_is_rejected_at_identity_validation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="dml-policy-v2-tampered-") as name:
            root = Path(name)
            self._git_repo(
                root,
                {policy.HISTORICAL_SOURCE_PATH: b"tampered historical source\n"},
            )
            with self.assertRaisesRegex(
                policy.PreparationError, "materialized historical library SHA-256 differs"
            ):
                policy._historical_library_identity(root)

    def test_candidate_policy_is_honestly_incomplete_and_fails_before_mutation(self) -> None:
        self.assertIsNone(policy.CANDIDATE_BINDING)
        with self.assertRaisesRegex(policy.PreparationError, "candidate policy is incomplete"):
            policy.prepare(
                Path("/does/not/exist/baseline"),
                Path("/does/not/exist/candidate"),
                next(iter(policy.APPROVED_BASELINE_REVISIONS)),
                "a" * 40,
                ROOT,
                Path("/tmp/dml-policy-v2-unwritten-receipt.json"),
            )

    def test_role_lock_identities_are_separate_and_candidate_values_are_fixed(self) -> None:
        self.assertEqual(set(policy.ROLE_LOCKS), {"workspace", "perf", "native"})
        for spec in policy.ROLE_LOCKS.values():
            self.assertNotEqual(spec["baseline"], spec["candidate"])
            self.assertEqual(len(spec["candidate"]["sha256"]), 64)
            self.assertGreater(spec["candidate"]["bytes"], 0)

    def test_dependency_policy_contains_exact_three_path_transitions(self) -> None:
        self.assertEqual(len(policy.DEPENDENCY_PATH_TRANSITIONS), 3)
        self.assertEqual(
            [item["package"] for item in policy.DEPENDENCY_PATH_TRANSITIONS],
            ["allsorts", "flate2", "zlib-rs"],
        )
        self.assertTrue(
            all(
                item["baseline_source"] == "registry"
                and item["candidate_source"] == "path"
                and item["candidate_path"].startswith("vendor/")
                for item in policy.DEPENDENCY_PATH_TRANSITIONS
            )
        )

    def test_sparse_vendor_checkout_is_rejected_after_git_tree_binding(self) -> None:
        with tempfile.TemporaryDirectory(prefix="dml-policy-v2-vendor-") as name:
            root = Path(name)
            files = {
                "vendor/allsorts/Cargo.toml": b"[package]\nname='allsorts'\n",
                "vendor/flate2/Cargo.toml": b"[package]\nname='flate2'\n",
                "vendor/zlib-rs/Cargo.toml": b"[package]\nname='zlib-rs'\n",
            }
            self._git_repo(root, files)
            identity = policy._vendor_tree_identity(root)
            self.assertEqual(identity["files"], 3)
            (root / "vendor/flate2/Cargo.toml").unlink()
            with self.assertRaisesRegex(policy.PreparationError, "required file is missing"):
                policy._vendor_tree_identity(root)

    def test_ignored_vendor_overlay_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory(prefix="dml-policy-v2-ignored-vendor-") as name:
            root = Path(name)
            self._git_repo(
                root,
                {
                    "vendor/allsorts/Cargo.toml": b"a\n",
                    "vendor/flate2/Cargo.toml": b"b\n",
                    "vendor/zlib-rs/Cargo.toml": b"c\n",
                    ".gitignore": b"vendor/flate2/ignored.txt\n",
                },
            )
            (root / "vendor/flate2/ignored.txt").write_bytes(b"ignored overlay\n")
            with self.assertRaisesRegex(policy.PreparationError, "untracked or ignored overlay"):
                policy._vendor_tree_identity(root)

    def test_vendor_tree_rejects_modified_tracked_blob(self) -> None:
        with tempfile.TemporaryDirectory(prefix="dml-policy-v2-vendor-mutated-") as name:
            root = Path(name)
            self._git_repo(
                root,
                {
                    "vendor/allsorts/Cargo.toml": b"a\n",
                    "vendor/flate2/Cargo.toml": b"b\n",
                    "vendor/zlib-rs/Cargo.toml": b"c\n",
                },
            )
            (root / "vendor/zlib-rs/Cargo.toml").write_bytes(b"changed\n")
            with self.assertRaisesRegex(policy.PreparationError, "tracked vendor blob differs"):
                policy._vendor_tree_identity(root)

    def test_endpoint_and_receipt_paths_must_be_disjoint(self) -> None:
        with tempfile.TemporaryDirectory(prefix="dml-policy-v2-paths-") as name:
            root = Path(name)
            baseline = root / "baseline"
            candidate = root / "candidate"
            baseline.mkdir()
            candidate.mkdir()
            with self.assertRaisesRegex(policy.PreparationError, "paths must be disjoint"):
                policy._assert_disjoint_paths({"baseline": baseline, "candidate": baseline / "nested"})
            with self.assertRaisesRegex(policy.PreparationError, "paths must be disjoint"):
                policy._assert_disjoint_paths({"candidate": candidate, "receipt": candidate / "receipt.json"})

    def test_manifest_declares_pending_v2_policy_without_candidate_hash_claim(self) -> None:
        manifest = json.loads(
            (ROOT / "tools/native_qualification/manifests/drawingml-raw-v2.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(manifest["schema"], "litchi.drawingml.raw-qualification.manifest.v2")
        self.assertEqual(manifest["status"], "pending_final_candidate_binding")
        self.assertIsNone(manifest["candidate"]["revision"])
        self.assertIsNone(manifest["candidate"]["manifests"])
        self.assertIsNone(manifest["candidate"]["vendor_tree"])
        self.assertEqual(
            manifest["historical_source"]["sha256"], policy.HISTORICAL_SOURCE_SHA256
        )
        self.assertEqual(manifest["toolchain_paths"], list(policy.TOOLCHAIN_PATHS))

    def test_candidate_policy_has_no_user_supplied_hash_parameter(self) -> None:
        self.assertNotIn("candidate_workspace_lock_sha256", policy.prepare.__annotations__)
        self.assertNotIn("perf_lock_sha256", policy.prepare.__annotations__)
        self.assertEqual(policy.ROLE_LOCK_PATHS, ("Cargo.lock", "tools/perf-baseline/Cargo.lock", "tools/native-resave/Cargo.lock"))

    def test_versioned_consumers_use_v2_policy_without_relabeling_v1(self) -> None:
        manifest = json.loads(
            (ROOT / "tools/native_qualification/manifests/drawingml-raw-v2.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(
            manifest["runner_module"],
            "tools/native_qualification/drawingml_raw_paired_runner_v2.py",
        )
        self.assertEqual(
            manifest["validator_module"],
            "tools/native_qualification/drawingml_raw_validator_v2.py",
        )
        runner = (ROOT / manifest["runner_module"]).read_text(encoding="utf-8")
        self.assertIn("prepare_drawingml_raw_inputs_v2", runner)
        self.assertNotIn("from native_qualification import prepare_drawingml_raw_inputs as preparation", runner)
        validator = (ROOT / manifest["validator_module"]).read_text(encoding="utf-8")
        self.assertIn("prepare_drawingml_raw_inputs_v2", validator)
        self.assertIn("litchi.drawingml.raw.validation.v2", validator)


if __name__ == "__main__":
    unittest.main()
