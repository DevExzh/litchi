#!/usr/bin/env python3
"""Source-only tests for the pending signed-native policy v2."""

from __future__ import annotations

import copy
import json
from pathlib import Path
import subprocess
import tempfile
import unittest

import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from native_qualification import prepare_signed_native_inputs_v2 as policy  # noqa: E402


class SignedNativePolicyV2Tests(unittest.TestCase):
    def _git_repo(self, root: Path, files: dict[str, bytes]) -> None:
        subprocess.run(["git", "init", "-q", str(root)], check=True)
        subprocess.run(
            ["git", "-C", str(root), "config", "user.email", "policy@example.invalid"],
            check=True,
        )
        subprocess.run(
            ["git", "-C", str(root), "config", "user.name", "policy test"],
            check=True,
        )
        for relative, payload in files.items():
            path = root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(payload)
        subprocess.run(["git", "-C", str(root), "add", "."], check=True)
        subprocess.run(
            ["git", "-C", str(root), "commit", "-qm", "fixture"], check=True
        )

    def test_checked_in_policy_is_pending_and_execution_fails_closed(self) -> None:
        self.assertEqual(
            policy.validate_policy(),
            {
                "status": "pending_final_source_freeze",
                "execution": "refused_pending_source_freeze",
            },
        )
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-output-") as name:
            output = Path(name) / "must-not-exist"
            with self.assertRaises(policy.IncompletePolicyError):
                policy.execute()
            self.assertFalse(output.exists())

    def test_prepare_fails_before_source_or_output_mutation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-prepare-") as name:
            root = Path(name)
            source = root / "source"
            output = root / "output"
            with self.assertRaisesRegex(policy.IncompletePolicyError, "candidate is incomplete"):
                policy.prepare(source, output)
            self.assertFalse(output.exists())

    def test_policy_cannot_be_activated_by_supplied_candidate_digest(self) -> None:
        document = policy.load_policy()
        changed = copy.deepcopy(document)
        changed["candidate"]["status"] = "bound"
        changed["candidate"]["source"] = {
            "commit": "0" * 40,
            "git_tree": "1" * 40,
            "tree_sha256": "2" * 64,
            "archive": {"path": "/tmp/arbitrary.tar", "bytes": 1, "sha256": "3" * 64},
        }
        with self.assertRaisesRegex(policy.PolicyError, "code binding is still incomplete"):
            policy._validate_policy_document(changed, execution=True)
        # The public loader does not take a path or digest override.
        self.assertNotIn("candidate", policy.execute.__annotations__)
        self.assertIsNone(policy.CANDIDATE_BINDING)

    def test_v1_receipts_are_not_relabelled(self) -> None:
        with self.assertRaisesRegex(policy.PolicyError, "cannot be relabelled"):
            policy.reject_v1_receipt({"schema": "litchi_signed_native_probe.v1"})
        with self.assertRaisesRegex(policy.PolicyError, "cannot be relabelled"):
            policy.reject_v1_receipt({"schema": "litchi_signed_native_qualification.v1"})

    def test_route_and_semantic_contract_covers_current_signature_matrix(self) -> None:
        document = policy.load_policy()
        self.assertEqual(tuple(document["formats"]), policy.NATIVE_FORMATS)
        self.assertEqual(tuple(document["routes"]), policy.ROUTES)
        self.assertEqual(set(document["correctness_only"]), set(policy.CORRECTNESS_ONLY))
        self.assertEqual(
            document["semantics"]["cancellation"],
            "cancelled_intermediate_edit_remains_signed_and_does_not_publish",
        )
        self.assertEqual(
            document["signed_batch"], {"status": "pending_owner_choice", "selection": None}
        )

    def test_fixture_manifest_payloads_and_signature_claims_are_bound(self) -> None:
        result = policy._fixture_policy()
        self.assertEqual(result["schema"], "litchi_signed_fixture_manifest.v1")
        self.assertEqual(result["fixtures"], 11)
        self.assertEqual(set(result["formats"]), set(policy.NATIVE_FORMATS))

        document = json.loads(policy.FIXTURE_MANIFEST.read_text(encoding="utf-8"))
        changed = copy.deepcopy(document)
        changed["fixtures"][0]["relative_path"] = "../outside.bin"
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-fixture-"):
            # Exercise the same path guard used by the fixture validator; no
            # external fixture is ever accepted merely because its digest was
            # supplied by a caller.
            with self.assertRaises(policy.PolicyError):
                policy._safe_relative(changed["fixtures"][0]["relative_path"], "fixture path")

    def test_vendor_tree_covers_all_tracked_blobs_and_rejects_sparse_checkout(self) -> None:
        files = {
            "vendor/allsorts/Cargo.toml": b"allsorts\n",
            "vendor/allsorts/src/lib.rs": b"a\n",
            "vendor/flate2/Cargo.toml": b"flate2\n",
            "vendor/zlib-rs/Cargo.toml": b"zlib-rs\n",
        }
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-vendor-") as name:
            root = Path(name)
            self._git_repo(root, files)
            identity = policy._vendor_tree_identity(root)
            self.assertEqual(identity["files"], len(files))
            self.assertEqual(set(identity["roots"]), set(policy.VENDOR_ROOTS))

            (root / "vendor/flate2/Cargo.toml").unlink()
            with self.assertRaisesRegex(policy.PolicyError, "required regular file is missing"):
                policy._vendor_tree_identity(root)

    def test_vendor_tree_rejects_modified_and_symlinked_tracked_blobs(self) -> None:
        files = {
            "vendor/allsorts/Cargo.toml": b"a\n",
            "vendor/flate2/Cargo.toml": b"b\n",
            "vendor/zlib-rs/Cargo.toml": b"c\n",
        }
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-vendor-mutated-") as name:
            root = Path(name)
            self._git_repo(root, files)
            (root / "vendor/zlib-rs/Cargo.toml").write_bytes(b"changed\n")
            with self.assertRaisesRegex(policy.PolicyError, "tracked vendor blob differs"):
                policy._vendor_tree_identity(root)

            (root / "vendor/zlib-rs/Cargo.toml").unlink()
            (root / "vendor/zlib-rs/Cargo.toml").symlink_to(root / "vendor/allsorts/Cargo.toml")
            with self.assertRaisesRegex(policy.PolicyError, "path contains a symlink"):
                policy._vendor_tree_identity(root)

    def test_bound_binding_requires_full_source_tree_archive_and_vendor_file_list(self) -> None:
        binding = {
            "status": "bound",
            "source_manifests": list(policy.SOURCE_MANIFEST_PATHS),
            "role_locks": policy.ROLE_LOCKS,
            "vendor_roots": list(policy.VENDOR_ROOTS),
            "fixture": policy.FIXTURE_INPUT,
            "probe_inputs": policy.PROBE_INPUTS,
            "toolchain_paths": list(policy.TOOLCHAIN_PATHS),
            "toolchain": policy.TOOLCHAIN_IDENTITIES,
            "source": {
                "commit": "0" * 40,
                "git_tree": "1" * 40,
                "tree_sha256": "2" * 64,
                "archive": {"path": "/tmp/source.tar", "bytes": 1, "sha256": "3" * 64},
            },
            "manifests": {path: {"bytes": 1, "sha256": "4" * 64} for path in policy.SOURCE_MANIFEST_PATHS},
            "vendor_tree": {"files": 1, "bytes": 1, "sha256": "5" * 64, "entries": []},
            "lock_graph": {},
            "features": {"package": "litchi-ops", "all_features": True, "selected": ["ops"]},
        }
        with self.assertRaisesRegex(policy.PolicyError, "vendor file list is incomplete"):
            policy._validate_bound_binding(binding)

    def test_relative_path_guard_rejects_absolute_traversal_and_upstream_paths(self) -> None:
        for value in ("/tmp/input", "../input", "a/../input", "test-data/3rdparty/file"):
            with self.assertRaises(policy.PolicyError):
                policy._safe_relative(value, "fixture")


if __name__ == "__main__":
    unittest.main()
