#!/usr/bin/env python3
"""Source-only tests for the pending signed-native policy v2."""

from __future__ import annotations

import copy
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tarfile
import tempfile
import unittest
from unittest import mock

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

    def test_prepare_rejects_an_unbound_source_before_output_mutation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-prepare-") as name:
            root = Path(name)
            source = root / "source"
            output = root / "output"
            with self.assertRaisesRegex(policy.IncompletePolicyError, "after endpoint is unavailable"):
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
            document["signed_batch"], policy.SIGNED_BATCH_POLICY
        )
        self.assertEqual(policy.ENDPOINT_API_BRANCH, {"before": "legacy_api", "after": "current_api"})
        self.assertEqual(policy.ENDPOINT_SOURCE_KINDS, {"before": "f2", "after": "candidate"})
        self.assertEqual(document["endpoints"]["after"]["source_kind"], "candidate")
        self.assertEqual(document["historical_harness"]["source_kind"], "472b")
        self.assertEqual(
            policy.ENDPOINT_ROUTE_STATUS["before"]["signed_leaf/unsigned_negative"],
            "not_applicable",
        )
        self.assertEqual(
            policy.ENDPOINT_ROUTE_STATUS["before"]["signed_leaf/noop_commit"],
            "capability_missing",
        )
        self.assertEqual(
            policy.ENDPOINT_ROUTE_STATUS["after"]["signed_leaf/unsigned_negative"],
            "expected_refusal",
        )

    def test_signed_batch_choice_cannot_be_overridden_by_policy_input(self) -> None:
        document = copy.deepcopy(policy.load_policy())
        document["signed_batch"] = {
            "status": "owner_confirmed",
            "selection": "resolve_like_edit",
        }
        with self.assertRaisesRegex(policy.PolicyError, "signed-batch owner choice changed"):
            policy._validate_policy_document(document, execution=False)

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
            with self.assertRaisesRegex(policy.PolicyError, "vendor tree contains a symlink"):
                policy._vendor_tree_identity(root)

    def test_vendor_tree_rejects_ignored_overlay_and_untracked_symlink(self) -> None:
        files = {
            ".gitignore": b"vendor/**/*.ignored\n",
            "vendor/allsorts/Cargo.toml": b"a\n",
            "vendor/flate2/Cargo.toml": b"b\n",
            "vendor/zlib-rs/Cargo.toml": b"c\n",
        }
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-vendor-overlay-") as name:
            root = Path(name)
            self._git_repo(root, files)
            ignored = root / "vendor/allsorts/overlay.ignored"
            ignored.write_bytes(b"ignored\n")
            with self.assertRaisesRegex(policy.PolicyError, "untracked or ignored overlay"):
                policy._vendor_tree_identity(root)

            ignored.unlink()
            symlink = root / "vendor/allsorts/overlay-link"
            symlink.symlink_to(root / "vendor/allsorts/Cargo.toml")
            with self.assertRaisesRegex(policy.PolicyError, "vendor tree contains a symlink"):
                policy._vendor_tree_identity(root)

    def _write_git_archive(self, root: Path, archive: Path, prefix: str) -> None:
        entries = policy._tree_entries(root)
        with tarfile.open(archive, mode="w") as stream:
            for entry in entries:
                source = root / entry["path"]
                info = stream.gettarinfo(str(source), arcname=prefix + entry["path"])
                with source.open("rb") as payload:
                    stream.addfile(info, payload)

    def test_archive_file_set_proves_selected_tree_and_rejects_different_tree(self) -> None:
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-archive-") as name:
            root = Path(name) / "source"
            archive = Path(name) / "source.tar"
            self._git_repo(root, {"src/lib.rs": b"original\n", "README": b"read me\n"})
            self._write_git_archive(root, archive, "capture/")
            tree = policy._tree_entries(root)
            archive_identity = policy._artifact_identity(archive, "source archive")
            descriptor = {
                "path": str(archive),
                "bytes": archive_identity["bytes"],
                "sha256": archive_identity["sha256"],
                "format": "git-tree-tar-v1",
                "prefix": "capture/",
                "git_tree": policy._tree_id(root, "HEAD"),
                "tree_sha256": policy._source_tree_digest(tree),
                "ignored_files": [],
            }
            verified = policy._archive_tree_identity(root, "HEAD", descriptor)
            self.assertEqual(verified["tracked_files"], len(tree))

            # The archive digest is updated to the actual different archive,
            # so a hash-only receipt would accept it.  The Git blob proof must
            # still reject it against the selected commit/tree.
            (root / "src/lib.rs").write_bytes(b"different tree\n")
            wrong = Path(name) / "different.tar"
            self._write_git_archive(root, wrong, "capture/")
            wrong_identity = policy._artifact_identity(wrong, "source archive")
            wrong_descriptor = dict(descriptor)
            wrong_descriptor.update(
                {
                    "path": str(wrong),
                    "bytes": wrong_identity["bytes"],
                    "sha256": wrong_identity["sha256"],
                }
            )
            with self.assertRaisesRegex(policy.PolicyError, "blob differs from Git tree"):
                policy._archive_tree_identity(root, "HEAD", wrong_descriptor)

    def test_archive_member_reads_are_bounded_and_bind_ignored_files(self) -> None:
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-streamed-archive-") as name:
            root = Path(name) / "source"
            archive = Path(name) / "source.tar"
            self._git_repo(root, {"src/lib.rs": b"tracked\n"})
            self._write_git_archive(root, archive, "capture/")
            ignored = root / "ignored.lock"
            ignored.write_bytes(b"ignored\n" * 250_000)
            with tarfile.open(archive, mode="a") as stream:
                info = stream.gettarinfo(str(ignored), arcname="capture/ignored.lock")
                with ignored.open("rb") as payload:
                    stream.addfile(info, payload)
            tree = policy._tree_entries(root)
            archive_identity = policy._artifact_identity(archive, "source archive")
            ignored_identity = policy._stream_sha256(ignored)
            descriptor = {
                "path": str(archive),
                "bytes": archive_identity["bytes"],
                "sha256": archive_identity["sha256"],
                "format": "git-tree-tar-v1",
                "prefix": "capture/",
                "git_tree": policy._tree_id(root, "HEAD"),
                "tree_sha256": policy._source_tree_digest(tree),
                "ignored_files": [
                    {
                        "path": "ignored.lock",
                        "mode": "100644",
                        "bytes": ignored_identity[0],
                        "sha256": ignored_identity[1],
                    }
                ],
            }
            original_read = tarfile.ExFileObject.read
            requests: list[int] = []

            def bounded_read(file_stream: object, size: int = -1) -> bytes:
                requests.append(size)
                if size < 0 or size > policy._STREAM_CHUNK_BYTES:
                    raise AssertionError(f"unbounded tar member read: {size}")
                return original_read(file_stream, size)

            with mock.patch.object(tarfile.ExFileObject, "read", bounded_read):
                verified = policy._archive_tree_identity(root, "HEAD", descriptor)
            self.assertEqual(verified["tracked_files"], len(tree))
            self.assertEqual(verified["ignored_files"], ["ignored.lock"])
            self.assertTrue(requests)

    def test_large_sparse_artifact_streams_under_a_memory_cap(self) -> None:
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-sparse-artifact-") as name:
            artifact = Path(name) / "candidate-source.tar"
            size = 192 * 1024 * 1024
            with artifact.open("wb") as payload:
                payload.truncate(size)
            expected_digest = hashlib.sha256()
            zero_chunk = b"\0" * policy._STREAM_CHUNK_BYTES
            remaining = size
            while remaining:
                count = min(remaining, len(zero_chunk))
                expected_digest.update(zero_chunk[:count])
                remaining -= count
            script = """
import importlib.util
import resource
import sys
from pathlib import Path

spec = importlib.util.spec_from_file_location("streaming_policy", sys.argv[3])
assert spec is not None and spec.loader is not None
policy = importlib.util.module_from_spec(spec)
spec.loader.exec_module(policy)

resource.setrlimit(resource.RLIMIT_AS, (128 * 1024 * 1024, 128 * 1024 * 1024))
identity = policy._artifact_identity(Path(sys.argv[1]), "sparse source archive")
assert identity["bytes"] == int(sys.argv[2])
print(identity["sha256"])
"""
            environment = dict(os.environ)
            result = subprocess.run(
                [
                    sys.executable,
                    "-c",
                    script,
                    str(artifact),
                    str(size),
                    os.environ.get(
                        "SIGNED_NATIVE_STREAMING_POLICY",
                        str(Path(__file__).parent / "native_qualification" / "prepare_signed_native_inputs_v2.py"),
                    ),
                ],
                check=True,
                capture_output=True,
                text=True,
                env=environment,
            )
            self.assertEqual(result.stdout.strip(), expected_digest.hexdigest())

    def test_source_archive_output_paths_must_be_fully_disjoint(self) -> None:
        with tempfile.TemporaryDirectory(prefix="signed-native-v2-paths-") as name:
            base = Path(name)
            source = base / "source"
            archive = source / "capture.tar"
            output = base / "output"
            with self.assertRaisesRegex(policy.PolicyError, "paths must be disjoint"):
                policy._assert_disjoint_paths(
                    {"candidate source": source, "candidate archive": archive, "preparation output": output}
                )
            with self.assertRaisesRegex(policy.PolicyError, "paths must be disjoint"):
                policy._assert_disjoint_paths(
                    {"candidate source": source, "candidate archive": base / "archive.tar", "preparation output": base}
                )
            policy._assert_disjoint_paths(
                {"candidate source": source, "candidate archive": base / "archive.tar", "preparation output": output}
            )

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
                "archive": {
                    "path": "/tmp/source.tar",
                    "bytes": 1,
                    "sha256": "3" * 64,
                    "format": "git-tree-tar-v1",
                    "prefix": "capture/",
                    "git_tree": "1" * 40,
                    "tree_sha256": "2" * 64,
                    "ignored_files": [
                        {
                            "path": "Cargo.lock",
                            "bytes": policy.ROLE_LOCKS["workspace"]["candidate"]["bytes"],
                            "sha256": policy.ROLE_LOCKS["workspace"]["candidate"]["sha256"],
                            "mode": "100644",
                        }
                    ],
                },
            },
            "source_assembly": {
                "path": "/tmp/source-assembly.json",
                "bytes": 1,
                "sha256": "6" * 64,
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

    def test_checked_in_candidate_perf_lock_tracks_the_repaired_consumer_lock(self) -> None:
        expected = policy.ROLE_LOCKS["perf"]["candidate"]
        manifest = json.loads(
            (ROOT / "tools/native_qualification/manifests/signed-native-v2.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(manifest["candidate"]["role_locks"]["perf"]["candidate"], expected)


if __name__ == "__main__":
    unittest.main()
