#!/usr/bin/env python3
"""Fail-closed unit tests for the ODF shared-triad protocol."""

from __future__ import annotations

import copy
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from native_qualification import odf_shared_triad_runner as runner  # noqa: E402


def _side(changed: bool) -> dict[str, object]:
    return {
        "changed": changed,
        "patch_noop": not changed,
        "patch_base_match": True,
        "diagnostics_changed": changed,
        "source_unchanged": True,
        "output_differs_from_source": changed,
        "publication_repeat_exact": True,
        "forward_replay_exact": True,
        "inverse_replay_restores_source": True,
        "source_bytes": 1,
        "output_bytes": 2 if changed else 1,
        "output_sha256": ("2" if changed else "1") * 64,
    }


def _member_identity(name: str = "meta.xml") -> dict[str, object]:
    return {
        "name": name,
        "name_raw_sha256": "b" * 64,
        "content_sha256": "a" * 64,
        "content_bytes": 1,
        "compression": "Stored",
        "compressed_bytes": 1,
        "crc32": "00000000",
        "version_needed": 10,
        "flags": 0,
        "compression_options": "Stored",
        "external_attributes": 0,
        "is_utf8": False,
        "using_data_descriptor": False,
        "large_file": False,
        "central_extra_sha256": None,
        "last_modified": None,
        "extra_sha256": None,
        "comment": "",
        "unix_mode": None,
        "version_made_by": [0, 0],
        "encrypted": False,
        "content_text": "<meta />",
    }


def _output(manifest: dict[str, object]) -> dict[str, object]:
    rows = []
    for route in manifest["routes"]:  # type: ignore[index]
        family = route["family"]
        odm = family == "odm"
        audit = None
        if odm:
            audit = {
                "status": "audited",
                "decision": "byte_exact",
                "allowed": True,
                "archive_comment_equal": True,
                "leaf_archive_comment_sha256": "c" * 64,
                "facade_archive_comment_sha256": "c" * 64,
                "manifest_semantics_equal": True,
                "leaf_manifest_semantics": {
                    "root_attributes": [["manifest:version", "1.0"]],
                    "ordered_entries": [[
                        ["manifest:full-path", "/"],
                        ["manifest:media-type", "application/vnd.oasis.opendocument.text"],
                    ], [
                        ["manifest:full-path", "meta.xml"],
                        ["manifest:media-type", "text/xml"],
                    ]],
                    "normalized_entries": [[
                        ["manifest:full-path", "/"],
                        ["manifest:media-type", "application/vnd.oasis.opendocument.text"],
                    ], [
                        ["manifest:full-path", "meta.xml"],
                        ["manifest:media-type", "text/xml"],
                    ]],
                    "other_events": [],
                },
                "facade_manifest_semantics": {
                    "root_attributes": [["manifest:version", "1.0"]],
                    "ordered_entries": [[
                        ["manifest:full-path", "/"],
                        ["manifest:media-type", "application/vnd.oasis.opendocument.text"],
                    ], [
                        ["manifest:full-path", "meta.xml"],
                        ["manifest:media-type", "text/xml"],
                    ]],
                    "normalized_entries": [[
                        ["manifest:full-path", "/"],
                        ["manifest:media-type", "application/vnd.oasis.opendocument.text"],
                    ], [
                        ["manifest:full-path", "meta.xml"],
                        ["manifest:media-type", "text/xml"],
                    ]],
                    "other_events": [],
                },
                "member_order_exact": True,
                "manifest_listing_tracks_order": True,
                "leaf_manifest_tracks_order": True,
                "facade_manifest_tracks_order": True,
                "archive_member_order_and_manifest_listing": False,
                "non_meta_content_exact": True,
                "non_manifest_content_exact": True,
                "non_meta_stored_metadata_exact": True,
                "non_manifest_stored_metadata_exact": True,
                "meta_stored_metadata_exact": True,
                "non_manifest_archive_record_exact": True,
                "meta_archive_record_exact": True,
                "leaf_order": ["meta.xml"],
                "facade_order": ["meta.xml"],
                "content_differences": [],
                "leaf_member_count": 1,
                "facade_member_count": 1,
                "members": [
                    {
                        "name": "meta.xml",
                        "facade_name": "meta.xml",
                        "content_equal": True,
                        "stored_metadata_equal": True,
                        "archive_record_equal": True,
                        "leaf": _member_identity(),
                        "facade": _member_identity(),
                    }
                ],
            }
        rows.append(
            {
                "family": family,
                "route": route["id"],
                "source": route["source"],
                "descriptor": route,
                "source_bytes": 1,
                "source_sha256": "1" * 64,
                "facade_source_bytes": 1,
                "facade_source_sha256": "1" * 64,
                "source_pair_exact": True,
                "leaf": {"changed": _side(True), "no_op": _side(False)},
                "facade": {"changed": _side(True), "no_op": _side(False)},
                "paired": {
                    "changed_output_exact": True,
                    "changed_output_sha256": "2" * 64,
                    "facade_changed_output_sha256": "2" * 64,
                    "semantic_equality": True,
                    "archive_audit": audit,
                },
                "timing": {
                    "status": "not_run",
                    "setup_outside_clock": list(runner.SETUP_PHASES),
                    "clock_phases": list(runner.PHASES),
                    "max_iterations": 128,
                },
            }
        )
    return {
        "schema": runner.OUTPUT_SCHEMA,
        "mode": "validate",
        "iterations": 1,
        "families": rows,
        "claim_scope": runner.CLAIM_SCOPE,
    }


class OdfSharedTriadRunnerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest, _ = runner.load_manifest(
            ROOT / "tools/native_qualification/manifests/odf-shared-triad.json"
        )

    def test_manifest_is_unresolved_until_final_candidate(self) -> None:
        self.assertIsNone(self.manifest["candidate_commit"])
        stale = copy.deepcopy(self.manifest)
        stale["candidate_commit"] = "15ab9324ef4f8c927acd6ecfc19d9707c3f886ee"
        with self.assertRaisesRegex(runner.RunnerError, "historical"):
            runner.validate_manifest(stale)

    def test_manifest_mutations_fail_closed(self) -> None:
        for field, value, message in (
            ("probe", {"features": ["odc"]}, "features"),
            ("routes", [], "routes"),
            ("phases", {"inside_clock": ["publication"]}, "inside_clock"),
        ):
            mutated = copy.deepcopy(self.manifest)
            if field == "probe":
                mutated["probe"].update(value)
            else:
                mutated[field].update(value) if isinstance(mutated[field], dict) else mutated.__setitem__(field, value)
            with self.assertRaisesRegex(runner.RunnerError, message):
                runner.validate_manifest(mutated)

    def test_probe_output_booleans_and_bindings_are_required(self) -> None:
        payload = _output(self.manifest)
        runner.validate_probe_output(self.manifest, payload)
        broken = copy.deepcopy(payload)
        broken["families"][0]["leaf"]["changed"]["forward_replay_exact"] = False
        with self.assertRaisesRegex(runner.RunnerError, "forward_replay_exact"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][1]["source_pair_exact"] = False
        with self.assertRaisesRegex(runner.RunnerError, "source bytes"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][1]["facade_source_sha256"] = "f" * 64
        with self.assertRaisesRegex(runner.RunnerError, "source_pair_exact"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][2]["paired"]["semantic_equality"] = False
        with self.assertRaisesRegex(runner.RunnerError, "semantic_equality"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][3]["paired"]["archive_audit"]["content_differences"] = ["content.xml"]
        with self.assertRaisesRegex(runner.RunnerError, "content difference"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][3]["paired"]["archive_audit"]["members"][0]["archive_record_equal"] = False
        with self.assertRaisesRegex(runner.RunnerError, "archive_record_equal"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][3]["paired"]["archive_audit"]["facade_order"] = ["other.xml"]
        with self.assertRaisesRegex(runner.RunnerError, "member names differ"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][3]["paired"]["archive_audit"]["decision"] = "metadata_member_only"
        with self.assertRaisesRegex(runner.RunnerError, "decision is unknown"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][3]["paired"]["archive_audit"]["facade_manifest_semantics"]["root_attributes"] = [
            ["manifest:version", "9.9"]
        ]
        with self.assertRaisesRegex(runner.RunnerError, "manifest root_attributes differs"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][3]["paired"]["archive_audit"]["members"][0]["facade"]["external_attributes"] = 1
        with self.assertRaisesRegex(runner.RunnerError, "unadmitted fields"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][3]["paired"]["archive_audit"]["facade_archive_comment_sha256"] = "f" * 64
        with self.assertRaisesRegex(runner.RunnerError, "archive comments differ"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][0]["paired"]["changed_output_exact"] = False
        with self.assertRaisesRegex(runner.RunnerError, "changed_output_exact"):
            runner.validate_probe_output(self.manifest, broken)

    def test_source_tree_probe_and_lock_are_bound(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-shared-runner-") as directory:
            root = Path(directory)
            (root / "crates/litchi-ops/examples").mkdir(parents=True)
            shutil.copy2(ROOT / "crates/litchi-ops/Cargo.toml", root / "crates/litchi-ops/Cargo.toml")
            probe = root / "crates/litchi-ops/examples/odf_shared_triad_qualification.rs"
            shutil.copy2(ROOT / "crates/litchi-ops/examples/odf_shared_triad_qualification.rs", probe)
            lock = root / "Cargo.lock"
            lock.write_bytes(b"synthetic lock identity\n")
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "config", "user.email", "test@example.invalid"], check=True)
            subprocess.run(["git", "-C", str(root), "config", "user.name", "runner-test"], check=True)
            subprocess.run(["git", "-C", str(root), "add", "."], check=True)
            subprocess.run(["git", "-C", str(root), "commit", "-qm", "fixture"], check=True)
            commit = subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip()
            lock_sha = hashlib.sha256(lock.read_bytes()).hexdigest()
            probe_sha = hashlib.sha256(probe.read_bytes()).hexdigest()
            manifest = copy.deepcopy(self.manifest)
            manifest["baseline_commit"] = commit
            manifest["probe"]["sha256"] = probe_sha
            manifest["source"]["lock_sha256"] = lock_sha
            identity = runner.source_identity(
                manifest,
                role="baseline",
                source_root=root,
                source_commit=None,
                lock_sha256=lock_sha,
            )
            self.assertEqual(identity["commit"], commit)
            self.assertGreater(identity["tracked_closure"]["file_count"], 0)
            self.assertIn("sha256", identity["cargo_configs"])
            subprocess.run(
                ["git", "-C", str(root), "update-index", "--skip-worktree", "Cargo.lock"],
                check=True,
            )
            with self.assertRaisesRegex(runner.RunnerError, "sparse/skip-worktree"):
                runner.source_identity(
                    manifest,
                    role="baseline",
                    source_root=root,
                    source_commit=None,
                    lock_sha256=lock_sha,
                )
            subprocess.run(
                ["git", "-C", str(root), "update-index", "--no-skip-worktree", "Cargo.lock"],
                check=True,
            )
            config_target = root.parent / f"cargo-config-{root.name}"
            config_target.write_text("[net]\noffline = true\n", encoding="utf-8")
            config_link = root.parent / f"cargo-config-link-{root.name}"
            config_link.symlink_to(config_target)
            with mock.patch.dict(os.environ, {"CARGO_CONFIG": str(config_link)}):
                with self.assertRaisesRegex(runner.RunnerError, "symlink"):
                    runner.source_identity(
                        manifest,
                        role="baseline",
                        source_root=root,
                        source_commit=None,
                        lock_sha256=lock_sha,
                    )
            with self.assertRaisesRegex(runner.RunnerError, "lock"):
                runner.source_identity(
                    manifest,
                    role="baseline",
                    source_root=root,
                    source_commit=None,
                    lock_sha256="f" * 64,
                )

    def test_order_only_odm_audit_requires_exact_manifest_delta(self) -> None:
        payload = _output(self.manifest)
        row = payload["families"][3]
        audit = row["paired"]["archive_audit"]
        manifest_leaf = _member_identity("META-INF/manifest.xml")
        manifest_facade = copy.deepcopy(manifest_leaf)
        manifest_facade.update(
            {
                "content_sha256": "d" * 64,
                "content_bytes": 2,
                "compressed_bytes": 2,
                "crc32": "11111111",
                "content_text": "<manifest><file-entry /></manifest>",
            }
        )
        audit.update(
            {
                "decision": "archive_member_order_and_manifest_listing",
                "non_meta_content_exact": False,
                "non_meta_stored_metadata_exact": False,
                "member_order_exact": False,
                "manifest_listing_tracks_order": False,
                "archive_member_order_and_manifest_listing": True,
                "leaf_order": ["meta.xml", "content.xml", "META-INF/manifest.xml"],
                "facade_order": ["content.xml", "meta.xml", "META-INF/manifest.xml"],
                "content_differences": ["META-INF/manifest.xml"],
                "leaf_member_count": 3,
                "facade_member_count": 3,
                "leaf_manifest_semantics": {
                    "root_attributes": [["manifest:version", "1.0"]],
                    "ordered_entries": [
                        [["manifest:full-path", "/"]],
                        [["manifest:full-path", "meta.xml"]],
                        [["manifest:full-path", "content.xml"]],
                    ],
                    "normalized_entries": [
                        [["manifest:full-path", "/"]],
                        [["manifest:full-path", "content.xml"]],
                        [["manifest:full-path", "meta.xml"]],
                    ],
                    "other_events": [],
                },
                "facade_manifest_semantics": {
                    "root_attributes": [["manifest:version", "1.0"]],
                    "ordered_entries": [
                        [["manifest:full-path", "/"]],
                        [["manifest:full-path", "content.xml"]],
                        [["manifest:full-path", "meta.xml"]],
                    ],
                    "normalized_entries": [
                        [["manifest:full-path", "/"]],
                        [["manifest:full-path", "content.xml"]],
                        [["manifest:full-path", "meta.xml"]],
                    ],
                    "other_events": [],
                },
                "members": [
                    {
                        "name": "meta.xml",
                        "facade_name": "meta.xml",
                        "content_equal": True,
                        "stored_metadata_equal": True,
                        "archive_record_equal": True,
                        "leaf": _member_identity("meta.xml"),
                        "facade": _member_identity("meta.xml"),
                    },
                    {
                        "name": "content.xml",
                        "facade_name": "content.xml",
                        "content_equal": True,
                        "stored_metadata_equal": True,
                        "archive_record_equal": True,
                        "leaf": _member_identity("content.xml"),
                        "facade": _member_identity("content.xml"),
                    },
                    {
                        "name": "META-INF/manifest.xml",
                        "facade_name": "META-INF/manifest.xml",
                        "content_equal": False,
                        "stored_metadata_equal": False,
                        "archive_record_equal": False,
                        "leaf": manifest_leaf,
                        "facade": manifest_facade,
                    },
                ],
            }
        )
        row["leaf"]["changed"]["output_sha256"] = "4" * 64
        row["facade"]["changed"]["output_sha256"] = "5" * 64
        row["paired"]["changed_output_exact"] = False
        row["paired"]["changed_output_sha256"] = "4" * 64
        row["paired"]["facade_changed_output_sha256"] = "5" * 64
        runner.validate_probe_output(self.manifest, payload)

        broken = copy.deepcopy(payload)
        broken_audit = broken["families"][3]["paired"]["archive_audit"]
        broken_audit["leaf_manifest_semantics"]["ordered_entries"] = [
            [["manifest:full-path", "/"]],
            [["manifest:full-path", "content.xml"]],
            [["manifest:full-path", "meta.xml"]],
        ]
        broken_audit["facade_manifest_semantics"]["ordered_entries"] = [
            [["manifest:full-path", "/"]],
            [["manifest:full-path", "meta.xml"]],
            [["manifest:full-path", "content.xml"]],
        ]
        normalized = [
            [["manifest:full-path", "/"]],
            [["manifest:full-path", "content.xml"]],
            [["manifest:full-path", "meta.xml"]],
        ]
        broken_audit["leaf_manifest_semantics"]["normalized_entries"] = copy.deepcopy(normalized)
        broken_audit["facade_manifest_semantics"]["normalized_entries"] = copy.deepcopy(normalized)
        with self.assertRaisesRegex(runner.RunnerError, "leaf_manifest_tracks_order must be false"):
            runner.validate_probe_output(self.manifest, broken)

        broken = copy.deepcopy(payload)
        broken["families"][3]["paired"]["archive_audit"]["leaf_manifest_tracks_order"] = False
        with self.assertRaisesRegex(runner.RunnerError, "leaf_manifest_tracks_order must be true"):
            runner.validate_probe_output(self.manifest, broken)

        audit["members"][2]["facade"]["flags"] = 1
        with self.assertRaisesRegex(runner.RunnerError, "unadmitted fields"):
            runner.validate_probe_output(self.manifest, payload)

    def test_receipt_tree_identity_is_a_git_tree_sha(self) -> None:
        receipt = runner.make_receipt(
            b"manifest",
            {
                "clean": True,
                "commit": "a" * 40,
                "tree": "b" * 40,
                "package_manifest": {
                    "path": "/source/crates/litchi-ops/Cargo.toml",
                    "sha256": "d" * 64,
                    "name": "litchi-ops",
                },
                "source_tree_sha256": "e" * 64,
                "tracked_closure": {
                    "sha256": "f" * 64,
                    "file_count": 1,
                    "flags_sha256": "1" * 64,
                    "index_sha256": "2" * 64,
                },
                "cargo_configs": {"sha256": "3" * 64, "files": []},
                "root": "/source",
                "probe": {
                    "path": "/source/crates/litchi-ops/examples/odf_shared_triad_qualification.rs",
                    "sha256": self.manifest["probe"]["sha256"],
                },
                "lock": {"path": "/source/Cargo.lock", "sha256": "c" * 64},
            },
            None,
        )
        runner.validate_receipt_bindings(self.manifest, receipt)
        receipt["source"]["tree"] = "d" * 64
        with self.assertRaisesRegex(runner.RunnerError, "commit/tree"):
            runner.validate_receipt_bindings(self.manifest, receipt)


if __name__ == "__main__":
    unittest.main()
