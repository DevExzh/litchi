#!/usr/bin/env python3
"""Fixture-only tests for the 0496 immutable seal helper."""

from __future__ import annotations

import json
import os
from types import SimpleNamespace
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch


import seal


class SealFixtureTests(unittest.TestCase):
    def setUp(self) -> None:
        self.holder = tempfile.TemporaryDirectory(prefix="litchi-seal-0496-")
        self.root = Path(self.holder.name) / "evidence"
        self.repo = Path(self.holder.name) / "repo"
        self.root.mkdir()
        (self.repo / "docs/adr").mkdir(parents=True)

    def tearDown(self) -> None:
        self.holder.cleanup()

    def test_inventory_excludes_seal_and_transient_bytecode(self) -> None:
        (self.root / "README.md").write_text("evidence\n", encoding="utf-8")
        (self.root / "seal.json").write_text("future\n", encoding="utf-8")
        transient = self.root / "__pycache__"
        transient.mkdir()
        (transient / "helper.pyc").write_bytes(b"bytecode")
        inventory = seal._inventory(self.root)
        self.assertEqual(set(inventory), {"README.md"})

    def test_inventory_rejects_links_special_files_and_external_bytecode(self) -> None:
        outside = Path(self.holder.name) / "outside"
        outside.write_text("outside\n", encoding="utf-8")
        (self.root / "link").symlink_to(outside)
        with self.assertRaises(seal.SealError):
            seal._inventory(self.root)
        (self.root / "link").unlink()
        (self.root / "unexpected.pyc").write_bytes(b"bytecode")
        with self.assertRaises(seal.SealError):
            seal._inventory(self.root)
        (self.root / "unexpected.pyc").unlink()
        fifo = self.root / "fifo"
        import os
        os.mkfifo(fifo)
        with self.assertRaises(seal.SealError):
            seal._inventory(self.root)

    def test_adr_binding_checks_every_recorded_hash(self) -> None:
        adr = self.repo / "docs/adr/0001-test.md"
        adr.write_text("accepted\n", encoding="utf-8")
        refresh = self.root / "adr-refresh.json"
        digest = seal._sha(adr)
        refresh.write_text(json.dumps({"reference": "fixture", "files": {"docs/adr/0001-test.md": digest}}) + "\n", encoding="utf-8")
        binding = seal._adr_binding(self.root, self.repo)
        self.assertEqual(binding["files"]["docs/adr/0001-test.md"], digest)
        adr.write_text("changed\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._adr_binding(self.root, self.repo)

    def test_primary_binding_includes_required_rust_and_manifest_hashes(self) -> None:
        source = self.repo / "src/primary.rs"
        source.parent.mkdir(parents=True)
        source.write_text("fn main() {}\n", encoding="utf-8")
        refresh = self.root / "protected-primary.json"
        refresh.write_text(json.dumps({"src/primary.rs": seal._sha(source)}) + "\n", encoding="utf-8")
        binding = seal._primary_binding(self.root, self.repo, paths=("src/primary.rs",))
        self.assertEqual(binding["files"]["src/primary.rs"]["sha256"], seal._sha(source))
        source.write_text("fn changed() {}\n", encoding="utf-8")
        changed = seal._primary_binding(self.root, self.repo, paths=("src/primary.rs",))
        self.assertNotEqual(binding["files"]["src/primary.rs"]["sha256"], changed["files"]["src/primary.rs"]["sha256"])

    def test_retained_tree_requires_only_four_authenticated_executables(self) -> None:
        temporary = self.root / "scratch"
        retained = temporary / "retained"
        builds = {}
        for phase in ("before", "after"):
            for role, filename in (("normal", "litchi-perf-baseline"), ("allocator", "litchi-perf-baseline-alloc")):
                path = retained / phase / role / filename
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(f"{phase}/{role}\n".encode())
                os.chmod(path, 0o755)
                builds[f"{phase}/{role}"] = {"binary": {"path": str(path), **seal._meta(path, "fixture binary")}}
        binding = seal._retained_tree(SimpleNamespace(TEMP=temporary), builds)
        self.assertEqual(set(binding), set(builds))
        (retained / "unexpected.txt").write_text("extra\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._retained_tree(SimpleNamespace(TEMP=temporary), builds)

    def test_projection_cleanup_removes_only_empty_owned_directory(self) -> None:
        capture = SimpleNamespace(TEMP=self.root / "scratch")
        projection = capture.TEMP / "projections"
        projection.mkdir(parents=True)
        seal._remove_empty_projection_parent(capture)
        self.assertFalse(projection.exists())
        projection.mkdir(parents=True)
        marker = projection / "unexpected"
        marker.write_text("keep\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._remove_empty_projection_parent(capture)
        self.assertTrue(marker.exists())

    def test_manifest_is_exclusive(self) -> None:
        (self.root / "seal.json").write_text("existing\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._manifest(self.root)

    def test_manifest_and_verification_timestamps_are_timezone_bound(self) -> None:
        (self.root / "README.md").write_text("fixture evidence\n", encoding="utf-8")
        evidence = {
            "protocol": {}, "builds": {}, "retained_binaries": {}, "formal": {},
            "analysis": {}, "verification": {}, "cleanup": {}, "primary_rust": {},
            "adr": {},
        }
        with patch.object(seal, "_evidence", return_value=evidence):
            manifest = seal._manifest(self.root)
        seal._timestamp(manifest["sealed_utc"], "sealed_utc")
        self.assertIn("README.md", manifest["files"])
        with self.assertRaises(seal.SealError):
            seal._timestamp("2026-09-10T12:00:00", "verified_utc")
        with self.assertRaises(seal.SealError):
            seal._timestamp("not-a-timestamp", "verified_utc")


if __name__ == "__main__":
    unittest.main()
