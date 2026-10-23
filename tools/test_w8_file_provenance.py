"""Tests for the W8 fixture and dependency provenance generator."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import tempfile
import unittest

from tools import w8_file_provenance as provenance


ROOT = Path(__file__).resolve().parents[1]
PROBE = ROOT / "crates/litchi-ops/examples/file_w8_qualified_probe.rs"
PACKAGE = ROOT / "crates/litchi-ops/Cargo.toml"
TOOLCHAIN = ROOT / "rust-toolchain.toml"


class FixtureDescriptor(unittest.TestCase):
    def test_current_probe_has_the_reviewed_workload(self):
        descriptor = provenance.fixture_descriptor(PROBE)
        self.assertEqual(descriptor["route_count"], 109)
        self.assertEqual(tuple(descriptor["families"]), provenance.EXPECTED_FAMILIES)
        self.assertEqual(descriptor["constants"]["paragraphs"], 1_000)
        self.assertEqual(descriptor["constants"]["shapes"], 100)
        self.assertEqual(descriptor["api"]["commit_extraction"], "into_file")
        self.assertEqual(descriptor["recipe"]["xlsb"]["import_rows"], 100)
        self.assertEqual(descriptor["recipe"]["odp"]["batch"], 100)

    def test_descriptor_bytes_are_deterministic(self):
        descriptor = provenance.fixture_descriptor(PROBE)
        with tempfile.TemporaryDirectory() as tmp:
            first = Path(tmp) / "first.json"
            second = Path(tmp) / "second.json"
            provenance.write_new(first, descriptor)
            provenance.write_new(second, descriptor)
            self.assertEqual(first.read_bytes(), second.read_bytes())
            self.assertEqual(provenance.digest(first), provenance.digest(second))

    def test_changed_workload_constant_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            altered = Path(tmp) / PROBE.name
            altered.write_text(
                PROBE.read_text(encoding="utf-8").replace(
                    "const PARAGRAPHS: usize = 1_000;",
                    "const PARAGRAPHS: usize = 999;",
                    1,
                ),
                encoding="utf-8",
            )
            with self.assertRaises(provenance.ProvenanceError):
                provenance.fixture_descriptor(altered)


class DependencyProjection(unittest.TestCase):
    def projection_args(self, root: Path, lock: Path, fixture: Path, probe: Path) -> argparse.Namespace:
        return argparse.Namespace(
            workspace=root,
            package_manifest=root / "crates/litchi-ops/Cargo.toml",
            lock=lock,
            toolchain=root / "rust-toolchain.toml",
            probe=probe,
            fixture=fixture,
            features=["docx,xlsx"],
            all_features=False,
            target="x86_64-unknown-linux-gnu",
            profile="release",
        )

    def test_projection_contains_real_lock_graph_and_workload_hashes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package_dir = root / "crates/litchi-ops"
            package_dir.mkdir(parents=True)
            (root / "Cargo.toml").write_text(
                "[workspace]\nmembers = [\"crates/*\"]\n[workspace.dependencies]\nserde = \"1\"\n",
                encoding="utf-8",
            )
            (root / "rust-toolchain.toml").write_text(
                "[toolchain]\nchannel = \"1.95.0\"\ncomponents = [\"clippy\"]\n",
                encoding="utf-8",
            )
            (package_dir / "Cargo.toml").write_text(
                "[package]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\n[features]\ndocx = []\nxlsx = []\n[dependencies]\nserde = { workspace = true }\n",
                encoding="utf-8",
            )
            lock = root / "Cargo.lock"
            lock.write_text(
                "version = 4\n\n[[package]]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\ndependencies = [\"serde\"]\n\n[[package]]\nname = \"serde\"\nversion = \"1.0.0\"\nsource = \"registry+https://example.invalid\"\nchecksum = \"" + "0" * 64 + "\"\n",
                encoding="utf-8",
            )
            fixture = root / "fixture.json"
            probe = root / "probe.rs"
            probe.write_text("probe", encoding="utf-8")
            fixture.write_text(
                json.dumps({
                    "schema": "litchi_w8_fixture_descriptor.v1",
                    "probe": {"sha256": provenance.digest(probe)},
                    "route_count": 109,
                }),
                encoding="utf-8",
            )
            projection = provenance.dependency_projection(self.projection_args(root, lock, fixture, probe))
            self.assertEqual(projection["lock"]["packages"][0]["name"], "litchi-ops")
            self.assertEqual(projection["features"]["selected"], ["docx", "xlsx"])
            self.assertEqual(projection["workload"]["route_count"], 109)

    def test_projection_rejects_fixture_for_another_probe(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package_dir = root / "crates/litchi-ops"
            package_dir.mkdir(parents=True)
            (root / "Cargo.toml").write_text("[workspace]\n", encoding="utf-8")
            (root / "rust-toolchain.toml").write_text("[toolchain]\nchannel = \"1.95.0\"\n", encoding="utf-8")
            (package_dir / "Cargo.toml").write_text(
                "[package]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\n[features]\ndocx = []\n",
                encoding="utf-8",
            )
            lock = root / "Cargo.lock"
            lock.write_text(
                "version = 4\n\n[[package]]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\n",
                encoding="utf-8",
            )
            probe = root / "probe.rs"
            fixture = root / "fixture.json"
            probe.write_text("actual", encoding="utf-8")
            fixture.write_text(json.dumps({
                "schema": "litchi_w8_fixture_descriptor.v1",
                "probe": {"sha256": "0" * 64},
            }), encoding="utf-8")
            with self.assertRaises(provenance.ProvenanceError):
                provenance.dependency_projection(self.projection_args(root, lock, fixture, probe))


if __name__ == "__main__":
    unittest.main()
