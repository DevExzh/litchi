#!/usr/bin/env python3
"""Mutation tests for the source-only OLE owner preparation protocol."""

from __future__ import annotations

import hashlib
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest import mock

from tools.native_qualification import prepare_ole_owner_inputs as preparation


PROBE_MARKERS = "\n".join(preparation.REQUIRED_PROBE_MARKERS) + "\n"
CLONE_MARKERS = "\n".join(preparation.REQUIRED_CLONE_MARKERS) + "\n"
ORACLE_MARKERS = "fn check_ole_outputs\nallocator_evidence\nowner_target\n"


def _git(root: Path, *arguments: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *arguments],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return result.stdout.strip()


def _source(root: Path, *, marker: str) -> str:
    root.mkdir()
    _git(root, "init", "--quiet")
    _git(root, "config", "user.email", "ole-owner-protocol@test.invalid")
    _git(root, "config", "user.name", "OLE owner protocol test")
    (root / "source.txt").write_text(marker + "\n", encoding="utf-8")
    cfb_source = root / preparation.CFB_SOURCE_PATH
    cfb_source.parent.mkdir(parents=True)
    if marker == "baseline":
        cfb_source.write_text(
            "pub fn open(reader: R) -> Result<Self, OleError> {\n"
            "pub fn open_with_limits(mut reader: R, limits: OleFileLimits) {\n",
            encoding="utf-8",
        )
    else:
        cfb_source.write_text(
            "pub fn open(mut reader: R, limits: OleFileLimits) -> Result<Self, OleError> {\n",
            encoding="utf-8",
        )
    _git(root, "add", ".")
    _git(root, "commit", "--quiet", "-m", "source")
    return _git(root, "rev-parse", "HEAD")


class PrepareOleOwnerInputsTests(unittest.TestCase):
    def _prepare(self, root: Path) -> tuple[Path, Path, Path, Path, dict[str, str]]:
        baseline = root / "baseline"
        candidate = root / "candidate"
        baseline_ref = _source(baseline, marker="baseline")
        candidate_ref = _source(candidate, marker="candidate")
        candidate_probe = candidate / preparation.PROBE_PATH
        candidate_probe.parent.mkdir(parents=True)
        candidate_probe.write_text(
            PROBE_MARKERS
            + "OleFile::open(\n        Cursor::new(package.into_inner()),\n        OleFileLimits::default(),\n"
            + "OleFileLimits::default()\n",
            encoding="utf-8",
        )
        _git(candidate, "add", ".")
        _git(candidate, "commit", "--quiet", "-m", "reviewed adapter")
        candidate_ref = _git(candidate, "rev-parse", "HEAD")
        reviewed = root / "reviewed.rs"
        reviewed.write_text(PROBE_MARKERS + "OleFile::open(Cursor::new(package.into_inner()))\n", encoding="utf-8")
        clone = root / "clone.rs"
        clone.write_text(CLONE_MARKERS, encoding="utf-8")
        reviewed_lock = root / "reviewed.Cargo.lock"
        reviewed_lock.write_text(
            "version = 4\n\n"
            "[[package]]\nname = \"bitflags\"\nversion = \"2.0.0\"\n"
            "source = \"registry+https://example.invalid/index\"\n"
            "checksum = \"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\"\n\n"
            "[[package]]\nname = \"sha2\"\nversion = \"0.11.0\"\n"
            "source = \"registry+https://example.invalid/index\"\n"
            "checksum = \"bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb\"\n\n"
            "[[package]]\nname = \"litchi-core\"\nversion = \"0.0.1\"\n\n"
            "[[package]]\nname = \"litchi-odraw\"\nversion = \"0.0.1\"\ndependencies = [\n \"bitflags\",\n]\n\n"
            "[[package]]\nname = \"litchi-perf-baseline\"\nversion = \"0.1.0\"\ndependencies = [\n \"litchi-odraw\",\n]\n",
            encoding="utf-8",
        )
        baseline_lock = root / "baseline.Cargo.lock"
        baseline_lock.write_text(
            reviewed_lock.read_text(encoding="utf-8").replace(
                ' "litchi-odraw",\n]\n', ' "litchi-odraw",\n "sha2 0.11.0",\n]\n'
            ),
            encoding="utf-8",
        )
        candidate_lock = root / "candidate.Cargo.lock"
        candidate_lock.write_text(
            baseline_lock.read_text(encoding="utf-8").replace(
                'name = "litchi-odraw"\nversion = "0.0.1"\ndependencies = [\n "bitflags",\n]',
                'name = "litchi-odraw"\nversion = "0.0.1"\ndependencies = [\n "bitflags",\n "litchi-core",\n]'
            ),
            encoding="utf-8",
        )
        oracle = root / "ole_owner_oracle.rs"
        oracle.write_text(ORACLE_MARKERS, encoding="utf-8")
        drivers = root / "drivers"
        receipt = root / "preparation.json"
        values = {
            "BASELINE_COMMIT": baseline_ref,
            "REVIEWED_PROBE_SHA256": hashlib.sha256(reviewed.read_bytes()).hexdigest(),
            "CANDIDATE_ADAPTER_SHA256": hashlib.sha256(candidate_probe.read_bytes()).hexdigest(),
            "CLONE_PROBE_SHA256": hashlib.sha256(clone.read_bytes()).hexdigest(),
            "ORACLE_SOURCE_SHA256": hashlib.sha256(oracle.read_bytes()).hexdigest(),
            "REVIEWED_LOCK_SHA256": hashlib.sha256(reviewed_lock.read_bytes()).hexdigest(),
            "BASELINE_LOCK_SHA256": hashlib.sha256(baseline_lock.read_bytes()).hexdigest(),
            "CANDIDATE_LOCK_SHA256": hashlib.sha256(candidate_lock.read_bytes()).hexdigest(),
        }
        with mock.patch.multiple(preparation, **values):
            preparation.prepare(
                baseline_root=baseline,
                candidate_root=candidate,
                candidate_ref=candidate_ref,
                reviewed_probe=reviewed,
                clone_probe=clone,
                oracle_probe=oracle,
                reviewed_lock=reviewed_lock,
                baseline_lock=baseline_lock,
                candidate_lock=candidate_lock,
                drivers_root=drivers,
                output=receipt,
            )
        return baseline, candidate, drivers, receipt, values

    def test_prepares_external_drivers_without_baseline_overlay(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-preparation-") as directory:
            root = Path(directory)
            baseline, candidate, drivers, receipt, pins = self._prepare(root)
            self.assertEqual(_git(baseline, "status", "--porcelain"), "")
            self.assertEqual(_git(candidate, "status", "--porcelain"), "")
            self.assertEqual(sorted(path.name for path in (drivers / "baseline" / "examples").iterdir()), [
                "ole_owner_oracle.rs",
                "ole_package_clone_probe.rs",
                "ole_w5_probe.rs",
            ])
            with mock.patch.multiple(preparation, **pins):
                preparation.validate_preparation_receipt(receipt)

    def test_rejects_candidate_probe_or_source_mutation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-preparation-") as directory:
            root = Path(directory)
            _, candidate, _, receipt, pins = self._prepare(root)
            (candidate / preparation.PROBE_PATH).write_text("mutated\n", encoding="utf-8")
            with mock.patch.multiple(preparation, **pins):
                with self.assertRaisesRegex(preparation.PreparationError, "source identity changed"):
                    preparation.validate_preparation_receipt(receipt)

    def test_rejects_driver_input_mutation_and_extra_file(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-preparation-") as directory:
            root = Path(directory)
            _, _, drivers, receipt, pins = self._prepare(root)
            driver_probe = drivers / "baseline" / "examples" / preparation.PROBE_PATH.name
            driver_probe.write_text(PROBE_MARKERS + "tampered\n", encoding="utf-8")
            with mock.patch.multiple(preparation, **pins):
                with self.assertRaisesRegex(preparation.PreparationError, "driver file identity changed"):
                    preparation.validate_preparation_receipt(receipt)

            driver_probe.write_text(
                PROBE_MARKERS
                + "OleFile::open(Cursor::new(package.into_inner()))\n",
                encoding="utf-8",
            )
            (drivers / "candidate" / "unexpected.txt").write_text("no\n", encoding="utf-8")
            with mock.patch.multiple(preparation, **pins):
                with self.assertRaisesRegex(preparation.PreparationError, "unexpected or missing"):
                    preparation.validate_preparation_receipt(receipt)

    def test_rejects_nonempty_baseline_overlay_policy(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-preparation-") as directory:
            root = Path(directory)
            baseline, _, _, receipt, pins = self._prepare(root)
            document = __import__("json").loads(receipt.read_text(encoding="utf-8"))
            document["baseline_overlay"]["allowed_paths"] = ["Cargo.lock"]
            receipt.write_text(__import__("json").dumps(document), encoding="utf-8")
            with mock.patch.multiple(preparation, **pins):
                with self.assertRaisesRegex(preparation.PreparationError, "production overlay"):
                    preparation.validate_preparation_receipt(receipt)
            (baseline / "unapproved.txt").write_text("changed\n", encoding="utf-8")

    def test_lock_patch_rejects_registry_identity_or_unreviewed_edge(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-preparation-") as directory:
            root = Path(directory)
            self._prepare(root)
            reviewed = root / "reviewed.Cargo.lock"
            candidate = root / "candidate.Cargo.lock"
            mutated_registry = root / "candidate-registry.Cargo.lock"
            mutated_registry.write_text(
                candidate.read_text(encoding="utf-8").replace("b" * 64, "c" * 64),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(preparation.PreparationError, "registry package identity"):
                preparation._lock_patch(reviewed, mutated_registry, "candidate")

            mutated_edge = root / "candidate-edge.Cargo.lock"
            mutated_edge.write_text(
                candidate.read_text(encoding="utf-8").replace(' "litchi-core",', ' "unreviewed-edge",'),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(preparation.PreparationError, "dependency patch differs"):
                preparation._lock_patch(reviewed, mutated_edge, "candidate")


if __name__ == "__main__":
    unittest.main()
