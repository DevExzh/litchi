#!/usr/bin/env python3
"""Fail-closed tests for the paired ODF shared-triad orchestration."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import tempfile
import unittest

from tools.native_qualification import odf_shared_triad_paired_runner as paired
from tools.native_qualification import odf_shared_triad_runner as shared
from tools.native_qualification import prepare_odf_shared_triad_inputs as preparation
from tools.test_prepare_odf_shared_triad_inputs import repository


ROOT = Path(__file__).resolve().parents[1]
MANIFEST_PATH = ROOT / "tools/native_qualification/manifests/odf-shared-triad.json"


class OdfSharedTriadPairedRunnerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest, _, _ = paired._load_json(MANIFEST_PATH, "manifest")

    def test_exact_paired_contract(self) -> None:
        paired.validate_manifest(self.manifest)
        for field, value, message in (
            ("abba_order", ["candidate", "baseline", "candidate", "baseline"], "ABBA order"),
            ("warmup_iterations_per_leg", 2, "warmups"),
            ("samples_per_leg", 29, "samples"),
            ("probe_iterations", 127, "probe iterations"),
        ):
            broken = copy.deepcopy(self.manifest)
            broken["paired"][field] = value
            with self.assertRaisesRegex(paired.RunnerError, message):
                paired.validate_manifest(broken)

    def test_allocator_and_overlay_mutations_fail_closed(self) -> None:
        broken = copy.deepcopy(self.manifest)
        broken["paired"]["allocator"]["elapsed_claimable"] = True
        with self.assertRaisesRegex(paired.RunnerError, "non-claimable"):
            paired.validate_manifest(broken)

        broken = copy.deepcopy(self.manifest)
        broken["paired"]["baseline_overlay"]["allowed_paths"] = ["Cargo.lock"]
        with self.assertRaisesRegex(paired.RunnerError, "overlay paths"):
            paired.validate_manifest(broken)

        broken = copy.deepcopy(self.manifest)
        broken["candidate_commit"] = shared.HISTORICAL_CANDIDATE
        with self.assertRaisesRegex(paired.RunnerError, "historical"):
            paired.validate_manifest(broken)

    def test_preparation_rejects_unreviewed_overlay(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-paired-prep-") as directory:
            root = Path(directory)
            baseline_root = root / "baseline"
            candidate_root = root / "candidate"
            candidate_probe = b"reviewed adapter\n"
            baseline_ref = repository(baseline_root, probe=None, lock=None)
            candidate_ref = repository(candidate_root, probe=candidate_probe, lock=b"candidate lock\n")
            baseline_lock = root / "baseline.lock"
            baseline_lock.write_bytes(b"baseline lock\n")
            old_commit = preparation.BASELINE_COMMIT
            preparation.BASELINE_COMMIT = baseline_ref
            try:
                manifest = copy.deepcopy(self.manifest)
                manifest["baseline_commit"] = baseline_ref
                manifest["probe"]["sha256"] = hashlib.sha256(candidate_probe).hexdigest()
                document = preparation.prepare(
                    baseline_root=baseline_root,
                    candidate_root=candidate_root,
                    candidate_ref=candidate_ref,
                    candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                    baseline_lock=baseline_lock,
                    output=root / "preparation.json",
                )
                valid = paired._validate_preparation(
                    document,
                    preparation_hash=hashlib.sha256(
                        (root / "preparation.json").read_bytes()
                    ).hexdigest(),
                    manifest=manifest,
                    baseline_root=baseline_root,
                    candidate_root=candidate_root,
                    candidate_commit=candidate_ref,
                    candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                )
                self.assertEqual(valid["baseline_probe_sha256"], manifest["probe"]["sha256"])
                broken = copy.deepcopy(document)
                broken["baseline"]["overlay_paths"] = ["Cargo.lock", "unexpected.txt"]
                with self.assertRaisesRegex(paired.RunnerError, "overlay paths"):
                    paired._validate_preparation(
                        broken,
                        preparation_hash="0" * 64,
                        manifest=manifest,
                        baseline_root=baseline_root,
                        candidate_root=candidate_root,
                        candidate_commit=candidate_ref,
                        candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                    )
            finally:
                preparation.BASELINE_COMMIT = old_commit


if __name__ == "__main__":
    unittest.main()
