#!/usr/bin/env python3
"""Fail-closed tests for the paired ODF shared-triad orchestration."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import subprocess
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

    def test_source_closure_rejects_sparse_and_skip_worktree_endpoints(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-paired-closure-") as directory:
            root = Path(directory) / "source"
            repository(root, probe=None, lock=None)
            self.assertGreater(paired._tracked_closure(root)["file_count"], 0)

            subprocess.run(
                ["git", "-C", str(root), "config", "core.sparseCheckout", "true"],
                check=True,
            )
            with self.assertRaisesRegex(paired.RunnerError, "sparse-checkout"):
                paired._tracked_closure(root)

            subprocess.run(
                ["git", "-C", str(root), "config", "core.sparseCheckout", "false"],
                check=True,
            )
            subprocess.run(
                ["git", "-C", str(root), "update-index", "--skip-worktree", "Cargo.toml"],
                check=True,
            )
            with self.assertRaisesRegex(paired.RunnerError, "sparse/skip-worktree"):
                paired._tracked_closure(root)

    def test_fresh_file_rejects_all_symlink_ancestors_and_existing_receipts(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-paired-fresh-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            baseline.mkdir()
            candidate.mkdir()
            runner = object.__new__(paired.PairedRunner)
            runner.baseline_root = baseline
            runner.candidate_root = candidate

            real = root / "real"
            real.mkdir()
            link = root / "link"
            link.symlink_to(real, target_is_directory=True)
            with self.assertRaisesRegex(paired.RunnerError, "ancestor"):
                runner._fresh_file(link / "nested" / "receipt.json", "receipt")

            existing = root / "existing.json"
            existing.write_text("old\n", encoding="utf-8")
            with self.assertRaisesRegex(paired.RunnerError, "must not overwrite"):
                runner._fresh_file(existing, "receipt")

    def test_build_receipt_rejects_reordered_event_journal(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-paired-journal-") as directory:
            root = Path(directory)
            output = root / "output"
            (output / "build").mkdir(parents=True)
            runner = object.__new__(paired.PairedRunner)
            runner.output = output
            runner.manifest_sha256 = "manifest"
            runner.preparation = {"sha256": "preparation"}
            runner.runner_hash = "runner"
            runner.cpu = 0
            runner.manifest = self.manifest
            runner._toolchain = lambda: {}
            events = [
                {"event_id": f"build-{role}-{variant}"}
                for variant in ("normal", "allocator")
                for role in paired.ROLES
            ]
            receipt = {
                "schema": paired.SCHEMA,
                "phase": "build",
                "status": "passed",
                "manifest": {"sha256": "manifest"},
                "preparation": {"sha256": "preparation"},
                "runner": {"sha256": "runner"},
                "claim_scope": "build identity only; no native acceptance claim",
                "toolchain": {},
                "cpu_affinity": 0,
                "features": list(paired.FEATURES),
                "routes": list(self.manifest["routes"]),
                "phases": dict(self.manifest["paired"]),
                "endpoints": {role: {} for role in paired.ROLES},
                "builds": {
                    f"{role}/{variant}": {}
                    for role in paired.ROLES
                    for variant in ("normal", "allocator")
                },
                "event_journal": {
                    "count": 4,
                    "events": list(reversed(events)),
                    "sha256": paired._sha256_bytes(
                        json.dumps(list(reversed(events)), sort_keys=True, separators=(",", ":")).encode()
                    ),
                },
            }
            (output / "build-receipt.json").write_text(
                json.dumps(receipt), encoding="utf-8"
            )
            with self.assertRaisesRegex(paired.RunnerError, "journal order"):
                runner._load_build_receipt()

    def test_probe_handoff_rejects_substituted_external_binary(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-paired-handoff-") as directory:
            root = Path(directory)
            runner = object.__new__(paired.PairedRunner)
            runner.baseline_root = root / "baseline"
            runner.candidate_root = root / "candidate"
            runner.target_root = root / "targets"
            runner.tmp_root = root / "tmp"
            fixed = runner._binary_path("baseline", "normal")
            fixed.parent.mkdir(parents=True)
            fixed.write_bytes(b"reviewed binary")
            fixed.chmod(0o755)
            external = root / "substituted"
            external.write_bytes(b"substituted binary")
            with self.assertRaisesRegex(paired.RunnerError, "fixed endpoint binary"):
                runner._run_probe(
                    build={
                        "role": "baseline",
                        "variant": "normal",
                        "binary": {
                            "path": str(external),
                            "sha256": hashlib.sha256(fixed.read_bytes()).hexdigest(),
                            "bytes": fixed.stat().st_size,
                        },
                    },
                    phase="normal",
                    role="baseline",
                    kind="sample",
                    index=0,
                    iterations=paired.PROBE_ITERATIONS,
                )


if __name__ == "__main__":
    unittest.main()
