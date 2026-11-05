#!/usr/bin/env python3
"""Source-only tests for the pristine ODF baseline preparation."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
import tempfile
import unittest

from tools.native_qualification import prepare_odf_shared_triad_inputs as prep


def git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return result.stdout.strip()


def repository(root: Path, *, probe: bytes | None, lock: bytes | None) -> str:
    root.mkdir(parents=True)
    git(root, "init", "-q", "--initial-branch=main")
    git(root, "config", "user.email", "odf-prep@example.invalid")
    git(root, "config", "user.name", "ODF preparation test")
    (root / "Cargo.toml").write_text("[workspace]\nmembers=[]\n", encoding="utf-8")
    if probe is not None:
        probe_path = root / prep.PROBE_PATH
        probe_path.parent.mkdir(parents=True, exist_ok=True)
        probe_path.write_bytes(probe)
    if lock is not None:
        (root / prep.LOCK_PATH).write_bytes(lock)
    git(root, "add", ".")
    git(root, "commit", "-qm", "fixture")
    return git(root, "rev-parse", "HEAD")


class OdfPreparationTests(unittest.TestCase):
    def test_prepares_only_probe_and_external_lock(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-prep-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            probe = b"reviewed adapter\n"
            baseline_ref = repository(baseline, probe=None, lock=None)
            candidate_ref = repository(candidate, probe=probe, lock=b"candidate lock\n")
            external_lock = root / "baseline.lock"
            external_lock.write_bytes(b"baseline lock\n")
            output = root / "preparation.json"
            old = prep.BASELINE_COMMIT
            prep.BASELINE_COMMIT = baseline_ref
            try:
                document = prep.prepare(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    candidate_ref=candidate_ref,
                    candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                    baseline_lock=external_lock,
                    output=output,
                )
            finally:
                prep.BASELINE_COMMIT = old
            self.assertEqual(document["baseline"]["commit"], baseline_ref)
            self.assertEqual(
                Path(document["baseline"]["probe"]["destination_path"]).read_bytes(),
                probe,
            )
            self.assertEqual(
                json.loads(output.read_text(encoding="utf-8"))["claim_scope"],
                "source preparation identity only; no build, timing, or acceptance claim",
            )
            self.assertEqual(
                sorted(prep._status_paths(document["baseline"]["after_overlay"]["status"])),
                ["Cargo.lock", str(prep.PROBE_PATH)],
            )

    def test_rejects_dirty_baseline_and_stale_candidate(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-prep-guard-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            baseline_ref = repository(baseline, probe=None, lock=None)
            candidate_ref = repository(candidate, probe=b"adapter\n", lock=b"candidate\n")
            external_lock = root / "baseline.lock"
            external_lock.write_bytes(b"baseline\n")
            old = prep.BASELINE_COMMIT
            prep.BASELINE_COMMIT = baseline_ref
            try:
                (baseline / "production.txt").write_text("mutated", encoding="utf-8")
                with self.assertRaisesRegex(prep.PreparationError, "clean"):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        candidate_ref=candidate_ref,
                        candidate_lock_sha256=hashlib.sha256(b"candidate\n").hexdigest(),
                        baseline_lock=external_lock,
                        output=root / "dirty.json",
                    )
                (baseline / "production.txt").unlink()
                with self.assertRaisesRegex(prep.PreparationError, "historical"):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        candidate_ref=prep.HISTORICAL_CANDIDATE,
                        candidate_lock_sha256=hashlib.sha256(b"candidate\n").hexdigest(),
                        baseline_lock=external_lock,
                        output=root / "stale.json",
                    )
            finally:
                prep.BASELINE_COMMIT = old

    def test_reuses_existing_approved_overlay_without_rewriting_it(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-prep-reuse-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            probe = b"reviewed adapter\n"
            baseline_ref = repository(baseline, probe=None, lock=None)
            candidate_ref = repository(candidate, probe=probe, lock=b"candidate lock\n")
            external_lock = root / "baseline.lock"
            external_lock.write_bytes(b"baseline lock\n")
            old = prep.BASELINE_COMMIT
            prep.BASELINE_COMMIT = baseline_ref
            try:
                approved_path = root / "approved.json"
                prep.prepare(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    candidate_ref=candidate_ref,
                    candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                    baseline_lock=external_lock,
                    output=approved_path,
                )
                before_probe = (baseline / prep.PROBE_PATH).read_bytes()
                before_lock = (baseline / prep.LOCK_PATH).read_bytes()
                before_status = git(baseline, "status", "--porcelain=v1", "--untracked-files=all")
                successor = root / "successor.json"
                document = prep.reuse_approved_overlay(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    candidate_ref=candidate_ref,
                    candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                    baseline_lock=external_lock,
                    approved_preparation=approved_path,
                    output=successor,
                )
            finally:
                prep.BASELINE_COMMIT = old
            self.assertEqual(document["candidate"]["commit"], candidate_ref)
            self.assertEqual((baseline / prep.PROBE_PATH).read_bytes(), before_probe)
            self.assertEqual((baseline / prep.LOCK_PATH).read_bytes(), before_lock)
            self.assertEqual(
                git(baseline, "status", "--porcelain=v1", "--untracked-files=all"),
                before_status,
            )
            self.assertEqual(
                document["baseline"]["lock"]["sha256"],
                hashlib.sha256(b"baseline lock\n").hexdigest(),
            )
            self.assertEqual(
                json.loads(successor.read_text(encoding="utf-8"))["claim_scope"],
                "retained approved source overlay rebound to selected candidate; no build, timing, or acceptance claim",
            )


if __name__ == "__main__":
    unittest.main()
