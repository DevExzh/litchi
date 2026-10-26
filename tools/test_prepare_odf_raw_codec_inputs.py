#!/usr/bin/env python3
"""Source-only tests for ODF raw-codec baseline preparation."""

from __future__ import annotations

import hashlib
from pathlib import Path
import subprocess
import tempfile
import unittest

from tools.native_qualification import prepare_odf_raw_codec_inputs as prep


def git(root: Path, *args: str) -> str:
    result = subprocess.run(["git", "-C", str(root), *args], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    return result.stdout.strip()


def repository(root: Path, *, probe: bytes | None, lock: bytes | None) -> str:
    root.mkdir(parents=True)
    git(root, "init", "-q", "--initial-branch=main")
    git(root, "config", "user.email", "odf-raw-prep@example.invalid")
    git(root, "config", "user.name", "ODF raw preparation test")
    (root / "Cargo.toml").write_text("[workspace]\nmembers=[]\n", encoding="utf-8")
    consumer = root / prep.CONSUMER_MANIFEST
    consumer.parent.mkdir(parents=True, exist_ok=True)
    consumer.write_text(
        "[package]\nname = \"litchi-perf-baseline\"\n"
        + ("# candidate adapter\n" if probe is not None else "# baseline adapter\n"),
        encoding="utf-8",
    )
    if probe is not None:
        path = root / prep.PROBE_PATH
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(probe)
    if lock is not None:
        path = root / prep.LOCK_PATH
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(lock)
    git(root, "add", ".")
    git(root, "commit", "-qm", "fixture")
    return git(root, "rev-parse", "HEAD")


class OdfRawPreparationTests(unittest.TestCase):
    def test_overlay_is_limited_to_adapter_and_lock(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-prep-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            baseline_ref = repository(baseline, probe=None, lock=None)
            candidate_ref = repository(candidate, probe=b"reviewed raw adapter\n", lock=b"candidate lock\n")
            external_lock = root / "baseline.lock"
            external_lock.write_bytes(b"baseline lock\n")
            output = root / "preparation.json"
            old = prep.BASELINE_COMMIT
            prep.BASELINE_COMMIT = baseline_ref
            try:
                value = prep.prepare(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    candidate_ref=candidate_ref,
                    candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                    baseline_lock=external_lock,
                    output=output,
                )
            finally:
                prep.BASELINE_COMMIT = old
            self.assertEqual(value["lane"], prep.LANE)
            self.assertEqual(
                sorted(prep._status_paths(value["baseline"]["after_overlay"]["status"])),
                sorted([str(prep.CONSUMER_MANIFEST), str(prep.LOCK_PATH), str(prep.PROBE_PATH)]),
            )
            self.assertEqual((baseline / prep.PROBE_PATH).read_bytes(), b"reviewed raw adapter\n")
            self.assertEqual((baseline / prep.LOCK_PATH).read_bytes(), b"baseline lock\n")
            self.assertIn("no build", value["claim_scope"])

    def test_rejects_historical_candidate_and_dirty_baseline(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-prep-guard-") as directory:
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
                (baseline / "dirty.txt").write_text("dirty", encoding="utf-8")
                with self.assertRaisesRegex(prep.PreparationError, "clean"):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        candidate_ref=candidate_ref,
                        candidate_lock_sha256=hashlib.sha256(b"candidate\n").hexdigest(),
                        baseline_lock=external_lock,
                        output=root / "dirty.json",
                    )
                (baseline / "dirty.txt").unlink()
                with self.assertRaisesRegex(prep.PreparationError, "historical"):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        candidate_ref=prep.HISTORICAL_CANDIDATE,
                        candidate_lock_sha256=hashlib.sha256(b"candidate\n").hexdigest(),
                        baseline_lock=external_lock,
                        output=root / "historical.json",
                    )
            finally:
                prep.BASELINE_COMMIT = old


if __name__ == "__main__":
    unittest.main()
