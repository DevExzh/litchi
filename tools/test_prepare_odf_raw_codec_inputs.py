#!/usr/bin/env python3
"""Source-only tests for ODF raw-codec baseline preparation."""

from __future__ import annotations

import hashlib
from pathlib import Path
import subprocess
import tempfile
import unittest
import json

from tools.native_qualification import prepare_odf_raw_codec_inputs as prep
from tools.native_qualification import odf_raw_codec_paired_runner as runner


def git(root: Path, *args: str) -> str:
    result = subprocess.run(["git", "-C", str(root), *args], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    return result.stdout.strip()


def repository(root: Path, *, probe: bytes | None, lock: bytes | None, consumer: bool) -> str:
    root.mkdir(parents=True)
    git(root, "init", "-q", "--initial-branch=main")
    git(root, "config", "user.email", "odf-raw-prep@example.invalid")
    git(root, "config", "user.name", "ODF raw preparation test")
    (root / "Cargo.toml").write_text("[workspace]\nmembers=[]\n", encoding="utf-8")
    production = root / "crates/litchi-ops/Cargo.toml"
    production.parent.mkdir(parents=True, exist_ok=True)
    production.write_text(
        "[package]\nname = \"litchi-ops\"\nversion = \"0.0.0\"\n",
        encoding="utf-8",
    )
    if consumer:
        consumer_manifest = root / prep.CONSUMER_MANIFEST
        consumer_manifest.parent.mkdir(parents=True, exist_ok=True)
        consumer_manifest.write_text(
            "[package]\nname = \"litchi-perf-baseline\"\n"
            + ("# reviewed adapter\n" if probe is not None else "# candidate placeholder\n"),
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
            adapter = root / "adapter"
            baseline_ref = repository(baseline, probe=None, lock=None, consumer=False)
            candidate_ref = repository(candidate, probe=None, lock=b"candidate lock\n", consumer=True)
            repository(adapter, probe=b"reviewed raw adapter\n", lock=None, consumer=True)
            external_lock = root / "baseline.lock"
            external_lock.write_bytes(b"baseline lock\n")
            output = root / "preparation.json"
            old = prep.BASELINE_COMMIT
            old_candidate = prep.CANDIDATE_COMMIT
            prep.BASELINE_COMMIT = baseline_ref
            prep.CANDIDATE_COMMIT = candidate_ref
            try:
                value = prep.prepare(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    adapter_root=adapter,
                    candidate_ref=candidate_ref,
                    candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                    baseline_lock=external_lock,
                    output=output,
                )
            finally:
                prep.BASELINE_COMMIT = old
                prep.CANDIDATE_COMMIT = old_candidate
            self.assertEqual(value["lane"], prep.LANE)
            self.assertEqual(
                sorted(prep._status_paths(value["baseline"]["after_overlay"]["status"])),
                sorted([str(prep.CONSUMER_MANIFEST), str(prep.LOCK_PATH), str(prep.PROBE_PATH)]),
            )
            self.assertEqual((baseline / prep.PROBE_PATH).read_bytes(), b"reviewed raw adapter\n")
            self.assertEqual((baseline / prep.LOCK_PATH).read_bytes(), b"baseline lock\n")
            self.assertIn("no build", value["claim_scope"])
            manifest = json.loads(
                (Path(__file__).resolve().parents[1] / "tools/native_qualification/manifests/odf-raw-codec.json").read_text(encoding="utf-8")
            )
            manifest["baseline_commit"] = baseline_ref
            manifest["candidate_commit"] = candidate_ref
            manifest["source"]["candidate_consumer_manifest_sha256"] = value["candidate"]["consumer_manifest"]["sha256"]
            manifest["source"]["candidate_lock_sha256"] = value["candidate"]["lock"]["sha256"]
            manifest["probe"]["sha256"] = value["candidate"]["probe"]["sha256"]
            old_runner_candidate = runner.CANDIDATE_COMMIT
            runner.CANDIDATE_COMMIT = candidate_ref
            try:
                manifest = runner.validate_manifest(manifest)
            finally:
                runner.CANDIDATE_COMMIT = old_runner_candidate
            checked = runner._validate_preparation(
                value,
                hashlib.sha256(output.read_bytes()).hexdigest(),
                manifest=manifest,
                baseline_root=baseline.resolve(),
                candidate_root=candidate.resolve(),
                adapter_root=adapter.resolve(),
                candidate_commit=candidate_ref,
                candidate_lock_sha256=value["candidate"]["lock"]["sha256"],
            )
            self.assertEqual(checked["adapter"]["root"], str(adapter.resolve()))
            mutated = json.loads(json.dumps(value))
            mutated["candidate"]["probe"]["sha256"] = "0" * 64
            with self.assertRaises(runner.RunnerError):
                runner._validate_preparation(
                    mutated,
                    hashlib.sha256(json.dumps(mutated, sort_keys=True).encode()).hexdigest(),
                    manifest=manifest,
                    baseline_root=baseline.resolve(),
                    candidate_root=candidate.resolve(),
                    adapter_root=adapter.resolve(),
                    candidate_commit=candidate_ref,
                    candidate_lock_sha256=value["candidate"]["lock"]["sha256"],
                )

    def test_rejects_historical_candidate_and_dirty_baseline(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-prep-guard-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            adapter = root / "adapter"
            baseline_ref = repository(baseline, probe=None, lock=None, consumer=False)
            candidate_ref = repository(candidate, probe=None, lock=b"candidate\n", consumer=True)
            repository(adapter, probe=b"adapter\n", lock=None, consumer=True)
            external_lock = root / "baseline.lock"
            external_lock.write_bytes(b"baseline\n")
            old = prep.BASELINE_COMMIT
            old_candidate = prep.CANDIDATE_COMMIT
            prep.BASELINE_COMMIT = baseline_ref
            prep.CANDIDATE_COMMIT = candidate_ref
            try:
                (baseline / "dirty.txt").write_text("dirty", encoding="utf-8")
                with self.assertRaisesRegex(prep.PreparationError, "clean"):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        adapter_root=adapter,
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
                        adapter_root=adapter,
                        candidate_ref=prep.HISTORICAL_CANDIDATE,
                        candidate_lock_sha256=hashlib.sha256(b"candidate\n").hexdigest(),
                        baseline_lock=external_lock,
                        output=root / "historical.json",
                    )
            finally:
                prep.BASELINE_COMMIT = old
                prep.CANDIDATE_COMMIT = old_candidate

    def test_output_path_is_rejected_before_any_overlay(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-prep-output-guard-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            adapter = root / "adapter"
            baseline_ref = repository(baseline, probe=None, lock=None, consumer=False)
            candidate_ref = repository(candidate, probe=None, lock=b"candidate\n", consumer=True)
            repository(adapter, probe=b"adapter\n", lock=None, consumer=True)
            external_lock = root / "baseline.lock"
            external_lock.write_bytes(b"baseline\n")
            before_candidate = {
                path: path.read_bytes() for path in (candidate / prep.CONSUMER_MANIFEST, candidate / prep.LOCK_PATH)
            }
            old = prep.BASELINE_COMMIT
            old_candidate = prep.CANDIDATE_COMMIT
            prep.BASELINE_COMMIT = baseline_ref
            prep.CANDIDATE_COMMIT = candidate_ref
            try:
                with self.assertRaisesRegex(prep.PreparationError, "outside source roots"):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        adapter_root=adapter,
                        candidate_ref=candidate_ref,
                        candidate_lock_sha256=hashlib.sha256(b"candidate\n").hexdigest(),
                        baseline_lock=external_lock,
                        output=candidate / "preparation.json",
                    )
            finally:
                prep.BASELINE_COMMIT = old
                prep.CANDIDATE_COMMIT = old_candidate
            self.assertEqual(before_candidate, {path: path.read_bytes() for path in before_candidate})
            self.assertFalse((baseline / prep.PROBE_PATH).exists())


if __name__ == "__main__":
    unittest.main()
