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
            manifest["source"]["candidate_lock_bytes"] = value["candidate"]["lock"]["bytes"]
            manifest["probe"]["sha256"] = value["candidate"]["probe"]["sha256"]
            old_runner_candidate = runner.CANDIDATE_COMMIT
            old_runner_lock_sha = runner.CANDIDATE_LOCK_SHA256
            old_runner_lock_bytes = runner.CANDIDATE_LOCK_BYTES
            runner.CANDIDATE_COMMIT = candidate_ref
            runner.CANDIDATE_LOCK_SHA256 = value["candidate"]["lock"]["sha256"]
            runner.CANDIDATE_LOCK_BYTES = value["candidate"]["lock"]["bytes"]
            try:
                manifest = runner.validate_manifest(manifest)
            finally:
                runner.CANDIDATE_COMMIT = old_runner_candidate
                runner.CANDIDATE_LOCK_SHA256 = old_runner_lock_sha
                runner.CANDIDATE_LOCK_BYTES = old_runner_lock_bytes
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

    def test_public_prepare_main_to_real_runner_admission(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-cli-composition-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            adapter = root / "adapter"
            baseline_ref = repository(baseline, probe=None, lock=None, consumer=False)
            candidate_ref = repository(candidate, probe=None, lock=b"candidate\n", consumer=True)
            repository(adapter, probe=b"adapter\n", lock=None, consumer=True)
            external_lock = root / "baseline.lock"
            external_lock.write_bytes(b"baseline\n")
            manifest = json.loads(
                (Path(__file__).resolve().parents[1] / "tools/native_qualification/manifests/odf-raw-codec.json").read_text(encoding="utf-8")
            )
            manifest["baseline_commit"] = baseline_ref
            manifest["candidate_commit"] = candidate_ref
            manifest["source"]["candidate_consumer_manifest_sha256"] = hashlib.sha256(
                (adapter / prep.CONSUMER_MANIFEST).read_bytes()
            ).hexdigest()
            manifest["source"]["candidate_lock_sha256"] = hashlib.sha256(b"candidate\n").hexdigest()
            manifest["source"]["candidate_lock_bytes"] = len(b"candidate\n")
            manifest["probe"]["sha256"] = hashlib.sha256(
                (adapter / prep.PROBE_PATH).read_bytes()
            ).hexdigest()
            manifest_path = root / "manifest.json"
            manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
            preparation_path = root / "preparation.json"
            old_baseline = prep.BASELINE_COMMIT
            old_candidate = prep.CANDIDATE_COMMIT
            old_runner_candidate = runner.CANDIDATE_COMMIT
            old_runner_lock_sha = runner.CANDIDATE_LOCK_SHA256
            old_runner_lock_bytes = runner.CANDIDATE_LOCK_BYTES
            old_build = runner.PairedRunner.build
            prep.BASELINE_COMMIT = baseline_ref
            prep.CANDIDATE_COMMIT = candidate_ref
            runner.CANDIDATE_COMMIT = candidate_ref
            runner.CANDIDATE_LOCK_SHA256 = manifest["source"]["candidate_lock_sha256"]
            runner.CANDIDATE_LOCK_BYTES = manifest["source"]["candidate_lock_bytes"]

            def admit_endpoints(instance: runner.PairedRunner) -> None:
                instance._endpoint("baseline", full=True)
                instance._endpoint("candidate", full=True)

            runner.PairedRunner.build = admit_endpoints  # type: ignore[method-assign]
            try:
                self.assertEqual(
                    prep.main(
                        [
                            "--baseline-root", str(baseline),
                            "--candidate-root", str(candidate),
                            "--adapter-root", str(adapter),
                            "--candidate-ref", candidate_ref,
                            "--candidate-lock-sha256", manifest["source"]["candidate_lock_sha256"],
                            "--baseline-lock", str(external_lock),
                            "--output", str(preparation_path),
                        ]
                    ),
                    0,
                )
                candidate_lock = manifest["source"]["candidate_lock_sha256"]
                self.assertEqual(
                    runner.main(
                        [
                            "--manifest", str(manifest_path),
                            "--preparation", str(preparation_path),
                            "--baseline-root", str(baseline),
                            "--candidate-root", str(candidate),
                            "--adapter-root", str(adapter),
                            "--candidate-commit", candidate_ref,
                            "--candidate-lock-sha256", candidate_lock,
                            "--output", str(root / "runner-output"),
                            "--target-root", str(root / "target"),
                            "--tmp-root", str(root / "tmp"),
                            "--cpu", "0",
                            "--phase", "build",
                        ]
                    ),
                    0,
                )
                bad_manifest = dict(manifest)
                bad_manifest["source"] = dict(manifest["source"])
                bad_manifest["source"]["candidate_lock_sha256"] = "0" * 64
                bad_manifest_path = root / "bad-manifest.json"
                bad_manifest_path.write_text(json.dumps(bad_manifest, indent=2) + "\n", encoding="utf-8")
                self.assertEqual(
                    runner.main(
                        [
                            "--manifest", str(bad_manifest_path),
                            "--preparation", str(preparation_path),
                            "--baseline-root", str(baseline),
                            "--candidate-root", str(candidate),
                            "--adapter-root", str(adapter),
                            "--candidate-commit", candidate_ref,
                            "--candidate-lock-sha256", candidate_lock,
                            "--output", str(root / "bad-output"),
                            "--target-root", str(root / "bad-target"),
                            "--tmp-root", str(root / "bad-tmp"),
                            "--cpu", "0",
                            "--phase", "build",
                        ]
                    ),
                    2,
                )
            finally:
                runner.PairedRunner.build = old_build
                runner.CANDIDATE_COMMIT = old_runner_candidate
                runner.CANDIDATE_LOCK_SHA256 = old_runner_lock_sha
                runner.CANDIDATE_LOCK_BYTES = old_runner_lock_bytes
                prep.BASELINE_COMMIT = old_baseline
                prep.CANDIDATE_COMMIT = old_candidate


if __name__ == "__main__":
    unittest.main()
