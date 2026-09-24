#!/usr/bin/env python3
"""Focused tests for the retained matched-profile comparison."""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path


HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import compare  # noqa: E402


class MatchedComparisonTests(unittest.TestCase):
    def test_defaults_bind_the_candidate_result_directory(self) -> None:
        self.assertEqual(compare.DEFAULT_BASELINE.name, "paired-baseline-e5c18ca")
        self.assertEqual(compare.DEFAULT_CANDIDATE.name, "paired-candidate-251d361f")

    def test_pair_identity_rejects_wrong_arm_pin(self) -> None:
        with self.assertRaises(compare.CompareError):
            compare.read_identity(
                compare.DEFAULT_CANDIDATE,
                "candidate",
                compare.BASELINE_PIN,
                compare.CANDIDATE_HEAD,
            )

    def test_fixture_mismatch_is_rejected(self) -> None:
        baseline = {
            "action_count": 128,
            "result_action_count": 128,
            "operation_count": 128,
            "input_bytes": 15380,
            "expected_success": True,
        }
        candidate = dict(baseline)
        candidate["input_bytes"] += 1
        with self.assertRaises(compare.CompareError):
            compare.validate_fixture_pair(baseline, candidate, "scalar_batch_scaled_128")

    def test_rss_rejects_multiple_exit_status_lines(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            receipt = Path(directory) / "sample.time.txt"
            receipt.write_text(
                "Maximum resident set size (kbytes): 1234\n"
                "Exit status: 0\n"
                "Exit status: 1\n"
            )
            with self.assertRaises(compare.CompareError):
                compare.read_rss(receipt)

    def test_retained_pair_has_unchanged_control_allocator_vectors(self) -> None:
        data = compare.compare_arms(compare.DEFAULT_BASELINE, compare.DEFAULT_CANDIDATE)
        self.assertTrue(data["control_allocator_and_peak_vectors_unchanged"])
        self.assertEqual(len(data["lanes"]), 34)
        self.assertEqual(data["baseline"]["arm"], "baseline")
        self.assertEqual(data["candidate"]["arm"], "candidate")

    def test_manifest_receipt_binds_actual_retained_bytes(self) -> None:
        needed = (
            "verification.json",
            "source-provenance.txt",
            "build-provenance.txt",
            "checkout-provenance.txt",
            "commands.txt",
            "binary.sha256",
            "binary-after.sha256",
            "source-manifest-before.txt",
            "source-manifest-after.txt",
        )
        with tempfile.TemporaryDirectory() as directory:
            result = Path(directory)
            for name in needed:
                shutil.copy2(compare.DEFAULT_BASELINE / name, result / name)
            manifest = result / "source-manifest-before.txt"
            manifest.write_text(manifest.read_text() + "tampered\n")
            with self.assertRaises(compare.CompareError):
                compare.read_identity(result, "baseline", compare.BASELINE_PIN, compare.BASELINE_HEAD)

    def test_manifest_commit_is_bound_even_if_provenance_hashes_are_rewritten(self) -> None:
        needed = (
            "verification.json",
            "source-provenance.txt",
            "build-provenance.txt",
            "checkout-provenance.txt",
            "commands.txt",
            "binary.sha256",
            "binary-after.sha256",
            "source-manifest-before.txt",
            "source-manifest-after.txt",
        )
        with tempfile.TemporaryDirectory() as directory:
            result = Path(directory)
            for name in needed:
                shutil.copy2(compare.DEFAULT_BASELINE / name, result / name)
            for name in ("source-manifest-before.txt", "source-manifest-after.txt"):
                manifest = result / name
                manifest.write_text(
                    manifest.read_text().replace(
                        f"git_commit={compare.BASELINE_HEAD}",
                        "git_commit=" + "0" * 40,
                    )
                )
            source_provenance = result / "source-provenance.txt"
            source_text = source_provenance.read_text()
            for name in ("source-manifest-before.txt", "source-manifest-after.txt"):
                digest = compare.sha256(result / name)
                field = "source_manifest_before_sha256" if "before" in name else "source_manifest_after_sha256"
                prefix = field + "="
                line = next(line for line in source_text.splitlines() if line.startswith(prefix))
                source_text = source_text.replace(line, prefix + digest)
            source_provenance.write_text(source_text)
            with self.assertRaises(compare.CompareError):
                compare.read_identity(result, "baseline", compare.BASELINE_PIN, compare.BASELINE_HEAD)

    def test_manifest_pinned_source_entry_is_bound_to_the_arm(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            manifest = Path(directory) / "source-manifest.txt"
            shutil.copy2(compare.DEFAULT_BASELINE / "source-manifest-before.txt", manifest)
            expected = compare.PINNED_SOURCE_ENTRIES["baseline"][
                "crates/litchi-drawingml/src/ink/actions.rs"
            ]
            manifest.write_text(manifest.read_text().replace(expected, "0" * 64, 1))
            with self.assertRaises(compare.CompareError):
                compare.validate_source_manifest(manifest, "baseline", compare.BASELINE_HEAD)

    def test_semantic_corruption_is_rejected_even_with_stale_verification(self) -> None:
        lane = "no_op_small_8"
        with tempfile.TemporaryDirectory() as directory:
            result = Path(directory)
            for process in (1, 2, 3):
                for suffix in ("json", "time.txt", "stderr.log"):
                    name = f"{lane}-p{process}.{suffix}"
                    shutil.copy2(compare.DEFAULT_BASELINE / name, result / name)
            receipt = result / f"{lane}-p1.json"
            value = json.loads(receipt.read_text())
            value["samples"][0]["semantic_ok"] = False
            receipt.write_text(json.dumps(value))
            with self.assertRaises(compare.CompareError):
                compare.load_lane(result, lane)


if __name__ == "__main__":
    unittest.main()
