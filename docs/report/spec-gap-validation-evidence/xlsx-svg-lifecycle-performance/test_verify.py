"""Adversarial receipt checks; these are not performance measurements."""

import json
import hashlib
import os
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import verify


class ReceiptChecks(unittest.TestCase):
    def setUp(self):
        self.directory = TemporaryDirectory(prefix="litchi-profile-verifier-")
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)

    def process(self, name="capture-p1.json"):
        path = self.root / name
        path.with_suffix(".stderr.log").write_bytes(b"")
        path.with_suffix(".time.txt").write_text("\tExit status: 0\n")
        return path

    def test_process_requires_empty_stderr_and_successful_exit(self):
        path = self.process()
        verify.verify_process_output(path)
        for stderr, timing in [
            (b"debug output\n", "Exit status: 0\n"),
            (b"", "Exit status: 1\n"),
            (b"", "Command terminated by signal 9\n"),
            (b"", "Exit status: 0\nExit status: 0\n"),
        ]:
            with self.subTest(stderr=stderr, timing=timing):
                path.with_suffix(".stderr.log").write_bytes(stderr)
                path.with_suffix(".time.txt").write_text(timing)
                with self.assertRaises(AssertionError):
                    verify.verify_process_output(path)

    def test_missing_stderr_is_not_evidence_of_silence(self):
        path = self.process()
        path.with_suffix(".stderr.log").unlink()
        with self.assertRaisesRegex(AssertionError, "stderr receipt missing"):
            verify.verify_process_output(path)

    def test_binary_identity_must_be_well_formed_and_stable(self):
        receipt = "a" * 64 + "  /removed/disposable/binary\n"
        (self.root / "binary.sha256").write_text(receipt)
        (self.root / "binary-after.sha256").write_text(receipt)
        self.assertEqual(verify.verify_binary_receipts(self.root), "a" * 64)
        (self.root / "binary-after.sha256").write_text(receipt.replace("aaa", "bbb", 1))
        with self.assertRaisesRegex(AssertionError, "changed during measurement"):
            verify.verify_binary_receipts(self.root)
        for name in ("binary.sha256", "binary-after.sha256"):
            (self.root / name).write_text("invalid  /binary\n")
        with self.assertRaisesRegex(AssertionError, "malformed executable SHA"):
            verify.verify_binary_receipts(self.root)

    def test_three_arbitrary_filenames_do_not_prove_process_identities(self):
        for number in (1, 2, 4):
            (self.root / f"capture-p{number}.json").write_text("{}")
        with patch.object(verify, "LANES", ("capture",)):
            with self.assertRaisesRegex(AssertionError, "process identities"):
                verify.verify_lanes(self.root)

    def test_input_size_must_match_even_when_digests_match(self):
        sample = {
            "expected_success": True, "actual_success": True,
            "semantic_ok": True, "output_exact": True,
            "direct_allocated_bytes": 0, "realloc_new_bytes": 0,
            "realloc_old_bytes": 0, "deallocated_bytes": 0,
            "live_before": 0, "live_after": 0, "requested_alloc_bytes": 0,
            "alloc_balance_ok": True, "alloc_invalid": False,
            "alloc_failed": 0, "error": None,
        }
        receipt = {
            "schema": "xlsx-svg-lifecycle-profile-v1", "lane": "capture",
            "warmup": 2, "sample_count": 20, "samples": [sample] * 20,
            "expected_success": True, "input_hash_fnv1a64": 1,
            "input_sha256": "a" * 64, "input_bytes": 100,
        }
        for number in range(1, 4):
            self.process(f"capture-p{number}.json").write_text(json.dumps(receipt))
        with patch.object(verify, "LANES", ("capture",)):
            verify.verify_lanes(self.root)
            receipt["input_bytes"] = 101
            (self.root / "capture-p3.json").write_text(json.dumps(receipt))
            with self.assertRaisesRegex(AssertionError, "fixture size changed"):
                verify.verify_lanes(self.root)

    def test_native_receipt_identity_must_match_documented_producer_fixture(self):
        native = self.root / "native.xlsx"
        native.write_bytes(b"bounded synthetic verifier input")
        digest = hashlib.sha256(native.read_bytes()).hexdigest()
        corpus = {"source_fixture": {"path": native.name, "sha256": digest}}
        receipt = {"input_sha256": digest, "input_bytes": native.stat().st_size}
        for number in range(1, 4):
            (self.root / f"capture_native_fixture-p{number}.json").write_text(json.dumps(receipt))
        verify.verify_native_identity(self.root, corpus, self.root)
        receipt["input_sha256"] = "b" * 64
        (self.root / "capture_native_fixture-p2.json").write_text(json.dumps(receipt))
        with self.assertRaisesRegex(AssertionError, "different fixture"):
            verify.verify_native_identity(self.root, corpus, self.root)

    @unittest.skipUnless(Path("/usr/bin/time").is_file(), "runner requires GNU time")
    def test_runner_cleanup_preserves_exploratory_and_unrelated_evidence(self):
        here = self.root / "docs/report/spec-gap-validation-evidence/profile"
        results = here / "results"
        results.mkdir(parents=True)
        script = here / "run_profile.sh"
        script.write_bytes(Path(__file__).with_name("run_profile.sh").read_bytes())
        keep = ["exploratory-review.json", "unrelated.json", "prior-review.log"]
        remove = ["capture_native_fixture-p1.json", "capture_native_fixture-p1.stderr.log"]
        for name in keep + remove:
            (results / name).write_bytes(b"retained evidence")
        bin_dir = self.root / "bin"
        bin_dir.mkdir()
        cargo = bin_dir / "cargo"
        cargo.write_text("#!/bin/sh\nexit 81\n")
        cargo.chmod(0o755)
        target = self.root / "disposable-target"
        env = os.environ.copy()
        for key in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP", "ALLOW_EXISTING_TARGET"):
            env.pop(key, None)
        env.update({
            "PATH": str(bin_dir) + os.pathsep + env.get("PATH", ""),
            "PROFILE_FROZEN": "1", "XLSX_SVG_PROFILE_API_WIRED": "1",
            "PROCESSES": "3", "WARMUP": "2", "SAMPLES": "20",
            "CARGO_TARGET_DIR": str(target),
        })
        run = subprocess.run(["bash", str(script)], env=env, capture_output=True, text=True)
        self.assertEqual(run.returncode, 81, run.stderr)
        for name in keep:
            self.assertEqual((results / name).read_bytes(), b"retained evidence")
        for name in remove:
            self.assertFalse((results / name).exists())
        self.assertFalse(target.exists(), "owned target must be cleaned on build failure")


if __name__ == "__main__":
    unittest.main()
