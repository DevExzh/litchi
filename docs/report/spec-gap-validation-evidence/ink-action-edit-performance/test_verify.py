"""Adversarial verifier and cleanup checks for the InkAction profile."""

from __future__ import annotations

import os
import hashlib
import subprocess
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

import verify


def valid_sample() -> dict[str, object]:
    return {
        "elapsed_ns": 1,
        "requested_alloc_bytes": 0,
        "direct_allocated_bytes": 0,
        "realloc_old_bytes": 0,
        "realloc_new_bytes": 0,
        "deallocated_bytes": 0,
        "alloc_calls": 0,
        "realloc_calls": 0,
        "dealloc_calls": 0,
        "live_before": 0,
        "live_after": 0,
        "peak_live_delta": 0,
        "alloc_balance_ok": True,
        "alloc_invalid": False,
        "alloc_failed": 0,
        "expected_success": True,
        "actual_success": True,
        "semantic_ok": True,
        "source_exact": True,
        "source_shared": True,
        "inverse_ok": True,
        "opaque_preserved": True,
        "output_exact": True,
        "rejection_ok": None,
        "rejection_resource": None,
        "rejection_limit": None,
    }


class ReceiptChecks(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = TemporaryDirectory(prefix="litchi-ink-profile-verifier-")
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)

    def test_rejects_negative_allocator_counter(self) -> None:
        sample = valid_sample()
        sample["requested_alloc_bytes"] = -1
        with self.assertRaisesRegex(AssertionError, "negative or malformed"):
            verify.verify_sample(sample, True, "no_op_small_8", 10, self.root / "sample.json")

    def test_peak_cannot_be_less_than_retained_live_growth(self) -> None:
        sample = valid_sample()
        sample.update(requested_alloc_bytes=10, direct_allocated_bytes=10, live_after=10)
        with self.assertRaises(AssertionError):
            verify.verify_sample(sample, True, "no_op_small_8", 10, self.root / "sample.json")

    def test_rss_requires_one_positive_value(self) -> None:
        path = self.root / "capture-p1.json"
        path.with_suffix(".stderr.log").write_bytes(b"")
        timing = path.with_suffix(".time.txt")
        for values in ("0", "1\nMaximum resident set size (kbytes): 2"):
            with self.subTest(values=values):
                timing.write_text(f"Maximum resident set size (kbytes): {values}\nExit status: 0\n")
                with self.assertRaises(AssertionError):
                    verify.verify_process_output(path)
                    verify.rss(timing)

    def test_rejects_non_boolean_applicable_field(self) -> None:
        sample = valid_sample()
        sample["semantic_ok"] = 1
        with self.assertRaisesRegex(AssertionError, "not boolean or null"):
            verify.verify_sample(sample, True, "no_op_small_8", 10, self.root / "sample.json")

    def test_process_requires_empty_stderr_and_exact_success(self) -> None:
        path = self.root / "capture-p1.json"
        stderr = path.with_suffix(".stderr.log")
        timing = path.with_suffix(".time.txt")
        stderr.write_bytes(b"")
        timing.write_text("Maximum resident set size (kbytes): 1\nExit status: 0\n")
        verify.verify_process_output(path)
        for stderr_bytes, timing_text in (
            (b"diagnostic\n", "Maximum resident set size (kbytes): 1\nExit status: 0\n"),
            (b"", "Maximum resident set size (kbytes): 1\nExit status: 1\n"),
            (b"", "Maximum resident set size (kbytes): 1\nCommand terminated by signal 9\n"),
        ):
            stderr.write_bytes(stderr_bytes)
            timing.write_text(timing_text)
            with self.assertRaises(AssertionError):
                verify.verify_process_output(path)

    def test_process_identity_receipts_require_p1_p2_p3(self) -> None:
        for number in (1, 2, 4):
            (self.root / f"capture-p{number}.json").write_text("{}")
        with patch.object(verify, "LANES", ("capture",)):
            with self.assertRaisesRegex(AssertionError, "process identities"):
                verify.verify_lanes(self.root)

    def test_binary_identity_must_be_well_formed_and_stable(self) -> None:
        missing_binary = self.root / "missing" / "ink-action-edit-profile"
        receipt = "a" * 64 + f"  {missing_binary}\n"
        (self.root / "binary.sha256").write_text(receipt)
        (self.root / "binary-after.sha256").write_text(receipt)
        (self.root / "build-provenance.txt").write_text(
            f"binary={missing_binary}\n"
            + receipt
        )
        self.assertEqual(verify.verify_binary_receipts(self.root), "a" * 64)
        (self.root / "binary-after.sha256").write_text(receipt.replace("aaa", "bbb", 1))
        with self.assertRaisesRegex(AssertionError, "changed during measurement"):
            verify.verify_binary_receipts(self.root)

    def test_binary_identity_rechecks_live_executable_when_retained(self) -> None:
        binary = self.root / "ink-action-edit-profile"
        binary.write_bytes(b"profile binary")
        digest = hashlib.sha256(binary.read_bytes()).hexdigest()
        receipt = f"{digest}  {binary}\n"
        (self.root / "binary.sha256").write_text(receipt)
        (self.root / "binary-after.sha256").write_text(receipt)
        (self.root / "build-provenance.txt").write_text(f"binary={binary}\n{receipt}")
        verify.verify_binary_receipts(self.root)
        binary.write_bytes(b"tampered profile binary")
        with self.assertRaisesRegex(AssertionError, "live profile executable hash"):
            verify.verify_binary_receipts(self.root)

    @unittest.skipUnless(Path("/usr/bin/time").is_file(), "runner requires GNU time")
    def test_runner_cleanup_preserves_unrelated_evidence(self) -> None:
        results = self.root / "results"
        results.mkdir()
        keep = ("exploratory-review.json", "unrelated.json", "prior-review.log")
        for name in keep:
            (results / name).write_bytes(b"retain")
        for name in (
            "draft_small_8-p1.json",
            "draft_small_8-p1.time.txt",
            "draft_small_8-p1.stderr.log",
        ):
            (results / name).write_bytes(b"remove")

        bin_dir = self.root / "bin"
        bin_dir.mkdir()
        cargo = bin_dir / "cargo"
        cargo.write_text("#!/bin/sh\nexit 81\n")
        cargo.chmod(0o755)
        git = bin_dir / "git"
        git.write_text(
            "#!/bin/sh\n"
            "if [ \"$1\" = -C ]; then shift 2; fi\n"
            "case \"$1\" in\n"
            "  rev-parse) echo 079cbbcbfc00c8d2412a38586a9688ae7eb0009e ;;\n"
            "  merge-base) exit 0 ;;\n"
            "  status) exit 0 ;;\n"
            "  *) exit 81 ;;\n"
            "esac\n"
        )
        git.chmod(0o755)
        target = self.root / "owned-target"
        env = os.environ.copy()
        for key in list(env):
            if key.startswith("CARGO_PROFILE_RELEASE_"):
                env.pop(key)
        for key in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP", "RUSTDOCFLAGS", "ALLOW_EXISTING_TARGET"):
            env.pop(key, None)
        env.update(
            {
                "PATH": str(bin_dir) + os.pathsep + env.get("PATH", ""),
                "PROFILE_FROZEN": "1",
                "PROFILE_RESULTS_DIR": str(results),
                "PROFILE_REPORT_OUTPUT": str(self.root / "report.md"),
                "CARGO_TARGET_DIR": str(target),
                "PROCESSES": "3",
                "WARMUP": "2",
                "SAMPLES": "20",
            }
        )
        script = Path(__file__).with_name("run_profile.sh")
        run = subprocess.run(["bash", str(script)], env=env, capture_output=True, text=True)
        self.assertEqual(run.returncode, 81, run.stderr)
        for name in keep:
            self.assertEqual((results / name).read_bytes(), b"retain")
        for name in (
            "draft_small_8-p1.json",
            "draft_small_8-p1.time.txt",
            "draft_small_8-p1.stderr.log",
        ):
            self.assertFalse((results / name).exists())
        self.assertFalse(target.exists(), "owned target must be cleaned on build failure")


if __name__ == "__main__":
    unittest.main()
