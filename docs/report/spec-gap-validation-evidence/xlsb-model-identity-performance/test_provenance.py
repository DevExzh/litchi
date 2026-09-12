"""Git-blob and provenance regressions for the XLSB smoke scaffold."""

from __future__ import annotations

import subprocess
import tempfile
import unittest
import hashlib
from pathlib import Path

import source_manifest
import verify


class GitBlobSnapshotTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory(prefix="litchi-xlsb-profile-git-")
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        subprocess.run(["git", "-C", str(self.root), "init", "-q"], check=True)
        subprocess.run(
            ["git", "-C", str(self.root), "config", "user.name", "Profile Test"],
            check=True,
        )
        subprocess.run(
            [
                "git",
                "-C",
                str(self.root),
                "config",
                "user.email",
                "profile@example.invalid",
            ],
            check=True,
        )
        self.source = self.root / "input.rs"
        self.source.write_text("fn source() {}\n")
        subprocess.run(["git", "-C", str(self.root), "add", "input.rs"], check=True)
        subprocess.run(
            ["git", "-C", str(self.root), "commit", "-qm", "source"],
            check=True,
        )
        self.commit = subprocess.check_output(
            ["git", "-C", str(self.root), "rev-parse", "HEAD"],
            text=True,
        ).strip()

    def assert_rejected(self) -> None:
        with self.assertRaisesRegex(SystemExit, "committed Git blobs"):
            source_manifest.verify_git_inputs([self.source], self.root, self.commit)
        with self.assertRaisesRegex(AssertionError, "committed snapshot"):
            verify.verify_git_snapshot([self.source], self.root, self.commit)

    def test_clean_input_is_accepted(self) -> None:
        source_manifest.verify_git_inputs([self.source], self.root, self.commit)
        verify.verify_git_snapshot([self.source], self.root, self.commit)

    def test_unstaged_dirty_input_is_rejected(self) -> None:
        self.source.write_text("fn changed() {}\n")
        self.assert_rejected()

    def test_staged_dirty_input_is_rejected(self) -> None:
        self.source.write_text("fn staged() {}\n")
        subprocess.run(["git", "-C", str(self.root), "add", "input.rs"], check=True)
        self.assert_rejected()

    def test_assume_unchanged_does_not_hide_dirty_input(self) -> None:
        subprocess.run(
            ["git", "-C", str(self.root), "update-index", "--assume-unchanged", "input.rs"],
            check=True,
        )
        self.source.write_text("fn hidden_change() {}\n")
        self.assert_rejected()

    def test_untracked_input_is_rejected(self) -> None:
        untracked = self.root / "untracked.rs"
        untracked.write_text("fn untracked() {}\n")
        with self.assertRaisesRegex(SystemExit, "not tracked"):
            source_manifest.verify_git_inputs([untracked], self.root, self.commit)
        with self.assertRaisesRegex(AssertionError, "not tracked"):
            verify.verify_git_snapshot([untracked], self.root, self.commit)

    def test_manifest_commit_mismatch_is_rejected(self) -> None:
        wrong_commit = "0" * 40
        with self.assertRaisesRegex(SystemExit, "Git head changed"):
            source_manifest.verify_git_inputs([self.source], self.root, wrong_commit)
        with self.assertRaisesRegex(AssertionError, "Git snapshot commit changed"):
            verify.verify_git_snapshot([self.source], self.root, wrong_commit)

    def test_metadata_provenance_mismatch_is_rejected(self) -> None:
        results = self.root / "results"
        results.mkdir()
        metadata = results / "metadata-before.json"
        metadata.write_text("{}\n")
        (results / "metadata-after.json").write_text("{}\n")
        source_manifest_hash = "a" * 64
        binary = results / "xlsb-model-identity-profile"
        binary.write_bytes(b"binary")
        binary_digest = hashlib.sha256(binary.read_bytes()).hexdigest()
        (results / "provenance.txt").write_text(
            "\n".join(
                (
                    f"git_head={self.commit}",
                    "git_baseline=host",
                    "neutral_baseline=neutral",
                    "git_status_relevant=",
                    "metadata_before_sha256=" + "b" * 64,
                    "metadata_after_sha256=" + "b" * 64,
                    f"source_manifest_before_sha256={source_manifest_hash}",
                    f"source_manifest_after_sha256={source_manifest_hash}",
                    f"binary_sha256={binary_digest}",
                    f"binary_after_sha256={binary_digest}",
                )
            )
            + "\n"
        )
        with self.assertRaisesRegex(SystemExit, "metadata hashes"):
            verify.verify_provenance(
                results,
                self.root,
                {"host_feature_commit": "host", "neutral_baseline_commit": "neutral"},
                self.commit,
                hashlib.sha256(metadata.read_bytes()).hexdigest(),
                source_manifest_hash,
                binary_digest,
            )

    def test_binary_receipt_mismatch_is_rejected(self) -> None:
        results = self.root / "binary-results"
        results.mkdir()
        binary = results / "xlsb-model-identity-profile"
        binary.write_bytes(b"binary")
        digest = hashlib.sha256(binary.read_bytes()).hexdigest()
        receipt = f"{digest}  {binary}\n"
        (results / "binary.sha256").write_text(receipt)
        (results / "binary-after.sha256").write_text("0" * 64 + f"  {binary}\n")
        (results / "build-provenance.txt").write_text(f"binary={binary}\n{receipt}")
        with self.assertRaisesRegex(SystemExit, "changed during smoke"):
            verify.verify_binary_receipts(results)


if __name__ == "__main__":
    unittest.main()
