"""Source provenance must follow bytes, including Git index edge cases."""

import subprocess
import tempfile
import unittest
from pathlib import Path

import source_manifest
import verify


class SourceSnapshotTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="litchi-root-source-snapshot-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.git("init", "-q")
        self.source = self.root / "input.rs"
        self.source.write_text("committed source\n")
        self.git("add", "input.rs")
        self.git("-c", "user.name=Fixture", "-c", "user.email=fixture@example.invalid", "commit", "-qm", "fixture")
        self.commit = self.git("rev-parse", "HEAD").strip()

    def git(self, *args):
        return subprocess.check_output(["git", "-C", str(self.root), *args], text=True)

    def assert_rejected(self, path):
        with self.assertRaises(SystemExit):
            source_manifest.verify_git_inputs([path], self.root, self.commit)
        with self.assertRaises(AssertionError):
            verify.verify_git_snapshot([path], self.root, self.commit)

    def test_clean_input_is_accepted(self):
        source_manifest.verify_git_inputs([self.source], self.root, self.commit)
        verify.verify_git_snapshot([self.source], self.root, self.commit)

    def test_modified_input_is_rejected_before_and_after_staging(self):
        self.source.write_text("different bytes\n")
        self.assert_rejected(self.source)
        self.git("add", "input.rs")
        self.assert_rejected(self.source)

    def test_ignored_untracked_input_is_rejected(self):
        (self.root / ".gitignore").write_text("ignored.rs\n")
        ignored = self.root / "ignored.rs"
        ignored.write_text("uncommitted dependency\n")
        self.assert_rejected(ignored)

    def test_assume_unchanged_does_not_hide_modified_input_bytes(self):
        self.git("update-index", "--assume-unchanged", "input.rs")
        self.source.write_text("different bytes\n")
        self.assert_rejected(self.source)


if __name__ == "__main__":
    unittest.main()
