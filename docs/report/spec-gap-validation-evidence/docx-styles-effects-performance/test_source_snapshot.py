"""Check production and descendant harness identity using disposable Git trees."""

import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest


class SourceSnapshotTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="litchi-docx-styles-source-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.root = self.directory / "checkout"
        self.root.mkdir()
        self.git("init", "-q")
        self.production = self.root / "crates/fixture/src/lib.rs"
        self.production.parent.mkdir(parents=True)
        self.production.write_text("pub fn value() -> u8 { 1 }\n")
        (self.production.parent.parent / "Cargo.toml").write_text(
            '[package]\nname = "fixture"\nversion = "0.1.0"\n'
        )
        self.commit("production")
        self.source_commit = self.git("rev-parse", "HEAD").stdout.strip()
        self.evidence = self.root / "docs/evidence"
        self.harness = self.evidence / "harness/main.rs"
        self.harness.parent.mkdir(parents=True)
        self.harness.write_text("fn main() {}\n")
        (self.harness.parent / "Cargo.toml").write_text(
            '[package]\nname = "harness"\nversion = "0.1.0"\n'
        )
        self.extra = self.evidence / "fixture.txt"
        self.extra.write_text("native fixture bytes\n")
        self.commit("harness")
        self.metadata = self.directory / "metadata.json"
        self.metadata.write_text(json.dumps({"packages": [
            {"name": "fixture", "version": "0.1.0", "source": None,
             "manifest_path": str(self.production.parent.parent / "Cargo.toml"),
             "targets": [{"src_path": str(self.production)}]},
            {"name": "harness", "version": "0.1.0", "source": None,
             "manifest_path": str(self.harness.parent / "Cargo.toml"),
             "targets": [{"src_path": str(self.harness)}]},
        ]}))
        self.output = self.directory / "manifest.txt"

    def git(self, *arguments):
        return subprocess.run(
            ["git", "-C", str(self.root), *arguments], check=True,
            capture_output=True, text=True,
        )

    def commit(self, message):
        self.git("add", ".")
        self.git("-c", "user.name=Guard Test", "-c", "user.email=guard@example.invalid",
                 "commit", "-qm", message)

    def capture(self):
        return subprocess.run([
            sys.executable, "-B", str(Path(__file__).with_name("source_manifest.py")),
            "--metadata", str(self.metadata), "--root", str(self.root),
            "--source-commit", self.source_commit, "--evidence", str(self.evidence),
            "--extra", str(self.extra), "--output", str(self.output),
        ], capture_output=True, text=True)

    def assert_refused(self):
        run = self.capture()
        self.assertNotEqual(run.returncode, 0, "uncommitted input was accepted")
        self.assertFalse(self.output.exists(), run.stderr)

    def test_committed_descendant_harness_and_fixture_are_admitted(self):
        run = self.capture()
        self.assertEqual(run.returncode, 0, run.stderr)
        manifest = self.output.read_text()
        self.assertIn(f"source_commit={self.source_commit}\n", manifest)
        self.assertIn(f"git_head={self.git('rev-parse', 'HEAD').stdout.strip()}\n", manifest)
        self.assertIn("docs/evidence/harness/main.rs", manifest)
        self.assertIn("docs/evidence/fixture.txt", manifest)

    def test_production_modification_is_refused_even_when_committed_in_descendant(self):
        self.production.write_text("pub fn value() -> u8 { 2 }\n")
        self.commit("changed production")
        self.assert_refused()

    def test_assume_unchanged_harness_modification_is_refused(self):
        self.git("update-index", "--assume-unchanged", "docs/evidence/harness/main.rs")
        self.harness.write_text("fn main() { panic!(); }\n")
        self.git("diff", "--quiet")
        self.assert_refused()

    def test_untracked_discovered_production_source_is_refused(self):
        (self.production.parent / "new.rs").write_text("pub const EXTRA: u8 = 1;\n")
        self.assert_refused()

    def test_new_production_source_outside_pin_is_refused_even_when_committed(self):
        (self.production.parent / "new.rs").write_text("pub const EXTRA: u8 = 1;\n")
        self.commit("new production module")
        self.assert_refused()

    def test_untracked_extra_is_refused(self):
        self.extra = self.evidence / "untracked.txt"
        self.extra.write_text("uncommitted fixture\n")
        self.assert_refused()

    def test_assume_unchanged_fixture_modification_is_refused(self):
        self.git("update-index", "--assume-unchanged", "docs/evidence/fixture.txt")
        self.extra.write_text("changed fixture\n")
        self.git("diff", "--quiet")
        self.assert_refused()


class EvidencePresentAtSourceCommitTests(unittest.TestCase):
    """Exercise descendant binding when evidence also exists at the source pin."""

    def setUp(self):
        temporary = TemporaryDirectory(prefix="litchi-docx-styles-source-pinned-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.root = self.directory / "checkout"
        self.root.mkdir()
        self.git("init", "-q")

        self.production = self.root / "crates/fixture/src/lib.rs"
        self.production.parent.mkdir(parents=True)
        self.production.write_text("pub fn value() -> u8 { 1 }\n")
        (self.production.parent.parent / "Cargo.toml").write_text(
            '[package]\nname = "fixture"\nversion = "0.1.0"\n'
        )

        self.evidence = self.root / "docs/evidence"
        self.harness = self.evidence / "harness/main.rs"
        self.harness.parent.mkdir(parents=True)
        self.harness.write_text("fn main() {}\n")
        (self.harness.parent / "Cargo.toml").write_text(
            '[package]\nname = "harness"\nversion = "0.1.0"\n'
        )
        self.extra = self.evidence / "fixture.txt"
        self.extra.write_text("native fixture bytes\n")

        # The approved production commit already contains the initial
        # evidence tree.  A descendant update must therefore bind to the
        # descendant HEAD rather than falling back to the old source blob.
        self.commit("production and initial evidence")
        self.source_commit = self.git("rev-parse", "HEAD").stdout.strip()
        self.metadata = self.directory / "metadata.json"
        self.metadata.write_text(json.dumps({"packages": [
            {"name": "fixture", "version": "0.1.0", "source": None,
             "manifest_path": str(self.production.parent.parent / "Cargo.toml"),
             "targets": [{"src_path": str(self.production)}]},
            {"name": "harness", "version": "0.1.0", "source": None,
             "manifest_path": str(self.harness.parent / "Cargo.toml"),
             "targets": [{"src_path": str(self.harness)}]},
        ]}))
        self.output = self.directory / "manifest.txt"

    def git(self, *arguments):
        return subprocess.run(
            ["git", "-C", str(self.root), *arguments], check=True,
            capture_output=True, text=True,
        )

    def commit(self, message):
        self.git("add", ".")
        self.git("-c", "user.name=Guard Test", "-c", "user.email=guard@example.invalid",
                 "commit", "-qm", message)

    def capture(self):
        return subprocess.run([
            sys.executable, "-B", str(Path(__file__).with_name("source_manifest.py")),
            "--metadata", str(self.metadata), "--root", str(self.root),
            "--source-commit", self.source_commit, "--evidence", str(self.evidence),
            "--extra", str(self.extra), "--output", str(self.output),
        ], capture_output=True, text=True)

    def assert_refused(self):
        run = self.capture()
        self.assertNotEqual(run.returncode, 0, "changed input was accepted")
        self.assertFalse(self.output.exists(), run.stderr)

    def test_committed_update_to_pinned_evidence_binds_descendant_head(self):
        self.harness.write_text("fn main() { /* descendant */ }\n")
        self.commit("updated pinned evidence")

        run = self.capture()
        self.assertEqual(run.returncode, 0, run.stderr)
        manifest = self.output.read_text()
        descendant_head = self.git("rev-parse", "HEAD").stdout.strip()
        self.assertIn(f"git_head={descendant_head}\n", manifest)
        self.assertIn("docs/evidence/harness/main.rs", manifest)

    def test_dirty_pinned_evidence_update_is_refused(self):
        self.harness.write_text("fn main() { /* dirty */ }\n")
        self.assert_refused()

    def test_untracked_pinned_evidence_is_refused(self):
        self.extra = self.evidence / "untracked.txt"
        self.extra.write_text("uncommitted evidence\n")
        self.assert_refused()

    def test_changed_production_input_remains_bound_to_source_commit(self):
        self.production.write_text("pub fn value() -> u8 { 2 }\n")
        self.commit("changed production")
        self.assert_refused()

    def test_new_production_input_outside_evidence_is_refused(self):
        (self.production.parent / "new.rs").write_text("pub const EXTRA: u8 = 1;\n")
        self.commit("new production input")
        self.assert_refused()


if __name__ == "__main__":
    unittest.main()
