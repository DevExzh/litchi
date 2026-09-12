"""Exercise manifest capture and replay against disposable real Git trees."""

import hashlib
import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest

import verify


class SourceSnapshotTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="litchi-xlsx-source-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.root = self.directory / "checkout"
        self.root.mkdir()
        self.git("init", "-q")
        self.package = self.root / "crates/fixture"
        self.source = self.package / "src/lib.rs"
        self.source.parent.mkdir(parents=True)
        self.source.write_text("pub fn value() -> u8 { 1 }\n")
        for index in range(4):
            (self.source.parent / f"helper{index}.rs").write_text("// retained source\n")
        (self.package / "Cargo.toml").write_text(
            '[package]\nname = "fixture"\nversion = "0.1.0"\n'
        )
        (self.package / "build.rs").write_text("fn main() {}\n")
        self.extra = self.root / "fixture.txt"
        self.extra.write_text("retained fixture\n")
        self.git("add", ".")
        self.git("-c", "user.name=Guard Test", "-c", "user.email=guard@example.invalid",
                 "commit", "-qm", "fixture")
        self.commit = self.git("rev-parse", "HEAD").stdout.strip()
        self.metadata = self.directory / "metadata.json"
        self.metadata.write_text(json.dumps({"packages": [{
            "name": "fixture", "version": "0.1.0", "source": None,
            "manifest_path": str(self.package / "Cargo.toml"),
            "targets": [{"src_path": str(self.source)}],
        }]}))
        self.manifest = self.directory / "manifest.txt"

    def git(self, *arguments):
        return subprocess.run(
            ["git", "-C", str(self.root), *arguments], check=True,
            capture_output=True, text=True,
        )

    def capture(self):
        return subprocess.run([
            sys.executable, "-B", str(Path(__file__).with_name("source_manifest.py")),
            "--metadata", str(self.metadata), "--root", str(self.root),
            "--git-commit", self.commit, "--extra", str(self.extra),
            "--output", str(self.manifest),
        ], capture_output=True, text=True)

    def test_committed_local_cargo_sources_and_extra_replay(self):
        run = self.capture()
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertEqual(verify.verify_manifest(self.manifest, self.root), 8)

    def test_capture_rejects_untracked_discovered_source_before_output(self):
        (self.source.parent / "untracked.rs").write_text("pub const EXTRA: u8 = 2;\n")
        run = self.capture()
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("not tracked", run.stderr)
        self.assertFalse(self.manifest.exists())

    def test_capture_and_replay_reject_assume_unchanged_source(self):
        self.assertEqual(self.capture().returncode, 0)
        self.git("update-index", "--assume-unchanged", "crates/fixture/src/lib.rs")
        self.source.write_text("pub fn value() -> u8 { 2 }\n")
        self.git("diff", "--quiet")
        run = self.capture()
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("committed Git blobs", run.stderr)
        with self.assertRaises(AssertionError):
            verify.verify_manifest(self.manifest, self.root)

    def test_replay_keeps_path_source_kind_after_manifest_hashes_are_updated(self):
        self.assertEqual(self.capture().returncode, 0)
        self.source.write_text("pub fn value() -> u8 { 9 }\n")
        lines = self.manifest.read_text().splitlines()
        entries = []
        for index, line in enumerate(lines):
            if not line.startswith("file="):
                continue
            fields = line.split("\t")
            fields[4] = hashlib.sha256((self.root / fields[3]).read_bytes()).hexdigest()
            entries.append((fields[3], fields[4]))
            lines[index] = "\t".join(fields)
        tree = hashlib.sha256("\n".join(f"{p}\t{h}" for p, h in entries).encode()).hexdigest()
        for index, line in enumerate(lines):
            if line.startswith("package="):
                fields = line.split("\t")
                self.assertEqual(fields[2], "path")
                fields[6] = tree
                lines[index] = "\t".join(fields)
        self.manifest.write_text("\n".join(lines) + "\n")
        with self.assertRaisesRegex(AssertionError, "committed snapshot"):
            verify.verify_manifest(self.manifest, self.root)

    def test_capture_refuses_untracked_explicit_extra(self):
        self.extra = self.root / "new-fixture.txt"
        self.extra.write_text("uncommitted fixture\n")
        run = self.capture()
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("not tracked", run.stderr)
        self.assertFalse(self.manifest.exists())


if __name__ == "__main__":
    unittest.main()
