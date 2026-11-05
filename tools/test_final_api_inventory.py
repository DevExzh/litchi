from __future__ import annotations

import csv
import json
from pathlib import Path
import tempfile
import unittest

from tools import final_api_inventory as inventory


def rustdoc_fixture(crate: str, *, includes_private: bool = False) -> dict:
    return {
        "root": 0,
        "includes_private": includes_private,
        "index": {"0": {
            "name": crate,
            "crate_id": 0,
            "inner": {"module": {"items": [], "is_crate": True}},
        }},
        "paths": {"0": {"path": [crate], "kind": "module", "crate_id": 0}},
    }


class FinalInventoryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name) / "checkout"
        for name, lib_name in (("alpha", "alpha"), ("odd-package", "odd_lib"), ("bin-only", None)):
            crate = self.root / "crates" / name
            (crate / "src").mkdir(parents=True)
            manifest = f'[package]\nname = "{name}"\nversion = "0.1.0"\n'
            if lib_name is None:
                (crate / "src" / "main.rs").write_text("fn main() {}\n", encoding="utf-8")
            else:
                if lib_name != name.replace("-", "_"):
                    manifest += f'[lib]\nname = "{lib_name}"\n'
                (crate / "src" / "lib.rs").write_text("pub struct Fixture;\n", encoding="utf-8")
            (crate / "Cargo.toml").write_text(manifest, encoding="utf-8")
        self.output = Path(self.tmp.name) / "out"
        self.target = Path(self.tmp.name) / "target"
        self.calls: list[list[str]] = []

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def runner(self, *, status: int = 0, includes_private: bool = False):
        def run(command, cwd, env, log_path):
            self.calls.append(list(command))
            self.assertEqual(env["RUSTC_BOOTSTRAP"], "1")
            log_path.write_text("log\n", encoding="utf-8")
            package = command[command.index("--package") + 1]
            lib = {"alpha": "alpha", "odd-package": "odd_lib"}[package]
            doc = Path(command[command.index("--target-dir") + 1]) / "doc"
            doc.mkdir(parents=True, exist_ok=True)
            (doc / f"{lib}.json").write_text(
                json.dumps(rustdoc_fixture(lib, includes_private=includes_private)),
                encoding="utf-8")
            return status
        return run

    def capture(self, *, commit: str = "a" * 40, reusable: bool = True, **runner_args):
        packages = inventory.enumerate_packages(self.root)
        return inventory.capture(self.root, self.output, self.target, packages,
                                 commit=commit, reusable=reusable,
                                 runner=self.runner(**runner_args), env={})

    def test_enumerates_library_packages_with_explicit_lib_name(self) -> None:
        packages = inventory.enumerate_packages(self.root)
        self.assertEqual([(p.name, p.lib_name) for p in packages],
                         [("alpha", "alpha"), ("odd-package", "odd_lib")])

    def test_capture_writes_json_and_manifest_and_reuses_same_commit(self) -> None:
        paths = self.capture()
        self.assertEqual([p.name for p in paths], ["alpha.json", "odd-package.json"])
        self.assertEqual(len(self.calls), 2)
        self.assertIn("--all-features", self.calls[0])
        self.assertIn("--document-hidden-items", self.calls[0])
        rows = inventory.read_manifest(self.output / "manifest.tsv")
        self.assertEqual(rows["odd-package"]["lib_name"], "odd_lib")
        self.assertEqual(rows["alpha"]["sha256"], inventory.sha256_file(paths[0]))
        self.capture()
        self.assertEqual(len(self.calls), 2, "same clean commit reuses every JSON")

    def test_changed_commit_or_dirty_checkout_recaptures(self) -> None:
        self.capture()
        self.capture(commit="b" * 40)
        self.assertEqual(len(self.calls), 4)
        self.capture(commit="b" * 40, reusable=False)
        self.assertEqual(len(self.calls), 6)

    def test_edited_json_is_recaptured(self) -> None:
        paths = self.capture()
        paths[0].write_text("{}", encoding="utf-8")
        self.capture()
        self.assertEqual(len(self.calls), 3)

    def test_rustdoc_failure_and_private_json_are_errors(self) -> None:
        with self.assertRaisesRegex(inventory.InventoryError, "rustdoc failed for alpha"):
            self.capture(status=101)
        with self.assertRaisesRegex(inventory.InventoryError, "private"):
            self.capture(includes_private=True)

    def test_output_inside_checkout_is_rejected(self) -> None:
        status = inventory.main(["--root", str(self.root), "--output",
                                 str(self.root / "inventory-out")])
        self.assertEqual(status, 2)
        self.assertFalse((self.root / "inventory-out").exists())

    def test_disposition_scaffold_requires_explicit_decision(self) -> None:
        source = Path(self.tmp.name) / "inventory.tsv"
        destination = Path(self.tmp.name) / "disposition.tsv"
        source.write_text("crate\tcategory\tpath\tdetail\ttarget\n"
                          "foo\tmut_owner\tfoo::Package::set\t&mut self\tEdit\n"
                          "bar\tmut_draft\tbar::builder::Edit::set\t&mut self\t\n"
                          "litchi-cfb\tmut_draft\tlitchi_cfb::Reader::open\t&mut self\t\n",
                          encoding="utf-8")
        inventory.write_disposition_scaffold(source, destination)
        with destination.open(encoding="utf-8", newline="") as stream:
            rows = list(csv.DictReader(stream, delimiter="\t"))
        self.assertEqual([r["decision"] for r in rows],
                         ["UNASSIGNED_REQUIRED_REVIEW", "DRAFT_CANDIDATE_REVIEW",
                          "SUBSTRATE_SCOPE_REVIEW"])
        self.assertEqual(rows[0]["evidence"], "")


if __name__ == "__main__":
    unittest.main()
