from __future__ import annotations

import csv
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from tools import final_api_inventory as inventory


def rustdoc_fixture(crate: str) -> dict:
    return {
        "root": 0,
        "includes_private": False,
        "index": {"0": {
            "name": crate,
            "crate_id": 0,
            "inner": {"module": {"items": [], "is_crate": True}},
        }},
        "paths": {"0": {"path": [crate], "kind": "module", "crate_id": 0}},
    }


class FinalInventoryTests(unittest.TestCase):
    def make_package_root(self) -> Path:
        root = Path(self.tmp.name)
        for name, lib_name in (("alpha", "alpha"), ("odd-package", "odd_lib")):
            crate = root / "crates" / name
            (crate / "src").mkdir(parents=True)
            manifest = f'''[package]\nname = "{name}"\nversion = "0.1.0"\n'''
            if lib_name != name.replace("-", "_"):
                manifest += f'''[lib]\nname = "{lib_name}"\n'''
            (crate / "Cargo.toml").write_text(manifest, encoding="utf-8")
            (crate / "src" / "lib.rs").write_text("pub struct Fixture;\n", encoding="utf-8")
        return root

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.external = tempfile.TemporaryDirectory(dir=Path.cwd())

    def tearDown(self) -> None:
        self.tmp.cleanup()
        self.external.cleanup()

    def context(self, root: Path, source: str = "a" * 40) -> inventory.SourceContext:
        return inventory.SourceContext(
            root=str(root), source_commit=source, lock_sha256="b" * 64,
            rustc="rustc 1.95.0", cargo="cargo 1.95.0",
            rust_toolchain_sha256="c" * 64, cargo_path="fake-cargo", rustc_path="fake-rustc",
        )

    def test_enumerates_direct_library_packages_and_explicit_lib_name(self) -> None:
        root = self.make_package_root()
        packages = inventory.enumerate_packages(root)
        self.assertEqual([(item.name, item.lib_name) for item in packages],
                         [("alpha", "alpha"), ("odd-package", "odd_lib")])

    def test_capture_command_receipt_and_exact_identity_reuse(self) -> None:
        root = self.make_package_root()
        packages = inventory.enumerate_packages(root)
        package = packages[1]
        context = self.context(root)
        output = root / "external" / "run"
        target = root / "external" / "target"
        calls: list[tuple[list[str], dict[str, str]]] = []

        def fake_runner(command, cwd, env, stdout, stderr, resource_log):
            calls.append((list(command), dict(env)))
            target_arg = Path(command[command.index("--target-dir") + 1])
            target_arg.joinpath("doc").mkdir(parents=True, exist_ok=True)
            lib = package.lib_name
            target_arg.joinpath("doc", f"{lib}.json").write_text(
                json.dumps(rustdoc_fixture(lib)), encoding="utf-8")
            stdout.write_text("fake stdout\n", encoding="utf-8")
            stderr.write_text("", encoding="utf-8")
            resource_log.write_text("Maximum resident set size (kbytes): 1234\n",
                                    encoding="utf-8")
            return 0

        first = inventory.capture_package(package, context, output, target,
                                          cargo="fake-cargo", runner=fake_runner)
        self.assertFalse(first.reused)
        command, environment = calls[0]
        for flag in ("--all-features", "--offline", "--locked", "--document-hidden-items"):
            self.assertIn(flag, command)
        self.assertEqual(environment["CARGO_BUILD_JOBS"], "1")
        self.assertEqual(environment["CARGO_INCREMENTAL"], "0")
        self.assertEqual(environment["RUSTC_BOOTSTRAP"], "1")
        self.assertEqual(environment["LITCHI_MEM_MAX"], "20G")
        self.assertEqual(environment["LITCHI_CAPPED"], "")
        receipt = json.loads(Path(first.receipt_path).read_text(encoding="utf-8"))
        self.assertEqual(receipt["resource"]["peak_rss_kib"], 1234)

        external_tmp = Path(self.tmp.name) / "external" / "tmp"
        self.assertEqual(inventory.capture_env({"TMPDIR": "/tmp", "LITCHI_CAPPED": "1"},
                                               tmpdir=external_tmp)["TMPDIR"],
                         str(external_tmp))

        def should_not_run(*_args):
            raise AssertionError("matching JSON receipt should be reused")

        second = inventory.capture_package(package, context, output, target,
                                           cargo="fake-cargo", runner=should_not_run)
        self.assertTrue(second.reused)
        self.assertEqual(first.json_sha256, second.json_sha256)

    def test_changed_freeze_cannot_reuse_json(self) -> None:
        root = self.make_package_root()
        package = inventory.enumerate_packages(root)[0]
        output = root / "external" / "run"
        target = root / "external" / "target"
        calls = 0

        def fake_runner(command, cwd, env, stdout, stderr, resource_log):
            nonlocal calls
            calls += 1
            target_arg = Path(command[command.index("--target-dir") + 1])
            target_arg.joinpath("doc").mkdir(parents=True, exist_ok=True)
            target_arg.joinpath("doc", "alpha.json").write_text(
                json.dumps(rustdoc_fixture("alpha")), encoding="utf-8")
            stdout.write_text("", encoding="utf-8")
            stderr.write_text("", encoding="utf-8")
            resource_log.write_text("Maximum resident set size (kbytes): 1234\n",
                                    encoding="utf-8")
            return 0

        inventory.capture_package(package, self.context(root), output, target,
                                  cargo="fake-cargo", runner=fake_runner)
        inventory.capture_package(package, self.context(root, "d" * 40), output, target,
                                  cargo="fake-cargo", runner=fake_runner)
        self.assertEqual(calls, 2)

    def test_target_limits_use_decimal_bytes(self) -> None:
        target = Path(self.tmp.name) / "target"
        with self.assertRaises(inventory.InventoryError):
            inventory.check_target(target, inventory.MIN_TARGET_BYTES, label="preflight",
                                   size_fn=lambda _: 19_200_000_000)
        with self.assertRaises(inventory.InventoryError):
            inventory.check_target(target, inventory.MAX_TARGET_BYTES, label="hard cap",
                                   size_fn=lambda _: 20_000_000_000)

    def test_cargo_shim_is_verified_and_inherited_bypass_is_rejected(self) -> None:
        shim = Path(self.tmp.name) / "cargo"
        shim.write_text(
            "#!/bin/sh\n"
            "if [ -n \"$LITCHI_CAPPED\" ]; then exec cargo \"$@\"; fi\n"
            "systemd-run --scope -p MemoryMax=${LITCHI_MEM_MAX:-26G} cargo \"$@\"\n",
            encoding="utf-8")
        shim.chmod(0o755)
        environment = {"LITCHI_MEM_MAX": "20G", "LITCHI_CAPPED": ""}
        with patch.object(inventory.shutil, "which", return_value="/usr/bin/systemd-run"):
            identity = inventory.verify_cargo_shim(str(shim), environment)
        self.assertEqual(identity["memory_max"], "20G")
        self.assertTrue(identity["litchi_capped_unset"])
        with self.assertRaises(inventory.InventoryError):
            with patch.object(inventory.shutil, "which", return_value="/usr/bin/systemd-run"):
                inventory.verify_cargo_shim(str(shim), {"LITCHI_MEM_MAX": "20G",
                                                        "LITCHI_CAPPED": "1"})

    def test_capture_all_marks_inventory_skip_incomplete(self) -> None:
        root = Path(self.tmp.name) / "source"
        root.mkdir()
        output = Path(self.external.name) / "output"
        target = Path(self.tmp.name) / "target"
        context = self.context(root)
        packages = [inventory.Package(f"p{i}", f"p{i}", f"crates/p{i}/Cargo.toml", "d" * 64)
                    for i in range(inventory.DEFAULT_PACKAGE_COUNT)]

        tmpdirs = []

        def fake_capture(package, *_args, **kwargs):
            tmpdirs.append(kwargs["tmpdir"])
            return inventory.CaptureResult(package.name, str(output / f"{package.name}.json"),
                                            str(output / f"{package.name}.receipt.json"),
                                            False, 1, "e" * 64)

        with patch.object(inventory, "source_context", return_value=context), \
                patch.object(inventory, "enumerate_packages", return_value=packages), \
                patch.object(inventory, "verify_cargo_shim", return_value={}), \
                patch.object(inventory, "capture_package", side_effect=fake_capture), \
                patch.object(inventory, "source_stable") as stable:
            with self.assertRaises(inventory.InventoryError):
                inventory.capture_all(root, output, target, run_final_inventory=False)
        marker = json.loads((output / "run.json").read_text(encoding="utf-8"))
        self.assertEqual(marker["status"], "incomplete")
        self.assertEqual(marker["phase"], "failed")
        self.assertEqual(marker["package_count"], inventory.DEFAULT_PACKAGE_COUNT)
        self.assertEqual(marker["package_names"], [package.name for package in packages])
        self.assertEqual(tmpdirs, [output / "tmp"] * inventory.DEFAULT_PACKAGE_COUNT)
        self.assertTrue((output / "tmp").is_dir())
        self.assertGreaterEqual(stable.call_count, 2 * inventory.DEFAULT_PACKAGE_COUNT)

    def test_capture_all_rejects_tmp_output_without_touching_it(self) -> None:
        root = Path(self.tmp.name) / "source"
        root.mkdir()
        output = Path("/tmp") / f"litchi-final-inventory-{self.tmp.name.split('/')[-1]}"
        output.mkdir()
        sentinel = output / "keep.me"
        sentinel.write_text("unrelated\n", encoding="utf-8")
        try:
            with self.assertRaisesRegex(inventory.InventoryError, "must not be under /tmp"):
                inventory.capture_all(root, output, Path(self.tmp.name) / "target")
            self.assertEqual(sentinel.read_text(encoding="utf-8"), "unrelated\n")
            self.assertFalse((output / "tmp").exists())
        finally:
            sentinel.unlink()
            output.rmdir()

    def test_capture_all_requires_the_fixed_68_package_set(self) -> None:
        root = Path(self.tmp.name) / "source"
        root.mkdir()
        context = self.context(root)
        package = inventory.Package("only", "only", "crates/only/Cargo.toml", "d" * 64)
        with patch.object(inventory, "source_context", return_value=context), \
                patch.object(inventory, "enumerate_packages", return_value=[package]), \
                patch.object(inventory, "verify_cargo_shim") as verify:
            with self.assertRaisesRegex(inventory.InventoryError, "68 direct library packages"):
                inventory.capture_all(root, Path(self.external.name) / "output",
                                      Path(self.external.name) / "target")
        verify.assert_not_called()

    def test_capture_all_aborts_on_midrun_source_change(self) -> None:
        root = Path(self.tmp.name) / "source"
        root.mkdir()
        output = Path(self.external.name) / "output"
        target = Path(self.external.name) / "target"
        context = self.context(root)
        packages = [inventory.Package(f"p{i}", f"p{i}", f"crates/p{i}/Cargo.toml", "d" * 64)
                    for i in range(inventory.DEFAULT_PACKAGE_COUNT)]
        stable = Mock(side_effect=[None, inventory.InventoryError("changed")])

        with patch.object(inventory, "source_context", return_value=context), \
                patch.object(inventory, "enumerate_packages", return_value=packages), \
                patch.object(inventory, "verify_cargo_shim", return_value={}), \
                patch.object(inventory, "capture_package", return_value=inventory.CaptureResult(
                    "p0", "p0.json", "p0.receipt.json", False, 1, "e" * 64)), \
                patch.object(inventory, "source_stable", stable):
            with self.assertRaises(inventory.InventoryError):
                inventory.capture_all(root, output, target)
        marker = json.loads((output / "run.json").read_text(encoding="utf-8"))
        self.assertEqual(marker["status"], "incomplete")
        self.assertEqual(marker["captured"], 1)

    def test_disposition_scaffold_requires_explicit_decision(self) -> None:
        root = Path(self.tmp.name)
        source = root / "inventory.tsv"
        destination = root / "disposition.tsv"
        source.write_text("crate\tcategory\tpath\tdetail\ttarget\n"
                          "foo\tmut_owner\tfoo::Package::set\t&mut self\tEdit\n"
                          "bar\tmut_draft\tbar::builder::Edit::set\t&mut self\t\n",
                          encoding="utf-8")
        inventory.write_disposition_scaffold(source, destination)
        with destination.open(encoding="utf-8", newline="") as stream:
            rows = list(csv.DictReader(stream, delimiter="\t"))
        self.assertEqual(rows[0]["decision"], "UNASSIGNED_REQUIRED_REVIEW")
        self.assertEqual(rows[1]["decision"], "DRAFT_CANDIDATE_REVIEW")
        self.assertEqual(rows[0]["evidence"], "")


if __name__ == "__main__":
    unittest.main()
