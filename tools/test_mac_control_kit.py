"""Tests for the macOS failure control kit (record 0431)."""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools" / "mac_queue"))

import control_kit  # noqa: E402

BASIC = ROOT / "test-data" / "iwork" / "numbers" / "basic.numbers"
CHART = ROOT / "test-data" / "iwork" / "numbers" / "chart-data-native.numbers"


class WireTests(unittest.TestCase):
    def test_iwa_literal_encoding_round_trips(self) -> None:
        for size in (0, 1, 59, 60, 300, 65536, 200_000):
            data = bytes(index % 251 for index in range(size))
            self.assertEqual(control_kit.iwa_decode(control_kit.iwa_encode(data)), data)

    def test_archive_round_trip_keeps_every_object(self) -> None:
        package = control_kit.Package(CHART.read_bytes())
        rebuilt = control_kit.Package(package.to_bytes(set(package.archives)))
        self.assertEqual(
            [(name, item["id"]) for name, item in package.objects()],
            [(name, item["id"]) for name, item in rebuilt.objects()],
        )
        for (_, left), (_, right) in zip(package.objects(), rebuilt.objects()):
            self.assertEqual(
                [message["data"] for message in left["messages"]],
                [message["data"] for message in right["messages"]],
            )


class AuditTests(unittest.TestCase):
    def test_native_files_list_every_reference(self) -> None:
        for path in (BASIC, CHART):
            self.assertEqual(
                control_kit.unlisted_references(control_kit.Package(path.read_bytes())), []
            )

    def test_native_chart_has_an_empty_sparse_array(self) -> None:
        arrays = control_kit.sparse_arrays(control_kit.Package(CHART.read_bytes()))
        self.assertEqual([(count, entries) for _, count, entries in arrays], [(0, [])])

    def test_variants_refuse_inputs_without_the_defect(self) -> None:
        with self.assertRaises(ValueError):
            control_kit.variant_list_references(BASIC.read_bytes())
        with self.assertRaises(ValueError):
            control_kit.variant_sparse_count(CHART.read_bytes(), 3)
        with self.assertRaises(ValueError):
            control_kit.variant_drop_objects(BASIC.read_bytes(), {1})


class AddTests(unittest.TestCase):
    def test_add_appends_and_refuses_a_repeat(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            kit = root / "kit"
            kit.mkdir()
            (kit / "manifest.json").write_text(json.dumps({"packages": []}))
            (kit / "queue.tsv").write_text("package_id\tapp\tpath\tsha256\trows\n")
            shutil.copy2(BASIC, root / "extra.numbers")
            listing = root / "list.tsv"
            listing.write_text(
                "id\tapp\tfile\texpect\tsource\n"
                "N-1\tNumbers\textra.numbers\topens\ta test\n"
            )
            self.assertEqual(control_kit.add_packages(kit, listing), 0)
            self.assertTrue((kit / "packages/Numbers/N-1.numbers").is_file())
            self.assertIn("N-1\tNumbers", (kit / "queue.tsv").read_text())
            with self.assertRaises(ValueError):
                control_kit.add_packages(kit, listing)


if __name__ == "__main__":
    unittest.main()
