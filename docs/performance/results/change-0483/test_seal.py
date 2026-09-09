#!/usr/bin/env python3
"""Synthetic tests for the portable 0483 SHA256SUMS boundary."""

from __future__ import annotations

import hashlib
from pathlib import Path
import tempfile
import unittest

import seal


class SealTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        (self.root / "nested").mkdir()
        (self.root / "nested" / "b.txt").write_bytes(b"bravo\n")
        (self.root / "a.txt").write_bytes(b"alpha\n")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_seal_is_sorted_exclusive_and_round_trips(self) -> None:
        receipt = seal.seal_write(self.root)
        self.assertEqual(receipt["files"], 2)
        self.assertEqual(
            (self.root / seal.MANIFEST_NAME).read_text(encoding="utf-8"),
            "\n".join(
                f"{hashlib.sha256(data).hexdigest()}  {name}"
                for name, data in (("a.txt", b"alpha\n"), ("nested/b.txt", b"bravo\n"))
            )
            + "\n",
        )
        self.assertEqual(seal.verify_inventory(self.root)["files"], 2)
        with self.assertRaisesRegex(seal.SealError, "existing seal"):
            seal.seal_write(self.root)

    def test_missing_file_is_rejected(self) -> None:
        seal.seal_write(self.root)
        (self.root / "nested" / "b.txt").unlink()
        with self.assertRaisesRegex(seal.SealError, "file set differs"):
            seal.verify_inventory(self.root)

    def test_extra_file_is_rejected(self) -> None:
        seal.seal_write(self.root)
        (self.root / "extra.bin").write_bytes(b"extra")
        with self.assertRaisesRegex(seal.SealError, "file set differs"):
            seal.verify_inventory(self.root)

    def test_changed_bytes_are_rejected(self) -> None:
        seal.seal_write(self.root)
        (self.root / "a.txt").write_bytes(b"changed\n")
        with self.assertRaisesRegex(seal.SealError, "digest differs"):
            seal.verify_inventory(self.root)

    def test_symlinks_are_rejected_for_seal_and_verify(self) -> None:
        target = self.root / "a.txt"
        link = self.root / "link.txt"
        try:
            link.symlink_to(target)
        except (NotImplementedError, OSError) as error:
            self.skipTest(f"symlinks unavailable: {error}")
        with self.assertRaisesRegex(seal.SealError, "symlinks are forbidden"):
            seal.seal_write(self.root)
        link.unlink()
        seal.seal_write(self.root)
        link.symlink_to(target)
        with self.assertRaisesRegex(seal.SealError, "symlinks are forbidden"):
            seal.verify_inventory(self.root)

    def test_pycache_and_pyc_are_rejected(self) -> None:
        cache = self.root / "__pycache__"
        cache.mkdir()
        (cache / "module.cpython-314.pyc").write_bytes(b"bytecode")
        with self.assertRaisesRegex(seal.SealError, "__pycache__ is forbidden"):
            seal.inventory(self.root)

    def test_unsafe_manifest_path_is_rejected(self) -> None:
        with self.assertRaisesRegex(seal.SealError, "newline"):
            seal.render_manifest({"bad\nname": "0" * 64})
        with self.assertRaisesRegex(seal.SealError, "path escapes"):
            seal.parse_manifest(("0" * 64 + "  ../escape\n").encode())

    def test_manifest_tampering_is_rejected_even_when_records_parse(self) -> None:
        seal.seal_write(self.root)
        manifest = self.root / seal.MANIFEST_NAME
        manifest.write_bytes(manifest.read_bytes().replace(b"  a.txt\n", b" a.txt\n"))
        with self.assertRaisesRegex(seal.SealError, "malformed digest"):
            seal.verify_inventory(self.root)


if __name__ == "__main__":
    unittest.main()
