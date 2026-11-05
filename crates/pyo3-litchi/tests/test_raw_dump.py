"""``File.dump`` and the raw layer (record 0336): ``parts``, ``raw``,
``add_part``, ``raw_set`` and ``validate``.

Oracles: a dump replayed with ``batch`` onto ``create`` rebuilds the text and
dumps the same items again; raw writes are visible through the unified read
API and never change the input File; the returned dicts' keys match the
stub's TypedDicts.
"""

from __future__ import annotations

import unittest

from _support import L, fixture, typed_dict_fields


class DumpTests(unittest.TestCase):
    def test_dump_replays_onto_create(self):
        file = L.open(fixture("ooxml/docx/table_footnotes.docx"))
        items = file.dump()
        self.assertEqual(items[0], {"command": "meta", "dumpVersion": 2})
        replay = L.create("docx").batch(items).file
        self.assertEqual(replay.view("text"), file.view("text"))
        self.assertEqual(replay.dump(), items)

    def test_every_workbook_and_document_format_dumps(self):
        for kind in ("docx", "odt", "rtf", "xlsx", "xlsb", "ods", "pptx"):
            with self.subTest(kind=kind):
                items = L.create(kind).dump()
                self.assertEqual(items[0]["command"], "meta")
                self.assertEqual(L.create(kind).batch(items).file.dump(), items)

    def test_strict_and_refused(self):
        with self.assertRaises(L.UnsupportedError) as raised:
            L.create("odg").dump()
        self.assertEqual(raised.exception.code, "unsupported_capability")
        # A document whose dump skips nothing passes strict.
        L.create("docx").dump(strict=True)


class RawTests(unittest.TestCase):
    def test_parts_and_raw(self):
        file = L.create("docx")
        parts = file.parts()
        self.assertTrue(any(part["name"] == "/word/document.xml" for part in parts))
        self.assertEqual(set(parts[0]), set(typed_dict_fields()["_PartInfo"]))
        xml = file.raw("/document")
        self.assertIsInstance(xml, bytes)
        self.assertIn(b"<w:body", xml)
        with self.assertRaises(L.PathError):
            file.raw("/nothing")
        with self.assertRaises(L.UnsupportedError):
            L.create("rtf").raw("/document")

    def test_raw_set_and_add_part(self):
        file = L.create("docx")
        before = file.to_bytes()
        commit = file.raw_set(
            "/document", "//w:body", "prepend", "<w:p><w:r><w:t>Raw</w:t></w:r></w:p>"
        )
        self.assertEqual(commit.value, {"matched": 1})
        # Raw edits are not transaction steps: no patch (record 0347).
        self.assertIsNone(commit.patch)
        self.assertEqual(commit.file.text().strip(), "Raw")
        self.assertEqual(file.to_bytes(), before)
        with self.assertRaises(L.PathError) as raised:
            file.raw_set("/document", "//w:tbl", "remove")
        self.assertEqual(raised.exception.code, "zero_matches")

        added = file.add_part("/document", "header")
        self.assertEqual(added.value["part"], "/word/header1.xml")
        self.assertIsNone(added.patch)
        self.assertTrue(added.value["relationshipId"].startswith("rId"))
        self.assertIn(b"<w:hdr", added.file.raw("/header[1]"))
        custom = file.add_part(
            "/",
            content_type="application/xml",
            relationship="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
            data=b"<root/>",
            name="/customXml/item9.xml",
        )
        self.assertEqual(custom.file.raw("/customXml/item9.xml"), b"<root/>")
        with self.assertRaises(L.PropertyError):
            file.add_part("/", content_type="application/xml")

    def test_validate(self):
        report = L.create("xlsx").validate()
        self.assertEqual(set(report), set(typed_dict_fields()["_Validation"]))
        self.assertTrue(report["valid"])
        self.assertEqual(report["format"], "xlsx")
        for kind in ("doc", "xls", "ppt", "rtf", "odt"):
            with self.subTest(kind=kind):
                self.assertTrue(L.create(kind).validate()["valid"])


if __name__ == "__main__":
    unittest.main()
