"""Field refresh (record 0279): ``File.refresh``.

Oracles: the summary's keys against the stub's ``_RefreshSummary``, the
written formula cache read back through ``evaluate``, the exact no-op, and
immutability of the refreshed File.
"""

from __future__ import annotations

import datetime
import unittest

from _support import L, typed_dict_fields


def workbook():
    book = L.create("xlsx").set("/Sheet1/A1", {"value": 21})
    return book.set("/Sheet1/A2", {"formula": "A1*2"})


class RefreshTests(unittest.TestCase):
    def test_a_workbook_reports_then_writes_its_caches(self):
        book = workbook()
        before = book.to_bytes()
        report = book.refresh()
        self.assertFalse(report.changed)
        summary = report.value
        self.assertEqual(set(summary), set(typed_dict_fields()["_RefreshSummary"]))
        field = summary["fields"][0]
        self.assertEqual(set(field), set(typed_dict_fields()["_RefreshField"]))
        self.assertEqual((field["type"], field["path"], field["status"]), ("FORMULA", "Sheet1!A2", "stale"))
        written = book.refresh(write_cache=True)
        self.assertTrue(written.changed)
        self.assertEqual(written.value["refreshed"], 1)
        self.assertEqual(written.file.evaluate("/Sheet1/A2"), {"value": 42.0, "origin": "cached"})
        self.assertFalse(written.file.refresh(write_cache=True).changed)
        self.assertEqual(book.to_bytes(), before)

    def test_a_document_without_fields_is_unchanged(self):
        doc = L.create("docx")
        doc, _ = doc.add("/body", "paragraph", {"text": "Plain"})
        commit = doc.refresh(clock=datetime.datetime(2026, 10, 14, 9, 30))
        self.assertFalse(commit.changed)
        self.assertEqual(commit.value["fields"], [])
        self.assertFalse(doc.refresh(clock=datetime.date(2026, 10, 14)).changed)

    def test_options_are_checked(self):
        doc = L.create("docx")
        with self.assertRaises(L.PropertyError):
            doc.refresh(clock="2026-02-30")
        with self.assertRaises(L.PropertyError):
            doc.refresh(page_numbers="guess")
        with self.assertRaises(L.PropertyError):
            doc.refresh(clock=3)
        with self.assertRaises(L.UnsupportedError):
            L.create("rtf").refresh()


if __name__ == "__main__":
    unittest.main()
