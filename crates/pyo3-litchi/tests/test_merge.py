"""Template merge (record 0274): ``File.merge``.

Oracles: read-back through the unified read API, immutability of the
template File, and the summary's keys against the stub's ``_MergeSummary``.
"""

from __future__ import annotations

import unittest

from _support import L, typed_dict_fields


def template():
    file = L.create("docx")
    file, _ = file.add("/body", "paragraph", {"text": "Dear {{name}}, {{ order.total }} due."})
    return file


class MergeTests(unittest.TestCase):
    def test_a_mapping_fills_placeholders(self):
        file = template()
        before = file.to_bytes()
        commit = file.merge({"name": "Ada", "order": {"total": 5200.5}})
        self.assertTrue(commit.changed)
        summary = commit.value
        self.assertEqual(set(summary), set(typed_dict_fields()["_MergeSummary"]))
        self.assertEqual(summary["replacements"], 2)
        self.assertEqual(summary["keys"], ["name", "order.total"])
        self.assertEqual(summary["unresolved"], [])
        self.assertEqual(commit.file.text(), "Dear Ada, 5200.5 due.")
        self.assertEqual(file.to_bytes(), before)

    def test_json_text_and_missing_keys(self):
        file = template()
        kept = file.merge(b'{"name": "Grace"}')
        self.assertEqual(kept.file.text(), "Dear Grace, {{ order.total }} due.")
        self.assertEqual(kept.value["unresolved"], ["order.total"])
        emptied = file.merge('{"name": "Grace"}', missing="empty")
        self.assertEqual(emptied.file.text(), "Dear Grace,  due.")
        none = file.merge({"other": 1})
        self.assertFalse(none.changed)

    def test_every_format_and_errors(self):
        book = L.create("xlsx").set("/Sheet1/A1", {"value": "Hi {{who}}"})
        self.assertEqual(book.merge({"who": "you"}).file.get("/Sheet1/A1", depth=0)["text"], "Hi you")
        with self.assertRaises(L.PropertyError):
            template().merge("[1]")
        with self.assertRaises(L.PropertyError):
            template().merge({"a": 1}, missing="drop")
        with self.assertRaises(L.PropertyError):
            template().merge('[{"a": 1}, {"a": 2}]')
        with self.assertRaises(L.UnsupportedError):
            L.create("odg").merge({})
        with self.assertRaises(L.LitchiError) as caught:
            template().merge({"name": "x", "order": {"total": 1}}, max_replacements=1)
        self.assertEqual(caught.exception.code, "limit_exceeded")


    def test_more_formats_merge(self):
        # Record 0295: binary, legacy, RTF and ODF templates merge too.
        book = L.create("xlsb").set("/Sheet1/A1", {"value": "Hi {{who}}"})
        self.assertEqual(book.merge({"who": "you"}).file.get("/Sheet1/A1", depth=0)["text"], "Hi you")
        text, _ = L.create("odt").add("/body", "paragraph", {"text": "Dear {{who}}"})
        self.assertEqual(text.merge({"who": "Ada"}).file.text(), "Dear Ada")
        sheet = L.create("ods").set("/Sheet1/A1", {"value": "{{n}} items"})
        self.assertEqual(sheet.merge({"n": 3}).file.get("/Sheet1/A1", depth=0)["text"], "3 items")

    def test_find_replace_on_more_formats(self):
        text, _ = L.create("odt").add("/body", "paragraph", {"text": "draft one, draft two"})
        commit = text.apply({"command": "set", "path": "/", "props": {"find": "draft", "replace": "final"}})
        self.assertEqual(commit.file.text(), "final one, final two")


if __name__ == "__main__":
    unittest.main()
