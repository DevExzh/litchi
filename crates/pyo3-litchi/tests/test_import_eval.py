"""CSV/TSV import and formula evaluation (record 0266): ``File.import_csv``,
``File.evaluate`` and ``File.recalculate``.

Oracles: read-back through the unified read API, immutability of the input
File, the returned dicts' keys against the stub's TypedDicts, and the cached
values the producing application stored.
"""

from __future__ import annotations

import unittest

from _support import L, fixture, typed_dict_fields

WORKBOOKS = ("xlsx", "xlsb", "xls", "ods")


def cell_text(file, path):
    return file.get(path, depth=0)["text"]


class ImportTests(unittest.TestCase):
    def test_typed_import_into_every_workbook_format(self):
        for kind in WORKBOOKS:
            with self.subTest(kind=kind):
                file = L.create(kind)
                before = file.to_bytes()
                commit = file.import_csv("/sheet[1]/B2", "name,score\nAda,36\n", header=True)
                self.assertTrue(commit.changed)
                summary = commit.value
                self.assertEqual(set(summary), set(typed_dict_fields()["_ImportSummary"]))
                self.assertEqual((summary["row"], summary["column"]), (1, 1))
                self.assertEqual(summary["numbers"], 1)
                self.assertEqual(cell_text(commit.file, "/sheet[1]/C3"), "36")
                self.assertEqual(file.to_bytes(), before)

    def test_options_bytes_and_noop(self):
        file = L.create("xlsx")
        data = "﻿a\t007\n".encode("utf-16-le")
        commit = file.import_csv("/Sheet1", data, delimiter="tab", infer=False)
        self.assertEqual(cell_text(commit.file, "/Sheet1/B1"), "007")
        empty = file.import_csv("/Sheet1", b"")
        self.assertFalse(empty.changed)
        formula = file.import_csv("/Sheet1", "2,=A1*21", formulas=True)
        self.assertEqual(formula.value["formulas"], 1)

    def test_errors_are_typed(self):
        file = L.create("xlsx")
        # Malformed text is "invalid_input", a CorruptFileError.
        with self.assertRaises(L.CorruptFileError) as caught:
            file.import_csv("/Sheet1", b'"open')
        self.assertEqual(caught.exception.code, "invalid_input")
        with self.assertRaises(L.PropertyError):
            file.import_csv("/Sheet1", "a", delimiter=";;")
        with self.assertRaises(L.LitchiError) as caught:
            file.import_csv("/Sheet1", "1,2,3", max_cells=2)
        self.assertEqual(caught.exception.code, "limit_exceeded")
        with self.assertRaises(L.UnsupportedError):
            L.create("docx").import_csv("/body", "a")


class EvaluateTests(unittest.TestCase):
    def test_evaluate_and_recalculate(self):
        file = L.create("xlsx").import_csv("/Sheet1", "20,=A1+22,=A1/0", formulas=True).file
        result = file.evaluate("/Sheet1/B1")
        self.assertEqual(set(result), set(typed_dict_fields()["_Evaluated"]))
        self.assertEqual(result, {"value": 42.0, "origin": "computed"})
        self.assertEqual(file.evaluate("/Sheet1/C1")["value"], "#DIV/0!")
        self.assertEqual(file.evaluate("/Sheet1/A1"), {"value": 20.0, "origin": "literal"})
        report = file.recalculate()
        self.assertEqual(set(report), set(typed_dict_fields()["_Recalculation"]))
        self.assertEqual(report["uncached"], 2)
        cell = report["cells"][0]
        self.assertEqual(set(cell), set(typed_dict_fields()["_RecalculatedCell"]))
        self.assertEqual(cell["computed_value"], 42.0)

    def test_cached_corpus_values_match(self):
        file = L.open(fixture("ooxml/xlsx/shared_formulas.xlsx"))
        report = file.recalculate()
        self.assertEqual(report["formulas"], report["match"])
        first = report["cells"][0]
        self.assertEqual(first["cached_value"], first["computed_value"])

    def test_unsupported_formulas_raise(self):
        file = L.create("xlsx").set("/Sheet1/A1", {"formula": 'A2&"x"'})
        with self.assertRaises(L.UnsupportedError):
            file.evaluate("/Sheet1/A1")


if __name__ == "__main__":
    unittest.main()
