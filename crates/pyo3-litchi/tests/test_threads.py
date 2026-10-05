"""Handles used from several threads (bug B5, fixed in S8a).

Every class is ``frozen`` and thread-safe over the ``Send + Sync``
``litchi::File``, and the document-wide reads release the GIL. These tests
share handles across threads and check that every thread reads the same
values the opening thread did. They run in a subprocess, as the B5 test does,
so a regression that panics cannot abort the whole run (release builds use
``panic = "abort"``).
"""

from __future__ import annotations

import unittest

from _support import TEST_DATA, run_isolated

SHARED_READS = '''
import sys, threading
import litchi_py as L

root = {root!r} + "/"
handles = {{
    "docx": L.Document.open(root + "ooxml/docx/table-alignment.docx"),
    "doc": L.Document.open(root + "ole/doc/HeaderFooterUnicode.doc"),
    "pages": L.Document.open(root + "iwork/pages/body-sections-unicode.pages"),
    "pptx": L.Presentation.open(root + "ooxml/pptx/shapes.pptx"),
    "ppt": L.Presentation.open(root + "ole/ppt/SampleShow.ppt"),
    "key": L.Presentation.open(root + "iwork/keynote/basic.key"),
    "xlsx": L.Workbook.open(root + "ooxml/xlsx/SimpleNormal.xlsx"),
    "xls": L.Workbook.open(root + "ole/xls/SimpleMultiCell.xls"),
    "numbers": L.Workbook.open(root + "iwork/numbers/cell-value-native.numbers"),
}}
paragraph = handles["docx"].paragraphs()[4]
table = handles["docx"].tables()[0]
slide = handles["pptx"].slides()[1]

def snapshot():
    values = []
    for name in ("docx", "doc", "pages"):
        document = handles[name]
        values.append(document.text())
        values.append([p.text() for p in document.paragraphs()])
        values.append([[c.text() for r in t.rows() for c in r.cells()] for t in document.tables()])
    for name in ("pptx", "ppt", "key"):
        presentation = handles[name]
        values += [presentation.text(), presentation.slide_count(), [s.text() for s in presentation.slides()]]
    for name in ("xlsx", "xls", "numbers"):
        workbook = handles[name]
        values += [workbook.text(), workbook.worksheet_count(), workbook.worksheet_names()]
    values.append([(r.text(), r.bold(), r.italic(), r.underline()) for r in paragraph.runs()])
    values.append([c.text() for r in table.rows() for c in r.cells()])
    values.append((slide.text(), repr(slide), repr(paragraph)))
    return values

expected = snapshot()
matches, failures = [], []

def work():
    try:
        for _ in range({rounds}):
            matches.append(snapshot() == expected)
    except BaseException as error:
        failures.append(repr(error))

threads = [threading.Thread(target=work) for _ in range({threads})]
for thread in threads:
    thread.start()
for thread in threads:
    thread.join()
print(len(matches), all(matches), failures)
'''

OPEN_IN_A_WORKER = '''
import threading
import litchi_py as L

root = {root!r} + "/"
opened = {{}}

def open_all():
    opened["document"] = L.Document.open(root + "ooxml/docx/Hyperlink.docx")
    opened["paragraphs"] = opened["document"].paragraphs()
    opened["workbook"] = L.Workbook.open(root + "iwork/numbers/basic.numbers")
    try:
        L.Presentation.open(root + "ooxml/docx/Hyperlink.docx")
    except L.UnsupportedError as error:
        opened["error"] = error

worker = threading.Thread(target=open_all)
worker.start()
worker.join()
# Handles and the exception made on the worker are used on the main thread.
print(repr(opened["document"].text()))
print([p.text() for p in opened["paragraphs"]])
print(opened["workbook"].worksheet_names())
print(opened["error"].code)
'''


class CrossThreadTests(unittest.TestCase):
    def test_shared_handles_read_the_same_values_on_every_thread(self):
        result = run_isolated(SHARED_READS.format(root=str(TEST_DATA), rounds=5, threads=8))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), "40 True []", result.stderr)

    def test_handles_opened_on_a_worker_are_read_on_the_main_thread(self):
        result = run_isolated(OPEN_IN_A_WORKER.format(root=str(TEST_DATA)))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(
            result.stdout.splitlines(),
            [
                repr("Following is a hyperlink.\nEricWhite.com"),
                str(["Following is a hyperlink.", "EricWhite.com"]),
                str(["Sheet 1"]),
                "unsupported_format",
            ],
        )


if __name__ == "__main__":
    unittest.main()
