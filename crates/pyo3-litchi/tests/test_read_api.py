"""Read-API oracle per format: Document, Presentation and Workbook families.

Every assertion here is a stable fact of a real fixture under test-data/
(counts, exact short strings, sheet names, run formatting). S8a must keep them
green: the Python read API stays source-compatible (design §7) while the
bindings move onto ``litchi::File``. ``ReprSnapshotTests`` is the one snapshot
class: its strings may change, but only deliberately and on the record.

Formats the bindings accept:
  Document     docx, doc, rtf, odt, pages
  Presentation pptx, ppt, odp, key
  Workbook     xlsx, xlsb, xls, ods, numbers
``.ppt`` and ``.xls`` are accepted since S8a (bug B1, see test_known_drift.py).
Opening from bytes is not offered by the bindings; only detection takes bytes
(test_formats_and_values.py).
"""

from __future__ import annotations

import gc
import os
import tempfile
import unittest
from pathlib import Path

from _support import (
    TEST_DATA,
    L,
    copy_fixture,
    fixture,
    paragraph_texts,
    replace_atomically,
    run_facts,
    slide_texts,
    table_grid,
    text_lines,
)


def assert_in_order(test: unittest.TestCase, text: str, parts: list) -> None:
    """Assert that each part occurs in ``text``, in order, without overlap."""
    position = 0
    for part in parts:
        found = text.find(part, position)
        test.assertGreaterEqual(found, 0, f"{part!r} missing after offset {position} of {text!r}")
        position = found + len(part)


# ---------------------------------------------------------------------------
# Document family
# ---------------------------------------------------------------------------


class DocxReadTests(unittest.TestCase):
    HYPERLINK = "ooxml/docx/Hyperlink.docx"
    TABLES = "ooxml/docx/table-alignment.docx"
    GRID = [
        ["Loren", "Ipsum", "Dolor"],
        ["Amet", "Consectetur", "adipiscing"],
        ["Quisque", "faucibus", "ex"],
    ]

    def test_open_accepts_str_pathlib_and_keyword(self):
        path = fixture(self.HYPERLINK)
        for document in (
            L.Document.open(path),
            L.Document.open(Path(path)),
            L.Document.open(path=path),
        ):
            self.assertIsInstance(document, L.Document)
            self.assertEqual(
                paragraph_texts(document), ["Following is a hyperlink.", "EricWhite.com"]
            )

    def test_text_holds_every_paragraph_in_order(self):
        # The separator between paragraphs is pinned by bug B3 (test_known_drift)
        # and test_text_is_one_line_per_paragraph.
        for relative in (self.HYPERLINK, self.TABLES):
            document = L.Document.open(fixture(relative))
            assert_in_order(self, document.text(), paragraph_texts(document))

    def test_text_is_one_line_per_paragraph(self):
        # S8a (bug B3): File::text() joins the paragraph texts with "\n", empty
        # and table-cell paragraphs included.
        self.assertEqual(
            L.Document.open(fixture(self.HYPERLINK)).text(),
            "Following is a hyperlink.\nEricWhite.com",
        )
        for relative in (self.HYPERLINK, self.TABLES):
            document = L.Document.open(fixture(relative))
            self.assertEqual(document.text().split("\n"), paragraph_texts(document))

    def test_paragraphs_and_runs(self):
        paragraphs = L.Document.open(fixture(self.HYPERLINK)).paragraphs()
        self.assertEqual(len(paragraphs), 2)
        self.assertEqual(
            [run_facts(p) for p in paragraphs],
            [[("Following is a hyperlink.", None, None)], [("EricWhite.com", None, None)]],
        )

    def test_run_formatting_is_tri_state(self):
        paragraphs = L.Document.open(fixture(self.TABLES)).paragraphs()
        self.assertEqual(run_facts(paragraphs[0]), [("Loren", True, None)])
        self.assertEqual(run_facts(paragraphs[3]), [("Amet", None, None)])
        self.assertEqual(
            run_facts(paragraphs[4]),
            [("Cons", None, None), ("ectet", True, True), ("ur", None, None)],
        )

    def test_tables_rows_and_cells(self):
        document = L.Document.open(fixture(self.TABLES))
        tables = document.tables()
        self.assertEqual(len(tables), 6)
        for table in tables:
            self.assertEqual(table.row_count(), 3)
            self.assertEqual(table_grid(table), self.GRID)
        self.assertEqual(L.Document.open(fixture(self.HYPERLINK)).tables(), [])

    def test_paragraphs_include_table_cell_paragraphs(self):
        texts = paragraph_texts(L.Document.open(fixture(self.TABLES)))
        self.assertEqual(len(texts), 60)
        self.assertEqual(
            texts[:10],
            ["Loren", "Ipsum", "Dolor", "Amet", "Consectetur", "adipiscing",
             "Quisque", "faucibus", "ex", ""],
        )


class DocReadTests(unittest.TestCase):
    NO_HEAD_FOOT = "ole/doc/NoHeadFoot.doc"
    UNICODE = "ole/doc/HeaderFooterUnicode.doc"
    TABLE = "ole/doc/commented-table.doc"
    PROPERTIES = "ole/doc/documentProperties.doc"

    def test_text_keeps_word_paragraph_marks(self):
        # Legacy Word text keeps the raw paragraph mark (\r) after each paragraph.
        document = L.Document.open(fixture(self.PROPERTIES))
        self.assertEqual(document.text(), "This is document text\r")
        self.assertEqual(paragraph_texts(document), ["This is document text"])

    def test_paragraphs(self):
        document = L.Document.open(fixture(self.NO_HEAD_FOOT))
        texts = paragraph_texts(document)
        self.assertEqual(len(texts), 10)
        self.assertEqual(
            texts[0],
            "This is a sample word document. It has two pages, but no headers or footers.",
        )
        self.assertEqual(texts[2], "HEADING TEXT")
        self.assertEqual(texts[4], "More on page one")
        self.assertEqual(texts[7], "End of page 1")
        self.assertEqual(texts[9], "This is page two. It also lacks headers and footers.")
        self.assertEqual(document.tables(), [])

    def test_run_formatting(self):
        paragraphs = L.Document.open(fixture(self.NO_HEAD_FOOT)).paragraphs()
        self.assertEqual(run_facts(paragraphs[2]), [("HEADING TEXT", True, None)])
        self.assertEqual(run_facts(paragraphs[4]), [("More on page one", None, None)])

    def test_unicode_text_and_formatting(self):
        document = L.Document.open(fixture(self.UNICODE))
        paragraphs = document.paragraphs()
        self.assertEqual(len(paragraphs), 34)
        self.assertEqual(paragraphs[3].text(), "\tGBP - £")
        self.assertEqual(paragraphs[4].text(), "\tEUR - €")
        self.assertEqual(run_facts(paragraphs[6]), [("\t", True, None), ("Molière", True, None)])
        self.assertEqual(
            run_facts(paragraphs[11]),
            [
                ("This is page two. ", None, None),
                ("Les Précieuses ridicules. ", None, True),
                ("The end.", None, None),
            ],
        )
        text = document.text()
        for expected in ("GBP - £", "EUR - €", "Molière", "L'Avare ou l'École du mensonge"):
            self.assertIn(expected, text)

    def test_table(self):
        document = L.Document.open(fixture(self.TABLE))
        tables = document.tables()
        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0].row_count(), 1)
        self.assertEqual(table_grid(tables[0]), [["A1", "B1"]])
        self.assertEqual(paragraph_texts(document)[:3], ["Before.", "A1", "B1"])


class RtfReadTests(unittest.TestCase):
    STYLES = "rtf/testStyles.rtf"
    HELLO = "rtf/hyperlink.rtf"
    TABLE = "rtf/table-cell-horizontal-merge.rtf"
    STYLE_LINES = [
        "Test 1 - normal.", "Test 2 - italic.", "Test 3 - bold.", "Test 4 -normal.",
        "Test 5 - Calibri 11.", "Test 6 - Arial 11.", "Test 7 - Calibri 11.",
        "Test 6 - Arial 12.", "Test 7 - Calibri 11.",
    ]

    def test_text(self):
        self.assertEqual(L.Document.open(fixture(self.HELLO)).text(), "Hello!\n")
        text = L.Document.open(fixture(self.STYLES)).text()
        self.assertTrue(text.startswith("\n".join(self.STYLE_LINES) + "\n"), repr(text))

    def test_paragraphs_and_runs(self):
        paragraphs = L.Document.open(fixture(self.HELLO)).paragraphs()
        self.assertEqual([p.text() for p in paragraphs], ["Hello!"])
        self.assertEqual(run_facts(paragraphs[0]), [("Hello", False, False), ("!", False, False)])

    def test_run_formatting_reports_explicit_false(self):
        # RTF answers False (not None) where the source has no bold/italic.
        paragraphs = L.Document.open(fixture(self.STYLES)).paragraphs()
        self.assertEqual([p.text() for p in paragraphs], self.STYLE_LINES)
        self.assertEqual(run_facts(paragraphs[0]), [("Test 1 - normal.", False, False)])
        self.assertEqual(run_facts(paragraphs[1]), [("Test 2 - italic.", False, True)])
        self.assertEqual(run_facts(paragraphs[2]), [("Test 3 - bold.", True, False)])

    def test_table(self):
        document = L.Document.open(fixture(self.TABLE))
        tables = document.tables()
        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0].row_count(), 1)
        self.assertEqual(table_grid(tables[0]), [["A1", ""]])
        # Unlike DOCX, DOC and ODT, RTF paragraphs() leaves out table-cell paragraphs.
        self.assertEqual(document.paragraphs(), [])


class OdtReadTests(unittest.TestCase):
    TABLE = "odf/corpus/writer-table.odt"
    STYLES = "odf/corpus/writer-paragraph-styles.odt"
    SPANS = "odf/odt/table-cell-column-span.odt"

    def test_text(self):
        self.assertEqual(
            L.Document.open(fixture(self.TABLE)).text(),
            "Cell 1\nCell 2\nCell 3 (bold)\nCell 4 (italic)\n",
        )

    def test_paragraphs(self):
        self.assertEqual(
            paragraph_texts(L.Document.open(fixture(self.STYLES))),
            [
                "Plain text", "Bold", "Italic", "Bold italic", "Underline",
                "Underline italic", "Underline bold italic", "Km2 - superscript",
                "H2O - subscript", "Strike-through",
            ],
        )

    def test_paragraphs_expose_no_runs(self):
        # Today ODT paragraphs return no runs, even for styled text.
        for relative in (self.TABLE, self.STYLES):
            for paragraph in L.Document.open(fixture(relative)).paragraphs():
                self.assertEqual(paragraph.runs(), [], paragraph.text())

    def test_table(self):
        document = L.Document.open(fixture(self.TABLE))
        tables = document.tables()
        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0].row_count(), 2)
        self.assertEqual(
            table_grid(tables[0]), [["Cell 1", "Cell 2"], ["Cell 3 (bold)", "Cell 4 (italic)"]]
        )
        self.assertEqual(
            paragraph_texts(document), ["Cell 1", "Cell 2", "Cell 3 (bold)", "Cell 4 (italic)", ""]
        )

    def test_column_spanned_cells(self):
        tables = L.Document.open(fixture(self.SPANS)).tables()
        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0].row_count(), 5)
        self.assertEqual(
            table_grid(tables[0]), [["x"], ["x", "", "x"], ["x"], ["x"], ["x", "", "x"]]
        )


class PagesReadTests(unittest.TestCase):
    """Pages through ``Document`` (record 0003 R6; since record 0306 the class reads the Pages ``File``)."""

    BASIC = "iwork/pages/basic.pages"
    UNICODE = "iwork/pages/body-sections-unicode.pages"
    TABLES = "iwork/pages/body-table-cells-native.pages"
    BASIC_TEXT = "Litchi native Pages fixture\nBuffa lazy-view migration verification\n2026-08-07"

    def test_text(self):
        self.assertEqual(L.Document.open(fixture(self.BASIC)).text(), self.BASIC_TEXT)
        self.assertEqual(
            L.Document.open(fixture(self.UNICODE)).text(),
            "Pages borrowed body 😀\nFirst section marker.\nSecond section marker — end.",
        )

    def test_one_paragraph_per_text_storage(self):
        paragraphs = L.Document.open(fixture(self.BASIC)).paragraphs()
        self.assertEqual([p.text() for p in paragraphs], [self.BASIC_TEXT])
        runs = paragraphs[0].runs()
        self.assertEqual(
            [(r.text(), r.bold(), r.italic(), r.underline()) for r in runs],
            [(self.BASIC_TEXT, None, None, None)],
        )
        self.assertEqual(
            paragraph_texts(L.Document.open(fixture(self.UNICODE))),
            ["Pages borrowed body 😀\nFirst section marker.", "Second section marker — end."],
        )

    def test_tables_are_not_exposed(self):
        # The fixture holds two native tables (its receipt lists "Alpha",
        # "Beta", "Café", "北京" cells). The Python Document exposes no table,
        # and text() carries one U+FFFC object placeholder per table.
        document = L.Document.open(fixture(self.TABLES))
        self.assertEqual(document.tables(), [])
        self.assertEqual(
            document.text(), "Pages hidden-axis native oracle — 2026-09-05\n￼\n￼"
        )
        self.assertEqual(L.Document.open(fixture(self.BASIC)).tables(), [])


# ---------------------------------------------------------------------------
# Presentation family
# ---------------------------------------------------------------------------


class PptxReadTests(unittest.TestCase):
    SAMPLE = "ooxml/pptx/sample.pptx"
    SHAPES = "ooxml/pptx/shapes.pptx"
    SAMPLE_SLIDES = [
        "Lorem ipsum dolor sit amet\nNunc at risus vel erat tempus posuere. Aenean non ante.",
        "Lorem ipsum dolor sit amet\nLorem\nipsum\ndolor\nsit\namet",
    ]

    def test_open_accepts_str_pathlib_and_keyword(self):
        path = fixture(self.SAMPLE)
        for presentation in (
            L.Presentation.open(path),
            L.Presentation.open(Path(path)),
            L.Presentation.open(path=path),
        ):
            self.assertIsInstance(presentation, L.Presentation)
            self.assertEqual(presentation.slide_count(), 2)

    def test_slides(self):
        presentation = L.Presentation.open(fixture(self.SAMPLE))
        self.assertEqual(presentation.slide_count(), 2)
        self.assertEqual(slide_texts(presentation), self.SAMPLE_SLIDES)

    def test_text_joins_slides_with_a_blank_line(self):
        presentation = L.Presentation.open(fixture(self.SAMPLE))
        self.assertEqual(presentation.text(), "\n\n".join(self.SAMPLE_SLIDES))

    def test_shape_text_includes_tables_and_skips_empty_slides(self):
        presentation = L.Presentation.open(fixture(self.SHAPES))
        texts = slide_texts(presentation)
        self.assertEqual(presentation.slide_count(), 6)
        self.assertEqual(len(texts), 6)
        self.assertEqual(texts[0].split("\n")[:2], ["Learning PPTX", "Cloud"])
        self.assertEqual(texts[1], "PPTX \nTitle\nSubtitle\nAnd second line")
        self.assertEqual((texts[2], texts[5]), ("", ""))
        self.assertIn("B3 and C3 are merged", texts[3].split("\n"))  # table cells
        self.assertIn("http://poi.apache.org/", texts[4].split("\n"))
        self.assertEqual(presentation.text(), "\n\n".join(t for t in texts if t))


class PptReadTests(unittest.TestCase):
    """Legacy PowerPoint, accepted since S8a enabled the ``ppt`` feature (B1)."""

    SHAPES = "ole/ppt/text_shapes.ppt"
    SAMPLE = "ole/ppt/SampleShow.ppt"

    def test_slides_and_text(self):
        presentation = L.Presentation.open(fixture(self.SHAPES))
        self.assertEqual(presentation.slide_count(), 1)
        # The title placeholder's text lives in the slide list (record 0260).
        self.assertEqual(
            slide_texts(presentation),
            ["Title Placeholder\nText in a TextBox\nRectangle\nEllipse\nOctagon"],
        )
        self.assertEqual(
            presentation.text(), "Title Placeholder\nText in a TextBox\nRectangle\nEllipse\nOctagon"
        )

    def test_text_joins_slides_with_a_blank_line(self):
        presentation = L.Presentation.open(fixture(self.SAMPLE))
        texts = slide_texts(presentation)
        self.assertEqual(presentation.slide_count(), 2)
        self.assertEqual(len(texts), 2)
        self.assertTrue(texts[0].startswith("FirstSlide\nTitle of the first slide\n"), texts[0])
        self.assertTrue(texts[1].startswith("This is the second slide\n"), texts[1])
        self.assertEqual(presentation.text(), "\n\n".join(texts))


class OdpReadTests(unittest.TestCase):
    LINKS = "odf/odp/tdf169979.odp"
    LAYOUTS = "odf/corpus/impress-master-layouts.odp"
    TABLE = "odf/odp/tdf105502.odp"

    def test_slides_and_text(self):
        presentation = L.Presentation.open(fixture(self.LINKS))
        expected = ["Link to BMP File", "Link to PDF File", "Link to Text"]
        self.assertEqual(presentation.slide_count(), 3)
        self.assertEqual(slide_texts(presentation), expected)
        self.assertEqual(presentation.text(), "\n\n".join(expected))

    def test_many_slides(self):
        presentation = L.Presentation.open(fixture(self.LAYOUTS))
        self.assertEqual(presentation.slide_count(), 5)
        self.assertEqual(slide_texts(presentation), ["1", "2", "3", "4", "5"])

    def test_table_shape_text(self):
        presentation = L.Presentation.open(fixture(self.TABLE))
        self.assertEqual(presentation.slide_count(), 1)
        self.assertEqual(slide_texts(presentation), ["A1\nB1\nA2\nB2"])


class KeynoteReadTests(unittest.TestCase):
    """Keynote through ``Presentation`` (record 0003 R6)."""

    BASIC = "iwork/keynote/basic.key"
    TABLE = "iwork/keynote/slide-table-read-native.key"
    TEXT = "Litchi native Keynote fixture\nBuffa lazy-view migration verification\n2026-08-07"

    def test_slides_and_text(self):
        presentation = L.Presentation.open(fixture(self.BASIC))
        self.assertEqual(presentation.slide_count(), 1)
        self.assertEqual(slide_texts(presentation), [self.TEXT])
        self.assertEqual(presentation.text(), self.TEXT)

    def test_table_cells_are_not_in_slide_text(self):
        # The slide holds a native table whose receipt lists A1 "Buffa
        # discovery" and A2 "Café 北京"; slide text leaves table cells out.
        presentation = L.Presentation.open(fixture(self.TABLE))
        self.assertEqual(presentation.slide_count(), 1)
        self.assertEqual(slide_texts(presentation), [self.TEXT])
        self.assertNotIn("Buffa discovery", presentation.text())


# ---------------------------------------------------------------------------
# Workbook family
# ---------------------------------------------------------------------------


class XlsxReadTests(unittest.TestCase):
    NORMAL = "ooxml/xlsx/SimpleNormal.xlsx"
    INLINE = "ooxml/xlsx/InlineStrings.xlsx"

    def test_open_accepts_str_pathlib_and_keyword(self):
        path = fixture(self.NORMAL)
        for workbook in (
            L.Workbook.open(path),
            L.Workbook.open(Path(path)),
            L.Workbook.open(path=path),
        ):
            self.assertIsInstance(workbook, L.Workbook)
            self.assertEqual(workbook.worksheet_names(), ["Sheet1", "Sheet Number 2"])

    def test_worksheet_names_and_count(self):
        workbook = L.Workbook.open(fixture(self.NORMAL))
        self.assertEqual(workbook.worksheet_count(), 2)
        self.assertEqual(workbook.worksheet_names(), ["Sheet1", "Sheet Number 2"])

    def test_text_rows_are_tab_separated_cell_values(self):
        lines = text_lines(L.Workbook.open(fixture(self.INLINE)))
        for row in (
            "Numbers\tStrings\tInline Strings\tFormulas",
            "12\tA\t1st Inline String\t12",
            "32\tB\t2nd Inline String\t44",  # formula cell: cached value
            "21\tLonger Text\tThe End\t9",
        ):
            self.assertIn(row, lines)

    def test_text_covers_every_sheet(self):
        lines = text_lines(L.Workbook.open(fixture(self.NORMAL)))
        self.assertIn("test\t\t1", lines)
        self.assertIn("test 2\t\t2", lines)
        self.assertTrue(any(line.startswith("This is sheet 2") for line in lines), lines)


class XlsbReadTests(unittest.TestCase):
    SAMPLE = "ooxml/xlsb/sample.xlsb"
    VALUES = "ooxml/xlsb/bug66682.xlsb"
    SIMPLE = "ooxml/xlsb/Simple.xlsb"

    def test_worksheet_names_and_count(self):
        workbook = L.Workbook.open(fixture(self.SAMPLE))
        self.assertEqual(workbook.worksheet_count(), 3)
        self.assertEqual(workbook.worksheet_names(), ["Sheet1", "rich test", "Sheet3"])

    def test_text_rows(self):
        lines = text_lines(L.Workbook.open(fixture(self.SAMPLE)))
        for row in ("Lorem\t111", "amet\t555", "Nunc\t999", "at\t4995"):
            self.assertIn(row, lines)
        self.assertTrue(
            any(line.startswith("The quick brown fox jumps over the lazy dog") for line in lines)
        )

    def test_boolean_and_error_values(self):
        workbook = L.Workbook.open(fixture(self.VALUES))
        self.assertEqual(workbook.worksheet_names(), ["test2"])
        lines = text_lines(workbook)
        for row in (
            "ID\tLiteral\tFornula",
            "1\tTRUE\tTRUE",
            "2\tFALSE\tFALSE",
            "4\t1.5\t1.5",
            "5\tabcd\tabcd",
            "6\tERROR\t#DIV/0!",
            "9\tERROR\t#N/A",
        ):
            self.assertIn(row, lines)

    def test_single_text_cell(self):
        workbook = L.Workbook.open(fixture(self.SIMPLE))
        self.assertEqual(workbook.worksheet_names(), ["Sheet1", "Sheet2", "Sheet3"])
        self.assertIn(
            "This is an example spreadsheet created with Microsoft Excel 2007 Beta 2.",
            text_lines(workbook),
        )


class XlsReadTests(unittest.TestCase):
    """Legacy Excel, accepted since S8a enabled the ``xls`` feature (B1)."""

    SIMPLE = "ole/xls/Simple.xls"
    MULTI = "ole/xls/SimpleMultiCell.xls"

    def test_worksheet_names_and_count(self):
        workbook = L.Workbook.open(fixture(self.SIMPLE))
        self.assertEqual(workbook.worksheet_count(), 3)
        self.assertEqual(workbook.worksheet_names(), ["Sheet1", "Sheet2", "Sheet3"])
        self.assertEqual(text_lines(workbook)[0], "replaceMe")

    def test_text_rows_are_tab_separated_cell_values(self):
        lines = text_lines(L.Workbook.open(fixture(self.MULTI)))
        for row in ("1\t\t\t\t", "2\t3\t\t\t", "7\t8\t9\t10\t", "11\t12\t13\t14\t15"):
            self.assertIn(row, lines)


class OdsReadTests(unittest.TestCase):
    TWO_SHEETS = "odf/corpus/calc-two-sheets.ods"
    CHINESE = "odf/corpus/calc-unicode-chinese.ods"
    FORMULAS = "odf/corpus/calc-formulas.ods"

    def test_worksheet_names_and_rows(self):
        workbook = L.Workbook.open(fixture(self.TWO_SHEETS))
        self.assertEqual(workbook.worksheet_count(), 2)
        self.assertEqual(workbook.worksheet_names(), ["sheet1", "sheet2"])
        lines = text_lines(workbook)
        for row in ("col A\tcol B\tcol C", "1\t2\t3", "a text\t\tanother"):
            self.assertIn(row, lines)

    def test_unicode_sheet_name(self):
        workbook = L.Workbook.open(fixture(self.CHINESE))
        self.assertEqual(workbook.worksheet_names(), ["工作表1"])
        self.assertEqual(workbook.text(), "test\ttest\ttest")

    def test_formula_values(self):
        workbook = L.Workbook.open(fixture(self.FORMULAS))
        self.assertEqual(workbook.worksheet_names(), ["Sheet1", "Sheet2", "Sheet3"])
        self.assertEqual(text_lines(workbook)[:4], ["3", "4", "5", "7"])


class NumbersReadTests(unittest.TestCase):
    """Numbers through ``Workbook`` (record 0003 R6).

    Unlike the other workbook formats, text() names each sheet and table and
    then lists one cell per line, with formula source instead of a value.
    """

    BASIC = "iwork/numbers/basic.numbers"
    TWO_SHEETS = "iwork/numbers/table-relocation-native.numbers"
    VALUES = "iwork/numbers/cell-value-native.numbers"

    def test_worksheet_names_and_text(self):
        workbook = L.Workbook.open(fixture(self.BASIC))
        self.assertEqual(workbook.worksheet_count(), 1)
        self.assertEqual(workbook.worksheet_names(), ["Sheet 1"])
        self.assertEqual(workbook.text(), "Sheet 1\nTable 1\nLitchi native Numbers fixture\n42")

    def test_multiple_sheets(self):
        workbook = L.Workbook.open(fixture(self.TWO_SHEETS))
        self.assertEqual(workbook.worksheet_names(), ["Sheet 1", "Sheet 2"])
        self.assertEqual(
            workbook.text(),
            "Sheet 1\nTable 1\nNative image adjustment marker\n"
            "Sheet 2\nDestination table\nDestination stays",
        )

    def test_cell_values_one_per_line(self):
        workbook = L.Workbook.open(fixture(self.VALUES))
        self.assertEqual(workbook.worksheet_names(), ["Sheet 1"])
        self.assertEqual(
            text_lines(workbook),
            [
                "Sheet 1", "shared-model", "Kind", "Value", "Note",
                "Number", "42.5", 'Text Café, "北京"',
                "Boolean", "true", "Text line one", "line two",
                "Date", "2026-09-10", "Native date",
                "Duration", "1h 2m 3s", "Native duration",
                "Formula", "=(B2+1)", "Cached number",
                "Error", "=(1/0)", "Cached error",
                "Empty", "Blank value",
            ],
        )


class IworkPackageDirectoryTests(unittest.TestCase):
    """Apple's package-directory form of the three iWork formats.

    Since S8a, ``litchi::open`` names the iWork format of a package directory
    and the semantic readers accept one, so the openers read it like the
    packaged file. Before S8a every opener refused a directory as an I/O
    error. Path detection still reports None for a directory
    (test_formats_and_values).
    """

    ROOT = TEST_DATA / "iwork" / "directory"

    def test_package_directories_open(self):
        for path in (
            self.ROOT / "pages" / "basic.pages",
            self.ROOT / "keynote" / "basic.key",
            self.ROOT / "numbers" / "basic.numbers",
        ):
            self.assertTrue(path.is_dir(), f"fixture directory missing: {path}")
        document = L.Document.open(self.ROOT / "pages" / "basic.pages")
        self.assertEqual(document.text(), PagesReadTests.BASIC_TEXT)
        presentation = L.Presentation.open(self.ROOT / "keynote" / "basic.key")
        self.assertEqual(slide_texts(presentation), [KeynoteReadTests.TEXT])
        workbook = L.Workbook.open(self.ROOT / "numbers" / "basic.numbers")
        self.assertEqual(workbook.worksheet_names(), ["Sheet 1"])
        self.assertEqual(workbook.text(), "Sheet 1\nTable 1\nLitchi native Numbers fixture\n42")

    def test_wrong_opener_names_the_directory_format(self):
        with self.assertRaises(L.UnsupportedError) as caught:
            L.Workbook.open(self.ROOT / "pages" / "basic.pages")
        self.assertIn("pages", str(caught.exception).lower())


# ---------------------------------------------------------------------------
# Cross-format behaviour
# ---------------------------------------------------------------------------


class SignedPackageReadTests(unittest.TestCase):
    """Signed OOXML reads normally today (record 0003 R7: S8a wraps it in ``Signed``)."""

    def test_signed_docx(self):
        document = L.Document.open(fixture("poi/test-data/xmldsign/hello-world-signed.docx"))
        self.assertEqual(document.text().strip(), "Hello world")
        self.assertEqual(paragraph_texts(document), ["Hello world", ""])

    def test_signed_pptx(self):
        presentation = L.Presentation.open(fixture("poi/test-data/xmldsign/hello-world-signed.pptx"))
        self.assertEqual(presentation.slide_count(), 1)
        self.assertEqual(presentation.text(), "Hello\nWorld")

    def test_signed_xlsx(self):
        workbook = L.Workbook.open(fixture("poi/test-data/xmldsign/hello-world-signed.xlsx"))
        self.assertEqual(workbook.worksheet_names(), ["Sheet1", "Sheet2", "Sheet3"])
        self.assertEqual(workbook.text(), "Hello\t123\n")


class ContentBasedDispatchTests(unittest.TestCase):
    """Openers classify by content; the file name's extension does not matter."""

    def test_extension_is_ignored(self):
        with tempfile.TemporaryDirectory() as directory:
            docx_as_xlsx = copy_fixture("ooxml/docx/Hyperlink.docx", directory, "doc.xlsx")
            docx_no_ext = copy_fixture("ooxml/docx/Hyperlink.docx", directory, "noextension")
            xlsx_as_docx = copy_fixture("ooxml/xlsx/SimpleNormal.xlsx", directory, "book.docx")
            doc_as_xls = copy_fixture("ole/doc/NoHeadFoot.doc", directory, "legacy.xls")
            ods_as_odt = copy_fixture("odf/corpus/calc-two-sheets.ods", directory, "calc.odt")
            pages_as_zip = copy_fixture("iwork/pages/basic.pages", directory, "basic.zip")
            for path in (docx_as_xlsx, docx_no_ext):
                self.assertEqual(
                    paragraph_texts(L.Document.open(path)),
                    ["Following is a hyperlink.", "EricWhite.com"],
                )
            self.assertEqual(
                L.Workbook.open(xlsx_as_docx).worksheet_names(), ["Sheet1", "Sheet Number 2"]
            )
            self.assertEqual(paragraph_texts(L.Document.open(doc_as_xls))[2], "HEADING TEXT")
            self.assertEqual(L.Workbook.open(ods_as_odt).worksheet_names(), ["sheet1", "sheet2"])
            self.assertEqual(
                L.Document.open(pages_as_zip).text(), PagesReadTests.BASIC_TEXT
            )


class HandleSemanticsTests(unittest.TestCase):
    """Handle-level guarantees: list returns, lifetimes, idempotence, snapshots."""

    def test_collections_are_lists_of_typed_handles(self):
        document = L.Document.open(fixture(DocxReadTests.TABLES))
        paragraphs = document.paragraphs()
        tables = document.tables()
        self.assertIs(type(paragraphs), list)
        self.assertIs(type(tables), list)
        self.assertTrue(all(isinstance(p, L.Paragraph) for p in paragraphs))
        self.assertTrue(all(isinstance(r, L.Run) for p in paragraphs for r in p.runs()))
        self.assertIs(type(paragraphs[0].runs()), list)
        self.assertTrue(all(isinstance(t, L.Table) for t in tables))
        rows = tables[0].rows()
        self.assertIs(type(rows), list)
        self.assertTrue(all(isinstance(r, L.TableRow) for r in rows))
        cells = rows[0].cells()
        self.assertIs(type(cells), list)
        self.assertTrue(all(isinstance(c, L.TableCell) for c in cells))
        slides = L.Presentation.open(fixture(PptxReadTests.SAMPLE)).slides()
        self.assertIs(type(slides), list)
        self.assertTrue(all(isinstance(s, L.Slide) for s in slides))
        names = L.Workbook.open(fixture(XlsxReadTests.NORMAL)).worksheet_names()
        self.assertIs(type(names), list)
        self.assertTrue(all(type(n) is str for n in names))

    def test_counts_agree_with_collections(self):
        for relative in (
            DocxReadTests.TABLES, DocReadTests.TABLE, OdtReadTests.SPANS, RtfReadTests.TABLE
        ):
            for table in L.Document.open(fixture(relative)).tables():
                self.assertEqual(table.row_count(), len(table.rows()), relative)
        for relative in (
            PptxReadTests.SHAPES, OdpReadTests.LAYOUTS, KeynoteReadTests.BASIC
        ):
            presentation = L.Presentation.open(fixture(relative))
            self.assertEqual(presentation.slide_count(), len(presentation.slides()), relative)
        for relative in (
            XlsxReadTests.NORMAL, XlsbReadTests.SAMPLE, OdsReadTests.TWO_SHEETS,
            NumbersReadTests.BASIC,
        ):
            workbook = L.Workbook.open(fixture(relative))
            self.assertEqual(
                workbook.worksheet_count(), len(workbook.worksheet_names()), relative
            )

    def test_children_outlive_their_parent_handle(self):
        paragraphs = L.Document.open(fixture(DocxReadTests.TABLES)).paragraphs()
        tables = L.Document.open(fixture(DocxReadTests.TABLES)).tables()
        slides = L.Presentation.open(fixture(PptxReadTests.SAMPLE)).slides()
        gc.collect()
        self.assertEqual(paragraphs[4].runs()[1].text(), "ectet")
        self.assertEqual(table_grid(tables[5]), DocxReadTests.GRID)
        self.assertEqual([s.text() for s in slides], PptxReadTests.SAMPLE_SLIDES)

    def test_reads_are_repeatable(self):
        document = L.Document.open(fixture(DocReadTests.UNICODE))
        self.assertEqual(document.text(), document.text())
        self.assertEqual(paragraph_texts(document), paragraph_texts(document))
        presentation = L.Presentation.open(fixture(PptxReadTests.SHAPES))
        self.assertEqual(slide_texts(presentation), slide_texts(presentation))
        workbook = L.Workbook.open(fixture(XlsbReadTests.SAMPLE))
        self.assertEqual(workbook.text(), workbook.text())

    def test_handles_keep_reading_what_they_opened(self):
        # Source-backed handles pin the opened file: replacing the path with a
        # new file, or unlinking it, does not change what an open handle reads.
        cases = (
            (L.Document, "ooxml/docx/Hyperlink.docx", "ooxml/docx/comment.docx", "text"),
            (L.Presentation, "ooxml/pptx/sample.pptx", "ooxml/pptx/trigger.pptx", "text"),
            (L.Workbook, "ooxml/xlsx/SimpleNormal.xlsx", "ooxml/xlsx/InlineStrings.xlsx", "text"),
            (L.Workbook, "ooxml/xlsb/Simple.xlsb", "ooxml/xlsb/sample.xlsb", "text"),
            (L.Document, "odf/corpus/writer-table.odt", "odf/corpus/writer-paragraph-styles.odt", "text"),
            (L.Workbook, "odf/corpus/calc-two-sheets.ods", "odf/corpus/calc-formulas.ods", "text"),
            (L.Presentation, "odf/odp/tdf169979.odp", "odf/odp/cellspan.odp", "text"),
            (L.Document, "ole/doc/NoHeadFoot.doc", "ole/doc/Lists.doc", "text"),
            (L.Document, "rtf/testStyles.rtf", "rtf/hyperlink.rtf", "text"),
            (L.Document, "iwork/pages/basic.pages", "iwork/pages/body-sections-unicode.pages", "text"),
            (L.Workbook, "iwork/numbers/basic.numbers", "ooxml/xlsx/SimpleNormal.xlsx", "text"),
        )
        for opener, original, replacement, method in cases:
            with self.subTest(original=original), tempfile.TemporaryDirectory() as directory:
                expected = getattr(opener.open(fixture(original)), method)()
                replaced = copy_fixture(original, directory, "replaced" + Path(original).suffix)
                unlinked = copy_fixture(original, directory, "unlinked" + Path(original).suffix)
                replaced_handle = opener.open(replaced)
                unlinked_handle = opener.open(unlinked)
                replace_atomically(replaced, fixture(replacement))
                os.remove(unlinked)
                self.assertEqual(getattr(replaced_handle, method)(), expected)
                self.assertEqual(getattr(unlinked_handle, method)(), expected)


class ReprSnapshotTests(unittest.TestCase):
    """``repr``/``str`` of handles (snapshot tier: may change only deliberately).

    Previews are cut by character (bug B2 fixed in S8a); multibyte text is
    covered by test_multibyte_previews_are_cut_by_character.
    """

    def test_handle_reprs(self):
        self.assertEqual(repr(L.Document.open(fixture(DocxReadTests.HYPERLINK))), "<Document>")
        self.assertEqual(
            repr(L.Presentation.open(fixture(PptxReadTests.SAMPLE))), "<Presentation: 2 slides>"
        )
        self.assertEqual(
            repr(L.Workbook.open(fixture(XlsxReadTests.NORMAL))), "<Workbook: 2 worksheets>"
        )
        table = L.Document.open(fixture(DocxReadTests.TABLES)).tables()[0]
        self.assertEqual(repr(table), "<Table: 3 rows>")
        self.assertEqual(repr(table.rows()[0]), "<TableRow: 3 cells>")
        self.assertEqual(repr(table.rows()[0].cells()[0]), "<TableCell: 'Loren'>")

    def test_text_previews(self):
        paragraphs = L.Document.open(fixture(DocxReadTests.HYPERLINK)).paragraphs()
        self.assertEqual(repr(paragraphs[0]), "<Paragraph: 'Following is a hyperlink.'>")
        self.assertEqual(repr(paragraphs[0].runs()[0]), "<Run: 'Following is a hyperlink.'>")
        self.assertEqual(
            repr(L.Presentation.open(fixture(OdpReadTests.LINKS)).slides()[0]),
            "<Slide: 'Link to BMP File'>",
        )

    def test_previews_truncate_long_text(self):
        # Paragraph and Slide previews keep 50 characters, Run and TableCell 30.
        paragraph = L.Document.open(fixture(DocReadTests.NO_HEAD_FOOT)).paragraphs()[0]
        text = paragraph.text()
        self.assertEqual(repr(paragraph), f"<Paragraph: '{text[:50]}...'>")
        self.assertEqual(repr(paragraph.runs()[0]), f"<Run: '{text[:30]}...'>")
        slide = L.Presentation.open(fixture(PptxReadTests.SAMPLE)).slides()[0]
        self.assertEqual(repr(slide), f"<Slide: '{slide.text()[:50]}...'>")

    def test_multibyte_previews_are_cut_by_character(self):
        # drawing.docx holds Russian text, where a cut at byte 50 or 30 used to
        # fall inside a letter and panic (bug B2).
        document = L.Document.open(fixture("ooxml/docx/drawing.docx"))
        paragraph = document.paragraphs()[3]
        run = document.paragraphs()[7].runs()[2]
        cell = document.tables()[1].rows()[2].cells()[0]
        for handle, owner, limit in ((paragraph, "Paragraph", 50), (run, "Run", 30), (cell, "TableCell", 30)):
            text = handle.text()
            preview = text if len(text) <= limit else text[:limit] + "..."
            self.assertEqual(repr(handle), f"<{owner}: '{preview}'>")


if __name__ == "__main__":
    unittest.main()
