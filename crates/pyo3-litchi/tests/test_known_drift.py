"""Known drift and bugs in the read API, and their fixes.

The oracle (docs/unified-ops/s8a-python-oracle.md) recorded each item as a
pair of tests:

  test_<id>_..._today       asserted what the bindings did then;
  test_<id>_..._intended    asserted the fix, marked @unittest.expectedFailure.

S8a (docs/unified-ops/s8a-python-on-file.md) fixed every item: each ``_today``
test is deleted and each ``_intended`` test is now an ordinary test. A new
item is added the same way, as a pair.

ID   Item                                                      Fixed in
D1   Odb detected as FileFormat.Doc; FileFormat lacks Odb      S8a
D2   Stub declares 14 of 21 formats                            S8a
D3   Stub declares FileFormat(Enum); runtime is no Enum        S8a (stub base dropped)
D4   Raised classes contradict the stub's "Raises:" sections   S8a (LitchiError families)
D5   README documents a Workbook/Worksheet API that is absent  S8a (README fixed)
F10  PPTX source change raises ValueError, others Exception    S8a (IoError for all)
B1   .ppt and .xls are not recognised (features not enabled)   S8a
B2   repr() panics when its preview splits a UTF-8 character   S8a
B3   DOCX Document.text() drops paragraph boundaries           S8a (File::text keeps them)
B4   Run.underline() is a placeholder that returns None        S8a
B5   Handles panic when used from another thread               S8a
"""

from __future__ import annotations

import os
import tempfile
import unittest

from _support import (
    CORE_FORMAT_NAMES,
    L,
    copy_fixture,
    fixture,
    modify_in_place,
    paragraph_texts,
    parse_stub,
    public_members,
    read_fixture,
    readme_api_references,
    resolve_stub_base,
    run_isolated,
    runtime_has,
)

FF = L.FileFormat
OPENERS = (L.Document, L.Presentation, L.Workbook)
STUB = parse_stub()


class D1OdbDetectedAsDoc(unittest.TestCase):
    """``From<litchi::FileFormat>`` (src/common.rs) sent unnamed variants to
    Doc. S8a added ``FileFormat.Odb``; an unknown future format is an error."""

    ODB = (
        "libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb",
        "libreoffice-core/extras/source/database/biblio.odb",
    )

    def test_d1_odb_is_reported_as_odb_intended(self):
        self.assertEqual(public_members(FF), set(CORE_FORMAT_NAMES))
        for relative in self.ODB:
            self.assertEqual(L.detect_file_format(fixture(relative)), FF.Odb)
            self.assertEqual(L.detect_file_format_from_bytes(read_fixture(relative)), FF.Odb)


class D2StubListsFourteenFormats(unittest.TestCase):
    """The stub declared 14 of the 21 formats; S8a declares all of them."""

    def test_d2_stub_declares_every_runtime_format_intended(self):
        self.assertEqual(set(STUB.classes["FileFormat"].members), public_members(FF))


class D3StubDeclaresEnum(unittest.TestCase):
    """The stub said ``class FileFormat(Enum)``; the runtime is a plain pyo3
    class. S8a dropped the base from the stub (the runtime is unchanged, so
    ``str(v) == name`` still holds, see test_formats_and_values)."""

    def test_d3_stub_bases_hold_at_runtime_intended(self):
        for base in STUB.classes["FileFormat"].bases:
            self.assertTrue(issubclass(FF, resolve_stub_base(base)))


class D4RaisedClassesContradictStub(unittest.TestCase):
    """The stub's ``Raises:`` promised OSError for unreadable files and
    ValueError for invalid or unsupported formats, but Workbook raised a bare
    Exception for everything and the other openers for most format errors.
    S8a maps every error code to a LitchiError family whose built-in base keeps
    those promises (IoError is an OSError; UnsupportedError, CorruptFileError
    and ProtectedError are ValueErrors)."""

    def _raised(self, opener, path):
        with self.assertRaises(Exception) as caught:
            opener.open(path)
        return caught.exception

    def _format_error_inputs(self, directory):
        plain = os.path.join(directory, "plain.txt")
        with open(plain, "wb") as handle:
            handle.write(b"not an office file\n" * 8)
        truncated = os.path.join(directory, "truncated.docx")
        docx = read_fixture("ooxml/docx/Hyperlink.docx")
        with open(truncated, "wb") as handle:
            handle.write(docx[: len(docx) // 2])
        return [
            (opener, path)
            for opener in OPENERS
            for path in (plain, truncated)
        ] + [
            (L.Document, fixture("ooxml/xlsx/SimpleNormal.xlsx")),
            (L.Presentation, fixture("ooxml/docx/Hyperlink.docx")),
            (L.Workbook, fixture("ooxml/pptx/sample.pptx")),
        ]

    def test_d4a_workbook_io_errors_are_oserror_intended(self):
        with tempfile.TemporaryDirectory() as directory:
            for path in (os.path.join(directory, "missing.xlsx"), directory):
                self.assertIsInstance(self._raised(L.Workbook, path), OSError)

    def test_d4b_format_errors_are_valueerror_intended(self):
        with tempfile.TemporaryDirectory() as directory:
            for opener, path in self._format_error_inputs(directory):
                self.assertIsInstance(self._raised(opener, path), ValueError)

    def test_d4c_workbook_unsupported_input_is_valueerror_intended(self):
        self.assertIsInstance(
            self._raised(L.Workbook, fixture("odf/flat/minimal-formula.fodf")), ValueError
        )


class D5ReadmeDocumentsMissingApi(unittest.TestCase):
    """README "API Reference" listed Workbook.worksheets(), worksheet_by_name()
    and Worksheet.name/row_count/column_count/cell_value/rows, none of which
    exist. S8a removed them from the README."""

    def _missing(self):
        return {ref for ref in readme_api_references() if not runtime_has(*ref)}

    def test_d5_readme_lists_only_existing_api_intended(self):
        self.assertEqual(self._missing(), set())


class F10PptxSourceChangeFlattened(unittest.TestCase):
    """Record 0003 F10: PPTX ``check_source()`` failures went through
    ``map_ooxml_error`` and surfaced as ValueError, other formats as a bare
    Exception. Through ``litchi::File`` every format reports ``io_error``,
    raised as IoError."""

    def _source_change_error(self, opener, relative):
        with tempfile.TemporaryDirectory() as directory:
            path = copy_fixture(relative, directory)
            handle = opener.open(path)
            modify_in_place(path)
            with self.assertRaises(Exception) as caught:
                handle.text()
            return caught.exception

    def test_f10_pptx_source_change_matches_other_formats_intended(self):
        pptx = self._source_change_error(L.Presentation, "ooxml/pptx/sample.pptx")
        docx = self._source_change_error(L.Document, "ooxml/docx/Hyperlink.docx")
        self.assertIs(type(pptx), type(docx))


class B1LegacyPptAndXlsNotEnabled(unittest.TestCase):
    """crates/pyo3-litchi/Cargo.toml enabled litchi's ``doc`` feature but not
    ``ppt`` or ``xls``, and OLE2 classification is gated per feature
    (crates/litchi-ops/src/detect/ole2.rs). S8a enables ``legacy``."""

    PPT = ("ole/ppt/text_shapes.ppt", "ole/ppt/SampleShow.ppt")
    XLS = ("ole/xls/Simple.xls", "ole/xls/SimpleMultiCell.xls")

    def test_b1_ppt_opens_as_presentation_intended(self):
        for relative in self.PPT:
            self.assertEqual(L.detect_file_format(fixture(relative)), FF.Ppt)
            self.assertGreaterEqual(L.Presentation.open(fixture(relative)).slide_count(), 1)

    def test_b1_xls_opens_as_workbook_intended(self):
        for relative in self.XLS:
            self.assertEqual(L.detect_file_format(fixture(relative)), FF.Xls)
            self.assertGreaterEqual(L.Workbook.open(fixture(relative)).worksheet_count(), 1)


class B2ReprPanicsOnMultibyteText(unittest.TestCase):
    """``__repr__`` previews sliced ``&text[..50]`` (Paragraph, Slide) or
    ``&text[..30]`` (Run, TableCell) by byte, which panicked when the cut fell
    inside a character (an interpreter abort in release builds). S8a cuts at
    50 or 30 characters. Still run in a subprocess, so a regression cannot
    abort the test run."""

    DRAWING = "ooxml/docx/drawing.docx"  # Russian text
    CASES = {
        "Paragraph": "document.paragraphs()[3]",  # byte 50 falls inside a letter
        "Run": "document.paragraphs()[7].runs()[2]",  # byte 30 falls inside a letter
        "TableCell": "document.tables()[1].rows()[2].cells()[0]",
    }

    def _repr_in_subprocess(self, expression):
        code = (
            "import litchi_py\n"
            f"document = litchi_py.Document.open({fixture(self.DRAWING)!r})\n"
            f"print(repr({expression}))\n"
        )
        return run_isolated(code)

    def test_b2_repr_previews_multibyte_text_intended(self):
        for owner, expression in self.CASES.items():
            result = self._repr_in_subprocess(expression)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue(result.stdout.startswith(f"<{owner}: '"), result.stdout)


class B3DocxTextDropsParagraphBoundaries(unittest.TestCase):
    """The DOCX text used ``extract_word_text``, which concatenates ``<w:t>``
    text. S8a makes ``File::text()`` (and so ``Document.text()``) join the
    paragraph texts with \\n, as the other Document formats separate
    paragraphs (DOC ends each with \\r, RTF/ODT/Pages use \\n)."""

    HYPERLINK = "ooxml/docx/Hyperlink.docx"

    def test_b3_docx_text_keeps_paragraph_boundaries_intended(self):
        document = L.Document.open(fixture(self.HYPERLINK))
        lines = document.text().splitlines()
        for text in paragraph_texts(document):
            if text:
                self.assertIn(text, lines)


class B4UnderlineIsAPlaceholder(unittest.TestCase):
    """``Run.underline`` in src/document.rs returned ``None`` unconditionally.
    S8a reads it through ``litchi::compat::word::Run::underline``."""

    UNDERLINED = "libreoffice-core/sw/qa/extras/rtfimport/data/tdf108947.rtf"

    def _underlined_run(self):
        # Paragraph 0 is "{\ul ... Sublet Agreement}" in the RTF source.
        run = L.Document.open(fixture(self.UNDERLINED)).paragraphs()[0].runs()[0]
        self.assertEqual(run.text(), "Sublet Agreement")
        return run

    def test_b4_underline_reports_underlined_text_intended(self):
        self.assertIs(self._underlined_run().underline(), True)


class B5HandlesPanicAcrossThreads(unittest.TestCase):
    """Every handle class was ``#[pyclass(unsendable)]``, so pyo3 panicked when
    one was used from a thread other than its creator. S8a makes the classes
    ``frozen`` and thread-safe over the ``Send + Sync`` ``litchi::File``.
    Subprocess for the same reason as B2."""

    CODE = (
        "import sys, threading, litchi_py\n"
        "document = litchi_py.Document.open({path!r})\n"
        "result = {{}}\n"
        "def read():\n"
        "    try:\n"
        "        result['text'] = document.text()\n"
        "    except BaseException as error:\n"
        "        result['error'] = error\n"
        "worker = threading.Thread(target=read)\n"
        "worker.start()\n"
        "worker.join()\n"
        "if 'error' in result:\n"
        "    print(type(result['error']).__name__, result['error'], file=sys.stderr)\n"
        "    sys.exit(3)\n"
        "print(result['text'])\n"
    )

    def _cross_thread_read(self):
        return run_isolated(self.CODE.format(path=fixture("ooxml/docx/Hyperlink.docx")))

    def test_b5_cross_thread_read_succeeds_intended(self):
        result = self._cross_thread_read()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("Following is a hyperlink.", result.stdout)


if __name__ == "__main__":
    unittest.main()
