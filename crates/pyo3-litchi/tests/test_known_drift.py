"""Known drift and bugs in the read API: today's behaviour and the intended fix.

Each item is a pair of tests:

  test_<id>_..._today       asserts what the bindings do now (passes now);
  test_<id>_..._intended    asserts the fix, marked @unittest.expectedFailure.

When a change fixes an item, its ``_today`` test fails and its ``_intended``
test becomes an unexpected success, which fails the run (unittest since
Python 3.4, and pytest). Then delete the ``_today`` test, drop the decorator,
and update docs/unified-ops/s8a-python-oracle.md.

ID   Item                                                      Fix expected in
D1   Odb detected as FileFormat.Doc; FileFormat lacks Odb      S8a (record 0003 §5.1)
D2   Stub declares 14 of 21 formats                            S8a (record 0003 §5.1)
D3   Stub declares FileFormat(Enum); runtime is no Enum        open
D4   Raised classes contradict the stub's "Raises:" sections   open (S8a remaps errors)
D5   README documents a Workbook/Worksheet API that is absent  open
F10  PPTX source change raises ValueError, others Exception    S8a (record 0003 F10, R7)
B1   .ppt and .xls are not recognised (features not enabled)   open
B2   repr() panics when its preview splits a UTF-8 character   open
B3   DOCX Document.text() drops paragraph boundaries           open
B4   Run.underline() is a placeholder that returns None        open
B5   Handles panic when used from another thread               open
"""

from __future__ import annotations

import enum
import os
import tempfile
import unittest

from _support import (
    CORE_FORMAT_NAMES,
    KNOWN_README_ONLY_API,
    L,
    STUB_FORMAT_NAMES_TODAY,
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
    """``From<litchi::FileFormat>`` (src/common.rs) sends unnamed variants to Doc."""

    ODB = (
        "libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb",
        "libreoffice-core/extras/source/database/biblio.odb",
    )

    def test_d1_odb_is_reported_as_doc_today(self):
        self.assertFalse(hasattr(FF, "Odb"))
        for relative in self.ODB:
            with self.subTest(fixture=relative):
                self.assertEqual(L.detect_file_format(fixture(relative)), FF.Doc)
                self.assertEqual(L.detect_file_format_from_bytes(read_fixture(relative)), FF.Doc)
                # ...yet the Document opener refuses it.
                with self.assertRaises(Exception):
                    L.Document.open(fixture(relative))

    @unittest.expectedFailure
    def test_d1_odb_is_reported_as_odb_intended(self):
        self.assertEqual(public_members(FF), set(CORE_FORMAT_NAMES))
        for relative in self.ODB:
            self.assertEqual(L.detect_file_format(fixture(relative)), FF.Odb)
            self.assertEqual(L.detect_file_format_from_bytes(read_fixture(relative)), FF.Odb)


class D2StubListsFourteenFormats(unittest.TestCase):
    def test_d2_stub_declares_14_of_21_formats_today(self):
        declared = set(STUB.classes["FileFormat"].members)
        self.assertEqual(declared, set(STUB_FORMAT_NAMES_TODAY))
        self.assertEqual(
            public_members(FF) - declared, {"Odg", "Odc", "Odf", "Odi", "Odm", "Oth"}
        )
        self.assertEqual(set(CORE_FORMAT_NAMES) - declared, {"Odg", "Odc", "Odf", "Odi", "Odm", "Oth", "Odb"})

    @unittest.expectedFailure
    def test_d2_stub_declares_every_runtime_format_intended(self):
        self.assertEqual(set(STUB.classes["FileFormat"].members), public_members(FF))


class D3StubDeclaresEnum(unittest.TestCase):
    """The stub says ``class FileFormat(Enum)``; the runtime is a plain pyo3 class.

    Either fix works: make the runtime an Enum (keeping ``str(v) == name``,
    see test_formats_and_values) or drop the base from the stub.
    """

    def test_d3_runtime_file_format_is_not_an_enum_today(self):
        self.assertEqual(STUB.classes["FileFormat"].bases, ("Enum",))
        self.assertFalse(issubclass(FF, enum.Enum))

    @unittest.expectedFailure
    def test_d3_stub_bases_hold_at_runtime_intended(self):
        for base in STUB.classes["FileFormat"].bases:
            self.assertTrue(issubclass(FF, resolve_stub_base(base)))


class D4RaisedClassesContradictStub(unittest.TestCase):
    """Stub ``Raises:``: IOError if the file cannot be read, ValueError if the
    format is invalid or unsupported. Workbook errors all pass through
    ``boxed_err_to_py_err`` (bare Exception); Document and Presentation map only
    Io, InvalidFormat and Unsupported (``to_py_err``)."""

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

    def test_d4a_workbook_io_errors_are_bare_exception_today(self):
        with tempfile.TemporaryDirectory() as directory:
            for path in (os.path.join(directory, "missing.xlsx"), directory):
                self.assertIs(type(self._raised(L.Workbook, path)), Exception)

    @unittest.expectedFailure
    def test_d4a_workbook_io_errors_are_oserror_intended(self):
        with tempfile.TemporaryDirectory() as directory:
            for path in (os.path.join(directory, "missing.xlsx"), directory):
                self.assertIsInstance(self._raised(L.Workbook, path), OSError)

    def test_d4b_format_errors_are_bare_exception_today(self):
        with tempfile.TemporaryDirectory() as directory:
            for opener, path in self._format_error_inputs(directory):
                with self.subTest(opener=opener.__name__, input=os.path.basename(path)):
                    self.assertIs(type(self._raised(opener, path)), Exception)

    @unittest.expectedFailure
    def test_d4b_format_errors_are_valueerror_intended(self):
        with tempfile.TemporaryDirectory() as directory:
            for opener, path in self._format_error_inputs(directory):
                self.assertIsInstance(self._raised(opener, path), ValueError)

    def test_d4c_workbook_unsupported_input_is_bare_exception_today(self):
        flat = fixture("odf/ods/grid-padding.fods")
        self.assertIs(type(self._raised(L.Workbook, flat)), Exception)
        # The same refusal is a ValueError through Document.
        self.assertIs(type(self._raised(L.Document, flat)), ValueError)

    @unittest.expectedFailure
    def test_d4c_workbook_unsupported_input_is_valueerror_intended(self):
        self.assertIsInstance(
            self._raised(L.Workbook, fixture("odf/ods/grid-padding.fods")), ValueError
        )


class D5ReadmeDocumentsMissingApi(unittest.TestCase):
    """README "API Reference" lists Workbook.worksheets(), worksheet_by_name()
    and Worksheet.name/row_count/column_count/cell_value/rows; none exist."""

    def _missing(self):
        return {ref for ref in readme_api_references() if not runtime_has(*ref)}

    def test_d5_readme_lists_unimplemented_api_today(self):
        self.assertEqual(self._missing(), set(KNOWN_README_ONLY_API))

    @unittest.expectedFailure
    def test_d5_readme_lists_only_existing_api_intended(self):
        self.assertEqual(self._missing(), set())


class F10PptxSourceChangeFlattened(unittest.TestCase):
    """Record 0003 F10: PPTX ``check_source()`` failures go through
    ``map_ooxml_error`` and surface as InvalidFormat, i.e. ValueError."""

    def _source_change_error(self, opener, relative):
        with tempfile.TemporaryDirectory() as directory:
            path = copy_fixture(relative, directory)
            handle = opener.open(path)
            modify_in_place(path)
            with self.assertRaises(Exception) as caught:
                handle.text()
            return caught.exception

    def test_f10_pptx_source_change_is_valueerror_today(self):
        pptx = self._source_change_error(L.Presentation, "ooxml/pptx/sample.pptx")
        docx = self._source_change_error(L.Document, "ooxml/docx/Hyperlink.docx")
        odp = self._source_change_error(L.Presentation, "odf/odp/tdf169979.odp")
        self.assertIs(type(pptx), ValueError)
        self.assertIs(type(docx), Exception)
        self.assertIs(type(odp), Exception)

    @unittest.expectedFailure
    def test_f10_pptx_source_change_matches_other_formats_intended(self):
        pptx = self._source_change_error(L.Presentation, "ooxml/pptx/sample.pptx")
        docx = self._source_change_error(L.Document, "ooxml/docx/Hyperlink.docx")
        self.assertIs(type(pptx), type(docx))


class B1LegacyPptAndXlsNotEnabled(unittest.TestCase):
    """crates/pyo3-litchi/Cargo.toml enables litchi's ``doc`` feature but not
    ``ppt`` or ``xls``, and OLE2 classification is gated per feature
    (crates/litchi-ops/src/detect/ole2.rs). README, module docs and the
    Presentation/Workbook docstrings claim .ppt and .xls support."""

    PPT = ("ole/ppt/text_shapes.ppt", "ole/ppt/SampleShow.ppt")
    XLS = ("ole/xls/Simple.xls", "ole/xls/SimpleMultiCell.xls")

    def test_b1_ppt_and_xls_are_not_recognised_today(self):
        for relative in self.PPT + self.XLS:
            with self.subTest(fixture=relative):
                path = fixture(relative)
                self.assertIsNone(L.detect_file_format(path))
                self.assertIsNone(L.detect_file_format_from_bytes(read_fixture(relative)))
                for opener in OPENERS:
                    with self.assertRaises(Exception) as caught:
                        opener.open(path)
                    self.assertEqual(str(caught.exception), "Not a valid Office file")

    @unittest.expectedFailure
    def test_b1_ppt_opens_as_presentation_intended(self):
        for relative in self.PPT:
            self.assertEqual(L.detect_file_format(fixture(relative)), FF.Ppt)
            self.assertGreaterEqual(L.Presentation.open(fixture(relative)).slide_count(), 1)

    @unittest.expectedFailure
    def test_b1_xls_opens_as_workbook_intended(self):
        for relative in self.XLS:
            self.assertEqual(L.detect_file_format(fixture(relative)), FF.Xls)
            self.assertGreaterEqual(L.Workbook.open(fixture(relative)).worksheet_count(), 1)


class B2ReprPanicsOnMultibyteText(unittest.TestCase):
    """``__repr__`` previews slice ``&text[..50]`` (Paragraph, Slide) or
    ``&text[..30]`` (Run, TableCell) by byte. When the cut falls inside a
    character, Rust panics: a dev build raises pyo3_runtime.PanicException (a
    BaseException that ``except Exception`` misses), a release build aborts the
    interpreter. Run in a subprocess for that reason. Slide shares the code
    (src/presentation.rs) but no fixture reaches it."""

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

    def test_b2_repr_panics_today(self):
        for owner, expression in self.CASES.items():
            with self.subTest(owner=owner):
                result = self._repr_in_subprocess(expression)
                self.assertNotEqual(result.returncode, 0, result.stdout)
                self.assertIn("is not a char boundary", result.stderr)

    @unittest.expectedFailure
    def test_b2_repr_previews_multibyte_text_intended(self):
        for owner, expression in self.CASES.items():
            result = self._repr_in_subprocess(expression)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue(result.stdout.startswith(f"<{owner}: '"), result.stdout)


class B3DocxTextDropsParagraphBoundaries(unittest.TestCase):
    """The source-backed DOCX route uses ``extract_word_text``, which
    concatenates ``<w:t>`` text. Every other Document format separates
    paragraphs (DOC ends each with \\r, RTF/ODT/Pages use \\n)."""

    HYPERLINK = "ooxml/docx/Hyperlink.docx"

    def test_b3_docx_text_concatenates_paragraphs_today(self):
        document = L.Document.open(fixture(self.HYPERLINK))
        self.assertEqual(document.text(), "Following is a hyperlink.EricWhite.com")
        tables = L.Document.open(fixture("ooxml/docx/table-alignment.docx"))
        self.assertEqual(tables.text(), "".join(paragraph_texts(tables)))

    @unittest.expectedFailure
    def test_b3_docx_text_keeps_paragraph_boundaries_intended(self):
        document = L.Document.open(fixture(self.HYPERLINK))
        lines = document.text().splitlines()
        for text in paragraph_texts(document):
            if text:
                self.assertIn(text, lines)


class B4UnderlineIsAPlaceholder(unittest.TestCase):
    """``Run.underline`` in src/document.rs returns ``Ok(None)`` unconditionally;
    the facade Run has no underline accessor to delegate to."""

    UNDERLINED = "libreoffice-core/sw/qa/extras/rtfimport/data/tdf108947.rtf"

    def _underlined_run(self):
        # Paragraph 0 is "{\ul ... Sublet Agreement}" in the RTF source.
        run = L.Document.open(fixture(self.UNDERLINED)).paragraphs()[0].runs()[0]
        self.assertEqual(run.text(), "Sublet Agreement")
        return run

    def test_b4_underline_returns_none_today(self):
        run = self._underlined_run()
        self.assertIsNone(run.underline())
        # The same run answers bold and italic explicitly.
        self.assertIs(run.bold(), False)
        self.assertIs(run.italic(), False)

    @unittest.expectedFailure
    def test_b4_underline_reports_underlined_text_intended(self):
        self.assertIs(self._underlined_run().underline(), True)


class B5HandlesPanicAcrossThreads(unittest.TestCase):
    """Every handle class is ``#[pyclass(unsendable)]``; pyo3 panics when one is
    used from a thread other than its creator (PanicException in dev builds,
    interpreter abort in release builds). Subprocess for the same reason as B2."""

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

    def test_b5_cross_thread_use_panics_today(self):
        result = self._cross_thread_read()
        self.assertNotEqual(result.returncode, 0, result.stdout)
        self.assertIn("is unsendable, but sent to another thread", result.stderr)

    @unittest.expectedFailure
    def test_b5_cross_thread_read_succeeds_intended(self):
        result = self._cross_thread_read()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("Following is a hyperlink.", result.stdout)


if __name__ == "__main__":
    unittest.main()
