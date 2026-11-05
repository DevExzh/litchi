"""The ``LitchiError`` exception hierarchy (S8a).

Every error the bindings raise for a file is a ``LitchiError`` carrying the
unified error's ``code``, ``message`` and ``suggestion``. Each family class
also derives from the built-in exception the oracle contract pins, so
``except OSError`` and ``except ValueError`` keep working. The class names are
recorded in docs/unified-ops/s8a-python-on-file.md for the editing API (M3).
"""

from __future__ import annotations

import copy
import os
import pickle
import tempfile
import unittest

from _support import L, copy_fixture, fixture, modify_in_place, read_fixture

FAMILIES = {
    "PathError": LookupError,
    "PropertyError": ValueError,
    "UnsupportedError": ValueError,
    "IoError": OSError,
    "CorruptFileError": ValueError,
    "ProtectedError": ValueError,
}

DEFAULT_CODES = {
    "LitchiError": "internal_error",
    "PathError": "path_not_found",
    "PropertyError": "invalid_value",
    "UnsupportedError": "unsupported_capability",
    "IoError": "io_error",
    "CorruptFileError": "corrupt_file",
    "ProtectedError": "document_protected",
}


def raised(opener, path):
    try:
        opener.open(path)
    except Exception as error:  # noqa: BLE001 - the class is what is checked
        return error
    raise AssertionError(f"{opener.__name__}.open({path!r}) did not raise")


class HierarchyTests(unittest.TestCase):
    def test_base_class(self):
        self.assertTrue(issubclass(L.LitchiError, Exception))
        self.assertEqual(L.LitchiError.__module__, "litchi_py")
        for name in ("code", "message", "suggestion"):
            self.assertIsInstance(getattr(L.LitchiError, name), property)

    def test_family_classes(self):
        for name, builtin in FAMILIES.items():
            with self.subTest(cls=name):
                cls = getattr(L, name)
                self.assertTrue(issubclass(cls, L.LitchiError))
                self.assertTrue(issubclass(cls, builtin))
                self.assertEqual(cls.__module__, "litchi_py")

    def test_families_do_not_overlap_the_other_builtins(self):
        self.assertFalse(issubclass(L.IoError, ValueError))
        self.assertFalse(issubclass(L.LitchiError, (OSError, ValueError, LookupError)))
        self.assertFalse(issubclass(L.PathError, ValueError))
        for name in ("PropertyError", "UnsupportedError", "CorruptFileError", "ProtectedError"):
            self.assertFalse(issubclass(getattr(L, name), (OSError, LookupError)), name)

    def test_construction(self):
        for name, code in DEFAULT_CODES.items():
            with self.subTest(cls=name):
                error = getattr(L, name)("went wrong")
                self.assertEqual(error.code, code)
                self.assertEqual(error.message, "went wrong")
                self.assertIsNone(error.suggestion)
                self.assertEqual(str(error), "went wrong")
                self.assertEqual(error.args, ("went wrong",))
        error = L.IoError("cannot read", "file_not_found", "check the path")
        self.assertEqual(
            (error.code, error.message, error.suggestion),
            ("file_not_found", "cannot read", "check the path"),
        )
        self.assertEqual(str(error), "cannot read (suggestion: check the path)")
        self.assertIsNone(error.errno)
        keyword = L.UnsupportedError(message="m", code="unsupported_format", suggestion="s")
        self.assertEqual((keyword.code, keyword.suggestion), ("unsupported_format", "s"))

    def test_attributes_are_read_only(self):
        error = L.LitchiError("m", "internal_error")
        for name in ("code", "message", "suggestion"):
            with self.assertRaises(AttributeError):
                setattr(error, name, "x")

    def test_pickle_and_copy_round_trip(self):
        for error in (
            L.IoError("cannot read", "file_not_found", "check the path"),
            L.UnsupportedError("detected Docx", "unsupported_format", None),
            raised(L.Presentation, fixture("ooxml/docx/Hyperlink.docx")),
        ):
            for clone in (pickle.loads(pickle.dumps(error)), copy.copy(error)):
                with self.subTest(error=repr(error)):
                    self.assertIs(type(clone), type(error))
                    self.assertEqual(
                        (clone.code, clone.message, clone.suggestion, str(clone)),
                        (error.code, error.message, error.suggestion, str(error)),
                    )


class RaisedErrorTests(unittest.TestCase):
    """The class, code and suggestion of each kind of open and read failure."""

    def assertError(self, error, cls, code, suggestion=None):
        self.assertIs(type(error), cls, repr(error))
        self.assertEqual(error.code, code)
        self.assertEqual(error.suggestion, suggestion)
        expected = error.message
        if suggestion is not None:
            expected += f" (suggestion: {suggestion})"
        self.assertEqual(str(error), expected)

    def test_unreadable_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            missing = os.path.join(directory, "missing.docx")
            for opener in (L.Document, L.Presentation, L.Workbook):
                with self.subTest(opener=opener.__name__):
                    self.assertError(raised(opener, missing), L.IoError, "file_not_found")
                    self.assertError(raised(opener, directory), L.IoError, "io_error")

    def test_inputs_that_are_not_office_files(self):
        with tempfile.TemporaryDirectory() as directory:
            plain = os.path.join(directory, "notes.txt")
            with open(plain, "wb") as handle:
                handle.write(b"plain text, not an office file\n" * 8)
            truncated = os.path.join(directory, "truncated.docx")
            docx = read_fixture("ooxml/docx/Hyperlink.docx")
            with open(truncated, "wb") as handle:
                handle.write(docx[: len(docx) // 2])
            for opener in (L.Document, L.Presentation, L.Workbook):
                with self.subTest(opener=opener.__name__):
                    error = raised(opener, plain)
                    self.assertError(error, L.UnsupportedError, "unsupported_format")
                    self.assertEqual(error.message, "not a recognized Office file")
                    self.assertError(raised(opener, truncated), L.CorruptFileError, "corrupt_file")

    def test_wrong_opener_names_the_right_one(self):
        cases = (
            (L.Presentation, "ooxml/docx/Hyperlink.docx", "Docx", "Document.open"),
            (L.Workbook, "iwork/pages/basic.pages", "Pages", "Document.open"),
            (L.Document, "ole/ppt/SampleShow.ppt", "Ppt", "Presentation.open"),
            (L.Workbook, "iwork/keynote/basic.key", "Keynote", "Presentation.open"),
            (L.Document, "ole/xls/Simple.xls", "Xls", "Workbook.open"),
            (L.Presentation, "iwork/numbers/basic.numbers", "Numbers", "Workbook.open"),
        )
        for opener, relative, detected, right in cases:
            with self.subTest(opener=opener.__name__, fixture=relative):
                error = raised(opener, fixture(relative))
                self.assertError(
                    error, L.UnsupportedError, "unsupported_format", f"open it with {right}"
                )
                self.assertEqual(
                    error.message, f"detected {detected}, which this operation does not handle"
                )

    def test_formats_no_opener_reads_get_no_suggestion(self):
        # Odg is detected, but no family opener reads OpenDocument drawings
        # (litchi_py.open holds them as File, record 0282).
        for opener in (L.Document, L.Presentation, L.Workbook):
            with self.subTest(opener=opener.__name__):
                error = raised(opener, fixture("libreoffice-core/desktop/qa/data/BlankDrawDocument.odg"))
                self.assertError(error, L.UnsupportedError, "unsupported_format")
                self.assertIsNone(error.suggestion)

    def test_unsupported_variants(self):
        self.assertError(
            raised(L.Document, fixture("ole/doc/word6-no-table-stream.doc")),
            L.UnsupportedError,
            "unsupported_format",
        )
        for opener in (L.Document, L.Presentation, L.Workbook):
            with self.subTest(opener=opener.__name__):
                self.assertError(
                    raised(opener, fixture("odf/flat/minimal-formula.fodf")),
                    L.UnsupportedError,
                    "unsupported_capability",
                )

    def test_encrypted_document(self):
        error = raised(L.Document, fixture("ole/doc/PasswordProtected.doc"))
        self.assertIs(type(error), L.ProtectedError)
        self.assertEqual(error.code, "document_protected")
        self.assertEqual(error.message, "the Doc document is encrypted")
        self.assertIn("password", error.suggestion)
        # The encrypted .doc is still a Doc: another opener names it.
        self.assertError(
            raised(L.Workbook, fixture("ole/doc/PasswordProtected.doc")),
            L.UnsupportedError,
            "unsupported_format",
            "open it with Document.open",
        )
        # Legacy encryption of the other formats opens as File::Locked too,
        # among them PowerPoint's RC4 CryptoAPI, which a build without the
        # `encryption` feature cannot decrypt (M0a S9; before, it was an
        # UnsupportedError).
        for opener, relative, name in (
            (L.Workbook, "ole/xls/password.xls", "Xls"),
            (L.Workbook, "poi/test-data/spreadsheet/xor-encryption-abc.xls", "Xls"),
            (L.Document, "poi/test-data/document/password_password_cryptoapi.doc", "Doc"),
            (L.Presentation, "poi/test-data/slideshow/Password_Protected-hello.ppt", "Ppt"),
        ):
            with self.subTest(fixture=relative):
                error = raised(opener, fixture(relative))
                self.assertIs(type(error), L.ProtectedError)
                self.assertEqual(error.message, f"the {name} document is encrypted")

    def test_source_change_is_one_io_error_for_every_format(self):
        cases = (
            (L.Document, "ooxml/docx/Hyperlink.docx"),
            (L.Presentation, "ooxml/pptx/sample.pptx"),
            (L.Workbook, "ooxml/xlsx/SimpleNormal.xlsx"),
            (L.Workbook, "ooxml/xlsb/Simple.xlsb"),
            (L.Document, "odf/corpus/writer-table.odt"),
            (L.Workbook, "odf/corpus/calc-two-sheets.ods"),
            (L.Presentation, "odf/odp/tdf169979.odp"),
        )
        for opener, relative in cases:
            with self.subTest(fixture=relative), tempfile.TemporaryDirectory() as directory:
                path = copy_fixture(relative, directory)
                handle = opener.open(path)
                modify_in_place(path)
                with self.assertRaises(L.IoError) as caught:
                    handle.text()
                self.assertError(caught.exception, L.IoError, "io_error")
                self.assertIn("source changed", caught.exception.message)

    def test_except_clauses(self):
        with tempfile.TemporaryDirectory() as directory:
            missing = os.path.join(directory, "missing.xlsx")
            for cls in (OSError, L.IoError, L.LitchiError, Exception):
                with self.assertRaises(cls):
                    L.Workbook.open(missing)
        wrong = fixture("ooxml/pptx/sample.pptx")
        for cls in (ValueError, L.UnsupportedError, L.LitchiError):
            with self.assertRaises(cls):
                L.Workbook.open(wrong)
        try:
            L.Workbook.open(wrong)
        except OSError:  # pragma: no cover - must not match
            self.fail("a format error must not be an OSError")
        except ValueError:
            pass


if __name__ == "__main__":
    unittest.main()
