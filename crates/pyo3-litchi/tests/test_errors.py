"""Error behaviour of the openers and of reads after the source changes.

Contract tier (keep green through S8a):
  * exception classes are checked with ``isinstance``, so S8a may raise
    subclasses (for example a ``LitchiError`` family, design §8) but must keep
    ``except OSError`` / ``except ValueError`` / ``except Exception`` working
    where they work today;
  * messages are checked only for semantic fragments: the detected format of
    a wrong-family input (record 0003 D7), "password", "source changed".

Snapshot tier: ``MessageSnapshotTests`` pins the exact messages. S8a changed
them deliberately (design §8: errors carry code/message/suggestion, and
``str()`` is the message followed by the suggestion); the S8a record
(docs/unified-ops/s8a-python-on-file.md) lists the differences. Change the
table only on purpose, and list the difference in a record.

The exception classes, codes and suggestions are pinned in test_exceptions.py;
the fixed class drift (D4, F10) in test_known_drift.py.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
import zipfile

from _support import L, copy_fixture, fixture, modify_in_place, read_fixture

OPENERS = (L.Document, L.Presentation, L.Workbook)

#: One fixture per accepted format: (fixture, detected name, accepting opener).
ACCEPTED = (
    ("ooxml/docx/Hyperlink.docx", "Docx", L.Document),
    ("ole/doc/NoHeadFoot.doc", "Doc", L.Document),
    ("rtf/testStyles.rtf", "Rtf", L.Document),
    ("odf/corpus/writer-table.odt", "Odt", L.Document),
    ("libreoffice-core/sw/qa/extras/odfexport/data/tdf121119.odm", "Odm", L.Document),
    ("iwork/pages/basic.pages", "Pages", L.Document),
    ("ooxml/pptx/sample.pptx", "Pptx", L.Presentation),
    ("odf/odp/tdf169979.odp", "Odp", L.Presentation),
    ("iwork/keynote/basic.key", "Keynote", L.Presentation),
    ("ooxml/xlsx/SimpleNormal.xlsx", "Xlsx", L.Workbook),
    ("ooxml/xlsb/sample.xlsb", "Xlsb", L.Workbook),
    ("odf/corpus/calc-two-sheets.ods", "Ods", L.Workbook),
    ("iwork/numbers/basic.numbers", "Numbers", L.Workbook),
)

#: Flat ODF the unified opener refuses: a flat formula (no flat reader) and
#: a flat ODP whose drawing body the leaf refuses. Flat ODS (record 0277) and
#: flat ODT (record 0292) open.
FLAT_ODF = ("odf/flat/minimal-formula.fodf", "odf/drawing/multicolor-gradient.fodp")
FLAT_ODS = "odf/ods/grid-padding.fods"
FLAT_ODT = "odf/odt/form-controls.fodt"

#: ODF object formats no family opener reads: ``litchi_py.open`` holds them
#: as ``File`` (record 0282), but they are neither documents,
#: presentations nor workbooks.
DETECTED_NOT_OPENED = (
    "libreoffice-core/desktop/qa/data/BlankDrawDocument.odg",
    "odf/odc-producer-evidence/odfdom-created.odc",
    "odf/native-resave/source/font-styles.odf",
    "libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb",
)


def write_non_office_inputs(directory: str) -> dict:
    """Create inputs that are not Office files; return {label: path}."""
    paths = {}

    def put(name: str, payload: bytes) -> None:
        path = os.path.join(directory, name)
        with open(path, "wb") as handle:
            handle.write(payload)
        paths[name] = path

    put("plain.txt", b"hello world, not an office file\n" * 10)
    put("random.bin", bytes(range(256)) * 4)
    put("empty.docx", b"")
    put("plain.docx", b"hello world, not an office file\n" * 10)
    docx = read_fixture("ooxml/docx/Hyperlink.docx")
    put("truncated.docx", docx[: len(docx) // 2])
    archive = os.path.join(directory, "unrelated.zip")
    with zipfile.ZipFile(archive, "w") as handle:
        handle.writestr("readme.txt", "an ordinary ZIP archive")
    paths["unrelated.zip"] = archive
    return paths


class PathErrorTests(unittest.TestCase):
    def test_missing_file(self):
        with tempfile.TemporaryDirectory() as directory:
            missing = os.path.join(directory, "missing.docx")
            for opener in (L.Document, L.Presentation):
                with self.subTest(opener=opener.__name__):
                    with self.assertRaises(OSError):
                        opener.open(missing)
            # Workbook raised a bare Exception before S8a (drift D4a, now an
            # OSError; see test_known_drift).
            with self.assertRaises(Exception):
                L.Workbook.open(missing)

    def test_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            for opener in (L.Document, L.Presentation):
                with self.subTest(opener=opener.__name__):
                    with self.assertRaises(OSError):
                        opener.open(directory)
            with self.assertRaises(Exception):
                L.Workbook.open(directory)

    def test_argument_types(self):
        for opener in OPENERS:
            for argument in (None, 42, ["a.docx"]):
                with self.subTest(opener=opener.__name__, argument=argument):
                    with self.assertRaises(TypeError):
                        opener.open(argument)


class NonOfficeInputTests(unittest.TestCase):
    def test_non_office_input_is_refused_as_a_format_error(self):
        with tempfile.TemporaryDirectory() as directory:
            for label, path in write_non_office_inputs(directory).items():
                for opener in OPENERS:
                    with self.subTest(input=label, opener=opener.__name__):
                        with self.assertRaises(Exception) as caught:
                            opener.open(path)
                        # A readable file in the wrong format is not an I/O error.
                        self.assertNotIsInstance(caught.exception, OSError)
                        self.assertTrue(str(caught.exception))


class WrongFamilyTests(unittest.TestCase):
    def test_wrong_family_refusal_names_the_detected_format(self):
        for relative, detected, accepting in ACCEPTED:
            path = fixture(relative)
            for opener in OPENERS:
                if opener is accepting:
                    continue
                with self.subTest(fixture=relative, opener=opener.__name__):
                    with self.assertRaises(Exception) as caught:
                        opener.open(path)
                    self.assertNotIsInstance(caught.exception, OSError)
                    self.assertIn(detected.lower(), str(caught.exception).lower())


class UnsupportedInputTests(unittest.TestCase):
    def test_flat_odf_is_refused_as_unsupported(self):
        for relative in FLAT_ODF:
            path = fixture(relative)
            for opener in (L.Document, L.Presentation):
                with self.subTest(fixture=relative, opener=opener.__name__):
                    with self.assertRaises(ValueError):
                        opener.open(path)
            with self.subTest(fixture=relative, opener="Workbook"):
                # A bare Exception before S8a (drift D4c, now a ValueError;
                # see test_known_drift).
                with self.assertRaises(Exception):
                    L.Workbook.open(path)

    def test_flat_odt_opens_as_a_document(self):
        # Record 0292: flat text documents read through the ODT adapter.
        document = L.Document.open(fixture(FLAT_ODT))
        self.assertIsInstance(document.text(), str)
        for opener in (L.Presentation, L.Workbook):
            with self.subTest(opener=opener.__name__):
                with self.assertRaises(ValueError):
                    opener.open(fixture(FLAT_ODT))

    def test_flat_ods_opens_as_a_workbook(self):
        names = L.Workbook.open(fixture(FLAT_ODS)).worksheet_names()
        self.assertEqual(names[:3], ["Padded", "Gaps", "Styled tail"])
        for opener in (L.Document, L.Presentation):
            with self.subTest(opener=opener.__name__):
                with self.assertRaises(ValueError):
                    opener.open(fixture(FLAT_ODS))

    def test_detected_odf_variants_without_an_opener_are_refused(self):
        for relative in DETECTED_NOT_OPENED:
            path = fixture(relative)
            for opener in OPENERS:
                with self.subTest(fixture=relative, opener=opener.__name__):
                    with self.assertRaises(Exception) as caught:
                        opener.open(path)
                    self.assertNotIsInstance(caught.exception, OSError)

    def test_password_protected_doc(self):
        with self.assertRaises(ValueError) as caught:
            L.Document.open(fixture("ole/doc/PasswordProtected.doc"))
        self.assertIn("password", str(caught.exception).lower())

    def test_unsupported_word_version(self):
        with self.assertRaises(ValueError):
            L.Document.open(fixture("ole/doc/word6-no-table-stream.doc"))


class SourceChangeTests(unittest.TestCase):
    """Reads after the opened file changes in place (same inode)."""

    SOURCE_BACKED = (
        (L.Document, "ooxml/docx/Hyperlink.docx", ("text", "paragraphs", "tables")),
        (L.Presentation, "ooxml/pptx/sample.pptx", ("text", "slides", "slide_count")),
        (L.Workbook, "ooxml/xlsx/SimpleNormal.xlsx", ("text", "worksheet_names", "worksheet_count")),
        (L.Workbook, "ooxml/xlsb/Simple.xlsb", ("text", "worksheet_names")),
        (L.Document, "odf/corpus/writer-table.odt", ("text", "paragraphs")),
        (L.Workbook, "odf/corpus/calc-two-sheets.ods", ("text",)),
        (L.Presentation, "odf/odp/tdf169979.odp", ("text", "slides")),
    )

    #: Read completely at open today, so a later change is not observed.
    EAGER = (
        (L.Document, "ole/doc/NoHeadFoot.doc"),
        (L.Document, "rtf/testStyles.rtf"),
        (L.Document, "iwork/pages/basic.pages"),
        (L.Presentation, "iwork/keynote/basic.key"),
        (L.Workbook, "iwork/numbers/basic.numbers"),
    )

    def test_source_backed_reads_refuse_a_changed_source(self):
        for opener, relative, methods in self.SOURCE_BACKED:
            for method in methods:
                with self.subTest(fixture=relative, method=method):
                    with tempfile.TemporaryDirectory() as directory:
                        path = copy_fixture(relative, directory)
                        handle = opener.open(path)
                        modify_in_place(path)
                        with self.assertRaises(Exception) as caught:
                            getattr(handle, method)()
                        self.assertIn("source changed", str(caught.exception).lower())

    def test_eagerly_read_formats_ignore_a_changed_source(self):
        # Snapshot of today's routing: if S8a makes one of these source-backed,
        # it must refuse like SOURCE_BACKED; move the row there deliberately.
        for opener, relative in self.EAGER:
            with self.subTest(fixture=relative):
                with tempfile.TemporaryDirectory() as directory:
                    path = copy_fixture(relative, directory)
                    handle = opener.open(path)
                    expected = handle.text()
                    modify_in_place(path)
                    self.assertEqual(handle.text(), expected)


@unittest.skipUnless(sys.platform.startswith("linux"), "I/O messages are platform text")
class MessageSnapshotTests(unittest.TestCase):
    """Exact messages (snapshot tier), as S8a left them."""

    NOT_OFFICE = "not a recognized Office file"
    NO_ZIP_MAGIC = "ZIP error: OOXML-suffixed input does not have ZIP magic"
    FLAT = "flat OpenDocument {} is detected, but only packaged OpenDocument opens"
    DETECTED = "detected {}, which this operation does not handle"
    OPENER = {L.Document: "Document.open", L.Presentation: "Presentation.open", L.Workbook: "Workbook.open"}

    def _message(self, opener, path) -> str:
        with self.assertRaises(Exception) as caught:
            opener.open(path)
        return str(caught.exception)

    def test_messages(self):
        with tempfile.TemporaryDirectory() as directory:
            inputs = write_non_office_inputs(directory)
            rows = [
                (opener, os.path.join(directory, "missing.docx"),
                 "No such file or directory (os error 2)")
                for opener in OPENERS
            ]
            rows += [
                (opener, directory, "positional filesystem source must be a regular file")
                for opener in OPENERS
            ]
            for opener in OPENERS:
                rows += [
                    (opener, inputs["plain.txt"], self.NOT_OFFICE),
                    (opener, inputs["random.bin"], self.NOT_OFFICE),
                    (opener, inputs["unrelated.zip"], self.NOT_OFFICE),
                    (opener, inputs["empty.docx"], self.NO_ZIP_MAGIC),
                    (opener, inputs["plain.docx"], self.NO_ZIP_MAGIC),
                    (opener, inputs["truncated.docx"], "ZIP error: Missing end of central directory"),
                ]
                rows += [
                    (opener, fixture(r), self.DETECTED.format(name))
                    for r, name in zip(DETECTED_NOT_OPENED, ("Odg", "Odc", "Odf", "Odb"))
                ]
            for relative, detected, accepting in ACCEPTED:
                rows += [
                    (
                        opener,
                        fixture(relative),
                        self.DETECTED.format(detected)
                        + f" (suggestion: open it with {self.OPENER[accepting]})",
                    )
                    for opener in OPENERS
                    if opener is not accepting
                ]
            rows += [
                (L.Document, fixture("ole/doc/PasswordProtected.doc"),
                 "the Doc document is encrypted (suggestion: remove the password in the"
                 " application that created the file; litchi_py cannot unlock encrypted"
                 " documents yet)"),
                (L.Document, fixture("ole/doc/word6-no-table-stream.doc"),
                 "Word 6.0 documents (nFib 0x0065) predate the Word 97 binary format and"
                 " are not supported"),
            ]
            for relative, name in zip(FLAT_ODF[:1], ("Odf",)):
                rows += [(opener, fixture(relative), self.FLAT.format(name)) for opener in OPENERS]
            # Flat ODP opens through the leaf's flat reader (record 0278),
            # which refuses this drawing-bodied fixture.
            rows += [
                (opener, fixture(FLAT_ODF[1]), "Invalid format: flat ODP has a misplaced or empty office:body")
                for opener in OPENERS
            ]
            for opener, path, expected in rows:
                with self.subTest(opener=opener.__name__, input=os.path.basename(path)):
                    self.assertEqual(self._message(opener, path), expected)

    def test_source_change_messages(self):
        cases = (
            (L.Document, "ooxml/docx/Hyperlink.docx", "OPC error: OPC source changed from SourceVersion {"),
            (L.Workbook, "ooxml/xlsx/SimpleNormal.xlsx", "OPC package error: OPC source changed from SourceVersion {"),
            (L.Presentation, "ooxml/pptx/sample.pptx", "PPTX OPC error: OPC source changed from SourceVersion {"),
        )
        for opener, relative, prefix in cases:
            with self.subTest(fixture=relative), tempfile.TemporaryDirectory() as directory:
                path = copy_fixture(relative, directory)
                handle = opener.open(path)
                modify_in_place(path)
                with self.assertRaises(Exception) as caught:
                    handle.text()
                self.assertTrue(str(caught.exception).startswith(prefix), str(caught.exception))


if __name__ == "__main__":
    unittest.main()
