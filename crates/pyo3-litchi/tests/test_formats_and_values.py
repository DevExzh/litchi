"""FileFormat values, detection functions, RGBColor and Length.

Contract tier: keep these green. S8a fixed the FileFormat vocabulary drift
(Odb mapped to Doc, the 14-of-21 stub); test_known_drift.py keeps the fixed
pairs' intended halves (D1, D2, D3).
"""

from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

from _support import L, RUNTIME_FORMAT_NAMES, fixture, read_fixture

FF = L.FileFormat

#: (fixture, expected FileFormat name). Covers every format a Python opener
#: accepts, the ODF variants that are detected but not opened, and flat ODF.
DETECTION_CASES = (
    ("ooxml/docx/Hyperlink.docx", "Docx"),
    ("ole/doc/NoHeadFoot.doc", "Doc"),
    ("rtf/testStyles.rtf", "Rtf"),
    ("odf/corpus/writer-table.odt", "Odt"),
    ("iwork/pages/basic.pages", "Pages"),
    ("ooxml/pptx/sample.pptx", "Pptx"),
    ("odf/odp/tdf169979.odp", "Odp"),
    ("iwork/keynote/basic.key", "Keynote"),
    ("ooxml/xlsx/SimpleNormal.xlsx", "Xlsx"),
    ("ooxml/xlsb/sample.xlsb", "Xlsb"),
    ("odf/corpus/calc-two-sheets.ods", "Ods"),
    ("iwork/numbers/basic.numbers", "Numbers"),
    # Legacy PowerPoint and Excel, accepted since S8a enabled them (B1).
    ("ole/ppt/SampleShow.ppt", "Ppt"),
    ("ole/xls/Simple.xls", "Xls"),
    # Detected, but no Python opener reads them (test_errors.py).
    ("libreoffice-core/desktop/qa/data/BlankDrawDocument.odg", "Odg"),
    ("odf/odc-producer-evidence/odfdom-created.odc", "Odc"),
    ("odf/native-resave/source/font-styles.odf", "Odf"),
    ("libreoffice-core/sw/qa/extras/odfexport/data/tdf121119.odm", "Odm"),
    ("libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb", "Odb"),
    # Flat (single-XML) ODF is classified as its packaged sibling.
    ("odf/odt/form-controls.fodt", "Odt"),
    ("odf/ods/grid-padding.fods", "Ods"),
    ("odf/drawing/multicolor-gradient.fodp", "Odp"),
    ("odf/drawing/hatch-angles.fodg", "Odg"),
)


class FileFormatTests(unittest.TestCase):
    def test_runtime_variant_names(self):
        # All 21 core variants (S8a added Odb; drift D1).
        names = {name for name in dir(FF) if not name.startswith("_")}
        self.assertEqual(names, set(RUNTIME_FORMAT_NAMES))

    def test_str_and_repr(self):
        for name in RUNTIME_FORMAT_NAMES:
            value = getattr(FF, name)
            self.assertIsInstance(value, FF)
            self.assertEqual(str(value), name)
            self.assertEqual(repr(value), f"FileFormat.{name}")

    def test_equality_and_hashing(self):
        self.assertEqual(FF.Docx, FF.Docx)
        self.assertNotEqual(FF.Docx, FF.Doc)
        values = [getattr(FF, name) for name in RUNTIME_FORMAT_NAMES]
        self.assertEqual(len(set(values)), len(RUNTIME_FORMAT_NAMES))
        detected = L.detect_file_format(fixture("ooxml/docx/Hyperlink.docx"))
        self.assertEqual(detected, FF.Docx)
        self.assertEqual({FF.Docx: "word"}[detected], "word")
        self.assertIn(detected, {FF.Docx, FF.Xlsx})

    def test_values_convert_to_distinct_ints(self):
        ints = {int(getattr(FF, name)) for name in RUNTIME_FORMAT_NAMES}
        self.assertEqual(len(ints), len(RUNTIME_FORMAT_NAMES))


class DetectionTests(unittest.TestCase):
    def test_detect_from_path(self):
        for relative, expected in DETECTION_CASES:
            with self.subTest(fixture=relative):
                path = fixture(relative)
                self.assertEqual(L.detect_file_format(path), getattr(FF, expected))
                self.assertEqual(L.detect_file_format(Path(path)), getattr(FF, expected))

    def test_detect_from_bytes(self):
        for relative, expected in DETECTION_CASES:
            with self.subTest(fixture=relative):
                self.assertEqual(
                    L.detect_file_format_from_bytes(read_fixture(relative)), getattr(FF, expected)
                )

    def test_keyword_arguments(self):
        path = fixture("ooxml/pptx/sample.pptx")
        self.assertEqual(L.detect_file_format(path=path), FF.Pptx)
        self.assertEqual(L.detect_file_format_from_bytes(data=read_fixture("ooxml/pptx/sample.pptx")), FF.Pptx)

    def test_detection_is_content_based(self):
        with tempfile.TemporaryDirectory() as directory:
            disguised = os.path.join(directory, "spreadsheet.xlsx")
            with open(disguised, "wb") as handle:
                handle.write(read_fixture("ooxml/docx/Hyperlink.docx"))
            self.assertEqual(L.detect_file_format(disguised), FF.Docx)
            text_as_docx = os.path.join(directory, "letter.docx")
            with open(text_as_docx, "w", encoding="utf-8") as handle:
                handle.write("plain text with an Office extension\n")
            self.assertIsNone(L.detect_file_format(text_as_docx))

    def test_unrecognised_input_returns_none(self):
        self.assertIsNone(L.detect_file_format_from_bytes(b""))
        self.assertIsNone(L.detect_file_format_from_bytes(b"PK\x03\x04"))
        self.assertIsNone(L.detect_file_format_from_bytes(b"plain text\n" * 20))
        self.assertIsNone(L.detect_file_format_from_bytes(bytes(range(256)) * 4))
        with tempfile.TemporaryDirectory() as directory:
            # Path detection reports None, it does not raise, for unreadable paths.
            self.assertIsNone(L.detect_file_format(os.path.join(directory, "missing.docx")))
            self.assertIsNone(L.detect_file_format(directory))

    def test_argument_types(self):
        with self.assertRaises(TypeError):
            L.detect_file_format(None)
        with self.assertRaises(TypeError):
            L.detect_file_format(42)
        with self.assertRaises(TypeError):
            L.detect_file_format_from_bytes("PK\x03\x04 is text, not bytes")
        with self.assertRaises(TypeError):
            L.detect_file_format_from_bytes(None)


class RGBColorTests(unittest.TestCase):
    def test_components(self):
        color = L.RGBColor(31, 78, 121)
        self.assertEqual((color.r, color.g, color.b), (31, 78, 121))
        keyword = L.RGBColor(r=0, g=128, b=255)
        self.assertEqual((keyword.r, keyword.g, keyword.b), (0, 128, 255))

    def test_str_and_repr(self):
        color = L.RGBColor(31, 78, 121)
        self.assertEqual(str(color), "RGB(31, 78, 121)")
        self.assertEqual(repr(color), "RGBColor(31, 78, 121)")

    def test_rejects_out_of_range_and_non_integers(self):
        for components in ((256, 0, 0), (0, -1, 0)):
            with self.assertRaises(OverflowError):
                L.RGBColor(*components)
        for components in ((1.5, 0, 0), ("1", 0, 0)):
            with self.assertRaises(TypeError):
                L.RGBColor(*components)


class LengthTests(unittest.TestCase):
    def test_conversions(self):
        for length in (
            L.Length.from_emus(914400),
            L.Length.from_points(72.0),
            L.Length.from_inches(1.0),
            L.Length.from_emus(emus=914400),
            L.Length.from_points(points=72.0),
            L.Length.from_inches(inches=1.0),
        ):
            self.assertEqual(length.to_emus(), 914400)
            self.assertEqual(length.to_points(), 72.0)
            self.assertEqual(length.to_inches(), 1.0)
        self.assertEqual(L.Length.from_points(12).to_emus(), 152400)
        self.assertEqual(L.Length.from_emus(12700).to_points(), 1.0)
        self.assertEqual(L.Length.from_inches(0.5).to_emus(), 457200)
        self.assertEqual(L.Length.from_emus(-914400).to_inches(), -1.0)

    def test_str_and_repr(self):
        self.assertEqual(str(L.Length.from_points(72)), "72.00 pt")
        self.assertEqual(repr(L.Length.from_points(72)), "Length(914400 EMUs)")
        self.assertEqual(str(L.Length.from_points(0.1)), "0.10 pt")

    def test_argument_types(self):
        with self.assertRaises(TypeError):
            L.Length.from_emus(1.5)
        with self.assertRaises(OverflowError):
            L.Length.from_emus(2**63)


if __name__ == "__main__":
    unittest.main()
