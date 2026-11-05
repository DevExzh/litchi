"""The unified write API: ``litchi_py.create``, the ``File`` write verbs
(``set``, ``add``, ``remove``, ``move``, ``swap``), transactions (``edit``,
``Edit``), OfficeCLI batches (``batch``, ``apply``) and publication
(``save``, ``save_as``, ``to_bytes``, ``unsign``) — M3, record 0255.

Oracles, strongest first:

* **Rust equivalence.** Two routes into the same Rust verb must publish the
  same bytes: a dict and OfficeCLI's ``"key=value"`` list, dict items and
  the same items as JSON text, one batch and the same sets one by one.
* **Read-back.** What a write did is read back through the unified read
  API (node views, ``text()``), from the returned File and after
  ``to_bytes``/``open_bytes`` or ``save_as``/``open``.
* **Immutability.** The File a verb is called on publishes the same bytes
  before and after the verb.
* **The capability report.** Every write verb a format's report marks
  unsupported raises ``UnsupportedError`` with the report's code.

Formats: every verb is exercised on every format whose capability report
supports it (DOCX, XLSX, PPTX, XLSB, DOC, XLS, PPT, RTF, ODS, ODT, ODP); no
format refuses every write verb. The cross-thread and GIL tests of the write verbs
are in test_threads.py.
"""

from __future__ import annotations

import json
import os
import pickle
import tempfile
import unittest

from _support import L, fixture, typed_dict_fields

SIGNED_DOCX = "poi/test-data/xmldsign/hello-world-signed.docx"
SIGNED_XLSX = "poi/test-data/xmldsign/hello-world-signed.xlsx"
LOCKED = ("ole/doc/PasswordProtected.doc", "ole/xls/password.xls")
#: The only PPT fixture whose shapes the native PPT editor accepts.
#: A PPT whose third shape holds plain text. (`datetime.ppt` holds only
#: date fields, whose text a set now refuses: record 0304.)
EDITABLE_PPT = "ole/ppt/text_shapes.ppt"
SLIDESHOW_PPT = "ole/ppt/SampleShow.ppt"
NOTES_PPT = "poi/test-data/slideshow/basic_test_ppt_file.ppt"
#: An RTF body of one contiguous root-level span, which the RTF structural
#: edits (remove, move, swap) require.
RTF_INPUT = rb"{\rtf1\ansi First\par Second\par Third\par Fourth}"

#: The formats ``create`` builds.
CREATABLE = (
    "docx", "xlsx", "pptx", "xlsb", "doc", "xls", "ppt", "rtf", "ods", "odt", "odp",
    # The remaining ODF formats (record 0282).
    "oth", "odm", "odg", "odc", "odi", "odb", "odf",
    # Numbers builds a blank workbook (record 0297).
    "numbers",
    # Pages builds its minimal document (record 0299), Keynote its minimal
    # presentation (record 0298).
    "pages", "keynote",
)
#: The formats File holds read-only, and a fixture of each: none since ODT
#: (0276), ODS (0277) and ODP (0278) gained writers.
READ_ONLY: tuple[tuple[str, str], ...] = ()
#: An ODP fixture for the presentation write verbs (record 0278).
ODP_INPUT = "odf/odp/tdf169979.odp"
#: Python suggestion of a document_protected error on a locked file.
PROTECTED_SUGGESTION = (
    "remove the password in the application that created the file; "
    "litchi_py cannot unlock encrypted documents yet"
)
SIGNED_SUGGESTION = (
    "remove the signature explicitly with File.unsign() before editing; "
    "the first write then publishes the document without it"
)


def raised(call):
    """The exception ``call`` raises; fails the test when it returns."""
    try:
        result = call()
    except Exception as error:  # noqa: BLE001 - the class is what is checked
        return error
    raise AssertionError(f"expected an exception, got {result!r}")


def texts(file, path):
    """The texts of the children of ``path``."""
    return [child.get("text") for child in file.get(path)["children"]]


def child_paths(file, path, kind=None):
    return [
        child["path"]
        for child in file.get(path)["children"]
        if kind is None or child["type"] == kind
    ]


def docx_with(*paragraphs):
    file = L.create("docx")
    for text in paragraphs:
        file, _ = file.add("/body", "paragraph", {"text": text})
    return file


def cell_text(file, path):
    return file.get(path, 0)["text"]


class WriteAssertions(unittest.TestCase):
    def assertError(self, error, cls, code, suggestion=...):
        self.assertIs(type(error), cls, repr(error))
        self.assertEqual(error.code, code)
        if suggestion is not ...:
            self.assertEqual(error.suggestion, suggestion)

    def assertUnchanged(self, file, before):
        """``file`` still publishes ``before``."""
        self.assertEqual(file.to_bytes(), before)

    def assertReopens(self, file):
        """``file``'s bytes reopen as the same format with the same text."""
        reopened = L.open_bytes(file.to_bytes())
        self.assertEqual(str(reopened.format), str(file.format))
        self.assertEqual(reopened.text(), file.text())
        return reopened


# ---------------------------------------------------------------------------
# create
# ---------------------------------------------------------------------------


class CreateTests(WriteAssertions):
    def test_every_creatable_format(self):
        for name in CREATABLE:
            with self.subTest(format=name):
                file = L.create(name)
                self.assertEqual(str(file.format).lower(), name)
                self.assertFalse(file.signed or file.locked or file.closed)
                self.assertEqual(file.get("/", 0)["path"], "/")
                self.assertReopens(file)
                self.assertEqual(file.capabilities()["verbs"]["create"], {"supported": True})

    def test_format_values_and_names(self):
        for value in (L.FileFormat.Docx, "docx", "DOCX", ".docx", "Docx"):
            with self.subTest(value=value):
                self.assertEqual(str(L.create(value).format), "Docx")
        self.assertEqual(str(L.create(format="xlsx").format), "Xlsx")

    def test_bad_format_arguments(self):
        error = raised(lambda: L.create("docz"))
        self.assertError(error, L.PropertyError, "invalid_argument")
        self.assertIn("docx", error.suggestion)
        self.assertError(raised(lambda: L.create(5)), L.PropertyError, "invalid_argument")

    def test_iwork_writes_round_trip(self):
        # Record 0296: the iWork leaves' focused writes through litchi_py.
        pages = L.open(fixture("iwork/pages/body-sections-unicode.pages"))
        edited = pages.set("/body/p[2]", {"text": "From Python"})
        self.assertEqual(L.open_bytes(edited.to_bytes()).get("/body/p[2]", 0)["text"], "From Python")
        numbers = L.open(fixture("iwork/numbers/basic.numbers"))
        edited = numbers.set("/sheet[1]/table[1]/B3", {"value": 12})
        cell = L.open_bytes(edited.to_bytes()).get("/sheet[1]/table[1]/B3", 0)
        self.assertEqual(cell["format"]["value"], 12)
        keynote = L.open(fixture("iwork/keynote/basic.key"))
        edited = keynote.set("/slide[1]", {"hidden": True})
        self.assertTrue(L.open_bytes(edited.to_bytes()).get("/slide[1]", 0)["format"]["hidden"])

    def test_keynote_structural_writes_round_trip(self):
        # Record 0298: Keynote slides, paragraphs and drawables.
        deck = L.open(fixture("iwork/keynote/multi-slide.key"))
        copied, path = deck.add("/", "slide", {"from": "/slide[2]"}, index=0)
        self.assertEqual(path, "/slide[1]")
        reopened = L.open_bytes(copied.to_bytes())
        self.assertEqual(reopened.get("/slide[1]/placeholder[title]", 0)["text"], "Agenda")
        removed = L.open_bytes(copied.remove("/slide[1]").to_bytes())
        self.assertEqual(removed.get("/slide[1]/placeholder[title]", 0)["text"], "Litchi multi-slide fixture")
        self.assertEqual(len(removed.query("slide")), len(deck.query("slide")))
        swapped = deck.swap("/slide[1]", "/slide[2]")
        self.assertEqual(swapped.get("/slide[1]/placeholder[title]", 0)["text"], "Agenda")
        moved = swapped.move("/slide[1]", index=1)
        self.assertEqual(moved.get("/slide[2]/placeholder[title]", 0)["text"], "Agenda")
        boxed, box = deck.add("/slide[1]", "textbox", {"text": "From Python", "x": "1in"})
        self.assertEqual(box, "/slide[1]/shape[2]")
        self.assertEqual(L.open_bytes(boxed.to_bytes()).get(box, 0)["text"], "From Python")
        body = "/slide[2]/placeholder[body]"
        added, paragraph = deck.add(body, "paragraph", {"text": "Fourth topic"})
        self.assertEqual(paragraph, "/slide[2]/placeholder[2]/paragraph[4]")
        edited = added.set(body + "/p[1]", {"text": "Changed"})
        self.assertEqual(edited.get(body + "/p[1]", 0)["text"], "Changed")
        # Record 0309: removing a title placeholder hides it (`hidden` is
        # set-only); setting `hidden` to false shows it again.
        hidden = deck.remove("/slide[1]/placeholder[title]")
        self.assertNotEqual(hidden.to_bytes(), deck.to_bytes())
        shown = L.open_bytes(hidden.to_bytes()).set("/slide[1]/placeholder[title]", {"hidden": False})
        self.assertEqual(shown.get("/slide[1]/placeholder[title]", 0)["text"], "Litchi multi-slide fixture")
        created = L.create("keynote")
        grown, _ = created.add("/", "slide")
        self.assertEqual(len(grown.query("slide")), len(created.query("slide")) + 1)

    def test_numbers_sheets_and_tables_are_added_and_removed(self):
        # Record 0301: Numbers sheet and table lifecycle through litchi_py.
        workbook = L.create("numbers")
        added, path = workbook.add("/", "sheet", {"name": "Extra"}, index=0)
        self.assertEqual(path, "/sheet[1]")
        with_table, table = added.add("/sheet[1]", "table", {"name": "Data", "rowCount": 3, "columnCount": 2})
        self.assertEqual(table, "/sheet[1]/table[1]")
        reopened = L.open_bytes(with_table.to_bytes())
        self.assertEqual(reopened.get(table, 0)["format"]["rowCount"], 3)
        written = reopened.set("/sheet[1]/table[1]/B2", {"value": 5})
        self.assertEqual(L.open_bytes(written.to_bytes()).get("/sheet[1]/table[1]/B2", 0)["format"]["value"], 5)
        removed = with_table.remove("/sheet[1]")
        self.assertEqual(L.open_bytes(removed.to_bytes()).get("/", 0)["format"]["sheetCount"], 1)

    def test_iwork_element_writes_round_trip(self):
        # Record 0309: Numbers rows and columns, character properties and
        # table cells across the iWork formats.
        workbook = L.create("numbers").set("/sheet[1]/table[1]/B2", {"value": 5})
        rows = workbook.get("/sheet[1]/table[1]", 0)["format"]["rowCount"]
        grown, row = workbook.add("/sheet[1]/table[1]", "row", {}, index=1)
        self.assertEqual(row, "/sheet[1]/table[1]/row[2]")
        reopened = L.open_bytes(grown.to_bytes())
        self.assertEqual(reopened.get("/sheet[1]/table[1]", 0)["format"]["rowCount"], rows + 1)
        self.assertEqual(reopened.get("/sheet[1]/table[1]/B3", 0)["format"]["value"], 5)
        shrunk = L.open_bytes(reopened.remove(row).to_bytes())
        self.assertEqual(shrunk.get("/sheet[1]/table[1]/B2", 0)["format"]["value"], 5)
        widened, column = workbook.add("/sheet[1]/table[1]", "col", {}, index=0)
        self.assertEqual(column, "/sheet[1]/table[1]/col[1]")
        self.assertEqual(L.open_bytes(widened.to_bytes()).get("/sheet[1]/table[1]/C2", 0)["format"]["value"], 5)
        bold = workbook.set("/sheet[1]/table[1]/B2", {"font.bold": True, "font.size": "18pt"})
        self.assertEqual(L.open_bytes(bold.to_bytes()).get("/sheet[1]/table[1]/B2", 0)["format"]["value"], 5)
        pages = L.open(fixture("iwork/pages/basic.pages"))
        styled = pages.set("/body/p[1]", {"font.italic": True})
        self.assertEqual(L.open_bytes(styled.to_bytes()).get("/body/p[1]", 0)["text"], pages.get("/body/p[1]", 0)["text"])
        tables = L.open(fixture("iwork/pages/body-table-catalog-native.pages"))
        written = tables.set("/body/tbl[1]/tr[1]/tc[1]", {"text": "From Python"})
        self.assertEqual(L.open_bytes(written.to_bytes()).get("/body/tbl[1]/tr[1]/tc[1]", 0)["text"], "From Python")
        deck = L.open(fixture("iwork/keynote/multi-slide.key"))
        boxed, box = deck.add("/slide[1]", "textbox", {"text": "Styled"})
        run = box + "/p[1]/run[1]"
        styled = L.open_bytes(boxed.set(run, {"bold": True, "color": "C80000"}).to_bytes())
        self.assertTrue(styled.get(run, 0)["format"]["bold"])
        cells = L.open(fixture("iwork/keynote/table-discovery.key"))
        cell = cells.query("tc")[0]["path"]
        self.assertEqual(L.open_bytes(cells.set(cell, {"text": "Cell"}).to_bytes()).get(cell, 0)["text"], "Cell")

    def test_a_created_file_has_no_path(self):
        error = raised(lambda: L.create("docx").save())
        self.assertError(error, L.IoError, "file_required")
        self.assertIsInstance(error, OSError)


# ---------------------------------------------------------------------------
# Verbs, per format
# ---------------------------------------------------------------------------


class DocxVerbTests(WriteAssertions):
    def test_set_typed_values_on_runs_and_paragraphs(self):
        file = docx_with("Hello")
        before = file.to_bytes()
        changed = file.set("/body/p[1]/r[1]", {"bold": True, "size": 14, "color": "1F4E79"})
        run = changed.get("/body/p[1]/r[1]", 0)["format"]
        self.assertEqual((run["bold"], run["size"], run["color"]), (True, "14pt", "#1F4E79"))
        self.assertEqual(changed.set("/body/p[1]/r[1]", {"size": 10.5}).get("/body/p[1]/r[1]", 0)["format"]["size"], "10.5pt")
        cleared = changed.set("/body/p[1]/r[1]", {"bold": None})
        self.assertNotIn("bold", cleared.get("/body/p[1]/r[1]", 0)["format"])
        aligned = file.set("/body/p[1]", {"align": "center", "text": "Bye"})
        self.assertEqual(aligned.get("/body/p[1]", 0)["format"]["align"], "center")
        self.assertEqual(aligned.get("/body/p[1]", 0)["text"], "Bye")
        self.assertUnchanged(file, before)
        self.assertEqual(L.open_bytes(changed.to_bytes()).get("/body/p[1]/r[1]", 0), changed.get("/body/p[1]/r[1]", 0))

    def test_set_through_a_scoped_selector(self):
        file = docx_with("a", "b")
        bold = file.set("/body/p[1]/r", {"bold": True})
        self.assertTrue(bold.get("/body/p[1]/r[1]", 0)["format"]["bold"])
        self.assertNotIn("bold", bold.get("/body/p[2]/r[1]", 0)["format"])
        error = raised(lambda: file.set("run", {"bold": True}))
        self.assertError(error, L.PathError, "bare_selector_rejected")

    def test_add_paragraphs_and_runs_with_placement(self):
        file = docx_with("A", "B")
        added, path = file.add("/body", "paragraph", {"text": "C"})
        self.assertEqual(path, "/body/p[3]")
        self.assertEqual(texts(added, "/body")[:3], ["A", "B", "C"])
        self.assertEqual(file.add("/body", "paragraph", {"text": "Z"}, index=0)[1], "/body/p[1]")
        self.assertEqual(texts(file.add("/body", "paragraph", {"text": "Z"}, after="/body/p[1]")[0], "/body")[:3], ["A", "Z", "B"])
        self.assertEqual(texts(file.add("/body", "paragraph", {"text": "Z"}, before="/body/p[1]")[0], "/body")[:3], ["Z", "A", "B"])
        with_run, run = file.add("/body/p[1]", "run", {"text": "!", "bold": True})
        self.assertEqual(run, "/body/p[1]/r[2]")
        self.assertEqual(with_run.get(run, 0)["format"]["bold"], True)
        self.assertEqual(file.add("/body", "p")[1], "/body/p[3]")
        self.assertEqual(file.add("/body", type="paragraph", props=["text=kw"])[0].get("/body/p[3]", 0)["text"], "kw")

    def test_remove_move_and_swap(self):
        file = docx_with("A", "B", "C")
        before = file.to_bytes()
        self.assertEqual(texts(file.remove("/body/p[2]"), "/body")[:2], ["A", "C"])
        self.assertEqual(texts(file.move("/body/p[3]", index=0), "/body")[:3], ["C", "A", "B"])
        self.assertEqual(texts(file.move("/body/p[1]", after="/body/p[3]"), "/body")[:3], ["B", "C", "A"])
        self.assertEqual(texts(file.move("/body/p[3]", before="/body/p[1]"), "/body")[:3], ["C", "A", "B"])
        self.assertEqual(texts(file.move("/body/p[1]"), "/body")[:3], ["B", "C", "A"])
        self.assertEqual(texts(file.move("/body/p[2]", to="/body", index=0), "/body")[:3], ["B", "A", "C"])
        moved = file.move("/body/p[1]/r[1]", to="/body/p[2]")
        self.assertEqual(texts(moved, "/body/p[2]"), ["B", "A"])
        self.assertEqual(texts(file.swap("/body/p[1]", "/body/p[3]"), "/body")[:3], ["C", "B", "A"])
        self.assertUnchanged(file, before)

    def test_placement_arguments(self):
        file = docx_with("A")
        for call in (
            lambda: file.add("/body", "paragraph", index=0, after="/body/p[1]"),
            lambda: file.move("/body/p[1]", index=0, before="/body/p[1]"),
            lambda: file.add("/body", "paragraph", index=-1),
            lambda: file.move("/body/p[1]", index=-2),
        ):
            self.assertError(raised(call), L.PropertyError, "invalid_argument")
        error = raised(lambda: file.move("/body/p[1]", to="/body", after="/body/p[1]"))
        self.assertError(error, L.UnsupportedError, "unsupported_capability")
        error = raised(lambda: file.add("/body", "paragraph", after="/body/p[@text=A]"))
        self.assertError(error, L.PathError, "invalid_path")


class OoxmlFollowUpTests(WriteAssertions):
    def test_record_0325_follow_ups(self):
        """Record 0325: story paragraphs, marks, level keys, links, note runs."""
        doc = L.create("docx")
        doc, _ = doc.add("/body", "paragraph", {"text": "Body"})
        doc, _ = doc.add("/", "header", {"text": "Top"})
        doc, second = doc.add("/header[1]", "paragraph", {"text": "Second"})
        self.assertEqual(second, "/header[1]/p[2]")
        doc = doc.set(second, {"align": "center"})
        doc, mark = doc.add("/body/p[1]", "bookmark", {"name": "Mark", "text": "here"})
        doc, _ = doc.add("/body/p[1]", "page", {})
        doc, _ = doc.add("/numbering", "abstractNum", {"level0.format": "upperRoman"})
        reopened = self.assertReopens(doc)
        self.assertEqual(reopened.get(second, 0)["format"]["align"], "center")
        self.assertEqual(reopened.get(mark, 0)["format"]["name"], "Mark")
        self.assertEqual(reopened.get("/field[1]", 0)["format"]["instruction"], "PAGE")
        deck = L.create("pptx")
        deck, _ = deck.add("/", "slide", {"layout": "Blank"})
        deck, shape = deck.add("/slide[1]", "shape", {"text": "Go", "width": "1in", "height": "1in", "link": "https://example.com/"})
        self.assertEqual(self.assertReopens(deck).get(shape, 0)["format"]["link"], "https://example.com/")
        book = L.create("xlsx")
        book, note = book.add("/Sheet1", "comment", {"ref": "B2", "text": "Bold", "font.bold": True})
        self.assertEqual(self.assertReopens(book).get(note, 0)["format"]["font.bold"], True)


class XlsxVerbTests(WriteAssertions):
    def test_cells(self):
        file = L.create("xlsx")
        before = file.to_bytes()
        changed = file.set("/sheet[1]/A1", {"value": 42})
        self.assertEqual(changed.get("/sheet[1]/A1", 0)["format"]["type"], "Number")
        self.assertEqual(cell_text(changed, "Sheet1!A1"), "42")
        formula = changed.set("/Sheet1/B1", {"formula": "SUM(A1:A5)"})
        self.assertEqual(formula.get("/sheet[1]/B1", 0)["format"]["formula"], "SUM(A1:A5)")
        text = file.set("/sheet[1]/A2", {"value": "hello"})
        self.assertEqual(cell_text(text, "/sheet[1]/A2"), "hello")
        self.assertUnchanged(file, before)
        self.assertEqual(cell_text(self.assertReopens(changed), "/sheet[1]/A1"), "42")

    def test_sheets(self):
        file = L.create("xlsx")
        file, path = file.add("/", "sheet", {"name": "Two"})
        self.assertEqual(path, "/sheet[2]")
        names = lambda f: [child["preview"] for child in f.get("/")["children"] if child["type"] == "sheet"]
        self.assertEqual(names(file), ["Sheet1", "Two"])
        self.assertEqual(names(file.move("/sheet[2]", index=0)), ["Two", "Sheet1"])
        self.assertEqual(names(file.swap("/sheet[1]", "/sheet[2]")), ["Two", "Sheet1"])
        self.assertEqual(names(file.remove("/sheet[2]")), ["Sheet1"])
        self.assertEqual(names(file.set("/sheet[2]", {"name": "Renamed"})), ["Sheet1", "Renamed"])


class PptxVerbTests(WriteAssertions):
    def setUp(self):
        file = L.create("pptx")
        # A blank layout: slides added from a layout with placeholders hold
        # those placeholders as shapes (record 0252).
        layout = "Blank"
        file, self.first = file.add("/", "slide", {"layout": layout})
        file, self.second = file.add("/", "slide", {"layout": layout})
        self.file = file

    def slide_names(self, file):
        return [file.get(path, 0)["format"]["name"] for path in child_paths(file, "/", "slide")]

    def test_slides(self):
        self.assertEqual((self.first, self.second), ("/slide[1]", "/slide[2]"))
        error = raised(lambda: L.create("pptx").add("/", "slide"))
        self.assertError(error, L.PropertyError, "missing_property")
        names = self.slide_names(self.file)
        self.assertEqual(self.slide_names(self.file.move("/slide[2]", index=0)), names[::-1])
        self.assertEqual(self.slide_names(self.file.swap("/slide[1]", "/slide[2]")), names[::-1])
        self.assertEqual(self.slide_names(self.file.remove("/slide[1]")), names[1:])
        hidden = self.file.set("/slide[2]", {"hidden": True})
        self.assertIs(hidden.get("/slide[2]", 0)["format"]["hidden"], True)

    def test_shapes(self):
        file, shape = self.file.add(self.first, "shape", {"text": "Box", "width": "2in", "height": "1in"})
        self.assertEqual(shape, "/slide[1]/shape[1]")
        self.assertEqual(file.get(shape, 0)["text"], "Box")
        changed = file.set(shape, {"text": "Changed", "x": "1in"})
        view = changed.get(shape, 0)
        self.assertEqual((view["text"], view["format"]["x"]), ("Changed", "2.54cm"))
        self.assertEqual(child_paths(changed.remove(shape), self.first), [])
        self.assertEqual(self.assertReopens(changed).get(shape, 0)["text"], "Changed")

    def test_smartart(self):
        # Record 0284: typed SmartArt CRUD passes through unchanged.
        file, art = self.file.add(self.first, "smartart", {"layout": "process", "text": "Plan\n\tScope\nShip"})
        self.assertEqual(art, "/slide[1]/smartart[1]")
        view = file.get(art, 0)
        self.assertEqual((view["text"], view["format"]["layout"]), ("Plan\n\tScope\nShip", "process"))
        file, node = file.add(art, "node", {"text": "Learn"})
        self.assertEqual(node, "/slide[1]/smartart[1]/node[3]")
        file = file.set("/slide[1]/smartart[1]/node[1]/node[1]", {"text": "Size"})
        file = file.remove("/slide[1]/smartart[1]/node[2]")
        self.assertEqual(self.assertReopens(file).get(art, 0)["text"], "Plan\n\tSize\nLearn")

    def test_record_0321_elements(self):
        # Record 0321: theme, master and layout names, hyperlinks and media.
        theme = self.file.set("/theme", {"accent1": "FF0000", "majorFont": "Georgia"})
        view = self.assertReopens(theme).get("/theme", 0)["format"]
        self.assertEqual((view["accent1"], view["headingFont"]), ("#FF0000", "Georgia"))
        renamed = self.file.set("/slidemaster[1]", {"name": "Brand"})
        self.assertEqual(renamed.get("/slidemaster[1]", 0)["format"]["name"], "Brand")
        file, _ = self.file.add(self.first, "shape", {"text": "Docs", "width": "2in", "height": "1in"})
        file, link = file.add("/slide[1]/shape[1]", "hyperlink", {"link": "slide[2]", "tooltip": "Next"})
        self.assertEqual(link, "/slide[1]/shape[1]/hyperlink")
        self.assertEqual(self.assertReopens(file).get(link, 0)["format"]["link"], "slide[2]")
        file = file.remove(link)
        self.assertEqual(len(list(file.query("hyperlink"))), 0)
        png = (
            "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAEklEQVR4"
            "nGP4z8DwHxkzoAsAAA8hD/EEN8afAAAAAElFTkSuQmCC"
        )
        mp4 = "data:video/mp4;base64,AAAAGGZ0eXBtcDQyAAAAAG1wNDJpc29t"
        file, video = self.file.add(self.first, "video", {"src": mp4, "poster": png})
        self.assertEqual(video, "/slide[1]/video[1]")
        self.assertEqual(self.assertReopens(file).get(video, 0)["type"], "video")
        error = raised(lambda: self.file.add(self.first, "video", {"src": mp4}))
        self.assertError(error, L.PropertyError, "missing_property")

    def test_record_0333_p0_residue(self):
        # Record 0333: paragraph, run, table and row adds, removals, cell
        # properties and comment removal.
        file, shape = self.file.add(self.first, "shape", {"text": "alpha", "width": "2in", "height": "1in"})
        file, paragraph = file.add(shape, "paragraph", {"text": "beta", "bold": True})
        self.assertEqual(paragraph, shape + "/paragraph[2]")
        file, run = file.add(paragraph, "run", {"text": "!", "italic": True})
        self.assertEqual(run, paragraph + "/run[2]")
        reopened = self.assertReopens(file)
        self.assertEqual(reopened.get(paragraph, 0)["text"], "beta!")
        self.assertEqual(reopened.get(run, 0)["format"]["italic"], True)
        file = file.remove(shape + "/paragraph[1]")
        self.assertEqual(file.get(shape, 0)["text"], "beta!")
        file, table = file.add(self.first, "table", {"data": "a,b;c,d"})
        file, row = file.add(table, "tr", {"c1": "e"})
        self.assertEqual(file.get(row + "/tc[1]", 0)["text"], "e")
        file = file.set(row + "/tc[2]", {"fill": "accent1", "valign": "center"})
        cell = self.assertReopens(file).get(row + "/tc[2]", 0)["format"]
        self.assertEqual((cell["fill"], cell["valign"]), ("accent1", "center"))
        file = file.remove(table + "/tr[1]")
        self.assertEqual(file.get(table, 0)["format"]["rows"], 2)
        error = raised(lambda: file.remove(shape + "/paragraph[1]"))
        self.assertError(error, L.UnsupportedError, "unsupported_capability")


class XlsbVerbTests(WriteAssertions):
    def test_cells_and_sheets(self):
        file = L.create("xlsb")
        changed = file.set("/sheet[1]/A1", {"value": 2})
        self.assertEqual(cell_text(changed, "/sheet[1]/A1"), "2")
        self.assertEqual(cell_text(changed.remove("/sheet[1]/A1"), "/sheet[1]/A1"), "(empty)")
        two, path = changed.add("/", "sheet", {"name": "Two"})
        self.assertEqual(path, "/sheet[2]")
        names = lambda f: [child["preview"] for child in f.get("/")["children"] if child["type"] == "sheet"]
        self.assertEqual(names(two.move("/sheet[2]", index=0)), ["Two", "Sheet1"])
        self.assertEqual(names(two.swap("/sheet[1]", "/sheet[2]")), ["Two", "Sheet1"])
        self.assertEqual(names(two.remove("/sheet[2]")), ["Sheet1"])
        self.assertEqual(cell_text(self.assertReopens(changed), "/sheet[1]/A1"), "2")


class LegacyVerbTests(WriteAssertions):
    def test_doc_sets_paragraph_text(self):
        file = L.create("doc")
        changed = file.set("/body/p[1]", {"text": "Hello DOC"})
        self.assertEqual(changed.get("/body/p[1]", 0)["text"], "Hello DOC")
        self.assertIn("Hello DOC", self.assertReopens(changed).text())

    def test_xls_cells_and_sheet_names(self):
        file = L.create("xls").set("/sheet[1]/A1", {"value": 7})
        self.assertEqual(cell_text(file, "/sheet[1]/A1"), "7")
        self.assertEqual(file.set("/sheet[1]", {"name": "Data"}).get("/sheet[1]", 0)["preview"], "Data")
        self.assertEqual(cell_text(file.remove("/sheet[1]/A1"), "/sheet[1]/A1"), "")
        self.assertEqual(cell_text(self.assertReopens(file), "/sheet[1]/A1"), "7")

    def test_ppt_shapes_and_slides(self):
        file = L.open(fixture(EDITABLE_PPT))
        changed = file.set("/slide[1]/shape[3]", {"text": "x", "x": "1in"})
        view = changed.get("/slide[1]/shape[3]", 0)
        self.assertEqual((view["text"], view["format"]["x"]), ("x", "2.54cm"))
        self.assertEqual(self.assertReopens(changed).get("/slide[1]/shape[3]", 0), view)
        show = L.open(fixture(SLIDESHOW_PPT))
        titles = [show.get(path, 0)["preview"] for path in child_paths(show, "/", "slide")]
        order = lambda f: [f.get(path, 0)["preview"] for path in child_paths(f, "/", "slide")]
        self.assertEqual(order(show.move("/slide[2]", index=0)), titles[::-1])
        self.assertEqual(order(show.swap("/slide[1]", "/slide[2]")), titles[::-1])
        self.assertEqual(order(show.remove("/slide[2]")), titles[:1])
        # PowerPoint-authored text of another length (record 0260): the
        # title is slide-list text, the subtitle has paragraphs.
        retitled = show.set("/slide[1]/shape[1]", {"text": "A longer replacement title"})
        self.assertEqual(retitled.get("/slide[1]/shape[1]", 0)["text"], "A longer replacement title")
        self.assertEqual(self.assertReopens(retitled).get("/slide[1]", 0)["preview"], "A longer replacement title")
        # A body with several formatting runs keeps the same-length contract.
        error = raised(lambda: show.set("/slide[1]/shape[2]/paragraph[3]", {"text": "x"}))
        self.assertError(error, L.UnsupportedError, "unsupported_capability")
        # Blank slides, paragraphs and speaker notes.
        added, path = show.add("/", "slide", index=0)
        self.assertEqual((path, order(added)), ("/slide[1]", ["(untitled)"] + titles))
        notes = L.open(fixture(NOTES_PPT))
        line = notes.set("/slide[1]/shape[2]/paragraph[2]", {"text": "Rewritten line"})
        self.assertEqual(line.get("/slide[1]/shape[2]/paragraph[2]", 0)["text"], "Rewritten line")
        noted = notes.set("/slide[1]/notes", {"text": "New notes"})
        self.assertEqual(noted.get("/slide[1]/notes", 0)["text"], "New notes")
        self.assertEqual(self.assertReopens(noted).get("/slide[1]/notes", 0)["text"], "New notes")


class RtfVerbTests(WriteAssertions):
    def test_add_and_set(self):
        file = L.create("rtf")
        file, first = file.add("/body", "paragraph", {"text": "One"})
        file, second = file.add("/body", "paragraph", {"text": "Two"})
        file, run = file.add(first, "run", {"text": " more"})
        self.assertEqual((first, second, run), ("/body/p[1]", "/body/p[2]", "/body/p[1]/r[2]"))
        self.assertEqual(texts(file, "/body"), ["One more", "Two"])
        self.assertEqual(file.set(first, {"align": "center"}).get(first, 0)["format"]["align"], "center")
        self.assertIs(file.set(first + "/r[1]", {"bold": True}).get(first + "/r[1]", 0)["format"]["bold"], True)
        self.assertEqual(self.assertReopens(file).text(), file.text())

    def test_structure(self):
        file = L.open_bytes(RTF_INPUT)
        self.assertEqual(texts(file.remove("/body/p[2]"), "/body"), ["First", "Third", "Fourth"])
        self.assertEqual(texts(file.move("/body/p[4]", index=0), "/body"), ["Fourth", "First", "Second", "Third"])
        self.assertEqual(texts(file.swap("/body/p[1]", "/body/p[4]"), "/body"), ["Fourth", "Second", "Third", "First"])


class OdtVerbTests(WriteAssertions):
    """ODT writes through the leaf body edits (record 0276)."""

    def test_round_trip(self):
        source = L.open(fixture("odf/corpus/writer-table.odt"))
        before = source.to_bytes()
        file, path = source.add("/body", "paragraph", {"text": "Added  text"})
        self.assertTrue(path.startswith("/body/p["), path)
        file = file.set(path, {"align": "center", "bold": True})
        file = file.set("/body/tbl[1]/tr[1]/tc[1]", {"text": "cell"})
        file = file.set("/metadata", {"title": "Round trip"})
        reopened = self.assertReopens(file)
        self.assertEqual(reopened.get(path, 0)["text"], "Added  text")
        self.assertEqual(reopened.get(path, 0)["format"]["align"], "center")
        self.assertIs(reopened.get(path, 0)["format"]["bold"], True)
        self.assertEqual(reopened.get("/body/tbl[1]/tr[1]/tc[1]", 0)["text"], "cell")
        self.assertEqual(reopened.get("/metadata", 0)["format"]["title"], "Round trip")
        self.assertUnchanged(source, before)

    def test_structure_and_create(self):
        file = L.create("odt")
        for text in ("a", "b", "c"):
            file, _ = file.add("/body", "paragraph", {"text": text})
        self.assertEqual(texts(file.remove("/body/p[2]"), "/body"), ["a", "c"])
        self.assertEqual(texts(file.move("/body/p[3]", index=0), "/body"), ["c", "a", "b"])
        self.assertEqual(texts(file.swap("/body/p[1]", "/body/p[3]"), "/body"), ["c", "b", "a"])
        batched = file.batch(
            [
                {"command": "set", "path": f"/body/p[{index}]", "props": {"italic": True}}
                for index in (1, 2, 3)
            ]
        ).file
        self.assertIs(batched.get("/body/p[3]", 0)["format"]["italic"], True)


class OdpVerbTests(WriteAssertions):
    """ODP writes through checked scene splices (record 0278)."""

    def test_round_trip(self):
        file = L.open(fixture(ODP_INPUT))
        before = file.to_bytes()
        titles = [file.get(path, 0)["preview"] for path in child_paths(file, "/", "slide")]
        changed = file.set("/slide[1]/shape[1]", {"text": "Retitled", "width": "20cm"})
        view = changed.get("/slide[1]/shape[1]", 0)
        self.assertEqual((view["text"], view["format"]["width"]), ("Retitled", "20cm"))
        added, path = changed.add("/slide[2]", "textbox", {"text": "Box", "width": "4cm", "height": "2cm"})
        self.assertEqual(path, "/slide[2]/shape[2]")
        order = lambda f: [f.get(p, 0)["preview"] for p in child_paths(f, "/", "slide")]
        self.assertEqual(order(file.swap("/slide[1]", "/slide[2]"))[:2], [titles[1], titles[0]])
        self.assertEqual(order(file.move("/slide[3]", index=0))[0], titles[2])
        self.assertEqual(len(order(file.remove("/slide[2]"))), len(titles) - 1)
        noted = added.set("/metadata", {"title": "ODP deck"})
        reopened = self.assertReopens(noted)
        self.assertEqual(reopened.get("/slide[1]/shape[1]", 0)["text"], "Retitled")
        self.assertEqual(reopened.get(path, 0)["text"], "Box")
        self.assertEqual(reopened.get("/", 0)["format"]["title"], "ODP deck")
        self.assertUnchanged(file, before)
        with tempfile.TemporaryDirectory() as directory:
            target = os.path.join(directory, "deck.odp")
            noted.save_as(target)
            self.assertEqual(L.open(target).get(path, 0)["text"], "Box")


class FlatOdfTests(WriteAssertions):
    """Flat ODF (record 0292): .fodt opens, and every flat family edits its
    packaged projection and republishes into the flat bytes."""

    def test_flat_text_paragraph_round_trip(self):
        source = L.open(fixture("odf/odt/title-field-invalidate.fodt"))
        before = source.to_bytes()
        paths = [node["path"] for node in source.query("p")]
        target = next(p for p in paths if source.get(p, 0)["text"] == "body text")
        edited = source.set(target, {"text": "new body"})
        published = edited.to_bytes()
        self.assertEqual(published, before.replace(b">body text<", b">new body<"))
        self.assertEqual(L.open_bytes(published).get(target, 0)["text"], "new body")
        self.assertUnchanged(source, before)

    def test_flat_drawing_and_moves_across_object_formats(self):
        drawing = L.open(fixture("libreoffice-core/xmloff/qa/unit/data/tdf161327_LatheEndAngle.fodg"))
        renamed = drawing.set("/page[1]", {"name": "Flat"})
        self.assertEqual(L.open_bytes(renamed.to_bytes()).get("/page[1]", 0)["format"]["name"], "Flat")
        # Record 0292: object-format verbs are element-level.
        chart = L.create("odc")
        for values in ("t.B2:B5", "t.C2:C5"):
            chart, _ = chart.add("/", "series", {"values": values})
        swapped = chart.swap("/series[1]", "/series[2]")
        self.assertEqual(swapped.get("/series[1]", 0)["format"]["values"], "t.C2:C5")


class OdfObjectTests(WriteAssertions):
    """ODF object formats (record 0288): producer drawing structure, the
    metadata family and typed chart writes."""

    def test_producer_drawing_structure_and_metadata(self):
        source = L.open(fixture("libreoffice-core/desktop/qa/data/3page.odg"))
        before = source.to_bytes()
        file, path = source.add("/", "page", {"name": "Front"}, index=0)
        self.assertEqual(path, "/page[1]")
        name = lambda f, p: f.get(p, 0)["format"]["name"]
        moved = file.move("/page[1]", index=3)
        self.assertEqual(name(moved, "/page[4]"), "Front")
        swapped = moved.swap("/page[1]", "/page[2]")
        self.assertEqual(name(swapped, "/page[1]"), "page2")
        removed = swapped.remove("/page[4]")
        noted = removed.set("/metadata", {"title": "Drawing", "author": "Ada"})
        reopened = self.assertReopens(noted)
        self.assertEqual(reopened.metadata()["title"], "Drawing")
        self.assertEqual(reopened.get("/", 0)["format"]["pageCount"], 3)
        self.assertUnchanged(source, before)
        batched = source.batch(
            [
                {"command": "set", "path": f"/page[{index}]", "props": {"name": f"P{index}"}}
                for index in (1, 2, 3)
            ]
        ).file
        self.assertEqual(name(batched, "/page[3]"), "P3")

    def test_typed_chart_and_fresh_image(self):
        chart = L.create("odc").set("/", {"title": "Sales", "chartType": "bar"})
        chart, path = chart.add("/", "series", {"values": "local-table.B2:B5"})
        self.assertEqual(path, "/series[1]")
        reopened = self.assertReopens(chart)
        self.assertEqual(reopened.get("/", 0)["format"]["title"], "Sales")
        self.assertEqual(reopened.get("/series[1]", 0)["format"]["values"], "local-table.B2:B5")
        image = L.create("odi").set("/frame[1]", {"x": "1cm", "y": "1cm", "width": "4cm", "height": "3cm"})
        self.assertEqual(self.assertReopens(image).get("/frame[1]", 0)["format"]["width"], "4cm")


    def test_xlsx_worksheet_features(self):
        """Record 0320: notes, conditional formats, validations and links."""
        book = L.create("xlsx")
        book, note = book.add("/Sheet1", "comment", {"ref": "B2", "text": "Check", "author": "Ann"})
        self.assertEqual(note, "/sheet[1]/comment[1]")
        book, rule = book.add(
            "/Sheet1",
            "conditionalformatting",
            {"type": "cellIs", "ref": "A2:A11", "operator": "greaterThan", "value": "80", "fill": "C6EFCE"},
        )
        self.assertEqual(rule, "/sheet[1]/cf[1]")
        book, _ = book.add("/Sheet1", "validation", {"type": "list", "ref": "C2:C9", "formula1": "a,b"})
        book = book.set("/Sheet1/D4", {"value": "Repo", "link": "https://example.test/"})
        reopened = self.assertReopens(book)
        self.assertEqual(reopened.get("/Sheet1/comment[1]", 0)["format"]["author"], "Ann")
        self.assertEqual(reopened.get("/Sheet1/cf[1]", 0)["format"]["type"], "cellIs")
        self.assertEqual(reopened.get("/Sheet1/dataValidation[1]", 0)["format"]["formula1"], '"a,b"')
        self.assertEqual(reopened.get("/Sheet1/hyperlink[1]", 0)["format"]["url"], "https://example.test/")

    def test_xlsb_worksheet_features(self):
        """Record 0322: XLSB takes the XLSX worksheet-feature vocabulary."""
        book = L.create("xlsb")
        book, note = book.add("/Sheet1", "comment", {"ref": "B2", "text": "Check", "author": "Ann"})
        self.assertEqual(note, "/sheet[1]/comment[1]")
        book, rule = book.add(
            "/Sheet1",
            "conditionalformatting",
            {"type": "cellIs", "ref": "A2:A11", "operator": "greaterThan", "value": "80", "fill": "C6EFCE"},
        )
        self.assertEqual(rule, "/sheet[1]/cf[1]")
        book, _ = book.add("/Sheet1", "validation", {"type": "list", "ref": "C2:C9", "formula1": "a,b"})
        book, _ = book.add("/Sheet1/D4", "hyperlink", {"url": "https://example.test/"})
        book, _ = book.add("/Sheet1", "rowbreak", {"row": 10})
        reopened = self.assertReopens(book)
        self.assertEqual(reopened.get("/Sheet1/comment[1]", 0)["format"]["author"], "Ann")
        self.assertEqual(reopened.get("/Sheet1/cf[1]", 0)["format"]["type"], "cellIs")
        self.assertEqual(reopened.get("/Sheet1/dataValidation[1]", 0)["format"]["formula1"], '"a,b"')
        self.assertEqual(reopened.get("/Sheet1/hyperlink[1]", 0)["format"]["url"], "https://example.test/")
        self.assertEqual(reopened.get("/Sheet1/rowbreak[1]", 0)["format"]["row"], 10)

    def test_xls_worksheet_features(self):
        """Record 0327: XLS takes the XLSX worksheet-feature vocabulary."""
        book = L.create("xls")
        book, note = book.add("/Sheet1", "comment", {"ref": "B2", "text": "Check", "author": "Ann"})
        self.assertEqual(note, "/sheet[1]/comment[1]")
        book, rule = book.add(
            "/Sheet1",
            "conditionalformatting",
            {"type": "cellIs", "ref": "A2:A11", "operator": "greaterThan", "value": "80", "fill": "00FF00"},
        )
        self.assertEqual(rule, "/sheet[1]/cf[1]")
        book, _ = book.add("/Sheet1", "validation", {"type": "list", "ref": "C2:C9", "formula1": "a,b"})
        book, _ = book.add("/Sheet1/D4", "hyperlink", {"url": "https://example.test/"})
        book, _ = book.add("/Sheet1", "rowbreak", {"row": 10})
        reopened = self.assertReopens(book)
        self.assertEqual(reopened.get("/Sheet1/comment[1]", 0)["format"]["author"], "Ann")
        self.assertEqual(reopened.get("/Sheet1/cf[1]", 0)["format"]["type"], "cellIs")
        self.assertEqual(reopened.get("/Sheet1/dataValidation[1]", 0)["format"]["formula1"], '"a,b"')
        self.assertEqual(reopened.get("/Sheet1/hyperlink[1]", 0)["format"]["url"], "https://example.test/")
        self.assertEqual(reopened.get("/Sheet1/rowbreak[1]", 0)["format"]["row"], 10)

    def test_ods_worksheet_features(self):
        """Record 0327: ODS takes the XLSX worksheet-feature vocabulary."""
        book = L.open(fixture("odf/corpus/calc-two-sheets.ods"))
        book, rule = book.add(
            "/sheet1",
            "conditionalformatting",
            {"type": "cellIs", "ref": "A2:A11", "operator": "greaterThan", "value": "80", "fill": "00FF00"},
        )
        self.assertEqual(rule, "/sheet[1]/cf[1]")
        book, _ = book.add("/sheet1", "validation", {"type": "list", "ref": "C2:C3", "formula1": "a,b"})
        book, _ = book.add("/sheet1/A1", "hyperlink", {"url": "https://example.test/"})
        book, _ = book.add("/sheet1", "rowbreak", {"row": 2})
        reopened = self.assertReopens(book)
        self.assertEqual(reopened.get("/sheet1/cf[1]", 0)["format"]["type"], "cellIs")
        self.assertEqual(reopened.get("/sheet1/dataValidation[1]", 0)["format"]["formula1"], '"a,b"')
        self.assertEqual(reopened.get("/sheet1/hyperlink[1]", 0)["format"]["url"], "https://example.test/")
        self.assertEqual(reopened.get("/sheet1/rowbreak[1]", 0)["format"]["row"], 2)


class CapabilityAgreementTests(WriteAssertions):
    """A write verb the report marks unsupported raises its code."""

    CALLS = {
        "set": lambda file: file.set("/", {"text": "x"}),
        "add": lambda file: file.add("/", "paragraph"),
        "remove": lambda file: file.remove("/x[1]"),
        "move": lambda file: file.move("/x[1]", index=0),
        "swap": lambda file: file.swap("/x[1]", "/x[2]"),
        "batch": lambda file: file.batch([{"command": "set", "path": "/x[1]", "props": {"text": "x"}}]),
        "save": lambda file: file.to_bytes(),
    }

    def files(self):
        for name in CREATABLE:
            yield name, L.create(name)
        for relative, name in READ_ONLY:
            yield name, L.open(fixture(relative))

    def test_unsupported_verbs_raise_the_reported_code(self):
        seen = set()
        for name, file in self.files():
            verbs = file.capabilities()["verbs"]
            for verb, call in self.CALLS.items():
                if verbs[verb]["supported"]:
                    continue
                seen.add(name)
                with self.subTest(format=name, verb=verb):
                    error = raised(lambda: call(file))
                    self.assertIsInstance(error, L.UnsupportedError)
                    self.assertEqual(error.code, verbs[verb]["code"])
        # DOC (0258), PPT (0260) and XLS (0259) support every write verb.
        # DOC (0258), PPT (0260), XLS (0259), ODS (0277), ODT (0276), ODP
        # (0278), OTH and ODM (0282) and the ODF object formats (0292) support
        # every write verb; object-format gaps are element-level.

    def test_read_only_formats_refuse_every_write(self):
        for relative, name in READ_ONLY:
            file = L.open(fixture(relative))
            for verb, call in self.CALLS.items():
                with self.subTest(format=name, verb=verb):
                    self.assertIsInstance(raised(lambda: call(file)), L.UnsupportedError)
            with self.subTest(format=name, verb="edit"):
                self.assertIsInstance(raised(lambda: file.edit(lambda tx: None)), L.UnsupportedError)


class OdsRoundTripTests(WriteAssertions):
    """ODS writes through the unified layer (record 0277)."""

    def test_cells_sheets_and_metadata_round_trip(self):
        source = L.open(fixture("odf/corpus/calc-two-sheets.ods"))
        before = source.to_bytes()
        file = source.set("/sheet1/B2", {"value": 42})
        file = file.set("/sheet1/E5", {"text": "new"})
        file, path = file.add("/", "sheet", {"name": "Extra"})
        self.assertEqual(path, "/sheet[3]")
        file = file.set("/metadata", {"title": "Budget"})
        self.assertUnchanged(source, before)
        reopened = self.assertReopens(file)
        self.assertEqual(cell_text(reopened, "/sheet1/B2"), "42")
        self.assertEqual(cell_text(reopened, "/sheet1/E5"), "new")
        self.assertEqual(reopened.metadata()["title"], "Budget")
        with tempfile.TemporaryDirectory() as directory:
            target = os.path.join(directory, "out.ods")
            file.save_as(target)
            self.assertEqual(cell_text(L.open(target), "/Extra/A1"), "")
        self.assertError(
            raised(lambda: source.set("/sheet1/A1", {"font.bold": True})),
            L.PropertyError,
            "unsupported_property",
        )


# ---------------------------------------------------------------------------
# Properties
# ---------------------------------------------------------------------------


class PropsTests(WriteAssertions):
    def test_dict_and_string_list_are_the_same_assignments(self):
        file = docx_with("a")
        self.assertEqual(
            file.set("/body/p[1]/r[1]", {"bold": True, "size": 14, "color": "FF0000"}).to_bytes(),
            file.set("/body/p[1]/r[1]", ["bold=true", "size=14", "color=FF0000"]).to_bytes(),
        )
        self.assertEqual(
            file.set("/body/p[1]/r[1]", ("italic=true",)).to_bytes(),
            file.set("/body/p[1]/r[1]", {"italic": True}).to_bytes(),
        )

    def test_dotted_keys(self):
        file = docx_with("a").set("/body/p[1]/r[1]", {"font.latin": "Arial", "lang.latin": "en-US"})
        run = file.get("/body/p[1]/r[1]", 0)["format"]
        self.assertEqual((run["font.latin"], run["lang.latin"]), ("Arial", "en-US"))

    def test_property_errors(self):
        file = docx_with("a")
        cases = (
            ({"bolld": True}, "unsupported_property", "did you mean 'bold'?"),
            ({"bold": "maybe"}, "invalid_value", ...),
            ({"bold": [1]}, "invalid_value", ...),
            ({"size": 2**200}, "invalid_value", ...),
            ({1: True}, "invalid_argument", ...),
            (["bold"], "invalid_argument", "write each property as key=value"),
            ([1], "invalid_argument", ...),
            (5, "invalid_argument", ...),
        )
        for props, code, suggestion in cases:
            with self.subTest(props=props):
                error = raised(lambda: file.set("/body/p[1]/r[1]", props))
                self.assertError(error, L.PropertyError, code, suggestion)
                self.assertIsInstance(error, ValueError)

    def test_typed_values(self):
        cell = L.create("xlsx")
        for value, text in ((7, "7"), (2.5, "2.5"), ("x", "x"), (True, "TRUE")):
            with self.subTest(value=value):
                self.assertEqual(cell_text(cell.set("/sheet[1]/A1", {"value": value}), "/sheet[1]/A1"), text)


# ---------------------------------------------------------------------------
# Patches (record 0347)
# ---------------------------------------------------------------------------


class PatchTests(WriteAssertions):
    """``Commit.patch``: ``apply``, ``inverse`` and ``is_noop``, as Rust's
    ``transaction::Patch``."""

    def assertRoundTrip(self, source, commit):
        """The patch replays onto ``source``, its inverse restores it, and
        a separately opened equal-content source is also accepted."""
        before, after = source.to_bytes(), commit.file.to_bytes()
        patch = commit.patch
        self.assertIsInstance(patch, L.Patch)
        self.assertFalse(patch.is_noop())
        self.assertEqual(patch.apply(source).to_bytes(), after)
        self.assertEqual(patch.inverse().apply(commit.file).to_bytes(), before)
        self.assertEqual(patch.inverse().inverse().apply(source).to_bytes(), after)
        reopened = L.open_bytes(before)
        try:
            self.assertEqual(patch.apply(reopened).to_bytes(), after)
        finally:
            reopened.close()
        for stale, target in ((patch, commit.file), (patch.inverse(), source)):
            error = raised(lambda stale=stale, target=target: stale.apply(target))
            self.assertIsInstance(error, L.LitchiError)
            self.assertEqual(error.code, "invalid_input")
        self.assertUnchanged(source, before)
        self.assertUnchanged(commit.file, after)

    def test_an_edit_patch_spans_every_run(self):
        # `add` and the reads close a run: this edit runs three Rust
        # transactions, and its one Patch replays and undoes all three.
        file = docx_with("A")
        with file.edit() as tx:
            path = tx.add("/body", "paragraph", {"text": "B"})
            tx.set(path, {"align": "center"})
            self.assertEqual(tx.get("/body/p[2]", 0)["text"], "B")
            tx.swap("/body/p[1]", "/body/p[2]")
        commit = tx.commit()
        self.assertEqual(texts(commit.file, "/body")[:2], ["B", "A"])
        self.assertRoundTrip(file, commit)
        self.assertRegex(repr(commit.patch), r"<Patch: noop=False>")

    def test_closure_batch_and_apply_patches(self):
        file = docx_with("A", "B")
        self.assertRoundTrip(file, file.edit(lambda tx: tx.set("/body/p[1]", {"align": "right"})))
        self.assertRoundTrip(
            file,
            file.batch([
                {"command": "set", "path": "/body/p[1]/r[1]", "props": {"bold": True}},
                {"command": "remove", "path": "/body/p[2]"},
            ]),
        )
        self.assertRoundTrip(
            file, file.apply({"command": "set", "path": "/body/p[2]/r[1]", "props": {"italic": True}})
        )

    def test_a_noop_patch(self):
        file = docx_with("A")
        commit = file.edit(lambda tx: None)
        self.assertFalse(commit.changed)
        self.assertTrue(commit.patch.is_noop())
        self.assertTrue(commit.patch.inverse().is_noop())
        self.assertEqual(commit.patch.apply(file).to_bytes(), file.to_bytes())

    def test_a_closed_file_refuses_a_patch(self):
        file = docx_with("A")
        commit = file.edit(lambda tx: tx.set("/body/p[1]", {"align": "right"}))
        source = L.open_bytes(file.to_bytes())
        source.close()
        error = raised(lambda: commit.patch.apply(source))
        self.assertIsInstance(error, L.LitchiError)

    def test_a_numbers_swap_patch(self):
        # Record 0347: a Numbers sheet swap, and its inverse.
        workbook = L.create("numbers")
        workbook, _ = workbook.add("/", "sheet", {"name": "Second"})
        commit = workbook.edit(lambda tx: tx.swap("/sheet[1]", "/sheet[2]"))
        self.assertEqual(commit.file.get("/sheet[1]", 0)["format"]["name"], "Second")
        self.assertRoundTrip(workbook, commit)
        with self.assertRaises(L.UnsupportedError):
            workbook.swap("/sheet[1]/table[1]/row[1]", "/sheet[1]/table[1]/row[2]")


# ---------------------------------------------------------------------------
# Batches
# ---------------------------------------------------------------------------


def set_items(count, column="A"):
    return [
        {"command": "set", "path": f"/sheet[1]/{column}{row + 1}", "props": {"value": row}}
        for row in range(count)
    ]


class BatchTests(WriteAssertions):
    def test_results_and_commit(self):
        file = docx_with("A")
        commit = file.batch(
            [
                {"command": "add", "parent": "/body", "type": "paragraph", "props": {"text": "B"}},
                {"op": "set", "path": "/body/p[2]/r[1]", "props": {"italic": True}},
                {"Command": "SET", "Path": "/body/p[1]", "Props": ["align=center"]},
                {"command": "meta", "dumpVersion": 2},
            ]
        )
        self.assertIsInstance(commit, L.Commit)
        self.assertTrue(commit.changed)
        self.assertEqual(commit.operations, 3)
        self.assertIsNone(commit.value)
        self.assertEqual(
            commit.results,
            [
                {"index": 0, "success": True, "changed": True, "path": "/body/p[2]"},
                {"index": 1, "success": True, "changed": True},
                {"index": 2, "success": True, "changed": True},
                {"index": 3, "success": True, "changed": False},
            ],
        )
        self.assertEqual(texts(commit.file, "/body")[:2], ["A", "B"])
        self.assertEqual(repr(commit), "<Commit: changed=True, operations=3>")
        fields = typed_dict_fields()["_ItemResult"]
        for result in commit.results:
            self.assertLessEqual(set(result), set(fields))

    def test_one_batch_equals_the_same_sets_one_by_one(self):
        file = L.create("xlsx")
        one_by_one = file
        for item in set_items(20):
            one_by_one = one_by_one.set(item["path"], item["props"])
        batch = file.batch(set_items(20)).file
        self.assertEqual(batch.to_bytes(), one_by_one.to_bytes())

    def test_json_text_and_bytes_equal_dict_items(self):
        file = L.create("xlsx")
        expected = file.batch(set_items(10)).file.to_bytes()
        text = json.dumps(set_items(10))
        self.assertEqual(file.batch(text).file.to_bytes(), expected)
        self.assertEqual(file.batch(text.encode()).file.to_bytes(), expected)
        self.assertEqual(file.batch(tuple(set_items(10))).file.to_bytes(), expected)

    def test_atomic_failure_changes_nothing(self):
        file = L.create("xlsx")
        before = file.to_bytes()
        items = set_items(3) + [{"command": "set", "path": "/sheet[9]/A1", "props": {"value": 1}}]
        error = raised(lambda: file.batch(items))
        self.assertError(error, L.PathError, "path_not_found")
        self.assertEqual(error.item, 3)
        self.assertUnchanged(file, before)

    def test_best_effort_reports_failed_items(self):
        file = L.create("xlsx")
        items = [set_items(1)[0], {"command": "set", "path": "/sheet[9]/A1", "props": {"value": 1}}]
        commit = file.batch(items, mode="best_effort")
        self.assertEqual(commit.results[0], {"index": 0, "success": True, "changed": True})
        failed = commit.results[1]
        self.assertEqual((failed["index"], failed["success"], failed["code"]), (1, False, "path_not_found"))
        self.assertIn("error", failed)
        self.assertIn("suggestion", failed)
        self.assertEqual(cell_text(commit.file, "/sheet[1]/A1"), "0")
        self.assertEqual(file.batch(items, mode="BestEffort").results, commit.results)

    def test_malformed_items(self):
        file = L.create("xlsx")
        cases = (
            [{"command": "set", "path": "/sheet[1]/A1", "PATH": "/x"}],  # repeated
            [{"command": "set", "op": "set"}],  # command and its alias
            [{"command": "set", "bogus": 1}],  # unknown key
            [{"command": "set", "path": 5}],  # wrong type
            [{"command": "add", "parent": "/", "type": "sheet", "index": 2**40}],  # not i32
            [{"command": "set", "path": "/sheet[1]/A1", "props": {"value": object()}}],
            [{"path": "/sheet[1]/A1"}],  # no command
            [1],
            "[{",
            b"{}",
        )
        for items in cases:
            with self.subTest(items=items if not isinstance(items, list) else repr(items)[:60]):
                error = raised(lambda: file.batch(items))
                self.assertError(error, L.PropertyError, "corrupt_batch")
        error = raised(lambda: file.batch(set_items(2) + [{"command": "set", "bogus": 1}]))
        self.assertEqual(error.item, 2)

    def test_argument_and_command_errors(self):
        file = L.create("xlsx")
        self.assertError(raised(lambda: file.batch(set_items(1), mode="eager")), L.PropertyError, "invalid_argument")
        self.assertError(raised(lambda: file.batch({"command": "set"})), L.PropertyError, "invalid_argument")
        self.assertError(raised(lambda: file.batch(5)), L.PropertyError, "invalid_argument")
        self.assertError(raised(lambda: file.apply({"command": "frobnicate"})), L.UnsupportedError, "unsupported_command")
        self.assertError(raised(lambda: file.apply({"command": "set"})), L.PropertyError, "missing_argument")
        error = raised(lambda: file.apply({"command": "get", "path": "/"}))
        self.assertError(error, L.UnsupportedError, "unsupported_capability")

    def test_operation_limit(self):
        error = raised(lambda: L.create("xlsx").batch(set_items(10_001)))
        self.assertError(error, L.LitchiError, "limit_exceeded")

    def test_apply(self):
        file = L.create("xlsx")
        commit = file.apply({"command": "set", "path": "/sheet[1]/A1", "props": {"value": 3}})
        self.assertEqual(commit.results, [{"index": 0, "success": True, "changed": True}])
        self.assertEqual(cell_text(commit.file, "/sheet[1]/A1"), "3")
        added = file.apply({"command": "add", "parent": "/", "type": "sheet", "props": {"name": "B"}})
        self.assertEqual(added.results[0]["path"], "/sheet[2]")
        error = raised(lambda: file.apply({"command": "set", "path": "/sheet[1]/A1", "bogus": 1}))
        self.assertError(error, L.PropertyError, "corrupt_batch")
        self.assertEqual(error.item, 0)

    def test_find_and_replace_report_the_match_count(self):
        # Record 0273: ``find``/``replace`` props on ``set``.
        file = L.create("xlsx")
        file = file.set("/sheet[1]/A1", {"value": "draft one, draft two"})
        commit = file.apply({"command": "set", "path": "/", "props": {"find": "draft", "replace": "final"}})
        self.assertEqual(commit.results, [{"index": 0, "success": True, "changed": True, "matched": 2}])
        self.assertEqual(cell_text(commit.file, "/sheet[1]/A1"), "final one, final two")
        none = file.apply({"command": "set", "path": "/", "props": {"find": "absent", "replace": "x"}})
        self.assertEqual(none.results[0]["matched"], 0)
        self.assertFalse(none.changed)
        self.assertEqual(none.file.to_bytes(), file.to_bytes())
        self.assertEqual(cell_text(file.set("/", {"find": "draft", "replace": "x"}), "/sheet[1]/A1"), "x one, x two")
        self.assertError(raised(lambda: file.set("/", {"find": "draft"})), L.PropertyError, "missing_property")
        self.assertError(
            raised(lambda: file.set("/", {"find": "d", "replace": "x", "regex": True})),
            L.UnsupportedError,
            "unsupported_capability",
        )

    def test_every_writable_format_batches(self):
        cases = (
            (L.create("docx"), [{"command": "add", "parent": "/body", "type": "paragraph", "props": {"text": "x"}}]),
            (L.create("xlsx"), set_items(3)),
            (L.create("xlsb"), set_items(3)),
            (L.create("xls"), set_items(3)),
            (L.create("doc"), [{"command": "set", "path": "/body/p[1]", "props": {"text": "x"}}]),
            (L.create("rtf"), [{"command": "add", "parent": "/body", "type": "paragraph", "props": {"text": "x"}}]),
            (L.open(fixture(EDITABLE_PPT)), [{"command": "set", "path": "/slide[1]/shape[3]", "props": {"text": "x"}}]),
            (L.open(fixture(SLIDESHOW_PPT)), [{"command": "swap", "path": "/slide[1]", "path2": "/slide[2]"}]),
        )
        for file, items in cases:
            with self.subTest(format=str(file.format)):
                commit = file.batch(items)
                self.assertTrue(commit.changed)
                self.assertTrue(all(result["success"] for result in commit.results))
                self.assertReopens(commit.file)


# ---------------------------------------------------------------------------
# Transactions
# ---------------------------------------------------------------------------


class EditTests(WriteAssertions):
    def test_context_manager_commits(self):
        file = docx_with("A")
        before = file.to_bytes()
        with file.edit() as tx:
            self.assertIsInstance(tx, L.Edit)
            path = tx.add("/body", "paragraph", {"text": "B"})
            tx.set(path, {"align": "right"})
            tx.set("/body/p[1]/r[1]", {"bold": True})
            self.assertEqual(repr(tx), "<Edit: 2 staged>")
            # Reads see every operation before them.
            self.assertEqual(tx.get(path, 0)["format"]["align"], "right")
            self.assertEqual(tx.find("/body/p[9]"), None)
            self.assertEqual(len(tx.query("paragraph")), 2)
        commit = tx.commit()
        self.assertIs(tx.commit(), commit)
        self.assertEqual(repr(tx), "<Edit: committed>")
        self.assertTrue(commit.changed)
        self.assertEqual(commit.operations, 3)
        self.assertIsNone(commit.results)
        self.assertEqual(texts(commit.file, "/body")[:2], ["A", "B"])
        self.assertTrue(commit.file.get("/body/p[1]/r[1]", 0)["format"]["bold"])
        self.assertUnchanged(file, before)
        error = raised(lambda: tx.set("/body/p[1]", {"text": "late"}))
        self.assertError(error, L.PropertyError, "invalid_argument")

    def test_an_exception_discards(self):
        file = docx_with("A")
        with self.assertRaises(KeyError):
            with file.edit() as tx:
                tx.set("/body/p[1]", {"text": "changed"})
                raise KeyError("inside")
        self.assertEqual(repr(tx), "<Edit: discarded>")
        self.assertError(raised(tx.commit), L.PropertyError, "invalid_argument")
        self.assertEqual(file.text().strip(), "A")

    def test_callable_form(self):
        file = docx_with("A")
        commit = file.edit(lambda tx: tx.add("/body", "paragraph", {"text": "B"}))
        self.assertEqual(commit.value, "/body/p[2]")
        self.assertEqual(texts(commit.file, "/body")[:2], ["A", "B"])
        nothing = file.edit(lambda tx: None)
        self.assertFalse(nothing.changed)
        self.assertEqual(nothing.file.to_bytes(), file.to_bytes())

        def failing(tx):
            tx.set("/body/p[1]", {"text": "changed"})
            raise ValueError("inside")

        with self.assertRaises(ValueError):
            file.edit(failing)
        self.assertEqual(file.text().strip(), "A")

    def test_a_staged_failure_names_its_operation_and_poisons(self):
        file = L.create("xlsx")
        tx = file.edit()
        tx.set("/sheet[1]/A1", {"value": 1})
        tx.set("/sheet[1]/A2", {"bolld": True})
        tx.set("/sheet[1]/A3", {"value": 3})
        error = raised(tx.commit)
        self.assertError(error, L.PropertyError, "unsupported_property")
        self.assertEqual(error.item, 1)
        again = raised(lambda: tx.set("/sheet[1]/A4", {"value": 4}))
        self.assertEqual((type(again), again.code, again.item), (type(error), error.code, 1))
        self.assertEqual(repr(tx), "<Edit: failed>")
        # A failure in a later run counts the operations of earlier runs.
        tx = file.edit()
        tx.set("/sheet[1]/A1", {"value": 1})
        tx.get("/sheet[1]/A1")
        tx.set("/sheet[9]/A1", {"value": 1})
        error = raised(tx.commit)
        self.assertEqual((error.code, error.item), ("path_not_found", 1))

    def test_structural_operations_and_placement(self):
        file = docx_with("A", "B", "C")
        with file.edit() as tx:
            tx.swap("/body/p[1]", "/body/p[3]")
            tx.move("/body/p[1]", index=1)
            tx.remove("/body/p[3]")
        self.assertEqual(texts(tx.commit().file, "/body")[:2], ["B", "C"])
        tx = file.edit()
        self.assertError(raised(lambda: tx.move("/body/p[1]", index=0, after="/body/p[2]")), L.PropertyError, "invalid_argument")
        self.assertError(raised(lambda: tx.add("/body", "paragraph", index=-1)), L.PropertyError, "invalid_argument")
        self.assertError(raised(lambda: tx.set("/body/p[1]", 5)), L.PropertyError, "invalid_argument")
        # Argument errors are raised before staging: the transaction is open.
        self.assertEqual(repr(tx), "<Edit: 0 staged>")

    def test_many_sets_coalesce_into_one_commit(self):
        file = L.create("xlsx")

        def fill(tx):
            for item in set_items(500):
                tx.set(item["path"], item["props"])

        commit = file.edit(fill)
        self.assertEqual(commit.operations, 500)
        self.assertEqual(commit.file.to_bytes(), file.batch(set_items(500)).file.to_bytes())

    def test_protected_files_refuse(self):
        signed = L.open(fixture(SIGNED_DOCX))
        before = signed.to_bytes()
        noop = signed.edit(lambda tx: None)
        self.assertFalse(noop.changed)
        self.assertEqual(noop.operations, 0)
        self.assertTrue(noop.patch.is_noop())
        self.assertEqual(noop.file.to_bytes(), before)
        self.assertTrue(noop.file.signed)
        replayed = noop.patch.apply(signed)
        self.assertTrue(replayed.signed)
        self.assertEqual(replayed.to_bytes(), before)

        error = raised(
            lambda: signed.edit(lambda tx: tx.set("/body/p[1]", {"text": "changed"}))
        )
        self.assertError(error, L.ProtectedError, "document_protected", SIGNED_SUGGESTION)
        for relative in LOCKED:
            with self.subTest(fixture=relative):
                error = raised(L.open(fixture(relative)).edit().commit)
                self.assertError(error, L.ProtectedError, "document_protected", PROTECTED_SUGGESTION)


# ---------------------------------------------------------------------------
# Publication and protection
# ---------------------------------------------------------------------------


class SaveTests(WriteAssertions):
    def test_save_as_and_save_round_trip_every_format(self):
        builders = {
            "docx": lambda: docx_with("saved"),
            "xlsx": lambda: L.create("xlsx").set("/sheet[1]/A1", {"value": "saved"}),
            "pptx": lambda: L.create("pptx"),
            "xlsb": lambda: L.create("xlsb").set("/sheet[1]/A1", {"value": "saved"}),
            "doc": lambda: L.create("doc").set("/body/p[1]", {"text": "saved"}),
            "xls": lambda: L.create("xls").set("/sheet[1]/A1", {"value": "saved"}),
            "ppt": lambda: L.create("ppt"),
            "rtf": lambda: L.create("rtf").add("/body", "paragraph", {"text": "saved"})[0],
        }
        with tempfile.TemporaryDirectory() as directory:
            for name, build in builders.items():
                with self.subTest(format=name):
                    file = build()
                    path = os.path.join(directory, f"out.{name}")
                    written = file.save_as(path)
                    self.assertEqual(written, os.path.getsize(path))
                    with open(path, "rb") as handle:
                        self.assertEqual(handle.read(), file.to_bytes())
                    opened = L.open(path)
                    self.assertEqual(opened.text(), file.text())
                    # save() writes back to the path the File was opened from.
                    self.assertEqual(opened.save(), os.path.getsize(path))
                    self.assertEqual(L.open(path).text(), file.text())

    def test_save_to_the_origin_and_overwrite(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "doc.docx")
            docx_with("one").save_as(path)
            opened = L.open(path)
            changed, _ = opened.add("/body", "paragraph", {"text": "two"})
            changed.save()
            self.assertEqual(texts(L.open(path), "/body")[:2], ["one", "two"])
            self.assertEqual(changed.save(path), os.path.getsize(path))
            error = raised(lambda: changed.save_as(path, overwrite=False))
            self.assertError(error, L.IoError, "file_exists")
            error = raised(lambda: changed.save(os.path.join(directory, "doc.xlsx")))
            self.assertError(error, L.PropertyError, "invalid_argument")
            self.assertIsInstance(changed.to_bytes(), bytes)

    def test_locked_files_refuse_every_write(self):
        for relative in LOCKED:
            file = L.open(fixture(relative))
            self.assertTrue(file.locked)
            with tempfile.TemporaryDirectory() as directory:
                for call in (
                    lambda: file.set("/body/p[1]", {"text": "x"}),
                    lambda: file.add("/", "sheet"),
                    lambda: file.remove("/sheet[1]"),
                    lambda: file.batch([]),
                    file.to_bytes,
                    lambda: file.save_as(os.path.join(directory, "out" + os.path.splitext(relative)[1])),
                ):
                    with self.subTest(fixture=relative):
                        error = raised(call)
                        self.assertError(error, L.ProtectedError, "document_protected", PROTECTED_SUGGESTION)

    def test_signed_files_keep_their_signature_until_unsigned(self):
        for relative, path, props in (
            (SIGNED_DOCX, "/body/p[1]", {"text": "changed"}),
            (SIGNED_XLSX, "/sheet[1]/A1", {"value": "changed"}),
        ):
            with self.subTest(fixture=relative):
                file = L.open(fixture(relative))
                self.assertTrue(file.signed)
                for call in (
                    lambda: file.set(path, props),
                    lambda: file.batch([{"command": "set", "path": path, "props": props}]),
                ):
                    error = raised(call)
                    self.assertError(error, L.ProtectedError, "document_protected", SIGNED_SUGGESTION)
                # An unchanged signed file publishes with its signature.
                self.assertTrue(L.open_bytes(file.to_bytes()).signed)
                unsigned = file.unsign()
                self.assertFalse(unsigned.signed)
                self.assertTrue(file.signed)
                self.assertIs(unsigned.unsign().signed, False)
                changed = unsigned.set(path, props)
                reopened = L.open_bytes(changed.to_bytes())
                self.assertFalse(reopened.signed)
                self.assertEqual(reopened.get(path, 0)["text"], "changed")


class ErrorClassTests(WriteAssertions):
    def test_item_attribute_pickles(self):
        error = raised(lambda: L.create("xlsx").batch(set_items(2) + [{"command": "set", "path": "/sheet[9]/A1", "props": {"value": 1}}]))
        clone = pickle.loads(pickle.dumps(error))
        self.assertEqual((type(clone), clone.code, clone.item, clone.message), (type(error), error.code, 2, error.message))
        self.assertIsNone(L.LitchiError("x").item)
        self.assertEqual(L.PathError("x", item=4).item, 4)

    def test_closed_files_refuse_writes(self):
        file = docx_with("A")
        file.close()
        for call in (
            lambda: file.set("/body/p[1]", {"text": "x"}),
            lambda: file.add("/body", "paragraph"),
            lambda: file.remove("/body/p[1]"),
            lambda: file.move("/body/p[1]"),
            lambda: file.swap("/body/p[1]", "/body/p[1]"),
            lambda: file.edit(),
            lambda: file.batch([]),
            lambda: file.apply({"command": "meta", "dumpVersion": 2}),
            file.save,
            file.to_bytes,
            file.unsign,
        ):
            error = raised(call)
            self.assertError(error, L.IoError, "file_required", "open it again with litchi_py.open")
        self.assertFalse(file.signed or file.locked)


class MetadataKeywordsTests(WriteAssertions):
    """The keywords are a list (record 0354): ``metadata()`` and the root
    node's ``format`` give a list of str; a set takes a list, or text split
    at ";" and ","."""

    FORMATS = ("docx", "xlsx", "pptx", "xlsb", "doc", "xls", "ppt", "odt", "ods", "odp")

    def test_list_and_text_round_trip_in_every_format(self):
        for name in self.FORMATS:
            with self.subTest(format=name):
                source = L.create(name)
                listed = source.set("/metadata", {"keywords": ["q4", "review", "a b"]})
                reopened = self.assertReopens(listed)
                self.assertEqual(reopened.metadata()["keywords"], ["q4", "review", "a b"])
                texted = listed.set("/metadata", {"keywords": "x; y,z"})
                self.assertEqual(
                    L.open_bytes(texted.to_bytes()).metadata()["keywords"], ["x", "y", "z"]
                )
                cleared = texted.set("/metadata", {"keywords": None})
                self.assertNotIn("keywords", L.open_bytes(cleared.to_bytes()).metadata() or {})

    def test_root_node_format_is_a_list(self):
        for name in ("docx", "xlsx", "pptx"):
            with self.subTest(format=name):
                file = L.create(name).set("/metadata", {"keywords": ("plan", "review")})
                self.assertEqual(file.get("/", 0)["format"]["keywords"], ["plan", "review"])

    def test_a_list_holds_only_str(self):
        file = L.create("docx")
        self.assertError(
            raised(lambda: file.set("/metadata", {"keywords": ["a", 1]})),
            L.PropertyError,
            "invalid_value",
        )


if __name__ == "__main__":
    unittest.main()
