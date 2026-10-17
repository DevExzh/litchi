"""The unified read API: ``litchi_py.open``/``open_bytes`` and ``File``
(M1 step R9, record 0010 §8.2).

Oracles, strongest first:

* **Rust goldens.** A node view dict must equal the Rust node view's JSON:
  the golden files of ``crates/litchi-ops/tests/golden/node_view/`` for
  real fixtures. The comparison is made after ``json.dumps``, so member
  order counts too. ``capabilities()`` must equal
  ``crates/litchi-ops/tests/golden/capabilities/<format>.json``, apart from
  ``retainsSource``, which only a file's report carries.
* **Round trips.** Every node a view or query lists re-resolves, through its
  ``path``, to an equal view.
* **Facts of real fixtures** under ``test-data/``: paths, counts, texts.

DOCX, XLSX, PPTX, XLSB, DOC, XLS, PPT and RTF have unified readers; ODT,
ODS and ODP refuse with ``unsupported_capability`` naming the follow-up
(``test_reads_agree_with_the_report_for_every_format``). The write verbs are
tested in test_write_api.py; the cross-thread and GIL tests are in
test_threads.py.
"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
import unittest
import warnings
from pathlib import Path

from _support import (
    STUB_PATH,
    TEST_DATA,
    L,
    copy_fixture,
    fixture,
    golden,
    modify_in_place,
    read_fixture,
    typed_dict_fields,
)

HYPERLINK = "ooxml/docx/Hyperlink.docx"
DRAWING = "ooxml/docx/drawing.docx"
SIMPLE = "ooxml/xlsx/SimpleNormal.xlsx"
TABLES = "ooxml/xlsx/ExcelTables.xlsx"
PPTX = "ooxml/pptx/sample.pptx"
SIGNED_DOCX = "poi/test-data/xmldsign/hello-world-signed.docx"
SIGNED_XLSX = "poi/test-data/xmldsign/hello-world-signed.xlsx"
LOCKED = (("ole/doc/PasswordProtected.doc", "Doc"), ("ole/xls/password.xls", "Xls"))

#: The node view members, in OfficeCLI's order; the second set may be absent.
MEMBERS = ("path", "type", "text", "preview", "style", "childCount", "format", "children")
OPTIONAL_MEMBERS = {"text", "preview", "style", "childCount"}

#: One fixture per format ``litchi::File`` holds.
FORMATS = (
    (HYPERLINK, "Docx"),
    (SIMPLE, "Xlsx"),
    (PPTX, "Pptx"),
    ("ooxml/xlsb/sample.xlsb", "Xlsb"),
    ("ole/doc/NoHeadFoot.doc", "Doc"),
    ("ole/xls/Simple.xls", "Xls"),
    ("ole/ppt/SampleShow.ppt", "Ppt"),
    ("rtf/testStyles.rtf", "Rtf"),
    ("odf/corpus/writer-table.odt", "Odt"),
    ("odf/corpus/calc-two-sheets.ods", "Ods"),
    ("odf/odp/tdf169979.odp", "Odp"),
)

#: Python suggestion of a document_protected error (the bindings cannot unlock).
PROTECTED_SUGGESTION = (
    "remove the password in the application that created the file; "
    "litchi_py cannot unlock encrypted documents yet"
)


def open_fixture(relative: str):
    return L.open(fixture(relative))


def raised(call):
    """The exception ``call`` raises; fails the test when it returns."""
    try:
        result = call()
    except Exception as error:  # noqa: BLE001 - the class is what is checked
        return error
    raise AssertionError(f"expected an exception, got {result!r}")


def walk(view):
    """``view`` and every view below it, depth first."""
    yield view
    for child in view["children"]:
        yield from walk(child)


class ErrorAssertions(unittest.TestCase):
    def assertError(self, error, cls, code, suggestion=...):
        self.assertIs(type(error), cls, repr(error))
        self.assertEqual(error.code, code)
        if suggestion is not ...:
            self.assertEqual(error.suggestion, suggestion)


# ---------------------------------------------------------------------------
# Opening
# ---------------------------------------------------------------------------


class OpenTests(ErrorAssertions):
    def test_open_accepts_str_pathlib_and_keyword(self):
        path = fixture(HYPERLINK)
        for file in (L.open(path), L.open(Path(path)), L.open(path=path)):
            self.assertIsInstance(file, L.File)
            self.assertEqual(str(file.format), "Docx")
            self.assertEqual(file.get("/body/p[1]")["text"], "Following is a hyperlink.")

    def test_open_bytes_reads_like_open(self):
        data = read_fixture(SIMPLE)
        for file in (L.open_bytes(data), L.open_bytes(data=data)):
            self.assertEqual(str(file.format), "Xlsx")
            self.assertEqual(file.get("/Sheet1", 2), open_fixture(SIMPLE).get("/Sheet1", 2))
            self.assertEqual(file.view("text"), open_fixture(SIMPLE).view("text"))

    def test_open_bytes_takes_bytes_only(self):
        for argument in ("text", None, 42):
            with self.subTest(argument=argument), self.assertRaises(TypeError):
                L.open_bytes(argument)

    def test_every_format_file_holds_opens(self):
        for relative, name in FORMATS:
            with self.subTest(fixture=relative):
                file = open_fixture(relative)
                self.assertEqual(str(file.format), name)
                self.assertEqual(repr(file), f"<File: {name}>")
                self.assertIsInstance(file.text(), str)

    def test_unreadable_and_foreign_inputs(self):
        with tempfile.TemporaryDirectory() as directory:
            missing = os.path.join(directory, "missing.docx")
            self.assertError(raised(lambda: L.open(missing)), L.IoError, "file_not_found")
            self.assertError(raised(lambda: L.open(directory)), L.IoError, "io_error")
            plain = os.path.join(directory, "notes.txt")
            with open(plain, "wb") as handle:
                handle.write(b"plain text, not an office file\n" * 8)
            truncated = os.path.join(directory, "truncated.docx")
            docx = read_fixture(HYPERLINK)
            with open(truncated, "wb") as handle:
                handle.write(docx[: len(docx) // 2])
            for call in (lambda: L.open(plain), lambda: L.open_bytes(b"plain text")):
                error = raised(call)
                self.assertError(error, L.UnsupportedError, "unsupported_format", None)
                self.assertEqual(error.message, "not a recognized Office file")
            self.assertError(raised(lambda: L.open(truncated)), L.CorruptFileError, "corrupt_file")
        self.assertError(
            raised(lambda: L.open(fixture("odf/flat/minimal-formula.fodf"))),
            L.UnsupportedError,
            "unsupported_capability",
        )

    def test_every_odf_format_opens_as_a_file(self):
        # Record 0282: the remaining ODF formats have File variants.
        for relative, name in (
            ("libreoffice-core/desktop/qa/data/BlankDrawDocument.odg", "Odg"),
            ("odf/odc-producer-evidence/odfdom-created.odc", "Odc"),
            ("odf/native-resave/source/font-styles.odf", "Odf"),
            ("libreoffice-core/sw/qa/extras/odfexport/data/tdf121119.odm", "Odm"),
            ("libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb", "Odb"),
        ):
            with self.subTest(fixture=relative):
                file = open_fixture(relative)
                self.assertEqual(str(file.format), name)
                self.assertEqual(file.get("/", 0)["path"], "/")

    def test_iwork_packages_open_as_files(self):
        # litchi::File holds Apple iWork since record 0296 (owner decision D1).
        for relative, name, path, text in (
            ("iwork/pages/basic.pages", "Pages", "/body/p[1]", "Litchi native Pages fixture"),
            (
                "iwork/keynote/basic.key",
                "Keynote",
                "/slide[1]/placeholder[title]",
                "Litchi native Keynote fixture",
            ),
            (
                "iwork/numbers/basic.numbers",
                "Numbers",
                "/sheet[1]/table[1]/B2",
                "Litchi native Numbers fixture",
            ),
        ):
            with self.subTest(fixture=relative):
                file = open_fixture(relative)
                self.assertEqual(str(file.format), name)
                self.assertEqual(file.get(path, 0)["text"], text)
                self.assertEqual(file.to_bytes(), read_fixture(relative))

    def test_iwork_package_directories_open_and_save_back_in_place(self):
        # Records 0299, 0301 and 0306: every iWork leaf converts the
        # app-authored directory to its single-file form; the directory is
        # the origin, so save() writes the bundle back.
        for relative, single, name, path in (
            ("iwork/directory/pages/basic.pages", "iwork/pages/basic.pages", "Pages", "/body/p[1]"),
            ("iwork/directory/keynote/basic.key", "iwork/keynote/basic.key", "Keynote",
             "/slide[1]/placeholder[title]"),
            ("iwork/directory/numbers/basic.numbers", "iwork/numbers/basic.numbers", "Numbers",
             "/sheet[1]/table[1]/B2"),
        ):
            with self.subTest(fixture=relative):
                directory = TEST_DATA / relative
                self.assertTrue(directory.is_dir(), directory)
                file = L.open(str(directory))
                self.assertEqual(str(file.format), name)
                self.assertEqual(file.get(path, 0)["text"], open_fixture(single).get(path, 0)["text"])
                with tempfile.TemporaryDirectory() as temp:
                    copy = os.path.join(temp, os.path.basename(relative))
                    shutil.copytree(directory, copy)
                    key = "value" if name == "Numbers" else "text"
                    edited = L.open(copy).set(path, {key: "Saved in place"})
                    edited.save()
                    self.assertTrue(os.path.isfile(os.path.join(copy, "Index.zip")))
                    self.assertEqual(L.open(copy).get(path, 0)["text"], "Saved in place")

    def test_open_takes_a_path(self):
        for argument in (None, 42, [fixture(HYPERLINK)]):
            with self.subTest(argument=argument), self.assertRaises(TypeError):
                L.open(argument)


# ---------------------------------------------------------------------------
# Node views: get and find
# ---------------------------------------------------------------------------


class NodeViewGoldenTests(unittest.TestCase):
    """Dicts built from the node view's accessors equal the Rust JSON."""

    def assertSameView(self, actual, expected):
        self.assertEqual(actual, expected)
        # Member order and JSON types too (True is not 1, 1 is not 1.0).
        self.assertEqual(json.dumps(actual), json.dumps(expected))

    def test_docx_fixtures(self):
        cases = (
            (DRAWING, "/body/p[10]", "docx/drawing_paragraph.json"),
            (DRAWING, "/body/tbl[1]", "docx/drawing_table.json"),
        )
        for relative, path, expected in cases:
            with self.subTest(fixture=relative, path=path):
                file = open_fixture(relative)
                self.assertSameView(file.get(path), golden(f"node_view/{expected}"))
                self.assertSameView(file.find(path), golden(f"node_view/{expected}"))

    def test_docx_stories_are_read(self):
        # Headers and notes were listed gaps (L-DOCX-R5) until record 0319.
        cases = (
            ("ooxml/docx/headerFooter.docx", "/header[@type=default]", "header", "/header[2]"),
            ("ooxml/docx/footnotes.docx", "/footnote[1]", "footnote", "/footnotes/footnote[1]"),
        )
        for relative, path, element, canonical in cases:
            with self.subTest(path=path):
                file = open_fixture(relative)
                node = file.get(path)
                self.assertEqual(node["type"], element)
                self.assertEqual(node["path"], canonical)
                self.assertTrue(node["text"])
                elements = {entry["element"]: entry for entry in file.capabilities()["elements"]}
                operations = elements[element]["operations"]
                for verb in ("get", "set", "add", "remove"):
                    self.assertTrue(operations[verb], verb)

    def test_docx_comments_are_read(self):
        file = open_fixture("ooxml/docx/comment.docx")
        comments = file.get("/comments")
        self.assertEqual(comments["type"], "comments")
        self.assertEqual(comments["childCount"], 1)
        comment = file.get("/comments/comment[1]")
        self.assertEqual(comment["format"]["author"], "Unbekannter Autor")
        # Three paragraphs, the middle one empty, one per line (record 0251c).
        self.assertEqual(comment["text"], "This is the first line\n\nThis is the second line")

    def test_xlsx_fixtures(self):
        cases = (
            (
                TABLES,
                ["/", "/ExcelTable", "/ExcelTable/table[1]", "/ExcelTable/C2", "/namedrange[2]"],
                "xlsx/fixture_excel_tables.json",
            ),
            (
                "ooxml/xlsx/table-sample.xlsx",
                ["/Tabelle1/table[1]", "/Tabelle1/G5"],
                "xlsx/fixture_table_sample.json",
            ),
        )
        for relative, paths, expected in cases:
            with self.subTest(fixture=relative):
                file = open_fixture(relative)
                self.assertSameView([file.get(path) for path in paths], golden(f"node_view/{expected}"))


class NodeViewTests(ErrorAssertions):
    def test_members_types_and_order(self):
        views = [
            open_fixture(DRAWING).get("/body", 2, max_nodes=10_000),
            open_fixture(TABLES).get("/", 2),
            open_fixture(SIMPLE).get("/Sheet Number 2", 2),
        ]
        seen = set()
        for view in (node for top in views for node in walk(top)):
            keys = list(view)
            self.assertEqual(keys, [key for key in MEMBERS if key in view], keys)
            self.assertTrue(set(MEMBERS) - OPTIONAL_MEMBERS <= set(keys), keys)
            self.assertIsInstance(view["path"], str)
            self.assertIsInstance(view["type"], str)
            for key in ("text", "preview", "style"):
                if key in view:
                    self.assertIsInstance(view[key], str)
            if "childCount" in view:
                self.assertIs(type(view["childCount"]), int)
            self.assertIsInstance(view["format"], dict)
            for key, value in view["format"].items():
                self.assertIsInstance(key, str)
                self.assertIn(type(value), (bool, int, float, str, type(None)), (key, value))
                seen.add(type(value))
            self.assertIsInstance(view["children"], list)
        self.assertTrue({bool, int, str} <= seen, seen)

    def test_stub_typed_dict_matches_the_members(self):
        fields = typed_dict_fields(STUB_PATH)["_NodeView"]
        self.assertEqual(list(fields), list(MEMBERS))
        self.assertEqual({key for key, required in fields.items() if not required}, OPTIONAL_MEMBERS)

    def test_format_values_keep_their_json_types(self):
        paragraph = open_fixture(DRAWING).get("/body/p[10]", 0)["format"]
        self.assertEqual(paragraph["spaceAfter"], "12pt")  # a length, as text
        self.assertIs(paragraph["spaceAfterAuto"], False)
        table = open_fixture(DRAWING).get("/body/tbl[1]", 0)["format"]
        self.assertIs(type(table["cellSpacing"]), int)
        excel_table = open_fixture(TABLES).get("/ExcelTable/table[1]")["format"]
        self.assertIs(excel_table["bandedRows"], True)
        self.assertEqual(open_fixture(SIMPLE).get("/", 0)["format"]["activeTab"], 1)

    def test_depth(self):
        file = open_fixture(HYPERLINK)
        flat = file.get("/body", 0)
        self.assertEqual((flat["childCount"], flat["children"]), (3, []))
        one = file.get("/body")
        self.assertEqual(one, file.get("/body", 1))
        self.assertEqual(
            [child["path"] for child in one["children"]],
            ["/body/p[1]", "/body/p[2]", "/body/section[1]"],
        )
        self.assertTrue(all(child["children"] == [] for child in one["children"]))
        two = file.get("/body", depth=2)
        self.assertEqual(two["children"][1]["children"][0]["path"], "/body/p[2]/r[1]")
        self.assertEqual(two["children"][1]["children"][0]["format"], {"rStyle": "Hyperlink"})
        # Depth is clamped at 64, as in Rust.
        self.assertEqual(file.get("/", depth=10**6), file.get("/", depth=64))

    def test_child_paths_resolve_to_equal_views(self):
        for relative, path in ((DRAWING, "/body"), (DRAWING, "/body/tbl[2]"), (SIMPLE, "/"), (TABLES, "/ExcelTable")):
            file = open_fixture(relative)
            parent = file.get(path, 2, max_nodes=10_000)
            for child in parent["children"]:
                with self.subTest(fixture=relative, path=child["path"]):
                    self.assertEqual(file.get(child["path"], 1), child)

    def test_find(self):
        file = open_fixture(HYPERLINK)
        self.assertEqual(file.find("/body/p[2]"), file.get("/body/p[2]"))
        self.assertIsNone(file.find("/body/p[99]"))
        self.assertIsNone(open_fixture(SIMPLE).find("/NoSuchSheet/A1"))
        self.assertError(raised(lambda: file.find("/body/p[")), L.PathError, "invalid_path")
        self.assertError(raised(lambda: file.find("/slide[1]")), L.PathError, "unsupported_path")

    def test_workbook_paths(self):
        file = open_fixture(SIMPLE)
        cell = file.get("/Sheet1/A1")
        self.assertEqual(
            cell,
            {
                "path": "/sheet[1]/A1",
                "type": "cell",
                "text": "test",
                "preview": "A1",
                "childCount": 0,
                "format": {"type": "SharedString"},
                "children": [],
            },
        )
        self.assertEqual(file.get("Sheet1!A1"), cell)
        self.assertEqual(file.get("/sheet[1]/A1"), cell)
        formula = file.get("/Sheet Number 2/A7")["format"]
        self.assertEqual((formula["formula"], formula["cachedValue"]), ("SUM(A3:D3)", "10"))
        # A coordinate with nothing stored is a cell, not a miss.
        self.assertEqual(file.get("/Sheet1/Z99")["text"], "(empty)")
        row = file.get("/Sheet1/row[1]")
        self.assertEqual([cell["path"] for cell in row["children"]], ["/sheet[1]/A1", "/sheet[1]/C1"])

    def test_dense_ranges(self):
        file = open_fixture(SIMPLE)
        sparse = file.get("/Sheet Number 2/A1:B3")
        self.assertEqual(sparse["childCount"], 4)
        dense = file.get("/Sheet Number 2/A1:B3", dense=6)
        self.assertEqual(
            [cell["path"].rsplit("/", 1)[1] for cell in dense["children"]],
            ["A1", "B1", "A2", "B2", "A3", "B3"],
        )
        self.assertEqual(dense["children"][1]["format"].get("empty"), True)

    def test_max_nodes_bounds_the_view(self):
        file = open_fixture(DRAWING)
        error = raised(lambda: file.get("/body", max_nodes=5))
        self.assertError(error, L.LitchiError, "output_too_large")
        self.assertIn("max_nodes", error.suggestion)
        self.assertEqual(file.get("/body/p[1]", 0, max_nodes=1)["path"], "/body/p[1]")
        self.assertError(raised(lambda: file.get("/body/p[1]", 0, max_nodes=0)), L.LitchiError, "output_too_large")

    def test_negative_arguments(self):
        file = open_fixture(HYPERLINK)
        for call in (
            lambda: file.get("/body", -1),
            lambda: file.get("/body", max_nodes=-1),
            lambda: file.find("/body", dense=-1),
        ):
            error = raised(call)
            self.assertError(error, L.PropertyError, "invalid_argument")
            self.assertIn("must not be negative", error.message)


# ---------------------------------------------------------------------------
# Query
# ---------------------------------------------------------------------------


class QueryTests(ErrorAssertions):
    def test_matches_in_document_order(self):
        self.assertEqual(
            [view["path"] for view in open_fixture(HYPERLINK).query("run")],
            ["/body/p[1]/r[1]", "/body/p[2]/r[1]"],
        )
        cells = open_fixture(SIMPLE).query("cell")
        self.assertEqual(len(cells), 19)
        self.assertEqual(
            [view["path"] for view in cells[:3]], ["/sheet[1]/A1", "/sheet[1]/C1", "/sheet[1]/A2"]
        )
        self.assertEqual(cells[-1]["path"], "/sheet[2]/C7")

    def test_limit(self):
        for relative, selector in ((DRAWING, "paragraph"), (SIMPLE, "cell")):
            file = open_fixture(relative)
            every = file.query(selector, max_nodes=10_000)
            self.assertEqual(file.query(selector, None, max_nodes=10_000), every)
            for limit in (0, 1, 2, 5, len(every), len(every) + 10):
                with self.subTest(selector=selector, limit=limit):
                    self.assertEqual(file.query(selector, limit), every[:limit])
                    self.assertEqual(file.query(selector, limit=limit), every[:limit])

    def test_depth_defaults_to_zero(self):
        file = open_fixture(SIMPLE)
        rows = file.query("row")
        self.assertTrue(rows)
        self.assertTrue(all(row["children"] == [] for row in rows))
        deep = file.query("row", depth=1)
        self.assertEqual([row["path"] for row in deep], [row["path"] for row in rows])
        self.assertEqual(
            [cell["path"] for cell in deep[0]["children"]], ["/sheet[1]/A1", "/sheet[1]/C1"]
        )

    def test_matches_equal_their_get_views(self):
        for relative, selector in ((DRAWING, "paragraph"), (DRAWING, "table"), (SIMPLE, "cell"), (TABLES, "table")):
            file = open_fixture(relative)
            matches = file.query(selector, max_nodes=10_000)
            self.assertTrue(matches, selector)
            for view in matches:
                with self.subTest(selector=selector, path=view["path"]):
                    self.assertEqual(file.get(view["path"], 0), view)

    def test_predicates(self):
        file = open_fixture(SIMPLE)
        self.assertEqual(
            [view["path"] for view in file.query("cell[formula]")],
            ["/sheet[2]/A7", "/sheet[2]/C7"],
        )
        numbers = file.query("cell[type=Number]")
        self.assertEqual(len(numbers), 15)
        self.assertTrue(all(view["format"]["type"] == "Number" for view in numbers))
        self.assertEqual(
            [view["path"] for view in file.query("Sheet1!A")], ["/sheet[1]/A1", "/sheet[1]/A2"]
        )
        styled = open_fixture(HYPERLINK).query("paragraph[style=Normal]")
        self.assertEqual([view["path"] for view in styled], ["/body/p[1]", "/body/p[2]"])

    def test_unknown_key_warns_and_matches_nothing(self):
        file = open_fixture(HYPERLINK)
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            self.assertEqual(file.query("paragraph[styl=Normal]"), [])
        self.assertEqual([warning.category for warning in caught], [UserWarning])
        message = str(caught[0].message)
        self.assertIn("unknown_key: unknown key", message)
        self.assertIn("did you mean 'style'?", message)
        # The warning points at the caller's line.
        self.assertEqual(caught[0].filename, __file__)
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            with self.assertRaises(UserWarning):
                file.query("paragraph[styl=Normal]")

    def test_errors(self):
        file = open_fixture(DRAWING)
        error = raised(lambda: file.query("p["))
        self.assertError(error, L.PathError, "invalid_selector")
        self.assertIsInstance(error, LookupError)
        self.assertError(raised(lambda: file.query("paragraph", -1)), L.PropertyError, "invalid_argument")
        self.assertError(raised(lambda: file.query("paragraph", depth=-1)), L.PropertyError, "invalid_argument")
        # One budget for the whole result: matches and their children.
        self.assertError(raised(lambda: file.query("paragraph", max_nodes=5)), L.LitchiError, "output_too_large")
        self.assertEqual(len(file.query("paragraph", 5, max_nodes=5)), 5)
        self.assertError(
            raised(lambda: file.query("paragraph", 5, depth=1, max_nodes=5)),
            L.LitchiError,
            "output_too_large",
        )


# ---------------------------------------------------------------------------
# Views
# ---------------------------------------------------------------------------


class ViewTests(ErrorAssertions):
    MODES = ("text", "annotated", "outline", "stats")

    def test_docx_text_and_annotated(self):
        file = open_fixture(HYPERLINK)
        self.assertEqual(
            file.view("text"), "[/body/p[1]] Following is a hyperlink.\n[/body/p[2]] EricWhite.com"
        )
        self.assertEqual(
            file.view("text", json=True),
            {
                "totalElements": 3,
                "elements": [
                    {"path": "/body/p[1]", "type": "paragraph", "text": "Following is a hyperlink."},
                    {"path": "/body/p[2]", "type": "paragraph", "text": "EricWhite.com"},
                ],
            },
        )
        annotated = file.view("annotated")
        self.assertEqual(
            annotated.splitlines()[0], "[/body/p[1]] 「Following is a hyperlink.」 ← Normal | (default)"
        )
        self.assertEqual(file.view("annotated", json=True), {"view": "annotated", "content": annotated})

    def test_docx_outline_and_stats(self):
        file = open_fixture(HYPERLINK)
        self.assertEqual(
            file.view("outline", json=True),
            {"paragraphs": 2, "tables": 0, "images": 0, "equations": 0, "headings": []},
        )
        stats = file.view("stats", json=True)
        self.assertEqual((stats["paragraphs"], stats["words"]), (2, 5))
        self.assertEqual(stats["styleDistribution"], {"Normal": 2})
        drawing = open_fixture(DRAWING)
        self.assertEqual(drawing.view("outline").splitlines()[0], "185 paragraphs | 5 tables | 15 images")

    def test_xlsx_views(self):
        file = open_fixture(SIMPLE)
        text = file.view("text")
        self.assertEqual(
            text.splitlines()[:3],
            ["=== Sheet: Sheet1 ===", "[/sheet[1]/row[1]] A1=test\tC1=1", "[/sheet[1]/row[2]] A2=test 2\tC2=2"],
        )
        data = file.view("text", json=True)
        self.assertEqual([sheet["name"] for sheet in data["sheets"]], ["Sheet1", "Sheet Number 2"])
        self.assertEqual(data["sheets"][0]["rows"][0], {"row": 1, "cells": {"A1": "test", "C1": "1"}})
        self.assertEqual(
            file.view("outline"),
            '├── "Sheet1" (6 rows × 3 cols)\n├── "Sheet Number 2" (5 rows × 6 cols, 2 formula(s))',
        )
        self.assertEqual(
            file.view("outline", json=True)["sheets"][1],
            {"name": "Sheet Number 2", "rows": 5, "cols": 6, "formulas": 2, "errorCells": 0, "tables": 0},
        )
        self.assertEqual(
            file.view("stats", json=True),
            {
                "sheets": 2,
                "totalCells": 19,
                "emptyCells": 0,
                "formulaCells": 2,
                "errorCells": 0,
                "dataTypeDistribution": {"Number": 15, "SharedString": 4},
            },
        )
        self.assertTrue(file.view("stats").startswith("Sheets: 2\nTotal Cells: 19\n"))
        annotated = file.view("annotated")
        self.assertIn("  A7: [10] ← =SUM(A3:D3)", annotated.splitlines())
        self.assertEqual(file.view("annotated", json=True), {"view": "annotated", "content": annotated})

    def test_every_mode_in_both_forms(self):
        for relative in (HYPERLINK, DRAWING, SIMPLE, TABLES):
            file = open_fixture(relative)
            for mode in self.MODES:
                with self.subTest(fixture=relative, mode=mode):
                    text = file.view(mode)
                    self.assertIsInstance(text, str)
                    self.assertTrue(text)
                    data = file.view(mode, json=True)
                    self.assertIsInstance(data, dict)
                    self.assertTrue(data)
                    # The JSON form is plain JSON data.
                    self.assertEqual(json.loads(json.dumps(data)), data)

    def test_mode_aliases_and_case(self):
        file = open_fixture(SIMPLE)
        for alias, mode in (("t", "text"), ("A", "annotated"), ("o", "outline"), ("STATS", "stats")):
            with self.subTest(alias=alias):
                self.assertEqual(file.view(alias), file.view(mode))

    def test_view_options_are_not_offered_yet(self):
        # litchi-ops keeps its view options (OfficeCLI's --start, --end,
        # --max-lines, --cols, --range) internal; record 0255 follow-up F2.
        file = open_fixture(DRAWING)
        for keyword in ("start", "end", "max_lines", "columns", "range"):
            with self.subTest(keyword=keyword):
                with self.assertRaises(TypeError):
                    file.view("text", **{keyword: 1})

    def test_argument_errors(self):
        docx = open_fixture(HYPERLINK)
        error = raised(lambda: docx.view("outlien"))
        self.assertError(error, L.PropertyError, "invalid_argument", "did you mean 'outline'?")
        error = raised(lambda: docx.view("nonsense"))
        self.assertError(error, L.PropertyError, "invalid_argument")
        self.assertIn("text, annotated, outline or stats", error.suggestion)

    def test_modes_outside_the_semantic_views_are_unsupported(self):
        file = open_fixture(HYPERLINK)
        for mode in ("issues", "forms", "html", "svg", "screenshot", "pdf"):
            for as_json in (False, True):
                with self.subTest(mode=mode, json=as_json):
                    error = raised(lambda: file.view(mode, json=as_json))
                    self.assertError(error, L.UnsupportedError, "unsupported_capability")


# ---------------------------------------------------------------------------
# Capabilities
# ---------------------------------------------------------------------------


class SchemaTests(unittest.TestCase):
    """``litchi_py.schema(format)`` is the Rust registry export
    (``crates/litchi-ops/tests/golden/registry/<format>.json``)."""

    def test_schema_equals_the_rust_registry_golden(self):
        for name in ("doc", "docx", "ppt", "pptx", "rtf", "xls", "xlsb", "xlsx"):
            with self.subTest(format=name):
                exported = L.schema(name)
                expected = golden(f"registry/{name}.json")
                self.assertEqual(exported, expected)
                self.assertEqual(json.dumps(exported), json.dumps(expected))
        self.assertEqual(L.schema(L.FileFormat.Docx), L.schema("docx"))

    def test_stub_typed_dict_matches_the_export(self):
        fields = typed_dict_fields(STUB_PATH)["_Schema"]
        self.assertEqual(list(L.schema("xlsx")), [key for key in fields if key != "pending"])
        # Every claimed format has a registry since record 0296, so no
        # format exports the `pending` shape in this build.
        self.assertEqual(L.schema("pages"), golden("registry/pages.json"))
        self.assertTrue(L.schema("numbers")["elements"])

    def test_unknown_format_names_raise(self):
        with self.assertRaises(L.PropertyError):
            L.schema("txt")


class CapabilitiesTests(unittest.TestCase):
    def test_reports_equal_the_rust_goldens(self):
        for relative, name in ((DRAWING, "docx"), (TABLES, "xlsx"), (PPTX, "pptx")):
            with self.subTest(format=name):
                report = open_fixture(relative).capabilities()
                # Only a file's report says whether its handle retained the
                # editor's source; Capabilities::of (the golden) does not.
                self.assertIs(type(report.pop("retainsSource")), bool)
                expected = golden(f"capabilities/{name}.json")
                self.assertEqual(report, expected)
                self.assertEqual(json.dumps(report), json.dumps(expected))

    def test_stub_typed_dicts_match_the_report(self):
        fields = typed_dict_fields(STUB_PATH)
        report = open_fixture(DRAWING).capabilities()
        self.assertEqual(list(report), list(fields["_Capabilities"]))
        support_keys = set(fields["_Support"])
        for support in (*report["verbs"].values(), *report["views"].values()):
            self.assertTrue(set(support) <= support_keys, support)

    def test_verbs_views_and_elements(self):
        report = open_fixture(SIMPLE).capabilities()
        self.assertEqual((report["format"], report["state"]), ("xlsx", "open"))
        for verb in ("get", "find", "query", "view", "set", "add", "edit", "batch", "save"):
            self.assertEqual(report["verbs"][verb], {"supported": True})
        for mode in ("text", "annotated", "outline", "stats"):
            self.assertTrue(report["views"][mode]["supported"])
        self.assertFalse(report["views"]["pdf"]["supported"])
        cell = next(element for element in report["elements"] if element["element"] == "cell")
        self.assertIn("font.bold", cell["properties"])
        paragraph = next(
            element
            for element in open_fixture(HYPERLINK).capabilities()["elements"]
            if element["element"] == "paragraph"
        )
        self.assertIn("style", paragraph["properties"])

    def test_reads_agree_with_the_report_for_every_format(self):
        for relative, _ in FORMATS:
            with self.subTest(fixture=relative):
                file = open_fixture(relative)
                verb = file.capabilities()["verbs"]["get"]
                if verb["supported"]:
                    self.assertEqual(file.get("/", 0)["path"], "/")
                else:
                    error = raised(lambda: file.get("/"))
                    self.assertIsInstance(error, L.LitchiError)
                    self.assertEqual(error.code, verb["code"])
                    self.assertIn(verb["followUp"], error.suggestion)


class PptxReadTests(unittest.TestCase):
    """PPTX reads through the unified reader (M1 step R4c, on trunk)."""

    def test_reads_answer_slide_paths(self):
        file = open_fixture(PPTX)
        root = file.get("/")
        self.assertEqual(root["type"], "presentation")
        slides = [child["path"] for child in root["children"] if child["type"] == "slide"]
        self.assertEqual(slides, ["/slide[1]", "/slide[2]"])
        self.assertEqual(file.find("/slide[1]"), file.get("/slide[1]"))
        self.assertIsNone(file.find("/slide[9]"))
        for view in file.query("shape"):
            with self.subTest(path=view["path"]):
                self.assertEqual(file.get(view["path"], 0), view)
        self.assertTrue(file.view("text"))
        self.assertIsInstance(file.view("outline", json=True), dict)

    def test_the_file_reports(self):
        file = open_fixture(PPTX)
        self.assertEqual(str(file.format), "Pptx")
        report = file.capabilities()
        self.assertEqual(report["state"], "open")
        self.assertEqual(report["verbs"]["get"], {"supported": True})
        self.assertEqual(file.text(), L.Presentation.open(fixture(PPTX)).text())
        self.assertEqual(file.metadata()["created"], "2007-08-15T06:51:09Z")


# ---------------------------------------------------------------------------
# Signed and locked files
# ---------------------------------------------------------------------------


class SignedAndLockedTests(ErrorAssertions):
    def test_signed_files_read_through(self):
        docx = open_fixture(SIGNED_DOCX)
        self.assertEqual(repr(docx), "<File: Docx, signed>")
        self.assertEqual(docx.capabilities()["state"], "signed")
        self.assertEqual(docx.get("/body/p[1]")["text"], "Hello world")
        self.assertEqual(
            [view["path"] for view in docx.query("paragraph")], ["/body/p[1]", "/body/p[2]"]
        )
        self.assertEqual(docx.view("text"), "[/body/p[1]] Hello world\n[/body/p[2]]")
        self.assertEqual(docx.text(), L.Document.open(fixture(SIGNED_DOCX)).text())
        xlsx = open_fixture(SIGNED_XLSX)
        self.assertEqual(xlsx.capabilities()["state"], "signed")
        self.assertEqual(xlsx.get("/")["children"][0]["path"], "/sheet[1]")
        self.assertTrue(xlsx.view("text").startswith("=== Sheet: Sheet1 ==="))

    def test_locked_files_open_and_refuse_every_read(self):
        for relative, name in LOCKED:
            with self.subTest(fixture=relative):
                file = open_fixture(relative)
                self.assertEqual(str(file.format), name)
                self.assertEqual(repr(file), f"<File: {name}, locked>")
                report = file.capabilities()
                self.assertEqual(report["state"], "locked")
                for verb in ("get", "find", "query", "view"):
                    self.assertEqual(report["verbs"][verb]["code"], "document_protected")
                for call in (
                    lambda: file.get("/"),
                    lambda: file.find("/"),
                    lambda: file.query("paragraph"),
                    lambda: file.view("text"),
                    lambda: file.view("stats", json=True),
                    file.text,
                    file.metadata,
                ):
                    error = raised(call)
                    self.assertError(error, L.ProtectedError, "document_protected", PROTECTED_SUGGESTION)
                    self.assertEqual(error.message, f"the {name} document is encrypted")


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


class ErrorTests(ErrorAssertions):
    def test_path_errors_keep_code_message_and_suggestion(self):
        file = open_fixture(DRAWING)
        error = raised(lambda: file.get("/body/paragrph[1]"))
        self.assertError(error, L.PathError, "invalid_path", 'did you mean "paragraph"?')
        # Messages are content-free (ADR 0034 §5): no path text in them.
        self.assertEqual(error.message, "unknown element at byte 6")
        self.assertIsInstance(error, LookupError)
        error = raised(lambda: file.get("/body/p[9999]"))
        self.assertError(
            error,
            L.PathError,
            "path_not_found",
            "available at the parent element: p[1]..p[185], tbl[1]..tbl[5], section[1]",
        )
        self.assertEqual(error.message, "requested p was not found at byte 6")
        self.assertEqual(
            str(error), f"{error.message} (suggestion: {error.suggestion})"
        )
        self.assertIsNone(error.item)
        error = raised(lambda: file.get("/slide[1]"))
        self.assertError(error, L.PathError, "unsupported_path")
        self.assertIn("Docx paths start at /body", error.suggestion)
        error = raised(lambda: file.get("/body/p["))
        self.assertError(error, L.PathError, "invalid_path", None)
        self.assertIn("at byte 8", error.message)
        error = raised(lambda: open_fixture(SIMPLE).get("/NoSheet/A1"))
        self.assertError(error, L.PathError, "path_not_found", "available at the parent element: sheet[1]..sheet[2]")
        error = raised(lambda: open_fixture(SIMPLE).get("/body"))
        self.assertError(error, L.PathError, "unsupported_path")

    def test_reads_after_the_source_changed(self):
        for relative, path in ((HYPERLINK, "/body/p[1]"), (SIMPLE, "/Sheet1/A1")):
            with self.subTest(fixture=relative), tempfile.TemporaryDirectory() as directory:
                copy = copy_fixture(relative, directory)
                file = L.open(copy)
                modify_in_place(copy)
                for call in (lambda: file.get(path), file.text, lambda: file.view("text")):
                    error = raised(call)
                    self.assertError(error, L.IoError, "io_error")
                    self.assertIn("source changed", error.message)


# ---------------------------------------------------------------------------
# Lifecycle
# ---------------------------------------------------------------------------


class LifecycleTests(ErrorAssertions):
    def test_context_manager_closes(self):
        with open_fixture(HYPERLINK) as file:
            self.assertIsInstance(file, L.File)
            self.assertFalse(file.closed)
            self.assertEqual(file.get("/body/p[2]")["text"], "EricWhite.com")
        self.assertTrue(file.closed)
        self.assertEqual(repr(file), "<File: Docx, closed>")

    def test_exceptions_propagate_and_the_file_closes(self):
        with self.assertRaises(KeyError):
            with open_fixture(HYPERLINK) as file:
                raise KeyError("inside")
        self.assertTrue(file.closed)
        self.assertIs(file.__exit__(None, None, None), False)

    def test_reads_after_close(self):
        file = open_fixture(SIMPLE)
        file.close()
        file.close()  # idempotent
        self.assertTrue(file.closed)
        self.assertEqual(str(file.format), "Xlsx")  # still known
        for call in (
            lambda: file.get("/"),
            lambda: file.find("/"),
            lambda: file.query("cell"),
            lambda: file.view("text"),
            file.capabilities,
            file.text,
            file.metadata,
            file.__enter__,
        ):
            error = raised(call)
            self.assertError(error, L.IoError, "file_required", "open it again with litchi_py.open")
            self.assertEqual(error.message, "the file is closed")
            self.assertIsInstance(error, OSError)

    def test_views_outlive_the_file(self):
        file = open_fixture(HYPERLINK)
        view = file.get("/body", 2)
        file.close()
        del file
        self.assertEqual(view["children"][0]["text"], "Following is a hyperlink.")

    def test_results_are_independent_copies(self):
        file = open_fixture(HYPERLINK)
        first = file.get("/body")
        first["children"].clear()
        first["format"]["changed"] = True
        self.assertEqual(len(file.get("/body")["children"]), 3)
        self.assertEqual(file.get("/body")["format"], {})


# ---------------------------------------------------------------------------
# Text and metadata
# ---------------------------------------------------------------------------


class TextAndMetadataTests(unittest.TestCase):
    def test_text_is_the_family_readers_text(self):
        for relative, family in ((HYPERLINK, L.Document), (DRAWING, L.Document), (SIMPLE, L.Workbook), (PPTX, L.Presentation)):
            with self.subTest(fixture=relative):
                self.assertEqual(open_fixture(relative).text(), family.open(fixture(relative)).text())

    def test_metadata(self):
        self.assertEqual(
            open_fixture(SIMPLE).metadata(),
            {
                "author": "Nick Burch",
                "last_modified_by": "Nick Burch",
                "created": "2015-03-13T11:12:37Z",
                "modified": "2015-03-13T11:16:26Z",
                # docProps/app.xml joins the core properties (record 0253).
                "application": "Microsoft Excel",
                "security": 0,
            },
        )
        docx = open_fixture(HYPERLINK).metadata()
        self.assertEqual((docx["author"], docx["revision"]), ("Eric White", "2"))
        declared = typed_dict_fields(STUB_PATH)["_Metadata"]
        for relative, _ in FORMATS:
            metadata = open_fixture(relative).metadata()
            with self.subTest(fixture=relative):
                if metadata is not None:
                    self.assertTrue(set(metadata) <= set(declared), set(metadata) - set(declared))


if __name__ == "__main__":
    unittest.main()
