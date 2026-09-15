#!/usr/bin/env python3
"""Change 0597: build the synthetic first-error matrix.

Each case rewrites `xl/worksheets/sheet1.xml` of one small valid package with a
worksheet that is malformed (or merely ineligible) at a chosen position, in
both a `<cols>`-bearing shape (the pre-gate fires) and a `<cols>`-free control
shape (the pre-gate declines and today's streaming scan runs). The reader's
first typed error for each case is then compared across measurement legs.
"""
import os, shutil, sys, zipfile

BASE = "/home/zhuhe/code/litchi-worktrees/0597/test-data/libreoffice-core/sc/qa/unit/data/xlsx/page_scale.xlsx"
SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

COLS = '<cols><col min="1" max="3" width="12"/></cols>'
HEAD = '<dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
BODY = '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData>'
TAIL = '<mergeCells count="1"><mergeCell ref="B1:B2"/></mergeCells><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'


def sheet(head=HEAD, cols="", body=BODY, tail=TAIL, root_attrs="", suffix="</worksheet>"):
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{SML}"{root_attrs}>{head}{cols}{body}{tail}{suffix}'
    ).encode("utf-8")


def cases():
    """(name, position, xml) triples; every case is emitted with and without <cols>."""
    out = []

    def add(name, position, **kw):
        out.append((name, position, kw))

    # --- well formed, merely ineligible ------------------------------------
    add("valid", "none")
    add("row-style", "in-sheetdata",
        body='<sheetData><row r="1" s="3" customFormat="1"><c r="A1"><v>1</v></c></row></sheetData>')

    # --- malformed strictly inside the gate prefix -------------------------
    add("prefix-bad-entity", "before-cols",
        head='<dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0" topLeftCell="A&bogus;1"/></sheetViews>')
    add("prefix-unclosed-element", "before-cols",
        head='<dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"/>')
    add("prefix-mismatched-end", "before-cols",
        head='<dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"/></sheetViewsX>')
    add("prefix-bad-dimension", "before-cols", head='<dimension ref="1:102"/>')
    add("prefix-stray-lt", "before-cols",
        head='<dimension ref="A1:C3"/><sheetViews>< </sheetViews>')
    add("prefix-undeclared-prefix", "before-cols",
        head='<dimension ref="A1:C3"/><zz:sheetPr xmlns:qq="urn:x"/>')

    # --- malformed after the gate window, before sheetData -----------------
    add("post-cols-mismatched-end", "after-cols",
        body='<sheetFormatPr/></zzz>' + BODY)
    add("post-cols-duplicate-dimension", "after-cols", body='<dimension ref="A1:C3"/>' + BODY)

    # --- malformed inside sheetData ---------------------------------------
    add("cell-bad-boolean", "in-sheetdata",
        body='<sheetData><row r="1"><c r="A1" t="b"><v>maybe</v></c></row></sheetData>')
    add("cell-ref-not-in-row", "in-sheetdata",
        body='<sheetData><row r="1"><c r="A2"><v>1</v></c></row></sheetData>')
    add("cell-unclosed", "in-sheetdata",
        body='<sheetData><row r="1"><c r="A1"><v>1</v></row></sheetData>')
    add("sheetdata-bad-entity", "in-sheetdata",
        body='<sheetData><row r="1"><c r="A1" t="str"><v>a&nope;b</v></c></row></sheetData>')
    add("row-out-of-order", "in-sheetdata",
        body='<sheetData><row r="2"><c r="A2"><v>2</v></c></row><row r="1"><c r="A1"><v>1</v></c></row></sheetData>')
    add("shared-formula-not-first", "in-sheetdata",
        body='<sheetData><row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2"/><v>1</v></c></row></sheetData>')

    # --- malformed after sheetData ----------------------------------------
    add("merge-after-successor", "after-sheetdata",
        tail='<hyperlinks/><mergeCells count="1"><mergeCell ref="B1:B2"/></mergeCells>')
    add("merge-count-mismatch", "after-sheetdata",
        tail='<mergeCells count="2"><mergeCell ref="B1:B2"/></mergeCells>')
    add("merge-bad-ref", "after-sheetdata",
        tail='<mergeCells count="1"><mergeCell ref="B1"/></mergeCells>')
    add("duplicate-mergecells", "after-sheetdata",
        tail='<mergeCells count="1"><mergeCell ref="B1:B2"/></mergeCells>'
             '<mergeCells count="1"><mergeCell ref="C1:C2"/></mergeCells>')
    add("dimension-after-sheetdata", "after-sheetdata", tail='<dimension ref="A1:C3"/>')
    add("trailing-unclosed-root", "after-sheetdata", suffix="")
    add("second-root", "after-sheetdata", suffix=f'</worksheet><worksheet xmlns="{SML}"/>')
    add("tail-mismatched-end", "after-sheetdata", tail='<pageMargins/></nope>')

    # --- MCE / x14ac -------------------------------------------------------
    add("mce-ignorable-undeclared", "root",
        root_attrs=f' xmlns:mc="{MC}" mc:Ignorable="nosuchprefix"')
    add("mce-alternatecontent-empty", "after-cols",
        body=f'<mc:AlternateContent xmlns:mc="{MC}"/>' + BODY)
    add("mce-alternatecontent-bad-choice", "after-cols",
        body=f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice><x/></mc:Choice></mc:AlternateContent>' + BODY)
    add("x14ac-bad-descent", "in-sheetdata",
        root_attrs=f' xmlns:mc="{MC}" mc:Ignorable="x14ac" xmlns:x14ac="{X14AC}"',
        body='<sheetData><row r="1" x14ac:dyDescent="notanumber"><c r="A1"><v>1</v></c></row></sheetData>')
    add("x14ac-descent-ok", "in-sheetdata",
        root_attrs=f' xmlns:mc="{MC}" mc:Ignorable="x14ac" xmlns:x14ac="{X14AC}"',
        body='<sheetData><row r="1" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c></row></sheetData>')
    return out


def write_case(out_dir, name, xml):
    path = os.path.join(out_dir, f"{name}.xlsx")
    with zipfile.ZipFile(BASE) as src, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = xml
            dst.writestr(item.filename, data)


def main():
    out_dir = sys.argv[1]
    shutil.rmtree(out_dir, ignore_errors=True)
    os.makedirs(out_dir)
    count = 0
    for name, position, kw in cases():
        for shape, cols in (("cols", COLS), ("nocols", "")):
            write_case(out_dir, f"{position}__{name}__{shape}", sheet(cols=cols, **kw))
            count += 1
    # A raw non-UTF-8 worksheet cannot be produced through `sheet()`.
    for shape, cols in (("cols", COLS), ("nocols", "")):
        body = b'<sheetData><row r="1"><c r="A1" t="str"><v>\xff\xfe</v></c></row></sheetData>'
        xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{SML}">'
               f'{HEAD}{cols}').encode() + body + f'{TAIL}</worksheet>'.encode()
        write_case(out_dir, f"in-sheetdata__invalid-utf8__{shape}", xml)
        count += 1
    print(f"{count} synthetic packages in {out_dir}")


main()
