#!/usr/bin/env python3
"""Synthetic twins that isolate each value-only admission gate (change 0602).

Every package here is minimal, compact and inside the value-only element and
attribute allow-lists except for the single feature its variant adds, so the
refusal it produces names exactly one gate.

  numeric  numeric cells only                       expected: admitted
  inline   plain inline strings <is><t>             expected: admitted
  rich     rich inline strings <is><r><t>           expected: element 'r'
  sst      t="s" cells + a sharedStrings part+rel   expected: workbook relationship
  sstfree  t="s" cells, no sharedStrings part       expected: no shared-string part
  cm       cm="1" on every <c>                      expected: attribute 'cm' on 'c'
  vm       vm="1" on every <c>                      expected: attribute 'vm' on 'c'

usage: synth.py <outdir> <rows> <cols> <variant>...
"""
import os, sys, zipfile

SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
DR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

def col(i):
    s = ""
    while True:
        s = chr(ord("A") + i % 26) + s
        i = i // 26 - 1
        if i < 0:
            return s

def sheet(rows, cols, variant):
    out = [DECL, '<worksheet xmlns="%s">' % SML,
           '<dimension ref="A1:%s%d"/>' % (col(cols - 1), rows), "<sheetData>"]
    strings, nstr = [], 0
    extra = ' cm="1"' if variant == "cm" else (' vm="1"' if variant == "vm" else "")
    for r in range(1, rows + 1):
        out.append('<row r="%d">' % r)
        for c in range(cols):
            ref = "%s%d" % (col(c), r)
            # Every third column carries text, the rest numerals: a producer-like mix.
            if c % 3 == 2 and variant in ("inline", "rich", "sst", "sstfree"):
                text = "label-%d-%d" % (r, c)
                if variant == "inline":
                    out.append('<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (ref, extra, text))
                elif variant == "rich":
                    out.append('<c r="%s"%s t="inlineStr"><is><r><t>%s</t></r></is></c>' % (ref, extra, text))
                else:
                    strings.append(text)
                    out.append('<c r="%s"%s t="s"><v>%d</v></c>' % (ref, extra, nstr))
                    nstr += 1
            else:
                out.append('<c r="%s"%s><v>%d</v></c>' % (ref, extra, r * 1000 + c))
        out.append("</row>")
    out += ["</sheetData>", "</worksheet>"]
    return "".join(out).encode(), strings

def build(path, rows, cols, variant):
    ws, strings = sheet(rows, cols, variant)
    want_sst = variant == "sst"
    wbrels = ['<Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>' % DR]
    if want_sst:
        wbrels.append('<Relationship Id="rId2" Type="%s/sharedStrings" Target="sharedStrings.xml"/>' % DR)
    parts = {
        "[Content_Types].xml": DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            + ('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' if want_sst else "")
            + "</Types>",
        "_rels/.rels": DECL + '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/></Relationships>' % (PR, DR),
        "xl/workbook.xml": DECL + '<workbook xmlns="%s" xmlns:r="%s"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>' % (SML, DR),
        "xl/_rels/workbook.xml.rels": DECL + '<Relationships xmlns="%s">%s</Relationships>' % (PR, "".join(wbrels)),
        "xl/worksheets/sheet1.xml": ws,
    }
    if want_sst:
        items = "".join("<si><t>%s</t></si>" % s for s in strings)
        parts["xl/sharedStrings.xml"] = (DECL + '<sst xmlns="%s" count="%d" uniqueCount="%d">%s</sst>'
                                         % (SML, len(strings), len(strings), items))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for k, v in parts.items():
            z.writestr(k, v if isinstance(v, bytes) else v.encode())
    return len(ws)

def main():
    outdir, rows, cols = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
    os.makedirs(outdir, exist_ok=True)
    for variant in sys.argv[4:]:
        p = os.path.join(outdir, "synth-%dx%d-%s.xlsx" % (rows, cols, variant))
        n = build(p, rows, cols, variant)
        print("%s cells=%d sheet_bytes=%d" % (p, rows * cols, n))

main()
