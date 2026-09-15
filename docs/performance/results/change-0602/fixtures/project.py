#!/usr/bin/env python3
"""Project a real .xlsx onto the value editor's admission surface (change 0602).

`degate.py` opens the three relationship gates. This opens the remaining two:
the value-only element allow-list and attribute allow-list in
`crates/litchi-xlsx/src/cell_values/validation.rs:367-620`. Everything outside
those lists is removed from `xl/workbook.xml` and from each worksheet, while
`<sheetData>` keeps the producer's exact row and cell geometry: the same rows,
the same `<c>` records in the same order with the same styles, the same
numerals and the same string lengths. That geometry is what planning and commit
cost is proportional to, so the projection is the largest real-producer input
the editor can be measured on today.

Variants (applied inside <sheetData>):
  plain    inline strings stay <is><t>text</t></is>            admitted
  rich     inline strings become <is><r><t>text</t></r></is>   refused: <r> is
           not in the element allow-list, so Stored.inline_rich is unreachable
  cm       every <c> gains cm="0"                              refused: cm is
           not in the <c> attribute allow-list, so cell_metadata is unreachable
"""
import re, sys, zipfile

WS_ELEMENTS = {b"worksheet", b"dimension", b"sheetViews", b"sheetView", b"pane",
               b"selection", b"sheetFormatPr", b"cols", b"col", b"sheetData",
               b"row", b"c", b"f", b"v", b"is", b"t"}
WB_ELEMENTS = {b"workbook", b"fileVersion", b"workbookPr", b"bookViews",
               b"workbookView", b"sheets", b"sheet", b"calcPr"}
WB_ATTRS = {
    b"fileVersion": {b"appName", b"lastEdited", b"lowestEdited", b"rupBuild", b"codeName"},
    b"workbookPr": {b"date1904", b"showObjects", b"showBorderUnselectedTables", b"filterPrivacy",
                    b"promptedSolutions", b"showInkAnnotation", b"backupFile",
                    b"saveExternalLinkValues", b"updateLinks", b"codeName", b"hidePivotFieldList",
                    b"showPivotChartFilter", b"allowRefreshQuery", b"publishItems",
                    b"checkCompatibility", b"autoCompressPictures", b"refreshAllConnections",
                    b"defaultThemeVersion"},
    b"workbookView": {b"visibility", b"minimized", b"showHorizontalScroll", b"showVerticalScroll",
                      b"showSheetTabs", b"xWindow", b"yWindow", b"windowWidth", b"windowHeight",
                      b"tabRatio", b"firstSheet", b"activeTab", b"autoFilterDateGrouping"},
    b"sheet": {b"name", b"sheetId", b"state", b"r:id"},
}
WS_ATTRS = {
    b"dimension": {b"ref"},
    b"sheetView": {b"windowProtection", b"showFormulas", b"showGridLines", b"showRowColHeaders",
                   b"showZeros", b"rightToLeft", b"tabSelected", b"showRuler",
                   b"showOutlineSymbols", b"defaultGridColor", b"showWhiteSpace", b"view",
                   b"topLeftCell", b"colorId", b"zoomScale", b"zoomScaleNormal",
                   b"zoomScaleSheetLayoutView", b"zoomScalePageLayoutView", b"workbookViewId"},
    b"pane": {b"xSplit", b"ySplit", b"topLeftCell", b"activePane", b"state"},
    b"selection": {b"pane", b"activeCell", b"activeCellId", b"sqref"},
    b"sheetFormatPr": {b"baseColWidth", b"defaultColWidth", b"defaultRowHeight", b"customHeight",
                       b"zeroHeight", b"thickTop", b"thickBottom", b"outlineLevelRow",
                       b"outlineLevelCol"},
    b"col": {b"min", b"max", b"width", b"style", b"hidden", b"bestFit", b"customWidth",
             b"phonetic", b"outlineLevel", b"collapsed"},
    b"row": {b"r", b"spans", b"s", b"customFormat", b"ht", b"hidden", b"customHeight",
             b"outlineLevel", b"collapsed", b"thickTop", b"thickBot", b"ph"},
    b"c": {b"r", b"s", b"t"},
    b"f": {b"t", b"ref", b"si", b"dt2D", b"dtr", b"del1", b"del2", b"r1", b"r2", b"ca", b"bx"},
    b"t": {b"xml:space"},
}
ATTRRE = re.compile(rb'([\w:.\-]+)\s*=\s*"([^"]*)"')
TAGRE = re.compile(rb'<(/?)([\w:.\-]+)((?:\s[^<>]*?)?)(/?)>', re.S)

def filter_attrs(tag, raw, allowed):
    keep = []
    for m in ATTRRE.finditer(raw):
        name = m.group(1)
        if name == b"xmlns" or name.startswith(b"xmlns:"):
            keep.append(m.group(0))
        elif name in allowed.get(tag, set()):
            keep.append(m.group(0))
    return (b" " + b" ".join(keep)) if keep else b""

def project(xml, elements, attrs):
    """Drop every element outside `elements` with its subtree; filter attributes."""
    out, depth_drop, stack = bytearray(), 0, []
    pos = 0
    for m in TAGRE.finditer(xml):
        if depth_drop == 0:
            out += xml[pos:m.start()]
        pos = m.end()
        closing, name, raw, selfclose = m.group(1), m.group(2), m.group(3) or b"", m.group(4)
        local = name.rsplit(b":", 1)[-1] if b":" in name else name
        bad = (name != local) or (local not in elements)
        if closing:
            if depth_drop:
                if stack and stack[-1] == name:
                    stack.pop()
                    if not stack:
                        depth_drop = 0
                continue
            out += m.group(0)
            continue
        if depth_drop:
            if not selfclose:
                stack.append(name)
            continue
        if bad:
            if not selfclose:
                depth_drop, stack = 1, [name]
            continue
        out += b"<" + name + filter_attrs(local, raw, attrs) + (b"/>" if selfclose else b">")
    out += xml[pos:]
    return bytes(out)

def variantize(ws, variant):
    if variant == "rich":
        ws = re.sub(rb'<is>(\s*)<t([^>]*)>(.*?)</t>\s*</is>',
                    rb'<is><r><t\2>\3</t></r></is>', ws, flags=re.S)
    elif variant == "cm":
        ws = re.sub(rb'<c ', b'<c cm="0" ', ws)
    return ws

def main():
    src, dst, variant = sys.argv[1], sys.argv[2], sys.argv[3]
    zin = zipfile.ZipFile(src)
    blobs = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    wb = "xl/workbook.xml"
    blobs[wb] = project(blobs[wb], WB_ELEMENTS, WB_ATTRS)
    n = 0
    for name in list(blobs):
        if re.match(r"xl/worksheets/sheet[^/]*\.xml$", name):
            blobs[name] = variantize(project(blobs[name], WS_ELEMENTS, WS_ATTRS), variant)
            n += 1
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
        for k, v in blobs.items():
            out.writestr(k, v)
    print("%s -> %s variant=%s worksheets=%d" % (src, dst, variant, n))

main()
