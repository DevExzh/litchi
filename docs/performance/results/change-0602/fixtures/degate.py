#!/usr/bin/env python3
"""Derive an admissible twin of a real .xlsx for change 0602's sizing legs.

The source-backed value editor refuses every real fixture at its package-root
relationship allow-list, so no real producer file can be measured as shipped.
This tool removes the minimum that each admission gate demands, and nothing
else, so the derived file keeps the producer's row and cell geometry, its
styles, its markers and its worksheet envelope:

  G1  drop every package-root relationship that is not the officeDocument
      owner, and the parts they reach (docProps/app.xml, docProps/core.xml,
      thumbnails, vendor metadata).
  G2  drop every workbook relationship outside the value editor's allow-list.
      When sharedStrings goes, its strings are folded back into the worksheet
      as inline strings so no cell content is lost.
  G3  drop each worksheet's own relationships and the r:id-bearing worksheet
      children that reach them (hyperlinks, drawings, tableParts, oleObjects,
      controls, printerSettings via pageSetup).

Variants:
  plain  shared strings become <is><t>text</t></is>       (reduced readback ON)
  rich   shared strings become <is><r><t>text</t></r></is> (reduced readback OFF,
         because Stored.inline_rich flips stored_entry_is_supported to false)

The two variants differ by exactly the six bytes of one <r> wrapper per former
shared-string cell, so a paired comparison isolates change 0525's gate.
"""
import re, sys, zipfile

B = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
S = "http://purl.oclc.org/ooxml/officeDocument/relationships/"
OD, SOD = B + "officeDocument", S + "officeDocument"
DSO = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
WB_OK = {B+"worksheet", S+"worksheet", B+"styles", S+"styles", B+"theme",
         B+"calcChain", S+"calcChain"}
RELRE = re.compile(rb'<Relationship\b[^>]*/>')
ATTR = re.compile(rb'(\w[\w:]*)\s*=\s*"([^"]*)"')

# Worksheet children that carry or contain an r:id.
CONTAINERS = ["hyperlinks", "tableParts", "oleObjects", "controls",
              "webPublishItems", "extLst"]
EMPTIES = ["drawing", "legacyDrawing", "legacyDrawingHF", "picture",
           "pageSetup", "drawingHF"]

def relname(part):
    if part == "":
        return "_rels/.rels"
    d, f = part.rsplit("/", 1) if "/" in part else ("", part)
    return (d + "/_rels/" + f + ".rels") if d else "_rels/" + f + ".rels"

def resolve(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    d = base.rsplit("/", 1)[0] if "/" in base else ""
    stack = []
    for p in ((d.split("/") if d else []) + target.split("/")):
        if p in ("", "."):
            continue
        if p == "..":
            if stack: stack.pop()
        else:
            stack.append(p)
    return "/".join(stack)

def parse_rels(blob):
    out = []
    for m in RELRE.finditer(blob or b""):
        out.append((m.group(0), {k.decode(): v.decode() for k, v in ATTR.findall(m.group(0))}))
    return out

def read_sst(blob):
    """Each <si> flattened to its concatenated <t> text, XML-escaped as found."""
    items = []
    for si in re.findall(rb'<si\b[^>]*>(.*?)</si>|<si\b[^>]*/>', blob, re.S):
        body = si if isinstance(si, bytes) else b""
        items.append(b"".join(re.findall(rb'<t\b[^>]*>(.*?)</t>', body, re.S)))
    return items

def inline_shared(ws, sst, rich):
    """Rewrite every t="s" cell into an inline string of the chosen shape."""
    def repl(m):
        head, body = m.group(1), m.group(2)
        v = re.search(rb'<v\b[^>]*>(.*?)</v>', body, re.S)
        if v is None:
            return m.group(0)
        try:
            text = sst[int(v.group(1).strip())]
        except (ValueError, IndexError):
            return m.group(0)
        head = re.sub(rb'\st\s*=\s*"s"', b' t="inlineStr"', head)
        inner = (b"<is><r><t>" + text + b"</t></r></is>") if rich else \
                (b"<is><t>" + text + b"</t></is>")
        return b"<c" + head + b">" + inner + b"</c>"
    return re.sub(rb'<c((?:[^>]*?)\st\s*=\s*"s"(?:[^>]*?))>(.*?)</c>', repl, ws, flags=re.S)

def strip_rid_children(ws):
    for tag in EMPTIES:
        ws = re.sub((r'<%s\b[^>]*\br:id\s*=[^>]*/>' % tag).encode(), b"", ws)
        ws = re.sub((r'<%s\b[^>]*\br:id\s*=[^>]*?>.*?</%s>' % (tag, tag)).encode(), b"", ws, flags=re.S)
        # pageSetup keeps its attributes but loses the printerSettings link
        ws = re.sub((r'(<%s\b[^>]*?)\sr:id\s*=\s*"[^"]*"' % tag).encode(), rb"\1", ws)
    for tag in CONTAINERS:
        def drop(m, tag=tag):
            return b"" if b"r:id" in m.group(0) else m.group(0)
        ws = re.sub((r'<%s\b[^>]*?>.*?</%s>' % (tag, tag)).encode(), drop, ws, flags=re.S)
        ws = re.sub((r'<%s\b[^>]*\br:id\s*=[^>]*/>' % tag).encode(), b"", ws)
    return ws

def main():
    src, dst, variant = sys.argv[1], sys.argv[2], sys.argv[3]
    rich = variant == "rich"
    zin = zipfile.ZipFile(src)
    blobs = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    # G1: package-root relationships.
    pkg = parse_rels(blobs.get("_rels/.rels"))
    keep_pkg = [(raw, a) for raw, a in pkg
                if a.get("TargetMode") != "External" and a.get("Type") in (OD, SOD, DSO)]
    wb = next(resolve("", a["Target"]) for _, a in keep_pkg if a["Type"] in (OD, SOD))

    # G2: workbook relationships.
    wbrels = parse_rels(blobs.get(relname(wb)))
    keep_wb = [(raw, a) for raw, a in wbrels
               if a.get("TargetMode") != "External" and a.get("Type") in WB_OK]
    dropped_sst = [resolve(wb, a["Target"]) for _, a in wbrels
                   if a.get("Type", "").endswith("/sharedStrings")]
    sst = read_sst(blobs[dropped_sst[0]]) if dropped_sst and dropped_sst[0] in blobs else []

    sheets = [resolve(wb, a["Target"]) for _, a in keep_wb
              if a["Type"] in (B+"worksheet", S+"worksheet")]

    # G3 plus the shared-string fold, applied per worksheet.
    for ws in sheets:
        blob = blobs[ws]
        if sst:
            blob = inline_shared(blob, sst, rich)
        blob = strip_rid_children(blob)
        blobs[ws] = blob
        blobs.pop(relname(ws), None)

    # Reachability over the surviving relationship graph.
    keep = {"[Content_Types].xml", "_rels/.rels", relname(wb)}
    frontier, seen = [wb], set()
    keep.add(wb)
    while frontier:
        part = frontier.pop()
        if part in seen:
            continue
        seen.add(part)
        keep.add(part)
        rn = relname(part)
        if rn in blobs:
            keep.add(rn)
            for _, a in parse_rels(blobs[rn]):
                if a.get("TargetMode") == "External":
                    continue
                t = resolve(part, a["Target"])
                if t in blobs:
                    frontier.append(t)

    # Rewrite the two relationship parts we pruned.
    def rebuild(entries):
        body = b"".join(raw for raw, _ in entries)
        return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + body + b'</Relationships>')
    blobs["_rels/.rels"] = rebuild(keep_pkg)
    blobs[relname(wb)] = rebuild(keep_wb)

    # Content types: drop overrides for parts that no longer exist.
    ct = blobs["[Content_Types].xml"]
    def keep_override(m):
        name = m.group(0)
        pn = re.search(rb'PartName\s*=\s*"([^"]*)"', name)
        return name if pn and pn.group(1).decode().lstrip("/") in keep else b""
    ct = re.sub(rb'<Override\b[^>]*/>', keep_override, ct)
    blobs["[Content_Types].xml"] = ct

    order = [n for n in blobs if n in keep]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
        for n in order:
            out.writestr(n, blobs[n])
    print("%s -> %s  variant=%s parts=%d sst_items=%d sheets=%d"
          % (src, dst, variant, len(order), len(sst), len(sheets)))

main()
