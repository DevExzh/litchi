#!/usr/bin/env python3
"""Admission ladder for the XLSX source-backed value editor (change 0602).

Reimplements, over the raw package bytes, the four structural gates that
`litchi_xlsx::cell_values` applies before change 0525's reduced readback can
run, and reports for each fixture which gate refuses it first and which gates
would refuse it at all.

  G1 validate_package_relationships  snapshot.rs:1924-1952
     package-root rels must all be officeDocument / strict officeDocument /
     digital-signature origin, and exactly one officeDocument owner.
  G2 validate_workbook_relationships snapshot.rs:1704-1751
     workbook rels must all be worksheet / styles / theme / calcChain
     (Transitional or Strict); at most one styles, theme, calcChain.
     sharedStrings is NOT in the allow-list.
  G3 worksheet relationship refusal   snapshot.rs:441-443
     the selected worksheet part must carry no relationships at all.
  G4 shared-string planning parse     raw/worksheet/semantic.rs:134-146
     every cell_values parse passes `|| Ok(None)`, so a `t="s"` cell with a
     <v> child is a hard refusal at planning.
  G5 stored_entry_is_supported        snapshot.rs:365-370
     reduced readback off when any entry is inline-rich or metadata-bearing
     (its shared_string clause is unreachable because G4 fires first).
"""
import os, re, sys, zipfile

NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
SOD = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
DSO = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
B = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
S = "http://purl.oclc.org/ooxml/officeDocument/relationships/"
WB_OK = {B+"worksheet", S+"worksheet", B+"styles", S+"styles", B+"theme",
         B+"calcChain", S+"calcChain"}
COUNTED = {B+"styles": "styles", S+"styles": "styles", B+"theme": "theme",
           B+"calcChain": "chain", S+"calcChain": "chain"}
SHEETRT = {B+"worksheet", S+"worksheet"}

RELRE = re.compile(rb'<Relationship\b[^>]*>')
ATTR = re.compile(rb'(\w+)\s*=\s*"([^"]*)"')

def rels(z, part):
    d, f = (part.rsplit("/", 1) if "/" in part else ("", part))
    name = (d + "/_rels/" + f + ".rels") if d else "_rels/" + f + ".rels"
    if part == "":
        name = "_rels/.rels"
    try:
        data = z.read(name)
    except KeyError:
        return []
    out = []
    for m in RELRE.finditer(data):
        a = {k.decode(): v.decode() for k, v in ATTR.findall(m.group(0))}
        out.append(a)
    return out

def resolve(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    d = base.rsplit("/", 1)[0] if "/" in base else ""
    parts = (d.split("/") if d else []) + target.split("/")
    stack = []
    for p in parts:
        if p in ("", "."):
            continue
        if p == "..":
            if stack: stack.pop()
        else:
            stack.append(p)
    return "/".join(stack)

def gates(path):
    r = dict(g1=None, g2=None, g3=None, g4=None, g5=None, sheet=None)
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        pkg = rels(z, "")
        owners = 0
        for rel in pkg:
            rt = rel.get("Type", "")
            if rel.get("TargetMode") == "External" or rt not in (OD, SOD, DSO):
                if r["g1"] is None:
                    r["g1"] = rt
            if rt in (OD, SOD):
                owners += 1
                wb = resolve("", rel.get("Target", ""))
        if owners != 1 and r["g1"] is None:
            r["g1"] = "owner-count=%d" % owners
        if owners == 0:
            return r
        wbrels = rels(z, wb)
        seen = {}
        for rel in wbrels:
            rt = rel.get("Type", "")
            if rel.get("TargetMode") == "External":
                if r["g2"] is None: r["g2"] = "external"
            elif rt not in WB_OK:
                if r["g2"] is None: r["g2"] = rt
            if rt in COUNTED:
                seen[COUNTED[rt]] = seen.get(COUNTED[rt], 0) + 1
        if r["g2"] is None and any(v > 1 for v in seen.values()):
            r["g2"] = "duplicate-" + ",".join(k for k, v in seen.items() if v > 1)
        sheets = [resolve(wb, rel["Target"]) for rel in wbrels
                  if rel.get("Type") in SHEETRT and rel.get("TargetMode") != "External"]
        sheets = [s for s in sheets if s in names]
        if not sheets:
            return r
        ws = sheets[0]
        r["sheet"] = ws
        if rels(z, ws):
            r["g3"] = "worksheet-rels"
        data = z.read(ws)
        # G4: a t="s" cell with a <v> child.
        for m in re.finditer(rb'<c\b[^>]*\bt\s*=\s*"s"[^>]*(/>|>)', data):
            if m.group(1) == b">":
                tail = data[m.end():m.end()+400]
                stop = tail.find(b"</c>")
                if stop >= 0 and b"<v" in tail[:stop]:
                    r["g4"] = "shared-string-value"
                    break
        # G5: inline rich runs, or cm=/vm= metadata attributes.
        if re.search(rb'<is\b[^>]*>\s*<r\b', data):
            r["g5"] = "inline-rich"
        elif re.search(rb'<c\b[^>]*\b(cm|vm)\s*=', data):
            r["g5"] = "cell-metadata"
    return r

def main():
    paths = []
    for a in sys.argv[1:]:
        if os.path.isdir(a):
            for d, _, fs in os.walk(a):
                for f in fs:
                    if f.lower().endswith(".xlsx"):
                        paths.append(os.path.join(d, f))
        else:
            paths.append(a)
    paths.sort()
    print("first\tg1\tg2\tg3\tg4\tg5\tsheet\tpath")
    tally = {}
    blocked = dict(g1=0, g2=0, g3=0, g4=0, g5=0)
    clear = 0
    for p in paths:
        try:
            g = gates(p)
        except Exception as e:
            print("ERR\t-\t-\t-\t-\t-\t-\t%s\t%s" % (p, e))
            continue
        first = next((k for k in ("g1", "g2", "g3", "g4", "g5") if g[k]), "admitted")
        tally[first] = tally.get(first, 0) + 1
        for k in blocked:
            if g[k]: blocked[k] += 1
        if first == "admitted": clear += 1
        print("%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s" % (
            first, g["g1"] or "-", g["g2"] or "-", g["g3"] or "-",
            g["g4"] or "-", g["g5"] or "-", g["sheet"] or "-", p))
    print()
    print("# files=%d" % len(paths))
    for k in ("g1", "g2", "g3", "g4", "g5", "admitted"):
        print("# first refusal %-9s %d" % (k, tally.get(k, 0)))
    print("# would refuse (independently):")
    for k in ("g1", "g2", "g3", "g4", "g5"):
        print("#   %s %d" % (k, blocked[k]))

main()
