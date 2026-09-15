#!/usr/bin/env python3
"""Structural census of the real .xlsx fixtures for change 0602.

For every .xlsx under the given roots, report the three structural facts that
decide whether change 0525's reduced readback and the source-backed value
editor can run at all:

  sst       sharedStrings part present in the package
  ts        the first worksheet contains at least one t="s" cell
  rels      the first worksheet carries a _rels/<sheet>.xml.rels part
  marker    the first worksheet contains mc:/AlternateContent/x14ac/dyDescent
  size      uncompressed bytes of the first worksheet part

Output is a TSV, sorted by path, plus a summary block.
"""
import os, re, sys, zipfile

MARKERS = (b"AlternateContent", b"x14ac", b"dyDescent", b"mc:Ignorable")

def first_sheet(z):
    names = [n for n in z.namelist() if re.match(r"xl/worksheets/sheet[^/]*\.xml$", n)]
    names.sort(key=lambda n: (len(n), n))
    return names[0] if names else None

def census(path):
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            sst = any(n.lower().endswith("sharedstrings.xml") for n in names)
            ws = first_sheet(z)
            if ws is None:
                return dict(sst=sst, ws=None, ts=False, rels=False, anyrels=False, marker=False, size=0, sheets=0)
            rel = ws.rsplit("/", 1)[0] + "/_rels/" + ws.rsplit("/", 1)[1] + ".rels"
            data = z.read(ws)
            ts = re.search(rb'\bt\s*=\s*"s"', data) is not None
            marker = any(m in data for m in MARKERS)
            allws = [n for n in names if re.match(r"xl/worksheets/sheet[^/]*\.xml$", n)]
            anyrels = any((n.rsplit("/",1)[0] + "/_rels/" + n.rsplit("/",1)[1] + ".rels") in names
                          for n in allws)
            return dict(sst=sst, ws=ws, ts=ts, rels=rel in names, anyrels=anyrels,
                        marker=marker, size=len(data), sheets=len(allws))
    except Exception as e:
        return dict(err=str(e))

def main():
    roots = sys.argv[1:]
    files = []
    for root in roots:
        for d, _, fs in os.walk(root):
            for f in fs:
                if f.lower().endswith(".xlsx"):
                    files.append(os.path.join(d, f))
    files.sort()
    print("sst\tts\trels\tanyrels\tmarker\tsheets\tsheet_bytes\tpath")
    n = sst = ts = rels = anyrels = marker = bad = 0
    sst_no_rels = 0
    for p in files:
        c = census(p)
        if "err" in c:
            bad += 1
            print("ERR\t-\t-\t-\t-\t-\t-\t%s\t%s" % (p, c["err"]))
            continue
        n += 1
        sst += c["sst"]; ts += c["ts"]; rels += c["rels"]; marker += c["marker"]
        anyrels += c["anyrels"]
        if c["sst"] and not c["rels"]:
            sst_no_rels += 1
        print("%d\t%d\t%d\t%d\t%d\t%d\t%d\t%s" % (c["sst"], c["ts"], c["rels"], c["anyrels"],
                                                  c["marker"], c["sheets"], c["size"], p))
    print()
    print("# files=%d unreadable=%d" % (n, bad))
    print("# sharedStrings part present:            %d" % sst)
    print("# first sheet has a t=\"s\" cell:          %d" % ts)
    print("# first sheet has worksheet relationships:%d" % rels)
    print("# any worksheet has relationships:        %d" % anyrels)
    print("# first sheet has mc/x14ac/dyDescent:     %d" % marker)
    print("# sst present AND no first-sheet rels:    %d" % sst_no_rels)

main()
