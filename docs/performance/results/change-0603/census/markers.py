#!/usr/bin/env python3
"""Marker census of every worksheet part in the real .xlsx fixture corpus.

Classifies each xl/worksheets/sheet*.xml part exactly as
`source_stream_eligible` (crates/litchi-xlsx/src/raw/worksheet/mod.rs) does,
and records which of the four marker gates each part trips, so the number of
parts admitted before and after a gate change can be counted rather than
estimated.
"""
import os, re, sys, zipfile

MCE = b"http://schemas.openxmlformats.org/markup-compatibility/2006"
X14AC = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
ALT = b"AlternateContent"
DYD = b"dyDescent"
MAX = 8 * 1024 * 1024

def classify(data):
    utf8 = True
    try:
        data.decode("utf-8")
    except UnicodeDecodeError:
        utf8 = False
    mce = MCE in data
    x14 = X14AC in data
    alt = ALT in data
    dyd = DYD in data
    # mc: prefixed usage beyond the declaration (prefix is conventionally mc,
    # but resolve it from the declaration to be exact).
    prefixes = set()
    for m in re.finditer(rb'xmlns:([A-Za-z0-9_.\-]+)\s*=\s*"' + re.escape(MCE) + rb'"', data):
        prefixes.add(m.group(1))
    x14prefixes = set()
    for m in re.finditer(rb'xmlns:([A-Za-z0-9_.\-]+)\s*=\s*"' + re.escape(X14AC) + rb'"', data):
        x14prefixes.add(m.group(1))
    mc_use = any(re.search(rb'[\s<]' + re.escape(p) + rb':', data) for p in prefixes)
    x14_use = any(re.search(rb'[\s<]' + re.escape(p) + rb':', data) for p in x14prefixes)
    return dict(utf8=utf8, size=len(data), mce=mce, x14=x14, alt=alt, dyd=dyd,
                mc_use=mc_use, x14_use=x14_use,
                eligible=(len(data) <= MAX and utf8 and not mce and not alt and not x14 and not dyd))

def main():
    roots = sys.argv[1:]
    files = []
    for root in roots:
        for d, _, fs in os.walk(root):
            for f in fs:
                if f.lower().endswith(".xlsx"):
                    files.append(os.path.join(d, f))
    files.sort()
    rows = []
    for p in files:
        try:
            z = zipfile.ZipFile(p)
        except Exception as e:
            print(f"# unreadable\t{p}\t{e}", file=sys.stderr)
            continue
        with z:
            names = [n for n in z.namelist() if re.match(r"xl/worksheets/sheet[^/]*\.xml$", n)]
            names.sort(key=lambda n: (len(n), n))
            for n in names:
                c = classify(z.read(n))
                c["path"] = p
                c["part"] = n
                rows.append(c)
    cols = ["eligible","utf8","mce","x14","alt","dyd","mc_use","x14_use","size","part","path"]
    print("\t".join(cols))
    for r in rows:
        print("\t".join(str(int(r[c])) if isinstance(r[c], bool) else str(r[c]) for c in cols))
    n = len(rows)
    def cnt(f): return sum(1 for r in rows if f(r))
    print(f"\n# worksheet parts: {n}", file=sys.stderr)
    print(f"# eligible today: {cnt(lambda r: r['eligible'])}", file=sys.stderr)
    print(f"# trip mce ns: {cnt(lambda r: r['mce'])}", file=sys.stderr)
    print(f"# trip x14ac ns: {cnt(lambda r: r['x14'])}", file=sys.stderr)
    print(f"# trip AlternateContent: {cnt(lambda r: r['alt'])}", file=sys.stderr)
    print(f"# trip dyDescent: {cnt(lambda r: r['dyd'])}", file=sys.stderr)
    print(f"# x14ac ns and no mce ns and no dyDescent and no alt: {cnt(lambda r: r['x14'] and not r['mce'] and not r['dyd'] and not r['alt'])}", file=sys.stderr)
    print(f"# mce ns, no alt: {cnt(lambda r: r['mce'] and not r['alt'])}", file=sys.stderr)
    print(f"# mce ns, no alt, no mc: usage: {cnt(lambda r: r['mce'] and not r['alt'] and not r['mc_use'])}", file=sys.stderr)
    print(f"# mce ns, no alt, no dyDescent: {cnt(lambda r: r['mce'] and not r['alt'] and not r['dyd'])}", file=sys.stderr)
    print(f"# would be eligible if only AlternateContent+size+utf8 gated: {cnt(lambda r: r['utf8'] and not r['alt'] and r['size'] <= MAX)}", file=sys.stderr)

main()
