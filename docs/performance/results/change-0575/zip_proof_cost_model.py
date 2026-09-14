#!/usr/bin/env python3
"""Positional-read cost model for the ZIP strict-layout proof candidates.

Pure-Python; reads the fixture bytes directly with `struct`.  It models the
number of `ReaderAt::read_exact_at` calls and the bytes those calls move, for
each candidate design in docs/performance/0575-zip-lazy-strict-layout-design.md.

Candidate key:
  d0  today's archive-wide proof              2 reads per central record
  d1  change 0573's archive-wide proof        1 read per central record
  e   coalesced archive-wide proof            1 read per header *run*
  a   zero-I/O pre-screen + target only       1 header, target only
  b   target + residual predecessors          1 + |residual| headers
  c   incremental memoised per-target proof   1 header per distinct member
"""
import os
import struct
import sys
import json

MAX_LOCAL_EXTRA = 65535
MAX_DESCRIPTOR = 24
RESIDUAL_WINDOW = MAX_LOCAL_EXTRA + MAX_DESCRIPTOR

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from zip_layout_census import parse_central  # noqa: E402

# One read of the 30-byte fixed local header, one of name+extra.
FIXED = 30


def header_bytes(e):
    """Bytes the strict path must read to resolve one entry's layout."""
    return FIXED + e["local_name_len"] + e["local_extra_len"]


def prepare(path):
    entries, cd_off, total = parse_central(path)
    for e in entries:
        e["min_end"] = (e["local_header_offset"] + FIXED + e["central_name_len"]
                        + e["compressed_size"])
        e["max_end"] = e["min_end"] + MAX_LOCAL_EXTRA + MAX_DESCRIPTOR
        e["hdr_bytes"] = header_bytes(e)
        e["hdr_end"] = e["local_header_offset"] + e["hdr_bytes"]
    return entries, cd_off, total


def residual_predecessors(entries, t):
    off_t = entries[t]["local_header_offset"]
    return [i for i in range(t)
            if entries[i]["max_end"] > off_t and entries[i]["min_end"] <= off_t]


def coalesce(entries, idxs, gap_threshold):
    """Group header reads into runs when the gap between the end of one header
    and the start of the next is at or below `gap_threshold`."""
    if not idxs:
        return 0, 0
    idxs = sorted(idxs, key=lambda i: entries[i]["local_header_offset"])
    reads = 0
    total = 0
    run_start = entries[idxs[0]]["local_header_offset"]
    run_end = entries[idxs[0]]["hdr_end"]
    for i in idxs[1:]:
        e = entries[i]
        if e["local_header_offset"] - run_end <= gap_threshold:
            run_end = max(run_end, e["hdr_end"])
        else:
            reads += 1
            total += run_end - run_start
            run_start = e["local_header_offset"]
            run_end = e["hdr_end"]
    reads += 1
    total += run_end - run_start
    return reads, total


def model(path, scenarios, gap_thresholds=(0, 512, 4096, 65536)):
    entries, cd_off, total_size = prepare(path)
    n = len(entries)
    by_name = {e["name"]: i for i, e in enumerate(entries)}
    all_idx = list(range(n))
    hdr_bytes_all = sum(e["hdr_bytes"] for e in entries)

    out = dict(path=path, entries=n, total_size=total_size,
               central_directory_offset=cd_off,
               header_bytes_all=hdr_bytes_all, scenarios={})

    # archive-wide variants are scenario-independent: they are paid once, in
    # full, on the first read of any member.
    d0 = dict(reads=2 * n, bytes=hdr_bytes_all)
    d1 = dict(reads=n, bytes=hdr_bytes_all)
    coalesced = {}
    for g in gap_thresholds:
        r, b = coalesce(entries, all_idx, g)
        coalesced[g] = dict(reads=r, bytes=b)
    out["archive_wide"] = dict(d0=d0, d1=d1, coalesced=coalesced)

    for label, names in scenarios.items():
        idxs = [by_name[nm] for nm in names if nm in by_name]
        missing = [nm for nm in names if nm not in by_name]
        # (a) zero-I/O screen, target-only validation, memoised
        a = dict(reads=len(idxs), bytes=sum(entries[i]["hdr_bytes"] for i in idxs))
        # (c) identical read cost to (a); they differ only in what is proven
        c = dict(a)
        # (b) target plus residual predecessors, memoised across the scenario
        touched = set()
        for t in idxs:
            touched.add(t)
            touched.update(residual_predecessors(entries, t))
        screen_only = touched - set(idxs)
        b = dict(reads=len(touched),
                 bytes=sum(entries[i]["hdr_bytes"] for i in touched),
                 headers=len(touched))
        # (b) refined: a non-target neighbour needs only the 30-byte fixed
        # local header to bound its span; name, extra, sizes and CRC are not
        # re-validated for an entry that is not being read.
        b_fixed = dict(reads=len(touched),
                       bytes=(sum(entries[i]["hdr_bytes"] for i in idxs)
                              + FIXED * len(screen_only)),
                       targets=len(idxs), screen_only=len(screen_only))
        out["scenarios"][label] = dict(
            members=len(idxs), missing=missing,
            a=a, b=b, b_fixed=b_fixed, c=c,
            per_target_residual={entries[t]["name"]: len(residual_predecessors(entries, t))
                                 for t in idxs},
        )
    return out


SCENARIOS = {
    "xlsx": {
        "one member (sheet1)": ["xl/worksheets/sheet1.xml"],
        "one-cell closure": [
            "[Content_Types].xml", "_rels/.rels",
            "xl/_rels/workbook.xml.rels", "xl/workbook.xml",
            "xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "xl/styles.xml",
        ],
    },
    "pptx": {
        "one member (slide1)": ["ppt/slides/slide1.xml"],
        "one-slide closure": [
            "[Content_Types].xml", "_rels/.rels",
            "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml",
            "ppt/slides/_rels/slide1.xml.rels", "ppt/slides/slide1.xml",
        ],
    },
}


def main():
    targets = [
        ("test-data/ooxml/xlsx/sheet-names.xlsx", "xlsx"),
        ("test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx", "xlsx"),
        ("test-data/ooxml/pptx/shapes.pptx", "pptx"),
    ]
    results = []
    for path, kind in targets:
        if not os.path.exists(path):
            print("MISSING", path, file=sys.stderr)
            continue
        results.append(model(path, SCENARIOS[kind]))

    for r in results:
        print(f"== {r['path']}  ({r['entries']} central records, {r['total_size']} bytes)")
        aw = r["archive_wide"]
        print(f"   archive-wide proof, paid once on the first read of any member:")
        print(f"     (d0) today, 2 reads/record       : {aw['d0']['reads']:>5} reads  {aw['d0']['bytes']:>8} B")
        print(f"     (d1) change 0573, 1 read/record  : {aw['d1']['reads']:>5} reads  {aw['d1']['bytes']:>8} B")
        for g, v in sorted(aw["coalesced"].items()):
            print(f"     (e)  coalesced, gap<= {g:>5} B    : {v['reads']:>5} reads  {v['bytes']:>8} B")
        for label, s in r["scenarios"].items():
            if s["members"] == 0:
                print(f"   scenario {label!r}: no members matched ({s['missing']})")
                continue
            print(f"   scenario {label!r} ({s['members']} members read; missing={s['missing']})")
            print(f"     (a/c) target-only, memoised      : {s['a']['reads']:>5} reads  {s['a']['bytes']:>8} B")
            print(f"     (b)   target + residual preds    : {s['b']['reads']:>5} reads  {s['b']['bytes']:>8} B  (full headers)")
            bf = s['b_fixed']
            print(f"     (b')  same, 30-byte neighbour screen: {bf['reads']:>4} reads  {bf['bytes']:>8} B  "
                  f"({bf['targets']} targets + {bf['screen_only']} screened)")
            print(f"           per-target residual        : {s['per_target_residual']}")
        print()

    dest = os.environ.get("MODEL_JSON")
    if dest:
        with open(dest, "w") as fh:
            json.dump(results, fh, indent=1)
        print("wrote", dest)


if __name__ == "__main__":
    main()
