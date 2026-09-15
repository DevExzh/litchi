#!/usr/bin/env python3
"""Corpus census of what a *prefix* shared-string index would have to walk.

Change 0608 (item XLS-2 of change 0587).  A deferred SST index is a prefix
index: resolving entry k requires walking entries 0..k, because a shared
string's extent is only known after every earlier string has been framed.  So
the question "does deferral pay?" is not "how many strings does the scenario
resolve?" but "what is the highest index it resolves?".

This walks the CFB container, the BIFF8 framing, the SST header and every
`LabelSst` record of each fixture with pure arithmetic -- it shares no code with
`litchi-xls`, so it is an independent oracle rather than a re-run of the scan it
describes.

The CFB/BIFF/SST framing walk is adapted from change 0584's
`results/change-0584/analysis/sst_walk.py`, which answered a different question
(sector ordinals for the chain hint).  Standard library only.

Emitted per fixture, as one JSON object per line:

  unique          `cstUnique`, the entries an eager index records at open
  total           `cstTotal`
  sst_bytes       logical length of the SST + `Continue` payloads
  segments        number of SST/`Continue` records
  labelsst        `LabelSst` cells -- what a full text resolves
  distinct        distinct SST indices those cells reference
  idx_max         highest index referenced by any cell
  prefix_text     idx_max + 1: entries a prefix index walks for a full text
  prefix_mean     mean(idx + 1) over every `LabelSst` cell: the expected walk
                  for one uniformly chosen string cell
  prefix_first    first_idx + 1: the walk for the first string cell in stream
                  order, which is what a top-left one-cell query hits
  saved_text      1 - prefix_text / unique
  saved_mean      1 - prefix_mean / unique
"""

import json
import os
import statistics
import struct
import sys


def cfb_workbook(path):
    d = open(path, "rb").read()
    ssz = 1 << struct.unpack_from("<H", d, 30)[0]
    mssz = 1 << struct.unpack_from("<H", d, 32)[0]
    nfat = struct.unpack_from("<I", d, 44)[0]
    dirstart = struct.unpack_from("<I", d, 48)[0]
    cutoff = struct.unpack_from("<I", d, 56)[0]
    ministart = struct.unpack_from("<I", d, 60)[0]
    nmini = struct.unpack_from("<I", d, 64)[0]
    difstart = struct.unpack_from("<I", d, 68)[0]
    ndif = struct.unpack_from("<I", d, 72)[0]
    fatsect = [struct.unpack_from("<I", d, 76 + 4 * i)[0] for i in range(109)]
    s = difstart
    while ndif > 0 and s < 0xFFFFFFFA:
        off = (s + 1) * ssz
        for i in range((ssz // 4) - 1):
            fatsect.append(struct.unpack_from("<I", d, off + 4 * i)[0])
        s = struct.unpack_from("<I", d, off + ssz - 4)[0]
        ndif -= 1
    fat = []
    for fs in fatsect[:nfat]:
        if fs >= 0xFFFFFFFA:
            continue
        off = (fs + 1) * ssz
        for i in range(ssz // 4):
            fat.append(struct.unpack_from("<I", d, off + 4 * i)[0])

    def chain(start):
        out = []
        s = start
        while s < 0xFFFFFFFA and len(out) < 4_000_000:
            out.append(s)
            s = fat[s] if s < len(fat) else 0xFFFFFFFE
        return out

    def sread(start):
        return b"".join(d[(s + 1) * ssz : (s + 2) * ssz] for s in chain(start))

    # The mini stream. A `Workbook` stream shorter than the header's cutoff --
    # 4,096 bytes in every fixture here -- lives in mini sectors inside the root
    # entry's stream, not in ordinary sectors, and reading it through the FAT
    # yields garbage.  Eleven of this repository's SST-bearing fixtures are that
    # small, so the census is not corpus-wide without this.
    dd = sread(dirstart)
    root = dd[:128]
    ministream = sread(struct.unpack_from("<I", root, 116)[0])
    minifat = []
    for sector in chain(ministart) if ministart < 0xFFFFFFFA and nmini else []:
        off = (sector + 1) * ssz
        for i in range(ssz // 4):
            minifat.append(struct.unpack_from("<I", d, off + 4 * i)[0])

    def mread(start):
        out = bytearray()
        s = start
        seen = 0
        while s < 0xFFFFFFFA and seen < 4_000_000:
            out += ministream[s * mssz:(s + 1) * mssz]
            s = minifat[s] if s < len(minifat) else 0xFFFFFFFE
            seen += 1
        return bytes(out)

    found = {}
    for i in range(len(dd) // 128):
        e = dd[i * 128 : (i + 1) * 128]
        nl = struct.unpack_from("<H", e, 64)[0]
        name = e[: max(nl - 2, 0)].decode("utf-16-le", "replace")
        typ = e[66]
        if typ == 2 and name in ("Workbook", "Book") and name not in found:
            found[name] = (struct.unpack_from("<I", e, 116)[0],
                           struct.unpack_from("<Q", e, 120)[0])
    # `select_workbook_stream` (crates/litchi-xls/src/workbook/source.rs) tries
    # "Workbook" and only then "Book". Seven fixtures here carry BOTH -- a
    # legacy BIFF5 `Book` beside the BIFF8 `Workbook` -- and `Book` comes first
    # in directory order, so taking whichever appears first reads the BIFF5
    # stream, which has no SST at all.
    for name in ("Workbook", "Book"):
        if name in found:
            start, size = found[name]
            reader = mread if size < cutoff else sread
            return reader(start)[:size], ssz
    raise ValueError("no Workbook stream")


def analyze(path):
    wb, _ssz = cfb_workbook(path)
    recs = []
    off = 0
    while off + 4 <= len(wb):
        rid, rlen = struct.unpack_from("<HH", wb, off)
        recs.append((rid, off + 4, rlen))
        off += 4 + rlen

    segs = []
    logical = 0
    i = 0
    while i < len(recs) and recs[i][0] != 0x00FC:
        i += 1
    if i == len(recs):
        return None
    j = i
    while j < len(recs) and (j == i or recs[j][0] == 0x003C):
        _, ds, dl = recs[j]
        segs.append((ds, logical, dl))
        logical += dl
        j += 1

    def rd(lo, n):
        out = bytearray()
        for so, llo, ln in segs:
            if lo < llo + ln and lo + n > llo:
                a = max(lo, llo)
                b = min(lo + n, llo + ln)
                out += wb[so + (a - llo) : so + (b - llo)]
        return bytes(out)

    hdr = rd(0, 8)
    if len(hdr) < 8:
        return None
    total, unique = struct.unpack("<II", hdr)
    # The header checks the production scan applies before it walks a string
    # (`scan_shared_string_records`, crates/litchi-xls/src/records.rs).  A
    # fixture that fails one of them is refused at open today and is refused at
    # open under every design in change 0608, so it is reported and excluded
    # from the aggregates rather than counted as a saving.
    refused = None
    if total > 0x7FFFFFFF or unique > 0x7FFFFFFF:
        refused = "SST counts must be non-negative signed integers"
    elif total < unique:
        refused = "SST total count is smaller than its unique count"
    elif unique > max(logical - 8, 0) // 3:
        refused = "SST declares %d strings but its records are too short" % unique
    idxs = []
    for rid, ds, dl in recs:
        if rid == 0x00FD and dl >= 10:
            idxs.append(struct.unpack_from("<I", wb, ds + 6)[0])
    inrange = [x for x in idxs if x < unique]

    # The worksheet substreams, in `BoundSheet8` order, so that the index the
    # harness's one-cell selector actually resolves can be named.  The harness
    # reads (`--worksheet-index` w, row 1, column 0) and counts only sheets
    # whose type is 0, which is what `SheetKind::WorksheetOrDialog` selects.
    starts = []
    for rid, ds, dl in recs:
        if rid == 0x0085 and dl >= 6:
            pos = struct.unpack_from("<I", wb, ds)[0]
            kind = wb[ds + 5]
            if kind == 0:
                starts.append(pos)
    by_offset = {ds - 4: k for k, (_rid, ds, _dl) in enumerate(recs)}
    one_cell = []
    for pos in starts:
        k = by_offset.get(pos)
        hit = None
        if k is not None:
            for rid, ds, dl in recs[k + 1:]:
                if rid == 0x000A:  # EOF ends the substream
                    break
                if rid == 0x00FD and dl >= 10:
                    row, col = struct.unpack_from("<HH", wb, ds)
                    if row == 1 and col == 0:
                        hit = struct.unpack_from("<I", wb, ds + 6)[0]
                        break
        one_cell.append(hit)

    row = dict(
        file=os.path.basename(path),
        path=path,
        refused=refused,
        unique=unique,
        total=total,
        sst_bytes=logical,
        segments=len(segs),
        labelsst=len(idxs),
        in_range=len(inrange),
        distinct=len(set(inrange)),
        worksheets=len(starts),
        one_cell_isst=one_cell,
    )
    if inrange:
        prefix_text = max(inrange) + 1
        prefix_mean = statistics.fmean(x + 1 for x in inrange)
        row.update(
            idx_min=min(inrange),
            idx_max=max(inrange),
            prefix_text=prefix_text,
            prefix_mean=round(prefix_mean, 2),
            prefix_first=inrange[0] + 1,
            saved_text=round(100.0 * (1.0 - prefix_text / unique), 2) if unique else 0.0,
            saved_mean=round(100.0 * (1.0 - prefix_mean / unique), 2) if unique else 0.0,
        )
    else:
        row.update(
            idx_min=None,
            idx_max=None,
            prefix_text=0,
            prefix_mean=0.0,
            prefix_first=0,
            saved_text=100.0 if unique else 0.0,
            saved_mean=100.0 if unique else 0.0,
        )
    return row


def main(argv):
    rows = []
    for p in argv:
        try:
            r = analyze(p)
        except Exception as exc:  # noqa: BLE001 - a census, not a parser
            print(json.dumps({"file": os.path.basename(p), "skip": f"{type(exc).__name__}: {exc}"}))
            continue
        if r is None:
            print(json.dumps({"file": os.path.basename(p), "skip": "no SST"}))
            continue
        rows.append(r)
        print(json.dumps(r))
    refused_count = sum(1 for r in rows if r["refused"])
    if not rows:
        return
    rows = [r for r in rows if not r["refused"]]
    with_strings = [r for r in rows if r["unique"]]
    resolving = [r for r in with_strings if r["in_range"]]
    print(json.dumps({
        "summary": True,
        "fixtures_with_sst": len(rows),
        "fixtures_refused_at_open": refused_count,
        "fixtures_with_entries": len(with_strings),
        "fixtures_with_labelsst": len(resolving),
        "fixtures_no_labelsst": len(with_strings) - len(resolving),
        "entries_total": sum(r["unique"] for r in with_strings),
        "labelsst_total": sum(r["labelsst"] for r in with_strings),
        "prefix_text_total": sum(r["prefix_text"] for r in with_strings),
        "saved_text_corpus": round(
            100.0 * (1.0 - sum(r["prefix_text"] for r in with_strings)
                     / max(sum(r["unique"] for r in with_strings), 1)), 2),
        "saved_mean_median_over_fixtures": round(
            statistics.median([r["saved_mean"] for r in resolving]), 2) if resolving else None,
        "fixtures_where_text_saves_nothing": sum(
            1 for r in with_strings if r["prefix_text"] >= r["unique"]),
    }))


if __name__ == "__main__":
    main(sys.argv[1:])
