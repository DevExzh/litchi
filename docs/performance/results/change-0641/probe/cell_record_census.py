#!/usr/bin/env python3
"""Corpus census of the worksheet cell-record mix a source-backed query frames.

Change 0641 (item XLS-5 of change 0587).  `CellRecord::parse`
(crates/litchi-xls/src/records.rs) materializes `Label { value: String }` and
`Formula { formula: Vec<u8>, metadata, .. }` for *every* cell record a
worksheet scan frames, and a selected-cell query drops all but one of them.  How
much that costs depends entirely on the corpus mix, and change 0584's census
counted `LabelSst` only.  This answers the question it left open.

The CFB container walk and the `Workbook`/`Book` selection are taken unchanged
from change 0608's `results/change-0608/probe/sst_prefix_census.py`, which took
them from change 0584's `sst_walk.py`.  Pure arithmetic over the wire bytes; it
shares no code with `litchi-xls`, so it is an independent oracle rather than a
re-run of the scan it describes.

Per worksheet substream -- the unit a `query_cell` scan actually walks, from the
`BoundSheet8` position to its `EOF` -- it counts every record, every cell record
by kind, the packed cells inside `MulRk`/`MulBlank`, and the bytes that the
materializing parse copies onto the heap:

  label_bytes     `XLUnicodeString` character bytes each `Label` transcodes
  label_utf16     `Label` records whose `fHighByte` is set (the ones that run
                  `String::from_utf16` rather than a byte-to-char map)
  formula_tokens  `Rgce` bytes each `Formula` copies with `.to_vec()`
  formula_extra   `Formula` records carrying a nonempty `RgbExtra` suffix (the
                  only ones that also build an `Ancillary`)

Emitted as one JSON object per fixture on stdout, plus a corpus summary on
stderr.
"""

import json
import os
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
            out += ministream[s * mssz : (s + 1) * mssz]
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
            found[name] = (
                struct.unpack_from("<I", e, 116)[0],
                struct.unpack_from("<Q", e, 120)[0],
            )
    for name in ("Workbook", "Book"):
        if name in found:
            start, size = found[name]
            reader = mread if size < cutoff else sread
            return reader(start)[:size]
    raise ValueError("no Workbook stream")


BOF = 0x0809
EOF_ = 0x000A
# The seven single-cell kinds `CellRecord::parse` accepts, plus the two packed
# kinds the scan expands through `visit_mul_*`.
KINDS = {
    0x0201: "blank",
    0x0203: "number",
    0x0204: "label",
    0x0205: "boolerr",
    0x027E: "rk",
    0x00FD: "labelsst",
    0x0006: "formula",
    0x00BD: "mulrk",
    0x00BE: "mulblank",
}
FORMULA_FIXED_SIZE = 22


def new_counts():
    counts = {name: 0 for name in KINDS.values()}
    counts.update(
        records=0,
        payload_bytes=0,
        mulrk_cells=0,
        mulblank_cells=0,
        label_bytes=0,
        label_utf16=0,
        formula_tokens=0,
        formula_extra=0,
        formula_pending=0,
    )
    return counts


def count_substream(wb, recs, first):
    """Counts one worksheet substream, from its BOF to the matching EOF."""
    counts = new_counts()
    for rid, ds, dl in recs[first:]:
        counts["records"] += 1
        counts["payload_bytes"] += dl
        if rid == EOF_:
            break
        name = KINDS.get(rid)
        if name is None:
            continue
        counts[name] += 1
        if rid == 0x0204 and dl >= 9:
            chars = struct.unpack_from("<H", wb, ds + 6)[0]
            high = wb[ds + 8] & 0x01
            counts["label_bytes"] += chars * (2 if high else 1)
            counts["label_utf16"] += 1 if high else 0
        elif rid == 0x0006 and dl >= FORMULA_FIXED_SIZE:
            token_len = struct.unpack_from("<H", wb, ds + 20)[0]
            counts["formula_tokens"] += token_len
            if FORMULA_FIXED_SIZE + token_len < dl:
                counts["formula_extra"] += 1
            # `parse_formula_value`: a string-valued Formula has 0xFF in the
            # two high bytes and 0x00 in the first byte of the cached value.
            if wb[ds + 12] == 0xFF and wb[ds + 13] == 0xFF and wb[ds + 6] == 0x00:
                counts["formula_pending"] += 1
        elif rid == 0x00BD and dl >= 6:
            counts["mulrk_cells"] += max(dl - 6, 0) // 6
        elif rid == 0x00BE and dl >= 6:
            counts["mulblank_cells"] += max(dl - 6, 0) // 2
    return counts


def analyze(path):
    wb = cfb_workbook(path)
    recs = []
    off = 0
    while off + 4 <= len(wb):
        rid, rlen = struct.unpack_from("<HH", wb, off)
        if off + 4 + rlen > len(wb):
            break
        recs.append((rid, off + 4, rlen))
        off += 4 + rlen
    by_offset = {ds - 4: k for k, (_rid, ds, _dl) in enumerate(recs)}

    sheets = []
    for rid, ds, dl in recs:
        if rid == 0x0085 and dl >= 6:
            pos = struct.unpack_from("<I", wb, ds)[0]
            kind = wb[ds + 5]
            if kind == 0:
                sheets.append(pos)

    per_sheet = []
    for pos in sheets:
        k = by_offset.get(pos)
        if k is None or recs[k][0] != BOF:
            per_sheet.append(None)
            continue
        per_sheet.append(count_substream(wb, recs, k))
    return per_sheet


def fold(rows):
    total = new_counts()
    for row in rows:
        for key, value in row.items():
            total[key] += value
    return total


def main():
    root = sys.argv[1] if len(sys.argv) > 1 else "test-data"
    paths = []
    for base, _dirs, names in os.walk(root):
        for name in sorted(names):
            if name.lower().endswith((".xls", ".xlt")):
                paths.append(os.path.join(base, name))
    paths.sort()

    corpus = new_counts()
    fixtures = 0
    refused = 0
    sheets_total = 0
    for path in paths:
        try:
            per_sheet = analyze(path)
        except Exception as error:  # noqa: BLE001 - a census reports, never fails
            refused += 1
            print(json.dumps({"fixture": path, "error": str(error)}))
            continue
        usable = [sheet for sheet in per_sheet if sheet is not None]
        folded = fold(usable)
        fixtures += 1
        sheets_total += len(usable)
        for key, value in folded.items():
            corpus[key] += value
        print(
            json.dumps(
                {
                    "fixture": path,
                    "sheets": len(per_sheet),
                    "usable_sheets": len(usable),
                    "max_sheet_records": max((s["records"] for s in usable), default=0),
                    "total": folded,
                    "per_sheet": per_sheet,
                },
                sort_keys=True,
            )
        )

    summary = {
        "fixtures": fixtures,
        "unreadable": refused,
        "worksheet_substreams": sheets_total,
        "corpus": corpus,
    }
    cells = sum(
        corpus[name] for name in ("blank", "number", "label", "boolerr", "rk", "labelsst", "formula")
    ) + corpus["mulrk_cells"] + corpus["mulblank_cells"]
    summary["cell_values"] = cells
    if cells:
        summary["share_label_pct"] = round(100.0 * corpus["label"] / cells, 4)
        summary["share_formula_pct"] = round(100.0 * corpus["formula"] / cells, 4)
        summary["share_allocating_pct"] = round(
            100.0 * (corpus["label"] + corpus["formula"]) / cells, 4
        )
    print(json.dumps(summary, sort_keys=True), file=sys.stderr)


if __name__ == "__main__":
    main()
