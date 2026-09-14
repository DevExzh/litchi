#!/usr/bin/env python3
"""How much of the BIFF8 globals substream does the source-backed XLS open
actually consume?

`parse_globals` (crates/litchi-xls/src/workbook/source.rs:1659) frames every
globals record and reads every globals byte into one retained buffer, then runs
a semantic pass that inspects the payload of a *fixed, small* set of record
kinds. This models the touched closure: header bytes (always needed to frame)
plus the payloads of the consumed kinds, against the bytes read today.

Deterministic. Standard library plus `olefile`. No repository code runs.
"""
import argparse, json, os, sys
from collections import Counter

import olefile

BOF, EOF, CONTINUE, SST = 0x0809, 0x000A, 0x003C, 0x00FC
BOUND_SHEET, CODEPAGE, FILEPASS = 0x0085, 0x0042, 0x002F

# Payload-consuming kinds, read out of the two match statements:
#   workbook/source.rs parse_globals  -> FILEPASS CODEPAGE BOUND_SHEET SST(+CONTINUE run)
#                                        and BOF/EOF framing checks
#   number_format/codec.rs Formatting::parse_globals
#       -> DATE1904 0x0022, FORMAT 0x041E, XF 0x00E0, XFCRC 0x087C,
#          XFExt 0x087D, DXF 0x088D
CONSUMED = {
    0x0809: "BOF", 0x000A: "EOF", 0x002F: "FilePass", 0x0042: "CodePage",
    0x0085: "BoundSheet8", 0x00FC: "SST",
    0x0022: "Date1904", 0x041E: "Format", 0x00E0: "XF", 0x087C: "XFCRC",
    0x087D: "XFExt", 0x088D: "DXF",
}
NAMES = dict(CONSUMED)
NAMES.update({0x003C: "Continue", 0x00EB: "MsoDrawingGroup", 0x00FF: "ExtSST",
              0x0293: "Style", 0x0892: "StyleExt", 0x0031: "Font",
              0x005C: "WriteAccess", 0x0018: "Lbl", 0x005A: "Crn",
              0x088E: "TableStyles", 0x0899: "Theme", 0x08C8: "PLV"})


def u16(b, o):
    return b[o] | (b[o + 1] << 8)


def frame(stream):
    out, off, n = [], 0, len(stream)
    while off + 4 <= n:
        kind = u16(stream, off)
        length = u16(stream, off + 2)
        end = off + 4 + length
        if end > n:
            return out, off, "truncated"
        out.append((off, kind, length))
        off = end
        if kind == EOF:
            return out, off, None
    return out, off, "no-eof"


def merge(spans, gap):
    """Coalesce [start,end) spans separated by fewer than `gap` bytes."""
    out = []
    for s, e in spans:
        if out and s - out[-1][1] < gap:
            out[-1][1] = max(out[-1][1], e)
        else:
            out.append([s, e])
    return out


def analyse(path, gaps):
    ole = olefile.OleFileIO(path)
    name = next((c for c in ("Workbook", "Book") if ole.exists(c)), None)
    if name is None:
        ole.close()
        return None
    data = ole.openstream(name).read()
    ole.close()
    records, globals_end, issue = frame(data)
    if not records or records[0][1] != BOF:
        return {"path": path, "error": "no BOF"}

    # Needed byte spans: every 4-byte header, plus consumed payloads.
    # A CONTINUE is consumed only when it continues an SST run.
    spans, prev_consumed_sst = [], False
    consumed_payload = skipped_payload = 0
    header_bytes = 4 * len(records)
    skipped_hist = Counter()
    for off, kind, length in records:
        spans.append([off, off + 4])
        if kind == SST:
            prev_consumed_sst = True
        elif kind != CONTINUE:
            prev_consumed_sst = False
        take = kind in CONSUMED or (kind == CONTINUE and prev_consumed_sst)
        if take and length:
            spans[-1][1] = off + 4 + length
            consumed_payload += length
        elif length:
            skipped_payload += length
            skipped_hist[kind] += length
    spans = merge(spans, 1)

    row = {
        "path": path,
        "file_size": os.path.getsize(path),
        "stream_len": len(data),
        "globals_end": globals_end,
        "globals_records": len(records),
        "issue": issue,
        "header_bytes": header_bytes,
        "consumed_payload_bytes": consumed_payload,
        "closure_bytes": header_bytes + consumed_payload,
        "skipped_payload_bytes": skipped_payload,
        "closure_share_of_globals": (header_bytes + consumed_payload) / globals_end
        if globals_end else None,
        "top_skipped": [[NAMES.get(k, f"0x{k:04X}"), v]
                        for k, v in sorted(skipped_hist.items(), key=lambda kv: -kv[1])[:5]],
    }
    for gap in gaps:
        m = merge([list(s) for s in spans], gap)
        row[f"reads_gap{gap}"] = len(m)
        row[f"bytes_gap{gap}"] = sum(e - s for s, e in m)
    return row


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default="/home/zhuhe/code/litchi")
    ap.add_argument("--out")
    ap.add_argument("--gaps", default="1,512,4096")
    args = ap.parse_args()
    gaps = [int(g) for g in args.gaps.split(",")]

    files = []
    for dirpath, _d, filenames in os.walk(os.path.join(args.repo, "test-data")):
        for fn in filenames:
            if fn.lower().endswith((".xls", ".xlt")):
                files.append(os.path.join(dirpath, fn))
    files.sort()

    rows = []
    for path in files:
        try:
            row = analyse(path, gaps)
        except Exception as exc:  # noqa: BLE001
            row = {"path": path, "error": repr(exc)}
        if row is None:
            continue
        row["path"] = os.path.relpath(row["path"], args.repo)
        rows.append(row)

    ok = [r for r in rows if "error" not in r and r.get("globals_end")]
    tot = lambda k: sum(r[k] for r in ok)
    summary = {
        "files": len(ok),
        "total_globals_bytes_read_today": tot("globals_end"),
        "total_header_bytes": tot("header_bytes"),
        "total_consumed_payload_bytes": tot("consumed_payload_bytes"),
        "total_closure_bytes": tot("closure_bytes"),
        "total_skipped_payload_bytes": tot("skipped_payload_bytes"),
        "closure_share": tot("closure_bytes") / tot("globals_end"),
    }
    for gap in gaps:
        summary[f"total_reads_gap{gap}"] = sum(r[f"reads_gap{gap}"] for r in ok)
        summary[f"total_bytes_gap{gap}"] = sum(r[f"bytes_gap{gap}"] for r in ok)
    doc = {"summary": summary, "consumed_kinds": {f"0x{k:04X}": v for k, v in CONSUMED.items()},
           "files": rows}
    if args.out:
        open(args.out, "w").write(json.dumps(doc, indent=2, sort_keys=True) + "\n")
    print(json.dumps(summary, indent=2))
    print()
    hdr = "%-56s %9s %9s %7s %6s %6s %9s" % (
        "path", "globals", "closure", "share", "r@512", "r@4k", "b@512")
    print(hdr)
    for r in sorted(ok, key=lambda r: -r["globals_end"])[:14]:
        print("%-56s %9d %9d %6.2f%% %6d %6d %9d" % (
            r["path"][-56:], r["globals_end"], r["closure_bytes"],
            100 * r["closure_share_of_globals"], r["reads_gap512"],
            r["reads_gap4096"], r["bytes_gap512"]))
    print()
    print("top skipped kinds, flagship:")
    for r in ok:
        if r["path"].endswith("ole/xls/ConditionalFormattingSamples.xls"):
            for n, v in r["top_skipped"]:
                print("   %-18s %10d" % (n, v))


if __name__ == "__main__":
    main()
