#!/usr/bin/env python3
"""Model the composition of the BIFF8 workbook-globals substream.

Answers: of the bytes a source-backed XLS open reads today, how many are
needed to identify the file, how many to list worksheets, and how many exist
only to build the shared-string index and the formatting tables.

Reads fixtures with `olefile`; no repository code is executed.
"""
import argparse, json, os, sys
from collections import Counter

try:
    import olefile
except ImportError:
    sys.exit("olefile required")

BOF = 0x0809
EOF = 0x000A
BOUNDSHEET = 0x0085
SST = 0x00FC
CONTINUE = 0x003C
CODEPAGE = 0x0042
FILEPASS = 0x002F
EXTSST = 0x00FF

# Records the semantic globals pass actually consumes, from
# crates/litchi-xls/src/workbook/source.rs parse_globals.
NAMES = {
    0x0809: "BOF", 0x000A: "EOF", 0x0085: "BoundSheet8", 0x00FC: "SST",
    0x003C: "Continue", 0x0042: "CodePage", 0x002F: "FilePass",
    0x00FF: "ExtSST", 0x00E0: "XF", 0x087D: "XFExt", 0x0293: "Style",
    0x0892: "StyleExt", 0x041E: "Format", 0x0031: "Font", 0x007D: "ColInfo",
    0x00FD: "LabelSst", 0x0018: "Lbl", 0x005C: "WriteAccess",
    0x0893: "StyleExt?", 0x08C8: "PLV", 0x0899: "Theme",
    0x089A: "MTRSettings", 0x089B: "CompressPictures", 0x089C: "ForceFullCalculation",
    0x0862: "SheetExt", 0x0863: "BookExt", 0x0161: "DSF", 0x01C0: "ExcelServerstuff",
    0x00EB: "MsoDrawingGroup", 0x00EC: "MsoDrawing", 0x088E: "TableStyles",
    0x008C: "Country", 0x0014: "Header", 0x0015: "Footer",
}


def u16(b, o):
    return b[o] | (b[o + 1] << 8)


def walk(stream):
    """Frame the globals substream; return the record list and the globals end."""
    out = []
    off = 0
    n = len(stream)
    depth = 0
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


def analyse(path):
    try:
        ole = olefile.OleFileIO(path)
    except Exception as exc:  # noqa: BLE001
        return {"path": path, "error": f"cfb: {exc}"}
    name = None
    for candidate in ("Workbook", "Book"):
        if ole.exists(candidate):
            name = candidate
            break
    if name is None:
        ole.close()
        return None
    data = ole.openstream(name).read()
    stream_len = len(data)
    ole.close()

    records, globals_end, issue = walk(data)
    if not records or records[0][1] != BOF:
        return {"path": path, "error": "no BOF"}

    # Byte offset past the last BoundSheet8 record in the globals.
    last_bs_end = None
    first_bs_off = None
    bs_count = 0
    for off, kind, length in records:
        if kind == BOUNDSHEET:
            bs_count += 1
            if first_bs_off is None:
                first_bs_off = off
            last_bs_end = off + 4 + length

    # SST + its Continue run.
    sst_bytes = 0
    sst_frames = 0
    sst_start = None
    sst_end = None
    in_sst = False
    for off, kind, length in records:
        if kind == SST:
            in_sst = True
            sst_start = off
        elif in_sst and kind != CONTINUE:
            in_sst = False
        if in_sst:
            sst_bytes += 4 + length
            sst_frames += 1
            sst_end = off + 4 + length

    hist = Counter()
    for off, kind, length in records:
        hist[kind] += 4 + length

    # A Continue record belongs to whichever record precedes its run, so
    # attribute its bytes there. This is what makes "the Continue bulk is the
    # drawing group, not the shared-string table" a measurement rather than an
    # inference.
    continued = Counter()
    previous = None
    for off, kind, length in records:
        if kind == CONTINUE and previous is not None:
            continued[previous] += 4 + length
        elif kind != CONTINUE:
            previous = kind

    # Bytes needed for each scenario, as a *lower bound* on any scheme that
    # still frames records in order.
    identify_bytes = 0                        # CFB directory alone suffices
    list_bytes = last_bs_end or globals_end    # must frame through the last BoundSheet8
    return {
        "path": path,
        "file_size": os.path.getsize(path),
        "stream": name,
        "stream_len": stream_len,
        "globals_end": globals_end,
        "globals_records": len(records),
        "issue": issue,
        "boundsheet8_count": bs_count,
        "first_boundsheet8_offset": first_bs_off,
        "last_boundsheet8_end": last_bs_end,
        "list_prefix_bytes": list_bytes,
        "list_prefix_share": (list_bytes / globals_end) if globals_end else None,
        "sst_frames": sst_frames,
        "sst_bytes": sst_bytes,
        "sst_start": sst_start,
        "sst_end": sst_end,
        "sst_share_of_globals": (sst_bytes / globals_end) if globals_end else None,
        "bytes_after_last_boundsheet8": globals_end - (last_bs_end or globals_end),
        "continuation_bytes_by_continued_record": [
            [NAMES.get(k, f"0x{k:04X}"), v]
            for k, v in sorted(continued.items(), key=lambda kv: -kv[1])[:4]
        ],
        "top_record_bytes": [
            [NAMES.get(k, f"0x{k:04X}"), v]
            for k, v in sorted(hist.items(), key=lambda kv: -kv[1])[:8]
        ],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default="/home/zhuhe/code/litchi")
    ap.add_argument("--out")
    args = ap.parse_args()

    roots = [os.path.join(args.repo, "test-data")]
    files = []
    for root in roots:
        for dirpath, _dirnames, filenames in os.walk(root):
            for fn in filenames:
                if fn.lower().endswith((".xls", ".xlt")):
                    files.append(os.path.join(dirpath, fn))
    files.sort()

    rows = []
    for path in files:
        try:
            row = analyse(path)
        except Exception as exc:  # noqa: BLE001
            row = {"path": path, "error": repr(exc)}
        if row is None:
            continue
        row["path"] = os.path.relpath(row["path"], args.repo)
        rows.append(row)

    ok = [r for r in rows if "error" not in r and r.get("globals_end")]
    total_globals = sum(r["globals_end"] for r in ok)
    total_list = sum(r["list_prefix_bytes"] for r in ok)
    total_sst = sum(r["sst_bytes"] for r in ok)
    summary = {
        "files_scanned": len(rows),
        "files_with_globals": len(ok),
        "total_globals_bytes": total_globals,
        "total_list_prefix_bytes": total_list,
        "list_prefix_share": total_list / total_globals if total_globals else None,
        "total_sst_bytes": total_sst,
        "sst_share": total_sst / total_globals if total_globals else None,
        "files_where_sst_over_half_of_globals": sum(
            1 for r in ok if (r["sst_share_of_globals"] or 0) > 0.5
        ),
    }
    doc = {"summary": summary, "files": rows}
    text = json.dumps(doc, indent=2, sort_keys=True)
    if args.out:
        with open(args.out, "w") as fh:
            fh.write(text + "\n")
    print(json.dumps(summary, indent=2))
    top = sorted(ok, key=lambda r: -r["globals_end"])[:12]
    print("\n%-62s %10s %10s %8s %10s %7s" % ("path", "globals", "list<=", "share", "sst", "sst%"))
    for r in top:
        print("%-62s %10d %10d %7.2f%% %10d %6.1f%%" % (
            r["path"][-62:], r["globals_end"], r["list_prefix_bytes"],
            100 * r["list_prefix_share"], r["sst_bytes"],
            100 * (r["sst_share_of_globals"] or 0)))


if __name__ == "__main__":
    main()
