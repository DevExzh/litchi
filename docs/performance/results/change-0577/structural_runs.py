#!/usr/bin/env python3
"""Model a run-coalesced structural prefetch for an OOXML open.

Today `read_structural_member` costs one 30-byte local-header read plus one
payload read per structural member (plus a 16-byte descriptor read where the
member carries one). This script asks what it would cost if the structural
members that are physically ADJACENT in the archive were fetched by one
bounded read per contiguous run, and how many bytes that would request.

It reads only the ZIP; it asks the library nothing. The structural set is the
rule change 0572 measured: `[Content_Types].xml`, every `*/_rels/*.rels`, and
the format's main part for XLSX and PPTX.
"""
import sys, os, json, struct, zipfile, posixpath

WINDOW_CAP = 64 * 1024   # litchi_opc MAX_SOURCE_READ_AHEAD_BYTES


def is_rels(name):
    if not name.lower().endswith(".rels"):
        return False
    d = posixpath.dirname(name)
    return d == "_rels" or d.endswith("/_rels")


def records(path):
    out = []
    with zipfile.ZipFile(path) as z, open(path, "rb") as f:
        for info in z.infolist():
            off = info.header_offset
            f.seek(off)
            hdr = f.read(30)
            if hdr[:4] != b"PK\x03\x04":
                raise SystemExit(f"{path}: bad local header for {info.filename}")
            n, e = struct.unpack("<HH", hdr[26:30])
            start = off + 30 + n + e
            desc = 16 if (info.flag_bits & 0x08) else 0
            out.append({
                "name": info.filename, "local_off": off,
                "payload_start": start,
                "end": start + info.compress_size + desc,
                "desc": bool(desc),
            })
    out.sort(key=lambda r: r["local_off"])
    return out


def analyse(path):
    recs = records(path)
    names = {r["name"] for r in recs}
    main = names & {"xl/workbook.xml", "ppt/presentation.xml"}
    def structural(r):
        return (r["name"].lower() == "[content_types].xml"
                or is_rels(r["name"]) or r["name"] in main)
    sel = [r for r in recs if structural(r)]
    pos = {r["name"]: i for i, r in enumerate(recs)}

    # Contiguous runs: structural members that are consecutive in physical order
    # AND leave no byte gap between one record's end and the next one's header.
    runs, cur = [], []
    for r in sel:
        if cur and pos[r["name"]] == pos[cur[-1]["name"]] + 1 \
                and r["local_off"] == cur[-1]["end"]:
            cur.append(r)
        else:
            if cur:
                runs.append(cur)
            cur = [r]
    if cur:
        runs.append(cur)

    # One read per run, each read capped at the existing 64 KiB window bound.
    reads_capped, bytes_capped = 0, 0
    for run in runs:
        span = run[-1]["end"] - run[0]["local_off"]
        reads_capped += max(1, -(-span // WINDOW_CAP))
        bytes_capped += span
    today_reads = sum(3 if r["desc"] else 2 for r in sel)
    # Today each structural member costs its 30-byte local header plus its
    # payload, and its 16-byte descriptor where it has one; the variable header
    # region between them is skipped. A run read covers the whole span instead.
    today_bytes = sum(30 + (r["end"] - r["payload_start"]) for r in sel)
    return {
        "fixture": os.path.basename(path),
        "size": os.path.getsize(path),
        "members": len(recs),
        "structural": len(sel),
        "descriptor_members": sum(1 for r in sel if r["desc"]),
        "runs": len(runs),
        "run_sizes": [len(r) for r in runs],
        "longest_run": max((len(r) for r in runs), default=0),
        "today_structural_reads": today_reads,
        "today_structural_bytes": today_bytes,
        "today_open_requests": today_reads + 3,
        "coalesced_structural_reads": reads_capped,
        "coalesced_open_requests": reads_capped + 3,
        "coalesced_bytes": bytes_capped,
        "structural_span": (sel[-1]["end"] - sel[0]["local_off"]) if sel else 0,
        "peak_run_bytes": max((run[-1]["end"] - run[0]["local_off"]
                               for run in runs), default=0),
    }


if __name__ == "__main__":
    rows = [analyse(p) for p in sys.argv[1:]]
    print(json.dumps(rows, indent=2))
