#!/usr/bin/env python3
"""Byte-level witness for the semantic delta in change 0575's candidate table.

Builds a three-member ZIP in which member A's *local extra field* is inflated
so that A's declared local span physically covers member B's complete local
record.  B's local record is embedded inside that extra field; C follows A's
span cleanly.  Every field the central directory carries is self-consistent,
local and central names and sizes agree for all three members, and each member
validates in isolation, so no central-directory-only analysis can detect the
overlap.  Only A's *local* header reveals it, and the local extra-field length
is the one span input the central directory never carries.

Prints the offsets, the central-directory-derived bounds, and which candidate
design refuses which read.  Writes the archive only when --out is given.
"""
import argparse
import struct
import zlib

MAX_LOCAL_EXTRA = 65535
MAX_DESCRIPTOR = 24
RESIDUAL_WINDOW = MAX_LOCAL_EXTRA + MAX_DESCRIPTOR
FIXED = 30
PAD_EXTRA_ID = 0xFACE  # unknown, therefore ignored, extra-field header id


def local_record(name: bytes, extra: bytes, payload: bytes) -> bytes:
    return (struct.pack("<IHHHHHIIIHH", 0x04034B50, 20, 0, 0, 0, 0,
                        zlib.crc32(payload) & 0xFFFFFFFF,
                        len(payload), len(payload), len(name), len(extra))
            + name + extra + payload)


def central_record(name: bytes, payload: bytes, offset: int) -> bytes:
    return (struct.pack("<IHHHHHHIIIHHHHHII", 0x02014B50, 20, 20, 0, 0, 0, 0,
                        zlib.crc32(payload) & 0xFFFFFFFF,
                        len(payload), len(payload), len(name), 0, 0, 0, 0, 0,
                        offset) + name)


def build(extra_len=4096, embed_at=2048):
    names = [b"A.bin", b"B.bin", b"C.bin"]
    payloads = [b"A" * 32, b"B" * 32, b"C" * 32]

    extra = struct.pack("<HH", PAD_EXTRA_ID, extra_len - 4) + b"\0" * (extra_len - 4)
    a_record = bytearray(local_record(names[0], extra, payloads[0]))
    a_off = 0
    a_end = len(a_record)

    b_record = local_record(names[1], b"", payloads[1])
    b_off = embed_at
    assert b_off + len(b_record) <= a_end, "B must fit inside A's span"
    a_record[b_off:b_off + len(b_record)] = b_record

    out = bytearray(a_record)
    c_off = len(out)
    out += local_record(names[2], b"", payloads[2])

    cd_off = len(out)
    for name, payload, off in zip(names, payloads, (a_off, b_off, c_off)):
        out += central_record(name, payload, off)
    cd_size = len(out) - cd_off
    out += struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, 3, 3, cd_size, cd_off, 0)
    return bytes(out), names, payloads, [a_off, b_off, c_off], cd_off


def spans(data, names, payloads, offsets):
    rows = []
    for name, payload, off in zip(names, payloads, offsets):
        elen = struct.unpack_from("<H", data, off + 28)[0]
        nlen = struct.unpack_from("<H", data, off + 26)[0]
        rows.append(dict(
            name=name.decode(), off=off, csize=len(payload), local_extra=elen,
            exact_end=off + FIXED + nlen + elen + len(payload),
            cd_min_end=off + FIXED + len(name) + len(payload),
            cd_max_end=off + FIXED + len(name) + len(payload) + RESIDUAL_WINDOW,
        ))
    return rows


def overlaps(x, y):
    """Do spans x and y intersect, using exact ends?"""
    return x["off"] < y["exact_end"] and y["off"] < x["exact_end"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--extra-len", type=int, default=4096)
    ap.add_argument("--embed-at", type=int, default=2048)
    ap.add_argument("--out")
    args = ap.parse_args()

    data, names, payloads, offsets, cd_off = build(args.extra_len, args.embed_at)
    rows = spans(data, names, payloads, offsets)

    print(f"archive bytes            : {len(data)}")
    print(f"central directory offset : {cd_off}")
    print()
    print(f"{'member':<8} {'local_off':>10} {'csize':>7} {'local_extra':>12} "
          f"{'exact_end':>10} {'cd_min_end':>11} {'cd_max_end':>11}")
    for r in rows:
        print(f"{r['name']:<8} {r['off']:>10} {r['csize']:>7} {r['local_extra']:>12} "
              f"{r['exact_end']:>10} {r['cd_min_end']:>11} {r['cd_max_end']:>11}")
    print()
    for i in range(3):
        for j in range(i + 1, 3):
            verdict = "OVERLAP" if overlaps(rows[i], rows[j]) else "disjoint"
            print(f"  ({rows[i]['name']}, {rows[j]['name']}): {verdict}")
    print()

    print("central-directory-only adjacency verdicts (zero I/O):")
    for i in range(2):
        nxt = rows[i + 1]["off"]
        if rows[i]["cd_max_end"] <= nxt:
            v = "non-overlap PROVEN"
        elif rows[i]["cd_min_end"] > nxt:
            v = "overlap REFUTED"
        else:
            v = "UNDECIDED (only this entry's local header can decide)"
        print(f"  ({rows[i]['name']}, {rows[i+1]['name']}): {v}")
    print()

    print("read outcomes, by candidate and by target:")
    hdr = f"  {'target':<8} {'today/(d)/(e)':<16} {'(a)':<12} {'(b)':<12} {'(c) alone':<12}"
    print(hdr)
    archive_wide_ok = all(
        not overlaps(rows[i], rows[j]) for i in range(3) for j in range(i + 1, 3))
    for t in range(3):
        # (a): target's own layout plus the zero-I/O screen only
        a_ok = all(rows[i]["cd_min_end"] <= rows[t]["off"]
                   for i in range(t))
        # (b): target's span must be disjoint from every other declared span
        b_ok = all(not overlaps(rows[t], rows[j]) for j in range(3) if j != t)
        # (c) with an empty memo: identical to (a)
        c_ok = a_ok
        def s(x):
            return "accept" if x else "REFUSE"
        print(f"  {rows[t]['name']:<8} {s(archive_wide_ok):<16} {s(a_ok):<12} "
              f"{s(b_ok):<12} {s(c_ok):<12}")
    print()
    print("(c) is order dependent: reading A then C accepts C, but reading B then C")
    print("    still accepts C, while reading A then B refuses B.  (b) and (a) are")
    print("    order independent: each target's verdict is a function of the bytes alone.")

    if args.out:
        with open(args.out, "wb") as fh:
            fh.write(data)
        print(f"\nwrote {args.out}")


if __name__ == "__main__":
    main()
