#!/usr/bin/env python3
"""Census of ZIP local-header geometry for the OOXML fixtures used by the
0575 lazy strict-layout design.

Reads only the archive bytes with `struct`; no cargo, no library under test.
For every central directory entry it records:
  - local_header_offset, compressed_size, central/local name length,
  - the ACTUAL local extra-field length (read from the local header),
  - whether the general-purpose bit 3 (data descriptor) is set.

It then answers the question the design needs: for a single-member read,
how many *predecessor* local headers must be touched to prove that no other
entry's span can reach the target, given that the only unknown in a
central-directory-derived span bound is the local extra length (<= 65535)
plus a data descriptor (<= 24 bytes)?
"""
import os
import struct
import sys
import json

EOCD_SIG = b"PK\x05\x06"
EOCD64_LOC_SIG = b"PK\x06\x07"
EOCD64_SIG = b"PK\x06\x06"
CEN_SIG = b"PK\x01\x02"
LOC_SIG = b"PK\x03\x04"

MAX_LOCAL_EXTRA = 65535       # u16 field
MAX_DESCRIPTOR = 24           # ZIP64 descriptor with signature
RESIDUAL_WINDOW = MAX_LOCAL_EXTRA + MAX_DESCRIPTOR  # 65559


def find_eocd(buf):
    start = max(0, len(buf) - 65536 - 22)
    idx = buf.rfind(EOCD_SIG, start)
    if idx < 0:
        raise ValueError("no EOCD")
    return idx


def parse_central(path):
    with open(path, "rb") as fh:
        buf = fh.read()
    eocd = find_eocd(buf)
    (_, _, _, total, cd_size, cd_off, _) = struct.unpack_from("<HHHHIIH", buf, eocd + 4)
    # ZIP64 locator
    loc = eocd - 20
    if loc >= 0 and buf[loc:loc + 4] == EOCD64_LOC_SIG:
        (_, eocd64_off, _) = struct.unpack_from("<IQI", buf, loc + 4)
        if buf[eocd64_off:eocd64_off + 4] == EOCD64_SIG:
            total = struct.unpack_from("<Q", buf, eocd64_off + 32)[0]
            cd_size = struct.unpack_from("<Q", buf, eocd64_off + 40)[0]
            cd_off = struct.unpack_from("<Q", buf, eocd64_off + 48)[0]
    entries = []
    p = cd_off
    for _ in range(total):
        assert buf[p:p + 4] == CEN_SIG, f"bad central sig at {p}"
        (ver, vneed, flags, method, mtime, mdate, crc, csize, usize,
         nlen, elen, clen, disk, iattr, eattr, loff) = struct.unpack_from(
            "<HHHHHHIIIHHHHHII", buf, p + 4)
        name = buf[p + 46:p + 46 + nlen]
        extra = buf[p + 46 + nlen:p + 46 + nlen + elen]
        # resolve zip64 extra
        z = extra
        while len(z) >= 4:
            hid, hsz = struct.unpack_from("<HH", z, 0)
            body = z[4:4 + hsz]
            if hid == 0x0001:
                q = 0
                if usize == 0xFFFFFFFF and len(body) >= q + 8:
                    usize = struct.unpack_from("<Q", body, q)[0]; q += 8
                if csize == 0xFFFFFFFF and len(body) >= q + 8:
                    csize = struct.unpack_from("<Q", body, q)[0]; q += 8
                if loff == 0xFFFFFFFF and len(body) >= q + 8:
                    loff = struct.unpack_from("<Q", body, q)[0]; q += 8
            z = z[4 + hsz:]
        # actual local header
        assert buf[loff:loff + 4] == LOC_SIG, f"bad local sig at {loff}"
        (lvneed, lflags, lmethod, lmtime, lmdate, lcrc, lcsize, lusize,
         lnlen, lelen) = struct.unpack_from("<HHHHHIIIHH", buf, loff + 4)
        entries.append(dict(
            name=name.decode("utf-8", "replace"),
            central_name_len=nlen,
            central_extra_len=elen,
            local_name_len=lnlen,
            local_extra_len=lelen,
            compressed_size=csize,
            uncompressed_size=usize,
            local_header_offset=loff,
            flags=flags,
            has_descriptor=bool(flags & 0x8),
            method=method,
        ))
        p += 46 + nlen + elen + clen
    entries.sort(key=lambda e: e["local_header_offset"])
    return entries, cd_off, len(buf)


def analyse(path):
    entries, cd_off, total_size = parse_central(path)
    n = len(entries)
    for e in entries:
        # exact span end, using the ACTUAL local extra length
        e["exact_end"] = (e["local_header_offset"] + 30 + e["local_name_len"]
                          + e["local_extra_len"] + e["compressed_size"])
        # central-directory-only bounds; the strict path forces
        # local_name_len == central_name_len
        e["min_end"] = (e["local_header_offset"] + 30 + e["central_name_len"]
                        + e["compressed_size"])
        e["max_end"] = e["min_end"] + MAX_LOCAL_EXTRA + MAX_DESCRIPTOR

    # Zero-I/O adjacency screen over the whole archive.
    zero_io_proven_pairs = 0
    zero_io_refuted_pairs = 0
    residual_pairs = 0
    for i in range(n - 1):
        nxt = entries[i + 1]["local_header_offset"]
        if entries[i]["max_end"] <= nxt:
            zero_io_proven_pairs += 1
        elif entries[i]["min_end"] > nxt:
            zero_io_refuted_pairs += 1
        else:
            residual_pairs += 1

    # Per-target residual: predecessors whose maximal span could reach the
    # target's local header offset and whose minimal span does not already
    # prove an overlap.
    residuals = []
    for t, tgt in enumerate(entries):
        off_t = tgt["local_header_offset"]
        need = 0
        refuted = 0
        for i in range(t):
            e = entries[i]
            if e["max_end"] <= off_t:
                continue
            if e["min_end"] > off_t:
                refuted += 1
                continue
            need += 1
        residuals.append(dict(index=t, name=tgt["name"], predecessors=t,
                              headers_needed=need, zero_io_refuted=refuted))

    extras = [e["local_extra_len"] for e in entries]
    cextras = [e["central_extra_len"] for e in entries]
    gaps = [entries[i + 1]["local_header_offset"] - entries[i]["exact_end"]
            for i in range(n - 1)]
    return dict(
        path=path, total_size=total_size, central_directory_offset=cd_off,
        entries=n,
        descriptors=sum(1 for e in entries if e["has_descriptor"]),
        stored=sum(1 for e in entries if e["method"] == 0),
        local_extra_min=min(extras), local_extra_max=max(extras),
        local_extra_distinct=sorted(set(extras)),
        central_extra_min=min(cextras), central_extra_max=max(cextras),
        central_extra_distinct=sorted(set(cextras)),
        gap_min=min(gaps) if gaps else None, gap_max=max(gaps) if gaps else None,
        gap_nonzero=sum(1 for g in gaps if g != 0) if gaps else 0,
        zero_io_proven_pairs=zero_io_proven_pairs,
        zero_io_refuted_pairs=zero_io_refuted_pairs,
        zero_io_residual_pairs=residual_pairs,
        residual_headers_max=max(r["headers_needed"] for r in residuals),
        residual_headers_mean=sum(r["headers_needed"] for r in residuals) / n,
        residual_headers_last=residuals[-1]["headers_needed"],
        residual_headers_first_content=residuals[0]["headers_needed"],
        per_target=residuals,
        entries_detail=entries,
    )


def main(paths):
    out = []
    for p in paths:
        if not os.path.exists(p):
            print(f"MISSING {p}", file=sys.stderr)
            continue
        out.append(analyse(p))
    for r in out:
        print(f"== {r['path']}")
        print(f"   size={r['total_size']}  central_directory_offset={r['central_directory_offset']}")
        print(f"   entries={r['entries']}  stored={r['stored']}  data_descriptors={r['descriptors']}")
        print(f"   local extra len: min={r['local_extra_min']} max={r['local_extra_max']} distinct={r['local_extra_distinct']}")
        print(f"   central extra len: min={r['central_extra_min']} max={r['central_extra_max']} distinct={r['central_extra_distinct']}")
        print(f"   inter-entry gap (exact end -> next offset): min={r['gap_min']} max={r['gap_max']} nonzero={r['gap_nonzero']}")
        print(f"   zero-I/O adjacency screen over whole archive: proven={r['zero_io_proven_pairs']} refuted={r['zero_io_refuted_pairs']} residual={r['zero_io_residual_pairs']} (of {r['entries']-1} pairs)")
        print(f"   per-target predecessor headers needed: max={r['residual_headers_max']} mean={r['residual_headers_mean']:.1f} last-entry={r['residual_headers_last']}")
        # show a few representative targets
        for tgt in r["per_target"]:
            if tgt["index"] in (0, 1, r["entries"] // 2, r["entries"] - 1):
                print(f"     target[{tgt['index']:>3}] {tgt['name'][:52]:<52} preds={tgt['predecessors']:>3} headers_needed={tgt['headers_needed']:>3}")
        print()
    dest = os.environ.get("CENSUS_JSON")
    if dest:
        slim = []
        for r in out:
            s = {k: v for k, v in r.items() if k not in ("entries_detail",)}
            slim.append(s)
        with open(dest, "w") as fh:
            json.dump(slim, fh, indent=1)
        print(f"wrote {dest}")


if __name__ == "__main__":
    main(sys.argv[1:])
