#!/usr/bin/env python3
"""Map every 30-byte local-header read in a strace capture to its ZIP member.

The member name comes from the fixture's own central directory (keyed by
`local_header_offset`), so the attribution is exact rather than inferred.
"""
from __future__ import annotations
import collections, pathlib, re, struct, sys, zipfile

PREAD  = re.compile(r"^(\d+)\s+pread64\((\d+),\s*(.*),\s*(\d+),\s*(\d+)\)\s*=\s*(-?\d+)")
OPENAT = re.compile(r'^(\d+)\s+openat\([^,]+,\s*"([^"]*)"[^)]*\)\s*=\s*(-?\d+)')
CLOSE  = re.compile(r"^(\d+)\s+close\((\d+)\)")
WRITE  = re.compile(r"^(\d+)\s+write\((\d+),\s*\"(.*?)\"")

def central_directory(path):
    """local_header_offset -> (name, compressed_size, uncompressed_size, method)"""
    out = {}
    with zipfile.ZipFile(path) as z:
        for i in z.infolist():
            out[i.header_offset] = (i.filename, i.compress_size, i.file_size, i.compress_type)
    return out

def analyze(trace, fixture):
    cd = central_directory(fixture)
    target = pathlib.Path(fixture).name
    fds, phase = {}, "pre"
    per_member = collections.Counter()
    per_member_phase = collections.defaultdict(collections.Counter)
    reads = []
    for line in pathlib.Path(trace).read_text(errors="replace").splitlines():
        m = OPENAT.match(line)
        if m:
            fd = int(m.group(3))
            if fd >= 0:
                fds[(int(m.group(1)), fd)] = m.group(2)
            continue
        m = CLOSE.match(line)
        if m:
            fds.pop((int(m.group(1)), int(m.group(2))), None); continue
        m = WRITE.match(line)
        if m and "#MARK#" in m.group(3):
            phase = m.group(3).split("#")[2]; continue
        m = PREAD.match(line)
        if not m:
            continue
        pid, fd, n, off, ret = (int(m.group(1)), int(m.group(2)), int(m.group(4)),
                                int(m.group(5)), int(m.group(6)))
        if not (fds.get((pid, fd)) or "").endswith(target):
            continue
        reads.append((phase, off, n, ret))
    # A 30-byte local-header read is a *full member read* only when a payload
    # read follows it. A header read followed by an 8-byte read at header+30 is
    # the ODF `mimetype` name sniff, which never reads the member.
    for i, (ph, off, n, _ret) in enumerate(reads):
        if n != 30 or off not in cd:
            continue
        nxt = reads[i + 1] if i + 1 < len(reads) else None
        sniff = nxt is not None and nxt[2] == 8 and nxt[1] == off + 30
        kind = "name-sniff" if sniff else "full-read"
        per_member[(cd[off][0], kind)] += 1
        per_member_phase[(cd[off][0], kind)][ph] += 1
    return cd, reads, per_member, per_member_phase

if __name__ == "__main__":
    trace, fixture = sys.argv[1], sys.argv[2]
    cd, reads, per_member, per_phase = analyze(trace, fixture)
    print(f"trace={pathlib.Path(trace).name}  fixture={pathlib.Path(fixture).name} "
          f"({len(cd)} members)  package pread64={len(reads)}")
    print(f"{'member':46s} {'reads':>5s}  by phase")
    for (name, kind), n in sorted(per_member.items(), key=lambda kv: (-kv[1], kv[0])):
        print(f"{name:46s} {n:5d}  {kind:10s} " + ", ".join(f"{k}={v}" for k, v in per_phase[(name, kind)].items()))
    unmatched = [r for r in reads if r[2] == 30 and r[1] not in cd]
    if unmatched:
        print(f"unmatched 30-byte reads: {unmatched}")
