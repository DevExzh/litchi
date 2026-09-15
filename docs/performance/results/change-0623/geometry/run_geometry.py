#!/usr/bin/env python3
"""Model change 0623's structural runs from package bytes alone.

Reimplements `IndexedArchive::local_span_hint` and
`read_ahead::structural_runs` over the central directory, so the retained
bytes and run counts can be stated per fixture without running the library.
"""
import struct, sys, os

MAX_RUN = 64 * 1024
MAX_TOTAL = 256 * 1024
ALLOWANCE = 640
FIXED = 30
DESC = 24

def geometry(b):
    for i in range(len(b) - 22, -1, -1):
        if b[i:i+4] == b'PK\x05\x06':
            eocd = i
            break
    else:
        return None
    cd_size, cd_off = struct.unpack_from('<II', b, eocd + 12)
    if cd_off == 0xFFFFFFFF or cd_size == 0xFFFFFFFF:
        return None
    members = []
    p = cd_off
    end = cd_off + cd_size
    while p < end:
        if b[p:p+4] != b'PK\x01\x02':
            return None
        flags, = struct.unpack_from('<H', b, p + 8)
        csize, = struct.unpack_from('<I', b, p + 20)
        nlen, xlen, clen = struct.unpack_from('<HHH', b, p + 28)
        lho, = struct.unpack_from('<I', b, p + 42)
        name = b[p+46:p+46+nlen].decode('utf-8', 'replace')
        members.append((lho, name, csize, bool(flags & 8)))
        p += 46 + nlen + xlen + clen
    members.sort(key=lambda m: m[0])
    return cd_off, members

def is_structural(name):
    if name.lower() == '[content_types].xml':
        return True
    if '/' not in name:
        return False
    directory, file = name.rsplit('/', 1)
    if not file.endswith('.rels'):
        return False
    last = directory.rsplit('/', 1)[-1]
    return last == '_rels'

def span(cd_off, members, index):
    lho, name, csize, desc = members[index]
    if lho >= cd_off:
        return None
    d = DESC if desc else 0
    nxt = cd_off
    for other in members[index+1:]:
        if other[0] > lho:
            nxt = other[0]
            break
    physical = min(nxt, cd_off) + d
    declared = lho + ALLOWANCE + csize + d
    length = min(physical, declared, cd_off) - lho
    if not (FIXED <= length <= MAX_RUN):
        return None
    return (lho, length)

def runs(spans):
    out = []
    cur = None
    for off, ln in spans:
        end = off + ln
        if cur is not None and cur[1] >= off:
            cur = (cur[0], max(cur[1], end), cur[2] + 1)
        else:
            if cur is not None and cur[2] >= 2 and 0 < cur[1] - cur[0] <= MAX_RUN:
                out.append(cur)
            cur = (off, end, 1)
    if cur is not None and cur[2] >= 2 and 0 < cur[1] - cur[0] <= MAX_RUN:
        out.append(cur)
    kept, total = [], 0
    for r in out:
        ln = r[1] - r[0]
        if total + ln <= MAX_TOTAL:
            total += ln
            kept.append(r)
    return kept

def main(listing):
    print("fixture\tmembers\tstructural\truns\trun_members\tretained_bytes\tlongest_run")
    worst = (0, None)
    worst_run = (0, None)
    for path in listing:
        try:
            b = open(path, 'rb').read()
        except OSError:
            continue
        g = geometry(b)
        if g is None:
            continue
        cd_off, members = g
        spans = []
        structural = 0
        for i, m in enumerate(members):
            if not is_structural(m[1]):
                continue
            structural += 1
            s = span(cd_off, members, i)
            if s is not None:
                spans.append(s)
        rs = runs(spans)
        retained = sum(r[1] - r[0] for r in rs)
        longest = max((r[1] - r[0] for r in rs), default=0)
        covered = sum(r[2] for r in rs)
        if retained > worst[0]:
            worst = (retained, path)
        if longest > worst_run[0]:
            worst_run = (longest, path)
        print(f"{path}\t{len(members)}\t{structural}\t{len(rs)}\t{covered}\t{retained}\t{longest}")
    print(f"# largest retained total: {worst[0]} B ({worst[1]})", file=sys.stderr)
    print(f"# longest single run: {worst_run[0]} B ({worst_run[1]})", file=sys.stderr)

if __name__ == '__main__':
    main([l.strip() for l in open(sys.argv[1]) if l.strip()])
