#!/usr/bin/env python3
"""Count IndexedArchive constructions per phase from a strace capture.

Only `pread64` on the package file descriptor counts; the dynamic loader's own
positional reads on its transient fd 3 are excluded via openat/close tracking.
A 22-byte read is the end-of-central-directory record: one archive construction.
"""
from __future__ import annotations
import collections, json, pathlib, re, sys

PREAD  = re.compile(r"^(\d+)\s+pread64\((\d+),\s*(.*),\s*(\d+),\s*(\d+)\)\s*=\s*(-?\d+)")
OPENAT = re.compile(r'^(\d+)\s+openat\([^,]+,\s*"([^"]*)"[^)]*\)\s*=\s*(-?\d+)')
CLOSE  = re.compile(r"^(\d+)\s+close\((\d+)\)")
WRITE  = re.compile(r"^(\d+)\s+write\((\d+),\s*\"(.*?)\"")

def summarize(path, target):
    fds, phase, cur = {}, "pre", None
    constructions, phase_reads, opens = [], collections.Counter(), 0
    for line in pathlib.Path(path).read_text(errors="replace").splitlines():
        m = OPENAT.match(line)
        if m:
            fd = int(m.group(3))
            if fd >= 0:
                fds[(int(m.group(1)), fd)] = m.group(2)
                if m.group(2).endswith(target):
                    opens += 1
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
        pid, fd, data, n, off, ret = (int(m.group(1)), int(m.group(2)), m.group(3),
                                      int(m.group(4)), int(m.group(5)), int(m.group(6)))
        if not (fds.get((pid, fd)) or "").endswith(target):
            continue
        phase_reads[phase] += 1
        if n == 22:
            cur = {"n": len(constructions) + 1, "phase": phase, "eocd_off": off,
                   "eocd_bytes": ret, "cd_bytes": 0, "cd_reads": 0,
                   "reads": 0, "bytes": 0, "member_reads": 0, "in_cd": True}
            constructions.append(cur)
        if cur is None:
            continue
        cur["reads"] += 1; cur["bytes"] += max(ret, 0)
        if n == 30:
            cur["in_cd"] = False; cur["member_reads"] += 1
        elif cur["in_cd"] and n != 22:
            cur["cd_reads"] += 1; cur["cd_bytes"] += max(ret, 0)
    total = sum(phase_reads.values())
    return {"trace": pathlib.Path(path).name, "file_opens": opens,
            "pread64_on_package": total, "reads_by_phase": dict(phase_reads),
            "constructions": constructions}

if __name__ == "__main__":
    root = pathlib.Path(sys.argv[1]); target = sys.argv[2]
    rows = []
    for p in sorted(root.iterdir()):
        if not p.name.endswith(".txt"):
            continue
        r = summarize(p, target if not len(sys.argv) > 3 else target)
        rows.append(r)
    for r in rows:
        cons = r["constructions"]
        per = collections.Counter(c["phase"] for c in cons)
        print(f"{r['trace'][:-4]:30s} opens={r['file_opens']} pread64={r['pread64_on_package']:6d} "
              f"archives={len(cons):3d}  by-phase: " +
              ", ".join(f"{k}={v}" for k, v in per.items()))
    print()
    print(json.dumps(rows, indent=2)[:0])
