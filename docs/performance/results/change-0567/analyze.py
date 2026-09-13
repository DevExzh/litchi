#!/usr/bin/env python3
"""Segment a `strace -f -k -e trace=pread64,openat,close,write` capture of one
library-level OOXML open into archive constructions, and attribute each
construction (and each structural-member read) to its litchi call chain.

An archive construction is marked by a 22-byte end-of-central-directory read on
the *package* file descriptor only; the dynamic loader's own `pread64` calls on
its transient fd 3 are excluded by tracking `openat`/`close`.
"""
from __future__ import annotations
import argparse, collections, json, pathlib, re, subprocess, sys

PREAD  = re.compile(r"^(\d+)\s+pread64\((\d+),\s*(.*),\s*(\d+),\s*(\d+)\)\s*=\s*(-?\d+)")
OPENAT = re.compile(r'^(\d+)\s+openat\([^,]+,\s*"([^"]*)"[^)]*\)\s*=\s*(-?\d+)')
CLOSE  = re.compile(r"^(\d+)\s+close\((\d+)\)")
WRITE  = re.compile(r"^(\d+)\s+write\((\d+),\s*\"(.*?)\",\s*\d+\)")
FRAME  = re.compile(r"^\s*>\s*(\S+?)(?:\(([^)]*)\))?\s*\[0x([0-9a-f]+)\]")

def parse(path, target):
    """-> list of events. Each is a dict."""
    events, fds, cur = [], {}, None
    lines = pathlib.Path(path).read_text(errors="replace").splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        m = FRAME.match(line)
        if m and cur is not None:
            cur["stack"].append((m.group(1), m.group(2) or "", int(m.group(3), 16)))
            i += 1
            continue
        cur = None
        m = OPENAT.match(line)
        if m:
            fd = int(m.group(3))
            if fd >= 0:
                fds[(int(m.group(1)), fd)] = m.group(2)
            i += 1
            continue
        m = CLOSE.match(line)
        if m:
            fds.pop((int(m.group(1)), int(m.group(2))), None)
            i += 1
            continue
        m = WRITE.match(line)
        if m and "#MARK#" in m.group(3):
            events.append({"kind": "mark", "tag": m.group(3).split("#")[2]})
            i += 1
            continue
        m = PREAD.match(line)
        if m:
            pid, fd, data, n, off, ret = (int(m.group(1)), int(m.group(2)), m.group(3),
                                          int(m.group(4)), int(m.group(5)), int(m.group(6)))
            name = fds.get((pid, fd))
            ev = {"kind": "pread", "pid": pid, "fd": fd, "file": name, "len": n,
                  "off": off, "ret": ret, "data": data, "stack": []}
            if name is not None and name.endswith(target):
                cur = ev
                events.append(ev)
            i += 1
            continue
        i += 1
    return events

class Sym:
    def __init__(self, binary):
        self.binary, self.cache = binary, {}
    def batch(self, addrs):
        todo = [a for a in addrs if a not in self.cache]
        if not todo:
            return
        # `-a` prefixes each address's block with `0x...`, which makes the
        # variable-length `-i` inline expansion unambiguous in one call.
        args = ["addr2line", "-a", "-f", "-C", "-i", "-e", self.binary] + [hex(a - 1) for a in todo]
        out = subprocess.run(args, capture_output=True, text=True).stdout.splitlines()
        order, blocks, cur = [], {}, None
        for line in out:
            if line.startswith("0x"):
                cur = int(line.strip(), 16)
                order.append(cur); blocks[cur] = []
            elif cur is not None:
                blocks[cur].append(line)
        for a in todo:
            body = blocks.get(a - 1, [])
            self.cache[a] = list(zip(body[0::2], body[1::2])) or [("??", "??")]

    def frames(self, addr):
        return self.cache.get(addr, [("??", "??")])

NOISE = ("core::", "alloc::", "std::", "<core::", "<alloc::", "__", "std_detect",
         "_start", "main", "?")

def chain(stack, sym, keep=("litchi", "soapberry")):
    """Innermost-first litchi/soapberry frames as 'func (file:line)'."""
    out = []
    for mod, symname, addr in stack:
        if not mod.endswith("opc-index-probe"):
            continue
        for func, loc in sym.frames(addr):
            short = loc.replace("/home/zhuhe/code/litchi/", "")
            if any(k in func for k in keep) or any(k in short for k in keep):
                out.append(f"{func} ({short})")
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("trace"); ap.add_argument("--binary", required=True)
    ap.add_argument("--target", required=True)
    ap.add_argument("--json", default=None)
    ap.add_argument("--full-chains", action="store_true")
    args = ap.parse_args()

    events = parse(args.trace, args.target)
    sym = Sym(args.binary)
    addrs = {a for e in events if e["kind"] == "pread" for (_m, _s, a) in e["stack"]}
    sym.batch(sorted(addrs))

    constructions, phase = [], "pre"
    reads_by_phase = collections.Counter()
    cur = None
    member_reads = []          # (phase, construction#, offset, len) local-header reads
    for e in events:
        if e["kind"] == "mark":
            phase = e["tag"]; continue
        if e["kind"] != "pread":
            continue
        reads_by_phase[phase] += 1
        if e["len"] == 22:
            cur = {"n": len(constructions) + 1, "phase": phase, "eocd_off": e["off"],
                   "chain": chain(e["stack"], sym), "reads": [], "bytes": 0}
            constructions.append(cur)
        if cur is not None:
            cur["reads"].append((e["off"], e["len"], e["ret"]))
            cur["bytes"] += max(e["ret"], 0)
        if e["len"] == 30:
            member_reads.append((phase, cur["n"] if cur else 0, e["off"], e["stack"]))

    total = sum(1 for e in events if e["kind"] == "pread")
    print(f"trace: {args.trace}")
    print(f"package-fd pread64 calls: {total}   archive constructions: {len(constructions)}")
    print("reads per phase: " + ", ".join(f"{k}={v}" for k, v in reads_by_phase.items()))
    print()
    for c in constructions:
        # Central-directory read(s): reads strictly after the EOCD read within
        # the same construction that are neither 30-byte local headers nor
        # 16-byte descriptors, taken before the first 30-byte header read.
        cd = []
        for off, n, ret in c["reads"][1:]:
            if n == 30:
                break
            cd.append((off, n, ret))
        print(f"--- construction #{c['n']}  phase={c['phase']}  eocd@{c['eocd_off']} (22 B)"
              f"  cd_reads={len(cd)} cd_bytes={sum(r for _o, _n, r in cd)}"
              f"  total_reads={len(c['reads'])} total_bytes={c['bytes']}")
        seen = set()
        for f in c["chain"]:
            if f in seen and not args.full_chains:
                continue
            seen.add(f)
            print(f"      {f}")
        print()

    if args.json:
        pathlib.Path(args.json).write_text(json.dumps({
            "trace": args.trace, "pread64_calls": total,
            "constructions": [{k: v for k, v in c.items() if k != "reads"} | {
                "read_count": len(c["reads"]), "bytes": c["bytes"]} for c in constructions],
            "reads_by_phase": dict(reads_by_phase),
        }, indent=2) + "\n")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
