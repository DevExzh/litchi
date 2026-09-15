#!/usr/bin/env python3
"""Rank OLE2 DOC/PPT read-path symbols by self instruction count.

Isolation method: the same child is profiled at samples=20 and samples=320.
Differencing the two annotations and dividing by 300 cancels process startup,
the warmup iteration and every one-time initialization, leaving per-operation
self Ir. A third leg at samples=170 checks that the growth is linear.
"""

import os
import re
import sys

ROOT = "/tmp/claude-1001/-home-zhuhe-code-litchi/14c44904-927d-4351-97ac-5611bafb5316/scratchpad/out-docppt/ann"
SMALL, MID, LARGE = 20, 170, 320

CELLS = [
    ("docbig",   "open", "ca.kwsymphony...Seat_Booking_Form.doc (1,619,457 B) / open"),
    ("docbig",   "text", "ca.kwsymphony...Seat_Booking_Form.doc (1,619,457 B) / open+text"),
    ("docmid",   "open", "FloatingPictures.doc (335,360 B) / open"),
    ("docmid",   "text", "FloatingPictures.doc (335,360 B) / open+text"),
    ("docsmall", "open", "saved-by-table.doc (65,024 B) / open"),
    ("docsmall", "text", "saved-by-table.doc (65,024 B) / open+text"),
    ("pptbig",   "open", "45543-transition-litchi.ppt (842,240 B) / open+slide_count"),
    ("pptbig",   "text", "45543-transition-litchi.ppt (842,240 B) / open+text"),
    ("pptmid",   "open", "45543.ppt (385,024 B) / open+slide_count"),
    ("pptmid",   "text", "45543.ppt (385,024 B) / open+text"),
    ("pptsmall", "open", "SampleShow.ppt (125,440 B) / open+slide_count"),
    ("pptsmall", "text", "SampleShow.ppt (125,440 B) / open+text"),
]

ANN_LINE = re.compile(r"^\s*([\d,]+)\s*\([^)]*\)\s+(\S.*?)\s*$")
OBJECT_SUFFIX = re.compile(r"\s*\[[^\]]*\]\s*$")
FILE_PREFIX = re.compile(r"^.*?:(?!:)")


def strip(name):
    name = OBJECT_SUFFIX.sub("", name)
    return FILE_PREFIX.sub("", name, count=1)


def annotate(path):
    totals = {}
    with open(path, encoding="utf-8", errors="replace") as handle:
        for line in handle:
            match = ANN_LINE.match(line)
            if not match:
                continue
            raw = match.group(2)
            if raw.startswith(("<", "*", "-")) or "PROGRAM TOTALS" in raw:
                continue
            name = strip(raw)
            totals[name] = totals.get(name, 0) + int(match.group(1).replace(",", ""))
    return totals


def program_total(path):
    with open(path, encoding="utf-8", errors="replace") as handle:
        for line in handle:
            if "PROGRAM TOTALS" in line:
                return int(re.match(r"^\s*([\d,]+)", line).group(1).replace(",", ""))
    return 0


CRATES = (
    "litchi_cfb", "litchi_ole_common", "litchi_doc", "litchi_ppt", "litchi_odraw",
    "litchi_crypto", "litchi_core", "litchi_sheet", "litchi_drawingml",
    "hashbrown", "smallvec", "memchr", "encoding_rs", "compiler_builtins",
)


def crate_of(symbol):
    for crate in CRATES:
        if crate in symbol:
            return crate.replace("_", "-")
    if "malloc" in symbol or "free" in symbol or symbol.startswith("__") or "_avx" in symbol:
        return "libc/runtime"
    if "::" not in symbol:
        return "libc/runtime"
    return "core/alloc/std"


def path_for(stem, op, n, kind="self"):
    return os.path.join(ROOT, f"{kind}-{stem}-{op}-s{n}.txt")


def cell(stem, op, label, top=15):
    small, large = path_for(stem, op, SMALL), path_for(stem, op, LARGE)
    if not (os.path.exists(small) and os.path.exists(large)):
        print(f"\n== {stem}/{op}: MISSING annotation ==")
        return
    low, high = annotate(small), annotate(large)
    span = LARGE - SMALL
    delta = {}
    for name in set(low) | set(high):
        value = (high.get(name, 0) - low.get(name, 0)) / span
        if abs(value) >= 1:
            delta[name] = value
    total = sum(delta.values())
    negative = sum(v for v in delta.values() if v < 0)
    whole = (program_total(large) - program_total(small)) / span

    print()
    print(f"== {stem}/{op}  --  {label} ==")
    print(f"   whole-operation Ir per sample (PROGRAM TOTALS delta/{span}) : {whole:,.0f}")
    print(f"   sum of per-symbol self deltas                              : {total:,.0f}")
    print(f"   negative-delta mass (noise floor)                          : {negative:,.0f}"
          f"  ({abs(negative) / whole * 100 if whole else 0:.4f}% of whole)")

    # linearity check against the mid leg
    mid = path_for(stem, op, MID)
    if os.path.exists(mid):
        a = (program_total(mid) - program_total(small)) / (MID - SMALL)
        b = (program_total(large) - program_total(mid)) / (LARGE - MID)
        drift = abs(a - b) / whole * 100 if whole else 0
        print(f"   linearity: 20->170 = {a:,.0f} Ir/op ; 170->320 = {b:,.0f} Ir/op"
              f"  (drift {drift:.4f}%)")

    print(f"   {'rank':>4}  {'self Ir/op':>12}  {'%':>6}  {'crate':<18} symbol")
    for index, (name, value) in enumerate(sorted(delta.items(), key=lambda kv: -kv[1])[:top], 1):
        share = value / whole * 100 if whole else 0
        print(f"   {index:>4}  {value:>12,.0f}  {share:>5.2f}%  {crate_of(name):<18} {name[:150]}")

    by_crate = {}
    for name, value in delta.items():
        by_crate[crate_of(name)] = by_crate.get(crate_of(name), 0) + value
    print("   -- by crate --")
    for crate, value in sorted(by_crate.items(), key=lambda kv: -kv[1]):
        print(f"         {value:>12,.0f}  {value / whole * 100 if whole else 0:>5.2f}%  {crate}")


TARGETS = (
    "next_chain_sector", "read_stream_range_hinted", "read_stream_range", "collect_exact",
    "validate_stream_allocations", "validate_physical_sector_layout", "load_fat",
    "open_stream", "read_sector_run",
)


def cfb_report():
    print("\n\n######## litchi-cfb / litchi-ole-common symbols, self Ir per operation ########")
    header = f"{'symbol':<72}" + "".join(f"{s+'/'+o:>17}" for s, o, _ in CELLS)
    rows = {}
    wholes = {}
    for stem, op, _label in CELLS:
        small, large = path_for(stem, op, SMALL), path_for(stem, op, LARGE)
        if not (os.path.exists(small) and os.path.exists(large)):
            continue
        low, high = annotate(small), annotate(large)
        span = LARGE - SMALL
        wholes[(stem, op)] = (program_total(large) - program_total(small)) / span
        for name in set(low) | set(high):
            if "litchi_cfb" not in name and "litchi_ole_common" not in name:
                continue
            value = (high.get(name, 0) - low.get(name, 0)) / span
            if abs(value) < 1:
                continue
            rows.setdefault(name, {})[(stem, op)] = value

    order = sorted(rows, key=lambda n: -sum(rows[n].values()))
    print(header)
    for name in order:
        cells = "".join(
            f"{rows[name].get((s, o), 0):>10,.0f}"
            + f"{(rows[name].get((s,o),0)/wholes[(s,o)]*100 if wholes.get((s,o)) else 0):>6.2f}%"
            for s, o, _ in CELLS
        )
        star = " *" if any(t in name for t in TARGETS) else "  "
        print(f"{name[:70]:<70}{star}{cells}")

    print("\n-- named targets present / absent --")
    for target in TARGETS:
        hits = [n for n in rows if target in n]
        if hits:
            for hit in hits:
                per = ", ".join(
                    f"{s}/{o}={rows[hit].get((s,o),0):,.0f}" for s, o, _ in CELLS
                    if rows[hit].get((s, o), 0)
                )
                print(f"  PRESENT  {target:<34} -> {hit[:100]}\n           {per}")
        else:
            print(f"  ABSENT   {target}")


if __name__ == "__main__":
    which = sys.argv[1:] 
    for stem, op, label in CELLS:
        if which and f"{stem}-{op}" not in which:
            continue
        cell(stem, op, label, top=int(os.environ.get("TOP", "15")))
    if not which:
        cfb_report()
