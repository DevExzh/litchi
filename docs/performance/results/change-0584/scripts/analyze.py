#!/usr/bin/env python3
"""Rank source-backed OOXML read-path symbols by self instruction count.

Differences a small-sample and a large-sample `callgrind_annotate` profile of
the same child and divides by the extra operation count, which cancels process
startup, warmups and one-time initialization. Two independent replicates of
each leg give a measured repeatability spread. Pure stdlib.
"""

import os
import re
import sys

ROOT = "/tmp/claude-1001/-home-zhuhe-code-litchi/14c44904-927d-4351-97ac-5611bafb5316/scratchpad/cg"

# stem -> (small, large, human label)
CELLS = {
    "cfs-open": (20, 120, "ConditionalFormattingSamples.xlsx | source-backed open"),
    "cfs-scan": (5, 30, "ConditionalFormattingSamples.xlsx | open + full sheet0 sweep"),
    "ndp-open": (20, 220, "no_drawing_patriarch.xlsx | source-backed open"),
    "ndp-scan": (2, 12, "no_drawing_patriarch.xlsx | open + full sheet0 sweep (75,770 cells)"),
    "49609-open": (20, 220, "49609.xlsx | source-backed open"),
    "49609-read": (3, 18, "49609.xlsx | open + list sheets + cell A1 (streaming selected cell)"),
    "saut-open": (20, 220, "saut_page.docx | source-backed open"),
    "saut-read": (10, 60, "saut_page.docx | open + full text"),
    "h123-read": (5, 30, "heading123.docx | open + full text (1,802 B)"),
    "aes-open": (20, 120, "ArtisticEffectSample.pptx | source-backed open"),
    "aes-read": (20, 120, "ArtisticEffectSample.pptx | open + full text"),
    "b62513-read": (5, 30, "bug62513.pptx | open + full text (19 slides, 4,268 B)"),
}

ANN_LINE = re.compile(r"^\s*([\d,]+)(?:\s*\([^)]*\))?\s+(\S.*?)\s*$")
OBJECT_SUFFIX = re.compile(r"\s*\[[^\]]*\]\s*$")

CRATES = (
    "soapberry_zip", "litchi_opc", "litchi_xlsx", "litchi_docx", "litchi_pptx",
    "litchi_ooxml_common", "litchi_core", "litchi_sheet", "litchi_drawingml",
    "litchi_crypto", "litchi_cfb", "litchi_ole_common", "litchi_formula",
    "quick_xml", "flate2", "miniz_oxide", "hashbrown", "smallvec", "memchr",
    "chrono", "url", "serde", "ooxml_ir_probe", "zlib_rs", "crc32fast",
)


def annotate(path):
    totals = {}
    with open(path, encoding="utf-8", errors="replace") as handle:
        for line in handle:
            match = ANN_LINE.match(line)
            if not match:
                continue
            name = OBJECT_SUFFIX.sub("", match.group(2))
            if name.startswith("Ir") or "PROGRAM TOTALS" in name or name.startswith("--"):
                continue
            try:
                count = int(match.group(1).replace(",", ""))
            except ValueError:
                continue
            totals[name] = totals.get(name, 0) + count
    return totals


def program_total(path):
    with open(path, encoding="utf-8", errors="replace") as handle:
        for line in handle:
            if "PROGRAM TOTALS" in line:
                match = re.match(r"^\s*([\d,]+)", line)
                if match:
                    return int(match.group(1).replace(",", ""))
    return 0


def bare(name):
    """`callgrind_annotate` prints `<source-file>:<symbol>`; keep the symbol."""
    return name.split(":", 1)[1] if ":" in name else name


def crate_of(name):
    symbol = bare(name)
    for crate in CRATES:
        if symbol.startswith(crate + "::") or ("<" + crate + "::") in symbol:
            return crate.replace("_", "-")
    if "::" not in symbol:
        return "libc/ld.so/vg"
    return "core/alloc/std"


def deltas(stem, rep):
    small, large, _ = CELLS[stem]
    lo_path = os.path.join(ROOT, f"ann-{stem}-small-{rep}.txt")
    hi_path = os.path.join(ROOT, f"ann-{stem}-large-{rep}.txt")
    if not (os.path.exists(lo_path) and os.path.exists(hi_path)):
        return None, None
    extra = large - small
    low, high = annotate(lo_path), annotate(hi_path)
    delta = {}
    for name in set(low) | set(high):
        value = (high.get(name, 0) - low.get(name, 0)) / extra
        if abs(value) >= 0.5:
            delta[name] = value
    whole = (program_total(hi_path) - program_total(lo_path)) / extra
    return delta, whole


def table(stem, top=25):
    label = CELLS[stem][2]
    rep1, whole1 = deltas(stem, 1)
    rep2, whole2 = deltas(stem, 2)
    if rep1 is None:
        print(f"\n== {stem} == MISSING")
        return
    merged = {}
    for name in set(rep1) | set(rep2 or {}):
        merged[name] = (rep1.get(name, 0) + (rep2 or {}).get(name, 0)) / (2 if rep2 else 1)
    total = sum(merged.values())
    negative = sum(v for v in merged.values() if v < 0)

    print()
    print(f"== {stem} :: {label} ==")
    print(f"   per-op total Ir, PROGRAM TOTALS delta : rep1 {whole1:,.0f}"
          + (f" | rep2 {whole2:,.0f} | spread {abs(whole1 - whole2) / whole1 * 100:.2f}%"
             if rep2 else ""))
    print(f"   per-op total Ir, summed symbol deltas : {total:,.0f}")
    print(f"   negative-delta mass (noise floor)     : {negative:,.0f}"
          f" ({abs(negative) / total * 100 if total else 0:.3f}% of total)")
    print(f"   {'#':>3} {'self Ir/op':>12} {'%':>6}  {'rep-spread':>10}  {'crate':<20} symbol")
    ranked = sorted(merged.items(), key=lambda kv: -kv[1])
    for index, (name, value) in enumerate(ranked[:top], start=1):
        share = value / total * 100 if total else 0
        if rep2 and name in rep1 and name in rep2 and value:
            spread = f"{abs(rep1[name] - rep2[name]) / abs(value) * 100:9.1f}%"
        else:
            spread = "        --"
        print(f"   {index:>3} {value:>12,.0f} {share:>5.2f}%  {spread}  "
              f"{crate_of(name):<20} {bare(name)}")

    by_crate = {}
    for name, value in merged.items():
        by_crate[crate_of(name)] = by_crate.get(crate_of(name), 0) + value
    print("   -- self Ir by owning crate --")
    for crate, value in sorted(by_crate.items(), key=lambda kv: -kv[1]):
        if total and abs(value) < total * 0.001:
            continue
        print(f"       {value:>12,.0f} {value / total * 100:>5.2f}%  {crate}")


def main():
    stems = sys.argv[1:] or list(CELLS)
    for stem in stems:
        table(stem, top=int(os.environ.get("TOP", "25")))


if __name__ == "__main__":
    main()
