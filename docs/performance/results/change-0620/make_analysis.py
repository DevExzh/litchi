#!/usr/bin/env python3
"""Renders the retained JSON summaries into the plain-text tables the record cites."""
import json, sys
from pathlib import Path

here = Path(__file__).parent
b = json.load(open(here / 'callgrind/callgrind-before.json'))
a = json.load(open(here / 'callgrind/callgrind-after.json'))
pb = json.load(open(here / 'perf/perf-before.json'))
pa = json.load(open(here / 'perf/perf-after.json'))

print("# Change 0620 attribution tables")
print("#")
print("# Callgrind: per-operation inclusive Ir, isolation pairs differenced and")
print("# divided by the extra operations, --separate-callers=2 so each call site")
print("# of a repeated function is its own row. One 'operation' is one whole")
print("# open + stage + commit + publish of the named scenario, including the")
print("# probe's own per-iteration copy of the fixture bytes.")
print("# Callgrind runs SHA-256 in software, so every `sha2` row is an upper")
print("# bound; the perf table below is the native counterpart.")
print()
for stem in ('54016', 'cv', 'formula'):
    print(f"================ {stem} ================")
    for op in ('open', 'number-plan', 'number-source-backed', 'number-generic',
               'string-generic', 'noop-generic'):
        key = f"{stem}/{op}"
        if key not in b:
            continue
        tb, ta = b[key]['PROGRAM TOTALS'], a[key]['PROGRAM TOTALS']
        print(f"\n-- {op}: {tb:,} -> {ta:,} Ir/op ({(ta - tb) / tb * 100:+.2f}%)")
        for name, ir in sorted(b[key].items(), key=lambda kv: -kv[1]):
            if ir < tb * 0.015 or "'" not in name:
                continue
            if name.startswith(('std::', '(below', '0x0000', 'events annotated',
                                '__libc', 'PROGRAM', 'main', 'xls_edit_probe::main')):
                continue
            short = name.split("'xls_edit_probe::main'")[0]
            print(f"   {ir:>13,} {100 * ir / tb:>6.2f}%  -> {a[key].get(name, 0):>13,}   {short}")
    print()

print("================ native perf stat, same isolation method ================")
print(f"{'scenario':32s} {'cycles before':>15s} {'cycles after':>15s} {'delta':>8s} "
      f"{'instr before':>15s} {'instr after':>15s} {'delta':>8s}")
for key in sorted(pb):
    cb, ca = pb[key].get('cycles', 0), pa.get(key, {}).get('cycles', 0)
    ib, ia = pb[key].get('instructions', 0), pa.get(key, {}).get('instructions', 0)
    print(f"{key:32s} {cb:>15,.0f} {ca:>15,.0f} {(ca - cb) / cb * 100:>7.2f}% "
          f"{ib:>15,.0f} {ia:>15,.0f} {(ia - ib) / ib * 100:>7.2f}%")
