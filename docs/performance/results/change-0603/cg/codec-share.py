#!/usr/bin/env python3
"""Per-operation inclusive Ir of the MCE codec, by isolation pair.

Reads the one `mce/codec.rs` row of `process_markup_compatibility` called from
`process_ooxml` -- the whole preprocessing pass, planning and commit together --
from the N=1 and N=4 annotations and differences them over M=3 operations.
"""
import re, sys

ROW = re.compile(r'^\s*([\d,]+) \(\s*[\d.]+%\)\s+(\S+)')

def pick(path, file_suffix, func_prefix):
    needle = file_suffix + ":" + func_prefix
    for line in open(path, encoding="utf-8", errors="replace"):
        m = ROW.match(line)
        if not m:
            continue
        if needle in m.group(2):
            return int(m.group(1).replace(",", ""))
    return 0

def total(path):
    for line in open(path, encoding="utf-8", errors="replace"):
        if "PROGRAM TOTALS" in line:
            return int(line.split("(")[0].strip().replace(",", ""))
    return 0

FUNC = "litchi_ooxml_common::mce::codec::process_markup_compatibility'litchi_ooxml_common::mce::codec::process_ooxml"
print("%-6s %-7s %14s %14s %14s" % ("fixture", "leg", "op Ir", "codec Ir", "codec %"))
for tag in ("fct", "dvtr", "sss", "mfe", "ndp"):
    for leg in ("before", "after"):
        base = "%s/cg/%s/%s" % (sys.argv[1], leg, tag)
        op = (total(base + "-n4.incl") - total(base + "-n1.incl")) / 3
        codec = (pick(base + "-n4.incl", "mce/codec.rs", FUNC)
                 - pick(base + "-n1.incl", "mce/codec.rs", FUNC)) / 3
        print("%-6s %-7s %14.0f %14.0f %13.2f%%" % (tag, leg, op, codec, 100.0 * codec / op))
