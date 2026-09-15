#!/usr/bin/env python3
"""Difference a callgrind isolation pair into per-operation inclusive Ir.

Profiles of N and N+M planning-plus-commit cycles against one retained editor
differ by exactly M operations, so (n7 - n2)/5 is one operation's inclusive
instruction count for every function in the profile. `--separate-callers=1`
keeps `parse` attributable to its planning caller and its commit caller.
"""
import re, sys

KEY = re.compile(r'^\s*([\d,]+) \(\s*[\d.]+%\)\s+(\S.*)$')

def load(path):
    out, started = {}, False
    for line in open(path, encoding="utf-8", errors="replace"):
        if line.startswith("Ir") and "file:function" in line:
            started = True
            continue
        if not started:
            continue
        if line.startswith("---") or not line.strip():
            continue
        m = KEY.match(line)
        if not m:
            continue
        name = m.group(2).strip()
        if name.endswith("]"):
            name = name[:name.rfind(" [")].strip()
        if ":" not in name:
            continue
        # callgrind_annotate emits one line per call context; " (Nx)" marks a
        # repeated context. Normalize and keep the outermost (largest) cost.
        name = re.sub(r" \(\d+x\)$", "", name)
        out[name] = max(out.get(name, 0), int(m.group(1).replace(",", "")))
    return out

def total(path):
    for line in open(path, encoding="utf-8", errors="replace"):
        if "PROGRAM TOTALS" in line:
            return int(line.split("(")[0].strip().replace(",", ""))
    return 0

WANT = [
    ("plan  MultiSnapshot::load_source_backed", "MultiSnapshot::load_source_backed'"),
    ("plan  validation::worksheet_xml_and_parse_source", "worksheet_xml_and_parse_source'"),
    ("plan  validation::worksheet_xml", "validation::worksheet_xml'"),
    ("both  mce::codec::process_markup_compatibility", "process_markup_compatibility'"),
    ("both  raw::worksheet::parse", "raw::worksheet::parse'"),
    ("cmt   MultiSourceEdit::commit", "MultiSourceEdit::commit'"),
    ("cmt   rewrite_value_only_with_provenance", "rewrite_value_only_with_provenance'"),
    ("cmt   Snapshot::from_rewritten_value_source", "from_rewritten_value_source'"),
    ("cmt   reduced_readback", "package::reduced_readback'"),
    ("cmt   Store::merge_omitted_cells", "Store::merge_omitted_cells'"),
    ("cmt   Snapshot::invalidated_workbook_xml", "invalidated_workbook_xml'"),
    ("cmt   validation::workbook_xml", "validation::workbook_xml'"),
]

def main():
    a, b, m = sys.argv[1], sys.argv[2], int(sys.argv[3])
    A, B = load(a), load(b)
    ta, tb = total(a), total(b)
    print("# pair %s -> %s, M=%d operations" % (a.split("/")[-1], b.split("/")[-1], m))
    print("%-52s %14s %14s" % ("function (inclusive, per operation)", "Ir/op", "% of op"))
    op = (tb - ta) / m
    print("%-52s %14.0f %13.2f%%" % ("TOTAL (one plan + commit)", op, 100.0))
    for label, needle in WANT:
        rows = [(k, B.get(k, 0) - A.get(k, 0)) for k in B if needle in k]
        rows = [(k, v) for k, v in rows if v > 0]
        if not rows:
            continue
        rows.sort(key=lambda kv: -kv[1])
        seen = set()
        for k, v in rows:
            if k in seen:
                continue
            seen.add(k)
            caller = k.split("'", 1)[1] if "'" in k else ""
            caller = caller.split("::")[-1][:30]
            print("%-52s %14.0f %13.2f%%" % ("%s <-%s" % (label, caller), v / m, 100.0 * v / m / op))
            break
main()
