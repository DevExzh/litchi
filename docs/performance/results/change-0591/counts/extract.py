#!/usr/bin/env python3
"""Extract per-lifecycle inclusive Ir and call counts from callgrind output.

Isolation pair: profile the harness at N=1 and N=3 samples of the same case and
divide the difference by 2, so corpus construction, process start-up and output
verification cancel out.
"""
import re, sys, os, json

SYMBOLS = [
    "litchi_docx::document::transaction::Snapshot::from_xml",
    "litchi_docx::document::transaction::scan_document_with_context",
    "litchi_docx::document::transaction::body_child_shape",
    "litchi_docx::package::package::document::<impl litchi_docx::package::model::Package>::document_snapshot",
    "litchi_docx::package::package::document::<impl litchi_docx::package::model::Package>::apply_document_patch",
    "litchi_docx::writer::doc::codec::compact_changed_document_xml",
]

def incl(path):
    """Inclusive Ir per function from `callgrind_annotate --inclusive=yes`."""
    out = {}
    total = None
    for line in open(path, errors="replace"):
        m = re.match(r"\s*([\d,]+) \(\s*[\d.]+%\)\s+(.*)$", line)
        if not m:
            m2 = re.match(r"\s*([\d,]+) \(100\.0%\)\s+PROGRAM TOTALS", line)
            if m2:
                total = int(m2.group(1).replace(",", ""))
            continue
        cost = int(m.group(1).replace(",", ""))
        name = m.group(2)
        if name.startswith("PROGRAM TOTALS"):
            total = cost
            continue
        name = re.sub(r"\s*\[[^\]]*\]\s*$", "", name)
        if ":" in name:
            name = name.split(":", 1)[1]
        out.setdefault(name, cost)
    return total, out

def calls(path):
    """Total call counts per callee name from a raw callgrind out file."""
    names = {}
    counts = {}
    pending = None
    for line in open(path, errors="replace"):
        line = line.rstrip("\n")
        m = re.match(r"(c?fn)=\((\d+)\)(?: (.*))?$", line)
        if m:
            kind, ident, name = m.group(1), m.group(2), m.group(3)
            if name:
                names[ident] = name
            pending = names.get(ident) if kind == "cfn" else None
            continue
        m = re.match(r"calls=(\d+)", line)
        if m and pending is not None:
            counts[pending] = counts.get(pending, 0) + int(m.group(1))
            continue
    return counts

def main():
    countdir, rawdir, leg = sys.argv[1], sys.argv[2], sys.argv[3]
    report = {}
    for case in ("docx_semantic_one_edit_save",
                 "docx_semantic_one_percent_edit_save",
                 "docx_semantic_noop_edit_save"):
        t1, i1 = incl(f"{countdir}/incl-{leg}-{case}-1.txt")
        t3, i3 = incl(f"{countdir}/incl-{leg}-{case}-3.txt")
        c1 = calls(f"{rawdir}/{leg}-{case}-1.out")
        c3 = calls(f"{rawdir}/{leg}-{case}-3.out")
        entry = {"whole_child_ir_n1": t1, "whole_child_ir_n3": t3,
                 "per_lifecycle_ir": (t3 - t1) // 2, "symbols": {}}
        for symbol in SYMBOLS:
            short = symbol.split("::")[-1]
            ir1, ir3 = i1.get(symbol, 0), i3.get(symbol, 0)
            k1 = c1.get(symbol, 0)
            k3 = c3.get(symbol, 0)
            entry["symbols"][short] = {
                "incl_ir_per_lifecycle": (ir3 - ir1) / 2,
                "calls_per_lifecycle": (k3 - k1) / 2,
                "incl_ir_n1": ir1, "calls_n1": k1,
            }
        report[case] = entry
    print(json.dumps(report, indent=1))

main()
