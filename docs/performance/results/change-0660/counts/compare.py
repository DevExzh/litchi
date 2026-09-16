#!/usr/bin/env python3
"""Difference the isolation pairs of both legs and print the counts table.

Isolation pair: profile the harness at N=1 and N=3 samples of the same case and
halve the difference, so corpus construction, process start-up and the
harness's own verification cancel out.
"""
import re, sys, json

SYMBOLS = [
    "litchi_docx::document::transaction::Snapshot::from_xml",
    "litchi_docx::document::transaction::scan_document_with_context",
    "litchi_docx::package::package::document::<impl litchi_docx::package::model::Package>::document_snapshot",
    "litchi_docx::package::package::document::<impl litchi_docx::package::model::Package>::apply_document_patch",
    "litchi_docx::writer::doc::codec::compact_changed_document_xml",
    "litchi_docx::document::transaction::compact_changed_paragraphs",
    "litchi_docx::document::transaction::publication_accepts_preserved_xml",
    "xml_minifier::audit::verify_with_policy",
    "litchi_docx::document::transaction::Edit::commit",
]

def incl(path):
    out, total = {}, None
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
    names, counts, pending = {}, {}, None
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
    return counts

CASES = ("docx_semantic_one_edit_save",
         "docx_semantic_one_percent_edit_save",
         "docx_semantic_noop_edit_save")

def main():
    countdir, rawdir = sys.argv[1], sys.argv[2]
    out = {}
    for leg in ("before", "after"):
        report = {}
        for case in CASES:
            t1, i1 = incl(f"{countdir}/incl-{leg}-{case}-1.txt")
            t3, i3 = incl(f"{countdir}/incl-{leg}-{case}-3.txt")
            c1 = calls(f"{rawdir}/{leg}-{case}-1.out")
            c3 = calls(f"{rawdir}/{leg}-{case}-3.out")
            entry = {"whole_child_ir_n1": t1, "whole_child_ir_n3": t3,
                     "per_lifecycle_ir": (t3 - t1) // 2, "symbols": {}}
            for symbol in SYMBOLS:
                short = symbol.split("::")[-1]
                entry["symbols"][short] = {
                    "incl_ir_per_lifecycle": (i3.get(symbol, 0) - i1.get(symbol, 0)) / 2,
                    "calls_per_lifecycle": (c3.get(symbol, 0) - c1.get(symbol, 0)) / 2,
                }
            report[case] = entry
        out[leg] = report
    json.dump(out, open(f"{countdir}/counts-summary.json", "w"), indent=1)
    for case in CASES:
        b, a = out["before"][case], out["after"][case]
        delta = (a["per_lifecycle_ir"] - b["per_lifecycle_ir"]) / b["per_lifecycle_ir"] * 100
        print(f"\n== {case}")
        print(f"  whole iteration Ir: {b['per_lifecycle_ir']:,} -> {a['per_lifecycle_ir']:,} ({delta:+.2f}%)")
        for symbol in SYMBOLS:
            short = symbol.split("::")[-1]
            bi = b["symbols"][short]["incl_ir_per_lifecycle"]
            ai = a["symbols"][short]["incl_ir_per_lifecycle"]
            bc = b["symbols"][short]["calls_per_lifecycle"]
            ac = a["symbols"][short]["calls_per_lifecycle"]
            if bi or ai or bc or ac:
                print(f"    {short:38s} Ir {bi:>14,.0f} -> {ai:>14,.0f}   calls {bc:>7,.1f} -> {ac:>7,.1f}")

main()
