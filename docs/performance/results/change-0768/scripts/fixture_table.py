#!/usr/bin/env python3
"""Render the per-fixture table of change 0768 from the probe_0768 JSONL runs.

Usage: fixture_table.py BASE.jsonl AFTER.jsonl
"""
import json
import sys


def load(path):
    return {row["file"]: row for row in map(json.loads, open(path))}


def producer(row):
    magic = row.get("magic34")
    return {"0x6A62": "Word", "0x6143": "LibreOffice"}.get(magic, "other" if magic else "-")


def edit(value):
    if value == "ok":
        return "edits"
    if "protected DOC publication" in value:
        return "refused: protection"
    if "ambiguous Selsf" in value:
        return "refused: Selsf CP"
    if "encrypted" in value:
        return "refused: encrypted"
    if "Word 97+ FIB" in value:
        return "refused: pre-Word 97 FIB"
    if "OLE error" in value:
        return "refused: invalid CFB"
    if "malformed CHPX FKP" in value:
        return "output fails to reopen (CHPX FKP)"
    if "exceeds ccpText" in value:
        return "refused: Selsf CP beyond the text"
    if "SPRM" in value:
        return "refused: malformed SPRM"
    return value[:40]


base, after = load(sys.argv[1]), load(sys.argv[2])
print("| file | producer | nFib / cbRgFcLcb / cswNew (nFibNew) | lcbDop | base | after | base tracked insert at CP 0 | after |")
print("| --- | --- | --- | ---: | --- | --- | --- | --- |")
for name in sorted(after, key=str.lower):
    b, a = base[name], after[name]
    shape = "-"
    if a.get("nfib"):
        new = f" ({a['nfib_new']})" if a.get("csw_new") else ""
        shape = f"{a['nfib']} / {a['cb_rg_fc_lcb']} / {a.get('csw_new')}{new}"
    print(
        f"| {name.split('__')[-1]} | {producer(a)} | {shape} | {a.get('lcb_dop', '-') if a.get('csw_new') is not None else '-'} "
        f"| {b['classification'][:12]} | {a['classification'][:12]} | {edit(b['edit_cp0'])} | {edit(a['edit_cp0'])} |"
    )
