#!/usr/bin/env python3
"""Per-open `next_chain_sector` calls and whole-open Ir, from threshold-100
callgrind isolation pairs. A threshold below 100 drops small edges and would
under-report the small fixtures, so every figure here is read from a
`--threshold=100` annotation of the same profile."""
import re, sys, json, pathlib

NUM = re.compile(r"^\s*([\d,]+) \(\s*[\d.]+%\)\s*")


def totals(path):
    for line in path.read_text().splitlines():
        if "PROGRAM TOTALS" in line:
            return int(NUM.match(line).group(1).replace(",", ""))
    raise SystemExit(f"no PROGRAM TOTALS in {path}")


def edges(path):
    out, pending = {}, []
    for line in path.read_text().splitlines():
        m = NUM.match(line)
        if not m:
            continue
        rest, ir = line[m.end():], int(m.group(1).replace(",", ""))
        if rest.startswith("*"):
            name = rest[1:].strip()
            name = name.split(":", 1)[1].split(" [")[0].strip() if ":" in name else name
            for caller, calls, cir in pending:
                out[(caller, name)] = (calls, cir)
            pending = []
        elif rest.startswith("<"):
            body = rest[1:].strip()
            cm = re.search(r"\((\d[\d,]*)x\)", body)
            calls = int(cm.group(1).replace(",", "")) if cm else 0
            name = body.split(":", 1)[1].split(" (")[0].strip() if ":" in body else body
            pending.append((name, calls, ir))
        else:
            pending = []
    return out


def fold(base, leg, stem, small, large):
    delta = large - small
    es = edges(base / f"tree100-{leg}-{stem}-s{small}.txt")
    el = edges(base / f"tree100-{leg}-{stem}-s{large}.txt")
    calls = {}
    for key in set(es) | set(el):
        if not key[1].endswith("next_chain_sector"):
            continue
        calls[key[0]] = (el.get(key, (0, 0))[0] - es.get(key, (0, 0))[0]) / delta
    ir = (totals(base / f"ann100-{leg}-{stem}-s{large}.txt")
          - totals(base / f"ann100-{leg}-{stem}-s{small}.txt")) / delta
    return {"whole_open_ir": ir, "chain_calls": calls, "chain_calls_total": sum(calls.values())}


SETS = [
    ("callgrind", [("flagship", 20, 220), ("cv", 20, 120), ("54016", 10, 60)]),
    ("callgrind-extra", [("mini-colours", 20, 220), ("mini-checkboxes", 20, 220),
                         ("mini-extstyles", 20, 220), ("small-simple", 20, 220),
                         ("mid-images", 20, 220)]),
]

root = pathlib.Path(sys.argv[1])
report = {}
for folder, stems in SETS:
    for stem, small, large in stems:
        for leg in ("before", "after"):
            report[f"{leg}/{stem}"] = fold(root / folder, leg, stem, small, large)
print(json.dumps(report, indent=1))
