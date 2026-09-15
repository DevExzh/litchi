#!/usr/bin/env python3
"""Re-derive change 0575's residual window and price the two candidates.

Change 0575 computes a predecessor's zero-I/O span bracket as

    min_end_i = offset_i + 30 + central_name_len_i + compressed_size_i
    max_end_i = min_end_i + 65535 + 24            # "the residual window is 65,559 bytes"

using the record's *central* name length.  That is sound for the record being
read, whose local name the strict path forces to equal its central name, but
candidate (b) never validates a predecessor's name.  A predecessor may declare
a local `file_name_length` its central record does not carry, and both halves
of the local variable region are `u16`, so the sound bracket is

    min_end_i = offset_i + 30 + compressed_size_i
    max_end_i = offset_i + 30 + 2*65535 + compressed_size_i + 24

This script reports, for each fixture and target, how many predecessors each
bracket leaves undecided — that is, how many 30-byte local-header reads the
proof must issue.  Pure Python over the fixture bytes; no build required.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "change-0575"))
from zip_layout_census import parse_central  # noqa: E402

FIXED = 30
U16 = 65535
MAX_DESCRIPTOR = 24
BRACKETS = {
    "0575 (one u16, central name length)": lambda e: (
        e["local_header_offset"] + FIXED + e["central_name_len"] + e["compressed_size"],
        e["local_header_offset"] + FIXED + e["central_name_len"] + e["compressed_size"]
        + U16 + MAX_DESCRIPTOR,
    ),
    "0580 (two u16, no name assumption)": lambda e: (
        e["local_header_offset"] + FIXED + e["compressed_size"],
        e["local_header_offset"] + FIXED + 2 * U16 + e["compressed_size"] + MAX_DESCRIPTOR,
    ),
}

SCENARIOS = {
    "xlsx": {
        "one member": ["xl/worksheets/sheet1.xml"],
        "one-cell closure": [
            "[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.xml.rels",
            "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/sharedStrings.xml",
            "xl/styles.xml",
        ],
    },
    "pptx": {
        "one member": ["ppt/slides/slide1.xml"],
        "one-slide closure": [
            "[Content_Types].xml", "_rels/.rels", "ppt/_rels/presentation.xml.rels",
            "ppt/presentation.xml", "ppt/slides/_rels/slide1.xml.rels",
            "ppt/slides/slide1.xml",
        ],
    },
}


def residual(entries, bracket, target):
    """Predecessors of `target` the bracket cannot decide without a read."""
    offset = entries[target]["local_header_offset"]
    out = []
    for index in range(target):
        low, high = bracket(entries[index])
        if high <= offset or low > offset:
            continue
        out.append(index)
    return out


def ordered_probe_reads(entries, bracket, order):
    """Probe reads a reader issues for `order`, with its own memo in force.

    Change 0575's candidate table counts the *union* of the records each target
    touches. The implementation memoises per record, so a record first probed
    as a neighbour and later read as a target costs two reads, not one, and the
    union under-counts by exactly that number.
    """
    known = set()
    reads = 0
    per_target = []
    for target in order:
        issued = 1  # the target's own local-header read
        offset = entries[target]["local_header_offset"]
        for index in range(target - 1, -1, -1):
            low, high = bracket(entries[index])
            if high <= offset:
                continue
            if low > offset:
                break
            neighbour = entries[index]["local_header_offset"]
            if neighbour not in known:
                issued += 1
                known.add(neighbour)
        known.add(offset)
        reads += issued
        per_target.append((entries[target]["name"], issued))
    return reads, per_target


def main():
    root = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
    targets = [
        ("test-data/ooxml/xlsx/sheet-names.xlsx", "xlsx"),
        ("test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx", "xlsx"),
        ("test-data/ooxml/pptx/shapes.pptx", "pptx"),
    ]
    results = []
    for relative, kind in targets:
        path = os.path.join(root, relative)
        if not os.path.exists(path):
            print("MISSING", path, file=sys.stderr)
            continue
        entries, cd_off, total = parse_central(path)
        by_name = {e["name"]: i for i, e in enumerate(entries)}
        record = {"fixture": relative, "records": len(entries), "bytes": total,
                  "brackets": {}}
        print(f"== {relative}  ({len(entries)} records, {total} bytes)")
        for label, bracket in BRACKETS.items():
            per_scenario = {}
            for scenario, names in SCENARIOS[kind].items():
                idx = [by_name[n] for n in names if n in by_name]
                touched = set(idx)
                for t in idx:
                    touched.update(residual(entries, bracket, t))
                per_scenario[scenario] = {
                    "members": len(idx),
                    "records touched": len(touched),
                    "per target": {entries[t]["name"]: len(residual(entries, bracket, t))
                                   for t in idx},
                }
                print(f"   {label:<38} {scenario:<18} "
                      f"{len(idx):>3} members -> {len(touched):>4} records touched")
            worst = max(len(residual(entries, bracket, t)) for t in range(len(entries)))
            per_scenario["worst single-target residual"] = worst
            print(f"   {label:<38} {'worst target':<18} {worst:>3} predecessors")
            for scenario, names in SCENARIOS[kind].items():
                idx = [by_name[n] for n in names if n in by_name]
                reads, per_target = ordered_probe_reads(entries, bracket, idx)
                per_scenario[scenario]["ordered reads"] = reads
                per_scenario[scenario]["ordered per target"] = per_target
                print(f"   {label:<38} {scenario:<18} "
                      f"{reads:>4} reads once the per-record memo is modelled")
            record["brackets"][label] = per_scenario
        results.append(record)
        print()

    destination = os.environ.get("RESIDUAL_JSON")
    if destination:
        with open(destination, "w") as handle:
            json.dump(results, handle, indent=1)
        print("wrote", destination)


if __name__ == "__main__":
    main()
