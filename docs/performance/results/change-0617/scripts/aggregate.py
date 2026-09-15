#!/usr/bin/env python3
"""Difference two callgrind profiles of the same operation at different sample
counts and attribute the per-operation self instructions to owners.

Isolation pair: profile LOW and HIGH iterations of the operation, difference the
per-function self costs, divide by (HIGH - LOW). Process start-up, the fixture
read and the probe's own printing appear identically in both legs and cancel.
"""

import re
import sys
from collections import defaultdict

FN = re.compile(r"^fn=(?:\((\d+)\))?\s*(.*)$")
CFN = re.compile(r"^cfn=(?:\((\d+)\))?\s*(.*)$")
COST = re.compile(r"^[-+*\d]")


def parse(path):
    names = {}
    self_cost = defaultdict(int)
    current = None
    skip_next_cost = False
    with open(path, encoding="utf-8", errors="replace") as handle:
        for line in handle:
            line = line.rstrip("\n")
            if not line:
                continue
            if line.startswith("fn="):
                ident, name = FN.match(line).groups()
                if name:
                    if ident:
                        names[ident] = name
                    current = name
                else:
                    current = names.get(ident, f"?{ident}")
                skip_next_cost = False
                continue
            if line.startswith("cfn="):
                ident, name = CFN.match(line).groups()
                if name and ident:
                    names[ident] = name
                continue
            if line.startswith("calls="):
                skip_next_cost = True
                continue
            if line.startswith(("fl=", "fi=", "fe=", "cfi=", "cfl=", "ob=", "cob=")):
                # File/object records may also carry a name table entry.
                parts = line.split("=", 1)[1]
                match = re.match(r"^\((\d+)\)\s*(.*)$", parts)
                if match and match.group(2):
                    pass
                continue
            if COST.match(line):
                if skip_next_cost:
                    skip_next_cost = False
                    continue
                fields = line.split()
                if len(fields) >= 2 and current is not None:
                    try:
                        self_cost[current] += int(fields[1])
                    except ValueError:
                        pass
                continue
    return self_cost


# Ownership buckets. The order matters: the first pattern that matches wins.
BUCKETS = [
    ("cfb-writer", (
        "litchi_cfb::writer",
        "litchi_cfb..writer",
        "OleWriter",
        "FatBuilder",
        "MiniFatBuilder",
        "DirectoryBuilder",
        "HeaderBuilder",
        "DifatBuilder",
    )),
    ("cfb-parser", ("litchi_cfb",)),
    ("ole-common-container", ("litchi_ole_common::object", "litchi_ole_common..object")),
    ("ole-common-other", ("litchi_ole_common",)),
    ("format-xls", ("litchi_xls",)),
    ("format-doc", ("litchi_doc",)),
    ("format-ppt", ("litchi_ppt",)),
    ("format-biff", ("litchi_biff",)),
    ("format-odraw", ("litchi_odraw",)),
    ("format-other", (
        "litchi_core", "litchi_codepage", "litchi_formula", "litchi_sheet",
        "litchi_word", "litchi_slide", "litchi_drawingml", "litchi_vba",
        "litchi_crypto", "litchi_fonts", "litchi_eval", "litchi_spreadsheet",
    )),
    ("probe", ("cfb_save_probe",)),
]


def bucket(name):
    for label, patterns in BUCKETS:
        for pattern in patterns:
            if pattern in name:
                return label
    if name.startswith("0x") or name in ("???",):
        return "runtime"
    return "runtime"


def main():
    low_path, high_path, low_n, high_n = sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4])
    low, high = parse(low_path), parse(high_path)
    divisor = high_n - low_n
    per_op = {}
    for name in set(low) | set(high):
        delta = high.get(name, 0) - low.get(name, 0)
        if delta > 0:
            per_op[name] = delta / divisor
    total = sum(per_op.values())
    by_bucket = defaultdict(float)
    for name, value in per_op.items():
        by_bucket[bucket(name)] += value
    print(f"# {high_path} - {low_path} over {divisor} operations")
    print(f"total_self_instructions_per_operation {total:.0f}")
    print()
    print("## by owner")
    for label, value in sorted(by_bucket.items(), key=lambda kv: -kv[1]):
        print(f"{label:24} {value:14.0f}  {100 * value / total:6.2f}%")
    print()
    print("## top 30 symbols by self instructions per operation")
    for name, value in sorted(per_op.items(), key=lambda kv: -kv[1])[:30]:
        print(f"{value:14.0f}  {100 * value / total:6.2f}%  {bucket(name):22} {name[:120]}")


if __name__ == "__main__":
    main()
