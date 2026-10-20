#!/usr/bin/env python3
"""Rebuild `observations.json` from the original `cached-results.json`.

usage: compact.py CACHED_RESULTS_JSON > observations.json

Each observation keeps every field; its literal cell closure becomes a list
of blocks: `sheet<TAB>first row<TAB>first column<TAB>width<TAB>cell...`,
the cells of consecutive rows that each span the same `width` consecutive
columns, row-major. A cell is `n<number>`, `t<text>` or `e` (empty). The
script decodes its output and asserts it equals the input.
"""
import json, sys

TYPES = {"number": "n", "text": "t", "empty": "e"}


def encode(cells):
    runs, current = [], None
    for cell in sorted(cells, key=lambda c: (c["sheet"], c["row"], c["column"])):
        code = TYPES[cell["type"]] + cell.get("value", "")
        assert "\t" not in code and "\n" not in code
        key = (cell["sheet"], cell["row"])
        if current and current[0] == key and current[1] + len(current[2]) == cell["column"]:
            current[2].append(code)
        else:
            current = [key, cell["column"], [code]]
            runs.append(current)
    blocks = []
    for (sheet, row), column, values in runs:
        last = blocks[-1] if blocks else None
        if (last and last[0] == sheet and last[2] == column and last[3] == len(values)
                and last[1] + len(last[4]) // last[3] == row):
            last[4].extend(values)
        else:
            blocks.append([sheet, row, column, len(values), list(values)])
    return ["\t".join([s, str(r), str(c), str(w)] + v) for s, r, c, w, v in blocks]


def decode(runs):
    names = {v: k for k, v in TYPES.items()}
    cells = []
    for run in runs:
        sheet, row, column, width, *values = run.split("\t")
        for offset, value in enumerate(values):
            cell = {"sheet": sheet, "row": int(row) + offset // int(width),
                    "column": int(column) + offset % int(width), "type": names[value[0]]}
            if value[0] != "e":
                cell["value"] = value[1:]
            cells.append(cell)
    return cells


source = json.load(open(sys.argv[1]))
out = [dict(entry, cells=encode(entry["cells"])) for entry in source]
key = lambda c: (c["sheet"], c["row"], c["column"])
for original, compact in zip(source, out):
    assert sorted(original["cells"], key=key) == decode(compact["cells"])
    assert {k: v for k, v in original.items() if k != "cells"} == {
        k: v for k, v in compact.items() if k != "cells"}
json.dump(out, sys.stdout, ensure_ascii=False, separators=(",", ":"))
sys.stdout.write("\n")
