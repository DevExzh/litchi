#!/usr/bin/env python3
"""Record 0377: reduce the raw harness outputs (kept outside the repository)
to the committed evidence.

    summarize.py callgrind <all.tsv> > callgrind.tsv
    summarize.py abba <dir with base-N.json, cand-N.json> > abba.json

`all.tsv` is the per-run table of cg-one.sh: side, group, route, leg,
iterations, total Ir of the counted region, exit status. The reduced table
holds Ir per call for each leg on each build. The ABBA reduction keeps,
per route and build, every run's per-block leaf and unified medians.
"""
import json
import sys
from pathlib import Path


def callgrind(path):
    runs = {}
    for line in open(path, encoding="utf-8"):
        f = line.rstrip("\n").split("\t")
        if len(f) != 7:
            continue
        side, group, route, leg, iters, total, rc = f
        if rc != "0" or total == "NA":
            raise SystemExit(f"failed run: {line.strip()}")
        runs.setdefault((group, route), {})[(side, leg)] = int(total) / int(iters)
    print("group\troute\tleaf_base_ir\tleaf_cand_ir\tunified_base_ir\tunified_cand_ir")
    for (group, route), v in runs.items():
        cells = [v.get((s, l)) for s, l in
                 (("base", "leaf"), ("cand", "leaf"), ("base", "unified"), ("cand", "unified"))]
        print("\t".join([group, route] + ["" if c is None else f"{c:.0f}" for c in cells]))


def abba(directory):
    out = {"schema": "litchi.unified-fin-perf.abba-summary.v1", "record": "0377", "runs": []}
    for path in sorted(Path(directory).glob("*-[0-9].json")):
        side, n = path.stem.rsplit("-", 1)
        doc = json.loads(path.read_text(encoding="utf-8"))
        out["runs"].append({
            "build": side, "order": int(n),
            "blocks": doc["blocks"], "rounds_per_block": doc["rounds_per_block"],
            "cases": [{k: c[k] for k in ("group", "route", "format", "operations",
                                         "equivalence", "leaf_block_median_ns",
                                         "unified_block_median_ns")} for c in doc["cases"]],
        })
    json.dump(out, sys.stdout, separators=(",", ":"))
    print()


if __name__ == "__main__":
    {"callgrind": callgrind, "abba": abba}[sys.argv[1]](sys.argv[2])
