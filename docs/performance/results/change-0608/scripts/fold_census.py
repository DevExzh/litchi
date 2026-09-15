#!/usr/bin/env python3
"""Fold `sst-prefix-census.jsonl` into the tables change 0608 cites."""
import json
import statistics
import sys


def main(path):
    rows = [json.loads(line) for line in open(path)]
    summary = next((r for r in rows if r.get("summary")), None)
    skips = [r for r in rows if "skip" in r]
    data = [r for r in rows if "skip" not in r and not r.get("summary")]
    refused = [r for r in data if r["refused"]]
    ok = [r for r in data if not r["refused"]]
    strings = [r for r in ok if r["unique"]]

    print("=" * 78)
    print("Corpus census: what a prefix shared-string index would have to walk")
    print("=" * 78)
    print(f"  .xls and .xlt fixtures examined                     {len(rows) - 1:>8}")
    print(f"  no SST record reachable                             {len(skips):>8}")
    print(f"  carrying an SST record                              {len(data):>8}")
    print(f"  refused at open by an SST *header* check            {len(refused):>8}")
    for r in refused:
        print(f"      {r['file']:<34} {r['refused']}")
    print(f"  indexed at open today                               {len(ok):>8}")
    print(f"  of those, declaring at least one shared string      {len(strings):>8}")
    print(f"  of those, with at least one LabelSst cell           "
          f"{sum(1 for r in strings if r['in_range']):>8}")
    print(f"  shared-string entries over the corpus               "
          f"{sum(r['unique'] for r in strings):>8,}")
    print(f"  LabelSst cells over the corpus                      "
          f"{sum(r['labelsst'] for r in strings):>8,}")
    print()
    print("  THE DECISIVE COUNT: fixtures where the highest SST index any cell")
    print("  references is the LAST entry (so a full text walks the whole table")
    print(f"  under a prefix index and deferral saves nothing)     "
          f"{sum(1 for r in strings if r['prefix_text'] >= r['unique']):>8} of {len(strings)}")
    print()
    print("  Entries a prefix index must walk, as a percentage of the table:")
    means = [100.0 * r["prefix_mean"] / r["unique"] for r in strings if r["in_range"]]
    firsts = [100.0 * r["prefix_first"] / r["unique"] for r in strings if r["in_range"]]
    for label, series in (("one uniformly chosen string cell", means),
                          ("the first string cell in stream order", firsts)):
        q = statistics.quantiles(series, n=4)
        print(f"    {label:<40} min {min(series):6.2f}%  p25 {q[0]:6.2f}%  "
              f"median {statistics.median(series):6.2f}%  p75 {q[2]:6.2f}%  "
              f"max {max(series):6.2f}%")
    print(f"    {'a full text, or all cells':<40} 100.00% on every fixture")
    print()

    print("=" * 78)
    print("The three fixtures this record profiles")
    print("=" * 78)
    want = {
        "test-data/ole/xls/ConditionalFormattingSamples.xls": "flagship, worksheet 1",
        "test-data/ole/xls/WithCustomViews.xls": "worksheet 1",
        "test-data/poi/test-data/spreadsheet/54016.xls": "worksheet 0",
    }
    print(f"{'fixture':<34} {'uniq':>6} {'SST B':>8} {'segs':>5} {'cells':>7} "
          f"{'idx_max':>8} {'mean pfx':>9} {'harness cell':>13}")
    for path_, note in want.items():
        r = next(r for r in ok if r["path"] == path_)
        sheet = int(note.split()[-1])
        cell = r["one_cell_isst"][sheet] if sheet < len(r["one_cell_isst"]) else None
        print(f"{r['file']:<34} {r['unique']:>6,} {r['sst_bytes']:>8,} {r['segments']:>5} "
              f"{r['labelsst']:>7,} {r['idx_max']:>8,} {r['prefix_mean']:>9,.1f} "
              f"{('isst ' + str(cell)) if cell is not None else 'no string':>13}")
    print()
    print("  'harness cell' is the cell `xls_source_attribution --operation one-cell`")
    print("  reads: row 1, column 0 of the worksheet index this record profiles.")
    print("  A prefix index walks `isst + 1` entries for it, and none at all when")
    print("  that cell is not a shared string.")
    print()

    print("=" * 78)
    print("Largest tables in the corpus")
    print("=" * 78)
    print(f"{'fixture':<38} {'uniq':>7} {'cells':>7} {'segs':>5} {'mean pfx %':>11}")
    for r in sorted(strings, key=lambda r: -r["unique"])[:12]:
        share = 100.0 * r["prefix_mean"] / r["unique"] if r["in_range"] else 0.0
        print(f"{r['file'][:38]:<38} {r['unique']:>7,} {r['labelsst']:>7,} "
              f"{r['segments']:>5} {share:>10.2f}%")
    if summary:
        print()
        print("Probe's own summary line:", json.dumps(summary))


if __name__ == "__main__":
    main(sys.argv[1])
