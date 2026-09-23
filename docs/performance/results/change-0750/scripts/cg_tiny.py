#!/usr/bin/env python3
"""Change 0750: where the source audit's fixed cost goes. Exclusive
instructions per audit of the 55-byte `source/tiny-document` probe case, by
function, from the audit probe's callgrind runs at 10,000 and 0 iterations
(scripts/probe_callgrind.sh), for both legs. Functions that change by fewer
than 20 instructions per audit are omitted.

Usage: cg_tiny.py <directory with cg-<leg>-source_tiny-<n>.out>"""
import re, subprocess, sys

ROW = re.compile(r"\s*([\d,]+) \([^)]*\)\s+(\S.*?)(?: \[.*)?$")


def exclusive(path):
    out = subprocess.run(["callgrind_annotate", "--threshold=100", "--inclusive=no", path],
                         capture_output=True, text=True, check=True).stdout
    costs, started = {}, False
    for line in out.splitlines():
        if line.strip().endswith("PROGRAM TOTALS"):
            costs["PROGRAM TOTALS"] = int(line.split()[0].replace(",", ""))
            continue
        if "file:function" in line:
            started = True
            continue
        match = ROW.match(line) if started else None
        if match and "=>" not in line:
            name = match.group(2).split(":", 1)[-1]
            costs[name] = costs.get(name, 0) + int(match.group(1).replace(",", ""))
    return costs


def main(directory):
    per = {}
    for leg in ("before", "after"):
        zero = exclusive(f"{directory}/cg-{leg}-source_tiny-0.out")
        many = exclusive(f"{directory}/cg-{leg}-source_tiny-10000.out")
        per[leg] = {name: (many.get(name, 0) - zero.get(name, 0)) / 10000 for name in set(zero) | set(many)}
    totals = {leg: per[leg].get("PROGRAM TOTALS", 0) for leg in per}
    print(f"per audit: before {totals['before']:,.0f}, after {totals['after']:,.0f}, "
          f"added {totals['after'] - totals['before']:,.0f} instructions")
    print(f"{'before':>9} {'after':>9} {'added':>9}  function (or inlined source line)")
    names = sorted(set(per["before"]) | set(per["after"]),
                   key=lambda name: -(per["after"].get(name, 0) - per["before"].get(name, 0)))
    for name in names:
        if name in ("PROGRAM TOTALS", "events annotated"):
            continue
        before, after = per["before"].get(name, 0), per["after"].get(name, 0)
        if abs(after - before) >= 20:
            print(f"{before:>9.0f} {after:>9.0f} {after - before:>+9.0f}  {name[:140]}")


if __name__ == "__main__":
    main(sys.argv[1])
