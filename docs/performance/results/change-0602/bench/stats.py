#!/usr/bin/env python3
"""Per-leg quantiles and paired deltas for change 0602's timing legs.

A1/B1/B2/A2 are set/insert/insert/set: A is a one-cell replacement, where
change 0525's reduced readback applies; B is a one-cell insert on an absent
row, which leaves the rewrite's omission list empty and takes the complete
candidate parse. S1..S4 are four identical `set` legs, so S1/S2 and S3/S4 give
the A/A floor measured in the same window.
"""
import os, statistics, sys

def q(xs, p):
    xs = sorted(xs)
    if not xs:
        return 0
    i = min(len(xs) - 1, int(round(p * (len(xs) - 1))))
    return xs[i]

def load(path):
    if not os.path.exists(path):
        return []
    return [int(l) for l in open(path) if l.strip().isdigit()]

def row(tag, leg, xs):
    return "%-6s %-3s n=%-3d p50=%-12d mean=%-12d p95=%-12d p99=%-12d" % (
        tag, leg, len(xs), q(xs, .50), int(statistics.fmean(xs)) if xs else 0,
        q(xs, .95), q(xs, .99))

def delta(a, b):
    pa, pb = q(a, .50), q(b, .50)
    return 0.0 if pa == 0 else 100.0 * (pb - pa) / pa

def main():
    d = sys.argv[1]
    for tag in ("fct", "dvtr", "sss", "ndp"):
        legs = {l: load(os.path.join(d, "%s-%s.txt" % (tag, l)))
                for l in ("A1", "B1", "B2", "A2", "S1", "S2", "S3", "S4")}
        if not legs["A1"]:
            continue
        print("==== %s ====" % tag)
        for l in ("A1", "B1", "B2", "A2", "S1", "S2", "S3", "S4"):
            if legs[l]:
                print(row(tag, l, legs[l]))
        print("  set->insert  A1->B1 %+.2f%%   A2->B2 %+.2f%%" %
              (delta(legs["A1"], legs["B1"]), delta(legs["A2"], legs["B2"])))
        print("  insert->set  B1->A1 %+.2f%%   B2->A2 %+.2f%%" %
              (delta(legs["B1"], legs["A1"]), delta(legs["B2"], legs["A2"])))
        if legs["S2"]:
            print("  A/A floor    S1->S2 %+.2f%%   S3->S4 %+.2f%%   S1->S4 %+.2f%%" %
                  (delta(legs["S1"], legs["S2"]), delta(legs["S3"], legs["S4"]),
                   delta(legs["S1"], legs["S4"])))
        print()

main()
