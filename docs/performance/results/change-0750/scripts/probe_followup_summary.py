#!/usr/bin/env python3
"""Change 0750 follow-up: instructions per audit of the audit probe for the
base and the fix (scripts/probe_callgrind_followup.sh), next to the first
candidate's from callgrind/probe-summary.txt.

Usage: probe_followup_summary.py <cg-probe-followup dir> <callgrind/probe-summary.txt>"""
import sys

CASES = [("source/ws-structured", 4), ("source/ws-patriarch", 2), ("source/docx-drawing", 20),
         ("source/docx-table-alignment", 100), ("source/pptx-slide11", 40), ("pair/ws-structured", 4),
         ("source/corpus-accepted", 1), ("authored/ws-patriarch", 2), ("authored/tiny", 10000),
         ("source/tiny", 10000)]


def total(path):
    for line in open(path):
        if line.startswith(("summary:", "totals:")):
            return int(line.split()[1])
    raise ValueError(path)


def main(directory, summary):
    first = {}
    for line in open(summary):
        parts = line.split()
        if parts and "/" in parts[0] and len(parts) >= 5:
            first[parts[0]] = (int(parts[2].replace(",", "")), int(parts[3].replace(",", "")))
    print("Instructions per audit (callgrind; the probe at N iterations minus 0 iterations, divided by N).")
    print("base and fix: this run; first candidate: callgrind/probe-summary.txt (de69fb407d), whose base column")
    print("this run's base reproduces within 0.8%.")
    print()
    print(f"{'case':<30}{'base':>16}{'first candidate':>17}{'fix':>16}{'fix vs base':>13}{'fix vs first':>14}")
    for case, count in CASES:
        tag = case.replace("/", "_")
        per = {leg: (total(f"{directory}/cg-{leg}-{tag}-{count}.out") - total(f"{directory}/cg-{leg}-{tag}-0.out")) / count
               for leg in ("before", "after")}
        base, fixed = per["before"], per["after"]
        candidate = first[case][1]
        print(f"{case:<30}{base:>16,.0f}{candidate:>17,}{fixed:>16,.0f}{(fixed - base) / base * 100:>+12.2f}%"
              f"{(fixed - candidate) / candidate * 100:>+13.2f}%")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
