#!/usr/bin/env python3
"""Change 0750: attribute the harness isolation pairs' per-sample instructions
to the auditor. For each case and leg, the inclusive cost of every
`xml_minifier::audit::verify*` entry point (callgrind_annotate --inclusive=yes)
at --samples 3 minus --samples 1, halved, is the audit's cost per sample; the
callers of each entry point come from --tree=caller.

Usage: cg_attribute.py <directory with cg-<leg>-<case>-s<n>.out>"""
import re, subprocess, sys

CASES = ["sb_one_medium", "sb_one_dense", "docx_sb_one", "pptx_sb_one"]
LINE = re.compile(r"\s*([\d,]+) \([^)]*\)\s+\S*?:(xml_minifier::audit::verify\S*) \[")


def inclusive(path):
    out = subprocess.run(["callgrind_annotate", "--inclusive=yes", "--threshold=100", path],
                         capture_output=True, text=True, check=True).stdout
    costs = {}
    for line in out.splitlines():
        match = LINE.match(line)
        if match:
            costs[match.group(2)] = int(match.group(1).replace(",", ""))
    return costs


def main(directory):
    print(f"{'case':<15}{'entry point':<48}{'before/sample':>16}{'after/sample':>16}{'delta':>13}{'change':>9}")
    for case in CASES:
        runs = {(leg, s): inclusive(f"{directory}/cg-{leg}-{case}-s{s}.out")
                for leg in ("before", "after") for s in (1, 3)}
        names = sorted(set().union(*runs.values()))
        for name in names:
            before = (runs["before", 3].get(name, 0) - runs["before", 1].get(name, 0)) / 2
            after = (runs["after", 3].get(name, 0) - runs["after", 1].get(name, 0)) / 2
            if abs(before) < 50_000 and abs(after) < 50_000:
                continue  # entry points that run once per process, not per sample
            print(f"{case:<15}{name.replace('xml_minifier::audit::', ''):<48}{before:>16,.0f}{after:>16,.0f}"
                  f"{after - before:>+13,.0f}{(after - before) / before * 100:>+8.1f}%")


if __name__ == "__main__":
    main(sys.argv[1])
