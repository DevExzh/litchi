#!/usr/bin/env python3
"""Change 0754: per one-edit iteration, how many times each xml-minifier audit
entry point is called and its inclusive instructions, from callgrind runs of
the probe's `loop one` at 1 and 3 iterations (difference / 2)."""
import re, subprocess, sys
PATTERN = re.compile(r"xml_minifier::audit::(verify_source_replacement|verify_source|verify_observed|verify_with_policy|VerifiedSource::verify_replacement|VerifiedSource::verify)$")

def calls(path):
    """Total call count per audit function, summed over every call arc."""
    names, counts = {}, {}
    current = None
    for line in open(path, errors="replace"):
        m = re.match(r"cfn=\((\d+)\)(?: (.*))?", line.strip())
        if m:
            fid, name = m.group(1), m.group(2)
            if name:
                names[fid] = name
            current = names.get(fid)
            continue
        m = re.match(r"calls=(\d+)", line.strip())
        if m and current:
            short = PATTERN.search(current)
            if short:
                counts[short.group(1)] = counts.get(short.group(1), 0) + int(m.group(1))
            current = None
    # callgrind also names functions in fn= lines; a cfn=(id) without a name
    # refers to an id defined by an earlier fn=/cfn=.
    return counts

def names_first(path):
    names = {}
    for line in open(path, errors="replace"):
        m = re.match(r"c?fn=\((\d+)\) (.*)", line.strip())
        if m:
            names[m.group(1)] = m.group(2)
    return names

def calls_full(path):
    names = names_first(path)
    counts = {}
    current = None
    for line in open(path, errors="replace"):
        line = line.strip()
        m = re.match(r"cfn=\((\d+)\)", line)
        if m:
            current = names.get(m.group(1))
            continue
        m = re.match(r"calls=(\d+)", line)
        if m and current:
            short = PATTERN.search(current)
            if short:
                counts[short.group(1)] = counts.get(short.group(1), 0) + int(m.group(1))
            current = None
    return counts

def inclusive(path):
    out = subprocess.run(["callgrind_annotate", "--inclusive=yes", "--threshold=100", path],
                         capture_output=True, text=True, check=True).stdout
    costs = {}
    for line in out.splitlines():
        m = re.match(r"\s*([\d,]+) \([^)]*\)\s+\S*?:(xml_minifier::audit::\S+) \[", line)
        if m:
            short = PATTERN.search(m.group(2))
            if short:
                costs[short.group(1)] = int(m.group(1).replace(",", ""))
    return costs

d = sys.argv[1]
print(f"{'leg':<7}{'entry point':<40}{'calls/iter':>12}{'Ir/iter':>16}")
for leg in ("PB", "PA"):
    c1, c3 = calls_full(f"{d}/cg-one-{leg}-1.out"), calls_full(f"{d}/cg-one-{leg}-3.out")
    i1, i3 = inclusive(f"{d}/cg-one-{leg}-1.out"), inclusive(f"{d}/cg-one-{leg}-3.out")
    for name in sorted(set(c3) | set(i3)):
        print(f"{'before' if leg == 'PB' else 'after':<7}{name:<40}{(c3.get(name, 0) - c1.get(name, 0)) / 2:>12.1f}{(i3.get(name, 0) - i1.get(name, 0)) / 2:>16,.0f}")
