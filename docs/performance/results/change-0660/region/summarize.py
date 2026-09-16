#!/usr/bin/env python3
"""Halve the N=3 minus N=1 difference of each commit-region isolation pair."""
import re, json, sys

def total(path):
    text = open(path, errors="replace").read()
    m = re.search(r"refs:\s+([\d,]+)", text)
    return int(m.group(1).replace(",", ""))

out = {}
for leg in ("before", "after"):
    for shape in ("compact", "noncompact"):
        for paragraphs in (24, 200, 10000):
            one = total(f"region-{leg}-{shape}-{paragraphs}-1.txt")
            three = total(f"region-{leg}-{shape}-{paragraphs}-3.txt")
            out.setdefault(shape, {}).setdefault(paragraphs, {})[leg] = (three - one) // 2
json.dump(out, open("region-summary.json", "w"), indent=1)
for shape in ("compact", "noncompact"):
    print(f"\n== {shape} source")
    for paragraphs in (24, 200, 10000):
        b = out[shape][paragraphs]["before"]
        a = out[shape][paragraphs]["after"]
        print(f"  {paragraphs:>6} paragraphs: {b:>12,} -> {a:>12,}  ({(a-b)/b*100:+.2f}%)")
