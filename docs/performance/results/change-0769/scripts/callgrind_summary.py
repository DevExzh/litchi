#!/usr/bin/env python3
"""Change 0769: summarize the Callgrind runs of the harness's
doc_semantic_open/large (3 samples, no warmup) for base, the first candidate
(v1) and the retained candidate: program totals, the inclusive cost of the
DOC parse and of every litchi_cfb function, and the largest per-function
self-cost differences. Reads callgrind/doc-large-ARM.out.

Usage: callgrind_summary.py > doc-semantic-open-large.json
"""

import collections
import json
import re
import subprocess

ARMS = ("base", "cand-v1", "cand")
LINE = re.compile(r"^\s*([0-9,]+) \([^)]*\)\s+(.*?)(?: \[.*\])?$")


def annotate(arm, inclusive):
    args = ["callgrind_annotate", "--threshold=100"]
    if inclusive:
        args.append("--inclusive=yes")
    text = subprocess.run(args + [f"callgrind/doc-large-{arm}.out"], capture_output=True, text=True, check=True).stdout
    costs = collections.Counter()
    total = None
    for line in text.splitlines():
        match = LINE.match(line)
        if not match:
            continue
        value = int(match.group(1).replace(",", ""))
        name = match.group(2).replace("???:", "")
        if name == "PROGRAM TOTALS":
            total = value
            continue
        costs[name] += value
    return total, costs


def main():
    report = {"case": "doc_semantic_open/large", "samples": 3, "warmup": 0, "arms": {}}
    selves = {}
    for arm in ARMS:
        total, inclusive = annotate(arm, True)
        _, self_costs = annotate(arm, False)
        selves[arm] = self_costs
        report["arms"][arm] = {
            "program_total_ir": total,
            "doc_parse_inclusive_ir": sum(value for name, value in inclusive.items() if "Document>::from_ole_with_options" in name and "closure" not in name),
            "litchi_cfb_inclusive_ir": {name: value for name, value in inclusive.items() if name.startswith("litchi_cfb::")},
        }
    diffs = []
    for name in set(selves["base"]) | set(selves["cand"]):
        delta = selves["cand"][name] - selves["base"][name]
        if delta:
            diffs.append((abs(delta), delta, name))
    diffs.sort(reverse=True)
    report["largest_self_cost_changes_cand_vs_base"] = [
        {"function": name, "delta_ir": delta, "base_ir": selves["base"][name], "cand_ir": selves["cand"][name]}
        for _, delta, name in diffs[:20]
    ]
    json.dump(report, __import__("sys").stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
