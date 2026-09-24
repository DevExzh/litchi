#!/usr/bin/env python3
"""Extract exact timed-region instruction counts from the 0768 callgrind runs.

Probe runs were collected with --toggle-collect on `timed_format`, so their
collected total is exactly the timed lifecycles. Harness runs were collected
in full; the timed region is the four body_text calls inside the sample loop
(Snapshot::edit, Edit::replace_paragraph, Edit::commit, Snapshot::finish),
read as inclusive costs. Both legs ran identical arguments, so any untimed
preflight calls of those functions are identical in both legs.

Usage: cg_extract.py CG_DIR OUT_JSON
"""

import json
import re
import subprocess
import sys

TIMED = [
    "litchi_doc::body_text::Snapshot::edit",
    "litchi_doc::body_text::Edit::replace_paragraph",
    "litchi_doc::body_text::Edit::commit",
    "litchi_doc::body_text::Snapshot::finish",
]
DETAIL = TIMED + [
    "litchi_doc::parts::protection::policy::classify",
    "litchi_doc::writer::fkp::ChpxFkpBuilder::generate_pages",
]


def collected(log):
    return int(re.search(r"Collected : (\d+)", open(log).read()).group(1))


def inclusive(out):
    text = subprocess.run(
        ["callgrind_annotate", "--inclusive=yes", "--threshold=100", out],
        capture_output=True, text=True, check=True,
    ).stdout
    costs = {}
    for line in text.splitlines():
        match = re.match(r"\s*([\d,]+)\s+\([^)]*\)\s+(\S.*?)(?: \[.*\])?$", line)
        if not match:
            continue
        name = match.group(2)
        # Drop the "file:" prefix callgrind_annotate prints ("???:" without
        # debug information, "path/file.rs:" with it).
        name = re.sub(r"^(?:\?\?\?|\S*?\.(?:rs|c|S|h)):(?!:)", "", name)
        value = int(match.group(1).replace(",", ""))
        costs[name] = max(costs.get(name, 0), value)
    return costs


def main():
    cg, out_json = sys.argv[1], sys.argv[2]
    result = {}
    for case, samples in [("docnohf", 8), ("docfloat", 4)]:
        legs = {leg: collected(f"{cg}/{case}-{leg}.log") for leg in "AB"}
        result[f"probe_{case}_format"] = {
            "method": "callgrind --toggle-collect=*timed_format*, warmups 0",
            "lifecycles": samples,
            "ir_total": legs,
            "ir_per_lifecycle": {leg: legs[leg] / samples for leg in legs},
            "after_over_before": legs["B"] / legs["A"],
        }
    for shape, samples in [("tiny", 8), ("large", 2)]:
        legs = {}
        for leg in "AB":
            costs = inclusive(f"{cg}/harness-{shape}-{leg}.out")
            legs[leg] = {name: costs.get(name, 0) for name in DETAIL}
            legs[leg]["timed_sum"] = sum(costs.get(name, 0) for name in TIMED)
            legs[leg]["program_total"] = collected(f"{cg}/harness-{shape}-{leg}.log")
        result[f"harness_doc_semantic_one_edit_save/{shape}"] = {
            "method": "callgrind full collection; inclusive Ir of the four timed body_text calls "
                      f"over one process with --samples {samples} --warmup 0 (includes identical preflight calls)",
            "legs": legs,
            "timed_sum_after_over_before": legs["B"]["timed_sum"] / legs["A"]["timed_sum"],
        }
    json.dump(result, open(out_json, "w"), indent=1)
    for key, value in result.items():
        if "ir_per_lifecycle" in value:
            a, b = value["ir_per_lifecycle"]["A"], value["ir_per_lifecycle"]["B"]
            print(f"{key}: {a:,.0f} -> {b:,.0f} Ir/lifecycle ({value['after_over_before'] - 1:+.2%})")
        else:
            a, b = value["legs"]["A"], value["legs"]["B"]
            print(f"{key}: timed {a['timed_sum']:,} -> {b['timed_sum']:,} ({value['timed_sum_after_over_before'] - 1:+.2%})")
            for name in DETAIL:
                print(f"    {name.split('::', 1)[1]}: {a[name]:,} -> {b[name]:,}")


if __name__ == "__main__":
    main()
