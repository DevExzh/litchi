#!/usr/bin/env python3
"""Reconcile exact sampled-IP periods with perf's flat self-symbol report."""
import argparse
import collections
import hashlib
import json
from pathlib import Path
import re

P = Path(__file__).resolve().parent
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--output", type=Path, default=P / "self-analysis.json")
args = parser.parse_args()
LEAF = re.compile(r"^\d+/\d+\s+[\d.]+:\s+(\d+) cycles:u:\s+([0-9a-f]+) (.*?) \((.*)\)$")
HEADER = re.compile(r"^\d+/\d+\s+[\d.]+:\s+(\d+) cycles:u:")

def key(dso, symbol):
    if dso == "[unknown]":
        return ("[unknown]", "[unknown]")
    if symbol == "[unknown]" or re.fullmatch(r"0x[0-9a-f]+", symbol):
        return (Path(dso).name, "[unknown]")
    return (Path(dso).name, symbol)

def group(dso, symbol):
    if dso == "[unknown]" or symbol == "[unknown]" or symbol.startswith("0x"):
        return "unresolved_sampled_ip"
    for prefix, name in (
        ("litchi_ooxml_common::mce::", "mce"),
        ("litchi_xlsx::raw::worksheet::x14ac::", "x14ac"),
        ("quick_xml::name::NamespaceResolver::", "xml_namespace_resolver"),
        ("quick_xml::reader::ns_reader::", "xml_namespace_reader"),
        ("quick_xml::", "other_quick_xml"),
        ("litchi_xlsx::raw::worksheet::", "other_worksheet_parser"),
        ("zlib_rs::", "zlib"),
    ):
        if symbol.startswith(prefix):
            return name
    return "other_resolved"

rows = []
for folder in sorted(P.glob("child-*")):
    leaves = collections.Counter()
    leaf_counts = collections.Counter()
    for line in (folder / "leaf-samples.txt").read_text().splitlines():
        match = LEAF.fullmatch(line.strip())
        if match is None:
            raise ValueError("unparsed sampled IP: " + line)
        period, _, symbol, dso = match.groups()
        k = key(dso, symbol)
        leaves[k] += int(period)
        leaf_counts[k] += 1
    report = (folder / "self-symbols.txt").read_text()
    if "# Total Lost Samples: 0" not in report:
        raise ValueError("lost samples require investigation")
    reported = collections.Counter()
    reported_counts = collections.Counter()
    for line in report.splitlines():
        if line.lstrip().startswith("#") or "|" not in line:
            continue
        fields = [value.strip() for value in line.split("|")]
        _, count, period, dso, symbol = fields[:5]
        symbol = re.sub(r"^\[.\] ", "", symbol)
        k = key(dso, symbol)
        reported[k] += int(period)
        reported_counts[k] += int(count)
    if leaves != reported or leaf_counts != reported_counts:
        raise ValueError("sampled-IP and self-symbol views disagree: " + folder.name)
    total = sum(leaves.values())
    if total != int(re.search(r"Event count \(approx\.\): (\d+)", report)[1]):
        raise ValueError("event period total mismatch")
    groups = collections.Counter()
    for (dso, symbol), period in leaves.items():
        groups[group(dso, symbol)] += period
    stack_period = empty_period = unknown_stack_period = 0
    for block in (folder / "call-stacks.txt").read_text().strip().split("\n\n"):
        lines = block.splitlines()
        period = int(HEADER.match(lines[0])[1])
        stack_period += period
        if len(lines) == 1:
            empty_period += period
        if any("[unknown]" in line for line in lines[1:]):
            unknown_stack_period += period
    if stack_period != total:
        raise ValueError("stack sample population differs")
    rows.append({"child": folder.name, "sample_count": sum(leaf_counts.values()), "total_period": total,
        "lost_samples": 0, "self_groups": {name: {"period": period, "percent": period / total * 100} for name, period in groups.items()},
        "empty_stack_percent": empty_period / total * 100, "any_unknown_stack_percent": unknown_stack_period / total * 100,
        "top_self_symbols": [{"dso": k[0], "symbol": k[1], "period": period, "samples": leaf_counts[k], "percent": period / total * 100}
                             for k, period in leaves.most_common(25)],
        "input_sha256": {name: hashlib.sha256((folder / name).read_bytes()).hexdigest()
                         for name in ("leaf-samples.txt", "self-symbols.txt", "call-stacks.txt")}})
result = {"schema": "litchi.xlsx-scan-self-cycles-0886.v1", "status": "pass", "rows": rows,
          "scope": "Exact period-weighted sampled-IP self attribution; categories are disjoint. No inclusive or caller cost inferred from incomplete stacks. Guest instruction observations must remain separate."}
with args.output.open("x") as f:
    json.dump(result, f, indent=2)
    f.write("\n")
print(json.dumps([{k: v for k, v in row.items() if k not in ("top_self_symbols", "input_sha256")} for row in rows], indent=2))
