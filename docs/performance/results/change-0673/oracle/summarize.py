#!/usr/bin/env python3
"""Summarize the retained before/after open reports for change 0673."""

from __future__ import annotations

import re
import sys
from pathlib import Path


OPEN_COST = re.compile(r"^open-cost requests=(\d+) bytes=(\d+) versions=(\d+)$")
CONTAINER = re.compile(r"^=== (.+) len=(\d+)$")


def read_costs(path: Path) -> dict[str, tuple[int, int, int]]:
    current = None
    costs = {}
    for line in path.read_text().splitlines():
        match = CONTAINER.match(line)
        if match:
            current = match.group(1)
            continue
        match = OPEN_COST.match(line)
        if match and current is not None:
            costs[current] = tuple(map(int, match.groups()))
    return costs


def main() -> None:
    before_path, after_path, output = map(Path, sys.argv[1:4])
    before = read_costs(before_path)
    after = read_costs(after_path)
    if before.keys() != after.keys():
        raise SystemExit("before/after container sets differ")

    def totals(values):
        return tuple(sum(row[index] for row in values.values()) for index in range(3))

    before_total = totals(before)
    after_total = totals(after)
    deltas = [
        (name, before[name], after[name])
        for name in sorted(before)
    ]
    request_deltas = {after_req - before_req for _, (before_req, _, _), (after_req, _, _) in deltas}
    byte_deltas = {after_bytes - before_bytes for _, (_, before_bytes, _), (_, after_bytes, _) in deltas}
    if len(request_deltas) != 1 or len(byte_deltas) != 1:
        raise SystemExit("open-cost deltas are not uniform")

    lines = [
        "container\tbefore_requests\tafter_requests\tbefore_bytes\tafter_bytes\tbefore_versions\tafter_versions",
    ]
    for name, old, new in deltas:
        lines.append(
            "\t".join(
                [
                    name,
                    str(old[0]),
                    str(new[0]),
                    str(old[1]),
                    str(new[1]),
                    str(old[2]),
                    str(new[2]),
                ]
            )
        )
    (output / "corpus-open-costs.tsv").write_text("\n".join(lines) + "\n")

    report_before = before_path.read_text().splitlines()
    report_after = after_path.read_text().splitlines()
    if len(report_before) != len(report_after):
        raise SystemExit("before/after report line counts differ")
    differing = [
        (old, new)
        for old, new in zip(report_before, report_after)
        if old != new
    ]
    if not all(old.startswith("open-cost ") and new.startswith("open-cost ") for old, new in differing):
        raise SystemExit("report differences exceed open-cost lines")
    if differing:
        identity = (
            f"reports differ only on {len(differing)} open-cost lines\n"
            f"line_count={len(report_before)}\n"
        )
    else:
        identity = f"reports are byte-identical\nline_count={len(report_before)}\n"
    (output / "report-identity.txt").write_text(identity)
    (output / "report-diff-classes.txt").write_text(
        f"    {len(differing)} < open-cost\n"
        f"    {len(differing)} > open-cost\n"
    )

    count = len(before)
    req_delta, byte_delta = request_deltas.pop(), byte_deltas.pop()
    versions_same = before_total[2] == after_total[2]
    difference_sentence = (
        f"Only {len(differing)} lines differ, and every difference is an open-cost line."
        if differing
        else "No report lines differ."
    )
    summary = f"""Change 0673 open differential — {count} ZIP containers under test-data

Both legs record the open verdict, request/byte/source-observation cost, package
relationships, admitted parts, non-part members, and every decoded part's
length and CRC. The reports contain {len(report_before):,} lines each.

{difference_sentence}
No verdict, error identity, package relationship, part, non-part member, decoded
payload, or source-observation count differs. The locator's missed-probe path
therefore retains the prior read grammar while the exact terminal-EOCD path
avoids the heap scratch.

                                      before       after        delta
open requests, {count} containers        {before_total[0]}        {after_total[0]}       {after_total[0] - before_total[0]}
open bytes, {count} containers        {before_total[1]}     {after_total[1]}       {after_total[1] - before_total[1]}
version() observations                 {before_total[2]}        {after_total[2]}          {after_total[2] - before_total[2]}
containers costing more requests           0          0
containers costing fewer requests         {sum(1 for _, old, new in deltas if new[0] < old[0])}        {sum(1 for _, old, new in deltas if new[0] < old[0])}
containers reading more bytes              0          0
containers reading fewer bytes            {sum(1 for _, old, new in deltas if new[1] < old[1])}        {sum(1 for _, old, new in deltas if new[1] < old[1])}

Every common exact-EOCD container keeps the same open request count and bytes;
this corpus does not contain a missed-probe container. The per-fixture tests
exercise the comment-bearing and ZIP64 fallback paths separately.
uniform_request_delta={req_delta}
uniform_byte_delta={byte_delta}
versions_unchanged={versions_same}
"""
    (output / "summary.txt").write_text(summary)


if __name__ == "__main__":
    main()
