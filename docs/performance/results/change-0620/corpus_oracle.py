#!/usr/bin/env python3
"""The correctness oracle for change 0620, over the retained corpus runs.

Three independent comparisons, in increasing strictness:

1. refusal identity on every row;
2. every reported field except the artifact digest, on every row;
3. the artifact digest, on every row whose digest is reproducible.

The exclusion list for (3) is derived from the 16-run repetition retained as
`corpus/nondeterminism-{before,after}.jsonl`, not from the three-run sweep: the
generic `Transaction::commit` re-renders the CFB directory in `HashSet`
iteration order, so two fixtures publish different bytes on every process and
three runs are not enough to notice.

Usage: corpus_oracle.py [<packet dir>]
"""
import collections
import json
import sys
from pathlib import Path

here = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent)
corpus = here / "corpus"


def load(leg, run):
    rows = {}
    for line in open(corpus / f"corpus-{leg}-r{run}.jsonl"):
        row = json.loads(line)
        rows[(row["path"], row["operation"])] = row
    return rows


before = [load("before", r) for r in (1, 2, 3)]
after = [load("after", r) for r in (1, 2, 3)]
keys = set(before[0])
assert all(set(x) == keys for x in before + after), "row sets differ between runs"

repeats = collections.defaultdict(lambda: collections.defaultdict(set))
for leg in ("before", "after"):
    for line in open(corpus / f"nondeterminism-{leg}.jsonl"):
        row = json.loads(line)
        if "sha256" in row:
            repeats[leg][(row["path"], row["operation"])].add(row["sha256"])
unstable = {k for leg in repeats for k, v in repeats[leg].items() if len(v) > 1}


def without_digest(row):
    return json.dumps({k: v for k, v in row.items()
                       if k not in ("path", "operation", "sha256")}, sort_keys=True)


refusals = [k for k in keys if before[0][k].get("refused") != after[0][k].get("refused")]
fields = [k for k in keys if without_digest(before[0][k]) != without_digest(after[0][k])]
digests = [k for k in keys - unstable
           if before[0][k].get("sha256") != after[0][k].get("sha256")]

opened = sum(1 for k in keys if k[1] == "open" and "refused" not in before[0][k])
refused = sum(1 for k in keys if "refused" in before[0][k])
print(f"fixtures 126, opened {opened}, refused at open {126 - opened}")
print(f"rows per run {len(keys)} ({refused} typed refusals, {len(keys) - refused} publishing)")
print(f"rows compared in total {6 * len(keys)}")
print(f"oracle 1, refusal-text mismatches: {len(refusals)}")
print(f"oracle 2, non-digest field mismatches: {len(fields)}")
print(f"oracle 3, digest mismatches over {len(keys) - len(unstable)} reproducible rows: {len(digests)}")
print(f"          rows excluded as not reproducible within one binary: {len(unstable)}")
for k in sorted(unstable):
    print(f"            {k[1]:16s} {k[0]}")
