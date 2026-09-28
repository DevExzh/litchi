# 0820 reader review — unfrozen handoff

This is an unfrozen reader handoff. It is not part of `FROZEN_PACKET_FILES`
and does not bind the quality or build gates. It records transient reader
findings so they are not mistaken for production-source or protocol findings.

## Transient findings during driver edits

While the 0820 drivers were being replaced, the reader set temporarily had
the following issues:

* `analyze.py` had an indentation error at line 1050.
* `validate.py` had a duplicated `and all(...)` predicate, the old observer
  aggregate cardinality, and stale 0819 cleanup/seal schemas.
* The admission reader indexed `ordinary["save_durability"]` for the default
  route even though serde omits that field when the value is absent. The
  default route now requires strict field absence, while explicit policies use
  an exact value check.

These were transient implementation findings, not capture evidence. They
must not be cited as measured failures or used to interpret any 0820 result.

## Current static handoff

After the final driver handoff, `python3 -B -m py_compile` passes for the
seven 0820 readers and the stale-state scan is clear for 0819 counts,
schemas, targets, scratch paths, and report cardinalities. The admission
reader has the strict default omission check, and validation carries the
216-report/4,488-sample and 0820 observer cardinalities.

No capture has run at this handoff. The analysis reader still requires its
post-capture review and must be checked against the actual reports before
results are interpreted. The root runner owns those checks and the heavy
execution; this review performed no Cargo, binary, workload, or reader
execution that could create result evidence.

## Required post-capture checks

After receipts and reports exist, the reader owner should run the planned
analysis and validation commands, inspect their schemas and counts, and
confirm that no reader-generated `__pycache__` or other temporary files enter
the packet. At minimum:

```text
python3 -B analyze.py --write
python3 -B analyze.py --check
python3 -B validate.py
```

Reader completion is required before interpreting captures. It is independent
of the frozen production source and protocol review, and this file should not
be sealed as a frozen packet input unless the root reviewer explicitly changes
that decision.

## Final static reader review

The finalized `reader-notes.md` agrees with the reader contracts checked here.
AST parsing passes for `analyze.py`, `validate.py`, `cleanup.py`, and
`seal.py`. The plan has 24 unique selector IDs. `custody.ordered_cases` emits
24 unique rows for every qualification, native, and observer block using the
declared forward/reverse group order and rotated policy order. The six native
format/phase groups therefore produce 18 non-default policy comparisons
(full, file-only, and no-sync); the six default rows remain self-ratio control
rows and are not an independent policy comparison.

The reader's policy fields match the producer: default requires strict absence
of `save_durability`, explicit policies require their exact level, and
`atomic_publication_steps` carries the policy-specific route. The generic
`timing_scope` values match `Phase::timing_scope()` and are treated as phase
metadata. Each native paired statistic uses six matched block p50 ratios and
the specified seed, resample count, and ranks. Observer and qualification
reports remain separate from native elapsed summaries, and the 32 procfs
controls are retained without subtraction.

One terminal integration issue remains outside this reader: `cleanup.py`
currently requires `results-review.md`, while the finalized handoff file is
`reader-notes.md` and no `results-review.md` exists in the packet. Root must
resolve that filename contract before cleanup and sealing. The cleanup binary
records themselves are path-sorted and matched to build descriptors; `seal.py`
would include all packet files, including reader notes, in its sealed map.
