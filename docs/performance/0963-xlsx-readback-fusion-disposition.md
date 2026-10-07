# 0963 — reduced-readback fusion does not justify implementation

Status: source and retained-evidence review complete. The proposed omission-aware
shared traversal is declined before implementation. Production is unchanged;
there is no new workload, baseline, or speedup claim. The broader non-iWork
performance goal remains active.

## Route census

The [source/evidence review](results/change-0963/readback-review.md) confirms
that all three 0962 cases use reduced readback and merge source cells:

| Case | Touched-sheet owners | Updates | Omitted owners | Merges per iteration |
| --- | ---: | ---: | ---: | ---: |
| Medium 1% | 9,216 | 93 | 9,123 | 4 |
| Dense-sparse 1% | 17,792 | 178 | 17,614 | 4 |
| Medium one edit | 2,304 | 1 | 2,303 | 1 |

The 0958 three-sample allocation captures reproduce the omitted-owner counts
times three in `Stored::clone`, and direct merge callsites reproduce 12/12/3
calls. This combines allocation route witnesses with current source inspection;
it is not a fresh runtime route counter or a CPU ranking. One-edit produces no
replacement hints but still has omissions: those two lists are not equivalent.

## What fusion would actually remove

`from_rewritten_value_source` validates the complete output, copies the
non-omitted ranges into a reduced document, parses that document, then merges
unchanged source cells. A shared full-output reader must still feed every event
to the authoritative validator. It can suppress semantic parser transitions
only for proven complete omitted-cell subtrees. It could avoid the current
append-and-sort merge call only by doing equivalent source-record carry and
index work while retaining an independent immutable store; those merge
semantics remain.

The [replayable subtraction](results/change-0963/readback-residual.json) removes
the nested validator from the rewritten-readback ancestry in the frozen 0962
packet. It leaves 810 samples / 3,550,864,288 period weights, or 6.7184% of pooled
qualified stage period. Per-child shares range from 4.2524% to 7.2645%. This
residual includes merging and other retained work; it is neither removable
cost nor an elapsed-time or end-to-end speedup bound. No multiplication by the
separately measured native phase fractions is used.

The older [0602](0602-xlsx-real-producer-admission-design.md) isolation profiles
also found reduced copying small on four real fixtures. Their guest-instruction
population, source revision and operations differ, so their percentages are
context only and do not quantify the current reduced parser.

Fusion would add omission-range checks and state on the full stream, retain
validator and parser state concurrently, and require conservative handling of
event overlap, row boundaries, MCE offsets, parser failures and allocation
refusals. Current evidence does not isolate enough removable work to justify
that complexity. This is a prioritization decision before coding, not a measured
rejection of a compiled candidate and not proof that every possible fusion
would be slower.

## Semantic boundaries and alternatives

An eventual implementation must preserve validator-first error precedence,
complete XML validation, exact source identity, unchanged-cell provenance,
typed failures, cancellation and budgets, MCE behavior, staged readback,
deterministic output and inverse restoration. Uncertain cases must retain the
current validator → reduced parse/merge → complete-parse fallback. A validator
bypass is not this candidate and has no adopted equivalence proof.

The [attribute review](results/change-0963/attribute-review.md) checks a second
idea against prior trials. Short-prefix duplicate-check changes are not a new
mechanism: 0799–0806 already tested them, and 0806 failed public-workflow benefit
qualification. Namespace declaration discovery and checked attribute validation
also have different error and scope-installation rules. The 0962 sampled
iterator leaf cannot by itself justify repeating those rejected mechanisms.

The next investigation should stay with current source admission, the larger
qualified ancestry owner, and isolate specific redundant event/namespace work
before proposing another implementation. Older producer-MCE rankings require
their later history and current route to be checked; 0587 alone is not current
qualification. No new optimization is selected solely from allocation traffic.

## Verification and cleanup

Reproduce the bounded residual with:

```sh
python3 docs/performance/results/change-0963/reduce_residual.py
```

The helper binds the committed 0962 input by size and SHA-256 and checks the
nested inequalities plus child/pooled sample and period conservation. Independent
review covers route witnesses, scope and arithmetic. No Rust build or correctness
suite is claimed for this documentation-only batch. The clean detached candidate
worktree was removed before coding; no candidate binary, build cache or profiler
output was created. Unrelated worktree changes remain untouched.

ADR 0001/0005 support prioritization by measured scope; 0003/0006 retain readback
and validation authority; 0008 requires reproducible evidence and scoped claims;
0032 does not authorize a new parsed-tree cache. No runtime architecture or
dependency boundary changes are proposed.
