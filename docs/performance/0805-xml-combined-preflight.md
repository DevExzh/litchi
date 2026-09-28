# 0805 — combined checked-attribute candidate versus production

The candidate passes the frozen preflight; production adoption remains unproven.

This fresh preflight compares current production with the exact 0804 candidate
source. The candidate combines the no-replay short prefix, first-request empty
check, bounded linear byte equality and ordered-map fallback. Historical
diagnostic ratios are not multiplied or pooled. Production remains unchanged
while this comparison is evaluated.

## Frozen policy and scope

The [packet](results/change-0805/README.md) preserves the 39-case fixture matrix
and 18 protected consume cases from 0802. Advancement requires at least 3%
benefit for both one- and two-attribute consumption, with ratio intervals
wholly below 1. A protected consume regression above 5% with its interval
wholly above 1 vetoes advancement. Other regressions remain review triggers.
Passing can authorize only fresh workflow, resource and cross-format trials;
this packet cannot adopt production code.

Six alternating native block pairs use 30 samples, three warmups and 4,096
iterations per child on CPU 12. Process p50 uses nearest rank; paired median
ratios use 10,000 bootstrap resamples with fresh seed 805080. Two separate
Callgrind repeats retain guest instruction and branch counters. The build
records a symbol witness and checks that all four requested owners exist.

Construction includes common dispatch, `black_box(&iterator)`, drop and
checksum work. Consumption includes construction and checksum work. These
hot repeated direct-helper inputs provide no workflow latency, heap-use,
hardware-branch-rate or allocation API result. Guest instruction changes do
not uniquely explain native timing; inclusive descendant costs overlap.

## Correctness and results

Both mirror legs pass warnings-denied Clippy and their exact helper tests:
70 production and 100 candidate tests, with no failures or ignores. The
release probe passes formatting, build, check, Clippy, independent fixture
verification and semantic self-check across all 39 cases, including clone
advances 0, 1, 2, 3, 4, 5, 32 and 33. These are isolated helper/probe gates,
not full production-workspace verification. Whole iterator size is measured
as 120 bytes for production and 128 for the candidate; this is not heap use.

The candidate avoids replaying earlier values, retains duplicate-before-value
raw-key checks, and bounds linear checking to 32 names before ordered fallback.
Its test comparison counter remains active for both linear and ordered checks.
The candidate archive is exact 0804 source, including inherited 0802 comments
and a stale immediate-map sentence; implementation and source review show the
actual bounded array. No source comment is treated as performance evidence.

The candidate passes the frozen preflight and is eligible for fresh workflow
trials. Both dominant-class benefit gates pass, and no protected consume case
triggers a veto. This is not adoption: four non-protected consumption cases
regress 10.726–28.682%, and two construction rows also trigger diagnostic flags.

| Consume case | After/before p50 | Change | 95% interval |
| --- | ---: | ---: | ---: |
| distinct-0 | 0.318976 | -68.102% | 0.318807–0.336264 |
| distinct-1 | 0.613544 | -38.646% | 0.588405–0.617521 |
| distinct-2 | 0.809200 | -19.080% | 0.791095–0.809967 |
| distinct-3 | 0.936681 | -6.332% | 0.917889–0.949648 |
| distinct-4 | 1.226491 | +22.649% | 1.198300–1.243788 |
| distinct-8 | 1.027703 | +2.770% | 1.023107–1.039633 |
| distinct-16 | 0.988219 | -1.178% | 0.984651–1.008824 |
| distinct-32 | 1.032739 | +3.274% | 1.012950–1.054385 |
| distinct-33 | 0.556748 | -44.325% | 0.537645–0.572747 |
| distinct-64 | 0.989845 | -1.015% | 0.985159–1.011261 |

All six flagged rows are retained below. None of the four consumption rows
belongs to the frozen 18-case protected set; their absence from that set does
not establish that they are harmless in real documents.

| Case | Mode | Ratio | 95% interval |
| --- | --- | ---: | ---: |
| distinct-17 | construct | 1.136238 | 1.036975–1.159219 |
| distinct-3 | construct | 1.129848 | 1.000927–1.164418 |
| distinct-4 | consume | 1.226491 | 1.198300–1.243788 |
| duplicate-valid-after-4 | consume | 1.107260 | 1.095358–1.119992 |
| syntax-equals-value-after-4 | consume | 1.286817 | 1.280224–1.297670 |
| syntax-flag-after-4 | consume | 1.286344 | 1.209093–1.299985 |

There are 41 process-p50 spread flags among 156 case/mode/leg groups: 33
construction and eight consumption. Every raw sample, paired ratio, interval
and spread remains in [summary.md](results/change-0805/summary.md) and its
machine-readable analysis. The small increases at eight and 32 attributes
also remain visible despite falling below the diagnostic median threshold.

Both Callgrind repeats agree on the selected instruction totals below. All
312 exact owners qualify with one incoming call, and all five counters conserve
across 312 positive and 312 empty termination dumps. All four named owners
were present in the build witness; no scope amendment or failed capture was
needed in this packet.

| Case and mode | Production Ir | Candidate Ir |
| --- | ---: | ---: |
| distinct-0/construct | 70 | 70 |
| distinct-0/consume | 170 | 122 |
| distinct-1/consume | 574 | 407 |
| distinct-2/consume | 910 | 870 |
| distinct-4/consume | 1,702 | 2,372 |
| distinct-16/consume | 9,428 | 10,010 |
| distinct-32/consume | 25,562 | 26,130 |
| distinct-33/consume | 50,231 | 27,487 |
| syntax-flag-after-0/consume | 253 | 274 |
| syntax-equals-value-after-4/consume | 1,798 | 2,633 |

The four-attribute regression is consistent with additional work at the first
bounded-array stage: its guest instruction total rises by 670. This is a
mechanism diagnostic, not unique attribution of native time or an allocator
measurement. For the distinct-33 fixture, the candidate proves the remaining
tail is whitespace and finishes without building an ordered map, while
production reparses its first 32 attributes and seeds the map. When more tail
remains after the 33rd item, the candidate still seeds a fresh ordered map from
stored names; it avoids parser replay, not map insertion work. The selected
33-attribute result cannot isolate either saving. End-to-end measurements must
determine whether the short-tag benefit outweighs the remaining costs.

## Disposition and next qualification

All 1,248 reports and 28,392 samples pass independent semantic/checksum and
numerical replay. Both minimal quality legs and the release probe pass; no
failed setup, quality, build, native or profile attempt occurred. The decision
records `advance_to_workflow_trials: true` and `production_adoption: false`.

The next qualification must compare fresh production/candidate builds over
public PPTX capture/commit/lifecycle, resource guards and the existing cross-format
matrix, with the dedicated four-attribute fixture specified in
[next-trials-review.md](results/change-0805/next-trials-review.md). The
four-attribute costs remain explicit review triggers. Neither the
0798 dominant short-tag census nor these hot helper ratios proves a public
workflow gain. No current baseline is replaced until those gates pass.

The derived debug output exposes the new private state and disabled nested
quick-xml checking; item/error/clone/fused semantics, public signatures and
visibility remain covered by the source review and tests. All 9,196 production
source hashes, 35 architecture inputs, unrelated files and other worktrees
remain unchanged. No memory, cold/range, producer, concurrency or CRUD
coverage is promoted. iWork is excluded and the broader GOAL remains incomplete.

Owned-target cleanup removed 178,441,816 logical bytes after exact executable
identity verification. The final binary witness is retained; no failed-build
binary existed. Post-cleanup replay and final sealing verify the retained
evidence and unchanged production source census.
