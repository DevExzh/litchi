# Change 0536: reject the CFB collector cold-error-layout candidate

Date: 2026-09-12
Status: rejected; the measured runtime was restored to the baseline
Performance claim: none

## Decision

The candidate moved the eight formatted diagnostics in
`SectorChainScratch::collect_exact` into private `#[cold] #[inline(never)]`
`chain_error_*` helpers. It left `CheckedBitSet`, the successful collector
loop, validation order, reset behavior, allocation labels, and all test and
harness source unchanged. The source review found the intended narrow layout
boundary and no contract change, but the native admission rule rejected the
candidate: every primary row improved in p50, while only four of eight rows
reached the required 3% in both paired repeats.

| Primary workflow | Repeat 1 p50 (baseline → candidate) | Repeat 2 p50 (baseline → candidate) |
| --- | ---: | ---: |
| `xls_source_backed_open` | 110,910 → 106,960 (**3.5614% improvement**, pass) | 107,245 → 105,000 (**2.0933% improvement**, fail) |
| `xls_source_backed_open_one_cell` | 113,385 → 112,500 (**0.7805% improvement**, fail) | 111,850 → 105,710 (**5.4895% improvement**, pass) |
| `xls_owned_source_open` | 99,650 → 97,910 (**1.7461% improvement**, fail) | 97,350 → 94,440 (**2.9892% improvement**, fail) |
| `xls_owned_source_open_one_cell` | 102,905 → 96,035 (**6.6761% improvement**, pass) | 102,185 → 95,640 (**6.4050% improvement**, pass) |

The matched native lane contains 48,000 samples: 24,000 per stage across
nine XLS and three CFB scenarios. The frozen rule requires all four named
workflows to pass in both repeats, so the four passing rows cannot authorize
adoption. CFB p50 behavior is diagnostic only; one `few-large` repeat is
slower and no broad CFB claim follows.

## Source and evidence custody

The experiment is bound to revision
`8876e87b8dbcb7dca54a3416386bbfc892c68eb4`. The formatted candidate patch is
one production-file diff with no test or harness hunk, SHA-256
`29f96e9a19d8eb34508b84a84d1ac535bc970ace3e91d78f41d9a791ec05d9ec`; its
formatted candidate source file has SHA-256
`d8e322fb598950f157e9d290fbaf57567d0d7b3d9d5ec2c8d6a943becfe86a02`.
The baseline and restored source file SHA-256 is
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`, as
bound by the [source review](../results/change-0536/source-review.md) and
the baseline/final manifests. The candidate normal binary grew from
60,139,544 to 60,142,152 bytes; this layout effect is not a retained
performance result.

The primary [comparison](../results/change-0536/comparison.json),
[decision](../results/change-0536/decision.json),
[profile comparison](../results/change-0536/profile-comparison.json), and
[instruction comparison](../results/change-0536/instruction-analysis-comparison.json)
are the authoritative evidence. The [adverse review](../results/change-0536/adverse-review.json)
and [variation review](../results/change-0536/variation-review.json) retain
all 32 matched over-5% flags and 60 same-build variations with arithmetic
reviews; no row, sample, or outlier is discarded and no cause is inferred.

## Allocation and profile guards

The separate allocation lane contains 1,440 samples. Its four selected guard
vectors—`allocation_calls`, `reallocation_calls`, `allocated_bytes`, and
`incremental_region_peak_live_bytes`—are exactly identical in all 24 compared
rows. This is a guard result, not an allocation speedup or memory gain claim.

The profile evidence contains 16 children, 80 timed constructor dumps, and
12 setup dumps. XLS-owned constructor inclusive Ir increases from 11,316,236
to 11,316,310 (+0.0006539%) in repeat 1 and from 11,318,721 to 11,321,634
(+0.0257361%) in repeat 2, so the required constructor profile gate fails.
Collector exclusive self Ir falls only 0.0046419% in both XLS-owned repeats
(5,601,140 → 5,600,880) and 0.0017947% in both CFB `few-large` repeats
(5,571,945 → 5,571,845). Instruction mapping records `collect_exact`
shrinking from 1,436 to 1,238 bytes and 313 to 277 instructions, with the
eight helper symbols present in the candidate. These are code-shape and
Callgrind diagnostics; they do not establish native latency or causality and
cannot override the native gate.

## Quality, compatibility, and disposition

The candidate-stage [quality summary](../results/change-0536/candidate-quality-summary.json)
passes 14 checks with 4,382 executed tests. No Rust test or harness change is
part of the candidate or retained runtime. The final source patch is empty
against the restored baseline; restored-source quality also passed all 14 gates and 4,382 executions, for
8,764 new executions in this batch. Five verifier tamper probes pass and all
96 receipt intervals are serial. Final cleanup and seal verification are
recorded in the bundle verification receipt.

The candidate stays within the accepted ADR boundary: private CFB
implementation code, no public API or dependency change, no unsafe access,
and no change to typed errors, exact messages, validation/error precedence,
fallible reservation order, scratch reset, MiniFAT/FAT separation, or
collect-before-claim ownership. The [source review](../results/change-0536/source-review.md)
records those obligations. The candidate is rejected and no runtime speedup
is retained. OLE2/OOXML remains the active priority; ODF is deferred until
that goal completes, and iWork remains excluded.

See the [mechanism review](../results/change-0536/mechanism-review.md) and
[next priority](../results/change-0536/next-priority-review.md).
