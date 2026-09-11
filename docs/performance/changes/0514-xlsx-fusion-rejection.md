# 0514: reject XLSX parser/layout fusion

The proposed XLSX worksheet parser/layout fusion is rejected at its declared
pilot gate. The candidate fused the semantic worksheet parser with the
lossless snapshot `Layout` scan for ordinary borrowed MCE-free sources. It
passed the owner correctness and source-preservation checks, but cold semantic
no-ops paid a large speculative scan cost and held layout state beside the
semantic parser. The dense cold no-op incremental peak also increased by
**27.04%** and **18.64%** for the one-cell and one-percent actions.

The decision is recorded as `performance_claim: none` and
`claim_authorized: false`. The candidate is not retained in production. The
exact 14-file candidate patch remains archived for review, while production
and focused-test sources have been restored to base revision `33a21e0f0`.

## Candidate mechanism and correctness result

On a cold ordinary cell, row or column edit, the candidate fed one
namespace-aware reader to the semantic parser first and the snapshot scanner
second. It kept the scanner's layout and any first scanner error in an
ephemeral transaction-local wrapper. Semantic parsing, style validation and
action projection retained precedence; an ineffective action discarded the
layout, while an effective rewrite reused it only when the source identity
matched. MCE-owned bytes, merge-derived input, metadata edits, source identity
mismatch and cache-hit paths used existing fallback behavior.

The candidate preserved semantic values and exact source bytes in all recorded
correctness checks. The owner XLSX unit receipt reports **966/966 tests
passing**, including 17 new fusion tests. The [unit](../results/change-0514/xlsx-unit-receipt.json),
[Clippy](../results/change-0514/clippy-receipt.json),
[formatting](../results/change-0514/fmt-receipt.json),
[boundary](../results/change-0514/boundaries-receipt.json), and
[strict claim](../results/change-0514/claims-receipt.json) receipts all exit
successfully. These receipts validate the tested behavior and recorded scope;
they do not override the failed performance admission gate.

The candidate patch adds no public API or dependency edge. Its 10 production paths and four test paths form the 14-path patch recorded in
[`candidate-patch.json`](../results/change-0514/candidate-patch.json). The
[restoration receipt](../results/change-0514/restoration.json) records exact
base source-manifest recovery and confirms that production and focused-test
sources were restored.

## Pilot protocol and changed-edit rows

The declared formal protocol was 500 samples with five warmups over 12 rows
(four commit/commit-save cases × three shapes) in serial ABBA order. The
candidate was rejected during the predeclared 20-sample/two-warmup pilot, so
no formal after ABBA capture was run. The pilot remains descriptive evidence,
not a registered speedup comparison. The table preserves all 12 main pilot
rows; values are p50 milliseconds and candidate deltas are relative to the
paired control.

| Case | Shape | Control p50 | Candidate p50 | Delta |
| --- | --- | ---: | ---: | ---: |
| `xlsx_one_cell_commit` | tiny | 0.163155 | 0.154865 | -5.08% |
| `xlsx_one_percent_commit` | tiny | 0.298641 | 0.284846 | -4.62% |
| `xlsx_one_cell_commit_save` | tiny | 0.204200 | 0.197031 | -3.51% |
| `xlsx_one_percent_commit_save` | tiny | 0.368021 | 0.352336 | -4.26% |
| `xlsx_one_cell_commit` | medium | 1.746216 | 1.644951 | -5.80% |
| `xlsx_one_percent_commit` | medium | 6.971116 | 6.573569 | -5.70% |
| `xlsx_one_cell_commit_save` | medium | 2.213778 | 2.096663 | -5.29% |
| `xlsx_one_percent_commit_save` | medium | 8.815908 | 8.268761 | -6.21% |
| `xlsx_one_cell_commit` | dense-wide | 106.424767 | 105.982994 | -0.42% |
| `xlsx_one_percent_commit` | dense-wide | 216.647700 | 240.667963 | +11.09% |
| `xlsx_one_cell_commit_save` | dense-wide | 153.957515 | 147.516981 | -4.18% |
| `xlsx_one_percent_commit_save` | dense-wide | 309.577947 | 301.240949 | -2.69% |

The changed-edit pilot is mixed: the dense one-percent commit regresses by
11.09% while the other dense rows improve slightly and the tiny/medium rows
show larger pilot reductions. Those preliminary results are not promoted to
a speedup claim because the cold no-op guard fails and the formal after lane
was intentionally skipped.

## Cold and warm public guard rows

The independent public guard uses `Edit::commit` or first public `Worksheet::cell`
load only. Fresh workbook open, store warming, edit preparation, output and
readback oracles, and commit/view drop remain outside its clock. It uses 20
samples and two warmups per row in the pilot. All 21 rows (three shapes ×
seven scenarios) are retained here as p50 milliseconds; the complete sample
vectors, source/output oracles and machine-readable metadata remain in the
[control](../results/change-0514/before/guard-pilot-report.json) and
[candidate](../results/change-0514/after/guard-pilot-report.json) reports.

| Shape | Scenario | Control p50 | Candidate p50 | Delta |
| --- | --- | ---: | ---: | ---: |
| tiny | `cold-first-cell-read` | 0.035240 | 0.035810 | +1.62% |
| tiny | `cold-same-one-cell` | 0.037200 | 0.062530 | **+68.09%** |
| tiny | `cold-same-one-percent` | 0.075220 | 0.123790 | **+64.57%** |
| tiny | `warm-same-one-cell` | 0.000930 | 0.000940 | +1.08% |
| tiny | `warm-same-one-percent` | 0.001330 | 0.001310 | -1.50% |
| tiny | `warm-changed-one-cell` | 0.124031 | 0.121880 | -1.73% |
| tiny | `warm-changed-one-percent` | 0.226720 | 0.219371 | -3.24% |
| medium | `cold-first-cell-read` | 0.467141 | 0.458272 | -1.90% |
| medium | `cold-same-one-cell` | 0.483862 | 0.800593 | **+65.46%** |
| medium | `cold-same-one-percent` | 0.976114 | 1.610566 | **+65.00%** |
| medium | `warm-same-one-cell` | 0.001360 | 0.001410 | +3.68% |
| medium | `warm-same-one-percent` | 0.003990 | 0.003930 | -1.50% |
| medium | `warm-changed-one-cell` | 1.292925 | 1.269865 | -1.78% |
| medium | `warm-changed-one-percent` | 2.645280 | 2.555810 | -3.38% |
| dense-wide | `cold-first-cell-read` | 30.365800 | 29.460050 | -2.98% |
| dense-wide | `cold-same-one-cell` | 31.451314 | 51.715152 | **+64.43%** |
| dense-wide | `cold-same-one-percent` | 62.569727 | 115.530340 | **+84.64%** |
| dense-wide | `warm-same-one-cell` | 0.011690 | 0.011480 | -1.80% |
| dense-wide | `warm-same-one-percent` | 0.963694 | 0.820253 | -14.88% |
| dense-wide | `warm-changed-one-cell` | 79.378938 | 77.395828 | -2.50% |
| dense-wide | `warm-changed-one-percent` | 163.265672 | 157.927507 | -3.27% |

The six cold same-value rows regress **+64.43% to +84.64%**. This is the
decisive rejection signal: the candidate performs speculative semantic/layout
work for an action that later projects to an exact no-op. The standalone
public guard remains reusable for a later design; rejecting this production
fusion does not reject that guard harness.

## Allocation overlap and absolute peak interpretation

Both roles have two separate allocator captures, each with 10 measured
samples and one warmup per row. They reproduce the dense cold-no-op overlap. The incremental value is derived as
`region_peak_live_bytes - live_bytes_before`; absolute values are retained
separately:

| Dense cold scenario | Control entry live | Control region peak | Control incremental | Candidate entry live | Candidate region peak | Candidate incremental | Incremental delta |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `cold-same-one-cell` | 21,672,890 | 48,951,183 | 27,278,293 | 21,672,889 | 56,328,006 | 34,655,117 | **+27.04%** |
| `cold-same-one-percent` | 34,523,457 | 74,098,991 | 39,575,534 | 34,523,456 | 81,475,814 | 46,952,358 | **+18.64%** |

The entry and region values are absolute process live-byte observations. The
incremental column isolates the extra high-water demand relative to the
operation's entry state, but remains allocator callback evidence rather than
physical RSS. The candidate holds semantic parser state and temporary layout
state together; dropping the layout later does not remove this overlap. The
allocator reports are operation-scoped and do not support a process-wide heap
claim.

## Diagnostic candidate profile

The candidate exact profile is retained for compiler-work attribution even
though it does not rescue the rejected design. Collected save-helper Callgrind totals
are:

| Profile | Instruction references |
| --- | ---: |
| Control | 14,458,336,036 |
| Candidate | 13,547,679,875 |
| Candidate minus control | **-6.30%** |

The candidate helper annotation contains three direct commit calls costing
**11,001,102,600** references and three direct writer calls costing
**2,546,576,642**. The remaining 633 references are annotated as 171 helper-
exclusive plus 462 `memcpy`. Both profiles emit a `brk segment overflow`
warning but exit successfully. The profile is therefore instruction-level
diagnostic evidence only. Descendant call metadata includes calls made while
collection was disabled; a descendant count such as `9` must not be read as
nine measured stores. Collected instruction costs remain scoped to the
selected helper region.

The full after native lane was not run because the cold no-op gate rejected
the candidate. The exact profile does not establish practical latency,
throughput, memory or retention benefit.

## Retained state and completed control repeats

The production source and candidate tests are restored exactly to base. Both
control repeats have completed: 500 samples and five warmups over the 12 main
rows, and 100 samples and three warmups over the 21 guard rows. Together they
retain 16,200 formal control durations, separate from the 1,320 matched pilot
durations and 1,320 instrumented allocator samples. No full candidate native
comparison was admitted.

Main control p50 drift is −1.67% to +2.29%; guard control p50 drift is −4.63%
to +3.88%. The guard has one absolute tail-drift flag: tiny cold-same-one-cell
p95 changes −11.07% against its 10% threshold. Main control whole-child RSS is
140,232 -> 137,348 KiB, while guard RSS is 169,312 -> 180,284 KiB (+6.48%,
above the 5% review threshold). These control-context observations are retained;
they are not candidate memory deltas and do not replace the repeated operation
allocation evidence. Full vectors and drift flags are in
[the replay summary](../results/change-0514/summary.json).

The existing owner
unit, Clippy, formatting, boundaries and claims receipts are candidate-era
checks; they document correctness and scope but do not authorize a
performance claim. The exact [decision](../results/change-0514/decision.json),
[pilot plan](../results/change-0514/plan.json), [main control pilot](../results/change-0514/before/pilot-report.json),
[main candidate pilot](../results/change-0514/after/pilot-report.json),
[public control guard](../results/change-0514/before/guard-pilot-report.json),
[public candidate guard](../results/change-0514/after/guard-pilot-report.json),
[candidate allocator captures](../results/change-0514/after/allocator-r1-report.json),
[guard allocator captures](../results/change-0514/after/guard-allocator-r1-report.json),
[candidate profile receipt](../results/change-0514/after/profile-receipt.json),
[restoration receipt](../results/change-0514/restoration.json),
[design review](../results/change-0514/design-review.md), and
[implementation review](../results/change-0514/implementation-review.md) retain
the machine-readable evidence and rejected-path rationale.

## Priority after rejection

The full OLE2/OOXML optimization goal remains active. Further ODF work stays
deferred until that goal is complete, and iWork remains excluded. The next
XLSX/OLE2/OOXML candidate must begin from a fresh scoped hotspot and preserve
the cold-no-op and incremental-overlap gates established here.
