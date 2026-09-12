# 0522 runtime evidence review

`scope: read-only review of the terminal 0522 runtime evidence`

`disposition: reject the production scanner candidate under the frozen plan`

`performance_claim: none`

This review treats the native total gate as the admission decision. The
allocator and Callgrind lanes are supporting mechanism evidence; neither can
override a mixed primary total result. The numeric and profile capture lanes
are terminal. All twelve candidate quality gates and the final baseline-bound codec gate
have passed. The independent review ran no build, test, capture, or source
edit; root updated the final disposition/custody notes after that review,
without changing its numerical findings or rejection verdict.

## Hash and custody binding

| artifact | SHA-256 |
| --- | --- |
| [`plan.json`](plan.json) | `d7560f2aa733fdae98621d0fb2a6d85d926acd23751b87a3adb09f3e8a4703d0` |
| [`comparison.json`](comparison.json) | `da7bcdac0a6c8b7736d3a415d2966ba562db3afc83671431662264bd3f7afb27` |
| [`profile-analysis.json`](profile-analysis.json) | `219bb045247614d6c8e1a12b050acf7a1927dc02c9df1e5611064324e8d8c6d7` |
| [`flag-review.json`](flag-review.json) | `72afe7d4f91594a45d86b023036cbb9283003e5c1263e6e32e9b40504b18acee` |
| [`candidate/source.patch`](candidate/source.patch) | `c4c79c56f188820c2bc7d800b9ce17994f9c4094933fcd04b9b09b911caba4f0` |
| [`candidate/source-manifest.json`](candidate/source-manifest.json) | `d9bac2799c7b75a8bf8829e0e0956ae94927fdda985def2aab58f8496bd773de` |
| measured candidate scanner ([patch](cell-reference-candidate.patch)) | `adc39e78ab5d3fed25e6d9d93259c0b4cdd5df9594f4751b7b661de187c32cc2` |

The comparison and profile reports both carry the frozen plan hash. The
comparison has 18 native rows and 820 samples per stage (1,640 total), four
allocator rows and 20 samples per stage. The numerical analyzer checks 22 non-overlapping capture receipts per stage
(18 native plus four allocator), with 8,584 source-manifest entries. The final
verifier checks all 69 build, capture and quality intervals. The baseline and candidate binary
identities are distinct as required, and every timing row (18/18) and
allocator row (4/4) has `identity_equal: true`. The profile report has four
baseline/candidate profile pairs and all of its validation predicates pass.

These are analyzer and custody results, not a claim that the candidate passed
the performance gate. The candidate was reverted as recorded in [disposition.json](disposition.json).
The final checkout matches the baseline manifest; all 19 final codec tests
passed. [verification.json](verification.json) records successful source/report/
annotation replay and cleanup. The measured scanner hash above belongs only
to the rejected candidate.

## Frozen primary admission

The plan requires useful, repeatable improvement in primary **total** time,
with semantic, output, and resource oracles unchanged. The total includes
open, planning, staged sets plus commit, and publication (including the
returned snapshot's publication drop). Reopen and post-operation oracles are
recorded separately and are not part of that total.

| Primary shape / repeat | Baseline p50 ms | Candidate p50 ms | p50 change | Mean change | Within-child p50 ratio interval |
| --- | ---: | ---: | ---: | ---: | --- |
| dense-sparse / 1 | 50.311 | 49.276 | −2.056% | −2.098% | 0.975126–0.983466 |
| medium / 1 | 25.476 | 25.804 | **+1.286%** | **+1.255%** | 1.011893–1.013817 |
| dense-sparse / 2 | 49.774 | 49.119 | −1.316% | −1.161% | 0.985519–0.988639 |
| medium / 2 | 25.609 | 25.275 | −1.303% | −1.397% | 0.984330–0.989662 |

Dense-sparse improves in both repeats, by only 2.06% and 1.32% at p50.
Medium changes sign: it regresses 1.29% in repeat 1 and improves 1.30% in
repeat 2. Medium repeat 1 also regresses p95, p99, and mean by 1.006%,
1.339%, and 1.255%. The bootstrap intervals are the plan's descriptive
within-child resampling of candidate median over baseline median; they are not
cross-build, cross-host, or causal confidence intervals. The sign change and
small magnitude fail the frozen useful-repeatable-primary-total requirement.

All 18 native identities remain equal, so this is a performance disposition,
not an output or semantic mismatch. A passing identity oracle does not turn a
non-repeatable total into an accepted optimization.

## Retained flags

The complete flag set is retained in [`comparison.json`](comparison.json) and
reviewed in [`flag-review.json`](flag-review.json); no adverse row is removed
or relabeled as noise.

| retained diagnostic | count | distribution |
| --- | ---: | --- |
| candidate adverse metrics above 5% | 43 | 20 open, 17 publication, 6 excluded reopen |
| same-build drift metrics above 5% | 67 | 36 open, 21 excluded reopen, 10 publication; 27 baseline-stage and 40 candidate-stage |

Representative candidate adverse rows include managed dense-sparse guard 1
repeat 1 open p95 **+31.208%** and p99 **+29.563%**, vendor-extension guard 2
repeat 2 open p95 **+27.884%** and p99 **+26.651%**, and medium primary repeat
1 publication p50 **+11.731%** and p95 **+12.845%**. The six reopen rows are
outside the primary total but remain visible because the plan requires every
adverse phase metric to be retained.

The same-build set is a separate repeatability diagnostic and cannot be
attributed to the candidate. Its largest row is baseline-stage medium primary
open p99 drift of **+46.629%** between repeats 1 and 2, with p95 drift of
**+39.179%**. Candidate-stage and baseline-stage variations remain separately
listed; neither is silently used to manufacture or erase a candidate effect.

## Supporting allocator and profile evidence

The canonical allocator region covers staged sets plus `MultiSourceEdit::commit`
only. Publication, reopen, and oracles are outside this lane. Its repeated
vectors are real allocator observations:

| shape | allocation calls | allocated bytes | reallocations | incremental region peak |
| --- | ---: | ---: | ---: | ---: |
| medium | 118,744 → 100,312 (**−15.522%**) | 20,344,427 → 19,724,459 (**−3.047%**) | 10,798 → 10,798 | 2,984,983 → 2,984,983 |
| dense-sparse | 225,771 → 190,187 (**−15.761%**) | 27,331,029 → 26,122,091 (**−4.423%**) | 19,564 → 19,564 | 7,335,225 → 7,335,225 |

The same values repeat for both allocator repeats. The call reduction and
small byte reduction do not establish a native total improvement, and the
incremental region peak is unchanged. These values are not whole-child RSS,
document peak, or per-cell event counts.

The selected Callgrind owner is
`litchi_xlsx::cell_values::source::MultiSourceEdit::commit`. Four profile
pairs report commit Ir changes of:

| repeat / shape | commit Ir change |
| --- | ---: |
| 1 / dense-sparse | −4.172% |
| 1 / medium | −3.757% |
| 2 / dense-sparse | −4.168% |
| 2 / medium | −4.102% |

The aggregate is 1,148,663,511 → 1,101,713,273 Ir (**−4.087%**). The
`Scanner::start_cell → wire::cell_tag` positive edge is present in every
baseline profile and absent in every candidate profile. The analyzer's raw
inner call metadata is intentionally not used as timed event or allocation
counts; collection-off setup and readback work can be represented there. The
validator inclusive row moves 273,489,650 → 274,068,257 Ir (+0.212%) and is
also diagnostic only. `profile-analysis.json` defines the boundary explicitly:
Callgrind Ir and inclusive rows are not native latency, hardware cycles,
allocation counts, cold-cache, range, scaling, or native Office-producer
measurements.

The dominant direct callee remains
`Snapshot::from_rewritten_source` (702,367,967 baseline Ir, 61.147% of the
aggregate) and the worksheet rewrite is 443,626,538 Ir (38.621%). The
candidate's rewrite reduction is not enough to admit the end-to-end change;
the reconstruction callee rises 0.640%. This is attribution for the next
investigation, not permission to skip validation or semantic parsing.

## Source and contract verdict

The source-level proof remains sound. The following line numbers refer to the
rejected candidate retained in [the patch](cell-reference-candidate.patch),
not the restored production scanner. The candidate helper at
`candidate scan.rs:124`
performs one checked attribute iteration, decodes the unqualified `r`, and
returns a raw-name proof. Empty cells at
`candidate scan.rs:689`
and start cells at
`candidate scan.rs:992`
call `cell_tag` only when that proof is false. `cell_address` at
`candidate scan.rs:1017`
keeps row lookup, A1 parsing, row mismatch, inferred-column, cursor update,
and typed address checks in order. The unchanged fallback at
[`wire.rs#L105`](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/codec/wire.rs#L105)
retains prefix, additional-attribute, UTF-8, normalized-value, and source-order
handling.

The source owner still checks execution, validates the rewritten worksheet,
fully reparses it, clones the snapshot state, stores owned bytes, and checks
execution again at
[`snapshot.rs#L686`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L686).
The semantic cell attribute pass and cell finalization at
[`codec.rs#L172`](../../../../crates/litchi-xlsx/src/raw/worksheet/codec.rs#L172)
and [`codec.rs#L989`](../../../../crates/litchi-xlsx/src/raw/worksheet/codec.rs#L989)
are unchanged. Validation, readback, limits, source/resource fences, and
error precedence therefore have no source-level blocker. This correctness
verdict is independent of, and does not rescue, the failed native admission.

## Disposition

The production `scan.rs` candidate was rejected and reverted under the frozen plan.
Keep the baseline-compatible codec tests, opt-in `noncompact` guard, and all
campaign evidence. Do not relax the useful-repeatable-total gate, infer a
speedup from allocator calls or Ir, or revive the rejected 0514/0516 fusion
designs. The next work belongs to a fresh measured OLE2/OOXML opportunity.
