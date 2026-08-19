# Change 0197: batched per-window row reparse in litchi-ods — WITHHELD, reverted

Date: 2026-08-18

## Verdict

**Not banked. The implementation was reverted; only its new regression tests
remain.** The mechanism (batching the per-row synthetic-document reparse into
one reparse per changed window) is real and reproducible on the phase it
targets, but the fullest measured scope — the one-percent edit-save lifecycle —
moved adversely for the candidate binary in both paired directions in both
independent runs. The house standard requires withheld statistics to show no
regression pattern (see 0196), so no claim is made and the code is reverted.

## What was implemented (and then reverted)

`crates/litchi-ods/src/worksheet/package.rs` `validate_rewritable_row` built
one synthetic XML document per changed row (ancestor opening tags + row slice +
closing tags) and reparsed each with a fresh `quick_xml::NsReader`, re-resolving
the ~30 root xmlns declarations per row. The candidate replaced the per-row
calls inside `changed_row_edits` with a single
`validate_rewritable_rows(xml, spans, &rows[prefix..old_end])`:

- One synthetic document per contiguous changed window: shared ancestor chain
  (all rows are direct children of the same table) emitted once, then every row
  slice concatenated, then the closing tags.
- Per-row parser state (`row_depth`, `in_paragraph`) provably returns to
  `0`/`false` at each row boundary because every row slice is a well-formed
  element verified by the earlier full-document `scan()`.
- Error ordering preserved exactly: the descendant-span check for row k + 1
  runs when the reader position reaches the end byte of row k, so any refusal
  in row k (span check or reparse) fires before row k + 1 is examined, with
  identical error texts. Row completion is detected by byte position
  (`position(&reader) == row_ends[k]`), which is robust even for pathological
  ancestor chains containing a `table-row` element.
- The per-row descendant walk was extracted unchanged as
  `validate_row_descendants`.

Reachable-path behavior was byte-for-byte equivalent; the only behavioral
deltas were on defense-in-depth paths that are unreachable from `scan()`
-verified spans (reparse error wording for malformed row slices, and span-index
lookup order inside the window loop).

## Measurement

Two frozen release binaries differ only in `worksheet/package.rs`; both carry
changes 0193-0196 as baseline and the identical 341-case selector matrix.
Control SHA-256 `db708afa17eddc7dab9911429c51d6d0cd676550f8c33f4458893f2ea1201cff`
(bit-identical to the 0196 candidate, confirming build reproducibility),
candidate SHA-256
`fe3da69fae3a125183f2a9416cb49cf78314bf41c2a9a7743db7bd2af7de0ef3`.
Standard protocol: fresh CPU-2-pinned processes, order A1/B1/B2/A2, 30 warmups,
500 samples per leg, drift ceilings 5%/5%/10%/15% (p50/mean/p95/p99), all
embedded verification flags true in every leg. The single allowed rerun was
spent on the one-percent workload.

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

44 rows validated per commit across two 22-row windows. All sixteen statistics
withheld: total/commit/publication directions straddle zero within 3.4%. The
stage phase — which executes no code touched by this change — measured the
candidate 4.20%-7.41% *slower* in both paired directions on p50/mean/p95 with
healthy drifts, an adverse pattern attributable only to code-layout shift in
the candidate binary.

### One-edit guardrail (`ods_source_backed_one_edit_save`)

Single-row window: the batched path degenerates to the old behavior. All eight
statistics withheld as neutral (directions straddle zero; commit p99 straddles
+4.15%/-9.08% with candidate drift +15.02% just over the 15% ceiling).

### One-percent guardrail (`ods_source_backed_one_percent_edit_save`), two runs

Commit phase accepted identically in both runs:

| Statistic | Run 1 A1→B1 / A2→B2 | Run 2 A1→B1 / A2→B2 | Decision |
|---|---:|---:|---|
| commit p50 | +1.54% / +1.97% | +1.34% / +2.10% | accepted (both runs) |
| commit mean | +1.43% / +1.80% | +1.32% / +1.77% | accepted (both runs) |
| commit p95 | +0.10% / +1.17% | +1.25% / +0.53% | accepted (both runs) |
| commit p99 | +0.69% / -0.77% | +2.27% / -0.97% | withheld (straddle) |

Lifecycle statistics, however, measured the candidate slower in **both paired
directions on all four statistics in both runs** (run 1 p50 -1.27%/-1.57%, run
2 p50 -0.65%/-0.12%; p99 run 1 -1.06%/-3.41%). Since the commit phase is
faster, the adverse component sits in the open/stage/publication phases, which
execute identical instruction streams under this change — i.e. a systematic
per-binary code-layout effect, not a semantic one. It is nonetheless a
consistent measured regression pattern on the headline lifecycle scope, which
is exactly what the paired-direction protocol exists to surface, so the change
is withheld rather than banked on the phase-scoped commit win.

## What remains in the tree

Two tests in `crates/litchi-ods/tests/source_cell_transactions.rs` pin the
pre-existing multi-row window behavior of the per-row validator and pass
identically on the reverted code:

- `source_cell_commit_validates_multi_row_windows_in_order` — a changed window
  spanning an unchanged row and an empty row commits exactly.
- `source_cell_commit_reports_the_first_failing_row_in_a_window` — refusals
  name the first failing row with the exact established error text for
  descendant-scan failures (two paragraphs per cell) and reparse failures
  (loose text, CData, unmodeled row attribute).

litchi-ods passes 342/342 with these tests on the reverted code; scoped
clippy, rustdoc, fmt, and crate-boundary gates pass.

## Follow-up note for the seam

The commit-phase win (~1.3-2.1% on a 21-cell commit) confirms the profile
attribution: per-row `NsReader` setup plus root-namespace re-resolution is a
real cost. A future attempt at this seam should remove the synthetic reparse
entirely (e.g. validate against the cached `ContentLayout` spans with
span-stored namespace bindings, reparsing only attribute lists) rather than
amortize it; the amortization win is too small to clear binary-layout noise on
the lifecycle scope.

Artifacts:

- repeated-edit: [summary](../results/ods-repeated-edit-0197-summary.json),
  [manifest](../results/ods-repeated-edit-0197-manifest.json)
- one-edit guardrail: [summary](../results/ods-one-edit-save-0197-summary.json),
  [manifest](../results/ods-one-edit-save-0197-manifest.json)
- one-percent guardrail: [summary](../results/ods-one-percent-edit-save-0197-summary.json),
  [manifest](../results/ods-one-percent-edit-save-0197-manifest.json)
- one-percent rerun: [summary](../results/ods-one-percent-edit-save-0197r-summary.json),
  [manifest](../results/ods-one-percent-edit-save-0197r-manifest.json)
- compressed A1/B1/B2/A2 raw reports listed in each manifest
