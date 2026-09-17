# 0675 — integrate the authorized third wave and close the residue queue

Change [0652](0652-owner-decisions-for-the-third-wave.md) authorized the ten
decision rows of [0651](0651-queue-refresh-after-the-second-wave.md). This
record combines the implementations already on the branch with the resumed
`perf/NNNN-*` worktrees and the bounded follow-ups discovered during review.
It makes no aggregate performance claim. Each implementation's packet owns
its measurements, floors, corpus scope and limitations.

## Queue disposition

| 0651 row | Result and evidence |
| ---: | --- |
| 1 | Scoped MCE namespace emission and consumer migration landed in [0653](0653-mce-namespace-emission-rewrite.md); [0666](0666-mce-rewrite-residue.md) handles the remaining worksheet admission predicates. |
| 2 | Original XML publication accepts valid noncompact source bytes in [0654](0654-opc-original-bytes-audit-loosened.md); [0665](0665-opc-eager-writer-publication-audit.md) carries the eager provenance path, and [0677](0677-xml-publication-bom-offsets.md) repairs marked-XML audit offsets. |
| 3 | [0661](0661-opc-lazy-part-decode.md) migrates OPC part decode and audits newly fallible consumers. |
| 4 | [0655](0655-pptx-memoized-revision-proof.md) implements the memo and durable patch format bump. |
| 5 | [0656](0656-pptx-cross-copy-candidate-budget.md) retains the validated candidate under an explicit budget. |
| 6 | [0662](0662-parallel-changed-member-deflate.md) implements changed-member parallel deflate; [0676](0676-execution-budget-composition.md) completes shared execution budgets across the existing scheduled read paths. |
| 7 | [0657](0657-xlsx-value-editor-d4-admission.md) widens the value editor, with [0666](0666-mce-rewrite-residue.md) MCE and [0667](0667-xlsx-value-editor-shared-strings.md) shared-string dependency admission follow-ups. |
| 8 | [0658](0658-xlsx-selected-cell-ineligibility-gate.md) implements the accepted ineligibility gate. |
| 9 | [0659](0659-cfb-single-scan-identity-entry-point.md) adds the DOC identity entry point. |
| 10 | [0660](0660-docx-compaction-policy.md) makes DOCX preservation the default; [0663](0663-cfb-sector-layout-policy.md) adds CFB sector policy and reuse. |
| 11 | [0668](0668-xls-query-residues.md) retains the ADR 0005 prerequisite: cross-query indexes and chain hints need a bounded weighted evictable cache. No unbudgeted cache is introduced. |
| 12 | [0668](0668-xls-query-residues.md) avoids unwanted packed-cell materialization and temporary SST pointer scratch. The full locator scan remains; framing fusion retains its coverage-proof prerequisite. |
| 13 | [0669](0669-xlsb-edit-residues.md) removes duplicate candidate parsing and fixes derived-state and relationship guards. 0674 supplies the workbook-structure selector. |
| 14 | [0670](0670-docx-parser-residues.md) finishes borrowed-event and paragraph-buffer work and records the paragraph-index admission design. Guarded source readers retain their bounded buffers. |
| 15 | [0672](0672-xlsx-stored-cell-allocation.md) reserves stored-cell results exactly. The selected scanner still retains its records until a refusal-before-result design permits streaming. |
| 16 | [0670](0670-docx-parser-residues.md), [0671](0671-doc-admission-residues.md) and [0677](0677-xml-publication-bom-offsets.md) resolve BOM offsets, Strict page units and facade options. Local specs retain strict FBKF uniqueness and body-final section placement; observed PAPX padding requires explicit compatibility opt-in. |
| 17 | [0664](0664-perf-harness-marker-bearing-corpora-and-save-allocations.md) supplies producer corpora, ordinary-save allocation regions and DOCX text sinks; [0674](0674-performance-gate-hygiene.md) adds PPTX transaction phase timing. |
| 18 | [0673](0673-zip-locator-scratch.md) defers the ZIP locator scratch allocation until its stack probe misses. |
| 19–20 | [0674](0674-performance-gate-hygiene.md) updates feature/allocator gates, lockfile and evidence hygiene and documents release-binary reproducibility limits. |

## Integration-specific corrections

The DOCX tail-append settings test still expected the old per-element namespace
amplification. The scoped MCE codec now fits that fixture's source-size ceiling.
The updated regression checks successful publication at the exact ceiling,
unchanged source settings bytes, and typed refusal before output one byte below
it. All 40 tail-append integration tests pass after this correction.

Combining lazy OPC ingress with the revision memo required fallible payload
access in both memo projection and fingerprinting. The merged consumer check
and all 79 PPTX opened-presentation tests pass with those accessors. Publication
conflicts preserve allocation provenance for original XML and decode-error
propagation in the XLSB relationship guards.

The complete merged suite exposed stale payload bytes beyond a shortened root
mini-stream's logical end. Clearing that padding restores byte-identical XLS
durable inversion. The original XLS regression passes unchanged, and a focused
CFB grow/shrink regression covers the physical-sector boundary directly.

The 0657 rollup also contained an obsolete pre-publication statement. Its
packet and shared audit now describe the final publication results consistently.

## Measured costs retained

The generic CFB Reuse/Rewrite control in 0663 records a roughly 40% p50
regression on `picture.doc` once composed-plan validation is included. It is
retained as the cost of the selected preservation policy, not a latency
speedup. The equal-length object overlay is a separate path; this control does
not establish its before/after performance. A separate `FloatingPictures.doc`
opened-DOC before/after observation is about +2.3% at p50 without a paired
floor, so it also supports no latency claim. Likewise, 0672's exact stored-cell
reservation saves allocations but records a small warm timing regression.
These costs remain visible in their packets rather than being averaged into
an aggregate optimization claim.

## Validation and cleanup

The final fourteen-crate test run passes 10,801 tests with 74 ignored and no
failures. The facade passes 382 tests, the harness 545, the isolated allocator
5, and the DOCX/ODT feature pair 104. All sixteen final gates pass, including
formatting, compilation, Clippy and rustdoc with warnings denied, strict and
structural claims, the 50 claim-checker tests, coverage and crate boundaries.
The boundary checker retains its eleven existing migration-debt entries.
The packet retains commands, exit statuses, integration logs, the merge ledger
and completed-worktree cleanup ledger. Fourteen task worktrees and four
baseline checkouts were removed; unrelated worktrees remain. Implementing commits are retained
in branch history; unrelated worktrees and the current checkout are preserved.

This wave does not claim that the remaining design prerequisites are solved:
weighted clean-value cache admission, framing coverage ordering, scanner
refusal-before-result streaming, and a priced paragraph-index memo remain
explicit follow-ups. Physical cold-cache, cross-platform, RSS and general
concurrency scaling results are not inferred from these changes.
