# 0362upa — Budget composition returned for correction

Parent: [0362up](0362up-budget-assembly.md). Status: returned; diagnostic
composition is not accepted into the source integration or trunk.

Review of exact committed shared proposal `158b7029f6` found correctness,
API and performance defects that must be repaired under the 0351 contract.
Compiler progress and successful unrelated family reviews do not waive them.

- Cell transfers must preserve pending typed styles and remove moved source
  content. Retained style maps must not recreate a removed source cell.
- Native no-op identity must include rich-text properties. Same visible text
  with changed formatting must remain an edit; genuinely canceled or identical
  edits must release their retention charge. Existing merges, cached values,
  absent cells and bounded/unbounded joins require the same effective-state
  accounting, including failure rollback.
- Metadata admission must charge actual source, target and native patch
  retention before capture. A `2 * source` or `3 * source` substitute can
  reject shrinking edits and miss expansion. The configured public entry
  point is `edit_metadata_with_options(EditOptions)` followed by `commit()`.
- Prepared metadata encoding must reuse admitted lengths/layouts rather than
  decode/scan/preflight the same source again. Default edits must preserve
  their prior path. Planning must release removed relationship IDs in the
  same order as final publication so tight limits remain exact.
- The combined sheet-property/auto-filter facade route must propagate limits
  and include its native patch charge through its intermediate owner.

The separate feature/chart proposal `5ae87e6e08` already projects native
source/target retention before patch capture and passed its scoped review.
It is not replaced by a broader claim based on the older shared leaf. DOCX
source admission and XLSB digest/census fixes also retain their scoped clear
reviews. Final combined compilation, meaningful boundary/fidelity tests and
performance qualification remain required for the complete composition.

## Second cell review and combined runtime

Pinned shared proposal `1929682421` repairs common metadata preparation but
still needs cell corrections: same-visible plain assignment must erase rich
runs; rich transfers must carry source cell styles; transfer failure must
restore every staged map; existing merges and equivalent typed styles must
normalize before finite admission; shrinking style payloads must release the
old charge; independent hyperlink ranges sharing a start must accumulate.
Inline-rich range transfers must avoid repeated whole-part scans per address.
These are required repairs, not accepted fidelity or performance exceptions.

The combined metadata library suite passes 482 tests on `af09b4fcf9`.
XLS source-backed tests reach 72 passes and seven fixture failures: the new
freshness helper assumed a numeric cell in `Simple.xls`, which has none.
The helper must create or select a valid numeric source before exercising
cold, cached, inflight and handoff fences; no freshness assertion is waived.
The complete XLS run and all subsequent qualification remain required.

The corrected generated XLS fixture passes the complete unit/integration
run on `374fb0e5d9`: 1,863 tests across 85 suites, one ignored. The XLSX
library suite on that source passes 1,842 tests and fails five. Picture-only
transfers lose their target; shared-string transfers lose their exact
index/dependency and projected identity; two staged/joined capture-budget
boundaries fail. The transfer and ledger owners must repair behavior while
preserving these existing assertions. The source remains diagnostic.
[xls-xlsx-runtime-v2.json](results/0362up/xls-xlsx-runtime-v2.json) retains
both commands, summaries and external log hashes.

## Complete facade runtime return

On `55051a8e4f`, the full all-feature `litchi-ops` library suite passes
993 tests and fails nine (211.58 seconds including compilation; 72.90
seconds in tests). The complete batch is returned by scope:

- DOCX: recorded replay refusal classification, retained-hyperlink access,
  and exact signed paragraph no-op behavior.
- ODP/ODS: malformed-input tests no longer find their expected typed cause
  through the new crate-owned error chain, despite `CorruptFile` classification.
- DOC: a signed paragraph no-op reports corrupt input.
- XLS: signed edit entry and cancellation disagree with the explicit policy.
- PPTX: a signed cancelling leaf edit still reports document protection.

The repair must distinguish stale assertions from lost capabilities using
the confirmed contract. Exact signed no-ops/cancellation, changed-edit
refusal, typed causes and content-free refusal reasons remain required. The
XLSB runtime group was not reached after the facade failure.
[facade-xlsb-runtime-v12.json](results/0362up/facade-xlsb-runtime-v12.json)
retains the source, command, complete log and debug executable hashes.

The separate cell proposal `06afe43fa0` is also returned. It conflates
explicit empty creation with clear-if-present, loses exact shared/inline
string representation, scans rich XML repeatedly, mutates staged removal
before fallible typed admission, clones complete edit payloads during joins,
and flattens normalization errors into strings. A current-source repair leaf
`b2a1e6e735` composes the proposal with the reviewed rich-transfer and metadata
work solely to repair this full batch without temporary compiler overlays.
It is not accepted into this diagnostic branch or the main integration.

## XLSB runtime return

On `d4d1820dbe`, the full all-feature XLSB library passes 769 tests and
fails three (54.10 seconds including compilation, 8.65 seconds in tests).
The retained-metadata test compares serialized producer defaults against an
empty metadata value; two OPC eager/deferred projection tests disagree on
content-types source presence. These are returned for diagnosis and repair
without weakening exact retention limits, source fidelity or projection
validation. The runner stopped before the queued PPTX/ODF group.
[xlsb-pptx-odf-runtime-v13.json](results/0362up/xlsb-pptx-odf-runtime-v13.json)
retains source, command, complete log and completed executable identities.

The first XLSB fixture correction passes the metadata boundary. On
`dfd7edc65f`, 770 library tests pass and the two OPC projection tests still
fail: retained content-types source bytes differ after reopened replay
(1,113/1,110 versus 1,024 bytes). The remaining repair must distinguish
original physical provenance from effective publication while preserving
exact output and untouched payload fidelity. DOCX was not reached after
the runner stopped. [xlsb-docx-runtime-v15.json](results/0362up/xlsb-docx-runtime-v15.json)
retains the complete failure receipt.

## DOCX shared-scanner fidelity return

The full all-feature DOCX library on `0fcc45dae0` passes 1,516 tests and
fails one: `cell_text_replacement_refuses_content_it_would_drop` now admits
a bookmarked cell paragraph. The paragraph scan change must be audited
across all callers: text-slot replacement preserves bookmarks, whereas a
cell-content replacement may discard them. Preserve the existing fidelity
refusal or prove marker preservation; the failing assertion is not waived.
[docx-runtime-v16.json](results/0362up/docx-runtime-v16.json) retains the
complete source, command, test summary and external log identity.
