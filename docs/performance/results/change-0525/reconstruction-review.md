# 0525 XLSX reconstruction and semantic-store reuse review

`scope: bounded, read-only source and design audit`

`base revision: f08daf3976714dffebe37e40bd895d87266aaf8d`

`performance_claim: none`

`production_change: none`

This review covers the value-only source-backed XLSX path and the proposed
elimination of unnecessary semantic parsing after a worksheet rewrite. It does
not build, test, capture, or edit production code. OLE2 and OOXML remain the
active optimization priority; ODF is deferred until that goal is complete.

## Decision

A blanket replay of staged `Content` into the old `Store` is not a safe
replacement for `raw::worksheet::parse`. It would prove what the caller asked
the writer to do, but would not prove that the emitted XML decodes to that
state. It also misses implicit row and cell coordinates, shared-formula
provenance, declared-dimension changes, and the parser's resource and error
boundary.

A narrower fast path is supportable. The writer may produce the normal full
worksheet bytes and a private, source-bound semantic-readback skeleton. The
skeleton retains the actual emitted worksheet context and only the cell owners
that need semantic parsing. The existing raw worksheet parser parses that
skeleton; a private Store merge reuses immutable semantics for omitted owners.
Any uncertain shape, proof mismatch, reduced-parse error, allocation failure,
or merge invariant failure must discard the candidate and invoke the existing
full parse on the actual output bytes. The fallback is part of the design, not
an exceptional error path.

This is an output readback proof with bounded omission. It does not feed writer
events into the parser and does not combine the source parser with the edit
scanner. The rejected 0514/0516 fusion designs and the rejected 0522 scanner
candidate remain out of scope.

## Evidence and owner map

The retained 0522 profile puts
`Snapshot::from_rewritten_source` at 702,367,967 aggregate Callgrind
instructions, or 61.147% of the selected commit, and worksheet rewrite at
443,626,538, or 38.621%. These are attribution values, not elapsed-time or
allocation evidence. The existing source owner at
[`snapshot.rs#L686`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L686)
checks execution, validates the complete output, reparses it with
`raw::worksheet::parse`, installs the resulting Store, and checks execution
again.

The candidate can remove RawCell collection and cell materialization for
omitted owners, together with the raw-parser scratch used for those owners.
It deliberately retains the combined Store rebuild, cell and row indexes, and
extent calculation. Shared-formula resolution is avoided only for the
worksheet-level cases that use full-parser fallback; the candidate makes no
general shared-formula claim. Unchanged semantic entries may still be cloned,
including numeric and formula lexical storage, so net allocation and latency
must be measured against the existing parser's allocations. The candidate
does not remove the edit-layout scan or the complete value-only XML
validation. `raw::worksheet::parse` also owns x14ac capture,
MCE processing, UTF-8 decoding, semantic limits, and shared-string callback
selection at
[`raw/worksheet/mod.rs#L28`](../../../../crates/litchi-xlsx/src/raw/worksheet/mod.rs#L28)
and
[`raw/worksheet/codec.rs#L265`](../../../../crates/litchi-xlsx/src/raw/worksheet/codec.rs#L265).

The current value-only ingress already establishes a useful closed world. It
runs `validation::worksheet_xml`, then parses with a `None` shared-string
callback and validates styles and scalar-cell restrictions in
[`snapshot.rs#L634`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L634).
The validator rejects foreign or mixed dialect elements, MCE markup, merges,
rich inline runs, unknown worksheet elements, and unsupported attributes. The
raw parse rejects shared-string cells when no table callback is supplied. A
future broadening of this grammar must disable this fast path until its proof
is updated.

The static corpus translation in
[`closure-coverage.json`](closure-coverage.json) rules out retaining only
whole changed rows as the useful mechanism:

| shape | stored cells | updates | touched rows | cells in touched rows | cells omitted by row-only reuse | cells omitted by cell omission |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| medium | 9,216 | 93 | 93 | 4,464 | 4,752 (51.56%) | 9,123 (98.99%) |
| dense-sparse | 17,792 | 178 | 142 | 16,769 | 1,023 (5.75%) | 17,614 (99.00%) |

The counts are source-derived closure evidence, not a performance claim. The
candidate therefore needs cell omission inside changed rows and whole-row
omission for rows with no changed owners.

## Why pure Store replay is unsafe

`write_cell` does more than copy the requested semantic object. It removes and
regenerates `r`, changes or removes `t`, serializes `<f>`, `<v>`, or inline
text, escapes XML, and drops formula caches. This behavior is visible at
[`sheet_data.rs#L273`](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs#L273).
For example, a direct `Cell::Value` assignment would accept the staged value
even if a future writer regression emitted a wrong type or malformed escaped
text. Parsing the emitted bytes is the independent check required by ADR 0003
and the existing source owner.

The raw parser infers an omitted row number from the previous row and an
omitted cell column from the previous cell in that row. A replacement-only row
has unchanged membership, so omission is safe only when each retained changed
cell has the explicit address that the writer promises to emit. Removing or
inserting a cell can change every later implicit follower; those rows must be
parsed in full. Adding a new row can change later implicit row numbers; the
initial fast path must use the existing full parser for that case.

Shared formulas are a second closure. The parser gathers all members before
translating follower formulas and attaching `SharedFormulaStorage` in
[`semantic.rs#L169`](../../../../crates/litchi-xlsx/src/raw/worksheet/semantic.rs#L169).
Parsing only part of a group either loses provenance or creates a false
missing-master error. The safe initial gate is therefore a worksheet-level
fallback when any stored shared-formula owner is present. A later extension may
retain every row in a proven group, but it must prove the entire group before
using the reduced parser.

The immutable `Store` currently owns boxed sorted cells, row starts, metadata,
merge indexes, and four extents at
[`cell.rs#L630`](../../../../crates/litchi-xlsx/src/cell.rs#L630) and
[`cell.rs#L731`](../../../../crates/litchi-xlsx/src/cell.rs#L731). It has no
cheap mutable clone. The candidate therefore uses a fallible combined Store
rebuild that reuses source `Stored` semantics where the proof permits while
reconstructing the sorted cell and row indexes, merge indexes, and extents.
Numeric and formula lexical clones compete with the old parser's semantic
allocations, while the reduced parser avoids raw scratch for omitted cells;
allocation, peak memory, and latency must be measured together. A future
copy-on-write or delta representation is a separate optimization and is not
required for this design.

## Concrete readback mechanism

The cell-values source path should call a private specialized constructor, such
as `from_rewritten_cell_readback`, only after `SourceEdit::commit` or
`MultiSourceEdit::commit` has produced an effective cell plan. The existing
generic `from_rewritten_source` must remain a full-parse owner because it is
also called by page setup, margins, protection, hyperlinks, validations,
defined names, visibility, and other worksheet codecs. This keeps the proof
local to the value-only action model.

The rewrite operation should return the ordinary output and a private
`ReadbackProof` built while writing. It must not scan the completed output a
second time just to rediscover spans. The proof records the source identity,
the output row mode, the actual changed-cell spans, the row spans, and the
dimension result. The ordinary output remains the only bytes stored in the
published snapshot and in the reversible patch.

The reduced document is assembled from spans of the actual emitted output:

* retain the XML declaration, worksheet root, namespace declarations, all
  worksheet metadata, columns, defaults, sheet data tag, row tags, inter-row
  text, and declared dimension exactly as emitted;
* for an untouched row, retain its opening and closing row shell and omit only
  its cell owner spans;
* for a replacement-only row, retain its opening and closing row shell, all
  non-cell gaps, and each actual emitted changed-cell span; omit unchanged cell
  owner spans;
* for an existing row with insertion or removal, retain the complete actual
  row, including unchanged followers; and
* retain a complete row for any other shape only when its dependency closure
  is explicitly proven. Otherwise select the full-parser fallback.

The row and cell names are never recreated by a synthetic wrapper. This keeps
prefixes, default namespace rebinding, row-local declarations, whitespace,
comments, and the writer's exact escaping context. Coalescing the copied gaps
between omitted cells is an output-buffer optimization; it must not change the
retained bytes.

The raw parser receives the reduced document with the same `strings` callback
as the ordinary value-only path. It independently decodes every retained
changed cell and every complete membership-changing row. It also parses the
retained row and worksheet metadata. Omitted cells are admitted only by the
source-bound copy proof and the immutable pre-edit Store.

The proof needs three row modes:

```text
OmitCells       no effective action in the row
KeepChanged     existing-cell replacements; membership unchanged
KeepAll         insertion/removal in an existing row
```

The mode is computed from the final `Action` map, not from the user's staged
requests. Same-value requests do not enter the rewrite and remain the exact
no-op path. `KeepChanged` is valid only when every retained changed-cell span
contains the canonical explicit `r` emitted by `write_cell` and its address
matches the action map.

## Store merge contract

The reduced parse returns row metadata, columns, defaults, merges, and the
candidate declared extent along with the retained cells. The merge must be
fallible, source-independent after the proof is checked, and atomic:

1. For `OmitCells` rows, retain every base `Stored` entry at its original
   address.
2. For `KeepChanged` rows, retain base entries except at action addresses and
   take the parsed entry for each changed address.
3. For `KeepAll` rows, discard every base entry in the row and take all parsed
   entries from the emitted row. This lets the raw parser establish implicit
   follower addresses after an insertion or removal.
4. Reject duplicate addresses, an unexpected parsed address, or a missing
   required changed address. Do not expose these as new public errors; discard
   the candidate and run the full parser. The proof maps each retained emitted
   span to its source identity and expected action, while the reduced parser
   remains the independent decoder of the changed bytes. Existing staged
   readback continues to check the action semantics; no second full semantic
   comparison scan is required.
5. Rebuild the combined sorted cell and cell-row indexes without mutating the
   source Store; the full Store rebuild remains part of this candidate.
   Preserve every omitted entry's style, shared-string identity, inline-rich
   flag, formula range, shared-formula storage, and metadata.
6. Use the reduced parse's declared dimension. `expanded_dimension` can rewrite
   `<dimension>` even for a replacement-only edit when the source declaration
   did not cover an existing stored cell. Copying the source extent would make
   the semantic snapshot disagree with the emitted bytes.
7. Recompute `stored`, `content`, and `styled` extents exactly. A pure
   value-to-value or scalar-formula replacement keeps occupancy and content
   shape; `Clear`, `Remove`, and `Insert` can change bounds and must use an
   exact bounds update or take the full-parser fallback.

The parsed changed cell is the readback authority. Constructing it directly
from `Content` and then checking it against the same `Content` is not an
independent proof. Existing staged readback in
[`source.rs#L817`](../../../../crates/litchi-xlsx/src/cell_values/source.rs#L817)
and the multi-sheet path at
[`source.rs#L1122`](../../../../crates/litchi-xlsx/src/cell_values/source.rs#L1122)
must remain after the merged snapshot is created.

## Eligibility by action shape

All action shapes can use the same proof and fallback machinery, but the
initial candidate should admit only shapes whose closure and Store update are
fully implemented. The following is the safe design boundary:

| action | reduced readback mode | eligibility |
| --- | --- | --- |
| existing `Set(Value)` | `KeepChanged` | admit when source and output provenance are scalar and the changed cell parses to the requested value |
| existing `Set(Formula)` | `KeepChanged` | admit only for cacheless scalar formulas with `formula_range == None` and `shared_formula == None` |
| existing `Clear` | `KeepChanged` | admit only with exact `Empty` and extent handling; otherwise fall back |
| existing-row `Insert` | `KeepAll` | admit only when the row exists and the full emitted row is retained and merged |
| existing-row `Remove` | `KeepAll` | admit only when the row exists and the full emitted row is retained and merged |
| insert into an absent row | none | full parse; do not introduce a row-number shift proof in this candidate |
| `SetSharedFormula` | none | full parse; shared-formula worksheet closure is intentionally deferred |
| row, column, default, style, merge, or metadata action | none | remain on their existing owners and full parse paths |

The source-value editor already rejects unknown cells and cell metadata, and
the ingress callback rejects shared strings. The fast path must still assert
that every omitted or changed entry has `shared_string == None`,
`inline_rich == false`, and no cell/value metadata. If any assertion fails,
retain the full parser. Do not turn a new provenance kind into a guessed
scalar.

For a future membership extension, all row tags after a newly inserted row and
all cells whose inferred addresses could move would need explicit address
proofs. The current no-new-row gate is smaller and preserves the raw parser's
row inference without a second address-reconstruction algorithm.

## Validation, errors, resources, and publication

The specialized constructor must preserve the current observable sequence:

1. Check the source execution context.
2. Run `validation::worksheet_xml` over the complete actual output bytes.
   Eligibility checks and proof construction may only choose fallback; they
   must not publish a new error before this validation boundary.
3. If the proof is eligible, parse the reduced actual-output skeleton with
   the existing raw parser. If reduced semantic parsing, Store merging, proof
   validation, or exact retained-span checks fail, drop all reduced state and
   run `raw::worksheet::parse` on the complete actual output. The reduced
   error is never exposed: the full parser's error and precedence remain
   authoritative. A scratch `try_reserve` failure is recoverable after the
   candidate state is dropped; full parsing then runs and supplies the
   canonical result.
4. Install the merged Store beside the complete output bytes, preserving the
   source lineage and PartState identity rules.
5. Check execution again at the same outer boundary, then keep calculation-chain
   invalidation, staged semantic readback, publication, and reopen checks
   unchanged.

The fast path must not return a reduced-skeleton error. This rule matters for
malformed retained fragments, unsupported future attributes, and partial
shared-formula groups: the exact full parser must decide the public result.
The no-op branch must continue returning a patch that shares the original
source bytes without constructing a proof or skeleton.

The complete output remains subject to the existing worksheet byte and
aggregate limits. Skeleton capacity, retained-span metadata, and merged Store
reservations need checked sizes and operation-local accounting. A failed
`try_reserve` makes the candidate ineligible and releases the reduced state
before full fallback; it must not become a new format limit. Avoid adding
per-cell cancellation checks between the existing validation and parse checks,
because that would alter cancellation precedence. Keep the existing outer
execution checks and caller budget ownership.

The candidate must conservatively fall back when the actual output contains
MCE controls, foreign or mixed namespace candidates, x14ac descent markers,
or any source feature outside the already validated value-only grammar. This
avoids changing `x14ac::capture` or `process_ooxml` ordering on a reduced input.
An unused namespace declaration is enough to choose fallback if the marker
preflight cannot prove that extension processing is unnecessary.

The reduced skeleton is temporary. The published source bytes are always the
full output, and the patch's reverse operation uses those full bytes. No
source-backed snapshot may retain a borrow into the temporary skeleton.

## Required differential checks before retention

Performance retention remains unproven. A later pilot must test
the reference and candidate for:

* exact output bytes and untouched package members;
* value, formula, empty, style, row, dimension, and extent readback;
* XML escaping, inline text, explicit and inferred row/cell addresses,
  prefixed/default namespace scopes, empty rows, and noncompact tags;
* replacement-only rows with omitted preceding cells;
* existing-row insertion and removal with implicit followers;
* absent-row insertion fallback and all shared-formula fallback cases;
* MCE, x14ac, foreign, rich-inline, shared-string, unknown, malformed, and
  resource-limit cases;
* staged readback, calculation-chain invalidation, save/reopen, source/version
  checks, cancellation, and exact no-op sharing; and
* fallback error type and display text for every reduced-parser rejection.

The profile should expose enough private mechanism evidence to distinguish
`KeepChanged`, `KeepAll`, fallback reasons, retained cells, omitted cells,
skeleton bytes, and merged Store work. The static closure counts justify this
instrumentation but do not predict a speedup. Native total and commit gates,
allocator/reduced-peak results, adverse rows, and retained baseline custody
remain the admission authority in the frozen 0525 plan.

## Safe fallback and next work

The candidate deliberately keeps the full combined Store rebuild, cell and
row indexes, merge indexes, and extents. Its opportunity is removing raw XML
decoding and semantic materialization for omitted cells; source numeric and
formula lexical clones remain, and the old parser already allocates semantic
lexicals alongside raw-parser scratch. Compare total allocations, peak
memory, and latency before claiming retention. Do not claim that
`Store::from_unsorted` is redundant merely because the edit scanner observed
sorted rows: raw source parsing intentionally accepts and sorts some cell
orderings, and changing that contract would alter source acceptance and error
precedence.

If the reduced-readback pilot fails its total or memory gate, retain the full
parser and profile the changed-output validation and compaction boundary
described by the 0514 follow-up. That work occurs only after an effective edit
and can be investigated independently without reviving parser/layout or
emitted-event fusion. OLE2/OOXML stays ahead of ODF.

## Applicable accepted ADRs

* ADR 0001 keeps raw grammar ownership private and typed refusal behavior.
* ADR 0003 requires immutable snapshots, atomic commit, exact no-ops, final
  semantic readback, and source-checked patches.
* ADR 0005 requires caller-owned I/O, memory, work, and cancellation budgets.
* ADR 0006 requires preservation, validation before publication, and fail-closed
  handling of unsupported markup.
* ADR 0010 and ADR 0011 keep archive and physical OPC ownership below the
  facade.
* ADR 0017 and ADR 0018 keep OOXML producer templates and calculation-chain
  ownership in their format owners.

This review amends no ADR and makes no production performance claim.
