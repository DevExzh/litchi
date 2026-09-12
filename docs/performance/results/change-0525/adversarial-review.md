# XLSX reconstruction reuse: adversarial review

Status: frozen read-only review of the preflight-3 source manifest at
`docs/performance/results/change-0525/preflight-3/source-manifest.json`.
This review made no production edits and ran no builds, tests, or captures.
The preflight-1 compile failure is retained below. Preflight-3 is the canonical
source snapshot for this review; its XLSX test and Clippy gates both completed
successfully. The normal release build was still live when this review froze,
so this review does not claim a release-build result.

Canonical source manifest SHA-256:
`0fab9bafc238611761659bcafdd2e7b5df2543aecb194ac5477c168a2d6e6096`.
The terminal gate records are
`docs/performance/results/change-0525/preflight-3/check-xlsx-tests.receipt.json`
and
`docs/performance/results/change-0525/preflight-3/check-clippy.receipt.json`.
The parent test record reports 1,288 XLSX tests; both receipts have exit code
0 (the test receipt's unit-test stream begins with 980 tests).

The reviewed implementation hashes are:

* `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs`:
  `71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5`
* `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs`:
  `b41af5d8c91a5c1e82f9030b8798a354ca4943072373846b8874c73996c03035`
* `crates/litchi-xlsx/src/cell_values/snapshot.rs`:
  `e974e6e94d4af3c4797d883a60410726b61081ca99adea46f786099485b6b4d6`
* `crates/litchi-xlsx/src/cell.rs`:
  `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7`
* `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot.rs`:
  `f1e4666ce3c08f4aaa6affe4e684e596b30453166f95c17576457a5b8491ea05`
* `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs`:
  `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8`
* `crates/litchi-xlsx/tests/source_backed_cell_values.rs`:
  `d8270d5682af4764be54428c520b979a2cec035d881862654bbf5b21425fd166`

## Decision

The frozen writer has no remaining source-owner or changed-address blocker in
the reviewed path. The reconstruction can preserve the semantic contract as
an explicitly guarded fast path. The important boundary is narrower than
“parse the changed cells”: every cell record that remains in the reduced XML
must have the same effective address that it had in the complete candidate.
The reduced parser may then validate changed records and worksheet-level
state, while records proven byte-identical to the source can be imported from
the already validated source `Store`.

Deleting an unchanged cell span is safe only when all source records omitted
from that row are absent from the reduced stream and every record retained in
that row has an explicit `r` written by the editor. If an unchanged implicit
cell is retained, or if an implicit cell precedes a retained cell without an
explicit reference, deleting the former changes the latter's inferred
address. A safer general fallback is to retain an empty `<c>` wrapper for
every omitted implicit cell; the wrapper consumes the same sequential column
while the source `Stored` record supplies the original payload after parsing.
The merge must replace that address carrier with the one source `Stored`
entry, rather than append both entries and trigger a duplicate.

Membership-changing rows should remain on the complete-parser path. The
candidate proposes this fallback row-locally: a row containing a `Remove` or
cell creation/insertion is rendered by the complete `write_row`, while
unaffected rows can still contribute omitted spans. Any newly materialized
row is likewise complete. A worksheet-wide fallback is a valid conservative
implementation, but it is not required by this proof boundary. Replacement
only updates (`Set`, `Clear`, and scalar `SetFormula`) can use the reduced path
when their changed records are emitted with explicit references and no
exceptional worksheet-wide dependency is present.
Eligibility must be computed from the final effective action map after
same-value requests have been elided, and the proof must carry the source
identity and execution/version fence that authorize reuse.

The existing `from_rewritten_source` constructor is shared by page setup,
visibility, protection, links, defined names, and other worksheet owners. Keep
that generic entry point on its complete-parser behavior. The reuse proof must
enter through a private value-only constructor that receives the writer's
source-bound spans and action plan; otherwise an unrelated worksheet rewrite
could accidentally inherit cell-value assumptions.

## Why the boundary matters

The raw worksheet parser infers a missing cell reference from the preceding
cell in the same row. Its `start_cell` resets the row-local column counter and
uses the preceding `last_column` when `r` is absent
(`raw/worksheet/codec.rs` around lines 735–760). The frozen writer's
`write_replacement_row` copies each unchanged `CellSlot` from the source and
records the actual output interval, while `write_cell` removes any old `r`
attribute and appends the canonical `cell.address.a1()` reference before
writing every changed existing cell (`raw/worksheet/edit/codec/snapshot/write/sheet_data.rs`
around lines 189–245 and 485–518). This is the fact that makes a
replacement-only row reducible. It does not emit references for untouched
records.

For example, this valid source has two inferred cells:

```xml
<row r="1">
  <c r="A1"><v>1</v></c>
  <c><v>2</v></c>
  <c><f>SUM(A1)</f></c>
</row>
```

If `A1` is replaced and the numeric cell is dropped while the unchanged
formula cell is retained, the reduced stream parses the formula as `B1`, not
`C1`. Importing the old `B1` record then either creates a duplicate address
or silently attaches the formula to the wrong coordinate. Keeping the cell
wrapper (`<c></c>`) or retaining an explicit `r="C1"` fixes the inference.
The same issue occurs whenever a policy retains a formula, shared-formula
member, metadata cell, or any other original record in a row after dropping a
preceding implicit record.

The existing writer supplies a stronger counterexample for membership changes.
Given `A1` explicit, `B1` explicit, and `C1` implicit, removing `B1` leaves
the unchanged `C1` span in the complete output without an `r`. A complete
parse therefore makes that record `B1`. The current publication readback
correctly rejects the result because `B1` still exists. A reduced parser that
imports the source `C1` record without parsing the changed row would miss this
failure. Parsing every membership-changing row completely is therefore
required even if other rows use the reuse path.

Rows themselves also carry inferred numbers. The reduced XML must retain each
row start/end (or an exact self-closing row), including rows whose bodies are
removed. This preserves `previous_row`, duplicate-row checks, row properties,
and namespace scope in `raw/worksheet/codec.rs` around lines 559–633. Removing
whole untouched row elements would make implicit row numbers depend on the
reduction and invalidate the proof.

## What remains independently validated

`Snapshot::from_rewritten_source` first performs the independent worksheet XML
validation and only then calls the semantic parser
(`cell_values/snapshot.rs` around lines 686–694). The fast path must retain
that order on the complete candidate. The complete validation is valuable
because it still sees every candidate byte, including copied untouched cells,
and checks the allowed SpreadsheetML namespace, element tree, attributes, and
XML text. It is not a replacement for the raw parser on changed rows: it does
not check numeric/date lexicals, row/cell ordering, formula invariants,
style bounds, or shared-formula resolution.

The reduced XML should preserve all non-row XML and all row boundary tags. Its
raw parse then needs to cover:

* worksheet root, dimension, views, defaults, columns, and row attributes;
* every changed cell record, with its generated explicit address;
* every row containing a membership change, in its complete candidate form;
* all members of any shared-formula group that is allowed through the fast
  path.

Any eligibility mismatch, reduced-parse error, merge invariant failure, or
fallible reservation failure should discard the reduced state and run the
existing complete parse on the actual output. The fast path must not introduce
a new observable error or change the complete parser's error precedence.

The source `Store` is a valid source-side proof for omitted bytes: source
capture already runs worksheet validation, raw parsing, style-reference
validation, and scalar-cell validation
(`cell_values/snapshot.rs` around lines 379–388 and 656–664). Since the writer
copies an untouched row byte for byte when there is no action for it
(`write/sheet_data.rs` around lines 52–72), those source records can be reused
after the complete candidate has passed XML validation. The merge rebuilds the
final `Store` through its duplicate checks and index/extent construction,
rather than leaving the candidate's partial extents or cell-row index in
place.

The touched-row proof is also present in the frozen renderer. Its run starts
after the exact source-to-output gap before the first unchanged cell, extends
only across unchanged cell spans and intervening source gaps, and is flushed
before each changed cell and at the row body end. The scanner rejects duplicate
or descending cell addresses before this renderer runs. Thus the recorded
first/last address range names exactly the source owners covered by that byte
interval; it does not include a changed cell. The final tests compare those
intervals against source slices and compare provenance output bytes with the
ordinary renderer.

The following cases need a conservative guard or fallback:

* Any `Remove`, insertion, or newly created row. These can change membership or
  inferred addresses, so the whole affected row must be parsed. The simplest
  implementation can exclude the worksheet from cell-span reduction whenever
  a membership-changing action is present, but the proof obligation is
  row-local: unaffected rows may still use reuse when the implementation can
  isolate the complete affected rows.
* Shared formulas. `resolve_shared_formulas` needs the master and every member
  together, and translates the master expression into each follower
  (`raw/worksheet/semantic.rs` around lines 169–285). A worksheet-level
  fallback is sound. A narrower implementation could admit a group only when
  every member is present in the reduced candidate, or when every member is
  omitted and imported as a complete source group; it must not admit a partial
  group.
* Array/data-table formulas and other formula-range records. Cell-value action
  guards refuse editing their covered cells, so source reuse is possible when
  all such records are omitted and imported intact. Retaining only a subset in
  the reduced stream adds unnecessary proof burden; fallback is safer until an
  exact range-ownership check exists.
* Shared strings. The source-backed value-only loader invokes the worksheet
  parser with no shared-string callback, and `materialize` requires that table
  for `t="s"` values (`raw/worksheet/codec.rs` around lines 384–402 and
  `raw/worksheet/semantic.rs` around lines 134–145). Current admission should
  preserve that existing refusal. If shared-string support is added later, the
  reduced parser must receive the same table and preserve index checks.
* MCE, foreign namespaces, and extension attributes. The value-only XML
  validator currently rejects foreign elements and namespaced attributes, so
  the current capability has no extension state to reconstruct. If that input
  boundary is widened, extension capture (`x14ac` and MCE branch selection)
  becomes a worksheet-wide dependency and the fast path needs a new proof or
  fallback.

For a replacement-only row, deleting all untouched cell elements is valid even
when the source cells used inferred references, because the only cells left in
that row are the changed records and the writer gives each one an explicit
`r`. This claim is encoded by the row membership check: a missing action owner
or `Remove` takes the row through `write_row`, and the replacement helper can
only retain existing `Action::Update` owners. Any future writer behavior that
retains an unchanged record or emits an implicit new record must preserve that
guard or force the row to complete parsing.

## Ownership and performance risk

The semantic merge is not free. `Store` owns a `Box<[Stored]>` and its indexes
(`cell.rs` around lines 631–799); the frozen source makes `Stored` cloneable,
while `Cell` and numeric/formula lexical strings remain owned. A merge that
clones omitted `Stored` records has a measurable ownership cost, especially
for owned numeric and formula lexical text. Compare that cost directly with
the complete parser's existing `Number` allocations and raw scratch in the
allocation, peak-live, RSS, and latency gates; this review does not predict
the net result.

The merge helper therefore needs fallible reservation, duplicate detection,
and measurements of clone allocations. A persistent base-plus-overlay store
could avoid deep cloning, but it would change the `Store` representation and
all binary-search, traversal, and extent consumers; that is a separate design
and should not be smuggled into this fast-path patch. The candidate should be
admitted only after the retained source store, reduced parse, and merge costs
are measured end to end.

## Required adversarial checks before admission

The frozen focused tests compare a normal complete candidate parse with the
reduced-plus-merged result, including cell kind, formula text/cache/range and
shared metadata, rows, columns, defaults, merges, and all four extents. They
also compare the provenance writer's complete bytes with the ordinary writer,
assert every changed cell's explicit `r`, and compare each recorded omission
interval with the corresponding source slice. Include or retain these
fixtures:

For each replacement-only touched row, also assert the reduction invariant
directly: every unchanged source cell owner is omitted, or every omitted
implicit owner has an explicit address carrier that the merge replaces with
exactly one source entry. No original implicit owner may remain as a second
semantic entry.

1. Replacement after one or more inferred cells, replacement before inferred
   cells, and rows whose references are all explicit.
2. Empty and self-closing rows/cells, whitespace/comments between records, and
   prefixed SpreadsheetML names with declarations on the row or cell.
3. Ordinary scalar formulas, array/data-table formulas in unrelated rows, and
   shared-formula groups that cross omitted rows.
4. Removal and insertion in a row with inferred references, proving the
   complete-row fallback and the existing readback rejection behavior.
5. Malformed candidate payloads and invalid numeric/date/style values in the
  changed row, proving that full validation and reduced raw parsing retain
  typed refusal and error precedence.

The remaining test gaps are bounded. Add an unchanged scalar formula to an
omitted run (including an implicit formula owner), and add an unrelated
array/data-table formula fixture, so source-owned formula metadata is exercised
when its XML body is removed. Add one invalid lexical value in a retained
changed cell to prove the reduced parser's failure falls back to the complete
candidate parser. A fault-injected metadata reservation test would directly
exercise the fail-open `recording` flag, although the frozen writer already
clears partial provenance and returns the complete output when either
reservation fails. These are test gaps, not a discovered semantic blocker.

Preflight-1 did fail to compile because the new writer was not re-exported from
`codec/snapshot.rs`; preflight-2 added that re-export, and preflight-3 is the
canonical source frozen for this review. The preflight-3 test and Clippy
receipts both have exit code 0. The normal release build was live at freeze,
with no terminal result available to this review.

The full provenance-writer output is published and remains byte-identical to
the ordinary rewrite. The reduced XML is an internal readback/semantic
construction only; it must never become the published worksheet payload. This
preserves the exact patch and source-provenance behavior in `SourceState` while
excluding only proven unchanged cell payloads from semantic XML parsing. Output
validation still scans the complete bytes, and reconstruction still rebuilds
the full `Store` indexes/extents and clones imported entries.
