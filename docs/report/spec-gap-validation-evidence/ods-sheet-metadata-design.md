# ODS sheet metadata lifecycle design

Status: bounded design and evidence only. This note does not change the ODS
production API, feature matrix, or audit row. It proposes one source-bound
metadata batch for `litchi-ods`; formula evaluation, external-source access,
data styles, and the existing source-adapter work remain outside this design.

The design was reviewed against the current worktree on 2026-09-12. The
important distinction is that the audit row groups several families with
different implementation states. Scenarios already have a public transaction,
and `table:cell-range-source` already has a typed worksheet read/write path.
Consolidation, label ranges, and detective metadata still lack a complete
public owner. A focused owner can close the group without replacing the
existing worksheet `Cell`/`Sheet` graph.

## Current implementation boundary

The current compiled ODS surface is as follows.

| Family | Existing code | What callers can do today | Remaining gap |
|---|---|---|---|
| Scenarios | `crates/litchi-ods/src/model/scenario.rs` and `model/scenario/transaction.rs` | `Spreadsheet` and `MutableSpreadsheet` expose source-bound `Snapshot`/`Edit`/`Commit`/`Patch` CRUD with exact-source checks and inert values | No work belongs in this batch. Reuse its transaction and facade patterns as a reference. |
| Consolidation | `crates/litchi-ods/src/model/consolidation.rs` | `Options`, validation, standalone `parse_consolidation`, and canonical `write_consolidation` exist | No `Spreadsheet`/`MutableSpreadsheet`/source-backed owner, no read/write set, inverse patch, package readback, or caller limits. The parser has no explicit resource-limit context. Its integration tests are under `cfg(any())` and are not executable evidence. |
| Label ranges | `crates/litchi-ods/src/model/label_range.rs` | `Range`, validation, standalone `parse`, and canonical `write` exist | No public facade or source-bound transaction. The disabled tests refer to missing facade methods. The parser/writer has no owner source spans, explicit limit context, or preservation contract. |
| Detective | `crates/litchi-ods/src/model/detective.rs` | Typed `Direction`, `OperationKind`, `HighlightedRange`, `Operation`, `Detective`, and a writer exist | There is no parser or typed worksheet field, so no public readback or safe replacement path exists. The live worksheet codec classifies `table:detective` as `Other`; it is not a detective implementation. |
| Cell range source | `crates/litchi-ods/src/model/source.rs`, `worksheet/model.rs`, `worksheet/codec.rs` | `Cell::range_source`, `set_range_source`, and `take_range_source`; direct empty `table:cell-range-source` is parsed in table and covered cells; existing worksheet edits can replace a complete cell and round-trip it | There is no focused metadata catalog or `set/clear` operation addressed by `CellSelector`, nor an owner-specific patch/read set. Existing row rewrites deliberately refuse noncanonical source-local lexical markup (`worksheet/package.rs:validate_cell_range_source_lexical`) rather than preserving it. Reuse this model and refusal seam; do not introduce a second `CellRange`. |

`crates/litchi-ods/src/model/cell.rs` contains a separate detective-bearing cell
model, but `model/mod.rs` does not export it. It is not the production
worksheet graph and must not be copied into the new facade. The live graph is
`crates/litchi-ods/src/worksheet/model.rs`, where `Cell` currently has a range
source but no detective field.

The current public worksheet transaction in
`crates/litchi-ods/src/worksheet/snapshot.rs` replaces whole typed cells or
whole sheets. It is useful infrastructure for locating repeated physical runs,
but it is too broad to serve as the metadata owner's preservation boundary: a
metadata edit must splice only the selected owner and its minimal insertion
anchor, then reopen the complete candidate.

## Normative ownership and grammar

The checked-in OpenDocument 1.4 archive is
`3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
The relevant members are:

* `schemas/OpenDocument-v1.4-schema.rng`, SHA-256
  `4034ec6be29205d5fc1ee5f42468ac6ef824287b3aba6d9289032af4fafbda7f`;
* `part3-schema/OpenDocument-v1.4-os-part3-schema.html`, SHA-256
  `43fb603f9f54f030db7082518aff6f136297d182abb20124a859d271a7969a15`.

The RNG is the authority for ordering and cardinality. The exact spreadsheet
child sequence is a prelude, zero or more `table:table` elements, and an
epilogue. The prelude admits optional tracked changes, `text:decls`, and the
`table-decls` group. Inside that group the order is calculation settings,
content validations, then optional `table:label-ranges`. The epilogue is the
`table-functions` group. Each member is optional, and their order when present
is named expressions, database ranges, DataPilot tables,
`table:consolidation`, then DDE links.
`table-decls` and `table-functions` are schema groups, not XML wrapper
elements: all four owners are children of their actual ODF parent. An
insertion must therefore anchor after the preceding recognized group member
and before the following recognized member, while preserving comments,
processing instructions, whitespace, and foreign siblings. If an unmodeled
direct child makes that anchor ambiguous, the edit returns a typed refusal.

* `table:consolidation` is empty and requires `table:function`,
  `table:source-cell-range-addresses`, and `table:target-cell-address`; the
  source attribute uses `cellRangeAddressList`, defined as `string` by the RNG
  without a non-empty restriction. Any stricter authoring validation must be
  identified as API policy rather than a schema requirement. The
  optional `table:use-labels` values are `none`, `row`, `column`, and `both`,
  and `table:link-to-source-data` is Boolean. The function accepts the listed
  standard names or an application-defined string. The editor retains it and
  never invokes it.
* `table:label-ranges` contains zero or more empty `table:label-range`
  children. Each child requires label and data cell-range addresses and an
  orientation of `row` or `column`.
* `table:cell-range-source` is empty, requires its table name and positive
  last row/column spans, and carries `xlink:type="simple"`, `xlink:href`, and
  the optional linked-source attributes. The Part 3 §9.3.1 prose says it
  represents a database or named range in another file,
  is in the first cell of the range, and is usable directly in both
  `table:table-cell` and `table:covered-table-cell`.
  Within `table:table-cell-content` the direct child order is exactly
  `table:cell-range-source`, optional `office:annotation`, optional
  `table:detective`, then zero or more text-content children. The editor may
  splice one of these direct slots but must not move a source past an
  annotation, place detective after text, or treat a nested same-named element
  as the cell owner.
* Part 3 §9.3.2 defines `table:detective` as a container for
  `table:highlighted-range` and `table:operation`, with no attributes, usable
  directly in either cell form. Section §9.3.3 makes `table:operation` empty
  with `table:name` and non-negative `table:index`; §9.3.4 makes
  `table:highlighted-range` empty, with either the valid-range attributes or
  `table:marked-invalid`. The highlighted state is a snapshot of what was
  calculated at the time; the editor must not recompute it.

The schema permits meaningful absence and emptiness that the existing leaf
codecs currently collapse:

| Owner | Absent | Present but empty | CRUD consequence |
|---|---|---|---|
| `table:consolidation` | No singleton declaration | Impossible: the element is empty but its required attributes still carry the declaration | `None` removes the owner; `Some(Options)` inserts or replaces the one owner. |
| `table:label-ranges` | No container | A present container with zero `table:label-range` children | The snapshot retains `present`; `ensure_label_ranges` may create an empty container, while `clear_label_ranges` removes the container. Replacing with an empty vector must not silently choose between those two states. |
| `table:cell-range-source` | No direct child in the selected cell | Impossible: it is empty but required attributes identify the source | `None` clears the direct child; `Some(CellRange)` inserts or replaces it. |
| `table:detective` | No direct child in the selected cell | A present container with zero highlights and zero operations | `None` removes it; `Some(Detective::new())` preserves an explicitly present empty container. |

Detective highlights and operations retain separate document-order vectors.
The schema requires all highlighted ranges before all operations; the numeric
`table:index` is operation metadata, not a sort key. `add_operation`,
`insert_operation`, `replace_operation`, and `remove_operation` therefore act
on checked source-order positions and never reorder by `index`. A changed
detective owner must preserve that order and the distinction between an empty
container and no container.

The codec must resolve namespace bindings by expanded URI and local name, not
by the source prefix. Encoded-equivalent attribute values and duplicate
expanded attributes are checked before projection. The source prefix, quote
style, entity spelling, comments, processing instructions, and declaration
ordering remain preservation data when a changed owner can retain them; a
changed operation that cannot retain them returns a typed unsupported-owner
error.

## Smallest complete batch

Add one focused `sheet_metadata` owner under
`crates/litchi-ods/src/sheet_metadata/` (or an equivalently scoped module)
with a borrowed scan/index and a source-backed transaction. The owner reuses
the four existing semantic models rather than introducing alternate
consolidation, label-range, detective, or cell-range-source values.

The initial public shape should be selector-first and source-bound:

```rust,ignore
pub mod sheet_metadata {
    pub struct Snapshot { /* immutable content.xml source and bounded index */ }
    pub struct Edit { /* source-checked staged operations */ }
    pub struct Commit { /* new Snapshot and reversible Patch */ }
    pub struct Patch { /* exact source-authorized forward bytes */ }
    pub struct Limits { /* caller profile within hard ceilings */ }
}

impl Spreadsheet {
    pub fn sheet_metadata(&self) -> Result<sheet_metadata::Snapshot>;
    pub fn sheet_metadata_with(
        &self,
        limits: sheet_metadata::Limits,
        context: &litchi_core::ExecutionContext,
    ) -> Result<sheet_metadata::Snapshot>;
    pub fn edit_sheet_metadata<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut sheet_metadata::Edit) -> Result<()>;
    pub fn edit_sheet_metadata_with_context<F>(
        &mut self,
        limits: sheet_metadata::Limits,
        context: &litchi_core::ExecutionContext,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut sheet_metadata::Edit) -> Result<()>;
    pub fn apply_sheet_metadata_patch(&mut self, patch: &sheet_metadata::Patch)
        -> Result<()>;
}

impl SourceBackedSpreadsheet {
    pub fn sheet_metadata(&self) -> Result<SourceMetadataSnapshot<'_>>;
    pub fn edit_sheet_metadata(&self) -> Result<SourceMetadataEdit<'_>>;
    pub fn apply_sheet_metadata_patch(
        &self,
        patch: &SourceMetadataPatch<'_>,
    ) -> Result<SourceMetadataCommit<'_>>;
}

impl<'source> SourceMetadataEdit<'source> {
    // The same consolidation, label-range, cell-range-source, and detective
    // CRUD verbs as sheet_metadata::Edit.
    pub fn commit(
        &mut self,
        context: &litchi_core::ExecutionContext,
    ) -> Result<SourceMetadataCommit<'source>>;
}

impl<'source> SourceMetadataCommit<'source> {
    pub fn patch(&self) -> &SourceMetadataPatch<'source>;
    pub fn write_to<W: std::io::Write>(
        &self,
        writer: W,
        options: SourceContentPublicationOptions,
    ) -> std::result::Result<
        SourceContentPublicationReport,
        SourceContentPublicationError,
    >;
}

impl<'source> SourceMetadataPatch<'source> {
    pub fn apply(
        &self,
        snapshot: &SourceMetadataSnapshot<'source>,
    ) -> Result<SourceMetadataCommit<'source>>;
    pub fn inverse(&self) -> Self;
}
```

`MutableSpreadsheet` forwards the same operations. `SourceBackedSpreadsheet`
must expose the same read snapshot, full CRUD, and source-version checks, and
must publish through its existing source replacement seam while preserving the
same source identity and lexical refusal rules. No API should expose package
member names, native IDs, relationship IDs, or a catalog-only string mode.
`Builder` methods for detached authoring are deliberately deferred; this batch
is for existing source-bound documents and can add an owner when absent.
The convenience methods without an explicit context must delegate to the same
bounded implementation using a named finite profile; they are not an
unbudgeted allocation path. The `_with`/`_with_context` forms are the required
route for callers that share a hierarchical budget or cancellation token.
`SourceBackedSpreadsheet` is not a read-only exception: its source-backed
snapshot, edit, commit, patch/inverse/apply, and `write_to` operations are part
of this batch. `SourceMetadataEdit::commit` takes `&mut self`; if validation,
limits, source freshness, rendering, or readback fails, the staged edit and
its original source snapshot remain available unchanged. A successful commit
does not mutate the source; its `SourceMetadataCommit::write_to` publishes the
candidate package through the source-backed publication writer.

The snapshot should expose typed reads such as:

```rust,ignore
snapshot.consolidation() -> Result<Option<&Options>>;
snapshot.label_ranges() -> LabelRangesView<'_>; // includes present/absent
snapshot.cell_metadata(sheet_metadata::CellSelector::by_name("Sheet1", 0, 0))
    -> Result<Option<CellMetadataView<'_>>>;
```

`CellMetadataView` is a single focused view over the already-public
`CellRange` plus the new typed detective value. It must distinguish a missing
physical cell from an existing cell without metadata, and report the physical
row/cell run and repetition when a logical selector lands in a repeated run.
The metadata API uses one `sheet_metadata::CellSelector` type whose sheet
component is the existing `worksheet::Selector`: `by_name` accepts an exact
sheet name and `by_position` accepts a checked zero-based position. The row
and column are always checked zero-based logical coordinates. There is no
second name-only overload and no implicit first match; exact-name ambiguity is
an error. The existing root `litchi_ods::CellSelector` remains a compatibility
lookup type for the current facade, but metadata CRUD must either adapt it
explicitly or use the focused selector everywhere.

Its conceptual constructors are `CellSelector::by_name(&str, row, column)`
and `CellSelector::by_position(Position, row, column)`; both resolve against
the same immutable snapshot and participate in the same read set.

The edit supports these operations:

* `set_consolidation(Option<Options>)` and `clear_consolidation()` address the
  singleton owner;
* `add_label_range`, `insert_label_range`, `replace_label_range`,
  `remove_label_range`, and `clear_label_ranges` use checked source-order
  positions. The staged list preserves order;
* `set_cell_range_source`, `clear_cell_range_source`, `set_detective`,
  `clear_detective`, and a checked detective editor address a logical cell
  selector. They use the repeated-run contract below;
* every operation is allowed to be a semantic no-op. A missing cell may be
  created only under an explicit, bounded cell-creation policy; the first
  batch should refuse implicit creation when the surrounding row/table
  structure is not already representable by the existing worksheet writer.

The `detective` model needs a bounded parser and exact readback before the
editor is exposed. Its existing writer is not sufficient evidence. The
parser must enforce the schema's valid-versus-invalid highlighted-range choice,
operation ordering, non-negative index, duplicate expanded attributes, empty
content, and bounded child counts.

### Selector and repeated-run contract

All cell operations use the focused `sheet_metadata::CellSelector` defined
above: either an exact sheet name or a checked zero-based sheet position, plus
logical row and column coordinates. Both forms are required in this batch and
resolve through `worksheet::Selector`; a native table ID or an implicit first
matching sheet is never accepted. The scan index
stores each physical row run as a checked half-open logical interval and each
physical cell run inside it as a checked column interval. Lookup returns the
covering physical owner and its repeat counts without expanding either run.

For a changed operation whose target falls inside a repeated row or cell, the
owner must either perform this exact split or return a typed
`UnsupportedRepeatedRun`/`OpaqueCell` error before rendering:

1. Split the physical row into prefix, one target row, and suffix runs, when
   needed. Copy the row attributes and untouched child spans into each output
   run; retain the original repeat values on prefix/suffix and emit the target
   with one logical row.
2. Within the target row, split the covering physical cell into prefix, one
   target cell, and suffix runs. Copy cell attributes and untouched child
   spans; retain `number-columns-repeated` on prefix/suffix and emit the
   target with one logical cell.
3. Splice only the target cell's direct metadata sequence. The target's
   source/detective operation applies to one logical cell; it is not silently
   broadcast to the rest of the original run. Clearing metadata on one member
   of a repeated run uses the same split. A semantic no-op must not split.
4. Validate the resulting run counts, covered-cell role, merge spans, child
   order, and row/table ownership before publishing. If an opaque child,
   foreign namespace fragment, unsupported text structure, or noncanonical
   source fragment cannot be copied through the split, return the typed
   refusal and retain the entire future repeated-run capability as explicit
   implementation debt. Never fall back to replacing the whole row or sheet.

The first batch must test both paths. A refusal is a safe bounded result, not
permission to treat a logical target as the entire physical run.

### Merge and implicit-covered coordinates

The selector contract also preserves the worksheet model's three physical
merge roles. `Merge::None` is an ordinary physical cell and can carry the
direct metadata sequence. `Merge::Span { rows, columns }` is an anchor; its
anchor coordinate is selectable, while coordinates inside the span that have
no physical `table:covered-table-cell` are implicit covered coordinates.
`Merge::Covered` is an explicit physical covered cell and may be selected for
the direct metadata allowed by the ODF cell-content grammar. The scan index
records the role, anchor, and checked span for every returned view.

Reads of an implicit covered coordinate return a distinct
`CellMetadataView::ImplicitCovered { anchor, span }`, never `Missing` and never
the anchor's direct metadata as though it belonged to the selected coordinate.
The first lifecycle batch refuses writes to that implicit coordinate with
`UnsupportedImplicitCoveredCell`; it does not materialize covered cells or
change merge geometry. Writes to `Merge::None`, a span anchor, or an explicit
`Merge::Covered` cell follow the repeated-run split contract and validate that
the source/annotation/detective/text child order remains legal. Materializing
implicit covered cells is retained as an explicit future scope item.

## Transaction and publication contract

`sheet_metadata::Snapshot` owns or borrows the exact `content.xml` source and a
compact owner index. It must not clone a complete `Spreadsheet` or expand
repeated rows and cells. The index records owner spans, direct parent kind,
namespace context needed for insertion, and physical row/cell locators. Typed
values are projected only for recognized, schema-valid owners; a bounded raw
owner with diagnostics is retained for malformed or unsupported markup so an
exact no-op can still preserve it. Selecting such an owner for a changed edit
returns a typed refusal before any bytes are published.

The edit records semantic operations and expected-state fingerprints rather
than mutating the source. Its read set includes the exact source identity and
version, selected owner spans, parent and sibling ordering, namespace context,
and every row/cell run boundary traversed to resolve a selector. Its write set
contains only the selected singleton, label child, or direct cell child and
the minimal insertion/split anchors. Two edits compose only when these sets
are proven disjoint; an overlap or stale source is a structured conflict, not
last-writer-wins behavior.

Commit must proceed in this order:

1. Validate all selectors, staged values, owner cardinalities, source
   fingerprints, and read/write-set conflicts without changing the source.
2. Precharge the output and scratch budgets. Render only changed owner
   fragments with the active namespace context, or return a preservation
   refusal if canonical rendering would discard source-local lexical data.
3. Splice the changed spans and anchors into a candidate `content.xml`; do
   not rebuild unrelated tables, rows, cells, or package members.
4. Reopen the complete candidate package under the retained limits, parse the
   focused owner, and verify every staged value plus the owner placement and
   unchanged dependency closure.
5. Publish the new immutable snapshot and a reversible `Patch`. The source
   snapshot remains unchanged. `Patch::apply` requires exact source bytes and
   `Patch::inverse` restores the accepted source byte-for-byte.

The ordinary `Edit` uses the same non-consuming signature as the source-backed
edit: `commit(&mut self, &ExecutionContext) -> Result<Commit>`. On error it
retains its original base snapshot and every staged operation, allowing the
caller to inspect, correct, or roll back the edit. On success it returns the
accepted immutable commit while retaining the staged ledger until an explicit
`clear`/`rollback` operation; a second commit is therefore deterministic and
source-checked rather than silently consuming the only recovery state.

The focused patch scope is exact `content.xml`, not an implicitly normalized
or whole-ZIP byte image. A `sheet_metadata::Snapshot` retains the source
`content.xml` bytes (or an immutable owner of them), a source identity/version
when supplied by the package adapter, and the source digest used for cheap
diagnostics. `Patch::apply` requires the exact original `content.xml` bytes and
matching source lineage; for `SourceBackedSpreadsheet` it also checks the
current `SourceVersion` immediately before and after the read/splice. The
patch stores exact before/after `content.xml` bytes so `inverse()` restores
that XML fragment exactly, including the original preamble and trailing bytes.

The ordinary owned `Spreadsheet` facade applies the accepted content patch
through the existing package `replace_content_xml` seam and then rehydrates
the complete package. ZIP local/central records, compression, and unrelated
members are package-adapter concerns; this design does not claim byte identity
for the reconstructed ZIP. A source-backed publication that observes a source
change, package signature policy failure, or candidate readback mismatch
returns before replacing the source. It never publishes a content-only patch
as though it were an archive-wide patch.

For `SourceBackedSpreadsheet`, publication captures the source version before
the focused scan, retains the exact source `content.xml` bytes used for the
read set, checks the version again after candidate readback, and invokes the
existing source/package replacement seam only while that version still
matches. If that pre-publication fence fails, the operation returns
`SourceChanged`, discards the candidate, and has not touched the destination.
Once replacement or sink writing begins, the adapter owns its atomicity: a
late source change or I/O failure is reported with the adapter's untouched or
partial-progress state, without a rollback or writer retry claim. The seam
should perform its own identity-aware atomic fence where it supports one, but
this design does not pretend that a late failure is reversible. After a
successful replacement, the returned source-backed commit carries the
accepted before/after `content.xml` bytes and the source version used to
authorize them, while the package adapter owns any ZIP reconstruction and
cleanup. A later external change after publication is a new source snapshot,
not permission to publish a stale second edit.

`SourceMetadataCommit::write_to` writes the complete accepted ODS package to a
caller-provided sequential sink through the existing bounded source-content
publication writer; the semantic patch itself remains content.xml-scoped. A
non-atomic sink can contain a prefix if I/O fails, the output limit is reached,
or cancellation arrives after bytes were emitted. The publication report must
carry the existing untouched/partial progress state and bytes written. The
API makes no all-or-nothing destination claim for such a sink; callers needing
atomic replacement must provide the package/file atomic-publication adapter.

An equal semantic edit must share the original source allocation and bytes,
skip candidate construction and readback, and still retain any malformed or
unknown source exactly. A changed cell-range-source edit must use the existing
lexical validation/refusal seam until the model can carry namespace and
attribute provenance. It must not silently turn a noncanonical
`xlink:type`, prefix, attribute order, or entity spelling into the canonical
writer form.

Owner placement is strict. A direct recognized owner under the expected ODF
parent is typed. A same-named element under a foreign element, an unknown
wrapper, or an unsupported compatibility branch is opaque and is not counted
as the effective owner. Add/replace/remove must refuse when that ignored
markup makes insertion ambiguous rather than creating a second effective
owner. Duplicate direct consolidation, label-ranges, cell-range-source, or
detective owners are malformed and block changed publication. Unrelated
comments, processing instructions, foreign siblings, whitespace, and unknown
top-level declarations remain untouched.

Foreign-namespace handling follows ODF preservation rules and does not import
OOXML Markup Compatibility semantics. The scanner does not interpret
`mc:AlternateContent`, `mc:Choice`, or `mc:Fallback`, does not select an
effective branch, and does not treat a foreign wrapper as transparent. A
recognized ODF local name is an owner only when its expanded ODF namespace and
direct parent are correct. Foreign attributes/elements are retained as opaque
source data where they are outside the changed owner; a changed operation that
would need to move, normalize, or discard them returns the typed preservation
refusal. This is not a claim that arbitrary foreign extensions are schema-valid
or editable.

Under ADR 0006, a package with an observed digital signature is read-only for
this batch: exact snapshots and exact no-op patches may be inspected, but any
changed metadata commit, content replacement, or inverse application returns a
typed signed-source refusal unless a future explicit unsign/resign policy has
been consumed. No signature part, manifest entry, or signature reference is
silently removed or regenerated here.

Staging is failure-atomic as well as publication-atomic. Each edit method
validates its selector and staged value against a cloned operation ledger
before recording it; a failed operation leaves the `Edit` ledger unchanged.
Commit then resolves all operations in deterministic owner order (singleton
owners, ordered label children, then sorted cell selectors), rejects duplicate
targets and overlapping write spans before allocating a candidate, and only
then performs one checked render/splice and one complete reopen. A failure at
any stage drops the candidate and reservations and leaves both the source
snapshot and the staged edit available unchanged. Composition joins operation
ledgers only when source identity/version and read/write sets match and all
target spans are disjoint; conflicting owner presence, repeated-run splits,
or overlapping anchors produce a structured conflict rather than
last-writer-wins behavior.

## Limits, work, and cancellation

`sheet_metadata::Limits` needs explicit input, output, depth, event/span,
namespace-context, sheet, row-run, cell-run, logical-repeat-work, label,
detective-range, detective-operation, address/function/URI/text,
aggregate-retained, scratch, patch, and cancellation budgets. The caller may
select a profile within hard structural ceilings; it cannot bypass integer,
nesting, XML, or package safety ceilings. Effective limits must be checked
before copying bytes or reserving destination capacity.

The owner must thread the existing `litchi_core::ExecutionContext` through
both the initial scan and commit rather than inventing an unobservable counter.
The concrete context contract is:

```rust,ignore
Snapshot::parse_with_context(xml, limits, &context)?;
edit.commit_with_context(&context)?;
```

The format-owned work ledger must use fixed units so a caller can set a finite
profile and reproduce a limit failure. One accepted profile is:

* scan: one unit per input byte visited, 16 units per XML event, 64 units per
  projected owner, and 8 units per selector comparison;
* mutation: 32 units per repeated-run split, 64 units per changed owner, and
  8 units per splice anchor;
* render/reopen: one unit per candidate byte emitted or reparsed and 16 units
  per candidate XML event.

All additions are checked `u64` arithmetic. The first hard profile should
reuse the existing ODS ceilings of 256 MiB `content.xml`, depth 1,024,
1,048,576 logical rows and columns, 4,194,304 logical cells, 4,096 cell
operations, and 16 MiB scalar text fields. It should additionally cap
metadata owners and each detective/label collection at 1,048,576 items and
cap cumulative work at the checked sum of the maximum scan, mutation, render,
and reopen formulas above, with an explicit hard ceiling of
`MAX_WORK_UNITS = 16_000_000_000`. A caller profile may lower the work limit or
raise a default up to these hard caps; it cannot use an unbounded or `u64::MAX`
sentinel as a production profile.

At scan start, call `context.check()` and reserve `Resource::InputBytes` for
the bounded source view. Before each event/span index entry, namespace-context
copy, typed string, or owner record is published, charge its checked size with
`context.reserve(Resource::Memory, bytes)` and its object count with
`context.consume(Resource::Objects, 1)`. Charge event bytes, selector
resolution, repeated-run splitting, validation, rendering, and candidate
reopen work through `context.consume(Resource::Work, amount)`; reserve one
`Resource::Depth` unit for each active XML stack frame and drop it on the
matching end event. Keep outstanding `Reservation`
guards alive for the corresponding source/index/staged/candidate allocation
and release them when that allocation is dropped. Use `Resource::OutputBytes`
for the checked candidate output and patch bytes. A per-owner `Limits` value
is an additional format-specific ceiling, not a substitute for the shared
budget.

Before constructing any `String`, `Vec`, namespace table, row split, or
candidate output, perform checked length arithmetic, reserve the exact or
bounded capacity, and acquire the matching context reservation. A failed
reservation must occur before the allocation and before the operation ledger
publishes a reference to it. The precharge must include both retained and
peak old-plus-new scratch capacity where a vector or hash table grows.
The `Edit` retains the context's child budget or an equivalent operation-local
ledger while staged values live, so `set_*` calls precharge their owned strings,
vectors, and operation records before recording them. `commit_with_context`
must use the same context lineage rather than starting a second root budget.
`context.check()` is required at entry, before and after every 4,096 XML events
or 64 KiB of input/output work (whichever comes first), before each
reservation, before candidate publication, and after candidate reopen. The
caller cancellation token therefore returns a typed cancellation error with
no source or destination mutation.

Retained semantic bytes and transient render/parser scratch are separate
ledger buckets. In particular, a candidate owner fragment, namespace-context
copy, row split, spliced output, and readback projection cannot be charged as
zero merely because the final XML length is within the limit. Failed edits must
release staged reservations, and cancellation during scan, validation,
render, splice, or candidate reopen leaves the source and destination package
unchanged.

The scanner should visit `content.xml` once, resolve all supported owner
parents, and build sorted physical locators. It must not rescan the complete
sheet for each selector or expand `table:number-*-repeated` into individual
objects. Logical selector work is charged explicitly, including checked
overflow and the chosen policy for a repeated run. If a row contains
unsupported markup that cannot be preserved by a focused splice, the editor
returns a bounded refusal instead of falling back to a whole-sheet
canonicalization.

No operation in this owner evaluates a formula, applies a consolidation,
recomputes detective arrows, matches labels, opens a linked document, refreshes
an external range, follows `xlink:href`, executes DDE, or renders a sheet.
`xlink:actuate="onRequest"` is retained as inert metadata.

## Verification plan

Unit and integration evidence should be added only with the implementation,
under a focused target such as
`crates/litchi-ods/tests/ods_sheet_metadata_transactions.rs`.

### Codec and parent grammar

* Parse and write valid consolidation and label-range owners at their schema
  positions, including URI/prefix aliases and encoded-equivalent values.
* Cover absent and present-empty `table:label-ranges` and `table:detective`
  separately, plus singleton consolidation and cell-range-source insertion,
  replacement, and removal. Assert the exact prelude/epilogue anchors rather
  than merely searching for a tag.
* Parse detective valid and invalid highlighted ranges, all five operation
  names, non-negative indices, ordering, empty content, direct cell and
  covered-cell parents, and duplicate/unknown attributes.
* Assert the direct cell sequence
  `cell-range-source -> office:annotation -> detective -> text-content`,
  including absent/empty slots and refusal of a source after annotation,
  detective after text, or any nested same-named owner.
* Reject duplicate direct singleton/cell owners, wrong parents, nested
  recognized names under opaque wrappers, unbound prefixes, malformed
  namespace declarations, invalid addresses, zero dimensions, and values over
  every stated bound.
* Retain comments, PIs, whitespace, foreign siblings, unrelated declarations,
  and unknown branches in exact no-op and changed-owner fixtures. A changed
  owner with data the writer cannot retain must return the typed preservation
  refusal.

### Source-bound lifecycle

* Read, add, replace, remove, clear, and inverse consolidation and label-range
  edits while preserving the schema prelude/epilogue order.
* Read, set, clear, and replace detective metadata on both ordinary and
  covered cells, including repeated row/cell runs. Verify physical-run
  splitting and that unrelated cell text, formula, cached value, style,
  annotation, source metadata, and unknown markup remain unchanged.
* Verify selector behavior at the first, interior, and last logical member of
  repeated rows and cells. A supported split must preserve prefix/target/suffix
  repeat counts; an unsupported opaque split must return the typed refusal and
  leave the edit unchanged.
* Exercise the existing `table:cell-range-source` fixture through the focused
  `CellSelector` API. Prove inert reopen and exact inverse; prove that a
  source-local namespace or noncanonical lexical fragment is preserved or
  refused, never silently normalized.
* Check exact semantic no-op byte identity, stale source/patch conflicts,
  overlapping and disjoint operations, failure atomicity, full candidate
  readback, `Spreadsheet` and `MutableSpreadsheet` forwarding, and
  `SourceBackedSpreadsheet` source-version behavior. Include signed-package
  changed-edit refusal and no-op inspection.

### Resource and package evidence

* Test exact limit boundaries and one-over-boundary cases for input, owner
  count, labels, detective children, repeated-run work, retained bytes,
  temporary scratch, output, patch, and cancellation. Include a small output
  limit with a large valid source to catch precharge-after-allocation bugs.
* Use a real `ExecutionContext` with a paired `CancellationSource` and assert
  `Resource::{Memory,InputBytes,OutputBytes,Objects,Depth,Work}` usage at
  successful, failed-reservation, and cancelled scan/commit points. Verify
  old-plus-new growth is charged before a destination Vec or hash table is
  grown.
* Run offline OpenDocument 1.4 schema validation on changed package outputs and
  verify untouched ZIP members remain unchanged except for the required
  `content.xml` publication closure.
* The bounded local scan used for this design covered 13 `.ods` files under
  `test-data/odf`, `test-data/libreoffice-core`, and `test-data/odfdo`, plus 19
  `.fods` files under those roots and the ODS test fixtures. It found no
  consolidation, label-range, detective, or cell-range-source occurrences.
  This is evidence only for the inspected corpus, not a claim about all native
  producers. Until a real producer fixture is found, use deterministic
  schema-valid synthetic fixtures and label them synthetic. Any future native
  claim must record the source path, producer/version, member names, and full
  artifact/member hashes.

## Ownership and prerequisites

The implementation can be assigned as one bounded ODS batch with these file
owners:

1. `crates/litchi-ods/src/sheet_metadata/` owns the scanner, owner index,
   source spans, `Snapshot`/`Edit`/`Commit`/`Patch`, limits, read/write sets,
   and focused package splice. It may reuse the scenario transaction's exact
   source/readback pattern but must not clone the full worksheet graph.
2. `crates/litchi-ods/src/model/detective.rs` adds the parser/validator and
   bounded read/write support for the existing detective values. The
   consolidation and label-range modules need limits-aware entry points or a
   metadata wrapper that preflights before their current allocating
   constructors/writers.
3. `crates/litchi-ods/src/worksheet/codec.rs` and `worksheet/package.rs` own
   the shared direct-cell locator, repeated-run split, namespace-context, and
   preserve-or-refuse seam. They must not duplicate `CellRange` or change the
   data-style/source-adapter owner.
4. `crates/litchi-ods/src/facade/mod.rs`, `facade/source.rs`, and
   `authoring/mutable.rs` add concise forwards. If package-level
   `document::Edit` integration is promised, its `content.xml` operation must
   compose this owner once and use the final complete reopen; it must not
   parse and publish a second independent candidate.
5. The focused integration target owns native/synthetic receipts, schema
   validation, malformed input, limits, exact inverse, and preservation tests.
   `docs/report/spec-gap-audit.md` and the feature matrix should be updated
   only after those tests and a committed implementation establish the new
   capability.

This batch deliberately does not implement scenarios, formula calculation,
label matching, consolidation execution, detective recomputation, DDE, linked
source refresh/import, external fetch, data-style vocabulary, source-adapter
lexical repairs, title/description preservation, rendering, or a detached
authoring builder. Those are separate owners or intentional inert boundaries;
they must not be hidden behind the sheet-metadata API.

The design follows [ADR 0003](../../adr/0003-snapshots-edits-and-patches.md),
[ADR 0004](../../adr/0004-semantic-api-design.md),
[ADR 0005](../../adr/0005-io-memory-and-performance.md),
[ADR 0006](../../adr/0006-validation-security-and-compatibility.md),
[ADR 0007](../../adr/0007-office-object-models.md), and
[ADR 0023](../../adr/0023-odf-family-crate-split.md). It is a design proposal,
not native interoperability approval or a claim that the current source has
these APIs.
