# ADR 0003: Snapshots, edits, patches, and concurrency

- Status: Accepted
- Date: 2026-07-31

## State model

Opened documents are immutable, cheap-to-share `Send + Sync` snapshots. Major
objects are small lifetime-free handles backed by hidden shared state; fine
traversal uses borrowed views. Borrowed zero-copy inputs produce scoped
`DocumentRef<'a>`-style types, and conversion to owned storage is explicit.

Mutation consumes or borrows a snapshot into an `Edit`. Attached document trees
are never publicly mutable. Transaction-scoped proxies expose short semantic
verbs such as `set`, `add`, `move_before`, `clear`, and `remove`. Detached
builders may be ordinary mutable values when every field combination is valid.
Identity-changing verbs such as `rename` update their modeled dependency
closure in the same transaction. They never expose a catalog-string-only mode
through the ordinary facade.

`commit()` validates the changed dependency closure atomically and returns a
named `Commit<T>` containing the new snapshot, a reversible patch, and
diagnostics. It never mutates the source snapshot. A full validation pass is
explicit. Untouched malformed or unknown content is preserved and may remain a
diagnostic, but a commit cannot worsen it.

## Identity and selection

Handles carry snapshot lineage and stable internal identity. Public lookup is
selector-first:

```rust,ignore
let sheet = book.sheet("Summary")?;
let first = book.sheet(0)?;
```

Both calls return `Result<Option<_>>`; numeric positions are checked and
zero-based. Names, semantic roles, A1/R1C1 references, and relative operations
are the main entry points. Raw physical IDs are advanced diagnostics only.

Structural insertion uses the same selectors as stable source-snapshot
anchors. `add_before` and `add_after` resolve a developer name or checked
zero-based position to semantic identity; they do not accept relationship IDs,
native sheet IDs, or part names. Base-object moves are resolved first,
before/after additions then stay attached to their anchor identity, repeated
additions retain explicit call or join order, and unanchored `add` remains the
concise tail operation. A transaction-local handle may report its current
projected position, but a later structural intent can shift it; the committed
patch is authoritative for final positions.

Serializable patches do not inject private IDs into Office files. They anchor
objects using available native identity, parent and semantic selectors, context
fingerprints, expected-state hashes, and patch-local IDs for inserted objects.
Ambiguous resolution is a conflict.

## Patches and conflicts

Every edit records semantic operations with read/write sets and yields a
versioned, format-independent patch. `Patch<Reversible>` has an inverse; a
consuming `seal()` creates `Patch<ForwardOnly>` for redaction and other
irreversible work. The canonical wire representation is deterministic JSON,
with optional compact encoding and content-addressed bounded blob bundles.

Independent `SubEdit`s may run concurrently. Join automatically merges only
provably disjoint effects. Overlap returns a structured `ConflictSet`; there is
no last-writer-wins behavior. Three-way merge produces a non-mutating plan and
commits only after every conflict is resolved.

Cross-document copying and moving use dependency-closure transfer primitives.
Move coordinates two snapshots and returns both results and patches. Resources
are reused only when equivalence is proven; name/style/theme collisions require
an explicit resolver. Undo/redo is an explicit budgeted `History<T>`.

## Deletion

`clear` removes primary payload while preserving the object. `remove` deletes
the selected object and required references. `gc` is a separate planned
reachability operation. Structural edits update unambiguous dependencies and
return a disposition requirement when content or incoming references would be
lost. Until a disposition is explicitly supplied, the safe facade returns a
typed refusal instead of guessing whether to cascade, detach, retarget, or
retain an orphan. Sanitization certification is available only when all
prohibited content is proven absent.


## iWork package transactions

The Keynote, Numbers, and Pages `Package` types apply this model to one
focused property or structure at a time. Each transaction family is a
read/edit/apply triple on `Package` (for example
`show_settings`/`edit_show_settings`/`apply_show_settings`) with its own
`Edit`, `Commit`, `Patch`, `Diagnostics`, `Error`, and `LimitKind` types, so
adding one family never changes the patch vocabulary of another. The rules in
this section apply to every family; the subsections below record only what a
family adds.

### Selection and staging

- Selectors resolve by exact name or checked zero-based position against the
  immutable base snapshot, never against values staged earlier in the same
  edit. Ambiguous names and out-of-range positions are typed failures.
- Public values carry no native identifier, component or member name,
  generated message, raw field, or source bytes.
- Staging is infallible or validates only the archive-free value; physical
  admission happens at commit. Commit never mutates the source package.
- A transaction edits existing objects only. It never creates or adopts a
  missing object unless the family says so; absence is a typed refusal.

### No-ops

An edit whose staged value equals the source value is an exact no-op. Its
commit shares the source allocation and bytes, reports zero touched components
and zero deleted previews, and skips every changed-only step: ownership proof,
dependency and lock checks, rendering invalidation, reassembly, and reopen.
Reads and no-ops therefore stay as broad as the reader, including for legacy
sources that a change would refuse.

### Changed edits

- **Exact provenance.** A change requires the complete flat ZIP artifact.
  Under the Preserve policy a physical legacy nested-`Index.zip` source stays
  readable and supports exact no-ops, but a change returns `UnsupportedSource`
  instead of normalizing its topology. A semantic-only prepared source is also
  `UnsupportedSource` for a change.
- **Ownership proof.** The changed path proves the rooted owner chain of the
  selected object (for example Document → Show → SlideTree → SlideNode →
  Slide): unique local objects, canonical selected message framing, exact
  aggregate reference metadata and, when present, matching field-local
  (`FieldInfo`) metadata, no merge/base/diff state, and no aliasing between
  roles. Zero, external, duplicate, cross-component, aliased, or ambiguous
  ownership fails closed. Malformed state is `InvalidSource`; a modeled but
  unsupported dependency is `UnsupportedDependency`. Neither is a
  normalization opportunity.
- **Raw splices.** Writers splice the complete raw records of the selected
  fields. Unknown and deprecated fields, sibling records, unselected objects
  and messages, `MessageInfo`/`FieldInfo` apart from required lengths, and
  unrelated ZIP members stay byte-exact; only derived ZIP size, checksum, and
  offset bookkeeping may change. Field presence is preserved: absent, explicit
  false, and true are distinct, and a write never materializes an absent
  default or erases an explicit one.
- **Rendering caches.** A change that can alter rendering deletes the zero to
  three canonical root previews (`preview.jpg`, `preview-micro.jpg`,
  `preview-web.jpg`) that exist, matched by exact name. A change that cannot
  alter rendering keeps them.
- **Publication.** The candidate is reassembled once and reopened completely
  under the source's retained limits. Publication happens only after the
  selected semantic value reads back as requested and locality (exactly the
  admitted delta, nothing else) is verified.

### Patches

- A patch retains the complete immutable source and target artifacts behind
  shared handles, the semantic before/after values, and the private ownership
  preconditions. Clone and `inverse` are constant-time handle swaps.
- Exact artifact bytes (or allocation identity) and the semantic
  preconditions authorize application. Fingerprints are diagnostic only.
- A no-op apply shares the source. A changed apply verifies the exact source,
  then reopens and verifies the retained target; it does not restage the
  operation. Stale, foreign, replayed, tampered, or wrong-direction
  application is `PatchConflict`.
- A valid inverse applied to the target restores the accepted source byte for
  byte, including deleted previews and caches.
- These patches are process-local and reversible but neither compact nor
  durable. The versioned format-independent operation encoding, read/write
  sets, deterministic JSON/blob serialization, composition, three-way merge,
  bounded history, and library-owned atomic file publication described above
  are not provided by them. `Package::write_to` is the output boundary.

### Codecs and budgets

- A strict handwritten wire pass validates canonical key and length framing
  before any recognized payload is interpreted, and before a borrowed lazy
  Buffa view is built over the selected fields. Strict and Buffa readings must
  agree. Raw source records, not generated messages, are the preservation and
  rewrite authority; production generated code for these projections has no
  repeated view or encoder.
- Reference-line edits follow the same rule at the archive adapter: malformed
  existing payloads are rejected before patching, nested unknown fields survive
  scalar replacement, repeated graph nodes are bounded before generated
  materialization, and a staged opaque-field candidate is validated before it
  is published.
- One transaction budget covers the codec's field, work, nesting, and
  reference reports, semantic ceilings, object and metadata structure, raw
  rewrites, component serialization, reassembly, reopen, and locality
  comparison. Work is charged before the output it pays for is allocated, so
  an allowance one below the requirement refuses with no component encode,
  output allocation, reassembly, or reopen. Scaling is checked by structural
  work counters, not wall-clock thresholds.

### Keynote slide skip state

`Package::edit()` selects one slide and stages one skip/include operation. A
change patches the one length-stable Boolean field and verifies the same
semantic position after reopen.

### Keynote slide order

`Package::{edit_slide_order, apply_slide_order}` stage one `move_slide`. The
destination is the moved slide's final zero-based position and must be less
than the base slide count; the staged source identifies the same base slide
even when removal shifts intermediate indexes. A change rewrites only the
component owning the Show and verifies the complete slide order. It moves
complete raw slide-reference records, so deprecated and unknown reference
fields travel with their slide.

### Keynote show settings

Show settings are a singleton value read by `Package::show_settings` and
edited as one complete `show::Settings`. A package whose root show reference
is null reads as `Settings::default()`; changing it is `UnsupportedSource`.
A change rewrites the unique Show-owner component. Size or slide-number
changes delete the root previews; playback-only changes keep them. Slide
components and slide-node caches are kept in both cases.

### Keynote slide transitions

`transition::Edit::{set, clear}` stage a complete archive-free value for one
slide. `set` requires an existing modern transition envelope and never
synthesizes one. `clear` keeps an existing envelope in Keynote's no-effect
representation, including delay, automatic start, random seed, and writing
direction; clearing an absent transition is an idempotent no-op.

The mutation closure is the selected Slide's field-4 transition subtree and,
when effect presence changes, the selected SlideNode field-7 marker. Strict
transition and marker views must agree with the semantic state; one
field/work budget covers the whole nested preflight. Mixed modern/legacy
transition fields are invalid. Legacy database-field transitions are readable
and survive no-ops but are not writable. Owner lookup uses the package's sorted
object locator for `O(slides log objects)` resolution.

### Keynote placeholder visibility

`Package::{slide_placeholder_visibility, edit_slide_placeholder_visibility,
apply_slide_placeholder_visibility}` select one existing
`slide::placeholder::Kind` (`Title`, `Body`, or `SlideNumber`) on one slide.
`Edit::{set, show, hide}` stage a complete `State`. A read returns `None` when
the slide has no such placeholder graph, which is not the same as hidden; an
edit then returns `Error::PlaceholderNotFound { kind }`. The transaction never
creates or adopts a placeholder.

Visibility is membership of the role reference in both the Slide's
owned-drawables list (field 7) and z-order list (field 42): once in each is
visible, in neither is hidden, anything else is invalid. Showing appends the
exact raw role-reference record after the last existing member of each list,
as Keynote does; hiding removes only that record. Other list members keep
their bytes and order. Hidden keeps the role reference, placeholder object,
text storage, and presentation.

- **Title and body.** The selected placeholder must be a unique local,
  unlocked object of the right kind whose parent is the slide. Selected build
  dependencies, direct title/body cache fields, and style-level visibility
  overrides refuse the change. A change also invalidates the selected
  SlideNode thumbnail payload and metadata and marks thumbnails dirty, in
  addition to deleting root previews.
- **Slide number.** Per-slide state is separate from the show-wide
  `show::Settings::slide_numbers_visible`, which is not changed. SlideNode
  field 18 must agree with list membership: absent or false with no
  membership, true with one membership in each list. Field 18 is patched in
  place or appended canonically. The placeholder's storage, when present, must
  be the native single attachment-character storage resolving to one
  slide-number attachment; layout visibility overrides, layering state,
  selected builds, and cross-slide aliases refuse. SlideNode caches are not
  invalidated: the node delta admits only field 18 and its message length.
  `Kind::SlideNumber` is not text: the slide-text boundary returns
  `SlideTextError::UnsupportedKind`.

SlideNode delta proofs build one payload index and one metadata index per side
rather than rescanning per identifier, and charge every `MessageInfo` and
`FieldInfo` record, including empty ones, before validation.

### Keynote soundtrack settings

`soundtrack::Settings` holds optional volume (finite, `0.0..=1.0`) and
optional `Mode` with independent presence; `Mode::Unknown(i32)` round-trips a
future discriminant and rejects a known value. An absent soundtrack reads as
`None` and an edit returns `Error::SoundtrackNotFound`; the transaction neither
creates nor deletes a soundtrack or media item.

The changed path proves Document → Show field 17 → a unique Soundtrack
object, and validates its field-3 media collection (data references, package
metadata ownership, and occurrence counts) without editing it. Media creation,
ordering, removal, and reclamation belong to the soundtrack-item operations.
A change rewrites only fields 1 and 2 in one component and keeps all previews.
Field-3 references are streamed through a bounded visitor rather than
collected.

### Keynote slide deletion

`Package::edit_slide_deletion` stages one `remove_slide`. Deleting the last
slide is refused, and there is no no-op: every accepted commit removes exactly
one slide. The changed path proves the flat Document → Show/SlideTree →
SlideNode → Slide path, exclusive inbound ownership, single-message selected
objects, and exact PackageMetadata component, UUID, external-edge, and
data-owner records. A deprecated root-node or secondary slide-list topology
refuses.

The commit removes the raw Show slide-reference record, deletes the Node and
Slide objects, and removes exactly their PackageMetadata entries
(object-to-UUID entries, the node-to-slide component reference when present,
and their data-reference owner entries). Component records, the last-object
identifier, global data-catalog records, and every `Data/` member remain.
Deletion is not `gc`: shared or uncertain media may stay present and
unreachable until a reclamation operation with its own reachability proof
exists.

### Keynote chart arrangement

`Package::{slide_chart_arrangement, edit_slide_chart_arrangement,
apply_slide_chart_arrangement}` edit the `locked` and
`constrain_proportions` flags of an existing chart selected by
`SlideSelector` and `ChartSelector`. The change cannot alter rendering, so
previews are kept. Legend visibility, persisted sort configuration, chart data,
geometry, titles, and other chart graph state are outside this transaction.

### Keynote table `Sort Now`

`Package::execute_slide_table_sort_order` and
`execute_slide_table_sort_order_to_rows` physically reorder body rows of an
existing slide table using its persisted sort order as read-only input.
`RowRange` is body-relative and half-open, so header and footer rows are never
moved. Only the canonical table-model route with proven tile, data-list,
header, UID, and empty sentinel shapes is admitted, with text, number, Boolean,
date, and duration keys. Text uses Rust lexical order, numeric, date, and
duration values use `f64::total_cmp`, mixed domains refuse, and equal keys keep
source order. Complete tile-row envelopes, sparse row headers, and row/column
UID mappings move together. Cross-tile moves, formulas and errors, rich text,
comments, merges, filters, groups, categories, pivots, spills, conditional
styles, hidden state, non-empty strokes, unknown mutable fields, and other
unproven row-affine state refuse.

### Numbers sheet and table names

`Package::edit_names()` opens one batch; `rename_sheet` and `rename_table`
resolve against the base and reject a repeated target. Commit validates one
final namespace (sheet names unique in the workbook, table names unique in
their sheet), so swaps and names vacated by the same batch are valid and a
final collision fails atomically. A change refuses a locked selected table, a
rooted non-empty volatile sheet/table-name dependency, and any rooted pivot
owner of an affected table. Each touched component is rewritten once and root
previews are deleted; `Index` and `ViewState` caches stay exact.

### Numbers table headers

`table::headers` settings (header rows, header columns, footer rows, and the
Boolean header options) are edited as one complete value. Equality is
presence-sensitive. Present counts stay within `1..=5`; header plus footer
rows must fit the row count and header columns the column count.

A pivot owner (TableModel field 85) refuses every change. Header-count changes
refuse category/grouping fields 81, 83 (non-empty), 84, and 86, any rooted
HeaderNameMgr, TableInfo fields 4, 5, 7, 8, 15, and 17, and a true field 16;
footer changes refuse active grouping (fields 81, 83, 86) and TableInfo
fields 5, 15, 16 (true), and 17. Repeating-header changes refuse the deprecated
sheet-level field 4. Native Numbers rewrites HeaderNameMgr and CalculationEngine
state when header counts change, which is why those dependencies refuse rather
than being edited. A change rewrites the TableModel component once and deletes
root previews; `Index/ViewState.iwa` stays exact. Modern and legacy
TableInfo/TableModel messages are edited in place without type promotion.

### Numbers sheet order

`Package::{edit_sheet_order, apply_sheet_order}` stage one `move_sheet`; a
second move is `OperationAlreadyStaged` and an empty commit is
`NoStagedOperation`. The read surface is the existing `Package::sheets()`.

Order has two co-located owners in `Index/Document.iwa`: Document field 1
(sheet references) and the sidebar-root TreeNode field 2 (child nodes,
positionally matched to the sheets). The writer permutes the complete raw
records in both lists and the matching subsequences of their aggregate
reference lists; record bytes and message lengths do not change. A
`FormBasedSheet` and any `FieldInfo` attribution of either order list refuse.
A change requires exactly one of each canonical root preview and deletes all
three.

### Numbers table titles

`table::title::Settings` holds TableModel field 22 (title visible) and field
37 (title outline), each with independent presence. A change refuses a locked
table and refuses any `Index/ViewState.iwa` message of the transient
table-name selection type, package-wide. Making the title visible requires a
present, finite, non-negative title height (field 33) and unique local
paragraph-style and shape-style references (fields 30 and 36); missing ones
are `UnsupportedDependency`. A hidden target does not require them. The writer
splices only fields 22 and 37 in `Index/CalculationEngine.iwa` and deletes
root previews.

### Numbers cell batches and formulas

`Package::{edit_table_cells, apply_table_cells}` stage a bounded batch of cell
changes. Staging rejects invalid addresses, out-of-bounds or duplicate
positions, count and text overflow, and failed allocation. Commit sorts the
final coordinates, elides no-ops, and plans against one final overlay. Clearing
a non-empty stored cell writes the stored-empty state; clearing a missing or
empty cell is a no-op and never erases sparse presence.

Supported changes are finite scalar writes and clears, string-list assignment
and release with exact reference counts, sparse-tile growth for non-text
scalars, in-place replacement of uniquely owned rich text, and formula
insertion, replacement, clearing, and cache refresh. The final formula set is
derived from source survivors plus the batch, so sequential authoring and
mixed scalar/formula batches keep full graph authority. Cache equality, cycle
analysis, dirty closure, and final-batch evaluation all run before one
execution authorization. New dependency tiles are written inline and tiled, as
native Numbers does.

Header-name-backed header cells, segmented string lists, shared or rich text
needing a reference transition, ambiguous rich ownership, and unsupported,
cyclic, range, deletion, or sparse formula closures are
`UnsupportedDependency`; affected merges, pivots, categories, spills, hidden
state, and conditional styles refuse by their dependency kind. A marked cycle
unrelated to the batch is preserved; a marked formula affected by the batch
refuses unless the batch removes it. A changed patch also retains the verified
source/target snapshot pair, so apply needs only locality verification and no
second reopen.

### Numbers dimension sizes

`Package::{table_dimension_size, edit_table_dimension_size,
apply_table_dimension_size}` edit one row height or column width.
`Size::Default` (absent or native zero, using the table's axis default) and
`Size::Points` are distinct even when numerically equal. Clearing a size-only
header may remove it; a header that also carries styles, hiding, cell counts,
or unknown fields is kept with the default size. A change rewrites one
`HeaderStorageBucket` component and deletes root previews. Table dimensions,
cells, formulas, merges, and row/column insertion are outside this
transaction.

### Numbers cell data formats

Number, Percentage, Currency, Scientific, and Fraction formats of one existing
cell each have a family selected by `SheetSelector`, `TableSelector`, and a
checked `CellPosition` (for example `Package::{table_cell_currency_format,
edit_table_cell_currency_format, apply_table_cell_currency_format}`). `None`
means no explicit format of that family, which is distinct from an explicit
format using automatic decimals; the transaction never invents a native
default object. A change validates the cell's format-family record, the native
format payload, format-list identity, a full reference census, and reference
counts, then applies copy-on-write. Stored values, formula caches, other cell
fields, and opaque tails are kept.

- Scientific fixes the minus-sign negative style and hides the thousands
  separator.
- Fraction admits nine `FractionAccuracy` strategies. Field 20
  (`requires_fraction_replacement`) is accepted only when absent or
  canonically `false`, and is preserved as found; `true` is rejected.

A format family never converts a cell to another family. Cell creation, values,
formulas, other format families, styles, and table topology are outside these
transactions.

### Pages section settings

`Package::edit_section_settings` edits one complete `section::Settings`: four
optional Booleans, three optional pagination values, and an optional name,
each with native presence. Names reject NUL; duplicate names are allowed, as in
Pages, and a later exact-name selection reports ambiguity. A change splices
only fields 17–22, 26, and 28 of the selected section message; layout caches
and root previews stay exact. `edit_section_name` and `edit_section_pagination`
are projections over the same patch core and do not decode or rewrite on their
own.

`edit_section_background` stages `set_solid` or `clear` and rewrites only
field 30; unsupported fills and reference-owned field-30 state refuse.

## iWork read snapshots

Each iWork crate has one archive-free semantic document type (`Presentation`,
`Workbook`, and Pages `Document`). Source-backed construction from a path,
borrowed bytes, or shared bytes completes bounded semantic projection eagerly,
then releases the physical source; `snapshot()` is a constant-time clone of
the shared state. Source-backed documents carry `metadata()` and `stats()`;
documents built from semantic values or taken from a `Package` return `None`
for both.

`Package` is the separate authority for exact bytes, physical validation, and
transaction provenance; its semantic document shares values but not source
diagnostics. Opening a package directory captures the bundle into a complete
ZIP before parsing; a semantic-only prepared source cannot become an editable
package. Semantic captures make no preservation claim for unselected metadata,
`Data/`, previews, media, or unknown sidecars.

Metadata binds to the exact canonical logical `Metadata/Properties.plist` path
(after stripping only a legacy outer-package prefix); near-names are inert.
Canonical metadata members are size- and compression-checked before they are
materialized. Keynote text visits only storages reachable in rooted
presentation order, and slide reads keep rich storage fragments. Numbers plain
text uses the validated exact output length and fails with a content-free
allocation error rather than returning partial text.

Directory capture is pinned and no-follow on Unix and version-checked on
other non-Windows targets. Package directories are refused on Windows until
reparse-safe pinned acquisition exists; byte-backed snapshots work on every
platform.
