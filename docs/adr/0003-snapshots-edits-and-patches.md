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

IWA reference-line edits apply the same boundary at the archive adapter:
malformed existing line payloads are rejected before patching, recognized
fields are checked for canonical framing, nested custom-value unknown fields
survive scalar replacement, and repeated graph nodes are bounded before
generated protobuf materialization. Typed graph updates preserve unknown raw
fields at every modeled reference-line nesting level and validate a staged
opaque-field candidate before publishing it. Public format editors publish
only after their staged CRUD operation and typed readback succeed.

The strict reference-line readers now parse through the common source-bound
`WireView<'a>` once and expose `WireFieldView<'a>` values tied to that source.
Canonical key and length framing is checked before a recognized payload is
interpreted; mutation continues through the shared bounded patch primitives.
This keeps borrowed inspection allocation-conscious without weakening the
transactional publication boundary.

The Numbers table-header adapter applies the same publication rule to
`litchi_numbers::table::headers::{Count, Settings}`. It validates every
recognized field's canonical presence and value before staging a wire patch,
retains unknown payload fields, validates table-section capacities, reparses
the staged package, and publishes only after typed readback equals the
requested settings. The leaf error is non-exhaustive and is mapped only at the
IWA boundary. Pages and Keynote share the value type; their remaining native
model-object selectors are an explicit migration debt and do not become a
precedent for new semantic APIs.

The first concrete Keynote package transaction applies this model to slide
playback state. `Package::edit()` resolves an exact-name or checked-position
`SlideSelector` against the immutable base snapshot, accepts one bounded
skip/include operation, and returns a named `Commit` with a new package,
diagnostics, and a reversible exact-source-checked `Patch`. Equal state reuses
the original shared allocation and bytes. A real change is published only
after the one length-stable Boolean wire field is patched, the complete package
is reopened under the retained limits, and the same semantic position reads
back the requested value. The package-local patch currently retains immutable
source and target byte handles; it is reversible in memory but is not yet the
durable deterministic JSON patch envelope required for cross-process use.

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

## 2026-08-08 amendment: Keynote slide-order transactions

Keynote structural ordering uses a separate, source-compatible transaction
family. `Package::edit_slide_order()` directly returns a `SlideOrderEdit<'_>`
that accepts one `move_slide` operation. Its source is a `SlideSelector` by
exact name or checked position; its `Position` destination means the final
zero-based position in the immutable base list and must be strictly less than
the base slide count. Selection and destination validation finish before any
candidate is published. Exact-name ambiguity and either out-of-range position
are typed failures.

The staged source position identifies the same base-snapshot slide even when
removal shifts intermediate vector indexes. Moving a slide to its existing
position is an exact no-op: the commit shares the original source allocation,
touches no component, and does not perform a redundant full reparse. A real
move rewrites only the owning show component, then reopens the entire candidate
under the original `ReadOptions` and verifies the complete semantic slide
order before publication. The source snapshot never changes.

`SlideOrderPatch` records only semantic source and destination positions in its
public vocabulary while retaining exact immutable source and target bytes
privately for conflict authorization. `Package::apply_slide_order()` requires
the exact source artifact, and `inverse()` restores the accepted source bytes.
`SlideOrderCommit`, `SlideOrderDiagnostics`, `SlideOrderError`, and
`SlideOrderLimitKind` remain distinct format-owned types, so the earlier
skip-state transaction keeps its established Boolean patch accessors.
The rewrite moves complete validated raw slide-reference field records,
including each encoded key, encoded length, and nested reference payload.
Deprecated and unknown reference fields therefore travel with their slide
instead of being normalized through Buffa or Prost. This in-memory reversible
patch is still not the durable deterministic JSON envelope required for
cross-process patch exchange.

## 2026-08-08 amendment: Keynote show-settings transactions

Presentation settings are a singleton semantic value and therefore require no
object selector. `Package::show_settings()` reads that value directly, while
`Package::edit_show_settings()` stages a complete checked `Settings` value
against one immutable package. `ShowSettingsCommit`, `ShowSettingsPatch`, and
`ShowSettingsDiagnostics` remain separate from the skip-state and slide-order
families so those established patch vocabularies do not change.

An exact semantic no-op shares the original source allocation, touches zero
components, and performs no candidate reopen. This includes a package whose
root show reference is null, which reads as `Settings::default()` and has no
physical Show owner to change. A real edit of an exact package rewrites only
the component owning the unique Show payload. It then reopens the complete
candidate under the source package's retained `ReadOptions` and verifies the
requested settings through the focused reader before publication. The source
snapshot never changes.

`ShowSettingsPatch` exposes only before/after semantic settings and compact
diagnostic fingerprints. Exact immutable source and target bytes remain
private and exact byte equality, not the fingerprint, authorizes application.
`inverse()` restores the accepted source artifact byte for byte. Semantic-only
prepared sources and changed legacy nested-`Index.zip` sources are refused as
`UnsupportedSource`; an exact no-op against a physical legacy source remains
valid. The host retains the separate legacy normalization compatibility path.
Like the other current Keynote patches, this reversible patch is in-memory and
is not yet ADR 0003's durable deterministic JSON envelope.

Preservation is measured at the changed dependency closure. Untouched ZIP
members and their raw local and central records, every non-setting Show field
record, including its encoded key and length header, nested unknown Size
records, and the immutable source artifact remain exact. The edited Show
payload's effective message type and length, its
`MessageInfo`, and any enclosing length/framing bytes required to represent a
size change are part of the mutation closure and must not be described as
untouched archive metadata. Raw source records, rather than Buffa's generated
view, preserve unknown content.

## 2026-08-08 amendment: Pages section-name transactions

`Package::edit_section_name(selector)` resolves an exact-name or checked
position selector against one immutable Pages snapshot and stages one optional
producer-visible name. `None` removes native field presence while `Some("")`
retains an explicitly present empty string. NUL is rejected before staging;
duplicate destination names are permitted because native Pages permits them,
and a later exact-name selection reports typed ambiguity instead of choosing
one occurrence.

An exact semantic no-op shares the source package allocation and bytes,
touches no component, and performs no redundant candidate reopen. This remains
valid for legacy nested-`Index.zip` input. A changed edit requires exact ZIP
provenance, rewrites one selected section message, reassembles one IWA member,
reopens the whole package under the retained limits, and verifies every
published section field before returning a new immutable package. A changed
legacy source returns `UnsupportedSource` rather than silently normalizing its
physical topology.

`SectionNamePatch` exposes only the semantic position and optional before/after
names plus content-free diagnostic fingerprints. Exact private source and
target artifacts authorize application; the fingerprint is diagnostic only.
The inverse restores the accepted source artifact byte for byte, and applying
against any other exact artifact yields `PatchConflict`. This reversible patch
is intentionally in-memory and does not yet satisfy the durable deterministic
JSON envelope required for cross-process patch exchange.

## 2026-08-10 amendment: hardened Keynote show-settings transaction

This amendment supersedes the 2026-08-08 show-settings names and host-retention
claim. `litchi_keynote::Package::{show_settings, edit_show_settings,
apply_show_settings}` reads, stages, and applies the singleton value against an
immutable package. The canonical focused vocabulary is
`litchi_keynote::show::{Settings, Edit, Commit, Patch, Diagnostics, Error,
LimitKind}`; the flat `ShowSettings*` transaction names are not root aliases.
The family remains separate from skip-state and slide-order, and its public
method/type signatures expose no native identity, component/member name,
generated message, raw field, source bytes, or retained artifact accessor.

No-op and changed transaction rules remain exact-source rules. An equal edit
shares the original source allocation, reports zero components and deletions,
and skips cache inspection, reassembly, and reopen. A real edit rewrites the
unique Show-owner component and publishes only after a retained-limit full
reopen and semantic/ownership verification. Size or slide-number changes may
delete zero to three root previews; playback-only changes preserve them, and
both paths preserve slide components and slide-node caches. Exact source bytes,
not the compact diagnostic fingerprint, authorize patch application, and a
valid inverse restores the complete accepted source artifact.

The explicit Preserve policy admits reads and exact no-ops for physical legacy
nested-`Index.zip` sources but returns `show::Error::UnsupportedSource` for a
change; semantic-only prepared sources and a changed null-root Show are also
unsupported. The former normalizing fallback is deliberately deleted rather
than silently changing physical provenance. The direct host mutation surface
`KeynoteEditor::{show_settings, set_show_settings}`, its private
`editor::show_settings` module and source, `examples/edit_keynote_show.rs`, and
their direct mutation tests are removed rather than shimmed. This is not all
Show ownership: read-only `KeynoteDocument::show` still decodes a Prost
`KN.ShowArchive`, and other host creation and graph paths remain.

`show::Patch` retains complete immutable source and target artifacts behind
process-local shared allocations. Clone and inversion are `O(1)` shared-handle
operations, while equality and exact-artifact authorization can read
`O(package bytes)`. It is reversible in memory but neither compact nor durable.
It does not yet satisfy ADR 0003's versioned format-independent semantic
operation model, read/write sets, deterministic JSON/blob serialization,
composition, three-way merge, or bounded history.

## 2026-08-10 amendment: combined Numbers sheet/table name transactions

This amendment supersedes the host's separate immediate sheet- and table-name
mutations. `Package::edit_names()` opens one infallible `O(1)` batch against an
immutable base snapshot. Consuming `names::Edit::{rename_sheet, rename_table}`
calls resolve exact-name or checked-position selectors against that base, not
against names staged earlier in the batch, and reject a repeated semantic
target. `commit` validates one simultaneous final namespace: sheet names are
unique across the workbook and table names are unique within their owning
sheet. Consequently swaps and names vacated by the same batch are valid, while
an unresolved final collision fails atomically without sequential retargeting.

An empty batch or a batch whose selected names all remain equal is an exact
no-op. It shares the source allocation and skips changed-only ownership,
dependency, lock, cache-deletion, reassembly, and reopen work. A changed batch
must prove the rooted Sheet or FormBasedSheet owner and each affected
TableInfo/TableModel chain. It refuses a selected locked table, a rooted
nonempty volatile sheet/table-name dependency, and any rooted pivot owner for
an affected table. Publication rewrites each distinct touched component once,
removes the zero to three root preview members that exist, then performs a
retained-limit full reopen and semantic/locality verification. `Index` and
`ViewState` caches, unrelated entries, objects, messages, and the immutable
source remain exact.

`names::Patch` is an exact-source-checked, reversible, process-local value. It
privately retains two complete package artifacts plus its semantic/native plan;
the public value exposes content-free operation and diagnostic counts, no
authored names or source bytes. Exact package equality and semantic-before
checks, not a diagnostic fingerprint, authorize application. A valid inverse
restores the accepted source artifact byte for byte, including deleted
previews. This is deliberately not a compact or durable patch: ADR 0003's
versioned format-independent operation encoding, read/write sets,
deterministic JSON/blob serialization, composition, three-way merge, bounded
history, and durable/atomic file publication remain deferred.

Preserve mode accepts unambiguous canonical and alternate legacy flat message
encodings in place. A physical legacy nested-`Index.zip` source remains
readable and supports an exact no-op, but a changed batch returns
`names::Error::UnsupportedSource` rather than normalizing its topology. The
public host methods `NumbersEditor::{rename_sheet, rename_table}`, their direct
tests, and the raw-ID `rename_numbers_items` example are deleted rather than
shimmed. The private cross-format `rename_table_in_package` helper remains for
Pages and Keynote table creation/edit flows; its retention is not retention of
the retired public Numbers rename API.
