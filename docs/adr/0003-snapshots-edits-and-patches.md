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

## 2026-08-10 amendment: hardened Keynote slide-transition transaction

This amendment supersedes the earlier flat transaction names and host-retention
claim for slide-transition mutation. Exact-name or checked-position selection
resolves one slide against an immutable package. The consuming
`transition::Edit::{set, clear}` methods stage a complete archive-free value;
`set` requires an existing modern envelope and never synthesizes one. `clear`
retains an existing modern envelope in Keynote's no-effect representation,
including its delay, automatic-start, random-seed, and writing-direction
semantics. Clearing an already absent transition is idempotent and remains
absent.

An equal edit, including absent `clear`, is an exact no-op. It shares the source
artifact, reports zero touched components, and skips changed-only ownership,
framing, reassembly, and reopen work. A changed edit proves the rooted Document
to Show edge, the Show/SlideTree path `[3, 2]` to the selected SlideNode, and
that node's field-2 reference to the selected SlideArchive. Each nonzero edge,
component/object, and expected typed message must be unique. The rooted owner
proof uses the package's sorted object locator for indexed
`O(slides log objects)` lookup and charges aggregate node/reference work to
`LimitKind::WireWork`. Aggregate reference metadata is exact and optional
field-local metadata, when present, must match the one followed path. Strict
transition and node-marker views must agree with the selected semantic state;
one shared field/work budget governs the transition codec's complete nested
preflight rather than resetting at each descended payload.

Changed-only guards reject noncanonical selected component/object framing,
group-bearing selected payloads, merge/base/diff metadata, aliased ownership,
and mixed modern/legacy transition fields. The mutation closure contains only
the selected SlideArchive field-4 transition subtree and, when effect presence
changes, the selected SlideNode field-7 marker. Co-located owners rewrite one
component and split owners rewrite two, each exactly once. Retained-limit
candidate reopening re-proves selected ownership, requested settings, marker
agreement, and exact locality; all unselected members, objects, messages,
unknown fields, reference metadata, previews, and caches remain exact.

`transition::Patch` privately retains complete immutable source and target
artifacts, the selected semantic before/after values, and private owner
preconditions. Exact bytes and semantic/ownership preconditions, not diagnostic
fingerprints, authorize application. A no-op apply shares the source; a changed
apply reopens and verifies the retained target; replay, tamper, and the wrong
source conflict. `inverse` swaps shared artifacts and semantic preconditions so
a valid inverse-on-target restores the accepted source byte for byte.

Preserve policy keeps physical legacy nested-`Index.zip` sources readable and
exact on no-op paths, while a changed edit returns
`transition::Error::UnsupportedSource`. Legacy database-field transition
representations may be read and retained by a no-op but are not writable by the
focused transaction; a mixed modern/legacy changed owner is invalid rather
than normalized. These process-local two-artifact patches are reversible but
neither compact nor durable. ADR 0003's stable semantic operation encoding,
read/write sets, deterministic serialization, composition, merge, bounded
history, and library-owned atomic durable publication remain deferred.

The host methods `KeynoteEditor::{slide_transition, set_slide_transition,
clear_slide_transition}`, the `transition_lifecycle` module/source, their
five whole direct mutation tests, and the clear/edit/set-effect host examples
are deleted rather than shimmed. This retires direct host editing, not
transition snapshot or creation support. `KeynoteSlideInfo.transition` and
`transition_wire.rs` remain for `KeynoteEditor::slides()` aggregate decoding
and no-op validation; the separate `creation.rs::transition()` helper and the
`create_keynote_transition` workflow remain for creation.

## 2026-08-10 amendment: Numbers table-header settings transaction

This amendment supersedes the immediate mutable Numbers editor operation with
a selector-first immutable transaction. `Package::edit_table_headers` resolves
one exact-name or checked-position sheet/table pair against a base snapshot and
retains its checked semantic path plus private owner preconditions. Infallible
consuming `table::headers::transaction::Edit::set(self, Settings) -> Self`
stages one complete archive-free value; commit never mutates the source
package.

Settings equality is presence-sensitive. An equal edit, including one against
a locked table or a supported physical legacy source, is an exact no-op that
shares the source artifact, preserves root previews, reports no touched
component or deletion, and skips changed-only ownership, lock, dependency,
transaction-work, reassembly, and reopen work. A changed edit proves the
rooted Document-to-Sheet/FormBasedSheet-to-TableInfo-to-TableModel chain and
its exact aggregate plus optional matching field-local reference metadata. It
rejects ambiguous, detached, aliased, merge/diff, noncanonical, or
interactively locked selected ownership before publication.

Changed admission also fails closed with typed `UnsupportedDependency`. A
valid TableModel pivot owner in field 85 refuses every changed edit. Header-row
or header-column count changes refuse present category/haunted/category-owner
fields 81, 84, or 86, a nonempty group-by field 83, and any rooted
HeaderNameMgr. Footer changes additionally refuse active grouping decoded from
fields 81 or 86 and nonempty field 83. Selected TableInfo fields 4, 5, 7, 8,
15, and 17, or a true field 16, refuse header-count changes; fields 5, 15, and
17, or a true field 16, also refuse footer changes. Reference-bearing
dependencies must be unique local objects with exact declared metadata and may
not alias the document, selected
sheet, TableInfo, or TableModel; malformed dependency state is `InvalidSource`,
not a normalization opportunity. Repeating-header changes also refuse the
deprecated sheet-level repeating-header field 4.

Present header/footer counts retain the checked `1..=5` domain. Effective
header plus footer rows must fit the declared row count, and effective header
columns must fit the declared column count. The seven selected TableModel
fields preserve absence independently from explicit Boolean values; changed
publication neither materializes absent false values nor erases explicit false
values. Input/output, package, payload, reference, wire-output, wire
field/nesting/work, transaction-work, and allocation failures remain typed and
bounded.

An admitted changed commit conservatively preflights aggregate source work,
rewrites the selected TableModel component once, deletes the zero to three
existing root previews as the explicit rendering-cache closure, fully reopens
the candidate under retained limits, and verifies requested presence, bounds,
ownership, and byte locality. `Index/ViewState.iwa` and all unrelated ZIP/IWA
entries, objects, messages, unknown fields, and detached state remain exact.

`table::headers::transaction::Patch` privately retains exact immutable source
and target artifacts, exact selected source and target payloads for a change,
semantic before/after settings, the checked path, and private owner
preconditions. Exact bytes and selected payload state, not a diagnostic
fingerprint, authorize apply. No-op apply shares the source. Changed apply
conflict-checks the source settings and retained source payload, charges the
source topology plus distinct retained target bytes against the aggregate
transaction-work ceiling before reopening the target, and verifies the exact
retained target payload and locality. It applies only that already-verified
artifact; it does not restage or merge a semantic operation. A valid inverse
swaps both artifacts and payloads and restores the complete source including
previews.

Unique role-specific modern and legacy TableInfo/TableModel messages are
edited in place without type promotion; mixed or duplicate candidates fail
closed. The Preserve policy keeps physical legacy nested-`Index.zip` sources
readable and exact on no-op paths, while changed publication returns
`transaction::Error::UnsupportedSource` rather than normalizing the physical
topology.

Two native Numbers 14.4 oracles define the admitted closure. Changing source
SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`
to two header rows and columns preserved B2/B3 but produced SHA-256
`5c2323b509e5ea9a975b5f254bbd46cf42657aa1c3858d2c7e98f30f07e4b40c`
after changing TableModel, HeaderNameMgr, a new manager tile, and CalcEngine
formula/dependency state. That is refusal evidence, not count-parity evidence.
A Boolean-only native Save As changed TableModel field 12 from explicit true
to absent while preserving B2/B3 and all counts. Its freeze-off SHA-256
`015568e6b922e80fbfb760491dc49994ccc2218356ed197131beb46c1bd75850`
and same-state native control SHA-256
`df44ed7d0b12c1d372dad7ad7361ed1140d41967921ee42b71a4072b78615721`
regenerated semantically equivalent ViewState topology and payload with only
allocated reference identities differing. This supports the focused writer's
exact preservation of the source ViewState; it does not claim byte-identical
native Save As output.

The public host cut removes only
`NumbersEditor::{table_header_settings, set_table_header_settings}`, their
direct Numbers mutation tests, the duplicate host count test, and
`edit_numbers_table_headers.rs`. Private
`table_header_settings_in_package`/`set_table_header_settings_in_package`
bridges used by Pages/Keynote table workflows and the lower attached-table
primitives used by Numbers structural edits remain. The patch is therefore a
process-local two-artifact capability, not a compact or durable operation.
ADR 0003's stable semantic serialization, read/write sets, composition, merge,
bounded history, and library-owned atomic durable publication remain deferred.
