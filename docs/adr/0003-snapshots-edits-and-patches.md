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

## 2026-08-10 amendment: Keynote title/body placeholder visibility

This amendment supersedes direct mutable title/body placeholder visibility.
`Package::edit_slide_placeholder_visibility` resolves an exact navigator name
or checked position and one semantic
`slide::placeholder::Kind::{Title, Body}` against an immutable base snapshot.
Infallible consuming `Edit::{set, show, hide}` stages one complete `State`;
commit never mutates the source package.

`Package::slide_placeholder_visibility` returns `Option<State>` because a
layout-provided role may be absent. `None` means no existing title/body
placeholder graph, not hidden. The edit entry point returns typed
`Error::PlaceholderNotFound { kind }` for that absence and never creates or
adopts a placeholder. `State::Hidden` instead retains the role reference,
placeholder object, text storage, text, and presentation while removing the
placeholder from both per-slide drawing ownership lists.

An exact equal edit shares the source artifact and preserves every preview and
cache. A changed edit proves Document-to-Show-to-SlideTree-to-SlideNode-to-
SlideArchive ownership, the selected SlideArchive title/body reference, its
unique local Placeholder owner, exact aggregate and optional matching
field-local metadata, role kind, parent slide, and unlocked state. The selected
placeholder must occur either once in both SlideArchive owned-drawable field 7
and z-order field 42 (`Visible`) or in neither (`Hidden`); disagreement,
duplicate ownership, role aliasing, merge/diff state, group-bearing or
noncanonical selected framing, and ambiguous rooted ownership fail closed.

Showing appends the exact existing raw title/body role-reference record at the
end of both ownership lists; hiding removes its single exact record from both.
All other list members retain their bytes and order. Independent native Keynote
title and body oracles established that append-at-end behavior; neither role is
inferred from the other.

The initial native gate used pristine SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
The Rust title-hidden artifact was
`df119410433b97b9993d46619764a8ffb75f257b16c0680cd54faabd9a453cdd`;
its diagnostics reported `changed=true`, two touched components, and three
deleted root previews. Keynote 14.4 opened it without warning and showed Title
off and Body on while retaining the body and date content. Native Save As,
close, and reopen produced a 475,102-byte artifact with SHA-256
`c5c996415191758b9fc638a8fdf024a912a6fe2ac4c3989970f0cb611e0670e3`
and the same UI state. Applying the exact Rust inverse restored the pristine
SHA-256. This closes title-hide interoperability and exact inverse restoration;
the independent title-show and body-show oracles close the opposite direction.

Both focused show directions also pass exact retained-artifact gates from
Apple-authored hidden sources. Apple-hidden title SHA-256
`d61a92b212d8a0f001bdfc24490d846e065b96885f0d0d0b86ef0be9f10e7580`
produced Rust-shown
`3d36d31c6222b7622cab180f6dd9559ccf43f4b481e6b245c9d2c56fe8852b2c`
and its inverse restored the exact Apple-hidden title source. Apple-hidden body
`05ca9617ea5a23c57252c28c3029af96d4ec54345de331571d89b612566b8416`
produced Rust-shown
`3e8855e954c16bd32350e057665b5ee4758a02e85ad23c3c6543f1caef177b13`
and its inverse restored the exact Apple-hidden body source. Each forward show
reported `changed=true`, two touched components, and three deleted root
previews. Together with the native oracles, these gates establish both show
directions and exact reversibility over Apple-authored hidden artifacts.

The focused mutation refuses selected build dependencies, direct slide
title/body cache fields, and style-level title/body visibility overrides rather
than rewriting broader layout or animation state. An admitted changed commit
rewrites the SlideArchive ownership lists and invalidates rendering: it removes
the zero to three root previews and clears the selected SlideNode thumbnail
payload/metadata while marking thumbnails dirty. Co-located owners rewrite one
component and split owners two, each once; candidate reopening verifies the
requested state, invalidated render caches, rooted ownership, and exact
locality. Its direction-aware SlideNode exact-delta proof builds one linear
payload occurrence/kind index and one metadata declaration index per side,
rather than rescanning payload and metadata for every identifier. Doubling the
distinct-reference regression from 4,096 to 8,192 consumes no more than 2.3
times the measured work. Work also charges every `MessageInfo` and every
`FieldInfo` plus its path, including structurally empty records. The 4,096
empty-`FieldInfo` regression proves that zero-work and payload-only allowances
reject conditional mutation and exact-equality verification without changing
the source node. Full structural precharge covers selected and nonselected
message payload bytes; each `MessageInfo` record including its scalar/base
state, version/diff vectors, diff and removal paths; every `FieldInfo` record,
path, version vector, and feature identifier; and one unit of both Work and
References for each aggregate or `FieldInfo` reference occurrence. Conditional
invalidation performs that precharge before broad validation and additionally
charges `header_length` before the core replacement walk; exact verification
precharges the same complete structure for both source and candidate before
archive equality. The 256 KiB sibling payload with 2,048-element
reference/vector metadata regression proves that low Work and References
allowances reject both exact verification and conditional mutation atomically.
Conditional invalidation scans the selected payload once, reuses that scan for
any rewrite, and exact verification compares source and candidate without
cloning a node or rerunning invalidation. Both receive the transaction's
remaining work/reference allowance and merge their exact reports into the same
budget. The slide ownership router separately charges exact
`6 * source.len() + output_len + 2 * parsed_fields` work before allocating
output, so exhausted transaction work cannot fail only after that allocation.

Placeholder text/storage, the other placeholder role, slide-number state,
layout/style identity, builds, transitions, notes, and unrelated
members/objects/messages/unknown fields remain exact.

`slide::placeholder::Patch` privately retains complete exact source and target
artifacts, semantic before/after state, the checked position and role, and
private placeholder/owner preconditions. Exact artifact bytes and selected
semantic/ownership state, not diagnostic fingerprints, authorize apply. A
no-op apply shares the source; a changed apply reopens and verifies the retained
target and cache direction; replay, tamper, and the wrong source conflict.
`inverse` swaps shared artifacts and semantic/cache direction so a valid
inverse-on-target restores the complete accepted source including SlideNode
caches and root previews.

Preserve policy keeps physical legacy nested-`Index.zip` sources readable and
exact on no-op paths, while a changed edit is `Error::UnsupportedSource` rather
than physical normalization. These complete-source/target patches remain
process-local, reversible, and non-durable; stable semantic serialization,
read/write sets, composition, merge, bounded history, and library-owned atomic
durable publication remain deferred.

The completed host cut removed
`KeynoteEditor::{set_slide_text_placeholder_visible, set_slide_title_visible,
set_slide_body_visible}`, the public `KeynoteSlideTextPlaceholder` type, the
150-line `placeholder_visibility` module/source, two whole direct mutation
tests, their exclusive `TEST_TITLE_PLACEHOLDER_FIELD` constant, and the 30-line
`set_keynote_placeholder_visibility.rs` example. Mixed layout reads now use
the canonical `slide::placeholder::Kind`; the focused `SlideTextRole`
discriminator was consolidated into that same type rather than retained as a
second title/body enum. Host `KeynoteSlideTextRole` remains for aggregate
title/body/text-box/shape reads and is not a transaction alias. The cut retains
`KeynoteSlideInfo::{is_title_visible, is_body_visible}` snapshot reads and the
private ownership substrate still used by aggregate title/body slide decoding
and layout changes. This title/body slice does not retire or claim semantic
ownership of `set_slide_layout`, layout placeholder creation/adoption, or
slide-number placeholders; the separate per-slide slide-number handoff below
retires only its direct host mutator.

The title/body cut gate passed 94/94 focused library tests, 18/18 preview tests,
5/5 placeholder-visibility tests, 25/25 slide-text tests, 8/8 facade tests,
7/7 doctests, and 129/129 boundary-checker tests. The all-target check, strict
library Clippy, strict rustdoc, bounded fuzz smoke, exact patch/inverse gates,
and native Keynote interoperability also passed.

## 2026-08-11 amendment: Keynote per-slide slide-number visibility

This amendment extends the existing exact-source placeholder-visibility
transaction to `slide::placeholder::Kind::SlideNumber`; it supersedes the
title/body amendment's earlier slide-number exclusion without adding another
edit, patch, commit, diagnostics, error, or limit family. The same
`Package::{slide_placeholder_visibility, edit_slide_placeholder_visibility,
apply_slide_placeholder_visibility}` methods and consuming
`Edit::{set, show, hide}` operations select one existing role on one rooted
slide. A missing SlideArchive field 20 reads as `None` when the node is also
hidden, while edit returns `Error::PlaceholderNotFound { kind }`; this
transaction never creates a slide-number graph.

Per-slide state is separate from the show-wide setting. The transaction owns
SlideNode field 18 and the selected field-20 reference's membership in
SlideArchive owned-drawables field 7 and z-order field 42. It preserves
KeynoteShow field 6 and therefore does not change
`show::Settings::slide_numbers_visible`. `Kind::SlideNumber` is also not text:
the slide-text read/edit boundary returns
`SlideTextError::UnsupportedKind { kind: Kind::SlideNumber }` rather than
exposing or rewriting the native attachment character.

The semantic state requires all three facts to agree. Hidden is absent or
false SlideNode field 18 and zero occurrences of the field-20 reference in
both lists; visible is true field 18 and exactly one occurrence in each list.
Duplicate or one-sided list membership, a true node without field 20, or any
node/list disagreement is `Error::InvalidSource`. Showing copies the exact
field-20 reference payload after the final existing field-7 member and after
the final existing field-42 member, matching the native per-list insertion
rule; a source with no insertion anchor for either list refuses rather than
inventing an ordering. Hiding removes only the selected membership records.
Node field 18 is patched in place when present or appended canonically when
absent. Exact artifacts, not normalization, let inverse restore an absent
versus explicit-false source representation byte-for-byte.

Changed admission proves the rooted Document-to-Show-to-SlideTree-to-SlideNode-
to-SlideArchive chain, unique package-wide ownership of the selected field-20
identifier by that slide, exact aggregate and field-path metadata, canonical
framing, and slide/placeholder co-location. The selected type-7 placeholder
must have native kind 1, the selected slide as parent, and unlocked state; its
modern/deprecated storage references must agree. When nonzero storage exists,
it is a local type-2001 storage with kind absent or 3, `in_document=true`, one
object-replacement character, and exactly one character-zero attachment-table
entry. That entry resolves to a local type-2043 slide-number attachment whose
kind is absent or zero and whose string equivalent is absent or empty. The
storage stylesheet/attribute-table dependencies, attachment, placeholder
presentation dependencies, other slide roles, and relevant node references
must be unique, non-aliased, and exactly declared in aggregate and field-local
metadata. Merge/base/diff state, groups, noncanonical framing, layout
visibility override field 6, layering/cached slide state, selected builds,
cross-slide aliases, or ambiguous closure fail closed.

A changed commit rewrites the selected SlideArchive and SlideNode in one
component when co-located or two components when split, then deletes the zero
to three root previews and reopens the complete candidate. Unlike title/body
visibility, it does not invalidate SlideNode thumbnail payloads or metadata:
the exact node delta admits only field 18 and its enclosing message length, so
all existing node caches remain byte-exact. The exact Slide delta admits only
the role-membership splice and message length. `Index/ViewState.iwa`,
KeynoteShow field 6, placeholder/storage/attachment content, other roles, and
all unrelated members, objects, messages, metadata, and unknown fields remain
exact. No-op commit/application shares the source and skips changed-only
guards; changed apply verifies the retained exact source and target deltas;
inverse swaps the process-local complete artifacts and restores the complete
source including deleted previews.

The private Buffa closure projects only SlideNode field 18, storage fields
1/9/10, and the textual attachment super fields 1/2. The repeated attachment
table remains borrowed opaque bytes and is validated by the bounded strict
router, so generated production code contains no repeated view or encoder.
Strict and Buffa snapshots must agree; typed byte/field/work/nesting limits are
enforced, and successful exact field/work reports consume the transaction's
remaining aggregate budget. The generated closure is five files and 112,101
bytes under a 116-KiB cap, with aggregate SHA-256
`eacce4103b5c9f9f32fd98639b81249ae1d15fcd63da6fe636569e0a2a324c30`.
The scalar node writer and its direction-aware verifier use the same bounded
full-structure precharge as the title/body cache verifier; the common slide
router precharges
`6 * source.len() + output_len + 2 * parsed_fields` before output allocation.

The exact native gate began from the 500,058-byte source SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
Rust show produced a 455,859-byte artifact with SHA-256
`a2dafcd4ffc57bafc3bbf7d7cd4ee8131bab2c06dd52adc292632d4208c126be`,
reported `changed=true`, two touched components, and three deleted previews,
and its inverse restored the exact source. Keynote 14.4 (7043.0.93) opened the
candidate without warning, showed Slide Number checked and canvas number 1,
and retained title, body, and date. Save As, close, and reopen preserved that
state in a 500,192-byte artifact with SHA-256
`b1edd073d309157d27508baf4aedbe93d6dee0687f727dd71f1e8232f6171882`;
Apple regenerated previews, which is semantic compatibility rather than raw
native-cache equality; cached Data9074 remained byte-exact at SHA-256
`575645e2455199d7cc0c65fab8002b9e025765ba19b8b03c6e51c000f4915e89`.

Apple-only controls independently changed hidden 500,024-byte SHA-256
`9a456ccda73da47e81f0781fc831482da30938d043986c54bea43394cd2ad5e9`
to visible 500,191-byte
`5f4d4dbe1264446107649342b0b29587e7054282c238b45f42b2ad9b6d65fa5b`,
then visible resave 500,192-byte
`01749e2ed0e963e35cb7bc77f8f26cf60df3def59468becfca3169a3f73e2774`
and rehidden 500,006-byte
`d70e365a2784fdb927e1772978a9917094973f9963a06cc70aa3b5914b6eb499`.
The selected native delta changed only field 18 in the SlideNode and inserted
the exact field-20 reference payload at the end of each field-7/42 list;
SlideNode caches, KeynoteShow field 6, and the placeholder/storage/attachment
closure stayed exact.

The completed host cut removes the one public
`KeynoteEditor::set_slide_number_visible` method with its whole 172-line
`keynote/editor/slide_number.rs` source and module declaration, two whole
direct mutation tests, their four exclusive constants, and the
`test_package_with_slide_number` fixture helper, plus the 23-line
`set_keynote_slide_number.rs` example. The retained source-free
`create_keynote_slide_numbers.rs` example now hands off to the focused
transaction. `KeynoteSlideInfo::is_slide_number_visible` remains a snapshot
read; `KeynoteDocumentBuilder::slide_number_visible`, the private creation
module, add-slide materialization, `set_slide_layout`, aggregate title/body
reads, and the shared host ownership helper remain. This cut transfers only
direct mutation of an existing per-slide role; it does not retire creation,
layout, snapshot, show-wide settings, or slide-number graph construction.

The patch remains a process-local two-artifact capability with no stable
semantic serialization, read/write sets, composition, merge, bounded history,
or library-owned atomic durable publication. Changed physical nested
`Index.zip` sources still return `Error::UnsupportedSource`; exact reads and
no-ops preserve them. The frozen handoff passed 8/8 codec tests, 98/98 Keynote
library tests, 7/7 focused placeholder-visibility tests, 9/9 facade tests,
22/22 preview tests, and 7/7 doctests, together with the all-target check,
strict library Clippy and rustdoc, bounded fuzz smoke, exact patch/inverse
gates, and native interoperability. The boundary unit suite passed 138/138;
the live slide-number host, placeholder host, and focused boundary audits were
empty, while the full checker retained only the unchanged 14 dependency-policy
baselines.

## 2026-08-11 amendment: Keynote soundtrack settings

The focused immutable transaction is rooted directly in
`litchi_keynote::soundtrack::{Mode, Settings, Edit, Patch, Commit, Diagnostics,
Error, LimitKind}`. `Package::{soundtrack_settings, edit_soundtrack_settings,
apply_soundtrack_settings}` read, stage, commit, and apply one presentation's
existing playback settings without exposing a native identifier, component
name, reference payload, or source bytes. `Edit::settings` returns the staged
value; consuming, infallible `Edit::set` replaces it, and consuming
`Edit::commit` publishes only after exact verification. An absent soundtrack
reads as `Ok(None)` and edit returns `Error::SoundtrackNotFound`; an existing
soundtrack with both scalar fields absent reads as `Some(Settings::default())`.
The transaction neither creates nor deletes a soundtrack object or media item.

`Settings` preserves the independent presence of optional volume and mode.
Volume must be finite and in `0.0..=1.0`; `Mode::{PlayOnce, Loop, DoNotPlay}`
are the canonical known values, while `Mode::Unknown(i32)` round-trips a
genuinely future discriminant and rejects a known value hidden inside
`Unknown`. `Settings::{set_volume, set_mode}` validate before changing the
value. Their semantic validation errors remain at the crate boundary; the
transaction's `soundtrack::Error` separately reports absent/invalid/unsupported
sources, typed resource and allocation failures, verification failure, and
exact-source patch conflict without leaking content.

Selection proves one rooted Document object 1 field 2 reference to the Show,
then one nonexternal Show type-2 field 17 reference to a unique type-21
Soundtrack object. Soundtrack field 1 is optional fixed64 volume and field 2 is
optional varint mode. Changed admission additionally proves exact aggregate
and any field-local object-reference metadata for the Document/Show route,
canonical selected-message framing, no merge/base/diff state, disjoint use of
the soundtrack identifier in other known Show and slide-tree reference roles,
and a valid selected-component media closure. Physical nested `Index.zip`
sources remain readable and support exact no-op commits, but a change returns
`Error::UnsupportedSource` under Preserve.

Soundtrack field 3 remains the ordered media collection and is not an editable
settings field. Each strict field-3 payload must be a canonical nonzero local
data reference whose order and multiplicity equal the selected message's
aggregate and any field-3 data-reference metadata. For populated media,
the package-metadata component must identify the selected component, each data
reference must name a rooted data record and one package member, and the
soundtrack owner's recorded occurrence count must agree. These proofs admit a
settings rewrite; they do not transfer ownership of soundtrack media creation,
replacement, ordering, removal, or resource garbage collection.

A changed commit rewrites exactly the component containing the selected
Soundtrack object and reports one touched component. Present fields 1/2 are
replaced in place, cleared values remove only their selected records, and
newly present values are appended canonically in field order. Field 3, unknown
fields, every other message and object in that component, message metadata
apart from the selected payload length, and all other package members remain
exact; only necessary ZIP checksum, size, and offset bookkeeping may differ.
No rendering cache is invalidated and all root previews are retained. Complete
candidate reopen checks the requested settings, rooted source contract, exact
field-1/2 delta, media closure, and package locality before returning.

No-op commit and no-op apply share the immutable source, report zero touched
components, and skip changed-only physical admission and reopen. Changed apply
first authorizes the exact retained source artifact and prior semantic value,
then reopens and verifies the retained target once. `Patch::inverse` swaps the
complete process-local source and target artifacts in constant shared-handle
work, so inverse restores all original field presence, unknown bytes, ZIP
records, and package members exactly. The patch has no stable semantic
serialization, composition, merge, read/write sets, bounded history, or
library-owned atomic durable publication; `Package::write_to` remains the
bounded output boundary.

The private Buffa projection contains only Soundtrack fields 1/2. A strict
handwritten pass validates those scalars, streams field-3 data references
through a bounded visitor without retaining an input-width vector, and must
agree with the Buffa scalar snapshot. Production generated code exposes no
repeated view or encoder. The deterministic closure is five generated files
and 27,753 bytes under a 32-KiB cap, with aggregate SHA-256
`458206e0b57d8ec5ae4c3fc706bf793ccd385ab867b7e92ac30d66ab1858b4d3`.
Typed byte/field/work/nesting and semantic-reference ceilings are enforced;
the codec's successful field/work report and streamed media counts contribute
to the transaction budget. The final performance review found no P0 or P1
issue. A test-only scaling gate in the media validator exercises realistic
4,096- and 8,192-record metadata/media states through the real streaming path:
references double exactly, while measured fields, work, and references each
remain within 2.3 times the smaller case. This structural gate uses no
wall-clock threshold and changes no production path.

The populated native gate began from the Apple-resaved 506,640-byte source
SHA-256
`69795554212651b261f5ffd71dd5cf511544f285cab680d724a9de7d3f04b14d`.
Rust changed it to Loop with volume `0.35`, producing a 506,640-byte candidate
SHA-256
`6367e38a2edeebe6e65b148d0fd2aae555ee219dc1a65c339954047eb533ce1a`;
only `Index/Document.iwa` differed and inverse restored the exact source.
Keynote opened the candidate without warning, reported Loop and volume
`0.3499999940395355` in the Audio UI, retained the one-second `ringin` item,
and played it. Save As produced a 506,651-byte artifact SHA-256
`e264f4e714b0c44fca420b2c7b43e18f2ed1be99a766d25fe901f68d5f8bc299`.
Every root preview and `Data/ringin-9075.m4a` stayed exact; the media SHA-256
was `5a08f48c4f86074e14a763d4f19f49ca31196a7a5f52fb48960e76b6f3d3d96b`.
Reopening the native-resaved artifact and restaging its normalized volume was
an exact no-op.

The completed host cut retires the two direct
`KeynoteEditor::{soundtrack_settings, set_soundtrack_settings}` methods and the
whole 68-line `keynote/editor/soundtrack.rs` settings module. The production
host diff is exactly two insertions and 91 deletions, including removal of the
settings-only parts of the shared wire module and its module declaration. It
also deletes 157 lines of direct host settings tests and the complete 29-line
`edit_keynote_soundtrack.rs` legacy example. The structure inspector and host
README migrate settings reads and examples to the focused `Package` API rather
than retaining a compatibility shim.

The ordered `KeynoteEditor` soundtrack-item
read/add/insert/replace/move/remove CRUD, `KeynoteSoundtrackItemInfo`, the
media-item example, soundtrack creation, and the shared wire lookup,
media-reference, metadata-repair, and replacement substrate remain. This cut
retires settings ownership only; it does not move or delete media-item CRUD.
The frozen gates passed 5/5 codec tests, the 1/1 focused scaling unit gate, 4/4
focused soundtrack-settings integration tests, 99/99 Keynote library tests,
10/10 facade tests, and 8/8 doctests. Formatting, strict library checks,
all-target checks, example builds, the live host audit, and the final diff
check also passed. The boundary unit suite passed 152/152; the live host and
focused audits were empty, while the full checker retained only the unchanged
14 baselines: six development-only annotations and eight edge classifications.

## 2026-08-11 amendment: Numbers sheet order

The focused immutable transaction is rooted directly in
`litchi_numbers::sheet::order::{Edit, Patch, Commit, Diagnostics, Error,
LimitKind}`. `Package::edit_sheet_order()` is an infallible borrow-only entry;
consuming `Edit::move_sheet(selector, destination)` stages exactly one move and
consuming `Edit::commit` publishes it. `Package::apply_sheet_order` applies an
exact retained patch. Existing semantic `Package::sheets()` and Document sheet
iteration remain the read surface, so this mutation family adds no duplicate
order reader, native ID, archive/component handle, protobuf value, or source
bytes.

Selectors resolve against the immutable base semantic document by exact name
or checked zero-based position. The destination is the moved sheet's final
zero-based position after removal and must be less than the base sheet count.
A second move returns `Error::OperationAlreadyStaged`; committing an empty edit
returns `Error::NoStagedOperation`. Moving a sheet to its current position is
an exact positional no-op: the source package and exact artifact are shared,
diagnostics report unchanged, zero touched components, zero deleted previews,
and no full candidate reopen. This fast path does not resolve the native order
owners.

A change proves two co-located order owners in `Index/Document.iwa`. Root
Document object 1/type 1 owns the ordered sheet-reference records at field 1
and a required nonexternal sidebar-root reference at field 5. That reference
resolves to a unique type-205 sidebar-root TreeNode whose field 2 is the
ordered child-node sequence and whose field 3 must be absent. Each child node
must be local and positionally associate its field-3 object reference with the
corresponding Document field-1 sheet. Child identifiers, descendant nodes,
and all sheet/table subgraphs are retained, unique across their roles, and
co-located with both owners. Current changed support admits one ordinary
`TN.Sheet` message per rooted sheet; a `FormBasedSheet` returns
`Error::UnsupportedSource` because its native reorder contract is not yet
proven.

Both owner messages must use canonical selected framing without merge/base/diff
state. Their MessageInfo aggregate object-reference lists may contain unrelated
roles, but the selected Document sheet-reference subsequence and sidebar child-
reference subsequence must each occur exactly once and in payload order. Any
FieldInfo attribution of either selected order subsequence is refused as
`UnsupportedSource`; the writer does not guess how a field-attributed order
should be restated. Field-local declarations for the non-order Document field
5, child field 3, and descendant field 2 may be absent, but when present must
use the exact accepted path and reference. Zero, external, duplicate,
cross-component, role-aliased, mismatched, or ambiguously owned references fail
closed.

The raw writer reorders the complete existing Document field-1 records and
sidebar-root field-2 records; it never decodes and rebuilds a selected
reference. All unknown/deprecated fields inside each reference therefore move
with that exact record. It reorders only the corresponding selected
subsequences in the two MessageInfo aggregate lists, leaving unrelated
aggregate references and every FieldInfo byte exact. Because the moved records
are byte-identical, both payload lengths and message lengths remain unchanged.
Every child node, sheet object, table and data sidecar, ViewState,
CalculationEngine state, unknown field, sibling object/message, and unrelated
package member remains exact apart from necessary ZIP offsets. The change
is admitted only when the source contains exactly one each of the three
canonical root previews: `preview.jpg`, `preview-micro.jpg`, and
`preview-web.jpg`; a changed edit refuses a missing or non-unique member. It
rewrites one component, deletes all three previews, reassembles once, and fully
reopens one candidate before exact locality verification.

The patch privately retains complete process-local source and target artifacts,
the semantic positions, the moved-sheet identity, and directional compact
native/reopen proofs. Source and target fingerprints are diagnostic only;
authorization uses allocation identity or complete byte equality. Changed
apply verifies the exact source, retained target, moved-sheet position, dual-
owner order, complete locality, and the exact three-to-zero preview direction.
`Patch::inverse` swaps artifacts, positions, preview counts, and native proofs
in constant shared-handle work, verifies the reverse zero-to-three direction,
and restores the accepted source byte-for-byte including its three previews.
Stale, replayed, cross-base, byte-different, or directionally inconsistent
application conflicts.
The patch has no stable serialization, composition, merge, bounded history, or
library-owned durable publication; output remains `Package::write_to`.
Physical nested legacy sources retain exact no-op behavior, while a changed
edit is `UnsupportedSource` under Preserve.

The private projection is exactly
`TNNumbersSheetReferenceArchive.proto`, containing only the three scalar fields
of one `TSP.Reference`: required identifier and optional deprecated type and
external marker. A handwritten two-pass strict router owns Document fields 1/5
and TreeNode fields 2/3, canonical framing, exact reservation, and aggregate
byte/field/work/reference accounting; every selected scalar reference is
forced through a borrowed Buffa parity check. The raw writer is handwritten,
and generated production code contains no repeated view or encoder. The frozen
closure is five files and 32,579 bytes under a 33-KiB cap, with SHA-256
`2a0850fd82cfbf337ed48e582d4a998bd27e5046eb63c61f6939fa5ff1a09854`.

One transaction budget merges exact strict-codec reports with semantic sheet
and reference ceilings, object/message/MessageInfo/FieldInfo structure,
indexed reference lookup, raw and aggregate reorder work, component
decompression/allocation/serialization, package reassembly, target reopen, and
exact locality comparison. The 4,096-to-8,192 regression exercises the strict
codec, raw record reorder, and core aggregate-reference reorder; measured
work, references, and payload remain within 2.3 times the smaller case plus a
fixed 32-unit allowance, with no wall-clock assertion. Exact max-minus-one
reference allowance rejects in preflight. Performance review found no P0/P1
issue and no quadratic sheet scan.

The accepted bounded P2 costs are explicit: a changed patch retains roughly
four reference snapshots per sheet, bounded by the 4,096-sheet semantic cap;
Vec-to-Arc target conversion may transiently duplicate the target bytes; and
authorization of an equal but separately allocated source may perform one
bounded full-package comparison before transaction charging, while the usual
allocation-identity path is constant time. Native-unproven `FormBasedSheet`
support also remains fail-closed P2 debt rather than an inferred writer right.

The native oracle began from the 133,740-byte Apple-authored source SHA-256
`781181e89c655da5c92b677b9ba5c939c85379e7b33ccf10e3846fe8588f9c5b`;
Rust inverse restored that artifact exactly. Matched Apple control and reorder
artifacts were respectively 133,594-byte SHA-256
`f9c5cbec4f422484c63d1d39bd8d09da122d011596561a5feb2ad1e812574990`
and 153,498-byte SHA-256
`7b3bcbc853346a433e84ee815d28671d01fc3da857e43b8b7d29b310f94e7e1a`.
They established the same Document field-1 and sidebar-root field-2
permutation plus matching aggregate subsequences while leaving child-to-sheet
and table topology intact. The no-op control preserved all previews; native
reorder regenerated `preview.jpg` as 47,868-byte SHA-256
`db372ed754b8702fb964760f5087cedb2b2cfac09ff2d898947458822446c1f6`,
`preview-micro.jpg` as 1,177-byte
`582e37b9fddd5e669e1929d64f54e31da3c2c22f13cbd0df1e74dfad34543f5e`,
and `preview-web.jpg` as 7,217-byte
`6c7a226b0a64d5946cabbc517c5b416a677e23871a7ffd040fb4f225b1ac339d`.

The Rust moved artifact SHA-256
`97c76894503a2628c1828babd93d9a9a891794d86c86177cab60f09333997a68`
opened in Numbers 14.4 without warning, repair, or conversion. The UI showed
tabs `FirstCreated`, `SecondCreated` in the requested order and retained the
`A-new`, `A-old`, and `B-only` cell markers. Save As, close, and exact-path
reopen produced a 103-member artifact SHA-256
`4aa257e4db61a3c03950360b29267c9495985d460ae22b6f679bee31f2693217`
with the same UI state. Its three previews matched the Apple-native reorder
hashes above exactly. Fresh tree/ViewState identifiers, revision/document
identity, metadata and physical-object ordering, and CalculationEngine field-14
cache culling observed in Apple Save As are native normalization, not required
minimal Rust deltas; warning-free native acceptance proves preservation of the
existing IDs and caches is benign. Restaging the same position on the
native-resaved artifact reported unchanged, zero touched components, zero
deleted previews, no reopen, and byte-exact output and inverse.

The production gates passed 7/7 codec tests, 132/132 proto tests, 109/109
Numbers library tests, 4/4 private transaction tests, and 1/1 public focused
integration test, plus all-target/all-feature checks, formatting, and diff
checks. Sheet-order code has no strict-Clippy finding; remaining whole-crate
strict diagnostics are unrelated legacy baseline. The migration changes no
package, feature, or dependency edge: topology remains 64 workspace packages,
235 internal declarations, 14 host declarations, and 14 ordered debts.

Host retirement is scoped only to direct `NumbersEditor::move_sheet`, its
exclusive `selectors::sheet_index`, legacy example, direct/mixed move coverage,
and README call. That cut removes exactly 58 production lines, changes host
tests by +2/-43, and deletes the 23-line move example; the retained remove-sheet
example's selector migration is +2/-6. Existing sheet reads,
add/duplicate/remove CRUD, table moves, drawable layering, creation, and the
shared `update_numbers_document` writer remain; none is an order shim. The
boundary suite passes 165/165, the live-host and focused audits are empty, and
the full checker retains only the unchanged 14 baselines: six development-only
annotations and eight edge classifications. Its sheet-order inventory is the
private error/resolve/rewrite helper tuple and five source files.

## 2026-08-11 amendment: Numbers table-title exact patches

The focused immutable owner is
`litchi_numbers::table::title::{Settings, Edit, Patch, Commit, Diagnostics,
Error, LimitKind, Path}` with
`Package::{table_title_settings, edit_table_title, apply_table_title}`. A
semantic sheet selector followed by a sheet-scoped table selector replaces the
former raw model identifier. `Settings` preserves the independent presence of
TableModel field 22 (title visibility) and field 37 (title outline): absent,
explicit false, and true are distinct transaction values. That nine-state Rust
contract is not a native-oracle claim for explicit false or outline behavior.

An edit is bound to one immutable source snapshot and exposes the selected
semantic path and staged settings. Consuming `Edit::set(self, Settings) ->
Self` only replaces the staged value. If it equals the source value, commit
shares the exact source artifact, reports unchanged with zero touched
components and previews, and skips all changed-only ownership, rendering, and
candidate-reopen work. No-op apply and inverse retain the same exact behavior.

A change first requires exact flat package provenance. It reuses the table-
header owner's rooted Document field-1 -> Sheet or FormBasedSheet drawable ->
TableInfo field-2 -> TableModel proof, including unique local objects,
canonical selected messages, exact aggregate and optional FieldInfo reference
metadata, nonaliased roles, and the effective TableInfo lock. Canonical and
unambiguous legacy TableInfo/TableModel message types remain readable through
that shared resolver; a changed nested/non-exact physical source returns
`UnsupportedSource` under Preserve rather than being normalized. A locked
selected table returns `TableLocked`. Changed admission also scans every
message in `Index/ViewState.iwa` and returns `UnsupportedSource` if any message
has native type 6284, the transient table-title selection state. The scan is
conservative and package-wide; it does not infer that a type-6284 occurrence
belongs to the selected table.

When the requested target is effectively visible, changed admission also
requires the source's field-33 title height to be present, finite, and
nonnegative. Field-30 paragraph-style and field-36 shape-style references must
be present, local, distinct from each other and the rooted table roles, occur
exactly once in aggregate metadata, and use only their exact optional
FieldInfo paths. They must resolve uniquely to canonical type-2022 paragraph
style and type-2025 shape style messages with valid required-super framing.
Missing rendering state returns `UnsupportedDependency`; external, aliased,
ambiguous, or malformed state is invalid. A hidden requested target does not
invent or require those rendering dependencies.

The raw writer splices only fields 22 and 37 of the selected TableModel and
retains each field's requested presence. It rewrites one
`Index/CalculationEngine.iwa` component, removes every canonical root preview
present in the source (zero through three), reassembles once, fully reopens the
candidate, and verifies semantic readback plus exact locality. Field 33, fields
30/36, all unknown records, MessageInfo and FieldInfo apart from the necessary
selected payload length, every nonselected object/component, ViewState, table
data and styles, and unrelated package members remain exact; only derived ZIP
size, checksum, and offset bookkeeping may differ. Reads and exact semantic
no-ops do not run the changed-only type-6284 guard, so those operations remain
broad. For an accepted changed source with no type-6284 message, all other
ViewState content remains byte-exact.

`Patch` retains the complete exact source and target artifacts, before/after
settings, selected semantic target, directional preview counts, and bounded
reopen proof process-locally. Changed apply authorizes the exact source and
prior semantic value, charges retained target work, completely reopens the
target, and repeats semantic and locality verification. Drift, replay on a
byte-different source, or a mismatched direction conflicts. `Patch::inverse`
swaps the two artifacts and directional proof and restores the accepted source
byte-for-byte, including deleted previews. Fingerprints are diagnostic rather
than authorization. The patch has no stable serialization, composition,
merge, bounded history, or library-owned durable publication;
`Package::write_to` remains the output boundary.

The private scalar projection is
`TSTTableTitleSettingsArchive.proto`, containing only fields 22, 33, and 37.
A strict handwritten pass validates their canonical wire forms and routes the
field-30/36 `TSP.Reference` payloads through the existing private scalar
`TNNumbersSheetReferenceArchive` lazy view. Raw caller-owned records remain
the preservation and rewrite authority; generated production code has no
encoder, `RepeatedView`, or `LazyRepeatedView`. The deterministic closure is
five files and 32,332 bytes under 33 KiB with SHA-256
`56cfd70666ffa6079175bdab0a63a4ddd055099edf3c771ed3ad8b3051596ee1`.
Typed byte, field, work, nesting, reference, archive, and aggregate transaction
ceilings apply before publication.

The final performance review found no P0 or P1 issue. A structural gate drives
the real rooted `Package` transaction path over 4,096 and 8,192 native states.
Strict fields rise from 53,307 to 108,363 (2.0326 times), wire work from 315,936
to 636,752 (2.0155 times), references from 16,386 to 32,770, exactly `2 + 4N`
(1.9999 times), and transaction work from 9,084,384 to 18,298,157 (2.0142
times). Every measured resource remains below the 2.3 structural ceiling, and
a maximum-minus-one allowance rejects atomically before output. This is a
deterministic work gate, not a wall-clock benchmark. The accepted P2 costs are
bounded linear selector temporary vectors and redundant strict decodes on the
changed path; neither grants unbounded retention or quadratic work.

The authoritative matched Numbers 14.4 native pair is the 136,204-byte
control resave SHA-256
`25c9fc858ca4fb4f1fedeafb944e96afb81af03a082a41be297ecf6f2542dbdb`
and the 136,273-byte hidden-title artifact SHA-256
`ac8a7117ad6256b0da2e6d191b9e64f721b689d71696a89ac0f78bc6aa513a28`.
Their selected native delta establishes only field 22 `Some(true)` to absent.
The final Rust gate began from the 136,357-byte source SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`.
Rust produced a 136,351-byte hidden-title artifact SHA-256
`4c7f6340b6f2675240577c5b59d5c154de24c8a7e763a31257c56a9899a8e40c`,
and inverse restored the exact source artifact. Numbers 14.4 opened it without
warning, showed the title checkbox off, retained the 22 by 7 table, `B2`
marker, and `B3 = 42`, and after native resave and exact-path reopen produced a
136,353-byte artifact SHA-256
`5b162f8431f45333f0ae9a8654dfa724794f2ec2b391ea11f6a5eee7822cbb10`
with the same UI state. Neither gate proves a native explicit-false spelling,
field-37 outline mutation, or a right to rewrite type-6284 ViewState; the
focused writer refuses that transient type on changed admission and leaves all
accepted ViewState exact.

The frozen host cut removes only
`NumbersEditor::{table_title_settings, set_table_title_settings}`: 32
production lines, four whole direct tests totaling 245 lines, and the complete
39-line `edit_numbers_table_title.rs` example. Private
`table_title_settings_in_package`/`set_table_title_settings_in_package`, their
wire module, the shared title `Settings`, Pages/Keynote table-title methods and
tests, and all format-specific table CRUD remain because the two other formats
still use that cross-format substrate. This is a Numbers direct-owner handoff,
not retirement of Pages or Keynote title editing.

The frozen gates pass 9/9 focused codec tests, 141/141 protobuf tests, 111/111
Numbers library tests, 2/2 private transaction tests, and 5/5 public focused
integration tests. The boundary unit suite passes 173/173; live host and
focused audits are empty, and the full checker is restored to only the
unchanged 14 baselines: six development-only annotations and eight edge
classifications. No edge or debt closes: final topology is 64 workspace
packages, 237 internal declarations, 14 `litchi-iwa` dependency declarations,
and 14 ordered debts, including debt 015 (`litchi-iwa -> litchi-numbers`).

## 2026-08-11 amendment: aggregate Pages section-settings exact patches

`Package::edit_section_settings(selector)` resolves an exact producer name or
checked position once against the immutable source snapshot and stores only the
content-free `Position`. The edit owns one complete `section::Settings`
replacement. `Edit::settings` borrows the staged value and consuming
`Edit::set(Settings)` validates the archive-free name and pagination invariants
before replacing it. The four optional Booleans, three optional pagination
values, and optional name all retain native presence; equality of the complete
value is the only semantic no-op.

The no-op path shares the original package and exact source artifact, reports
zero touched components and preview deletions, and does not inspect template
dependencies, plan a rewrite, reassemble, or reopen. This ordering
also admits reads and exact no-ops for supported legacy nested packages. A
changed edit requires exact package provenance and rechecks the resolved
section, its eight-field source value, rooted template prerequisites, and
exact cache/preview state against the same immutable artifact before planning
output.

The changed dependency closure contains only the selected type-10011 section
message and required enclosing IWA/ZIP length framing. The transaction
raw-splices only fields 17--22, 26, and 28, retaining source position for
replaced records and appending newly present selected records in numeric order.
Rooted document/view-state layout caches, their exact reference metadata, and
all canonical root previews remain byte-exact. The complete candidate is
reopened under retained limits and must reproduce the requested settings,
stable package statistics, unchanged cache/previews, and exact unrelated
semantics and physical locality.

`section::settings::Patch` exposes the resolved path, complete before/after
semantic values, diagnostic fingerprints, `is_noop`, and `inverse`. Exact
source and target artifacts, selected-payload evidence, and locality proof
remain private. Fingerprints are never
authorization. Application requires exact artifact bytes and the directional
semantic/physical preconditions; a changed apply reopens the retained target
rather than replaying a rewrite. Replay, stale or competing bases, source or
target tampering, and using an inverse in the wrong direction yield
`PatchConflict`. Inversion swaps every directional artifact and proof and
restores the accepted source byte-for-byte.

The retained `edit_section_name` and `edit_section_pagination` transactions are
projection-scoped adapters over this patch core. A name adapter changes only
`Settings::name`; a pagination adapter changes only fields 20--22. Each retains
its established semantic facade and exact-patch behavior, but neither performs
an independent payload decode, rewrite, reassembly, reopen, or locality check.
Concurrent patches do not merge implicitly; exact source authorization keeps
the immutable snapshots independent.

The patch remains process-local. Durable deterministic serialization,
versioned semantic operation/read-write sets, composition, three-way merge,
bounded history, and library-owned atomic filesystem publication remain open
ADR 0003 work. The frozen aggregate codec is exactly five generated files and
80,202 bytes under 80 KiB, has no generated repeated view, and has aggregate
SHA-256
`2202f4b1d394346450cb9f88a41c2784ab476cff23b181fffbab6f37b4a42b62`.
The real rooted-package scaling gate doubles total objects from 4,096 to 8,192:
fields stay 77, `WireWork` stays 564, references stay 4, and
`TransactionWork` changes from 292,154 to 587,222 (2.0100x). Both sizes perform
one output allocation and one reopen. A maximum-minus-one transaction-work
budget fails before output with zero allocations and reopens. Focused
integration passes 7/7, four private production/security tests cover budget
observation, object scaling, alias-metadata refusal, and repeated-reference
scaling/max-minus-one refusal, and the final locality review is clean. The
matched native pairs in ADR 0008, rather than an unrecorded Rust artifact, are
the application UI
oracle for this transaction. The full Pages library/integration total is
118/118, boundary regressions are 181/181, both focused facade/host audits are
empty, and the full checker reports only the unchanged 14 baselines.

## 2026-08-12 amendment: Numbers cell-batch exact patches

`litchi_numbers::table::cells::Edit` is a consuming selector-first batch. It
accepts bounded `Change` values through `set`, `set_a1`, `clear`, `clear_a1`,
`change`, and `extend`; staging rejects invalid addresses, out-of-bounds and
duplicate positions, update-count overflow, owned-text overflow, and fallible
allocation before publication. `commit` sorts the final coordinate set,
elides semantic no-ops, and gives the physical owner one immutable final
overlay. A changed clear of an existing non-empty stored cell materializes the
format's stored-empty state; clearing `Missing` or already-empty storage is an
exact no-op, never a request to erase sparse presence.

A changed `Patch` is an exact directional process-local capability. In
addition to exact source/target artifacts and bounded message/reference
evidence, it privately retains a verified source/target `PackagePair`. Forward
apply borrows the patch, authorizes the retained source artifact and read
profile, verifies directional locality against the retained target snapshot, and publishes that
snapshot without a second reopen. `inverse()` swaps artifacts, snapshots,
reference transitions, and exact preview membership in constant time. A no-op
patch shares the source and performs no physical ownership or dependency scan.
Conflict, malformed evidence, or directional-locality failure is atomic.
Reads and exact no-ops remain broad; a changed source without an exact physical
`SourceCatalog`, including a nested legacy layout, fails as `UnsupportedSource`.

The accepted changed matrix is intentionally smaller than the semantic value
model: finite scalar writes and clears; direct/unsegmented string-list
assignment and release with exact refcounts; missing sparse-tile growth through
the synthetic 513-row boundary for finite non-text scalars; in-place
authored-text replacement in uniquely owned rich backing while retaining
key/storage identity and releasing exact style references; and formula-cache refresh
for a strict supported AST/dependency closure evaluated from the batch's final
overlay. Sparse text-to-missing-tile changes refuse as `SharedString`.
HeaderNameMgr-backed header cells, segmented string lists, shared/COW or
rich text requiring any FieldInfo reference transition, noncanonical or
ambiguous FieldInfo rich ownership, existing formula or error cells, and
modeled unsupported/cyclic/range/deletion/sparse formula closure fail as
`UnsupportedDependency`. Impacted active merge, pivot, category, spill,
hidden, and conditional-style state refuses by its matching dependency kind;
unrelated or inert state remains exact. A modeled missing storage prerequisite
is `UnsupportedDependency { CellStorage }`, malformed storage is
`InvalidSource`, and an unmodeled stored BNC value/source kind is
`UnsupportedSource`. This patch does not construct formula ASTs or mutate
format/control state.

Canonical payload field-1-to-storage and storage field-2-to-style FieldInfo
metadata may be present on the unique rich path and remains exact when no
field-specific reference transition is required.

The strict dependency-only formula projection used for the supported cache
closure remains five generated files/201,539 bytes with SHA-256
`ccd972b3dcd76b6142342d36435f2f76a305c029265853ced04d64c1e2bf1752`.
Its focused codec gate passes 7/7 and the current full protobuf suite passes
178/178.
The adjacent PackageMetadata projection is five files/145,681 bytes with zero
repeated generated views and SHA-256
`ee49927f75c6b632c83055f9b7e647920b389be41bec10e25871a6ef7b56ab31`;
its focused gate passes 7/7.

The transaction charges staging first, passes a remaining allowance through
sequential strict leaf reports, reserves each component before output, charges
exact component costs, and retains publication/locality work through the final
verification. Numeric, unique-text, same-tile, and formula 4,096-to-8,192
transaction-work ratios are respectively 1.1899x, 1.2245x, 1.1396x, and
1.8021x; every governed subterm is at most 2.0x. Required-minus-one formula and
sparse limits reject with zero component, reassembly, output, reopen, and
locality work. These are deterministic counters, not latency or RSS evidence.

The retained `PackagePair` reduces apply to bounded locality verification but
is also explicit process-local memory and durability debt: serialization,
versioning, operation encoding, composition, merge, and history remain absent.
Native numeric B3=43 scalar and no-impact rich-text commit/apply/inverse gates
pass; the
latter preserves its independent formula/cache and is not impacted-formula
refresh proof. The completed host cut retires the three direct
`NumbersEditor` cell writers, two raw-ID model writers,
`TableCellBatch::apply_numbers`, 15 obsolete direct tests, and the legacy
example while retaining shared attached-table and fixture-only adapters.

## 2026-08-12 amendment: Keynote existing-slide deletion patches

`Package::edit_slide_deletion()` creates a one-operation immutable
`slide::delete::Edit`. `Edit::remove_slide` resolves an exact navigator name or
checked zero-based `Position` against the base semantic snapshot and stores
only that semantic position. Missing and ambiguous names, missing positions, a
second staged operation, an empty edit, and deletion of the presentation's
final slide are typed refusals. Deletion has no semantic no-op: every accepted
commit removes exactly one existing slide.

Commit first validates the complete source and proves the selected flat
Document -> Show/SlideTree -> SlideNode -> Slide ownership path. Exact
aggregate occurrences and any present `FieldInfo` attribution, package-wide
exclusive inbound ownership, single-message selected objects, absence of
merge/base/diff state, and exact current PackageMetadata component, UUID,
external-edge, and data-owner records are preconditions. Unsupported hierarchy,
a deprecated root-node or secondary slide-list topology, duplicate identities
or references, surviving inbound ownership, versioned ownership, mismatched
locators/counts, and malformed or ambiguous metadata refuse before publication.

A changed commit removes the one raw Show slide-reference field record while
preserving its siblings, deletes the selected Node and Slide objects while
retaining co-located objects and component registrations, and applies the
exact PackageMetadata transition: two object-to-UUID entries, one unversioned
Node-component-to-Slide-component object reference when that ownership form is
present, and each selected Node/Slide data-reference owner/count entry are
removed. A supported component-level edge is preserved instead. The last-object
identifier, the component records themselves, global data-catalog records and
payloads, unrelated ownership, and unknown records remain exact. A component
data-reference record remains with its surviving owners or is removed when no
owners survive. The exact root preview names are invalidated; a case-distinct
or nested near-name is not.

One package reassembly and complete reopen must reproduce the source order
minus the selected semantic position before the immutable candidate is
published.

`slide::delete::Patch` is a directional, process-local capability. Publicly it
exposes the resolved position and diagnostic fingerprints; privately it keeps
the exact source/target artifacts, complete before/after slide counts and
navigator-name sequence, and locality evidence. `Package::apply_slide_deletion`
requires the exact directional source, reopens and validates the retained
target, and checks the complete semantic sequence. `inverse()` swaps the same
exact artifacts and restores the accepted source byte-for-byte. Fingerprints
are diagnostic only; stale, unrelated, or wrong-direction application is
`PatchConflict`.

This patch does not serialize a semantic operation or satisfy ADR 0003's
durable JSON, versioning, composition, merge, history, or library-owned atomic
filesystem publication goals. The focused example supplies synchronized
sibling-temporary, distinct-output, no-clobber publication, but that command is
not the library durability contract.

Deletion is not `gc`. Data-reference ownership records attached to the two
deleted objects are removed only after exact proof, while every `Data/` member
and global PackageMetadata data-catalog record stays preserved. Shared or
uncertain media may therefore remain physically present and unreachable. A
future reclamation transaction needs its own package-wide reachability and
disposition proof.

## 2026-08-12 amendment: formula-cache planning foundation only

The internal Numbers cell-cache planner preserves an unrelated cycle marker
byte-for-byte, refuses when a marked formula impacted by a scalar edit survives
the final same-batch overlay, and succeeds when that overlay removes the marked
formula. Graph work has an exact max-minus-one refusal regression, while
scratch and allocation remain bounded by the planner limits. These checks
remain implementation foundations for the existing scalar cell transaction;
they do not create an `Edit`, `Patch`, or `Commit` formula-authoring operation.

Public formula insertion, host retirement, formula-native validation, and
formula-authoring performance evidence remain unclaimed.

## 2026-08-13 amendment: Pages section-background patches

A Pages section-background `Edit` binds one resolved semantic position to an
immutable package snapshot. It can stage only `set_solid` or `clear`; it cannot
manufacture an unsupported native fill. Exact no-ops share the source snapshot.
A changed edit requires an exact source, proves the selected section/payload and
field-30 reference ownership, rewrites only field 30, reopens the complete
candidate, verifies semantic readback and locality, then publishes one
immutable package.

`Patch` is directional and process-local. Its private exact source and target
artifacts authorize apply and inverse; public diagnostics expose only semantic
path, change state, touched-component count, and reparse state. Replayed,
stale, tampered, or wrong-source application fails with `PatchConflict`, and
inverse restores the accepted source exactly. Unsupported fills and ambiguous
or reference-owned field-30 state refuse changed edits before publication.

This is not a durable operation log, merge format, history system, or
library-level atomic filesystem-save guarantee. The focused CLI publishes to a
distinct no-clobber destination through a synchronized sibling temporary; it
demonstrates command-line handling, not a broader package-save contract.

Apple Pages 14.4.1 accepted both focused candidate artifacts without repair or
conversion, saved and reopened their exact paths, and retained the requested
`Color Fill` dark-red and `No Fill` UI states. This native check confirms the
supported semantic transition; Pages' own resave is not used as an exact
locality oracle.

## 2026-08-13 amendment: canonical Keynote read snapshots

Keynote read snapshots now have one focused owner with two provenance levels.
`litchi_keynote::Document::{open, open_with_options}` accepts either a complete
ZIP or a frozen app-authored package directory, completes bounded semantic
projection eagerly, and publishes a cheaply clonable archive-free full `Show`,
rooted text, source-derived metadata, and source statistics.
`litchi_keynote::Package::snapshot` instead cheaply shares the immutable
complete regular-file artifact and its semantic state; `semantic_snapshot`
returns a cheap shared archive-free `Document` whose source diagnostics are
intentionally absent, so its `metadata()` and `stats()` are `None`. The retired
`KeynoteDocument` had duplicated these responsibilities with its own `Bundle`,
`ObjectIndex`, and `OnceLock<Document>`.

Together, `Document` and `Package` retain the supported read capabilities.
Semantic path reads, snapshots, text, slides, metadata, show, validation, and
source statistics are available through source-backed `Document`; exact ZIP
byte ingress and artifact-backed statistics remain on `Package`.
`from_archive_bytes` was merely an alias for `from_bytes` and is intentionally
not preserved under a second name; exact byte callers use
`Package::from_bytes`. `KeynoteDocumentStats::application` was the constant
`Keynote`, so the focused stats type need not repeat it.

Directory semantics follow a different invariant from exact artifact
ownership. Focused `Document` freezes app-authored directories through
`PreparedSource`; the cross-format coordinator can delegate through that same
boundary. `Package::open` refuses directories because exact `write_to` and edit
provenance require the complete ZIP artifact. Directory capture never promotes
an `Index.zip` fragment into a supposedly complete writable source, and the
semantic snapshot makes no preservation claim for other sidecars, `Data/`, or
previews.

Semantic corrections are part of the cutover. `Package::text` visits only
storages reachable in rooted presentation order. Slide reads preserve rich
`Storage` fragments rather than flattening body/date content into legacy text
vectors. Metadata and validation use the focused package's stricter bounded
rules and may recover plist revision/content-status data differently. These
differences forbid an object-for-object equality claim while preserving every
supported read capability.

Metadata lookup binds to the exact canonical logical
`Metadata/Properties.plist` path. A hostile near-name with the same basename is
inert. Catalog-normalized legacy nested-ZIP wrapper prefixes are supported;
arbitrary flat wrapper prefixes are not treated as canonical metadata paths.
Source-backed metadata is always present because it begins with semantic Show
fields; canonical properties, when present, contribute only narrowly decoded
scalar diagnostics. Their independent hard admission ceiling is 64 KiB.

No edit, patch, publication, or native mutation is introduced by deleting the
duplicate reader. Existing `Package` transactions continue to bind their
patches to exact immutable source snapshots.

## 2026-08-13 amendment: canonical Pages read snapshots

`litchi_pages::Document` captures a complete ZIP path or app-authored package
directory where stable path ingress is supported, or borrowed ZIP bytes or
caller-owned shared ZIP bytes on every supported platform, then eagerly
completes the bounded Pages projection, drops the physical components and
unselected sidecars, and
publishes one archive-free `Arc`-backed state. `snapshot()` is a constant-time
clone of that state. Its source-derived `metadata()` and `stats()` are present;
documents built from semantic values return `None` for both. Semantic
`validate()` walks the retained values without reconstructing a package or
allocating another document.

Windows Pages path capture deliberately fails closed because stable,
reparse-safe identity is not yet pinned there; borrowed and shared byte ingress
remain available. Source-open failures cross the public archive-free boundary
as content-free `ReadError` categories and numeric bounds. The three canonical
Pages metadata members in a ZIP are declared-size and compression preflighted
before package entries are materialized. Selection compares exact raw logical
name bytes after stripping only the selected legacy outer-package prefix, so
raw near-names remain excluded. Unrelated supported members can still be
expanded under the generic source limits before projection drops them.

`litchi_pages::Package::snapshot` has a different contract: it shares the
immutable complete ZIP artifact and remains the authority for exact bytes,
physical validation, and transaction provenance. Its borrowed
`semantic_document()` shares archive-free semantics but intentionally has no
source diagnostics; callers use `Package::{metadata, stats}` for artifact
diagnostics. Directory-backed semantic values cannot be used to edit or claim
preservation of metadata beyond the selected authorities, `Data/`,
previews, media, unknown sidecars, or complete package bytes.

Deleting the duplicate host reader adds no edit, patch, commit, publication,
or save operation. Existing focused Pages edits remain bound to exact
regular-file or byte-backed package snapshots, and their patch limitations are
unchanged.

Snapshot and ingress behavior is covered by the 15/15 document-reader gate;
the complete focused Pages suite passes 153/153. Supporting archive and
detector suites pass 93/93 and 32/32, the host generated-roundtrip passes 1/1,
and the 227/227 boundary suite reports zero live retirement or focused-public-
API findings.

## 2026-08-13 amendment: canonical Numbers read snapshots

`litchi_numbers::Document` now has one shared immutable state containing the
rooted semantic sheets, validated plain-text length, and optional source
diagnostics. `snapshot()` is a constant-time clone of that state, and
`shared_sheets()` shares the same semantic allocation. Source-backed path,
borrowed-byte, and shared-byte construction eagerly complete bounded capture
and semantic projection before publication, then release the physical source.
The resulting snapshot contains no archive, package member, native identifier,
protobuf/Buffa value, preview, or media payload.

Source diagnostics follow the snapshot's provenance. A source-backed document
may retain canonical projected metadata and content-free `DocumentStats`; a
document constructed from semantic sheets or returned by
`Package::document()`/`document_snapshot()` returns `None` for both. Rooted
`plain_text()` is available for every document because it is derived from
retained workbook semantics, not retained source diagnostics. It uses the
validated exact output length and reports a content-free allocation error
rather than publishing partial text. Empty rendered Text/Formula values do not
contribute a line or separator.

`litchi_numbers::Package::snapshot` remains a different capability: it shares
the complete immutable regular-file/byte artifact, parsed components, exact
write source, and transaction provenance. Its package-derived `Document`
shares semantic values but does not acquire source metadata or statistics.
Directory semantic capture cannot be upgraded into exact-artifact ownership
and makes no preservation promise for unselected metadata, `Data/`, previews,
media, or unknown sidecars.

Unix path-backed snapshots use pinned, no-follow capture. Other non-Windows
targets use version-checked path capture; Windows path-backed snapshots fail
closed until reparse-safe pinned acquisition exists. Byte-backed snapshots
remain portable.

Deleting the duplicate host reader creates no edit, patch, commit,
publication, or filesystem-save operation. Existing Numbers transactions
remain bound to exact package snapshots, including their locality, conflict,
inverse, and native verification rules. The reader retirement establishes no
latency, RSS, allocator-event, zero-copy, or end-to-end Buffa-laziness claim;
its supported memory statement is structural sharing plus release of package
state after eager semantic projection.

Frozen verification passes 16/16 focused reader cases, with a seventeenth
Windows-configured case; 240 Numbers library cases pass and four are ignored;
compatibility and name gates pass 5/5 and 10/10. Archive coverage passes 127
cases (125 unit plus two integration), detector coverage passes 40/40, and the
host library passes 1,397/1,397, while generated-roundtrip and doctest gates
pass 1/1 and nine passed with three ignored. Host all-target check and no-run,
strict scoped host Clippy, focused
all-target Clippy, strict focused rustdoc, formatting, and diff checks pass. The
boundary units pass 237/237 and both live retirement/API audits report zero
findings. Broad host all-target Clippy remains blocked by unrelated existing
lints; the global boundary policy still reports 14 unrelated
`soapberry-zip`/`xml-minifier` debt findings.
