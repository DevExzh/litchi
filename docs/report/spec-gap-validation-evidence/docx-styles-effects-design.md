# DOCX `stylesWithEffects` ownership and source-backed API design

Status: bounded design note for the next implementation owner. This note records
local specification and fixture evidence; it makes no production changes and does
not claim full Word effects or native rendering support. Root implementation
decisions are: fresh glossary creation keeps the existing four-resource seed and
does not auto-seed `stylesWithEffects`; lifecycle support is resource-level; the
typed projection is read-only until a lossless style rewrite exists; and package
facade `put`/`remove` methods follow the existing `Result<bool>` convention.

## Requirement and current gap

The checked-in `[MS-DOCX]` section `2.1.1` defines `stylesWithEffects` as a part
that stores a copy of the Style Definitions part and says that a package **MUST
NOT contain more than two** such parts. The local Open XML SDK part metadata
confirms the wire identity:

```text
content type:    application/vnd.ms-word.stylesWithEffects+xml
relationship:    http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects
target stem:     stylesWithEffects
root element:    styles
version:         Office2010
```

The feature matrix marks this row as topology-only (`crates/litchi-docx/docs/FEATURE_MATRIX.md`,
the `stylesWithEffects part` row). The separate `rPr` effects row is already
typed through `Run::effects` and the mutable run APIs. The missing owner is
therefore the package resource and its source-preserving lifecycle; it is not a
request to add a renderer or to duplicate the existing run-effects model.

This slice follows the accepted snapshot/edit/patch and glossary ownership ADRs:
immutable source-backed snapshots, isolated edits, atomic validation and
publication, exact-source no-ops, source-checked reversible patches, and no
physical part or relationship identifiers in ordinary CRUD. A loaded snapshot
must retain its `PartData`/source token, or an equivalent owner-bound shared
source allocation, rather than copying `blob()` into a detached `Arc<[u8]>` that
has lost source provenance. `glossary::raw` stays the low-level graph escape
hatch.

## Fixture evidence

All hashes below are SHA-256. Part hashes are over the uncompressed XML member,
not the ZIP entry bytes.

| Fixture | Package hash | Effects parts and observed owners | Effects XML size/hash |
| --- | --- | --- | --- |
| `test-data/poi/test-data/document/Bug54849.docx` | `f54182713ea5ce5d77b9593d3d9d24e645460043cec0b40ef59c932385f084d3` | main `word/document.xml` → `rId3` → `word/stylesWithEffects.xml`; glossary `word/glossary/document.xml` → `rId2` → `word/glossary/stylesWithEffects.xml` | main 19,883 / `799de1f7a4ce43f0ca101dc750e8a8bd6e75bcb721f7744d4787dad576cda3b1`; glossary 16,138 / `d27f6ced340ffa173b3b861b08a4006e76673a411687b4145f5e46dc6dae13eb` |
| `test-data/poi/test-data/xmldsign/ms-office-2010-signed.docx` | `bc55c0362722818823a6dd95f8e0ca9869e179ace972a0915241feb4677bde5f` | main `word/document.xml` → `rId2` → `word/stylesWithEffects.xml`; no glossary | 15,710 / `00c5cda7671bf545a8c97312f14b2b8bc0ee7fa469b36c25ee158c8a5c1c1568` |
| `test-data/ooxml/docx/ComplexNumberedLists.docx` | `297a085a7d433af2eeee7661e8db21539452cb5850964847744a1e9f5f258b0b6` | main `word/document.xml` → `rId3` → `word/stylesWithEffects.xml`; no glossary | 15,955 / `b4bf5d355a45daf0a1085e73fe27041b5db22bfa23820f18ffaa7f9c8cb70f18` |
| `test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/testGlossary.docx` | `8ccd581d8f0ae102b220228ad26b3974821a7ce8e3ff7df4b78f7da8a0d06ed9` | main `word/document.xml` → `rId3` → `word/stylesWithEffects.xml`; glossary `word/glossary/document.xml` → `rId2` → `word/glossary/stylesWithEffects.xml` | main 20,117 / `e72df38e71a351ebaaf7b102e8ab7862b7e5eb04cc86e187f8510746b70d3f54`; glossary 16,244 / `78112f02ff4e0b94a6c99f8688d86c6fa91d54d6c57c504001453b9c73e3dad1` |

Each effects member has a WordprocessingML `w:styles` root and an exact content
type override. In the two glossary fixtures the main and glossary copies are
different sizes and bytes, so they are independent resources; do not silently
deduplicate them or assume that either copy is byte-identical to its ordinary
`styles.xml` counterpart. A bounded scan of these four producer files found no
`glow`, `shadow`, `reflection`, `textOutline`, `textFill`, `scene3d`, `props3d`,
or `w14:` effects token inside the effects parts. That is evidence for the
part-copy/topology contract only and does not reduce the normative extension
surface elsewhere in the document model.

The signed fixture is a required publication case: an exact no-op must retain
the signature, while a changed effects resource must follow the existing
explicit unsign/resign policy.

## Ownership contract

Use a semantic owner enum with two values, for example
`styles::effects::Owner::{MainDocument, Glossary}`. The enum is intentionally
not a package path, relationship ID, or ZIP implementation detail.

* `MainDocument` resolves only from the package main document part.
* `Glossary` resolves only from the glossary document part reached through the
  package's glossary-document relationship. The glossary graph remains owned by
  `litchi-docx::glossary`.
* Each owner has zero or one internal relationship with the exact
  `stylesWithEffects` relationship type. Absence is a valid read result; loading
  must not synthesize a part.
* Across the package, the owner closure contains zero, one, or two effects parts.
  A second part for the same owner, a third package part, an orphan effects part,
  an external target, a duplicate target, or a wrong target content type fails
  closed for the typed owner. The low-level raw graph can still preserve an
  untouched unknown resource when a caller chooses the raw preservation path.
* The target must have the exact effects content type and parse as a bounded
  `w:styles` document in the package's established Strict/Transitional dialect.
  The effects relationship itself has the Microsoft URI above; do not invent a
  Strict replacement that is not in the local evidence.
* Treat the effects part as a leaf resource. No fixture has an outbound
  relationship from it. If a future producer supplies one, retain it through
  the raw graph or reject it under the typed profile rather than silently
  dropping it.

The existing glossary code is a useful starting boundary: it already declares
the exact relationship/content-type constants, classifies the relationship as
`stylesWithEffects`, and validates a glossary target as an internal leaf with
the exact content type. `seed_semantic_graph` currently seeds only four
glossary auxiliary resources (styles, settings, font table, and web settings),
and that four-resource seed is retained. Fresh glossary publication does **not**
auto-seed a fifth effects resource. A glossary effects resource is created only
by an explicit resource `put`, with the same owner and content-type checks; an
existing effects resource is loaded and preserved independently of the ordinary
styles resource.

## Recommended public seam

Keep `crates/litchi-docx/src/styles.rs` in place and add a public
`styles::effects` child. Rust can resolve that child from
`crates/litchi-docx/src/styles/effects.rs` without moving the established
standard-styles module. Private codec and package helpers can live below
`src/styles/effects/` as implementation grows. This avoids a gratuitous module
rename while keeping the contextual API short.

The public source-backed types should be:

```text
styles::effects::Owner
styles::effects::Resource
styles::effects::Conformance
styles::effects::Snapshot
styles::effects::Transaction
styles::effects::Commit
styles::effects::Patch
```

`Snapshot` is an owner-wide state, so it exists even when the owner has no
effects resource. Its private state is `Option<Resource>` (the same absence
shape as `glossary::Snapshot`'s optional graph), plus owner/conformance and a
source revision/token. `Snapshot::resource()` returns `Option<&Resource>` and
`Snapshot::is_empty()` reports absence. A present `Resource` retains exact XML
and a typed, read-only style-definition projection.

The loaded resource must keep a provenance-bearing payload handle: use
`PartData` and its source lineage when the owner is opened through the
source-backed OPC path, or `Part::blob_arc()` together with a private
owner/package source token when using the existing materialized `OpcPackage`.
The latter is the minimum zero-copy fallback; never construct a snapshot by
copying `part.blob()` into an unrelated `Arc<[u8]>`. `Resource::xml_bytes()` may
borrow the retained bytes, but it must not detach the source token or expose
physical part identity. A caller-created replacement resource may own one
validated source allocation; it does not inherit the package's source token
until publication.

The current `Styles<'a>`/`Style` model is a useful read projection, but it does
not contain the complete style property grammar or unknown direct-child source,
so it cannot be used as a lossy rewrite model. Resource-level publication is
the supported write boundary; no style-level setter is implied.

Precise suggested module and facade signatures are:

```text
effects::load(&OpcPackage, Owner) -> Result<Snapshot>
effects::apply_patch(&mut OpcPackage, &Patch) -> Result<Snapshot>
effects::apply_commit(&mut OpcPackage, Commit) -> Result<Snapshot>
effects::put(&mut OpcPackage, Owner, Resource) -> Result<bool>
effects::remove(&mut OpcPackage, Owner) -> Result<bool>

Package::styles_with_effects(owner) -> Result<Snapshot>
Package::put_styles_with_effects(owner, resource) -> Result<bool>
Package::remove_styles_with_effects(owner) -> Result<bool>

Resource::from_xml(xml, conformance) -> Result<Resource>
Snapshot::{owner, conformance, resource, is_empty, edit}
Transaction::replace_resource(Option<Resource>) -> Result<&mut Transaction>
Transaction::clear_resource() -> &mut Transaction
Transaction::commit() -> Result<Commit>
Commit::{changed, snapshot, patch, into_parts}
Patch::{before, after, is_empty, inverse, apply(&mut OpcPackage), undo(&mut OpcPackage)}
```

`load` follows `glossary::Snapshot::load`; the important contract is that the
return is an owner snapshot whose `resource()` can be `None`. The mutable
package `put` and `remove` methods intentionally return `Result<bool>`:
`false` is a semantic no-op and preserves bytes/signatures, matching
`Package::put_glossary`, `Package::put_fonts`, and `Package::remove_fonts`.
`Document::styles_with_effects()` can be a short main-owner read convenience,
but package publication remains the owner-aware facade. Glossary callers use
the same owner-bound resource through the glossary package service, while
`glossary::raw::{Graph, Part, Rel}` remains the explicit low-level graph API.

For the first implementation slice, `Resource` construction validates a
complete `w:styles` XML source and parses the existing typed read projection;
it does not rewrite individual style definitions. Resource-level
add/replace/remove is the complete CRUD needed to close the topology gap. Do
not expose a `set_effects` API that implies the full cascade or visual-effects
vocabulary. A later complete style-definition model may add selector-first
`add`, `replace`, `rename`, and `remove` operations over style IDs/names, but
only after its rewrite codec proves preservation of direct formatting, extension
children, MCE branches, and untouched order.

## Commit, inverse, and graph/resource rules

`Snapshot::edit()` creates an isolated transaction over the optional resource.
The transaction supports `replace_resource(Some(Resource))` and
`replace_resource(None)`/`clear_resource()`; it does not pretend to edit the
individual style definitions. A no-op commit returns the same source handle
and performs no graph or signature mutation. A changed commit must:

1. rewrite only the selected effects resource, preserving root attributes,
   unknown children, namespace spelling, and unrelated package bytes;
2. parse the candidate back into a `Snapshot` and validate the owner, exact
   relationship/content type, package-wide two-part limit, and inbound closure;
3. return a `Commit { snapshot, patch }`, where `Patch` stores private exact
   before/after owner states (including `None` for an absent resource), source
   tokens, and graph preconditions, supports `inverse()`, and rejects stale or
   differently owned sources; and
4. publish the part, relationship, and content-type changes atomically. A failed
   validation or stale patch leaves all package bytes and signatures unchanged.

Removing an owner resource applies the `after.resource == None` state: it
removes the relationship and deletes the target only when no other inbound edge
owns it. Main and glossary resources are not merged. Creating a resource
allocates a collision-free package target and relationship ID internally, then
reopens it through the same owner validator. For signed packages, an unchanged
source goes through the existing exact no-op path; changes require the caller's
explicit signature disposition. `Patch::apply` returns the owner-wide
`Snapshot`, whose `resource()` is `None` after a successful removal, so the
result type does not lose the removal state.

## Concrete ownership and test files

The implementation can be split along these boundaries:

* `crates/litchi-docx/src/styles.rs`: add `pub mod effects;` while retaining the
  established standard-styles module.
* `crates/litchi-docx/src/styles/effects.rs` and private children under
  `src/styles/effects/`: source snapshot/resource handle, typed projection,
  bounded validation, patch/inverse, and errors. The payload handle must retain
  `PartData`/source lineage or an owner-bound shared `Part::blob_arc()` token.
* `crates/litchi-docx/src/package/package/styles_with_effects.rs` (or the
  existing `parts.rs`): owner lookup and atomic publication, with facade methods
  matching `Package::put_* -> Result<bool>` and `Package::remove_* ->
  Result<bool>`.
* `crates/litchi-docx/src/glossary/{graph.rs,package.rs}`: glossary-owner
  binding, explicit-only creation policy, inbound closure, and raw-graph
  handoff. Keep the already-recognized relationship/content-type mapping in one
  place; retain the existing four-resource semantic seed.
* `crates/litchi-docx/src/document/package/codec.rs`: main-document convenience
  accessor; it currently searches only the ordinary `styles` relationship.
* `crates/litchi-docx/src/validation.rs` and the DOCX constants owner: recognize
  the exact effects relationship/content type in package validation. The
  Microsoft DOCX extension constants can stay in `litchi-docx`; `litchi-opc`
  should not acquire a format-specific semantic owner.
* Tests: a new native-fixture matrix covering all four files above; source
  no-op/changed/inverse/stale-patch cases; explicit absent/present/remove
  states; main-only, glossary-only, both, duplicate, third-part, orphan,
  wrong-content-type, external-target, and shared-inbound failures; and signed
  no-op/change publication. Existing glossary graph tests that assert four
  seeded parts remain four and should gain a separate explicit-effects put case.

## Boundaries and blockers

This design does not establish style inheritance/cascade, layout, rendering,
native Word acceptance, or a complete independent visual-effects model. The
four fixtures exercise topology and copies of style definitions, not effect
elements themselves. The checked-in `crates/litchi-docx/src/resources/stylesWithEffects.xml`
is currently unused by the package constructor; it should be audited as a
canonical authoring template before being used for fresh resource creation.

The principal implementation blocker is wiring the owner-aware resource service
into both the materialized `Package` facade and the source-backed OPC path while
retaining source tokens/managed payload handles. Completing a
source-preserving style-definition rewrite is deliberately deferred; until
then, resource-level CRUD plus typed read projection is the safe boundary.
