# API Glossary

The living lookup table for the unified cross-format API vocabulary and
default policies. The rationale and the decision itself live in
[ADR 0034](adr/0034-unified-api-glossary-and-default-policies.md), adopted
by owner instruction on 2026-10-08. The table describes the target contract;
existing conformance labels are inherited from the draft and require
per-wave verification. No new rename is claimed by adoption alone.

Rules of reading:

- **One concept, one name.** A canonical name listed here is the only legal
  spelling in any public API. Banned names are deleted, not deprecated
  (ADR 0001: backward compatibility is not a constraint).
- **Layers**: *structural* vocabulary is identical across all formats;
  *domain nouns* differ per document family on purpose (`Document` for word
  processing, `Workbook` for spreadsheets, `Presentation` for decks).
- Status: ✅ already conformant · 🔴 pending rename/change · ➖ not
  applicable to that family. The **Wave** column names the rollout wave
  (W1 iWork, W2 OOXML, W3 ODF, W4 legacy) in which a 🔴 flips.
- Any public API change must update this table in the same commit; once
  `tools/check_api_glossary.py` lands, drift is a CI failure.

## 1. Owner types and layers

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Immutable owner snapshot (physical + semantic, all state verbs) | `Package` | — | ✅ pptx/numbers/pages/keynote/doc/ods · 🔴 docx immutable public owner (0019 shares the physical graph; mutable APIs remain) · 🔴 odt/odp (split out of the merged type, W3) · 🔴 xlsb (W2) |
| Read-only semantic projection, word processing | `Document` | — | ✅ pages · ✅ docx owned `Package::document()` projection (0019) |
| Read-only semantic projection, spreadsheets | `Workbook` | `Spreadsheet` | 🔴 numbers `Document`→`Workbook` (W1) · 🔴 xlsx `Workbook` demoted to projection (W2) · 🔴 ods `Spreadsheet`→`Workbook` (W3) · 🔴 xls (W4) |
| Read-only semantic projection, decks | `Presentation` | — | 🔴 keynote `Document`→`Presentation` (W1) · ✅ pptx view / odp name (shape changes W2/W3) · 🔴 ppt (W4) |
| Projection accessor | `Package::{document, workbook, presentation}()` | — | owned projections primary; borrowed `*_ref()` views secondary |
| Option-bearing fresh authoring | per-format `Builder` in a focused module | `WorkbookWriter` (xlsb), `Writer` (doc), `*DocumentBuilder` prefixes (litchi-iwa) | 🔴 W2–W4; `Builder::build()` returns `Package`, never bytes |
| iWork document identity (explicit, required) | `Identity` with `Identity::generate()` / fixed bytes | implicit random identity inside builders | 🔴 W1; `build()` without identity is a typed error |

## 2. Transaction vocabulary (iWork canonical form)

Canonical shape: scoped families live in
`module::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` with package
entry points `Package::{x, edit_x, apply_x}`; the one primary document
transaction is the bare `Package::edit()`.

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Edit entry, primary transaction | `Package::edit()` | `edit_document()` (docx), `opened_presentation_transaction()` (pptx), `Workbook::edit()` (xlsx — moves to `Package`) | ✅ docx eager/source entry name (0018; eager package-bound commit 0022, source 0031) · ✅ iWork/xlsx-name · 🔴 pptx/xlsx (W2), odt already `edit` but moves onto `Package` (W3) |
| Edit entry, scoped family | `Package::edit_x()` | — | ✅ numbers/keynote/pages families |
| Transaction type | `Edit` | `Transaction` (pptx), `WorkbookEdit` (xlsb) | 🔴 W2 |
| Commit | `Edit::commit() -> Result<Commit>` | — | ✅ everywhere; shape fixed by the rows below |
| New snapshot accessor | `Commit::package()` / `Commit::into_package()` | `Commit::snapshot()` (docx/pptx), `Commit::{workbook, into_workbook}` (xlsx) | ✅ iWork/docx eager (0022) and source (0031) · 🔴 other W2 |
| Diagnostics on every commit | `Commit::diagnostics()` | commits without diagnostics (xlsx, pptx) | ✅ docx/iWork · 🔴 xlsx/pptx (W2) |
| Split commit | `Commit::into_parts() -> (Package, Patch)` | — | ✅ docx eager (0022) and source (0031); 🔴 remaining owner shapes |
| Patch application | `Package::apply(&patch)`; scoped `Package::apply_x(&patch)` | `publish_document_commit()`, `apply_document_patch()` (docx), `apply_opened_presentation_commit/patch()` (pptx), `Workbook::apply` (xlsx — moves) | ✅ iWork · ✅ docx eager immutable `apply` (0020) and source (0030); 🔴 legacy publishers and other W2 owners |
| No-op introspection | `commit.diagnostics().changed()`; `patch.is_noop()` | `Transaction::is_changed()`, `Commit::is_changed()` (pptx), `Patch::is_empty()` (xlsx) | ✅ iWork/docx · 🔴 W2 |
| Inverse patch | `Patch::inverse()` | — | ✅ everywhere |
| Rollback | *dropping the `Edit`* | `Transaction::rollback()` (pptx) | 🔴 W2 |
| Mutable writer on a package | *none — `edit()` is the only write path* | `document_mut()` (docx), `presentation_mut()` (pptx) | 🔴 W2: authoring verbs move into `Edit`; `UnsafeEdit` refusals disappear with the writers |

Deferred (own tracks, not this glossary): `History`/undo-redo, patch
composition (`join`, `three_way`), durable patch formats beyond pptx/xlsx.

## 3. Open, create, save

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Open from path | `open(path)` | — | ✅ everywhere |
| Open from bytes (copy allowed) | `from_bytes(&[u8])` | `from_bytes(Vec<u8>)` signatures (xlsx/odt/odp), `from_slice` (xlsx) | ✅ DOCX owning/lazy Package (0038/0040); 🔴 W2/W3 |
| Open from shared bytes (zero-copy) | `from_shared_bytes(Arc<[u8]>)` | — | ✅ numbers; offered only where zero-copy is real |
| Open from reader | `from_reader(R)` | — | ✅ docx/xlsx; extended per wave |
| Configured open | `open_with_options(path, OpenOptions)`, `from_bytes_with_options(bytes, OpenOptions)` | `open_with_limits`, `open_with` (doc), `open_with_options` under other type names | envelope decided in §5 |
| Minimal fresh artifact | `Package::new()` | `Workbook::create()` duplicate (xlsx) | ✅ docx/pptx · 🔴 xlsx (W2); all formats gain it |
| Save to path | `save(path)` — **always atomic** (sibling temp + rename) | non-atomic `Builder::save` (odt/odp) | ✅ OOXML, including lazy DOCX (0041) · 🔴 ODF (W3, behavior change with CRUD evidence) |
| Save with explicit durability | `save_with_durability(path, Durability)` | — | ✅ OOXML; extended everywhere |
| Bytes output | `to_bytes()` | `Builder::build() -> Vec<u8>` as output (ODF) | `build()` returns `Package`; output lives on `Package` |
| Stream output | `write_to(impl Write)` | `to_stream`, `write_to_stream` | ✅ lazy DOCX Package (0041); 🔴 owning/specialized routes and W2 |
| Open encrypted | `open_with_password(path, password)` | — | ✅ where encryption exists; ODF keeps same name |
| Save encrypted / re-encrypted | `save_encrypted(path, pw, Mode)` / `save_reencrypted(path, pw)` | — | ✅ OOXML |
| Explicit plain output | `save_plain(path)` / `to_plain_bytes()` | — | ✅ OOXML; the only plain output of an encrypted source |
| Query source encryption | `encryption() -> Option<Mode>` | — | ✅ OOXML |

## 4. Limits and configuration

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Physical ingress limits | `SourceLimits` | `ReadLimits` (OPC re-exports in docx/pptx/xlsx/xlsb), `SourcePackageLimits` (odt), `package::Limits` (doc), `DocumentSourceLimits` (iWork) | ✅ OPC `SourceLimits`/`SourceLimitsBuilder` and OOXML `SourceLimits` re-exports, record 0035 · 🔴 ODF/legacy/iWork |
| Semantic limits | `SemanticLimits` | `DocumentLimits` (numbers), `DocumentSemanticLimits` (keynote) | ✅ pages · 🔴 W1 |
| Open configuration envelope | `OpenOptions` (per format, extensible fields) | `PackageOpenOptions` (doc — merges), `ReadOptions` (numbers), `iwork::Options` (facade) | ✅ DOCX ingress envelope (0038–0040); owning lazy-policy support and W1–W4 pending |
| Construction idiom | consuming builder, fallible `max_<dimension>` setters, fallible `build()` | positional `new(...)` constructors (iWork) | ✅ OPC · 🔴 W1 |
| Build-time invalid limit | `Error::InvalidLimit { kind, value, maximum }` | standalone `*LimitsError` types (doc, iWork) | ✅ OPC `InvalidLimit` with contextual maximum (0036); other owners pending |
| Ingress limit exceeded | `Error::Limit { kind, observed, maximum }` | `OpcError::ReadLimit`, facade `ResourceLimit` mapping drift | ✅ OPC kind/observed/maximum (0036); broader taxonomy pending |
| Limit kind enums | `SourceLimitKind` / `SemanticLimitKind` | `ReadResource`, per-crate one-off kinds | ✅ OPC and OOXML `SourceLimitKind` (0036); semantic kinds pending |
| Dimension naming | nouns (`input_bytes`, `members`, `entry_bytes`, `aggregate_bytes`, `depth`, `events`, …); setters `max_<noun>` | a second name for an existing dimension | ✅ OPC common nouns (0037); shared defaults remain 512 MiB input and XML depth 256; other families pending |
| Coverage policy | caller-selected limits on **every** public ingress path | — | 🔴 gaps: ppt/xls (W4), ODF (W3) |

### OPC source dimensions

`SourceLimits` and `SourceLimitsBuilder` use `max_<dimension>` for both
getters and consuming setters. `SourceLimitKind` identifies the same bound.
The shared dimensions below are canonical in OPC (record 0037). Other names
are distinct OPC dimensions, not aliases for a shared one.

| Dimension | Kind | Meaning |
|---|---|---|
| `input_bytes` | `InputBytes` | Compressed package input bytes |
| `members` | `Members` | Non-directory ZIP members |
| `entry_bytes` | `EntryBytes` | Declared uncompressed bytes in one ZIP member |
| `aggregate_bytes` | `AggregateBytes` | Aggregate declared uncompressed ZIP bytes |
| `depth` | `Depth` | Nesting depth in checked XML |
| `events` | `Events` | XML events per checked part or operation |
| `archive_total_entries` | `ArchiveTotalEntries` | ZIP entries including directory records |
| `archive_member_name_bytes` | `ArchiveMemberNameBytes` | Bytes in one ZIP member name |
| `archive_metadata_bytes` | `ArchiveMetadataBytes` | Central-directory metadata bytes |
| `archive_compressed_bytes` | `ArchiveCompressedBytes` | Declared compressed bytes in one ZIP member |
| `parts` | `Parts` | Materialized OPC part count |
| `part_bytes` | `PartBytes` | Bytes in one materialized OPC part |
| `total_part_bytes` | `TotalPartBytes` | Aggregate materialized OPC part bytes |
| `content_types_bytes` | `ContentTypesBytes` | Content-types XML bytes |
| `content_type_mappings` | `ContentTypeMappings` | Content-type defaults and overrides |
| `relationship_parts` | `RelationshipParts` | Relationship part count |
| `relationship_xml_bytes` | `RelationshipXmlBytes` | Bytes in one relationship XML part |
| `total_relationship_xml_bytes` | `TotalRelationshipXmlBytes` | Aggregate relationship XML bytes |
| `relationships_per_part` | `RelationshipsPerPart` | Relationships in one part |
| `total_relationships` | `TotalRelationships` | Aggregate package relationships |
| `relationship_graph_nodes` | `RelationshipGraphNodes` | Discovered relationship graph nodes |
| `total_relationship_xml_events` | `TotalRelationshipXmlEvents` | Aggregate events across relationship XML parts |
| `xml_attribute_bytes` | `XmlAttributeBytes` | Bytes in one checked XML attribute |
| `relationship_target_bytes` | `RelationshipTargetBytes` | Bytes in one relationship target |

`members` and `archive_total_entries` intentionally differ: directory records
count only toward the latter. Likewise, declared ZIP `aggregate_bytes` and
decoded `total_part_bytes` protect different stages. The XML bounds are consumed
by content-types, relationships and supported source-preserving XML operations;
this ledger does not imply that every format-semantic reader uses this profile.
Local validation/custom-data XML policies retain their own names pending their
semantic-limit migrations.

## 5. Errors

- Each focused module owns its non-exhaustive `Error`; crate roots keep only
  ingress/egress error types. God enums (`xlsx::Error`, …) split per module,
  staged per wave. Boundary mapping is typed; string flattening
  (`map_ooxml_error` pattern) is retired.
- Shared variant taxonomy (payload shapes fixed):

| Concept | Canonical variant | Banned |
|---|---|---|
| Recognized but wrong family | `UnexpectedFormat { detected: FileFormat }` | `NotPages` / `NotKeynote` / `NotNumbers`, string messages (`"Not an ODT file: …"`) |
| Unrecognized input | `NotOfficeFile` | — |
| Source file changed after open | `SourceChanged { expected, observed }` | — |
| Patch base mismatch | `PatchConflict { .. }` | `StaleSource` (merges; the two concepts stay distinct) |
| Limit exceeded | `Limit { kind, observed, maximum }` | per-crate duplicates |
| Allocation failure | `Allocation { resource, .. }` | `CollectionAllocation` |
| Unexpected internal failure | `Other(String)` (narrowed) | `ParseError(String)` / `InvalidFormat(String)` for expected failures |

- Facade: `sheet::Workbook` returns typed `litchi_core::Error` (the
  `Box<dyn Error + Send + Sync>` split is removed); facade wrappers preserve
  leaf errors as typed sources. Folded into the UNIFIED_OPS M0a migration.
- Content redaction: errors, diagnostics, and `Debug` never carry authored
  content (user text, entry names, sheet names); structural metadata (part
  paths, limit figures, positions, `FileFormat`) is allowed. ✅ iWork ·
  🔴 OOXML/legacy call-site cleanup per wave.

## 6. Metadata

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Semantic metadata type | `Metadata` (merged field union; typed `Keywords`; real time types) | `ooxml_common::Props`; the `creator` field name (→ `author`) | 🔴 W2 |
| Metadata read | `Package::metadata()` | `Package::props()` (docx) | ✅ facade/ODF/iWork reads · 🔴 docx rename (W2) · 🔴 pptx/xlsx typed entry added (W2) |
| Metadata mutation | `Package::{edit_metadata, apply_metadata}` + `metadata::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | `put_props()`, `props_mut()`, `clear_props()` | 🔴 W2 |
| Custom properties | `custom_props()` / `edit_custom_props()` | — | ✅ name (docx/pptx/xlsx); mutation transactionalized W2 |
| Application properties (OOXML-only) | `app_props()` | — | format-specific layer, OOXML only |
| Fresh-authoring metadata | `Builder::set_metadata(Metadata)` | — | ✅ ODF; not an opened-document mutation |
| iWork metadata write | typed unsupported until native gates pass | — | no capability is claimed ahead of evidence |

## 7. Selectors and coordinates

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Catalog selector | `Selector<'a> = Name(Cow<'a, str>) \| Position(Position)` | the `Id` variant (raw IDs live in raw modules); `Index(usize)` (numbers); `At`/`Index` selector forms (pages footnotes) | ✅ core/pages/keynote shape · 🔴 numbers (W1) |
| Positional selection word | `Position` — checked, zero-based value type | `Index` as a selector variant; raw `usize` in public signatures (facade `slide(usize)`, `paragraph(usize)`) | 🔴 facade fold into M0a |
| Selector-taking methods | `impl Into<Selector>` parameter, `Result<Option<_>>` return; missing → `Ok(None)`, ambiguous/invalid → typed error, never panics, no `Index` trait | method-level splits that bypass the selector channel | pptx `scene::get/at` converge (W2) |
| Grid coordinates (not selectors) | `At` / `Area` with A1 parsing at the edge | treating coordinates as catalog selectors | ✅ xlsx; rule recorded |
| Name identity policy | follows the producing application; fresh authoring defaults to Unicode caseless | silently mixing case policies | ✅ behavior unchanged; per-format registration below |

Registered name-identity policies: xlsx sheet names — Unicode caseless
(Excel semantics); docx glossary entries — NFD/case-fold/NFD (Word
semantics); pages section names — case-sensitive (Pages semantics); keynote
slide names — exact (Keynote semantics).

## 8. Default policies

| Policy | Canonical rule | Banned behavior |
|---|---|---|
| Write path | every mutation of any `Package` — fresh or opened — goes through `edit()`/`edit_x()` + `commit()` | hydrating an opened package into a mutable writer (docx `document_mut()`), per-format refusal entries (`UnsafeEdit`) for a second write path |
| Signatures | changed write on a signed package → `Error::Signed`; exact no-op preserves signatures; state changes only via `sign()` / `resign()` / `unsign()` | pptx commit auto-`unsign()`; policy-parameter entry points (`PlaceholderSignaturePolicy`) |
| Encryption | plain `save()`/`to_bytes()` on an encrypted source → `Error::EncryptionPolicy`; explicit `save_plain()`/`to_plain_bytes()`; `save_reencrypted()` keeps the source profile; unsupported profiles → typed `Error::Unsupported` | silently emitting plaintext; silently re-encrypting with default parameters |
| Determinism | builders/writers never consult ambient clock, randomness, locale, or network; iWork identity is an explicit required `Identity` parameter | `IWorkDocumentIdentity::generate()` inside builders; hidden timestamps in generated parts (audit item) |
| Tolerance reports | fail-closed readers report `Strict(FileFormat)` with `is_clean() == true`; readers that repair report a tracked `ToleranceReport` | `Untracked` (never-clean reporting of a repair-free reader); any silent normalization found by the iWork audit becomes tracked or a typed error |
| Atomic save | `save()` is a sibling-temp + rename replacement in every format | `std::fs::write` whole-package overwrite (odt/odp builders) |

## 9. Enforcement and waves

- `tools/check_api_glossary.py` (W0): rustdoc JSON banned-name lint over
  public items + transaction-family six-piece shape check; generalizes the
  `check_iwork_public_api.py` pattern workspace-wide. Each wave registers exact banned names and public path scopes in
  `tools/api_glossary.json`; registered old names in that surface fail CI.
- Waves: **W0** governance (ADR 0034, this table, checker skeleton) ·
  **W1** iWork renames · **W2** OOXML in docx → pptx → xlsx order ·
  **W3** ODF (package split, atomic save, limits coverage) · **W4** legacy
  (doc/ppt/xls). Facade-surface cleanup folds into the UNIFIED_OPS M0a
  migration so the facade breaks once; leaf-crate unification lands before
  M0a so the `File` enum maps onto unified names.
- Per-wave gates: warning-denied build/tests/Clippy; checker + this table
  updated in the same commit; docs, README, and examples rewritten in the
  same commit; behavior-changing waves add CRUD-checklist evidence.

## 10. Adoption status (2026-10-08)

W0 is complete: the ADR, glossary, rustdoc checker, real-rustdoc regression
tests and CI entry are tracked. The `core-minimal` profile inspects the current
`litchi-core` tree with no bans or registered families, as required for W0.
This verifies the gate mechanism; it does not certify any pending rename or
policy migration in the table above. M0a and M0b
already landed, so the leaf waves adapt the existing facade in the same
changes as leaf renames. iWork remains independently owned and must be
coordinated through API requests. See [record 0016](unified-ops/0016-rebase-and-glossary-adoption.md).


## 11. Registering and running a migration gate

Run `python3 tools/check_api_glossary.py` to build and inspect every registered
profile. By default it uses a temporary build directory outside the checkout
and removes that directory on exit. Pass `--target-dir target/api-glossary`
to keep a local/CI cache. Cargo may create the ordinary root lockfile if absent.
The checker does not remove a caller's lockfile.

`tools/api_glossary.json` has version 1 and an explicit `profiles` array. Each
profile contains `id`, `package`, `default_features`, `features`,
`banned_names`, and `families`. Feature combinations are separate profiles;
results are checked immediately after each build, before the next profile
can overwrite its rustdoc JSON. Register every feature surface affected by a
migration in the same commit as the code and glossary status update.

A banned-name rule contains `name`, a fully qualified `scope`, and a
`replacement` explanation. Optional `recursive: false` restricts the ban to
direct children of the scope; the default checks all descendants. XLSB uses
this for its root physical-limit name while nested semantic limits migrate
separately. For example, a future completed selector migration
could ban `Id` within `litchi_core::selector::Selector`; that does not ban a
format's semantic `glossary::Id`. Names and path segments match exactly, with
no substring matching or implicit exceptions. A scope must exist in the
exported surface or in the nominal defining paths of types reached through
that surface; a misspelled or vanished scope fails rather than disabling a
rule. Private siblings of those nominal types are not inspected. Methods are checked beneath
their public owner paths, including type aliases and re-exports. Public fields,
variants, trait members and `#[doc(hidden)]` APIs are included; private items,
function argument names, documentation and generated blanket impl methods are
not exported vocabulary. Nominal types reached through public signatures are
also inspected, so a public alias cannot hide a banned nominal type.
Dependency type names are attributed to the public signature that exposes
them. Rustdoc sometimes omits opaque/generated types from both its index and
path table; their use-site spelling is checked, but their unavailable members
are not invented. Known-local paths with missing items and missing public
child relations fail as malformed input.

`families` lists fully qualified public module paths that have completed the
shape migration. Each must expose types named `Edit`, `Patch`, `Commit`,
`Diagnostics`, `Error` and `LimitKind` directly; nested or private names and
functions with those names do not count. Re-exported modules and local globs
are resolved. This six-name check does not prove commit signatures, immutable
behavior, diagnostics contents, reversibility or strict default policies;
those still need Rust and CRUD scenario tests. Unregistered families remain
pending work, not implicitly conformant.

An unresolved public glob (for example a dependency module whose children
rustdoc omits) refuses the profile instead of claiming to have checked its
contents. Such profiles need dependency export data support before activation.
This is not an exemption for dependency APIs or a claim of whole-workspace
coverage. The W0 profile has no such unresolved glob.

For an already generated document, use
`python3 tools/check_api_glossary.py --profile core-minimal --json /path/to/litchi_core.json`.
This mode checks the document's crate identity but relies on the caller for
feature/build provenance. Exit codes are 0 for a checked pass, 1 for vocabulary
or family violations, and 2 for invalid policy/input or a failed build.

Regression tests: `python3 -m unittest tools.test_check_api_glossary`.
They compile a small temporary Rust library with the installed rustdoc; no
compiled fixture or rustdoc JSON is committed. CI runs the tests and checker
in the Rust-toolchain job. See [record 0017](unified-ops/0017-glossary-gate.md).


## 12. W2 DOCX entry and focused error names (record 0018)

The primary edit entry is now `Package::edit()` on both DOCX owners.
`document::Error` and `document::Result` replace the prefixed transaction
error/result names. Old public names are removed, with no compatibility
aliases. Private imports may use disambiguating local names when a module
also handles the package ingress error.

`docx-minimal` and `docx-full` enforce these three retired names and the
eager `apply_document_patch` removal (0020). The full
profile enables `automatic-fonts`, `encryption`, `sign` and `vba-inspection`,
which include every optional DOCX feature. Neither profile registers the
complete document family yet: eager `Commit` carries the immutable package
(0022) and the source owner (0031); legacy history retains transitional XML commits. The eager read-only
projection is now owned (0019), and eager `apply(&patch) -> Result<Package>`
is immutable (0020), and eager package-bound commit is implemented (0022).
Specialized immutable ownership, legacy source publication replacement, caller-selected semantic limits,
and mutable-writer removal remain pending. Scoped transactions and other
formats are not marked migrated by this step.

## 13. W2 DOCX immutable eager patch application (record 0020)

Eager `Package` is cheaply cloneable: the physical graph, metadata and retained
legacy authoring state are shared until mutation. `apply(&patch)` returns a new
package without changing its source; a no-op shares state and keeps signatures.
`apply_document_patch` is removed from the public DOCX surface. The source-backed
publisher, legacy commit/history publication methods and mutable authoring
entries still require migration. The complete transaction-family contract is
not yet registered as conformant.

## 14. W2 DOCX shared source owners (record 0021)

Source-backed `Package` is cheaply cloneable, sharing source identity, index,
read caches, admission and repair-report state. Existing stream publishers now
borrow the package, allowing repeated publication and retry while retaining the
original source. Execution reservations follow the lifetime of retained shares.
These stream entries remain transitional; records 0030–0031 implement the
ordinary package-bound triad and immutable source-overlay application. No additional full
transaction family is marked conformant by this step.

## 15. W2 DOCX eager package-bound commit (record 0022)

Root `document::{Edit, Commit}` retain a complete eager package. `commit()`
validates and publishes through the existing immutable application engine;
`Commit` exposes `package`, `patch`, `diagnostics`, `into_package` and
`into_parts -> (Package, Patch)`. Eager `publish_document_edit` and
`publish_document_commit` are removed. Minimal/full profiles ban their return
and ban `snapshot` on the canonical commit type.

Standalone XML and source-backed publication explicitly use the transitional
`document::snapshot::{Edit, Commit}` types. They are distinct nominal types,
not aliases for the canonical package family. Snapshot composition/history and
source stream publishers still require owner migration; no complete family is
registered until that work and the remaining limit/policy contracts are finished. The root patch still
uses main-XML equality for ordinary eager sources; full package lineage remains
required for owner-wide conformance. Broad content-safe Debug cleanup remains
pending beyond the new canonical Edit/Commit implementations.

## 16. Retained OPC topology foundation (record 0023)

`PreparedTopology` shares its effective catalog, generated members, source
proofs and reservations. Its stream publisher borrows the retained candidate;
clones can read or publish it again while physical source checks remain active.
This supplies retained topology ownership for the source-backed DOCX wave.
This foundation alone did not make source-backed `Package::edit()` canonical;
records 0024–0031 supply effective XML authorization and package-bound commits.
Physical archive lineage/version remains distinct from logical candidate state.

## 17. OPC XML logical revision foundation (record 0024)

`SourceXmlPart::begin_revision()` creates a distinct logical range-proof scope
for the current validated XML. Clones share that scope; independent calls do
not. Changed splices require a fresh revision before another edit. Physical
archive identity and original bytes remain separate publication authority.
This enables per-part chaining without archive materialization. It does not
supply package-wide effective catalog revisions or migrate source DOCX commits;
no complete transaction family is registered by this step.

## 18. Retained effective XML application (record 0025)

`PreparedTopology::edit_xml(&self, partname, callback)` returns another retained
candidate. Its checked transaction reads the effective source XML and preserves
previous topology changes. The resulting state is flat against the original
archive, with no parent-candidate chain. This is a low-level publication
primitive, not a registered six-type transaction family. General topology
application, added/authored XML provenance and canonical source DOCX owner
integration remain required.

## 19. Owned effective read foundation (record 0026)

`PreparedTopology::part`, `iter_parts` and `main_document_part` return owned
`PreparedPartView` values. `data()` returns `PreparedPartData`, which retains
its candidate and payload admissions. Source XML proof capture reads the
current effective payload and shares the edit path's provenance checks.
`physical_source_*` accessors explicitly expose original archive identity;
that identity is not a logical revision/cache key. These are low-level OPC
owner facilities. Records 0029–0031 integrate source DOCX projections/commits;
complete transaction-family registration remains pending.

## 20. Main-document patch and Debug policy (record 0027)

The primary DOCX `document::Patch` exposes `is_noop()`; its former `changed()`
method is removed and banned in both DOCX profiles. Commit diagnostics retain
`changed()`. In-tree publishers, examples and performance consumers are migrated.

Main-document transaction Debug output now omits XML, authored text, transfer
metadata and local composition identifiers/conflicts. Error Display and shared
core diagnostics remain pending; this does not register a complete family.

## 21. Effective owning conversion (record 0028)

`PreparedTopology::to_opc_package()` explicitly materializes the current effective
catalog, preserving retained graph changes. It refuses managed base and donor
sources and rechecks all source guards. Record 0029 integrates source DOCX;
this is a conversion seam, not a complete transaction-family registration.

## 22. Source DOCX effective ownership (record 0029)

Source DOCX can retain an effective OPC graph for main-document reads, owning
conversion and publication. Legacy physical operations and the now-fallible
`source_backed_package()` accessor refuse effective owners. Semantic caches are
isolated per adopted owner; pinned views retain donor execution checks.
Records 0030–0031 integrate canonical source Edit/Commit/Patch. Complete family
registration remains pending.


## 23. Source DOCX package-bound commit (records 0030–0031)

Both owner entry points return `document::Edit<P>` and commit to
`document::Commit<P>`, with `P` defaulting to the eager Package. Source-backed
edits retain their complete owner and return a new effective Package without ZIP
serialization/reopening. Common staging methods share one implementation;
private construction and specialized commit implementations constrain ownership
to the supported eager/source types without a public owner trait.

`package`, `into_package`, `patch`, `diagnostics` and `into_parts` have the same
shape for both owners. Source commit checks exact XML identity, protection,
signatures and source/execution fences through immutable `apply`. Generic Debug
never formats the owner or authored content.

Legacy stream/history publishers explicitly consume raw
`document::snapshot::Commit`; their callers start with
`document_snapshot()?.edit()`. This is not a compatibility alias for the ordinary
owner entry point. Record 0032 supplies typed `LimitKind`; full family registration still needs
caller-selected semantic limits, remaining policy vocabulary and specialized ownership migration. Managed donor snapshot
refusal and unmanaged copy/rescan costs recorded in 0030 remain open.


## 24. Primary DOCX typed limit errors (record 0032)

`document::LimitKind` replaces free-form resource strings in the primary
transaction error. `Error::Limit { kind, observed, maximum }` is canonical;
`resource`, `actual` and `max` are removed and banned under that variant in both
DOCX profiles. Durable count/path helpers, history and package transfer adapters
use typed kinds as well. Aggregate transfer-graph payload limits are checked
before copying and report the dedicated typed byte kind.

This does not migrate wrapped OPC/source-story errors or add caller-selected
SemanticLimits. Full family registration and the remaining policy contracts are
still pending. Existing overflow guards retain their documented saturated
`usize::MAX` sentinel.


## 25. Source DOCX effective metadata (record 0033)

`source_backed::Package::metadata()` now reads retained effective core properties
through a shared OOXML graph adapter and decoder. Canonical commit owners can
read metadata; property payloads, relationship redirection and absence follow
that owner's graph. Source/donor fences and per-owner repair deduplication remain
active. Other specialized read/transaction routes and canonical metadata writes
remain pending.
