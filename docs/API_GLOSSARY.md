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
| Immutable owner snapshot (physical + semantic, all state verbs) | `Package` | — | ✅ numbers/pages/keynote/doc/ods · ✅ XLS (0197–0198) · 🔴 pptx remaining direct mutators (0127 removes the package writer) · 🔴 docx immutable public owner (0019 shares the physical graph; mutable APIs remain) · 🔴 odt/odp (split out of the merged type, W3) · 🔴 xlsb (W2) |
| Read-only semantic projection, word processing | `Document` | — | ✅ pages · ✅ docx owned `Package::document()` projection (0019) |
| Read-only semantic projection, spreadsheets | `Workbook` | `Spreadsheet` | 🔴 numbers `Document`→`Workbook` (W1) · 🔴 xlsx `Workbook` demoted to projection (W2) · 🔴 ods `Spreadsheet`→`Workbook` (W3) · ✅ XLS (0197) |
| Read-only semantic projection, decks | `Presentation` | — | 🔴 keynote `Document`→`Presentation` (W1) · ✅ pptx view / odp name (shape changes W2/W3) · 🔴 ppt (W4) |
| Projection accessor | `Package::{document, workbook, presentation}()` | — | owned projections primary; borrowed `*_ref()` views secondary |
| Option-bearing fresh authoring | per-format `Builder` in a focused module | `WorkbookWriter` (xlsb), `Writer` (doc), `*DocumentBuilder` prefixes (litchi-iwa) | ✅ XLSB (0175) and XLS (0200) `Builder::build()` return `Package`; 🔴 remaining W2–W4 |
| iWork document identity (explicit, required) | `Identity` with `Identity::generate()` / fixed bytes | implicit random identity inside builders | 🔴 W1; `build()` without identity is a typed error |

## 2. Transaction vocabulary (iWork canonical form)

Canonical shape: scoped families live in
`module::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` with package
entry points `Package::{x, edit_x, apply_x}`; the one primary document
transaction is the bare `Package::edit()`.

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Edit entry, primary transaction | `Package::edit()` | `edit_document()` (docx), `opened_presentation_transaction()` (pptx), `Workbook::edit()` (xlsx — moves to `Package`) | ✅ docx eager/source entry name (0018; eager package-bound commit 0022, source 0031) · ✅ iWork/xlsx-name · ✅ pptx eager entry (0116) · 🔴 xlsx owner (W2), odt already `edit` but moves onto `Package` (W3) |
| Edit entry, scoped family | `Package::edit_x()` | — | ✅ numbers/keynote/pages families |
| PPTX slide metadata | `Edit::{set_slide_name, set_slide_hidden}` | — | Eager/source primary edits and unified slide `name`/`hidden` setters; eager, source primary and focused source edits accept combined `set_slide_properties` (0162–0163) |
| PPTX combined shape properties | `Edit::set_shape_properties`, `ShapePropertiesReplacement` | — | Eager/source primary edits share checked `Position` targets and optional text/bounds in one atomic operation (0153); broader shape properties remain pending |
| PPTX chunked shape publication | `Edit::{set_shape_text_chunks, set_shape_property_chunks}` | — | Eager/source primary and focused source edits admit bounded, globally disjoint chunks together for one same-slide publication (0158); single-chunk setter limits are retained |
| Transaction type | `Edit` | `Transaction` (pptx), `WorkbookEdit` (xlsb) | ✅ pptx primary `opened::Edit` (0116), XLSB structural and cell-value `Edit/Commit/Patch` families, diagnostics and Package commit triads (0165–0171); 🔴 remaining XLSB mutators/error families and scoped PPTX families (W2) |
| Commit | `Edit::commit() -> Result<Commit>` | — | ✅ everywhere; shape fixed by the rows below |
| New snapshot accessor | `Commit::package()` / `Commit::into_package()` | `Commit::snapshot()` (docx/pptx), `Commit::{workbook, into_workbook}` (xlsx) | ✅ iWork/docx eager (0022) and source (0031), PPTX eager (0118) · 🔴 other W2 |
| Diagnostics on every commit | `Commit::diagnostics()` | commits without diagnostics (xlsx, pptx) | ✅ docx/iWork/PPTX eager (0118) · 🔴 remaining owners (W2) |
| Split commit | `Commit::into_parts() -> (Package, Patch)` | — | ✅ docx eager (0022) and source (0031), PPTX eager (0118); 🔴 remaining owner shapes |
| Patch application | `Package::apply(&patch)`; scoped `Package::apply_x(&patch)` | `publish_document_commit()`, `apply_document_patch()` (docx), `apply_opened_presentation_commit/patch()` (pptx), `Workbook::apply` (xlsx — moves) | ✅ iWork · ✅ docx eager immutable `apply` (0020) and source (0030); ✅ PPTX primary immutable apply (0118); 🔴 other W2 owners |
| No-op introspection | `commit.diagnostics().changed()`; `patch.is_noop()` | `Transaction::is_changed()`, `Commit::is_changed()` (pptx), `Patch::is_empty()` (xlsx) | ✅ iWork/docx/PPTX primary (0118) · 🔴 remaining W2 |
| Inverse patch | `Patch::inverse()` | — | ✅ everywhere |
| Rollback | *dropping the `Edit`* | `Transaction::rollback()` (pptx) | ✅ primary PPTX edit (0116); 🔴 scoped PPTX families (W2) |
| Mutable writer on a package | *none — `edit()` is the only write path* | `document_mut()` (docx), `presentation_mut()` (pptx) | PPTX `presentation_mut` removed (0127); 🔴 remaining W2 authoring verbs move into `Edit`, direct mutators and `UnsafeEdit` vocabulary still require migration |

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
| Minimal fresh artifact | `Package::new()` | `Workbook::create()` duplicate (xlsx) | ✅ docx; ✅ canonical PPTX constructor (0127); ✅ XLSB (0175); ✅ XLS (0199) · 🔴 xlsx (W2); all formats gain it |
| Save to path | `save(path)` — **always atomic** (sibling temp + rename) | non-atomic `Builder::save` (odt/odp) | ✅ OOXML, including lazy DOCX (0041) · 🔴 ODF (W3, behavior change with CRUD evidence) |
| Save with explicit durability | `save_with_durability(path, Durability)` | — | ✅ OOXML; extended everywhere |
| Bytes output | `to_bytes()` | `Builder::build() -> Vec<u8>` as output (ODF) | ✅ owning DOCX (0044); lazy managed DOCX bytes pending. `build()` returns `Package`; output lives on `Package` |
| Stream output | `write_to(impl Write)` | `to_stream`, `write_to_stream` | ✅ owning/lazy DOCX Package (0041–0042); 🔴 specialized routes and remaining W2 |
| Open encrypted | `open_with_password(path, password)` | — | ✅ where encryption exists; ODF keeps same name |
| Save encrypted / re-encrypted | `save_encrypted(path, pw, Mode)` / `save_reencrypted(path, pw)` | — | ✅ OOXML |
| Explicit plain output | `save_plain(path)` / `to_plain_bytes()` | — | ✅ OOXML, including owning DOCX bytes (0044); the only plain output of an encrypted source |
| Query source encryption | `encryption() -> Option<Mode>` | — | ✅ OOXML |

## 4. Limits and configuration

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Physical ingress limits | `SourceLimits` | `ReadLimits` (OPC re-exports in docx/pptx/xlsx/xlsb), `SourcePackageLimits` (odt), `package::Limits` (doc), `DocumentSourceLimits` (iWork) | ✅ OPC `SourceLimits`/`SourceLimitsBuilder` and OOXML `SourceLimits` re-exports, record 0035 · ✅ DOC (0192); ✅ XLS eager/retained policies (0196–0198); 🔴 ODF/other legacy/iWork |
| Semantic limits | `SemanticLimits` | `DocumentLimits` (numbers), `DocumentSemanticLimits` (keynote) | ✅ pages · 🔴 W1 |
| Open configuration envelope | `OpenOptions` (per format, extensible fields) | `PackageOpenOptions` (doc — merges), `ReadOptions` (numbers), `iwork::Options` (facade) | ✅ DOCX ingress envelope (0038–0040); owning lazy-policy support and W1–W4 pending |
| Construction idiom | consuming builder, fallible `max_<dimension>` setters, fallible `build()` | positional `new(...)` constructors (iWork) | ✅ OPC and DOC (0192) · 🔴 other W1 |
| Build-time invalid limit | `Error::InvalidLimit { kind, value, maximum }` | standalone `*LimitsError` types (doc, iWork) | ✅ OPC `InvalidLimit` with contextual maximum (0036); DOC (0192); other owners pending |
| Ingress limit exceeded | `Error::Limit { kind, observed, maximum }` | `OpcError::ReadLimit`, facade `ResourceLimit` mapping drift | ✅ OPC (0036) and DOC (0192) kind/observed/maximum; broader taxonomy pending |
| Limit kind enums | `SourceLimitKind` / `SemanticLimitKind` | `ReadResource`, per-crate one-off kinds | ✅ OPC/OOXML (0036) and DOC (0192) `SourceLimitKind`; semantic kinds pending |
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
  Owning DOCX publication preserves OPC causes and accepted-output counts
  (0043); broader leaf error cleanup remains pending.
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


## 26. DOCX direct property projections (record 0045)

`properties::{Run, Paragraph, Table, Row, Cell}` represent the properties stated
by one element's own property container. `Run::direct_properties` and
`Paragraph::direct_properties` avoid conflating those values with existing part
properties; table, row and cell views expose `properties()`. These are read
projections, without style inheritance or a new transaction family.

`Paragraph::addressable_runs()` supplies the positional navigation run set;
`runs()` retains its existing outermost-run contract. `InlineKind` classifies
inert inline content by resolved namespace and name. Retained XML views resolve
ancestor namespace bindings. Property values that fail supported schema parsing
are reported through `invalid_properties`, not repaired into defaults.

Record 0046 adds direct-body formatting and paragraph/run structural writes.
Effective relationship integration, caller semantic limits, canonical owner
consolidation and complete family registration remain pending.


## 27. DOCX formatting and body structure (record 0046)

Package-bound `document::Edit<P>` now exposes `format_run`, `format_paragraph`,
`insert_formatted_paragraph`, `move_paragraph`, `remove_paragraph`, `insert_run`,
`append_run` and `remove_run`. `BodyAnchor` addresses a gap relative to a typed
`BodyBlock` or the body end; block kinds retain separate checked `Position`
sequences. Run positions here count direct run children, matching
`replace_run_text`; the broader navigation set remains `addressable_runs`.

`NewParagraph` and `NewRun` carry text and typed local properties for these edit
operations. They are not additional package owners or transaction families.
Public transactions retain the existing `commit().into_package()` and immutable
`apply(&patch)` paths. XML-only callers still use `document::snapshot::Edit`.

`LimitKind::FormattingStagingBytes` identifies a conservative fixed staging
bound for property plans and generated markup. Managed formatting operations
currently refuse; caller-selected semantic limits and managed retention remain
required before the primary transaction family can be registered complete.


## 28. Unified node and value primitives (record 0047)

`node::{Node, View, ViewOptions, NodePath, PathStep}` provide borrowed generic
projection and typed paths. Numeric navigation uses `Position`; native identity
construction is explicit in `node::raw::{Ident, identity_step}`. There is no
ordinary native identity constructor. Test-support fixtures may exercise raw
paths; production property registries must not advertise native identity fields.

`value::{Length, Color, ValueType, Scalar, Test}` share coercion and canonical
rendering between future registry writes and query compilation. `WordSpacing`
reads bare unsigned twips and rounds unit conversions at the twip boundary.
Diagnostics and Debug omit authored values; explicit rendering/serialization
produces content. Excessive view depth and non-finite JSON values are refused.

These are format-independent prerequisites, exercised through test-only fake
nodes. Concrete format node payloads remain uninhabited; File query, registry
and write dispatch are not implemented by this record. Grouped Word spacing
input and long grouped-number allocation remain documented follow-up work.


## 29. Unified path input and resolution (record 0048)

`path::{IntoPath, PathArg}` accepts text and reusable core `Path` values.
`path::raw::path` explicitly opts into native identity resolution. Ordinary
resolution accepts positions and registered semantic keys; unknown keys still
produce an invalid-path error before any raw-identity gate. Grammar types carry
syntax without authorizing native resolution. Native `NodePath` conversions
must also pass through `path::raw::path`; ordinary adapters must emit positions
or semantic names.

`Source::Selector` retains typed grammar failures and byte offsets. Selector
Debug and diagnostic rendering omit authored names, values and encountered
characters; explicit Display and data accessors retain content rendering.
`ParseError::offset_by` adjusts a suffix error to the original input offset.

This is shared resolution infrastructure with test-only adapters. Concrete
format navigation, File get/find/query and capability generation remain pending.

## 30. DOCX indexed blocks (record 0049)

`Blocks`, `IndexedBlock` and `BlockContentControl` are leaf-owned read views.
Block selection and returned ordinals use `Position`; kind-specific ordinals
count direct visible children, flattening custom-XML wrappers. Display equations
have their own kind. `transaction_position()` refuses unless raw addressing is
proven, and never treats a wrapped child as a direct body edit target.

Native paragraph/control identity access added by this record lives under
`document::raw`. Paragraph identifier lookup returns the first matching block;
it is not an ordinary semantic selector. Existing legacy identity methods remain
migration debt. Namespace bindings are retained with indexed ranges.

Generation memos are bounded clean-value caches. `trim_document_memos()` releases
unpinned entries. Managed source projections refuse detachable semantic values,
including effective documents backed by managed donor parts. Configurable
semantic limits, concrete unified nodes and complete family registration remain
pending.

## 31. Unified query plans (record 0050)

`query::{IntoQuery, QueryArg, Plan}` accepts selector text, typed syntax and
reusable compiled plans. Ordinary predicates cannot select native IDs;
`query::raw::query` explicitly admits them and native identity scope paths.
A raw compiled plan retains that authorization on reuse. Semantic names remain
ordinary selectors. Query grammar constructors alone grant no raw access.

Typed syntax is bounded before recursive compilation. Query walkers stream in
document order, stop on the first error or requested limit, and refuse excessive
frame depth or visits. These fixed bounds are prerequisites for configurable
semantic budgets, not their final API.

The design's unknown-key diagnostics and absent-value comparison semantics stay
intact. Debug and diagnostic rendering omit authored selector content; explicit
syntax rendering and data accessors remain available. The engine currently runs
through test-only format adapters; production File dispatch remains pending.

## DOCX unified read adapter (record 0051)

DOCX nodes use positional canonical paths. Paragraph/content-control identity
selection requires raw paths or queries; ordinary node projections hide these
identity fields. Typed leaf access and transaction-position reads validate the
retained owner revision and return typed errors. The static query vocabulary
and internal document views are integrated; File dispatch and production
registry/capability publication remain pending.

## File read dispatch (record 0052)

DOCX unified reads are now available from File, including lazy query Matches
and typed JSON rendering. Ordinary paths/results remain positional; native
selection requires explicit raw admission. The facade exports the same node,
path, query and view modules. View options remain internal pending typed public
integration; other format adapters and capability publication remain pending.

## Capability discovery (record 0053)

`File::capabilities()` and `capabilities::Capabilities::of` expose typed
read-support metadata. Ordinary element/property discovery excludes native
identity fields and native selector templates using the query compiler's
classification. Positional paths and semantic selectors remain discoverable;
raw selection is still explicit. `view::Mode::ALL` enumerates supported and
refused modes for inspection. Source freshness is validated by reads, not by
static capability discovery. Full property-registry and gap coverage remain
pending.

## XLSX source-grid prerequisites (record 0054)

New grid and cache-diagnostic types use the focused names
`workbook::source::{Grid, CacheDiagnostics}` without root aliases. XLSX minimal
and full API profiles ban the candidate names `SourceGrid` and
`SemanticCacheDiagnostics`. Grid addressing retains `At`/`Area`; row/column
coordinates remain checked. This does not claim migration of the existing
XLSX package/projection or transaction API families.


## XLSX readable formatting (record 0055)

Source cell records use `workbook::source::Record`; their optional style keys
reuse opaque, lineage-bound `style::StyleKey`. Keys from another workbook are
refused, while cloned owners share the lineage. No numeric source key accessor
is public. The stored number-format ID is available only through explicit
`style::cell_format::raw::number_format_id`; ordinary NumberFormat reads expose
format codes and classification. Both XLSX glossary profiles register these
bans. Legacy table-model physical metadata and existing XLSX transaction-family
migration remain pending.


## XLSX unified read projection (record 0056)

File XLSX reads return concrete nodes with canonical positional sheet/table/
defined-name paths and checked A1 coordinates. Authored sheet-name reads are
fallible; structural positions are infallible. Native `numFmtId` metadata is
explicitly marked raw in the static property registry, hidden from ordinary
views/discovery and admitted only by raw queries. Raw properties are separate
from resolver identities. The source leaf's checked `style_key` yields the
existing opaque lineage-bound key without cloning a cell. Legacy XLSX owner,
transaction and raw table-model migration remain pending.


## Independently retained PPTX scenes (record 0057)

The new `shape::owned` module uses `Scene`, `Handle`, `Handles`, `Selector`
and `Error`. Selection admits only exact case-sensitive producer names and
`Position(Position)`; missing values return `Ok(None)`, ambiguity has a
content-free typed error. Handles expose `position() -> Position`; raw owner
XML lives in `owned::raw`. Debug omits authored content. Minimal/full PPTX
profiles enforce these names for the new module. Existing borrowed Shape,
scene Key/index APIs and broad PPTX owner/default-policy migration remain open.


## Source PPTX scene retention (record 0058)

`SourceSlide::scene()` returns `shape::owned::Scene`, retaining no presentation
owner. Memo diagnostics use focused `presentation::scene::CacheDiagnostics`;
the prefixed candidate spelling is banned in the minimal/full PPTX profiles.
Diagnostics expose only structural counters. New source scene reads enforce
physical XML limits before MCE and source/execution fences around memo access.
Existing source-owner positional/native accessors and semantic execution-budget
accounting remain migration debt.


## PPTX lazy shape metadata (record 0059)

Common/Shape expose semantic NonVisual values and typed transforms. Newly
introduced relationship IDs and physical spans are confined to `shape::raw`;
minimal/full profiles ban ordinary Payload/PayloadAttribute exports and
payload/span methods. Raw ambiguity and namespace failures are content-free
typed errors. Common/shape/non-visual/payload Debug omits authored data.
Borrowed scenes, iterators, selectors and lookup errors also redact Debug/Display.
Legacy borrowed shape IDs/XML access and selector/error field migration remain pending.


## Shared DrawingML text bodies (record 0060)

PPTX and SpreadsheetDrawing use one strict `text::body::reader`; the new
DrawingML glossary profile bans `Profile` and `read_with`. Native field and
hyperlink/picture relationship IDs are private and use `body::raw` access.
Both PPTX profiles ban ordinary `text_body_span`. Authored model/reader Debug
is redacted. Shared writers return typed results, validate before mutation,
and refuse incomplete or unsupported local projections. Broad source-limit
and owner vocabulary migration remains open.


## Source PPTX related snapshots (record 0061)

`presentation::related` names the retained `Layout`, `Notes`, `Comments`,
`Comment` and `Chart` projections without source-owner prefixes. Semantic
accessors expose authored values; `related::raw` owns physical part names,
relationship identifiers and bytes. Both PPTX profiles ban native identities
on the ordinary snapshot types. Indexed chart selection accepts `Position`.
Source acquisition is fallible; detached snapshot accessors read already
validated values. Broad source-owner vocabulary migration remains pending.


## Source PPTX presentation catalog (record 0062)

Detached masters use `presentation::related::Master`; layouts reuse `Layout`.
Catalog selection uses `Position` and preserves authored XML order. Section
projections use `presentation::sections::{List, Section}` with semantic slide
positions; native GUIDs, slide IDs and opaque extensions use `sections::raw`.
Captured main-root bytes use `presentation::raw`. Ordinary snapshot diagnostics
omit names and native identities. Broad source-owner naming and policy-envelope
migration remains pending.


## Source PPTX modern comments (record 0063)

`presentation::modern_comments::{List, Comment, Reply}` provides detached
semantic projections. `entries`, checked positional selectors, `text` and
`author_name` expose authored content; wire IDs, assignments, namespace bindings,
XML and physical metadata use `modern_comments::raw`. Missing and present-empty
text bodies remain distinct; status and progress accessors apply schema defaults.
Both PPTX glossary profiles register the new native-identity bans. Full package
owner vocabulary and configurable semantic-policy migration remain pending.

## Unified batch payloads (record 0073)

`BatchItem` carries the full OfficeCLI command envelope. Rust uses `kind` for
the wire key `type` and `dump_version` for `dumpVersion`; `op` is an input
alias of `command`, not a second Rust field. Wire field matching is ASCII
case-insensitive. Unknown or repeated envelope fields refuse, including
`command` together with `op`. `Props` retains its ordered assignments and
typed null semantics. Authored fields are omitted from `Debug` and decoder
error messages. `BatchMode::Atomic` is the default; `BestEffort` is explicit.

These are transport inputs, not an alternate document model. Numeric wire
fields retain OfficeCLI's signed 32-bit representation; future command
admission must convert positions to `Position` and validate command-specific
ranges. Batch execution, staged navigation, commit diagnostics and composite
patches remain pending; payload decoding alone grants no capability.

`batch::SourceLimits` and its consuming `SourceLimitsBuilder` use
`max_input_bytes` and `max_items`; `SourceLimitKind::{InputBytes, Items}`
identify those dimensions. `batch::Error` carries `InvalidLimit`, `Limit`,
`Malformed` (line/column only) and `Allocation`. Bounded array ingress returns
the unified error with that typed source and, where applicable, item index.
Out-of-range wire integers fail transport; in-range negative values are
retained for command-specific validation.

## Unified closure transactions (record 0075)

`File::{edit, edit_with_limits}` use `transaction::Edit`, replacing the draft
proxy name `Transaction`. `transaction::{Commit, Patch, Diagnostics, Error,
LimitKind}` name their result/replay family; `SemanticLimits` and its consuming
builder bound mutation calls and retained main XML. The normalized public
`Commit<T>` fields retain the file, patch, diagnostics and closure value.
`Diagnostics::changed()` and `Patch::is_noop()` describe the final logical
artifact, not how many operations ran. Other counts remain explicit facts.

Patches are process-local and bound to exact immutable owner handles (clones
accepted). `Patch::inverse()` reverses endpoint direction; `Patch::apply(&File)`
returns the retained validated endpoint after source guards. It cannot apply to
an independent open or reverse explicit signature removal. `Error::PatchConflict`
remains distinct from physical stale-source failures. `capabilities::Verb::Edit`
advertises this DOCX closure subset; command `Batch` execution stays pending.

## Unified command execution (record 0076)

`File::apply` executes one command and returns `transaction::Commit<Applied>`;
`File::{batch, batch_with_limits}` return `BatchCommit` with `file`, `patch`,
`diagnostics` and `results`. `ItemResult::{Applied, Failed}` preserves input
order; failures retain their original item index. `Applied` carries `changed`
and `path: Option<Box<Path<'static>>>`, populated by creation. The optional
path is indirect so ordinary result entries do not embed the grammar's large
inline path storage; callers borrow it with `as_deref()`.

Atomic and best-effort modes share one transaction and resource profile.
Ordinary best-effort failures roll back one item, while global resource,
source/protection/cancellation failures abort the whole batch. The DOCX core
command subset is discoverable through `Verb::Batch`; contextual gaps retain
follow-ups for command families and text/dump semantics not yet implemented.

## XLSX unified creation and publication (record 0095)

`create(FileFormat::Xlsx)` adopts the native `Workbook::create` owner directly.
`File::{to_bytes, write_to, save, save_as, save_with}` share the existing
publication vocabulary and strict policies across DOCX and XLSX. Path-backed
XLSX handles retain their origin through unlock; bytes and newly created
handles require a destination. Ordinary encrypted output refuses protection
removal, and explicit `SaveOptions::reencrypt` retains the native profile.
Template, Strict and initial-metadata creation options remain discoverable
refusals. The initial XLSX cell mutation surface is described below.

XLSX edit admission now borrows the handle (`editor(&self)`, internal) rather
than consuming it or choosing an engine from reference counts (record 0096).
Successful exact-source promotion shares a native workbook per open; explicit
signature removal has a separate owner slot.

## XLSX unified cell writes (record 0097)

`File::set` accepts direct cell paths with registry properties `value`,
`formula`, and `clear`; the same operations participate in `File::edit` and
atomic or best-effort batches. `ValueType::Scalar` preserves text, boolean,
64-bit integer, and finite-number kinds. Text stays literal; native formulas
omit the leading `=`. Null clears `value` or `formula`.

Transactions retain native patches and immutable endpoint owners. Replay is
bound to the input snapshot, and inverse replay returns the retained source.
Native content comparison detects final no-ops without ZIP serialization.
Structural commands, command coalescing, OfficeCLI string inference/extended
clear modes, and encrypted-owner edits remain explicit capability gaps.

Record 0098 adds rectangle and scoped-selector cell assignments. Ranges include
blank coordinates; formulas broadcast without relative translation. All
selected cells stage in one native commit. `SemanticLimits` now admits
`max_expanded_targets` and `max_staged_bytes` for XLSX fanout, defaulting to
65,536 distinct cells and a 16 MiB payload/action charge. Caller-selected
profiles flow through `File::edit_with_limits` and batch execution. Unknown
predicates, zero matches and excessive expansion are typed failures.

Record 0099 adds `font.bold` and `font.italic` assignments alongside cell
payloads. Boolean false remains distinct from null, which removes the authored
font flag. The native parameter `style::FontChange` stages through
`WorksheetEdit::font`; it exposes no font or cell-format indexes. Stylesheet
resource compaction across separate commits remains an explicit capability
gap.

Record 0100 narrows the XLSX font-authoring profile: the native source-span
editor accepts `mc:Ignorable` stylesheets when the shared MCE admission proof
establishes source/projection agreement. Opaque extensions stay in the
published XML. Alternate content, unproven directives, escaped namespace
bindings and inherited formatting remain explicit gaps.

Record 0101 adds native `WorksheetEdit::state` for the cell-local staged state,
using the same `State` and shared-string identity as native patches. It is not
commit validation or a final workbook preview; pending font changes refuse.
Unified closure edits and atomic batches group distinct direct cell payload
assignments into one native edit. Repeated targets, ranges, selectors, fonts,
reads, metadata and the cumulative staged-byte bound close a group. Best-effort
items remain isolated. Broader coalescing and structural commands remain open.

Record 0102 adds XLSX worksheet creation/removal and tab rename, visibility,
move and swap through the native workbook edit. The sheet registry shares
`name` and `hidden` admission between set and add. Omitted names use
`Sheet{count+1}`; explicit placements use the existing `Placement` vocabulary.
Scoped sheet selectors deduplicate and enforce expansion/staging limits before
publication. Native dependency, protection and last-visible-tab checks remain
authoritative. Structural commands close pending cell runs and retain native
patches for replay/inverse. Non-worksheet removal, OfficeCLI placeholder reuse,
`ifExists`, additional sheet properties and broader coalescing remain open.

Record 0103 adds row/column layout assignment through native worksheet editors.
Rows accept `height` in the existing point-based length vocabulary; columns
accept `width` in native character units. Both accept `hidden`, `outline`
(aliases `outlineLevel` and `group`) and `collapsed`. Null invokes the native
reset/default operation. Existing `outlineLevel` readback remains canonical.
Bounded selectors deduplicate sheet/axis/index identities and retain one native
commit; mixed dimension families refuse atomically. Protection and native
capability guards retain distinct typed error categories. Whole-axis cell
payloads, structural displacement and unit-qualified column width conversion
remain open.

Record 0104 adds native merge geometry writes and XLSX range `merge` boolean
readback. `merge` accepts a checked scalar grammar: range booleans, matching
A1 text, explicit cell/sheet range lists, and range-only `sweep`; null clears
through the same exact-unmerge policy. Cell unmerge is anchor-only. A nonexact
overlapping range reports `merge_not_exact`. Mixed cell or sheet assignments
share one native commit. Selected nodes and geometry intents have separate
bounds, and combined staged actions obey the transaction limits. Raw text
is bounded before trimming or deduplication. Native `TableOverlap` prevents
new merges crossing tables; native preservation guards remain authoritative.
Blank merge anchors now retain their merge readback even without a cell record.


Record 0106 adds native `FontChange::strike` and `FontChange::underline`, used
by unified cell/range `strike` and `underline` assignments. Underline accepts
five native tokens and documented textual aliases through the static registry.
Null removes the authored element; false strike and underline `none` remain
explicit native values. Existing readback omits disabled underline. Key aliases
resolve to the same setter for duplicate rejection. These properties share
native commits with payloads and remain barriers to payload-only coalescing.


Record 0107 extends native `FontChange` with checked `name` and `size` builders
and a typed `script` assignment. It retains shared owned name text and is
`Clone`, not `Copy`. Unified `font.name`, `font.size`, `superscript` and
`subscript` assignments use this native value. Names are byte-budgeted before
ownership; size uses the shared point grammar and the 0.5–409 point authoring
domain. Script false authors baseline, null removes the authored element,
and two script selectors in one property list are a typed conflict. Alias
admission and range broadcast retain the same limits and native commit path.

Record 0108 adds native typed `FontChange::color` and unified `font.color`
(`color` alias). Spreadsheet RGB/ARGB and theme colors retain checked tint;
null removes the authored element. Parsed source lexical markup is excluded
from native change ownership and equality. The facade preserves its existing
read projection: theme tint is exposed, while RGB tint remains native-only.
Color assignments share native commits, limits, conflict checks and payload
coalescing barriers with other font properties.

Record 0109 extends native `FontChange` and the unified registry/read vocabulary
with `font.outline`, `font.shadow`, `font.condense`, `font.extend`, `font.charset`,
`font.family` and `font.scheme`. Family and charset builders are fallible and
check the authoring domains 0–5 and 0–255. Explicit false, zero and scheme
`none` remain distinct from null removal. The facade reads these fields from
the existing native format owner and uses the same shared scheme token table
for read metadata and write coercion.

Record 0110 adds native `AlignmentChange` and worksheet `alignment` staging.
The request retains typed tri-state assignments for the seven curated
alignment attributes. Native style lowering combines it with `FontChange`
per cell before cloning shared resources. Unified `alignment.*` setters use
these owners directly, with OfficeCLI aliases and checked rotation/indent
domains. Null clears an attribute; contextual reading order also clears at
the facade boundary, while native callers may explicitly author context.

Record 0111 adds native `ProtectionChange` and worksheet `protection` and
`quote_prefix` staging. Alignment requests also expose `relative_indent` and
`justify_last_line`, using the native signed integer and optional boolean.
`Alignment::justify_last_line` now preserves absence versus explicit false as
`Option<bool>`; this is an intentional alpha API change to the canonical model.
The unified setters use `alignment.relativeIndent`,
`alignment.justifyLastLine`, `protection.locked`, `protection.hidden` and
`quotePrefix`. Native lowering combines all formatting facets in one request,
without facade-owned style records or a parallel formatting model.

Record 0112 adds worksheet `number_format` staging by checked code. The native
`NumberFormat::validate_code` shares authoring rules with the unified property
admission; `NumberFormat` remains the canonical resource model. Resource IDs
are resolved by the native stylesheet editor. The facade only normalizes
OfficeCLI keyword aliases for the existing `numberformat` vocabulary.

Record 0113 adds native `FillChange` and worksheet `fill` staging using the
existing `Pattern` and `Color` types, reexported as `FillPattern` and `FillColor`
for authoring. The unified writer resolves OfficeCLI incremental/replacement
policy against native effective formats, including the base format for missing
cells. Ordinary node readback keeps its existing base-format omission policy.

Record 0114 extends the existing native `Fill`, `PatternFill`, `GradientFill`
and `GradientStop` types with checked authoring and whole-fill replacement.
`FillChange::replace` owns a shared native fill; the request is `Clone`, not
`Copy`, and its equality/hash use semantic values rather than pointer identity.
A later replacement supersedes earlier granular fields, while granular edits
can compose with a replacement pattern. A granular pattern edit cannot
implicitly convert a replacement gradient. The unified layer admits gradient
text into these native values and retains the existing worksheet transaction.

Record 0115 adds `format::xlsx::Node::cell_format()`, returning the borrowed
native `style::Format` already cached by the node. Complete native gradient
geometry and color values remain in their leaf-owned types; the unified API
does not introduce an aggregate formatting model. The accessor retains source
freshness checks even after the format has been cached.

Native `BorderChange` stages whole-side replacements or independent line/color
changes using the existing `Side`, `Line` and `BorderColor` types. Worksheet
`border` participates in the same Edit/Commit/Patch path as other cell format
facets. Null line and color fields clear only that attribute; whole-side
replacement is explicit. Direction and outline fields distinguish omission,
authored false and removal. The unified `border.*` properties and shorthands
lower into these native changes.

## PPTX primary edit entry (0116)

`Package::edit()` returns the detached `opened::Edit`; `Snapshot::edit()`
returns the same type. The primary `Transaction` name, verbose package entry,
and `rollback()` method are removed without aliases. Dropping the edit keeps
the caller's source unchanged. The PPTX glossary profile rejects all three
removed names. Scoped editors and source-backed publication, package-carrying
commits, diagnostics, canonical patch introspection, and immutable patch apply
remain migration work; this entry-point change does not certify the whole
PPTX transaction family.

## PPTX primary committed package (0118)

`opened::Commit` now exposes `package()`, `into_package()`, `patch()`,
`diagnostics()`, and `into_parts() -> (Package, Patch)`. `Package::apply(&patch)`
returns a new package. The old mutable publication entry points, public commit
`snapshot()`/`is_changed()`, edit `is_changed()`, and patch `is_empty()` are
removed; `Diagnostics::changed()` and `Patch::is_noop()` provide introspection.

Direct publication retains the edit's candidate. Durable patch replay validates
its write set against the destination and retains that destination's unrelated
state. Both carry native tolerance, physical-source and encryption policy.
Changed signed replay refuses; exact no-ops keep signatures and encryption
provenance. Encrypted results still require explicit re-encryption or plaintext
output policy. Commit clones share the returned package; consuming a shared
clone copies its OPC graph while retained source buffers remain shared.

Source-backed commit ownership, scoped families, new-package authoring, and
remaining selector/error vocabulary remain migration work. The glossary checker
therefore enforces the completed primary names without claiming complete family
alignment.

## PPTX primary deferred-source owner (0119)

`pptx::source::{Package, Presentation, Slide, Edit, Commit, Patch, Diagnostics}`
now provides immutable retained-source package revisions and an owned read-only
projection. Shape-text edits compose and return the canonical commit triad;
package `apply` is immutable, patches expose `is_noop` and `inverse`, and saves
use guarded atomic replacement. Managed payload reservations follow retained
candidates, patches, scenes, and shape handles. `from_presentation` shares
already-admitted source state without requiring exclusive projection ownership.

Record 0120 adds bounded durable `Patch::to_bytes` / `Patch::from_bytes`
transport with root, selected-relationship and exact before-XML guards. Local
patches retain owner-lineage checks; decoded patches can replay after reopening.
Structural editing, source-family renames, primary creation, and facade write
dispatch remain pending. This scoped entry does not certify the full source
transaction family or replace the broader glossary gates.


## PPTX facade native-owner promotion (0121)

The internal facade adapter now borrows its handle to obtain an immutable native
source or eager package. The consuming `into_editor` handoff and sharing-only
refusal are removed. Source promotion reuses admitted state with live readers;
successful results are cached, failed attempts remain retryable, and every
lookup preserves source/execution fences. Explicit signature removal has a
separate owner cache. Unified PPTX mutation dispatch remains pending.


## PPTX source revision readback (0122)

Native `source::Slide::{text, text_with_tolerance, text_and_name}` reads the
captured immutable revision through existing native slide readers. Facade
source-revision adoption now carries that revision through text, shape-scene,
paragraph/run and tolerance reads and subsequent native edits. Non-text facts
remain on the admitted source projection under the private same-owner,
shape-text-commit adoption contract. Unified PPTX mutation dispatch is still
pending; these readback adapters do not advertise new write capabilities.


## PPTX unified shape-text writes (0123)

`File::set` now lowers slide shape/placeholder `text` assignments to one native
immutable commit, with scoped selectors, explicit clearing, bounded expansion,
transactions and ordered batches. Original snapshots remain readable; native
proofs back process-local replay and inverse. Source/eager publication shares
the uniform atomic save and origin policy. Signature removal preserves source
freshness; managed unsigned promotion retains the native materialization refusal.
Other properties, structural writes and unified re-encryption remain gaps in
capability reports. No second facade document model or XML editor is introduced.


## PPTX immutable encrypted publication (0124)

`Package::write_reencrypted` accepts shared immutable access, a sink, password
and explicit crypto limits for an already-opened encrypted package. It preserves
retained provenance and refuses mutable authoring state. The facade's existing
`SaveOptions::reencrypt` now delegates to it under common atomic save policies.
Plaintext output stays protected, and encrypted signature-removal requests still
refuse until native edits can fulfill them. Typed crypto errors retain their
classification instead of becoming generic corrupt-file errors.


## PPTX builder and unified creation (0125)

The focused `builder::Builder` owns native typed authoring and returns a
canonical `Package` from `build()`. It has no publication methods. Unified
PPTX creation adopts that result directly. No duplicate `Package::create`
entry remains. Existing mutable `Package::new` consumers still require migration
to finish the canonical fresh-package/builder split; this wave does not certify
that broader migration or its complete glossary family.

Both PPTX profiles reject `Package::create` and publication methods on the
new builder; output remains on its built package.


## PPTX builder consumer migration (0126)

Public PPTX authoring examples, the unified write fixture and live performance
fixture generators now configure `builder::Builder::presentation()` and finish
with `build()` before package edits or publication. Finalization-only archive
round trips are removed. Native source-ingress reopens remain explicit.
`Builder::font_embedding(mode)` configures automatic font embedding before
finalization; font discovery/embedding errors propagate from `build()`.

Native authoring fixtures still use the legacy mutable `Package::new` and
`presentation_mut` surface. Removing that surface and making `Package::new`
canonical remain required; this consumer wave does not mark them conformant.
See [record 0126](unified-ops/0126-pptx-builder-consumer-migration.md).


## PPTX canonical constructor and private authoring (0127)

`Package::new()` now returns the canonical empty presentation, directly usable
with reads, `edit()`, immutable replay and publication. The focused builder owns
private authoring state and finalizes it without serializing/reopening a ZIP.
The public `Package::presentation_mut`, `set_font_embedding` and
`with_font_embedding` methods are deleted. Configure fonts through
`Builder::font_embedding` before `build()`. Unified creation now calls the native
canonical constructor. Native fixtures and public doctests use the builder.

Both PPTX glossary profiles ban the removed names. This completes the constructor /
package-writer separation described as pending in 0125–0126. Remaining package
mutators and authoring verbs still require migration into edit/commit families;
this change does not claim the full single-write-path policy or removal of every
`UnsafeEdit` variant. See [record 0127](unified-ops/0127-pptx-canonical-constructor.md).


Native PPTX slide selection now uses `slide::Selector` with borrowed/owned names
or `Position(Position)` (record 0130). The old slide/root `Key` and its `Index`
variant are removed. Primary opened slide moves take a semantic source and
checked destination, swaps take two semantic selectors, and complete reordering
takes a selector permutation; raw slide IDs remain in private lowering. Other
legacy order families and positional read signatures remain pending.

## PPTX source shape removal (0161)

Primary source `Edit::remove_shape(slide, shape)` and focused slide
`remove_shape(shape)` use the same verb and semantic shape key as the eager
editor. They participate in the existing Edit/Commit/Patch family; no mutable
package escape or parallel transaction type is introduced. The unified facade
routes top-level shape targets through those edits. Source removal preserves
unused relationship resources for exact replay, while shared native admission
refuses retained connector/animation references to removed shape identities.

## XLSB tab state selection (0177)

The primary `Package::edit()` family gains `set_sheet_state(selector, State)`.
It accepts `sheet::Selector` (semantic name or checked `Position`) over the
complete tab catalog and returns `Result<Option<bool>>`: missing is `None`,
while `Some` reports the staged state change. The existing `sheet::State`
vocabulary supplies Visible/Hidden/VeryHidden. No mutable writer escape or
parallel transaction family is added. Record 0183 completes name selection with `set_sheet_name(selector, String)`,
using the same complete catalog and `Result<Option<bool>>` contract. The old
`rename_sheet` method is removed. Selectors resolve against the immutable source
catalog; repeated assignments report the staged value, and missing selectors
stage nothing. Durable tag 1 retains worksheet-ordinal decoding and normalizes
to catalog-based tag 6 before replay or merge.

## OPC batch payload replacement (0181)

The raw OPC escape `OpcPackage::replace_part_payloads` accepts an owning batch
of existing `(PackURI, Vec<u8>)` replacements. It resolves exact names before
ASCII-case-insensitive aliases, preflights targets and deferred payloads, and
rejects repeated resolved targets. Empty input preserves exact-source authority.
Nonempty batches require built-in parts; custom-part packages receive typed
`UnsupportedCustomPartPayload` before mutation because arbitrary setters cannot
provide atomic failure. One revocation records signature policy before the
built-in setters invalidate replaced XML source provenance. XLSB uses this seam
for workbook and worksheet view payloads staged by its structural edit.


## Shared sheet-name dependencies (0184)

The focused `litchi_ooxml_common::xml::sheet_name` helper owns lexical XML and
formula-prefix surgery shared by XLSX and XLSB. Its caller supplies semantic
name equivalence and the original catalog titles; it exposes typed parse,
allocation, scan-limit and unsupported-carrier errors. Format owners retain
transaction, package, selector and publication policy. New XLSB name patches
use tag 7 for dependency-aware replay; historical tags 1/6 retain their earlier
metadata-only behavior.

## XLSB complete tab ordering (0185)

The primary `Package::edit()` family gains `reorder_sheets(&[sheet::Selector])`.
The slice names every tab exactly once using semantic names or checked positions
in the immutable source catalog. Repeated calls replace the staged permutation;
an identity permutation restores the original order. The result reports whether
the staged order changed. No mutable package escape or parallel transaction
family is introduced. Ordering is a native structural barrier; the unified
`move_to` and `swap` adapters publish through this family and flush pending edit
families at transaction boundaries.

## XLSB worksheet insertion (0186)

The primary `Package::edit()` family gains
`insert_sheet(Position, String, sheet::State) -> Result<()>`. The position is a
zero-based slot in the staged complete catalog, including chart tabs. Repeated
insertions use sequential positions; a supplied visibility state authors the
new tab atomically. Other native edit families remain structural barriers.
The unified `File::add` and transaction adapters resolve workbook placements,
coerce name/hidden/visibility properties, and delegate publication to this
native family. No mutable worksheet escape or alternate transaction API is
introduced. Durable insertions use tag 9; existing patch tags retain their
meaning.

## XLSB sheet removal (0187)

The primary `Package::edit()` family gains
`remove_sheet(impl Into<sheet::Selector>) -> Result<Option<bool>>`. Selectors
resolve the immutable source catalog, including chart tabs. Missing selectors
return `None`; repeated removals return `Some(false)`. Distinct removals extend
one sorted source-position set, encoded with durable tag 10. Other edit
families remain structural barriers. The unified sheet-path and homogeneous
sheet-query adapters delegate to one native edit and commit. Additional
internal content dependencies require explicit disposition and currently
receive typed refusal; removal does not implicitly run garbage collection.

## DOC builder and publication entry (record 0188)

Legacy DOC now exports `Builder::build() -> Package`; `Writer` and its direct
output methods are removed. `Package` owns byte, sequential-stream and atomic
path publication, and `from_bytes` borrows its input slice. The DOC minimal and
full glossary profiles guard this completed naming slice. Native package
transaction ownership and the merged ingress envelope follow in records
0190–0191. See [0188](unified-ops/0188-doc-paragraph-loop.md).


## DOC primary package transactions (record 0191)

The immutable DOC `Package` exposes `edit`, `edit_with_options`, and `apply`.
`package::edit` exports the canonical `Edit`, `Commit`, `Patch`, `Diagnostics`,
`Error`, and `LimitKind` family. Commits expose `package`, `patch`, and
`diagnostics`; `into_parts` returns `(Package, Patch)`. Focused property-set
transactions live under `package::property_set` instead of claiming the
primary package names. The unified adapter retains and edits the native
package, including its shared parsed document projection.

Physical source-limit naming and checked construction follow in record 0192.
The remaining focused/raw vocabulary audit is still pending. Registering this transaction
family establishes its shape, not complete DOC CRUD or performance acceptance.

## DOC source limits (record 0192)

DOC exposes `SourceLimits`, a checked consuming `SourceLimitsBuilder`, and
`SourceLimitKind` with input/entry/aggregate byte dimensions. Configuration and
ingress errors carry the canonical inline fields. `OpenOptions` consumes
`source_limits`, `password`, and `leniency`; `DocOptions` uses `source_limits`.
Defaults are 512 MiB input, 512 MiB per stream, and 2 GiB aggregate. Canonical
and raw reader admission checks all declared logical streams before payload
reads, including metadata and opaque streams. Already-parsed adoption cannot
retroactively constrain CFB parsing allocations. Record 0193 extends this preflight to the focused body Snapshot's strict
owner before payload capture, including candidate and profiled reopens.
Complete allocation-budget coverage for raw escape paths and full DOC
acceptance remain open. See [0192](unified-ops/0192-doc-source-limits.md) and
[0193](unified-ops/0193-doc-strict-owner-admission.md).

## XLS focused cell edits (record 0194)

`cell_values::Edit` is now the concrete type; `Transaction` and the duplicate
`Snapshot::transaction()` entry are removed. `Patch` and `SemanticPatch` use
`is_noop()`, and `Diagnostics::changed()` reports exact artifact change. The
focused commit splits into `(Snapshot, Patch)`, with diagnostics read before
consumption. The XLS glossary profile guards these removed spellings. The
immutable `Package`/read-only `Workbook` migration, package commit accessors,
source/options vocabulary, and remaining focused aliases are still pending;
this family is not yet registered as fully conformant. See
[0194](unified-ops/0194-xls-edit-vocabulary.md).

## XLS retained policies (record 0195)

The focused cell-values Snapshot retains its object and physical CFB profiles
through edits and replay. Exact patch application uses the destination's
profile; joins and history recording refuse mismatched profiles. Shared CFB
and OLE ownership retains the selected bounds through candidate validation,
copy-through and layout changes. This policy fix precedes the canonical XLS
SourceLimits/OpenOptions and Package migration; it does not mark that API
family complete. See [0195](unified-ops/0195-xls-retained-limits.md).

## XLS eager source policies (record 0196)

The eager reader now exposes `SourceLimits` / `SourceLimitsBuilder` with
`max_input_bytes`, `max_entry_bytes`, `max_aggregate_bytes`, and checked
`build`. `SourceLimitKind` identifies the dimension in the inline
`Error::InvalidLimit` and `Error::Limit` variants. `OpenOptions::source_limits`
sets the profile; `get_source_limits` reads it. The canonical consuming
`password`, `leniency`, and `compatibility_profile` setters replace the
`with_*` spellings; getters use `get_*`. Credentials are redacted in Debug.
The XLS glossary gate bans all three replaced setters.

This applies to eager reader ingress and raw CFB adoption. Positional-reader
options, focused edit profile convergence, and the immutable Package/Workbook
split remain open; this is not a complete canonical owner registration.
See [0196](unified-ops/0196-xls-source-limits.md).

## XLS package ownership (record 0197)

`litchi_xls::Package` owns immutable source state and returns an owned read-only
`Workbook`. The old generic reader is available only as `raw::Reader`; its
constructors are banned on the canonical Workbook. Native Package publication
preserves exact source bytes and applies source/encryption fences. Facade reads
delegate to this owner. Package primary edits, Builder creation and positional
option-envelope migration remain open; no complete transaction family is
registered by this wave. See [0197](unified-ops/0197-xls-immutable-owner.md).

## XLS primary transactions (record 0198)

The primary `package::edit` family exports Edit, Commit, Patch, Diagnostics,
Error and LimitKind. Package owns edit/apply; Commit exposes the canonical
triad and pair split, while Workbook remains read-only. No-op and inverse
operations retain the recipient's effective profiles. The operation bound is
an inline typed error; physical admission errors keep their typed source.
Builder and full option-envelope unification remain open. See
[0198](unified-ops/0198-xls-primary-transactions.md).

Record 0199 adds XLS `Package::new()` for one empty worksheet and connects the
unified creator and publisher to that owner. Option-bearing XLS authoring still
uses `Writer` pending its complete Builder and encrypted-publication migration.

Record 0200 replaces XLS `Writer` with concrete `Builder`. `Builder::build()`
returns a validated Package, and all publication belongs to that owner,
including explicit encrypted image export. The old builder output methods are
removed and covered by the glossary audit.


Record [0203](unified-ops/0203-xls-unified-reads.md) connects unified XLS
`get`/`find`/`query`/`view` to canonical `Package` reads. Native `Workbook::sheets()`
returns the complete tab catalog, independent of parsed worksheet ordinals;
source-backed cell projections retain `is_formula()`. The first read adapter
still materializes sparse grids for traversal; early termination and measured
facade overhead remain open acceptance work.
