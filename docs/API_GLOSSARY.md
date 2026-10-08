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
| Immutable owner snapshot (physical + semantic, all state verbs) | `Package` | — | ✅ docx/pptx/numbers/pages/keynote/doc/ods · 🔴 odt/odp (split out of the merged type, W3) · 🔴 xlsb (W2) |
| Read-only semantic projection, word processing | `Document` | — | ✅ pages · ✅ docx (`Package::document()` view) — becomes owned (W2) |
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
| Edit entry, primary transaction | `Package::edit()` | `edit_document()` (docx), `opened_presentation_transaction()` (pptx), `Workbook::edit()` (xlsx — moves to `Package`) | ✅ iWork/xlsx-name · 🔴 docx/pptx (W2), xlsx (W2), odt already `edit` but moves onto `Package` (W3) |
| Edit entry, scoped family | `Package::edit_x()` | — | ✅ numbers/keynote/pages families |
| Transaction type | `Edit` | `Transaction` (pptx), `WorkbookEdit` (xlsb) | 🔴 W2 |
| Commit | `Edit::commit() -> Result<Commit>` | — | ✅ everywhere; shape fixed by the rows below |
| New snapshot accessor | `Commit::package()` / `Commit::into_package()` | `Commit::snapshot()` (docx/pptx), `Commit::{workbook, into_workbook}` (xlsx) | ✅ iWork · 🔴 W2 |
| Diagnostics on every commit | `Commit::diagnostics()` | commits without diagnostics (xlsx, pptx) | ✅ docx/iWork · 🔴 xlsx/pptx (W2) |
| Split commit | `Commit::into_parts() -> (Package, Patch)` | — | ✅ xlsx/docx/pptx exist; name fixed |
| Patch application | `Package::apply(&patch)`; scoped `Package::apply_x(&patch)` | `publish_document_commit()`, `apply_document_patch()` (docx), `apply_opened_presentation_commit/patch()` (pptx), `Workbook::apply` (xlsx — moves) | ✅ iWork · 🔴 W2 |
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
| Open from bytes (copy allowed) | `from_bytes(&[u8])` | `from_bytes(Vec<u8>)` signatures (xlsx/odt/odp), `from_slice` (xlsx) | 🔴 W2/W3 |
| Open from shared bytes (zero-copy) | `from_shared_bytes(Arc<[u8]>)` | — | ✅ numbers; offered only where zero-copy is real |
| Open from reader | `from_reader(R)` | — | ✅ docx/xlsx; extended per wave |
| Configured open | `open_with_options(path, OpenOptions)`, `from_bytes_with_options(bytes, OpenOptions)` | `open_with_limits`, `open_with` (doc), `open_with_options` under other type names | envelope decided in §5 |
| Minimal fresh artifact | `Package::new()` | `Workbook::create()` duplicate (xlsx) | ✅ docx/pptx · 🔴 xlsx (W2); all formats gain it |
| Save to path | `save(path)` — **always atomic** (sibling temp + rename) | non-atomic `Builder::save` (odt/odp) | ✅ OOXML · 🔴 ODF (W3, behavior change with CRUD evidence) |
| Save with explicit durability | `save_with_durability(path, Durability)` | — | ✅ OOXML; extended everywhere |
| Bytes output | `to_bytes()` | `Builder::build() -> Vec<u8>` as output (ODF) | `build()` returns `Package`; output lives on `Package` |
| Stream output | `write_to(impl Write)` | `to_stream`, `write_to_stream` | 🔴 W2 |
| Open encrypted | `open_with_password(path, password)` | — | ✅ where encryption exists; ODF keeps same name |
| Save encrypted / re-encrypted | `save_encrypted(path, pw, Mode)` / `save_reencrypted(path, pw)` | — | ✅ OOXML |
| Explicit plain output | `save_plain(path)` / `to_plain_bytes()` | — | ✅ OOXML; the only plain output of an encrypted source |
| Query source encryption | `encryption() -> Option<Mode>` | — | ✅ OOXML |

## 4. Limits and configuration

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Physical ingress limits | `SourceLimits` | `ReadLimits` (OPC re-exports in docx/pptx/xlsx/xlsb), `SourcePackageLimits` (odt), `package::Limits` (doc), `DocumentSourceLimits` (iWork) | 🔴 all waves |
| Semantic limits | `SemanticLimits` | `DocumentLimits` (numbers), `DocumentSemanticLimits` (keynote) | ✅ pages · 🔴 W1 |
| Open configuration envelope | `OpenOptions` (per format, extensible fields) | `PackageOpenOptions` (doc — merges), `ReadOptions` (numbers), `iwork::Options` (facade) | 🔴 W1–W4 |
| Construction idiom | consuming builder, fallible `max_<dimension>` setters, fallible `build()` | positional `new(...)` constructors (iWork) | ✅ OPC · 🔴 W1 |
| Build-time invalid limit | `Error::InvalidLimit { kind, value, maximum }` | standalone `*LimitsError` types (doc, iWork) | folds into the module `Error` |
| Ingress limit exceeded | `Error::Limit { kind, observed, maximum }` | `OpcError::ReadLimit`, facade `ResourceLimit` mapping drift | — |
| Limit kind enums | `SourceLimitKind` / `SemanticLimitKind` | `ReadResource`, per-crate one-off kinds | — |
| Dimension naming | nouns (`input_bytes`, `members`, `entry_bytes`, `aggregate_bytes`, `depth`, `events`, …); setters `max_<noun>` | a second name for an existing dimension | shared dimensions share defaults (512 MiB input, XML depth 256) |
| Coverage policy | caller-selected limits on **every** public ingress path | — | 🔴 gaps: ppt/xls (W4), ODF (W3) |

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
`replacement` explanation. For example, a future completed selector migration
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
