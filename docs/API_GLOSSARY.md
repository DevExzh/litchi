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
- Status: ✅ already conformant · 🔴 pending rename/change · ⏸ deferred by
  a numbered decision of record 0343 (Dn) that the owner memo
  (`unified-ops/0331-owner-decision-memo.md`) lists for confirmation · ➖ not
  applicable to that family. The **Wave** column names the rollout wave
  (W1 DOCX, W2 PPTX, W3 PPT, W4 XLSX/XLSB, W5 legacy, W6 ODF, W7 iWork,
  W8 File) in which a 🔴 flips. Record 0343
  closed every 🔴 cell: each is now ✅ or ⏸. Owner decision
  [0350](unified-ops/0350-owner-api-debt-decision.md) (2026-10-20) made
  D1–D7, D10 and D11 required work: their cells are 🔴 again, and the cells
  of D1–D6 point at the migration waves of
  [record 0351](unified-ops/0351-api-contract.md), tracked line by line in
  §2a.
- Any public API change must update this table in the same commit; once
  `tools/check_api_glossary.py` lands, drift is a CI failure.

## 1. Owner types and layers

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Immutable owner snapshot (physical + semantic, all state verbs) | `Package` | — | ✅ numbers/pages/keynote/doc/ods · ✅ DOCX `Package`, `source_backed::Package` and `Document` owner/projection boundary (0355, 0355a; raw/model helper rows remain assigned below) · ✅ PPTX `Package` and `source::Package` (0356, 0356a, 0356b; no stale 77-method exception) · ✅ RTF owner split (0223; atomic publication 0226) · ✅ XLS (0197–0198) · ✅ XLSX/XLSB W4/W4b owner and read-only projection source implementation is assembled (`0360`, `0360b`, merge-bases verified); native qualification remains · ✅ PPT immutable owner/publication and primary edit (0211–0212) · ✅ odt/odp/ods immutable `Package` owners split from the read-only `Document`/`Presentation`/`Workbook` projections (0268) · ✅ oth/odm (`Document`), odg (`Drawing`), odc (`Chart`), odi (`Image`), odb (`Database`) and odf-formula (`Formula`) `Package` owners with read-only projections (0282) |
| Read-only semantic projection, word processing | `Document` | — | ✅ pages · ✅ RTF shared owned projection (0223) · ✅ docx owned `Package::document()` projection (0019) |
| Read-only semantic projection, spreadsheets | `Workbook` | `Spreadsheet` | ✅ numbers `Document`→`Workbook`, owned `Package::workbook()` / borrowed `workbook_ref()`, module `litchi_numbers::workbook` (0297) · ✅ xlsx `Workbook` owned projection via `Package::workbook()` (0250) · ✅ ods `Spreadsheet`→`Workbook` owned projection via `Package::workbook()` (0268) · ✅ XLS (0197) |
| Read-only semantic projection, decks | `Presentation` | — | ✅ keynote (0298): `Document`→`presentation::Presentation`, `Package::semantic_snapshot()`→`Package::presentation()` · ✅ pptx view / odp `Presentation` owned projection via `Package::presentation()` (0268) · ✅ PPT shared owned projection (0211) |
| Projection accessor | `Package::{document, workbook, presentation}()` | — | owned projections primary; borrowed `*_ref()` views secondary |
| Option-bearing fresh authoring | per-format `Builder` in a focused module | `WorkbookWriter` (xlsb), `Writer` (doc), `PagesDocumentBuilder`/`KeynoteDocumentBuilder`/`NumbersDocumentBuilder` prefixes (litchi-iwa) | ✅ XLSB (0175), XLS (0200), PPT (0213) and RTF (0225) `Builder::build()` return `Package`; ✅ ODT/ODS/ODP `Builder::build()` return `Package` (0268); ✅ OTH/ODM/ODG/ODC/ODI/ODB/ODF-formula `Builder::build()` return `Package` (0282); ✅ pages `builder::Builder::build() -> Package` (0299), with `language` and `locale` (`PackageError::{InvalidLanguage, InvalidLocale}`, 0371); ✅ Numbers `litchi_numbers::create::Builder::build() -> Package` (0297), with `table_dimensions`, `language` and `locale` (0370i); ✅ Keynote `litchi_keynote::builder::Builder::build()` returns `Package`, identity required (0298); ✅ DOCX `Builder` (0251), PPTX `builder::Builder` (0127), DOC `Builder` (0188); XLSX has no option-bearing authoring and creates through `Package::new()` (its `create()` duplicate deleted, 0250) · ✅ the unpublished migration host `litchi-iwa`: `PagesDocumentBuilder`, `KeynoteDocumentBuilder`, `NumbersDocumentBuilder` → `litchi_iwa::{pages, keynote, numbers}::Builder`, and every other `Pages*`/`Keynote*`/`Numbers*` type or alias inside those modules loses its redundant prefix (`pages::Editor`, `keynote::SlideInfo`, `numbers::TableCellTextFont`, 168 names); the old names are banned in the new `iwa` profile (0352) |
| iWork document identity (explicit, required) | `Identity` with `Identity::generate()` / fixed bytes | implicit random identity inside builders | ✅ pages `builder::Identity::{generate, from_bytes, from_parts}` with document, version and private UUIDs (`version_uuid`, 0371); `Builder::build()` without one is `PackageError::MissingIdentity`; fixed identities give byte-identical builds (0299) · ✅ Numbers `create::Identity` (`generate`, `from_bytes`, fixed `BLANK`); `Builder::build()` without identity is `create::Error::MissingIdentity` (0297) · ✅ Keynote `litchi_keynote::Identity` (`generate()`, `from_bytes`, fixed `Identity::BLANK`, `Package::identity()`); `Builder::build()` without identity is `builder::Error::MissingIdentity` (0298) |

## 2. Transaction vocabulary (iWork canonical form)

Canonical shape: scoped families live in
`module::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` with package
entry points `Package::{x, edit_x, apply_x}`; the one primary document
transaction is the bare `Package::edit()`.

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Edit entry, primary transaction | `Package::edit()` | `edit_document()` (docx), `opened_presentation_transaction()` (pptx), `Workbook::edit()` (xlsx — moves to `Package`) | ✅ docx eager/source entry name (0018; eager package-bound commit 0022, source 0031) · ✅ iWork/xlsx-name · ✅ pptx eager entry (0116) · ✅ xlsx `Package::edit()` (0250) · ✅ odt/ods/odp `Package::edit()` (0268) |
| Edit entry, scoped family | `Package::edit_x()` | — | ✅ numbers/keynote/pages families; ✅ keynote family types live in focused modules (`slide::text::Edit`, `chart::caption::Edit`, …), prefixed root names deleted (0298) |
| PPTX slide metadata | `Edit::{set_slide_name, set_slide_hidden}` | — | Eager/source primary edits and unified slide `name`/`hidden` setters; eager, source primary and focused source edits accept combined `set_slide_properties` (0162–0163) |
| PPTX animation effect class | `animations::Effect::{preset_class, preset_id}` | `Effect::to_preset` | `preset_class` is the sole class accessor; the deprecated alias is removed from both PPTX profiles (0362ul156) |
| PPTX combined shape properties | `Edit::set_shape_properties`, `ShapePropertiesReplacement` | — | Eager/source primary edits share checked `Position` targets and optional text/bounds in one atomic operation (0153); broader shape properties remain pending |
| PPTX chunked shape publication | `Edit::{set_shape_text_chunks, set_shape_property_chunks}` | — | Eager/source primary and focused source edits admit bounded, globally disjoint chunks together for one same-slide publication (0158); single-chunk setter limits are retained |
| PPTX chunked text-property staging | `opened::Edit::{set_text_property_chunks, set_text_property_chunks_per_edit, set_notes_text_property_chunks, set_notes_text_property_chunks_per_edit}`; `source::Edit::{set_text_property_chunks, set_text_property_chunks_per_edit}`; `table::{Table::cell_text_index, CellTextIndex}`, `shape::owned::Handle::cell_text_index`; `shape::{Common::text_index, Common::read_text_index, BodyTextIndex}` | — | ✅ chunks of at most 256 edits staged in one part pass (0333); `_per_edit` returns each edit's change in order, as one call per edit applied one after the other reports it, and notes chunks stage in one notes-part pass (0348); `CellTextIndex` gives a table's cells' paragraph counts and text-run positions from one pass, each body validated by the text-body reader (bodies alike but for their plain text read once) and a refused one indexed as none, as `Cell::text_body` sees it; `BodyTextIndex` a shape body's as `text_body` reads it (`read_text_index`: one pass over a body already read) (0348) |
| PPT chunked text publication | `Edit::set_shape_text_chunks` | — | At most 16384 globally distinct targets (4096 before 0264) across at most 4096 chunks, each bounded at 256; one native publication and complete readback per call (0220, 0264); ordinary `set_shape_texts` remains bounded at 256 |
| ODP checked scene splices | `Edit::{scene, splice_scene}`, `scene::{Scene, Change}` | — | ✅ paragraph/run/story text, object attributes, text-box/object/page insertion, removal, move and swap as checked `content.xml` source splices on the primary `Package::edit()` (0278) |
| ODG checked structure splices | `edit::Edit::{insert_page, add_page, remove_page, move_page, insert_shape, add_shape, remove_shape, move_shape, move_shape_to_page}` | — | ✅ page and text-rectangle insertion, page and shape removal, and page/shape moves publish as checked `content.xml` source splices, so LibreOffice drawings (non-compact XML) keep every byte outside the touched elements; adjacent markup splices join (0288); ✅ `move_shape_to_page` moves a top-level shape into another page with its bytes unchanged, refusing group members and identity references that would cross pages (0292) |
| ODB element order | `Edit::{move_table, move_column, move_query}` | — | ✅ name-selected table, column and query moves to a final sibling position, relocating the element's bytes within its collection; a self move is an exact no-op (0292) |
| Flat ODF projection and republication | `litchi_odf_common::flat::Parts::{split, to_bytes, merge, merge_package}`; `odt::flat::Document::{to_package, republish}` | — | ✅ a flat document's root children split into package parts with their bytes kept and published as an in-memory package the family's packaged reader and editor work on; an edit of that package republishes into the flat document element by element, keeping every other byte, and refuses edits a flat document cannot hold (changed part roots, added or removed top-level elements, added package members) (0292) |
| ODP retained metadata edit | `Edit::{metadata, set_metadata}` | — | ✅ title/author/subject/description/keywords patched into the retained `meta.xml`; other fields refused (0278) |
| Web Extensions custom-function request | `web::custom_function::Change` | `web::custom_function::Edit` | ✅ semantic input to the durable patch planner, with no independent publication path (0360) |
| XLSX staged handles and source changes | `workbook::edit::{Sheet, Tab, Row, Column, Defaults, FormControl}`, `cell_values::{Change, SheetChange}`, `row_visibility::Change` | `WorksheetEdit`, `TabEdit`, `RowEdit`, `ColumnEdit`, `DefaultsEdit`, `FormControlEdit`, `CellValueEdit`, `SheetCellValueEdit`, `RowVisibilityEdit` | ✅ contextual names; borrowed handles stage in the parent edit and changes are patch inputs (0360) |
| XLSX duplicate family readers and setters | `Snapshot::load`; collection `entries`, `info`, `objects` or `hyperlinks`; `hyperlinks::source::Edit::{put, remove, set}`; `workbook::edit::Sheet::put_hyperlink` | duplicate `read`, `catalog`, `map_info`, `value`, `links`, `surveys`, `ole_objects`, `values`, `put_hyperlink`, `replace_hyperlink`, `remove_hyperlink`, `replace`, `replace_all` in the explicitly scoped gate rows | ✅ 0360: 58 duplicate entries removed, with consumers migrated; bans are scoped to the affected types |
| XLSX drawing and pivot projections | `read_payload`, `svg`, `read_svg_image`, `owner_element().range()`, `server_formats().formats()`, `cached_unique_names().names()` | duplicate `payload`, `svg_descriptor`, `svg_image`, `owner_range`, `formats`, `names` on the affected projection types | ✅ 0360; the underlying typed values and lazy payload readers remain available |
| Shared OOXML duplicate verbs | `custom_xml::Edit::{set_xml, set_properties}`; `custom_data::Limits::with_uid_units`; `mce::StreamError::{raw_callback_error, active_callback_error}`; `process_markup_compatibility_stream` | corresponding `set_item_xml`, `set_props`, `with_max_uid_units`, `raw_error`, `active_error`, `process_markup_compatibility_stream_with_active_observer` | ✅ 0360; both shared profiles enforce the removals |
| Transaction type | `Edit` | `Transaction` (pptx), `WorkbookEdit` (xlsb) | ✅ keynote: every scoped family is a focused `module::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` with unprefixed type definitions, 39 families registered in `keynote-minimal`/`keynote-full` (0298); ✅ pptx primary `opened::Edit` (0116), XLSB structural and cell-value `Edit/Commit/Patch` families, diagnostics and Package commit triads (0165–0171); ✅ XLSB `cell_values`, `cell_values::structure` and `cell_values::worksheet` complete with `Error` and `LimitKind` and registered (0261); ✅ XLS scoped families (`sheet_visibility`, `ole_object`, `real_time_data`, `revision_records`, `user_names`, `picture_compression`, `chart_area`, `web_pub`, `external_link`) (0259); ✅ scoped PPTX families renamed `Edit` (0252); ✅ ODP `Transaction`→`Edit`, ODT/ODS primary `Edit` on `Package` (0268); ✅ every XLSB scoped family (`calculation_chain`, `comments::threaded`, `custom_data`, `data_model`, `external_link`, `connections`, `web_extension_bindings`, `shared_workbook`, `slicer`, `theme`, `timeline`, `volatile_dependencies`, `xml_maps`) renamed `Edit`, `rollback` and the `CustomDataTransaction` alias deleted, `Transaction` banned crate-wide (0343) |
| Commit | `Edit::commit() -> Result<Commit>` | — | ✅ everywhere; shape fixed by the rows below |
| New snapshot accessor | `Commit::package()` / `Commit::into_package()` | `Commit::snapshot()` (docx/pptx), `Commit::{workbook, into_workbook}` (xlsx) | ✅ iWork, DOCX registered primary/family triads, PPTX primary/source, assembled XLSX/XLSB families, ODF and W5 primary families (0022, 0031, 0118, 0355a, 0356a, 0358, 0359, 0360, 0360b, 0362q, 0362u, 0362usd); ✅ DOCX raw/model/helper boundaries are focused raw projections or canonical owner commits, with no detached snapshot family (`0362ra`, `0362rb`, `0362rc`, `0362rf`, `0362rm`) |
| Diagnostics on every commit | `Commit::diagnostics()` | commits without diagnostics (xlsx, pptx) | ✅ DOCX registered primary/family triads, PPTX, assembled XLSX/XLSB families, iWork, ODF and W5 families (0118, 0355a, 0356a, 0358, 0359, 0360, 0360b, 0362q, 0362u, 0362usd); ✅ DOCX raw/model/helper boundaries are focused raw projections and owner commits, with no separate undiagnosed commit family (`0362ra`, `0362rb`, `0362rc`, `0362rf`, `0362rm`) |
| Split commit | `Commit::into_parts() -> (Package, Patch)` | — | ✅ DOCX registered primary/family triads, PPTX, assembled XLSX/XLSB families, ODF/W5 primary families and every iWork scoped family (0022, 0031, 0118, 0268, 0282, 0359, 0360, 0360b, 0361, 0362q, 0362u, 0362usd); ✅ DOCX raw/model/helper boundaries keep publication in the owner triads and detached work in raw trees (`0362ra`, `0362rb`, `0362rc`, `0362rf`, `0362rm`) |
| Patch application | `Package::apply(&patch)`; scoped `Package::apply_x(&patch)` | `publish_document_commit()`, `apply_document_patch()` (docx), `apply_opened_presentation_commit/patch()` (pptx), `Workbook::apply` (xlsx — moves) | ✅ primary and registered-family replay shape for PPT/PPTX, DOCX, assembled XLSX/XLSB, iWork, ODF and W5 (0020, 0030, 0118, 0212, 0268, 0282, 0359, 0360, 0360b, 0362u); content-matching proof is implemented across the unified format consumers (0362m, 0362u); ✅ DOCX raw/model/helper replay uses focused raw patches or canonical owner apply (`0362ra`, `0362rb`, `0362rc`, `0362rf`, `0362rm`) · 🔴 native qualification and final workspace gate remain |
| No-op introspection | `commit.diagnostics().changed()`; `patch.is_noop()` | `Transaction::is_changed()`, `Commit::is_changed()` (pptx), `Patch::is_empty()` (xlsx) | ✅ iWork, DOCX registered primary/family triads, PPTX, assembled XLSX/XLSB, ODF and XLS/W5 scoped families (0118, 0252, 0268, 0359, 0360, 0360b, 0362q, 0362u); ✅ DOCX raw/model/helper boundaries expose focused raw no-op/patch behavior without a public `is_changed` transaction (`0362ra`, `0362rb`, `0362rc`, `0362rf`, `0362rm`) |
| Inverse patch | `Patch::inverse()` | — | ✅ everywhere |
| Rollback | *dropping the `Edit`* | `Transaction::rollback()` (pptx) | ✅ PPT, PPTX, DOCX, XLS, XLSB, ODF, OLE, ograph and crypto source/records show no public rollback; crate-wide bans are recorded by 0252, 0258, 0260, 0268, 0359 and 0360b. Any private transaction helper is outside this public row. |
| Mutable writer on a package | *none — `edit()` is the only write path* | `document_mut()` (docx), `presentation_mut()` (pptx) | ✅ DOCX/PPTX/ODF/W5/XLSX/XLSB owner boundaries (0251, 0356, 0358, 0359, 0360, 0360b); ✅ DOCX raw/model/helper projections use focused raw trees while package mutation remains in owner edits (`0362rc`, `0362rf`, `0362rm`) |

Deferred (own tracks, not this glossary): `History`/undo-redo, patch
composition (`join`, `three_way`), durable patch formats beyond pptx/xlsx.

Pages W1 (record 0299): every focused family lives in its module with the
six canonical names — `section::{text, name, settings, background,
paragraphs}`, `section::pagination::transaction`,
`page_layout::transaction`, `header_footer::text`, `footnote::text`,
`footnote::body::transaction`, `chart::{arrangement, data}`,
`drawable::order`, `image::adjustments`, `metadata`, `document_settings`,
`table::{title, lock}` and `table::{appearance, dimension, headers,
hidden_axes, merge, name, sort}::transaction`; the crate-root prefixed names
(`SectionTextCommit`, `BodyTableNameEdit`, …) and the root aliases
(`BodyTableHeaderSettings`, `BodyTableLockState`, `BodyTableName`,
`BodyTableNameValueError`) are deleted and banned by the `pages` glossary
profile. `Package::{document, document_ref}` replace `semantic_document`;
`Document` keeps no ingress verbs (`open`, `from_bytes`, …), and
`Package::open` reads package directories.

RTF primary `edit::Commit` now exposes `package`, `patch`, `diagnostics`,
`into_package` and `(Package, Patch)` `into_parts` (0224). Primary replay is
`Package::apply`; `Patch::is_noop` and diagnostics use exact artifact equality.
Replaying into an equal-byte owner preserves its retained opening policy.
Records 0227–0228 add checked operation-count and replacement-byte
`SemanticLimits` through `EditOptions`. Replacement admission precedes incoming
payload copies; text uses UTF-8 bytes and pictures use preserved wire bytes.
Logical-tail commits now use the same package accessors and `into_parts`;
replay is `Package::apply_tail_append`, preserves recipient policy, and exposes
`Patch::is_noop` (0229). Record 0230 registers
`tail_append::{Edit, Commit, Patch, Diagnostics, Error, LimitKind}` and
`Package::edit_tail_append`; old prefixed names and duplicate forwarding
methods are removed. Transaction limit errors carry a typed dimension and
observed/maximum counts. Record 0231 replaces unchecked `TailAppendLimits`
with checked `tail_append::SemanticLimits` and `EditOptions` for retained and
source-backed edits. Publication-window policies, source-publication naming
and redaction vocabulary remain pending.

## 2a. Edit, commit and refusal contract (record 0351)

[Record 0351](unified-ops/0351-api-contract.md) is the normative contract for
owner decision 0350 items 14–19. Its foundation is shared code in
`litchi_core::edit`; the format crates migrate onto it in the waves listed
below.

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Immutable owner | no public `&mut self` method on a `Package` or its read-only projection; outputs (`save*`, `to_*`, `write_to*`) take `&self`; signature verbs return a new `Package` | direct `put_*`/`remove_*`/`set_*`/`apply_*` mutators on a package, `custom_props_mut`, `Package::into_workbook` (xlsb) | ✅ DOCX/PPTX/ODF/W5/XLSX/XLSB package and projection owners in the assembled source (0355, 0356, 0358, 0359, 0360, 0360b); ✅ OLE owner/edit publication is commit-only after `39a7a2729b`; ✅ DOCX raw/model/helper projections are non-owner raw trees and canonical owner edits (`0362ra`, `0362rb`, `0362rc`, `0362rf`, `0362rm`) |
| Shared commit triad | `litchi_core::edit::Commit<P, T, D = Diagnostics>` (`new`, `package`, `patch`, `diagnostics`, `into_package`, `into_parts`); a family may alias it as its `Commit` | `Commit::{snapshot, into_snapshot}` | ✅ 0351 (`litchi-core`); families with extra data keep their own `Commit` with the same accessors |
| Minimal diagnostics | `litchi_core::edit::Diagnostics` (`new`, `UNCHANGED`, `CHANGED`, `changed`) | commits without diagnostics | ✅ 0351; families that report more keep their own `Diagnostics` with `changed()` |
| Edit refusal | module `Error::Refused(litchi_core::edit::Refusal)`; `RefusalKind::{UnsupportedOperation, UnsupportedContent, UnsupportedTarget, UnsupportedProperty, InvalidValue, NotFound, Protected}`; litchi-ops classifies by kind alone | `UnsafeEdit` (docx, pptx, xls, ograph) | ✅ 0351 classification and the W1/W2/W5 records; a source scan at this audit head finds no `UnsafeEdit` declaration in the DOCX, PPTX, XLS or ograph crates. Keep this row assigned to the DOCX helper and W4/W4b inventory rechecks; it is not an exception. |
| Discard staged changes | drop the `Edit`; in-place reset is `Edit::reset()` | `rollback` | ✅ public rollback is removed/banned in PPT, ODF, DOCX, OLE, ograph and crypto (0252, 0258, 0260, 0268, 0359); private helpers are not public API. |
| Edit configuration | `Package::edit_with_options(EditOptions)`, `edit_x_with_options(x::EditOptions)`; defaults inherit the limits the package was opened with | `*_with_limits` on edit, apply and output verbs | ✅ W1/W2/W3/W5/W7/W4 migrated profiles use options or required raw limits (0355, 0356, 0357, 0359, 0360, 0360b, 0361); ✅ W6 ODF raw archive/manifest and writer paths now require caller-selected typed limits (`0362ro`, source closure `49b447f673`); ✅ CFB raw ingress and writer limits are canonicalized with required arguments (`0362zad`); ✅ generic soapberry ZIP transport is canonicalized in `0362zae`; the old 0351 count is stale. |
| Options accessors | consuming setter named by the noun (`source_limits(self, ..)`); getter `<noun>_value()` | `get_source_limits`, `get_leniency`, `get_password`, `get_compatibility_profile`; `with_*` setters on option types | ✅ DOCX/PPT/W5/W7/W4 noun setters and `<noun>_value()` accessors (0355, 0357, 0359, 0360, 0360b, 0361); ✅ W6 ODF raw archive/manifest and writer paths use required typed limits (`0362ro`, source closure `49b447f673`); ✅ signature `Policy::{limits, limits_value}` (0362za1); ✅ CFB `SequentialWriterOptions::{limits, limits_value}` (0362zad). |
| Raw CFB ingress, source-layout adoption and storage moves | `OleFile::open(reader, OleFileLimits)`; `SharedOleFile::{open, open_owned, open_owned_vec}(..., SharedOleFileLimits)`; `validate_source(source, SharedOleFileLimits, ValidationLimits)`; `OleWriter::adopt_source_layout(source, OleFileLimits)`; `OleWriter::move_storage(..., StorageMoveLimits)`; `SequentialWriterOptions::{limits, limits_value}`; `OleWriter::create_stream` (the former `update_stream` alias is removed) | public `*_with_limits` methods, limit-less twins and `update_stream` alias | ✅ every public CFB raw ingress and source-layout adoption path requires caller-selected finite typed limits; former default consumers pass their prior defaults explicitly, retained object/PPT profiles flow through adoption, custom bounds remain unchanged, and duplicate writer aliases are removed (0362zad) |
| MS-OVBA raw project and codec boundaries | `litchi_vba::{codec::{encode, decode}, dir::Dir::parse, project::Project::read, build::{Project::finish, Payload::read}}(..., &Limits)`; retained `Payload::write_into` reuses its admission profile; DOC/XLS raw `Reader::vba(..., &Limits)`; PPT `Presentation::vba()` inherits `OpenOptions::vba_project_limits` | `Project::read_with`, default raw project twins, nested CFB defaults, PPT `Presentation::vba_with` | ✅ selected VBA limits are required at every parse/decompress/write boundary; nested CFB directory/allocation ceilings derive from the selected VBA CFB budget without changing outer PPT source limits, and DOCX/XLSX/XLSB plus PPT projections retain their selected profiles (0362ul67/0362ul67a) |
| Generic ZIP transport limits and vocabulary | `soapberry_zip::PreservationIndex::new`, `office::{ArchiveReader::new, IndexedArchive::{from_reader, from_zip_archive, preservation_index}, LazyArchiveReader::new(data, archive_limits, cache_limits), LocatedArchive::index, StreamingArchiveWriter::{new, with_writer, with_writer_and_spool*}}`; `LazyArchiveCacheLimits::new` takes all four cache/flight values and exposes `active_flight_limits` plus `*_value()` getters; owned entries use `ZipArchiveWriter::{start_file_owned, start_file_owned_zip64}` and I/O uses `ErrorKind::IO` | the former limit-less twins, `new_with_cache_limits`, cache `new_with_*_flight_limits`/`with_*_flight_limits`, `max_active_key_bytes`, `*_with_limits`/`with_limits`, `start_entry_owned`, `start_entry_owned_zip64`, and `ErrorKind::Io` | ✅ every tracked ZIP consumer, test, fuzz target, example and performance tool migrated with prior defaults or custom profiles retained; entry aliases and the I/O compatibility variant are deleted, and the canonical names are registered in the corrected Soapberry profile (0362ul86) |
| ODT checked transaction selectors | `transaction::Edit::{remove_rdf_triple_at, move_rdf_triple_to, remove_form_control_at, move_form_control_to, remove_form_at, move_form_to, remove_embedded_chart_at, remove_embedded_object_at, remove_embedded_image_at, move_embedded_object_to, move_embedded_image_to}` | the 11 unchecked `usize` forwarders with the corresponding names without `_at`/`_to` | ✅ checked `Position` selectors retain every operation and the deprecated public forwarders are deleted; the ODT profile bans their reintroduction (0362ul156) |
| Imgconv raw metafile limits | `litchi_imgconv::{pict::PictParser::new, emf::{EmfParser::{new, new_compatible, from_owned, from_owned_compatible, new_lazy}, EmfRecord::parse, EmfRecordRef::{parse_ref, to_owned}, parse_pointl_array, parse_points_array, parse_poly_counts, parse_dib}, wmf::{WmfParser::{new, from_owned}, parser::WmfRecord::parse}}(..., &Limits)` | `pict::PictParser::new_bounded`; public limit-less raw parser/record twins and any `_with_limits` suffix | ✅ input/record/path-array admission uses the caller-selected `litchi_imgconv::Limits` before copying or reserving; fixed-width metafile headers and scalar readers remain unchanged; every imgconv codec/SVG/test consumer passes its retained `Options::limits` or explicit facade default, including compatibility converters (0362ul47) |

Gate rules added by record 0351 (`tools/api_glossary.json`, §11): affix bans
(`"match": "prefix" | "suffix"`), `immutable_owners` and `family_methods`.
`UnsafeEdit`, `rollback`, `Transaction` (exact and as a suffix), `is_changed`,
`into_snapshot`, the `Mutable` prefix and the `_with_limits` suffix are banned
in every profile whose crate already has none; each wave adds the bans for its
crates when it removes the last use.
Record 0362b also inventories the exact name `with_limits`; each migrated
profile bans that name alongside the suffix. Raw and record constructors take
their limits as a required argument.

Record 0362ul47 registers the `litchi-imgconv` raw profile: variable PICT,
EMF, and WMF parser/record boundaries and allocating EMF fragment helpers take
the required `litchi_imgconv::Limits` argument.  `new_bounded` is retired;
there is no public compatibility twin.  The high-level conversion facades may
select their documented default owner policy, but pass that profile into the
raw parser.  Boundary tests cover exact input, record-count, record-byte and
path-array admission; no native or full-workspace qualification is implied.

Record 0362d registers the ODraw picture family: `picture::Package` owns the
record, `View` is read-only, and both edit commit and `Package::apply` return
the shared triad. `Snapshot`/`Owned` publication is removed. Picture limits
enter through options for owners and a required argument for raw projections.
Record 0362v (commit `841e730798`) closes the ODraw raw, error and builder
API migration; only coordinator validation and performance qualification
remain, with no API exception recorded.

Record 0362i names ODraw traversal resources `LimitKind` and image resources
`image::LimitKind`. The former root names `Limit` and `ImageLimit` are deleted
and banned; error variants and their existing observations stay unchanged.

Record 0362ul69 registers the Formula raw codec profile. `mtef::MtefParser`,
`mtef::MtefWriter`, `omml::OmmlParser` and `omml::OmmlWriter` retain the
caller-selected finite `mtef::Limits`/`omml::Limits` policy, and the public
MTEF/OMML conversion helpers require those policies at every raw boundary.
Input, depth, work, item, string and output ceilings are checked before
retention, allocation or emission; the conversion and parser tests preserve
wire and semantic round trips.

### Migration status by wave (record 0351)

Each wave flips only its own lines, so parallel waves do not edit the same
line. A line is ✅ when the wave's gate registrations pass.

| Wave | Crates | Immutable owner | Commit triad / no snapshot | No `Transaction`/`rollback`/`is_changed` | `Refused` replaces `UnsafeEdit` | Limits through options | Gate |
|---|---|---|---|---|---|---|---|
| F0 foundation | `litchi-core`, gate | ➖ | ✅ shared types | ✅ bans on clean crates | ✅ shared types, ops classification | ➖ | ✅ 0351 |
| W1 | DOCX | ✅ (0355; `litchi_docx::Package`, `source_backed::Package`, `Document`) | ✅ (0355, 0355a; 23 families registered; the 13 source-backed families use `edit_x` → `Edit` → `commit` → `Package::apply_x`, with `Package::write_to`/`save` as the only output; the former part-level publishers and `*Publication` results are removed) | ✅ (0355; banned in both `docx` profiles) | ✅ (0355; `Refused`/`PatchConflict`/`Internal`) | ✅ (0355; `SemanticLimits` on `OpenOptions`) | ✅ 0355, 0355a |
| W2 | PPTX | ✅ `Package`, `source::Package` (0356, 0356a, 0356b) | ✅ 47 registered families; source-backed editing on `source::Package` (`edit`, `edit_svg_images`, `edit_cross_slide_copy`), no public raw `OpcPackage` mutator (0356a), no public `SourceBackedPresentationEditor` (0356b) | ✅ (0252) | ✅ (0356) | ✅ (0356) | ✅ (0356, 0356a) |
| W3 | PPT | ✅ | ✅ 12 scoped families publish a `Package` (0357) | ✅ banned crate-wide (0357) | ➖ (typed `slide_order`/`text_edit` refusals gain `kind()`) | ✅ `raw::Reader`/`SourceBackedPackage` `*_with_options`, `_with_limits` banned (0357) | ✅ 0357 |
| W4 | XLSX, XLSB, `litchi-ooxml-common` | ✅ XLSX/XLSB owner and projection source implementation assembled (`0360`, `0360b`; both merge-bases verified) | ✅ XLSX/XLSB scoped commit/accessor implementation assembled (`0360`, `0360b`) | ✅ XLSX/XLSB family bans and replay implementation assembled (`0360`, `0360b`, `0362u`) | ✅ XLSX/XLSB typed refusal/limit implementation assembled (`0360`, `0360b`) | ✅ XLSX/XLSB options and typed `LimitKind` implementation assembled (`0360`, `0360b`) | 🔴 final native/performance qualification and full workspace gate remain |
| W5 | DOC, XLS, RTF, `litchi-ograph`, `litchi-ole-common`, `litchi-crypto` | ✅ (0359; DOC `raw::Package`→`raw::Reader`, RTF `raw::Document`→`raw::Tree`, mutable non-owner codec models under 0351 §2.1 rule 6; DOC/XLS/RTF `Package` and the DOC/XLS/RTF source-backed owners registered, 0359a) | ✅ 12 DOC and 7 XLS scoped families publish a `Package` (0359); codec-level edits alias `litchi_core::edit::Commit<Snapshot, Patch>`; the record-level engines `doc::body_text` and `xls::cell_values` are crate-private behind `Package::edit` and the `Package` routes, and the source-backed publishers are the deferred families DOC `source_backed::paragraph`, XLS `source_backed::{cells, comments, sheet_visibility}` and RTF `source_backed::tail_append` (0359a) | ✅ banned in the doc, xls, rtf, ole-common, crypto and ograph profiles (0359) | ✅ XLS and ograph `Refused`/`PatchConflict`/`Internal` (0359) | ✅ `_with_limits` banned in all six profiles; `<noun>_value()` getters and noun setters (0359); user-defined OLE property overlays take required typed limits and return the shared OLE error (0359h) | ✅ 0359, 0359a, 0359h, 0362rd |
| W6 | ODF (`litchi-odf-common`, ODT, ODS, ODP, ODG, ODC, ODI, ODM, OTH, ODB, ODF-formula) | ✅ (0358; every `Package`, projection and flat owner, ODT `generic::{Package, flat::Document}` and `odc::Document`; `MutableDocument` crate-private behind `Package::edit`, ODS `builder::Workbook` draft) | ✅ (0358; primary, metadata, flat, generic, chart and ODM title/link families registered; part-level drafts are closure arguments of `Package::edit`; 0358a: the source-backed families `ods::source_backed::{cells, dde, sheet_metadata}` and `odp::tail_append` commit a deferred `SourceBackedSpreadsheet`/`SourceBackedPresentation`, written by its `write_to`/`save`) | ✅ (0358; banned crate-wide in all eleven profiles) | ✅ (0358; `Refused`/`PatchConflict`/`Signed`, refusal `kind()`s; 0358a: no string-carrying `Unsupported` refusal left) | ✅ ODF raw archive/manifest and writer paths require caller-selected typed limits (`0362ro`, source closure `49b447f673`; integrated validation and native qualification remain open) | 🔴 source closure is integrated, but the raw/codec profile, integrated validation/native qualification and final workspace gate remain (`0362ro`) |
| W7 | iWork (Pages, Keynote, Numbers; host `litchi-iwa`) | ✅ (host `litchi_iwa::Package`, `Document`; host editors are `Builder` drafts, 0361) | ✅ host primary raw package family `litchi_iwa::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` with `Package::{edit, apply}` (0361la), plus `into_parts` on 93 commits, `family_methods` on `pages`, `numbers-*`, `keynote-*`, six single-operation lifecycle families with an `Edit` (0361) | ✅ | ✅ typed per-family `Refusal` with `kind()` in `Error::Refused`, `Limit` for `LimitExceeded` (0361) | ✅ `MergeReader::*_with_options`, `Workbook::from_sheets_with_options`, `media::OpenOptions`; `_with_limits` banned (0361) | ✅ 0361, 0361la |
| W8 | unified `File` (`litchi-ops`, facade, CLI, Python) | ➖ | ✅ Part A and primary Commit/BatchCommit accessor shape is assembled (0362, 0362u) | ➖ | ✅ DOC body/source-backed refusal classification is assembled (0362) | ✅ Part A options/required raw-limit surface is assembled (0362) | ✅ content-matching replay implementation is assembled across the unified format consumers (0362m, 0362u); 🔴 native/ABBA qualification, signed-native review and the final workspace gate remain |

### Final contract audit boundary (2026-10-24)

This glossary reconciliation is now bound to the fresh all-68 API capture at
`472b6705475b073e7fe08f6e11ee17fb1e29f838`, tree
`ebc1cee05f124c0c63977aff07a23288c19e7d00`, recorded in
[0362ul21](unified-ops/0362ul21-final-api-capture-and-glossary-return.md).
All 68 packages ran fresh with hidden items, all features, offline/locked
Cargo, one job and the 20 GB cap; no capture artifact was reused. The fresh
84-row inventory is SHA-256
`161754e5dbe5b7d149b8631a507e6c1a1b222bf2b780e14f84cb48cdec5039c8`.
The external blank disposition scaffold is explicitly not a review: its 84
`UNASSIGNED_REQUIRED_REVIEW` rows have SHA-256
`57d486415c0e940160a3e0ed29bee08b79b14b013773973a06827e2b17445529`.
The reviewed 84-row disposition remains the 0362ul14 file (SHA-256
`3b0d48c38a4cc53e319c720d5e5ef4a1378631b10de5ec13190b03f5d3ad1a3e`), and
the independent 472b receipt (SHA-256
`6bd44bf07454ce459f9b0cfb27c91a0718bd12a2b9bc02b08a011e3c3ed2b2f3`)
revalidates those roles against the fresh source. This is API evidence only;
it does not turn the all-feature JSON into a feature-specific gate. A red cell
below is assigned required engineering work, not an accepted exception:

- W4 XLSX and W4b XLSB source implementation is assembled: both
  `git merge-base --is-ancestor 3c55c62986 HEAD` and
  `git merge-base --is-ancestor b2e21c0ab4 HEAD` return success. Their final
  native/performance qualification and workspace gate remain open.
- DOCX primary and registered-family triads, including the earlier 0362q
  return-contract work, are integrated. The raw/model/helper boundary closure
  is also integrated as focused raw projections and canonical owner edits
  (`0362ra`, `0362rb`, `0362rc`, `0362rf`, `0362rm`); the old 0362q inventory
  and the earlier `uo2/w8-docx-document` assignment are historical. Runtime,
  allocation and performance qualification remain separate coordinator work.
- W8 content-matching replay is implemented for the unified format consumers
  (`0362m`, `0362u`), including PPTX cross-copy, ODP tail-append and XLS/RTF
  source-backed paths. Native/ABBA qualification and the final workspace gate
  remain open; this is not a performance acceptance claim.
- Signed-native replay is integrated at this source boundary: the original
  implementation (`2b594e1757`) and the reviewed follow-ups (`7dc296655c`,
  `77bc2d3ac2`, `df2d948477`) are ancestors of the composed target. Native/
  ABBA qualification and the final workspace gate remain open; this is not a
  performance acceptance claim. OLE's commit-only/fixed-error correction is
  already integrated (`39a7a2729b`, `1a5f56f45d`, `02c9f44c78`).
- 0362rd's separate CFB and generic ZIP substrate ownership is historical. The
  CFB rows are canonicalized and consumer-closed by 0362zad, and the generic
  ZIP rows and cache-policy aliases are canonicalized with every tracked
  consumer migrated by 0362zae/0362zaea. The corrected Soapberry profile
  passes in the fresh OPC/DOCX validation receipt below; no generic ZIP API
  debt is implied by the old separate-follow-up wording. Any public format
  re-export or foreign surface remains assigned to its owning format lane.
- The old OPC `*_with_limits`/alias findings in 0362rd are closed in the
  assembled source. 0362ul19 records fresh feature-specific `opc-minimal` and
  `opc-full` rustdoc/glossary passes, alongside `docx-minimal`, `docx-full`
  and the corrected `soapberry-zip` profile. These checks are source/API
  evidence, not native qualification or the final workspace gate.
- W6's package owners and public ODF raw/codec limit paths are source-closed:
  `litchi-odf-common::package::Archive::new(data, ArchiveLimits)`,
  `Archive::read_manifest_xml(XmlStreamLimits)`, and the bounded
  `PackageWriter` constructors are recorded by `0362ro` and integrated by
  `49b447f673`. The raw/codec profile, integrated validation, native
  qualification and ODF owner/performance work remain open and must not be
  treated as acceptance of those results.

The 0362j record closes the 0357b PPT first-commit/replay regression for its
measured workload and preserves its native measurements; it is not a bounded
residual exception or final new native qualification. The current PPT native
qualification remains open. XLS concrete-iterator/deferred-reopen and ODraw
API residuals are closed by `0362y`/`904bf4df8c`/`91a578db22` and
`0362v`/`841e730798`, respectively; their coordinator qualification is not
recast as an API exception. Item 22's redaction audit is the only permitted
audit deferral.

## 3. Open, create, save

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Open from path | `open(path)` | — | ✅ everywhere |
| Open from bytes (copy allowed) | `from_bytes(&[u8])` | `from_bytes(Vec<u8>)` signatures (xlsx/odt/odp), `from_slice` (xlsx), `from_vec` (pptx, removed 0252) | ✅ DOCX owning/lazy Package (0038/0040); ✅ XLSX `Package` (0250; copies into the owned source archive); ✅ PPTX (0252); ✅ XLSB `Package` (0261; copies into the owned source archive; `from_slice*` deleted) · ✅ ODT/ODS/ODP `Package::from_bytes(&[u8])` (0268) · ✅ OTH/ODM/ODG/ODC/ODI/ODB/ODF-formula `Package::from_bytes(&[u8])` (0282); ✅ flat owners `odt::flat::Document`, `ods::FlatSpreadsheet`, `odg::FlatDrawing`, `odc::FlatChart`, `odp::FlatPresentation` and `odp` flat `Snapshot` take `from_bytes(&[u8])` (0343; the gate checks names, not signatures, so this row is verified by compilation) |
| Open from shared bytes (zero-copy) | `from_shared_bytes(Arc<[u8]>)` | — | ✅ numbers; ✅ pages `Package::from_shared_bytes{,_with_options}` (0299); offered only where zero-copy is real |
| Open from reader | `from_reader(R)` | — | ✅ docx/xlsx; ✅ XLSB `Package`/`SourceBackedWorkbook` (0261); extended per wave |
| Configured open | `open_with_options(path, OpenOptions)`, `from_bytes_with_options(bytes, OpenOptions)` | `open_with_limits`, `open_with` (doc), `open_with_options` under other type names | envelope decided in §5 · ✅ DOCX (0038) · ✅ XLSX `OpenOptions` (source limits, leniency, encryption limits; 0250) · ✅ PPTX owning `Package` and positional-source presentation owners (0252) · ✅ XLSB `OpenOptions` (source, external-link and cache limits, leniency, drawing decoding, execution context) on `Package`, the `Workbook` projection constructors and `SourceBackedWorkbook` (0261) · ✅ DOC raw owner and part editors, `open_with_limits*` banned crate-wide (0258) · ✅ PPT `raw::Reader::{open_with_options,from_reader_with_options,from_ole_file_with_options}`; no public `*_with_limits` remains (0357; source audit at this head) · ✅ PPT `text_edit::SourceSnapshot::open_with_options` and `slide_order::Snapshot::from_bytes_with_options` (0260) · ✅ keynote `Package`, `Presentation` and `MergeReader` `open_with_options`/`from_bytes_with_options` (`open_with_limits`/`from_bytes_with_limits` deleted, 0298) · ✅ ODT/ODS/ODP `open_with_options`/`from_bytes_with_options`/`from_reader_with_options` with the `litchi_odf_common::OpenOptions` envelope (source + semantic limits, 0268) · ✅ the other seven ODF `Package` owners `open_with_options`/`from_bytes_with_options`/`from_reader_with_options` (0282) · ✅ pages `Package::{open_with_options, from_bytes_with_options, from_shared_bytes_with_options}`; `open_with_limits`/`from_bytes_with_limits` banned (0299) · ✅ Numbers `Package::{open_with_options, from_bytes_with_options, from_shared_bytes_with_options}` and the `Workbook` projection (0297) · ✅ assembled XLSX/XLSB W4/W4b ingress source surface (0360, 0360b); final qualification remains |
| Minimal fresh artifact | `Package::new()` | `Workbook::create()` duplicate (xlsx) | ✅ docx; ✅ canonical PPTX constructor (0127); ✅ XLSB (0175); ✅ XLS (0199); ✅ RTF (0225); ✅ XLSX (0250); ✅ pages (deterministic built-in template, 0299); ✅ Numbers (0297; byte-identical, fixed `Identity::BLANK`); ✅ Keynote, deterministic with the fixed `Identity::BLANK` (0298) · all formats gain it |
| Save to path | `save(path)` — **always atomic** (sibling temp + rename) | non-atomic `Builder::save` (odt/odp) | ✅ OOXML, including lazy DOCX (0041), and RTF (0226) · ✅ ODF: every ODT/ODS/ODP/ODC/formula/generic whole-package save publishes through `litchi_core::atomic` with failed-save CRUD evidence (0268); `Package::{save, save_with_durability}` on all seven remaining ODF owners (0282) |
| Save with explicit durability | `save_with_durability(path, Durability)` | — | ✅ OOXML and RTF (0226); extended everywhere |
| Bytes output | `to_bytes()` | `Builder::build() -> Vec<u8>` as output (ODF) | ✅ owning DOCX (0044); lazy managed DOCX bytes pending. `build()` returns `Package`; output lives on `Package` |
| Stream output | `write_to(impl Write)` | `to_stream`, `write_to_stream` | ✅ owning/lazy DOCX Package (0041–0042); ✅ DOCX source tail-append `Plan`/`Commit` (0251); ✅ OPC `OpcPackage::write_to`, `PackageWriter::write_to`, `SourceArtifact::write_to{,_with_accounting}` and `SourcePartSplicePlan::write_to{,_with_accounting,_with_expected_artifact}`, the `to_stream`/`write_to_stream*` names banned in the `opc` profiles; PPTX/XLSX/XLSB export neither name (0343) |
| Open encrypted | `open_with_password(path, password)` | — | ✅ where encryption exists; ODF keeps same name |
| Save encrypted / re-encrypted | `save_encrypted(path, pw, Mode)` / `save_reencrypted(path, pw)` | — | ✅ OOXML |
| Explicit plain output | `save_plain(path)` / `to_plain_bytes()` | — | ✅ OOXML, including owning DOCX bytes (0044); the only plain output of an encrypted source |
| Query source encryption | `encryption() -> Option<Mode>` | — | ✅ OOXML |

## 4. Limits and configuration

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Physical ingress limits | `SourceLimits` | `ReadLimits` (OPC re-exports in docx/pptx/xlsx/xlsb), `SourcePackageLimits` (odt), `package::Limits` (doc), `RecordLimits` (ppt), `DocumentSourceLimits` (iWork) | ✅ keynote (0298) `SourceLimits` (input, entries, entry/aggregate/component bytes, archive objects/messages/message bytes/header fields) · ✅ OPC `SourceLimits`/`SourceLimitsBuilder` and OOXML `SourceLimits` re-exports, record 0035 · ✅ DOC (0192); ✅ XLS eager/retained policies (0196–0198); ✅ PPT (0210); ✅ ODF `litchi_odf_common::SourceLimits` (`SourcePackageLimits` deleted, 0268); ✅ pages `SourceLimits` (input bytes, members, entry/aggregate/component bytes; `DocumentSourceLimits` and the `Limits` re-export deleted, 0299); ✅ Numbers `litchi_numbers::SourceLimits` (`DocumentSourceLimits`, `PackageLimits` deleted; adds the format-specific `header_fields` dimension; default = hard ceilings, 0297) |
| Semantic limits | `SemanticLimits` | `DocumentLimits` (numbers), `DocumentSemanticLimits` (keynote) | ✅ keynote (0298), `DocumentSemanticLimits` merged into `SemanticLimits`; ✅ pages; ✅ RTF primary operation count and replacement bytes (0227–0228); ✅ XLSB nested `data_model`, `calculation_chain`, `volatile_dependencies` and `xml_maps` `ReadLimits` and their `Limits` aliases renamed `SemanticLimits` (0261) · ✅ ODT/ODS/ODP `SemanticLimits` (content/styles bytes, units; 0268) · ✅ pages builder form, `SemanticLimitsError` deleted (0299) · ✅ Numbers one `SemanticLimits` for the package and the `Workbook` projection (`DocumentLimits`, `PackageSemanticLimits` deleted, 0297) · ✅ W1 complete; `DocumentLimits` and `DocumentSemanticLimits` are banned in the core, numbers and keynote profiles and no format crate exports either name (0343; `xml_minifier::audit::DocumentLimits` is an XML-audit profile, not a document semantic limit). Per-dimension naming of the other families' semantic limits is the *Dimension naming* row |
| Open configuration envelope | `OpenOptions` (per format, extensible fields) | `PackageOpenOptions` (doc — merges), `ReadOptions` (numbers), `iwork::Options` (facade; deleted with `litchi::iwork`, 0306) | ✅ keynote (0298) `OpenOptions { source_limits, semantic_limits }` on `Package`, `Presentation` and `MergeReader`; ✅ DOCX ingress envelope (0038–0040); ✅ XLSB (0261, `OpenPolicy` refusals for owner-specific policies); ✅ pages `OpenOptions` (`DocumentReadOptions` deleted, 0299); ✅ Numbers `litchi_numbers::OpenOptions` on `Package`, `Workbook` and `MergeReader` ingress (`ReadOptions`, `DocumentReadOptions` deleted, 0297); ✅ assembled W4 XLSX/XLSB option envelopes (0360, 0360b); native qualification remains |
| Signature policy configuration | `litchi_sign::Policy::limits(Limits)`; `Policy::limits_value()` | `Policy::with_limits`; a `limits()` getter with no `_value` suffix | ✅ immutable verification policy uses the noun setter and `<noun>_value()` accessor (0362za1) |
| Raw CFB signature editor | `litchi_sign::cfb::Editor::open(bytes, format, Limits)` | `Editor::with_limits`; default `Editor::open(bytes, format)` | ✅ one raw constructor with caller-selected limits; no limit-less twin (0362za1) |
| Construction idiom | consuming builder, fallible `max_<dimension>` setters, fallible `build()` | positional `new(...)` constructors (iWork) | ✅ keynote (0298) `SourceLimits::builder()`/`SemanticLimits::builder()`; ✅ OPC, DOC (0192), XLS (0196) and PPT (0210) · ✅ pages `SourceLimits::builder()`/`SemanticLimits::builder()` (0299) · ✅ Numbers `SourceLimits::builder()`, `SemanticLimits::builder()` (0297) |
| Build-time invalid limit | `Error::InvalidLimit { kind, value, maximum }` | standalone `*LimitsError` types (doc, iWork) | ✅ keynote (0298) `Error::InvalidLimit { kind: LimitKind, value, maximum }`; ✅ OPC `InvalidLimit` with contextual maximum (0036); DOC (0192), XLS (0196), PPT (0210); pages `PackageError::InvalidLimit` (0299); ✅ Numbers `limits::Error::InvalidLimit { kind: LimitKind, .. }` (`SemanticLimitsError`, `DocumentSourceLimitsError`, `DocumentLimitsError` deleted, 0297); other owners pending |
| Ingress limit exceeded | `Error::Limit { kind, observed, maximum }` | `OpcError::ReadLimit`, facade `ResourceLimit` mapping drift | ✅ OPC (0036), DOC (0192), XLS (0196), PPT (0210) kind/observed/maximum; ✅ PPTX root/codec dimensions (0362ul84); ✅ XLSB cell-values transaction bounds (`PackageError::Limit`, 0261); broader taxonomy pending |
| Limit kind enums | `SourceLimitKind` / `SemanticLimitKind` | `ReadResource`, per-crate one-off kinds | ✅ keynote (0298) `SourceLimitKind`/`SemanticLimitKind` with `as_str()` nouns and `LimitKind::{Source, Semantic}`; ✅ OPC/OOXML (0036), DOC (0192), XLS (0196), PPT (0210) `SourceLimitKind`; ✅ PPTX root/codec dimensions (0362ul84); ✅ XLSB `cell_values::LimitKind` (0261); other semantic kinds pending |
| Dimension naming | nouns (`input_bytes`, `members`, `entry_bytes`, `aggregate_bytes`, `depth`, `events`, …); setters `max_<noun>` | a second name for an existing dimension | ✅ OPC common nouns (0037); shared defaults remain 512 MiB input and XML depth 256; other families pending |
| Coverage policy | caller-selected limits on **every** public ingress path | — | ✅ PPT package/record ingress (0210), XLS eager ingress (0196); ✅ ODT/ODS/ODP `Package`, source-backed and catalog ingress via `OpenOptions` with typed `Error::{Limit, InvalidLimit}` (0268); ✅ OTH/ODM/ODG/ODC/ODI/ODB/ODF-formula `Package` ingress, `from_prepared_package` included (0282); ✅ ODF flat owners `flat::{Document, Workbook, Presentation, Drawing, Chart, Image}` and `generic::flat::Document` take `OpenOptions`; the ODC chart `Limits` fold into `SemanticLimits` (0358) |

### RTF source dimensions

RTF uses checked `SourceLimits`/`SourceLimitsBuilder`, `SourceLimitKind`, and
`Error::{InvalidLimit, Limit}` for eight ingress dimensions (record 0222):
input bytes, lexer events, individual and aggregate binary bytes, decompressed
bytes, opaque node count, individual opaque node bytes and aggregate opaque
bytes. Input defaults to 512 MiB; other existing finite defaults are retained.
Zero explicitly denies a dimension. Values above `isize::MAX` are invalid.
`OpenOptions::{source_limits, leniency}` carries the profile and explicit repair
policy, including through unified `OpenOptions::rtf_options`.

Raw validation is explicit: `validation::{validate, validate_str}` and
`ValidationReport::{from_bytes, from_str}` require a caller-selected
`read::OpenOptions`; no default validation twins remain. Transport encoding is
similarly bounded: `transport::compress(data, compress, CompressionLimits)`
checks the complete framed output before each fallible growth and final frame
allocation. `transport::decompress` retains its required
`DecompressionLimits` argument.

This migration does not rename the separate semantic edit limits or structural
`LimitExceeded` errors. Record 0223 adds the concrete immutable `Package` and owned read-only
`Document` projection. Record 0224 aligns primary transaction vocabulary, and
0225 adds native `Builder`, `Package::new()` and unified RTF creation. Record
0226 adds native atomic publication and unified save/stream dispatch. Records
0227–0231 migrate the focused transaction families. Records 0232–0240 add
native-backed paragraph reads, text replacement, removal, move, insertion, swap,
alignment and layout through the unified API. Native primary body paragraph edit
selectors now consistently use `Position`, including
`ParagraphTextReplacement`, `ParagraphAlignmentUpdate`,
`ParagraphLayoutUpdate`, `TransferPlan::plain_paragraph`, and the paragraph
lifecycle/text-property methods;
the duplicate `split_paragraph_at` spelling is removed. Other retained-story
selectors and semantic limit vocabulary still need migration. Richer
element/property coverage and performance qualification remain pending.

### PPT source dimensions

PPT uses checked `SourceLimits`/`SourceLimitsBuilder` and typed
`Error::InvalidLimit`/`Error::Limit` with `SourceLimitKind` (record 0210).
`max_input_bytes` bounds the complete CFB input, `max_entry_bytes` each declared
stream, and `max_aggregate_bytes` the sum of all declared stream lengths,
including metadata and unrelated streams. The root mini-stream backing is
excluded to avoid counting its child payloads twice. Admission precedes stream
payload reads, including adoption of an already parsed CFB owner.

Complete-input, entry and aggregate defaults are 512 MiB, 512 MiB and 2 GiB;
record depth defaults to 256. PPT also bounds encoded record bytes, record
events, individual record payload bytes and aggregate record payload bytes.
`to_builder()` preserves all selected dimensions when changing one ceiling.
`OpenOptions` exposes `password`, `leniency`, `source_limits` and corresponding
`get_*` accessors; its debug output redacts credentials. Package projections
retain the stricter combination of admitted and explicitly requested limits.
The immutable primary owner and shared owned projection are implemented in
0211. The canonical primary `package::edit` family and `Package::{edit, apply}`
are implemented in 0212. Record 0213 adds canonical `Builder::build()` and
`Package::new()`; publication belongs to `Package`, and `raw::encode` is the
explicit codec escape hatch. Encrypted authoring consumes caller-supplied
one-use `EncryptionEntropy`, including on failed admission. Unified PPT
creation/publication is integrated; navigation and mutation remain open.
Record 0214 adds `package::edit::Error::Signed`: signed primary transactions
admit exact no-ops and refuse changed writes before native encoding. Read-only
inspection retains selected source limits and leniency; exact no-op patches
preserve the complete original signed artifact.

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

### ODF semantic dimensions

`litchi_odf_common::SemanticLimits` (in `OpenOptions::semantic_limits`)
bounds the projections every ODF crate parses; `SemanticLimitKind::as_str()`
is the noun and `Error::Limit { kind: LimitKind::Semantic(..), .. }` the
exceeded error (record 0358). A flat (single-file) owner is bounded by the
smaller of the source `input_bytes` and `content_bytes`. Each default is the
most permissive bound an ODF reader enforced before the dimension was shared;
a caller tightens it.

| Dimension | Default | Ceiling | Meaning |
|---|---|---|---|
| `content_bytes` | 256 MiB | 256 MiB | `content.xml` (and a flat document) bytes |
| `styles_bytes` | 256 MiB | 256 MiB | `styles.xml` bytes |
| `units` | 1,048,576 | 1,048,576 | semantic units (paragraphs, cells, shapes) of one projection |
| `depth` | 4,096 | 4,096 | XML nesting depth |
| `text_bytes` | 16 MiB | 16 MiB | one text run, paragraph or map text |
| `edits` | 1,024 | 1,024 | staged edits of one flat transaction |
| `shapes` | 1,000,000 | 1,000,000 | drawing shapes |
| `frames` | 1,000,000 | 1,000,000 | image frames |
| `map_areas` | 100,000 | 100,000 | image-map areas |
| `output_bytes` | 256 MiB | 256 MiB | serialized output of one commit |
| `axes` | 16,384 | 16,777,216 | chart axes |
| `series` | 65,536 | 16,777,216 | chart series |
| `data_points` | 16,777,216 | 16,777,216 | chart data points |
| `cached_rows` | 16,777,216 | 16,777,216 | rows of a chart's cached table |
| `cached_cells` | 16,777,216 | 16,777,216 | cells of a chart's cached table |
| `range_items` | 65,536 | 16,777,216 | items of one cell-range list |
| `resources` | 100,000 | 16,777,216 | package resources one projection references |
| `history` | 4,096 | 16,777,216 | retained history steps |
| `scalar_bytes` | 64 KiB | 16 MiB | one scalar value (attribute, cached cell text) |

## 5. Errors

- Each focused module owns its non-exhaustive `Error`; crate roots keep only
  ingress/egress error types. God enums (`xlsx::Error`, …) split per module,
  staged per wave. Boundary mapping is typed; string flattening
  (`map_ooxml_error` pattern) is retired.
- Error extensibility is part of that boundary: public format `Error` enums
  are `#[non_exhaustive]`, and downstream matches retain their current typed
  classifications with a fixed fallback for future variants. The final sweep
  closes the 18 assigned format rows in 0362uvb; the concurrent fonts
  obfuscation row remains assigned to its own lane until its declaration and
  consumers land.
- Shared variant taxonomy (payload shapes fixed):

| Concept | Canonical variant | Banned |
|---|---|---|
| Recognized but wrong family | `UnexpectedFormat { detected: FileFormat }` | `NotPages` / `NotKeynote` / `NotNumbers`, string messages (`"Not an ODT file: …"`) (✅ keynote (0298): `ReadError` and `presentation::Error`; ✅ pages `PackageError::UnexpectedFormat`, unrecognized input `PackageError::NotOfficeFile`; `NotPages` banned, 0299; ✅ Numbers package and workbook `UnexpectedFormat`, 0297) |
| Unrecognized input | `NotOfficeFile` | — |
| Source file changed after open | `SourceChanged { expected, observed }` | — |
| Edit refused (well-formed request this route cannot perform losslessly) | `Refused(litchi_core::edit::Refusal)`, classified by `RefusalKind` | `UnsafeEdit` (docx, pptx, xls, ograph); string-carrying refusal variants (record 0351, §2a) |
| Patch base mismatch | `PatchConflict { .. }` | `StaleSource` (merges; the two concepts stay distinct; ✅ PPTX 0252, ✅ DOCX `document::Error`, the source-backed `story_text`, `paragraph_copy` and `tail_append` errors (0343); banned in the `docx` profiles) |
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
  ⏸ OOXML/legacy call-site audit (0343 D9).

## 6. Metadata

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Semantic metadata type | `Metadata` (merged field union; keywords as a list; real time types) | `ooxml_common::Props`; the `creator` field name (→ `author`) | ✅ one `litchi_core::Metadata`; `Props` deleted and banned in the ooxml-common profiles; `author` is the only spelling (`dc:creator` stays in the codec); `Timestamp` replaces the `*_local`/`*_time` pairs (`created`, `modified`, `last_printed`, `last_backup`); OOXML-only `category`/`content_status`/`version`/`last_printed` are fields (0253) · ✅ `Metadata.keywords: Vec<String>` (empty = none, skipped when serialized) in every format, with the JSON views, node views (`ValueType::List`, `node::Value::List`), the CLI and the `litchi-py` dictionaries writing arrays; string formats split at `;`/`,` and join with `, ` (`litchi_core::metadata::keywords::{split, join}`), ODF stores one `meta:keyword` each, and the OOXML codec keeps `Keywords` mixed content losslessly for an unchanged list (owner decision 0350 overrides 0343 D7; 0354) |
| Unstorable metadata field | typed `metadata::Error::Refused(Refusal)` with `kind() == UnsupportedProperty` and a fixed reason | silently dropping a field | ✅ OOXML transaction (0253): `content_type`, `last_backup`, `codepage` |
| Metadata read | `Package::metadata()` | `Package::props()` (docx) | ✅ facade/ODF/iWork reads · ✅ pptx/xlsx/xlsb `Package::metadata()` (0253) · ✅ docx `Package::metadata() -> Option<&Metadata>` (0251) |
| Metadata mutation | `Package::{edit_metadata, apply_metadata}` + `metadata::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | `put_props()`, `props_mut()`, `clear_props()` | ✅ pptx/xlsx/xlsb `Package` and the shared `litchi_ooxml_common::properties` engine; families registered in the gate (0253) · ✅ XLSB source owner: `SourceBackedWorkbook::edit_metadata` with `metadata::source::{Edit, Commit}`, published over the retained source through `properties::Patch::{compose, check_source_backed, prepare_source_backed}` (0261) · ✅ doc/xls/ppt `package::metadata` (0254) · ✅ docx `metadata` family over the shared core codec, registered in `docx-minimal`/`docx-full` (0251) · ✅ odt `Package::{edit_metadata, apply_metadata}` and `litchi_odt::metadata` (splices `meta.xml` children, no clock), registered in the `odt` profile (0276); ✅ docx app/custom staging: `metadata::Edit::{app_props, set_app_props, custom_props, set_custom_props, remove_custom_props}`, written by the shared engine's exact part patch, and `Package::app_props()` (0342, 0362ac) · ✅ ods/odp/oth/odm/odg/odc/odi/odb/odf-formula `Package::{edit_metadata, apply_metadata}` and `litchi_<crate>::metadata` over the shared `litchi_odf_common::metadata` engine (element splices of `meta.xml`, or a new `meta.xml` when absent; no clock; signed/encrypted refused), registered in each profile; odt's family uses the same engine (0288) · ✅ ODS keeps the primary `document::Edit::set_metadata` beside the family (0277 M1); its former `metadata::Commit` is now `metadata::XmlCommit` |

| Custom properties | `custom_props()` / `edit_custom_props()` | — | ✅ name (docx/pptx/xlsx); OOXML custom properties are staged through `metadata::Edit::set_custom_props`, while `remove_custom_props()` explicitly removes the custom-properties graph (an equal empty assignment remains an exact no-op) (0253, 0362ac) |
| Application properties (OOXML-only) | `app_props()` | — | ✅ `litchi_ooxml_common::properties::app` (splice writer), pptx/xlsx/xlsb `Package::app_props()` (0253), docx `Package::app_props()` (0342) |
| Fresh-authoring metadata | `Builder::set_metadata(Metadata)` | `Builder::set_property` (doc/xls/ppt) | ✅ ODF · ✅ doc/xls/ppt (0254); not an opened-document mutation · OOXML creation commits a metadata `Edit` on the fresh package (0253) |
| iWork metadata write | typed unsupported until native gates pass | — | ✅ pages `Package::{metadata, edit_metadata, apply_metadata}` with `metadata::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`: title/author/keywords/description in `Metadata/Properties.plist`, read back and verified; the Pages app does not display these keys, so no native acceptance is claimed (0299) · ✅ keynote/numbers: a root metadata `set` is a typed refusal and leaves the file unchanged (`iwork_unified::iwork_metadata_write`), and Numbers creation with initial metadata is refused (0343) |

## 7. Selectors and coordinates

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Catalog selector | `Selector<'a> = Name(Cow<'a, str>) \| Position(Position)` | the `Id` variant (raw IDs live in raw modules); `Index(usize)` (numbers); `At`/`Index` selector forms (pages footnotes) | ✅ core/pages/keynote shape · ✅ XLS `cell_values`/`sheet_visibility` `Position(Position)` (0259) · ✅ pages `ImageSelector`/`BodyChartSelector` `Position(Position)`, footnote `Selector::{Anchor, Position}` (0299) · ✅ numbers `SheetSelector`/`TableSelector`/`ChartSelector::Position(Position)`, `#[non_exhaustive]` (0297) |
| Positional selection word | `Position` — checked, zero-based value type | `Index` as a selector variant; raw `usize` in public signatures (facade `slide(usize)`, `paragraph(usize)`) | ✅ the facade's family views `DocumentView::{paragraph, paragraph_text}` and `PresentationView::slide`, the PPTX handle's `slide`, and the compat table `row_at`/`cell_at`/`grid_span_at` take `impl Into<Position>` (`From<usize>` keeps call sites terse); a rustdoc-JSON scan of `litchi-ops` finds no other positional `usize` parameter (counts, limits and capacities remain `usize`) (0343) |
| Selector-taking methods | `impl Into<Selector>` parameter, `Result<Option<_>>` return; missing → `Ok(None)`, ambiguous/invalid → typed error, never panics, no `Index` trait | method-level splits that bypass the selector channel | pptx `scene::get/at` converge (W2) |
| Grid coordinates (not selectors) | `At` / `Area` with A1 parsing at the edge | treating coordinates as catalog selectors; numbers `Position`/`Range` coordinate aliases | ✅ xlsx; rule recorded · ✅ numbers `Position`/`Range` aliases (root, `table`, facade `litchi::numbers`) deleted (0300) · ✅ iWork: `litchi_iwa_common::table::coordinate::{CellPosition, CellRange}` → `At`/`Area` (re-exported by numbers, keynote `slide::table::cells` and the facade `litchi::numbers`), the Keynote-local `number_format::CellPosition`/`CellPositionError` → `At`/`AtError` and the transaction error variant `CellPositionNotFound` → `CellNotFound`, the `table::model` `Position`/`Range` aliases deleted; every workspace caller migrated, the old names banned in the `iwa-common`, `numbers-*`, `keynote-*`, `pages` and `facade-iwork` profiles (0352) |
| Name identity policy | follows the producing application; fresh authoring defaults to Unicode caseless | silently mixing case policies | ✅ behavior unchanged; per-format registration below |

Registered name-identity policies: xlsx sheet names — Unicode caseless
(Excel semantics); xls sheet names — the same Unicode caseless key (NFD,
default case fold, NFD) for the reader's duplicate check, selector lookup,
rename and tab insertion (`litchi_xls::sheet_name`, record 0259); docx glossary entries — NFD/case-fold/NFD (Word
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
`banned_names`, and `families`, and optionally `delegated_globs`: the crates
whose glob re-exports (`pub use litchi_pages::*` in the facade) the profile
leaves to those crates' own profiles (record 0306); any other glob without
local module data is refused. Feature combinations are separate profiles;
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

Record 0351 adds three rule kinds. A ban may carry `"match": "prefix"` or
`"match": "suffix"` (default `"exact"`): the name then matches as the start or
end of a public name, never the bare affix itself (`_with_limits` bans
`open_with_limits`; `Mutable` bans `MutableDocument`). Optional
`immutable_owners` lists fully qualified owner types (a format's `Package`
and its read-only projections); any public method on them, inherent or from
a trait impl, that takes `&mut self` fails, and a listed owner that is not
exported fails rather than passing. Optional `family_methods: true` extends
the six-name family check to methods: `Commit` must expose `package`,
`into_package`, `patch`, `diagnostics` and `into_parts` and no `snapshot`,
`into_snapshot` or `is_changed`; `Diagnostics` must expose `changed`; `Edit`
must expose `commit` and no `rollback`, `is_changed` or `snapshot`; `Patch`
must expose `is_noop` and `inverse`. A member that is another crate's type
(`pub type Commit = litchi_core::edit::Commit<..>`, a re-exported shared
family type) is checked by that crate: rustdoc does not list a dependency's
inherent methods.

The physical iWork substrate and archive profiles (`iwa-core` and
`iwa-archive`, record 0361d) also enforce the raw-operation limit vocabulary.
Archive parse, mutation, serialization and catalog ingress require limits on
their canonical methods; Snappy decoding and encoding require `SnappyLimits`.
There are no default-limits overloads or `*_with_limits` aliases. The physical
archive policy is `litchi_iwa_core::Limits`; its former `ArchiveLimits` and
`ResourceLimits` aliases are removed. The archive's catalogs and frozen source
projections are registered as immutable owners.

The `iwa-text-wire` profile (record 0361g) applies the same rule to raw text
projection, validation, decoding, boundary scans and rewrites. `from_bytes`
requires `Limits`; the other raw operations require `RewriteLimits`, including
the `with_behavior` variants. Their former `_with_limits` and `_and_limits`
names are removed, with no default-policy twin. The storage-scoped text
engines of record 0374 follow it too: `paragraph_style::{effective, plan,
uniform_style, repoint, stylesheet_with_variation,
stylesheet_without_variation}` take a required `WireLimits` and are pinned
as `required_signatures`. An engine owns payloads only; a set/reset run of
one host property is one `paragraph_style::Change::{Set(Property),
Reset(PropertyKind)}`, and its `Plan` (`Rewrite`, `Derive`, `Collapse`) is
published by the format package's own family. Named paragraph styles
(`paragraph_style::named`: `catalog`, `applied`, `plan_apply`,
`named_copy`, `renamed`, `redefine` and the stylesheet and theme
rewrites) and highlights and text comments (`annotation`: `annotations`,
`annotation`, `graph` and the highlight and comment-storage codecs) follow
the same rule; a refused catalog request is
`paragraph_style::Error::Refused(Refusal)`.

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

Standalone XML editing uses `document::raw::Tree` with required caller-selected
`document::raw::Limits`; it emits bytes and reversible patches without a raw
snapshot-publishing commit. Package-bound eager and source-backed edits both
publish the canonical `document::Commit<P>`. Bounded planning lives under
`document::composition`: `Prepared` and `Composition` retain recoverable join
and three-way state, while `Package::apply_composition` performs the one
owner validation/publication pass. Owner-created compositions bind the complete
artifact, so equal-content reopened packages are accepted and a changed
untouched part is refused; a raw XML composition retains XML-only scope.

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

Record 0362ul94sa adds two explicit OPC publication seams for format-owned
opaque XML: `OpcPackage::try_add_authored_xml_part_bytes(name, content_type,
bytes, limits)` and `OpcPackage::try_replace_authored_xml_part_bytes(partname,
expected, replacement, limits)`. Both require `SourceLimits`, admit part,
aggregate-part, XML-event/depth, and XML-attribute ceilings before mutation,
and retain exact bytes only as publication provenance. They never manufacture
an `OwnedXmlPart` or `SourceXmlPart` range-splice token; the strict
source-splice namespace proof remains unchanged. Equal replacement bytes are
an audited no-op, including on a signed package, and every failed add or
replace leaves the package and its signature state untouched.

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

History is exposed by `document::composition::History` and records the
canonical owner commit. Source-backed publication retains deferred output and
source freshness checks. Record 0032 supplies typed `LimitKind`; raw XML
admission additionally bounds depth, element count, attributes and attribute
bytes before layout retention.


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
`apply(&patch)` paths. XML-only callers use `document::raw::Tree`; composition
planning and publication use `document::composition` and an owner route.

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

`File::{edit, edit_with_options}` use `transaction::Edit`, replacing the draft
proxy name `Transaction`. `transaction::{Commit, Patch, Diagnostics, Error,
LimitKind}` name their result/replay family; `SemanticLimits` and its consuming
builder bound mutation calls and retained main XML, configured through
`EditOptions` (record 0362). `Commit<T>` exposes the file, patch, diagnostics
and closure value through the triad accessors `file`, `into_file`, `patch`,
`diagnostics`, `value`, `into_value` and `into_parts` (record 0362).
`Diagnostics::changed()` and `Patch::is_noop()` describe the final logical
artifact, not how many operations ran. Other counts remain explicit facts.

Patches are process-local. The original immutable owner and its clones use a
constant-work identity path; a separately opened same-variant file is accepted
when its published base passes the bounded length-and-digest proof. A changed
publication, format/security mismatch, stale source, signed input or locked
input remains refused. `Patch::inverse()` reverses endpoint direction;
`Patch::apply(&File)` returns the retained validated endpoint after source
guards. `Error::PatchConflict` remains distinct from physical stale-source
failures. `capabilities::Verb::Edit` advertises this DOCX closure subset;
command `Batch` execution stays pending.

## Unified command execution (record 0076)

`File::apply` executes one command and returns `transaction::Commit<Applied>`;
`File::{batch, batch_with_options}` return `BatchCommit` with the accessors
`file`, `into_file`, `patch`, `diagnostics`, `value`, `into_value` and
`into_parts` (record 0362). `ItemResult::{Applied, Failed}` preserves input
order; failures retain their original item index. `Applied` carries `changed`
and `path: Option<Box<Path<'static>>>`, populated by creation. The optional
path is indirect so ordinary result entries do not embed the grammar's large
inline path storage; callers borrow it with `as_deref()`.

Atomic and best-effort modes share one transaction and resource profile.
Ordinary best-effort failures roll back one item, while global resource,
source/protection/cancellation failures abort the whole batch. The DOCX core
command subset is discoverable through `Verb::Batch`; contextual gaps retain
follow-ups for command families and text/dump semantics not yet implemented.

## Unified dump and raw layer (record 0336)

| Concept | Canonical | Banned | Notes / status |
|---|---|---|---|
| Document as a replayable batch | `File::dump() -> batch::Dump` | `dump_batch`, `to_batch` | ✅ OfficeCLI v2 items (`meta` with `dumpVersion: 2` first); `Dump::{items, into_items, warnings}`; skipped content is one `unsupported_element` (or `unsupported_property`) warning per node, never silent. DOCX/ODT/RTF/DOC, XLSX/XLSB/ODS/XLS and PPTX (record 0336); other formats refuse `unsupported_capability` naming the format |
| Package part listing / read | `File::parts() -> Vec<raw::PartInfo>`, `File::raw(part) -> raw::Part` | `raw_part`, `read_part`, `get_part` on `File` | ✅ OPC parts (DOCX, XLSX, PPTX, XLSB) and packaged ODF members, by part name or OfficeCLI alias; read from the published package (record 0336) |
| Part addition | `File::add_part(parent, raw::NewPart) -> raw::AddedPart` | `insert_part`, `create_part` | ✅ `NewPart::{of_type, new, named}`; `AddedPart::{file, part, relationship_id}`; XML payloads pass the OPC publication audit (record 0336) |
| XPath-addressed XML edit | `File::raw_set(part, xpath, raw::Action, xml) -> raw::Edited` | `raw_edit`, `xpath_set` | ✅ location-path subset; `Action::{Append, Prepend, InsertBefore, InsertAfter, Replace, Remove, SetAttr}` (`FromStr` takes OfficeCLI's spellings); spliced, then audited through `try_replace_owned_xml_part_bytes` (record 0336) |
| Structural validation | `File::validate() -> raw::Validation` | `check`, `lint`, `verify` | ✅ `Validation::{format, report, is_valid, errors}` over `litchi_core::ValidateReport`; `raw::Error` when a validator cannot complete; OPC part-graph checks are `litchi_opc::validate_parts` (record 0336) |
| Capability verbs | `Verb::{Dump, Raw, Validate, AddPart, RawSet}` (`dump`, `raw`, `validate`, `add-part`, `raw-set`) | — | ✅ record 0336; `Verb::ALL` has 23 verbs |

## XLSX unified creation and publication (record 0095)

`create(FileFormat::Xlsx)` adopts the native `Package::new` owner directly
(record 0250 replaced the removed `Workbook::create`).
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

Transactions retain native patches and immutable endpoint owners. Replay uses
the input identity fast path; a same-variant reopened publication is accepted
after the unified bounded content proof, and inverse replay returns the
retained source. Native content comparison detects final no-ops without ZIP
serialization.
Structural commands, command coalescing, OfficeCLI string inference/extended
clear modes, and encrypted-owner edits remain explicit capability gaps.

Record 0098 adds rectangle and scoped-selector cell assignments. Ranges include
blank coordinates; formulas broadcast without relative translation. All
selected cells stage in one native commit. `SemanticLimits` now admits
`max_expanded_targets` and `max_staged_bytes` for XLSX fanout, defaulting to
65,536 distinct cells and a 16 MiB payload/action charge. Caller-selected
profiles flow through `File::edit_with_options` and batch execution. Unknown
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

Record [0204](unified-ops/0204-xls-eager-query-cursors.md) adds complete-catalog
`Workbook::sheet_count()` and selector-based `Workbook::sheet()`, distinct from
ordinary worksheet ordinals. Eager cell-only queries now consume the native
iterator without a facade grid. Source-backed and axis projections still need
bounded cursor work before full read-performance acceptance.

Record [0205](unified-ops/0205-xls-unified-writes.md) connects unified XLS scalar
`set`, clear/null, sheet names and generic edit/batch/history to the same
canonical native `Package::edit` family. Scalar record replacements preserve
XF styles; native SST closure owns string-table publication. Formula-cell
replacement and unrepresented dependency/inverse closures remain explicit
element-level gaps. No parallel mutable workbook model or transaction family
is introduced.

## XLS unified P0 and W4 residue (record 0259)

Unified XLS adds sheet add/remove/move/swap, row and column height, width,
hidden state and empty insertion/removal, formula writes with `formula` and
`cachedValue` reads, cell alignment/border/fill writes through `XF` records,
and read-only `namedrange` nodes. Every write is a native `Package::edit`
operation: `Edit::{reorder_sheets, insert_sheet, remove_sheet, set_formula,
set_row_height, set_row_hidden, set_column_width, set_column_hidden,
set_cell_format}` with `cell_values::FormatChange`. Sheet-name identity uses
the registered caseless key.

W4 residue: the scoped XLS families name their transaction `Edit`; the
`Transaction` names, `UserNamesTransaction`/`UserNamesPackageTransaction`,
`rollback`, `is_changed`, `Patch::is_empty` and
`sheet_visibility::Snapshot::transaction` are removed and banned in the
`xls-minimal` profile. The `cell_values` and `sheet_visibility` selectors
carry a checked `litchi_core::Position`. `property_set::Transaction` remains
because the record 0254 metadata owner uses it. See
[0259](unified-ops/0259-xls-unified-p0.md).

## Unified PPT navigation (0215)

The PPT adapter now supports unified `get`, `find`, `query` and semantic views
using native `Presentation` and `Slide::shared_shapes` projections.
`Presentation::check_source` fences cached positional reads, including reads
after `retain_source(false)`. `ShapeContainer` accepts concrete `ShapeEnum`
children; `Shape::clone_box` is removed. Unified nodes retain shared native
arrays and index trails. Native identities require raw paths or queries;
unimplemented stories and properties remain explicit capability gaps.
Unified PPT mutations and full API/ADR migration remain open.

## Unified PPT setters (0216)

`File::set`, explicit edits and batches route top-level PPT shape text/anchor
changes and slide visibility through canonical `Package::edit` and its
`Commit`/`Patch` family. Native source-order targets are distinct from physical
navigation ordinals. The facade adopts the committed package directly and keeps
the native proof in snapshot-bound history. Anchor inputs round to native master
units; unchanged bounds preserve the original anchor encoding. Signed, encrypted,
released and stale owners keep explicit refusals. Native visibility-record
insertion, nested writes and consecutive-command coalescing remain open.

## Unified PPT slide structure (0217)

`File::{remove,move_to,swap}` and explicit edits now use the canonical native
PPT primary transaction for slide structure. Positional destinations follow the
shared post-detachment contract; removal selectors resolve before descending
native staging. A swap composes at most two native slide moves in one edit.
Exact self moves/swaps preserve snapshot identity. Slide and shape insertion,
shape removal, and consecutive-command coalescing remain open.

## PPT P0 read and write surface (0260)

The native leg adds `package::edit::Edit::{insert_slide, set_notes_text}` (and
the same methods on `slide_order::Transaction`), named after the PPTX
`Edit::insert_slide` verb; `insert_slide(position)` creates a blank-layout
slide. Shapes expose `ShapeEnum::{name, description, source_index, id}`
(L-PPT-R3/R4) and slide-list placeholder text is read and written through
`OutlineTextBody` (L-PPT-R1). `text_edit::Refusal::NoNotesPage` names the
missing-notes refusal. The unified PPT adapter reads `/slide[N]/notes`,
`.../paragraph[N]`, `name`, `alt` and raw `sourceIndex`, sets notes and
paragraph `text`, and adds slides. Removed and banned in both PPT profiles:
`text_edit::SourceSnapshot::open_with_limits`, `rollback` on
`text_edit::{Transaction, SourceTransaction}` and `slide_order::Transaction`,
and `slide_order::Snapshot::from_bytes_with_limits`. `ppt-full` now also
bans `Builder::set_property` and checks the `package::metadata` family, as
`ppt-minimal` already did. See [0260](unified-ops/0260-ppt-p0-surface.md).

## Unified PPT shape runs (0218)

Consecutive direct shape setters use one canonical native PPT `Edit` and retain
its `Commit`/`Patch` proof. Reads, selectors, repeated targets and structural
commands publish the preceding run before resolving against its new snapshot.
Native operation and semantic retention limits remain enforced across run
boundaries. This does not add a mutable presentation model or change public
operation names.

## RTF unified borrowed run axis (0241)

RTF unified nodes expose the canonical `Node::text` projection for both
paragraphs and character runs; the former paragraph-specific text accessor is
removed. Run nodes retain borrowed native `litchi_rtf::text::Run` values and
provide `as_run()`, `run_position()`, and their containing
`paragraph_position()`. Paragraph children are lazy native `runs()` cursors,
so structural break tokens do not become run nodes and empty paragraphs keep
their zero-child behavior.

The `/body/p[N]/r[N]` axis accepts both `r` and `run` spellings and projects
read-only local text, font, size, character flags, foreground color, and each
native underline variant as a distinct token. Run properties and the run
element support `get` and `query`; `set`, `add`, `remove`, `move`, and `swap`
remain explicit capability gaps until the native atomic run editor exists.

## RTF run character transactions (0242)

`edit::RunFormatPatch` carries omitted or explicit local character facets;
`RunFormatUpdate` pairs it with typed source paragraph/run `Position` selectors.
`Edit::patch_body_run_formats` admits one ordered batch atomically, and
`patch_run_format` is the single-target entry point. Compound facets share one
native operation per run. These operations use the canonical primary
`Edit`/`Commit`/`Patch` lifecycle and immutable `Package` endpoints.

The unified run registry now marks the ten integrated character properties as
settable. Run text, font/color resource edits and structural run verbs retain
explicit gaps. No facade-owned parallel run model or property-by-property
publication is introduced.

### RTF run text replacement (0243)

`edit::RunTextReplacement` carries typed paragraph/run `Position` selectors,
owned replacement text and an optional `RunFormatPatch`. Primary `Edit` exposes
`replace_run_text` and atomic ordered `replace_body_run_texts`; text and supplied
character facets publish together as one operation per target. The new semantic
change retains before/after text for exact source-bound patch history; durable
replay remains an explicit gap. Unicode may change native run fragmentation,
while text and local formatting coverage remain invariant outside selected text.
The unified run registry now advertises `text` as writable alongside the ten
character facets; resource and structural run gaps remain explicit.

### Unified RTF run removal (0244)

Run `remove` uses typed paragraph/run positions and the existing native
`Edit::replace_body_run_texts` transaction with empty replacements. No parallel
public native transaction family is introduced. Removing all runs preserves
paragraphs; paragraph removal remains a distinct structural action. The shared
facade admission pass now charges source text retained by both run replacement
and removal before native staging, with native patch weight governing history.

### RTF opaque run-edit anchors (0245)

The existing native run transaction now projects retained opaque body anchors
through text replacements and removals using semantic UTF-8 byte positions.
The emitted terminal paragraph delimiter participates in projection. No public
transaction alias or facade-specific rewrite API is introduced; exact source
bytes, candidate validation and source-bound inverse history remain native-owned.

### RTF run insertion (0246)

`RunTextInsertion` carries typed paragraph and run-boundary `Position` values.
`Edit::insert_run` and `Edit::insert_body_runs` publish through the native owner;
source-relative batch boundaries are nondecreasing and equal positions preserve
input order. The unified `add` verb supplies placement and property coercion,
returning the canonical path of the whole inserted run. Formatting inherits the
right neighbor, last left neighbor at append, or the empty paragraph delimiter.

### RTF direct-set batching (0247)

Consecutive compatible facade setters use existing native ordered batch methods
and the canonical native `Edit`/`Commit`/`Patch` family. No public alias or second
editable document model is introduced. Temporary source indexes are bounded;
staging windows close at their limits, while scalar admission stays available
when the optimization index cannot fit. Source-changing boundaries publish
before resolving later positional paths.

### RTF destination stories (0275)

Native retained-destination text edits (`set_table_cell_text`,
`set_header_footer_text`, `set_note_text`, `set_annotation_text`) keep their
names and the canonical `Edit`/`Commit`/`Patch` family; publication now splices
the replacement into recorded source provenance instead of reserializing the
document. New read accessors follow existing vocabulary: `text::Paragraph::
text_range`, `ParagraphFormat::{style, list_override, list_level}`,
`Format::character_style`, and `Package`/`Document::picture_body_offset`
(a UTF-8 text offset, not a `Position` selector). No alias is introduced.

### Unified RTF stories (0275)

The RTF adapter projects `/body/tbl[N]/tr[M]/tc[K]`, `/header[N]` and
`/footer[N]` (with `p[M]` and `[@type=default|first|even|odd]`),
`/footnote[N]`, `/endnote[N]`, `/comments/comment[N]`,
`/body/p[N]/hyperlink[M]`, `/body/p[N]/picture[M]`, `/styles/<name>` and
`/numbering/{abstractNum,num}[N]` (with `level[N]`) through the canonical Word
vocabulary; native identities (`commentId`, list `id`, `ilvl`) stay raw-path
only. Paragraphs read `style`, `numId`, `numLevel`; runs read `rStyle`. Cell,
header/footer, note and comment `text` is writable through the native
destination editor, pictures are removable; every other new element or
property is a typed capability gap. Unscoped `p` queries stay in the main
story. No public alias or second document model is introduced.

## Gate hardening (record 0249)

Record 0249 extends the registered gate without changing any public API. Every
added rule matches code that is already conformant; banned names that still
exist are listed below as wave residue, not as failing rules.

- **Families.** `litchi_docx::document` is registered in `docx-minimal` and
  `docx-full`; it is the only OOXML module exporting all six family types.
  **Pending, not registered:** pptx `opened`, `source` and `change_tracking`
  (no `Error`, `LimitKind`); the scoped xlsx families
  (`workbook::edit` is registered by record 0250) (no `Diagnostics`, `Error`, `LimitKind`, or no `Edit`); docx
  `content_control`, `ink`, `section::layout`, `revision::conflict` and the
  `source_backed::*` edit modules. Each registers in the change that completes it.
- **pptx-full** now carries the nine primary-transaction bans `pptx-minimal`
  already had (`opened::Transaction`, `opened_presentation_transaction`,
  `apply_opened_presentation_{commit,patch}`, `opened::Edit::{rollback,
  is_changed}`, `opened::Commit::{snapshot, is_changed}`,
  `opened::Patch::is_empty`).
- **ppt-full** builds every PPT feature with the `ppt-minimal` rules. XLS, RTF
  and core declare no Cargo features, so their single profile is already the
  full surface; no duplicate profile is added.
- **core-minimal** bans the §4 configuration names that `litchi-core` must not
  grow (`ReadLimits`, `ReadLimitsBuilder`, `ReadResource`,
  `SourcePackageLimits`, `RecordLimits`, `DocumentSourceLimits`,
  `DocumentLimits`, `DocumentSemanticLimits`, `PackageOpenOptions`,
  `ReadOptions`, `LimitsError`). §4 bans `Limits` only as DOC's
  `package::Limits` (already banned in the DOC profiles); core's execution
  budget `litchi_core::Limits` is not a banned name.
- **ooxml-common-minimal/full** register the transaction, open/output and
  configuration names banned in §2–§4 that `litchi-ooxml-common` does not expose.
- **Facade.** A `litchi` profile cannot be activated: even the featureless
  facade re-exports dependency crates through public globs
  (`common::*` → `litchi_core::*`), whose children rustdoc JSON omits, so the
  checker refuses the profile (record 0017). It needs dependency export data.

Wave residue that still exposes a glossary-banned name (owner: record 0248):

| Crate | Public name | Canonical | Owner |
|---|---|---|---|
| ooxml-common | `custom_data::Limits` | `SemanticLimits` | follow-up of record 0253 (re-exported by xlsx/xlsb, which uo2/p4 and step 7d own) |
| core | `Limits`, `Error::ResourceLimit` (execution budgets) | not banned by §4; rename only on an owner decision | owner decision |
| pptx, xlsx, opc, ops | names listed in record 0248 §2 | §2–§4 | Phase B steps 4–5, ops follow-ups |
| xlsb | `Package::into_workbook`; mutating `Workbook` verbs (`edit_opc`, `put_*`, `set_*`, `sign`); scoped `x_with_limits`/`edit_x_with_limits` read and edit entries | `Package::workbook()` projection; `edit_x()`/`apply_x()` families | record 0261 §8 |

## XLSB W2 residue and P0 writes (record 0261)

- Ingress: `litchi_xlsb::OpenOptions` (`source_limits`, `external_link_limits`,
  `leniency`, `parse_drawings`, `cache_limits`, `execution_context`) on
  `Package::{open, from_bytes, from_reader, from_opc}_with_options`,
  `Workbook::{new, from_opc_package}_with_options` and
  `SourceBackedWorkbook::{from_reader, from_read_at}_with_options`. A policy
  an owner cannot honour is `PackageError::UnsupportedOpenPolicy { policy:
  OpenPolicy }`. `Package::from_bytes` borrows `&[u8]`.
- Edits: `Package::edit_with_options(structure::EditOptions)` and
  `Package::edit_cell_values_with_options(selector, cell_values::EditOptions)`.
- Renames: `raw::Writer` → `raw::Encoder`; nested `ReadLimits`/`Limits` →
  `SemanticLimits` (data_model, calculation_chain, volatile_dependencies,
  xml_maps). `cell_values::LimitKind` with `PackageError::Limit`.
- Writes: `cell_values::Edit::{set_format(reference, &Format),
  set_formula_text}` (`Format`, `BorderLine` are partial changes; unnamed
  attributes keep their records); `structure::Edit::{merge_cells,
  unmerge_cells, set_defined_name, remove_defined_name}` (durable tags 11 and
  12). Reads: `named_ranges::Entry` via `{Workbook,
  SourceBackedWorkbook}::defined_name_entries`.
- Source-owner metadata: `SourceBackedWorkbook::{metadata, app_props,
  edit_metadata, write_to, has_metadata_edits}` with
  `metadata::source::{Edit, Commit}`.

## Legacy OLE metadata (record 0254)

DOC, XLS and PPT author and edit document properties through the shared
`litchi_ole_common::property_set::SummaryProperties` model
(`SummaryInformation`, `DocumentSummaryInformation` and the user-defined
section). Fresh authoring is `Builder::set_metadata(Metadata)` plus
`Builder::{summary_properties, summary_properties_mut}` for typed and custom
properties; the refusing (DOC) or silently dropping (PPT) `set_property` is
deleted and banned. Opened packages expose `Package::{metadata,
edit_metadata, apply_metadata}` and the registered
`package::metadata::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`
family; `LimitKind` has no caller-selected dimension yet. The one
`litchi_core::Metadata` conversion is `SummaryProperties::{from_metadata,
merge_metadata}`; fields without a property-set slot (`identifier`,
`hyperlink_base`, unzoned timestamps) are refused with
`MetadataError::Unsupported`, never dropped. The unified layer accepts
initial metadata in `create_with` and `set("/metadata", ..)` with
root-metadata keys and `custom.<name>`. See
[0254](unified-ops/0254-legacy-ole-metadata.md).

The inert `_PID_LINKBASE` and `_PID_HLINKS` overlay uses
`litchi_ole_common::property_set::user_defined::{Properties, Edit}` with a
required `Limits` argument and the shared `litchi_ole_common::Error`; its
blob, link-count and UTF-16 caps retain distinct `LimitKind` dimensions.
The former `with_limits` constructors are removed and banned (0359h).

## XLSX W2 owner reshape (record 0250)

`xlsx::Package` is the immutable owner of every state verb: `new`, `open`,
`from_bytes(&[u8])`, `from_reader`, the `*_with_options(.., OpenOptions)`
ingress, `edit`, `apply`, the save/bytes/stream outputs and the security
verbs. `xlsx::Workbook` is the owned read-only projection
(`Package::workbook()`, borrowed `workbook_ref()`). `workbook::edit` is a
registered family: `Commit::{package, patch, diagnostics, into_package,
into_parts}`, `Patch::{is_noop, operation_count}`, `Diagnostics::changed`,
`LimitKind` with `Error::Limit { kind, observed, maximum }`; `Error` is still
the crate error (module split pending). Removed: `Package::create`,
`Workbook::create`, `from_slice*`, `from_bytes(Vec)`, `*_with_limits` and
`*_with_leniency` ingress, `Package::into_workbook`, `Commit::{workbook,
into_workbook}`, `Patch::{is_empty, len}`, `Workbook::unsigned` (use
`Package::unsign`) and the pivot `Workbook{Transaction,Patch,Commit}` types,
now `pivot::{server_formats, table_data, cached_unique_names}::{Edit, Patch,
Commit, Diagnostics, Error, LimitKind}` with `Package::{edit_pivot_table,
edit_pivot_table_data(_with_options), edit_pivot_cache(cache, field)}` and
`Package::apply_pivot_*` (0360).
Scoped `&mut` families and `SourceBackedWorkbook` ingress names remain W2
residue.

## PPTX ingress options and patch conflict (0252)

`litchi_pptx::OpenOptions` is the one ingress envelope: `source_limits`,
`cache_limits`, `execution_context` and, with `encryption`,
`encryption_limits`. The owning `Package` gains `open_with_options`,
`from_reader_with_options`, `from_bytes_with_options` and the
`*_with_password_and_options` routes; it refuses cache and execution policies
with the typed `Error::UnsupportedOpenPolicy { policy: OpenPolicy }` before
reading input. `Package::{open_with_limits, from_reader_with_limits,
from_bytes_with_limits, from_vec, from_vec_with_limits,
from_vec_with_password, from_vec_with_password_and_limits,
*_with_password_and_limits}` are deleted. `SourceBackedPresentation` and
`SourceBackedPresentationEditor` replace their six `open_with_*` combinations
with `open_with_options`; their positional `from_path_*`/`from_read_at_*`
constructors remain a pending migration. `Error::StaleSource` is renamed
`Error::PatchConflict`. Both PPTX profiles ban the removed names, and
pptx-full now carries the primary-transaction bans pptx-minimal had.
See [record 0252](unified-ops/0252-pptx-w2-residue-and-p0-writes.md).

## PPTX primary family completion and rich text edits (0252)

`litchi_pptx::opened` now exports the complete family
`{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` and is registered in
both PPTX glossary profiles. `opened::Error` is the crate `Error` re-exported
(a focused split of the PPTX god-enum is deferred). The crate-wide typed
boundary uses `Error::Limit { kind, observed, maximum }` and
`Error::InvalidLimit { kind, value, maximum }`; `opened::LimitKind` and
`opened::CodecLimitKind` re-export the canonical dimensions (0362ul84).

Rich formatting of existing text bodies uses one verb on every primary edit:
`opened::Edit::set_text_properties(slide, &[TextEdit])`,
`source::Edit::set_text_properties(slide, &[TextEdit])` (the focused
`SourceBackedSlideEdit` is crate-private since record 0356b). A `TextEdit`
pairs a `TextTarget` (`Shape`, `Paragraph`, `Run`, `Cell`, all checked
`Position`s) with `TextProperties` built by `with_*` setters whose
`Option<Option<_>>` arguments distinguish omission, removal and replacement.
`table::Cell::fill()` reads a cell's local RGB fill. No parallel transaction
type or mutable writer is introduced.

## PPTX scoped family vocabulary (0252)

The scoped PPTX edit families (notes, embedded content parts/OLE/controls,
presentation-properties metadata tracks/changes/structure/custom shows/
guides/events/colour map/designer tags, read-only recommendation, change
tracking, ink actions, shape designer/classification, placeholder type
extensions, comment collaboration, source-backed slide/SVG/order edits) now
use the canonical names: `Transaction` types are `Edit`
(`AuthorPresenceTransaction`/`CommentThreadingTransaction` are
`AuthorPresenceEdit`/`CommentThreadingEdit`), `rollback()` is deleted
(dropping the edit is the rollback), edits and commits expose
`diagnostics() -> litchi_pptx::diagnostics::Diagnostics` with `changed()`
(families that already had their own diagnostics keep them), and patches
expose `is_noop()`; public `is_changed()`, `changed()` aliases and patch
`is_empty()` are removed. Both PPTX profiles ban `Transaction`,
`AuthorPresenceTransaction`, `CommentThreadingTransaction`, `rollback` and
`is_changed` crate-wide. `presentation_mut` was already removed (0127).
These families remain snapshot-scoped (`Snapshot::edit()`); moving them to
`Package::{x, edit_x, apply_x}` entries with full `Error`/`LimitKind`
families is a follow-up.


## PPTX notes, comments, pictures and layout slides (0252)

`opened::Edit::set_notes_text` now creates the slide's notes page from the
presentation's single notes master when the slide has none (transitional
decks; other cases are typed `UnsafeEdit` refusals). Legacy comments gain
positional verbs on the same edit: `add_comment_text(slide, author,
initials, text, location, date_time)` (creates the named author when
absent; the caller supplies any timestamp, the editor reads no clock) and
`set_comment_at(slide, Position, text, location)`. Slide insertion on both
primary edits instantiates the layout's placeholders (title, body, object,
… with type, orientation, size and index; date, footer, slide-number,
header and slide-image placeholders are not instantiated).
`TextTarget::Shape` accepts text, so an empty placeholder body can be filled.
No parallel transaction type is introduced.

## OOXML metadata transaction (record 0253)

`litchi_core::Metadata` is the one metadata type. Its timestamps are
`metadata::Timestamp` (`Utc` instant or `Unzoned` wall clock), `metadata::Field`
names a field in errors and diagnostics, and `metadata::Keywords` is the
lossless keyword model shared with the OOXML codec.
`litchi_ooxml_common::properties::{Snapshot, Edit, Commit, Patch, Change,
Diagnostics, Error, LimitKind, Family}` is the exact, reversible engine over
the core, application and custom parts; `properties::app::Props` is the
application-properties model. pptx, xlsx and xlsb expose
`Package::{metadata, app_props, edit_metadata, apply_metadata}`, and each
re-exports
`metadata::{Patch, Diagnostics, Error, LimitKind}` beside its owner-bound
`metadata::{Edit, Commit}` (`Commit::{package, into_package, patch,
diagnostics, into_parts}`). The unified layer accepts initial metadata in
`create_with` and `set("/metadata", ..)` with record 0254's 22
root-metadata keys, `custom.<name>` and `null` removal; `/metadata` stays
set-only for every format.

## DOC P0 surface and W4 residue (record 0258)

Positional DOC readers take `Position`, never raw `usize`:
`Document::{body_paragraph, body_table, footnote, endnote, comment,
header_footer_slot, section_header_footer}` and `StoryBlocks::{get,
get_kind, paragraph, table, element, paragraph_text}`; `HeaderFooterSlot`
reports its sections as `Position`. The primary `doc::Edit` gains
`set_character_value(TextTarget, CharacterValue)`, `insert_paragraph`,
`remove_paragraph` and `move_paragraph`; `TextTarget::BodySpan` addresses
one run. Part-level ingress is `open(bytes)`, `open_with_policy` and
`open_with_options(bytes, &EditOptions)` (revision limits plus protection
policy); `SourceSnapshot` keeps `open_with_options(source,
SourceBackedOptions)`; the raw owner's configured open is
`raw::Package::open_with_options(path, OpenOptions)`; smart tags adopt a
bounded `Snapshot::parse_with_options_and_policy` result through
`Editor::from_snapshot`. `open_with_limits`, `open_with_limits_and_policy`
and every `rollback` are removed and banned in both DOC profiles (dropping a
transaction is the rollback). See
[0258](unified-ops/0258-doc-p0-surface.md).

## Registry schema export (record 0263)

`litchi_ops::schema(format) -> schema::Schema` is the one export of the
property registry (`litchi::schema` in the facade, `litchi schema <type>` in
the CLI, `litchi_py.schema(format)` in Python). `Schema::{format,
registered, elements, element, gaps}` reuses the capability report's
`Element`/`Property`/`Gap` types; it is not a second element model.
`capabilities::Property` gains `aliases()`, `constraint()`, `clearable()`
and `write_gap()`, and `capabilities::Constraint` (`None`, `Int`, `Number`,
`Length`, `Chars`, `NotEmpty`) names the registry's value domain. The
format-level report stays `Capabilities::of`; the file-level one stays
`File::capabilities`. No alias names are added. See
[0263](unified-ops/0263-registry-schema-export.md).

## DOCX metadata transaction, Builder and stream output (0251)

`litchi_docx::metadata::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`
is the core-properties family: `Package::metadata()` reads the typed value
(`None` is an absent part), `Package::edit_metadata()` stages `set`/`update`/
`clear`, `Edit::commit()` returns a new package, and
`Package::apply_metadata(&patch)` replays a patch whose base matches. A
changed write to a signed package is `metadata::Error::Signed` at commit; an
equal value is an exact no-op that keeps the signature. The family wraps the
shared `litchi_ooxml_common::properties` codec; its value type is
`litchi_core::Metadata`, re-exported as `metadata::Metadata` (0253).
`Package::{props, props_mut, put_props, clear_props}` are deleted.

`Package::document_mut()` is deleted. Fresh authoring uses
`Builder::{new, new_template, document, package, font_embedding, build}`;
`build()` materializes the authored document into the package graph without a
ZIP round trip. An opened package changes only through `edit()` and its scoped
families; it never hydrates into the legacy writer.

The source-backed tail-append `Plan::write_to_stream` and
`Commit::write_to_stream` (both tail-append families) are renamed `write_to`.
Both DOCX profiles ban the six retired names and register the
`litchi_docx::metadata` family.

## DOCX element paths (0251a)

`document::{ElementPath, ElementStep, ElementAnchor, NewTable}` address and
create main-story elements beyond the direct body: a path is the direct-child
steps from `w:body`, each a `Position` among same-named direct children
(`Paragraph`, `Run`, `Hyperlink`, `Table`, `Row`, `Cell`, `ContentControl`).
The element verbs carry an `_at` suffix on the existing `Edit`
(`format_paragraph_at`, `format_run_at`, `replace_text_at`,
`replace_cell_text_at`, `insert_run_at`, `insert_paragraph_at`,
`insert_table_at`, `insert_row_at`, `remove_element_at`, `move_element_at`)
and record `Operation::{ReplaceElement, InsertElement, RemoveElement}`. No
second transaction family is introduced.

## DOCX package-graph changes, hyperlinks, comments, pictures and lists (0251b–0251e)

`document::GraphChange` is the package-graph part of a main-document patch:
relationships of the main part, new related parts, and exact part
replacements. It is staged by verbs on the eager package-bound `Edit`
(`add_external_relationship`, `add_related_part`, `replace_part`, with the
readers `part_payload` and `related_part`) and travels in the same `Patch`;
no second transaction family is introduced. `document::HyperlinkTarget`
(`Url`, `Anchor`) names a hyperlink's target for `insert_hyperlink_at` and
`set_hyperlink_target_at`. `document::NewComment` (`new`, `with_initials`,
`with_date`) seeds `add_comment_at`; `remove_comment(id)` takes the native
comment identifier. `Package::comments` on the source-backed owner matches
the eager `Document::comments`.

| Name | Status |
|---|---|
| `document::GraphChange` | adopted (0251b) |
| `document::HyperlinkTarget` | adopted (0251b) |
| `Edit::{add_external_relationship, add_related_part, replace_part, part_payload, related_part}` | adopted (0251b) |
| `Edit::{insert_hyperlink_at, set_hyperlink_target_at}` | adopted (0251b) |
| `document::NewComment` | adopted (0251c) |
| `Edit::{add_comment_at, remove_comment}` | adopted (0251c) |
| `source_backed::Package::comments` | adopted (0251c) |
| `document::NewPicture` | adopted (0251d) |
| `Edit::insert_picture_at` | adopted (0251d) |
| `numbering::{ListKind, AddedList}` | adopted (0251e) |
| `Edit::{add_list, list_kind}` | adopted (0251e) |

## Command-channel property policy, PPTX shape fill and slide transitions (record 0269)

`batch::Unsupported::{Refuse, Warn}` names what a command does with a property
its target refuses with `unsupported_property`. `Refuse` is the default
everywhere and the only behaviour of the library verbs (`set`, `add`, `edit`
closures); `Warn` is OfficeCLI's "apply the rest and warn" and exists only on
the command channel. `batch::Options` is the command-channel envelope (batch
mode, unsupported-property policy, the transaction's `EditOptions`) with
consuming setters `mode`, `unsupported` and `edit_options` (record 0362);
`File::apply_with` and `File::batch_with` take it, and
`File::batch_with_options` delegates to `batch_with`. Skipped
properties are reported per item in `Applied::warnings`, each the refusing
`Error` with its property and item index; no second diagnostics type is
introduced and `transaction::Diagnostics` is unchanged.

PPTX leaf: `opened::ShapeFill::{Solid, None}` is a shape's own `p:spPr` fill;
`TextProperties::{with_shape_fill, with_geometry}` assign it and a preset
geometry (`litchi_drawingml::geom::Preset`) on a `TextTarget::Shape` edit, so
both owners reuse `set_text_properties`. `opened::Edit::set_slide_transition`
and the source-backed primary `Edit::set_slide_transition` set (`Some`) or
clear (`None`) a slide's direct transition, mirroring
`SourceBackedSlideEdit::{set_transition, clear_transition}`.

| Name | Status |
|---|---|
| `batch::Unsupported::{Refuse, Warn}` | adopted (0269) |
| `batch::Options` (`mode`, `unsupported`, `edit_options`) | adopted (0269); `limits` → `edit_options` (0362) |
| `File::{apply_with, batch_with}` | adopted (0269) |
| `Applied::warnings` | adopted (0269) |
| `litchi_pptx::opened::ShapeFill` | adopted (0269) |
| `TextProperties::{with_shape_fill, with_geometry}` | adopted (0269) |
| `opened::Edit::set_slide_transition`, `source::Edit::set_slide_transition` | adopted (0269) |

## Typed chart CRUD (record 0270)

`litchi_drawingml::chart::authoring` is the host-neutral way to author a
chart: a `Spec` (a `Variant` — `Kind`, `Grouping`, 3D flag — with title,
categories, `Series`, legend, axis titles, an `Appearance` and a `Binding`)
writes a deterministic chart-space part (`Spec::to_xml`) and describes the
worksheet its references use (`Spec::sheet_layout`, `Cell`). Hosts place
it: PPTX through the primary `opened::Edit::add_chart` (chart part,
deterministic embedded workbook, relationships and frame) and replace it
through `Edit::set_chart`; XLSX through the scoped family
`chart::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` with
`Package::{edit_charts, apply_charts}` and the reads
`Package::{chart_count, chart_xml}` and `SourceWorksheet::{chart_count,
chart_xml}`. The source-preserving chart editor gains
`set_legend`, `set_axis_title`, `set_series_values` and `set_categories`
on the focused `chart::edit::Edit` family. Its immutable `chart::Package`
publishes a core `Commit<Package, Patch>` through `Package::apply`, preserving
the exact source bytes on no-op edits; `Commit::package()` and
`Commit::diagnostics()` are the only publication accessors.

| Name | Status |
|---|---|
| `litchi_drawingml::chart::authoring::{Spec, Variant, Kind, Grouping, Series, Binding, Cell, Appearance, Labels, Line}` | adopted (0270) |
| `litchi_drawingml::chart::edit::Edit::{set_legend, set_axis_title, set_series_values, set_categories}` | adopted (0270; DrawingML family closure 0362dv) |
| `litchi_pptx::opened::Edit::{add_chart, chart, set_chart}` | adopted (0270); `add_chart` takes a `Spec`, the old `crate::chart::Chart` argument is removed |
| `litchi_pptx::presentation::embedded::excel::Workbook::from_spec` | adopted (0270) |
| `litchi_pptx::slide::Slide::chart_xml`, `SourceSlide::chart_xml`, `source::Slide::chart_xml` | adopted (0270) |
| `litchi_xlsx::chart::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | adopted (0270); family registered in the gate |
| `litchi_xlsx::Package::{edit_charts, apply_charts, chart_count, chart_xml}` | adopted (0270) |
| `litchi_xlsx::SourceWorksheet::{chart_count, chart_xml}` | adopted (0270) |

## DrawingML detached family closure (record 0362dv)

The host-neutral DrawingML editors now use the same focused family vocabulary
as package owners. `chart::Package`, `transform::Package`, and
`theme::family::Package` retain exact source bytes, expose `edit()` and
`apply(&Patch) -> Result<Commit>`, and publish the core
`Commit<Package, Patch, Diagnostics>` triad. Their commits expose
`package()`/`into_package()`, `patch()`, `diagnostics()`, and `into_parts()`;
no snapshot accessors or `Transaction` aliases remain. `Patch::is_noop()` and
`Diagnostics::changed()` provide no-op introspection without eager
serialization. The action profile follows the same triad with `Prepared` as
its checked package value.

Raw authoring remains explicitly bounded: `ink::Prepared::from_bytes` requires
an `ink::builder::AuthoringLimits` value, and
`ink::actions::Edit::{new, from_profile}` require
`actions::builder::EditOptions`. Fresh action and InkML authoring
types live under `ink::actions::builder` and `ink::builder`; their old
compatibility spellings are deleted without aliases or deprecations. Action
owners expose `Prepared::{edit, edit_with_options, apply}` and publish only
through `Edit::commit`; raw `Patch::apply` is private. Every chart,
transform, theme-family, and ink-action consumer uses the focused
`Edit`/`Commit` shape in the same migration.

| Name | Status |
|---|---|
| `litchi_drawingml::{chart, transform}::{Package, Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | adopted (0362dv; registered family triads and replay) |
| `litchi_drawingml::theme::family::{Package, Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | adopted (0362dv; exact opaque-markup preservation) |
| `litchi_drawingml::ink::actions::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | adopted (0362dv; `Prepared` package triad) |
| `litchi_drawingml::ink::actions::builder::{Draft, ActionDraft, ActionGroupDraft, ActionDataDraft, DataGroupDraft, PropertyDraft, OpaquePayload, EditOptions}` | adopted (0362dva; canonical fresh-action authoring) |
| `litchi_drawingml::ink::builder::{Draft, Prepared, AuthoringLimits}` | adopted (0362dva; canonical fresh InkML authoring) |
| `litchi_drawingml::ink::Prepared::from_bytes(bytes, AuthoringLimits)` | adopted (0362dv; required raw limit) |
| `litchi_drawingml::chart::Package::apply`, `transform::Package::apply`, `theme::family::Package::apply` | adopted (0362dv; source-checked replay returns `Commit`) |
| `litchi_drawingml::ink::actions::Prepared::{edit, edit_with_options, apply}` | adopted (0362dva; O(1) shared owner routes and commit triads) |
| `litchi_drawingml::theme::family::part::{Limits, replace_family, add_family, add_family_with_uri, remove_family}` | adopted (0362dvb; every raw Theme-family mutation requires the module `Limits` policy; removed `_with_limit` twins are banned) |
| `litchi_drawingml::chart::Limits` and `chart::{reader::read, authoring::{read, Parsed::is_authored}, chartex::read, style::{parse, parse_color}, writer::{write, write_with_rels}}` | adopted (0362dve; one required chart raw policy across source bytes, XML structure, authored data, ChartEx, styles, and bounded writer output) |

## Per-series chart keys and warn-mode failures (record 0271)

`litchi_drawingml::chart::authoring::Series` gains `color` (an `RRGGBB` or
`AARRGGBB` fill that takes precedence over the series' `Appearance::colors`
entry) and the consuming setter `with_color`, the same `with_*` shape as
`Spec`. No new type: OfficeCLI's dotted add keys `series{N}.name`,
`series{N}.values` and `series{N}.color` are registry properties of the
chart element, not API names. `litchi_ops::Error::warnings` returns the
properties an `Unsupported::Warn` command skipped before it failed, reusing
the name and element type of `Applied::warnings`; it is empty for every
other error.

| Name | Status |
|---|---|
| `litchi_drawingml::chart::authoring::Series::{color, with_color}` | adopted (0271) |
| `litchi_ops::Error::warnings` | adopted (0271) |

## Unified CSV/TSV import and formula evaluation (record 0266)

`File::import(target, source, &import::Options)` and
`File::import_with_options` (record 0362) import delimited text into a worksheet of an XLSX,
XLSB or XLS workbook as one mutation and one native commit, returning
`transaction::Commit<import::Summary>`; `transaction::Edit::import` is the
same mutation inside a closure transaction. `import` is the design §13 verb
name; Python spells it `File.import_csv` because `import` is a Python
keyword. Options are one consuming `import::Options` (`csv`, `tsv`,
`delimiter`, `quote`, `header`, `encoding`, `infer`, `infer_numbers`,
`infer_booleans`, `infer_dates`, `formulas`); limits are the transaction's
`SemanticLimits` (`max_expanded_targets` bounds fields, `max_staged_bytes`
bounds the source and staged payload), not a second limits type. Parse
failures are the focused, content-free `import::Error` (`InvalidOption`,
`Encoding`, `UnterminatedQuote`, `InvalidCharacter`, `FieldTooLong`) with
`import::OptionKind`.

`File::evaluate(path, &calc::Options)` and `File::recalculate(&calc::Options)`
(feature `eval`) read formula results through `litchi-eval`. The module is
`calc` because the facade's `litchi::eval` already re-exports the engine
crate. `calc::Value`/`ErrorValue` are the evaluated values,
`calc::Evaluated`/`Origin` one cell's result, `calc::Report`/`Formula`/`Status`
a recalculation. Bounds are `calc::Options::{max_cells, max_depth}` with
`calc::LimitKind::{Cells, Depth}` folded into the focused `calc::Error`. The
engine gains `litchi_eval::FormulaEvaluator::with_max_depth` and the
`litchi_eval::DepthLimit` error.

| Name | Status |
|---|---|
| `File::{import, import_with_options}`, `transaction::Edit::import` | adopted (0266); `import_with_limits` → `import_with_options` (0362) |
| `import::{Options, Encoding, Summary, Error, OptionKind}` | adopted (0266) |
| `File::{evaluate, recalculate}` | adopted (0266) |
| `calc::{Options, Value, ErrorValue, Evaluated, Origin, Report, Formula, Status, Error, LimitKind}` | adopted (0266) |
| `capabilities::Verb::{Import, Evaluate}` | adopted (0266) |
| `litchi_eval::{FormulaEvaluator::with_max_depth, DepthLimit}` | adopted (0266) |
| Python `File.import_csv`, `File.evaluate`, `File.recalculate` | adopted (0266) |

## Worksheet cell cursors (record 0267)

Each source-backed worksheet handle gains `cell_cursor()`, a lazy iterator
over its stored cells in row-major order that decodes only as many cells as
the caller reads. The name follows the existing `Grid`/`Cells` vocabulary of
each leaf: a cursor yields coordinates (and, for XLS, the owned cell the
whole-sheet visitor reports), and the cells are read from the grid the cursor
names. No second grid type or alias is introduced.

| Name | Status |
|---|---|
| `litchi_xlsx::workbook::source::CellCursor`, `SourceWorksheet::cell_cursor() -> Result<CellCursor>` (items `(Address, RecordIndex)`; `CellCursor::{grid, is_complete}`) | adopted (0267) |
| `litchi_xlsb::sheet::CellCursor`, `SourceBackedWorksheet::cell_cursor() -> Result<CellCursor>` (items `(u32, u32)`; `CellCursor::{grid, is_complete}`) | adopted (0267) |
| `litchi_xls::CellCursor`, `SourceBackedWorksheet::cell_cursor() -> Result<Option<CellCursor>>` (items `SourceBackedCell`; `CellCursor::needs_formula_context`) | adopted (0267) |

See [0267](unified-ops/0267-streaming-cell-cursors.md).

## Unified template merge (record 0274)

`File::merge(&merge::Record, &merge::Options)` and `File::merge_with_options`
fill a DOCX, PPTX or XLSX template's `{{key}}` placeholders (OfficeCLI's
grammar) and, for DOCX, its `MERGEFIELD` fields from one record, as one
mutation and at most one native commit, returning
`transaction::Commit<merge::Summary>`; `transaction::Edit::merge` is the same
mutation inside a closure transaction. `merge` is the design §13 verb name
and always means template (mail) merge; document merging is a separate P2
concern. Data sources are `merge::Data::{from_json, from_csv}`, each taking
its `SemanticLimits` as a required argument (record 0362; records in source
order); `merge::Record` is one
record of text values. Options are one `merge::Options` (`missing`,
`fields`) with `merge::Missing::{Keep, Empty}`; limits are the transaction's
`SemanticLimits` (`max_expanded_targets` bounds replacements,
`max_staged_bytes` bounds their payload and the data source), not a second
limits type. Data failures are the focused, content-free `merge::Error`
(`Json`, `TooDeep`, `NotAnObject`, `NotARecord`), classified `invalid_json`
or `limit_exceeded`.

The DOCX leaf gains a formatting-preserving story substitution beside the
main-document edit: `document::Edit::{story, substitute}` for the main
document and `document::Edit<Package>::{story_parts, part_story,
substitute_part}` for headers, footers, notes and comments, over the types of
the `document::substitute` module. A source-backed handle materializes its
owning package once for these writes (`Docx` keeps the promotion).

| Name | Status |
|---|---|
| `File::{merge, merge_with_options}`, `transaction::Edit::merge` | adopted (0274); `merge_with_limits` → `merge_with_options` (0362) |
| `merge::{Record, Data, Options, Missing, Summary, Error}` | adopted (0274) |
| `capabilities::Verb::Merge` | adopted (0274) |
| `litchi_docx::document::Edit::{story, substitute}`; `Edit<Package>::{story_parts, part_story, substitute_part}` | adopted (0274); `story_parts` returns `StoryPart`s (0285) |
| `litchi_docx::document::substitute::{Story, Paragraph, Field, FieldKind, Replacement, Outcome, OBJECT}` | adopted (0274) |
| CLI `litchi merge`, Python `File.merge` | adopted (0274) |

See [0274](unified-ops/0274-template-merge.md).

## Field and table-of-contents refresh (record 0279)

`File::refresh(&refresh::Options)` and `File::refresh_with_options` recompute
the derived results a document caches, as one mutation and at most one native
commit, returning `transaction::Commit<refresh::Summary>`;
`transaction::Edit::refresh` is the same mutation inside a closure
transaction. `refresh` is the design §13 verb name (OfficeCLI's command).
Options are one `refresh::Options` (`clock`, `page_numbers`, `mark_dirty`,
`write_cache`, and `calc` with the `eval` feature) with `refresh::Clock` (an
explicit date and time; the system clock is never read) and
`refresh::PageNumbers::{Keep, Require}`; limits are the transaction's
`SemanticLimits`. The report is `refresh::Summary` of `refresh::Field`s, each
with a `refresh::Status` (`Refreshed`, `Unchanged`, `Stale`, `NeedsLayout`,
`NeedsClock`, `Locked`, `Unresolved`, `Unsupported`). A table of contents
whose page numbers are required but unknown is `refresh_failed`.

Leaf primitives for recomputing derived field results. The DOCX leaf indexes
the main story's paragraphs, simple and complex fields (instruction, cached
result, nesting, lock and dirty state) and bookmarks, and rewrites field
results without evaluating anything: `document::Edit::{fields,
refresh_fields}` over the `document::fields` module (`Fields`, `Paragraph`,
`Field`, `Bookmark`, `Update::{Text, Toc, Bookmark, Instruction, RenameBookmark,
RemoveBookmark}` (the last three record 0325), `TocEntry`, `Outcome`).
The XLSX worksheet edit stores a formula's cached result while keeping its
formula markup (`WorksheetEdit::set_cached`). PPTX run-text edits accept a
field run (`a:fld`) as their target, writing its cached result.

| Name | Status |
|---|---|
| `File::{refresh, refresh_with_options}`, `transaction::Edit::refresh` | adopted (0279); `refresh_with_limits` → `refresh_with_options` (0362) |
| `refresh::{Options, Clock, PageNumbers, Summary, Field, Status}` | adopted (0279) |
| `capabilities::Verb::Refresh` | adopted (0279) |
| CLI `litchi refresh`, Python `File.refresh` | adopted (0279) |
| `litchi_docx::document::Edit::{fields, refresh_fields}` (snapshot and package edits) | adopted (0279) |
| `litchi_docx::styles::Style::declares_outline` | adopted (0279) |
| `litchi_pptx::{Package, source::Package}::first_slide_number` | adopted (0279) |
| `litchi_docx::document::fields::{Fields, Paragraph, Field, Bookmark, Update, TocEntry, Outcome}` | adopted (0279) |
| `litchi_xlsx::WorksheetEdit::set_cached` | adopted (0279) |
| `litchi_pptx::opened::TextTarget::Run` text on `a:fld` runs | adopted (0279) |

See [0279](unified-ops/0279-field-refresh.md).

## Formatting-aware find/replace (record 0273)

Find/replace is not a verb: as OfficeCLI's `set --find`, a `File::set`
(and `Edit::set`, `File::apply`, batch `set` items) whose properties name
`find` searches the target's subtree. The keywords are OfficeCLI's `find` and
`replace`, plus Litchi's `matchCase` (default `true`) and `wholeWord`
(default `false`); `regex` is admitted only as `false` (record 0010 Q3).
Every other property of the call formats the replaced, or matched, text. The
match count is `Applied::matched() -> Option<usize>` (CLI `--json` and Python
result dicts: `matched`, OfficeCLI's name). `capabilities::Verb::Replace`
(`"replace"`) reports format support, and the schema export lists the
keywords under `setKeywords`. Leaf primitives: the DOCX transaction lists a
paragraph's searchable `w:t` nodes and rewrites them, optionally isolating
byte ranges into runs of their own (`TextNode`, `TextNodeRewrite`,
`ParagraphRewrite`, `Edit::{text_nodes, rewrite_text_nodes}`); PPTX text
edits gain `TextTarget::RunSpan` (a byte span of one run, split out with a
copy of the run's properties); XLSX worksheets report rich-text cells
(`Worksheet::has_rich_text`).

| Name | Status |
|---|---|
| `find`, `replace`, `matchCase`, `wholeWord`, `regex` properties of `set` | adopted (0273) |
| `Applied::matched` | adopted (0273) |
| `capabilities::Verb::Replace` | adopted (0273) |
| `litchi_docx::document::{TextNode, TextNodeRewrite, ParagraphRewrite}`, `Edit::{text_nodes, rewrite_text_nodes}` (snapshot and package edits) | removed and banned (0285): `document::substitute` is the one primitive |
| `litchi_pptx::opened::TextTarget::RunSpan` | adopted (0273) |
| `litchi_xlsx::Worksheet::has_rich_text` | adopted (0273) |

See [0273](unified-ops/0273-find-replace.md).

## Find/replace and template merge on one primitive (record 0285)

Find/replace and template merge share one DOCX leaf primitive, the story
substitution of `litchi_docx::document::substitute`: a `Story` lists each
paragraph's text with its `Segment`s (one per `w:t`, grouped where markup
separates text) and fields; `Replacement`s rewrite ranges of it, optionally
`formatted` (the new text, or a `keep` range's text, split into runs of its
own that a run-property closure formats). Field results are text (OfficeCLI's
find reads them), instructions are fixed content, and field state is followed
across paragraphs. `Edit::story_within` reads only the paragraphs a set of
element paths names; `Edit::substitute_formatted` and
`Edit<Package>::substitute_part_formatted` apply formatted replacements;
`Edit<Package>::story_parts` lists `StoryPart`s (`name`, `kind`:
`StoryKind::{Header, Footer, Footnotes, Endnotes, Comments}`). The
record-0273 text-node API is removed. PPTX find and merge share one run
rewrite and one text-body walk in `litchi-ops`; PPTX text edits address
table-cell runs (`opened::TextTarget::{CellRun, CellRunSpan}`, read through
`table::Cell::text_body`) and notes-page shapes
(`opened::Edit::set_notes_text_properties`).

| Name | Status |
|---|---|
| `litchi_docx::document::Edit::{story_within, substitute_formatted}` (snapshot and package edits); `Edit<Package>::substitute_part_formatted`; `litchi_docx::source_backed::Package::story_parts` | adopted (0285) |
| `litchi_docx::document::substitute::{Segment, StoryKind, StoryPart}`; `Replacement::{keep, formatted, text}`; `Outcome::formatted`; `Paragraph::{segments, item}` | adopted (0285) |
| `litchi_pptx::opened::TextTarget::{CellRun, CellRunSpan}`, `litchi_pptx::table::Cell::text_body` | adopted (0285) |
| `litchi_pptx::opened::Edit::set_notes_text_properties` | adopted (0285) |

See [0285](unified-ops/0285-find-merge-consolidation.md).

## DOCX run-level markup compatibility (record 0307)

Document transactions admit a main document whose `mc:AlternateContent`
wrappers all lie inside runs (drawings and text boxes with a fallback). The
shared proof that establishes it, `litchi_ooxml_common::mce::admission::
admit_held`, takes the holder element (`Holder`) and reports the verdict
outside the wrappers and their count (`Held`); it does not cover processing
inside a wrapper, which the caller runs. The story of
`litchi_docx::document::substitute` lists each wrapper as one `OBJECT` and
`Paragraph::alternate_content` gives their offsets. A text set over a wrapper
whose branches hold text is refused with `document::Refusal::
AlternateContentText`.

| Name | Status |
|---|---|
| `litchi_ooxml_common::mce::admission::{admit_held, Holder, Held}` | adopted (0307) |
| `litchi_docx::document::substitute::Paragraph::alternate_content` | adopted (0307) |
| `litchi_docx::document::Refusal::AlternateContentText` | adopted (0307) |

See [0307](unified-ops/0307-docx-run-markup-compatibility.md).

## Typed SmartArt CRUD (record 0284)

`litchi_drawingml::diagram::authoring` is the host-neutral way to author a
SmartArt graphic: a `Spec` (a built-in `Layout` and a tree of `Node`s, also
parsed from a tab-indented text outline) writes the data model, the layout
definition and Litchi's own deterministic cached drawing; `quick_style_xml`
and `colors_xml` write the style parts; `Spec::{from_parts, recognize}`
recognize parts this module wrote. `litchi_drawingml::diagram::edit` reads a
data model's `Outline` (content `Entry`s in document order, layout and
drawing reference) and makes source-preserving changes (`set_text`,
`drawing_text`, `set_drawing_reference`); it is a set of pure functions over
part bytes, not a transaction family. PPTX places and replaces the parts
through the primary `opened::Edit`.

| Name | Status |
|---|---|
| `litchi_drawingml::diagram::authoring::{Spec, Layout, Node, quick_style_xml, colors_xml}` | adopted (0284) |
| `litchi_drawingml::diagram::edit::{Outline, Entry, set_text, drawing_text, set_drawing_reference}` | adopted (0284) |
| `litchi_pptx::opened::Edit::{add_smartart, smartart, set_smartart}`, `litchi_pptx::opened::SmartArtParts` | adopted (0284) |
| `litchi_pptx::slide::Slide::diagram_xml`, `SourceSlide::diagram_xml`, `source::Slide::diagram_xml` | adopted (0284) |

See [0284](unified-ops/0284-typed-smartart-crud.md).

## Remaining ODF leaves, W3 (record 0282)

`litchi-oth`, `litchi-odm`, `litchi-odg`, `litchi-odc`, `litchi-odi`,
`litchi-odb` and `litchi-odf-formula` follow the ODT/ODS/ODP split of record
0268. Each has one immutable `Package` owner with every state verb (`new`,
`open`, `open_with_options`, `from_bytes(&[u8])`, `from_bytes_with_options`,
`from_reader(_with_options)`, `from_prepared_package`, password routes where
the crate had them, `edit`, `apply`, `save`, `save_with_durability`,
`to_bytes`, `write_to`, `as_bytes`) and a read-only projection: OTH and ODM
`Document`, ODG `Drawing`, ODC `Chart`, ODI `Image`, ODB `Database`,
ODF formula `Formula`. The primary families are
`litchi_oth`, `litchi_odm::transaction`, `litchi_odg::edit`, `litchi_odc`,
`litchi_odi`, `litchi_odb` and `litchi_odf_formula`, each exporting
`{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` (`Error`/`LimitKind`
are `litchi_odf_common`'s). `Commit` exposes `package`, `patch`,
`diagnostics`, `into_package`, `into_parts`; `Patch::is_noop` is added and
public `Patch::apply` removed. `Builder::build()` returns `Package`.
Ingress applies `litchi_odf_common::OpenOptions` through the additive
`core::family::Package::*_with_options` constructors.

| Name | Status |
|---|---|
| `litchi_oth::Template`, `litchi_odm::Master`, `litchi_odg::Package*` root aliases, `Snapshot`, `Transaction` | removed and banned (profiles `oth`, `odm`, `odg`) |
| projection state verbs (`Drawing::open`, `Chart::from_bytes`, `Image::edit`, `Database::into_bytes`, …) | removed and banned |
| `Commit::{changed, template, master, snapshot, chart, image, database, formula}` | removed and banned per profile |
| `litchi_odi::flat::Commit::{semantic_patch, source}` | removed and banned; semantic planning and source provenance are on `litchi_odi::flat::Patch` (0362usd) |
| ODG joined composition owner routes (`Package::apply_joined*`) | removed; stage a bounded `JoinedEdits` with `Edit::join`, publish the shared `Commit` triad, and replay with `Package::apply(&Patch)` |
| metadata `edit_metadata`/`apply_metadata` family on these crates | not adopted (gap; metadata writes stay in the primary `Edit` where the crate has them) |
| flat `Flat*` owners, ODC `Definition*`, history/merge vocabularies | unchanged (out of scope, record 0282 gaps); ODI flat semantic planning now lives on `flat::Patch::semantic_patch` (0362usd) |


## Unified ODF formats (record 0282)

| Name | Status |
|---|---|
| `litchi_ops::File::{Oth, Odm}` (hold `format::odt::Odt`; `Odt::format()`) | adopted (0282) |
| `litchi_ops::File::{Odg, Odc, Odi, Odb, Odf}`, `format::odf_tree::{OdfTree, Node}`, `node::Node::Object`, `ToleranceReport::Object` | adopted (0282) |
| `litchi_odt::Package::from_prepared_text_package` (ODT/OTH/ODM prepared packages) | adopted (0282, additive reuse hook) |
| `litchi_odf_common::core::family::Package::*_with_options` | adopted (0282, additive) |

## DOCX charts and extended chart types (record 0283)

`litchi_drawingml::chart::authoring::Kind` gains the extended families
(`Waterfall`, `Funnel`, `Treemap`, `Sunburst`, `Histogram`, `Pareto`,
`BoxWhisker`; `Kind::is_extended`), which `Spec::to_xml` writes as a
`cx:chartSpace` part through the new `chart::chartex` module.
`authoring::read` turns a `c:chartSpace` or `cx:chartSpace` part back into
the `Spec` that writes it (`Parsed`, `Parsed::is_authored`), so hosts can
rebuild a chart whose family or data changes without losing anything the
model reads. `Spec::workbook_members` replaces the PPTX-private worksheet
generator so DOCX shares it.

| Name | Status |
|---|---|
| `litchi_drawingml::chart::authoring::Kind::{Waterfall, Funnel, Treemap, Sunburst, Histogram, Pareto, BoxWhisker, is_extended}` | adopted (0283) |
| `litchi_drawingml::chart::authoring::{Bins, QuartileMethod, ParentLabels}`, `Appearance::{subtotals, increase_color, decrease_color, total_color, bins, quartile_method, parent_labels}` | adopted (0283) |
| `litchi_drawingml::chart::authoring::{read, Parsed}`, `Parsed::is_authored`, `Spec::of` | adopted (0283) |
| `litchi_drawingml::chart::authoring::Spec::{category_axis_title, value_axis_title, series_kinds, with_variant, without_title, with_data, series_mut, appearance_mut, workbook_members}` | adopted (0283) |
| `litchi_drawingml::chart::chartex::{read, style_xml, colors_xml, NAMESPACE, CONTENT_TYPE, RELATIONSHIP_TYPE, GRAPHIC_DATA_URI, STYLE_CONTENT_TYPE, STYLE_RELATIONSHIP_TYPE, COLORS_CONTENT_TYPE, COLORS_RELATIONSHIP_TYPE, MAX_XML_BYTES}` | adopted (0283) |
| `litchi_drawingml::chart::edit::Edit::set_series_color` | adopted (0283; 0362dv) |
| `litchi_drawingml::chart::chartex::is_chartex` | adopted (0283) |
| `litchi_pptx::opened::Edit::add_chart` with an extended `Spec` (a `cx:chartSpace` part, its companions and a `cx:chart` frame); `Slide::chart_xml`, `SourceSlide::chart_xml` and `chart`/`set_chart` read and replace extended parts too | adopted (0283) |
| `litchi_pptx::opened::Edit::set_shape_properties` bounds also move and resize chart graphic frames (`p:graphicFrame/p:xfrm`) | adopted (0283) |
| `litchi_xlsx::chart::edit::Edit::set_anchor`, `litchi_xlsx::Package::chart_anchor`, `litchi_xlsx::SourceWorksheet::chart_anchor`; `Edit::add` stages extended parts; `chart::Anchor` gains `PartialEq`/`Eq` | adopted (0283) |
| `litchi_docx::chart::InlineChart`, `Paragraph::charts`, `Package::chart_xml`, `source_backed::Package::chart_xml` | adopted (0283) |
| `litchi_docx::document::{NewChart, chart_workbook}`, `Edit::{insert_chart_at, chart_parts, set_chart, set_chart_extent_at, remove_chart_at}` | adopted (0283) |

See [0283](unified-ops/0283-docx-charts-and-chartex.md).

## Find/replace and merge on every text format (record 0295)

Find/replace (`set` with `find=`) and template merge (`File::merge`) cover
XLSB, DOC, PPT, XLS, RTF, ODT, OTH, ODM, ODS and ODP through one
crate-private text-surface engine (`find::surface`) and per-format adapters
over their existing native text and cell edits; no new unified verb or
option. XLSX rich-text cells are rewritten run by run through a new leaf
read and write pair: `Worksheet::rich_text` reads a cell's runs (a formatted
shared-string item or an inline rich string) with each run's properties as
stored, `RichText::set_run_text` changes one run's text, and
`WorksheetEdit::set_rich_text` writes the runs as the cell's own inline rich
string.

| Name | Status |
|---|---|
| `litchi_xlsx::{RichText, RichRun}` (`RichText::{plain, runs, text, set_run_text}`, `RichRun::{text, is_formatted}`) | adopted (0295) |
| `litchi_xlsx::Worksheet::rich_text` | adopted (0295) |
| `litchi_xlsx::WorksheetEdit::set_rich_text` | adopted (0295) |
| `litchi_odt::body::{TextField, text_fields}` (`TextField::{name, range, attribute, text}`), `litchi_odt::transaction::Edit::set_field_text` | adopted (0295): ODT field refresh |
| `litchi_odp::scene::Change::SetFieldText` | adopted (0295): ODP field refresh |
| `litchi_docx::document::NewSmartArt` (`new`, `with_size`, `with_name`, `with_description`, `size`), `Edit::insert_smartart_at` | adopted (0295): DOCX SmartArt insertion |
| DOCX `smartart` element (`/smartart[N]`, `/body/smartart[N]`; `layout`, `layoutId`, `nodeCount`, `text`; add-only `name`, `alt`, `width`, `height`) | adopted (0295) |
| `litchi_xlsb` cell-value formula caches changing type (`set_formula_*_cache` across the four `BrtFmla*` families) | behaviour (0295): XLSB cache write-back |

See [0295](unified-ops/0295-m4-engines-on-more-formats.md).

## W1 Numbers (record 0297)

`litchi-numbers` completes its ADR 0034 W1 slice. The immutable `Package`
owns every state verb (`open`, `from_bytes`, `from_shared_bytes`, their
`*_with_options` forms over one `OpenOptions`, `new`, `edit_*`, `apply_*`,
`save`, `write_to`); `Workbook` (formerly `Document`, module
`litchi_numbers::workbook`) is the read-only projection, owned through
`Package::workbook()` and borrowed through `workbook_ref()`. Like the Pages
`Document`, the projection keeps read-only `open`/`from_bytes` ingress for
app-authored package directories, which the package owner refuses.

- Limits: `litchi_numbers::limits::{SourceLimits, SourceLimitsBuilder,
  SemanticLimits, SemanticLimitsBuilder, SourceLimitKind, SemanticLimitKind,
  LimitKind, OpenOptions, Error}`; invalid values are the module's
  `limits::Error::InvalidLimit` (kept off the package error so the
  read-only projection's API does not reach the physical `package` module;
  `tools/check_iwork_leaf_public_api.py`). The
  source default stays at the hard ceilings (2 GiB input): the focused
  transactions derive their work budgets from the physical profile.
- Creation: `Package::new()` (one sheet, one 10 × 5 table, fixed
  `Identity::BLANK`, byte-identical) and `create::Builder` with a required
  `create::Identity`.
- Transaction families moved into focused modules with short names:
  `table::lock`, `table::merge::transaction`, `sheet::chart::{arrangement,
  data, metadata}`, `shape::image::adjustments`, `cell::comment`,
  `cell::comment::transaction`; `table::dimension::transaction::Error` replaces
  `TransactionError`. The prefixed root aliases are deleted.
- Selectors: `Position(Position)`; `index(usize)` and `From<usize>` stay as
  constructors (the Pages precedent); `as_index` is deleted.
- Errors: `Error::UnexpectedFormat { detected }` and
  `workbook::ReadError::UnexpectedFormat { detected }` replace `NotNumbers`;
  the root keeps the ingress `Error`/`Result` only.

Profiles `numbers-minimal` and `numbers-full` (`internal-iwork-source`) in
`tools/api_glossary.json` ban 103 retired names and register 27 complete
transaction families.

## W1 Keynote (record 0298)

ADR 0034 W1 for `litchi-keynote`, part 1 (Presentation projection and limits):

- `document` module → `presentation`: `Document`→`Presentation` (root
  re-export), `DocumentStats`→`presentation::Stats`,
  `DocumentIoKind`→`presentation::IoKind`,
  `DocumentReadError`→`presentation::Error`,
  `DocumentReadLimitKind`→`presentation::LimitKind`;
  `Package::semantic_snapshot()`→`Package::presentation()`.
- `pub mod limits` (root re-exports), mirroring `litchi_odf_common::limits`:
  `SourceLimits`/`SourceLimitsBuilder`/`SourceLimitKind`,
  `SemanticLimits`/`SemanticLimitsBuilder`/`SemanticLimitKind`,
  `LimitKind::{Source, Semantic}`, `OpenOptions` and
  `Error::InvalidLimit { kind, value, maximum }`. Defaults equal the physical
  archive maxima, so default opens are unchanged. Deleted: the root `Limits`
  re-export, `ReadOptions`, `SemanticLimitsError`, positional
  `SemanticLimits::new`, every `Document*Limits*`/`DocumentReadOptions` name,
  the root `MAX_*` constants, `Package::{open_with_limits,
  from_bytes_with_limits, limits, read_options}` and
  `MergeReader::read_options` (now `Package::{options, source_limits}` and
  `MergeReader::options`).
- `NotKeynote` → `UnexpectedFormat { detected: litchi_core::FileFormat }` in
  `ReadError` and `presentation::Error`.
- Registered profiles `keynote-minimal` and `keynote-full`
  (`internal-iwork-source`; the hidden archive seams legitimately take the
  neutral IWA `Limits`, so only `keynote-minimal` bans that name).
  `Package::new()`, `Builder` and `Identity` are not part of this step.

ADR 0034 W1 for `litchi-keynote`, part 2 (focused transaction families):

- The primary transaction is `litchi_keynote::edit::{Edit, Patch, Commit,
  Diagnostics, Error, LimitKind}` (`EditError` → `edit::Error`; the new
  `edit::LimitKind` names no dimension because the one-operation skip/include
  rewrite bounds no resource of its own).
- Every prefixed family moved to a focused module with renamed type
  definitions: `chart::{arrangement, axis_title, caption, legend, title,
  value_axis, data, metadata}`, `background`, `build::order`,
  `slide::image::adjustments`, `slide::movie::{caption, geometry, playback,
  title, creation}`, `slide::audio::{position, creation}`,
  `slide::{notes, order, text}`, `slide::comment::transaction`
  (+`DrawableSummary`), `slide::media::{data (+Data, MediaPart), properties,
  lifecycle}`, `slide::table::{appearance, dimension, lock, name, sort, title,
  number_format, merge (+MergeReader), headers (+InvalidReason),
  physical_sort}::transaction` and `slide::table::cells::{Error, LimitKind,
  Path}`. The alias modules `slide::audio::transaction` and
  `slide::media::{geometry, playback}::transaction` are folded into
  `slide::audio::position` and `slide::movie::{geometry, playback}`; the root
  value aliases `SlideTableName`/`SlideTableNameValueError` are deleted in
  favour of `slide::table::name::{Name, Error}`.
- Every deleted prefixed name is banned under `litchi_keynote`. The 39
  registered families include the already-canonical `slide::delete`,
  `slide::placeholder`, `transition`, `show`, `soundtrack`,
  `soundtrack::{order, items}`. Not registered (no `Edit`, by design):
  `slide::media::lifecycle` and `slide::{audio, movie}::creation` (package
  operations return the commit directly), `chart::metadata` and
  `slide::table::cells` (shared error vocabularies, not transactions).

ADR 0034 W1 for `litchi-keynote`, part 3 (slide creation):

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_keynote::slide::creation::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | Registered family (40th) in `keynote-minimal` / `keynote-full` | ✅ 0298 |
| `Package::edit_slide_creation() -> creation::Edit<'_>` | Several operations per edit, applied in order in one workspace, one commit | ✅ 0298 |
| `Edit::duplicate_slide(impl Into<SlideSelector>, Position)` | Selector resolved against the base snapshot; destination = final position when the operation runs (`0..=count`) | ✅ 0298 |
| `Edit::add_slide(impl Into<LayoutSelector>, Position)` | `LayoutSelector::{Name, Position}`; `&str`, `Position`, `usize`, `&Layout` convert | ✅ 0298 |
| `Package::{slide_layouts, default_slide_layout}` | `creation::Layout { name(), position() }`; no native identifiers | ✅ 0298 |
| `Package::apply_slide_creation(&Patch)` | Exact source check, reopen; `Patch::{is_noop, inverse, source_slide_count, target_slide_count, created_slides}` | ✅ 0298 |
| `Commit::{package, into_package, patch, diagnostics, into_parts}` | `Diagnostics::{changed, created_slides, touched_components, full_reparse_performed}` | ✅ 0298 |
| `creation::Error::UnsupportedContent { reason: Unsupported }` | Typed refusals: `ExternalContent`, `UnmodelledContent`, `RegistryShape`, `LayoutState`, `LayoutMedia`, `NotesTemplate`, `SharedComponent` | ✅ 0298 |

ADR 0034 W1 for `litchi-keynote`, part 4 (slide drawables):

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_keynote::slide::drawable::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | Registered family (41st) in `keynote-minimal` / `keynote-full` | ✅ 0298 |
| `Package::slide_drawable_list(impl Into<SlideSelector>) -> Box<[drawable::Drawable]>` | Source order of the slide's owned drawables (the `DrawableSelector` order); `Drawable::{selector, kind, geometry, text, name, is_locked, z_position}`; `drawable::Kind` (text box, shape, line, image, movie, audio, title/body/slide-number/object/other placeholder, table, chart, group, other); no native identifiers. `slide_drawables` stays the comment family's summary list | ✅ 0298 |
| `drawable::Geometry { x, y, width, height, rotation: Option<f32> }` | Points and degrees from `TSD.GeometryArchive`; `Geometry::{new, with_rotation}` | ✅ 0298 |
| `Package::edit_slide_drawables() -> drawable::Edit<'_>` | Several operations per edit, one workspace, one commit; selectors resolve against the base snapshot | ✅ 0298 |
| `Edit::{set_geometry, remove, move_to}(slide, drawable, …)` | `move_to` takes the back-to-front z-order `Position`; owned order (and so selectors) unchanged | ✅ 0298 |
| `Edit::{add_text_box(slide, Geometry, &str), add_shape(slide, Preset, Geometry, &str), add_image(slide, &[u8], &str, Geometry)}` | `drawable::Preset::{Rectangle, RoundedRectangle { radius }, Ellipse}`; `ImageFormat::{Png, Jpeg, Gif}` sniffed; `Commit::created() -> &[Created { slide(), drawable() }]` | ✅ 0298 |
| `Edit::set_text_style(slide, drawable, TextSpan, Style)` | Each run in the span gets a `TSWP.CharacterStyleArchive` variation of its style in the storage's stylesheet; `Unsupported::TextStyling` only for a style that sets nothing; `Package::slide_drawable_text_runs` reads `TextRun { span, text, style }` with the parent chain resolved | ✅ 0309 (refused in 0298) |
| `Package::apply_slide_drawables(&Patch)` | Exact source check, reopen; `Patch::{is_noop, inverse, operations}` | ✅ 0298 |

## Numbers structure (record 0301)

Three new complete `litchi-numbers` transaction families, each with the
canonical `{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` set and
`Package::{edit_x, apply_x}` entry points:

- `sheet::lifecycle` — `Package::{edit_sheet_lifecycle,
  apply_sheet_lifecycle}`; `Edit::{add, remove}` (one operation per edit),
  plus `sheet::lifecycle::InvalidName`; `Edit::duplicate` since record 0370j.
- `table::lifecycle` — `Package::{edit_table_lifecycle,
  apply_table_lifecycle}` over one sheet selector; `Edit::{add, remove,
  duplicate}` (`duplicate` since record 0370h).
- `table::order` — `Package::{edit_table_order, apply_table_order}` over one
  sheet selector; `Edit::move_table(selector, destination)` with a final
  zero-based `destination`, as `sheet::order` does.

Both profiles now register 30 families.

## Numbers rows and columns (record 0309)

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_numbers::table::topology::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Unsupported}` | Registered family (31st) in `numbers-minimal` / `numbers-full`; the section-relative `RowInsertion`, `ColumnInsertion`, `RowDeletion`, `ColumnDeletion` values stay reachable in the same module, re-exported from their canonical owner `litchi_iwa_common::table::topology` (0368) | ✅ 0309, 0368 |
| `Package::edit_table_topology(sheet, table) -> topology::Edit<'_>` | `Edit::{insert_row, insert_column, remove_row, remove_column}`; several operations per edit, applied in order, one commit | ✅ 0309 |
| `Package::apply_table_topology(&Patch)` | Exact source check; `Patch::{is_noop, inverse}` | ✅ 0309 |
| `topology::Error::Unsupported { reason: Unsupported }` | `Formulas`, `ReferencedByOtherOwners`, `Merges`, `Filters`, `Categories`, `Pivot`, `SortRules`, `HiddenAxis`, `Borders`, `Comments`, `StorageLayout` | ✅ 0309 |
| `litchi_numbers_wire::{ListKind, BncCell::list_keys}` | The table data lists a BNC cell's keys index | ✅ 0309 |

## iWork character styling (record 0309)

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_iwa_text_wire::{CharacterProperties, CharStyleRun, CharStyleError, char_style_runs, styles_in_span, restyle_span, character_style_variation, character_style_properties, character_style_parent, stylesheet_with_style}` | Shared TSWP `table_char_style` runs and the character-style variation records a span restyle writes | ✅ 0309 |
| `litchi_numbers::cell::text_style::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, CharacterProperties}` | Registered family (32nd) in `numbers-minimal` / `numbers-full`; `Package::{edit_table_cell_text_style, apply_table_cell_text_style, table_cell_text_style}`; `Edit::set(At, CharacterProperties)` on stored cells | ✅ 0309 |
| `litchi_numbers::cell::text::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, Property, PropertyKind}` | Registered family in `numbers-minimal` / `numbers-full`; `Package::{edit_table_cell_text, apply_table_cell_text, table_cell_text}`; `Edit::set(At, Property)` and `Edit::reset(At, PropertyKind)`: one property verb over the shared paragraph-style engine (`litchi_iwa_text_wire::paragraph_style`) replacing the host's per-property `set_table_cell_*`/`reset_table_cell_*`/`table_cell_*` methods | ✅ 0370 |
| `litchi_numbers::cell::list::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, Property, PropertyKind, ParagraphList…, TextPosition}` | Registered family in `numbers-minimal` / `numbers-full`; `Package::{edit_table_cell_lists, apply_table_cell_lists, table_cell_list, table_cell_lists, table_cell_list_levels, table_cell_list_level, table_cell_list_numbering, table_cell_list_property}`; `Edit::{set_list, set_lists, set_level, set_numbering, set, reset}` over the shared list engine (`litchi_iwa_text_wire::list`), promoting a plain cell to rich text first; replaces the host's 31 `set_table_cell_paragraph_list*`/`reset_table_cell_paragraph_list*`/`table_cell_paragraph_list*` methods | ✅ 0370b |
| `litchi_numbers::table::cell_appearance::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, Appearance, Fill, Stroke, Layout, Inset, Insets, TextWrap, VerticalAlignment, BorderSide, Borders}` | Registered family in `numbers-minimal` / `numbers-full`; `Package::{table_cell_appearance, edit_table_cell_appearance, apply_table_cell_appearance}`; `Edit::{set_fill, reset_fill, set_layout, reset_layout, set_border, clear_border}` over the shared `litchi_iwa_table::cell_appearance` engine (the Pages and Keynote families' engine), published on the shared structural working set (`cell_appearance::working::publish`); replaces the host's nine `table_cell_{fill,layout,borders}` methods | ✅ 0412b |
| `litchi_numbers::cell::conditional_highlight::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, Rule, Condition, Style, …}` | Registered family in `numbers-minimal` / `numbers-full`; `Package::{edit_table_cell_conditional_highlights, apply_table_cell_conditional_highlights, table_cell_conditional_highlight}`; `Edit::{set(At, rules), clear(At)}` over a prost-free port of the host's conditional-style graph; replaces the host's `set_/clear_cell_conditional_highlighting`, `cell_conditional_highlighting` and `cell_conditional_highlight_rules`; volatile date rules are refused | ✅ 0370d |
| `litchi_numbers::table::resize::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind}` | Registered family in `numbers-minimal` / `numbers-full`; `Package::{edit_table_resize, apply_table_resize}`; `Edit::resize(rows, columns)` at the trailing edges; replaces the host's `resize_table` | ✅ 0370e |
| `litchi_numbers_wire::{Limits, LimitKind}` and `litchi_numbers_wire::{BncCell, BncCellView, cell_value::decode_cell_value, pre_bnc::PreBncCellView}` | Every raw parse, projection and owned encode takes the caller-selected `&Limits`; unknown wire fields and opaque tails remain source-authoritative | ✅ 0362ul66 |
| `litchi_numbers_wire::table_write::{tile_cell, tile_set_cell, header_adjust, list_intern, list_release, cell_bytes, cell_formula_state, table_formula_state, CellInput, FormulaState, TileCellWrite, WriteError}` | Shared TST cell writer for Pages and Keynote tables; each variable-size reader/writer takes `&Limits`, preflights input/output/items, and uses fallible staging | ✅ 0309, 0362ul66 |
| `litchi_pages::table::values::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Value}` | Registered family in `pages`; `Package::{edit_body_table_values, apply_body_table_values}`; `Edit::set(row, column, Value)` | ✅ 0309 |
| `litchi_keynote::slide::drawable::{Edit::set_table_cell, TableCellValue, Unsupported::{TableFormulas, TableStorage}}` | Slide-table cell values in the slide-drawable commit | ✅ 0309 |
| `litchi_pages::table::format::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, DataFormat, At}` | Registered family in `pages`; `Package::{body_table_cell_format, edit_body_table_cell_format, apply_body_table_cell_format}`; `Edit::{set(DataFormat), reset}`: one family over the neutral `DataFormat` and the shared `litchi_iwa_table::cell_format` engine, replacing the host's per-family `set_table_cell_<family>_format`/`reset_table_cell_<family>_format`/`table_cell_<family>_format` methods (controls, Pop-Up Menus and numeral systems are read but refused for writing) | ✅ 0379 |
| `litchi_keynote::slide::table::format::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, DataFormat, At}` | Registered family in `keynote-minimal` / `keynote-full`; `Package::{slide_table_cell_format, edit_slide_table_cell_format, apply_slide_table_cell_format}`; the same verbs and engine as the Pages family; the dedicated `edit_slide_table_cell_number_format` family stays | ✅ 0379 |
| `litchi_pages::table::cell_text::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, Property, PropertyKind, At}` | Registered family in `pages`; `Package::{body_table_cell_text, edit_body_table_cell_text, apply_body_table_cell_text}`; `Edit::set(At, Property)` and `Edit::reset(At, PropertyKind)`: the Numbers `cell::text` family's verbs over the shared `litchi_iwa_table::cell_rich_text::text` engine, replacing the host's per-property `set_table_cell_*`/`reset_table_cell_*`/`table_cell_*` methods on body tables | ✅ 0412 |
| `litchi_pages::table::cell_list::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, Property, PropertyKind, ParagraphList…, TextPosition, At}` | Registered family in `pages`; `Package::{edit_body_table_cell_lists, apply_body_table_cell_lists, body_table_cell_list, body_table_cell_lists, body_table_cell_list_levels, body_table_cell_list_level, body_table_cell_list_numbering, body_table_cell_list_property}`; `Edit::{set_list, set_lists, set_level, set_numbering, set, reset}` over the shared `litchi_iwa_table::cell_rich_text::list` engine (plain cells promoted to rich text first); replaces the host's 31 `table_cell_paragraph_list*` methods on body tables | ✅ 0412 |
| `litchi_keynote::slide::table::cell_text::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, Property, PropertyKind, At}` | Registered family in `keynote-minimal` / `keynote-full`; `Package::{slide_table_cell_text, edit_slide_table_cell_text, apply_slide_table_cell_text}`; the Pages family's verbs over the same engine, replacing the host's `set_slide_table_cell_*`/`reset_slide_table_cell_*`/`slide_table_cell_*` property methods | ✅ 0412 |
| `litchi_keynote::slide::table::cell_list::{Edit, Patch, Commit, Diagnostics, Error, Refusal, LimitKind, Property, PropertyKind, ParagraphList…, TextPosition, At}` | Registered family in `keynote-minimal` / `keynote-full`; `Package::{edit_slide_table_cell_lists, apply_slide_table_cell_lists, slide_table_cell_list, slide_table_cell_lists, slide_table_cell_list_levels, slide_table_cell_list_level, slide_table_cell_list_numbering, slide_table_cell_list_property}`; the Pages family's verbs over the same engine, replacing the host's 31 `slide_table_cell_paragraph_list*` methods | ✅ 0412 |
| `litchi_pages::table::comments::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Path}` | Prepared table-cell root-comment family plus read-only replies from `Comment::replies()`; `Package::{body_table_cell_comment, edit_body_table_cell_comment, apply_body_table_cell_comment}`; selector-first checked cells, exact-source inverse/no-op and candidate readback; reply authoring remains an explicit host-only follow-up (0379) | 🔴 0379 |
| `litchi_keynote::slide::table::comments::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Path}` | Prepared slide-table cell root-comment family plus read-only replies from `Comment::replies()`; `Package::{slide_table_cell_comment, edit_slide_table_cell_comment, apply_slide_table_cell_comment}`; selector-first checked slide/table/cell coordinates, exact-source inverse/no-op and candidate readback; reply authoring remains an explicit host-only follow-up (0379) | 🔴 0379 |
| `litchi_iwa_table::{TableSource, ResolvedObject, ResolveError}`, `litchi_iwa_table::cell_format::{Target, Path, Error, Refusal, LimitKind, TransactionBudget, ListKeys, control, display, data_format, publish, pop_up_menu, pop_up_menu_native, pop_up_menu_metadata}` | Format-neutral table-family engines over one resolved table model (profile `iwa-table`); private to the three leaves, never re-exported by a facade | ✅ 0379 |
| `litchi_iwa_table::cell_rich_text::{Error, Refusal, LimitKind, Geometry, Table, Located, text, list, storage_list}`, `litchi_iwa_table::structure::wire` | Format-neutral whole-cell rich text over one table model (profile `iwa-table`): character/paragraph properties (`text::{read, apply}`), paragraph lists (`list::{Reader, Operation, apply, verify}`) and the storage-list publisher (`storage_list::Publisher`), lifted from `litchi-numbers` (records 0370, 0370b); generic over `structure::TableLeaf`, private to the three leaves | ✅ 0412 |
| `litchi_iwa_table::comments_native::{Selection, Profile, RootOperation, ListKeys, Fault, LimitKind, Mutation, rewrite_root}` | Shared private native root-comment graph/rewrite engine: one bounded archive census and copy-on-write implementation for Pages, Keynote and Numbers; each facade retains selector resolution, publication, and public error mapping (0379) | 🔴 0379 |
| `litchi_iwa_table::structure::{Working, Published, TableLeaf, ReopenError, Fault, FaultLimit, budget, collect, formula, metadata, plan, references}` | The exact-source structural working set and its phases, moved from `litchi-numbers`; Numbers keeps `type Working<'p> = Working<'p, Package>` | ✅ 0379 |
| `litchi_iwa_table::hidden_axes::{Leaf, Budget, Target, Graph, Published, Error, Refusal, LimitKind, resolve_graph, publish, verify_replay, previews, support}` | The hidden-row/hidden-column engine, moved from `litchi-pages` (`body_table_hidden_axes`); the leaf supplies its selector, budget and strict reopen | ✅ 0379 |
| `litchi_keynote::slide::table::hidden_axes::{AxisIndex, HiddenAxes, transaction::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal}}` | Registered family in `keynote-minimal` / `keynote-full`; `Package::{slide_table_hidden_axes, edit_slide_table_hidden_axes, apply_slide_table_hidden_axes}`; `Edit::{set, clear}` over the shared `litchi_iwa_table::hidden_axes` engine (`hidden_axes::WireBudget`, root previews kept) | ✅ 0379 |
| `litchi_numbers::table::hidden_axes::{AxisIndex, HiddenAxes, transaction::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal}}` | Registered family in `numbers-minimal` / `numbers-full`; `Package::{table_hidden_axes, edit_table_hidden_axes, apply_table_hidden_axes}`; `Edit::{set, clear}` over the shared engine (root previews kept) | ✅ 0379 |
| `litchi_pages::section::text_style::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Run, CharacterProperties}` | Registered family in `pages`; `Package::{edit_section_text_style, apply_section_text_style, section_text_runs}`; `Edit::set(TextSpan, CharacterProperties)` (UTF-16, several per edit) | ✅ 0309 |

## iWork lazy tables and section lifecycle (record 0312)

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_pages::section::lifecycle::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | Registered family in `pages` (request 11); `Package::{edit_section_lifecycle, apply_section_lifecycle}`; `Edit::{add(after, Option<&str>), remove(section)}`, one operation per edit; `Error::{InitialSection, EmptySection, UnsupportedContent, UnsupportedDependency}` | ✅ 0312 |

## iWork table creation, `#REF!` and notes runs (record 0329)

Table creation is one verb on the owner (as `insert` is elsewhere), not a
staged family: `litchi_pages::table::creation` carries `{Commit, Patch,
Diagnostics, Error, LimitKind, NewTable}` without an `Edit`, so it is not
registered in the `pages` profile.

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_numbers_wire::table_create::{Shape, Role, TableUuid, Frame, OwnerGeometry, admits_size, model_payload, info_payload, empty_storage, owner_payload, engine_payload, …}` | The payload half of Numbers table creation (0301), shared by the Keynote and Pages owners | ✅ 0329 |
| `litchi_keynote::slide::drawable::{NewTable, Edit::add_table, Unsupported::TableTemplate, Error::InvalidTableSize}` | A slide table cloned from the deck's first slide table, in the slide-drawable commit | ✅ 0329 |
| `litchi_pages::Package::{insert_body_table, apply_body_table_creation}`, `litchi_pages::table::creation::{NewTable, Commit, Patch, Diagnostics, Error, LimitKind}` | A body table cloned from the document's first body table, anchored at a UTF-16 body `Position`; `Error::NoTemplate` without one | ✅ 0329 |
| `litchi_numbers_wire::BncCell::{set_formula_cached_error, set_formula_identifier}` | `#REF!` result caches and shared-formula splits of the topology shift | ✅ 0329 |
| `litchi_keynote::slide::drawable::Edit::set_notes_text_style`, `litchi_keynote::Package::slide_notes_text_runs` | Speaker-notes character styles (request 10) | ✅ 0329 |

## iWork neutral table vocabulary and formula graph (record 0368)

Wave W0-B of record 0363 (decision D4). A moved type keeps one canonical
definition; the format owner re-exports it at its established path, which is
not an alias (no second name exists).

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_iwa_common::table::cell::data_format::{DataFormat, Text, Checkbox, StarRating, control, currency, custom, date_time, duration, fraction, number, numeral_system, percentage, pop_up_menu, scientific, text}` | Archive-free cell display and control formats, moved from `litchi_numbers::cell::data_format`; Numbers re-exports every value at its existing path beside its `transaction` modules, so no public path changes | ✅ 0368 |
| `litchi_iwa_common::table::topology::{RowInsertion, ColumnInsertion, RowDeletion, ColumnDeletion}` | Section-relative topology coordinates, moved from `litchi_numbers::table::topology` (re-exported there) | ✅ 0368 |
| `litchi_numbers_wire::formula_graph::{remap_owner, remap_cell_record_tile, remap_range_precedents_tile, append_owners_to_engine, next_owner_id, Remap, Limits, LimitKind, Error}` | Bounded source-preserving formula-graph remaps for table/sheet duplication, moved from the host's `formula_clone`; required `&Limits`, typed `Error::{Limit, InvalidLimit, InvalidRemap, UnsupportedDependency, InvalidSource, Allocation}`; no `*_with_limits` twin; `Remap::new` checks the maps once and returns `Result` (0368a); `next_owner_id` is the map-aware owner allocator (0368a) | ✅ 0368, 0368a |
| `litchi_numbers_wire::formula_owner::{Uuid, TABLE_OWNER_KIND, Error, encode_empty_table}` | Formula-owner identity arithmetic (`for_table`, `cfuuid_words`, `translate_from`) and the host's empty table owner, byte for byte | ✅ 0368 |
| `litchi_numbers_wire::table_lifecycle::{ReferenceRemap, TABLE_MODEL_REFERENCE_PATHS, TABLE_INFO_REFERENCE_PATHS, Error, rewrite_table_model, rewrite_table_info, rewrite_table_data_list, rewrite_reference, remap_reference_paths}` | Table duplicate/remove reference rewrites (donor `numbers_table_lifecycle_codec`), required `&Limits` | ✅ 0368 |
| `litchi_numbers_wire::{formula_cache::{plan_final_cache, plan_final_cache_with_authored, plan_final_cache_with_logical_graph, collect_formula_hosts, StrictEvaluator, CacheSource, CacheLimits, CachePlan, SuppliedCache, Error, LimitKind, Unsupported}, formula_dependencies::{prepare_graph, SourceGraph, PreparedGraph, Limits, Error, LimitKind}, formula_list::{prepare, PreparedPlan, Limits, Error, LimitKind}}` | The format-neutral formula-dependency removal/insertion and dependent-cache refresh engine, moved from `litchi-numbers` `table_cell_edit::{cache, formula_metadata, formula_list}` and shared by the three table owners; plan-only, every entry point takes its required limits; typed `Error::Limit { kind, observed, maximum }` (the cache planner's `Failure::LimitExceeded` and the dependency rewrite's string-named resource are gone); the authored cache is a borrowed `SuppliedCache`, not `litchi_numbers::formula::CachedValue` | ✅ 0376 |
| `litchi_numbers_wire::formula_cache::refresh::{plan, cache_limits, Source, Write, Plan, TileCaches, TileCaches::apply}` and `formula_cache::Unsupported::Storage` | The dependent formula-cache refresh of value writes into a Keynote slide table or a Pages body table that a formula reads: plans against the source graph through an object resolver, required `&Limits`, publishes nothing; the owner applies `TileCaches` to its staged tile; a formula-cell overwrite or a dependent formula outside the strict evaluator is `Error::UnsupportedDependency`, which Keynote maps to `Unsupported::TableFormulas` and Pages to `Refusal::Formulas` | ✅ 0376 |
| `litchi_numbers_wire::formula_cache::author::{encode, plan, table_owner, Object, Source, Formula, Change, Edit}`, `table_write::{formula_free, with_formula}` and `formula_cache::Unsupported::Graph` | Formula authoring and removal for the Keynote and Pages table owners: `encode` lowers a same-table `FormulaExpression` (a reference to another table, which the expression names by native identifier, or a pivot category is refused), `plan` plans the dependency records, the engine count and the formula list through the moved engine and reports the cell-record tile it creates (with the caller's `next_identifier`) or empties; required `&Limits` | ✅ 0376a |
| `litchi_keynote::slide::drawable::Edit::set_table_formula_cached`, `litchi_pages::table::values::Edit::set_formula_cached`, `litchi_pages::table::formula` | Formula authoring in a slide table and a body table with a supplied cache (Numbers' `set_formula_cached` name); a value write over a formula cell removes the formula | ✅ 0376a |
| `litchi_iwa_protos::{theme_codec, keynote_theme_codec, theme_presets_codec, shape_properties_codec, keynote_shape_creation_codec, number_attachment_codec}` | Bounded Buffa codecs ported from the donors; each entry point takes its `DecodeOptions`/`EncodeOptions`; no limit-less twin | ✅ 0368 |
| `shape_properties_codec::{decode_shape_properties, rewrite_shape_properties}` | banned donor aliases: use `decode_drawable_archive`, `rewrite_drawable_archive` (never ported) | ✅ 0368 |
| `{theme_codec, keynote_theme_codec}::ThemeReferenceSnapshot::encode -> Result<Vec<u8>, DecodeError>` | a zero identifier is a typed error; the panicking `encode`/`encode_checked` pair is replaced, `encode_checked` banned | ✅ 0368a |

## Pages body media, shapes and text out of the host (record 0371)

Wave W1-B of record 0363 moves the host's Pages body editors into
`litchi-pages` families. Host runs of `x()`/`set_x()`/`flip_x()`/`restore_x()`
collapse into one family with property verbs; selectors carry kind and
text-flow position, never a native identifier.

| Name | Meaning | Status |
|---|---|---|
| `litchi_pages::DrawableSelector::{Image, Movie, Audio, Shape}` (`image`, `movie`, `audio`, `shape`, `position`; also `litchi_pages::drawable::DrawableSelector`) | One body image, movie, audio clip or ordinary shape by kind and zero-based text-flow position among that kind (audio = audio-only movie; live video is neither; text boxes are not shapes) — the 0363 §3.2 Pages drawable selector | ✅ 0371 |
| `litchi_pages::drawable::{Geometry, GeometryFlags, Properties, Point, Size, FlipAxis, MediaPlaybackSettings, MediaLoopMode, MediaVolume}` | Archive-free drawable frame (optional position, size, layout flags, angle; `Geometry::flipped`), the reflection-bearing flag set, the four shared drawable properties (hyperlink, locked, aspect-ratio locked, accessibility description) and the shared playback values | ✅ 0371 |
| `litchi_pages::drawable::properties::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Snapshot}` | Registered family in `pages`; `Package::{body_drawable_properties, edit_body_drawable_properties, apply_body_drawable_properties}`; `Edit::{set_geometry, set_position, flip, restore_original_size, set_properties, set_playback}` replace the host's 25 image/movie/audio/shape geometry, flip, Original Size, property, position and playback methods; a verb that does not apply to the kind is `Refusal::UnsupportedOperation` | ✅ 0371 |
| `litchi_pages::drawable::text::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal}` | Registered family in `pages`; `Package::{body_drawable_text, edit_body_drawable_text, apply_body_drawable_text}` over a `DrawableSelector` (ordinary shapes; other kinds `Refusal::NoText`); `Edit::{replace, insert, delete, set, clear}`, one splice per edit; a splice removing a referenced object is `Refusal::DependentContent` | ✅ 0371 |
| `litchi_pages::drawable::caption::{Kind, Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal}` | Registered family in `pages`; `Package::{body_drawable_caption, edit_body_drawable_caption, apply_body_drawable_caption}` over a `DrawableSelector` and `Kind::{Title, Caption}` (images, movies, shapes; audio `Refusal::NoLabel`); `Edit::{set, clear}` replace the host's 15 image/movie/shape title/caption methods; setting an absent label authors the native inline graph, clearing a present one attaches a fresh stand-in | ✅ 0371 |
| `litchi_pages::drawable::lifecycle::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal}` | Registered family in `pages`; `Package::{edit_body_media, apply_body_media}`; `Edit::{add_image, add_movie, add_audio, remove, duplicate, replace_data, replace_poster, replace_asset}` replace the host's `add_body_{image,movie,audio}`, `{remove,duplicate,replace_*_data}_body_*`, `replace_body_movie_poster` and `replace_media`; reads `Package::{body_media, media_assets, section_media_assets, extract_media}` (`body_media` → `lifecycle::Media`) return W1-E's `media::embedded::Embedded` (shapes `Refusal::NotMedia`, a commented drawable `Refusal::Commented`) | ✅ 0371 |
| `litchi_pages::table::sort::execution::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Unsupported}` | Registered family in `pages`; `Package::{edit_body_table_sort_execution, apply_body_table_sort_execution}`; `Edit::{apply, apply_to_rows}` replace the host's `apply_table_sort_order[_to_rows]` | ✅ 0371 |
| `litchi_pages::table::size::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal}` | Registered family in `pages`; `Package::{edit_body_table_size, apply_body_table_size}`; `Edit::set(rows, columns)` replaces the host's `resize_table` | ✅ 0371 |
| `litchi_pages::table::creation::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, NewTable, Refusal}` | Registered family in `pages`; `Package::{edit_body_table_creation, apply_body_table_creation}`; one operation per edit: `Edit::insert(Position, NewTable)` replaces the host's `add_table`, `Edit::duplicate(BodyTableSelector, Position)` its `duplicate_table` (a populated copy named `"<name> copy"`, on the shared `litchi_iwa_table::structure::duplicate` engine); `Refusal::{TableNotFound, AmbiguousTableName, UnsupportedDependency}` | ✅ 0371 |
| `litchi_pages::drawable::style::{Property, PropertyKind, Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Fill, Stroke, Shadow, Effects, Layout, Endpoint, Endpoints}` | Registered family in `pages`; `Package::{body_drawable_style, edit_body_drawable_style, apply_body_drawable_style}`; `Edit::{set, reset}` replace the host's `{,set_,reset_}body_shape_{fill,stroke,shadow,effects,text_layout}` and `{,set_,reset_}body_line_endpoints`, on the shared `litchi_iwa_style::shape` engine and `litchi_iwa_style::line_end` codec | ✅ 0371 |
| `litchi_pages::table::cell_appearance::{Appearance, Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Fill, Layout, Insets, Inset, TextWrap, VerticalAlignment, Borders, BorderSide, Stroke}` | Registered family in `pages`; `Package::{body_table_cell_appearance, edit_body_table_cell_appearance, apply_body_table_cell_appearance}`; `Edit::{set_fill, reset_fill, set_layout, reset_layout, set_border, clear_border}` replace the host's `{set,reset}_table_cell_{fill,layout}`, `{set,clear}_table_cell_border` and `table_cell_{fill,layout,borders}`, on the shared `litchi_iwa_table::cell_appearance` engine | ✅ 0371 |
| `litchi_pages::bookmark::{Bookmark, Selector, Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, TextRange, Name, Settings, Visibility}` | Registered family in `pages`; `Package::{body_bookmarks, edit_body_bookmarks, apply_body_bookmarks}`; `Edit::{add(TextRange, Settings), update(selector, TextRange, Settings), remove(selector)}`, one operation per edit; `Selector::{Position, Name}`; `Bookmark::{range, settings}` carries no native identifier | ✅ 0371 |
| `litchi_pages::date_time::{Field, Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, TextRange, DisplayText, Format, FormatterStyle, Instant, LocaleIdentifier, Settings, UpdatePlan}` | Registered family in `pages`; `Package::{body_date_time_fields, edit_body_date_time_fields, apply_body_date_time_fields}`; `Edit::{add(TextRange, Settings), insert(TextPosition, DisplayText, Settings), update(position, TextRange, Settings), remove(position)}`, one operation per edit; `Field::{range, settings}` | ✅ 0371 |
| `litchi_pages::number_attachment::{Attachment, Target, Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, TextNumberAttachmentSettings, TextNumberAttachmentKind, TextNumberAttachmentFormat, TextNumberAttachmentText}` | Registered family in `pages`; `Package::{number_attachments, edit_number_attachments, apply_number_attachments}` over `Target::{Body, HeaderFooter(HeaderFooterSelector)}`; `Edit::{insert(TextPosition, settings), update(position, settings), remove(position)}`; `Attachment::{position, settings}` carries no native identifier (the host's `TextNumberAttachmentId` stays out) | ✅ 0371 |

## Pages text boxes, annotations, fields, shapes and comments out of the host (record 0413)

Wave W4-P of record 0363 moves the host's Pages text-box, shape and
drawable-comment editors into `litchi-pages` families, ported from the
Numbers sheet-drawable families of records 0378 and 0378a. Text boxes are
body-anchored drawables selected by text-flow position, never by a native
identifier.

| Name | Meaning | Status |
|---|---|---|
| `litchi_pages::DrawableSelector::TextBox` (`text_box`) | One body text box (a `TSWP.ShapeInfoArchive` flagged as a text box that owns its storage) by its text-flow position among body text boxes; `drawable::properties`, `drawable::text` and `drawable::style` accept it | ✅ 0413 |
| `litchi_pages::drawable::style::{Property::TextColumns, PropertyKind::TextColumns, Columns}` | The text frame's columns on the shared `litchi_iwa_style::shape` engine; with `TextLayout` they replace the host's `{,set_,reset_}text_box_{text_layout,columns}` | ✅ 0413 |
| `litchi_iwa_style::shape_path::{path_source, line_path_source, preset, with_preset, Error}` | The host's `shape_path_source`/`shape_preset`/`set_shape_preset` over the shared `litchi_iwa_common::shape::path::Preset`; the Numbers and Keynote authors and the Keynote preset reader use it (their copies are deleted, record 0413b) | ✅ 0413b |
| `litchi_pages::drawable::text_box_text::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, appearance, character, font, paragraph}` | Registered family in `pages`; `Package::{body_text_box_text(drawable, PropertyKind), edit_body_text_box_text(drawable), apply_body_text_box_text}`; `Edit::{set(Property), reset(PropertyKind)}` over the shared paragraph-style engine replace the host's 72 `set_/reset_/text_box_{text,paragraph}_*` methods and the following-style triple; text that is not uniformly styled is `Refusal::MixedParagraphStyles`, a drawable without text `Refusal::NoText` | ✅ 0413 |
| `litchi_iwa_text_wire::list::publish::{Publisher, Workspace, Failure, WorkspaceStyles, STORAGE_MESSAGE_TYPE, STYLESHEET_MESSAGE_TYPE}` | The record 0370b storage-list publisher lifted into the text crate: `Publisher::new(workspace, limits)` with `set_list`, `set_lists`, `set_level`, `set_numbering`, `property`, publishing the list engine's plans through a leaf's `Workspace` (objects, headers, identifiers, registry transition); `litchi_iwa_table::structure::list_workspace::{ListWorkspace, ReadStyles, storage_payload}` implement it over the shared structural working set for Pages text boxes and Numbers text boxes and cells (the Numbers copy is deleted) | ✅ 0413 |
| `litchi_pages::drawable::list::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, TextPosition, ParagraphList…}` | Registered family in `pages`; `Package::{body_text_box_list, body_text_box_lists, body_text_box_list_levels, body_text_box_list_level, body_text_box_list_numbering, body_text_box_list_property, edit_body_text_box_lists, apply_body_text_box_lists}`; `Edit::{set_list, set_lists, set_level, set_numbering, set, reset}` on the W1-D list engine through the shared storage-list publisher `litchi_iwa_text_wire::list::publish`; replaces the host's 30 text-box list methods | ✅ 0413 |
| `litchi_pages::drawable::named_style::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName}` | Registered family in `pages`; `Package::{body_text_box_named_styles, body_text_box_applied_named_style, edit_body_text_box_named_styles, apply_body_text_box_named_styles}`; `Edit::{apply, create, rename_style, delete, delete_applied, redefine}` on the W1-D named-style engine replace the host's eight `*_text_box_named_paragraph_style*` methods | ✅ 0413 |
| `litchi_pages::drawable::language::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, TextPosition, TextLanguage, TextLanguageRun}` | Registered family in `pages`; `Package::{body_drawable_text_languages, body_drawable_text_language(.., TextPosition), edit_body_drawable_text_languages, apply_body_drawable_text_languages}`; `Edit::{set, remove_boundary, reset}` on the W1-D language engine replace the host's five text-box language methods | ✅ 0413 |
| `litchi_pages::drawable::drop_cap::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, TextPosition, DropCap, LineCount, CharacterCount, RaisedLines, Wrap, Padding, Outdent}` | Registered family in `pages`; `Package::{body_drawable_drop_caps, edit_body_drawable_drop_caps, apply_body_drawable_drop_caps}`; `Edit::{set(TextPosition, DropCap), remove(TextPosition)}` on the W1-D Drop Cap engine with the theme's Drop Cap base (`Refusal::MissingBaseStyle` when it lists none) | ✅ 0413 |
| `litchi_pages::drawable::hyperlink::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Hyperlink, TextHyperlinkTarget, TextPosition, TextRange}` | Registered family in `pages`; `Package::{body_drawable_hyperlinks, edit_body_drawable_hyperlinks, apply_body_drawable_hyperlinks}`; `Edit::{add, update(position, ..), remove(position)}` by text-order position on the W1-D ranged smart-field engine | ✅ 0413 |
| `litchi_pages::drawable::date_time::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, DateTimeField, DisplayText, Format, FormatterStyle, Instant, LocaleIdentifier, Settings, UpdatePlan, TextPosition, TextRange}` | Registered family in `pages`; `Package::{body_drawable_date_time_fields, edit_body_drawable_date_time_fields, apply_body_drawable_date_time_fields}`; `Edit::{add, insert, update, remove}` on the W1-D ranged and smart-field engines (the body's own fields stay `litchi_pages::date_time`) | ✅ 0413 |
| `litchi_pages::drawable::annotation::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Highlight, Comment, TextPosition, TextRange}` | Registered family in `pages`; `Package::{body_drawable_highlights, body_drawable_comments, edit_body_drawable_annotations, apply_body_drawable_annotations}`; `Edit::{add_highlight, update_highlight, remove_highlight, add_comment, update_comment, remove_comment, add_reply, update_reply, remove_reply}` by text-order position on the W1-D annotation engine, with the generated annotation author created and removed as the host does | ✅ 0413 |
| `litchi_pages::number_attachment::{Target::Drawable, Refusal::NoText}` | The storage a body text box or shape owns as a number-attachment target; replaces the host's four `*_text_box_number_attachment*` methods | ✅ 0413 |
| `litchi_pages::drawable::lifecycle::Edit::remove` (text boxes and shapes) | Removal now also takes `DrawableSelector::{TextBox, Shape}`: the drawable, its owned storage, stand-ins, attachment, z-order entry, UUIDs and references leave; replaces the host's `remove_text_box` and `remove_body_shape` | ✅ 0413 |
| `litchi_pages::drawable::lifecycle::Edit::{add_text_box, add_shape, add_line}`, `Refusal::{InvalidGeometry, MissingThemeStyle}` | New body text boxes, preset shapes and straight lines anchored at a UTF-16 body position (their own storage inheriting the body's text styles, the theme's first text-box or shape style); with-fill, with-endpoints and with-style creations compose with `drawable::style`; replaces the host's `add_text_box` and `add_body_{rectangle,shape,shape_with_fill,line,line_with_endpoints,line_with_style}` | ✅ 0413b |
| `litchi_pages::drawable::lifecycle::Edit::{duplicate (shapes, lines and text boxes), duplicate_text_box}` | Copies of body shapes, lines and text boxes in the host's graph order, offset 12 pt, with the storage copied (or, for `duplicate_text_box`, the new text set); replaces the host's `duplicate_body_shape` and `duplicate_text_box`. `Package::body_drawable_properties` (preset, line segment, frame, properties) with `body_drawable_text` and `body_drawable_style` replace the host's `body_shapes` listing | ✅ 0413b |
| `litchi_pages::drawable::style::{Edit::set_image_fill, ImageTechnique, Refusal::InvalidImage}` | A body shape's or text box's image fill: the bytes become a new `Data/` member and `DataInfo`, the shape's style variation (the shared `shape` plan layer) takes the fill and owns the data reference, as Keynote's image fills and the media family's data-owner transitions do; replaces the host's `set_body_shape_image_fill`. Replacing or resetting an existing image fill (`Edit::set`/`reset` of `Fill`, or a second `set_image_fill`) releases the old data reference and owner and the unreferenced media, on the shared `shape::property::ReleasingFills` and `shape::stored_fill` (record 0415); only two image fills staged in one edit are `Refusal::ImageFill` | ✅ 0413b, 0415 |
| `litchi_pages::drawable::properties::{Edit::set_preset, Snapshot::preset, Preset}` | A body shape's or text box's preset path, read and replaced at the frame size through `litchi_iwa_style::shape_path`; replaces the host's `{,set_}body_shape_preset` | ✅ 0413b |
| `litchi_pages::drawable::properties::{Edit::set_line_segment, Snapshot::line_segment, Segment}` | A native straight line's directed segment: the read derives it from the frame and path width, the verb rewrites both through the shared `shape_style_codec::rewrite_line`, keeping the stored direction; replaces the host's `{,set_}body_line_segment` | ✅ 0413 |
| `litchi_pages::drawable::comment::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, DrawableComment}` | Registered family in `pages`; `Package::{body_drawable_comment, edit_body_drawable_comment, apply_body_drawable_comment}` over body text boxes, shapes, images, movies and audio; `Edit::{set, clear, add_reply, set_reply, remove_reply}` with replies by ordinal (no native storage identifier) on the host's copy-on-write comment graph with the generated author lifecycle; replaces the host's `*_drawable_comment*` routes and, with the per-selector read, `drawables` | ✅ 0413 |

## Numbers sheet drawables out of the host (record 0378)

Wave W2-A of record 0363 moves the host's Numbers sheet-drawable editors
into `litchi-numbers` families. Host runs of `x()`/`set_x()`/`flip_x()`/
`restore_x()` per drawable kind collapse into one property family; selectors
carry a list position or a kind and position, never a native identifier.

| Name | Meaning | Status |
|---|---|---|
| `litchi_numbers::DrawableSelector::{Position, Of { kind, position }}` (`position`, `index`, `of`, `text_box`, `shape`, `image`, `movie`, `audio`, `chart`, `kind`, `as_position`, `as_index`; also `litchi_numbers::sheet::drawable::DrawableSelector`; `From<usize>`, `From<Position>`, `From<ImageSelector>`, `From<ChartSelector>`) | One sheet drawable by its back-to-front list position or by its position among one kind — the 0363 §3.2 Numbers drawable selector | ✅ 0378 |
| `litchi_numbers::sheet::drawable::{Kind, Drawable, Geometry, GeometryFlags, Properties, Point, Size, FlipAxis, MediaPlaybackSettings, MediaLoopMode, MediaVolume, ImageOptions, MovieOptions, AudioOptions}` | `Kind::{TextBox, Shape, Image, Movie, Audio, Chart, Table, Other}` (`has_text`, `name`); `Drawable::{selector, kind_selector, kind, z_position, geometry}`; the archive-free frame, flag set and shared properties as in Pages | ✅ 0378 |
| `litchi_numbers::sheet::drawable::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal}` | Registered family in `numbers-minimal`/`numbers-full`; `Package::{sheet_drawable_list, edit_sheet_drawables, apply_sheet_drawables}`; `Edit::{move_to(drawable, Position), remove(drawable), add_text_box(Geometry, &str), add_shape(Preset, Geometry, &str), add_line(Segment), duplicate(drawable), duplicate_text_box(drawable, &str), add_image(&str, &[u8], ImageOptions), add_movie((&str, &[u8]), (&str, &[u8]), MovieOptions), add_audio(&str, &[u8], AudioOptions)}`, several per edit; `Refusal::InvalidMedia`, selectors resolve against the base, new drawables join the front; `Refusal::InvalidGeometry`; `Preset` is the shared `litchi_iwa_common::shape::path::Preset`; replaces `sheet_drawable_order`, `set_sheet_drawable_order`, `move_sheet_drawable`, `remove_sheet_{text_box,shape,image,movie,audio}`, `add_sheet_{text_box,rectangle,shape}`, `duplicate_sheet_{image,movie,audio,text_box}`, `add_sheet_{image,movie,audio}`; removal clears the drawable's direct comment and releases unreferenced embedded data. W4-D (0416): `add_line(Segment)` replaces `add_sheet_line`; with `style` it replaces `add_sheet_line_with_{endpoints,style}` and `add_sheet_shape_with_fill`; `duplicate` also copies shapes and lines (`duplicate_sheet_shape`) | ✅ 0378, 0416 |
| `litchi_numbers::sheet::drawable::properties::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Snapshot}` | Registered family; `Package::{sheet_drawable_properties, edit_sheet_drawable_properties, apply_sheet_drawable_properties}(sheet, drawable)`; `Edit::{set_geometry, set_position, flip, restore_original_size, set_properties, set_playback}` replace the host's 30 text-box/shape/image/movie/audio/chart geometry, flip, Original Size, property, audio position and playback methods; a verb that does not apply to the kind is `Refusal::UnsupportedOperation`. W4-D (0416): `Edit::{set_preset(Preset), set_line_segment(Segment)}`, `Snapshot::{preset, line_segment}` and the re-exported `Preset`, `Segment` replace `{,set_}sheet_shape_preset` and `{,set_}sheet_line_segment` through the shared `litchi_iwa_style::shape_path`; with `Package::sheet_drawable_list` the snapshot replaces the host's `sheet_{shapes,images,movies,audio}` lists | ✅ 0378, 0416 |
| `litchi_numbers::sheet::drawable::style::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, Fill, Stroke, Effects, Layout, Columns}` | Registered family in `numbers-minimal`/`numbers-full` (W2-A 0380); `Package::{sheet_drawable_style, edit_sheet_drawable_style, apply_sheet_drawable_style}(sheet, drawable)`; `Edit::{set(Property), reset(PropertyKind)}` over the shared `litchi_iwa_style::shape` graph for shape fill, stroke, effects and text layout plus text-box columns and text layout; image fills, presets, lifecycle and other drawable kinds remain typed refusals or separate families. W4-D (0416): `Property::{Shadow, LineEndpoints}` (`Shadow`, `Endpoint`, `Endpoints` re-exported) replace `{,set_,reset_}sheet_shape_shadow` and `{,set_,reset_}sheet_line_endpoints`; endpoints go through the shared `litchi_iwa_style::line_end::{endpoints, set_endpoints}`, a reset stores the undecorated default, a shape that is not a line is `Refusal::UnsupportedProperty` | ✅ 0380, 0416 |
| `litchi_numbers::sheet::drawable::text::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, appearance, character, font, paragraph}` (the last four re-export the `litchi_iwa_text` value modules the property values are built from) | Registered family; `Package::{sheet_text_box_text(sheet, drawable, PropertyKind), edit_sheet_text_box_text(sheet, drawable), apply_sheet_text_box_text}`; `Edit::{set(Property), reset(PropertyKind)}` over the shared paragraph-style engine replace the host's 69 `set_/reset_/sheet_text_box_{text,paragraph}_*` methods; text that is not uniformly styled is `Refusal::MixedParagraphStyles`, a drawable without text `Refusal::NoText` | ✅ 0378 |
| `litchi_numbers::sheet::drawable::plain_text::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, TextPosition, TextSpan}` | Registered family; `Package::{sheet_drawable_plain_text, edit_sheet_drawable_plain_text, apply_sheet_drawable_plain_text}(sheet, drawable)` over text boxes and shapes; `Edit::{replace, insert, delete, set, clear}`, one UTF-16 splice per edit, as Pages' `drawable::text`; replaces `{set,replace,clear}_sheet_{text_box,shape}_text`; a splice removing a referenced object is `Refusal::DependentContent` | ✅ 0378 |
| `litchi_numbers::sheet::chart::style::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, RadarGridShape, corner, donut, gaps, legend, pie, radar, chart3d}` | Registered family; `Package::{sheet_chart_style(sheet, chart, PropertyKind), edit_sheet_chart_style(sheet, chart), apply_sheet_chart_style}`; `Property::{PieStartAngle, DonutInnerRadius, RadarStartAngle, SeriesGap3d, Rotation3d, Depth3d, RoundedCorners, LegendFrame, LegendVisible, BorderVisible, IncludesHiddenData, BarShape3d, RadarGridShape, GapSpacing}` over the shared `litchi_iwa_common::{chart, chart3d}` values; `Edit::{set, reset}` (reset writes the native default) on the W1-E chart style engine; a property the chart kind does not expose is `Refusal::UnsupportedKind` | ✅ 0378 |
| `litchi_numbers::sheet::chart::lifecycle::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Kind, Direction}` | Registered family (W4-D 0416); `Package::{sheet_chart_direction(sheet, chart), edit_sheet_chart_lifecycle(sheet, chart), apply_sheet_chart_lifecycle}`; `Edit::{set_kind(Kind), set_direction(Direction)}` rewrite the chart archive's kind and series direction in place through the shared `litchi_iwa_style::chart_lifecycle`; replaces the host's `set_sheet_chart_{kind,direction}`; with `sheet_chart_metadata`, `sheet_chart_data`, `sheet_chart_arrangement` and the drawable properties it replaces the `sheet_charts` list; `sheet::chart::data::Edit::set` replaces `set_sheet_chart_data`; an undefined kind or direction is `Refusal::{UnsupportedKind, UnsupportedDirection}` | ✅ 0416 |
| `litchi_numbers::sheet::drawable::language::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, TextPosition, TextLanguage, TextLanguageRun}` | Registered family; `Package::{sheet_drawable_text_languages, sheet_drawable_text_language(.., TextPosition), edit_sheet_drawable_text_languages, apply_sheet_drawable_text_languages}`; `Edit::{set(TextPosition, TextLanguage), remove_boundary(TextPosition), reset()}` on the W1-D language engine; replaces the host's five text-box language methods | ✅ 0378 |
| `litchi_numbers::sheet::drawable::list::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, TextPosition, ParagraphList…}` | Registered family; `Package::{sheet_text_box_list, sheet_text_box_lists, sheet_text_box_list_levels, sheet_text_box_list_level, sheet_text_box_list_numbering, sheet_text_box_list_property, edit_sheet_text_box_lists, apply_sheet_text_box_lists}`; `Edit::{set_list, set_lists, set_level, set_numbering, set, reset}` on the W1-D list engine through the shared storage-list publisher (record 0370b); replaces the host's 30 text-box list methods | ✅ 0378 |
| `litchi_numbers::sheet::drawable::drop_cap::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, TextPosition, DropCap, LineCount, CharacterCount, RaisedLines, Wrap, Padding, Outdent}` | Registered family; `Package::{sheet_drawable_drop_caps, edit_sheet_drawable_drop_caps, apply_sheet_drawable_drop_caps}`; `Edit::{set(TextPosition, DropCap), remove(TextPosition)}` on the W1-D Drop Cap engine; `Refusal::MissingBaseStyle` when the theme lists no Drop Cap base | ✅ 0378 |
| `litchi_numbers::sheet::drawable::hyperlink::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Hyperlink, TextHyperlinkTarget, TextPosition, TextRange}` | Registered family; `Package::{sheet_drawable_hyperlinks, edit_sheet_drawable_hyperlinks, apply_sheet_drawable_hyperlinks}`; `Edit::{add(TextRange, TextHyperlinkTarget), update(position, TextRange, TextHyperlinkTarget), remove(position)}` (positions in text order; no native field identifier) on the W1-D ranged smart-field engine | ✅ 0378 |
| `litchi_numbers::sheet::drawable::date_time::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, DateTimeField, DisplayText, Format, FormatterStyle, Instant, LocaleIdentifier, Settings, UpdatePlan, TextPosition, TextRange}` | Registered family; `Package::{sheet_drawable_date_time_fields, edit_sheet_drawable_date_time_fields, apply_sheet_drawable_date_time_fields}`; `Edit::{add(TextRange, Settings), insert(TextPosition, DisplayText, Settings), update(position, TextRange, Settings), remove(position)}` on the W1-D ranged and smart-field engines | ✅ 0378 |
| `litchi_numbers::sheet::drawable::annotation::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Highlight, Comment, TextPosition, TextRange}` | Registered family; `Package::{sheet_drawable_highlights, sheet_drawable_comments, edit_sheet_drawable_annotations, apply_sheet_drawable_annotations}`; `Edit::{add_highlight, update_highlight, remove_highlight, add_comment, update_comment, remove_comment, add_reply, update_reply, remove_reply}` by text-order positions (no native identifier) on the W1-D annotation and ranged engines; the generated annotation author is created and removed as the host does | ✅ 0378 |
| `litchi_numbers::sheet::drawable::named_style::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName}` | Registered family; `Package::{sheet_text_box_named_styles, sheet_text_box_applied_named_style, edit_sheet_text_box_named_styles, apply_sheet_text_box_named_styles}`; `Edit::{apply, create, rename_style, delete, delete_applied, redefine}` on the W1-D named-style engine; the host `*_sheet_text_box_named_paragraph_style*` routes collapse into one transaction | ✅ 0378 |
| `litchi_numbers::sheet::drawable::comment::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, DrawableComment}` | Registered family; `Package::{sheet_drawable_comment, edit_sheet_drawable_comment, apply_sheet_drawable_comment}`; `Edit::{set, clear, add_reply, set_reply, remove_reply}` with replies by ordinal (no native storage identifier); the host copy-on-write comment graph, generated author lifecycle and save-token stamps | ✅ 0378 |
| `litchi_numbers::sheet::drawable::media::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, AssetId, Embedded, Type}` | Registered family; `Package::{media_assets, sheet_media_assets, extract_media, edit_media, apply_media}`; `Edit::{replace_asset, replace_data, replace_poster}` — one replacement verb for the host `replace_media` and `replace_sheet_{image,movie,audio}_data` routes, by asset identifier or drawable selector; the `Data/` member and its `DataInfo` digest and length change, the identifier stays | ✅ 0378 |
| `litchi_numbers::sheet::chart::series::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, ConnectionLine, WedgeExplosion, LabelDistance, LabelVisibility, Symbol, SymbolShape, SymbolSize, Trendline, TrendlineType, PolynomialOrder, MovingAveragePeriod, Visibility, Stroke, StrokePattern, SymbolFill, Fill, Gradient, error_bar}` | Registered family; `Package::{sheet_chart_series(sheet, chart, PropertyKind), edit_sheet_chart_series, apply_sheet_chart_series}`; `Edit::set(Property)` with one value per series on the W1-E graph engine (`litchi_iwa_style::series`) and series-style write rules; the host all-series and one-series `set_sheet_chart_*` routes of connection lines, wedge explosions, label distances, data symbols, trendlines, error bars, pie labels, value-label visibility, series strokes, fills, symbol fills and symbol outlines collapse into it; `Edit::reset(PropertyKind, series)` replaces `reset_sheet_chart_series_{stroke,fill,symbol_fill,symbol_outline}`; `Refusal::{UnsupportedKind, SeriesCount, SeriesNotFound, NotResettable, SharedStyle, ImageFill}` | ✅ 0378 |
| `litchi_numbers::sheet::chart::axis::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, Axis, TickMarkLocation, Visibility}` | Registered family; `Package::{sheet_chart_axis(sheet, chart, Axis, PropertyKind), edit_sheet_chart_axis, apply_sheet_chart_axis}`; `Edit::set(Axis, Property)` on the primary value or category axis style (line, major/minor gridlines, minor tick marks, tick-mark location, the value axis minimum label) or non-style (labels, the category axis series names), refusing a shared axis style; replaces `set_sheet_chart_axis_{line,major_gridlines,minor_gridlines,minor_tick_marks}_visible`, `set_sheet_chart_axis_tick_mark_location`, `set_sheet_chart_value_axis_minimum_label_visible`, `set_sheet_chart_axis_labels_visible`, `set_sheet_chart_category_axis_series_names_visible` and their reads | ✅ 0378 |

## W1-E chart title and axis-title facades (record 0375b)

The bounded generated-extension title and category/value-axis-title rewrite
engine lives in `litchi_iwa_style::chart_text`. Format owners retain selector
resolution, chart-reference ownership and exact package publication; caption
authoring is explicitly deferred and Keynote keeps its native codec route.

| Canonical name | Meaning | Status |
|---|---|---|
| `litchi_numbers::sheet::chart::{title, axis_title}::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal}` | Registered selector-first families; `Package::{sheet_chart_title, edit_sheet_chart_title, apply_sheet_chart_title}` and `Package::{sheet_chart_axis_title, edit_sheet_chart_axis_title, apply_sheet_chart_axis_title}` | ⏳ 0375b: parity and native queue pending |
| `litchi_pages::chart::{title, axis_title}::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal}` | Registered selector-first families; `Package::{body_chart_title, edit_body_chart_title, apply_body_chart_title}` and `Package::{body_chart_axis_title, edit_body_chart_axis_title, apply_body_chart_axis_title}` | ⏳ 0375b: parity and native queue pending |
| `litchi_pages::chart::lifecycle::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Kind, Direction}` | Registered family (W4-D 0416); `Package::{body_chart_direction, body_chart_geometry, edit_body_chart_lifecycle, apply_body_chart_lifecycle}(BodyChartSelector)`; `Edit::{set_kind(Kind), set_direction(Direction)}` on the shared `litchi_iwa_style::chart_lifecycle`; replaces the host's `set_body_chart_{kind,direction}`; with `body_chart_metadata`, `body_chart_data`, `body_chart_arrangement` and `body_chart_geometry` it replaces the `body_charts` list; `chart::data::Edit::set` replaces `set_body_chart_data` | ✅ 0416 |

## W2-A Numbers chart-font facade (record 0385)

The Numbers chart-font facade is a thin selector adapter over the shared
`litchi_iwa_style::chart_font` graph engine. Its admission and typed-error
surface remains provisional until the shared chart-font bounded-limit contract
is accepted; no legacy-host fallback is implied by a refusal.

| Canonical name | Meaning | Status |
|---|---|---|
| `litchi_numbers::sheet::chart::font::{Property, PropertyKind, Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, ChartFont, TextPointSize}` | Selector-first `Package::{sheet_chart_font, edit_sheet_chart_font, apply_sheet_chart_font}` for chart font identity/traits and point size; `Edit::{set, reset}` uses family-scoped reads, exact-source patches, inverse/no-op behavior and candidate readback. The facade does not expose native IDs, archives or protobuf values; mixed unrelated font/size state is not rejected merely because the other family is non-uniform. | ⏳ 0385: host differential and public-contract tests prepared; shared bounded-error contract, native queue, final performance and integrated gate pending |

## W2-B Pages chart-series symbols (record 0400)

The Pages facade is a thin selector adapter over the existing bounded
`litchi_iwa_style::series::set_symbols` graph engine. Host single-series and
all-series symbol setters collapse into one semantic vector property; native
identifiers and archive values do not cross the public boundary.

| Canonical name | Meaning | Status |
|---|---|---|
| `litchi_pages::chart::series::{Property, PropertyKind, Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Symbol, SymbolShape, SymbolSize}` | Selector-first `Package::{body_chart_series, edit_body_chart_series, apply_body_chart_series}` for per-series visibility, shape and optional point size. `Property::Symbols(Vec<Option<Symbol>>)` is the sole write value; `Edit::set` preserves exact-source no-op/inverse/reopen/locality and maps unsupported kinds, series-count mismatch, shared styles and resource ceilings to typed refusals/errors. | ⏳ 0400: Pages host differential and scoped facade checks pass; native queue, full-facade performance and final integrated gate pending |

## W4-C chart style families on the shared chart engines (record 0414)

The chart-level, axis-level and per-series chart properties of all three
formats are thin selector adapters over `litchi_iwa_style::{chart_style,
series, series::style}`. The property values are the engines' shared
vocabulary (`chart_style::{chart, axis}::{Property, PropertyKind}`), re-exported
at each family path; host single-series and all-series setters collapse into
one per-series vector property, host per-property setters into one
`Edit::set` per family. Native identifiers and archive values do not cross the
public boundary.

| Canonical name | Meaning | Status |
|---|---|---|
| `litchi_keynote::chart::style::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, RadarGridShape, corner, donut, gaps, legend, pie, radar, chart3d}` | Registered family; `Package::{slide_chart_style(slide, chart, PropertyKind), edit_slide_chart_style(slide, chart), apply_slide_chart_style}`; `Edit::{set(Property), reset(PropertyKind)}` (reset writes the native default) for pie/donut/radar angles, 3D gap/rotation/depth/bar shape, rounded corners, legend frame, border, radar grid and gaps; a shared style is `Refusal::SharedStyle`, a property the kind lacks `Refusal::UnsupportedKind` | ⏳ 0414: Keynote host differential (Tier B) passes; native queue pending |
| `litchi_keynote::chart::axis::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, Visibility, TickMarkLocation}` (beside the existing axis values) | Registered family; `Package::{slide_chart_axis(slide, chart, Axis, PropertyKind), edit_slide_chart_axis, apply_slide_chart_axis}`; `Edit::set(Axis, Property)` for the primary axis's line, gridlines, minor ticks, tick-mark location, minimum label, labels and series names | ⏳ 0414: Tier B; native queue pending |
| `litchi_keynote::chart::series::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, …values}` | Registered family; `Package::slide_chart_series(slide, chart, PropertyKind)` (was symbols-only, no kind); `Edit::{set(Property), reset(PropertyKind, series)}` for connection lines, wedge explosions, label distances, symbols, trendlines, error bars, pie labels, value labels, value-label affixes and number formats, strokes, fills, symbol fills and outlines (`set` now takes a `Property`, was a symbol vector) | ⏳ 0414: Tier B; native queue pending |
| `litchi_pages::chart::style::{…}` and `litchi_pages::chart::axis::{…}` | Registered families with the Keynote meanings over a `BodyChartSelector`: `Package::{body_chart_style, edit_body_chart_style, apply_body_chart_style}`, `Package::{body_chart_axis, edit_body_chart_axis, apply_body_chart_axis}`; by-value `Edit` builders as the other Pages chart families | ⏳ 0414: byte-identical to the host on the tracked fixture; native queue pending |
| `litchi_pages::chart::series::{…}` | `Package::body_chart_series(selector, PropertyKind)` (was symbols-only); `Edit::{set(Property), reset(PropertyKind, series)}` over the same per-series property set as Keynote | ⏳ 0414: native queue pending |
| `litchi_numbers::sheet::chart::{style, axis}::{Property, PropertyKind, RadarGridShape}` | Now re-exports of the engine values (same names and variants); `sheet::chart::series::Property::{ValueLabelAffixes, ValueLabelNumberFormats}` added with `LabelAffixes`, `NumberFormat` | ⏳ 0414: native queue pending |
| `{litchi_keynote::chart, litchi_pages::chart, litchi_numbers::sheet::chart}::axis::{GridlineStroke, LabelAngle, LabelAffixes, NumberFormat, LabelPosition3d, Stroke}`; `axis::Property::{MajorGridlineStroke, MinorGridlineStroke, LabelAngle, LabelAffixes, NumberFormat, LabelPosition3d}`; `style::Property::BorderStroke` | Slice 2: gridline strokes (`GridlineStroke::{Inherited, None, Stroke}`), label angle, label affixes and number format, the 3D value-axis label position (`Refusal::UnsupportedKind` on a 2D chart) and the chart border stroke; axis values are no longer `Copy` (`axis::Patch::{before, after}` return references) | ⏳ 0414 slice 2: native queue pending |
| `{litchi_keynote::chart, litchi_pages::chart, litchi_numbers::sheet::chart}::legend_style::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Refusal, Property, PropertyKind, LegendFill, LegendStroke, LegendShadow, Fill, Gradient, Drop, Stroke}` | Registered families; `Package::{slide,body,sheet}_chart_legend_style(.., PropertyKind)`, `edit_*_chart_legend_style`, `apply_*_chart_legend_style`; `Edit::{set(Property), reset(PropertyKind)}` (reset restores inheritance) on `litchi_iwa_style::legend_style` (derive a private legend-style variation, write, collapse); image fills `Refusal::ImageFill`. Numbers `sheet::chart::style` gains `Refusal::SharedStyle` | ⏳ 0414 slice 3: native queue pending |
| `{litchi_keynote::chart, litchi_pages::chart, litchi_numbers::sheet::chart}::style::{Lighting3d}`; `style::Property::{BackgroundFill, Lighting3d}` | Slice 4: the chart background fill (`Fill::{None, Solid, Gradient}`; the native default is the 67%-gray-to-white linear gradient, whose write removes the stored fill; image fills `Refusal::ImageFill`) and the 3D lighting preset (`Lighting3d::{Default, SoftLight, SoftFill, MediumCenter, MediumRight, MediumLeft, Glossy}`, `inspector_label()`; a 2D chart is `Refusal::UnsupportedKind`); chart `style::Property` is no longer `Copy` | ⏳ 0414 slice 4: native queue pending |
| `{litchi_keynote::chart, litchi_pages::chart, litchi_numbers::sheet::chart}::style::{Shadow, Drop}`, `style::Property::Shadow`; `axis::{Frequency, Interval, Layout}`, `axis::Property::CategoryLabels` | Slice 5: the chart shadow (`Shadow::{None, IndividualSeries(Drop), Grouped(Drop)}`, `Shadow::NATIVE_DEFAULT`; stored in every series style and the chart style; a series style another chart uses is `Refusal::SharedStyle`) and the category label layout on the category axis (`Layout::new(Frequency, show_last_category)`; `Frequency::None` hides the labels and keeps the stored interval). `litchi_iwa_style::series::chart_series_styles`, `series::style::{series_shadow, rewrite_series_shadow, NATIVE_DEFAULT_SHADOW}`, `chart_style::chart::shadow`, `chart_style::axis::{category_label_style, category_labels, rewrite_category_labels}` are the payload-level engine rules | ⏳ 0414 slice 5: native queue pending |
| `{litchi_keynote::chart, litchi_pages::chart, litchi_numbers::sheet::chart}::legend_style::{LegendFont, LegendFontSize, Font, TextPointSize}`; `legend_style::Property::{Font, FontSize}` | Slice 6: the legend label typography (`LegendFont::{Inherited, Font { font, bold, italic }}`, `LegendFontSize::{Inherited, Size(TextPointSize)}`) on `litchi_iwa_style::legend_style` (the host's legend-font graph: a private paragraph-style variation the legend style selects by index, inside the legend-style derive and collapse). `litchi_iwa_text_wire::paragraph_style::{VariationOverrides, variation_overrides, variation_with}` expose a variation's full override set | ⏳ 0414 slice 6: native queue pending |
| `litchi_iwa_style::chart_style::{chart, axis, read_chart, write_chart, read_axis, write_axis}`, `litchi_iwa_style::series::{read, write, visible_series, Refusal, style::{Kind, Value, effective, rewrite, read, write}}` | Format-neutral engines (not re-exported by a facade except the value enums): bounded workspace reads/writes returning the source-derived `MutationScope` (now with `allowed_chart_fields`) the owners publish under | engine, record 0414 |

## iWork facade retirement and iWork open options (record 0306)

| Canonical name | Banned | Status |
| --- | --- | --- |
| `litchi::open`/`open_with` → `File::{Pages, Keynote, Numbers}` for packages **and** package directories | `litchi::iwork` (module), `litchi::iwork::{Document, Snapshot, Summary, Table, Slide, Section, SectionKind, TextRole, Text, Value, CellView, Options, SourceLimits, SnapshotLimits, Error, ErrorKind, Stage, Resource, Format}` | ✅ 0306: deleted; `facade-iwork` profile bans `litchi::iwork` |
| `litchi_ops::OpenOptions::{pages_options, keynote_options, numbers_options}` taking the leaf `OpenOptions` | — | ✅ 0306 |
| `Package::save_directory(path)` on `litchi_keynote` and `litchi_numbers` (as on `litchi_pages`), `SaveError::{Bundling, Bundle}` | — | ✅ 0306 |
| `litchi_iwa_archive::bundle::{capture, save, converted_limits, is_directory, CaptureLimits, CaptureLimitKind, CaptureError, INDEX_ZIP}` | — | ✅ 0306 (shared by the three leaves) |

## XLS anchor behaviour (record 0316)

`[MS-XLS]` 2.5.193 sets `fMove`/`fSize` when a shape is *kept intact* as its
cells move or resize. The read enum named the three states the other way
round (`Fixed` for 0, `MoveAndSize` for 3), and the writer carried a second
enum with the same inversion. One enum now serves both, named after Apache
POI's `AnchorType` and ECMA-376 `ST_EditAs`; wire values are unchanged.

| Canonical name | Banned | Status |
| --- | --- | --- |
| `litchi_xls::AnchorBehavior::MoveAndResize` (wire 0, `twoCell`) | `AnchorBehavior::Fixed` | ✅ 0316 |
| `litchi_xls::AnchorBehavior::MoveWithoutResize` (wire 2, `oneCell`) | `AnchorBehavior::Size` | ✅ 0316 |
| `litchi_xls::AnchorBehavior::Absolute` (wire 3, `absolute`) | `AnchorBehavior::MoveAndSize` | ✅ 0316 |
| `litchi_xls::writer::shape::Anchor::{new, cells, behavior}` take and return `AnchorBehavior` | `litchi_xls::writer::shape::Behavior` | ✅ 0316 |

The `xls-minimal` profile bans the old spellings. See
[0316](unified-ops/0316-xls-odp-structural-follow-ups.md).

## ODP character formatting through styles, links and fields (record 0316)

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_odp::scene::CommonStyles::parse` | Indexes `office:styles` of `styles.xml` or a flat document, a family's `style:default-style` under the empty name | ✅ 0316 |
| `litchi_odp::scene::AutomaticStyles::{resolve, resolve_paragraph, resolve_run}` | Character properties through `style:parent-style-name` (automatic, then common styles) and the family's default style; a link's text span over its link style | ✅ 0316 |
| `litchi_odp::scene::TextProperties::inherit` | Fills unset properties from a parent; the font is one property | ✅ 0316 |
| `litchi_odp::scene::Run::link_text_style_name` | The style of the one `text:span` a link's content is | ✅ 0316 |
| `RunKind::Field` for a `text:span` holding one field with text | The span's style formats the field; `Change::SetFieldText` rewrites the field inside it | ✅ 0316 |
| `Change::SetTextProperties` on link and field runs | A link keeps its link style and its content span carries the formatting; a bare field is wrapped in a span | ✅ 0316 |

Unified ODP `shape`, `placeholder`, `tc` and `notes` nodes read and set
`bold`, `italic`, `underline`, `size`, `color` and `font` as the value every
run of their text shares; paragraph and run reads resolve their style chain
through common styles, a run over its paragraph.

## XLS worksheet repairs (record 0316)

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_xls::WorksheetDefect::{ZoomLevel, SharedFormulaFlag}`, `Defect::Worksheet` | Excel-compatible repairs under every `Leniency`: a WINDOW2 zoom outside `10..=400` reads as the default; a stale `fShrFmla` on a formula with its own tokens is ignored | ✅ 0316 |

## DOCX located parts: styles, numbering, headers, footers and notes (record 0319)

L-DOCX-R5/R6. Owned, source-located catalogs read a part once and keep
every entry's byte range, so a read answers the typed direct-property keys
and an edit splices the raw part. They are owner-neutral: both owners and
the eager `Edit` return the same catalog type. Edit verbs stage on the eager
package-bound `document::Edit` and travel in its `Patch` as graph changes
and main-story element replacements; no new transaction family is
introduced. Native identifiers (`w:styleId`, `w:abstractNumId`, `w:numId`,
note `w:id`) are the selectors these verbs take, as `remove_comment(id)`
does.

| Name | Notes | Status |
|---|---|---|
| `styles::catalog::{Catalog, Entry, Flag}` | `Catalog::{from_xml, xml, entries, get, resolve, paragraph_properties, run_properties, is_strict}`; `resolve` matches an id or a display name | ✅ 0319 |
| `styles::catalog::{NewStyle, StyleChange}` | `with_*` builders; an inner `None` removes an element | ✅ 0319 |
| `Edit::{style_catalog, add_style, set_style, remove_style}` | | ✅ 0319 |
| `numbering::catalog::{Catalog, AbstractNum, Num, Lvl}` | `Catalog::{abstract_nums, nums, abstract_num, num, level_of, format_of, level_paragraph_properties, level_run_properties, next_abstract_num_id, next_num_id}`; `Lvl` is named after `w:lvl` (the typed `numbering::Level` stays the parsed model) | ✅ 0319 |
| `numbering::catalog::{LevelChange, AbstractNumChange, NewNum, NewLevel, NewAbstractNum}` | `NewLevel::{start, paragraph}` read a template level back (0325) | ✅ 0319, 0325 |
| `Edit::{numbering_catalog, set_level, set_abstract_num, set_num_definition, add_num, add_abstract_num, add_level, remove_num, remove_abstract_num, remove_level}` | | ✅ 0319 |
| `header_footer::catalog::Entry` | role, relationship identifier and part name, in relationship order | ✅ 0319 |
| `header_footer::Story::from_shared_bytes` | | ✅ 0319 |
| `document::NewStory`; `Edit::{header_footer_entries, add_header_footer, remove_header_footer, set_header_footer_kind, set_story_text, format_story}` | | ✅ 0319 |
| `document::{NewNote, NoteKind}`; `Edit::{add_note_at, set_note_text, remove_note}` | | ✅ 0319 |
| `footnote::Note::{footnotes_from_xml, endnotes_from_xml}` | | ✅ 0319 |
| `Package::{style_catalog, numbering_catalog, header_footer_entries, part_xml}` and `source_backed::Package::{style_catalog, numbering_catalog, header_footer_entries, part_xml, footnotes, endnotes}` | both owners | ✅ 0319 |
| `document::ElementStep::SectionProperties` | the body-final `w:sectPr`, position 0, first step only | ✅ 0319 |
| `Edit::insert_break_at(paragraph, at, RunBreakType) -> Position` | a new run holding one `w:br`, before direct run `at` or appended; returns the new run's position (record 0319, P1 marks) | ✅ 0319 |
| `Edit::{insert_bookmark_at(paragraph, at, name, text) -> i64, insert_field_at(paragraph, at, instruction, result) -> Position}` (snapshot and package edits); `document::fields::Update::{Instruction, RenameBookmark, RemoveBookmark}` | a bookmark around optional text (the identifier after the largest), a complex field with its cached result, before direct run `at` or appended; field instruction sets mark the field `w:dirty`; bookmark renames and marker-only removals, spliced through `refresh_fields` (record 0325) | ✅ 0325 |
| `Edit::edit_story(StoryTarget, FnOnce(&mut raw::Tree) -> Result<T, E>)`, `document::StoryTarget::{Part, Note}` | runs the main document's element-path verbs on a header, footer or note story wrapped as a document body, and splices the result back into the part; an unchanged story stages nothing (record 0325) | ✅ 0325 |
| `source_backed::Document::{bookmarks, text_boxes}` | the main story's bookmarks and text boxes, as the eager `Document` reads them | ✅ 0319 |

## XLSX worksheet feature families (record 0320)

Five scoped families publish worksheet features through immutable
part-level edits sharing one staged engine: each `Edit` stages part
replacements, `commit()` returns the triad `package()`, `patch()`,
`diagnostics()` with `into_package()`/`into_parts()`, and an unchanged
staging returns the source package (an exact no-op keeps every byte).

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_xlsx::conditional_formatting::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`, `Package::{edit_conditional_formats, apply_conditional_formats}` | `Edit::{formats, set_formats, add_differential}`; `DifferentialStyle::{read, is_empty}`, `Underline` | ✅ 0320; family registered in the gate |
| `litchi_xlsx::data_validation::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`, `Package::{edit_data_validations, apply_data_validations}` | `Edit::{validations, set_validations}` | ✅ 0320; family registered in the gate |
| `litchi_xlsx::auto_filter::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`, `Package::{edit_auto_filter, apply_auto_filter}` | `Edit::{auto_filter, set_auto_filter}` | ✅ 0320; family registered in the gate |
| `litchi_xlsx::workbook::comments::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`, `Package::{edit_comments, apply_comments}` | `Edit::{comments, set, add, remove, set_runs, add_runs}` (`add` refuses an annotated cell and returns the row-major position); splices notes and their VML note shapes, keeping the other notes' bytes; rich notes as `TextRun` with `RunFont` and `RunColor` (`Comment::runs` reads them, record 0325); whole typed graph `Edit::{replace, set_comment, set_author, rename_author, clear, reset}` (the former `workbook::comments::Transaction`, 0360) | ✅ 0320, 0325, 0360; family registered in the gate |
| `litchi_xlsx::drawing::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`, `Package::{drawing_elements, edit_drawings, apply_drawings}` | `Edit::{elements, add_picture, add_shape, set, remove}`; read model `Element`, `ElementKind`, `Placement`, `Marker`, `Image`, `Fill`, `Line`, `Font`, `Align`, `VerticalAlign`; specs `PictureSpec`, `ShapeSpec`, `Change` | ✅ 0320; family registered in the gate |
| `litchi_xlsx::Worksheet::comments`, `SourceWorksheet::{comments, page_breaks, drawing_elements}` | reads on both owners | ✅ 0320 |

## PPTX theme, columns, hyperlinks, media, equations, masters and layouts (record 0321)

New verbs live on the primary `opened::Edit` (no new transaction family);
reads live on the owners and slide views that already answer charts and
SmartArt. Slide hyperlink targets are semantic (`Target::{Url, Slide,
Action}`); the authored relationship ID of a `Click` stays crate-private.

| Name | Status |
|---|---|
| `litchi_pptx::Package::theme`, `SourceBackedPresentation::theme`, `shape::theme::presentation_theme` | adopted (0321) |
| `litchi_pptx::opened::Edit::{theme, set_theme_colors, set_theme_fonts}` | adopted (0321) |
| `litchi_pptx::opened::Edit::{set_master_name, set_layout_name, set_master_shape_text, set_layout_shape_text}` | adopted (0321) |
| `litchi_pptx::table::Table::grid_columns`; `opened::Edit::{set_table_column_width, insert_table_column, remove_table_column}` | adopted (0321) |
| `litchi_pptx::hyperlinks::{Click, Link, Target}`; `Click::{of_shape, from_text, action, tooltip}` | adopted (0321) |
| `litchi_pptx::slide::Slide::hyperlink`, `SourceSlide::hyperlink`, `source::Slide::hyperlink` | adopted (0321) |
| `litchi_pptx::opened::Edit::{set_shape_hyperlink, set_run_hyperlink}` | adopted (0321); `set_shape_hyperlink` also writes a picture's or a group's `p:cNvPr` click; `hyperlinks::mentions_click` is the cheap no-click pre-check of `Click::of_shape` (0325) |
| `litchi_pptx::media_parts::{Summary, summary}`; `opened::{Edit::add_media, MediaFile}` | adopted (0321) |
| `litchi_pptx::shape::math::{Equation, equations}`; `Slide::equations`, `SourceSlide::equations`, `source::Slide::equations` | adopted (0321) |

See [0321](unified-ops/0321-pptx-element-gaps.md).

## PPTX P0 residue: text structure, rows, cell properties, comments, pictures (record 0333)

New verbs live on the primary `opened::Edit`, and the slide-local ones also
on the source primary `source::Edit` (no new transaction family). The verb
names follow the glossary's `insert_*`/`remove_*`/`set_*` spellings of the
other text and table owners; positions are `litchi_core::Position`s, and
native comment identities stay behind the positional verbs.

| Name | Status |
|---|---|
| `litchi_pptx::opened::TextBody::{Shape, Cell, Notes}`; `opened::Edit::{insert_paragraph, remove_paragraph, insert_run, remove_run}`; `source::Edit::{insert_paragraph, remove_paragraph, insert_run, remove_run}` (slide bodies) | adopted (0333) |
| `litchi_pptx::opened::TextTarget::CellParagraph` | adopted (0333) |
| `litchi_pptx::opened::{CellFill, CellMargin}`; `TextProperties::{with_fill(CellFill), with_cell_anchor, with_cell_margin, with_underline, with_strike}` | adopted (0333); `with_fill` takes a `CellFill` (was an `Rgb`) |
| `litchi_pptx::table::{CellProperties, Cell::properties, Row::height, Table::row_cell_counts}` | adopted (0333) |
| `litchi_pptx::opened::Edit::{insert_table_row, remove_table_row, set_table_row_height}` | adopted (0333); column insertion and removal keep merged cells consistent where safe |
| `litchi_pptx::opened::Edit::set_text_property_chunks`; `source::Edit::set_text_property_chunks`; `SourceBackedSlideEdit::set_text_property_chunks` | adopted (0333), the text-formatting sibling of `set_shape_text_chunks` |
| `litchi_pptx::opened::Edit::{remove_comment_at, remove_modern_comment_at, remove_modern_comment_reply_at}` | adopted (0333), positional siblings of `set_comment_at` |
| `litchi_pptx::opened::Edit::{set_picture_image, set_shape_description, set_shape_name}` | adopted (0333) |
| `litchi_ops` PPTX kinds: `paragraph` and `run` under `tc` and `notes` (by path), `tr` (`height`, `cols`, `c1`…`c20`), `table` add (`rows`, `cols`, `data`), `tc` `valign`/`padding.*`, `picture` `alt`/`src`/frame set | adopted (0333) |

See [0333](unified-ops/0333-pptx-p0-residue.md).

## XLSB worksheet feature families (record 0322)

Seven scoped families publish XLSB worksheet features through immutable
part-level edits sharing one staged engine (`litchi_xlsb::part_edit`,
crate-private): worksheet streams are edited at the record level, so every
unedited record keeps its bytes. Each family starts from a `Package` or,
without copying, from a `Workbook`; `commit()` returns the triad
`package()`, `patch()`, `diagnostics()` with `into_package()`/
`into_parts()`, and an unchanged staging returns the source snapshot.
Positions are complete-catalog sheet positions.

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_xlsb::page_breaks::{Break, PageBreaks, read}`, `page_breaks::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`, `{Package, Workbook}::edit_page_breaks`, `Package::apply_page_breaks` | `Edit::{page_breaks, set_page_breaks}`; `BrtBeginRwBrk`/`BrtBeginColBrk` collections | ✅ 0322; family registered in the gate |
| `litchi_xlsb::auto_filter::{AutoFilter, read}`, `auto_filter::edit::{…}`, `{Package, Workbook}::edit_auto_filter`, `Package::apply_auto_filter` | `Edit::{auto_filter, set_auto_filter}`; a moved range keeps the filter columns still inside it | ✅ 0322; family registered in the gate |
| `litchi_xlsb::hyperlinks::edit::{…}`, `{Package, Workbook}::edit_hyperlinks`, `Package::apply_hyperlinks` | `Edit::{hyperlinks, add, set, remove}`; external targets live in worksheet relationships | ✅ 0322; family registered in the gate |
| `litchi_xlsb::data_validation::edit::{…}`, `{Package, Workbook}::edit_data_validations`, `Package::apply_data_validations` | `Edit::{validations, add, set, remove}`; classic `BrtDVal` only | ✅ 0322; family registered in the gate |
| `litchi_xlsb::conditional_formatting::edit::{…}`, `{Package, Workbook}::edit_conditional_formats`, `Package::apply_conditional_formats` | `Edit::{formats, add, set, remove, differential_formats, add_differential}`; `DifferentialStyle::{parse, encode}`, `Underline` | ✅ 0322; family registered in the gate |
| `litchi_xlsb::comments::edit::{…}`, `{Package, Workbook}::edit_comments`, `Package::apply_comments` | `Edit::{comments, add, set, remove}` with VML note shapes | ✅ 0322; family registered in the gate |
| `litchi_xlsb::drawing::edit::{…}`, `{Package, Workbook}::edit_drawings`, `Package::apply_drawings` | `Edit::{elements, add_picture, add_shape, set, remove}`; the XLSX drawing vocabulary with `Rgb = u32`; `ChartPart` (a frame's placement, name and chart markup) | ✅ 0322; family registered in the gate |
| `litchi_spreadsheet_drawing::objects::{Color, Element, ElementKind, Placement, Marker, Image, Fill, Line, Font, Align, VerticalAlign, PictureSpec, ShapeSpec, Change, ChartFrame, Detached, Limits}`, `objects::{read, chart_frames, drawing_of, count, next_id, append, check_picture, picture_fragment, shape_fragment, change, detach}` (each variable raw operation takes `&objects::Limits`); `litchi_spreadsheet_drawing::shape::{Limits, read(xml, &Limits)}`; `litchi_spreadsheet_drawing::vml::{Limits, Shape, shapes(xml, &Limits), next_shape_id, note_shape, new_drawing}` | the format-neutral half of the XLSX and XLSB drawing editors and their VML note shapes; the model is generic over the format's `Color` (XLSX `Rgb`, XLSB `u32`), and each `drawing::edit` re-exports it (`Fill`, `Line`, `Font`, `Element`, `ShapeSpec` are aliases); owner-derived policies retain the hard ceilings and typed limit failures across object, shape, geometry, image, and VML boundaries | ✅ 0325; raw policy/signature closure 0362ul51 |
| `litchi_xlsb::sheet::{Features, FeatureSet}`, `Workbook::{sheet_features, sheet_comments, drawing_elements, drawing_charts, differential_formats}`, `SourceBackedWorksheet::{features, comments, drawing_elements, drawing_charts}`, `SourceBackedWorkbook::differential_formats` | feature reads on both owners, without decoding cells; `FeatureSet` selects the kinds (only validations and formats render formulas) | ✅ 0322 |
| `litchi_xlsx::{Worksheet, SourceWorksheet}::threaded_comments`, `litchi_xlsx::{Workbook, SourceBackedWorkbook}::people` | threaded-comment and people reads on both owners (record 0320 gap 3, reading only) | ✅ 0322 |

## DOC styles, lists, hyperlinks, bookmarks, tabs and sections (record 0324)

The primary `doc::Edit` gains the record-0324 verbs; each stages one
MS-DOC structure (a paragraph's PAPX, the main story and its `PlcffldMom`
markers, `SttbfBkmk`/`PlcfBkf`/`PlcfBkl`, or one section's SEPX) and keeps
every other byte. Native indexes (`istd`, field begin CP) are the selectors
these verbs take; the unified layer resolves style ids and names.

| Name | Notes | Status |
|---|---|---|
| `Edit::set_paragraph_style(Position, u16)` | applies a paragraph style by `istd`; refuses a non-paragraph slot (`Refusal::StyleNotFound`) | ✅ 0324 |
| `body_text::CharacterValue::Style(u16)`, `CharacterValueProperty::Style` | `sprmCIstd` through `set_character_value`; index 10 clears it | ✅ 0324 |
| `Edit::{paragraph_tabs, set_paragraph_tabs}` | effective stops in, one `sprmPChgTabsPapx` out | ✅ 0324 |
| `body_text::{HyperlinkInsert, FieldPart}`, `TextTarget::Field { start, part }` | a field's instruction or result as a text target | ✅ 0324 |
| `Edit::{insert_hyperlink, remove_field}` | a whole `HYPERLINK` field in, a whole field (nested ones included) out | ✅ 0324 |
| `Edit::{bookmarks, add_bookmark, rename_bookmark, remove_bookmark}` | standard bookmarks by name | ✅ 0324 |
| `body_text::SectionValue`, `Edit::set_section_values(Position, &[SectionValue])` | typed `sprmS*` page setup | ✅ 0324 |
| `body_text::ChangeRef::{ParagraphStyle, ParagraphTabs, Field, Bookmark, Section}` | patch change descriptors; the durable envelope refuses them | ✅ 0324 |
| `body_text::Refusal::{StyleNotFound, InvalidProperty, FieldNotFound, BookmarkName, BookmarkNotFound, SectionNotFound}` | | ✅ 0324 |
| `litchi_ops` DOC kinds `Styles`, `Style`, `Numbering`, `AbstractNum`, `Num`, `Level`, `Hyperlink`, `Bookmark`, `Field`, `Toc`, `Section`, `Tab`, `TextBox`, `Shape` | the DOCX element names and paths (`/styles/<id or name>`, `/numbering/abstractNum[N]/level[L]`, `/body/p[N]/hyperlink[M]`, `/bookmark[@name=X]`, `/field[N]`, `/toc[N]`, `/section[N]`, `/body/p[N]/tab[M]`, `/body/textbox[N]`, `/body/shape[N]`); a DOC style's id is its name without spaces and punctuation, each word of a built-in name capitalized | ✅ 0324 |

## PPT element surface (record 0323)

Element verbs live on the primary `package::edit::Edit` and the root
`slide_order::Transaction` (no new transaction family). Each publishes
complete replacements of the persisted records it changes and keeps every
other record byte for byte; `slide_order::Refusal::{UnsupportedElement,
ElementNotFound, InvalidElementValue}` name the element refusals.

| Name | Notes | Status |
|---|---|---|
| `litchi_ppt::comments::{NewComment, CommentChange, DateTime}` | `DateTime::{parse, day_of_week}`; no ambient clock: an absent date is an all-zero `SYSTEMTIME` | ✅ 0323 |
| `Edit::{add_comment, set_comment, remove_comment}` (and on `slide_order::Transaction`) | `add_comment` returns the comment's `Position`; the author's next index raises a document-level author seed | ✅ 0323 |
| `slide_order::Refusal::{UnsupportedElement, ElementNotFound, InvalidElementValue}` | | ✅ 0323 |
| `Edit::set_table_cell_text(target, row, column, text)` (and on `slide_order::Transaction`) | a cell's shape text under the shape-text closure rules; an empty PowerPoint cell gains its first `TextCharsAtom` | ✅ 0323 |
| `shapes::shape_enum::TableShape::{cell_shape_id, row_heights, column_widths}` | the grid read from the table group's cell anchors, in master units | ✅ 0323 |
| `litchi_ppt::hyperlink::{Click, Target}`; `Click::resolve`; `Presentation::shape_click` | `Target::{Url, Slide(Position), Jump(InteractionJump)}`, named after the PPTX `hyperlinks::{Click, Target}` | ✅ 0323 |
| `Edit::set_shape_hyperlink(target, Option<&Click>)` (and on `slide_order::Transaction`) | a shape's `InteractiveInfo` click; URL and slide links append an `ExHyperlinkContainer` (and an `ExHyperlink9Container` screen tip) | ✅ 0323 |
| `text_edit::{TextRuns, CharacterRun, TextLink, RunFormat}`; `Slide::shape_text_runs`; `Presentation::resolve_click` | a body's character runs (`TextCFException` over UTF-16 units, covering the final paragraph mark) and text hyperlinks; outline-backed placeholders read their slide-list body | ✅ 0323 |
| `Edit::{set_run_format, set_text_hyperlink}` (and on `slide_order::Transaction`) | split the body's `StyleTextPropAtom` character runs (a body without one gains it) / write an `InteractiveInfo` + `TextInteractiveInfoAtom` pair; outline bodies are edited in the live document | ✅ 0323 |
| `litchi_ppt::theme::{Theme, Fonts, COLOR_SLOTS, parse, from_package}`; `Presentation::theme` | the `RoundTripThemeAtom` theme part's color and font schemes, read only | ✅ 0323 |

## XLS worksheet feature family (record 0327)

One scoped family edits the worksheet features of an opened BIFF8
workbook at the record level (`litchi_xls::sheet_features`): legacy notes,
hyperlinks, drawing pictures and shapes, conditional formats, data
validations, the auto-filter and manual page breaks. Every record a change
does not touch keeps its bytes; the `BoundSheet8`, `Index` and `ExtSST`
offsets that move with a splice are repaired, and the drawing group's
cluster table and blip store follow the shapes. `commit()` returns the
triad `package()`, `patch()`, `diagnostics()` with `into_package()`/
`into_parts()`; an unchanged edit returns the source package. Positions
are complete-catalog sheet positions.

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_xls::sheet_features::{Features, FeatureSet, Area, PageBreak, Note, Drawing, DrawingKind, Fill, Line, ConditionalFormat, Rule, RuleType, Style, Threshold, ColorScale, DataBar, IconSet, Validation, AutoFilter}`, `Package::sheet_features` | reads one worksheet without decoding its cells; `FeatureSet` selects the kinds (only formats and validations render formulas) | ✅ 0327 |
| `sheet_features::edit::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`, `Package::{edit_sheet_features, apply_sheet_features}` | `Edit::{features, set_row_breaks, set_column_breaks, add_hyperlink, set_hyperlink, remove_hyperlink, add_validation, set_validation, remove_validation, add_format, set_format, set_format_ranges, remove_format, set_auto_filter, add_note, set_note, remove_note, add_picture, add_shape, set_drawing, remove_drawing, commit}` | ✅ 0327; family registered in the gate |
| `sheet_features::edit::{Hyperlink, Validation, Style, Format, FormatRule, Threshold, Picture, Shape, ShapeGeometry, Anchor, DrawingChange}` | the authored values; formulas are BIFF8 token streams, colors palette indexes (styles) or RGB (drawings, CF12 visuals) | ✅ 0327 |
| `sheet_features::edit::compile_formula`, `Package::palette` | formula text to BIFF8 tokens with every built-in worksheet function (not `CALL`, `REGISTER.ID`, `EXEC`); the workbook palette a style's indexes name | ✅ 0327 |

## XLS raw record boundaries (record 0362ul46)

Variable-payload XLS fragment codecs require the selected source profile at
their public boundary. The shared-formula and hyperlink parsers retain their
historical decoded models and hard BIFF8 checks while refusing a payload over
`SourceLimits::max_entry_bytes()` before copying tokens or strings. The
worksheet/package readers and cell-value structural checks pass the retained
owner profile; the worksheet hyperlink collector applies the same check to
`HLinkTooltip` before UTF-16 decoding, and source-backed `reader_ref()` keeps
the retained physical raw profile when it materializes. The logical
`SourceBackedLimits` worksheet traversal budget remains on the source-backed
owner and is not converted into a per-entry CFB ceiling. There is no public
limit-less twin.

| Canonical name | Notes | Status |
| --- | --- | --- |
| `litchi_xls::formula_metadata::shared::parse(data, &SourceLimits)` | `ShrFmla` token payload; selected `entry_bytes` is checked before retention | ✅ 0362ul46; source migration |
| `litchi_xls::hyperlinks::parse_hlink_record(data, &SourceLimits)` | variable display/location/frame/moniker payload; selected `entry_bytes` is checked before string decoding | ✅ 0362ul46; source migration |
| collector-internal `hyperlinks::parse_tooltip(data, &SourceLimits)` | `HLinkTooltip` UTF-16 payload; selected `entry_bytes` is checked before decode | ✅ 0362ul46b/c; owner profile propagation |
| `litchi_xls::validation::{Limits, Error}` | focused validator names; the former `XlsValidationLimits` and `XlsValidationError` spellings are banned with no aliases | ✅ 0362ul46; all tracked consumers |

## ODS worksheet feature family (record 0327)

The worksheet features of an ODS sheet (`litchi_ods::advanced::sheet_features`)
are read from `content.xml` without decoding cells and edited through the
existing `document::Edit` transaction as provenance splices of `content.xml`
(and `styles.xml` for a conditional format's named cell style): every
element an edit does not touch keeps its bytes. No new transaction family:
the verbs extend `Edit`, and `commit()` keeps its triad.

| Canonical name | Notes | Status |
| --- | --- | --- |
| `sheet_features::{read, Features, Drawing, DrawingKind, Validation, AutoFilter, Rect, format_styles, FormatStyle}` | `calcext:conditional-format`, `table:content-validation` bound by `table:content-validation-name`, `table:database-range` with filter buttons, `fo:break-before` row/column styles, `draw:frame`/`draw:custom-shape` | ✅ 0327 |
| `sheet_features::{ValidationSpec, DrawingSpec, DrawingChange, parse_address, address, column_letters, to_odff, from_odff}` | authored values and ODF address / OpenFormula reference conversion | ✅ 0327 |
| `Edit::{add_sheet_format, set_sheet_format, remove_sheet_format, add_sheet_validation, set_sheet_validation, remove_sheet_validation, set_sheet_auto_filter, set_sheet_breaks, add_sheet_drawing, set_sheet_drawing, remove_sheet_drawing}` | sheet-scoped verbs of the document transaction; drawings by kind and position | ✅ 0327 |

## ODT, ODP and RTF element parts (record 0328)

`litchi-odt` gains an on-demand part index beside the body index; it keeps
offsets only and never changes the body index, so every body read and edit
is unaffected. The primary `Edit` gains the header and footer text set; tab
stops ride on the existing paragraph-property edit.

| Name | Notes | Status |
|---|---|---|
| `litchi_odt::body::{Parts, PartsView}` (`Parts::{build, view}`; `PartsView::{list_styles, list_style, regions, tables_of_contents, shapes, objects, columns}`) | one pass over each of `content.xml` and `styles.xml` | ✅ 0328 |
| `litchi_odt::body::parts::{ListStyle, ListLevel, ListLevelKind, Region, RegionKind, TableOfContents, Shape, EmbeddedObject, Column, TabStop, tab_stops, formula_text}` | borrowed views; attributes and text read from the source bytes | ✅ 0328 |
| `litchi_odt::body::fields`, `TextField::at` | every body text field (the `text_fields` set and the variable, reference, sequence, chapter, sender, template, database and conditional fields); `text_fields` keeps its set for field refresh | ✅ 0328 |
| `litchi_odt::body::TabStopEdit`, `ParagraphPropertyEdit::tab_stops` | replaces an automatic paragraph style's `style:tab-stops` | ✅ 0328 |
| `litchi_odt::transaction::Edit::set_region_text` | a master-page header or footer region's text as one paragraph; content a text set cannot keep is refused (`Unsupported`) | ✅ 0328 |
| `litchi_odt::Document::member_data` | the eager counterpart of `SourceBackedDocument::member_data` | ✅ 0328 |
| `litchi_ops` ODT kinds `Header`, `Footer`, `Numbering`, `AbstractNum`, `Num`, `Level`, `Field`, `Toc`, `Tab`, `Shape`, `Equation`, `Column` | the DOCX element names and paths (`/header[N]`, `/footer[N]`, `/numbering/abstractNum[N]/level[L]`, `/field[N]`, `/toc[N]`, `/body/p[N]/tab[M]`, `/body/shape[N]`, `/body/equation[N]`, `/body/tbl[N]/col[C]`), shared by OTH and ODM | ✅ 0328 |
| `litchi_odp::scene::parts::{TableColumn, table_columns, line_breaks, Plugin, MediaKind, plugin, Click, shape_click, link_title, inline_math, formula_text}` | read the bytes of one indexed scene node on demand; the scene index is unchanged | ✅ 0328 |
| `litchi_odp::scene::{Change::{SetRunLink, LinkRun, SetShapeClick, InsertLineBreak, RemoveLineBreak}, ShapeAction}` | a link run's target and title (or its removal, keeping the text), a plain run made a link, a shape's click listener or `draw:a` link, and line breaks, as checked scene splices | ✅ 0328 |
| `litchi_odp::{Presentation, SourceBackedPresentation}::member_data` | one package member (an embedded object's part), `None` when absent | ✅ 0328 |
| `litchi_ops` ODP kinds `Hyperlink`, `LineBreak`, `Col`, `Video`, `Audio`, `Equation`, `Series`, `Axis` | the PPTX paths of record 0321 (`/slide[N]/shape[M]/hyperlink`, `…/run[L]/hyperlink`, `…/paragraph[K]/linebreak[B]`, `/slide[N]/table[M]/col[C]`, `/slide[N]/video[M]`, `/slide[N]/audio[M]`, `/slide[N]/equation[K]`, `/slide[N]/chart[M]/series[K]`, `/slide[N]/chart[M]/axis[@role=ROLE]`); video, audio and equations are addressed among the media and embedded objects they are, and views list the structural kinds only | ✅ 0328 |
| `litchi_rtf::edit::Edit::{set_field_instruction, rename_bookmark, remove_bookmark, insert_bookmark, set_paragraph_tab_stops, set_section_values}` | mark edits: each splices only the `\fldinst` text, the `\bkmkstart`/`\bkmkend` groups, the paragraph's tab-stop controls or the section's page-setup controls, then reparses the candidate and refuses any other change; they publish as one `mark.transfer` root transfer, which replays durably; mixing them with body work is `Refused(UnsupportedOperation)` | ✅ 0328; 0359b |
| `litchi_rtf::edit::SectionValue` | one section page-setup value (break type, page size, orientation, margins, header and footer distance, columns, column spacing, title page, vertical alignment) | ✅ 0328 |
| `litchi_rtf::ParagraphFormat::tab_stops` | a paragraph's custom tab stops in declaration order | ✅ 0328 |
| `litchi_ops` RTF kinds `Section`, `Bookmark`, `Field`, `Toc`, `Tab`, `TextBox`, `Shape` | the DOCX paths (`/section[N]`, `/bookmark[N]`, `/bookmark[@name=X]`, `/field[N]`, `/toc[N]`, `/body/p[N]/tab[M]`, `/body/textbox[N]`, `/body/shape[N]`) over the native collections; sections, bookmark names, field instructions, TOC switches and tab stops are writable, tab stops are added and removed, bookmarks removed | ✅ 0328 |

## DOCX picture, comment and SmartArt edits; DOCX chart series and axes (record 0335)

The eager package-bound `document::Edit` gains the set and remove halves of
the DOCX picture, comment and SmartArt families whose add halves landed in
records 0251c, 0251d and 0295. Changes are `with_*` builders, as
`styles::catalog::StyleChange` is; each part left unset keeps the current
value. Every verb travels in the main-document `Patch` (graph changes and
element replacements), so replay and the exact inverse apply; no second
transaction family is introduced.

| Name | Notes | Status |
|---|---|---|
| `litchi_docx::document::PictureChange` (`new`, `with_size`, `with_name`, `with_description`, `with_image`, `is_empty`), `Edit::{set_picture_at, remove_picture_at}` | addressed as `insert_picture_at` returns them: a paragraph and the picture's position among `Paragraph::images`; a new image is a new media part, the previous one stays unreferenced (the removed-chart policy) | ✅ 0335 |
| `litchi_docx::document::CommentChange` (`new`, `with_author`, `with_initials`, `with_date`, `with_text`, `is_empty`), `Edit::set_comment` | by native comment identifier, as `remove_comment`; text replaced as `set_note_text` replaces a note's | ✅ 0335 |
| `litchi_docx::document::SmartArtParts` (`data`, `layout`, `quick_style`, `colors`, `drawing`), `Edit::{smartart_parts, set_smartart, remove_smartart}` | by the data-model relationship `insert_smartart_at` returns; the PPTX `opened::SmartArtParts` shape | ✅ 0335 |
| `litchi_docx::smartart::Diagram::extent` | the drawing's `wp:extent`, which a regenerated cached drawing takes | ✅ 0335 |
| DOCX `picture` set (`width`, `height`, `alt`, `name`, `src`; aliases `w`, `h`, `altText`, `description`, `path`), remove, query | OfficeCLI's picture keys; pictures are paragraph children for the query walker | ✅ 0335 |
| DOCX `comment` set (`author`, `initials`, `date`, `text`); `parentId` and `done` property gaps (follow-up 0335-F2) | a reply or resolution key is refused (`unsupported_capability`) on add and set | ✅ 0335 |
| DOCX `smartart` set (`layout`, `text`) and remove; `node` element (`/smartart[N]/node[K]`, nested; `level`, write-only `text`) | the PPTX `node` vocabulary of record 0284 | ✅ 0335 |
| DOCX `series` (`/body/p[N]/chart[M]/series[K]`, `/chart[N]/series[K]`; `name`, `values`, `categories`, `color`) and `axis` (`.../axis[@role=ROLE]`, `.../axis[N]`; `role`, `title`, `min`, `max`) | the PPTX and XLSX `chart-series` and `chart-axis` vocabulary | ✅ 0335 |

## XLSX P0 residue, root metadata and two correctness fixes (record 0338)

XLSX rich-text runs get a typed read and change surface over the existing
`RichText`/`RichRun` pair (record 0295); a picture's image is replaced in the
drawing edit; a chart series can refer to a range anywhere in the workbook.
The unified layer routes OfficeCLI's root metadata spelling (`set / title=…`)
to the metadata writers. No banned name is introduced.

| Name | Notes | Status |
|---|---|---|
| `litchi_xlsx::RichText::{new, run_mut, insert_run, remove_run}` | build runs, change one in place, insert or remove one | ✅ 0338 |
| `litchi_xlsx::RichRun::{new, set_text, font, apply_font}` | an unformatted run; its text; its `rPr` read as the stylesheet `Font`; a `FontChange` applied in place (unnamed properties keep their markup) | ✅ 0338 |
| `litchi_xlsx::SourceWorksheet::rich_text` | the source-backed counterpart of `Worksheet::rich_text` | ✅ 0338 |
| `litchi_xlsx::drawing::edit::Edit::replace_image` | a picture's image as a new media part, its relationship retargeted, the old part removed once unreferenced | ✅ 0338 |
| `litchi_drawingml::chart::authoring::Spec::with_series_reference` | per-series values (and categories) references, overriding the `Binding` for that series | ✅ 0338 |
| XLSX `run` element (`/<sheet>/<A1>/run[N]`, alias `r[N]`; `text`, `bold`, `italic`, `strike`, `underline`, `superscript`, `subscript`, `size`, `color`, `font`) | get, query, set, add, remove | ✅ 0338 |
| XLSX cell `type`, `runs`, `ref`; cell `add` under a sheet, row or cell | `type` converts the value (or the stored one); `runs` is a JSON array; `ref` (alias `address`) is add-only | ✅ 0338 |
| DOCX, XLSX and PPTX root `set` of document properties | routed to the `/metadata` writer; DOCX writes the core properties (`litchi_docx::metadata`) | ✅ 0338 |
| `litchi_ops::PathArg` `Display` | a fixed point of `PathArg::parse`: a relative first step that would read as sheet notation is quoted | behaviour (0338) |
| workspace `serde_json` `float_roundtrip` | JSON numbers decode correctly rounded | behaviour (0338) |

See [0338](unified-ops/0338-xlsx-p0-root-metadata-json-paths.md).

## DOCX and PPTX v1 closure (record 0342)

No new transaction family: the DOCX verbs live on the primary
`document::Edit` (both owners) and the standalone `raw::Tree`, beside
`format_paragraph_at`/`format_run_at`; the PPTX verb on the primary
`opened::Edit`, beside the record 0333 `set_table_*` verbs.

| Name | Status |
|---|---|
| `litchi_docx::document::Edit::{format_table_at, format_row_at}` | adopted (0342): splice `w:tblPr` (style, borders) and `w:trPr` (height, `cantSplit`, `tblHeader`, `hidden`); any other property is `Refusal::UnmodelledProperty` |
| `litchi_docx::properties::Height::new` | adopted (0342): the non-exhaustive row height's constructor, as `Border::new` |
| `litchi_docx::metadata::{Edit::app_props, Edit::set_app_props, Edit::custom_props, Edit::set_custom_props, Edit::remove_custom_props, Patch::parts, Error::Properties}` and `Package::app_props` | adopted (0342, 0362ac): the xlsx spellings of the same staging, with explicit custom-properties graph removal |
| `litchi_pptx::opened::Edit::set_table_properties` | adopted (0342): the `a:tblPr` style switches through a closure over `table::Properties`; the style id is refused |


## DOCX P0 final: cells, grid columns, hyperlink tips, comment threads (record 0345)

No new transaction family: the grid and tooltip verbs live on the primary
`document::Edit` (both owners) and the standalone `raw::Tree`, beside
`insert_row_at` and `set_hyperlink_target_at`; the comment verbs stay on the
eager package-bound `Edit` beside `add_comment_at`/`set_comment`. Native
comment identifiers remain the selectors these verbs take; `w14:paraId`
values are assigned by the leaf and never addressed.

| Name | Notes | Status |
|---|---|---|
| `litchi_docx::document::NewCell` (`new`, `with_text`, `with_width`, `text`, `width`), `Edit::insert_cell_at` | OfficeCLI's `AddCell`: one paragraph of text, `w:tcW`; the grid gains average-width columns and the other rows padding cells when the row outgrows it | ✅ 0345 |
| `litchi_docx::document::Edit::{insert_column_at, remove_column_at, set_column_width_at}` | a grid column by its position: `w:gridCol` and each row's cell at that slot, one table replacement; merged slots and the last column refused | ✅ 0345 |
| `litchi_docx::document::Edit::set_hyperlink_tooltip_at` | `w:tooltip` on the hyperlink start tag; `None` removes it | ✅ 0345 |
| `litchi_docx::document::NewComment::{with_parent, with_done, with_start_after, with_range}`, `CommentChange::with_done` | replies and resolution through `commentsExtended` (`w15:paraIdParent`, `w15:done`, keyed by the `w14:paraId` of the comment's last paragraph, [MS-DOCX] `CT_CommentEx`; an existing `commentsIds` part gains the new mappings); `with_start_after(Position)` is OfficeCLI's 1-based `runStart` minus one; `with_range(false)` writes a point comment | ✅ 0345 |
| `litchi_docx::comment::Comment::paragraph_id`, `Document::comments_extended`, `source_backed::Package::comments_extended` | the last paragraph's `w14:paraId` (the `commentsExtended` key); the parsed `commentsExtended` entries (`modern_comments::Extended`) on both owners | ✅ 0345 |
| DOCX `cell` add under a row (`text`, `width`; aliases `tc`) and `text` written by add and set | OfficeCLI's `table-cell` schema | ✅ 0345 |
| DOCX `column` element (`/body/tbl[N]/col[M]`, alias `col`): `width` read, set and added, `text` added; add and remove | OfficeCLI's virtual `table-column`; not a listed table child, so table views stay the rows | ✅ 0345 |
| DOCX hyperlink `tooltip` add and set | OfficeCLI's `hyperlink.tooltip` (add only there) | ✅ 0345 |
| DOCX comment `anchoredTo` (read), `done` (alias `resolved`; read, add, set), `parentId` (read, add), `runStart`, `range` (add) | OfficeCLI's `comment` schema; the record 0335 property gaps are deleted | ✅ 0345 |

See [0345](unified-ops/0345-docx-p0-final.md).

## Numbers swap, Python patches (record 0347)

No new transaction family. Numbers `swap` composes the existing sheet-order,
table-order, table-relocation and names transactions. The Python bindings
gain the ADR 0034 §3 patch vocabulary over `litchi_ops::transaction::Patch`.

| Name | Status |
|---|---|
| Python `Commit.patch` (None for `File.add_part` and `File.raw_set`) | adopted (0347) |
| Python `Patch.{apply, inverse, is_noop}` | adopted (0347, 0362m): the Rust `transaction::Patch` names; one Python `Patch` spans every Rust run of an `Edit`, and same-format reopened equal-content inputs pass the bounded replay proof |
| `litchi_numbers::table::relocation::Error::DuplicateTableName` | adopted (0347): a relocation into a sheet that already has a table of that name, refused before any rewrite |
| Numbers `swap` of `sheet`s and `table`s | adopted (0347): two moves applied one at a time (four leaf steps across sheets) |
| `litchi_xlsx::drawing::edit::Edit::replace_element_image` | adopted (0347): `replace_image` for a picture element the caller already read; `replace_image` delegates to it |
| `litchi_rtf::Error::UnsupportedCodePage` | adopted (0347): an `\ansicpg` with no exact codec (10001, 10007); `unsupported_capability` in the unified layer |

See [0347](unified-ops/0347-numbers-swap-python-patch-xlsx-picture.md).

## PPTX P0 closure: comments, modern comments, notes, placeholders, slides (record 0346)

No new transaction family: every verb lives on the primary `opened::Edit`,
beside the record 0252 and 0333 positional comment verbs. Comment and
reply identities stay native (the modern-comment `id` is read through the
raw query surface only).

| Name | Status |
|---|---|
| `litchi_pptx::opened::Edit::{set_comment_date_at, set_comment_author_at}` | adopted (0346), positional siblings of `set_comment_at`; re-attribution renames an author only when the comment is their only one, otherwise moves the comment to an author of that name and initials at the author's next index |
| `litchi_ops` PPTX `comment` `author`/`initials`/`date` set, `direction` (alias `dir`) on add | adopted (0346) |
| `litchi_pptx::opened::Edit::{add_modern_comment_text, add_modern_comment_reply_text, set_modern_comment_at}`, `opened::ModernCommentChange` | adopted (0346), positional siblings of `remove_modern_comment_at`; new identities are GUIDs derived from the edit's inputs, unique in the comment graph (no clock or random source) |
| `litchi_pptx::presentation::modern_comments::{Comment, Reply}::author_initials` | adopted (0346), beside `author_name` |
| `litchi_ops` PPTX `modernComment` add (`text`, `author`, `initials`, `created`, `resolved`, `parent`) and set; `reply` set; reads `initials`, `resolved`, `parent`; raw `id` | adopted (0346); `created` is required on add (no clock) |
| `litchi_pptx::opened::Edit::{add_placeholder, put_placeholder_text}` | adopted (0346): a placeholder bound to the layout slot of its type (a title adopts the layout's title-family type; other types the slot's index, else the smallest free one); `put_` fills the slide's matching placeholder or adds one, OfficeCLI's slide `title`/`text` shorthand |
| `litchi_pptx::opened::Edit::set_slide_layout` | adopted (0346): retargets the slide's layout relationship to the layout of that exact name; shapes are kept |
| `litchi_pptx::table::Table::cell_contents` | adopted (0346): every cell's properties and text body, by row; a table without namespace declarations resolves its scope once |
| `litchi_ops` PPTX `placeholder` add (`phType`, alias `type`; `text`), returning the canonical `/slide[N]/shape[K]`; `notes` add (`text`); slide `title`/`text` on add and `layout` set | adopted (0346) |

See [0346](unified-ops/0346-pptx-p0-final.md).

## Sheet traits: static dispatch (record 0353)

Owner decision 0350 (memo item 29): the `litchi-core::sheet` traits no
longer return trait objects. Each implementation names its concrete types
as generic associated types, and the iterators are standard `Iterator`s.
The traits are therefore not object-safe: generic code takes
`W: WorkbookTrait`, and a heterogeneous collection is an enum over the
concrete formats. The removed iterator traits are deleted, not aliased.

| Name | Status |
|---|---|
| `litchi_core::sheet::Worksheet::{Cell, Cells, Rows}`, `WorkbookTrait::{Worksheet, Worksheets}` (associated types) | adopted (0353), replacing `Box<dyn Cell>`, `Box<dyn Worksheet>` and the boxed iterators |
| `litchi_core::sheet::{CellIterator, RowIterator, WorksheetIterator}` (and their `litchi::sheet` re-exports) | deleted (0353): `Worksheet::cells`, `Worksheet::rows` and `WorkbookTrait::worksheets` return `Iterator`s |
| `Cell for &T`, `Cell for Cow<'_, T>`, `Worksheet for &T`, `Worksheet for Arc<T>` (blanket impls in `litchi_core::sheet`) | adopted (0353); the XLS `Cell for &Cell` and `Worksheet for &Worksheet` impls they supersede are removed |
| `litchi_ops::view::sheet::{Worksheet, Cell, Cells, Rows, Worksheets}` | adopted (0353): the associated types of `WorkbookView`, `Xlsx`, `Xlsb` and `Xls`, enums over the per-format types |
| `litchi_xls::worksheet::{Cells, Rows}`, `litchi_xls::workbook::Worksheets` | adopted (0353): the XLS reader's trait iterators (`Worksheet::cell` lends `Cow<'_, Cell>`, `WorkbookTrait` lends `&Worksheet`) |
| `litchi_xlsb::sheet::{StoredCells, Rows}`, `litchi_xlsb::workbook::Worksheets` | adopted (0353); `litchi_xlsb::workbook::WorksheetIterator` renamed `Worksheets`. The eager `Workbook` lends the cache's `Arc<Worksheet>` instead of a clone of the worksheet |

See [0353](unified-ops/0353-sheet-trait-static-dispatch.md).

## Text workbooks in their own leaf (record 0354)

Owner decision 0350 (memo item 26): the facade's `litchi::sheet::text`
module moves into the leaf crate `litchi-sheet-text`; the facade re-exports
it as `litchi::sheet::text` behind the `sheet-text` feature. Every reader
and writer stays; the names follow the focused-module rule. The
`sheet-text` and `facade-sheet-text` gate profiles ban every old name.

| Name | Status |
|---|---|
| `litchi_sheet_text::Workbook` (one sheet, immutable: `new(rows, options)`, `with_name`, `open`, `open_with_options`, `from_bytes`, `from_bytes_with_options`, `from_reader`, `from_reader_with_options`, `options`, `name`, `rows`, `bom`, `worksheet`, `write_to`, `to_bytes`) | adopted (0354); replaces `TextWorkbook` and its `from_path_with_config`, `worksheet_name`, `set_worksheet_name`, `set_data`, `detected_bom`, `config` (no `&mut self` method, record 0351 §2.1) |
| `litchi_sheet_text::{Worksheet, Cell, Worksheets, Cells, Rows}` | adopted (0354); `Worksheet<'a>` borrows the workbook's rows; the iterators are the statically dispatched associated types of the 0353 sheet traits; replace `TextWorksheet`, `TextCell`, `TextWorksheetIterator`, `TextCellIterator`, `TextRowIterator` |
| `litchi_sheet_text::delimited::{Options, Dialect, Tokenizer, declared_delimiter, read, write}` | adopted (0354); `Options` merges `TextConfig` and `DelimitedConfig`; `Tokenizer` replaces `TextParser` and import's private parser; `read`/`write` replace `read_delimited`/`write_delimited` |
| `litchi_sheet_text::{sylk, dif, fixed_width}::{Options, read, write}` | adopted (0354); replace `SylkConfig`/`DifConfig`/`FixedWidthConfig`, the `*Parser` types and `read_*`/`write_*` |
| `litchi_sheet_text::{Encoding, SourceLimits, SemanticLimits, LimitKind, SourceLimitKind, SemanticLimitKind, Error, Format, OptionKind}` | adopted (0354); every reader takes caller limits (`max_input_bytes`, `max_cells`); `litchi_ops::import::Encoding` is this `Encoding` |

## Markdown reader owner policy (record 0362ul70)

`litchi-markdown::reader::Snapshot` is an immutable exact-source owner, not a
detached raw codec. Its finite CommonMark defaults remain valid owner
convenience constructors; a selected grammar or stricter policy enters through
one retained `OpenOptions` value and is reused for edits, rendering, and patch
replay. The public limit-bearing `Snapshot::read_with` entry is removed; raw
record-limit signatures are not applicable because the parser is crate-private.

| Name | Status |
|---|---|
| `litchi_markdown::{OpenOptions, reader::OpenOptions}` | adopted (0362ul70; `dialect`/`read_limits` consuming setters and `*_value()` getters) |
| `litchi_markdown::reader::Snapshot::{read, read_bytes, read_with_options, read_bytes_with_options}` | adopted (0362ul70; default finite owner ingress plus canonical retained options) |
| `litchi_markdown::reader::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | adopted (0362ul70/165ud; reader family exposes typed limit dimensions, exact `Patch::is_noop`, the shared `Commit` alias, and owner replay result) |

Structured keywords (owner decision 0350, memo item 20):

| Name | Status |
|---|---|
| `litchi_core::Metadata::keywords: Vec<String>`; `litchi_core::metadata::keywords::{split, join, is_plain, SEPARATORS, JOINER}`, `Keywords::{list, from_list}`, `keywords::Error::Separator` | adopted (0354) |
| `litchi_ops::value::{ValueType::List, Scalar::List}`, `litchi_ops::node::Value::List`, `litchi_ops::props::PropValue::List` | adopted (0354): a list is a JSON array of strings in batches and views, a Python list of str; query comparison holds when one item does |
| `litchi_odf_common::metadata::{Assignments::set_keywords, Edit::set_keywords}`, `litchi_odt::metadata::Edit::set_keywords` | adopted (0354): one `meta:keyword` per keyword |
| `litchi_pages::metadata::Edit::set_keywords(&[String])`, `metadata::Error::InvalidKeyword`; `litchi_ods::metadata::Editor::set_keywords(Vec<String>)`; `litchi_odi` `set_keywords(Vec<String>)` and `MetadataFields::keywords() -> &[String]` | adopted (0354) |

See [0354](unified-ops/0354-sheet-text-and-structured-keywords.md).

## Sheet traits: typed errors and a synchronous engine (record 0353a)

Owner decision 0350 and coordinator decision 0353a: no boxed error on the
sheet traits and no boxed future in the formula engine.

| Name | Status |
|---|---|
| `litchi_core::sheet::{Error, LimitKind}` (non-exhaustive; `Refused(edit::Refusal)`, `Limit { kind, observed, maximum }`, `SourceChanged`, `Io`, `Malformed(&'static str)`, `Internal(&'static str)`), `Error::{refused, refusal}` | adopted (0353a): the format-neutral, content-free sheet and evaluation error (0351 §2.4 shape) |
| `litchi_core::sheet::Result<T, E = Error>` | adopted (0353a), replacing `Result<T, Box<dyn Error + Send + Sync>>` |
| `litchi_core::sheet::Worksheet::Error`, `WorkbookTrait::Error` (associated types; `WorkbookTrait::Worksheet<'a>: Worksheet<Error = Self::Error>`) | adopted (0353a): each implementation reports its own typed error (`litchi_ops::Error`, `litchi_xls::Error`, `litchi_xlsb::package::PackageError`, `litchi_sheet_text::Error`, `sheet::Error`) |
| `impl From<litchi_ops::Error \| litchi_xls::Error \| litchi_xlsb::package::PackageError \| litchi_sheet_text::Error> for litchi_core::sheet::Error` | adopted (0353a): classification with static reasons, which `FormulaEvaluator` requires (`sheet::Error: From<W::Error>`) |
| `litchi_eval::DepthLimit` | deleted (0353a): a too-deep chain is `sheet::Error::Limit { kind: LimitKind::Depth, .. }` |
| `litchi_eval::FormulaEvaluator::{evaluate_cell, evaluate_sheet}` | kept `async` (0353a); the engine core is synchronous and the futures only await the caller's `Fetch` for `WEBSERVICE` between passes |

See [0353a](unified-ops/0353a-sheet-errors-and-sync-engine.md).

## Numbers-wire raw boundaries (records 0362ul66, 0362ul66a)

The standalone `litchi-numbers-wire` leaf has one caller-selected raw policy:
`Limits` and its typed `LimitKind` error dimensions. `BncCell::{parse, encode}`
and `BncCellView::parse`, the legacy `PreBncCellView::parse`, and
`cell_value::decode_cell_value` all require `&Limits`. The shared table-cell
reader keeps its required aggregate read budget and also receives `&Limits`.
The borrowed BNC rewrite and clear operations use canonical names and the same
typed profile: `rewrite_duration_format_identifier`, `rewrite_scalar`,
`clear_value`, `clear_comment`, `rewrite_formula_cache`, `rewrite_formula`, and
`rewrite_formula_without_cache`.
The variable-size table writers (`tile_cell`, `tile_set_cell`, `header_adjust`,
`list_intern`, `list_release`, `cell_bytes`, and the formula-state readers)
likewise require the same policy. Input, output, field, opaque-tail and item
counts are checked before copies or staging allocations; vectors use fallible
reservation. No default policy is selected inside these boundaries, and no
`*_with_limits` twin or compatibility alias was introduced.

Owned BNC cells copy fixed fields and opaque tails only after the selected
profile admits their exact sizes. Re-encoding and all table rewrites retain
unknown protobuf fields and tail bytes byte-for-byte. Numbers, Pages and Keynote
consumers carry their owner-selected profile through every raw call; the
profile is derived from the package ingress and semantic budgets at the owner
boundary. The raw leaf's source tests cover exact and one-under input,
field, tail and output caps plus owned round-trip fidelity.

See [0362ul66](unified-ops/0362ul66-numbers-wire-raw.md) and the review return
[0362ul66a](unified-ops/0362ul66a-numbers-wire-raw-review.md).

## Unified `File` configuration and commit accessors (record 0362)

`litchi_ops::EditOptions` (re-exported as `litchi::EditOptions`) is the
per-call configuration of one `File` transaction: the `SemanticLimits` set
with the consuming `semantic_limits(limits)` and read with
`semantic_limits_value()`. Every configured `File` write takes it through a
`_with_options` form; the forms without options use `EditOptions::default()`.
`transaction::Commit<T>` and `BatchCommit` have no public fields: the triad
accessors are `file()` (the facade's spelling of `package()`),
`into_file()`, `patch()`, `diagnostics()`, `into_parts()` and the value
(`value()`/`into_value()` on both `Commit` and `BatchCommit`). Detection
functions take a required `&SourceLimits`; an owner derives that policy from
its `OpenOptions` before calling detection (their OOXML source limits bound
the OPC probe). The OOXML probes of `detect::ooxml` and the record parsers
take their limits as a required argument.

| Name | Status |
|---|---|
| `EditOptions` (`new`, `semantic_limits`, `semantic_limits_value`) | adopted (0362) |
| `File::{edit_with_options, batch_with_options, import_with_options, merge_with_options, refresh_with_options}` | adopted (0362); replace the `*_with_limits` forms (banned) |
| `transaction::Commit::{file, into_file, patch, diagnostics, value, into_value, into_parts}` | adopted (0362); public fields removed |
| `BatchCommit::{file, into_file, patch, diagnostics, value, into_value, into_parts}` | adopted (0362); public fields removed |
| `batch::Options::edit_options` | adopted (0362); replaces `limits` (banned) |
| `batch::from_bytes(bytes, &SourceLimits)` | adopted (0362); `from_bytes_with_limits` merged |
| `merge::Data::{from_json(source, &SemanticLimits), from_csv(source, &import::Options, &SemanticLimits)}` | adopted (0362); `from_json_with_limits`, `from_csv_with` merged (banned) |
| `detect::{detect_format_smart, detect_format_smart_with_tail, detect_and_prepare, detect_and_prepare_with_tail}` taking `&SourceLimits` | adopted (0362); the `*_with_limits` and `*_with_limits_and_tail` forms merged (banned) |
| `detect::ooxml::{detect_zip_format, try_detect_zip_format, detect_zip_format_from_reader, try_detect_zip_format_from_reader}` taking `&SourceLimits` | adopted (0362); the aliases `detect_ooxml_format`, `detect_ooxml_format_from_bytes` and their limit forms deleted (banned) |
| `format::docx::Docx::section_inventory(&section::Limits)` | adopted (0362); `section_inventory_with_limits` merged |
| `litchi_doc::{body_text, source_backed}::Refusal::kind` | adopted (0362) |

## Crypto raw codec limits (record 0362ul52)

The raw crypto codecs do not hide a process-wide capability behind a limit-less
entry point. Every parser, inspector, and serializer receives its module-owned
profile; package owners retain the profile selected at ingress.

| Name | Status |
|---|---|
| `litchi_crypto::ooxml::{open, inspect, load, encrypt, rekey}(..., &ooxml::Limits)` | adopted (0362ul52); duplicate `*_with` wrappers deleted, while explicit integrity-policy variants remain `*_with_policy` |
| `litchi_crypto::spaces::{inspect, inspect_bytes, parse_*, write_*}(..., &spaces::Limits)` | adopted (0362ul52); profile covers stream, aggregate, entry, component, string, XML-depth/node and output ceilings |
| `litchi_crypto::labels::{parse, write}(..., &labels::Limits)` | adopted (0362ul52); profile covers XML bytes, labels, extensions, depth/nodes, attributes and output |
| `spaces::Error::{InvalidLimit, LimitExceeded, Allocation}` and `labels::Error::{InvalidLimit, LimitExceeded, Allocation}` | adopted (0362ul52); typed limit dimensions preserve invalid-argument, limit-exceeded and allocation classification through OOXML, DOC and ops consumers |
| `spaces::Snapshot` / `spaces::Patch` profile retention and OOXML StrongEncryption propagation | adopted (0362ul52); inert IRM/license/label payloads remain source data |
| `spaces::Patch::write_to` and DataSpaces directory/source capture | adopted (0362ul52); directory traversal, aggregate payload admission, and final CFB publication are bounded before or during allocation |
| `litchi_doc::{OpenOptions, raw::Reader}::{data_spaces_limits, data_spaces_limits_value, data_spaces, data_spaces_snapshot}` | adopted (0362ul52); the DOC raw reader retains the ingress DataSpaces profile and passes it to both graph projections |

## Spreadsheet chart authoring placement (record 0362zab)

Record 0351 §2.1.7 treats a freshly authored chart as a builder draft, not
as a read-only package projection. The shared chart model is therefore
published under the focused `builder` module, and both spreadsheet hosts
expose the same nested `builder::chart` paths. The re-exports are the same
definitions; no wrapper or compatibility alias was added.

| Name | Status |
|---|---|
| `litchi_spreadsheet_drawing::chart::builder::{Chart, Series, Anchor, ExternalDataPart, ExternalDataTarget, Relationship, Target, UserShapesPart}` | adopted (0362zab §2.1.7); shared chart authoring values |
| `litchi_xlsx::builder::chart::{Chart, Series, Anchor, ExternalDataPart, ExternalDataTarget, Relationship, Target, UserShapesPart}` | adopted (0362zab §2.1.7); canonical XLSX authoring paths |
| `litchi_xlsb::builder::chart::{Chart, Series, Anchor, ExternalDataPart, ExternalDataTarget, Relationship, Target, UserShapesPart}` | adopted (0362zab §2.1.7); canonical XLSB authoring paths |
| `litchi_{xlsx,xlsb}::chart::{Chart, Series, Anchor, ExternalDataPart, ExternalDataTarget, Relationship, Target, UserShapesPart}` | deleted (0362zab §2.1.7); `chart::read` and `chart::write` remain codec entry points |
| `litchi_xlsx::drawing::Chart` and XLSB drawing/package projections | retained (0362zab §2.1.7); read-only projections remain distinct from authoring drafts |

The XLSX and XLSB glossary profiles ban every removed direct child of their
`chart` modules and point to the corresponding `builder::chart` path. The
builder paths are not listed as immutable owners: the authoring values retain
their existing construction and serialization behavior, while package and
projection owners retain their immutable API. All examples, tests, facade
consumers and performance harness imports were migrated in the same change.

See [0362zab](unified-ops/0362zab-spreadsheet-chart-builder.md).

## PPTX raw property and model3d limits (record 0362ul63)

PPTX view-properties and presentation-properties codecs expose one explicit
limits policy for every raw parse, tolerant/strict parse and XML write. Package
and source-backed owners derive that policy from their retained `SourceLimits`;
detached callers select the module-local `Limits` directly. Unknown XML and MCE
tolerance remain preserved, while input, output, depth, node, string, semantic
item and opaque-fragment budgets are checked before retention. Model3d reads,
writes, validation and edits retain the same owner-derived policy for XML,
relationships, links and embedded model/preview payloads.

| Name | Status |
|---|---|
| `litchi_pptx::view_properties::ViewProperties::{parse, parse_with_tolerance, parse_strict, to_xml}`; `view_properties::{load_from_package, load_from_package_with_tolerance}` | adopted (0362ul63); every raw codec and package loader takes the required `view_properties::Limits` |
| `litchi_pptx::presentation_properties::Properties::{parse, parse_with_tolerance, parse_strict, to_xml}`; `presentation_properties::{load_from_package, load_from_package_with_tolerance}` | adopted (0362ul63); strict/tolerant paths share the selected `presentation_properties::Limits` |
| `litchi_pptx::presentation_properties::read_only_recommended::Snapshot::{from_xml, from_xml_owned}` | adopted (0362ul63); detached source snapshots require the presentation-properties limits policy |
| `litchi_pptx::model3d::codec::{read, write}`; `model3d::validation::validate` | adopted (0362ul63); model3d package, selected-shape and edit consumers propagate the retained owner policy, including payload and relationship ceilings |

See [0362ul63](unified-ops/0362ul63-pptx-raw-properties.md).

## PPTX structural metadata policies (record 0362ul81)

The eight PresentationML structural metadata families share one required
`presentation_properties::metadata::Limits` policy.  Package owners derive
that policy from their retained `SourceLimits`; detached and raw operations
receive it explicitly.  Changed replay validates the candidate payload under
the recipient's policy, while an equal-content no-op keeps the recipient
profile and exact bytes without a second semantic rewrite.

| API | Status / contract |
|---|---|
| `litchi_pptx::presentation_properties::metadata::Limits` and its re-exports from `custom_show`, `guides`, `sections`, `structure`, `color_map`, `handout`, `template` and `protection` | adopted (0362ul81); one checked policy carries input/output bytes, tree/nodes, strings, extensions, family counts, references, attributes and relationships |
| `custom_show::List::{parse_xml, to_xml, to_xml_with_rel_map}` | adopted (0362ul81); required metadata limits, bounded MCE selection, relationship-aware output and no fixed codec ceilings |
| `guides::Guides::{from_xml, to_xml}`, `sections::{List, Section}::{from_xml, to_xml}` | adopted (0362ul81); required limits through DOM, extension, reference and writer paths, with opaque extension bytes retained |
| `structure::{load, load_snapshot}` and `structure::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | adopted (0362ul81); package-aware graph edits use the shared triad, derive owner policy, and replay against equal content rather than patch lineage |
| `color_map::{parse_master, parse_override}` and `color_map::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` | adopted (0362ul81); master/layout/slide sources retain their exact owner bytes and apply changed payloads under recipient limits |
| `handout::Master::{parse_xml, to_xml}`, `protection::Settings::{parse_xml, to_xml, to_pres_props_xml}`, `template::pres_props_xml_with_handout` | adopted (0362ul81); required limits cover input, output, attributes, strings, nodes and password verifier fields; detached handout/protection writers retain canonical typed-model semantics, while static template constants remain format-invariant |
| `custom_show`, `guides`, `structure` and `color_map` replay | adopted (0362ul81); no-op patches preserve exact recipient content and policy, changed patches perform bounded recipient validation, and no public `_with_limits`/default twins remain; `Snapshot::read`, projection aliases and draft `snapshot` wrappers (including `color_map::Draft::snapshot`) are banned in these scopes |

See [0362ul81](unified-ops/0362ul81-pptx-metadata-structural.md).

## PPTX metadata history limits (record 0362ul82)

The Changes Information, Revision Information, slide-synchronization,
caption-track and slide-show-event codecs now take their module-local typed
`Limits` at every raw parse, read, write and package-load boundary. Owners
retain the selected profiles from `OpenOptions`; edits accept them through
`EditOptions`, snapshots and patches retain them, projections use the same
profile, and replay admits changed payloads under the recipient profile.
Format constraints such as the slide-synchronization date lexical rules and
the inert treatment of extension/MCE payloads remain unchanged. No default
only twin, alias or `_with_limits` entry point was added.

| Name | Status |
|---|---|
| `litchi_pptx::presentation_properties::metadata::{changes,revision}::{Info::{parse,to_xml},load}` | adopted (0362ul82); explicit typed limits at raw XML and package boundaries |
| `litchi_pptx::presentation_properties::metadata::slide_sync::{Properties::{parse,to_xml},load}` | adopted (0362ul82); selected profile retains the existing date/year constraints |
| `litchi_pptx::presentation_properties::metadata::tracks::{File::{parse,to_bytes},load,load_media}` | adopted (0362ul82); WebVTT and slide metadata bounds share the retained profile |
| `litchi_pptx::presentation_properties::metadata::events::{load,load_snapshot}` | adopted (0362ul82); slide XML event admission remains inert and opaque outside the typed island |
| `OpenOptions::{changes_limits,revision_limits,slide_sync_limits,tracks_limits,events_limits}` and each family's `EditOptions::limits` | adopted (0362ul82); profiles survive owners, projections, edits, snapshots, patches and replay |

See [0362ul82](unified-ops/0362ul82-pptx-metadata-history.md).

## PPTX metadata history bounded admission (record 0362ul82a)

The 82a correction applies the selected semantic profile before MCE
expansion, keeps modern-comment Changes operations on the retained owner
profile, and performs checked, fallible event-fragment and slide-splice
allocation. Escaped slide-synchronization identifiers are counted before
output admission. These are corrections to the 82 contracts, not aliases or
additional default entry points.

| Name | Status |
|---|---|
| `litchi_pptx::modern_comments::{load_modern_comment_changes, Edit::{changes, update_changes}}` | adopted (0362ul82a); selected Changes limits are intersected with physical source limits and used for reads and writes |
| `litchi_pptx::presentation_properties::metadata::{changes,revision,slide_sync}::Limits` | adopted (0362ul82a); bounded MCE admission uses the selected profile before expansion |
| `litchi_pptx::presentation_properties::metadata::events::store` and slide extension splice | adopted (0362ul82a); checked exact sizing and fallible reservation use the selected output ceiling |

See [0362ul82a](unified-ops/0362ul82a-pptx-metadata-history-admission.md).

## Public support projections and owned record views (record 0362ul157)

The support leaves publish one spelling for each read projection. The PPT
record parser's owned traversal is explicitly named for its cloning behavior;
the borrowed traversal remains the zero-copy alternative. The iWork reference
graph accepts only validated `ObjectId` values, so the protobuf null sentinel
(`0`) is rejected at conversion rather than silently ignored by a raw-ID
forwarder. `ObjectIndex` exposes the complete immutable graph as a borrowed
typed snapshot; cloning that snapshot remains available without publishing
mutable state.

| Name | Status |
|---|---|
| `litchi_ppt::RecordParser::find_records_owned()` | adopted (0362ul157; corrected 0362ul157a); preserves the complete owned-clone traversal without a misleading unused filter argument, while `find_records_ref` remains the borrowed zero-copy view; the deprecated `find_records` spelling is deleted |
| `litchi_iwa_graph::ReferenceGraph::{add_object_reference, incoming, outgoing, all_object_ids, has_cycle, reachable, statistics}` | adopted (0362ul157); all public identities are `ObjectId`, `ObjectId::try_from(0)` is rejected, and the former raw-`u64` wrappers are deleted |
| `litchi_iwa_index::ObjectIndex::reference_snapshot` | adopted (0362ul157); borrows the complete immutable `ReferenceGraphSnapshot`, retaining its typed `Clone` capability; the deprecated `reference_graph` spelling is deleted |
