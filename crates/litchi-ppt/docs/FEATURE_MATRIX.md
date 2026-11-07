# PowerPoint Binary Feature Matrix (PPT)

This document tracks the public feature families implemented by the litchi-ppt crate for
legacy PowerPoint binary presentations. It is the format-specific split of the repository
feature matrix; it does not replace the repository-level matrix.

The matrix describes library support, not rendering fidelity or complete conformance with every
revision of the PowerPoint binary specification. A row can be supported while intentionally
treating external links, scripts, OLE payloads, media, and macros as inert data.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the feature scope described in the Notes cell |
| 🟡 | Bounded, partial, metadata-only, pass-through, or otherwise limited support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to the format or direction |

Read and Write describe the public direction independently. A 🟡 direction usually means a
subset of the model, lossless preservation, or an inert serializer rather than full semantic
CRUD. Cryptographic verification means integrity/signature verification only; it does not
establish certificate trust or revocation status.

## Implementation and specification scope

The implementation owner is litchi-ppt, exposed independently by the umbrella crate's ppt
feature. The audit covers the record parser, OLE package/editor, Escher and shape model, text
model, slide and notes model, animation and transition model, external-object storage, and
writers.

The specification audit used the read-only [MS-PPT] ToC.md and Front Matter.md, its overview
and structure families for file streams, document, slide, slide-show, shape, animation, text,
external-object, other, and common types, plus its structure examples. Companion audits
covered [MS-CFB], [MS-ODRAW], [MS-OLEDS], [MS-OGRAPH], [MS-OFFCRYPTO], and [MS-OVBA] for
compound files, OfficeArt, OLE activation metadata, native charts, encryption, and VBA.

## Additional native metadata coverage

These rows make previously grouped [MS-PPT] features explicit. Record codecs
and document-stream transactions do not imply an unrestricted opened-package
editor, playback, or native application behavior.

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Photo album settings (`PhotoAlbumInfo10Atom`) | 🟡 | ✅ | ❌ | `DocumentProperties10::photo_album` contains typed layout/frame shape and photo-album flags via `PhotoAlbumSettings`. No dedicated semantic photo-album writer; source preservation is distinct from authoring an album. |
| Presentation Advisor (`PP9DocBinaryTagExtension` preferences) | 🟡 | ✅ | 🟡 | `Presentation::presentation_advisor_settings()` exposes `PresentationAdvisorSettings`; its checked record parser and `to_record` retain the warning-rule settings. No advisor execution or ordinary package-level settings transaction. The current 11-file native binary PowerPoint corpus has no PresAdvisorFlags9Atom; its PP9 navigation tags are separate evidence. |
| Document sound collection (`SoundCollection`, `Sound`, `SoundData`) | 🟡 | ✅ | ✅ | Typed bounded collection inspection validates IDs, interaction references, and embedded-WAV references. The writer can add/replace/remove embedded sounds and build the collection; writer-supported references are checked during serialization. This does not fetch linked sounds or play media. `tests/ppt_sound_collection.rs` and `ppt_interaction_sound_writer.rs`. |
| WMF recoloring (`RecolorInfo`) | 🟡 | ✅ | 🟡 | `RecolorInfo::{parse_payload,to_payload,to_record}` models source colors, brushes, patterns, and bounded recolor entries. Codec support does not rasterize WMF or provide an ordinary selected-shape recolor transaction. |
| Custom table-style record (`RoundTripCustomTableStyles12Atom`) | 🟡 | ✅ | 🟡 | `DocumentStructure` models presence/placement; `Presentation::powerpoint12_document_properties()` validates and retains the embedded OPC table-style part. `Package::edit()` inserts/replaces/removes/moves the record envelope with source-checked patches. Native tests distinguish a populated 2393-byte package from empty 1986-byte lists, exercise the envelope lifecycle and valid native-donor replacement, and preserve unrelated records and CFB streams. The replacement API takes a record and does not validate its embedded OPC/DrawingML table-style package or supply an effective style resolver. |

## Package, records, and presentation structure

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open, create, and save | ✅ | ✅ | ✅ | OLE2 presentation streams, in-memory packages, path workflows, and transactional writer/editor APIs |
| CFB streams, storages, and property sets | ✅ | ✅ | ✅ | Compound-file graph access and package-preserving edits; unknown storage payloads remain inert |
| PowerPoint document stream and persist mapping | ✅ | ✅ | ✅ | Current-user, user-edit, persist-directory, persist-object, and document records with bounded validation; changed append-only publications use a UserEdit-linked `PersistPtrIncrementalBlock` (6002), while readers retain support for full (6001) and incremental directories. `document_structure::{Snapshot, Transaction, Commit, Patch}` provides source-checked terminal document metadata edits without rebuilding unrelated records. The package-level `slide_order` root composes exact-source shape text/anchor, slide visibility, advance timing, visual transitions, and inert external-media path/playback edits with slide-list changes, durable replay, inverse patches, merge planning, and history. Disjoint structural and media edits can be staged in either order: structural changes are checked and rebased over the current live-document owner before one append-only publication. |
| Slide list, ordering, IDs, and visibility | ✅ | ✅ | ✅ | Slide records, list-with-text containers, stable identifiers, writer-created slide graphs, and fixed-width `SlideShowSlideInfoAtom.fHidden` mutation through the durable package root. `SlidePersistAtom.fNonOutlineData` is exposed only as non-outline-content metadata and is preserved by visibility edits. |
| Main masters and master text styles | ✅ | ✅ | ✅ | Typed main-master records, placeholder drawing, master text-style levels, color schemes, PowerPoint 12 round-trip metadata, and bounded TemplateNameAtom design-name authoring |
| Title, notes, and handout master records | 🟡 | ✅ | 🟡 | Master contexts and references are parsed, including placeholder and color-scheme metadata; contextual SlideNameAtom authoring covers all four masters, and the notes-master `RT_RoundTripNotesMasterTextStyles12Atom` owner builds a bounded PresentationML `txStyles` package atomically; general OfficeArt master synthesis remains outside scope |
| Presentation metadata and summary information | ✅ | ✅ | ✅ | Document properties, summary information, document-summary information, bookmarks, privacy, print options, and related bounded records; bookmark-summary semantics are exposed through the layered `bookmark_summary::Summary` owner. Native PPT/PPS/POT properties retain Mac codepage, timestamp precision and saved counts; title/custom-property CRUD preserves all presentation and inert macro payload streams with exact inverse after reopen. |
| View state, guides, and print settings | ✅ | ✅ | ✅ | Normal/slide/notes view state, pane splitters, zoom metadata, guides, notes-view scale, handout targets, and display preferences are typed metadata and do not drive layout. View accessors resolve only the live document's direct DocInfoList, excluding superseded incremental records and rejecting duplicate live owners/views. Native tests cover zoom/origins, notes views, opaque pane bytes and preservation after reorder; these specimens have no nonempty guides or PrintOptionsAtom. |

## Slides, text, drawing, and themes

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Text boxes, paragraphs, runs, and formatting | ✅ | ✅ | ✅ | Text extraction and authoring with fonts, colors, paragraphs, runs, alignment, tabs, fills, and line/paragraph formatting. `text_edit::{Snapshot, Target, Transaction, Commit, Patch}` additionally provides source-checked reversible edits for an existing shape's single text atom. Length changes update record framing and are accepted for no-style or single-paragraph/single-character-run text; range-bearing dependencies are refused. Patches can use the shared bounded deterministic-JSON envelope with exact-artifact and expected-text preconditions. |
| Text extensions and outline metadata | ✅ | ✅ | ✅ | Language and spelling defaults, Kinsoku settings, text special information, master defaults, metachar placeholders, bookmarks, and OutlineTextRefAtom outline references. Native tests pin EN-US/alternate-zero/spelling-false defaults and exact detached record serialization. The four inspected native decks use inline text and have no outline references/bodies; their absence does not qualify outline binding or package language editing. |
| Picture bullets | ✅ | ✅ | ✅ | Bounded MS-PPT `BlipCollection9Container`/`BlipEntityAtom` parsing, OfficeArt BLIP/FBSE validation, and source-checked package snapshots/transactions for replacing, appending, and removing bullets; an existing collection is required and absent-owner creation is not exposed. All 11 native binary PowerPoint fixtures lack that owner, recorded as an evidence gap rather than positive CRUD coverage. Null text-style references and delayed FBSE stores remain inert, and rendering or native resource resolution is outside scope. |
| OfficeArt shapes and groups | ✅ | ✅ | ✅ | AutoShapes, groups, fills, gradients, lines, shadows, geometry, shape flags, and Escher records; native solver-container headers count their rules, and those children consume the shared record budget; parsed shape values carry source lineage and atomically refuse general text/fill/line mutation when no faithful existing-presentation transaction exists. The focused `text_edit` owner handles single-text-atom replacement, including bounded length changes when the complete formatting dependency closure is one paragraph and one character run; detached values remain editable. `client_anchor::{Snapshot, Transaction, Patch}` owns bounded MS-PPT host anchors with exact no-op retention and source-checked reversible edits, also surfaced by the durable package root for ordinary shapes and tables. Bounded transfer rewrites ordinary `OfficeArtSp`, connector, arc, callout, deleted-shape, drawing-selection, linked-shape, round-trip, and BLIP-bearing property references into target-owned IDs. Connection-site indexes and the flag bits surrounding deleted-shape IDs remain exact. References outside the selected drawing and animation/build graphs are typed refusals. |
| Shape-scoped OfficeArtClientData | 🟡 | ✅ | ✅ | `Package::edit_client_data` and `client_data::{Snapshot, Edit, Patch}` provide bounded checked child insertion, replacement, removal, and reversible source-validated edits. Known children retain their typed grammar while ordered producer-defined records remain opaque and inert. Native placeholder/round-trip/external-reference reads and opaque-child CRUD preserve all unowned drawing leaves, slide/Document records and CFB streams. Invalid candidate grammar produces an InvalidValue refusal; malformed source wire remains a corruption error. |
| Placeholders and inherited layout hints | ✅ | ✅ | ✅ | Placeholder kinds and contexts for slides, main masters, title masters, notes masters, and handout masters are validated; unresolved inheritance is not rendered |
| Pictures and BLIP resources | ✅ | ✅ | ✅ | JPEG/PNG and supported OfficeArt BLIP payloads with picture frames, image relationships, and writer support. Bounded slide transfer reuses semantically identical target images or atomically appends exact donor BLIP framing, a validated FBSE, and the target `Pictures` stream while extending or creating the document BStore; unrelated CFB streams remain exact and durable replay carries the same closure. |
| Tables | ✅ | ✅ | ✅ | OfficeArt table groups, grids, cells, rows, columns, cell text, dimensions, and table authoring. Native content-placeholder tables retain their typed table and cells; source title/cell edits reopen and invert exactly. |
| Legacy color schemes and theme-like round-trip metadata | 🟡 | ✅ | ✅ | Eight-color scheme atoms, per-slide and master scheme inventory, and bounded PowerPoint 12 theme records; this is not a full OOXML theme model |
| Native Graph and Excel charts | 🟡 | ✅ | 🟡 | Typed inert chart inventory over Graph and Excel-hosted OGraph books, persist mapping, frame/slide selectors, bounded decompression, and per-object failures; standalone Graph packages expose source-bound replacement and the writer authors the bounded Bar/Line/Pie Graph profile with matching PPT object closure, while Excel-hosted authoring and arbitrary semantic mutation remain outside the profile |
| Diagram and SmartArt objects | 🟡 | ✅ | 🟡 | A bounded native inventory groups the MS-PPT build identity and metadata with matching MS-ODRAW shape references; `diagram::{Snapshot, Transaction, Commit, Patch}` and `diagram::package::{SlideSnapshot, SlideEditor, SlideCommit, SlidePatch}` safely publish supported build mode and shape-reference edits while retaining OfficeArt bytes, but do not calculate layout, render, or author SmartArt |
| Diagram build records | 🟡 | ✅ | ✅ | Bounded `DiagramBuildContainer`/`DiagramBuildAtom` metadata follows [MS-PPT] §§2.8.13–2.8.14 and 2.13.7, retaining fixed-width unknown enum values and reserved bytes; source-checked transactions publish only fixed-width fields inside the owning slide envelope and do not render, author, or play SmartArt diagrams |

## Interaction, notes, review, and slide show

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Hyperlinks and slide navigation | ✅ | ✅ | ✅ | URL, internal slide, file, and named-show targets with relationship and frame metadata; targets are never opened or followed. Bounded slide transfer resolves each referenced donor hyperlink to semantically equivalent target metadata, rewrites fixed-width interaction IDs (including `OfficeArtClientData`/textbox hosts), and exposes the native ID map on `TransferPlan`; missing targets, macros, programs, and active OLE actions are refused. |
| Action and interaction settings | 🟡 | ✅ | 🟡 | Typed click/mouse-over actions, jumps, links, triggers, OLE verbs, sound references, flags, custom-show names, and inert macro names; text actions use bounded UTF-16 ranges and writer emission is limited to validated canonical records |
| Speaker notes | ✅ | ✅ | ✅ | Notes containers, notes text, notes drawings, notes IDs, and notes-page writer APIs |
| Classic comments | ✅ | ✅ | ✅ | Comment 2000 records, authors, anchors, text, indices, and package-level aggregation/authoring. Bounded transfer resolves each comment by author name against the target catalog and verifies the target's comment-index seed; unrelated catalog entries and colors need not be byte-identical. Missing authors or insufficient seeds are typed refusals. |
| Modern comments | ❌ | ❌ | ❌ | The binary format implementation covers legacy Comment 2000 records; PresentationML modern comment parts, replies, presence, and reactions are not a PPT feature surface |
| Animations and timing trees | ✅ | ✅ | ✅ | Build steps, paragraph/chart/diagram build metadata, triggers, motion paths, color/effect/motion/rotation/scale/set/command behaviors, conditions, sequences, and transactional animation editing |
| Transitions and slide advance timing | ✅ | ✅ | ✅ | Transition kind, speed, direction, sound references, click advance, timed advance, and slide timing records use the exact `[MS-PPT]` 2.6.6 effect table. Unsupported newer-format effects are rejected instead of assigned conflicting binary values. The durable package root independently replaces either `slideTime`/manual/automatic flags or the fixed-width visual effect type/direction/speed while retaining sound state, untouched timing/flags, unused bytes, and record framing exactly. Genuine producer-corpus files are fully reopened after replay/inverse, and a LibreOffice 26.2.5 same-format changed-save probe preserved the Litchi-authored Box/Out-to-Cover/FromLeft visual transition through native save and semantic readback. Missing, duplicate, noncanonical, invalid-direction, or out-of-range atoms are typed refusals; record insertion and playback remain unsupported. |
| Named/custom slide shows | ✅ | ✅ | ✅ | Named-show containers, names, ordered slide-ID lists, and CRUD. `Presentation::custom_shows()` returns `Result`, uses the strict bounded owner and live document, and resolves stable IDs against current order rather than subtracting 256. Native summary-slide insertion and append-only reorder are covered. |
| Headers and footers | ✅ | ✅ | ✅ | Presentation, slide, notes, and handout defaults, per-slide overrides, date formats, and typed metachar placeholder positions |

## Media and embedded content

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Audio and video external-object metadata | 🟡 | ✅ | 🟡 | Embedded and linked WAV/AIFF, MIDI, CD audio, AVI/movie, MCI, and external path metadata are validated; `external_media::{Snapshot, Transaction, Commit, Patch}` adds inert source-checked path/flag edits and media lifecycle operations. The package-level root publishes path/playback replacements through that owner, reopens the genuine PPT, and emits deterministic reversible operations. Bounded transfer matches inert media and embedded sounds by semantic content, rewrites their fixed-width target IDs, and reports non-identity mappings. Active OLE, unknown external-object children, executable actions, and missing equivalents remain typed refusals. Broader movie and linked-media authoring remains limited. |
| Media playback, rendering, and external activation | ❌ | N/A | N/A | Media bytes and targets are stored or exposed as inert data; no audio/video playback, browser launch, resource fetch, interpolation, or rendering is performed |
| Embedded OLE objects and ActiveX controls | ✅ | ✅ | ✅ | Embedded and linked OLE/control frames, ProgIDs, storages, compressed or uncompressed payloads, add/remove/reorder operations, and package-preserving edits; content-placeholder frames retain their picture/media/OLE identity; payloads are never activated |
| Custom XML data storage | ✅ | ✅ | ✅ | Bounded, lossless MsoDataStore item/property XML with item GUIDs, schema references, known-family classification, and IRM markers; schema URIs are never resolved |
| Embedded font records and payload metadata | 🟡 | ✅ | 🟡 | Font collections, font entities, embedding flags, and bounded embedded-font data are recognized; this is not a complete font extraction or licensing-aware authoring pipeline |

## Metadata extensions, security, and macros

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Smart tags | ✅ | ✅ | ✅ | PowerPoint 11 smart-tag stores, property bags, Unicode type/string tables, and shape-run mappings; recognizers, download URLs, and schemas remain inert. The current native binary PowerPoint corpus has no PP11 smart-tag store; ordinary programmable tags do not qualify this feature. |
| Programmable tags | ✅ | ✅ | ✅ | Document, slide, and shape ProgTags/ProgBinaryTag containers for PP9 through PP12, ordered binary-tag records, scope checks, duplicate rejection, byte-exact payload retention, and high-level accessors; payloads are never executed. Native document tags exercise PP9 navigation records, PP10 grid/font flags and PP12 document flags with exact codec and incremental-history retention. |
| Slide-library synchronization and round-trip records | ✅ | ✅ | ✅ | `slide_sync::{Snapshot, Synchronization, Editor}` exposes bounded typed server IDs, HTTP library URLs, and validated SYSTEMTIME values; atomic set/clear edits preserve unrelated and unknown slide records with reversible revisions, while no server is contacted |
| Broadcast, HTML-publish, routing-slip, envelope, and privacy metadata | 🟡 | ✅ | ✅ | Typed bounded records for presentation broadcast, web publishing, routing recipients, mail-envelope state, and privacy flags; `broadcast::{Snapshot, Transaction, Commit, Patch}` adds source-checked broadcast edits; no mail client, browser, network, or recipient workflow is invoked |
| Document comparison and review diff metadata | 🟡 | ✅ | 🟡 | Typed diff flags, reviewer names, slide/master/shape/text/table/external-object diff containers, and bounded record serialization; comparison generation and accept/reject workflow are not implemented |
| Modify-password and protection metadata | ✅ | ✅ | ✅ | Modify-password records and writer support; the protection policy is metadata and is not enforced by the library |
| Password encryption | ✅ | ✅ | ✅ | Supported legacy PowerPoint password profiles, encrypted summary handling, encrypted open/save, and password-change workflows; unsupported encryption kinds fail explicitly |
| VBA projects | ✅ | ✅ | ✅ | MS-PPT VBAInfo and VbaProjectStg persistence, bounded standalone-CFB and MS-OVBA project/module source parsing, deterministic cache-free authoring/removal, context-correct compressed storage, and encrypted round trips; VBA source is inert and never executed. Native sync-flushed outer storage is exposed as `Compression::ZlibSyncFlush`; bounded decoding requires exact output size and a verified flush boundary. It does not invent an original checksum |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral CFB signature verification and signing with transactional edits; certificate trust and revocation are outside the model |

## Important boundaries and explicit gaps

| Feature family exposed by the specifications | Status | Read | Write | Notes |
|-----------------------------------------------|--------|------|-------|-------|
| General title-, notes-, and handout-master authoring | 🟡 | ✅ | 🟡 | Contextual master inventory resolves main, title, notes, and handout records from the persist graph; master_layout Snapshot/Transaction provides bounded lossless record-tree edits, the metadata editor authors checked SlideNameAtom values, the nested master::metadata::template facade authors main-master TemplateNameAtom values, and notes_styles authors a bounded validated `txStyles` OPC package for notes masters with atomic validation and unknown-record preservation. Full append-only package rebuilding and semantic OfficeArt master synthesis remain outside this scope |
| Complete native chart authoring | 🟡 | ✅ | 🟡 | The neutral chart inventory and source-bound Graph replacement remain available; the writer now creates a bounded standalone Graph profile (one primary Bar, Line, or no-axis Pie group) with matching `ObjectSubtype::Graph`, `MSGraph.Chart`, persist/storage, and OfficeArt frame closure. Excel-hosted charts, secondary axes, mixed groups, rendering, evaluation, and arbitrary semantic mutation remain outside this profile |
| Native diagram or SmartArt authoring | ❌ | ❌ | ❌ | The read-only inventory is not a layout/rendering engine or full SmartArt authoring surface; MS-PPT diagram-build records remain animation metadata |
| Full OfficeArt record and vendor-extension semantics | 🟡 | 🟡 | 🟡 | The record graph handles the implemented record families and preserves bounded unknown data where possible; unsupported record semantics are not promoted to typed APIs |
| External resource resolution and OLE/macro/media execution | ❌ | N/A | N/A | URLs, paths, links, OLE verbs, macro names, scripts, and media remain inert even when their containing records are supported |
| Rendering and slide-show playback | ❌ | N/A | N/A | The crate serializes and exposes presentation data; it does not calculate layout, render shapes, play transitions, or run animation timelines |

## Native Microsoft PowerPoint regressions

`tests/microsoft_office_native.rs` covers all ten unencrypted PPT/PPS/POT decks at
slide/shape materialization scope, including zero dropped-shape repairs. Focused
checks cover grouped drawings, twelve native table cells, image slot counts and
exact preview hashes, media/OLE metadata, custom shows, source edits and inverses,
correct/wrong/missing passwords, and passive VBA module text. The no-encryption
build has a separate typed unsupported test. These tests do not execute macros,
refresh targets, render slides, or establish a fresh native-app reopen.

FBSE pixel-data identifiers and BLIP file-data identifiers remain distinct values;
image resolution follows slot offsets, kinds and lengths and retains both UIDs.
No equality between those two identifier fields is inferred.

## Native timing and master qualification

`tests/microsoft_office_native_metadata.rs` checks the saved native Fly In tree,
its eight nodes, 1.25-second x/y behaviors, shape target, keyframes, sequence
conditions and expanded paragraph build. The binary condition instance is 0–5,
as required by `TimeConditionContainer` and emitted by native PowerPoint; the
separate `ConditionEnum` table labels the same roles 1–6. Unicode timing strings
stop at their optional terminator. Authored strings with an embedded terminator
are refused because they cannot round-trip as a complete semantic string.

Native keyframes may store an empty optional formula beside an explicit value;
the codec retains that presence. It does not accept an empty formula as the
only value or as satisfying Formula calculation mode. Native expanded paragraph
builds may omit templates whose effects already live in the timing tree;
unexpanded empty builds and malformed level/node pairs remain errors.

The same target qualifies five main-master/persist identities, all eight master
text-style types, 17 embedded layout packages, Ion colors/typefaces, inherited
footer/date text, the handout master context and the legacy font-conversion
boundary. The legacy font specimen has three font entities and no EOT payload;
that absence is not font-extraction evidence. The PPTX companion separately
checks the exact embedded Cambria Math Regular payload.

Native duration, transition/advance and master design-name edits reopen to the
expected typed state, preserve unrelated slide/master records and CFB streams,
and invert exactly. Changed timing serialization can normalize optional string
terminators within the owned timing tree. The native Excel.Chart.8 inventory
checks its frame, series and all 24 category/value cache entries. Public operations
query the semantic title placeholder on file-backed Auto and byte Eager routes
while retaining its physical shape path through text/visibility edits.

These bounded assertions are indexed in `tools/native_office_tests.json`. They
do not establish full master inheritance/layout, arbitrary chart authoring,
rendering, calculation, slideshow playback or fresh PowerPoint acceptance.

`tests/microsoft_office_native_missing_owners.rs` audits all 11 native binary
presentations, including password-backed encrypted ingress. None has an
`MsoDataStore` or promotion-marker storage, a live PP10 reviewer tree, a reviewing
toolbar state, or a modify-password atom. It decodes the live document's
`BinaryTagData` payloads explicitly, inventories 12 PP10 records, and checks
`NotFound` plus exact no-op preservation for absent reviewer edits on the ten
unencrypted files. Open encryption is distinct from a modify password. These
are fixture gaps in `tools/native_office_evidence_reviews.json`; they do not
count as passing custom-XML, comparison or modify-password semantics.

The same audit target checks all 49 native slides for `RecolorInfoAtom`,
`DiagramBuild` and `DiagramBuildAtom`, entering both OfficeArt client-data
records and PP10 binary-tag payloads with exact record framing. The full ten
unencrypted document streams add historical/master/notes coverage. The encrypted
specimen contributes its decrypted slide and image inventory, not a claim about
every encrypted persisted record. Across all eleven files there are 26 PNG,
four JPEG and two EMF images, with no WMF. One paragraph build remains distinct
from diagram builds; six unencrypted slide catalogs contain no diagram builds,
and 42 slides refuse the absent build-list target. Recolor and diagram owners
still need native specimens; these two reviews do not qualify their semantics.

The runtime boundaries `ppt-375`, `ppt-393`
are explicitly reviewed in `tools/native_office_evidence_reviews.json`.
The operations test `microsoft_office_native_runtime_boundaries` checks all 139
Office files through Auto/file and Eager/byte ingress: HTML, SVG, screenshot and
PDF requests return the advertised refusal through string, JSON and sink APIs,
without writing to the sink. Locked files first require a password; supported
unlocks still refuse visual output, and encrypted XLSB retains its separate
unsupported-decryption boundary. Reads preserve publication bytes and all input
files. Cited owner tests qualify saved data and bounded edits, not the excluded
application rendering, layout or execution.

`microsoft_office_native_signature_boundary` authors disposable XML signatures
over seven native Office parents. These are Litchi derivatives, not native
Office signatures. Mathematical validity, resource coverage and signature
presence remain distinct from `Trust::NotChecked`. Workbook cases also check
outer/inner signed-handle edit refusals, explicit unsign, complete graph/data
equivalence to an ordinary scalar edit, and invalid-payload detection. The
shared verifier's older Microsoft-signed certificate fixtures explicitly retain
unchecked trust too. These tests do not claim native application acceptance,
legacy CryptoAPI signature authoring, trust-chain validation or revocation.
Native signed-fixture handoffs remain open despite this bounded operation
coverage; historical authoring statuses are unchanged.
