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

## Package, records, and presentation structure

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open, create, and save | ✅ | ✅ | ✅ | OLE2 presentation streams, in-memory packages, path workflows, and transactional writer/editor APIs |
| CFB streams, storages, and property sets | ✅ | ✅ | ✅ | Compound-file graph access and package-preserving edits; unknown storage payloads remain inert |
| PowerPoint document stream and persist mapping | ✅ | ✅ | ✅ | Current-user, user-edit, persist-directory, persist-object, and document records with bounded validation; changed append-only publications use a UserEdit-linked `PersistPtrIncrementalBlock` (6002), while readers retain support for full (6001) and incremental directories. `document_structure::{Snapshot, Transaction, Commit, Patch}` provides source-checked terminal document metadata edits without rebuilding unrelated records. The package-level `slide_order` root composes exact-source shape text/anchor, slide visibility, advance timing, visual transitions, and inert external-media path/playback edits with slide-list changes, durable replay, inverse patches, merge planning, and history. Disjoint structural and media edits can be staged in either order: structural changes are checked and rebased over the current live-document owner before one append-only publication. |
| Slide list, ordering, IDs, and visibility | ✅ | ✅ | ✅ | Slide records, list-with-text containers, stable identifiers, writer-created slide graphs, and fixed-width `SlideShowSlideInfoAtom.fHidden` mutation through the durable package root. `SlidePersistAtom.fNonOutlineData` is exposed only as non-outline-content metadata and is preserved by visibility edits. |
| Main masters and master text styles | ✅ | ✅ | ✅ | Typed main-master records, placeholder drawing, master text-style levels, color schemes, PowerPoint 12 round-trip metadata, and bounded TemplateNameAtom design-name authoring |
| Title, notes, and handout master records | 🟡 | ✅ | 🟡 | Master contexts and references are parsed, including placeholder and color-scheme metadata; contextual SlideNameAtom authoring covers all four masters, and the notes-master `RT_RoundTripNotesMasterTextStyles12Atom` owner builds a bounded PresentationML `txStyles` package atomically; general OfficeArt master synthesis remains outside scope |
| Presentation metadata and summary information | ✅ | ✅ | ✅ | Document properties, summary information, document-summary information, bookmarks, privacy, print options, and related bounded records; bookmark-summary semantics are exposed through the layered `bookmark_summary::Summary` owner |
| View state, guides, and print settings | ✅ | ✅ | ✅ | Normal/slide/notes view state, pane splitters, zoom metadata, guides, notes-view scale, handout targets, and display preferences are typed metadata and do not drive layout |

## Slides, text, drawing, and themes

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Text boxes, paragraphs, runs, and formatting | ✅ | ✅ | ✅ | Text extraction and authoring with fonts, colors, paragraphs, runs, alignment, tabs, fills, and line/paragraph formatting. `text_edit::{Snapshot, Target, Transaction, Commit, Patch}` additionally provides source-checked reversible edits for an existing shape's single text atom. Length changes update record framing and are accepted for no-style or single-paragraph/single-character-run text; range-bearing dependencies are refused. Patches can use the shared bounded deterministic-JSON envelope with exact-artifact and expected-text preconditions. |
| Text extensions and outline metadata | ✅ | ✅ | ✅ | Language and spelling defaults, Kinsoku settings, text special information, master defaults, metachar placeholders, bookmarks, and OutlineTextRefAtom outline references |
| OfficeArt shapes and groups | ✅ | ✅ | ✅ | AutoShapes, groups, fills, gradients, lines, shadows, geometry, shape flags, and Escher records; parsed shape values carry source lineage and atomically refuse general text/fill/line mutation when no faithful existing-presentation transaction exists. The focused `text_edit` owner handles single-text-atom replacement, including bounded length changes when the complete formatting dependency closure is one paragraph and one character run; detached values remain editable. `client_anchor::{Snapshot, Transaction, Patch}` owns bounded MS-PPT host anchors with exact no-op retention and source-checked reversible edits, also surfaced by the durable package root for ordinary shapes and tables. Bounded transfer rewrites ordinary `OfficeArtSp`, connector, arc, callout, deleted-shape, drawing-selection, linked-shape, round-trip, and BLIP-bearing property references into target-owned IDs. Connection-site indexes and the flag bits surrounding deleted-shape IDs remain exact. References outside the selected drawing and animation/build graphs are typed refusals. |
| Shape-scoped OfficeArtClientData | 🟡 | ✅ | ✅ | `client_data::{Snapshot, Transaction, Patch}` provides bounded checked child insertion, replacement, removal, and reversible source-validated edits; known children retain their typed grammar while ordered producer-defined records remain opaque and inert |
| Placeholders and inherited layout hints | ✅ | ✅ | ✅ | Placeholder kinds and contexts for slides, main masters, title masters, notes masters, and handout masters are validated; unresolved inheritance is not rendered |
| Pictures and BLIP resources | ✅ | ✅ | ✅ | JPEG/PNG and supported OfficeArt BLIP payloads with picture frames, image relationships, and writer support. Bounded slide transfer reuses semantically identical target images or atomically appends exact donor BLIP framing, a validated FBSE, and the target `Pictures` stream while extending or creating the document BStore; unrelated CFB streams remain exact and durable replay carries the same closure. |
| Tables | ✅ | ✅ | ✅ | OfficeArt table groups, grids, cells, rows, columns, cell text, dimensions, and table authoring |
| Legacy color schemes and theme-like round-trip metadata | 🟡 | ✅ | ✅ | Eight-color scheme atoms, per-slide and master scheme inventory, and bounded PowerPoint 12 theme records; this is not a full OOXML theme model |
| Native Graph and Excel charts | 🟡 | ✅ | 🟡 | Typed inert chart inventory over Graph and Excel-hosted OGraph books, persist mapping, frame/slide selectors, bounded decompression, and per-object failures; standalone Graph packages expose a transaction that replaces a Graph-framed chart substream and a typed host bridge that stages the result back into the matching `ExOleObjStg` while preserving its compression envelope and OfficeArt anchor, while PptWriter::add_chart still refuses fresh authoring because a complete Office-compatible binary chart grammar is not implemented |
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
| Named/custom slide shows | ✅ | ✅ | ✅ | Named-show containers, names, ordered slide-ID lists, and CRUD |
| Headers and footers | ✅ | ✅ | ✅ | Presentation, slide, notes, and handout defaults, per-slide overrides, date formats, and typed metachar placeholder positions |

## Media and embedded content

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Audio and video external-object metadata | 🟡 | ✅ | 🟡 | Embedded and linked WAV/AIFF, MIDI, CD audio, AVI/movie, MCI, and external path metadata are validated; `external_media::{Snapshot, Transaction, Commit, Patch}` adds inert source-checked path/flag edits and media lifecycle operations. The package-level root publishes path/playback replacements through that owner, reopens the genuine PPT, and emits deterministic reversible operations. Bounded transfer matches inert media and embedded sounds by semantic content, rewrites their fixed-width target IDs, and reports non-identity mappings. Active OLE, unknown external-object children, executable actions, and missing equivalents remain typed refusals. Broader movie and linked-media authoring remains limited. |
| Media playback, rendering, and external activation | ❌ | N/A | N/A | Media bytes and targets are stored or exposed as inert data; no audio/video playback, browser launch, resource fetch, interpolation, or rendering is performed |
| Embedded OLE objects and ActiveX controls | ✅ | ✅ | ✅ | Embedded and linked OLE/control frames, ProgIDs, storages, compressed or uncompressed payloads, add/remove/reorder operations, and package-preserving edits; payloads are never activated |
| Custom XML data storage | ✅ | ✅ | ✅ | Bounded, lossless MsoDataStore item/property XML with item GUIDs, schema references, known-family classification, and IRM markers; schema URIs are never resolved |
| Embedded font records and payload metadata | 🟡 | ✅ | 🟡 | Font collections, font entities, embedding flags, and bounded embedded-font data are recognized; this is not a complete font extraction or licensing-aware authoring pipeline |

## Metadata extensions, security, and macros

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Smart tags | ✅ | ✅ | ✅ | PowerPoint 11 smart-tag stores, property bags, Unicode type/string tables, and shape-run mappings; recognizers, download URLs, and schemas remain inert |
| Programmable tags | ✅ | ✅ | ✅ | Document, slide, and shape ProgTags/ProgBinaryTag containers for PP9 through PP12, ordered binary-tag records, scope checks, duplicate rejection, byte-exact payload retention, and high-level accessors; payloads are never executed |
| Slide-library synchronization and round-trip records | ✅ | ✅ | ✅ | `slide_sync::{Snapshot, Synchronization, Editor}` exposes bounded typed server IDs, HTTP library URLs, and validated SYSTEMTIME values; atomic set/clear edits preserve unrelated and unknown slide records with reversible revisions, while no server is contacted |
| Broadcast, HTML-publish, routing-slip, envelope, and privacy metadata | 🟡 | ✅ | ✅ | Typed bounded records for presentation broadcast, web publishing, routing recipients, mail-envelope state, and privacy flags; `broadcast::{Snapshot, Transaction, Commit, Patch}` adds source-checked broadcast edits; no mail client, browser, network, or recipient workflow is invoked |
| Document comparison and review diff metadata | 🟡 | ✅ | 🟡 | Typed diff flags, reviewer names, slide/master/shape/text/table/external-object diff containers, and bounded record serialization; comparison generation and accept/reject workflow are not implemented |
| Modify-password and protection metadata | ✅ | ✅ | ✅ | Modify-password records and writer support; the protection policy is metadata and is not enforced by the library |
| Password encryption | ✅ | ✅ | ✅ | Supported legacy PowerPoint password profiles, encrypted summary handling, encrypted open/save, and password-change workflows; unsupported encryption kinds fail explicitly |
| VBA projects | ✅ | ✅ | ✅ | MS-PPT VBAInfo and VbaProjectStg persistence, bounded standalone-CFB and MS-OVBA project/module source parsing, deterministic cache-free authoring/removal, context-correct compressed storage, and encrypted round trips; VBA source is inert and never executed |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral CFB signature verification and signing with transactional edits; certificate trust and revocation are outside the model |

## Important boundaries and explicit gaps

| Feature family exposed by the specifications | Status | Read | Write | Notes |
|-----------------------------------------------|--------|------|-------|-------|
| General title-, notes-, and handout-master authoring | 🟡 | ✅ | 🟡 | Contextual master inventory resolves main, title, notes, and handout records from the persist graph; master_layout Snapshot/Transaction provides bounded lossless record-tree edits, the metadata editor authors checked SlideNameAtom values, the nested master::metadata::template facade authors main-master TemplateNameAtom values, and notes_styles authors a bounded validated `txStyles` OPC package for notes masters with atomic validation and unknown-record preservation. Full append-only package rebuilding and semantic OfficeArt master synthesis remain outside this scope |
| Complete native chart authoring | 🟡 | ✅ | ❌ | Graph and Excel chart payloads can be inspected through the neutral chart inventory; bounded standalone Graph package replacement and host-side `Graph::replace_package` staging are available, but they do not claim fresh chart creation or arbitrary semantic mutation |
| Native diagram or SmartArt authoring | ❌ | ❌ | ❌ | The read-only inventory is not a layout/rendering engine or full SmartArt authoring surface; MS-PPT diagram-build records remain animation metadata |
| Full OfficeArt record and vendor-extension semantics | 🟡 | 🟡 | 🟡 | The record graph handles the implemented record families and preserves bounded unknown data where possible; unsupported record semantics are not promoted to typed APIs |
| External resource resolution and OLE/macro/media execution | ❌ | N/A | N/A | URLs, paths, links, OLE verbs, macro names, scripts, and media remain inert even when their containing records are supported |
| Rendering and slide-show playback | ❌ | N/A | N/A | The crate serializes and exposes presentation data; it does not calculate layout, render shapes, play transitions, or run animation timelines |
