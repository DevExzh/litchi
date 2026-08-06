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
| PowerPoint document stream and persist mapping | ✅ | ✅ | ✅ | Current-user, user-edit, persist-directory, persist-object, and document records with bounded validation |
| Slide list, ordering, IDs, and visibility | ✅ | ✅ | ✅ | Slide records, list-with-text containers, stable identifiers, hidden state, and writer-created slide graphs |
| Main masters and master text styles | ✅ | ✅ | ✅ | Typed main-master records, placeholder drawing, master text-style levels, color schemes, PowerPoint 12 round-trip metadata, and bounded TemplateNameAtom design-name authoring |
| Title, notes, and handout master records | 🟡 | ✅ | 🟡 | Master contexts and references are parsed, including placeholder and color-scheme metadata; contextual SlideNameAtom authoring covers all four masters, and the notes-master `RT_RoundTripNotesMasterTextStyles12Atom` owner builds a bounded PresentationML `txStyles` package atomically; general OfficeArt master synthesis remains outside scope |
| Presentation metadata and summary information | ✅ | ✅ | ✅ | Document properties, summary information, document-summary information, bookmarks, privacy, print options, and related bounded records; bookmark-summary semantics are exposed through the layered `bookmark_summary::Summary` owner |
| View state, guides, and print settings | ✅ | ✅ | ✅ | Normal/slide/notes view state, pane splitters, zoom metadata, guides, notes-view scale, handout targets, and display preferences are typed metadata and do not drive layout |

## Slides, text, drawing, and themes

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Text boxes, paragraphs, runs, and formatting | ✅ | ✅ | ✅ | Text extraction and authoring with fonts, colors, paragraphs, runs, alignment, tabs, fills, and line/paragraph formatting |
| Text extensions and outline metadata | ✅ | ✅ | ✅ | Language and spelling defaults, Kinsoku settings, text special information, master defaults, metachar placeholders, bookmarks, and OutlineTextRefAtom outline references |
| OfficeArt shapes and groups | ✅ | ✅ | ✅ | AutoShapes, groups, anchors, fills, gradients, lines, shadows, geometry, shape flags, and Escher records |
| Placeholders and inherited layout hints | ✅ | ✅ | ✅ | Placeholder kinds and contexts for slides, main masters, title masters, notes masters, and handout masters are validated; unresolved inheritance is not rendered |
| Pictures and BLIP resources | ✅ | ✅ | ✅ | JPEG/PNG and supported OfficeArt BLIP payloads with picture frames, image relationships, and writer support |
| Tables | ✅ | ✅ | ✅ | OfficeArt table groups, grids, cells, rows, columns, cell text, dimensions, and table authoring |
| Legacy color schemes and theme-like round-trip metadata | 🟡 | ✅ | ✅ | Eight-color scheme atoms, per-slide and master scheme inventory, and bounded PowerPoint 12 theme records; this is not a full OOXML theme model |
| Native Graph and Excel charts | 🟡 | ✅ | 🟡 | Typed inert chart inventory over Graph and Excel-hosted OGraph books, persist mapping, frame/slide selectors, bounded decompression, and per-object failures; standalone Graph packages expose a transaction that replaces a Graph-framed chart substream and a typed host bridge that stages the result back into the matching `ExOleObjStg` while preserving its compression envelope and OfficeArt anchor, while PptWriter::add_chart still refuses fresh authoring because a complete Office-compatible binary chart grammar is not implemented |
| Diagram and SmartArt objects | 🟡 | ✅ | ❌ | A bounded native inventory groups the MS-PPT build identity and metadata with matching MS-ODRAW shape references and inert OfficeArt payload handles; it does not calculate layout, render, or author SmartArt |
| Diagram build records | 🟡 | ✅ | ✅ | Bounded `DiagramBuildContainer`/`DiagramBuildAtom` metadata follows [MS-PPT] §§2.8.13–2.8.14 and 2.13.7, retaining fixed-width unknown enum values and reserved bytes; the diagram inventory reuses that owner and does not render, author, or play SmartArt diagrams |

## Interaction, notes, review, and slide show

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Hyperlinks and slide navigation | ✅ | ✅ | ✅ | URL, internal slide, file, and named-show targets with relationship and frame metadata; targets are never opened or followed |
| Action and interaction settings | 🟡 | ✅ | 🟡 | Typed click/mouse-over actions, jumps, links, triggers, OLE verbs, sound references, flags, custom-show names, and inert macro names; text actions use bounded UTF-16 ranges and writer emission is limited to validated canonical records |
| Speaker notes | ✅ | ✅ | ✅ | Notes containers, notes text, notes drawings, notes IDs, and notes-page writer APIs |
| Classic comments | ✅ | ✅ | ✅ | Comment 2000 records, authors, anchors, text, indices, and package-level aggregation/authoring |
| Modern comments | ❌ | ❌ | ❌ | The binary format implementation covers legacy Comment 2000 records; PresentationML modern comment parts, replies, presence, and reactions are not a PPT feature surface |
| Animations and timing trees | ✅ | ✅ | ✅ | Build steps, paragraph/chart/diagram build metadata, triggers, motion paths, color/effect/motion/rotation/scale/set/command behaviors, conditions, sequences, and transactional animation editing |
| Transitions and slide advance timing | ✅ | ✅ | ✅ | Transition kind, speed, direction, sound references, click advance, timed advance, and slide timing records; playback is never performed |
| Named/custom slide shows | ✅ | ✅ | ✅ | Named-show containers, names, ordered slide-ID lists, and CRUD |
| Headers and footers | ✅ | ✅ | ✅ | Presentation, slide, notes, and handout defaults, per-slide overrides, date formats, and typed metachar placeholder positions |

## Media and embedded content

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Audio and video external-object metadata | 🟡 | ✅ | 🟡 | Embedded and linked WAV/AIFF, MIDI, CD audio, AVI/movie, MCI, and external path metadata are validated; sound collections have bounded authoring, while broader movie and linked-media authoring remains limited |
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
| Broadcast, HTML-publish, routing-slip, envelope, and privacy metadata | 🟡 | ✅ | ✅ | Typed bounded records for presentation broadcast, web publishing, routing recipients, mail-envelope state, and privacy flags; no mail client, browser, network, or recipient workflow is invoked |
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
