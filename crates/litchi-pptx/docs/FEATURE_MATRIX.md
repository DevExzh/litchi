# PresentationML Feature Matrix (PPTX)

This document tracks the public feature families implemented by the litchi-pptx crate for
PresentationML packages, including the package graphs used by macro-enabled and slide-show
variants where the implementation exposes them. It is the format-specific split of the
repository feature matrix; it does not replace the repository-level matrix.

The matrix describes library support, not rendering fidelity or complete conformance with every
revision of ISO/IEC 29500, the Microsoft extensions, or every producer extension. A row can be
supported while intentionally treating external links, scripts, embedded payloads, web
extensions, comments, and media controls as inert data.

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

The implementation owner is litchi-pptx. Its public package facade owns bounded OPC parts,
relationships, strict/transitional XML, semantic slide owners, and transactional graph editing.
Raw or unknown XML retention is not counted as typed semantic support unless a row says so.

The specification audit used the read-only [MS-PPTX] ToC.md and Front Matter.md, its overview
and applicability statement, the core extension structure families, versioned schemas, and the
transition, media, sections, and slide-show examples. Companion audits covered [MS-ODRAWXML]
for DrawingML, charts, diagrams, ink, math, and 3D extensions, [MS-OWEXML] for web extensions,
[MS-CFB] for embedded compound files, [MS-OFFCRYPTO] for encryption and signatures,
[MS-OVBA] for VBA, and [MS-OI29500]/[MS-OE376] for the surrounding Office Open XML conventions.

## Package, relationships, and presentation graph

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open, create, and save | ✅ | ✅ | ✅ | Path, bytes, reader, in-memory OPC, and deterministic package workflows with bounded XML and part validation |
| OPC parts, relationships, and content types | ✅ | ✅ | ✅ | Strict/transitional relationship and content-type handling, graph validation, transactional edits, and relationship-aware cleanup |
| Slides, IDs, ordering, duplication, and visibility | ✅ | ✅ | ✅ | Enumerate, add, insert, duplicate, move, remove, resolve stable slide identities, and expose hidden state |
| Opened-presentation atomic composition | ✅ | ✅ | ✅ | `opened::{Snapshot, Transaction, Commit, Patch}` composes slide ordering/removal; shape text and common text-box, rectangle, ellipse, picture creation/removal; dependency-complete common-shape transfer; notes; table frames/styles; charts; media; masters/layouts; legacy and modern comments/replies/typed extensions; and general relationship-closure transfer in one immutable-root transaction. Exact before/after OPC resource patches cover part and relationship creation/removal, validate a complete candidate before atomic publication, and provide durable inverse, deterministic compatible joins, non-mutating three-way conflict plans, and bounded undo/redo history while preserving untouched parts and unknown XML. Group subtrees and externally attached connectors compute a complete top-level endpoint closure, then remap every common non-visual identity, connector endpoint, relationship-namespace attribute, and copied dependency part. Recognized table, classic-chart, DiagramML, and inert OLE graphic frames use that closure; DiagramML `r:dm/lo/qs/cs` targets are covered explicitly. `ShapeTransferRefusal` classifies nested selections, identity-less `contentPart` shapes, unknown extension shapes, unclassified graphic-frame payloads, missing identities, and unresolved connector endpoints before atomic publication. |
| Source-backed cross-slide image copy | 🟡 | ✅ | 🟡 | `SourceBackedPresentationEditor::{plan_cross_slide_copy,publish_cross_slide_copy_to_stream}` extends the dependency-free source-backed slide closure to a nonempty caller-bounded set of direct `p:pic` leaves under exactly one direct `p:spTree`; each selected picture has exactly one direct `p:blipFill/a:blip r:embed` reference to an internal, relationship-free `/ppt/media/` `image/*` leaf. Distinct media targets are copied once, deterministic destination media URIs are allocated, and selected image relationship IDs are allocated/rewritten without XML normalization. Semantic picture parsing preserves bounded foreign, non-MCE, non-relationship `a:blip` attributes opaquely and accepts one valid unqualified `cstate` token; namespace-safe copy rewrites only the full-slide resolved relationship-namespace `r:embed`. A full-slide unbound lexical `r:embed` returns `UnsupportedRelationship`; `r:link`, unknown relationship attributes, MCE, and duplicate/ambiguous resolved embeds refuse. The destination anchor permits and preserves other valid existing relationships while anchoring exactly one dialect-correct internal `slideLayout`; the full-slide namespace-aware `SourceSlide::images` inventory is a fence, source catalog relationship reconstruction is fallible, physical ZIP media deduplication is asserted, and strict XML end-name and unresolved-prefix fences remain active. Unselected XML and package members remain preserved while source/destination dialect, layout/master/theme equivalence, limits, freshness, lineage, signatures, destination physical members, publication cancellation, partial-sink, and resource fences remain checked. Selected image relationship-ID collisions are handled by allocation/rewrite; layout, non-selected, and unsupported relationship-ID collisions still refuse. Wrong-type non-selected and non-anchor slide bindings are rejected at open, and planning revalidates every binding defense-in-depth; no malformed-object planner test is claimed. Missing, mistyped, external/linked, outbound, unreferenced, shared, or otherwise unsupported media relationships, duplicate shape trees, malformed/misplaced/ambiguous blips, and broader dependency closures fail closed before output. Image decoding, conversion, rendering, and durable inverse are not provided. |
| Presentation size, root properties, and view settings | 🟡 | ✅ | ✅ | Typed slide/notes dimensions, size type, presentation and view settings, PowerPoint 2010 browse-mode state, and bounded extension serialization; settings do not render or drive a host UI |
| Sections | ✅ | ✅ | ✅ | Typed section IDs and names, resolved slide-index membership, and graph-safe ordered CRUD |
| Custom slide shows | ✅ | ✅ | ✅ | Typed named subsets and graph-safe ordered CRUD; `presentation_properties::metadata::custom_show::{Snapshot, Transaction, Commit, Patch}` adds source-checked XML inverse edits |
| Slide masters and layouts | ✅ | ✅ | ✅ | Semantic master/layout reading with shape and placeholder inventory, typed layout references, matching/type metadata, relationship validation, new-master and layout authoring, placeholder add/replace, and unreferenced-layout removal |
| Handout master | ✅ | ✅ | ✅ | Root relationship resolution, handout settings, and header/footer metadata |
| Speaker notes and notes masters | ✅ | ✅ | ✅ | `Package::{notes, put_notes, remove_notes, clear_notes}` uses semantic slide selectors, while `notes::{Snapshot, Transaction, Commit, Patch}` provides bounded source-checked notes-slide/master/theme/text lifecycle edits; opaque XML and exclusively owned descendants are handled losslessly, slide comments remain slide-owned, and notes are never rendered or activated |
| Themes, color maps, and overrides | ✅ | ✅ | ✅ | Master-, layout-, and slide-scoped theme resolution, typed color maps and overrides, 12-slot color schemes, major/minor font schemes, attachment, replacement, removal, and orphan cleanup; contextual `color_map::{Snapshot, Transaction, Commit, Patch}` adds source-checked map edits, while fmtScheme authoring is not covered |
| Presentation property structure and extended guides | ✅ | ✅ | ✅ | `presentation_properties::metadata::structure::{Snapshot, Transaction, Commit, Patch}` and `guides::{Snapshot, Transaction, Commit, Patch}` edit ordered sections/custom-show references and bounded `p15` guide XML while retaining unrelated extension namespaces |
| Headers, footers, and visibility settings | ✅ | ✅ | ✅ | Presentation and slide header/footer flags and related inherited settings are typed and serialized; layout is not calculated |

## Slides, text, shapes, and drawing content

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Text boxes, paragraphs, runs, and bullets | ✅ | ✅ | ✅ | Bounded DrawingML text extraction and formatted text-box authoring, including text defaults and paragraph/run content |
| Text styles and Kinsoku settings | ✅ | ✅ | ✅ | Presentation default text-style inventory and Kinsoku line-breaking settings are exposed as typed metadata |
| Basic shapes, groups, transforms, and geometry | ✅ | ✅ | ✅ | Rectangles, ellipses, lines, text boxes, nested groups, formatting, non-visual IDs, deterministic depth-first shape inventory, and semantic shape selection |
| Placeholders and master/layout inheritance | ✅ | ✅ | ✅ | Placeholder metadata and add/replace operations are validated against master/layout relationships; effective rendering is not performed |
| Images and picture backgrounds | ✅ | ✅ | ✅ | Embedded and linked picture resources, photo-album defaults, solid/gradient/pattern/picture backgrounds, inheritance, and relationship-resolved resources. `SourceBackedPresentation` additionally exposes a metadata-only scene-order inventory and exact zero-based reads for validated direct `p:pic` images under `/ppt/media/`; external targets remain inert and backgrounds, inherited owners, notes, charts, OLE previews, and ambiguous/MCE picture branches are refused |
| Tables and table styles | ✅ | ✅ | ✅ | Table extraction and authoring, table style inventories, and bounded style graph mutation |
| Hyperlinks and slide-jump targets | ✅ | ✅ | ✅ | External URL, email, and slide-jump values with relationship-aware read/write; targets are never opened or followed |
| Action and interaction settings | 🟡 | ✅ | 🟡 | Bounded inert click/hover actions, reserved PowerPoint action values, validated slide jumps, and declared targets; no action is activated or executed |
| Classic charts | ✅ | ✅ | ✅ | Per-slide chart inventory with relationship and part identity plus basic type, title, and legend metadata; chart creation and package storage are typed |
| Extended charts and embedded workbooks | ✅ | ✅ | ✅ | Multiple chart types, chart style/color parts, ChartEx, and embedded workbook resources with graph-aware package editing |
| SmartArt and diagrams | ✅ | ✅ | ✅ | Diagram data, layout, style, color part graphs, graphic frames, and builder support; unsupported producer extensions remain inert |

## Media, animation, and transitions

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Audio, video, posters, and media relationships | ✅ | ✅ | ✅ | Typed slide-level media resources, embedded/linked relationships, poster frames, trim/fade/bookmark metadata, and bounded payload authoring |
| Media tracks, captions, and narration metadata | 🟡 | ✅ | 🟡 | `presentation_properties::metadata::tracks::{Snapshot, Transaction, Commit, Patch}` provides bounded source-checked edits for `p173:tracksInfo`, caption identity/language/display location, and authored `p15:isNarration` values while preserving unknown extension XML and inert media targets; this is not a full track/player model |
| Animations and timing trees | ✅ | ✅ | ✅ | Shape effects, sequences, triggers, timing metadata, chart/diagram relationships, and timing on slides, layouts, and masters |
| Transitions and slide advance timing | ✅ | ✅ | ✅ | Typed effect/option combinations, direction/axis/corner/origin/shape/ripple/spoke variants, inherited slide/layout/master transitions, duration, speed, click, and timed advance |
| PowerPoint 2010 transition extensions | ✅ | ✅ | ✅ | Compatibility-choice ripple effects, typed direction and duration, deterministic fade fallback, bounded unknown-child retention, and safe prefix handling |
| Slide-show laser traces | ✅ | ✅ | ✅ | Typed inert p14:laserTraceLst coordinates and time offsets with validated extension storage; traces are never replayed or rendered |
| Slide-show event records | ✅ | ✅ | ✅ | `presentation_properties::metadata::events::{Snapshot, Transaction, Commit, Patch}` provides bounded source-checked trigger/media event CRUD with opaque extension preservation; events are never replayed or executed |

## Comments, embedded content, and extensibility

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Classic comments | ✅ | ✅ | ✅ | Validated authors, comments, anchors, dates, IDs, package relationships, and graph-safe CRUD |
| Modern comments | ✅ | ✅ | ✅ | Validated authors, anchors, replies, status, author references, and bounded task, reaction, moniker, and V2 command/change package models; command data is inert, unknown XML is preserved, unsafe mutations are rejected, and collaboration behavior is never executed |
| Embedded OLE and package objects | ✅ | ✅ | ✅ | `presentation::embedded::ole::slide::{Snapshot, Transaction, Commit, Patch}` adds source-checked slide-owned metadata, anchor, link, payload, add/remove/detach, and replace edits while preserving OLE/XML/MCE bytes; objects are never activated |
| ActiveX/control payloads | 🟡 | ✅ | 🟡 | Contextual `presentation::embedded::controls::slide::{Snapshot, Transaction, Commit, Patch}` edits bounded ActiveX metadata, binary replacement, detach, and orphan cleanup while preserving MCE choice/fallback branches and opaque payloads; controls are not instantiated, rendered, or executed |
| Embedded fonts | ✅ | ✅ | ✅ | Typed font inventory, faces, PANOSE/pitch/charset metadata, standard raw-font and PowerPoint EOT payload profiles, licensing checks, shared payload ownership, and ordered CRUD; PPTX exposes no Word-only font obfuscation path |
| Tags and customer data | 🟡 | ✅ | ✅ | Strict/transitional tag lists, direct and shape-owned tag CRUD, relationship inventory, customer-data and schema preservation, graph-safe shared-anchor handling, and inert values/extension markup |
| Revision and change information | 🟡 | ✅ | 🟡 | `presentation_properties::metadata::changes::{Snapshot, Transaction, Commit, Patch}` provides typed source-checked change/review metadata CRUD with opaque XML retention and relationship/orphan validation; this is not a complete collaborative change-history engine |
| Web extensions and Office Add-ins | 🟡 | ✅ | 🟡 | Bounded task-pane/web-extension graph CRUD, shared payload bytes, typed inert links, extension-site handling, and graph budgets; add-ins are never loaded or executed and native Office behavior is not implied |
| Ink annotations | ✅ | ✅ | ✅ | Bounded inert InkML content-part inventory and validated slide storage through p:contentPart and customXml relationships; no handwriting recognition, replay, rendering, or execution |

## Security, macros, and explicit extension coverage

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Password encryption | ✅ | ✅ | ✅ | Optional `encryption` feature supporting Standard AES-128/ECB with SHA-1 and Agile AES-128/CBC with SHA-1; bounded password-aware ingress keeps crypto and OPC limits independent, retains source mode, refuses implicit plaintext downgrade and semantic mutation, and provides explicit plaintext, selected-mode, retained-mode/rekey, and atomic save paths without changing inner OPC signature parts |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral OPC XMLDSig verification, signing, re-signing, and clearing for RSA-SHA256/ECDSA; certificate trust and revocation are outside the model |
| VBA projects and PPTM/PPSM/POTM macros | 🟡 | 🟡 | 🟡 | Macro-enabled relationship metadata plus bounded vbaProject.bin CFB/MS-OVBA project/module source parsing, deterministic cache-free authoring/removal, and kind-preserving package mutation; VBA is never executed |
| Selected 2010/2012 extension metadata | ✅ | ✅ | ✅ | Typed image-edit flags and DPI, chart-tracking reference mode, browse mode, laser color, media-control visibility, narration flags, plus `[MS-PPTX]` 2.2.9 slide `creationId` and shape `modId` values through a source-bound `change_tracking::{Snapshot, Edit, Commit, Patch}` owner; per-slide modification-ID uniqueness and unknown extension bytes are preserved under limits |

## Important unsupported or incomplete specification families

| Feature family exposed by the specifications | Status | Read | Write | Notes |
|-----------------------------------------------|--------|------|-------|-------|
| PowerPoint math extension (a14:m) | 🟡 | ✅ | ✅ | Typed bounded presentation-level `a14:m` math defaults (`brkBin`/`brkBinSub`), strict/transitional OMML validation, opaque unrelated-extension preservation, and transactional package CRUD; equation content and rendering remain outside the owner |
| Section, slide, and summary zoom objects | 🟡 | 🟡 | 🟡 | [MS-PPTX] 2.2.15 is covered by the contextual `shape::zoom::Owner`: typed target/property metadata, lossless fallbacks and unknown choices, package target/relationship validation, and transactional CRUD are implemented; rendering/layout is intentionally out of scope |
| 3D models and animated model3d content | 🟡 | ✅ | ✅ | Contextual `model3d::{Model, Scene, Asset, Preview}` owns bounded graphic-frame discovery, shared DrawingML scene metadata, inert GLB/preview resources, relationship validation, and atomic replace/remove operations; model3d animation semantics and rendering remain outside the owner |
| Designer/design-element/designer-property/designer-tag metadata | 🟡 | ✅ | ✅ | `[MS-PPTX]` 2.2.17/2.5 `p15:designElem`, 2.2.19/2.17.1 shape-scoped `p202:designPr`, and 2.2.20/2.17.3 slide-ID-scoped `p202:designTagLst` have typed bounded codecs/models. `shape::designer` and `presentation_properties::metadata::designer_tags` provide source-bound snapshots, candidate-reparsed transactions, stale-source checks, reversible patches, and atomic OPC publication; new-presentation writer shapes/slides can author properties/tags under the same limits. Ordered duplicate tag values are retained as inert data, while duplicate outer tag extensions are inventoried and refuse ambiguous singular mutation. Recognized extension ranges are edited without rewriting unknown neighboring extension markup; a public package-facade path for the new owners is not claimed. Designer Service invocation, rendering, fetching, and execution are never performed. |
| Classification metadata | ✅ | ✅ | ✅ | `[MS-PPTX]` 2.2.18 and 2.15 classification extensions are exposed through shape-scoped `classification::{Outcome, Snapshot, Editor}` and package/slide facades; the `none`, `hdr`, `ftr`, and `watermark` outcomes are typed, unknown extension entries are preserved, and selected-shape edits are atomic. No host classification or rendering behavior is performed |
| Modern comment task, reaction, and V2 change-command families | ✅ | ✅ | ✅ | Bounded typed semantic, wire, and package support for [MS-PPTX] 2.18 through 2.21; command data is inert, unknown XML is preserved, unsafe mutations are rejected, and collaboration behavior is never executed |
| Non-Ink content parts | ✅ | ✅ | ✅ | `presentation::embedded::content_parts::{Snapshot, Transaction, Commit, Patch}` provides bounded slide-owned anchor/relationship/target and payload lifecycle edits while retaining opaque XML and inert external URI values; no payload vocabulary is interpreted or executed |
| Full media TracksInfo/player semantics | 🟡 | ✅ | 🟡 | Selected media and narration metadata is bounded and inert; the full track schema, playback events, and player behavior are not a semantic runtime |
| Full extension-schema conformance | 🟡 | 🟡 | 🟡 | Unknown extension XML can be retained at supported owners under limits, but schema presence alone does not grant typed support for every versioned [MS-PPTX] or [MS-ODRAWXML] family |
| Rendering, layout, playback, and action execution | ❌ | N/A | N/A | The crate reads and writes package data; it does not render slides, resolve effective layout, play media or animation, follow links, or execute macros/add-ins |
