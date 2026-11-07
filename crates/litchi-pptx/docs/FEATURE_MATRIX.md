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
The primary deferred-source owner `source::Package` additionally supports immutable
shape-text edits with package-carrying commits, local and durable inverse/replay,
owned read projections, and guarded atomic save. Durable patches bind the
presentation root, selected relationships and exact before XML; transport is
bounded to 32 slide deltas and 128 MiB. Destination XML and execution limits
remain enforced. Structural verbs remain pending.

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
| Opened-presentation atomic composition | ✅ | ✅ | ✅ | `opened::{Snapshot, Edit, Commit, Patch, Diagnostics}` composes slide ordering/removal; shape text and common text-box, rectangle, ellipse, picture creation/removal; dependency-complete common-shape transfer; notes; table frames/styles; charts; media; masters/layouts; legacy and modern comments/replies/typed extensions; and general relationship-closure transfer in one immutable-root transaction. Exact before/after OPC resource patches cover part and relationship creation/removal, validate a complete candidate before atomic publication, and provide durable inverse, deterministic compatible joins, non-mutating three-way conflict plans, and bounded undo/redo history while preserving untouched parts and unknown XML. Group subtrees and externally attached connectors compute a complete top-level endpoint closure, then remap every common non-visual identity, connector endpoint, relationship-namespace attribute, and copied dependency part. Recognized table, classic-chart, DiagramML, and inert OLE graphic frames use that closure; DiagramML `r:dm/lo/qs/cs` targets are covered explicitly. `ShapeTransferRefusal` classifies nested selections, identity-less `contentPart` shapes, unknown extension shapes, unclassified graphic-frame payloads, missing identities, and unresolved connector endpoints before atomic publication. |
| Source-backed cross-slide image/chart leaf copy | 🟡 | ✅ | 🟡 | `source::Package::{plan_cross_slide_copy,edit_cross_slide_copy}` (record 0356b; `Edit::stage`/`copy`, `commit`, then `write_to` streams the publication) extends the dependency-free source-backed slide closure to a nonempty caller-bounded set of direct `p:pic` leaves and zero or more direct ordinary chart graphic-frame leaves under exactly one direct `p:spTree`. Pictures retain the 0382 requirement for exactly one direct `p:blipFill/a:blip r:embed` reference to an internal, relationship-free `/ppt/media/` `image/*` leaf. Each admitted chart has exactly one namespace-resolved `r:id` binding to an internal canonical `/ppt/charts/` relationship-free leaf with the destination dialect, chart content type, and valid chart root; ChartEx, embedded workbooks, `externalData`, style/color, chart drawing/userShapes, chart outbound relationships, and broader graph closures are refused. Distinct image/media and chart targets are copied once, deterministic destination part URIs and relationship IDs are allocated, and only exact namespace-resolved `r:embed`/`r:id` byte values are rewritten without XML normalization. Semantic picture parsing preserves bounded foreign, non-MCE, non-relationship `a:blip` attributes opaquely and accepts one valid unqualified `cstate` token; namespace-safe copy rewrites only the full-slide resolved relationship namespace. A full-slide unbound lexical relationship attribute returns `UnsupportedRelationship`; MCE, DTD, PI, unresolved/rebound namespaces, and duplicate/ambiguous resolved bindings refuse. The destination anchor permits and preserves other valid existing relationships while anchoring exactly one dialect-correct internal `slideLayout`; the full-slide namespace-aware source inventory is a fence, source catalog relationship reconstruction is fallible, physical ZIP media/chart deduplication is asserted, and strict XML end-name fences remain active. A bounded namespace/owner/URI proof admits passive native `a16:creationId`, `p14:creationId`/`modId`, leaf `themeFamily` (normative and native discriminator URIs), and empty presentation `p15:sldGuideLst` declarations without changing their source bytes; the newly admitted metadata rejects extra children/attributes, while unknown source-slide extension owners, nonempty guides and relationship-bearing extensions remain outside this copy profile. Existing core-only layout/master/theme admission is retained. Scalar decoding retains cancellation/memory accounting. Unselected XML and package members remain preserved while source/destination dialect, layout/master/theme equivalence, limits, freshness, lineage, signatures, destination physical members, publication cancellation, partial-sink, and resource fences remain checked. Selected image/chart relationship-ID collisions are handled by allocation/rewrite; layout, non-selected, and unsupported relationship-ID collisions still refuse. Wrong-type non-selected and non-anchor slide bindings are rejected at open, and planning revalidates every binding defense-in-depth. Missing, mistyped, external/linked, outbound, unreferenced, shared, or otherwise unsupported dependencies, duplicate shape trees, malformed/misplaced/ambiguous hosts, and broader closures fail closed before output. Image/chart decoding, conversion, rendering, and durable inverse are not provided. |
| Presentation size, root properties, and view settings | 🟡 | ✅ | ✅ | Typed slide/notes dimensions, size type, presentation and view settings, PowerPoint 2010 browse-mode state, and bounded extension serialization; settings do not render or drive a host UI |
| Sections | ✅ | ✅ | ✅ | Typed section IDs and names, resolved slide-index membership, and graph-safe ordered CRUD |
| Custom slide shows | ✅ | ✅ | ✅ | Typed named subsets and graph-safe ordered CRUD; `presentation_properties::metadata::custom_show::{Snapshot, Edit, Commit, Patch}` adds source-checked XML inverse edits |
| Slide masters and layouts | ✅ | ✅ | ✅ | Semantic master/layout reading with shape and placeholder inventory, typed layout references, matching/type metadata, `p232:phTypeExt` cameo/unknown placeholder metadata, source-checked reversible add/remove/update transactions across slide, master, and layout owners, relationship validation, new-master and layout authoring, unknown-shape-byte preservation, and unreferenced-layout removal |
| Handout master | 🟡 | ✅ | 🟡 | Direct header/footer/date/slide-number placeholder text and visibility, explicit `p:hf` overrides, and direct solid RGB background; bounded fresh XML authoring retains represented fields. The codec does not preserve source shapes/styles or serialize the separate print-layout/slide-image hints. |
| Speaker notes and notes masters | ✅ | ✅ | ✅ | `Package::{notes, put_notes, remove_notes, clear_notes}` uses semantic slide selectors, while `notes::{Snapshot, Edit, Commit, Patch}` provides bounded source-checked notes-slide/master/theme/text lifecycle edits; opaque XML and exclusively owned descendants are handled losslessly, slide comments remain slide-owned, and notes are never rendered or activated |
| Themes, color maps, and overrides | ✅ | ✅ | ✅ | Master-, layout-, and slide-scoped theme resolution, typed color maps and overrides, 12-slot color schemes, major/minor font schemes, attachment, replacement, removal, and orphan cleanup; contextual `color_map::{Snapshot, Edit, Commit, Patch}` adds source-checked map edits, while fmtScheme authoring is not covered |
| Presentation property structure and extended guides | ✅ | ✅ | ✅ | `presentation_properties::metadata::structure::{Snapshot, Edit, Commit, Patch}` and `guides::{Snapshot, Edit, Commit, Patch}` edit ordered sections/custom-show references and bounded `p15` guide XML while retaining unrelated extension namespaces |
| PowerPoint 2017/10 read-only recommendation | 🟡 | ✅ | ✅ | The normative `presentationPr/extLst/ext` view-mode owner for `p1710:readonlyRecommended` is exposed through `presentation_properties::readonly_recommended::{Snapshot, Edit, Commit, Patch}` and `Package` CRUD. The required `xsd:boolean` value is typed and inert: it recommends opening with an option to edit, never enforces host access policy. MCE branch selection, strict/transitional namespaces, unknown extension bytes, source-checked inverse patches, lexical no-op writes, and explicit signed-source mutation policy are bounded and preserved. |
| Headers, footers, and visibility settings | ✅ | ✅ | ✅ | Presentation and slide header/footer flags and related inherited settings are typed and serialized; layout is not calculated |

## Slides, text, shapes, and drawing content

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Text boxes, paragraphs, runs, and bullets | ✅ | ✅ | ✅ | Bounded DrawingML text extraction and formatted text-box authoring, including text defaults and paragraph/run content |
| Text styles and Kinsoku settings | 🟡 | 🟡 | 🟡 | Presentation `defaultTextStyle` and `kinsoku` are retained as source XML by package operations; the current presentation owner has no dedicated typed accessor or semantic transaction for these children. The Kinsoku asset test inspects raw XML and does not establish a typed facade. Native tests pin saved default-style values and exact preservation during an independent property edit, without style resolution or line breaking. |
| Basic shapes, groups, transforms, and geometry | ✅ | ✅ | ✅ | Rectangles, ellipses, lines, text boxes, nested groups, formatting, non-visual IDs, deterministic depth-first shape inventory, and semantic shape selection |
| Placeholders and master/layout inheritance | ✅ | ✅ | ✅ | Placeholder metadata and add/replace operations are validated against master/layout relationships; title-family types and subtitle indices bind to native layout slots, and bound shapes inherit geometry through an empty `spPr`. Effective rendering is not performed |
| Images and picture backgrounds | ✅ | ✅ | ✅ | Embedded and linked picture resources, photo-album defaults, solid/gradient/pattern/picture backgrounds, inheritance, and relationship-resolved resources. `SourceBackedPresentation` additionally exposes a metadata-only scene-order inventory and exact zero-based reads for validated direct `p:pic` images under `/ppt/media/`; the `[MS-ODRAWXML]` 2.26 `asvg:svgBlip` extension is typed as an optional SVG descriptor with deferred embedded payload reads. Native vector-only pictures without a raster reference use the SVG as their primary image; linked SVG targets remain inert. External targets remain inert and backgrounds, inherited owners, notes, charts, OLE previews, and ambiguous/MCE picture branches are refused |
| Source-backed SVG attachment lifecycle | 🟡 | ✅ | 🟡 | `source::Package::edit_svg_images` (`Edit::attachment`, `attach`, `detach`; record 0356a) selects an existing direct slide picture with an internal PNG fallback. Borrowed `attach_svg`/advanced attachment requests and `detach` stage source-preserving changes; fallible `commit()` validates the changed dependency closure before publication, which rechecks source freshness and atomically publishes slide XML, relationships, and SVG media. Shared references retain their media, and inverse patches restore the authorized in-memory source exactly. Inherited namespaces, BOMs, wrapper metadata, and opaque foreign extension payloads are preserved within explicit XML and caller limits. Imported SVG bytes retain producer whitespace under the bounded source-XML audit (well-formed XML, one root, no DTD); library-authored XML still requires compact output. Empty source relationship elements can be retargeted while preserving unrelated relationship XML; opaque child content remains refused. SVG media remains inert and is not rendered. Picture creation, ambiguous/MCE picture ownership, external SVG attachment, rendering, native application acceptance, and durable cross-process inverse transport are outside this scope. |
| Tables and table styles | ✅ | ✅ | ✅ | Table extraction and authoring, table style inventories, and bounded style graph mutation |
| Hyperlinks and slide-jump targets | ✅ | ✅ | ✅ | External URL, email, and slide-jump values with relationship-aware read/write; targets are never opened or followed |
| Action and interaction settings | 🟡 | ✅ | 🟡 | Bounded inert click/hover actions, including shape `a:hlinkHover` and text-run `a:hlinkMouseOver`, reserved PowerPoint action values, validated slide jumps, and declared targets; no action is activated or executed |
| Classic charts | ✅ | ✅ | ✅ | Per-slide chart inventory with relationship and part identity plus basic type, title, and legend metadata; chart creation and package storage are typed |
| Extended charts and embedded workbooks | ✅ | ✅ | ✅ | Multiple chart types, chart style/color parts, ChartEx, and embedded workbook resources with graph-aware package editing. Native attribute-form subtotal indices are read alongside the historical text form; mixed forms, duplicate indices and overflow remain errors. |
| SmartArt and diagrams | ✅ | ✅ | ✅ | Diagram data, layout, style, color part graphs, graphic frames, and builder support; unsupported producer extensions remain inert |

## Media, animation, and transitions

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Audio, video, posters, and media relationships | ✅ | ✅ | ✅ | Typed slide-level media resources, embedded/linked relationships, poster frames, trim/fade/bookmark metadata, and bounded payload authoring  Native `p14:extLst` is retained, and an internal `p14:media@r:embed` can supply the payload beside an inert external legacy placeholder. Both relationship types and modes remain checked. |
| Media tracks, captions, and narration metadata | 🟡 | ✅ | 🟡 | `presentation_properties::metadata::tracks::{Snapshot, Edit, Commit, Patch}` provides bounded source-checked edits for `p173:tracksInfo`, caption identity/language/display location, and authored `p15:isNarration` values while preserving unknown extension XML and inert media targets; this is not a full track/player model |
| Animations and timing trees | ✅ | ✅ | ✅ | Shape effects, sequences, triggers, timing metadata, chart/diagram relationships, and timing on slides, layouts, and masters. Native single-target presets may have several behaviors sharing that target; absent preset duration uses the longest direct behavior duration. Distinct targets in one preset have a typed UnsupportedContent projection boundary; no playback is performed. |
| Transitions and slide advance timing | ✅ | ✅ | ✅ | Typed effect/option combinations, direction/axis/corner/origin/shape/ripple/spoke variants, inherited slide/layout/master transitions, duration, speed, click, and timed advance |
| PowerPoint 2010 transition extensions | ✅ | ✅ | ✅ | Compatibility-choice ripple plus the 18 remaining normative p14 effects (`conveyor`, `doors`, `ferris`, `flash`, `flip`, `flythrough`, `gallery`, `glitter`, `honeycomb`, `pan`, `prism`, `reveal`, `shred`, `switch`, `vortex`, `warp`, `wheelReverse`, and `window`), typed direction/options, exact bounded `h`, `min`, `s`, `ms`, `µs`, `ns`, and unitless decimal durations, deterministic fade fallback, bounded unknown-child retention, and safe prefix handling |
| Morph and preset transitions | ✅ | ✅ | ✅ | Typed `p159:morph` (`byObject`, `byWord`, and `byChar`) and `p15:prstTrans` (`prst`, `invX`, and `invY`) values, MCE choices with deterministic fade fallbacks, bounded preset names, source-backed direct-transition editing, and exact source XML preservation for semantic no-op writes |
| Slide-show laser traces | ✅ | ✅ | ✅ | Typed inert p14:laserTraceLst coordinates and time offsets with validated initial extension storage and reversible patches. Replacing an existing trace through `edit_laser_traces().set` is explicitly refused; traces are never replayed or rendered |
| Slide-show event records | ✅ | ✅ | ✅ | `presentation_properties::metadata::events::{Snapshot, Edit, Commit, Patch}` provides bounded source-checked trigger/media event CRUD with opaque extension preservation; events are never replayed or executed |

## Comments, embedded content, and extensibility

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Classic comments | ✅ | ✅ | ✅ | Validated authors, comments, anchors, dates, IDs, package relationships, graph-safe CRUD, and source-bound inert `p15:presenceInfo`/`p15:threadingInfo` metadata with exact owner-URI closure |
| Modern comments | ✅ | ✅ | ✅ | Validated authors, anchors, replies, status, author references, and bounded task, reaction, moniker, and V2 command/change package models; command data is inert, unknown XML is preserved, unsafe mutations are rejected, and collaboration behavior is never executed |
| Embedded OLE and package objects | ✅ | ✅ | ✅ | `presentation::embedded::ole::slide::{Snapshot, Edit, Commit, Patch}` adds source-checked slide-owned metadata, anchor, link, payload, add/remove/detach, and replace edits while preserving OLE/XML/MCE bytes; objects are never activated |
| ActiveX/control payloads | 🟡 | ✅ | 🟡 | Contextual `presentation::embedded::controls::slide::{Snapshot, Edit, Commit, Patch}` edits bounded ActiveX metadata, binary replacement, detach, and orphan cleanup while preserving MCE choice/fallback branches and opaque payloads; controls are not instantiated, rendered, or executed |
| Embedded fonts | ✅ | ✅ | ✅ | Typed font inventory, faces, PANOSE/pitch/charset metadata, standard raw-font and PowerPoint EOT payload profiles, licensing checks, shared payload ownership, and ordered CRUD; PPTX exposes no Word-only font obfuscation path |
| Tags and customer data | 🟡 | ✅ | ✅ | Strict/transitional tag lists, direct and shape-owned tag CRUD, relationship inventory, customer-data and schema preservation, graph-safe shared-anchor handling, and inert values/extension markup |
| Revision and change information | 🟡 | ✅ | 🟡 | `presentation_properties::metadata::changes::{Snapshot, Edit, Commit, Patch}` provides typed source-checked change/review metadata CRUD with opaque XML retention and relationship/orphan validation; this is not a complete collaborative change-history engine |
| Web extensions and Office Add-ins | 🟡 | ✅ | 🟡 | Bounded task-pane/web-extension graph CRUD, shared payload bytes, typed inert links, extension-site handling, and graph budgets; add-ins are never loaded or executed and native Office behavior is not implied |
| Ink annotations | ✅ | ✅ | ✅ | Bounded inert InkML content-part inventory and validated slide storage through p:contentPart and customXml relationships, including native PowerPoint 2010 MCE choices; no handwriting recognition, replay, rendering, or execution |
| PowerPoint 2014 InkAction owners | 🟡 | 🟡 | 🟡 | `presentation::embedded::ink_actions` reads existing slide-owned MS-PPTX 2.2.3.1 `mc:AlternateContent` owners with exactly the supported `p14main`/`iact` Choice, `p:contentPart`, and sole `p:pic` fallback shape. Transitional and Strict slides require their matching `customXml` relationship URI and an explicitly declared `text/xml` target; target profiles reuse the shared source-preserving `iact:actions` model. Existing target scalar profile edits, exact no-op/inverse patches, stale-source/read-policy/content-type/relationship-closure checks, shared-target fan-out, signed-source refusal, and bounded readback are supported while inactive MCE/unknown bytes remain intact. This batch does not create or remove owners/parts, infer an InkML MIME or path, execute/render/recognize actions, or claim native Office interoperability. See the [owner design](../../../docs/report/spec-gap-validation-evidence/pptx-ink-actions-design.md) and [focused tests](../tests/pptx_ink_actions.rs). |

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
| Section, slide, and summary zoom objects | 🟡 | 🟡 | 🟡 | [MS-PPTX] 2.2.15 is covered by the contextual `shape::zoom::Owner`: typed target/property metadata, lossless fallbacks and unknown choices, package target/relationship validation, and transactional CRUD are implemented; rendering/layout is intentionally out of scope  Native graphic-frame hosts and Zoom-specific `zmPr` namespaces are read; source property edits retain frame identity/transforms, fallback and unrelated package members. Fresh detached Zoom authoring is not native-qualified by these regressions. |
| 3D models and animated model3d content | 🟡 | ✅ | ✅ | Contextual `model3d::{Model, Scene, Asset, Preview}` owns bounded graphic-frame discovery, shared DrawingML scene metadata, inert GLB/preview resources, relationship validation, and atomic replace/remove operations; model3d animation semantics and rendering remain outside the owner  Native assets use `model/gltf.binary`; historical `model/gltf-binary` inputs remain readable. Native GLB replacement preserves the preview and scene children. |
| Designer/design-element/designer-property/designer-tag metadata | 🟡 | ✅ | ✅ | `[MS-PPTX]` 2.2.17/2.5 `p15:designElem`, 2.2.19/2.17.1 shape-scoped `p202:designPr`, and 2.2.20/2.17.3 slide-ID-scoped `p202:designTagLst` have typed bounded codecs/models. `shape::designer` and `presentation_properties::metadata::designer_tags` provide source-bound snapshots, candidate-reparsed transactions, stale-source checks, reversible patches, and atomic OPC publication; new-presentation writer shapes/slides can author properties/tags under the same limits. Ordered duplicate tag values are retained as inert data, while duplicate outer tag extensions are inventoried and refuse ambiguous singular mutation. Recognized extension ranges are edited without rewriting unknown neighboring extension markup. Public package families `edit_shape_design_element`, `edit_shape_designer_properties` and `edit_slide_designer_tags` expose immutable commits and reversible patches. Designer Service invocation, rendering, fetching, and execution are never performed. |
| Classification metadata | ✅ | ✅ | ✅ | `[MS-PPTX]` 2.2.18 and 2.15 classification extensions are exposed through shape-scoped `classification::{Outcome, Snapshot, Edit, Commit, Patch}` and package/slide facades; the `none`, `hdr`, `ftr`, and `watermark` outcomes are typed, unknown extension entries are preserved, and selected-shape edits are atomic. No host classification or rendering behavior is performed |
| Modern comment task, reaction, and V2 change-command families | ✅ | ✅ | ✅ | Bounded typed semantic, wire, and package support for [MS-PPTX] 2.18 through 2.21. Native text-anchored threads cover eager/source reads, task-field updates, reply CRUD, and library-authored task-history/reaction extension add/update/remove with exact unowned-part preservation and reversible patches. The native threads contain no task/reaction extensions or V2 commands; derivatives use caller-selected extension URIs and do not establish Office recognition or collaboration behavior. V2 command native qualification remains open. |
| Non-Ink content parts | ✅ | ✅ | ✅ | `presentation::embedded::content_parts::{Snapshot, Edit, Commit, Patch}` provides bounded slide-owned anchor/relationship/target and payload lifecycle edits, including inert typed `p14:bwMode` metadata on `p:contentPart` anchors (all `ST_BlackWhiteMode` tokens; authored absence remains distinct from `auto`), while retaining opaque XML and inert external URI values; no payload vocabulary is interpreted or executed |
| Full media TracksInfo/player semantics | 🟡 | ✅ | 🟡 | Selected media and narration metadata is bounded and inert. Native caption identity, language and display-location edits retain exact trim/fade metadata, WebVTT cues and unselected package members. Language removal/reinsertion, failure atomicity and inverse replay are covered. The native specimen has no narration owner; creating one is refused. Caption-part creation is separate from media-descriptor creation and does not replace an existing track. The full track schema, playback events and player behavior remain outside this runtime. |
| Full extension-schema conformance | 🟡 | 🟡 | 🟡 | Unknown extension XML can be retained at supported owners under limits, but schema presence alone does not grant typed support for every versioned [MS-PPTX] or [MS-ODRAWXML] family |
| Rendering, layout, playback, and action execution | ❌ | N/A | N/A | The crate reads and writes package data; it does not render slides, resolve effective layout, play media or animation, follow links, or execute macros/add-ins |

### Unified shape-text publication (record 0123)

The `litchi-ops` adapter supports immutable slide shape/placeholder text sets,
scoped selectors, transactions, batches, replay/inverse and atomic save through
the source and eager native owners. Native `Package::{same_content, write_to}`
provide content comparison and immutable publication for already-opened eager
packages; pending mutable authoring state is refused. The source counterpart
compares normalized overlays within one source lineage. Managed signature
removal remains an explicit integration gap. Record 0124 integrates immutable
retained-profile re-encryption through `Package::write_reencrypted` and the
unified atomic save options, while ordinary plaintext output remains protected.

### Native builder and unified creation (record 0125)

`builder::Builder` configures native typed authoring and finalizes to a canonical
package for immutable editing and publication. Unified `create(Pptx)` adopts it
without a ZIP round trip. It creates an empty Transitional presentation; template,
Strict and initial unified metadata options remain explicit capability gaps.
Record 0127 completes the public constructor split: `Package::new()` returns a
canonical empty presentation directly usable with immutable edits and publication.
`Package::presentation_mut`, `set_font_embedding` and `with_font_embedding` are
removed; native fixtures and doctests use the builder, whose font option precedes
finalization. Remaining direct package mutators still require edit/commit migration.

## Native Microsoft PowerPoint regressions

`tests/microsoft_office_native.rs` exercises all fourteen unencrypted PPTX, POTX,
PPSX, PPTM, POTM and PPSM specimens plus the encrypted PPTX. The tests distinguish
ZIP recompression from byte preservation of every OPC member. Focused checks cover
visibility/layout metadata, native animation composition, hover actions, Zooms,
ChartEx data, 3D payloads, video/captions, vector-only SVG, sections/custom shows,
notes and comment replies. Native Zoom, 3D and macro-variant text edits are reopened
and inverted with unrelated content checked; Agile AES-256/SHA-512 input is rekeyed
without changing clear member payloads. No Office application, macro, model, media
player, external target or animation is executed.

The native `powerpoint-fonts-math-handout.pptx` regression checks the single
Cambria Math Regular face and the exact 129261-byte EOT payload/SHA-256 before
and after package serialization. Font bytes remain inert; this is not glyph
coverage, registration, rendering or font-edit evidence.

The same fixture now checks the exact native handout header/footer labels,
slide-number presence, absent date, and `00B0F0` background. The parser recognizes
direct placeholder owners when `p:hf` is absent, uses schema Boolean defaults
when it is present, and excludes foreign/nested shadow elements. Fresh XML
authoring retains the represented text, flags and RGB value; it does not
preserve native shapes, styles or layout. Generated field IDs have the required
8-4-4-4-12 GUID segments. The canceled print-dialog selection saved no `p:prnPr`,
so no six-slide layout is inferred. `layout` and `show_slide_images` remain
authoring hints outside this codec.

Native slide equations are checked separately from presentation-level math
defaults: shape IDs, exact fraction/radical/superscript/sum formulae and raw
`m:t` text, plus unchanged source XML and fallback PNGs after package reopen.
The matrix remains a documented text-only projection; row/column semantics,
editing, calculation and rendering are not covered.

The master-motion fixture checks the exact Ion palette, system-color fallbacks,
Century Gothic fonts and owning theme relationship/hash. Its transition checks
cover Push/Up/Medium with 3000ms automatic advance, Morph/ByObject with a saved
1500ms extension duration, absent transitions and semantic fragment write/read.
The Push part has no explicit duration despite the historical UI observation.
These read/fragment tests do not qualify every transition extension, theme
overrides or native playback. Native source edits are qualified separately below.

`tests/microsoft_office_native_visual.rs` checks all twelve native table cells,
grid widths, row heights and style switches. Column-width reads retain the
graphic frame's namespace declarations and the admitted standalone table's
fragment prefix; source edits still require actual namespace bindings. Native
row/column insertion, removal, width and style-switch edits preserve unrelated
package members, synchronize the frame extent and reopen with the expected
values. Public operations cover table queries, cell text and row height on
file-backed Auto and eager inputs.

The visual-object PPTX/POTX/PPSX charts expose their exact clustered-column
family, title, series formulas and all cached values. A typed chart-title edit
keeps the workbook, slide and companion parts unchanged; the public facade
also reads all twenty cells of the embedded workbook. Native BasicProcess
SmartArt tests bind the exact document/content IDs to the three labels and
update both a content node and its cached drawing while retaining layout,
style, color and relationship parts. Grouped-shape tests cover child order,
IDs, bounds, coordinate transforms and a source-preserving child-text edit.
These are inert data and mutation checks, not chart refresh, SmartArt layout,
rendering or native-app reopen evidence. Merged-table examples, connector
endpoint binding and broader visual inheritance remain separate gaps.

`tests/microsoft_office_native_ink.rs` covers the native ink stroke and replay
decks. Specialized InkML discovery selects the PowerPoint 2010 content-part
choice, reports one trace and retains the exact payload/PNG hashes. Shared
InkML reads verify the native brush dimensions/color, trace references and
raw trace bytes; source serialization is exact. Replay timing retains its
`entr:63` preset identity, target, click trigger and 2000ms duration without
execution. These specimens have no distinct PowerPoint 2014 InkAction anchors.

Native content-part transactions change an existing `bwMode` attribute and
replace an inert payload while preserving the original anchor markup,
fallback, timing and relationships. Source edits select only active anchors,
even when inactive alternatives reuse a relationship ID. Relationship cleanup
recognizes namespace aliases, quoting and escaped values in all retained
branches. Adding the native payload under a fresh relationship/part also
reopens and inverses exactly. These tests do not regenerate the raster preview,
author Replay behavior, position newly attached ink, or claim fresh native-app
validation. Removing the baseline's active anchor collects its orphan InkML
payload and retains the inactive picture fallback; inverse restores the source.
Whole visual-object deletion, replay-target cleanup and shared-payload removal
remain open.

`tests/microsoft_office_native_transitions.rs` qualifies direct transition
edits beside native Zoom MCE. A separate source admission mode retains the
complete shape tree while allowing transition splices; source positional shape
edits and removal keep their guards. Creating, replacing, clearing and
reinserting transitions after reopen preserves shape XML, animation timing,
inactive transition fallbacks and unrelated package members. Inherited PML
namespace aliases remain resolvable when locating the active transition.

The variant sweep applies transitions to all 52 slides in the fourteen
unencrypted native presentation, template, slideshow and macro-enabled files
through both eager and source-backed APIs. Macro payloads, media, embedded
workbooks and other resources remain inert and byte exact. Source overlays
retain original archive bytes for no-op/inverse; eager publication follows its
existing ZIP recompression and compact changed-XML declaration policy. Public
operations cover atomic Push/Morph updates and clear/reinsert in both modes.
Durable transition-patch wire grammar, general shape-MCE mutation and playback
are not promoted by these tests.

Zoom discovery now respects default namespace resets (`xmlns=""`), including
the native table and equation slides, while named-prefix undeclarations remain
refused. Its presentation-structure dependency accepts all six container MIME
types; native catalog tests bind exact IDs, part names and relationships and
reject a non-presentation main-part content type. Extended guides and variant
master/layout authoring need separate qualification.

`tests/microsoft_office_native_ole.rs` qualifies the native `Excel.Sheet.12`
frame, workbook payload and EMF preview. Package relationships retain the
embedded file's declared MIME (XLSX in this specimen); binary OLE relationships
still require the OLE-object MIME. Metadata and payload edits retain that MIME
and the original Strict/Transitional relationship identity. The source editor
uses the inventory's MCE choice policy, so it edits the active native object
while keeping the inactive preview branch byte exact.

Native tests cover metadata changes, opaque payload replacement, addition,
removal/detach, invalid/shared-target rollback and inverse publication. The
source shape-removal route additionally covers durable forward and inverse
replay while retaining resource relationships by contract. Graph-owning
removal collects the embedded workbook; the preview, which belongs to the
inactive branch, remains outside the active object's resource ownership.
Facade tests read the child chart/worksheet catalog, four formulas and stored
caches, edit an unrelated child cell and put the changed workbook back while
preserving all other parent members. OLE objects stay inert; no activation,
formula calculation, preview regeneration or fresh Office validation occurs.

`microsoft_office_native_copy.rs` qualifies source-backed JPEG slide copying
between independent instances of the native PPTX, PPSX and POTX visual-object
specimens, with exact source slide/media data, preserved manifest entries,
deterministic publication, and reopened inverse using the retained process-local
patch. This does not provide a durable inverse wire format. Copying between
unequal native variant dependency graphs remains refused, as do the native chart
with embedded-workbook/style/color relationships, the SVG extension binding and
an MCE-bearing master. Native p14 creation/modification identifier reads and
updates have separate exact-byte preservation assertions; DrawingML creation
GUIDs remain unchanged and no collaboration history is inferred.

## Native presentation preference qualification

`tests/microsoft_office_native_preferences.rs` exercises the visual-object and
passive-macro presentation, slideshow and template variants. These native files
have no read-only recommendation or programmable-tag owners. The tests create
Litchi derivatives, read and update their typed values, remove their owners, and
verify file reopen, source-checked replay/inverses and unchanged payloads. An
inert read-only recommendation still permits edits. Slide tags stay distinct
from title-shape tags, and failed selectors retain prior staged work. All macro,
chart, image and other unrelated resources remain exact.

The native default style has nine levels, 457200-EMU indentation steps,
914400-EMU default tabs and 1800/1200 size/kerning values. These are explicit raw
source assertions and preservation checks, not a typed Kinsoku/default-style
API or layout calculation. Borrowed-buffer ingress preserves member payloads
while it may re-encode the ZIP; reader ingress also retains exact no-op archive
bytes. Graph-changing inverses compare relationships and effective content
types, with byte-exact comparison of all other members.

Laser-trace creation is a Litchi-authored addition to native visual, macro and
motion slides. Show-event CRUD targets the existing video shape (ID 4 on slide
2) in both native media presentations, after adding a laser trace to that slide.
Tests preserve existing creation IDs, shape content and media, retain a prior laser extension during event edits, and
check exact replacement/empty-event refusals without losing staged work. The
corpus specimens do not supply native-authored laser/event data for these tests.
No fresh PowerPoint reopen, playback, rendering or macro execution is claimed.

`tests/microsoft_office_native_evidence_review.rs` audits all 15 native
PresentationML files, including the explicitly decrypted encrypted fixture.
Typed catalogs and all-part XML/relationship inventories find no ActiveX control,
Revision Information or Changes Information owner. The two modern comment parts
and four OLE anchors (including compatibility alternatives) are distinct evidence.
The only two `contentPart` anchors belong to the Ink fixtures and target their
2424-byte InkML payloads; inactive branches and orphan XML parts add no other
anchors. Non-Ink content parts still need a native specimen. These three gaps
are recorded separately from positive semantic coverage, without changing the
historical native-authoring inventory or claiming fresh PowerPoint validation.

Full extension-schema conformance is an explicit scope boundary in
`tools/native_office_evidence_reviews.json`. The native chart-slide identifier
regression changes only two existing `p14` values while retaining opaque `a16`
creation GUIDs, every other XML byte and the full chart/style/color/workbook
graph through file reopen and inverse. That concrete preservation profile does
not establish universal MS-PPTX/MS-ODRAWXML grammar or application semantics.
The report counts the boundary separately from missing native specimens and
positive semantic assertions; concrete extension owners remain independently
qualified by their tests.

The runtime boundaries `pptx-452`
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

The operations refresh adapter also updates supported slide-number and date
field caches. Native reordered slides without MCE refresh through both ingress
routes; library-authored date-field derivatives require an explicit clock.
Changed fields on native MCE-bearing slides retain the text editor's precise
`UnsupportedContent` refusal and leave the source unchanged. This is field-data
publication, not slide layout or playback.

Native-parent task-pane tests qualify the public package owner on ordinary and
macro-enabled sources: two-pane lifecycle, inert provider/binding/custom-function
metadata, shared snapshot bytes, opaque extension retention, semantic no-op and
stale/resource-collision refusals. File reopen/replay/inverse retains all data
parts and complete effective graphs; owned root-relationship/content-type XML may
normalize during generic replay. Additions are explicit library derivatives, with
no add-in activation, worksheet binding resolution or fresh Office validation.

Native-parent classic-comment tests now cover graph creation, author/comment
CRUD and ordering, eager/source reader agreement, exact invalid-input/stale
refusals, and file reopen/replay/inverse beside existing modern threads and VBA.
The whole-graph `Edit::set` creates a previously absent graph; existing records
use the dedicated verbs. Removing the final records retains an empty author part.
Classification and Designer flag/property/tag derivatives retain shape identity,
text, geometry, unselected shape XML and every unowned package member. Optional
values, explicit empty lists and removal remain distinct. No comment service,
Designer service or host classification action is performed.

The classification owner now resolves a selected shape in its actual inherited
namespace scope, including aliases/defaults and Strict dialects. InkAction
discovery no longer treats text inside unrelated MCE branches as malformed action
content; action-capable branches retain their whitespace restrictions, including
nested content. All 15 native PPTX-family files lack 2014 InkAction source/active
owners and payloads, while the two native InkML roots remain present. pptx-438
is a reviewed fixture gap, without positive action-payload credit. Native-authored
classic/Designer/classification specimens and fresh Office validation remain open.
