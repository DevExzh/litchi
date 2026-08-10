# ODP Feature Matrix

This document tracks the public and source-level feature coverage of the
`litchi-odp` OpenDocument Presentation implementation for packaged ODP files.
OTP template-package preservation is not currently claimed. This is a capability matrix, not a claim of complete ODF conformance,
PowerPoint compatibility, playback fidelity, or rendering fidelity.

The ODF common package layer is shared with the other OpenDocument families.
Rows below distinguish semantic slide/shape support from inert XML metadata,
opaque embedded resources, and source files that are not connected to the
public crate exports.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the feature scope described in the Notes cell |
| 🟡 | Bounded, partial, metadata-only, pass-through, source-level, or otherwise limited support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to the format or direction |

`Read` and `Write` describe the public direction independently. A 🟡 direction
must not be read as full semantic CRUD. URLs, scripts, media plugins, actions,
animations, and embedded payloads are inert unless a row says otherwise.

## Audit scope

The public ODP path includes package opening, slide parsing, slide/text and
shape authoring, package-contained image/media access, page-layout models,
presentation settings, declarations, page metadata, custom shows, metadata,
RDF graph editing, password opening, and a unified source-checked transaction
for slides, shapes, media, charts, layouts, masters, annotations, and RDF. Durable
patches can materialize the RDF/non-RDF merge cases the conservative planner proves
independent, bounded undo/redo timelines have a validated durable envelope, and
dependency-closed charts, rich-text boxes, tables, and inert form controls can transfer
between deck snapshots with deterministic style/resource collision remapping.
Attached mutable presentation roots are private implementation
details and cannot be obtained through the public API.

The Microsoft `[MS-PPTX]` Front Matter and ToC describe extensions to OOXML
PresentationML, not ODP. Their Part Enumerations, Extensions, structure
families, examples, and security sections were used as a gap checklist only;
the existence of a PPTX feature is never treated as ODP support.

## Package and shared ODF features

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open, create, and save ODP packages | ✅ | ✅ | ✅ | `Presentation` opens paths or bytes and saves/to-bytes; `Builder` creates presentation packages. OTP template-package opening and authoring are not claimed by this facade. |
| Password-protected ODP opening | ✅ | ✅ | N/A | `open_with_password` and `from_bytes_with_password` use the shared ODF encrypted-package reader; this does not imply password/encryption authoring |
| ZIP package, manifest, and MIME validation | ✅ | ✅ | ✅ | Shared ODF package validation and deterministic package writing cover manifest/resource access and the ODP presentation media type |
| `content.xml`, `styles.xml`, `meta.xml`, and `settings.xml` | ✅ | ✅ | ✅ | The public presentation/package and builder paths parse or write the core package XML parts; unsupported children may remain inert/preserved package data |
| ODF metadata and statistics | ✅ | ✅ | ✅ | `Presentation::metadata` and `Builder::set_metadata` use the shared typed metadata model, including Dublin Core, user-defined values, template/reload/link metadata, and statistics |
| Common styles and data styles | 🟡 | 🟡 | 🟡 | Direct presentation/drawing properties and bounded shared style values are supported; there is no complete public style-graph resolver/editor for every ODF style family |
| Images and package media | ✅ | ✅ | ✅ | Package-contained and linked image/media references are typed; package bytes are read without fetching external URLs, and builder media embedding creates package resources |
| Embedded objects and OLE-like payloads | 🟡 | ✅ | 🟡 | `embedded_objects()` provides a bounded inert inventory for regular objects, OLE payloads, applets, plugins, and floating frames, including storage/link classification and applet/plugin parameters. Payloads are never opened, fetched, activated, executed, or rendered |
| Embedded charts | 🟡 | ✅ | 🟡 | Bounded chart parts and frame/storage context are typed; `chart_snapshot()` and the unified `edit::Transaction` provide source-checked add/remove/replace commits, complete typed-definition replacement, fine series/cached-cell CRUD, dependency-closed cross-deck transfer with package-resource collision remapping, typed readback, and reversible exact-source patches. There is no recalculation or rendering engine |
| Annotations/comments | ✅ | ✅ | ✅ | `annotation::{Anchor, Info, Position}` inventories shared rich ODF annotations at validated pages or uniquely named shapes; `Presentation` and the unified transaction provide atomic add/replace/remove while untouched XML and no-op bytes remain preserved |
| Hyperlinks and external references | ✅ | ✅ | ✅ | Shape links, XLink targets, show/actuate values, page jumps, and action metadata are typed and serialized; targets are never opened or followed |
| Forms and controls | ✅ | ✅ | ✅ | The unified transaction authors common inert button/text/checkbox pairs; `content::FormControlModel` inventories and atomically whole-owner edits arbitrary producer declaration/control pairs while preserving attributes and children. Controls are never executed |
| Scripts, events, and macros | 🟡 | 🟡 | 🟡 | Event/action and script-binding metadata can be represented as inert values; no script, macro, or event execution occurs |
| RDF metadata graphs | ✅ | ✅ | ✅ | Graph and triple inventory plus ordered add/replace/remove/move operations are exposed on `Presentation`; the same CRUD can be committed atomically with slide, shape, media, chart, design, and annotation edits through `edit::Transaction` |
| ODF encryption authoring | 🟡 | ✅ | ❌ | Password opening is supported and `edit::SecurityPolicy` exposes encrypted packages as read-only; transactions and durable patch application refuse encrypted sources. The public builder has no encrypt/password-change operation |
| ODF digital signatures | 🟡 | 🟡 | ❌ | `edit::SecurityPolicy` detects document/macro signature owners and makes signed packages read-only before staging or durable patch application. Cryptographic verification and sign/add/clear operations are not provided |
| Unknown package-part preservation | 🟡 | ✅ | 🟡 | The owned package can retain unrelated resources around supported edits; preservation is not semantic understanding or guaranteed lossless mutation of every extension |

## Presentation content and layout

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Slides, ordering, and slide count | ✅ | ✅ | ✅ | `Presentation` parses slide order/count; `Builder` constructs detached decks; `edit::{Snapshot, Transaction, Commit, Patch}` provides checked add/insert/remove publication without mutating the source snapshot |
| Slide titles and body text | ✅ | ✅ | ✅ | Plain title/body text is extracted and authored; `Slide::all_text` includes labeled shape text |
| Rich text runs, paragraphs, and lists | ✅ | ✅ | ✅ | Common paragraph/run values are typed; `content::TextBoxModel` inventories arbitrary producer stories and nested lists as compact namespace-complete source fragments with paragraph/list projections and whole-owner editing, retaining unmodeled inline markup |
| Text boxes and basic shapes | ✅ | ✅ | ✅ | Rectangles, ellipses, lines, connectors, groups, frames/text boxes, and their core geometry/properties are represented by `Shape` and emitted by the builder |
| Advanced ODF drawing shapes | 🟡 | 🟡 | 🟡 | Polylines, polygons, paths, captions, measures, custom shapes, and 3D scene/object kinds are recognized or retained with inert/unmodeled geometry attributes; complete rendering semantics are not implemented |
| Shape stacking, transforms, and anchors | ✅ | ✅ | ✅ | Shape order, coordinates/extents, transform data, presentation roles, and group nesting are modeled within the supported shape scope |
| Images in slide frames | ✅ | ✅ | ✅ | Embedded or linked image references and package resources are parsed; builder insertion embeds supported payloads and preserves inert external links |
| Audio and video plugins | ✅ | ✅ | ✅ | `draw:plugin` references include MIME type, XLink show/actuate, IDs, and parameters; package-contained media is embeddable/readable, but playback is never attempted |
| Tables | ✅ | ✅ | ✅ | The unified transaction authors common rectangular rich-cell tables; `content::TableModel` inventories and whole-owner edits arbitrary producer tables, retaining formulas, repeats, spans, covered cells, and extension markup in compact source fragments |
| Charts and chart data ranges | 🟡 | ✅ | 🟡 | Chart frames and shared range/series views are bounded; complete typed definitions with series, ranges, axes, local styles, and cached table data can be added or atomically replaced. Checked in-place series and cached-cell edits plus dependency-closed transfer between immutable deck snapshots are supported. Recalculation and rendering remain unavailable |
| Presentation page layouts | ✅ | ✅ | ✅ | Named presentation page layouts and typed placeholder roles/geometry are parsed, validated, added, replaced, reordered, removed, and serialized through public builder/package and unified transaction APIs |
| Slide page metadata and layout references | ✅ | ✅ | ✅ | Page names, IDs, page-layout/master references, and related declaration bindings are inspected and authored through the public page metadata model |
| Master pages | ✅ | ✅ | ✅ | Typed master-page metadata, shared ODF regions/children, lossless XML fragments, package CRUD, ordering, and slide master/layout assignment are exposed through the layered `master` facade and unified transaction |
| Handout master | ✅ | ✅ | ✅ | `handout_master::Master` provides typed XML, shared drawing children, package set/replace/clear, and bounded one-hop presentation-layout resolution |
| Headers, footers, and date-time declarations | ✅ | ✅ | ✅ | Typed declaration collections cover header/footer/date-time values and page bindings; field expansion and host clock identity remain inert |
| Backgrounds and named drawing resources | 🟡 | 🟡 | 🟡 | Direct background/drawing properties and source-level named fill-image, gradient, hatch, marker, opacity, and stroke-dash values are bounded. Cross-deck rich-owner transfer recursively closes and collision-remaps these producer resource families and their package payloads; no renderer is provided |
| Speaker notes | ✅ | ✅ | ✅ | Slide notes are extracted and builder-generated notes stories are serialized; notes are text-oriented rather than a complete notes-layout editor |
| Animations and SMIL timing trees | ✅ | ✅ | ✅ | ODF/SMIL animation nodes, attributes, namespaces, timing structure, and legacy presentation effects are typed, ordered, validated, and serialized; no playback or interpolation is performed |
| Slide transitions and automatic timing | ✅ | ✅ | ✅ | Legacy transition type/style/speed/direction plus SMIL subtype/duration/fade/sound and automatic-advance metadata are typed and written; no visual transition playback occurs |
| Presentation settings and custom slide shows | ✅ | ✅ | ✅ | Animation/transition flags, start/end pages, click behavior, named custom shows, and page references are parsed, validated, and authored |
| Presentation actions and event bindings | ✅ | ✅ | ✅ | URL actions, page-jump targets, sound/media actions, and inert script/event bindings are represented and serialized; actions are never activated |
| Protection and edit restrictions | ❌ | ❌ | ❌ | No public ODP document/slide protection model or policy enforcement is available; XML flags, if present in an unsupported part, remain inert data |

## Explicit gaps from the audited feature families

| Feature family | Status | Read | Write | Notes |
|----------------|--------|------|-------|-------|
| Full rich presentation authoring | ✅ | ✅ | ✅ | The durable root combines common typed authoring with compact source-backed whole-owner read/edit models for arbitrary stories/lists, tables, and inert forms. Cross-deck transfer closes named styles, list/data styles, drawing-resource families, and package payloads with deterministic collision remapping; inverse patches and bounded history preserve exact source bytes |
| Rendering, layout, and playback | ❌ | ❌ | ❌ | No slide renderer, font/layout engine, pagination engine, animation/timing player, media player, chart renderer, or transition compositor is included |
| PPTX sections and zoom objects | ❌ | ❌ | ❌ | `[MS-PPTX]` structures for sections, section zoom, slide zoom, and summary zoom have no ODP typed counterpart here |
| PPTX vendor transitions and design extensions | ❌ | ❌ | ❌ | Morph, newer 2017-2023 transition families, design elements, and other Microsoft PresentationML extension parts are not implemented or converted to ODF |
| PPTX threaded comments, presence, and collaboration commands | ❌ | ❌ | ❌ | Threading/presence metadata, comment authors/replies, revision command monikers, and master/layout/shape change descriptors listed by `[MS-PPTX]` are not supported |
| PPTX media extensions | ❌ | ❌ | ❌ | Media bookmarks, fades, trims, playback event records, narration/presence flags, and media-control extensions are not typed; ODF plugin metadata is not equivalent support |
| PPTX guides and application-specific UI state | ❌ | ❌ | ❌ | Slide/master guides, browse/window state, laser traces, and other application-specific extension structures are not exposed |
| Full chart/table/form object models | 🟡 | ✅ | 🟡 | Charts have typed-definition add/replace plus checked in-place series/cached-cell changes. Tables and inert form pairs combine common typed values with arbitrary compact source-backed read/edit models. Form behavior and formulas are never executed or recalculated, and charts still have no rendering/recalculation engine |
| Macro/script execution and external fetching | ❌ | ❌ | ❌ | No macro, script, hyperlink, external media, DDE, or database source is resolved, fetched, executed, or refreshed |
| Signature writing and security policy | 🟡 | ✅ | ❌ | Public package policy distinguishes editable, signed, encrypted, and signed-plus-encrypted sources. Transactions and durable patch application fail closed before invalidating signatures or encrypted entries; cryptographic verification, signing, encryption writing, and document-protection enforcement remain unavailable |
