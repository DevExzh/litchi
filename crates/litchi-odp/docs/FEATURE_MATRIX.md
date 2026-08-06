# ODP/OTP Feature Matrix

This document tracks the public and source-level feature coverage of the
`litchi-odp` OpenDocument Presentation implementation for packaged ODP and OTP
files. It is a capability matrix, not a claim of complete ODF conformance,
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
RDF graph editing, and password opening. The source tree also contains richer
mutable-presentation and master-page helpers, but those modules are not
connected to the current `litchi_odp` module exports; their source-level
vocabulary is therefore marked bounded rather than presented as public CRUD.

The Microsoft `[MS-PPTX]` Front Matter and ToC describe extensions to OOXML
PresentationML, not ODP. Their Part Enumerations, Extensions, structure
families, examples, and security sections were used as a gap checklist only;
the existence of a PPTX feature is never treated as ODP support.

## Package and shared ODF features

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open, create, and save ODP/OTP packages | ✅ | ✅ | ✅ | `Presentation` opens paths or bytes and saves/to-bytes; `Builder` creates presentations and templates with presentation MIME handling |
| Password-protected ODP opening | ✅ | ✅ | N/A | `open_with_password` and `from_bytes_with_password` use the shared ODF encrypted-package reader; this does not imply password/encryption authoring |
| ZIP package, manifest, and MIME validation | ✅ | ✅ | ✅ | Shared ODF package validation and deterministic package writing cover manifest/resource access and presentation/template media types |
| `content.xml`, `styles.xml`, `meta.xml`, and `settings.xml` | ✅ | ✅ | ✅ | The public presentation/package and builder paths parse or write the core package XML parts; unsupported children may remain inert/preserved package data |
| ODF metadata and statistics | ✅ | ✅ | ✅ | `Presentation::metadata` and `Builder::set_metadata` use the shared typed metadata model, including Dublin Core, user-defined values, template/reload/link metadata, and statistics |
| Common styles and data styles | 🟡 | 🟡 | 🟡 | Direct presentation/drawing properties and bounded shared style values are supported; there is no complete public style-graph resolver/editor for every ODF style family |
| Images and package media | ✅ | ✅ | ✅ | Package-contained and linked image/media references are typed; package bytes are read without fetching external URLs, and builder media embedding creates package resources |
| Embedded objects and OLE-like payloads | 🟡 | 🟡 | 🟡 | Frames/object references can remain as bounded opaque content/resources; payloads are not opened, activated, executed, or rendered as native documents |
| Embedded charts | 🟡 | 🟡 | 🟡 | Shared ODF chart views can inspect chart XML vocabulary, but ODP exposes no complete typed chart-part model or chart data/editor; chart frames are bounded/opaque |
| Annotations/comments | ✅ | ✅ | ✅ | `annotation::{Anchor, Info, Position}` inventories shared rich ODF annotations at validated pages or uniquely named shapes; `Presentation` provides atomic add/replace/remove while untouched XML and no-op bytes remain preserved |
| Hyperlinks and external references | ✅ | ✅ | ✅ | Shape links, XLink targets, show/actuate values, page jumps, and action metadata are typed and serialized; targets are never opened or followed |
| Forms and controls | ❌ | ❌ | ❌ | No public ODP form/control model or authoring surface is exposed; a control-shaped XML payload is not treated as typed form support |
| Scripts, events, and macros | 🟡 | 🟡 | 🟡 | Event/action and script-binding metadata can be represented as inert values; no script, macro, or event execution occurs |
| RDF metadata graphs | ✅ | ✅ | ✅ | Graph and triple inventory plus ordered add/replace/remove/move operations are exposed on `Presentation` |
| ODF encryption authoring | ❌ | ✅ | ❌ | Password opening is supported, but the public ODP builder has no encrypt/password-change operation; common encryption code alone is not end-to-end write support |
| ODF digital signatures | ❌ | ❌ | ❌ | Shared XMLDSig models do not become ODP sign/verify/add/clear APIs in the current public crate |
| Unknown package-part preservation | 🟡 | ✅ | 🟡 | The owned package can retain unrelated resources around supported edits; preservation is not semantic understanding or guaranteed lossless mutation of every extension |

## Presentation content and layout

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Slides, ordering, and slide count | ✅ | ✅ | ✅ | `Presentation` parses slide order/count; `Builder` adds slides and slide elements with validated names/IDs; public mutable deletion/reordering helpers are not exported |
| Slide titles and body text | ✅ | ✅ | ✅ | Plain title/body text is extracted and authored; `Slide::all_text` includes labeled shape text |
| Rich text runs, paragraphs, and lists | 🟡 | 🟡 | 🟡 | Shape text and ODF whitespace/spans have bounded parser support, but the public builder primarily authors string stories rather than a complete rich paragraph/list tree |
| Text boxes and basic shapes | ✅ | ✅ | ✅ | Rectangles, ellipses, lines, connectors, groups, frames/text boxes, and their core geometry/properties are represented by `Shape` and emitted by the builder |
| Advanced ODF drawing shapes | 🟡 | 🟡 | 🟡 | Polylines, polygons, paths, captions, measures, custom shapes, and 3D scene/object kinds are recognized or retained with inert/unmodeled geometry attributes; complete rendering semantics are not implemented |
| Shape stacking, transforms, and anchors | ✅ | ✅ | ✅ | Shape order, coordinates/extents, transform data, presentation roles, and group nesting are modeled within the supported shape scope |
| Images in slide frames | ✅ | ✅ | ✅ | Embedded or linked image references and package resources are parsed; builder insertion embeds supported payloads and preserves inert external links |
| Audio and video plugins | ✅ | ✅ | ✅ | `draw:plugin` references include MIME type, XLink show/actuate, IDs, and parameters; package-contained media is embeddable/readable, but playback is never attempted |
| Tables | 🟡 | 🟡 | 🟡 | Table-shaped frames or opaque table XML can be retained within the bounded drawing model; there is no public typed cell/row/table editor |
| Charts and chart data ranges | 🟡 | 🟡 | 🟡 | Chart frames and shared chart range/series views are bounded metadata; no public ODP chart authoring, recalculation, or rendering engine is available |
| Presentation page layouts | ✅ | ✅ | ✅ | Named presentation page layouts and typed placeholder roles/geometry are parsed, validated, added, replaced, and serialized through public builder/package APIs |
| Slide page metadata and layout references | ✅ | ✅ | ✅ | Page names, IDs, page-layout/master references, and related declaration bindings are inspected and authored through the public page metadata model |
| Master pages | ✅ | ✅ | ✅ | Typed master-page metadata, shared ODF regions/children, lossless XML fragments, package CRUD, ordering, and slide master/layout assignment are exposed through the layered `master` facade |
| Handout master | ✅ | ✅ | ✅ | `handout_master::Master` provides typed XML, shared drawing children, package set/replace/clear, and bounded one-hop presentation-layout resolution |
| Headers, footers, and date-time declarations | ✅ | ✅ | ✅ | Typed declaration collections cover header/footer/date-time values and page bindings; field expansion and host clock identity remain inert |
| Backgrounds and named drawing resources | 🟡 | 🟡 | 🟡 | Direct background/drawing properties and source-level named fill-image, gradient, hatch, marker, opacity, and stroke-dash values are bounded; no complete public resolver or renderer is provided |
| Speaker notes | ✅ | ✅ | ✅ | Slide notes are extracted and builder-generated notes stories are serialized; notes are text-oriented rather than a complete notes-layout editor |
| Animations and SMIL timing trees | ✅ | ✅ | ✅ | ODF/SMIL animation nodes, attributes, namespaces, timing structure, and legacy presentation effects are typed, ordered, validated, and serialized; no playback or interpolation is performed |
| Slide transitions and automatic timing | ✅ | ✅ | ✅ | Legacy transition type/style/speed/direction plus SMIL subtype/duration/fade/sound and automatic-advance metadata are typed and written; no visual transition playback occurs |
| Presentation settings and custom slide shows | ✅ | ✅ | ✅ | Animation/transition flags, start/end pages, click behavior, named custom shows, and page references are parsed, validated, and authored |
| Presentation actions and event bindings | ✅ | ✅ | ✅ | URL actions, page-jump targets, sound/media actions, and inert script/event bindings are represented and serialized; actions are never activated |
| Protection and edit restrictions | ❌ | ❌ | ❌ | No public ODP document/slide protection model or policy enforcement is available; XML flags, if present in an unsupported part, remain inert data |

## Explicit gaps from the audited feature families

| Feature family | Status | Read | Write | Notes |
|----------------|--------|------|-------|-------|
| Full rich presentation authoring | 🟡 | 🟡 | 🟡 | Text stories, shapes, notes, layouts, media, settings, and animations have public bounded paths, but complete tables, charts, master pages, forms, style resolution, and arbitrary story editing are not exposed |
| Rendering, layout, and playback | ❌ | ❌ | ❌ | No slide renderer, font/layout engine, pagination engine, animation/timing player, media player, chart renderer, or transition compositor is included |
| PPTX sections and zoom objects | ❌ | ❌ | ❌ | `[MS-PPTX]` structures for sections, section zoom, slide zoom, and summary zoom have no ODP typed counterpart here |
| PPTX vendor transitions and design extensions | ❌ | ❌ | ❌ | Morph, newer 2017-2023 transition families, design elements, and other Microsoft PresentationML extension parts are not implemented or converted to ODF |
| PPTX threaded comments, presence, and collaboration commands | ❌ | ❌ | ❌ | Threading/presence metadata, comment authors/replies, revision command monikers, and master/layout/shape change descriptors listed by `[MS-PPTX]` are not supported |
| PPTX media extensions | ❌ | ❌ | ❌ | Media bookmarks, fades, trims, playback event records, narration/presence flags, and media-control extensions are not typed; ODF plugin metadata is not equivalent support |
| PPTX guides and application-specific UI state | ❌ | ❌ | ❌ | Slide/master guides, browse/window state, laser traces, and other application-specific extension structures are not exposed |
| Full chart/table/form object models | ❌ | 🟡 | ❌ | Opaque frames and bounded common XML views must not be mistaken for typed chart series CRUD, table-cell CRUD, form controls, or recalculated embedded objects |
| Macro/script execution and external fetching | ❌ | ❌ | ❌ | No macro, script, hyperlink, external media, DDE, or database source is resolved, fetched, executed, or refreshed |
| Signature writing and security policy | ❌ | ❌ | ❌ | Digital-signature authoring/verification and document-protection enforcement are not public ODP capabilities |
