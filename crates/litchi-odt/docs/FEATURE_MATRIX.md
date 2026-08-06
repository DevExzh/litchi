# ODT/OTT Feature Matrix

This document tracks the public OpenDocument Text support implemented by Litchi. It covers
packaged text documents and templates (`.odt` and `.ott`) plus flat text XML (`.fodt`). It
describes library data models, parsing, and authoring APIs, not visual rendering fidelity or
complete conformance to every ODF revision.

A raw package entry, cached field value, opaque XML child, or preserved binary object is not
counted as typed semantic support. External resources, formulas, scripts, macros, and links are
never executed or fetched merely because their metadata is present.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the scope described in the Notes cell |
| 🟡 | Bounded, partial, metadata-only, pass-through, or otherwise limited support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to the format or direction |

`Read` and `Write` describe the public direction independently. A `🟡` direction can mean that
only a subset is modeled, the original XML/binary is retained, or a serializer writes inert
metadata rather than performing the behavior implied by the format.

## Package, document lifecycle, and variants

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open, validate, and save ODT bytes or paths | ✅ | ✅ | ✅ | `Document`, `Package`, `Builder`, and `MutableDocument` support in-memory and path workflows with ODF MIME and XML validation. |
| OTT template packages | ✅ | ✅ | ✅ | The package model recognizes the text-template MIME family and preserves the template classification; template content uses the same ODF Text structures as ODT. |
| Package MIME and template classification | ✅ | ✅ | ✅ | `Package::family`, `mimetype`, `extension`, and `is_template` expose validated family metadata. |
| ZIP package, `mimetype`, and manifest | ✅ | ✅ | ✅ | ODF package structure, manifest entries, media types, paths, and required `content.xml` validation are handled by the common package layer. |
| Core XML parts | ✅ | ✅ | ✅ | `content.xml`, optional `styles.xml`, `meta.xml`, and `settings.xml` are parsed as namespace-aware XML and rewritten through package-aware serializers. |
| Unknown and auxiliary package entries | 🟡 | ✅ | ✅ | Arbitrary files can be listed and extracted and safe auxiliary entries are preserved; typed mutation exists only for supported resource families. |
| Embedded media inventory | 🟡 | ✅ | ✅ | Referenced, inline, missing, linked, and package-backed media can be discovered and selected resources can be replaced or removed; this is not a general media codec or renderer. |
| Flat OpenDocument Text XML (`.fodt`) | ✅ | ✅ | ✅ | `FlatDocument` validates a single `office:document`, exposes text-family readers and XML mutation, and saves the exact XML representation without rebuilding a ZIP package. |
| Flat text templates | ❌ | ❌ | ❌ | `FlatDocument` rejects template MIME types; there is no standard flat `.ott` package equivalent in this API. |
| Package-only parts on FODT | N/A | N/A | N/A | Manifest entries, ZIP resources, per-entry encryption, and package signatures are not available in a single flat XML document. |
| Other ODF families through the generic package wrapper | 🟡 | ✅ | 🟡 | The format-neutral `Package` can validate other ODF family MIME types, but this document does not claim the ODT text model for spreadsheets, presentations, drawings, charts, formulas, images, master documents, or databases. |

## Text structure and navigation

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Paragraphs and rich text spans | ✅ | ✅ | ✅ | Typed `text:p`, spans, character data, whitespace, tabs, line breaks, and common inline elements can be extracted and mutated. |
| Headings and outline levels | ✅ | ✅ | ✅ | Typed heading text, level, and style-name access with builder and mutable authoring. |
| Numbered paragraphs | ✅ | ✅ | ✅ | `text:numbered-paragraph` blocks and their numbering attributes are retained as inert numbering metadata; labels are not regenerated. |
| Text blocks and ordering | ✅ | ✅ | ✅ | Paragraph, heading, list, table, and supported framed content can be inserted, replaced, removed, and enumerated in document order. |
| Lists and list items | ✅ | ✅ | ✅ | Ordered and unordered lists, nested items, list headers, style references, and common list authoring are typed. |
| Outline styles and label alignment | ✅ | ✅ | ✅ | Outline declarations and modern list-level label alignment have typed inspection and mutation; computed labels and pagination remain outside the model. |
| Hyperlinks | ✅ | ✅ | ✅ | `text:a` targets, XLink show/actuate values, and link text are typed and writable; targets are inert and never followed. |
| Bookmarks | ✅ | ✅ | ✅ | Point and range bookmark targets have typed parsing and insertion, replacement, and removal with XML identifier checks. |
| Reference marks | ✅ | ✅ | ✅ | Point/range `text:reference-mark` targets support typed CRUD; cross-reference consumers do not resolve them automatically. |
| Ruby annotations | 🟡 | ✅ | 🟡 | Ruby pairs, base/annotation text, named ruby styles, and range-aware wrapping are supported; legal inline structure is bounded and no layout engine renders ruby. |
| Frames and text boxes in text flow | 🟡 | ✅ | ✅ | Anchored frames and text boxes are discovered and selected text-box/image authoring is available; arbitrary nested frame content is not a complete general block model. |
| Header and footer text | ✅ | ✅ | ✅ | Master-page header/footer content and properties can be inspected and changed, including XML-backed content when a typed convenience method is insufficient. |
| Page breaks and inline structural markers | ✅ | ✅ | ✅ | Common page-break and inline marker elements are parsed and serialized as structure; no page calculation is performed. |
| Text extraction as rendered text | 🟡 | ✅ | N/A | Flattened text is useful for content access but does not represent pagination, visual line wrapping, hidden-field evaluation, or every nested rich-content boundary. |

## Styles, page layout, and master pages

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Text and character styles | ✅ | ✅ | ✅ | Named, automatic, and default text styles with typed text properties and XML replacement support. |
| Paragraph styles | ✅ | ✅ | ✅ | Alignment, borders, breaks, drop caps, flow, line spacing, margins, tab stops, writing mode, and related paragraph properties have dedicated models. |
| Table, row, column, and cell styles | ✅ | ✅ | ✅ | Table family declarations and row/column/cell property sets support typed inspection and targeted XML mutation. |
| List and numbering styles | ✅ | ✅ | ✅ | List-style declarations and level properties are parsed and serialized, including bullet/number metadata and outline references. |
| Page layouts | ✅ | ✅ | ✅ | Page usage, dimensions, margins, background, columns, breaks, footnote separators, and additional bounded page-layout attributes are exposed. |
| Page-layout columns | ✅ | ✅ | ✅ | Column counts, widths, gaps, and related properties are typed where modeled. |
| Line-numbering configuration | ✅ | ✅ | ✅ | Document line-numbering policy is inspected and mutated as configuration; line numbers are not generated. |
| Master pages | ✅ | ✅ | ✅ | Master-page declarations and page-layout references support add, insert, replace, remove, and metadata access. |
| Explicit `text:page-sequence` | ✅ | ✅ | ✅ | Master-page assignment sequences are validated and can be inserted, replaced, or removed; this is not a pagination engine. |
| Headers, footers, and their properties | ✅ | ✅ | ✅ | Header/footer region metadata and content/properties are typed for supported regions; cached fields inside them remain inert. |
| Font-face declarations | ✅ | ✅ | ✅ | Declarations in both `content.xml` and `styles.xml` can be inspected, replaced, and cleared; no system font loading or substitution is performed. |
| Data and number styles | 🟡 | ✅ | ✅ | ODF data-style declarations and field references are retained and bounded; values are not formatted by a layout or calculation engine. |
| Drawing style resources | ✅ | ✅ | ✅ | Named gradients, hatches, fill images, markers, opacity definitions, stroke dashes, and related resource metadata have typed adapters. |
| Style registry and inheritance | 🟡 | ✅ | 🟡 | Style declarations and registry relationships are available, but complete cascade resolution, style-use normalization, and producer-specific fallback behavior are not guaranteed. |
| Pagination, line layout, and visual rendering | ❌ | ❌ | ❌ | No page-layout engine, font shaping engine, field refresh layout, drawing renderer, or pixel-level fidelity claim is made. |

## Sections, fields, variables, and generated content

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Text sections | ✅ | ✅ | ✅ | Sections support add, wrap, unwrap, replace, remove, and clear operations with names, protection, visibility conditions, linked-source metadata, and inert DDE metadata. |
| Section properties and backgrounds | ✅ | ✅ | ✅ | Section style properties, writing mode, background color/image, repeat behavior, and bounded lengths are validated and serialized. |
| Section protection policy | 🟡 | ✅ | ✅ | Protection keys and flags are retained as metadata; the library does not enforce edit protection or unlock protected sections. |
| Inline field vocabulary | ✅ | ✅ | ✅ | Typed ODF field models cover page/date/time, file/template/sheet/chapter, user and document metadata, sequence/reference, placeholder, conditional/hidden text, expression, variable, drop-down, measure, note, statistic, script, and meta-field families. |
| Dynamic field CRUD | ✅ | ✅ | ✅ | Fields can be parsed and inserted, replaced, or removed in document order with validated cached display values and field attributes. |
| Field cached values and instructions | 🟡 | ✅ | ✅ | Instructions, conditions, display switches, source names, and producer-cached text are modeled or preserved; a cached result is not evidence that the field was evaluated by Litchi. |
| Database fields and sources | 🟡 | ✅ | ✅ | Database field kinds, table/query/command source metadata, columns, conditions, row values, connection-resource references, and cached text have typed inert CRUD. |
| DDE connections | 🟡 | ✅ | ✅ | Connection declarations and uses can be retained or changed as metadata; no DDE conversation, refresh, or external process is started. |
| Field evaluation and refresh | ❌ | ❌ | ❌ | No page, date, user identity, sequence, expression, condition, database, DDE, sheet-state, formula, or document-statistic field is recalculated or refreshed. |
| Variable declarations and user fields | ✅ | ✅ | ✅ | Variable declaration groups and user-field metadata have typed ordered inspection and mutation; values remain inert. |
| Content validation rules | 🟡 | ✅ | ✅ | Validation conditions, messages, and ranges are represented as bounded XML metadata; there is no form/UI enforcement engine. |
| Tables of contents and indexes | 🟡 | ✅ | ✅ | TOC, user, alphabetical, illustration, table, object, and bibliography index sources plus cached bodies/templates are structurally parsed and writable; automatic generation, page numbers, and refresh are not implemented. |
| Index marks | ✅ | ✅ | ✅ | TOC, user, alphabetical, and bibliography marks support typed ordered CRUD, including range metadata where present. |
| Bibliography configuration and records | 🟡 | ✅ | ✅ | Bibliography policy, sort keys, source marks, and inert records are typed; citation lookup and automatic entry generation are not performed. |
| Concordance/auto-mark metadata | 🟡 | ✅ | ✅ | Auto-mark file declarations are retained as typed metadata; no file-driven index regeneration is performed. |

## Tables and tabular behavior

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Tables, rows, columns, and cells | ✅ | ✅ | ✅ | Nested tables and typed table element models support enumeration, construction, insertion, replacement, and removal. |
| Merged cells | ✅ | ✅ | ✅ | Column and row spans, including ODF `table:number-columns-spanned` and `table:number-rows-spanned`, are represented in table access and serialization. |
| Repeated rows and columns | 🟡 | ✅ | ✅ | ODF repeated structures can be expanded for semantic access and written within bounded table mutation; the original repeat strategy is not a layout calculation. |
| Cell values | ✅ | ✅ | ✅ | Empty, text, number, boolean, date, currency, percentage, and time cell values have typed vocabulary. |
| Table, row, column, and cell properties | ✅ | ✅ | ✅ | Widths, borders, backgrounds, alignment, padding, break behavior, and modeled style properties support typed or targeted XML mutation. |
| Table formulas and cached results | 🟡 | ✅ | 🟡 | Formula attributes and cached values can remain in the document, but the ODT crate does not evaluate table formulas or recalculate dependent cells. |
| Database ranges and external table sources | 🟡 | ✅ | ✅ | Source/range metadata is preserved where modeled; no database connection, query, import, or refresh is performed. |
| Automatic table layout and pagination | ❌ | ❌ | ❌ | Column sizing, row splitting, repeated heading layout, and page placement are not rendered or computed. |

## Notes, annotations, and tracked changes

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Footnotes and endnotes | 🟡 | ✅ | ✅ | Typed note identity, class, citation, label, plain text, and validated rich note bodies support insertion, replacement, and removal; nested fields, links, scripts, and macro metadata remain inert. |
| Note configuration | ✅ | ✅ | ✅ | Footnote/endnote numbering, position, prefix/suffix, restart, and separator configuration are inspected and mutated. |
| Footnote separators | ✅ | ✅ | ✅ | Separator line properties and related page-layout metadata have typed access. |
| Standard ODF annotations/comments | ✅ | ✅ | ✅ | Point and range anchors, authors, initials, dates, display state, body paragraphs, names, and namespaced body elements support package-aware CRUD. |
| Rich annotation bodies | 🟡 | ✅ | ✅ | Common paragraph and arbitrary bounded namespaced child content is retained; comments are not rendered or threaded by a collaboration service. |
| Threaded comments, replies, reactions, and people graphs | ❌ | ❌ | ❌ | The ODT model is classic ODF annotation-oriented and does not implement Microsoft modern-comment conversation semantics. |
| Tracked-change declarations | ✅ | ✅ | ✅ | Insertion, deletion, and format-change declarations expose IDs, authors, dates, comments, style references, content, and container policy. |
| Tracked-change marking and XML CRUD | ✅ | ✅ | ✅ | Change ranges and deletions can be marked, updated, unmarked, removed, or cleared with protection-key metadata retained but never used as an unlock mechanism. |
| Accept/reject and revision merge engine | ❌ | ❌ | ❌ | There is no general semantic accept/reject, conflict merge, style resolution, or review-state recalculation engine. |

## Forms, controls, and events

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| ODF `office:forms` containers | ✅ | ✅ | ✅ | Nested forms, names, properties, controls, shape references, event metadata, and order are typed. |
| Text, textarea, button, checkbox, radio, list, and combo controls | ✅ | ✅ | ✅ | Broad classic control families have typed inspection and insertion/replacement/removal APIs, including labels, IDs, current state, and options. |
| Number, date, time, password/file, image-frame, and generic controls | ✅ | ✅ | ✅ | Specialized control models cover typed values, visual attributes, file/password metadata, image frames, and generic fallback controls. |
| Value-range, typed-value, selection, and grid controls | ✅ | ✅ | ✅ | Control-specific properties and nested options/columns have typed authoring and mutation. |
| Form property values | ✅ | ✅ | ✅ | Boolean, number, text, date, time, list, and void property values are validated and serialized. |
| Form events, bindings, and external data sources | 🟡 | ✅ | ✅ | Event listeners, XLink targets, bindings, and connection-resource metadata are retained or mutated as inert declarations; no event or datasource is run. |
| Form validation and UI behavior | ❌ | ❌ | ❌ | The library does not display controls, enforce validation interactively, execute listeners, submit forms, or maintain live widget state. |
| OOXML content controls and ActiveX controls | ❌ | 🟡 | ❌ | ODF classic forms are not `w:sdt`, ActiveX, or other Microsoft control models; an opaque embedded payload may be preserved but is not typed. |

## Drawings, images, charts, and embedded objects

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| ODF image frames | ✅ | ✅ | ✅ | Image references, anchors, dimensions, and selected frame metadata are discoverable; image-frame authoring accepts sniffed PNG, JPEG, and GIF payloads with verbatim package storage. |
| Image replacement and removal | ✅ | ✅ | ✅ | Referenced package media can be replaced or removed with resource/path validation; no image editing or rendering pipeline is implied. |
| Linked and missing images | 🟡 | ✅ | ✅ | Linked targets and missing package parts are reported as inert source states; Litchi does not fetch or repair them. |
| Client-side image maps | 🟡 | ✅ | N/A | Rectangle, circle, and polygon `draw:image-map` areas and link metadata are read as typed/inert content; no general image-map authoring API is exposed. |
| Text boxes and frame anchors | ✅ | ✅ | ✅ | Text-box insertion and frame anchor metadata are supported for the modeled forms. |
| ODF geometric shapes and connectors | 🟡 | 🟡 | 🟡 | Existing drawing XML and named resources can be preserved and bounded properties are available, but there is no complete typed CRUD or renderer for every rectangle, ellipse, path, connector, custom shape, enhanced geometry, or group variant. |
| Drawing resources | ✅ | ✅ | ✅ | Gradients, hatches, fill images, markers, opacity, stroke dashes, and page drawing properties have typed package and flat-document adapters. |
| Embedded ODF charts | ✅ | ✅ | ✅ | Embedded chart subdocuments and inline chart content support discovery, add/replace/remove, cached tables, series, axes, labels, and modeled chart styles through the shared `litchi-odf-common::chart::authoring` content owner; package topology and host mutation remain ODT-owned. |
| Chart calculation and visual rendering | 🟡 | ✅ | ❌ | Chart XML and cached values are modeled, but formulas, live data sources, layout, and rendering are not performed by the ODT crate. |
| Embedded objects and subdocuments | 🟡 | ✅ | ✅ | `draw:object`, `draw:object-ole`, applet, plugin, floating-frame, OpenDocument, MathML, package-file, and linked sources are discoverable as typed source/kind metadata with bounded resource CRUD. |
| Embedded MathML | 🟡 | ✅ | 🟡 | MathML roots can be identified or carried as embedded content; the ODT object layer does not provide a complete MathML semantic evaluator or layout engine. |
| Embedded object activation or conversion | ❌ | ❌ | ❌ | Objects are never loaded into host applications, activated, converted, recalculated, rendered, or executed. |
| External drawing/object links | 🟡 | ✅ | ✅ | Hrefs and relationship-like targets are retained as inert metadata; no network, filesystem, OLE, or plugin activation occurs. |

## Metadata, settings, and package graphs

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Dublin Core and ODF document metadata | ✅ | ✅ | ✅ | Title, creator, subject, description, dates, language, keywords, editing information, generator, and related `meta.xml` fields are typed through the common metadata model. |
| User-defined metadata | ✅ | ✅ | ✅ | Named user-defined values support string, boolean, date, time, float, and related ODF value types with validation. |
| Document statistics | 🟡 | ✅ | ✅ | Stored paragraph, word, character, table, image, object, page, and editing statistics are read/written as metadata; they are not recomputed from content. |
| Template, auto-reload, and hyperlink-behavior metadata | 🟡 | ✅ | ✅ | ODF metadata/configuration is represented, but attached templates, reload targets, and link policies are not followed or applied. |
| RDF metadata graphs | ✅ | ✅ | ✅ | RDF package graph and triple CRUD are available where the package carries ODF RDF metadata. |
| Settings tree | ✅ | ✅ | ✅ | `office:settings`, config sets, maps, items, scalar values, and flat/package adapters support structural inspection and mutation. |
| Settings-driven application behavior | ❌ | 🟡 | ❌ | View, update, compatibility, spell-check, printer, cursor, and other configuration values may be preserved, but no office UI/runtime consumes them. |
| Document protection policy metadata | 🟡 | ✅ | ✅ | Typed form, bookmark, read-only, and tracked-change key metadata is read and transactionally rewritten while unknown settings XML remains opaque; no policy enforcement or unlock behavior is provided. |
| Protection and visibility settings | 🟡 | ✅ | ✅ | Document/section protection flags, hidden conditions, and related policy metadata are retained; they are not enforced or evaluated. |
| Package graph mutation and signature invalidation | ✅ | ✅ | ✅ | Known resource edits update manifest/package relationships through transactional writers and do not pretend that a prior signature remains valid after mutation. |

## Scripts, macros, and external behavior

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Document script declarations | 🟡 | ✅ | ✅ | Embedded script metadata, language/name/source declarations, libraries, and document event listeners have typed bounded CRUD. |
| Script package resources | 🟡 | ✅ | ✅ | Script files/resources can be listed, added, replaced, moved, or removed as inert package entries with manifest metadata. |
| Macro and event execution | ❌ | ❌ | ❌ | Scripts, macros, listeners, form actions, DDE, database connections, and external commands are never executed. |
| Mail merge and database execution | ❌ | 🟡 | ❌ | Recipient/source/field metadata may be parsed or written, but data sources are never opened and no merge is generated. |
| External link refresh | ❌ | 🟡 | ❌ | URLs, file targets, and update policies are retained as metadata only; no target is fetched, opened, or refreshed. |
| VBA/OVBA projects and ActiveX | ❌ | 🟡 | ❌ | An opaque package payload may be exposed or preserved, but there is no VBA project/module parser, ActiveX model, signature execution, or host automation. |

## Encryption and digital signatures

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| ODF per-entry password encryption | ✅ | ✅ | ✅ | `OwnedPackage` opens encrypted entries with a password and `PackageWriter` can encrypt payload entries using validated ODF profiles. |
| ODF encryption profiles | ✅ | ✅ | ✅ | The common layer models AES CBC/GCM variants, Blowfish CFB8, SHA-1/SHA-256 start keys, PBKDF2, and Argon2id where the profile validator permits the combination; unsupported combinations fail explicitly. |
| Encrypted package mutation | 🟡 | ✅ | 🟡 | Password opening and new encrypted package authoring are supported; high-level content rewrites of an existing encrypted package can be rejected rather than silently producing plaintext or invalid encryption. |
| ODF document signatures | ✅ | ✅ | ✅ | Package document signatures can be parsed, cryptographically verified, and generated with RSA-SHA256 or ECDSA P-256/SHA256 through the common signature layer. |
| Macro-signature metadata | 🟡 | ✅ | 🟡 | ODF macro-signature containers can be inspected as signature metadata, but no macro code is executed and signing does not establish macro safety. |
| Certificate trust, revocation, and identity policy | ❌ | 🟡 | ❌ | Verification is integrity/signature verification only; certificate-chain trust, revocation, identity policy, and platform trust stores are outside the model. |
| Flat-document encryption and signatures | N/A | N/A | N/A | The FODT wrapper has no ODF ZIP manifest entry channel for per-file encryption or `META-INF` package signature containers. |

## Explicit Microsoft-format gaps

The local Microsoft Open Specifications material was used to keep these rows explicit. The
`[MS-*]` documents describe Microsoft-specific binary/package structures, not ODF feature
requirements. An ODF embedded resource that happens to contain one of these payloads is still
opaque data.

| Feature family | Status | Read | Write | Notes |
|---------------|--------|------|-------|-------|
| Compound File Binary storage and streams (`[MS-CFB]`) | ❌ | 🟡 | 🟡 | No typed CFB header, sector, FAT, directory, storage, stream, mini-stream, or DIFAT model exists in ODT; raw embedded bytes may remain accessible as a package resource. |
| OLE1/OLE2 embedded and linked objects (`[MS-OLEDS]`) | ❌ | 🟡 | 🟡 | ODF object/source metadata is not an OLE object server, link manager, presentation cache, or OLE1/OLE2 parser. |
| OLE property sets (`[MS-OLEPS]`) | ❌ | ❌ | ❌ | ODF `meta.xml` and user-defined metadata are separate from SummaryInformation, DocumentSummaryInformation, property-set streams, code pages, and OLE property identifiers. |
| Microsoft Office encryption and data spaces (`[MS-OFFCRYPTO]`) | ❌ | ❌ | ❌ | ODF manifest encryption is not Standard/Agile Office encryption, data spaces, encrypted package streams, write protection, or binary-document crypto. |
| Microsoft Office binary signatures (`[MS-OFFCRYPTO]`) | ❌ | ❌ | ❌ | ODF XML package signatures are distinct from the CFB/binary document signature structures described by the Microsoft specification. |
| VBA project streams and modules (`[MS-OVBA]`, `[MS-OSHARED]`) | ❌ | 🟡 | ❌ | No dir/PROJECT/PROJECTwm/PROJECTlk/module stream parser, VBA source model, reference resolver, or VBA execution is provided. |
| OfficeArt shapes and records (`[MS-ODRAW]`) | ❌ | 🟡 | ❌ | ODF `draw:*` frames and resource models do not implement OfficeArt containers, property tables, shape paths, signature lines, or OfficeArt algorithms. |
| OOXML DrawingML and extensions (`[MS-ODRAWXML]`) | ❌ | 🟡 | ❌ | No `a:*`, WordprocessingDrawing, spreadsheetDrawing, diagram, ink, SVG-extension, or Markup Compatibility model is claimed for ODT. |
| Excel binary chart/graphics structures (`[MS-OGRAPH]`) | ❌ | 🟡 | ❌ | ODF chart XML is not the OGRAPH compound-file/chart-record vocabulary; an embedded binary chart is opaque. |
| EMF, EMF+, and WMF metafiles (`[MS-EMF]`, `[MS-EMFPLUS]`, `[MS-WMF]`) | ❌ | 🟡 | 🟡 | These binary drawing payloads are not parsed, authored, rasterized, or rendered by the ODT model; a package may preserve an opaque media entry. |
| OOXML-specific controls, modern comments, custom XML, and alternate content | ❌ | 🟡 | ❌ | ODT has separate ODF forms, annotations, RDF, and XML vocabularies; no `w:sdt`, modern-comment, custom-XML-store, `mc:AlternateContent`, or `altChunk` semantics are implemented here. |

## Remaining semantic non-goals

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Full ODF version/conformance validation | 🟡 | 🟡 | 🟡 | The implementation targets modeled ISO/IEC 26300/OpenDocument structures and preserves supported extensions, but it is not a complete schema, profile, producer-compatibility, or version-conformance validator. |
| Office-style rendering fidelity | ❌ | ❌ | ❌ | No visual comparison, pagination, font shaping, print layout, accessibility rendering, or application UI behavior is part of this crate. |
| Recalculation of generated content | ❌ | ❌ | ❌ | Fields, indexes, table/chart formulas, database sources, DDE, scripts, and external links are not refreshed as a side effect of reading or writing. |

## Implementation and specification boundary

The typed ODT surface is implemented in `crates/litchi-odt/src`, with package, metadata,
encryption, signatures, RDF, embedded-resource, annotation, media, drawing, and style
vocabulary shared through `crates/litchi-odf-common/src`. The Microsoft boundary rows correspond
to the local `[MS-CFB]`, `[MS-OLEDS]`, `[MS-OLEPS]`, `[MS-OFFCRYPTO]`, `[MS-OVBA]`, `[MS-OSHARED]`,
`[MS-ODRAW]`, `[MS-ODRAWXML]`, `[MS-OGRAPH]`, `[MS-EMF]`, `[MS-EMFPLUS]`, and `[MS-WMF]` ToC and
Front Matter scopes. No Microsoft protocol feature is promoted to ODT support solely because an
ODF package can carry an opaque payload with a related name or media type.
