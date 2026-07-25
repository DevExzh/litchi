# Office File Format Feature Matrix

This document tracks the public feature families implemented by Litchi.

The matrix describes library support, not rendering fidelity or complete conformance with every
revision of a file-format specification. A row can be supported while intentionally treating
external links, scripts, database commands, mail-merge sources, or embedded payloads as inert data.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the feature scope described in the Notes cell |
| 🟡 | Bounded, partial, metadata-only, pass-through, or otherwise limited support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to the format or direction |

`Read` and `Write` describe the public direction independently. A 🟡 direction usually means a
subset of the model, lossless preservation, or an inert serializer rather than full semantic CRUD.
Cryptographic verification means integrity/signature verification only; it does not establish
certificate trust or revocation status.

## Supported format families

- **Microsoft Office OOXML:** DOCX, XLSX, XLSB, PPTX
- **Microsoft Office binary/OLE:** DOC, XLS, PPT
- **OpenDocument packages and templates:** ODT/OTT, ODS/OTS, ODP/OTP, ODG/OTG, ODC/OTC,
  ODF/OTF, ODI/OTI, ODM/OTM, OTH, and ODB
- **Flat OpenDocument XML:** FODT, FODS, FODP, FODG, FODC, FODI, and the extended FODF convention
- **Rich Text Format:** RTF, including compressed RTF payloads
- **Apple iWork:** Pages, Keynote, and Numbers archives
- **Tabular interchange APIs:** CSV, TSV, configurable delimited text, PRN/fixed-width, SYLK,
  and DIF

Feature-gated families require the corresponding Cargo feature. The default umbrella build enables
legacy Office, OOXML, OOXML encryption, and the formula evaluator; ODF, RTF, iWork, formula
conversion, fonts, and image conversion are optional.

## Cross-format capabilities

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Smart format detection | ✅ | ✅ | N/A | Detects the listed Office, ODF, RTF, and iWork families from content, not only extensions |
| Unified document facade | 🟡 | ✅ | ❌ | Common text/metadata access for DOC, DOCX, RTF, Pages, and ODT; authoring remains format-specific |
| Unified presentation facade | 🟡 | ✅ | ❌ | Common access for PPT, PPTX, Keynote, and ODP; authoring remains format-specific |
| Unified workbook facade | ✅ | ✅ | ❌ | Common sheet names/count, text, and metadata for XLS, XLSX, XLSB, ODS, and Numbers |
| Workbook trait API | 🟡 | ✅ | ❌ | Implemented by XLS, XLSX, XLSB, text workbooks, and immutable ODS evaluation snapshots; Numbers uses the unified facade |
| OOXML OPC package editing | ✅ | ✅ | ✅ | Parts, relationships, content types, strict/transitional XML, and transactional graph updates |
| OLE/CFB package editing | ✅ | ✅ | ✅ | Streams, storages, property sets, and package-preserving editors |
| ODF package editing | ✅ | ✅ | ✅ | ZIP package, manifest, metadata, styles, settings, resources, and MIME validation |
| OOXML encryption | ✅ | ✅ | ✅ | Standard 2007 and Agile encryption; requires `ooxml_encryption` |
| Legacy Office encryption | ✅ | ✅ | ✅ | Format-specific DOC, XLS, and PPT password profiles |
| ODF encryption | ✅ | ✅ | ✅ | Package authoring/opening with supported AES/Blowfish, PBKDF2, and Argon2 profiles |
| OOXML digital signatures | ✅ | ✅ | ✅ | Verify, add, re-sign, and clear RSA-SHA256/ECDSA package signatures |
| Legacy Office digital signatures | ✅ | ✅ | ✅ | Verify, add, re-sign, and clear signatures in CFB packages |
| ODF digital signatures | ✅ | ✅ | ✅ | Sign and verify package documents with RSA or ECDSA |
| Core/extended/custom properties | ✅ | ✅ | ✅ | OOXML properties and OLE property-set editing; ODF metadata has its own model |
| Spreadsheet formula evaluation | 🟡 | ✅ | N/A | Shared async evaluator for workbook-trait adapters; broad function set but not complete Excel semantics |
| Equation parsing and conversion | 🟡 | ✅ | N/A | OMML/MTEF-to-LaTeX conversion plus semantic ODF MathML parsing; reverse conversion and layout are not implemented |
| Markdown export | 🟡 | ✅ | ✅ | Document/presentation conversion with optional parallel processing; fidelity varies by source format |
| Image conversion | 🟡 | ✅ | ✅ | Feature-gated EMF, WMF, and PICT conversion to common raster outputs |
| Font discovery/embedding helpers | 🟡 | ✅ | ✅ | Feature-gated system-font lookup plus format-specific embedded-font models |

## Word documents (DOCX)

### Structure and formatting

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open/create/save | ✅ | ✅ | ✅ | Path and in-memory package workflows |
| Text, paragraphs, and runs | ✅ | ✅ | ✅ | CRUD with character and paragraph formatting |
| Tables and cells | ✅ | ✅ | ✅ | Rows, cells, merges, borders, widths, and properties |
| Sections and page setup | ✅ | ✅ | ✅ | Margins, paper, orientation, columns, borders, numbering, and line numbering |
| Styles | ✅ | ✅ | ✅ | Paragraph, character, table, and numbering style models |
| Numbering and lists | ✅ | ✅ | ✅ | Abstract numbering, instances, overrides, and list formatting |
| Headers and footers | ✅ | ✅ | ✅ | Default, first-page, and odd/even stories |
| Footnotes and endnotes | ✅ | ✅ | ✅ | Notes, separators, continuations, and references |
| Hyperlinks | ✅ | ✅ | ✅ | Internal and external links |
| Bookmarks | ✅ | ✅ | ✅ | Range and point bookmarks |
| Fields | ✅ | ✅ | ✅ | Typed field delimiters/instructions plus inert `MACROBUTTON`, `ADDIN`/`CONTROL`/`HTMLCONTROL` kind/instruction/cached-result metadata, `GLOSSARY`/`AUTOTEXT` entry-name/unknown-switch/cached-result metadata, `AUTOTEXTLIST` display/style/tip/unknown-switch/cached-result metadata, `GOTOBUTTON` destination/button metadata, `USERADDRESS`/`USERINITIALS`/`USERNAME` kind/override/formatting metadata, `ADVANCE` point-adjustment metadata, `DDE`/`DDEAUTO`, `LINK`, `RD`, `INCLUDETEXT`/`INCLUDEPICTURE` source/option metadata, `DOCPROPERTY` property-name/switch/cached-result metadata, `TITLE`/`SUBJECT`/`AUTHOR`/`KEYWORDS`/`COMMENTS`/`LASTSAVEDBY` kind/switch/cached-result metadata, `DOCVARIABLE`/`MERGEFIELD` name/switch/cached-result metadata, `MERGEREC`/`MERGESEQ` kind/cached-result, `NEXT` cached-result/state, `NEXTIF`/`SKIPIF` kind/unparsed-comparison/cached-result, `COMPARE` unparsed-comparison/cached-result, `IF` unparsed-expression/cached-result, `SET` target/opaque-expression/cached-result metadata, `SEQ` identifier/bookmark/opaque-tail/cached-result metadata, `=` formula/cached-result metadata, `STYLEREF` style/options/unknown-switch/cached-result metadata, `ASK`/`FILLIN` prompt/default-response metadata, and `ADDRESSBLOCK`/`GREETINGLINE` recipient-layout/country/locale/fallback metadata; fields are not recalculated, macros are never resolved or executed, add-ins and controls are never loaded, instantiated, rendered, or executed, building-block fields never look up stored entries, read templates, show a selection UI, insert content, change bookmarks, or refresh, navigation fields never resolve or activate a destination, user-identity fields never read or modify host identity data, `ADVANCE` fields never move text, change layout, or reflow content, `COMPARE` fields never parse or evaluate a comparison, `SET` fields never evaluate expressions, look up or change bookmarks, change document state, or refresh, `SEQ` fields never look up bookmarks, increment or reset sequences, calculate numbers, or refresh, formula fields never parse or evaluate formulas, read table cells or bookmarks, resolve field values, or refresh, `STYLEREF` fields never look up styled text, search document stories, calculate paragraph numbers or relative positions, resolve page layout, or refresh, DDE never starts a conversation, `LINK` never activates OLE, prompt fields never display a dialog or capture a response, document properties and built-in document-information metadata are never read or resolved, document-information fields never read or modify host identity data, document variables are never resolved, mail merge never opens data sources or runs, recipient templates are never expanded or rendered, and external/referenced documents are never opened, resolved, or refreshed |
| Bookmark-reference fields | ✅ | ✅ | N/A | Typed inert `REF`/`PAGEREF`/historical `FTNREF`/`NOTEREF` kind/target/option/unknown-switch/cached-result metadata; bookmarks and notes are never resolved or read, links are never created, page and relative-position values are never calculated, and fields are never refreshed |
| Equation fields | ✅ | ✅ | N/A | Typed inert `EQ` expression/cached-result metadata; equation syntax is never parsed, calculated, formatted, rendered, or refreshed |
| Hyperlink fields | ✅ | ✅ | N/A | Typed inert `HYPERLINK` external-target/bookmark/tooltip/frame/image-map-coordinate/new-window/unknown-switch/cached-result metadata; targets are never opened, resolved, followed, activated, or refreshed |
| Table-of-contents entry fields | ✅ | ✅ | N/A | Typed inert `TC` entry/list-identifier/level/page-number-omission/switch/cached-result metadata; hidden-text state is never changed, page numbers are never calculated, and tables of contents are never generated or refreshed |
| Quote fields | ✅ | ✅ | N/A | Typed inert `QUOTE` text-argument/switch/cached-result metadata; character codes and nested fields are never interpreted, and text is never inserted or refreshed |
| Symbol fields | ✅ | ✅ | N/A | Typed inert `SYMBOL` character-argument/switch/cached-result metadata; character codes are never mapped, fonts are never read, glyphs are never inserted, and formatting or layout is never changed |
| Legacy automatic-number fields | ✅ | ✅ | N/A | Typed inert `AUTONUM`/`AUTONUMLGL`/`AUTONUMOUT` kind/switch/cached-result metadata; paragraph numbers are never calculated, heading or style state is never read, paragraphs or layout are never changed, and fields are never refreshed |
| List-number fields | ✅ | ✅ | N/A | Typed inert `LISTNUM` optional-list-name/switch/cached-result metadata; lists, level and start state, numbers, and layout are never read, calculated, changed, or refreshed |
| Printer-control fields | ✅ | ✅ | N/A | Typed inert `PRINT` opaque printer-instruction/cached-result metadata; printer-control text is never interpreted, sent to a printer, or refreshed |
| Embedded-object fields | ✅ | ✅ | N/A | Typed inert `EMBED` opaque object-instruction/cached-result metadata; objects are never loaded, inspected, deserialized, activated, rendered, executed, or refreshed |
| Barcode fields | ✅ | ✅ | N/A | Typed inert `BARCODE` opaque barcode-instruction/cached-result metadata; barcode data and symbology are never parsed, validated, generated, rendered, or refreshed |
| Bidirectional-outline fields | ✅ | ✅ | N/A | Typed inert `BIDIOUTLINE` opaque-instruction/cached-result metadata; right-to-left language, paragraph outline, numbering, and layout are never read, resolved, calculated, or refreshed |
| Drawing-canvas anchor fields | ✅ | ✅ | N/A | Typed inert `SHAPE` opaque-instruction/cached-result metadata; drawings and canvases are never located, linked, loaded, positioned, laid out, rendered, or refreshed |
| Legacy form-code fields | ✅ | ✅ | N/A | Typed inert `FORMTEXT`/`FORMCHECKBOX`/`FORMDROPDOWN` kind/opaque-instruction/cached-result metadata; form-property XML is never read, forms are never filled, selections and checkbox state are never changed, and entry or exit macros are never invoked |
| Legacy private-data fields | ✅ | ✅ | N/A | Typed inert `PRIVATE` opaque-instruction/cached-result metadata; conversion data is never converted, interpreted, made visible, laid out, or refreshed, and the field is not treated as a confidentiality mechanism |
| Historical external-include aliases | ✅ | ✅ | N/A | Typed inert `INCLUDE`/`IMPORT` aliases for text/picture external include metadata; sources are never opened, resolved, imported, fetched, transformed, converted, evaluated, executed, or refreshed |
| Database-query fields | ✅ | ✅ | N/A | Typed inert `DATABASE` opaque-instruction/cached-result metadata; data sources and databases are never opened, connection information is never used, SQL is never executed, tables are never generated or inserted, and fields are never refreshed |
| Legacy INFO fields | ✅ | ✅ | N/A | Typed inert explicit `INFO` property-selector/optional-replacement/switch/cached-result metadata; document and template properties are never read, resolved, modified, or refreshed |
| Database-query fields | ✅ | ✅ | N/A | Typed inert `DATABASE` opaque-instruction/cached-result metadata; data sources and databases are never opened, connection information is never used, SQL is never executed, tables are never generated or inserted, and fields are never refreshed |
| Mail-merge data-source fields | ✅ | ✅ | N/A | Typed inert `DATA` data-source/header-source/switch/cached-result metadata; sources are never opened, read, connected to, resolved, modified, selected, merged, or refreshed |
| Built-in document-information state and statistics | ✅ | ✅ | N/A | Typed inert `CREATEDATE`/`SAVEDATE`/`PRINTDATE`/`REVNUM`/`EDITTIME`/`NUMPAGES`/`NUMWORDS`/`NUMCHARS` kind/switch/cached-result metadata; dates, revision state, and statistics are never read from package metadata, calculated, resolved, or refreshed |
| Built-in document-context and runtime fields | ✅ | ✅ | N/A | Typed inert `FILENAME`/`TEMPLATE`/`DATE`/`TIME`/`PAGE`/`FILESIZE`/`SECTION`/`SECTIONPAGES` kind/switch/cached-result metadata; document paths, attached templates, host filesystem state or file size, current clock values, and page or section layout are never read, resolved, calculated, or refreshed |
| Document statistics | ✅ | ✅ | N/A | Word, character, paragraph, line, and page metadata where present |
| Core, extended, and custom properties | ✅ | ✅ | ✅ | Typed package properties |

### Collaboration, package parts, and advanced content

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Classic comments | ✅ | ✅ | ✅ | Comment bodies, authors, anchors, and CRUD |
| Modern comment metadata | 🟡 | ✅ | ✅ | Extensible comments, people, presence, reactions, and ID mappings; metadata-oriented model |
| Track changes | ✅ | ✅ | ✅ | Insert/delete/move plus paragraph, table, row, cell, and property revisions |
| Content controls | ✅ | ✅ | ✅ | Structured document tags and properties |
| Custom XML data stores and bindings | ✅ | ✅ | ✅ | Package graph load/store and ordered CRUD |
| Document variables | ✅ | ✅ | ✅ | Settings-backed variables plus inert stored `DOCVARIABLE` references; field values are never resolved or refreshed |
| Attached templates | ✅ | ✅ | ✅ | Internal/external template relationship metadata; external targets remain inert |
| Smart tags | ✅ | ✅ | ✅ | Typed smart-tag properties and writer support |
| Glossary/building blocks | ✅ | ✅ | ✅ | Glossary package graph and entry CRUD |
| Web settings | ✅ | ✅ | ✅ | Framesets, HTML divisions, and web-document settings |
| Mail-merge settings and recipients | 🟡 | ✅ | ✅ | Typed package authoring; no data-source fetch or merge execution |
| AltChunk | 🟡 | ✅ | ✅ | Internal/external alternative-format part CRUD; content is not imported or rendered |
| Charts | ✅ | ✅ | ✅ | Classic chart, style, color-style, and embedded-workbook part graphs; no rendering |
| Images | ✅ | ✅ | ✅ | Inline/floating image resources and relationships |
| Drawing/VML shapes | ✅ | ✅ | ✅ | Shape extraction plus VML authoring: `MutableVmlShape` writes rect/roundrect/ellipse/line presets with fill/stroke colors, inline or floating positions, and `v:textbox` stories that round-trip through the text-box inventory; rotation/z-index and WordArt text paths remain read-only |
| Embedded fonts | ✅ | ✅ | ✅ | Font table, payloads, obfuscation, licensing checks, and ordered CRUD |
| Embedded OLE/package objects | ✅ | ✅ | ✅ | Package-level embedded-part discovery plus inert OLE object authoring: validated ProgIDs, Word-convention shape id allocation, `/word/embeddings` parts with content types and relationships, optional preview images, and byte-identical payload round-trips |
| Web extensions/Office Add-ins | 🟡 | ✅ | 🟡 | Shared package-level task-pane create/replace/remove, bounded web-extension parsing/serialization, inert embedded/external snapshot-resource CRUD, typed CT_Blip compression/effect trees, and self-contained mixed-content `extLst` preservation at every MS-OWEXML site; add-ins and links are never activated or fetched |
| Themes | ✅ | ✅ | ✅ | Theme colors, fonts, and related package parts |
| Document protection | 🟡 | ✅ | ✅ | Protection settings and hashes; the library does not enforce editing policy |
| Table of contents | 🟡 | ✅ | ✅ | Typed inert discovery of simple/complex TOC fields, switches, cached results, and dirty/lock state plus field/content authoring; no pagination or automatic refresh |
| Watermarks | ✅ | ✅ | ✅ | Typed VML text-watermark discovery in headers plus generated and arbitrary text watermark authoring (layout, semitransparency, font/size/color) with full-fidelity round-trips, image watermark authoring (format/dimension sniffing, scaling, shared media part) with anchor/payload discovery, and removal |
| Office Math equations in-document | ✅ | ✅ | ✅ | Exact OMML extraction plus validated inline/display equation and math-paragraph authoring; layout and equation evaluation remain renderer responsibilities |
| SmartArt | ✅ | ✅ | ✅ | Typed inert diagram inventory: parsed `dgm:dataModel` node/connection trees, layout/quick-style/colors part metadata, and pre-rendered drawing references in both dialects; authoring generates the definition parts and `dgm:relIds` anchor and round-trips through the read inventory |
| DrawingML text boxes and WordArt | ✅ | ✅ | ✅ | Typed inert text-box/WordArt inventory: DrawingML `wps:txbx` and VML `v:textbox` fallbacks (via MCE), body properties (insets, vertical anchor, direction, wrap, autofit), rich story text with basic run formatting, and WordArt warp presets with inert styling flags; inline text-box authoring with presets, extents, formatted runs, and full bodyPr round-trips through the read inventory; WordArt and floating-anchor authoring are not covered |
| Citations, bibliography, index, and TOA | 🟡 | ✅ | ✅ | Typed inert `CITATION` source-tag/multi-source and `BIBLIOGRAPHY` field discovery, Custom XML bibliography source-store/scalar-value metadata, and TOA/TA plus INDEX/XE metadata expose stored switches, cached results, and dirty/lock state; typed `CITATION` authoring writes caller-supplied tags, locale, volume, prefix/suffix, multi-source order, and optional cached text, while typed `BIBLIOGRAPHY` authoring writes caller-supplied display/filter locales, selected source-tag order, and optional cached text; bibliography styles remain opaque and no citation/table/index generation or refresh occurs; the bibliography source store supports typed source add/replace/remove with graph-preserving mutation |
| IRM/Rights Management | 🟡 | ✅ | 🟡 | Shared MS-OFFCRYPTO DataSpaces inspection validates version/map/definition/transform graphs for OOXML and legacy binary IRM; typed codecs cover inert publishing licenses, cached end-user licenses, certificate chains, protected/viewer-content envelopes, sensitivity labels (including removed-label tombstones and lossless future extensions), EncryptedSIHash/EncryptedDSIHash property integrity verification, and legacy Custom XML promotion semantics. It never contacts rights services, evaluates licenses, or decrypts protected content |
| RibbonX customization | 🟡 | ✅ | ✅ | Word, Excel, and PowerPoint package wrappers retain bounded package-level Custom UI XML parts in each documented relationship family; PowerPoint additionally exposes read-only presentation accessors. All paths validate root relationships and namespaces without executing callbacks, macros, commands, or linked content |
| VBA projects/DOCM macros | 🟡 | 🟡 | ❌ | DOCM/DOTM main parts and the MS-OFFMACRO2 VBA Project → Word Supplemental Data relationship graph are validated as inert metadata; payload contents are never inspected, parsed, or executed by this API, and no macro authoring exists |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral OPC verification and signing |
| Password encryption | ✅ | ✅ | ✅ | Standard/Agile encrypted OOXML wrapper |

## Excel workbooks (XLSX)

### Workbook, cells, and formatting

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open/create/save | ✅ | ✅ | ✅ | Path and in-memory workflows |
| Multiple worksheets | ✅ | ✅ | ✅ | Add, access, update, and serialize sheets |
| Cell values and ranges | ✅ | ✅ | ✅ | String, rich text, number, boolean, error, date/time, and range access |
| Formula cells | ✅ | ✅ | ✅ | Formula strings, cached values, shared formulas, and array formulas |
| Formula evaluation | 🟡 | ✅ | N/A | Shared evaluator with many math, lookup, text, date, financial, and statistical functions |
| Shared strings and rich text | ✅ | ✅ | ✅ | Plain and formatted shared/inline strings |
| Named ranges/defined names | ✅ | ✅ | ✅ | Workbook and sheet scopes, built-ins, comments, and print names |
| Cell and table styles | ✅ | ✅ | ✅ | Fonts, fills, borders, alignment, protection, and number formats |
| Merged cells | ✅ | ✅ | ✅ | Read/write merge ranges |
| Row and column properties | ✅ | ✅ | ✅ | Sizes, visibility, outline, spans, and defaults |
| Freeze/split panes and selections | ✅ | ✅ | ✅ | Worksheet views, panes, selections, and active cells |
| Metadata and properties | ✅ | ✅ | ✅ | Core, extended, custom, and workbook metadata |

### Analysis, drawings, external data, and package features

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Classic charts | ✅ | ✅ | ✅ | Worksheet anchors, chart graphs, styles, user shapes, images, and external-data parts |
| ChartEx | 🟡 | ✅ | ✅ | Extended chart part model and serialization; integration is more limited than classic charts |
| Chart sheets | ✅ | ✅ | ✅ | Views, protection, print settings, chart resources, and package graph; `add_chart_sheet` authoring with interleaved sheet-order preservation, typed chartsheet/drawing emission, and pivot-chart support validated like worksheet pivot charts |
| Pivot tables and caches | ✅ | ✅ | ✅ | Definitions, cache fields/records, filters, grouping, layouts, and writer support |
| Pivot charts | ✅ | ✅ | ✅ | Typed inert chart-to-pivot binding: per-worksheet and chartsheet pivot-chart enumeration, `c:pivotSource` parsing, validated pivot-name resolution, and per-series drop-zone metadata; authoring binds classic charts by name with save-time validation/normalization to canonical sheet-qualified names, round-tripping through the read inventory |
| Structured tables/ListObjects | ✅ | ✅ | ✅ | Columns, formulas, totals, table types, and styles |
| Structured-reference evaluation | 🟡 | ✅ | N/A | Evaluator supports bounded table references, not complete Excel semantics |
| Data validation | ✅ | ✅ | ✅ | Standard and extension collections, formulas, prompts, and ranges |
| Conditional formatting | ✅ | ✅ | ✅ | Standard/extension rules, data bars, color scales, icon sets, and differential formats |
| Classic comments/notes | ✅ | ✅ | ✅ | Comment text, authors, and VML-backed notes |
| Threaded comments | ✅ | ✅ | ✅ | People, mentions, replies, resolution state, and graph-safe CRUD |
| Images and drawing anchors | ✅ | ✅ | ✅ | Pictures, drawing resources, and worksheet anchors |
| Arbitrary DrawingML shapes/text boxes | ✅ | ✅ | ✅ | Typed inert worksheet drawing inventory: two-cell/one-cell/absolute anchors with typed EMU coordinates, ~100 typed preset geometries, hidden/locked flags, text bodies with bodyPr properties and basic run formatting, connection shapes, nested groups, and inert legacy OLE object metadata; `xdr:sp` text-box shape authoring with anchors/bodyPr/runs round-trips through the inventory; group/connection-shape and styling authoring are not covered |
| Hyperlinks | ✅ | ✅ | ✅ | Internal/external links and tooltips |
| Auto-filter and sort state | ✅ | ✅ | ✅ | Values, custom/dynamic/color/icon filters, Top10, and multi-key sorts |
| Sparklines | ✅ | ✅ | ✅ | Groups, axes, colors, and extension markup |
| Slicers and slicer caches | ✅ | ✅ | ✅ | Package-aware load/store and ordered CRUD |
| Timelines and timeline caches | ✅ | ✅ | ✅ | Package-aware load/store and ordered CRUD |
| External workbook, DDE, and OLE links | ✅ | ✅ | ✅ | Typed inert links, cached sheet data, names, and targets; never refreshed automatically |
| Connections and query tables | ✅ | ✅ | ✅ | Typed package CRUD; external queries are never executed |
| OLE objects | ✅ | ✅ | ✅ | Worksheet object metadata, anchors, payload resources, and package graph |
| ActiveX controls | 🟡 | ✅ | ❌ | Typed control/property discovery; no worksheet control authoring graph |
| Web extensions/Office Add-ins | 🟡 | ✅ | 🟡 | Shared package-level task-pane create/replace/remove, bounded web-extension parsing/serialization, inert embedded/external snapshot-resource CRUD, typed CT_Blip compression/effect trees, self-contained mixed-content `extLst` preservation at every MS-OWEXML site, and typed worksheet `x15:webExtensions` range bindings exposed through worksheet/workbook read and transactional mutation APIs with MS-OWEXML `appRef` cross-validation; add-ins and links are never activated or fetched |
| XML maps | 🟡 | ✅ | ✅ | Typed inert MapInfo/schema/data-binding package CRUD with strict/transitional relationships; mappings, schema locations, and bound files are never resolved or executed |
| Volatile dependencies | 🟡 | ✅ | ✅ | Typed inert workbook-scoped RTD/OLAP dependency package CRUD; never contacts servers/connections or evaluates formulas |
| Data model/custom data/XLDM | 🟡 | ✅ | ✅ | Inert model/custom-data package storage plus bounded XLDM inspection/writing |
| Workbook revisions | ✅ | ✅ | ✅ | Revision headers, users, logs, and package storage; revisions are not replayed |
| Calculation properties | ✅ | ✅ | ✅ | Calculation mode, IDs, iteration, precision, and reference mode |
| Calculation chain | 🟡 | ✅ | ✅ | Typed inert parse/store of caller-authored calculation order; no dependency rebuilding or formula evaluation |
| Named sheet views | ✅ | ✅ | ✅ | Typed filters, sorts, ranges, color-sort differential formats, and extensions; validated worksheet-scoped package/workbook CRUD and construction/mutation, with open-ended differential-format and extension XML retained as bounded inert markup |
| Page setup, margins, and print options | ✅ | ✅ | ✅ | Orientation, paper, scaling, fit-to-page, margins, and options |
| Print areas/titles | ✅ | ✅ | ✅ | Built-in defined names |
| Headers and footers | ✅ | ✅ | ✅ | Odd/even/first sections and formatting codes |
| Page breaks and printer settings | ✅ | ✅ | ✅ | Horizontal/vertical breaks and printer-resource graphs |
| Sheet protection/protected ranges | ✅ | ✅ | ✅ | Legacy and strong hashes plus protected-range metadata |
| Workbook protection | 🟡 | ✅ | ✅ | Passive reader preserves structure/window/revision locks and legacy/strong verifier metadata; writer remains structure/window focused, with no password checking or policy enforcement |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral OPC verification and signing |
| Password encryption | ✅ | ✅ | ✅ | Standard/Agile encrypted OOXML wrapper |
| VBA projects/XLSM macros | 🟡 | 🟡 | ❌ | The MS-OFFMACRO2 Workbook → VBA Project relationship graph is validated as inert metadata; payload contents are never inspected, parsed, or executed by this API, and no macro authoring exists |

## PowerPoint presentations (PPTX)

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open/create/save | ✅ | ✅ | ✅ | Path and in-memory package workflows |
| Slides and ordering | ✅ | ✅ | ✅ | Add, delete, duplicate, move, resize, enumerate slides and stable IDs, expose hidden state, and validate root slide/master relationships |
| Text, text boxes, and bullets | ✅ | ✅ | ✅ | Text extraction and formatted text-box authoring, plus presentation-default text-style inventory and Kinsoku line-breaking settings |
| Basic shapes and groups | ✅ | ✅ | ✅ | Rectangles, ellipses, text boxes, nested groups, formatting, non-visual IDs, and placeholder inventories |
| Images and backgrounds | ✅ | ✅ | ✅ | Picture resources, photo-album defaults, plus solid, gradient, pattern, and relationship-resolved picture backgrounds with slide/layout/master inheritance |
| Tables | ✅ | ✅ | ✅ | Table extraction and authoring |
| Classic and extended charts | ✅ | ✅ | ✅ | Per-slide classic-chart inventory with relationship/part identity and basic type/title/legend metadata, plus multiple chart types, chart/style/color parts, ChartEx, and embedded workbook resources |
| SmartArt | ✅ | ✅ | ✅ | Diagram data/layout/style/color part graphs and builder support |
| Audio, video, posters, and captions | ✅ | ✅ | ✅ | Typed slide-level media resources, embedded/linked media, trim/fade/bookmark metadata, and high-level text-track inventory |
| Animations and timing trees | ✅ | ✅ | ✅ | Shape effects, sequences, triggers, chart/diagram timing relationships, and timing metadata on slides, layouts, and masters |
| Transitions and slide advance timing | ✅ | ✅ | ✅ | Typed base effects, including directional, corner, orientation, and wheel-spoke variants; through-black and split in/out options; slide/layout/master effective inheritance; MCE-backed custom durations; speed, sound, click, and timed advance |
| PowerPoint 2010 transition extensions | ✅ | ✅ | ✅ | Compatibility-choice ripple effects on slides, layouts, and masters, with typed corner/center direction and extended duration; deterministic fade fallback authoring |
| Hyperlinks and slide-jump actions | 🟡 | ✅ | ✅ | Strict/transitional hyperlink relationships, validated inline slide navigation, plus bounded inert click/hover action-setting inventory for PowerPoint-reserved action values and declared targets; no target is followed, opened, activated, or executed |
| Classic comments | ✅ | ✅ | ✅ | Validated high-level graph, legacy adapters, authors, slide comment parts, and package-aware CRUD |
| Modern comments | ✅ | ✅ | ✅ | High-level validated graph, authors, anchors, replies, status, and package-aware CRUD |
| Speaker notes and notes masters | ✅ | ✅ | ✅ | High-level complete notes graph load/store with resources and themes |
| Slide masters and layouts | ✅ | ✅ | ✅ | Semantic reading, including master/layout shape and placeholder inventory, header/footer and master-content visibility flags, matching/type and master/layout retention metadata, typed layout-reference identifiers, master text-style level inventory, and slide/layout/master relationship resolution; authoring of new masters (default text styles), typed layouts with placeholders, placeholder add/replace, and unreferenced-layout removal, all re-validated against the read-side graph; layout repointing and master deletion are not covered |
| Handout master | ✅ | ✅ | ✅ | Presentation-root relationship resolution plus layout and header/footer settings |
| Themes | ✅ | ✅ | ✅ | Master-, layout-, and slide-scoped validated theme resolution, typed color maps/overrides, and presentation inventory; theme authoring with typed 12-slot color schemes and major/minor font schemes, master attachment with graph validation, scheme replacement on existing theme parts, and theme override model + authoring on slides and layouts (parse/store/replace/remove with orphaned-part cleanup); fmtScheme authoring is not covered |
| Sections | ✅ | ✅ | ✅ | Typed section readers with stable IDs and resolved slide-index membership, plus graph-safe CRUD |
| Custom slide shows | ✅ | ✅ | ✅ | High-level typed inventory of named subsets, plus graph-safe ordered CRUD |
| Presentation/slide protection | ✅ | ✅ | ✅ | Protection and password metadata, including root modification-verifier inspection; policy is not enforced by the library |
| Embedded fonts | ✅ | ✅ | ✅ | High-level typed inventory, payloads, obfuscation, licensing checks, and ordered CRUD |
| Embedded OLE/package objects | ✅ | ✅ | ✅ | Typed per-slide inert OLE inventory with shape, ProgID, relationship, and payload metadata; inert OLE authoring with validated ProgIDs, `/ppt/embeddings` parts, `p:oleObj` frames, and byte-identical payload round-trips; never activated |
| View and presentation properties, and guides | 🟡 | ✅ | ✅ | Typed root/package presentation settings, high-level extended-guide inventory, slide/notes surface dimensions and size type, including PowerPoint 2010 browse-mode metadata, with bounded serialization |
| Tags, changes, and revision information | 🟡 | ✅ | 🟡 | Inert programmable tags with slide-aware package/presentation inventory, root smart-tag and customer-data relationship references, revision/change readers, and add-only validated package storage |
| Web extensions/Office Add-ins | 🟡 | ✅ | 🟡 | Shared package-level task-pane create/replace/remove, bounded web-extension parsing/serialization, inert embedded/external snapshot-resource CRUD, typed CT_Blip compression/effect trees, and self-contained mixed-content `extLst` preservation at every MS-OWEXML site; add-ins and links are never activated or fetched |
| Ink annotations | ✅ | ✅ | ✅ | Bounded inert InkML content-part inventory with slide/relationship/part identity and stored trace counts, plus validated inert InkML storage onto slides (`p:contentPart` + `customXml` relationship, dialect-preserving); no handwriting recognition, rendering, replay, or execution |
| Laser pointer traces | ✅ | ✅ | ✅ | Bounded inert PowerPoint 2010 slide-show trace inventory with stored time offsets and coordinates, plus validated inert trace storage (new `p14:laserTraceLst` extension, existing/empty/missing `p:extLst` handled, dialect-preserving, read-back self-check); never replayed, rendered, interpolated, or executed |
| Slide-show event records | ✅ | ✅ | ✅ | Bounded inert PowerPoint 2010 trigger/media event inventory with slide, event, target object, and stored timeline metadata, plus validated inert event storage (`p14:showEvtLst` via the shared extension patcher); events are never replayed or executed |
| VBA projects/PPTM macros | 🟡 | 🟡 | ❌ | High-level PPTM/PPSM/POTM macro-project metadata validates the MS-OFFMACRO2 Presentation → VBA Project relationship graph; payload contents are never inspected, parsed, or executed by this API, and no macro authoring exists |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral OPC verification and signing |
| Password encryption | ✅ | ✅ | ✅ | Standard/Agile encrypted OOXML wrapper |

## Word binary documents (DOC)

### Document and internal structures

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open/create/save | ✅ | ✅ | ✅ | OLE2 DOC reading, new-document writing, and transactional package editors |
| Text, paragraphs, and runs | ✅ | ✅ | ✅ | Piece-table text with character/paragraph SPRMs |
| Tables | ✅ | ✅ | ✅ | TAP structures, rows, cells, and formatting |
| Sections and page layout | ✅ | ✅ | ✅ | Margins, paper, columns, borders, page numbering, and line numbering |
| Styles and fonts | ✅ | ✅ | ✅ | Style sheet and font-table generation |
| Headers and footers | ✅ | ✅ | ✅ | Per-section story ranges and types |
| Footnotes and endnotes | ✅ | ✅ | ✅ | References, text stories, and writer APIs |
| Numbering and lists | ✅ | ✅ | ✅ | List templates, overrides, names, and numbering formats |
| Hyperlinks and fields | ✅ | ✅ | ✅ | HYPERLINK plus every parsed field story’s stored instruction and cached-result text, with typed inert `MACROBUTTON` macro-name/button/cached-result, `ADDIN`/`CONTROL`/`HTMLCONTROL` kind/instruction/cached-result metadata, `GLOSSARY`/`AUTOTEXT` entry-name/switch/cached-result and `AUTOTEXTLIST` display/style/tip/unknown-switch/cached-result metadata, `INDEX` configuration/unknown-switch/cached-result metadata, `REF`/`PAGEREF`/`FTNREF`/`NOTEREF` kind/bookmark/options/cached-result metadata, `SET` target/opaque-expression/cached-result metadata, `=` formula/cached-result metadata, `SEQ` identifier/bookmark/opaque-tail/cached-result metadata, `STYLEREF` style/options/unknown-switch/cached-result metadata, `TOC`/`TOA` options/unknown-switch/cached-result metadata, `GOTOBUTTON` destination/button metadata, `USERADDRESS`/`USERINITIALS`/`USERNAME` kind/override/formatting metadata, `ADVANCE` point-adjustment metadata, `DDE`/`DDEAUTO` application/source/item and representation/storage metadata, `LINK` application-type/source/item/result/formatting metadata, `INCLUDETEXT`/`INCLUDEPICTURE` and historical `INCLUDE`/`IMPORT` source/bookmark/converter/XML-option metadata, `MERGEFIELD` data-column/switch/cached-result, `DATA` data/header-source/switch/cached-result metadata, `DOCPROPERTY` property-name/switch/cached-result, `DOCVARIABLE` name/switch/cached-result, `MERGEREC`/`MERGESEQ` kind/cached-result, `NEXT` cached-result/state, `NEXTIF`/`SKIPIF` kind/unparsed-comparison/cached-result, `IF` unparsed-expression/cached-result, `COMPARE` unparsed-comparison/cached-result, `ASK`/`FILLIN` prompt/default-response metadata, and `ADDRESSBLOCK`/`GREETINGLINE` recipient-layout/country/locale/fallback metadata; fields are never evaluated, document properties are never read or resolved, document variables are never resolved, `COMPARE` fields never parse or evaluate a comparison, prompt fields never display a dialog or capture a response, mail merge never opens data sources or runs, recipient templates are never expanded or rendered, navigation fields are never activated, add-ins and controls are never loaded, instantiated, rendered, or executed, building-block fields never look up stored entries, read templates, show a selection UI, or insert content, index fields never scan markers, read bookmarks, calculate page numbers, sort entries, paginate, generate an index, or refresh, bookmark-reference fields never look up bookmarks, read referenced ranges, resolve page or note numbers, create links, calculate relative positions, or refresh, `SET` fields never evaluate expressions, look up or change bookmarks, change document state, or refresh, formula fields never parse or evaluate formulas, read table cells or bookmarks, resolve field values, or refresh, `SEQ` fields never look up bookmarks, increment or reset sequences, calculate numbers, or refresh, `STYLEREF` fields never look up styled text, search document stories, calculate paragraph numbers or relative positions, resolve page layout, or refresh, table-of-contents fields never scan entries, read bookmarks, resolve links, paginate, regenerate a table, or refresh, table-of-authorities fields never find citations, scan hidden text, read bookmarks, follow links, calculate page numbers, paginate, regenerate a table, or refresh, user-identity fields never read or modify host identity data, `ADVANCE` fields never move text, change layout, or reflow content, DDE never starts a conversation or opens a source, LINK never activates OLE or opens a source, external includes never open, resolve, import, fetch, refresh, transform, convert, evaluate, or execute sources, and macros and external sources remain inert |
| Equation fields | ✅ | ✅ | N/A | Typed inert `EQ` native-type/expression/cached-result metadata; equation syntax is never parsed, calculated, formatted, rendered, or refreshed |
| Hyperlink fields | ✅ | ✅ | N/A | Typed inert `HYPERLINK` native-type/external-target/bookmark/tooltip/frame/image-map-coordinate/new-window/unknown-switch/cached-result metadata; targets are never opened, resolved, followed, activated, or refreshed |
| Table-of-contents entry fields | ✅ | ✅ | N/A | Typed inert `TC` story/marker-position/entry/option/unknown-switch/cached-result metadata; marker characters are scanned from stored story text because native `Plcfld` metadata omits them, and entries never change hidden text, calculate page numbers, generate a table, or refresh |
| Table-of-authorities entry fields | ✅ | ✅ | N/A | Typed inert `TA` story/marker-position/option/unknown-switch/cached-result metadata; marker characters are scanned from stored story text because native `Plcfld` metadata omits them, and entries never find citations, change hidden text, follow bookmarks, calculate page numbers, generate a table, or refresh |
| Index-entry fields | ✅ | ✅ | N/A | Typed inert `XE` story/marker-position/entry/option/unknown-switch/cached-result metadata; marker characters are scanned from stored story text because native `Plcfld` metadata omits them, and entries never change hidden text, resolve bookmarks, calculate page numbers, sort entries, generate an index, or refresh |
| Referenced-document fields | ✅ | ✅ | N/A | Typed inert `RD` story/marker-position/source/relative-path/switch/cached-result metadata; marker characters are scanned from stored story text because native `Plcfld` metadata omits them, and references never open, resolve, read, import, refresh, evaluate, or execute a referenced document |
| Legacy private-data fields | ✅ | ✅ | N/A | Typed inert `PRIVATE` story/marker-position/opaque-instruction/cached-result metadata; marker characters are scanned from stored story text because native `Plcfld` metadata omits them, conversion data is never converted, interpreted, made visible, laid out, or refreshed, and the field is not treated as a confidentiality mechanism |
| Quote fields | ✅ | ✅ | N/A | Typed inert `QUOTE` native-type/text-argument/switch/cached-result metadata; character codes and nested fields are never interpreted, and text is never inserted or refreshed |
| Symbol fields | ✅ | ✅ | N/A | Typed inert `SYMBOL` native-type/character-argument/switch/cached-result metadata; character codes are never mapped, fonts are never read, glyphs are never inserted, and formatting or layout is never changed |
| Legacy automatic-number fields | ✅ | ✅ | N/A | Typed inert `AUTONUM`/`AUTONUMLGL`/`AUTONUMOUT` native-type/kind/switch/cached-result metadata; paragraph numbers are never calculated, heading or style state is never read, paragraphs or layout are never changed, and fields are never refreshed |
| List-number fields | ✅ | ✅ | N/A | Typed inert `LISTNUM` native-type/optional-list-name/switch/cached-result metadata; lists, level and start state, numbers, and layout are never read, calculated, changed, or refreshed |
| Printer-control fields | ✅ | ✅ | N/A | Typed inert `PRINT` native-type/opaque-printer-instruction/cached-result metadata; printer-control text is never interpreted, sent to a printer, or refreshed |
| Embedded-object fields | ✅ | ✅ | N/A | Typed inert `EMBED` native-type/opaque-object-instruction/cached-result metadata; objects are never loaded, inspected, deserialized, activated, rendered, executed, or refreshed |
| Barcode fields | ✅ | ✅ | N/A | Typed inert `BARCODE` native-type/opaque-barcode-instruction/cached-result metadata; barcode data and symbology are never parsed, validated, generated, rendered, or refreshed |
| Bidirectional-outline fields | ✅ | ✅ | N/A | Typed inert `BIDIOUTLINE` native-type/opaque-instruction/cached-result metadata; right-to-left language, paragraph outline, numbering, and layout are never read, resolved, calculated, or refreshed |
| Drawing-canvas anchor fields | ✅ | ✅ | N/A | Typed inert `SHAPE` native-type/opaque-instruction/cached-result metadata; drawings and canvases are never located, linked, loaded, positioned, laid out, rendered, or refreshed |
| Legacy form-code fields | ✅ | ✅ | N/A | Typed inert `FORMTEXT`/`FORMCHECKBOX`/`FORMDROPDOWN` native-type/kind/opaque-instruction/cached-result metadata; form properties are never read, forms are never filled, selections and checkbox state are never changed, and entry or exit macros are never invoked |
| Legacy INFO fields | ✅ | ✅ | N/A | Typed inert `INFO` native-type/property-selector/optional-replacement/switch/cached-result metadata; document and template properties are never read, resolved, modified, or refreshed |
| Built-in document-information fields | ✅ | ✅ | N/A | Typed inert `TITLE`/`SUBJECT`/`AUTHOR`/`KEYWORDS`/`COMMENTS`/`LASTSAVEDBY`/`CREATEDATE`/`SAVEDATE`/`PRINTDATE`/`REVNUM`/`EDITTIME`/`NUMPAGES`/`NUMWORDS`/`NUMCHARS` native-kind/switch/cached-result metadata; document metadata and host identity are never read, resolved, calculated, or modified |
| Built-in document-context and runtime fields | ✅ | ✅ | N/A | Typed inert `FILENAME`/`TEMPLATE`/`DATE`/`TIME`/`PAGE`/`FILESIZE`/`SECTION`/`SECTIONPAGES` native-kind/switch/cached-result metadata; document paths, attached templates, host filesystem state or file size, current clock values, and page or section layout are never read, resolved, calculated, or refreshed |
| Bookmarks | ✅ | ✅ | ✅ | Bookmark ranges and writer support |
| Comments | ✅ | ✅ | ✅ | Annotation ranges, authors, and reply metadata |
| Track changes | ✅ | ✅ | ✅ | Transactional add/update/remove/accept/reject editing |
| FIB, piece tables, FKPs, and BinTable | ✅ | ✅ | ✅ | Core DOC storage and formatting structures |
| SPRM properties and DOP versions | ✅ | ✅ | ✅ | Typed properties with unknown-data preservation where applicable |
| Associated strings, saved-by, proofing, and revision tables | 🟡 | ✅ | 🟡 | Typed auxiliary tables; mutation coverage varies by table |
| Glossary/AutoText | 🟡 | ✅ | 🟡 | Typed glossary structures with bounded authoring |

### Advanced and package features

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Images | 🟡 | ✅ | 🟡 | Inline/floating picture and blip extraction; `DocWriter::insert_picture` writes inline PNG/JPEG pictures (OfficeArtWordDrawing in the Data stream) and `insert_floating_picture` writes floating ones (0x0008 anchor + PlcfSpa + fcDggInfo OfficeArtContent) |
| Drawings and shapes | 🟡 | ✅ | 🟡 | OfficeArt/Escher shape extraction (Data stream + fcDggInfo drawing group) and textbox story reading (main + header); `DocWriter::insert_floating_shape`/`insert_floating_text_box` write main-story shapes and text boxes, `insert_header_text_box`/`insert_header_picture` write odd/even/first-page header text boxes and pictures (PlcfSpaHdr + ccpHdrTxbx + PlcfHdrtxbxTxt) |
| Embedded OLE/package objects | ✅ | ✅ | ✅ | Add, remove, reorder, and preserve embedded object storages; payloads remain inert |
| Custom XML data storage | ✅ | ✅ | ✅ | Bounded, lossless `MsoDataStore` item/property XML with typed item GUIDs, schema references, known item-family classification, and IRM redundant/modified promotion markers; schema URIs are never resolved |
| Smart tags/factoids | 🟡 | ✅ | ❌ | Bounded MS-OSHARED `PropertyBagStore` decoding plus validated `SttbfBkmkFactoid`, start/end bookmark PLCFs, positional property bags, and `Plcffactoid` recognizer-state ranges; recognizers, VBA callbacks, download URLs, and schemas remain inert |
| MathType/MTEF equations | 🟡 | ✅ | ❌ | Equation Native extraction and conversion; no DOC equation authoring |
| Summary/document properties | ✅ | ✅ | ✅ | OLE property-set reading and editing |
| Document protection settings | 🟡 | ✅ | ✅ | Typed settings/hashes; policy is not enforced |
| Password encryption | ✅ | ✅ | ✅ | Supported DOC encryption profiles and encrypted writer output |
| Macro-security metadata | ✅ | ✅ | ✅ | Passive DOP metadata only; macros are never executed |
| VBA project/code modules | 🟡 | 🟡 | ❌ | Directory-only inert MS-OVBA project-storage topology discovery; candidate module names are reported, but module/source bytes are never opened, decompressed, parsed, or executed |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral CFB verification and signing |

## Excel binary workbooks (XLS)

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| BIFF versions | ✅ | ✅ | ✅ | Reads BIFF2-BIFF8; writes BIFF8 |
| Workbooks, worksheets, and cells | ✅ | ✅ | ✅ | Multiple sheets and all principal cell value records |
| Formula tokens, shared/array formulas | ✅ | ✅ | ✅ | Ptg token streams and cached results |
| Formula evaluation | 🟡 | ✅ | N/A | Shared evaluator via `WorkbookTrait`; not complete Excel semantics |
| Shared strings and rich text | ✅ | ✅ | ✅ | SST/CONTINUE handling and formatting runs |
| Defined names | ✅ | ✅ | ✅ | Workbook/sheet names and extended metadata |
| Styles and number formats | ✅ | ✅ | ✅ | Fonts, fills, borders, alignment, XF/DXF, and custom formats |
| Merged cells | ✅ | ✅ | ✅ | BIFF merge ranges |
| Rows, columns, outlines, and views | ✅ | ✅ | ✅ | Dimensions, hidden state, freeze/split panes, selections, and window settings |
| Conditional formatting | ✅ | ✅ | ✅ | Classic and extended rule records |
| Data validation | ✅ | ✅ | ✅ | Validation collections, prompts, and ranges |
| Hyperlinks | ✅ | ✅ | ✅ | URL, file, and internal monikers |
| Comments/notes | ✅ | ✅ | ✅ | NOTE/OBJ/TXO text and object records |
| Images and primitive drawing shapes | 🟡 | ✅ | ✅ | OfficeArt extraction plus bounded primitive shape CRUD |
| Charts and chart sheets | 🟡 | ✅ | ✅ | Typed embedded/chart-sheet substreams and transactional CRUD; no renderer |
| Pivot tables and caches | ✅ | ✅ | ✅ | Cache values, grouping, fields, filters, and view/editor support |
| Structured tables/ListObjects | ✅ | ✅ | ✅ | ListObject, AutoFilter12, web/XML, and external-source metadata |
| Auto-filter and sort | ✅ | ✅ | ✅ | Filter conditions, filter modes, and sort records |
| External workbook, DDE, and OLE links | ✅ | ✅ | ✅ | Inert links, caches, names, and monikers; never refreshed automatically |
| Embedded OLE objects | ✅ | ✅ | ✅ | Package editor CRUD; embedded payloads remain inert |
| Custom XML data storage | ✅ | ✅ | ✅ | Bounded, lossless `MsoDataStore` item/property XML with typed item GUIDs, schema references, known item-family classification, and IRM redundant/modified promotion markers; schema URIs are never resolved |
| Page setup, headers/footers, and breaks | ✅ | ✅ | ✅ | Print/page records and page-break authoring |
| Protection | ✅ | ✅ | ✅ | Sheet, object, scenario, workbook, and password records |
| Calculation, scenarios, and consolidation | ✅ | ✅ | ✅ | Typed settings and inert scenario/consolidation metadata |
| Codepage handling | 🟡 | ✅ | 🟡 | Reader honors BIFF codepages; writer is centered on BIFF8/Windows-1252 |
| Password encryption | ✅ | ✅ | ✅ | XOR and supported RC4/CryptoAPI profiles |
| VBA project metadata | 🟡 | ✅ | 🟡 | Inert BIFF markers/code names plus directory-only MS-XLS `_VBA_PROJECT_CUR` → MS-OVBA topology discovery; no macro stream is opened, decompressed, parsed, or executed, and writing is limited to a module-free metadata scaffold |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral CFB verification and signing |

## Excel binary OOXML workbooks (XLSB)

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open/create/save | ✅ | ✅ | ✅ | Binary workbook and worksheet part support |
| Worksheets and cell values | ✅ | ✅ | ✅ | Blank, numeric, string, boolean, error, date, and shared-string cells |
| Formula records | ✅ | ✅ | ✅ | Numeric/string/boolean/error results plus shared and array formulas |
| Formula evaluation | 🟡 | ✅ | N/A | Shared evaluator via `WorkbookTrait`; not complete Excel semantics |
| Shared strings and styles | ✅ | ✅ | ✅ | Fonts, fills, borders, number formats, XF, and alignment |
| Named ranges | ✅ | ✅ | ✅ | BrtName parsing and writing |
| Merged cells | ✅ | ✅ | ✅ | New-workbook and existing-package mutation |
| Row/column information | ✅ | ✅ | ✅ | Widths, heights, spans, and hidden state |
| Hyperlinks and comments | ✅ | ✅ | ✅ | Locations, tooltips, authors, and comment text |
| Data validation | ✅ | ✅ | ✅ | Binary validation records and writer support |
| Conditional formatting | ✅ | ✅ | ✅ | Core and extension rule records |
| Auto-filter and sort | ✅ | ✅ | ✅ | Typed binary filter/sort models |
| Sheet protection | ✅ | ✅ | ✅ | Protection flags and password metadata |
| Calculation properties | ✅ | ✅ | ✅ | Workbook calculation settings |
| Pivot tables/caches | ✅ | ✅ | ✅ | Typed inert PivotCache definition stream model (MS-XLSB 2.1.7.38): refresh metadata, worksheet/consolidation sources, cache fields with shared items of every value type, range/discrete grouping, OLAP hierarchies and tuple caches, calculated items/members with inert formula tokens, and Excel 2010/2014 extensions, exposed per cache id alongside the existing pivot views; authoring serializes the full model with lossless-or-refuse semantics and workbook-stream cache-id wiring |
| Charts and drawings | 🟡 | ✅ | ❌ | Typed inert chart-sheet metadata (tab color, views, protection, page setup, drawing links), bounded SpreadsheetDrawing inventory (anchors, shapes, pictures, graphic frames), and embedded-chart resolution into the shared DrawingML chart model; no authoring |
| Structured tables | ✅ | ✅ | ✅ | Typed inert ListObject model (MS-XLSB 2.1.7.51): identity, ranges, table types, header/totals metadata, DXF style ids, typed totals-row functions, inert calculated-column/totals formulas, and style-info flags; worksheet list parts are resolved eagerly at load; authoring serializes tables with validated display names/ranges/columns and per-sheet BrtListPart wiring |
| External links and connections | ✅ | ✅ | ❌ | Typed inert external-workbook, DDE, and OLE link targets, sheet names, and declared name/item metadata, plus the MS-XLSB 2.1.7.24 External Data Connections part: DBType/CmdType enums, refresh and credential metadata, ODBC/OLE DB/OLAP/Web properties, typed parameters, and Web query tables; connection strings, commands, URLs, and credentials are stored verbatim and never resolved, contacted, refreshed, or executed |
| Web extensions/Office Add-ins | 🟡 | ✅ | ✅ | Bounded `WEBEXTENSIONS` collection and `BrtWebExtension` payload codecs preserve exact FRT formula bytes, enforce one REFERENCE-class 3D cell/area range, validate workbook-resolved internal-sheet XTI indices and exact UTF-16 `appRef` values, cross-check MS-OWEXML bindings, and integrate with immutable worksheet reads and mutable worksheet emission; records remain inert and are never activated |
| VBA project/code modules | 🟡 | 🟡 | ❌ | Inert MS-XLSB Workbook → VBA Project topology plus declared legacy/Agile signature-part metadata; project/signature payload contents are never inspected, parsed, verified, or executed |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral OPC verification and signing |
| Password encryption | ✅ | ✅ | ✅ | Standard/Agile encrypted OOXML wrapper |

## PowerPoint binary presentations (PPT)

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open/create/save | ✅ | ✅ | ✅ | OLE2 presentation streams and writer/editor APIs |
| Slides, masters, and persist mapping | ✅ | ✅ | ✅ | Slide management, main masters, and persist-object lookup |
| Text, runs, and formatting | ✅ | ✅ | ✅ | Text boxes, placeholders, fonts, colors, paragraphs, and runs |
| Shapes, groups, and OfficeArt | ✅ | ✅ | ✅ | AutoShapes, groups, anchors, fills, gradients, lines, and Escher records |
| Pictures | ✅ | ✅ | ✅ | JPEG/PNG/BLIP resources and writer support |
| Tables | ✅ | ✅ | ✅ | Table group/grid/cell extraction and table authoring (rows, columns, cell text, cell dimensions) |
| Native charts | ✅ | ✅ | ✅ | Typed inert inventory of embedded MSGraph/Excel.Chart OLE payloads (persist-mapping resolution, zlib decompression, full BIFF8 chart models) with slide-frame attribution and per-object failure isolation, plus chart authoring: `PptWriter::add_chart` builds bar/line/pie BIFF8 chart workbooks, embeds them as ExOleObjStg with shared ExObjList allocation, and round-trips through the inventory; never activated or rendered |
| Hyperlinks | ✅ | ✅ | ✅ | URLs and slide navigation |
| Action/interaction settings | 🟡 | ✅ | 🟡 | Typed action, jump, trigger, and macro metadata with bounded writer integration |
| Notes | ✅ | ✅ | ✅ | Speaker-note records |
| Comments | ✅ | ✅ | ✅ | Comment2000 records and presentation aggregation |
| Animations | ✅ | ✅ | ✅ | Build steps, triggers, motion paths, and transactional editor |
| Transitions and slide timings | ✅ | ✅ | ✅ | Transition type/speed/direction and advance timing |
| Custom slide shows | ✅ | ✅ | ✅ | Named show containers and slide-ID lists |
| Headers and footers | ✅ | ✅ | ✅ | Presentation and slide header/footer records |
| View information and guides | ✅ | ✅ | ✅ | View state, guides, and related settings |
| Audio/video | 🟡 | ✅ | 🟡 | Sound collections and linked/embedded media metadata; authoring coverage is bounded |
| Embedded OLE objects | ✅ | ✅ | ✅ | Add, remove, reorder, and preserve package storages; payloads remain inert |
| Custom XML data storage | ✅ | ✅ | ✅ | Bounded, lossless `MsoDataStore` item/property XML with typed item GUIDs, schema references, known item-family classification, and IRM redundant/modified promotion markers; schema URIs are never resolved |
| Smart tags | 🟡 | ✅ | ❌ | PowerPoint 11 smart-tag stores use the shared bounded MS-OSHARED property-bag codec with codepage-aware ANSI strings; recognizers, download URLs, and schemas remain inert |
| Presentation settings/metadata | ✅ | ✅ | ✅ | Slide-show, print, HTML publish, broadcast, envelope, routing, and privacy metadata |
| Modify password/protection | ✅ | ✅ | ✅ | Password and protection metadata; policy is not enforced |
| Password encryption | ✅ | ✅ | ✅ | Supported PPT encryption profiles |
| VBA project metadata | 🟡 | ✅ | ✅ | Inert MS-PPT `VBAInfo`/`VbaProjectStg` persist ID, compression, and payload-size metadata; storage is never decompressed or parsed as CFB, and no project/module code is exposed or executed |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral CFB verification and signing |

## OpenDocument common package features

These rows apply to packaged ODF families unless a format-specific row says otherwise.

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Package, manifest, and MIME handling | ✅ | ✅ | ✅ | ZIP package validation, resource access, and deterministic writing |
| `content.xml`, `styles.xml`, `meta.xml`, `settings.xml` | ✅ | ✅ | ✅ | Namespace-aware parsing and package writing |
| Document/template MIME families | ✅ | ✅ | ✅ | Standard document and template media types listed in Compatibility |
| Metadata | ✅ | ✅ | ✅ | Dublin Core and ODF metadata fields |
| Styles and data styles | ✅ | ✅ | ✅ | Common, automatic, master, page, text, table, and number styles |
| Embedded resource discovery/mutation | 🟡 | ✅ | ✅ | Images, objects, and subdocuments; creation support varies by host family |
| Annotation package CRUD | ✅ | ✅ | ✅ | Text, spreadsheet, and presentation anchors with ordered graph-safe mutation |
| Forms and controls | ✅ | ✅ | ✅ | Typed nested forms, properties, events, and broad control families |
| Scripts | 🟡 | ✅ | ✅ | Script package artifacts can be preserved/mutated but are never executed |
| RDF metadata graphs | ✅ | ✅ | ✅ | Package graph and triple CRUD |
| External references | ✅ | ✅ | ✅ | Parsed/serialized as inert links; never fetched or refreshed automatically |
| Encryption | ✅ | ✅ | ✅ | Per-entry package encryption and password opening |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral signing and cryptographic verification |

## OpenDocument text (ODT/OTT)

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open/create/save | ✅ | ✅ | ✅ | Document and template packages, paths, readers, and bytes |
| Text, paragraphs, spans, and headings | ✅ | ✅ | ✅ | Rich text extraction and mutation |
| Tables | ✅ | ✅ | ✅ | Nested tables, properties, rows, and cells |
| Lists and outline styles | ✅ | ✅ | ✅ | Ordered/unordered lists, labels, outline levels, alignment, and typed outline-style inspection/mutation; no label regeneration |
| Sections | ✅ | ✅ | ✅ | Add, wrap, unwrap, replace, remove, and protected/linked metadata |
| Styles and page layouts | ✅ | ✅ | ✅ | Paragraph/text/table styles, separate `content.xml`/`styles.xml` font-face declaration inspection and mutation, typed named fill-image, gradient, hatch, marker, opacity, and stroke-dash inspection, document line-numbering configuration, columns, drop caps, tab stops, page properties, and read-only explicit `text:page-sequence` master-page assignments; no pagination, line-number generation, font-resource loading, link following, style-use resolution, or rendering |
| Master pages, headers, and footers | ✅ | ✅ | ✅ | Master-page CRUD and header/footer content/properties, including typed cached page/navigation/statistic, reference/variable/sequence, conditional/formula/DDE/meta, database, document identity/revision, sender identity/contact, and script/macro metadata; fields remain inert |
| Hyperlinks | ✅ | ✅ | ✅ | Typed inert `text:a` insertion with XLink target/show/actuate and office/text metadata; links are never followed |
| Footnotes and endnotes | 🟡 | ✅ | ✅ | Typed footnote/endnote configuration inspection/mutation and separator support, plus validated inert `text:note-body` construction, parsing, and replacement (paragraphs/lists/tables/selected drawing content); `Note::rich_body` exposes existing rich bodies as namespace-resolved nodes for structural edits, while links, fields, scripts, and macro metadata remain inert |
| Bookmarks and reference marks | ✅ | ✅ | ✅ | Point/range targets and typed insertion/replacement/removal |
| Comments/annotations | ✅ | ✅ | ✅ | Point/range annotations and package-aware CRUD |
| Track changes | ✅ | ✅ | ✅ | Change metadata, regions, policy, and mutation APIs |
| Dynamic/database fields | ✅ | ✅ | ✅ | Date/time/page/user/variable/drop-down/database families plus inert inline `text:script` metadata, including cached `text:author-name`/`text:author-initials`, `text:sender-*`, `text:file-name`, `text:template-name`, `text:sheet-name`, and `text:chapter`; no external query execution, script execution or link opening, host identity/contact or path reads, template lookups, or live outline/sheet-state resolution |
| Variables and declarations | ✅ | ✅ | ✅ | Typed declarations and mutation |
| Ruby annotations | 🟡 | ✅ | 🟡 | Typed `text:ruby` insertion, named ruby styles, and mutable CRUD; append or validate a UTF-8 range wholly inside one text/CDATA/entity node without splitting surrounding markup |
| TOC, indexes, and source marks | 🟡 | ✅ | ✅ | Typed structures and cached-body authoring; no pagination or automatic regeneration |
| Bibliography records | 🟡 | ✅ | ✅ | Typed bibliography policy inspection/mutation plus inert records and source marks; no automatic entry generation or citation resolution |
| Images and drawing frames | ✅ | ✅ | ✅ | Semantic discovery and resource replacement/removal, plus frame authoring: `insert_image` (PNG/JPEG/GIF sniffed, verbatim `Pictures/` payloads, typed lengths/anchors) and `insert_text_box` round-trip through the read inventory |
| Embedded charts | ✅ | ✅ | ✅ | Package-subdocument/inline chart add, replace, and remove |
| Embedded objects and MathML | 🟡 | ✅ | 🟡 | Semantic discovery and resource mutation; formulas/objects remain inert |
| Forms | ✅ | ✅ | ✅ | Broad typed form-control creation and mutation |
| Mail-merge/database sources | 🟡 | ✅ | ✅ | Field/source metadata only; no merge or database execution |
| Protection/visibility settings | 🟡 | ✅ | ✅ | Section and document metadata; policy is not enforced |
| Scripts, RDF, encryption, signatures | ✅ | ✅ | ✅ | Common ODF package capabilities; scripts remain inert |

## OpenDocument spreadsheet (ODS/OTS)

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open/create/save | ✅ | ✅ | ✅ | Spreadsheet and template packages, paths, readers, and bytes |
| Sheets, rows, columns, and cells | ✅ | ✅ | ✅ | Add/remove sheets, range access, insert/delete rows/columns, and typed values |
| Cell value types | ✅ | ✅ | ✅ | String, number, boolean, date/time, duration, percentage, currency, and error-like values |
| Formula strings and references | ✅ | ✅ | ✅ | OpenFormula text and cached values |
| Formula evaluation | 🟡 | ✅ | N/A | Immutable ODS snapshots implement the shared workbook trait; common OpenFormula A1 references and semicolon arguments are normalized for the evaluator, while unsupported grammar remains an explicit formula error |
| Repeated and merged cells/rows | ✅ | ✅ | ✅ | Semantic expansion and deterministic serialization |
| Styles and full cell formatting | ✅ | ✅ | ✅ | Text, alignment, borders, backgrounds, number/data styles, protection styles, and read-only named fill-image/gradient/hatch/marker/opacity/stroke-dash inspection; no link following, style-use resolution, or rendering |
| Conditional cell styles | ✅ | ✅ | ✅ | ODF style-map conditions and ordered mutation; not the full Excel rule family |
| Content validation | ✅ | ✅ | ✅ | Conditions, prompts, error messages, events, definitions, and cell bindings |
| Comments/annotations | ✅ | ✅ | ✅ | Rich text/lists, creator/date, geometry, extensions, and CRUD |
| Hyperlinks | 🟡 | ✅ | 🟡 | Typed inert `text:a` anchors preserve and author non-overlapping UTF-8 text ranges with XLink and office/text metadata; inline rich-text styling is flattened |
| Images | ✅ | ✅ | ✅ | Sheet image resources, alternatives, and mutation |
| General drawing shapes | 🟡 | ✅ | 🟡 | Semantic frames/shapes with bounded authoring compared with ODG/ODP |
| Embedded charts | ✅ | ✅ | ✅ | Chart add, replace, and remove with package or inline storage |
| Embedded objects | 🟡 | ✅ | 🟡 | Semantic discovery and resource replacement/removal; payloads remain inert |
| Database ranges, filters, and sorts | ✅ | ✅ | ✅ | Recursive filters, sort keys, subtotals, and inert query/source metadata |
| Named ranges and expressions | ✅ | ✅ | ✅ | Global and sheet-local definitions with CRUD |
| DataPilot/pivot tables | ✅ | ✅ | ✅ | Sources, fields, levels, references, groups, and mutation |
| Sheet/document protection | ✅ | ✅ | ✅ | Keys, permissions, direct cell flags, and protection styles; no policy enforcement |
| Print and page settings | ✅ | ✅ | ✅ | Sheet print ranges/settings and page styles |
| Calculation settings and consolidations | ✅ | ✅ | ✅ | Calculation, label-range, scenario, and inert consolidation metadata |
| DDE sources and caches | ✅ | ✅ | ✅ | Typed declarations/cached tables; no DDE execution |
| Tracked changes | ✅ | ✅ | ✅ | Typed change families and mutation |
| Forms, scripts, and RDF | ✅ | ✅ | ✅ | Common package CRUD; scripts remain inert |
| Encryption and signatures | ✅ | ✅ | ✅ | Common ODF security support |
| CSV export | ✅ | ✅ | N/A | Sheet/table export utility |

## OpenDocument presentation (ODP/OTP)

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open/create/save | ✅ | ✅ | ✅ | Presentation and template packages, paths, readers, and bytes |
| Slides and text | ✅ | ✅ | ✅ | Add/insert/remove/reorder slides and extract/edit text |
| Text boxes and basic shapes | ✅ | ✅ | ✅ | Text, rectangle, ellipse, line, connector, groups, and shape properties |
| Images | ✅ | ✅ | ✅ | Embedded/linked images and package resources |
| Tables | ✅ | ✅ | ✅ | Table-shaped content and serialization |
| Embedded charts | ✅ | ✅ | ✅ | Chart add, replace, and remove with package or inline storage |
| Audio and video | ✅ | ✅ | ✅ | Linked or package-contained media, parameters, and mutation |
| Slide layouts | ✅ | ✅ | ✅ | Presentation page-layout definitions and mutation |
| Master pages | ✅ | ✅ | ✅ | Presentation master-page CRUD and reference repair |
| Styles and backgrounds | ✅ | ✅ | ✅ | Presentation styles, page styles, drawing properties, and read-only named fill-image, gradient, hatch, marker, opacity, and stroke-dash definitions; no link following, style-use resolution, or rendering |
| Animations | ✅ | ✅ | ✅ | ODF/SMIL trees and legacy effects with ordered mutation |
| Transitions and timings | ✅ | ✅ | ✅ | Transition type/style/speed/direction/duration/sound and automatic timing |
| Speaker notes | ✅ | ✅ | ✅ | Notes extraction and authoring |
| Comments/annotations | ✅ | ✅ | ✅ | Slide annotations and package-aware CRUD |
| Hyperlinks and presentation actions | ✅ | ✅ | ✅ | URLs, page jumps, events, and inert script bindings |
| Custom slide shows/settings | ✅ | ✅ | ✅ | Named page subsets, declarations, and presentation settings |
| Forms and controls | ✅ | ✅ | ✅ | Common ODF form package support |
| Embedded objects | 🟡 | ✅ | 🟡 | Semantic discovery and resource mutation; payloads remain inert |
| Scripts and RDF | ✅ | ✅ | ✅ | Common package CRUD; scripts remain inert |
| Encryption and signatures | ✅ | ✅ | ✅ | Common ODF security support |

## Additional OpenDocument families

| Family | Extensions | Status | Read | Write | Notes |
|--------|------------|--------|------|-------|-------|
| Drawing | `.odg`, `.otg` | ✅ | ✅ | ✅ | Pages, layers, standard 2D shapes, groups, text, geometry, metadata, typed named fill-image/gradient/hatch/marker/opacity/stroke-dash inspection from immutable and mutable drawings, resources, builder, and mutable CRUD; no link following, style-use resolution, or rendering |
| Standalone chart | `.odc`, `.otc` | ✅ | ✅ | ✅ | Titles, legends, plot areas, axes, series, data points, analytics nodes, cached tables, and semantic mutation |
| Formula document | `.odf`, `.otf` | ✅ | ✅ | ✅ | MathML mixed-content model, annotations, lossless source save, validated formula/template package construction from direct MathML roots, and a typed inert MathML tree editor (validated mutation, well-formed serialization with regenerated namespace prefixes, typed schemata builders, atomic `set_math` repackaging); no evaluation |
| Image document | `.odi`, `.oti` | 🟡 | ✅ | 🟡 | Frames, linked/package/base64 images, text boxes, objects, tables, maps, and exact lossless save |
| Master document | `.odm`, `.otm` | ✅ | ✅ | ✅ | Paragraphs, linked sections/subdocuments, indexes, styles, encryption, signing, builder, and mutable CRUD |
| Web template | `.oth` | ✅ | ✅ | ✅ | Text semantic reader, exact lossless save, and a dedicated authoring model (`WebDocumentBuilder` create-from-scratch, `MutableWebDocument` convert/edit) that re-validates every package through the web reader; `text-web` remains a legacy producer MIME |
| Database front end | `.odb` | 🟡 | ✅ | 🟡 | Connections, settings, forms, reports, queries, tables, schemas, keys, indices, and package mutation; nothing is executed |
| Flat OpenDocument | `.fodt`, `.fods`, `.fodp`, `.fodg`, `.fodc`, `.fodi`, `.fodf` | ✅ | ✅ | ✅ | Family validation and exact lossless save through `FlatOpenDocument`, family-typed wrappers that open flat files through the full packaged semantic readers (text/tables, sheets/cells, slides/notes, drawings, charts, image frames), and mutable text/spreadsheet/presentation variants that splice edits back into the flat XML while preserving unmodified sections byte-identically; `.fodg`/`.fodc`/`.fodi` stay read-only and new binary media cannot be represented in flat saves |

## Rich Text Format (RTF)

### Content, layout, and formatting

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Text, paragraphs, and runs | ✅ | ✅ | ✅ | Unicode/ANSI text, formatting groups, and deterministic serialization |
| Sections | ✅ | ✅ | ✅ | Typed multi-section properties, headers/footers, explicit `\sect` boundaries, and inherited sections round-trip in source order |
| Page layout, columns, borders, and numbering | ✅ | ✅ | ✅ | Orientation, dimensions, margins, facing pages, columns, page borders, and line/page numbering |
| Headers and footers | ✅ | ✅ | ✅ | Header/footer story content and types |
| Tables | ✅ | ✅ | ✅ | Nested/floating tables, merges, geometry, borders, shading, distances, banding, and story ownership |
| Character formatting | ✅ | ✅ | ✅ | Fonts, sizes, colors, bold/italic/underline, borders, shading, positioning, scaling, and kerning |
| Paragraph formatting | ✅ | ✅ | ✅ | Alignment, indents, spacing, tabs, borders, shading, bidi, flow, drop caps, and style references |
| Stylesheets and latent styles | ✅ | ✅ | ✅ | Paragraph, character, section, table, inheritance, latent styles, filters, and restrictions |
| Lists and numbering | ✅ | ✅ | ✅ | Modern list tables/overrides plus legacy section and paragraph numbering |
| Languages and bidirectional text | ✅ | ✅ | ✅ | Document defaults, character languages, LTR/RTL, and East Asian controls |
| Pictures and alternatives | ✅ | ✅ | ✅ | Common raster/metafile types, crop/layout metadata, identities, and compatibility alternatives |
| Shapes, groups, and text frames | ✅ | ✅ | ✅ | Geometry, anchors, wrapping, fills, gradients, themes, binary properties, stories, and mutation |
| Legacy drawings and text boxes | ✅ | ✅ | ✅ | Primitive/callout models and canonical round trips |

### Fields, review, metadata, and advanced destinations

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Fields and hyperlinks | ✅ | ✅ | ✅ | Parsed field codes, status, nested fields, URLs, bookmarks, page breaks, typed inert `HYPERLINK` target/bookmark/display-option metadata, `REF`/`PAGEREF`/`NOTEREF` kind/bookmark/switch metadata, `MACROBUTTON` macro-name/display metadata, `ADDIN`/`CONTROL`/`HTMLCONTROL` kind/instruction/cached-result metadata, `GLOSSARY`/`AUTOTEXT` entry-name/unknown-switch/cached-result metadata, `AUTOTEXTLIST` display/style/tip/unknown-switch/cached-result metadata, `GOTOBUTTON` destination/button metadata, `USERADDRESS`/`USERINITIALS`/`USERNAME` kind/override/formatting metadata, `ADVANCE` point-adjustment metadata, `DDE`/`DDEAUTO` application/source/item and representation/storage metadata, `LINK` application-type/source/item/result/formatting metadata, `RD` source/relative-path/switch metadata, `INCLUDETEXT`/`INCLUDEPICTURE` source/converter/XML-option metadata, `TOC`/`TC`/`TA`/`TOA`/`INDEX`/`XE` configuration/entry metadata, `CITATION`/`BIBLIOGRAPHY` source-tag/filter/locale metadata, and typed `DOCPROPERTY` property-name/switch/cached-result, `TITLE`/`SUBJECT`/`AUTHOR`/`KEYWORDS`/`COMMENTS`/`LASTSAVEDBY` kind/switch/cached-result, `DOCVARIABLE`/`MERGEFIELD` name/switch/cached-result, `MERGEREC`/`MERGESEQ` kind/cached-result, `NEXT` cached-result/state, `NEXTIF`/`SKIPIF` kind/unparsed-comparison/cached-result, `IF` unparsed-expression/cached-result, `COMPARE` unparsed-comparison/cached-result, `SET` target/opaque-expression/cached-result, `SEQ` identifier/bookmark/opaque-tail/cached-result metadata, `=` formula/cached-result metadata, `STYLEREF` style/options/unknown-switch/cached-result metadata, `ASK`/`FILLIN` prompt/default-response metadata, and `ADDRESSBLOCK`/`GREETINGLINE` recipient-layout/country/locale/fallback metadata; no field recalculation, comparison evaluation, expression evaluation, formula evaluation, table-cell or bookmark reads, field-value resolution, bookmark lookup or mutation, sequence calculation, style-reference lookup, document-story search, paragraph-number or relative-position calculation, page-layout resolution, document-state changes, layout changes, text movement, reflow, prompt display or response capture, macro lookup/execution, add-in or control loading, instantiation, rendering, or execution, building-block lookup, template reads, selection UI, content insertion, bookmark changes, or refresh, navigation-target resolution or activation, host-identity reads or modification, document-property or document-information resolution, DDE contact, OLE activation, bibliography-source lookup/style application/content generation, document-variable resolution, mail merge, recipient-template expansion/rendering, generated-content refresh, or external-source resolution |
| Quote fields | ✅ | ✅ | N/A | Typed inert `QUOTE` text-argument/switch/cached-result metadata; character codes and nested fields are never interpreted, and text is never inserted or refreshed |
| Symbol fields | ✅ | ✅ | N/A | Typed inert `SYMBOL` character-argument/switch/cached-result metadata; character codes are never mapped, fonts are never read, glyphs are never inserted, and formatting or layout is never changed |
| Legacy automatic-number fields | ✅ | ✅ | N/A | Typed inert `AUTONUM`/`AUTONUMLGL`/`AUTONUMOUT` kind/switch/cached-result metadata; paragraph numbers are never calculated, heading or style state is never read, paragraphs or layout are never changed, and fields are never refreshed |
| List-number fields | ✅ | ✅ | N/A | Typed inert `LISTNUM` optional-list-name/switch/cached-result metadata; lists, level and start state, numbers, and layout are never read, calculated, changed, or refreshed |
| Printer-control fields | ✅ | ✅ | N/A | Typed inert `PRINT` opaque printer-instruction/cached-result metadata; printer-control text is never interpreted, sent to a printer, or refreshed |
| Embedded-object fields | ✅ | ✅ | N/A | Typed inert `EMBED` opaque object-instruction/cached-result metadata; objects are never loaded, inspected, deserialized, activated, rendered, executed, or refreshed |
| Barcode fields | ✅ | ✅ | N/A | Typed inert `BARCODE` opaque barcode-instruction/cached-result metadata plus `DISPLAYBARCODE`/`MERGEBARCODE` kind/data-argument/type/switch metadata; barcode data and symbology are never validated, merge fields are never resolved, and barcodes are never generated, rendered, or refreshed |
| Bidirectional-outline fields | ✅ | ✅ | N/A | Typed inert `BIDIOUTLINE` opaque-instruction/cached-result metadata; right-to-left language, paragraph outline, numbering, and layout are never read, resolved, calculated, or refreshed |
| Drawing-canvas anchor fields | ✅ | ✅ | N/A | Typed inert `SHAPE` opaque-instruction/cached-result metadata; drawings and canvases are never located, linked, loaded, positioned, laid out, rendered, or refreshed |
| Legacy INFO fields | ✅ | ✅ | N/A | Typed inert explicit `INFO` property-selector/optional-replacement/switch/cached-result metadata; document and template properties are never read, resolved, modified, or refreshed |
| Mail-merge data-source fields | ✅ | ✅ | N/A | Typed inert `DATA` data-source/header-source/switch/cached-result metadata; sources are never opened, read, connected to, resolved, modified, selected, merged, or refreshed |
| Built-in document-information state and statistics | ✅ | ✅ | N/A | Typed inert `CREATEDATE`/`SAVEDATE`/`PRINTDATE`/`REVNUM`/`EDITTIME`/`NUMPAGES`/`NUMWORDS`/`NUMCHARS` kind/switch/cached-result metadata; dates, revision state, and statistics are never read from document metadata, calculated, resolved, or refreshed |
| Built-in document-context and runtime fields | ✅ | ✅ | N/A | Typed inert `FILENAME`/`TEMPLATE`/`DATE`/`TIME`/`PAGE`/`FILESIZE`/`SECTION`/`SECTIONPAGES` kind/switch/cached-result metadata; document paths, attached templates, host filesystem state or file size, current clock values, and page or section layout are never read, resolved, calculated, or refreshed |
| Bookmarks and navigation entries | ✅ | ✅ | ✅ | Bookmark ranges, index entries, TOC entries, and page references |
| Footnotes/endnotes and separators | ✅ | ✅ | ✅ | Note bodies, numbering/options, section overrides, and separator stories |
| Comments/annotations | ✅ | ✅ | ✅ | Point/range comments, identity, positions, and mutation |
| Track changes | ✅ | ✅ | ✅ | Author table, insert/delete ranges, revision metadata, and mutation |
| Legacy form-code fields | ✅ | ✅ | N/A | Typed inert `FORMTEXT`/`FORMCHECKBOX`/`FORMDROPDOWN` kind/opaque-instruction/cached-result metadata; form properties are not reconciled, forms are never filled, selections and checkbox state are never changed, and entry or exit macros are never invoked |
| Legacy private-data fields | ✅ | ✅ | N/A | Typed inert `PRIVATE` opaque-instruction/cached-result metadata; conversion data is never converted, interpreted or revealed, laid out, or refreshed, and the field is not treated as a confidentiality mechanism |
| Historical external-include aliases | ✅ | ✅ | N/A | Typed inert `INCLUDE`/`IMPORT` aliases for text/picture external include metadata; sources are never opened, resolved, imported, fetched, transformed, converted, evaluated, executed, or refreshed |
| Form fields | ✅ | ✅ | ✅ | Text, checkbox, dropdown, help/status, defaults, and positional mutation |
| Mail-merge metadata | ✅ | ✅ | ✅ | Data sources, field mappings, and recipients as inert metadata; no merge execution |
| Document variables and user properties | ✅ | ✅ | ✅ | Typed values, lexical forms, links, Unicode, and mutation |
| Embedded OLE objects | ✅ | ✅ | ✅ | OLE1 header decoding, object data/results, positions, and mutation; payloads remain inert |
| Equations/math | 🟡 | ✅ | 🟡 | Typed inert `EQ` field discovery and caller-authored `EQ` field serialization, a typed syntactic model of the ECMA-376 `EQ` instruction switches (fractions, radicals, scripts, integrals/sums/products, arrays, brackets, boxes, overstrikes, lists, and displacements), plus embedded equation objects and math-property metadata; equations are never calculated, formatted, or rendered |
| Embedded fonts | ✅ | ✅ | ✅ | `fontemb`/`fontfile` destinations and inline data |
| Themes and data stores | ✅ | ✅ | ✅ | Inert theme/data-store bytes and typed mutation |
| File table and external references | ✅ | ✅ | ✅ | Bounded inert external-file metadata; targets are never resolved |
| XML namespaces and XSL transform metadata | ✅ | ✅ | ✅ | Namespace table, transform location/usage, and XML policies; no transform execution |
| Document protection and write reservations | ✅ | ✅ | ✅ | Protection controls, users, hashes, reservations, and save preferences; no policy enforcement |
| Document/view/print/compatibility policies | ✅ | ✅ | ✅ | Typed RTF 1.9.1 settings across layout, rendering, privacy, revision, save, style, and compatibility groups |
| Document info and generator/origin metadata | ✅ | ✅ | ✅ | Title/author/timestamps, generator, origin, caption, and revision-save metadata |
| Compressed RTF | ✅ | ✅ | ✅ | LZFu compression/decompression |
| Digital signatures | N/A | N/A | N/A | RTF does not define package signatures |
| File encryption | N/A | N/A | N/A | RTF does not define a standard encrypted-file wrapper |

## Apple iWork archives

All iWork support is read-only. Models decode IWA object streams and extract useful structure; they
do not attempt lossless archive rewriting.

### Shared IWA infrastructure

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Bundle and IWA parsing | ✅ | ✅ | ❌ | ZIP bundle, Snappy framing, protobuf messages, varints, and archive metadata |
| Object index/reference graph | ✅ | ✅ | ❌ | Message-type lookups and relationship resolution |
| Text extraction | ✅ | ✅ | ❌ | TSWP storage extraction across Pages, Keynote, and Numbers |
| Structured extraction | ✅ | ✅ | ❌ | Tables, slides, sections, shape text, and chart metadata |
| Media discovery/extraction | ✅ | ✅ | ❌ | Images, audio, video, PDFs, and other bundle assets |
| Chart metadata | 🟡 | ✅ | ❌ | Chart type, title, row/column labels, series count, and cached grid metadata; no renderer/editor |
| Password-protected archives | ❌ | ❌ | ❌ | Not implemented |

### Pages

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Text and text styles | ✅ | ✅ | ❌ | Paragraph/character text storage and style extraction |
| Sections | ✅ | ✅ | ❌ | Body, header, footer, and floating section structure |
| Headers and footers | 🟡 | ✅ | ❌ | Text/section extraction; layout fidelity is limited |
| Floating drawables and media | 🟡 | ✅ | ❌ | Shape text and referenced asset extraction |
| Tables | 🟡 | ✅ | ❌ | Shared structured table extraction where table archives are present |
| Charts | 🟡 | ✅ | ❌ | Shared chart-metadata extraction |
| Comments, revisions, hyperlinks, and notes | ❌ | ❌ | ❌ | No public semantic models |

### Keynote

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Slides and text | ✅ | ✅ | ❌ | Titles, content, text storages, and ordering |
| Master references | ✅ | ✅ | ❌ | Master-slide identifiers |
| Build animations | 🟡 | ✅ | ❌ | Target/duration and bounded effect classification |
| Slide transitions | 🟡 | ✅ | ❌ | Duration and bounded transition classification |
| Speaker notes | ✅ | ✅ | ❌ | Referenced note text extraction |
| Multimedia references/assets | ✅ | ✅ | ❌ | Bundle reference and media extraction |
| Tables and charts | 🟡 | ✅ | ❌ | Shared table and chart-metadata extraction |
| Hyperlinks, actions, comments, and themes | ❌ | ❌ | ❌ | No public semantic models |

### Numbers

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Sheets and tables | ✅ | ✅ | ❌ | Sheet hierarchy, tables, row/column dimensions, and headers |
| Cell values | ✅ | ✅ | ❌ | Text, number, boolean, date-like, duration-like, formula, and empty values |
| Formula reconstruction | 🟡 | ✅ | ❌ | Reconstructs bounded formula AST expressions; no Numbers evaluator integration |
| Cell formatting | 🟡 | ✅ | ❌ | Selected stored format/style information |
| CSV export | ✅ | ✅ | N/A | Table export utility |
| Charts | 🟡 | ✅ | ❌ | Shared chart-metadata extraction |
| Filters, pivots, conditional highlighting, and named ranges | ❌ | ❌ | ❌ | No public semantic models |
| Comments, hyperlinks, and protection | ❌ | ❌ | ❌ | No public semantic models |

## Tabular interchange formats

These APIs are part of the unified sheet module but are not Office package formats.

| Format/feature | Status | Read | Write | Notes |
|----------------|--------|------|-------|-------|
| CSV | ✅ | ✅ | ✅ | Configurable delimiter/quote/escape, encoding policy, type inference, and workbook-trait access |
| TSV | ✅ | ✅ | ✅ | Tab-delimited preset over the text-workbook engine |
| Custom delimited text | ✅ | ✅ | ✅ | Caller-provided `TextConfig` |
| PRN/fixed-width | ✅ | ✅ | ✅ | Configurable fixed-column parsing and writing |
| SYLK | ✅ | ✅ | ✅ | Symbolic Link reader/writer mapped to text-workbook data |
| DIF | ✅ | ✅ | ✅ | Data Interchange Format reader/writer mapped to text-workbook data |
| Formula evaluation | 🟡 | ✅ | N/A | Text workbooks implement `WorkbookTrait`; formulas depend on parsed cell content and evaluator coverage |

## Performance, safety, and API quality

| Feature | Status | Notes |
|---------|--------|-------|
| In-memory APIs | ✅ | Major formats can open from bytes/readers and save to bytes/writers |
| Zero-copy/borrowed parsing | 🟡 | Used in selected binary/XML hot paths; many semantic models intentionally own data |
| Lazy/on-demand loading | 🟡 | Package parts and some semantic resources load on demand; not universal |
| SIMD acceleration | 🟡 | Used for selected signature, numeric, and formula/XML paths; not a global parser guarantee |
| Parallel processing | 🟡 | Rayon is used in Markdown conversion and selected ZIP/crypto work |
| Streaming | 🟡 | Reader/writer APIs exist, but many package formats still materialize parts or complete archives |
| Memory-mapped files | ❌ | No mmap-backed public API |
| Bounded parsers/resource limits | ✅ | Advanced package editors enforce size, count, depth, and decoded-string limits |
| Transactional mutation | ✅ | Many package graph editors validate staged changes before replacing original data |
| Typed errors/results | ✅ | Format-specific and shared error models use `Result`-based APIs |
| Feature-gated dependencies | ✅ | Optional ODF, RTF, iWork, fonts, formula, and image-conversion stacks |
| Examples and API documentation | ✅ | Crate docs and workspace examples cover reading, authoring, conversion, and evaluation |
| Unit/integration/fixture tests | ✅ | Extensive unit and integration suites, including real-producer fixtures and round trips |
| Fuzz targets | ✅ | Parser fuzz targets exist across core format crates |
| API stability | 🟡 | The project is under active development and does not promise a stable public API yet |

## Compatibility summary

The `Read` and `Write` columns below mean that a public high-level or format-specific API exists.
They do not override the per-feature limits above.

### Microsoft Office

| Format | Extensions | Read | Write | Compatibility scope |
|--------|------------|------|-------|---------------------|
| Word OOXML | `.docx` | ✅ | ✅ | Office 2007+ OOXML; strict/transitional package handling with feature limits above |
| Excel OOXML | `.xlsx` | ✅ | ✅ | Office 2007+ OOXML; broad worksheet/package authoring |
| Excel binary OOXML | `.xlsb` | ✅ | ✅ | Office 2007+ binary workbook parts; advanced feature gaps above |
| PowerPoint OOXML | `.pptx` | ✅ | ✅ | Office 2007+ OOXML; broad slide/package authoring |
| Word binary | `.doc` | ✅ | ✅ | Office 97-2003 OLE2; new writer plus targeted existing-package editors |
| Excel BIFF | `.xls` | ✅ | ✅ | Reads BIFF2-BIFF8 and writes/edits BIFF8 |
| PowerPoint binary | `.ppt` | ✅ | ✅ | Office 97-2003 OLE2; writer and targeted package editors |

Macro-enabled OOXML variants (`.docm`, `.dotm`, `.xlsm`, `.xltm`, `.pptm`, `.ppsm`, and `.potm`)
and macro-enabled XLSB workbooks have package-level support. Their published VBA APIs validate and
report only inert relationship metadata (including declared XLSB signature-part metadata); binary
project, signature, and code-module payloads remain opaque and are never executed. Litchi has no
typed VBA code-module model or macro-authoring support.

### OpenDocument

| Format family | Extensions | Read | Write | Compatibility scope |
|---------------|------------|------|-------|---------------------|
| Text/document template | `.odt`, `.ott` | ✅ | ✅ | Semantic reader, builder, and mutable document |
| Spreadsheet/template | `.ods`, `.ots` | ✅ | ✅ | Semantic reader, builder, and mutable spreadsheet |
| Presentation/template | `.odp`, `.otp` | ✅ | ✅ | Semantic reader, builder, and mutable presentation |
| Drawing/template | `.odg`, `.otg` | ✅ | ✅ | Semantic reader, builder, and mutable drawing |
| Standalone chart/template | `.odc`, `.otc` | ✅ | ✅ | Semantic chart model and bounded mutation |
| Formula/template | `.odf`, `.otf` | ✅ | 🟡 | Semantic MathML reader, lossless source save, and validated formula/template package construction |
| Image/template | `.odi`, `.oti` | ✅ | 🟡 | Semantic reader and lossless save |
| Master document/template | `.odm`, `.otm` | ✅ | ✅ | Semantic reader, builder, and mutable master document |
| Web template | `.oth` | ✅ | ✅ | Text-compatible reader, lossless save, and dedicated authoring |
| Database front end | `.odb` | ✅ | 🟡 | Semantic configuration plus bounded package mutation; no database execution |
| Flat OpenDocument | `.fodt`, `.fods`, `.fodp`, `.fodg`, `.fodc`, `.fodi`, `.fodf` | ✅ | 🟡 | Validation, semantic reading through the packaged family models, and exact lossless save |

ODF models target ISO/IEC 26300 structures and retain producer extensions where their typed or opaque
models allow it. The matrix does not claim complete ODF-version conformance validation.

### RTF and Apple iWork

| Format | Extensions | Read | Write | Compatibility scope |
|--------|------------|------|-------|---------------------|
| Rich Text Format | `.rtf` | ✅ | ✅ | Broad RTF 1.9.1 semantic reader/writer; multi-section limitation noted above |
| Apple Pages | `.pages` | ✅ | ❌ | Read-only IWA text/section/asset extraction |
| Apple Keynote | `.key` | ✅ | ❌ | Read-only IWA slide/notes/build/transition extraction |
| Apple Numbers | `.numbers` | ✅ | ❌ | Read-only IWA sheet/table/cell/formula extraction |

## Source map and contributions

The implementation and its tests are the source of truth for this matrix. Principal locations are:

- OOXML shared/package features: `crates/litchi-opc/src/` and `crates/litchi-ooxml/src/`
- DOCX, XLSX, XLSB, PPTX: `crates/litchi-ooxml/src/docx/`, `xlsx/`, `xlsb/`, and `pptx/`
- OLE/CFB infrastructure: `crates/litchi-cfb/src/` and `crates/litchi-ole/src/`
- DOC, XLS, PPT: `crates/litchi-ole/src/doc/`, `xls/`, and `ppt/`
- OpenDocument: `crates/litchi-odf/src/`
- Rich Text Format: `crates/litchi-rtf/src/`
- Apple iWork: `crates/litchi-iwa/src/`
- Formula parsing/conversion and evaluation: `crates/litchi-formula/src/` and `crates/litchi-eval/src/`
- Unified facades, Markdown, image conversion, and text workbooks: `crates/litchi/src/`

When changing support, update the relevant public API/tests and this matrix in the same change. Use a
🟡 status whenever support is intentionally bounded, pass-through, or metadata-only, and describe the
boundary in the Notes cell.
