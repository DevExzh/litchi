# Word Binary (DOC) Feature Matrix

This document tracks the public feature support of litchi-doc for the Word binary
.doc format described by [MS-DOC] and stored in an OLE2 Compound File Binary
container. It is a compatibility matrix, not a claim of complete Word-version,
rendering, or legacy-application conformance.

## Scope, status, and specification references

| Mark | Meaning |
|------|---------|
| ✅ | Typed public support for the scope in Notes |
| 🟡 | Bounded, partial, metadata-only, pass-through, or inert support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to the format or direction |

Read and Write describe the public direction independently. A row can be supported as
inert metadata or a preserved payload without opening external data, executing macros,
evaluating fields, or rendering OfficeArt. Generic CFB/OLE stream access is not counted as
typed support for every structure in a stream.

The primary audit sources are [MS-DOC] ToC.md, Front Matter.md, and 2 Structures:
2.1 File Structure, 2.2 Fundamental Concepts, 2.3 Document Parts, 2.4 Document
Content, 2.5 The File Information Block, 2.6 Single Property Modifiers, 2.7
Document Properties, 2.8 PLCs, and 2.9 Basic Types.

Related references are [MS-CFB] for the compound-file container, [MS-OLEPS] for
property sets, [MS-OFFCRYPTO] for encryption and IRM DataSpaces, [MS-OVBA] for VBA
projects, [MS-OSHARED] for shared Office metadata, and [MS-ODRAW] for binary OfficeArt.
The implementation owner is crates/litchi-doc/src/; litchi-cfb, litchi-sign, and
shared Office codecs are counted only where the DOC API exposes them.

## Core Word binary document model


## [MS-DOC] file structure and version audit

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| OLE2/CFB container and Word streams | ✅ | ✅ | ✅ | Package validates the WordDocument stream and the 0Table or 1Table stream, exposes bounded OLE editing, and writes new or transactionally edited packages |
| Word 97+ FIB and versioned FibRgFcLcb tables | ✅ | ✅ | ✅ | FIB base/version fields, table-stream pointers, Word 97 through later extension ranges, fixed-layout older records, and declared pointer bounds are handled by the DOC reader/writer |
| Word 6/95 and pre-Microsoft-Office binary profiles | ❌ | ❌ | ❌ | The public reader intentionally rejects files predating the Word 97 binary profile rather than treating incompatible table layouts as DOC |
| WordDocument, table, Data, and story ranges | ✅ | ✅ | ✅ | Piece-table text, formatting tables, auxiliary Data content, and story character counts are exposed through typed document and writer APIs |
| Custom XML data storage | ✅ | ✅ | ✅ | Bounded, lossless MsoDataStore item/property XML with typed GUIDs, schema references, IRM promotion markers, and inert schema URIs |
| Summary and document-summary property streams | ✅ | ✅ | ✅ | OLE property-set reading and editing with typed summary, document-summary, and user-defined property access |
| Encryption stream and password-to-open profiles | ✅ | ✅ | ✅ | The supported DOC encryption profiles cover the implemented XOR/RC4-compatible modes and typed password errors; unsupported profiles fail explicitly |
| XML signatures storage and signatures stream | ✅ | ✅ | ✅ | Trust-neutral CFB signature verification and transactional signing/editing are exposed; certificate trust and revocation are not established |
| ObjectPool and embedded OLE/package objects | ✅ | ✅ | ✅ | Object storages can be inventoried, added, removed, reordered, and preserved with inert payloads; objects are never activated or opened |
| Macros storage and VBA project payload | ✅ | ✅ | ✅ | Bounded MS-OVBA compressed-container, dir, PROJECT, module, and cache-free source metadata are read and authored; code is never compiled or executed |
| Information Rights Management Data Space and protected content | ❌ | ❌ | ❌ | [MS-DOC] 2.1.12-2.1.13 and [MS-OFFCRYPTO] define IRM/protected-content storage, but litchi-doc exposes no typed DOC rights/license/decryption API; generic OLE preservation is not semantic support |
| Unknown streams and storages | 🟡 | ✅ | 🟡 | Package-preserving editors can retain unrelated CFB topology where supported, but an arbitrary stream has no DOC semantic model or write guarantee |

## [MS-DOC] PLC, property, and auxiliary-family audit

| Feature family | Status | Read | Write | Notes |
|----------------|--------|------|-------|-------|
| CP, PLC, STTB, SPRM/PRL, and property storage primitives | ✅ | ✅ | ✅ | [MS-DOC] 2.2 defines character positions, piece and property storage, string tables, and single-property modifiers used by the typed model |
| Piece table, compressed/uncompressed Unicode text, FKPs, BTEs, and BinTable | ✅ | ✅ | ✅ | Core text and formatting indices are decoded and generated with bounds checks and unknown property data retained where applicable |
| Main, footnote, endnote, header, comment, textbox, and header-textbox parts | ✅ | ✅ | ✅ | [MS-DOC] 2.3 story ranges and the corresponding PLCs are exposed through typed document, note, comment, header/footer, picture, shape, and textbox APIs |
| Bookmark PLCs and names | ✅ | ✅ | ✅ | Range and point bookmarks are typed and editable, including repair/validation behavior for malformed ranges |
| Field PLCs and non-Plcfld text-only fields | ✅ | ✅ | ✅ | Native field delimiters, instruction/result text, marker positions, nesting, and the five text-only field families are reconstructed and authored with balanced graphs |
| Field evaluation, document navigation, and generated results | ❌ | ❌ | ❌ | Cached results and inert typed instructions are available, but fields never resolve bookmarks, pages, styles, documents, databases, prompts, or host identity and are never refreshed |
| Tables, nested tables, cell/row marks, merges, and TAP properties | ✅ | ✅ | ✅ | [MS-DOC] 2.4 table-depth and terminating-mark rules are represented by rows/cells and typed table/paragraph properties |
| Character, paragraph, table, section, and picture SPRM families | ✅ | ✅ | ✅ | [MS-DOC] 2.6 property families map to typed formatting, page setup, borders, widths, shading, and picture options |
| Stylesheet, list templates, list overrides, names, and numbering | ✅ | ✅ | ✅ | STSH/LSTF/LFO/LVLF and related style/list records have typed reading and writer generation |
| Section layout, columns, borders, page/line numbering, and note placement | ✅ | ✅ | ✅ | Section properties are typed and serialized; pagination and line-breaking calculations are not performed |
| PLRSID, saved-by, associated strings, proofing, grammar, and spelling PLCs | ✅ | ✅ | ✅ | Bounded auxiliary tables and state are read and mutated with mandatory/optional emission rules where required |
| Auto-summary (Asumyi and PlcfAsumy) | 🟡 | ✅ | 🟡 | The auto-summary state/ranges have a typed read model; authoring is bounded and does not run Word’s summary algorithm |
| Captions and AutoCaption tables | 🟡 | ✅ | ❌ | Caption definitions, labels, locations, and chapter-numbering metadata are typed for inspection; automatic caption insertion and refresh are not implemented |
| Smart tags and factoids (Plcffactoid, property bags, and related STTBs) | ✅ | ✅ | ✅ | Bounded property-bag/factoid codecs and typed entries support CRUD; recognizers, schema downloads, VBA callbacks, and URLs remain inert |
| Master-document subdocuments (PlcfWKB, WKB, FNPI, SttbFnm) | 🟡 | ✅ | 🟡 | Directory, outline, referenced-file names, and metadata are typed; referenced files are stored verbatim and never opened or followed |
| Revision marks, author tables, and property revisions | ✅ | ✅ | ✅ | Insert/delete/move and paragraph, table, cell, row, and property revision records support transactional edit/accept/reject behavior |
| Document variables, attached-template and web-export metadata | 🟡 | ✅ | ✅ | Stored settings and relationship-like metadata are available where exposed; templates and web targets are not loaded and web layout is not generated |
| Mail-merge settings, sources, filters, and recipients (Pms, Pmfs, Rfs, ODSO) | 🟡 | ✅ | ✅ | Word 97 and Word 2002+ source descriptors, SQL/connection text, recipient filtering, sorting, inclusion, and field mappings are typed bounded metadata; data sources are never contacted and no merge is run |
| Legacy form fields and FFData | ✅ | ✅ | ✅ | Text, checkbox, dropdown, defaults, selections, help/status text, and verbatim entry/exit macro names are typed; macros and host form behavior are inert |
| ActiveX/OCX control semantics | ❌ | ❌ | ❌ | Inert [MS-DOC] OcxInfo/RgxOcxInfo metadata is available through `parts::ole_controls`; ObjectPool payloads still have no public control lifecycle, property, event, or rendering API |
| Document protection settings and range-level protected bookmarks | 🟡 | ✅ | ✅ | Protection modes, hashes, SttbfBkmkProt ranges, editor assignments, and usernames are typed; editing policy is not enforced |
| Document statistics and DOP version metadata | ✅ | ✅ | ✅ | Word/character/paragraph/line/page counts where stored, versioned DOP records, compatibility options, typography, macro-security metadata, and related state are exposed within bounded models |

## Drawings, equations, and payloads

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| OfficeArt/Escher drawing group and shape extraction | ✅ | ✅ | 🟡 | Data-stream OfficeArt, drawing-group anchors, shape properties, and text boxes are extracted; writer support covers main/header shapes and bounded shape kinds rather than every OfficeArt record |
| Pictures and BLIP payloads | ✅ | ✅ | ✅ | EMF, WMF, PICT, JPEG, PNG, DIB/BMP, and TIFF payload families are extracted and authored with OfficeArt identifiers and required metafile/BMP normalization |
| Floating shapes, textboxes, and header drawings | 🟡 | ✅ | 🟡 | Main/header anchors, text-box stories, pictures, positions, and selected presets are authored; full OfficeArt geometry, WordArt, rotation/z-order, and layout are not a complete model |
| Embedded OLE and package payloads | ✅ | ✅ | ✅ | Payload bytes and storage topology round-trip without activation or deserialization |
| MathType/MTEF Equation.3 objects | ✅ | ✅ | ✅ | Native equation extraction/conversion and bounded MTEF/Equation.3 authoring include registration streams, ObjInfo, preserved CLSIDs, and real image previews; layout/evaluation remain external |
| Embedded font programs | 🟡 | ✅ | 🟡 | Font-table records and embedded-font metadata are available; the writer does not claim a complete typed model for every embedded program and licensing behavior |
| Generic rendering and pagination | ❌ | ❌ | ❌ | The library edits document data and drawing payloads; it does not lay out pages, calculate page fields, rasterize OfficeArt, or produce Word-compatible visual output |

## Explicit unsupported or intentionally inert families

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Command bars, keymaps, menus, toolbars, and UI customizations | ❌ | ❌ | ❌ | [MS-DOC] 2.9 includes CTB, Customization, PlfMcd, PlfKme, PlfAcd, and related UI records, but no public semantic API exposes them |
| Routing slips and route-slip protection | 🟡 | ✅ | ✅ | Typed, lossless `RouteSlip`/`RouteSlipInfo` metadata is parsed and serialized through the FIB/table-stream seam; `Document` does not yet own the route-slip lifecycle or enforce its protection policy |
| Macro or control execution | ❌ | ❌ | ❌ | VBA source, macro names, form metadata, OLE objects, and control payloads are passive data only |
| External document/database/include resolution | ❌ | ❌ | ❌ | RD, include/link/DDE, mail-merge paths, SQL, connection strings, and referenced subdocuments are never opened, contacted, imported, or refreshed |
| Field calculation and generated tables/indexes | ❌ | ❌ | ❌ | TOC, TOA, INDEX, sequence, formula, style-reference, page, statistics, and navigation fields retain instructions/results but do not calculate or generate content |
| IRM license evaluation and protected-content access | ❌ | ❌ | ❌ | IRM metadata and protected streams are not decrypted, authorized, or evaluated |
| Complete later-version Word binary behavior | 🟡 | 🟡 | 🟡 | Versioned FIB/DOP structures are handled within the implemented profile, but unknown future records, application behavior, and undocumented compatibility quirks are not promised |

## Implementation map

- crates/litchi-doc/src/package.rs owns DOC opening, OLE access, encryption selection, properties, signatures, custom XML, and macro/storage access.
- crates/litchi-doc/src/document.rs, paragraph.rs, table.rs, section.rs, sprm.rs, parts/, and writer/ own CP/text, PLC, FKP, style/list, story, and formatting behavior.
- crates/litchi-doc/src/parts/fields.rs, revisions.rs, mail_merge.rs, route_slip/, structured_tags.rs, ole_controls.rs, smart_tags.rs, document_properties*.rs, and PLC helpers provide bounded metadata codecs.
- crates/litchi-doc/src/shapes.rs, image.rs, equation.rs, embedded_object.rs, and vba.rs own payload-oriented OfficeArt, picture, MTEF, OLE, and VBA support.
