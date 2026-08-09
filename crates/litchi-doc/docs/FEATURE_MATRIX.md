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

## Read limits and defaults

`Package::open`, `Package::from_reader`, and `Package::from_ole_file` use
`Limits::default()`: 128 MiB for the outer CFB package, 64 MiB for each
DOC-owned input stream, and 96 MiB for the aggregate of `WordDocument`, the
selected table stream, and optional `Data` stream. Separate 1 GiB/512 MiB/768
MiB hard ceilings bound explicitly requested values accepted by
`Limits::try_new` and its checked builder methods.

`Package::open` uses those finite defaults; `Package::open_with` accepts a
`PackageOpenOptions` value for ergonomic explicit limits without replacing the
minimal API. `Package::{open,from_reader,from_ole_file}_with_limits` remains
available for direct limit passing. `document_with_limits` and
`document_with_options_and_limits` combine their supplied limits component by
component with those package limits, retaining the stricter value in each
dimension. Limit failures are structured `Error::ResourceLimit` values with
the exceeded resource kind, observed size, configured limit, and a stream path
when applicable. These limits govern package and main DOC-stream ingestion;
individual typed codecs can impose additional format-specific bounds.

Password-to-open input is owned by non-cloneable, zeroizing `Password` and is
supplied through `OpenOptions::with_password`; it is redacted in diagnostics.
`Document::text` preserves stored source text. `body_text::Snapshot` also
projects ordinary main-story paragraphs as stored, accepted, or rejected text.
Its bounded source-checked transaction supports length-changing replacements
across multiple ordinary paragraphs and direct-bold changes atomically. It
appends Unicode pieces, rebuilds CLX/CHPX data, shifts modeled main/all-story CP
tables, updates the main-story count, and performs both CFB and full DOC reopen
validation before publication. Structural or tracked text, mixed character
formatting, interior modeled CP boundaries, and known unmodeled CP-indexed structures
are typed refusals. The same semantic changes drive source-checked reversible
patches, deterministic durable patches, disjoint composition, and bounded
undo/redo history. Paragraph selection uses the format-neutral
`litchi_core::Position`, with collection resolution returning a typed not-found
refusal.

## Core Word binary document model

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| `MsoEnvelopeCLSID` / `MsoEnvelope` metadata | ✅ | ✅ | ✅ | `parts::envelope` owns the `[MS-DOC]` `fcMsoEnvelope` FIB range and the documented `[MS-OSHARED]` Office 6/8 body, recipient property bags, and attachments. Unknown CLSIDs and supported-body tail bytes are preserved as bounded opaque data; `Editor` stages clone-first package/FIB edits and publishes source-checked reversible semantic/CFB patches. `Document::envelope()` remains inert: no mail transport, recipient resolution, attachment activation, or external behavior. |


## [MS-DOC] file structure and version audit

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| OLE2/CFB container and Word streams | ✅ | ✅ | ✅ | Package validates the WordDocument stream and the 0Table or 1Table stream, exposes bounded OLE editing, and writes new or transactionally edited packages |
| Word 97+ FIB and versioned FibRgFcLcb tables | ✅ | ✅ | ✅ | FIB base/version fields, table-stream pointers, Word 97 through later extension ranges, fixed-layout older records, and declared pointer bounds are handled by the DOC reader/writer |
| Versioned document properties (DopBase through Dop2013) | ✅ | ✅ | ✅ | `Document::document_properties()` exposes the deferred validated DOP record; `VersionedDocumentProperties` covers Word 95/97/2000/2002/2003/2007/2010/2013, including typed lossless `DopMth` equation defaults, image-save settings, paragraph-ID context, and chart tracking. Fixed/reserved fields and specification value domains are checked, undefined bytes/bits are preserved, and typed extension codecs write back without normalizing older generations |
| Word 6/95 and pre-Microsoft-Office binary profiles | ❌ | ❌ | ❌ | The public reader intentionally rejects files predating the Word 97 binary profile rather than treating incompatible table layouts as DOC |
| WordDocument, table, Data, and story ranges | ✅ | ✅ | ✅ | Piece-table text, formatting tables, auxiliary Data content, and story character counts are exposed through typed document and writer APIs |
| Custom XML data storage | ✅ | ✅ | ✅ | Bounded, lossless MsoDataStore item/property XML with typed GUIDs, schema references, IRM promotion markers, and inert schema URIs; shared `litchi_ole_common::custom_xml::{Snapshot, Transaction, Commit, Patch}` supplies the layered source-checked owner |
| Summary and document-summary property streams | ✅ | ✅ | ✅ | OLE property-set reading and editing with typed summary, document-summary, and user-defined property access |
| Reserved user-defined hyperlink metadata | ✅ | ✅ | ✅ | `[MS-OSHARED]` §§2.3.3.1.18-21 and §2.4.2 define the typed `VtHyperlink`/vector/blob forms and hash used by `_PID_LINKBASE` and `_PID_HLINKS`. They are named values in the `UserDefinedProperties` section of `DocumentSummaryInformation`, never PIDDSI `0x15`. Shared `Properties` lazily decodes only the named value requested, while shared `Edit` applies caller-bounded typed writes for blob size, link count, and UTF-16 strings; these bounds apply to the typed overlay and secondary decoding after generic property-set parsing, not initial property-stream allocation. `Package::user_defined_hyperlinks()` exposes possible `FieldCandidate` values rather than asserting an association: `dwApp` can be an `FcCompressed` or collide across the seven field stories. The caller must select an exact candidate with `resolve_field` before a changed write may canonically order entries per `[MS-DOC]` §2.4.7: Main, Footnote, Header, Comment, Endnote, Textbox, HeaderTextbox, then decreasing `aFld` index. An unresolved changed write refuses; an exact no-op remains preservable. OfficeArt, direct-picture, and ambiguous data remain inert and source-relative. This is an existing-artifact property-set transaction, not a claim that fresh-document generation automatically emits the properties. Targets and locations remain raw inert strings: Litchi never parses, normalizes, resolves, fetches, opens, or executes them. |
| Encryption stream and password-to-open profiles | ✅ | ✅ | ✅ | The supported DOC encryption profiles cover the implemented XOR/RC4-compatible modes and typed password errors; unsupported profiles fail explicitly |
| XML signatures storage and signatures stream | ✅ | ✅ | ✅ | Trust-neutral CFB signature verification and transactional signing/editing are exposed; certificate trust and revocation are not established |
| ObjectPool and embedded OLE/package objects | ✅ | ✅ | ✅ | `embedded_object::transaction::{Snapshot, Transaction, Commit, Patch}` provides source-checked, atomic add/remove/reorder/replace operations plus typed ODT/OLEDS metadata edits; unknown streams and inert payloads are preserved and objects are never activated or opened |
| Macros storage and VBA project payload | ✅ | ✅ | ✅ | Bounded MS-OVBA compressed-container, dir, PROJECT, module, and cache-free source metadata are read and authored; code is never compiled or executed |
| Information Rights Management Data Space and protected content | 🟡 | ✅ | ❌ | `Package::data_spaces` exposes the validated, inert `litchi_doc::spaces::Graph` from the shared MS-OFFCRYPTO owner, including legacy-binary IRM transform/license topology, labels, and integrity sidecars; rights evaluation, decryption, and protected-content access remain outside the API |
| Unknown streams and storages | 🟡 | ✅ | 🟡 | Package-preserving editors can retain unrelated CFB topology where supported, but an arbitrary stream has no DOC semantic model or write guarantee |

## [MS-DOC] PLC, property, and auxiliary-family audit

| Feature family | Status | Read | Write | Notes |
|----------------|--------|------|-------|-------|
| CP, PLC, STTB, SPRM/PRL, and property storage primitives | ✅ | ✅ | ✅ | [MS-DOC] 2.2 defines character positions, piece and property storage, string tables, and single-property modifiers used by the typed model |
| Piece table, compressed/uncompressed Unicode text, FKPs, BTEs, and BinTable | ✅ | ✅ | ✅ | Core text and formatting indices are decoded and generated with bounds checks and unknown property data retained where applicable. `body_text::{Snapshot, Edit, Commit, Patch}` adds bounded multi-paragraph length-changing text plus direct-bold transactions, modeled CLX/CHPX/PLCF/FIB updates, full reopen validation, reversible and durable patches, disjoint composition, and bounded history; unmodeled dependencies are refused before mutation. |
| Main, footnote, endnote, header, comment, textbox, and header-textbox parts | ✅ | ✅ | ✅ | [MS-DOC] 2.3 story ranges and the corresponding PLCs are exposed through typed document, note, comment, header/footer, picture, shape, and textbox APIs |
| Bookmark PLCs and names | ✅ | ✅ | ✅ | Range and point bookmarks are typed and editable, including repair/validation behavior for malformed ranges |
| Field PLCs and non-Plcfld text-only fields | ✅ | ✅ | ✅ | Native field delimiters, instruction/result text, marker positions, nesting, and the five text-only field families are reconstructed and authored with balanced graphs |
| Field evaluation, document navigation, and generated results | ❌ | ❌ | ❌ | Cached results and inert typed instructions are available, but fields never resolve bookmarks, pages, styles, documents, databases, prompts, or host identity and are never refreshed |
| Tables, nested tables, cell/row marks, merges, and TAP properties | ✅ | ✅ | ✅ | [MS-DOC] 2.4 table-depth and terminating-mark rules are represented by rows/cells and typed table/paragraph properties |
| Character, paragraph, table, section, and picture SPRM families | ✅ | ✅ | ✅ | [MS-DOC] 2.6 property families map to typed formatting, page setup, borders, widths, shading, and picture options |
| Stylesheet, list templates, list overrides, names, and numbering | ✅ | ✅ | ✅ | STSH/LSTF/LFO/LVLF and related style/list records have typed reading and writer generation |
| Section layout, columns, borders, page/line numbering, and note placement | ✅ | ✅ | ✅ | Section properties are typed and serialized; pagination and line-breaking calculations are not performed |
| PLRSID, saved-by, associated strings, proofing, grammar, and spelling PLCs | ✅ | ✅ | ✅ | Bounded auxiliary tables and state are read and mutated with mandatory/optional emission rules where required |
| Auto-summary (Asumyi and PlcfAsumy) | ✅ | ✅ | ✅ | Typed summary state and checked `PlcfAsumy` ranges support bounded CRUD and exact PLC authoring; Word’s summary algorithm remains inert |
| Captions and AutoCaption tables | ✅ | ✅ | ✅ | Contextual `captions::{Tables, Editor, Snapshot, Transaction}` owns bounded FIB/table-stream codecs, typed labels, numbering/location options, validated ProgID-to-label references, and failure-atomic package CRUD; edits append new ranges and clear only FIB pointers, preserving unrelated table bytes. Word caption insertion, field refresh, host activation, and macro execution remain inert |
| Smart tags and factoids (Plcffactoid, property bags, and related STTBs) | ✅ | ✅ | ✅ | `parts::smart_tags::{Snapshot, Transaction, Commit, Patch}` adds source-checked bookmark, property-bag, and recognizer-range edits while preserving opaque FIB/table bytes; recognizers, schema downloads, VBA callbacks, and URLs remain inert |
| Master-document subdocuments (PlcfWKB, WKB, FNPI, SttbFnm) | 🟡 | ✅ | 🟡 | `parts::subdocuments::{Snapshot, Transaction, TablePatch}` owns source-checked, failure-atomic semantic edits, exact table encoders, and package publication with checked FIB pointer relocation for `PlcfWKB`/`SttbFnm`; unrelated table bytes and undefined FNIF/WKB bits remain preserved. Referenced files remain inert and are never opened or followed |
| Revision marks, author tables, and property revisions | ✅ | ✅ | ✅ | Insert/delete/move and paragraph, table, cell, row, and property revision records support transactional edit/accept/reject behavior; `tracked_revision::{Snapshot, Transaction, Commit, Patch}` adds source-checked binary story edits with reversible snapshots |
| Document variables, attached-template and web-export metadata | 🟡 | ✅ | ✅ | Stored settings and relationship-like metadata are available where exposed; templates and web targets are not loaded and web layout is not generated |
| Mail-merge settings, sources, filters, and recipients (Pms, Pmfs, Rfs, ODSO) | 🟡 | ✅ | ✅ | Word 97 and Word 2002+ source descriptors, SQL/connection text, recipient filtering, sorting, inclusion, and field mappings are typed bounded metadata; data sources are never contacted and no merge is run |
| Legacy form fields and FFData | ✅ | ✅ | ✅ | Text, checkbox, dropdown, defaults, selections, help/status text, and verbatim entry/exit macro names are typed; macros and host form behavior are inert |
| ActiveX/OCX control semantics | 🟡 | ✅ | ✅ | `litchi_doc::ole_controls` exposes inert, typed [MS-DOC] OcxInfo/RgxOcxInfo, story-specific `ifld` validation, ObjectPool `ODT`/`ODTPersist1`/`ODTPersist2`, exact `ObjInfo`/`OCXDATA`/presentation-stream metadata, and snapshot editing; CFB payloads, lifecycle, properties, events, rendering, and activation remain intentionally absent |
| Document protection settings and range-level protected bookmarks | 🟡 | ✅ | ✅ | Protection modes, hashes, SttbfBkmkProt ranges, editor assignments, and usernames are typed; editing policy is not enforced |
| Document statistics and DOP version metadata | ✅ | ✅ | ✅ | Word/character/paragraph/line/page counts where stored, versioned DOP records, compatibility options, typography, macro-security metadata, and related state are exposed within bounded models |

## Drawings, equations, and payloads

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| OfficeArt/Escher drawing group and shape extraction | ✅ | ✅ | 🟡 | Data-stream OfficeArt, drawing-group anchors, shape properties, and text boxes are extracted; writer support covers main/header shapes and bounded shape kinds rather than every OfficeArt record |
| Pictures and BLIP payloads | ✅ | ✅ | ✅ | EMF, WMF, PICT, JPEG, PNG, DIB/BMP, and TIFF payload families are extracted and authored with OfficeArt identifiers and required metafile/BMP normalization |
| Floating shapes, textboxes, and header drawings | 🟡 | ✅ | 🟡 | Main/header anchors, text-box stories, pictures, positions, and selected presets are authored; full OfficeArt geometry, WordArt, rotation/z-order, and layout are not a complete model |
| Embedded OLE and package payloads | ✅ | ✅ | ✅ | Payload bytes, storage topology, and bounded typed `ObjInfo`/`ODTPersist2` metadata round-trip through the source-checked embedded-object transactions without OLE activation or payload deserialization |
| MathType/MTEF Equation.3 objects | ✅ | ✅ | ✅ | Native equation extraction/conversion and bounded MTEF/Equation.3 authoring include registration streams, ObjInfo, preserved CLSIDs, and real image previews; layout/evaluation remain external |
| Embedded font programs | 🟡 | ✅ | 🟡 | Font-table records and embedded-font metadata are available; the writer does not claim a complete typed model for every embedded program and licensing behavior |
| Generic rendering and pagination | ❌ | ❌ | ❌ | The library edits document data and drawing payloads; it does not lay out pages, calculate page fields, rasterize OfficeArt, or produce Word-compatible visual output |

## Explicit unsupported or intentionally inert families

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Command bars, keymaps, menus, toolbars, and UI customizations | 🟡 | ✅ | ✅ | Bounded `CommandBars` metadata covers the FIB `Tcg` seam, inert `PlfMcd`/`PlfAcd`/`PlfKme` records, typed shared `TBCGeneralInfo`/`TBCExtraInfo`, and lossless CTBWRAPPER/TBC tails; shared `toolbar::{Snapshot, Transaction, Commit, Patch}` edits text, icons, and flags without executing macros/UI, while ambiguous variable boundaries and unknown Tcg records are refused |
| Routing slips and route-slip protection | 🟡 | ✅ | ✅ | `litchi_doc::route_slip` owns typed, lossless `RouteSlip`/`RouteSlipInfo` metadata, exact narrow bytes, checked recipient selectors, immutable snapshots, and transactional package edits; `Document::route_slip()` exposes deferred optional metadata, while route protection rejects lifecycle edits unless the policy is `Off`; mail transport, authentication, and host routing remain inert |
| Macro or control execution | ❌ | ❌ | ❌ | VBA source, macro names, form metadata, OLE objects, and control payloads are passive data only |
| External document/database/include resolution | ❌ | ❌ | ❌ | RD, include/link/DDE, mail-merge paths, SQL, connection strings, and referenced subdocuments are never opened, contacted, imported, or refreshed |
| Field calculation and generated tables/indexes | ❌ | ❌ | ❌ | TOC, TOA, INDEX, sequence, formula, style-reference, page, statistics, and navigation fields retain instructions/results but do not calculate or generate content |
| IRM license evaluation and protected-content access | ❌ | ❌ | ❌ | IRM metadata and protected streams are not decrypted, authorized, or evaluated |
| Complete later-version Word binary behavior | 🟡 | 🟡 | 🟡 | Versioned FIB/DOP structures are handled within the implemented profile, but unknown future records, application behavior, and undocumented compatibility quirks are not promised |

## Implementation map

- `src/package/{codec,model,property_set}.rs` owns DOC package opening, OLE access, package limits, encryption selection, properties, signatures, custom XML, and macro/storage access.
- `src/document/{codec,mod}.rs`, `src/paragraph/`, `src/table.rs`, `src/section/`, `src/sprm.rs`, `src/sprm_operations/`, `src/parts/`, and `src/writer/` own CP/text, PLC, FKP, style/list, story, and formatting behavior.
- The contextual metadata codecs live under `src/parts/`, with the corresponding facade modules in `src/lib.rs`; `src/route_slip.rs` layers route-slip codecs, validation, package editors, snapshot transactions, and its contextual facade. Envelope metadata is implemented by `src/parts/envelope/` over its FIB owner.
- `src/shape/`, `src/image.rs`, `src/equation.rs`, `src/embedded_object/`, and `src/vba.rs` own payload-oriented OfficeArt, picture, MTEF, OLE, and VBA support.
