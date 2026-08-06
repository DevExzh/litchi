# WordprocessingML (DOCX) Feature Matrix

This document tracks the public feature support of litchi-docx for WordprocessingML
packages (.docx, and the related .docm, .dotx, and .dotm package profiles where
the API explicitly exposes them). It is a compatibility matrix, not a claim of complete
ISO/IEC 29500 conformance or rendering fidelity.

## Scope, status, and specification references

| Mark | Meaning |
|------|---------|
| ✅ | Typed public support for the scope in Notes |
| 🟡 | Bounded, partial, metadata-only, pass-through, or inert support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to the format or direction |

Read and Write describe the public direction independently. A typed row can still be
inert: reading or writing field instructions, macros, links, or embedded payloads does not
evaluate, execute, fetch, or render them. Generic OPC part access and preservation of an
unmodified XML part are not counted as typed support for the XML feature inside that part.

The primary audit sources are [MS-DOCX] ToC.md, Front Matter.md, and 2 Structures:
2.1 Part Enumerations, 2.2 Extensions, 2.3 compatSetting elements, 2.4 numFmt
Extensions, and the Word 2010, 2012, 2015, 2016, 2018, 2020, 2023, and 2024 extension
namespace documents. The base WordprocessingML vocabulary is the ISO/IEC 29500 / ECMA-376
model referenced by [MS-DOCX].

Related package and payload references used for this matrix are [MS-CFB] and [MS-OLEPS]
for compound files and property sets, [MS-OFFCRYPTO] for OOXML encryption and IRM
DataSpaces, [MS-OWEXML] for web extensions, [MS-OREACTXML] for comment reactions,
[MS-OVBA] for VBA projects, and [MS-ODRAWXML] for WordprocessingML DrawingML.
The implementation owner is crates/litchi-docx/src/; shared OPC, signature, crypto,
embedded-object, web-extension, and common OOXML models are called out where they affect
DOCX behavior.

## Core WordprocessingML document model

## Shared specification map

The protocol ToCs and front matter under `3rdparty/specs/` are the audit input. A detailed matrix
must use the most specific format specification available and then account for shared dependencies.

| Family | Role in the audit |
|--------|-------------------|
| `[MS-DOC]`, `[MS-DOCX]` | Word binary and WordprocessingML feature families |
| `[MS-XLS]`, `[MS-XLSB]`, `[MS-XLSX]` | Excel BIFF, binary OOXML, and SpreadsheetML feature families |
| `[MS-PPT]`, `[MS-PPTX]` | PowerPoint binary and PresentationML feature families |
| `[MS-CFB]`, `[MS-OLEPS]`, `[MS-OLEDS]` | Compound File Binary, property sets, and embedded OLE data |
| `[MS-ODRAW]`, `[MS-ODRAWXML]`, `[MS-OGRAPH]`, `[MS-WMF]`, `[MS-EMF]`, `[MS-EMFPLUS]` | OfficeArt, DrawingML, graph, and drawing payloads |
| `[MS-OE376]`, `[MS-OI29500]`, `[MS-OWEXML]` | Shared OOXML relationships, package, and compatibility behavior |
| `[MS-OFFCRYPTO]` | Office encryption and password-protection envelopes |
| `[MS-OVBA]`, `[MS-VBAL]` | Macro project/module streams and VBA codec boundaries |
| `[MS-OSHARED]`, `[MS-DTYP]`, `[MS-LCID]`, `[MS-UCODEREF]` | Shared Office types, code pages, locale identifiers, and Unicode references |

The specifications describe what a producer or consumer may encounter. They do not by themselves
prove that Litchi implements a feature. Each row must be grounded in the public API and its tests,
with an honest boundary in `Notes`.

## Audit and maintenance rules

1. Keep one row per meaningful public feature family, not one row per protocol paragraph.
2. Include `Status`, `Read`, `Write`, and `Notes` columns in every detailed matrix.
3. Use `🟡` for bounded, metadata-only, inert, pass-through, or lossless-but-untyped behavior.
4. Record important protocol feature families that are not implemented as `❌`; do not silently omit
   them merely because the current API has no type for them.
5. State whether external targets, scripts, macros, database connections, embedded payloads,
   formulas, signatures, and encryption are resolved, executed, verified, or only preserved.
6. Keep package-level support separate from semantic support. A valid container or relationship
   graph does not imply support for every part carried by that graph.
7. When implementation changes, update the affected detailed matrix in the same change. Update
   this index only when ownership, shared semantics, or the supported format set changes.
8. Do not claim complete standards or application compatibility. Cite the relevant `[MS-*]` family
   or ISO/IEC structure in the detailed document's section or Notes cell when it helps delimit scope.

## Project-wide boundaries

Feature-gated families require the corresponding Cargo feature. Optional ODF, formula, RTF, iWork,
font, image-conversion, and related stacks are not implied by the default umbrella build. Public
APIs generally support in-memory and path-based workflows, but streaming, lazy loading, zero-copy
parsing, and rendering fidelity vary by crate. Typed errors and resource bounds are part of the
support claim where the detailed matrix says so.

The project treats untrusted Office and ODF content as data: external references are not fetched,
macros are not executed, and embedded objects are not activated. Formula evaluation and conversion
must describe their supported function/AST subset and any caller-provided capabilities. Signature
verification is integrity-oriented and does not establish certificate trust.

## Source map

- Shared package and signature infrastructure: `crates/litchi-opc/`, `crates/litchi-cfb/`, `crates/litchi-ole-common/`, `crates/litchi-sign/`, and `crates/litchi-crypto/`
- OOXML shared models and drawings: `crates/litchi-ooxml-common/`, `crates/litchi-drawingml/`, and `crates/litchi-fonts/`
- Legacy Office infrastructure: `crates/litchi-codepage/`, `crates/litchi-odraw/`, and `crates/litchi-ograph/`
- VBA project codec: `crates/litchi-vba/`
- OpenDocument shared models: `crates/litchi-odf/`, `crates/litchi-odf-common/`, and `crates/litchi-odf-formula/`
- Unified facades and conversion APIs: `crates/litchi/`, `crates/litchi-core/`, `crates/litchi-formula/`, `crates/litchi-eval/`, `crates/litchi-markdown/`, and `crates/litchi-sheet/`


## Package graph and shared Office features

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| OPC parts, relationships, and content types | ✅ | ✅ | ✅ | Package owns a validated OPC graph with transactional relationship and part edits; internal targets, content types, and graph cardinality are checked rather than treated as arbitrary ZIP entries |
| Transitional and Strict package forms | ✅ | ✅ | ✅ | DOCX package and WordprocessingML readers accept the supported Transitional and Strict namespaces; writers select the conformance form for owned parts, while unsupported strict-only extension semantics are not silently reinterpreted |
| Markup Compatibility (mc:Ignorable, AlternateContent) | 🟡 | ✅ | ✅ | MCE-aware readers select active branches for supported models and preserve bounded raw or opaque branches where possible; inactive or foreign content is not thereby semantically implemented |
| Main document, settings, styles, numbering, theme, font table, comments, notes, headers, footers, glossary, and properties parts | ✅ | ✅ | ✅ | The package graph covers the standard Word part families needed by the typed document model, including relationship and content-type maintenance; `glossary::{Snapshot, Transaction, Commit, Patch}` adds source-checked catalog and auxiliary-part CRUD with orphan-safe publication |
| Web settings part | ✅ | ✅ | ✅ | `web::{Snapshot, Transaction, Commit, Patch}` owns bounded frameset, layout, screen, and border edits with source-preserving opaque XML, graph/content-type validation, and atomic package publication; URLs and rendering remain inert |
| stylesWithEffects part | 🟡 | ✅ | 🟡 | [MS-DOCX] 2.1.1 part topology is recognized in glossary/package graph handling and the primary styles model is typed; the extended visual-effects vocabulary is not a complete independent style-effects model |
| Core, extended, and custom properties | ✅ | ✅ | ✅ | Bounded typed CRUD with deterministic package graph updates and preservation of unrelated parts |
| Embedded OLE and package objects | ✅ | ✅ | ✅ | Relationship inventory and inert object authoring preserve validated ProgIDs, embedding parts, content types, optional previews, and payload bytes; payloads and external targets are never opened |
| Embedded fonts | ✅ | ✅ | ✅ | Bounded Strict/Transitional font-table XML and OPC topology with typed add/replace/remove/reorder, obfuscation, licensing checks, and orphan-safe media cleanup; font programs remain inert |
| Images, charts, and DrawingML/VML relationships | ✅ | ✅ | ✅ | Resources and relationship graphs are authored and preserved; chart calculation, image rendering, and shape layout are outside the library |
| Digital signatures | ✅ | ✅ | ✅ | Trust-neutral OPC XMLDSig verification, signing, re-signing, and clearing; strict verification is the safe default and certificate trust/revocation is not established |
| Password encryption | ✅ | ✅ | ✅ | Supported Standard and Agile compatibility profiles are AES-128/SHA-1; unsupported AES-192/256 and SHA-2 Agile profiles are typed as unsupported rather than misread. Encrypted sources do not implicitly downgrade to plaintext on ordinary save |
| Information Rights Management and protected content | 🟡 | ✅ | 🟡 | [MS-OFFCRYPTO] DataSpaces, publishing-license, certificate, sensitivity-label, and integrity metadata have bounded shared codecs; rights services, license evaluation, protected-content decryption, and access enforcement are never performed |
| VBA projects in DOCM/DOTM | 🟡 | 🟡 | 🟡 | Bounded vbaProject.bin CFB/MS-OVBA project/module metadata and cache-free source payload authoring are available with package/content-type graph maintenance; VBA is never compiled or executed |
| Web extensions and Office Add-ins | 🟡 | ✅ | 🟡 | Typed task-pane/web-extension graph editing with bounded links, bindings, snapshots, extension lists, and shared payloads; callbacks, commands, remote services, and add-in code are inert |
| Mail-merge settings and recipient metadata | ✅ | ✅ | ✅ | `mail_merge::{Snapshot, Transaction, Commit, Patch}` provides bounded source-checked settings, recipient, and field-map edits with opaque XML preservation and reversible atomic publication; data sources are never contacted and no merge is run |
| External relationships and linked content | 🟡 | ✅ | ✅ | Relationship targets and link metadata can be inspected or authored, but Litchi does not fetch, resolve, open, refresh, or execute external targets |

## [MS-DOCX] part and extension audit

These rows make the extension specification explicit. A ❌ row means that the extension
family has no typed model even if an untouched part can remain in a package-preserving edit.

| Feature family | Status | Read | Write | Notes |
|----------------|--------|------|-------|-------|
| commentsExtended, people, commentsIds, and commentsExtensible parts | ✅ | ✅ | ✅ | [MS-DOCX] 2.1.2-2.1.5 are represented by bounded modern-comment metadata, people/presence, paragraph-to-durable-ID mappings, extension lists, UTC strings, and relationship/content-type maintenance |
| Modern comment reactions and extension lists | ✅ | ✅ | ✅ | commentsExtensible reaction metadata follows [MS-OREACTXML] as inert typed data; the library does not contact collaboration services or interpret reaction policy |
| rPr extension effects (glow, shadow, reflection, textOutline, textFill, scene3d, props3d) | ✅ | ✅ | ✅ | [MS-DOCX] 2.2.1 and 2.6 expose the seven Word 2010 visual effects through `Run::effects`, cached `RunProperties::effects`, and writer `MutableRun::{effects, effects_mut, set_effects}`. The typed model validates schema domains, preserves ordered bounded unknown direct children byte-for-byte, and does not render or evaluate effects |
| rPr OpenType extensions (ligatures, numForm, numSpacing, stylisticSets, cntxtAlts) | ✅ | ✅ | ✅ | The contextual `font::open_type::{Snapshot, Transaction, Commit, Patch}` owner exposes bounded Word 2010 run-property edits with absent-versus-authored `ST_OnOff` semantics, source-checked patches, and ordered unknown-child preservation; font shaping and rendering remain outside the crate |
| Settings extension elements (chartTrackingRefBased, docId, conflictMode, discardImageEditingData, defaultImageDpi) | ✅ | ✅ | ✅ | [MS-DOCX] 2.2.2 Word 2010/2012 settings extensions expose `settings::extensions::{Snapshot, Transaction, Commit, Patch}` with typed absent/default-on/explicit On/Off semantics, bounded MCE-aware source-preserving edits, ordered opaque direct-child preservation, stale/atomic validation, and malformed/round-trip coverage |
| Generic compatSetting records and compatibility mode | ✅ | ✅ | ✅ | Settings expose known compatibility flags, compatibilityMode, and bounded arbitrary name/URI/value records. This records compatibility intent; it does not implement Word layout algorithms |
| SDT extension controls (entityPicker, checkbox, repeatingSection, repeatingSectionItem, appearance, color, dataBinding) | ✅ | ✅ | ✅ | The content-control model recognizes and authors the supported extension markers and properties; host selection, repetition, entity lookup, and bound-data refresh are not executed |
| SDT web-extension links (webExtensionsLinked, webExtensionCreated) | 🟡 | 🟡 | 🟡 | Web-extension package graphs are typed separately, but these SDT-to-add-in behaviors are not a semantic host integration |
| Paragraph and table-row extension identifiers (paraId, textId) and paragraph noSpellErr | ✅ | ✅ | ✅ | `[MS-DOCX]` 2.2.4 and 2.6.2 are modeled by `paragraph::extensions::{Id, Ids, Extensions}`. Paragraphs and table rows expose checked eight-digit `ST_LongHexNumber` values; paragraph `noSpellErr` preserves absent versus explicit `0`/`1`, and writers emit local `w14` plus `mc:Ignorable` declarations. `textId` requires `paraId`; document-wide uniqueness policy, spelling detection, comments synchronization, and rendering remain outside this bounded owner |
| Conflict revision markup (conflictIns, conflictDel, and custom XML conflict ranges) | ❌ | ❌ | ❌ | [MS-DOCX] 2.2.5 conflict families are not represented by the ordinary tracked-revision model |
| anchorId on object and pict | ✅ | ✅ | ✅ | `[MS-DOCX]` 2.2.6 legacy-anchor inventories distinguish `w:object` and `w:pict` and validate the Word 2010 `AnchorId` range shared with `[MS-DOCX]` 2.6.2.1 and `[MS-ODRAWXML]` 2.18.2.1; `MutableOleObject` and `MutableVmlShape` author checked `w14:anchorId` values with a local `mc:Ignorable` declaration. OLE, VML, image payloads, layout, and rendering remain inert |
| Um Al-Qura umalqura calendar extension | ❌ | ❌ | ❌ | [MS-DOCX] 2.2.7 is not a typed calendar or date-rendering model |
| sectPr footnoteColumns | ✅ | ✅ | ✅ | `[MS-DOCX]` Word 2012 `w12:footnoteColumns` is exposed through `section::footnote_columns::{Layout, Snapshot, Transaction}` and `Section` read/edit helpers; absence and explicit zero remain distinct, inherited package namespaces and `mc:Ignorable` are retained out-of-band, value edits preserve authored attributes/lexical source, and patches are source-checked and atomic; layout/pagination are not implemented |
| pPr collapsed | ✅ | ✅ | ✅ | `[MS-DOCX]` 2.5.1.3 `w12:collapsed` is exposed as the typed `Collapsed` value with paragraph read/replace, authoring helpers, and atomic body-paragraph edits; explicit true/false remains distinct from absence, unknown `pPr` content is preserved, and display/layout behavior is not implemented |
| Numbering restartNumberingAfterBreak | ✅ | ✅ | ✅ | `[MS-DOCX]` Word 2012 `w12:restartNumberingAfterBreak` is exposed on contextual numbering definitions with namespace-aware parsing, absent-versus-explicit state, bounded source-preserving transactions, and package publication; numbering behavior itself remains inert |
| Run symEx symbol extension | ✅ | ✅ | ✅ | [MS-DOCX] 2.2.11 and 2.7 expose ordered `run_symbols::{Symbol, Symbols}` values, optional `font`/Unicode-scalar `char` state, bounded validation, source-preserving snapshots/transactions/patches, raw `Run` edit helpers, and ergonomic mutable-run authoring; unknown run content remains inert and byte-preserved |
| Data-binding storeItemChecksum | ❌ | ❌ | ❌ | Custom XML bindings and integrity relationships are typed, but the Word 2020 checksum extension is not calculated or validated |
| Word 2023 dateUtc and Word 2024 formattingAllowed extensions | 🟡 | 🟡 | 🟡 | Modern comment metadata carries bounded UTC strings; the 2024 SDT formatting-lock extension is not part of the content-control model |
| Extended note/list number formats from [MS-DOCX] 2.4 | 🟡 | ✅ | 🟡 | Common note and numbering formats have typed enums and serialization; the full Microsoft extension enumeration is not a layout/numbering engine |

## Explicit semantic gaps

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Pagination, line breaking, page fields, and rendering | ❌ | ❌ | ❌ | Section/page properties are data models only; there is no Word layout engine or renderer |
| Field calculation and refresh | ❌ | ❌ | ❌ | Field delimiters, instructions, switches, and cached results are typed or preserved, but fields are never recalculated, refreshed, or resolved |
| Mail-merge execution | ❌ | ❌ | ❌ | Mail-merge settings, recipients, filters, and source targets are typed bounded metadata; no data source is opened or contacted and no merge output is generated |
| AltChunk import and foreign-content conversion | ❌ | ❌ | ❌ | AltChunk anchors and payloads have bounded opaque CRUD, but HTML, XHTML, RTF, text, XML, and MIME payloads are never imported or rendered |
| Macro, ActiveX, form-control, and add-in execution | ❌ | ❌ | ❌ | VBA projects, OLE payloads, legacy form state, web extensions, and control metadata are inert; code, callbacks, events, and host UI are never executed |
| Chart, SmartArt, DrawingML, VML, and embedded-workbook rendering | ❌ | ❌ | ❌ | Graphs and definition parts are typed or preserved, but geometry, chart calculation, SmartArt layout, and rasterization are renderer responsibilities |
| Protection enforcement | 🟡 | ✅ | ✅ | Protection modes, hashes/settings, and document state can be read or written; the library does not prevent edits or enforce a Word editing policy |
| Signature trust and certificate status | 🟡 | ✅ | ✅ | Integrity/signature operations are available, but certificate-chain trust, revocation, identity, and enterprise policy are outside the API |
| Unsupported Word extension namespace content | 🟡 | 🟡 | 🟡 | Unknown extension parts or XML may survive bounded package-preserving operations, but no semantic behavior or round-trip guarantee is claimed for unmodeled markup |

## Implementation map

- crates/litchi-docx/src/package.rs owns package graphs, parts, relationships, properties, mail merge, custom XML, glossary, web settings, encryption, and save modes.
- crates/litchi-docx/src/document.rs, paragraph.rs, table.rs, section.rs, styles.rs, and numbering/ own the main typed WordprocessingML model.
- crates/litchi-docx/src/field/, content_control.rs, revision.rs, comment.rs, modern_comments/, custom_xml.rs, alt/, drawing/, textbox.rs, chart.rs, smartart.rs, math.rs, font/, and writer/ provide feature-specific APIs and bounded codecs.
