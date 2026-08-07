# Office File Format Feature Matrix

This document is the index and shared guideline for the format-specific feature matrices. It
intentionally does not duplicate the detailed feature rows. The authoritative, format-specific
claims live in the linked documents under the corresponding crate.

The matrices describe public Litchi API support, not rendering fidelity or complete conformance
with every revision of a Microsoft protocol or OpenDocument standard. A protocol feature can be
listed as unsupported even when the underlying package can preserve an opaque part. Conversely,
bounded metadata or pass-through support must be marked partial rather than presented as semantic
read/write support.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | The documented feature scope has a public typed implementation |
| 🟡 | Support is bounded, partial, metadata-only, pass-through, or otherwise limited |
| ❌ | No public typed support is currently available |
| N/A | The concept does not apply to the format or direction |

`Read` and `Write` are independent directions. A `🟡` direction can mean a subset of the model,
lossless preservation without semantic access, or an inert serializer. Notes are normative for the
scope of each row. Cryptographic verification means integrity/signature verification only; it does
not establish certificate trust or revocation status. Macros, external links, database commands,
mail-merge sources, and embedded payloads remain inert unless a detailed matrix explicitly says
otherwise.

## Detailed matrices

| Crate | Format family | Primary protocol/specification families | Detailed matrix |
|-------|---------------|------------------------------------------|------------------|
| `litchi-doc` | Word binary (`.doc`) | [MS-DOC], [MS-CFB], [MS-OLEPS], [MS-ODRAW], [MS-OSHARED] | [Word binary feature matrix](../crates/litchi-doc/docs/FEATURE_MATRIX.md) |
| `litchi-docx` | WordprocessingML (`.docx`, related OOXML packages) | [MS-DOCX], [MS-OE376], [MS-OI29500], [MS-OWEXML], [MS-ODRAWXML], [MS-OFFCRYPTO], [MS-OVBA] | [WordprocessingML feature matrix](../crates/litchi-docx/docs/FEATURE_MATRIX.md) |
| `litchi-xls` | Excel BIFF (`.xls`) | [MS-XLS], [MS-CFB], [MS-OLEPS], [MS-ODRAW], [MS-OSHARED] | [Excel BIFF feature matrix](../crates/litchi-xls/docs/FEATURE_MATRIX.md) |
| `litchi-xlsb` | Excel Binary Workbook (`.xlsb`) | [MS-XLSB], [MS-OE376], [MS-OI29500], [MS-OWEXML], [MS-ODRAWXML], [MS-OFFCRYPTO], [MS-OVBA] | [Excel binary OOXML feature matrix](../crates/litchi-xlsb/docs/FEATURE_MATRIX.md) |
| `litchi-xlsx` | SpreadsheetML (`.xlsx`, related OOXML packages) | [MS-XLSX], [MS-OE376], [MS-OI29500], [MS-OWEXML], [MS-ODRAWXML], [MS-OFFCRYPTO], [MS-OVBA] | [SpreadsheetML feature matrix](../crates/litchi-xlsx/docs/FEATURE_MATRIX.md) |
| `litchi-ppt` | PowerPoint binary (`.ppt`) | [MS-PPT], [MS-CFB], [MS-OLEPS], [MS-ODRAW], [MS-OGRAPH], [MS-OSHARED] | [PowerPoint binary feature matrix](../crates/litchi-ppt/docs/FEATURE_MATRIX.md) |
| `litchi-pptx` | PresentationML (`.pptx`, related OOXML packages) | [MS-PPTX], [MS-OE376], [MS-OI29500], [MS-ODRAWXML], [MS-OFFCRYPTO], [MS-OVBA] | [PresentationML feature matrix](../crates/litchi-pptx/docs/FEATURE_MATRIX.md) |
| `litchi-odt` | OpenDocument text (`.odt`, `.ott`) | ISO/IEC 26300 structures and shared ODF package/style models | [OpenDocument text feature matrix](../crates/litchi-odt/docs/FEATURE_MATRIX.md) |
| `litchi-ods` | OpenDocument spreadsheet (`.ods`, `.ots`) | ISO/IEC 26300 structures and shared ODF package/style/formula models | [OpenDocument spreadsheet feature matrix](../crates/litchi-ods/docs/FEATURE_MATRIX.md) |
| `litchi-odp` | OpenDocument presentation (`.odp`, `.otp`) | ISO/IEC 26300 structures and shared ODF package/style/drawing models | [OpenDocument presentation feature matrix](../crates/litchi-odp/docs/FEATURE_MATRIX.md) |

The detailed matrices are the source of truth for per-format rows. Shared ODF implementation lives
primarily in `litchi-odf` and its companion crates; shared Office package and drawing behavior lives
in the crates named by the specification map below.

iWork parsing is split by concrete application format: `litchi-pages` owns
Pages (`.pages`), `litchi-keynote` owns Keynote (`.key`), and `litchi-numbers`
owns Numbers (`.numbers`). Their shared IWA archive, protocol, and package
layers do not constitute a fourth user-facing format. At the `litchi` facade,
the `pages`, `keynote`, and `numbers` feature leaves enable those parsers
independently; `iwork` is only the aggregate of all three.

## Cross-format capability index

This index records capability families that span more than one detailed matrix. The linked matrices
give the authoritative status and limitations for each concrete format.

| Capability family | Overall scope | Primary owners |
|--------------------|---------------|----------------|
| Format detection and unified facades | Office, ODF, RTF, iWork, and tabular APIs where exposed; facade authoring is not implied | `litchi`, `litchi-core`, detailed format crates |
| OOXML OPC package editing | Parts, relationships, content types, strict/transitional XML, and graph validation. Ingestion has configurable `ReadLimits`; default ceilings protect package input, ZIP members/names/metadata/compressed and uncompressed sizes, materialized parts, `[Content_Types].xml`, and relationship XML, attributes, targets, events, depth, and graph traversal. These are safety policy rather than spec maxima, grounded in ECMA-376 Part 2 §7.3.6/§10 and [MS-OI29500] §2.1.1749-1752. | `litchi-opc`, `litchi-ooxml-common`, DOCX/XLSX/XLSB/PPTX matrices |
| OLE/CFB package editing | Streams, storages, property sets, inert OLEDS links, directory catalogs, custom XML stores, VBA signature metadata, and package-preserving editors | `litchi-cfb`, `litchi-ole-common`, DOC/XLS/PPT matrices; `object::{directory,link}`, `custom_xml`, and `vba_signature` provide bounded source-checked transactions with reversible patches and opaque-tail preservation |
| ODF package and flat XML editing | ZIP package, manifest, metadata, styles, settings, resources, and bounded flat-document handling | `litchi-odf` and ODF matrices |
| Encryption and integrity | Format-specific Office and ODF profiles; bounded password and integrity handling | `litchi-crypto`, `litchi-sign`, format matrices |
| Properties and metadata | Core, extended, custom, document, and package metadata according to each format. PPTX/XLSX custom properties use typed, inert `litchi_ooxml_common::custom::Props`/`Value` parsing and writing per vendored `[MS-OI29500]` §3.11. Legacy Office reserves `_PID_LINKBASE` and `_PID_HLINKS` as named values in the `UserDefinedProperties` section, not PIDDSI `0x15`: `litchi_ole_common::property_set::user_defined::{Properties, Edit}` lazily decodes and boundedly writes their `[MS-OSHARED]` §§2.3.3.1.18-21 / §2.4.2 wire forms. Those limits bound the typed overlay and secondary decoding after generic property-set parsing, not initial property-stream allocation. Only DOC exposes possible field candidates under `[MS-DOC]` §2.4.7 because `dwApp` can be an `FcCompressed` or collide across stories; callers must resolve a candidate explicitly before canonical reordering or a changed write. PPT/XLS do not claim contextual association. These strings are preserved as inert text and are never normalized, resolved, fetched, opened, or executed. | `litchi-ooxml-common`, `litchi-ole-common`, and format matrices |
| Drawing, media, and embedded objects | Typed support varies by format; opaque preservation is not semantic rendering | `litchi-drawingml`, `litchi-odraw`, `litchi-ograph`, format matrices |
| Formula and equation conversion/evaluation | Shared formula/equation infrastructure with intentionally incomplete host semantics | `litchi-formula`, `litchi-eval`, ODF/XLS/XLSB/XLSX matrices |
| VBA and macro-enabled packages | Bounded codepage-aware project/module metadata and preservation; VBA is never compiled, interpreted, or executed | `litchi-vba`, OOXML matrices |
| Conversion and interchange | Markdown, images, RTF, iWork, CSV/TSV and other text-workbook APIs are separate from the ten matrices | `litchi-markdown`, `litchi-imgconv`, `litchi-rtf`, `litchi-pages`, `litchi-keynote`, `litchi-numbers`, `litchi-sheet` |

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

Feature-gated families require the corresponding Cargo feature. Optional ODF, formula, RTF, iWork
format leaves, font, image-conversion, and related stacks are not implied by the default umbrella build. Public
APIs generally support in-memory and path-based workflows, but streaming, lazy loading, zero-copy
parsing, and rendering fidelity vary by crate. Typed errors and resource bounds are part of the
support claim where the detailed matrix says so.

The project treats untrusted Office and ODF content as data: external references are not fetched,
and macros, VBA, ActiveX, controls, OLE objects, and embedded code are never executed or activated.
Those payloads may only be retained, inspected, validated, inventoried, or edited as inert blobs.
Actions, formulas, and links are likewise never executed as document content. Formula evaluation and conversion
must describe their supported function/AST subset and any caller-provided capabilities. Signature
verification is integrity-oriented and does not establish certificate trust.

## Source map

- Shared package and signature infrastructure: `crates/litchi-opc/`, `crates/litchi-cfb/`, `crates/litchi-ole-common/`, `crates/litchi-sign/`, and `crates/litchi-crypto/`
- OOXML shared models and drawings: `crates/litchi-ooxml-common/`, `crates/litchi-drawingml/`, and `crates/litchi-fonts/`
- Legacy Office infrastructure: `crates/litchi-codepage/`, `crates/litchi-odraw/`, and `crates/litchi-ograph/`
- VBA project codec: `crates/litchi-vba/`
- OpenDocument shared models: `crates/litchi-odf/`, `crates/litchi-odf-common/`, and `crates/litchi-odf-formula/`
- iWork format owners and shared IWA layers: `crates/litchi-pages/`, `crates/litchi-keynote/`, `crates/litchi-numbers/`, and `crates/litchi-iwa-*/`
- Unified facades and conversion APIs: `crates/litchi/`, `crates/litchi-core/`, `crates/litchi-formula/`, `crates/litchi-eval/`, `crates/litchi-markdown/`, and `crates/litchi-sheet/`
