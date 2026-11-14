# Document Format Feature Matrix

Index of the per-format feature checklists. Each linked matrix lists features and their
support status for the public Litchi API of one format crate. The matrices describe API support,
not rendering fidelity or complete conformance with any specification.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported |
| 🟡 | Partial: a subset, metadata-only, preservation-only, or otherwise limited |
| ❌ | Unsupported |
| N/A | The concept does not apply to the format or direction |

`Read` and `Write` are independent directions. Every format below also reads and writes through
the unified `litchi::File` API.

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
| `litchi-rtf` | Rich Text Format (`.rtf`) | RTF 1.9.1, [MS-OXRTFCP] | [Rich Text Format feature matrix](../crates/litchi-rtf/docs/FEATURE_MATRIX.md) |
| `litchi-pages` | Apple Pages (`.pages`) | Apple IWA package plus Pages, text, table, and shared schema projections | [Pages feature matrix](../crates/litchi-pages/docs/FEATURE_MATRIX.md) |
| `litchi-keynote` | Apple Keynote (`.key`) | Apple IWA package plus Keynote, presentation, text, table, chart, and shared schema projections | [Keynote feature matrix](../crates/litchi-keynote/docs/FEATURE_MATRIX.md) |
| `litchi-numbers` | Apple Numbers (`.numbers`) | Apple IWA package plus Numbers, table, formula, chart, and shared schema projections | [Numbers feature matrix](../crates/litchi-numbers/docs/FEATURE_MATRIX.md) |
| `litchi-odt` | OpenDocument text (`.odt`, `.ott`) | ISO/IEC 26300 structures and shared ODF package/style models | [OpenDocument text feature matrix](../crates/litchi-odt/docs/FEATURE_MATRIX.md) |
| `litchi-ods` | OpenDocument spreadsheet (`.ods`, `.ots`) | ISO/IEC 26300 structures and shared ODF package/style/formula models | [OpenDocument spreadsheet feature matrix](../crates/litchi-ods/docs/FEATURE_MATRIX.md) |
| `litchi-odp` | OpenDocument presentation (`.odp`, `.otp`) | ISO/IEC 26300 structures and shared ODF package/style/drawing models | [OpenDocument presentation feature matrix](../crates/litchi-odp/docs/FEATURE_MATRIX.md) |
| `litchi-odg` | OpenDocument drawing (`.odg`) | ISO/IEC 26300 package and drawing structures | [OpenDocument drawing feature matrix](../crates/litchi-odg/docs/FEATURE_MATRIX.md) |
| `litchi-odc` | OpenDocument chart (`.odc`) | ISO/IEC 26300 package and chart structures | [OpenDocument chart feature matrix](../crates/litchi-odc/docs/FEATURE_MATRIX.md) |
| `litchi-odi` | OpenDocument image (`.odi`) | ISO/IEC 26300 package and image structures | [OpenDocument image feature matrix](../crates/litchi-odi/docs/FEATURE_MATRIX.md) |
| `litchi-odb` | OpenDocument database (`.odb`) | ISO/IEC 26300 package and database structures | [OpenDocument database feature matrix](../crates/litchi-odb/docs/FEATURE_MATRIX.md) |
| `litchi-oth` | OpenDocument HTML template (`.oth`) | ISO/IEC 26300 package and text-web structures | [OpenDocument HTML-template feature matrix](../crates/litchi-oth/docs/FEATURE_MATRIX.md) |
| `litchi-odm` | OpenDocument master document (`.odm`) | ISO/IEC 26300 package and master-document structures | [OpenDocument master-document feature matrix](../crates/litchi-odm/docs/FEATURE_MATRIX.md) |
| `litchi-odf-formula` | OpenDocument formula document (`.odf`, `.otf`) | ISO/IEC 26300 formula packages and W3C MathML | [OpenDocument formula feature matrix](../crates/litchi-odf-formula/docs/FEATURE_MATRIX.md) |

## Project-wide boundaries

- External references are never fetched. Macros, VBA, ActiveX, controls, OLE objects, scripts,
  actions, database commands and embedded code are never executed or activated; they are only
  retained, inspected or edited as inert data.
- Signature verification checks integrity only; it does not establish certificate trust or
  revocation status.
- No format crate renders documents.

## Maintenance rules

1. A matrix is a checklist: one row per feature family with its `Status`, `Read` and `Write`
   marks. Add a short note only where a mark alone would mislead.
2. List important unimplemented feature families as `❌` rather than omitting them.
3. Update the affected matrix in the same change as the implementation.
