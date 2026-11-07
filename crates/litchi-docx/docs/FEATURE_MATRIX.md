# WordprocessingML (DOCX) Feature Matrix

Feature checklist. ✅ Supported · 🟡 Partial · ❌ Unsupported · N/A Not applicable.

## Package graph and shared Office features

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| OPC parts, relationships, and content types | ✅ | ✅ | ✅ |
| Transitional and Strict package forms | ✅ | ✅ | ✅ |
| Markup Compatibility (mc:Ignorable, AlternateContent) | 🟡 | ✅ | ✅ |
| Main document, settings, styles, numbering, theme, font table, comments, notes, headers, footers, glossary, and properties parts | ✅ | ✅ | ✅ |
| Main-document source-preserving transactions | 🟡 | ✅ | 🟡 |
| Cross-story hyperlink batches | ✅ | ✅ | ✅ |
| Source-backed secondary-story text | 🟡 | ✅ | 🟡 |
| Ordered paragraph and run content projection | 🟡 | ✅ | N/A |
| Web settings part | ✅ | ✅ | ✅ |
| stylesWithEffects part | 🟡 | ✅ | 🟡 |
| Core, extended, and custom properties | ✅ | ✅ | ✅ |
| Embedded OLE and package objects | ✅ | ✅ | ✅ |
| Embedded fonts | ✅ | ✅ | ✅ |
| Images, charts, and DrawingML/VML relationships | ✅ | ✅ | ✅ |
| Main-story source SVG picture inventory | 🟡 | ✅ | ❌ |
| Digital signatures | ✅ | ✅ | ✅ |
| Password encryption | ✅ | ✅ | ✅ |
| Information Rights Management and protected content | 🟡 | ✅ | 🟡 |
| VBA projects in DOCM/DOTM | 🟡 | 🟡 | 🟡 |
| Web extensions and Office Add-ins | 🟡 | ✅ | 🟡 |
| Mail-merge settings and recipient metadata | ✅ | ✅ | ✅ |
| External relationships and linked content | 🟡 | ✅ | ✅ |

## [MS-DOCX] part and extension

| Feature family | Status | Read | Write |
| --- | --- | --- | --- |
| commentsExtended, people, commentsIds, and commentsExtensible parts | ✅ | ✅ | ✅ |
| Modern comment reactions and extension lists | ✅ | ✅ | ✅ |
| rPr extension effects (glow, shadow, reflection, textOutline, textFill, scene3d, props3d) | ✅ | ✅ | ✅ |
| rPr OpenType extensions (ligatures, numForm, numSpacing, stylisticSets, cntxtAlts) | ✅ | ✅ | ✅ |
| Settings extension elements (chartTrackingRefBased, docId, conflictMode, discardImageEditingData, defaultImageDpi) | ✅ | ✅ | ✅ |
| Generic compatSetting records and compatibility mode | ✅ | ✅ | ✅ |
| SDT extension controls (entityPicker, checkbox, repeatingSection, repeatingSectionItem, appearance, color, dataBinding) | 🟡 | 🟡 | 🟡 |
| SDT web-extension links (webExtensionsLinked, webExtensionCreated) | 🟡 | ✅ | 🟡 |
| Paragraph and table-row extension identifiers (paraId, textId) and paragraph noSpellErr | ✅ | ✅ | ✅ |
| W14 conflict revision markup (conflictIns, conflictDel, and custom XML conflict ranges) | 🟡 | ✅ | 🟡 |
| anchorId on object and pict | ✅ | ✅ | ✅ |
| Um Al-Qura umalqura calendar extension | ✅ | ✅ | ✅ |
| sectPr footnoteColumns | ✅ | ✅ | ✅ |
| pPr collapsed | ✅ | ✅ | ✅ |
| Numbering restartNumberingAfterBreak | ✅ | ✅ | ✅ |
| Run symEx symbol extension | ✅ | ✅ | ✅ |
| Data-binding storeItemChecksum | ✅ | ✅ | ✅ |
| Word 2024 formattingAllowed | ✅ | ✅ | ✅ |
| Word 2023 dateUtc | 🟡 | ✅ | ✅ |
| Extended note/list number formats from [MS-DOCX] 2.4 | 🟡 | ✅ | 🟡 |

## Base WordprocessingML typed owners and operations

| Feature family | Status | Read | Write |
| --- | --- | --- | --- |
| Tracked revision metadata and property changes | 🟡 | ✅ | 🟡 |
| Revision application and accept/reject | 🟡 | ✅ | 🟡 |
| Fields and cached field results | 🟡 | ✅ | ✅ |
| Office Math (OMML) fragments | 🟡 | ✅ | ✅ |
| Bibliography source stores | 🟡 | ✅ | ✅ |
| Smart tags | 🟡 | ✅ | ✅ |
| Bookmarks and bookmark anchors | 🟡 | ✅ | ✅ |
| Text boxes and WordArt inventory | 🟡 | ✅ | ✅ |
| Custom XML data stores and content-control bindings | 🟡 | ✅ | ✅ |
| Streaming document authoring | ✅ | N/A | ✅ |
| Validation and document statistics | 🟡 | ✅ | N/A |
| Direct (local) formatting properties | 🟡 | ✅ | N/A |
| Source-bound sanitization and external-hyperlink redaction | 🟡 | ✅ | ✅ |

## Explicit semantic gaps

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Pagination, line breaking, page fields, and rendering | ❌ | ❌ | ❌ |
| Field calculation and refresh | 🟡 | 🟡 | 🟡 |
| Mail-merge execution | ❌ | ❌ | ❌ |
| AltChunk import and foreign-content conversion | ❌ | ❌ | ❌ |
| Macro, ActiveX, form-control, and add-in execution | ❌ | ❌ | ❌ |
| Ink / inkML model | 🟡 | 🟡 | 🟡 owning authoring and durable replay |
| MS-OWEXML custom-function payloads | 🟡 | 🟡 | 🟡 |
| Chart, SmartArt, DrawingML, VML, and embedded-workbook rendering | ❌ | ❌ | ❌ |
| Protection enforcement | 🟡 | ✅ | ✅ |
| Signature trust and certificate status | 🟡 | ✅ | ✅ |
| Unsupported Word extension namespace content | 🟡 | 🟡 | 🟡 |
