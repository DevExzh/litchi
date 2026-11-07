# Rich Text Format feature matrix

Feature checklist. ✅ Supported · 🟡 Partial · ❌ Unsupported · N/A Not applicable.

## Input, output, and transactions

| Feature | Read | Write |
| --- | --- | --- |
| RTF groups, control words, escaped text, binary payloads | ✅ | ✅ |
| Unicode and legacy code pages | ✅ | 🟡 |
| Immutable package and read-only projection views | ✅ | N/A |
| Streaming fresh creation | N/A | 🟡 |
| Existing-package transactions | ✅ | 🟡 |
| Reversible patches, composition, and history | ✅ | ✅ |
| Cross-document transfer | ✅ | 🟡 |
| Tail append | ✅ | 🟡 |
| Unknown control words and starred destinations | 🟡 | 🟡 |
| Validation and redaction | 🟡 | 🟡 |
| Compressed RTF transport, [MS-OXRTFCP] | ✅ | ✅ |

## Document and content vocabulary

| Feature | Read | Write |
| --- | --- | --- |
| Character formatting and paragraph properties | ✅ | 🟡 |
| Paragraph, line, page, column, and soft breaks | ✅ | ✅ |
| Font/color tables and embedded fonts | ✅ | 🟡 |
| Stylesheets, Quick Styles, and personal/compose/reply styles | ✅ | 🟡 |
| Table styles and conditional formatting | ✅ | 🟡 |
| Latent styles and style/format restrictions | ✅ | 🟡 |
| Tables, nested tables, merges, and floating tables | ✅ | 🟡 |
| Lists and picture bullets | ✅ | 🟡 |
| Sections and header/footer stories | ✅ | 🟡 |
| Footnotes and endnotes | ✅ | 🟡 |
| Comments/annotations | ✅ | 🟡 |
| Bookmarks and navigation entries | ✅ | 🟡 |
| Tracked revisions | ✅ | 🟡 |
| Fields and hyperlinks | ✅ | 🟡 |
| Form fields and editable/protected regions | ✅ | 🟡 |
| Math and equations | ✅ | 🟡 |
| Pictures, drawings, shape groups, and text boxes | ✅ | 🟡 |
| OLE/embedded objects and Publisher objects | 🟡 | 🟡 |
| Document information and producer metadata | ✅ | 🟡 |
| Mail merge, external references, and XSL transforms | 🟡 | 🟡 |
| Document layout/compatibility/view/print policies | 🟡 | 🟡 |
| Bidirectional, associated-font, and East Asian formatting | 🟡 | 🟡 |
| Legacy document protection password | 🟡 | 🟡 |
| Write-reservation metadata | 🟡 | 🟡 |

## Explicit semantic gaps and exclusions

| Feature | Read | Write |
| --- | --- | --- |
| Positioned paragraph objects and frames | ✅ | 🟡 |
| SmartTag/factoid data | ✅ | 🟡 |
| Move bookmarks | ✅ | 🟡 |
| Modern protection password hash | ✅ | 🟡 |
| Word 6J-era East Asian controls | 🟡 | 🟡 |
| Rendering, pagination, shaping, and field/formula calculation | ❌ | ❌ |
| Macro/control/object execution and external refresh | ❌ | ❌ |
