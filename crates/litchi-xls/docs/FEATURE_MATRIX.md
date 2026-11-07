# Excel BIFF8 (.xls) Feature Matrix

Feature checklist. ✅ Supported · 🟡 Partial · ❌ Unsupported · N/A Not applicable.

## Package, workbook, and core BIFF

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| OLE Compound File container | ✅ | ✅ | ✅ |
| BIFF8 workbook and worksheet record streams | ✅ | ✅ | ✅ |
| Generic BIFF record access | ✅ | ✅ | ✅ |
| Unknown and future BIFF records | 🟡 | ✅ | 🟡 |
| Workbook and sheet directory | ✅ | ✅ | ✅ |
| Cell values and cached results | ✅ | ✅ | ✅ |
| Shared strings and rich-string properties | ✅ | ✅ | ✅ |
| Formula tokens and cached formula values | 🟡 | ✅ | 🟡 |
| Array, shared, and data-table formulas | 🟡 | ✅ | 🟡 |
| Defined names and built-in names | ✅ | ✅ | ✅ |
| Workbook calculation properties | ✅ | ✅ | ✅ |
| Workbook views and window state | ✅ | ✅ | ✅ |
| Core, summary, and custom properties | ✅ | ✅ | ✅ |
| Theme, fonts, palette, number formats, and XF styles | ✅ | ✅ | ✅ |
| Workbook protection | ✅ | ✅ | ✅ |
| BIFF password encryption | ✅ | ✅ | ✅ |
| Legacy digital signatures | ✅ | ✅ | ✅ |
| VBA project and module storage | 🟡 | 🟡 | ✅ |
| Custom XML data store | 🟡 | 🟡 | 🟡 |
| Legacy macrosheet and dialog-sheet behavior | 🟡 | ✅ | 🟡 |

## Worksheet content and interaction metadata

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Cell and worksheet CRUD | ✅ | ✅ | ✅ |
| Row and column dimensions | ✅ | ✅ | ✅ |
| Cell formatting lookup | ✅ | ✅ | ✅ |
| Merged cells | ✅ | ✅ | ✅ |
| Hyperlinks | ✅ | ✅ | ✅ |
| Classic comments and notes | ✅ | ✅ | ✅ |
| Worksheet protection | ✅ | ✅ | ✅ |
| Data validation | ✅ | ✅ | ✅ |
| Conditional formatting | ✅ | ✅ | ✅ |
| AutoFilter and sort state | ✅ | ✅ | ✅ |
| Structured tables and List12 objects | ✅ | ✅ | ✅ |
| Scenarios and what-if data tables | 🟡 | ✅ | 🟡 |
| Query tables and external-data metadata | 🟡 | ✅ | 🟡 |
| Worksheet views and panes | ✅ | ✅ | ✅ |
| Page setup and print layout | ✅ | ✅ | ✅ |
| Background and header/footer pictures | 🟡 | ✅ | 🟡 |
| Cell watches and formula-error features | 🟡 | ✅ | 🟡 |
| Phonetic information | 🟡 | ✅ | 🟡 |
| Data consolidation | 🟡 | ✅ | 🟡 |
| Web publishing | 🟡 | ✅ | 🟡 |
| Real-time data and RTD topics | 🟡 | ✅ | 🟡 |
| Revision log and shared-workbook collaboration | 🟡 | ✅ | 🟡 |
| User Names stream and shared-workbook user log | 🟡 | ✅ | 🟡 |
| Custom views and workbook environment options | 🟡 | ✅ | 🟡 |
| Office Toolbars stream (`XCB`) | 🟡 | ✅ | ✅ |

## Pivots, charts, drawings, and external references

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Pivot caches | ✅ | ✅ | ✅ |
| PivotTable views | ✅ | ✅ | ✅ |
| Pivot charts | 🟡 | ✅ | 🟡 |
| BIFF charts and chart sheets | 🟡 | ✅ | 🟡 |
| OfficeArt drawings and shape groups | 🟡 | ✅ | 🟡 |
| OLE objects and form controls | 🟡 | ✅ | 🟡 |
| Embedded packages and external payloads | 🟡 | ✅ | 🟡 |
| Supporting books and external links | 🟡 | ✅ | 🟡 |
| DDE and OLE link behavior | 🟡 | ✅ | 🟡 |
| MDX and cell/value metadata | 🟡 | ✅ | 🟡 |

## Explicit gaps

| Feature family exposed by [MS-XLS] | Status | Read | Write |
| --- | --- | --- | --- |
| Threaded comments, persons, mentions, and modern comment threads | ❌ | ❌ | ❌ |
| Slicers and timelines | ❌ | ❌ | ❌ |
| Rich data types and dynamic-array spill semantics | ❌ | ❌ | ❌ |
| XML maps and mapped XML import/export | 🟡 | ✅ | 🟡 |
| Complete Excel formula evaluation | ❌ | ❌ | ❌ |
| Chart rendering and complete chart authoring | ❌ | ❌ | ❌ |
| Macro, control, DDE, RTD, database, and external-link execution | ❌ | ❌ | ❌ |
| Certificate trust, revocation, and signed-content policy | ❌ | ❌ | ❌ |
