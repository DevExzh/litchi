# Excel Binary Workbook (.xlsb) Feature Matrix

Feature checklist. ✅ Supported · 🟡 Partial · ❌ Unsupported · N/A Not applicable.

## Package, BIFF12, workbook, and core data

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Fresh workbook creation | ✅ | ✅ | ✅ |
| OPC package and BIFF12 part graph | ✅ | ✅ | ✅ |
| Workbook Theme part | 🟡 | 🟡 | 🟡 |
| BIFF12 raw record codec | ✅ | ✅ | ✅ |
| Worksheet binary index | 🟡 | 🟡 | 🟡 |
| Workbook and worksheet binary streams | ✅ | ✅ | ✅ |
| Workbook and sheet directory | ✅ | ✅ | ✅ |
| Cell values and cached results | ✅ | ✅ | ✅ |
| Shared strings | ✅ | ✅ | ✅ |
| Formula tokens and cached formula values | 🟡 | ✅ | 🟡 |
| Array, shared, and dynamic-array formulas | 🟡 | ✅ | 🟡 |
| Defined names and named ranges | ✅ | ✅ | ✅ |
| Calculation properties | ✅ | ✅ | ✅ |
| Spreadsheet Data Model records and model part | 🟡 | ✅ | 🟡 |
| XLDM 140 table names and time groupings | 🟡 | ✅ | 🟡 |
| Calculation Chain | 🟡 | ✅ | 🟡 |
| Volatile Dependencies | 🟡 | ✅ | 🟡 |
| Styles, fonts, fills, borders, alignment, and number formats | ✅ | ✅ | ✅ |
| Workbook and worksheet views | 🟡 | ✅ | 🟡 |
| Merged cells | ✅ | ✅ | ✅ |
| Row and column layout | ✅ | ✅ | ✅ |
| Core/package properties | 🟡 | 🟡 | 🟡 |

## Worksheet features

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Worksheet cell CRUD | ✅ | ✅ | ✅ |
| Hyperlinks | ✅ | ✅ | ✅ |
| Classic comments and notes | ✅ | ✅ | ✅ |
| Threaded comments and persons | ✅ | ✅ | ✅ |
| Data validation | ✅ | ✅ | ✅ |
| Conditional formatting | ✅ | ✅ | ✅ |
| Sparklines | ✅ | ✅ | ✅ |
| AutoFilter, filter criteria, and sort state | 🟡 | ✅ | 🟡 |
| Structured tables | ✅ | ✅ | ✅ |
| Page setup and print metadata | 🟡 | 🟡 | 🟡 |
| Sheet protection | 🟡 | ✅ | 🟡 |
| Scenarios and what-if analysis | 🟡 | ✅ | ✅ |
| Cell watches and phonetic metadata | ✅ | ✅ | ✅ |
| Shared-workbook revision records | 🟡 | ✅ | ✅ |

## Charts, drawings, pivots, and package objects

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Chart sheets | 🟡 | ✅ | 🟡 |
| Embedded charts and chart resources | 🟡 | ✅ | 🟡 |
| Worksheet drawings and anchors | 🟡 | ✅ | ✅ |
| Images | 🟡 | ✅ | ✅ |
| OLE and embedded package objects | 🟡 | ✅ | 🟡 |
| Pivot caches | ✅ | ✅ | ✅ |
| PivotTable views | ✅ | ✅ | ✅ |
| Pivot charts | 🟡 | ✅ | 🟡 |
| Slicers and slicer caches | 🟡 | ✅ | ✅ |
| Timelines | 🟡 | ✅ | ✅ |

## External references, connections, and extensibility

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| External workbook links | ✅ | ✅ | 🟡 |
| External formulas and cached values | 🟡 | ✅ | 🟡 |
| DDE and OLE links | 🟡 | ✅ | 🟡 |
| External data connections | 🟡 | ✅ | 🟡 |
| Query tables and web imports | 🟡 | ✅ | 🟡 |
| Custom XML Maps catalogs and BIFF12 bindings | 🟡 | ✅ | ✅ |
| Ribbon and custom UI parts | 🟡 | ✅ | 🟡 |
| Task panes and web-extension bindings | 🟡 | ✅ | 🟡 |
| VBA project and code modules | 🟡 | 🟡 | ✅ |
| OPC digital signatures | ✅ | ✅ | ✅ |
| Password encryption | ❌ | ❌ | ❌ |
| Generic package part editing | ✅ | ✅ | ✅ |

## Explicit gaps

| Feature family exposed by [MS-XLSB] | Status | Read | Write |
| --- | --- | --- | --- |
| Complete Excel formula evaluation and recalculation | ❌ | ❌ | ❌ |
| Rich values, rich-data metadata, and modern data types | ❌ | ❌ | ❌ |
| MDX/cell/value metadata semantics | 🟡 | 🟡 | 🟡 |
| Complete chart grammar and rendering | ❌ | ❌ | ❌ |
| Slicers, timelines, and their filter behavior | 🟡 | ✅ | ✅ |
| Mapped XML import/export, schema/XPath evaluation, and Custom XML Data Storage processing | ❌ | ❌ | ❌ |
| ActiveX/form-control execution | ❌ | ❌ | ❌ |
| Macro execution and external-link/connection refresh | ❌ | ❌ | ❌ |
| Certificate trust and revocation | ❌ | ❌ | ❌ |
