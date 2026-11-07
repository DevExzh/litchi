# SpreadsheetML (.xlsx) Feature Matrix

Feature checklist. ✅ Supported · 🟡 Partial · ❌ Unsupported · N/A Not applicable.

## OPC package and workbook

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| OPC package, parts, relationships, and content types | ✅ | ✅ | ✅ |
| Workbook and worksheet lifecycle | ✅ | ✅ | ✅ |
| Workbook properties and metadata | 🟡 | 🟡 | 🟡 |
| Data Model descriptor and payload | 🟡 | 🟡 | 🟡 |
| Custom Data and Custom Data Properties | 🟡 | 🟡 | 🟡 |
| Cell values and cached results | ✅ | ✅ | ✅ |
| Shared strings and rich text runs | ✅ | ✅ | ✅ |
| Formula text and cached formula results | 🟡 | ✅ | 🟡 |
| Extension function names (`_xlfn`) | 🟡 | ✅ | 🟡 |
| Array and shared formulas | 🟡 | ✅ | 🟡 |
| Calculation properties | ✅ | ✅ | ✅ |
| Calculation chain and volatile dependencies | 🟡 | ✅ | 🟡 |
| Defined names | 🟡 | ✅ | 🟡 |
| Styles and theme | ✅ | ✅ | ✅ |
| Rows, columns, outlines, and dimensions | ✅ | ✅ | ✅ |
| Worksheet views, panes, selections, and named sheet views | ✅ | ✅ | ✅ |
| Merged cells | ✅ | ✅ | ✅ |
| Scenarios and what-if data | ✅ | ✅ | ✅ |
| Ignored errors and cell watches | 🟡 | ✅ | 🟡 |
| Cell smart tags | 🟡 | ✅ | ✅ |
| Phonetic properties | 🟡 | ✅ | ❌ |

## Worksheet interaction and presentation

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Hyperlinks | 🟡 | ✅ | 🟡 |
| Legacy comments and VML note placeholders | 🟡 | ✅ | ✅ |
| Threaded comments, persons, and mentions | 🟡 | ✅ | 🟡 |
| [MS-OREACTXML] `commentsExtensible` reactions | N/A | N/A | N/A |
| Data validation | ✅ | ✅ | ✅ |
| Conditional formatting | 🟡 | ✅ | 🟡 |
| AutoFilter and sort | ✅ | ✅ | ✅ |
| Structured tables | ✅ | ✅ | ✅ |
| Surveys attached to tables | 🟡 | 🟡 | 🟡 |
| Query tables and external-data table metadata | 🟡 | ✅ | 🟡 |
| Data consolidation | 🟡 | ✅ | 🟡 |
| Page margins | ✅ | ✅ | ✅ |
| Page setup and print options | ✅ | ✅ | ✅ |
| Row and column page breaks | ✅ | ✅ | ✅ |
| Headers, footers, and printer settings | ✅ | ✅ | ✅ |
| Worksheet and workbook protection | ✅ | ✅ | ✅ |
| Revisions and collaboration metadata | 🟡 | ✅ | 🟡 |
| Sparkline groups | 🟡 | ✅ | 🟡 |

## Charts, drawings, pivots, and external references

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Charts and chart sheets | 🟡 | ✅ | 🟡 |
| Worksheet drawings and anchors | 🟡 | ✅ | 🟡 |
| Images | ✅ | ✅ | ✅ |
| Existing worksheet SVG pictures | 🟡 | ✅ | 🟡 |
| Core worksheet drawing content parts | 🟡 | ✅ | ❌ |
| OLE objects and embedded packages | 🟡 | ✅ | 🟡 |
| Pivot caches | 🟡 | ✅ | 🟡 |
| PivotTable views | 🟡 | ✅ | 🟡 |
| PivotTable server formats (`pivotTableServerFormats`) | 🟡 | ✅ | 🟡 |
| PivotValues (`pivotTableData`) | 🟡 | ✅ | 🟡 |
| Cached unique names (`cachedUniqueNames`) | 🟡 | 🟡 | 🟡 |
| Pivot charts and non-worksheet PivotTables | 🟡 | ✅ | 🟡 |
| Slicers and slicer caches | 🟡 | ✅ | ✅ |
| Timelines and timeline caches | 🟡 | ✅ | ✅ |
| External workbook links | 🟡 | ✅ | 🟡 |
| DDE/OLE links and external data | 🟡 | ✅ | 🟡 |
| External connections | ✅ | ✅ | ✅ |
| XML maps and XML data bindings | ✅ | ✅ | ✅ |
| Custom XML data stores | ✅ | ✅ | ✅ |
| Web extensions and Office add-ins | 🟡 | ✅ | 🟡 |

## Security, macros, and package extensions

| Feature | Status | Read | Write |
| --- | --- | --- | --- |
| Form-control properties (`x14:formControlPr`) | 🟡 | ✅ | 🟡 |
| VBA projects in macro-enabled packages | 🟡 | 🟡 | 🟡 |
| ActiveX and form-control metadata | 🟡 | ✅ | 🟡 |
| OOXML digital signatures | 🟡 | 🟡 | 🟡 |
| OOXML package encryption | ✅ | ✅ | ✅ |
| Persisted Office Add-in task panes and web-extension parts | ✅ | ✅ | ✅ |
| Ribbon and custom UI XML | 🟡 | 🟡 | 🟡 |
| Unknown package parts and extension XML | 🟡 | ✅ | 🟡 |

## Explicit gaps

| Feature family exposed by [MS-XLSX] | Status | Read | Write |
| --- | --- | --- | --- |
| Rich values and modern rich-data objects | 🟡 | ✅ | ✅ |
| Feature property bags and checkbox/XF-control extensions | 🟡 | ✅ | ✅ |
| Python in Excel and external code services | ❌ | ❌ | ❌ |
| Complete PivotTable refresh/calculation | ❌ | ❌ | ❌ |
| Complete chart/drawing rendering and extension grammar | ❌ | ❌ | ❌ |
| Complete Excel formula evaluation | ❌ | ❌ | ❌ |
| External-link, connection, query, DDE, OLE, and add-in execution | ❌ | ❌ | ❌ |
| Macro execution and trusted ActiveX behavior | ❌ | ❌ | ❌ |
| Certificate trust and revocation | ❌ | ❌ | ❌ |
