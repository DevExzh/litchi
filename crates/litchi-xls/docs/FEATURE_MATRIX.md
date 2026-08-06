# Excel BIFF8 (.xls) Feature Matrix

This document tracks the public feature families implemented by litchi-xls. It describes
typed library support, not Excel's rendering, recalculation, trust, or complete conformance
with every BIFF revision.

The audit uses the authoritative Microsoft [MS-XLS] Front Matter and ToC, especially the
compound-file and stream grammar in [MS-XLS] 2.1, the workbook and worksheet substream
grammars in [MS-XLS] 2.1.7.20, the cell/formula/chart/pivot/style model in [MS-XLS] 2.2,
and the examples for conditional formatting, defined names, tables, filters, external
references, charts, formatting, workbooks, and PivotTables in [MS-XLS] 3.1-3.10.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the scope described in the Notes cell |
| 🟡 | Bounded, partial, metadata-only, pass-through, or inert support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to this format or direction |

Read and Write are independent. A supported row does not imply recalculation,
rendering, external I/O, macro execution, or certificate trust.

## Package, workbook, and core BIFF

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| OLE Compound File container | ✅ | ✅ | ✅ | Opens and writes the CFB storages and streams used by .xls; package edits are bounded and transactional at the owning writer boundary. |
| BIFF8 workbook and worksheet record streams | ✅ | ✅ | ✅ | Typed workbook, worksheet, chart-sheet, macro-sheet, and dialog-sheet record sequences are decoded and emitted through the XLS reader/writer. |
| Generic BIFF record access | ✅ | ✅ | ✅ | Public record utilities expose record headers and payloads; the typed model does not claim a semantic implementation of every record in the [MS-XLS] enumeration. |
| Unknown and future BIFF records | 🟡 | ✅ | 🟡 | Selected Continue, FRT, extension, and opaque payload families are retained or replayed; unknown records are not guaranteed to survive every typed mutation or fresh authoring path. |
| Workbook and sheet directory | ✅ | ✅ | ✅ | Sheet names, order, visibility, stable metadata, sheet kinds, code names, and sheet lookup are modeled. |
| Cell values and cached results | ✅ | ✅ | ✅ | Numbers, strings, shared-string references, booleans, errors, blanks, serial dates, and cell formatting references are supported. |
| Shared strings and rich-string properties | ✅ | ✅ | ✅ | BIFF SST/ExtSST lookup plus bounded string-property metadata is available; this is not a general rich-data object model. |
| Formula tokens and cached formula values | 🟡 | ✅ | 🟡 | BIFF formula token rendering, formula metadata, and bounded formula authoring are available. Tokens are not a complete Excel parser or evaluator, unsupported token graphs may be preserved or refused, and cached values are not recalculated. |
| Array, shared, and data-table formulas | 🟡 | ✅ | 🟡 | Array/table metadata and formula records are typed where implemented; Excel's full dependency graph, spill behavior, and recalculation semantics are not implemented. |
| Defined names and built-in names | ✅ | ✅ | ✅ | Workbook and sheet-scoped names, print areas, print titles, filter-database names, and name metadata are typed and writable. Name formulas remain inert token expressions. |
| Workbook calculation properties | ✅ | ✅ | ✅ | Calculation mode, reference mode, iteration settings, precision, multithreading flags, recalculation markers, and force-full-calculation metadata are supported. |
| Workbook views and window state | ✅ | ✅ | ✅ | Workbook windows, selected/active tabs, stable tab identifiers, sheet views, panes, selections, zoom, and related view metadata are supported. |
| Core, summary, and custom properties | ✅ | ✅ | ✅ | CFB property sets are exposed for summary information, document summary information, and user-defined properties. |
| Theme, fonts, palette, number formats, and XF styles | ✅ | ✅ | ✅ | BIFF fonts, palette colors, number formats, cell/style XFs, XF extensions, differential formats, table styles, and theme metadata are typed. |
| Workbook protection | ✅ | ✅ | ✅ | Workbook protection flags and legacy password records are modeled; the legacy hash is not cryptographic protection. |
| BIFF password encryption | ✅ | ✅ | ✅ | Supported BIFF8 password-to-open profiles are decrypted and emitted through the XLS encryption integration; unsupported profiles fail rather than being treated as plaintext. |
| Legacy digital signatures | ✅ | ✅ | ✅ | CFB signature reports can be verified and signature graphs can be edited through the signature integration. Verification is integrity/signature verification only and does not establish certificate trust or revocation. |
| VBA project and module storage | 🟡 | 🟡 | ✅ | VBA project topology, metadata, and bounded module/source payloads can be inspected or attached/replaced. VBA is never executed, resolved, or trusted; replacing project content invalidates stale signatures. |
| Custom XML data store | 🟡 | 🟡 | 🟡 | The workbook exposes bounded custom XML storage through the CFB/package layer, but it is preserved as payload-oriented data rather than a general XML-mapping authoring model. |
| Legacy macrosheet and dialog-sheet behavior | 🟡 | ✅ | 🟡 | Sheet kinds and selected records are recognized as metadata. Macro formulas, dialog controls, and host automation are not evaluated or executed. |

## Worksheet content and interaction metadata

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Cell and worksheet CRUD | ✅ | ✅ | ✅ | Worksheets can be inspected and authored with typed cells, dimensions, shared strings, and sheet metadata. |
| Row and column dimensions | ✅ | ✅ | ✅ | Row heights, hidden/collapsed state, outline levels, column widths, best-fit flags, formatting, and phonetic-guide flags are supported. |
| Cell formatting lookup | ✅ | ✅ | ✅ | Effective cell formats, number formats, borders, fills, alignment, fonts, colors, and date/time interpretation are available. |
| Merged cells | ✅ | ✅ | ✅ | MERGECELLS ranges are read and emitted with checked range boundaries. |
| Hyperlinks | ✅ | ✅ | ✅ | URL, file, item, relative/absolute, display, tooltip, frame, and location metadata are typed; following a link is never performed. |
| Classic comments and notes | ✅ | ✅ | ✅ | BIFF NOTE comments, authors, visibility, anchors, and writer support are available. |
| Worksheet protection | ✅ | ✅ | ✅ | PROTECT, object/scenario protection, password records, and sheet protection flags are modeled; protected content is not a cryptographic boundary. |
| Data validation | ✅ | ✅ | ✅ | BIFF data-validation settings, ranges, operators, formulas, prompts, and error messages are typed and writable within the supported record forms. |
| Conditional formatting | ✅ | ✅ | ✅ | Classic conditional-format rules and the implemented BIFF12-style extension metadata are available; rule formulas are stored/rendered, not evaluated by this crate. |
| AutoFilter and sort state | ✅ | ✅ | ✅ | AutoFilter columns, criteria, filter mode, sort metadata, and extended filter records are typed and writable. |
| Structured tables and List12 objects | ✅ | ✅ | ✅ | List objects, ranges, columns, totals metadata, styles, AutoFilter state, and selected external/web/XML source metadata are serialized. Source queries and calculated-column formulas remain inert. |
| Scenarios and what-if data tables | 🟡 | ✅ | 🟡 | Scenario manager and TABLE records have typed metadata and bounded writing; Excel's what-if calculation and scenario application are not performed. |
| Query tables and external-data metadata | 🟡 | ✅ | 🟡 | Query-table, connection, web-table, and external-table records are typed where implemented. Commands, credentials, and refresh behavior are never used. |
| Worksheet views and panes | ✅ | ✅ | ✅ | Window, zoom, pane/freeze, selection, and related worksheet view records are supported. |
| Page setup and print layout | ✅ | ✅ | ✅ | Margins, paper, scale/fit, orientation, print order, headers/footers, page breaks, gridlines, comments/errors display, and print flags are typed. Printer-driver bytes are inert. |
| Background and header/footer pictures | 🟡 | ✅ | 🟡 | BIFF background and header/footer picture records are recognized with bounded payload handling; no renderer is provided. |
| Cell watches and formula-error features | 🟡 | ✅ | 🟡 | CellWatch, error-checking, and related worksheet flags are exposed as metadata and can be retained in supported paths; no watch-window or diagnostic engine is implemented. |
| Phonetic information | 🟡 | ✅ | 🟡 | Phonetic string format and visible-range metadata are typed; phonetic layout and language conversion are not implemented. |
| Data consolidation | 🟡 | ✅ | 🟡 | Consolidation directories and source metadata are available as inert typed records; consolidation is not executed. |
| Web publishing | 🟡 | ✅ | 🟡 | Web publication destinations, ranges, source types, and page metadata are modeled; publishing and network I/O are not performed. |
| Real-time data and RTD topics | 🟡 | ✅ | 🟡 | RTD topic/cell metadata and cached values are inert; no server connection, refresh, or execution occurs. |
| Revision log and shared-workbook collaboration | 🟡 | ✅ | 🟡 | Revision streams and selected revision records are parsed as typed metadata; conflict resolution, shared locking, and collaboration behavior are not implemented. |
| Custom views and workbook environment options | 🟡 | ✅ | 🟡 | Custom views, link-update modes, object-display modes, access/write strings, backup flags, and related behavior options are exposed as metadata. They do not cause host actions. |
| Office Toolbars stream (`XCB`) | 🟡 | ✅ | ✅ | `Workbook::toolbar` reads and `Writer::set_toolbar` emits typed `CTBWRAPPER`/`CTBS`/`CTB` metadata through the root `XCB` stream, preserving reserved fields, optional visual bytes, `TBCCmd`, and bounded shared `TBCGeneralInfo`/`TBCExtraInfo` payloads for non-ActiveX controls. Ambiguous or malformed variable boundaries are rejected, and macros/UI/ActiveX are never executed. |

## Pivots, charts, drawings, and external references

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Pivot caches | ✅ | ✅ | ✅ | Typed PivotCache definitions include shared items, value types, grouping, worksheet/consolidation sources, OLAP hierarchies/tuples, calculated items/members, and supported extension records. |
| PivotTable views | ✅ | ✅ | ✅ | Pivot fields/items, axes, page selections, data items, layouts, functions, view extensions, and the constrained pivot-view editor are available. The crate does not refresh or calculate a PivotTable. |
| Pivot charts | 🟡 | ✅ | 🟡 | Pivot-chart links and related SXView/PivotChart records are recognized; chart rendering and complete pivot-chart authoring remain bounded by the chart graph support. |
| BIFF charts and chart sheets | 🟡 | ✅ | 🟡 | litchi-xls::chart provides chart-sheet discovery, strict BOF/EOF framing, typed bounded series formulas/cache cells, transactional chart-sheet data patches with byte-preserved unknown records, and exact pristine replay. Full chart grammar certification, fresh chart authoring, and unsafe embedded-chart graph edits are refused. |
| OfficeArt drawings and shape groups | 🟡 | ✅ | 🟡 | Shape extraction and selected shape/group geometry, anchors, text, and drawing metadata are available. The complete OfficeArt drawing-group graph, renderer, and arbitrary shape mutation are not claimed. |
| OLE objects and form controls | 🟡 | ✅ | 🟡 | OLE object records and selected controls such as buttons, checkboxes, list boxes, edit boxes, and drop-downs have typed metadata and bounded payloads. Objects are never instantiated, activated, rendered, or executed. |
| Embedded packages and external payloads | 🟡 | ✅ | ❌ | Embedded/OLE payloads can be inventoried or retained in supported CFB paths; general fresh embedded-payload authoring is not provided. |
| Supporting books and external links | 🟡 | ✅ | 🟡 | Supporting-book records, external names, source paths, external references, link flags, and cached values are typed where implemented. Paths and formulas are stored verbatim and never opened, fetched, refreshed, or evaluated. |
| DDE and OLE link behavior | 🟡 | ✅ | 🟡 | DDE/OLE link metadata is retained as inert records. No conversation, COM activation, data refresh, or external program launch occurs. |
| MDX and cell/value metadata | 🟡 | ✅ | 🟡 | BIFF metadata, MDX records, and selected metadata blocks are exposed as bounded metadata. Cube connections and MDX expressions are not executed. |

## Explicit gaps

| Feature family exposed by [MS-XLS] | Status | Read | Write | Notes |
|-------------------------------------|--------|------|-------|-------|
| Threaded comments, persons, mentions, and modern comment threads | ❌ | ❌ | ❌ | The .xls model supports classic NOTE comments only; modern threaded-comment parts are not a BIFF8 typed feature. |
| Slicers and timelines | ❌ | ❌ | ❌ | No typed slicer/timeline cache or view model is provided for the BIFF8 crate. |
| Rich data types and dynamic-array spill semantics | ❌ | ❌ | ❌ | The [MS-XLS] cell/formula records do not become a modern rich-value or spill-calculation engine in this crate. |
| XML maps and mapped XML import/export | 🟡 | ✅ | 🟡 | `[MS-XLS]` XML-map metadata has a typed `MapInfo`/schema/map/data-binding owner, bounded XML codec, workbook `XML`-stream loading, and list-column dependency checks; schema resolution and mapped-cell import/export remain inert |
| Complete Excel formula evaluation | ❌ | ❌ | ❌ | Formula tokens and cached results can be read, but there is no Excel-compatible recalculation engine, dependency graph, external-function execution, or volatile refresh. |
| Chart rendering and complete chart authoring | ❌ | ❌ | ❌ | The bounded chart model is intentionally not a renderer or a complete fresh chart writer; unsafe graph edits fail rather than producing a corrupt workbook. |
| Macro, control, DDE, RTD, database, and external-link execution | ❌ | ❌ | ❌ | These protocol families are represented only as inert metadata or cached payloads. |
| Certificate trust, revocation, and signed-content policy | ❌ | ❌ | ❌ | Signature verification establishes cryptographic integrity/signature validity only; it does not establish trust. |
