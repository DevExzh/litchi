# SpreadsheetML (.xlsx) Feature Matrix

This document tracks the public feature families implemented by litchi-xlsx. It describes
typed SpreadsheetML and OPC package support, not Excel rendering, recalculation, network
access, macro execution, certificate trust, or complete support for every Microsoft
extension schema.

The audit uses the authoritative Microsoft [MS-XLSX] Front Matter and ToC, the part,
extension, conceptual, and security material in [MS-XLSX] 2.1-2.7 and 4, and the slicer
example in [MS-XLSX] 3.1. The conceptual review explicitly accounts for PivotTable
what-if data, slicers, non-worksheet PivotTables, PivotValues, timelines, rich data,
threaded comments, named sheet views, feature property bags, and Python in Excel.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the scope described in the Notes cell |
| 🟡 | Bounded, partial, metadata-only, pass-through, or inert support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to this format or direction |

Read and Write are independent. A typed package row does not imply that formula,
query, macro, chart, pivot, or external content is executed.

## OPC package and workbook

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| OPC package, parts, relationships, and content types | ✅ | ✅ | ✅ | Package owns SpreadsheetML parts, relationship graphs, content types, transactional edits, and package saves. Strict and Transitional SpreadsheetML are handled where the feature codec supports both. |
| Workbook and worksheet lifecycle | ✅ | ✅ | ✅ | Workbook creation/opening, sheet names/order, insert/remove/reorder operations, worksheet lookup, and chart-sheet removal are supported. |
| Workbook properties and metadata | ✅ | ✅ | ✅ | Core/extended/custom workbook metadata and bounded workbook metadata/protection models are available through the package/workbook APIs. |
| Cell values and cached results | ✅ | ✅ | ✅ | Numbers, dates, strings, inline/shared strings, booleans, errors, blanks, and cell references are typed and serialized. |
| Shared strings and rich text runs | ✅ | ✅ | ✅ | Shared-string tables and supported rich text run properties are modeled; modern rich-value objects are a separate unsupported feature family below. |
| Formula text and cached formula results | 🟡 | ✅ | 🟡 | Formula references, function text, shared/array formulas, table/pivot references, and bounded formula edits are supported. No Excel-compatible evaluator, dependency graph, volatile refresh, or external-function execution is provided. |
| Array and shared formulas | 🟡 | ✅ | 🟡 | Array/shared formula ownership and supported autofit/edit behavior are covered; the full dynamic-array spill engine is not. |
| Calculation properties | ✅ | ✅ | ✅ | Bounded, inert typed reads and package transactions cover authored `calcPr` values and ordered, duplicate-preserving `calcFeatures`, including exact no-op/removal and source-checked reversible patches. Changed edits projected through MCE markup are refused. Formula edits only invalidate cached calculation state: Litchi does not recalculate formulas, infer calculation features, or implement complete calculation-engine semantics. Changed signed packages and mutation of encrypted-source facades are refused; encrypted input requires an explicit plaintext declassification/reopen workflow before editing. |
| Calculation chain and volatile dependencies | 🟡 | ✅ | 🟡 | Chain/volatile-dependency parts have bounded ownership and edit support; they are not recalculated and cannot establish Excel's full dependency semantics. |
| Defined names | ✅ | ✅ | ✅ | Workbook/sheet-scoped names, built-in names, name formulas, and validated name edits are supported. Names are not resolved into external I/O. |
| Styles and theme | ✅ | ✅ | ✅ | Fonts, fills, borders, alignment, number formats, cell formats, differential formats, themes, colors, and style inheritance are typed and writable. |
| Rows, columns, outlines, and dimensions | ✅ | ✅ | ✅ | Row/column height and width, hidden/collapsed state, outline levels, defaults, and dimension metadata are supported. |
| Worksheet views, panes, selections, and named sheet views | ✅ | ✅ | ✅ | Sheet views, frozen panes, selections, tab/view settings, named sheet views, and supported reconciliation/filter metadata are available. |
| Merged cells | ✅ | ✅ | ✅ | Merge ranges have checked read and CRUD operations, including lookup by worksheet name. |
| Scenarios and what-if data | ✅ | ✅ | ✅ | Scenario metadata and supported package relationships have typed CRUD; Excel's what-if calculation is not performed. |
| Ignored errors and cell watches | ✅ | ✅ | ✅ | Ignored-error ranges and worksheet cell-watch entries have bounded parse/write support. |
| Cell smart tags | 🟡 | ✅ | ✅ | The layered `smart_tags::{model,codec,package,validation,transaction}` owner provides inert typed annotations, checked cell references, strict Office type/property bounds, source-preserving worksheet edits, and semantic package selection. Action providers and smart-tag execution are intentionally unsupported. |
| Phonetic properties | ✅ | ✅ | ✅ | Worksheet phonetic properties and supported attributes are typed; phonetic layout/conversion is not a renderer. |

## Worksheet interaction and presentation

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Hyperlinks | ✅ | ✅ | ✅ | External and internal hyperlink metadata, locations, display text, and relationships are typed; targets are never followed. |
| Legacy comments and VML note placeholders | 🟡 | ✅ | ✅ | The contextual `workbook::comments::{snapshot,transaction,patch}` owner provides checked cell/author CRUD for classic notes, exact semantic no-ops, relationship-safe source-checked patches, and inert VML shape-ID retention. VML layout and rendering remain outside the owner. |
| Threaded comments, persons, and mentions | ✅ | ✅ | ✅ | Threaded comment threads, persons, mentions, replies, legacy placeholders, relationship IDs, and CRUD are supported as metadata; no collaboration service is contacted. |
| Data validation | ✅ | ✅ | ✅ | Validation ranges, list/custom/date/time/whole/decimal/text-length rules, formulas, prompts, and error messages are typed and writable within supported schemas. |
| Conditional formatting | ✅ | ✅ | ✅ | Rules, differential formats, ranges, formula expressions, color scales, data bars, icon sets, and supported extension forms are typed and serialized; conditions are not evaluated. |
| AutoFilter and sort | ✅ | ✅ | ✅ | AutoFilter ranges, filters, sort states, custom/dynamic filters, and supported extension metadata are available. Filtering is not executed against cells. |
| Structured tables | ✅ | ✅ | ✅ | Table identity, display names, ranges, columns, totals, styles, AutoFilter, calculated-column formulas, and package relationships are typed and writable. |
| Query tables and external-data table metadata | ✅ | ✅ | ✅ | Query-table identity, fields, refresh/import metadata, relationships, and supported connection references are modeled. No query is fetched or refreshed. |
| Data consolidation | 🟡 | ✅ | 🟡 | Consolidation metadata and source references have bounded typed coverage; consolidation calculations are not performed. |
| Page margins, setup, print options, and breaks | ✅ | ✅ | ✅ | Margins, paper, scale/fit, orientation, print order, comments/errors display, gridlines, print options, page breaks, and setup relationships are supported. |
| Headers, footers, and printer settings | ✅ | ✅ | ✅ | Header/footer sections, odd/even/first-page settings, pictures/relationships, and supported printer-setting parts are typed; printing is not performed. |
| Worksheet and workbook protection | ✅ | ✅ | ✅ | Sheet/workbook protection flags, legacy hashes, and supported modern hash/salt/spin metadata are validated and writable. Protection is not encryption or authorization. |
| Revisions and collaboration metadata | 🟡 | ✅ | 🟡 | `revisions::{Snapshot, Transaction, Commit, Patch}` provides source-checked package CRUD for revision users, headers, and inert logs with relationship/orphan validation and exact no-op preservation; live coauthoring, conflict resolution, locking, and recalculation are not implemented. |
| Sparkline groups | 🟡 | ✅ | 🟡 | Supported sparkline group and worksheet relationships are bounded; the crate does not render sparklines or calculate their source values. |

## Charts, drawings, pivots, and external references

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Charts and chart sheets | 🟡 | ✅ | 🟡 | Typed chart/chart-sheet resources, series, references, axes, labels, layouts, and supported relationships are available. Coverage is bounded by the implemented chart schema; there is no renderer or claim of complete Microsoft extension grammar. |
| Worksheet drawings and anchors | 🟡 | ✅ | 🟡 | Two-cell/one-cell/absolute anchors, groups, shapes, connections, text bodies, geometry, client data, images, and supported drawing relationships are typed. Complex unmodeled drawing extensions are preserved or refused rather than guessed. |
| Images | ✅ | ✅ | ✅ | Supported image relationships and drawing image resources can be read and authored; image conversion and final layout/rendering are outside this crate. |
| OLE objects and embedded packages | 🟡 | ✅ | 🟡 | `ole_objects::{Snapshot, Transaction, Commit, Patch}` provides source-checked anchor/relationship metadata edits, MCE choice/fallback preservation, and bounded inert payload lifecycle support. Payloads are never activated, deserialized, or executed. |
| Pivot caches | 🟡 | ✅ | 🟡 | Pivot cache definitions, fields, shared items, grouping, filters, styles, and supported extension data have typed bounded coverage. Cache refresh, cube access, and arbitrary unknown extensions are not implemented. |
| PivotTable views | 🟡 | ✅ | 🟡 | Pivot view and field metadata are read and retained through the pivot package model; full report authoring, refresh, layout calculation, and all product-specific extensions are not claimed. |
| Pivot charts and non-worksheet PivotTables | 🟡 | ✅ | 🟡 | Pivot-chart resources and non-worksheet chart relationships have bounded support; chart/pivot rendering and calculation are not implemented. |
| Slicers and slicer caches | 🟡 | ✅ | ✅ | The layered `slicer::{model,codec,package,validation,transaction}` owner provides bounded typed cache/view graphs, extension retention, clone-staged CRUD, and explicit graph validation. Slicer filtering, refresh, recalculation, and UI rendering remain inert. |
| Timelines and timeline caches | 🟡 | ✅ | ✅ | The layered `timeline::{model,codec,package,validation,transaction}` owner provides bounded typed cache/view state, extension retention, clone-staged CRUD, and relationship/orphan validation. Timeline filtering, refresh, recalculation, and UI rendering remain explicitly unsupported. |
| External workbook links | ✅ | ✅ | ✅ | `external_links::{Snapshot, Transaction, Commit, Patch}` adds source-checked package CRUD and inert metadata/cache edits over external link parts, source relationships, defined names, and supported extension properties. Paths and targets are retained verbatim and never opened or refreshed. |
| DDE/OLE links and external data | 🟡 | ✅ | 🟡 | Link targets, connection strings, commands, credentials, and DDE/OLE metadata are inert typed data. No conversation, COM activation, URL access, refresh, or execution occurs. |
| External connections | ✅ | ✅ | ✅ | `connections::{Snapshot, Transaction, Commit, Patch}` provides source-checked create/update/remove edits for database, OLE DB/ODBC, OLAP, Web, text-import, parameters, credential, and query-table metadata while preserving opaque XML and graph topology. The connection is never contacted. |
| XML maps and XML data bindings | ✅ | ✅ | ✅ | `xml_maps::{Snapshot, Transaction, Commit, Patch}` owns typed XML map schemas, data bindings, map IDs, XPath metadata, package relationships, and bounded mapped-data payloads with source-checked CRUD. XML is not fetched or transformed by a host service. |
| Custom XML data stores | ✅ | ✅ | ✅ | Bounded custom XML payloads and package relationships are typed and editable through the shared `litchi_ooxml_common::custom_xml` owner; arbitrary application semantics are not inferred. |
| Web extensions and Office add-ins | 🟡 | ✅ | 🟡 | Worksheet range bindings, task panes, web-extension relationships, IDs, app references, and supported binding-formula rewrites are handled with bounded validation. Add-in activation and provider-specific behavior remain unsupported. |

## Security, macros, and package extensions

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| VBA projects in macro-enabled packages | 🟡 | 🟡 | 🟡 | Macro-enabled package relationships and bounded project/ActiveX payload metadata can be preserved or edited through package/ActiveX paths. VBA source is not a typed formula/runtime model, never executes, and is not trusted. |
| ActiveX and form-control metadata | 🟡 | ✅ | 🟡 | ActiveX relationships, control metadata, and supported VBA-linked payload references are bounded and inert. Controls are not instantiated, rendered, or run. |
| OOXML digital signatures | 🟡 | 🟡 | 🟡 | The shared OPC graph/signature adapter can inspect and edit supported XML signatures at package level; this crate's workbook model does not validate certificate trust, revocation, or every producer signature profile. |
| OOXML package encryption | ✅ | ✅ | ✅ | With the optional `encryption` feature, `Package` and `Workbook` support the Standard AES-128/SHA-1 and Agile AES-128/CBC/SHA-1 managed-package profiles through bounded password-aware ingress, retained-mode re-encryption, atomic encrypted saves, and explicit plaintext declassification. Other algorithm families are rejected. Crypto and inner OPC limits are independent; encrypted-source mutation and implicit plaintext output are refused. Encryption does not make formula or macro behavior safe. |
| Persisted Office Add-in task panes and web-extension parts | ✅ | ✅ | ✅ | The layered `task_panes::{package,transaction}` owner exposes typed MS-OWEXML task-pane/add-in CRUD through clone-staged `Package::edit_task_panes`; existing relationship IDs, unrelated package graph entries, and supported opaque extension XML are retained. Add-in activation and provider behavior remain inert. |
| Ribbon and custom UI XML | 🟡 | 🟡 | 🟡 | Ribbon/customUI parts and relationships remain bounded/pass-through where exposed; UI activation and arbitrary extension semantics are not implemented. |
| Unknown package parts and extension XML | 🟡 | ✅ | 🟡 | OPC graph editing preserves supported opaque parts and validates relationship topology. Typed writers do not claim lossless preservation through every semantic mutation of unknown XML. |

## Explicit gaps

| Feature family exposed by [MS-XLSX] | Status | Read | Write | Notes |
|-------------------------------------|--------|------|-------|-------|
| Rich values and modern rich-data objects | 🟡 | ✅ | ✅ | `rich_values` owns typed `rvData`, `rvStructures`, and `arrayData`; rich styles, supporting bags, type metadata, and web-image payloads remain bounded opaque documents. Package snapshots retain complete relationship topology. |
| Feature property bags and checkbox/XF-control extensions | 🟡 | ✅ | ✅ | Typed `FeaturePropertyBags`, checkbox defaults, `XFControls`, `XFComplement`, `XFComplements`, and `DXFComplements` validation are inert; unknown XML and relationship IDs/topology are retained. |
| Python in Excel and external code services | ❌ | ❌ | ❌ | Python environments, scripts, parameter encodings, external-code-service parts, and execution are not implemented. |
| Complete PivotTable refresh/calculation | ❌ | ❌ | ❌ | Pivot metadata can be read or edited in bounded scope, but caches are not refreshed, cube connections are not queried, and report layouts are not calculated. |
| Complete chart/drawing rendering and extension grammar | ❌ | ❌ | ❌ | Typed resources do not constitute an Excel renderer or unrestricted support for every chart, DrawingML, slicer, timeline, or producer extension schema. |
| Complete Excel formula evaluation | ❌ | ❌ | ❌ | Formula text and cached values are supported, but there is no complete dependency graph, spill engine, volatile recalc, external function, Python, DDE, or query execution. |
| External-link, connection, query, DDE, OLE, and add-in execution | ❌ | ❌ | ❌ | Targets and metadata are inert; no network, filesystem, COM, credential, or provider action is performed. |
| Macro execution and trusted ActiveX behavior | ❌ | ❌ | ❌ | VBA and control payloads remain inert package data and are never loaded into a host runtime. |
| Certificate trust and revocation | ❌ | ❌ | ❌ | Signature verification can establish cryptographic integrity/signature validity only, not trust-chain or revocation status. |
