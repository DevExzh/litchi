# Excel Binary Workbook (.xlsb) Feature Matrix

This document tracks the public feature families implemented by litchi-xlsb. It describes
typed support for the BIFF12 binary parts and their OPC package, not Excel rendering,
recalculation, network access, macro execution, or complete conformance with every extension.

The audit uses the authoritative Microsoft [MS-XLSB] Front Matter and ToC, the file and
conceptual structures in [MS-XLSB] 2.1-2.2, and the examples for conditional formatting,
defined names, tables, filters, external references, formatting, workbooks, PivotTables,
metadata, and slicers in [MS-XLSB] 3.1-3.10.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the scope described in the Notes cell |
| 🟡 | Bounded, partial, metadata-only, pass-through, or inert support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to this format or direction |

Read and Write are independent. Typed records and package parts do not imply that
formula results are recalculated or that external, macro, or embedded content is run.

## Package, BIFF12, workbook, and core data

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| OPC package and BIFF12 part graph | ✅ | ✅ | ✅ | Workbook, Package, relationships, content types, binary parts, and package saves are modeled over the OOXML package graph. |
| BIFF12 raw record codec | ✅ | ✅ | ✅ | litchi-xlsb::raw enforces BIFF12 kind and length encodings, distinguishes clean EOF from truncation, lends payloads, validates UTF-16 and Boolean fields, and writes checked records. |
| Workbook and worksheet binary streams | ✅ | ✅ | ✅ | Workbook, worksheet, chart-sheet, styles, shared-string, connection, external-link, pivot, table, and related BIFF12 streams are read and written through the package model. |
| Workbook and sheet directory | ✅ | ✅ | ✅ | Sheet names, order, indices, visibility, sheet lookup, and workbook lifecycle operations are typed. |
| Cell values and cached results | ✅ | ✅ | ✅ | Numbers, strings, shared strings, booleans, errors, blanks, serial dates, and style references are supported with checked binary encodings. |
| Shared strings | ✅ | ✅ | ✅ | Shared-string tables and cell references are decoded and emitted; rich data objects beyond the shared-string model are not claimed. |
| Formula tokens and cached formula values | 🟡 | ✅ | 🟡 | BIFF12 formula token models, formula references, table/pivot references, and supported formula text are available. The formula surface is bounded, unsupported expressions can be refused, and there is no Excel-compatible evaluator. |
| Array, shared, and dynamic-array formulas | 🟡 | ✅ | 🟡 | Supported array/shared formula records and cached values are retained in bounded form; modern spill calculation and dependency propagation are not implemented. |
| Defined names and named ranges | ✅ | ✅ | ✅ | Workbook and sheet-scoped names, name formulas, and validated named-range authoring are available. Names are not evaluated as external actions. |
| Calculation properties | ✅ | ✅ | ✅ | The specified BrtCalcProp form and the bounded early-artifact option-tail form are read; canonical 26-byte output, mode, reference mode, iteration, delta, thread count, and recalculation flags are validated. |
| Styles, fonts, fills, borders, alignment, and number formats | ✅ | ✅ | ✅ | Typed style tables, cell formats, differential formats, fonts, colors, borders, fills, alignment, and number formats are supported. |
| Workbook and worksheet views | 🟡 | ✅ | 🟡 | Sheet views, panes, selections, zoom, tab/window metadata, and chart-sheet views have typed coverage; all producer-specific view extensions are not modeled. |
| Merged cells | ✅ | ✅ | ✅ | Worksheet merge ranges have checked read and CRUD operations, including lookup by sheet name. |
| Row and column layout | ✅ | ✅ | ✅ | Row/column dimensions, hidden and collapsed state, outline levels, widths, heights, and supported worksheet defaults are available. |
| Core/package properties | 🟡 | 🟡 | 🟡 | Generic OPC package editing is public; there is no broad XLSB-specific typed property facade comparable to the binary package and workbook models. |

## Worksheet features

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Worksheet cell CRUD | ✅ | ✅ | ✅ | Typed worksheet snapshots and writer models support cell insertion, lookup, value changes, and sheet serialization. |
| Hyperlinks | ✅ | ✅ | ✅ | BIFF12 hyperlink targets, display/location metadata, and bounded authoring are supported; links are never followed. |
| Classic comments and notes | ✅ | ✅ | ✅ | Worksheet comment records have a typed model and writer support. Comment text and anchors are data only. |
| Threaded comments and persons | ✅ | ✅ | ✅ | Contextual `comments::threaded::{Snapshot, Transaction, Commit, Patch}` owns bounded people/thread/mention metadata, worksheet relationship validation, source-checked CRUD, exact no-op replay, and inert package publication; collaboration services are never contacted. |
| Data validation | ✅ | ✅ | ✅ | Validation ranges and supported rule/formula forms are read and written through the worksheet/package writer. |
| Conditional formatting | ✅ | ✅ | ✅ | Classic rules, differential formatting, formulas, ranges, icon sets, color scales, data bars, and implemented extension/FRT forms are validated and serialized. Formulas are not evaluated. |
| AutoFilter, filter criteria, and sort state | 🟡 | ✅ | 🟡 | Worksheet AutoFilter and criteria metadata are modeled; the complete Excel filter UI and every producer extension are not a semantic engine. |
| Structured tables | ✅ | ✅ | ✅ | List/table identity, ranges, columns, header/totals metadata, style IDs, typed totals functions, inert calculated-column formulas, and worksheet wiring are serialized. |
| Page setup and print metadata | 🟡 | ✅ | 🟡 | Supported worksheet and chart-sheet page setup, margins, paper, scale/fit, orientation, print flags, and relationships are bounded; printer-driver/rendering behavior is not implemented. |
| Sheet protection | 🟡 | ✅ | 🟡 | Worksheet and chart-sheet protection flags and the supported strong-protection metadata are typed. Protection is not a cryptographic authorization boundary. |
| Scenarios and what-if analysis | 🟡 | ✅ | ✅ | `litchi_xlsb::scenarios` exposes bounded `BrtBeginScenMan`/`BrtBeginSct`/`BrtSlc` snapshots and transactional worksheet replacement. Known metadata is typed; unknown records and source order are retained, unsafe or ambiguous edits are refused, and scenario values are never substituted or recalculated. |
| Cell watches and phonetic metadata | ✅ | ✅ | ✅ | `[MS-XLSB]` 2.4.21, 2.4.331, 2.4.378, and 2.4.744 are exposed as typed worksheet cell-watch collections and worksheet-wide phonetic defaults with bounded, source-checked transactional edits; watch-window monitoring, phonetic rendering, and language conversion are not performed. |
| Shared-workbook revision records | 🟡 | ✅ | ✅ | `shared_workbook::{Snapshot, Transaction, Commit, Patch}` owns bounded user, header, and inert revision-log metadata with package relationship validation, source-checked CRUD, unknown-record retention, and failure atomicity; collaboration, locking, conflict resolution, and recalculation remain inactive. |

## Charts, drawings, pivots, and package objects

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Chart sheets | 🟡 | ✅ | 🟡 | Chart-sheet names, visibility, views, protection, page setup, colors, and selected chart-resource relationships are modeled; this is not a complete chart grammar or renderer. |
| Embedded charts and chart resources | 🟡 | ✅ | 🟡 | Bounded chart resource discovery, embedded chart relationships, and selected OOXML chart/pivot-chart payloads are available. Arbitrary chart graph edits and full chart authoring are not claimed. |
| Worksheet drawings and anchors | 🟡 | ✅ | 🟡 | Typed drawing/anchor inventories, image relationships, and supported shape resources are available. Complex OfficeArt/OOXML shape semantics, arbitrary text geometry, and rendering remain bounded. |
| Images | 🟡 | ✅ | ✅ | Supported image formats can be read as drawing resources and authored through the writer image model; no image renderer or Office layout engine is included. |
| OLE and embedded package objects | 🟡 | ✅ | ❌ | Bounded inventory recognizes worksheet, macro-sheet, dialog-sheet, and external-link object policies; payload bytes remain inert and general binary object authoring is not implemented. |
| Pivot caches | ✅ | ✅ | ✅ | Typed inert PivotCache definitions cover refresh metadata, worksheet/consolidation sources, all shared-item value kinds, range/discrete grouping, OLAP hierarchies and tuple caches, calculated items/members with inert formula tokens, and supported Excel extensions. |
| PivotTable views | ✅ | ✅ | ✅ | Pivot views, fields, items, layouts, data fields, filters, and workbook cache-ID wiring are read and serialized with lossless-or-refuse validation. No pivot refresh or calculation is performed. |
| Pivot charts | 🟡 | ✅ | 🟡 | Pivot-chart resources and relationships are recognized through the chart package model; chart rendering and full pivot-chart authoring are bounded. |
| Slicers and slicer caches | 🟡 | ✅ | ✅ | Bounded native/table BIFF12 cache definitions, worksheet slicer views, workbook/sheet relationship wiring, and transactional inert CRUD are supported. OLAP item-range authoring and filter execution remain outside this slice. |
| Timelines | 🟡 | ✅ | ✅ | Bounded timeline cache/view XML, BIFF12 workbook/sheet relationship wiring, and transactional inert CRUD are supported. Extension payloads, active filtering, and refresh remain outside this slice. |

## External references, connections, and extensibility

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| External workbook links | ✅ | ✅ | 🟡 | `external_link::{Snapshot, Transaction, Commit, Patch}` adds source-checked workbook/DDE/OLE metadata, external-name, flag, and cached-matrix edits while preserving opaque BIFF12 records and validated relationship topology; unsupported link forms are refused. |
| External formulas and cached values | 🟡 | ✅ | 🟡 | The five permitted external token structures and checked Xnum/worksheet-dimension rules are covered. Paths, names, formulas, and cached values are inert and are never resolved or recalculated. |
| DDE and OLE links | 🟡 | ✅ | 🟡 | DDE/OLE topics, program IDs, flags, and items are typed metadata only; no conversation, COM activation, refresh, or external launch occurs. |
| External data connections | ✅ | ✅ | ✅ | `host::connections::{Snapshot, Transaction, Commit, Patch}` adds source-checked create/update/remove and parameter edits over the BIFF12 External Data Connections part, preserving opaque records and OPC graph topology. Connection strings, commands, URLs, paths, and credentials are never contacted or executed. |
| Query tables and web imports | 🟡 | ✅ | 🟡 | Connection/query-table metadata is bounded and inert; no network fetch, import, refresh, or credential use is performed. |
| Ribbon and custom UI parts | 🟡 | ✅ | 🟡 | Ribbon families and bounded XML payloads can be inspected, replaced, and removed through package APIs; UI behavior is not executed. |
| Task panes and web-extension bindings | 🟡 | ✅ | 🟡 | Task-pane graphs and worksheet binding metadata have bounded package support. Add-in activation and provider-specific behavior are not implemented. |
| VBA project and code modules | 🟡 | 🟡 | ✅ | vbaProject.bin CFB/MS-OVBA project/module payloads and legacy/Agile signature-part metadata can be inventoried and authored with bounded project operations. Source is never executed or trusted; replacement drops stale project signatures. |
| OPC digital signatures | ✅ | ✅ | ✅ | Workbook signing, re-signing, verification, and unsigning are exposed through the OPC signature integration. Verification establishes integrity/signature validity only, not certificate trust or revocation. |
| Password encryption | ✅ | ✅ | ✅ | Supported profile package encryption/decryption is available; the dedicated high-level safe-save facade remains separate from the core workbook editor. |
| Generic package part editing | ✅ | ✅ | ✅ | The underlying OPC graph can edit parts and relationships transactionally, including pass-through payloads. Typed XLSB models still refuse unsafe or semantically unknown mutations. |

## Explicit gaps

| Feature family exposed by [MS-XLSB] | Status | Read | Write | Notes |
|-------------------------------------|--------|------|-------|-------|
| Complete Excel formula evaluation and recalculation | ❌ | ❌ | ❌ | Formula tokens and cached values are data; there is no Excel-compatible dependency graph, volatile refresh, external-function execution, or pivot calculation engine. |
| Rich values, rich-data metadata, and modern data types | ❌ | ❌ | ❌ | The [MS-XLSB] metadata examples include MDX and cell/value metadata families, but no public rich-value model is provided. |
| MDX/cell/value metadata semantics | 🟡 | 🟡 | 🟡 | Raw/package records may be retained or inspected through lower-level paths, but cube metadata is not a typed semantic model and MDX is never evaluated. |
| Complete chart grammar and rendering | ❌ | ❌ | ❌ | Chart-sheet/resource coverage is deliberately bounded; the crate does not provide an Excel renderer or unrestricted chart graph editor. |
| Slicers, timelines, and their filter behavior | 🟡 | ✅ | ✅ | Typed cache/view snapshots and safe transactional CRUD are available; selection/filter behavior, refresh, calculation, and unsupported opaque structures are intentionally inert or refused. |
| XML maps and mapped XML import/export | ❌ | ❌ | ❌ | No XLSB XML-map/data-binding authoring model is exposed. |
| ActiveX/form-control execution | ❌ | ❌ | ❌ | Drawing and embedded-object payloads are inert; controls are not instantiated, activated, or run. |
| Macro execution and external-link/connection refresh | ❌ | ❌ | ❌ | VBA, DDE, OLE, URLs, commands, credentials, and cached external data are never executed, fetched, or refreshed. |
| Certificate trust and revocation | ❌ | ❌ | ❌ | Signature verification is cryptographic integrity/signature verification only. |
