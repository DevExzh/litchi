# ODS/OTS Feature Matrix

This document tracks the public and source-level feature coverage of the
`litchi-ods` OpenDocument Spreadsheet implementation for packaged ODS and OTS
files. It is a capability matrix, not a claim of complete ODF conformance,
formula compatibility, or rendering fidelity.

The ODF common package layer is shared with the other OpenDocument families.
Rows below call out when a capability is only available as inert metadata,
package preservation, a public scalar codec, or a typed source model that is
not yet connected to the ODS facade.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the feature scope described in the Notes cell |
| 🟡 | Bounded, partial, metadata-only, pass-through, source-level, or otherwise limited support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to the format or direction |

`Read` and `Write` describe the public direction independently. A 🟡 direction
must not be read as full semantic CRUD. External links, database sources,
DDE, scripts, macros, and embedded payloads are inert unless a row says
otherwise.

## Audit scope

The ODS-specific source vocabulary includes cell, row, sheet, formula,
conditional-format, validation, DataPilot, database-range, protection,
scenario, consolidation, detective, sparkline, DDE, and tracked-change models.
The current crate exports the package facade, named-definition APIs, formula
codec, resource inventories, RDF APIs, and the bounded worksheet graph through
`litchi_ods::worksheet`, `Spreadsheet`, `Builder`, and
`MutableSpreadsheet`. This is a worksheet graph, not a full calculation or
rendering engine.

The Microsoft `[MS-XLSX]` Front Matter and ToC describe extensions to OOXML
SpreadsheetML, not ODS. Their Part Enumerations, Extensions, Conceptual
Overview, and listed structure families were used as a gap checklist only;
the existence of an XLSX feature is never treated as ODS support.

## Package and shared ODF features

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Open, create, and save ODS/OTS packages | ✅ | ✅ | ✅ | `Package` and `Spreadsheet` open paths or bytes; `Builder` creates a valid spreadsheet package; `MutableSpreadsheet` saves a package after the supported package edits |
| ZIP package, manifest, and MIME validation | ✅ | ✅ | ✅ | Shared ODF package layer validates the manifest and spreadsheet/template media types and writes deterministic package entries |
| `content.xml` access | ✅ | ✅ | ✅ | Raw content XML is exposed and can be supplied to the minimal builder; semantic sheet CRUD is not implied |
| `styles.xml` access | 🟡 | ✅ | 🟡 | Styles XML is available for inspection and package preservation; no public ODS style-graph editor is attached to the facade |
| `meta.xml` and `settings.xml` | 🟡 | ✅ | 🟡 | `meta.xml` has a typed snapshot plus bounded common-field CRUD; literal application `settings.xml` remains inert and is preserved through package edits |
| ODF metadata and document statistics | 🟡 | ✅ | 🟡 | `Spreadsheet::odf_metadata()` exposes the shared typed ODF model; the facade edits the common projection through retained-source patches while extended fields remain lossless/read-only |
| Common styles and data styles | 🟡 | 🟡 | 🟡 | ODF common style/data-style models cover number, date/time, currency, text, percentage, table-cell, page, master, and related vocabulary; ODS package integration is bounded/source-level |
| Images and package media inventory | ✅ | ✅ | 🟡 | Package, inline, missing, and external image references are discovered; verified package/inline bytes can be read; adding/replacing image parts requires raw package/content authoring |
| Embedded objects and OLE-like payload inventory | 🟡 | ✅ | 🟡 | Package, linked, flat, and missing object sources can be classified as inert resources; payloads are not opened, executed, rendered, or semantically edited |
| Embedded charts | ✅ | ✅ | ✅ | `Spreadsheet::charts` exposes bounded content-level chart inventory with exact-name/checked-position selectors; `MutableSpreadsheet::edit_charts` clone-stages safe replacement of package-backed or inline chart parts, retains unknown chart XML, preserves unrelated package members, and keeps exact bytes for no-op edits. Full chart-series CRUD and rendering remain outside this slice |
| Annotations/comments | ✅ | ✅ | ✅ | Shared annotation trees plus contextual `annotations::{Snapshot, Transaction, Commit, Patch}` provide bounded rich bodies, exact sheet/cell selection, source-preserving CRUD, and failure-atomic package publication; non-cell collaboration semantics remain outside ODS |
| Hyperlinks and external references | 🟡 | 🟡 | 🟡 | Typed ODF/XLink values and ODS source metadata preserve targets and activation hints; links are never followed, fetched, or refreshed |
| Forms and controls | 🟡 | 🟡 | 🟡 | Common ODF form vocabulary can be retained as package/XML metadata; no public ODS form-control editor is exposed |
| Scripts, events, and macros | 🟡 | 🟡 | 🟡 | Script/event artifacts may remain inert package data; no script, macro, or event execution is performed |
| RDF metadata graphs | ✅ | ✅ | ✅ | Graph and triple inventory plus ordered add/replace/remove/move operations are exposed by `Spreadsheet` and `MutableSpreadsheet` |
| ODF encryption | ❌ | ❌ | ❌ | The shared ODF layer contains password/encryption primitives, but the ODS wrapper exposes no password-open, encrypt, or password-change operation; encrypted sources must not be inferred to work from the common dependency alone |
| ODF digital signatures | ❌ | ❌ | ❌ | Shared XMLDSig parsing and cryptographic support exists, but the ODS public package facade has no sign/verify/add/clear signature API |
| Unknown package-part preservation | 🟡 | ✅ | 🟡 | The owned package can retain unrelated parts while supported package edits are made; preservation is not semantic understanding or guaranteed lossless editing of every extension |

## Spreadsheet data model

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Sheets, rows, columns, and cells | ✅ | ✅ | ✅ | Public `worksheet::{Sheet, Row, Cell}` graph with logical lookup over physical repetition runs; `Builder` and `MutableSpreadsheet` provide atomic add/remove/set/clear operations |
| Cell scalar types | ✅ | ✅ | ✅ | Typed string, number, boolean, date/time, percentage, currency, and unknown value tokens are validated and serialized without expanding repeats |
| Rich cell text, spans, whitespace, fields, and hyperlinks | 🟡 | 🟡 | 🟡 | Structure-preserving source models retain mixed text and inert hyperlinks; no public cell range API exposes or mutates them |
| Formula strings and references | ✅ | ✅ | ✅ | Public `codec::formula` parses and represents OpenFormula-like functions, literals, A1 references, ranges, operators, and sheet references; this codec is independent of package cell CRUD |
| Formula evaluation and recalculation | 🟡 | 🟡 | N/A | An ODS evaluation adapter source normalizes common `of:=`/A1/semicolon syntax for the shared evaluator, but the adapter is not part of the current exported codec module; full OpenFormula semantics, recalculation, rendering, and external I/O are not provided |
| Repeated rows/cells and merged/covered cells | ✅ | ✅ | ✅ | Physical repeated runs are retained, logical lookup/edit splits only the affected run, and merge/covered metadata is typed and validated |
| Named ranges and named expressions | ✅ | ✅ | ✅ | Global and sheet-scoped definitions are parsed, validated, ordered, looked up, added, replaced, and removed through the package facade and minimal builder |
| Cell and sheet styles | 🟡 | 🟡 | 🟡 | Typed source structures cover alignment, borders, background, number/data styles, text properties, and protection flags; style-use resolution and public application to cells are not implemented |
| Conditional cell styles (`style:map`) | 🟡 | 🟡 | 🟡 | Typed inert conditions and apply-style names are modeled and preserved; conditions are not evaluated and the package editor is not exposed |
| Sheet conditional formatting | 🟡 | 🟡 | 🟡 | LibreOffice `calcext` condition, color-scale, data-bar, icon-set, custom-icon, and date-is models are bounded and inert; no public sheet CRUD or visual evaluation is available |
| Sparklines | 🟡 | 🟡 | 🟡 | LibreOffice `calcext:sparkline-groups` and `loext` theme-color transformations have typed inert models; no rendering, calculation, or public sheet integration is available |
| Content validation | 🟡 | 🟡 | 🟡 | Conditions, prompts, error messages, events, definitions, and bindings are represented in source models; validation is not enforced by a public workbook editor |
| Comments attached to cells | ✅ | ✅ | ✅ | `Spreadsheet::annotations` and `MutableSpreadsheet::edit_annotations` expose checked zero-based sheet/cell selectors, add/replace/remove operations, exact no-op package replay, and unrelated XML preservation |
| Database ranges, filters, sorts, and subtotals | 🟡 | 🟡 | 🟡 | Recursive filter expressions, sort keys, subtotal groups, and inert database/source metadata are typed in source models; no query execution or public range editor is available |
| DataPilot/pivot tables | ✅ | ✅ | ✅ | `data_pilot::Catalog` exposes typed sources, fields, levels, references, grouping, display, layout, sorting, and grand-total metadata; `MutableSpreadsheet::edit_data_pilots` provides clone-staged add/replace/update/remove CRUD with source-checked owner replacement and unknown-XML refusal. Pivot calculation, refresh, rendering, and external-source execution remain inert |
| Sheet/document protection | ✅ | ✅ | 🟡 | `protection::{Snapshot, Transaction, Styles}` exposes source-checked document/sheet flags, inert verifier metadata, LibreOffice permissions, automatic cell-protection styles, conditional rules, and failure-atomic `content.xml` edits; passwords are never verified and policy is not enforced |
| Print ranges and page settings | 🟡 | 🟡 | 🟡 | Sheet printability, print ranges, page style references, row/column visibility, grouping, and table structure are represented in source models; no pagination or rendering is performed |
| Calculation settings and iteration | ✅ | ✅ | ✅ | `settings::{Snapshot, Transaction, Editor}` and the spreadsheet facade provide bounded CRUD for `table:calculation-settings`; this does not make formula evaluation complete |
| Consolidations, label ranges, scenarios, and detective metadata | 🟡 | 🟡 | 🟡 | Typed inert models cover consolidation options, label ranges, what-if scenarios, formula-auditing highlights, and operations; no calculation, UI behavior, or public sheet CRUD is supplied |
| DDE sources and cached tables | 🟡 | 🟡 | 🟡 | DDE source/link declarations and cached tables can be represented as inert metadata; no DDE conversation, refresh, or source access occurs |
| External linked tables/ranges | 🟡 | 🟡 | 🟡 | URI, dimensions, filter, refresh-delay, and actuate metadata are retained without dereferencing or refreshing the external document |
| Tracked spreadsheet changes | ✅ | ✅ | ✅ | `tracked_changes::{Snapshot, Transaction, Commit, Patch, Limits}` and the spreadsheet facade provide bounded ODF 1.4 owner-state inspection, all four record families, source-preserving CRUD/reorder and acceptance-metadata edits, exact no-ops, reversible source-checked patches, and auxiliary-part preservation. Foreign extension records and rich/foreign markup inside recognized records are opaque, retained verbatim, and never interpreted; regeneration is refused when it cannot preserve them safely. Changed package publication drops invalidated document and macro signatures, while encrypted rewrites are refused. Formulas, links, and historical cell values remain inert: this API does not apply accepted/rejected records to live cells, rows, columns, or tables |
| CSV export | ❌ | ❌ | ❌ | No public ODS-to-CSV export utility is exposed by this crate |

## Explicit gaps from the audited feature families

| Feature family | Status | Read | Write | Notes |
|----------------|--------|------|-------|-------|
| Full public worksheet authoring | 🟡 | ✅ | ✅ | Bounded worksheet graph and transactional sheet/cell CRUD are public; advanced ODF table extensions, calculation, rendering, and rich-text editing remain outside this slice |
| Full OpenFormula semantics | ❌ | 🟡 | ❌ | The public parser is not an evaluator; unsupported grammar, external workbook references, dynamic arrays, volatile behavior, data tables, and host-service functions are not resolved |
| XLSX slicers and timelines | ❌ | ❌ | ❌ | `[MS-XLSX]` ToC families for slicer caches, slicers, timelines, and their extension parts are OOXML features, not typed ODS support |
| XLSX PivotTable/data-model extensions | ❌ | ❌ | ❌ | OOXML PivotTable caches, OLAP/data-model structures, rich pivot data, pivot UI/version/auto-refresh extensions, and related XML parts are not implemented; ODF DataPilot metadata is not an equivalent claim |
| XLSX external data and code services | ❌ | ❌ | ❌ | Query tables, connections, database/model tables, external links, web/rich values, external code services, Python environments/scripts, and refresh behavior are not supported or executed |
| XLSX threaded comments and rich values | ❌ | ❌ | ❌ | The `[MS-XLSX]` threaded-comment, mentions, rich-value, property-bag, and web-image families have no ODS typed counterpart in this crate |
| Rendering, pagination, and recalculation | ❌ | ❌ | ❌ | No spreadsheet layout engine, visual renderer, chart/sparkline renderer, print pagination engine, or automatic recalculation service is included |
| Security policy enforcement | ❌ | ❌ | ❌ | Protection flags are metadata and common encryption/signature code is not exposed as ODS end-to-end operations |
