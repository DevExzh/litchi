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
| Unified package transactions and durable patches | ✅ | ✅ | ✅ | `document::{Snapshot, Edit, Commit, Patch, ThreeWayPlan, History}` composes worksheet, rich cell runs, row/column/sheet structure, automatic style graphs, conditional formats, sparklines, typed inert forms/events, positioned/grouped text drawings, typed package-backed charts and dependency transfer, definition, annotation, RDF, protection, DataPilot, tracked-change, and bounded resource edits into one clone-staged atomic publication. Fine-grained XML changes use exact-source provenance splices, while patches retain canonical JSON, inverse replay, deterministic conservative joins, explicit three-way resolution, finite history, and full reopen. Changed signed input requires explicit stale-signature stripping; encrypted or already-protected input remains refused |
| `content.xml` access | ✅ | ✅ | ✅ | Raw content XML is exposed and can be supplied to the minimal builder; semantic sheet CRUD is not implied |
| `styles.xml` access | 🟡 | ✅ | 🟡 | Styles XML is available for inspection and package preservation; no public ODS style-graph editor is attached to the facade |
| `meta.xml` and `settings.xml` | 🟡 | ✅ | 🟡 | `meta.xml` has a typed snapshot plus bounded common-field CRUD; literal application `settings.xml` remains inert and is preserved through package edits |
| ODF metadata and document statistics | 🟡 | ✅ | 🟡 | `Spreadsheet::odf_metadata()` exposes the shared typed ODF model; the facade edits the common projection through retained-source patches while extended fields remain lossless/read-only |
| Common styles and data styles | 🟡 | 🟡 | 🟡 | ODF common style/data-style models cover number, date/time, currency, text, percentage, table-cell, page, master, and related vocabulary. The unified root can add or same-family replace a closed automatic number/text/table-cell graph, conservatively remove unreferenced automatic styles, and resolve the supported effective cell/text projection across common and automatic package styles |
| Images and package media inventory | ✅ | ✅ | 🟡 | Package, inline, missing, and external image references are discovered; verified package/inline bytes can be read. The unified transaction can add, transfer, replace, and remove bounded inert package resources with explicit collision policy and atomically publish named or positioned image-frame references |
| Embedded objects and OLE-like payload inventory | 🟡 | ✅ | 🟡 | Package, linked, flat, and missing object sources can be classified as inert resources; payloads are not opened, executed, rendered, or semantically edited |
| Embedded charts | ✅ | ✅ | ✅ | `charts::{Snapshot, Edit, Commit, Patch}` and `MutableSpreadsheet::{chart_snapshot, edit_charts, apply_chart_patch}` provide exact-name/checked-position selection and replacement. The unified root can serialize a typed chart definition, create its manifest-backed subdocument and drawing occurrence atomically, or transfer a compact chart content root and every relative package dependency to a new object root. Publications require compact XML, typed readback, and reversible exact-source patches. Rendering remains outside this slice |
| Annotations/comments | ✅ | ✅ | ✅ | Shared annotation trees plus contextual `annotations::{Snapshot, Transaction, Commit, Patch}` provide bounded rich bodies, exact sheet/cell selection, source-preserving CRUD, and failure-atomic package publication; non-cell collaboration semantics remain outside ODS |
| Hyperlinks and external references | 🟡 | 🟡 | 🟡 | Typed ODF/XLink values and ODS source metadata preserve targets and activation hints; links are never followed, fetched, or refreshed |
| Forms and controls | 🟡 | 🟡 | 🟡 | Unified transactions can replace bounded button/check-box/radio/list-box/combo-box/text/image/date/time controls and atomically publish checked linked-cell/source-range bindings, optional package image dependencies, and inert event/macro metadata using compact provenance-spliced XML. Events and macros are never executed, and the full ODF control vocabulary remains unsupported |
| Scripts, events, and macros | 🟡 | 🟡 | 🟡 | Script/event artifacts may remain inert package data; no script, macro, or event execution is performed |
| RDF metadata graphs | ✅ | ✅ | ✅ | `metadata_graphs::{Snapshot, Edit, Commit, Patch}` provides clone-staged graph/triple CRUD, checked triple positions, full package readback, compact authored XML, exact no-op retention, and reversible exact-source patches; `Spreadsheet` and `MutableSpreadsheet` retain concise compatibility verbs |
| ODF encryption | ❌ | ❌ | ❌ | The shared ODF layer contains password/encryption primitives, while the ODS unified writer exposes explicit `Refuse` and `DecryptAndReencrypt` intents. The latter returns a typed capability refusal because this root has no credential-bearing writer; snapshot capabilities distinguish not-applicable, credential-required decryption refusal, and unsupported re-encryption |
| ODF digital signatures | 🟡 | 🟡 | 🟡 | Shared XMLDSig inspection remains available. Changed unified publications default to refusal and require `SignatureWritePolicy::StripInvalidated` to deliberately omit stale signature containers. Snapshot capabilities explicitly report signing/resigning as unsupported |
| Unknown package-part preservation | 🟡 | ✅ | 🟡 | Unified transactions retain unrelated members while supported owners and bounded resources are edited; explicit resource removal conservatively refuses paths referenced by retained XML. Preservation and transfer do not imply semantic understanding, relationship authoring, or execution of payloads |

## Spreadsheet data model

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Sheets, rows, columns, and cells | ✅ | ✅ | ✅ | Public `worksheet::{Sheet, Row, Cell, Snapshot, Edit, Commit, Patch}` graph with logical lookup over physical repetition runs, exact-name/checked-position selection, clone-staged structure/cell CRUD, compact authored XML, typed readback, and exact-source reversible patches; `Builder` and `MutableSpreadsheet` retain concise authoring verbs |
| Cell scalar types | ✅ | ✅ | ✅ | Typed string, number, boolean, date/time, percentage, currency, and unknown value tokens are validated and serialized without expanding repeats |
| Rich cell text, spans, whitespace, fields, and hyperlinks | 🟡 | 🟡 | 🟡 | `document::{RichText, RichRun}` authors bounded paragraphs, text spans, explicit spaces, tabs, line breaks, and inert safe hyperlinks in existing non-repeated cells through provenance splices. The ordinary worksheet projection remains collapsed plain text, and arbitrary fields/foreign inline markup are not structurally editable |
| Formula strings and references | ✅ | ✅ | ✅ | Public `codec::formula` parses and represents OpenFormula-like functions, literals, A1 references, ranges, operators, and sheet references; worksheet and unified transactions can retain, set, or clear inert formula strings without evaluating dependencies |
| Formula evaluation and recalculation | ❌ | ❌ | N/A | No ODS evaluation adapter is exported; formulas and cached values remain inert and no recalculation, external I/O, or refresh occurs |
| Repeated rows/cells and merged/covered cells | ✅ | ✅ | ✅ | Physical repeated runs are retained, logical lookup/edit splits only the affected run, and merge/covered metadata is typed and validated |
| Named ranges and named expressions | ✅ | ✅ | ✅ | Global and sheet-scoped definitions are parsed and validated per ODF 1.4 Part 3 sections 9.4.11-9.4.13; `definitions::{Snapshot, Edit, Commit, Patch}` adds exact key/checked-position CRUD and reorder, compact authored XML, typed readback, exact no-op retention, and reversible exact-source patches alongside the minimal builder |
| Cell and sheet styles | 🟡 | 🟡 | 🟡 | Typed source structures cover alignment, borders, background, number/data styles, text properties, and protection flags. Unified transactions can add or same-family replace bounded automatic graphs, conservatively remove styles with no retained references, apply cell/rich-run references, and resolve supported inherited properties across `styles.xml` plus content automatic styles. Full property vocabulary and effective layout remain unsupported |
| Conditional cell styles (`style:map`) | 🟡 | 🟡 | 🟡 | Typed inert conditions and apply-style names are modeled and preserved; conditions are not evaluated and the package editor is not exposed |
| Sheet conditional formatting | 🟡 | 🟡 | 🟡 | `Spreadsheet::source_features` inventories bounded inert `calcext:conditional-format` occurrences, and the unified root can replace/remove a sheet catalog using the typed condition/color-scale/data-bar/icon/date model. Rules are never evaluated and effective styles are not resolved |
| Sparklines | 🟡 | 🟡 | 🟡 | `Spreadsheet::source_features` inventories bounded inert `calcext:sparkline-group` occurrences, and the unified root can replace/remove a sheet catalog using the typed LibreOffice `calcext`/`loext` model. Data ranges remain inert and no calculation or rendering occurs |
| Content validation | 🟡 | 🟡 | 🟡 | Conditions, prompts, error messages, events, definitions, and bindings are represented in source models; validation is not enforced by a public workbook editor |
| Comments attached to cells | ✅ | ✅ | ✅ | `Spreadsheet::annotations` and `MutableSpreadsheet::edit_annotations` expose checked zero-based sheet/cell selectors, add/replace/remove operations, exact no-op package replay, and unrelated XML preservation |
| Database ranges, filters, sorts, and subtotals | 🟡 | 🟡 | 🟡 | Recursive filter expressions, sort keys, subtotal groups, and inert database/source metadata are typed in source models; no query execution or public range editor is available |
| DataPilot/pivot tables | ✅ | ✅ | ✅ | `data_pilot::{Catalog, Snapshot, Edit, Commit, Patch}` exposes typed sources, fields, levels, references, grouping, display, layout, sorting, and grand-total metadata; the mutable facade provides clone-staged CRUD and exact-source patch application with typed readback and unknown-owned-XML refusal. Pivot calculation, refresh, rendering, and external-source execution remain inert |
| Sheet/document protection | ✅ | ✅ | 🟡 | `protection::{Snapshot, Transaction, Commit, Patch}` exposes source-checked document/sheet flags, inert verifier metadata, LibreOffice permissions, automatic cell-protection styles, conditional rules, failure-atomic `content.xml` edits, and reversible exact-source patch application through the mutable facade; passwords are never verified and policy is not enforced |
| Print ranges and page settings | 🟡 | 🟡 | 🟡 | Sheet printability, print ranges, page style references, row/column visibility, grouping, and table structure are represented in source models; no pagination or rendering is performed |
| Calculation settings and iteration | ✅ | ✅ | ✅ | `settings::{Snapshot, Transaction, Editor}` and the spreadsheet facade provide bounded CRUD for `table:calculation-settings`; this does not make formula evaluation complete |
| Consolidations, label ranges, scenarios, and detective metadata | 🟡 | 🟡 | ❌ | `scenario::Snapshot` provides bounded, source-bound inspection of required ranges/state and optional display/copy/protection metadata without applying scenario values; consolidation, label-range, and detective source models remain internal and no public CRUD or calculation is supplied |
| DDE sources and cached tables | 🟡 | ✅ | ❌ | `dde::Snapshot` provides bounded, source-bound inspection of sheet sources, formula-link sources, conversion/update flags, and exact cached table XML; DDE remains inert and no conversation, refresh, source access, or public mutation exists |
| External linked tables/ranges | 🟡 | 🟡 | 🟡 | URI, dimensions, filter, refresh-delay, and actuate metadata are retained without dereferencing or refreshing the external document |
| In-table shapes and images | 🟡 | 🟡 | 🟡 | `Spreadsheet::source_features` and `Spreadsheet::images` inventory drawings and image sources. The unified root atomically adds/removes named image frames, adds finite positioned/anchored frames with exact resources, bounded rich-text groups, and positioned rectangle/ellipse/line geometry, and transfers dependencies while rewriting the destination `xlink:href`; arbitrary paths and linked-resource fetching remain unsupported |
| Tracked spreadsheet changes | ✅ | ✅ | ✅ | `tracked_changes::{Snapshot, Transaction, Commit, Patch, Limits}` and the spreadsheet facade provide bounded ODF 1.4 owner-state inspection, all four record families, source-preserving CRUD/reorder and acceptance-metadata edits, exact no-ops, reversible source-checked patches, and auxiliary-part preservation. Foreign extension records and rich/foreign markup inside recognized records are opaque, retained verbatim, and never interpreted; regeneration is refused when it cannot preserve them safely. Changed package publication drops invalidated document and macro signatures, while encrypted rewrites are refused. Formulas, links, and historical cell values remain inert: this API does not apply accepted/rejected records to live cells, rows, columns, or tables |
| CSV export | ❌ | ❌ | ❌ | No public ODS-to-CSV export utility is exposed by this crate |

## Explicit gaps from the audited feature families

| Feature family | Status | Read | Write | Notes |
|----------------|--------|------|-------|-------|
| Full public worksheet authoring | 🟡 | ✅ | ✅ | Bounded worksheet and unified transactions cover sheet/cell CRUD, row append/removal, column declaration append/removal, formulas/styles, and checked rich text. Repeated-run structural splitting, arbitrary column/row grouping, calculation, rendering, and the complete extension vocabulary remain outside this slice |
| Full OpenFormula semantics | ❌ | 🟡 | ❌ | The public parser is not an evaluator; unsupported grammar, external workbook references, dynamic arrays, volatile behavior, data tables, and host-service functions are not resolved |
| XLSX slicers and timelines | ❌ | ❌ | ❌ | `[MS-XLSX]` ToC families for slicer caches, slicers, timelines, and their extension parts are OOXML features, not typed ODS support |
| XLSX PivotTable/data-model extensions | ❌ | ❌ | ❌ | OOXML PivotTable caches, OLAP/data-model structures, rich pivot data, pivot UI/version/auto-refresh extensions, and related XML parts are not implemented; ODF DataPilot metadata is not an equivalent claim |
| XLSX external data and code services | ❌ | ❌ | ❌ | Query tables, connections, database/model tables, external links, web/rich values, external code services, Python environments/scripts, and refresh behavior are not supported or executed |
| XLSX threaded comments and rich values | ❌ | ❌ | ❌ | The `[MS-XLSX]` threaded-comment, mentions, rich-value, property-bag, and web-image families have no ODS typed counterpart in this crate |
| Rendering, pagination, and recalculation | ❌ | ❌ | ❌ | No spreadsheet layout engine, visual renderer, chart/sparkline renderer, print pagination engine, or automatic recalculation service is included |
| Security policy enforcement | 🟡 | ✅ | 🟡 | Unified changed publication explicitly refuses encrypted input and protected structures, distinguishes refusal from caller-requested but unavailable decrypt/re-encrypt publication, defaults to refusing signed input, and permits deliberate stale-signature stripping. Password unlock, signing, and re-encryption are not implemented |
