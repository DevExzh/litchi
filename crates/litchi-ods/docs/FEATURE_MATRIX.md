# ODS/OTS and flat ODS Feature Matrix

Feature checklist for packaged ODS/OTS and flat single-XML ODS. ✅ Supported · 🟡 Partial ·
❌ Unsupported · N/A Not applicable. External links, database sources, DDE, scripts, macros and
embedded payloads are never executed, fetched or refreshed.

## Package and shared ODF features

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Open, create, and save ODS/OTS packages | ✅ | ✅ | ✅ | |
| Flat single-XML ODS (`.fods`) | 🟡 | ✅ | 🟡 | Cell edits only |
| ZIP package, manifest, and MIME validation | ✅ | ✅ | ✅ | |
| Unified package transactions and durable patches | ✅ | ✅ | ✅ | |
| `content.xml` access | ✅ | ✅ | ✅ | |
| `styles.xml` access | 🟡 | ✅ | 🟡 | Preserved; no style-graph editor |
| `meta.xml` and `settings.xml` | 🟡 | ✅ | 🟡 | Settings preserved, not edited |
| ODF metadata and document statistics | 🟡 | ✅ | 🟡 | Extended fields read-only |
| Common styles and data styles | 🟡 | 🟡 | 🟡 | Page, master and layout styles preserved, not authored |
| Table templates and banding (`table:table-template`) | 🟡 | ✅ | ✅ | Templates not applied |
| Images and package media inventory | ✅ | ✅ | 🟡 | |
| Embedded objects and OLE-like payload inventory | 🟡 | ✅ | 🟡 | Inert inventory only |
| Embedded charts | ✅ | ✅ | ✅ | |
| Annotations/comments | ✅ | ✅ | ✅ | |
| Hyperlinks and external references | 🟡 | 🟡 | 🟡 | Targets never followed |
| Forms and controls | 🟡 | 🟡 | 🟡 | No grid controls or producer extensions |
| Scripts, events, and macros | 🟡 | 🟡 | 🟡 | Inert package data only |
| RDF metadata graphs | ✅ | ✅ | ✅ | |
| ODF encryption | 🟡 | ✅ | 🟡 | Encrypted packages are not editable |
| ODF digital signatures | 🟡 | 🟡 | 🟡 | No signing; changed saves drop signatures only on request |
| Unknown package-part preservation | 🟡 | ✅ | 🟡 | Preserved, not understood |

## Spreadsheet data model

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Sheets, rows, columns, and cells | ✅ | ✅ | ✅ | |
| Positional existing-cell publication | 🟡 | ✅ | 🟡 | Ordinary scalar cells only |
| Cross-document scalar worksheet transfer | 🟡 | ✅ | ✅ | Plain scalar sheets without dependencies only |
| Cell scalar types | ✅ | ✅ | ✅ | |
| Rich cell text, spans, whitespace, fields, and hyperlinks | 🟡 | 🟡 | 🟡 | Fields and foreign inline markup not editable |
| Formula strings and references | ✅ | ✅ | ✅ | |
| OpenFormula expression syntax | ✅ | ✅ | N/A | |
| Formula evaluation and recalculation | 🟡 | 🟡 | N/A | Subset of functions; no dependency recalculation |
| Ordinary formula text functions (OpenFormula §6.20) | 🟡 | N/A | N/A | No alternate locale profiles |
| Byte-position formula text functions (OpenFormula §6.7) | 🟡 | N/A | N/A | UTF-8 octets; no code-page emulation |
| Formula value inspection and conversion (OpenFormula §6.13 subset) | 🟡 | N/A | N/A | |
| Formula reference geometry and sheet metadata (OpenFormula §6.13 subset) | 🟡 | N/A | N/A | |
| Formula lookup and reference construction (OpenFormula §6.14 subset) | 🟡 | N/A | N/A | |
| Formula date and time (OpenFormula §6.10) | 🟡 | N/A | N/A | Fixed Gregorian profile; no workbook date systems |
| Repeated rows/cells and merged/covered cells | ✅ | ✅ | ✅ | |
| Named ranges and named expressions | ✅ | ✅ | ✅ | |
| Cell and sheet styles | 🟡 | 🟡 | 🟡 | Partial property vocabulary; no layout |
| Conditional cell styles (`style:map`) | 🟡 | 🟡 | 🟡 | Conditions not evaluated |
| Sheet conditional formatting | 🟡 | 🟡 | 🟡 | Rules not evaluated |
| Sparklines | 🟡 | 🟡 | 🟡 | Not calculated or rendered |
| Content validation | 🟡 | 🟡 | 🟡 | Not enforced |
| Comments attached to cells | ✅ | ✅ | ✅ | |
| Database ranges, filters, sorts, and subtotals | ✅ | ✅ | ✅ | |
| DataPilot/pivot tables | ✅ | ✅ | ✅ | |
| Sheet/document protection | ✅ | ✅ | 🟡 | |
| Print ranges and page settings | 🟡 | 🟡 | 🟡 | No pagination |
| Calculation settings and iteration | ✅ | ✅ | ✅ | |
| Consolidations, label ranges, scenarios, and detective metadata | 🟡 | 🟡 | 🟡 | Metadata only; nothing applied or calculated |
| DDE sources and cached tables | 🟡 | ✅ | 🟡 | Never refreshed |
| External linked tables/ranges | 🟡 | 🟡 | 🟡 | Never refreshed |
| Cell-level external range source (`table:cell-range-source`) | 🟡 | ✅ | ✅ | Never refreshed |
| In-table shapes and images | 🟡 | 🟡 | 🟡 | No arbitrary SVG paths |
| Tracked spreadsheet changes | ✅ | ✅ | ✅ | |
| CSV export | ❌ | ❌ | ❌ | |

## Explicit gaps from the audited feature families

| Feature family | Status | Read | Write | Notes |
|---|---|---|---|---|
| Full public worksheet authoring | 🟡 | ✅ | ✅ | No repeated-run splitting or arbitrary grouping |
| Full OpenFormula semantics | 🟡 | 🟡 | ❌ | Incomplete function coverage; no recalculation |
| XLSX slicers and timelines | ❌ | ❌ | ❌ | |
| XLSX PivotTable/data-model extensions | ❌ | ❌ | ❌ | |
| XLSX external data and code services | ❌ | ❌ | ❌ | |
| XLSX threaded comments and rich values | ❌ | ❌ | ❌ | |
| Rendering, pagination, and recalculation | ❌ | ❌ | ❌ | |
| Security policy enforcement | 🟡 | ✅ | 🟡 | No signing or re-encryption of encrypted sources |
