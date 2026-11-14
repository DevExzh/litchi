# ODT/OTT Feature Matrix

Feature checklist for packaged ODT/OTT and flat `.fodt`. ✅ Supported · 🟡 Partial ·
❌ Unsupported · N/A Not applicable. External resources, formulas, scripts, macros and links are
never executed or fetched.

## Package, document lifecycle, and variants

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Open, validate, and save ODT bytes or paths | ✅ | ✅ | ✅ | |
| OTT template packages | ✅ | ✅ | ✅ | |
| Package MIME and template classification | ✅ | ✅ | ✅ | |
| ZIP package, `mimetype`, and manifest | ✅ | ✅ | ✅ | |
| Core XML parts | ✅ | ✅ | ✅ | |
| Immutable package transactions | ✅ | ✅ | ✅ | |
| Durable deterministic patches | 🟡 | ✅ | ✅ | Paragraph, run, hyperlink and resource operations only |
| Joined edits and three-way planning | ✅ | N/A | ✅ | |
| Bounded transaction history | ✅ | ✅ | ✅ | |
| Signed and encrypted transaction policy | ✅ | ✅ | 🟡 | Changed signed or encrypted packages are refused |
| Deterministic new-document metadata | ✅ | N/A | ✅ | |
| Unknown and auxiliary package entries | 🟡 | ✅ | ✅ | Preserved, not understood |
| Embedded media inventory | 🟡 | ✅ | ✅ | No media codec |
| Flat OpenDocument Text XML (`.fodt`) | ✅ | ✅ | ✅ | |
| Flat text templates | ❌ | ❌ | ❌ | |
| Package-only parts on FODT | N/A | N/A | N/A | |
| Other ODF families through the generic package wrapper | 🟡 | ✅ | 🟡 | Package validation only |

## Text structure and navigation

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Paragraphs and rich text spans | ✅ | ✅ | ✅ | |
| Headings and outline levels | ✅ | ✅ | ✅ | |
| Numbered paragraphs | ✅ | ✅ | ✅ | |
| Text blocks and ordering | ✅ | ✅ | ✅ | |
| Lists and list items | ✅ | ✅ | ✅ | |
| Outline styles and label alignment | ✅ | ✅ | ✅ | Labels not computed |
| Hyperlinks | ✅ | ✅ | ✅ | Targets never followed |
| Bookmarks | ✅ | ✅ | ✅ | |
| Reference marks | ✅ | ✅ | ✅ | |
| Ruby annotations | 🟡 | ✅ | 🟡 | No ruby layout |
| Frames and text boxes in text flow | 🟡 | ✅ | ✅ | Nested frame content not fully modeled |
| Header and footer text | ✅ | ✅ | ✅ | |
| Page breaks and inline structural markers | ✅ | ✅ | ✅ | |
| Text extraction as rendered text | 🟡 | ✅ | N/A | No pagination or line wrapping |

## Styles, page layout, and master pages

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Text and character styles | ✅ | ✅ | ✅ | |
| Paragraph styles | ✅ | ✅ | ✅ | |
| Table, row, column, and cell styles | ✅ | ✅ | ✅ | |
| List and numbering styles | ✅ | ✅ | ✅ | |
| Page layouts | ✅ | ✅ | ✅ | |
| Page-layout columns | ✅ | ✅ | ✅ | |
| Line-numbering configuration | ✅ | ✅ | ✅ | Line numbers not generated |
| Master pages | ✅ | ✅ | ✅ | |
| Explicit `text:page-sequence` | ✅ | ✅ | ✅ | |
| Headers, footers, and their properties | ✅ | ✅ | ✅ | |
| Font-face declarations | ✅ | ✅ | ✅ | No font loading |
| Data and number styles | 🟡 | ✅ | ✅ | Values not formatted |
| Drawing style resources | ✅ | ✅ | ✅ | |
| Style registry and inheritance | 🟡 | ✅ | 🟡 | No complete cascade resolution |
| Pagination, line layout, and visual rendering | ❌ | ❌ | ❌ | |

## Sections, fields, variables, and generated content

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Text sections | ✅ | ✅ | ✅ | |
| Section properties and backgrounds | ✅ | ✅ | ✅ | |
| Section protection policy | 🟡 | ✅ | ✅ | Not enforced |
| Inline field vocabulary | ✅ | ✅ | ✅ | |
| Dynamic field CRUD | ✅ | ✅ | ✅ | |
| Field cached values and instructions | 🟡 | ✅ | ✅ | Fields not evaluated |
| Database fields and sources | 🟡 | ✅ | ✅ | Inert metadata |
| DDE connections | 🟡 | ✅ | ✅ | Inert metadata |
| Field evaluation and refresh | ❌ | ❌ | ❌ | |
| Variable declarations and user fields | ✅ | ✅ | ✅ | |
| Content validation rules | 🟡 | ✅ | ✅ | Not enforced |
| Tables of contents and indexes | 🟡 | ✅ | ✅ | No generation or refresh |
| Index marks | ✅ | ✅ | ✅ | |
| Bibliography configuration and records | 🟡 | ✅ | ✅ | No entry generation |
| Concordance/auto-mark metadata | 🟡 | ✅ | ✅ | No index regeneration |

## Tables and tabular behavior

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Tables, rows, columns, and cells | ✅ | ✅ | ✅ | |
| Merged cells | ✅ | ✅ | ✅ | |
| Repeated rows and columns | 🟡 | ✅ | ✅ | |
| Cell values | ✅ | ✅ | ✅ | |
| Table, row, column, and cell properties | ✅ | ✅ | ✅ | |
| Table formulas and cached results | 🟡 | ✅ | 🟡 | Formulas not evaluated |
| Database ranges and external table sources | 🟡 | ✅ | ✅ | Never queried or refreshed |
| Automatic table layout and pagination | ❌ | ❌ | ❌ | |

## Notes, annotations, and tracked changes

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Footnotes and endnotes | 🟡 | ✅ | ✅ | |
| Note configuration | ✅ | ✅ | ✅ | |
| Footnote separators | ✅ | ✅ | ✅ | |
| Standard ODF annotations/comments | ✅ | ✅ | ✅ | |
| Rich annotation bodies | 🟡 | ✅ | ✅ | Not threaded |
| Threaded comments, replies, reactions, and people graphs | ❌ | ❌ | ❌ | |
| Tracked-change declarations | ✅ | ✅ | ✅ | |
| Tracked-change marking and XML CRUD | ✅ | ✅ | ✅ | |
| Accept/reject and revision merge engine | ❌ | ❌ | ❌ | |

## Forms, controls, and events

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODF `office:forms` containers | ✅ | ✅ | ✅ | |
| ODF XForms model declarations | 🟡 | ✅ | ✅ | Inert; no XForms processing |
| Text, textarea, button, checkbox, radio, list, and combo controls | ✅ | ✅ | ✅ | |
| Number, date, time, password/file, image-frame, and generic controls | ✅ | ✅ | ✅ | |
| Value-range, typed-value, selection, and grid controls | ✅ | ✅ | ✅ | |
| Form property values | ✅ | ✅ | ✅ | |
| Form events, bindings, and external data sources | 🟡 | ✅ | ✅ | Inert declarations |
| Form validation and UI behavior | ❌ | ❌ | ❌ | |
| OOXML content controls and ActiveX controls | ❌ | 🟡 | ❌ | Opaque payload preserved only |

## Drawings, images, charts, and embedded objects

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODF image frames | ✅ | ✅ | ✅ | PNG, JPEG and GIF authoring |
| Image replacement and removal | ✅ | ✅ | ✅ | |
| Linked and missing images | 🟡 | ✅ | ✅ | Never fetched or repaired |
| Client-side image maps | 🟡 | ✅ | N/A | |
| Text boxes and frame anchors | ✅ | ✅ | ✅ | |
| ODF geometric shapes and connectors | 🟡 | 🟡 | 🟡 | No complete shape CRUD |
| Drawing resources | ✅ | ✅ | ✅ | |
| Embedded ODF charts | ✅ | ✅ | ✅ | |
| Chart calculation and visual rendering | 🟡 | ✅ | ❌ | Not calculated or rendered |
| Embedded objects and subdocuments | 🟡 | ✅ | ✅ | |
| Embedded MathML | 🟡 | ✅ | 🟡 | Carried as embedded content |
| Embedded object activation or conversion | ❌ | ❌ | ❌ | |
| External drawing/object links | 🟡 | ✅ | ✅ | Inert metadata |

## Metadata, settings, and package graphs

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Dublin Core and ODF document metadata | ✅ | ✅ | ✅ | |
| User-defined metadata | ✅ | ✅ | ✅ | |
| Document statistics | 🟡 | ✅ | ✅ | Not recomputed |
| Template, auto-reload, and hyperlink-behavior metadata | 🟡 | ✅ | ✅ | Not applied |
| In-content RDFa and `text:meta` | 🟡 | ✅ | ✅ | Semantics inert |
| RDF metadata graphs | ✅ | ✅ | ✅ | |
| Settings tree | ✅ | ✅ | ✅ | |
| Settings-driven application behavior | ❌ | 🟡 | ❌ | |
| Document protection policy metadata | 🟡 | ✅ | ✅ | Not enforced |
| Protection and visibility settings | 🟡 | ✅ | ✅ | Not enforced |
| Package graph mutation and signature invalidation | ✅ | ✅ | ✅ | |

## Scripts, macros, and external behavior

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Document script declarations | 🟡 | ✅ | ✅ | Inert |
| Script package resources | 🟡 | ✅ | ✅ | Inert |
| Macro and event execution | ❌ | ❌ | ❌ | |
| Mail merge and database execution | ❌ | 🟡 | ❌ | |
| External link refresh | ❌ | 🟡 | ❌ | |
| VBA/OVBA projects and ActiveX | ❌ | 🟡 | ❌ | |

## Encryption and digital signatures

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| ODF per-entry password encryption | ✅ | ✅ | ✅ | |
| ODF encryption profiles | ✅ | ✅ | ✅ | |
| Encrypted package mutation | 🟡 | ✅ | 🟡 | Rewrites of existing encrypted packages may be refused |
| ODF document signatures | ✅ | ✅ | ✅ | |
| Macro-signature metadata | 🟡 | ✅ | 🟡 | |
| Certificate trust, revocation, and identity policy | ❌ | 🟡 | ❌ | |
| Flat-document encryption and signatures | N/A | N/A | N/A | |

## Explicit Microsoft-format gaps

| Feature family | Status | Read | Write | Notes |
|---|---|---|---|---|
| Compound File Binary storage and streams (`[MS-CFB]`) | ❌ | 🟡 | 🟡 | Opaque bytes only |
| OLE1/OLE2 embedded and linked objects (`[MS-OLEDS]`) | ❌ | 🟡 | 🟡 | Opaque bytes only |
| OLE property sets (`[MS-OLEPS]`) | ❌ | ❌ | ❌ | |
| Microsoft Office encryption and data spaces (`[MS-OFFCRYPTO]`) | ❌ | ❌ | ❌ | |
| Microsoft Office binary signatures (`[MS-OFFCRYPTO]`) | ❌ | ❌ | ❌ | |
| VBA project streams and modules (`[MS-OVBA]`, `[MS-OSHARED]`) | ❌ | 🟡 | ❌ | Opaque bytes only |
| OfficeArt shapes and records (`[MS-ODRAW]`) | ❌ | 🟡 | ❌ | Opaque bytes only |
| OOXML DrawingML and extensions (`[MS-ODRAWXML]`) | ❌ | 🟡 | ❌ | Opaque bytes only |
| Excel binary chart/graphics structures (`[MS-OGRAPH]`) | ❌ | 🟡 | ❌ | Opaque bytes only |
| EMF, EMF+, and WMF metafiles (`[MS-EMF]`, `[MS-EMFPLUS]`, `[MS-WMF]`) | ❌ | 🟡 | 🟡 | Opaque bytes only |
| OOXML-specific controls, modern comments, custom XML, and alternate content | ❌ | 🟡 | ❌ | Opaque bytes only |

## Remaining semantic non-goals

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Full ODF version/conformance validation | 🟡 | 🟡 | 🟡 | Not a schema validator |
| Office-style rendering fidelity | ❌ | ❌ | ❌ | |
| Recalculation of generated content | ❌ | ❌ | ❌ | |
