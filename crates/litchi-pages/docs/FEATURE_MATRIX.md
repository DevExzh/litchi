# Pages (`.pages`) Feature Matrix

Feature checklist for Apple Pages packages. ✅ Supported · 🟡 Partial · ❌ Unsupported ·
N/A Not applicable. Writes edit existing packages unless a row says otherwise.

## Package and semantic reader

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Pages package ingress and application detection | 🟡 | ✅ | 🟡 | |
| Exact package no-op and source preservation | 🟡 | ✅ | ✅ | No general package editor |
| Archive-free semantic `Document` | ✅ | ✅ | N/A | |
| Section selection and plain-text projection | ✅ | ✅ | N/A | |
| Package metadata sidecars | 🟡 | ✅ | ❌ | |
| Validation and bounded resource accounting | 🟡 | ✅ | N/A | |
| Unknown fields, opaque members, and preservation-only payloads | 🟡 | ✅ | 🟡 | Preserved, not understood |

## Existing-package text and document formatting

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Body and section text | 🟡 | ✅ | 🟡 | One plain-text splice per transaction |
| Section settings | 🟡 | ✅ | 🟡 | |
| Section name and pagination convenience transactions | 🟡 | ✅ | 🟡 | |
| Document and footnote formatter settings | 🟡 | ✅ | 🟡 | |
| Section backgrounds | 🟡 | ✅ | 🟡 | Solid fills only |
| Page layout | 🟡 | ✅ | 🟡 | |
| Headers and footers | 🟡 | ✅ | 🟡 | Text of existing slots only |
| Body footnotes | 🟡 | ✅ | 🟡 | |
| Body drawable stacking order | 🟡 | ✅ | 🟡 | |
| Endnote text lifecycle | ❌ | ❌ | ❌ | |

## Body tables

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Body-table catalog, selection, and rooted ownership discovery | 🟡 | ✅ | N/A | |
| Body-table hidden rows and columns | 🟡 | ✅ | 🟡 | |
| Table appearance | 🟡 | ✅ | 🟡 | |
| Table dimensions | 🟡 | ✅ | 🟡 | |
| Header, footer, freeze, and repeat settings | 🟡 | ✅ | 🟡 | |
| Table name | 🟡 | ✅ | 🟡 | |
| Persisted table sort configuration | 🟡 | ✅ | 🟡 | Does not sort cell data |
| Table title visibility and outline | 🟡 | ✅ | 🟡 | |
| Table interactive lock state | 🟡 | ✅ | 🟡 | |
| Body-table merged-cell geometry | 🟡 | ✅ | 🟡 | |
| Table cell values, formulas and rich formatting | 🟡 | ✅ | 🟡 | Formulas not evaluated |
| Body-table deletion | 🟡 | ✅ | 🟡 | |
| Table and row/column structural CRUD | 🟡 | ✅ | 🟡 | |

## Rich content and application features

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Rich-text runs, styles, lists, links, and annotations | 🟡 | 🟡 | 🟡 | No revisions or inline-object editing |
| Images, audio, movies, and media assets | 🟡 | ✅ | 🟡 | |
| Drawable geometry, grouping, text boxes, and shape lifecycle | 🟡 | ✅ | 🟡 | Groups cannot be authored |
| Charts and chart data | 🟡 | 🟡 | 🟡 | Arrangement flags only |
| Comments | 🟡 | ✅ | 🟡 | |
| Collaboration, revisions, and coauthoring | ❌ | ❌ | ❌ | |
| Rendering and export | ❌ | ❌ | ❌ | |
| Fresh package creation and structural document authoring | 🟡 | N/A | 🟡 | |

## Security, durability, and interoperability boundaries

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Password-protected or encrypted Pages packages | ❌ | ❌ | ❌ | |
| External resources, actions, macros, controls, and embedded code | 🟡 | 🟡 | 🟡 | Inert; never fetched or executed |
| Cryptographic signatures and trust | ❌ | ❌ | ❌ | |
| Durable patch serialization, composition, merge, and history | ❌ | N/A | ❌ | |
| Durable atomic filesystem publication | 🟡 | N/A | 🟡 | |
