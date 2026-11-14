# Numbers (`.numbers`) Feature Matrix

Feature checklist for Apple Numbers packages. ✅ Supported · 🟡 Partial · ❌ Unsupported ·
N/A Not applicable. Writes edit existing packages unless a row says otherwise. Nothing is
fetched, executed, or recalculated.

## Package and semantic workbook

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Open and classify an existing Numbers package | 🟡 | ✅ | 🟡 | |
| Archive-free workbook, sheets, tables, and selectors | ✅ | ✅ | N/A | |
| Directory-backed semantic read | 🟡 | ✅ | N/A | Not on Windows |
| Exact no-op and untouched-member preservation | ✅ | ✅ | ✅ | |
| Fresh focused `.numbers` package creation | 🟡 | N/A | 🟡 | |
| Global compatibility table projection | 🟡 | ✅ | ❌ | |
| Per-table CSV projection | 🟡 | ✅ | ❌ | No CSV import |

## Sheets, tables, and presentation settings

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Sheet order | 🟡 | ✅ | 🟡 | |
| Existing table relocation between rooted sheets | 🟡 | ✅ | 🟡 | |
| Sheet and table names | 🟡 | ✅ | 🟡 | |
| Table dimensions and row/column sizes | 🟡 | ✅ | 🟡 | Sizes only; no topology change |
| Header/footer, frozen, and repeating rows/columns | 🟡 | ✅ | 🟡 | |
| Table appearance | 🟡 | ✅ | 🟡 | |
| Table title visibility and outline | 🟡 | ✅ | 🟡 | |
| Table lock state | 🟡 | ✅ | 🟡 | |
| Persisted sort configuration | 🟡 | ✅ | 🟡 | Does not reorder rows |
| Merged-cell geometry | 🟡 | ✅ | 🟡 | |
| Cell-grid and row/column topology CRUD | 🟡 | ✅ | 🟡 | |

## Cells, formulas, controls, and annotations

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Sparse scalar cells and bounded dense ranges | 🟡 | ✅ | 🟡 | |
| Formula source and cached values | 🟡 | 🟡 | 🟡 | 13 authorable functions; no evaluation or recalculation |
| Existing-cell Checkbox, Star Rating, Slider, and Stepper controls | 🟡 | 🟡 | 🟡 | |
| Existing-cell Pop-Up Menu control/display format | 🟡 | 🟡 | 🟡 | |
| Existing-cell Number display format | 🟡 | ✅ | 🟡 | |
| Existing-cell Percentage display format | 🟡 | ✅ | 🟡 | |
| Existing-cell Currency display format | 🟡 | ✅ | 🟡 | |
| Existing-cell Scientific display format | 🟡 | ✅ | 🟡 | |
| Existing-cell Fraction display format | 🟡 | ✅ | 🟡 | |
| Existing-cell Text display format | 🟡 | ✅ | 🟡 | |
| Existing-cell Date & Time display format | 🟡 | ✅ | 🟡 | |
| Existing-cell Duration display format | 🟡 | ✅ | 🟡 | |
| Other generic cell display formats and rich styles | 🟡 | ✅ | 🟡 | |
| Root cell comments | 🟡 | 🟡 | 🟡 | Text only; no authors, mentions or attachments |
| Direct replies to root cell comments | 🟡 | ✅ | 🟡 | |
| Charts, shapes, text boxes, and media assets | 🟡 | 🟡 | 🟡 | Chart arrangement flags only |
| Pivots and conditional highlights | 🟡 | ✅ | 🟡 | No pivot authoring |
| Filters, general categories and groups, and validation | ❌ | ❌ | ❌ | |
| Print/page setup and workbook export | ❌ | ❌ | ❌ | |
| Source metadata | 🟡 | ✅ | ❌ | |
| Existing-cell Custom display format | 🟡 | ✅ | 🟡 | |

## Limits, security, and publication

| Feature | Status | Read | Write | Notes |
|---|---|---|---|---|
| Bounded package and semantic resource handling | 🟡 | ✅ | 🟡 | |
| Encrypted or password-protected packages | ❌ | ❌ | ❌ | |
| Digital signatures | ❌ | ❌ | ❌ | |
| External resources, actions, macros, and database refresh | 🟡 | 🟡 | ❌ | Inert; never fetched or executed |
| Durable save and serialized/composable patches | 🟡 | N/A | 🟡 | No durable patch format |
