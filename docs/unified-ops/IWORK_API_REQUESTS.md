# iWork public API requests from the unified operations API

Status: initial stub (record 0002, 2026-10-04). The owner has delayed iWork to tier 3; these requests are
recorded early so the iWork owners can plan. The unified-ops program **does not modify** `litchi-iwa*`,
`litchi-numbers*`, `litchi-pages` or `litchi-keynote` (`docs/UNIFIED_OPS_GOAL.md`, constraint 7). Until a
request lands, the gap stays visible in `capabilities()`.

## What already suffices (wiring only)

The facade already holds each `Package` (`crates/litchi/src/document/types.rs:51`,
`crates/litchi/src/presentation/types.rs:47`, `crates/litchi/src/sheet/workbook_types.rs:18`). The following unified
verbs need no new iWork API:

- **Save.** `Package::save` (atomic sibling-temp save) and `write_to`:
  - `crates/litchi-pages/src/package.rs:656,713`
  - `crates/litchi-keynote/src/package.rs:941,997`
  - `crates/litchi-numbers/src/package.rs:763,823`
- **Pages body and section text set/insert/delete.** `edit_body_text` and `edit_section_text`
  (`crates/litchi-pages/src/package/section_text.rs:227-312,650-682`).
- **Keynote slide title/body text, slide order, slide delete and skip.**
  - `crates/litchi-keynote/src/package/slide_text.rs:337-442,756-866`
  - `package/slide_order.rs:441`
  - `package/slide_delete.rs:635`
  - `package/edit.rs:99-156`
- **Numbers cell set and clear** for text, number, boolean, date, duration and formula.
  - `crates/litchi-numbers/src/package/table_cells.rs:233-277`
  - `table/cells.rs:54`
- **Numbers display formats, sheet move, table move.**
  - `package/sheet_order.rs:202`
  - `package/table_relocation.rs:434-567`

## Requests

Each request is described by intent. The signatures are suggestions; the owners choose the final shape.

**Cross-cutting**

1. **Create from scratch (all three formats).** For example `Package::create() -> Result<Package>`, or
   `Package::new_blank(template)`, producing a minimal native-valid package. Today only detached
   `Builder`s exist, and they do not produce a package (`crates/litchi-pages/src/section.rs:565`,
   `crates/litchi-keynote/src/slide.rs:281`, `crates/litchi-numbers/src/sheet.rs:255`). Without this,
   unified `create(Format::{Pages, Keynote, Numbers})` cannot be implemented.
2. **Uniform commit shape.** Ideally one root `edit()` transaction per format, or at least a common
   trait or shape across the many scoped `*Commit` types: `(package, patch with inverse(), optional
   diagnostics)`. This lets the unified batch executor coalesce work and roll it back. (Keynote already
   has `Package::edit()`, `crates/litchi-keynote/src/package.rs:914`.)
3. **Typed encrypted state.** A typed encrypted error or state, instead of mapping encrypted input to
   `InvalidFormat` (e.g. `crates/litchi-numbers/src/document.rs:1282`). The unified `File` needs this to
   surface `Locked`.

**Numbers**

4. **Sheet lifecycle:** add, remove and rename a sheet.
5. **Table lifecycle:** add and remove a table.
6. **Row and column insert/delete as package transactions.** The `RowDeletion` and `ColumnDeletion`
   types exist (`crates/litchi-numbers/src/table/topology.rs:69`), but no public package edit entry was
   found (*unverified*).
7. **Cell text-style edits:** bold, italic, font, colour.

**Keynote**

8. **Slide add and duplicate** from a master or layout.
9. **Shape and text-box CRUD:** add, remove, and change geometry.
10. **Text formatting** on slide text spans.

**Pages**

11. **Section add and remove.**
12. **Paragraph-level addressing:** set, insert, remove and move a paragraph by index. The current
    surface is span-based over a flat section or body text.
13. **Text formatting** on body text spans.
14. **Document metadata set** (title, author).
