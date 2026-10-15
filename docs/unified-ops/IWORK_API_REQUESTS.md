# iWork public API requests from the unified operations API

Status: initial stub (record 0002, 2026-10-04); superseded in part by owner
decisions D1 and D2 (2026-10-15, record 0248): the unified-ops program may now
edit `litchi-iwa*`, `litchi-numbers*`, `litchi-pages` and `litchi-keynote`, and
`File` holds the iWork packages directly (record 0296, section below). The
requests stay the iWork leaves' backlog; until a request lands, the gap stays
visible in `capabilities()`.

## What already suffices (wiring only)

The facade no longer holds an iWork `Package`: the wrappers that did (`crates/litchi/src/document/types.rs`,
`crates/litchi/src/presentation/types.rs`, `crates/litchi/src/sheet/workbook_types.rs`) were deleted in S9
(`docs/unified-ops/s9-deletion.md`). `litchi::iwork::Document` (`crates/litchi/src/iwork/model.rs:19`) holds a
decoded read-only snapshot, and `litchi::open` refuses iWork (`crates/litchi/src/lib.rs:173-176`). An ops adapter
would hold each leaf `Package` directly (corrected by record 0249). The following unified verbs need no new iWork
API:

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

## 2026-10-08 glossary coordination request (record 0016)

The owner-directed glossary is adopted as [ADR 0034](../adr/0034-unified-api-glossary-and-default-policies.md).
Its W1 remains with the iWork owners. Please expose the canonical `Package`
owner and owned `Workbook` (Numbers) / `Presentation` (Keynote) projections,
`SourceLimits` / `SemanticLimits` / `OpenOptions`, checked `Position` selectors,
and complete focused transaction families with required diagnostics. Fresh
builders need explicit identity input, deterministic fixed-identity output,
and typed refusal when identity is missing. Audit silent normalization before
claiming `ToleranceReport::Strict`. The unified layer will consume those public
APIs at M6; no iWork crate is modified by this adoption.

This contract supersedes the optional diagnostics and provisional create names
in requests 1–2 above: diagnostics are required and minimal creation is
`Package::new()`. Existing request evidence remains a historical inventory,
not a claim that those old facade paths still exist after M0a.

## 2026-10-13 typed keywords request (record 0253)

ADR 0034 §7 types `Metadata.keywords` as `litchi_core::metadata::Keywords`.
Three iWork construction sites assign an `Option<String>` to the field, so the
field stays `Option<String>` until they change. Please replace each assignment
with `.map(Keywords::from)` (one line each) when the field flips:

- `crates/litchi-numbers/src/package.rs:1575` (`metadata.keywords = diagnostic.keywords;`);
- `crates/litchi-pages/src/package.rs:891` (`keywords: self.keywords.clone(),`);
- `crates/litchi-keynote/src/package.rs:4767` (`metadata.keywords = diagnostic.keywords.map(PropertyScalar::into_string);`).

`Keywords` already exists in `litchi_core::metadata` with `From<String>` and
`From<&str>`, so the change compiles against today's `litchi-core` once the
field type flips. No other iWork code touches the fields record 0253 changed.

## 2026-10-15 iWork joins `File` (record 0296)

Owner decisions D1 and D2 (record 0248, 2026-10-15) let `litchi-ops` depend
on the iWork crates and let the unified-ops program edit them. Record 0296
makes `File::Pages`, `File::Keynote` and `File::Numbers` hold the leaf
`Package` directly, through the object-tree adapter `litchi_ops::format::iwork`.

**Consumed now.**

- Open: `litchi_iwa_detect::PreparedSource::from_shared_bytes`, then each
  leaf's hidden prepared-source ingress (`__from_prepared_source`; the
  `internal-iwork-source` feature). Record 0296 added the Numbers ingress
  (`crates/litchi-numbers/src/package.rs`, `Package::__from_prepared_source`).
- Reads: Pages `sections`/`section_text`/`metadata`; Keynote
  `slides`/`slide_text`/`metadata`; Numbers `sheets`/tables/`iter_cells`.
- Writes: Pages `edit_section_text` (paragraph text as a UTF-16 span of its
  section, and whole-section text); Keynote `edit_slide_text` (title and
  body), `edit` (slide skip), `edit_slide_order`, `edit_slide_deletion`;
  Numbers `edit_table_cells` (text, number, boolean, clear).
- Save: each leaf's `write_to`, staged by `litchi_core::atomic::save`.
- The Keynote deletion result and error are named through the leaf's
  canonical `litchi_keynote::slide::delete` module.

**Still open, as `capabilities()` gaps** (requests above unless noted):

- 1: `Package::new()` for all three formats; `create(Format::{Pages,
  Keynote, Numbers})` returns `unsupported_capability`.
- 3: typed encrypted state (no `Locked` for iWork yet).
- 4–6: Numbers sheet, table, row and column lifecycle.
- 7, 10, 13: text styling (Numbers cells, Keynote spans, Pages spans).
- 8, 9: Keynote slide add/duplicate and shape CRUD.
- 11, 12: Pages section add/remove and paragraph insert/remove/move.
- 14: Pages metadata set.
- **New 15 — public prepared-source ingress.** The unified layer uses the
  hidden `__from_prepared_source` seams. A documented, non-hidden
  `Package::from_prepared(PreparedSource, SemanticLimits)` per format (and a
  public way to keep the `Arc<[u8]>` source without a copy) would let the
  `internal-iwork-source` feature stay private to the facade.
- **New 16 — package directories.** The leaves refuse directory-backed
  sources ("support semantic projection only"), so `File` refuses an
  app-authored package directory and the facade's read-only
  `litchi::iwork::Document` stays its reader. Editing and saving a
  directory (or converting it to a single-file package) would let the
  facade reader go (plan step 28).
- **New 17 — Numbers formula writes through the unified registry.** The
  leaf authors formulas (`Edit::set_formula`); the unified `value` key does
  not map onto it yet (follow-up record 0299).

## 2026-10-15 Numbers parity (record 0297)

Record 0297 brings `litchi-numbers` to ADR 0034 W1 and wires its focused
transactions through the unified layer. Status of the Numbers requests:

- **1 (Numbers): done.** `litchi_numbers::Package::new()` returns a
  deterministic blank workbook (one sheet, one 10 x 5 table) and
  `litchi_numbers::Builder` names its sheet, table and identity; unified
  `create(Format::Numbers)` uses it. Formula authoring on the blank refuses
  (`table_cells::Error::UnsupportedDependency { kind: FormulaCache }`): the
  embedded package has no formula-owner dependency record for its table.
  **New 18:** emit a formula owner in the blank so formulas work on a
  created workbook.
- **4: rename done** (`edit_names().rename_sheet`), **sheet order done**
  (`edit_sheet_order`; it now also accepts a package without sheet
  previews). **Open:** sheet add and remove. The host editor in
  `litchi-iwa` has both (`numbers/editor/sheet_delete.rs`,
  `semantic/workbook.rs` `add_empty_sheet`), but ADR 0028 keeps
  `litchi-ops` off `litchi-iwa`, and porting them means owning the
  drawable, calculation-engine and view-state graphs in the leaf.
- **5: rename done** (`edit_names().rename_table`), **cross-sheet move
  done** (`edit_table_relocation`). **Open:** table add and remove, and a
  table order within one sheet (relocation only appends).
- **6: open.** Row and column insert and delete stay host-only
  (`numbers/editor/row_insert.rs`, `column_insert.rs`); the leaf has only
  the `RowDeletion`/`ColumnDeletion` value types.
- **7: open.** Cell text styles (bold, italic, font, colour) have no leaf
  transaction; the unified layer writes only the number and percentage
  display formats (`numberFormat`).
- **17: done.** `litchi_numbers::formula::Expression::parse` parses formula
  text with local A1 references, ranges, whole rows and columns, the
  operators and function calls; unified `set formula` uses it. References
  to other tables still need source-bound table handles, so they are
  refused.
- **New 19 — cell comment writes on Numbers-authored files.** The leaf
  comment transaction (`edit_table_cell_comments`) refuses every
  Numbers-authored corpus file (`InvalidSource` or
  `UnsupportedDependency`), and comments resolve per cell, so the unified
  `comment` key (read and write) stays a typed gap.

## 2026-10-15 Pages to parity (record 0299)

Record 0299 closes the Pages requests in the leaf (`litchi-pages`, owner
decision D2) and wires them into `File::Pages`:

- **1 (Pages)** — `Package::new()` returns the deterministic minimal document
  (a 5 KiB built-in template); `builder::Builder` with a required
  `builder::Identity` (`Identity::generate()` or fixed bytes) authors one
  with fresh identity sidecars and initial body text. Unified
  `create(Pages)` uses `Package::new()`.
- **2 (Pages)** — every Pages family is a focused module with the canonical
  `{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` (ADR 0034 W1); the
  unified write chain keeps each leaf patch.
- **12** — `section::paragraphs` (`Package::{section_paragraphs,
  edit_section_paragraphs, apply_section_paragraphs}`): set, insert, remove,
  relocate and swap by `Position`, published as one section-text splice
  when the change consumes no object anchor. Setting an anchored
  paragraph's text keeps its objects; removing or moving one is refused
  (`AnchoredParagraph`). The section-text transaction now also accepts the
  parent back-edges of objects anchored in the body (the source-built
  table files 0296 found refused).
- **14** — `metadata::{Edit, …}` (`Package::{edit_metadata,
  apply_metadata}`) writes title, author, keywords and description into
  `Metadata/Properties.plist`; the Pages app does not show these keys.
- **16 (Pages)** — `Package::open` reads an app-authored package directory
  by converting it to its single-file form; `File` opens it (no origin:
  `save_as` writes the single-file package).
- Also wired: body tables (read, cell text read, remove), headers and
  footers (read, text set), body footnotes (read, add at a paragraph end,
  text set, remove), section names.

**Still open for Pages:** 11 (section add/remove), 13 (text runs and
formatting), body-table cell writes and table creation (no Pages
cell-write transaction; the Numbers writer is not shared), directory
save-back to the bundle form, and `litchi::iwork::Document` removal
(blocked on the Keynote and Numbers directory and table gaps, records
0297–0298).
