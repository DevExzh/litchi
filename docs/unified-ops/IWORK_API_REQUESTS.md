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
   found (*unverified*). *Landed in 0309* for tables without formulas or merges (`table::topology`).
7. **Cell text-style edits:** bold, italic, font, colour. *Landed in 0309* (`cell::text_style`).

**Keynote**

8. **Slide add and duplicate** from a master or layout. *Landed in 0298:*
   `litchi_keynote::slide::creation` — `Package::edit_slide_creation()` with
   `Edit::{duplicate_slide(SlideSelector, Position), add_slide(LayoutSelector,
   Position)}` (several operations per edit, one commit),
   `Package::{slide_layouts, default_slide_layout, apply_slide_creation}`,
   exact-source `Patch::{is_noop, inverse}`; typed
   `Error::UnsupportedContent` for slides whose content lives in other
   components (tables, charts, comments). `Package::slide_layout_list()`
   names each slide's layout. Wired into `File::Keynote` (record 0298,
   unified part).
9. **Shape and text-box CRUD:** add, remove, and change geometry. *Landed in 0298:*
   `litchi_keynote::slide::drawable` — `Package::slide_drawable_list(slide)`
   (kind, `Geometry`, text, name, lock, z-order position, in
   `DrawableSelector` order) and `Package::edit_slide_drawables()` with
   `Edit::{set_geometry, remove, move_to, add_text_box, add_shape, add_image}`
   (several operations per edit, one commit; `Preset::{Rectangle,
   RoundedRectangle, Ellipse}`; PNG/JPEG/GIF images),
   `Package::apply_slide_drawables`, exact-source `Patch::{is_noop, inverse}`.
   Typed refusals: removing tables, charts, groups, movies/audio (see
   `slide::media::lifecycle`) and placeholders; removing a drawable referenced
   by builds or other objects; geometry of tables, charts, groups, lines and
   movies. Removed images keep their data items in the package.
   `Edit::set_text(slide, drawable, text)` writes a text box's or shape's
   plain text as a minimal-span storage rewrite (record 0298, unified part;
   `Unsupported::InlineObjects` for a change touching an inline object
   marker). Wired into `File::Keynote`.
10. **Text formatting** on slide text spans. *Partially landed in 0298:*
   read-only `Package::slide_drawable_text_runs(slide, drawable)` (character
   runs with bold, italic, underline, size, font and RGB colour where the
   character-style table names a decodable style); the write
   `drawable::Edit::set_text_style(slide, drawable, TextSpan, Style)` is a
   typed `Error::Unsupported { reason: Unsupported::TextStyling }` until a
   native character-style override fixture exists to port the host's layout
   from. *Landed in 0309* for text boxes and shapes (see below).

**Pages**

11. **Section add and remove.**
12. **Paragraph-level addressing:** set, insert, remove and move a paragraph by index. The current
    surface is span-based over a flat section or body text.
13. **Text formatting** on body text spans. *Landed in 0309* (`section::text_style`).
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
  Keynote, Numbers})` returns `unsupported_capability`. Keynote landed in
  the leaf in 0298: `Package::new()` (deterministic, fixed
  `Identity::BLANK`), `litchi_keynote::Builder` (`builder` module; title,
  subtitle, notes, slide size; `build()` without an `Identity` is
  `builder::Error::MissingIdentity`) and `litchi_keynote::Identity`
  (`generate()`, `from_bytes`). Unified `create(Format::Keynote)` uses
  `Package::new()` (record 0298); Pages landed in 0299; Numbers remains
  open.
- 3: typed encrypted state (no `Locked` for iWork yet).
- 4–6: Numbers sheet, table, row and column lifecycle.
- 7, 10, 13: text styling (Numbers cells, Keynote spans, Pages spans).
  Keynote character runs are read (`run`); their writes stay
  `Unsupported::TextStyling` in the leaf.
- 8, 9: closed for Keynote in 0298 (slide add/duplicate, shape, text box
  and picture CRUD, frames, z-order, drawable text), wired into
  `File::Keynote`.
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

## 2026-10-15 Keynote to parity (record 0298, unified part)

`File::Keynote` now reads slides with their `layout`, title/body
placeholders with frames, speaker notes, `paragraph`s and character `run`s,
and the drawables a slide owns (`shape` for text boxes and shapes,
`picture`, `connector`, `video`, `audio`, `table`/`tr`/`tc`, `chart`,
`group`), and writes through `slide::text`, `slide::notes`, `edit` (skip),
`slide::order`, `slide::delete`, `slide::creation` and `slide::drawable`.
Consecutive batch items coalesce into one slide-drawable commit, one
commit per changed text, or one slide-creation commit.

Leaf changes in this record: `slide::drawable::Edit::set_text`,
`Package::slide_layout_list`, and `slide::delete` removing the registry's
stale data-owner records of a slide whose preview was invalidated (the
first slide of `multi-slide.key` could not be deleted before; Phase G F8).

**Still open for Keynote:**

- 10 — character-style writes (`run` properties, `Unsupported::TextStyling`).
- Slide-table cell writes and table creation through the unified layer
  (the leaf's table families are wired for reads only).
- Removing tables, charts and groups (the drawable family refuses them);
  removing movies and audio (`slide::media::lifecycle`, not wired); adding
  movies and audio (`slide::movie`/`slide::audio` creation, not wired);
  frames of movies (`slide::movie::geometry`, not wired), lines, tables,
  charts and groups.
- Slide-number and object placeholders have no unified element; title and
  body placeholders are hidden, not removed (`slide::placeholder`
  visibility, not wired).
- Duplicating slides whose content lives in other components (tables,
  charts, comments; `slide::creation` `Error::UnsupportedContent`): 42 of
  the 56 corpus decks refuse the structural round trip for this reason.
- 16 (Keynote) — package directories: `File` still refuses them; the facade
  reader `litchi::iwork::Document` stays their reader.
- Slide-preview invalidation (slide text and notes edits) leaves the
  registry's data-owner record of the dropped thumbnail; deletion now
  tolerates it, the writers do not yet remove it.

## 2026-10-16 iWork structure (record 0301)

Record 0301 adds structural leaf transactions and wires them into
`File::Numbers` and `File::Pages`. Status:

- **4 (Numbers): done.** `sheet::lifecycle` (`Package::{edit_sheet_lifecycle,
  apply_sheet_lifecycle}`): add an empty sheet (the first sheet's settings
  without its private objects, a sidebar node) and remove a sheet with its
  tables. A sheet that holds charts, media or shapes is refused
  (`SheetNotEmpty`).
- **5 (Numbers): done**, with one gap. `table::lifecycle`: add an empty
  table (the workbook's first table's structure, emptied storage, a fresh
  UUID, a CalculationEngine owner with the native range contexts, a
  sidebar node) and remove a table with its private graph and formula
  owners (the TSCE phase ported from the Pages body-table deletion). Open:
  a workbook without any table (`NoTemplate`), and a removal on a formula
  graph the ported phase does not rewrite (a range tile with several
  from-to ranges) is refused, typed.
- **In-sheet table order: done.** `table::order` (`Package::edit_table_order`).
- **Multi-table cell writes** (found while testing): the table-cell writer
  refused every multi-table workbook (another table's undeclared native
  owner back-reference, and inert foreign owners without geometry). Both
  are accepted now; a workbook with formulas in another table still refuses
  a write (`UnsupportedDependency { FormulaCache }`): the cache planner
  needs every formula-bearing table's geometry.
- **6: open.** Row and column insert/delete: the host closure is 9–12k
  lines (formula dependency shifting, tiles, UID maps, stroke layers,
  merges, headers); not ported.
- **7, 19, cross-table formulas: open** (unchanged).
- **18: open, narrowed.** The blank's table owner lacks the native
  total/body range contexts (the cache planner refuses empty ones); with
  them added, formula authoring on the blank then fails one reference short
  in the formula-metadata planner's exact execution budget
  (`LimitExceeded { References, observed 9, maximum 8 }`). The asset was
  not changed.
- **16 (Pages): done.** `Package::save_directory` writes the bundle form
  (`Index.zip` with the `Index/` records unchanged, every other member a
  file) through a sibling staging directory; reopening gives the same
  single-file bytes. A `File::Pages` opened from a directory keeps it as
  its origin, so `save()` writes the directory back.
- **Pages anchored paragraphs: partly done.** `Package::section_paragraph_tables`
  maps each paragraph to the body tables it anchors; unified `remove` of a
  paragraph whose every anchor is a body table removes the tables first.
  Paragraphs anchoring charts, images or footnotes, and moving any anchored
  paragraph, stay refused.
- **11, 13, Pages table cells/creation: open.** Section add/remove (the host
  clones the section's page-master graph and inserts a U+0004 break through
  its text editor), text runs, and a Pages cell writer were not ported.

## 2026-10-16 package directories and the facade reader (record 0306)

- **16 (Keynote, Numbers): done.** `Package::open[_with_options]` accepts an
  app-authored package directory and converts it to its single-file form;
  `Package::save_directory` writes the bundle form back
  (`SaveError::{Bundling, Bundle}`). The capture, conversion and bundle
  publication are shared by the three leaves in
  `litchi_iwa_archive::bundle`; the Pages directory module now delegates to
  it. A bundle whose components sit loose in an `Index/` folder converts too
  (and saves back in the `Index.zip` layout).
- `File` opens, edits and saves package directories of all three formats in
  place; `litchi::iwork::Document` is deleted (plan step 28).
- Still open: the Numbers object tree refuses a source whose table-merge
  storage is not native (the synthetic compatibility fixture: text reads,
  `query` is `invalid_input`); a merge-reader tolerance request for the
  Numbers owner.

## 2026-10-16 iWork element gaps (record 0309)

Record 0309 adds the element-level leaf families the unified layer was
missing and wires them into `File`. Every edit keeps untouched bytes exact
and returns an exact-source patch with `inverse()`; what is not ported is a
typed refusal.

- **6 (Numbers): done, with typed gaps.** `table::topology`
  (`Package::{edit_table_topology, apply_table_topology}`;
  `Edit::{insert_row, insert_column, remove_row, remove_column}` with
  header/body/footer-relative `RowInsertion`/`ColumnInsertion`/`RowDeletion`/
  `ColumnDeletion`). It splices the tiles (pre-BNC offsets too), the row and
  column info buffers, header counts, the layout cache, stroke layers and
  the table-data lists (removed cells release their keys). Refused, typed
  (`Error::Unsupported(..)`): tables holding formulas or read by a formula
  or chart elsewhere (no formula-coordinate shift is ported), merged cells,
  filters, categories, pivots, column edits under sort rules, removing a
  hidden, bordered or commented axis, and non-native storage layouts. Corpus:
  28 tables insert and remove a row and a column round trip; unified
  `add`/`remove` of `row[N]`/`col[N]` under a table.
- **7 (Numbers): done.** `cell::text_style`
  (`Package::{edit_table_cell_text_style, apply_table_cell_text_style,
  table_cell_text_style}`): bold, italic, underline, strikethrough, size,
  font and colour as a paragraph-style variation of the cell's style; an
  empty cell is refused (`EmptyCell`). 33 corpus cells styled. Unified
  `font.*` on `cell` (set-only).
- **10 (Keynote): done for text boxes and shapes.**
  `slide::drawable::Edit::set_text_style` writes character-style variations
  into the storage's stylesheet (registry UUIDs and external references
  kept). Unified `run` properties `bold`, `italic`, `underline`, `size`,
  `color` and `font.latin`. **Open:** placeholder and notes runs (their
  text keeps the layout's styles; typed refusal).
- **13 (Pages): done at paragraph level.** `section::text_style`
  (`Package::{edit_section_text_style, apply_section_text_style,
  section_text_runs}`). Unified `font.*` on `p` (set-only). **Open:** run
  elements in the Pages tree; a whole-section text `set` after styling is
  refused (`DependentContent`) because it would drop the style references.
- **18 (Numbers): done.** The blank workbook asset now carries its table
  owner's extents and the formula-metadata preflight counts the owners, so
  formulas work on a created workbook.
- **Table cell writes (Pages, Keynote): done.** `table::values` for Pages
  (`Package::{edit_body_table_values, apply_body_table_values}`; 13 corpus
  tables written, the 4 holding formulas refused) and
  `slide::drawable::Edit::set_table_cell` for Keynote
  (`Unsupported::{TableFormulas, TableStorage}`), on the shared
  `litchi_numbers_wire::table_write` writer. Unified `tc` `text`/`value`.
- **Keynote media removal and placeholder visibility: wired.**
  `remove` of a `video`/`audio` goes through `slide::media::lifecycle`;
  `remove` of a title or body placeholder hides it and `hidden` = false
  shows it again (`slide::placeholder` visibility).

**Still open after 0309:** Numbers formula-coordinate shifting for
topology (tables with formulas), merges and filters under topology; Numbers
comments (19) and cross-table formulas; the Numbers cell writer's
pre-existing refusals on several corpus files; Pages section add/remove
(11), run elements, Markdown body tables and table creation; Keynote
placeholder and notes run styles, table creation, movie and audio
creation, duplicating slides with other-component content, image data
removal and stale thumbnail records; a readable placeholder `hidden`.

## 2026-10-16 lazy tables and the next element gaps (record 0312)

- **11 (Pages): done.** `section::lifecycle`
  (`Package::{edit_section_lifecycle, apply_section_lifecycle}`;
  `Edit::{add(after, Option<&str>), remove(section)}`, one operation per
  edit). Add inserts a U+0004 break at the end of a section and a
  section-table entry for an independent copy of that section's private
  graph (page templates, headers and footers, the guide map and its first
  guide storage), as the native editor does; remove drops a non-initial
  boundary and its private graph, and the removed section's text joins the
  preceding one. On `basic.pages` and `body-sections-unicode.pages` the
  rewritten `Index/Document.iwa` is identical to the retired host editor's
  (`append_pages_section`, `remove_pages_section`); only the registry
  differs (fresh UUIDs; the host also lowers the identifier watermark).
  Typed refusals: `InitialSection`, `EmptySection` (a break must follow
  content), `UnsupportedContent` (a header or footer owning records of its
  own, data references), `UnsupportedDependency`. 31 of 31 corpus
  documents add and remove a section. Unified `add` of `section` under `/`
  and `remove` of `/section[N]`.
- **10 (Keynote placeholders): done.** `slide::drawable::Edit::set_text_style`
  accepts title and body placeholders (78 corpus placeholders styled and
  inverted exactly). **Open:** speaker-notes runs (the tree lists no runs
  under notes; the notes storage needs a run reader and a style writer).
- **Keynote placeholder `hidden`: readable** through
  `Package::slide_placeholder_visibility` (no leaf change).
- **19 (Numbers comments): open, diagnosed.** Probing
  `set_table_cell_comment` on every corpus workbook (record 0312 §5): 31
  files fail because the creation census is package-wide and refuses a
  workbook holding more than one comment list (native Numbers gives every
  table its own) — selecting the target table's list moves that refusal to
  the next check, which requires the comment list, the table model and the
  tile in one component (native files keep tiles in their own
  components, so creation needs cross-component registry references); 6
  more fail the strict header-metadata census of other lists
  (`validate_comment_message_metadata`). Not changed in this record.
- **Still open:** Keynote and Pages table creation; Numbers topology on
  tables with formulas (the formula-coordinate shift is not ported).

## 2026-10-16 batch equivalence and the open element gaps (record 0314)

- **New leaf APIs (record 0314).** `litchi_iwa_text_wire::{StorageBoundaries,
  storage_boundaries_with_limits}` read where a text storage's attribute
  tables place their entries; `StorageBoundaries::admits_combined_splice`
  answers whether a sequence of splices inside a region publishes what one
  combined splice publishes. Keynote exposes it per text:
  `slide_text::Edit::admits_combined_splice`,
  `slide_notes::Edit::admits_combined_splice`,
  `slide_drawable::Edit::text_admits_combined_splice`.
- **Pages paragraph edits: fixed data loss.**
  `section::paragraphs::Edit::commit` published several staged sets as one
  splice from the first text to the last, which dropped the paragraph-style
  entries of the paragraphs between the changed ones. It now publishes what
  the operations publish one at a time: text sets share one splice only
  where the boundaries prove it equal; structural operations and other
  runs commit one transaction each.
- **19 (Numbers comments): still open, diagnosed further.** The 31 refusals
  are the package-wide census at `creation_graph` (more than one comment
  list); the 6 others are the strict field-info rule of
  `validate_comment_message_metadata` in mutation mode, applied to every
  list in the workbook. Native packages keep each `TST.TableDataList` in
  its own `Index/Tables/DataList-<id>.iwa`, the table model in
  `Index/CalculationEngine.iwa`, tiles in `Index/Tables/Tile*.iwa`, and the
  comment storages in `Index/Tables/DataList.iwa`
  (`comment-reader-native-created.numbers`: list 905275 in its own member,
  storage 905277 beside list 904482). Creation in that layout needs a new
  data-list member with its registry entry and a cross-member reference
  from the model; the strict same-member writer cannot express it.
  Follow-up: a per-table census, a two-member writer (list member + tile
  member) for tables that already own a list, then the new-member path.
- **Numbers topology on tables with formulas: still open.** The retired
  host editor's shift (`litchi-iwa/src/numbers/editor/formula_dependency_shift{,/ast,/wire}.rs`,
  about 4,000 lines over Prost archives: dependency records, range
  precedents, formula ASTs, incoming references from other owners) is the
  reference; `litchi-numbers` works on exact wire replacements without
  Prost, so the port is a rewrite. `Unsupported::{Formulas,
  ReferencedByOtherOwners}` stay.
- **Keynote and Pages table creation: not started.**

## 2026-10-16 formula-coordinate shift and cell comments (record 0318)

- **6 (Numbers topology on tables with formulas): done, with typed gaps.**
  `table::topology` now shifts every formula coordinate that addresses the
  edited table (`package/table_topology/shift.rs`, a wire-level port of the
  retired host editor's `formula_dependency_shift`): the table's own
  formula ASTs (local references, colon tracts, thunks, footer ranges),
  dependency records and tiles, range dependencies and tiles, UUID
  coordinate sets, merge proxies, table extents; and, beyond the host
  editor, other tables' and sheets' formulas that read the table, packed
  edge arrays, and whole-axis sentinels. New typed reasons
  `Unsupported::{DeletedReference, SharedFormulas}`; `Formulas` and
  `ReferencedByOtherOwners` now name only what the shift refuses
  (volatile/spanning/error/spill state, charts and non-table readers, UUID
  references). Differential test against the host editor:
  `litchi-iwa/tests/numbers_topology_formula_differential.rs` (oracle:
  `NumbersEditor::__edit_table_topology_unverified`). **Open:** `#REF!`
  rewriting on deletion (refused), splitting a shared formula entry whose
  cells diverge (refused), the cell writer's clear refusals on some removed
  axes.
- **19 (Numbers comments): done.** Native multi-member packages: per-table
  census, FieldInfo-free comment lists, effective component locators,
  creation over the list, tile and author-storage members with the
  storage's cross-member author edge registered, clear of a UUID-registered
  storage. New `Package::table_cell_comments(sheet, table)`. Unified: cell
  `comment` read and set (`null` clears), `add` of `comment` under a cell,
  `remove` of `<cell>/comment`. **Open:** a new comment list for a table
  without one goes beside the table model rather than into a member of its
  own; a comment on an empty cell where the tile has no row (`NeedSparse`);
  packages with several authors.
- **Keynote and Pages table creation: not started** (plan in record 0318
  §5).
- **Keynote speaker-notes run styles: not started.**

## 2026-10-17 table creation, deleted references and the remaining gaps (record 0329)

- **Keynote and Pages table creation: done, with a typed gap.**
  `litchi_numbers_wire::table_create` holds the payload half of
  `litchi-numbers`' table creation (record 0301) for the Keynote and Pages
  owners. Keynote: `slide::drawable::Edit::add_table(slide, NewTable)`
  (`NewTable::{new(rows, columns, Geometry), with_name, with_cell}`) clones
  the deck's first slide table (drawable, model, emptied storage, its own
  formula owner, registry UUIDs and cross-member references), appends it to
  the slide's drawables and z-order, writes seeded cells with the shared
  cell writer and reads the table back through `slide_table_cells`. Pages:
  `Package::insert_body_table(anchor, table::creation::NewTable)` clones the
  first body table and its drawable attachment, anchors it with a `U+FFFC`
  and an attachment-table entry at a UTF-16 body position, and writes
  seeded cells through `table::values`; `apply_body_table_creation` replays
  a patch. The Pages structural toolkit's registry transition gained
  external-reference additions. Unified `add` of `table`/`tbl` (Keynote,
  under a slide: `rows`, `cols`, `name`, `data`, frame) and of `tbl` (Pages,
  under `/body`: an empty paragraph at the placement, then the table).
  **Open:** a deck or document without a table to copy refuses
  (`Unsupported::TableTemplate`, `table::creation::Error::NoTemplate`): a
  built-in template (the theme's table style preset with a storage
  skeleton) is the follow-up.
- **6 (Numbers): `#REF!` and shared-formula splits done.** Removing a row or
  column a surviving formula reads writes `REFERENCE_ERROR_NODE` (`#REF!`)
  for a reference to it, or for a range wholly on it, caches an error
  result (`BncCell::set_formula_cached_error`), shrinks a range that loses
  it (mixed absolute/relative tract endpoints included) and drops the
  precedent edges and range dependencies into it. A formula-list entry
  whose cells diverge is split, one entry per formula
  (`BncCell::set_formula_identifier`). Still refused: several ranges of one
  tract with one removed, merge-range endpoints (`DeletedReference`), a
  split in a segmented formula list (`SharedFormulas`).
- **Topology clears: done.** Comments on a removed axis go through the
  comment transaction; where the cell writer cannot clear the removed
  cells (`HeaderNameIndex`, `FormulaCache`, `InvalidSource`, the created
  workbook's reference budget), the structural pass drops them with their
  strings, formulas and dependency records (an emptied dependency tile
  goes too), refusing (`Formulas`) a range over them whose cache would go
  stale. Every removal the retired host editor publishes over the
  differential corpus, the leaf now publishes.
- **19 (Numbers comments): two of three edge cases done.** A comment on an
  empty cell in a row its tile does not store materialises the row (31
  corpus tables); with several commenting authors a new comment is the
  generated author's (reused, or appended to the author storage).
  **Open:** a new comment list in a member of its own needs a component
  registration; `litchi_iwa_protos::package_metadata_codec` edits existing
  components only (follow-up: a `ComponentAddition` in its batches).
- **10 (Keynote notes runs): done.** `slide::drawable::Edit::set_notes_text_style`
  and `Package::slide_notes_text_runs`; unified notes runs.
