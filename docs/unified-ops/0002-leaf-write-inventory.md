# 0002 — Leaf write-capability inventory (all 21 claimed formats)

Status: record (survey; no code changed). Date: 2026-10-04.
Driven by `docs/UNIFIED_OPS_GOAL.md` (OWNER OVERRIDE: every claimed format needs read **and** write
through the unified API) and the M2b rule that each write stream starts from this inventory.

Priority (owner, 2026-10-04): OOXML (docx/xlsx/pptx/xlsb) and legacy OLE (doc/xls/ppt) are surveyed in
depth; RTF, ODF and iWork are summarized (entry points, whether edit/commit/save exist, gap class).

## Method and conventions

- Every claim was checked against source in this worktree by reading code (`grep`/`read`), not
  against `FEATURE_MATRIX.md`. Nothing was built or executed. Where a conclusion is from reading the
  code path rather than a test, it says *code-path inference*; where not confirmed, *unverified*.
- Citations are `path:line` relative to the worktree root. Per-section path prefixes are declared at
  the top of each section (e.g. `X/` = `crates/litchi-xlsx/src/`). `F/` always means the facade
  `crates/litchi/src/`.
- Gap classes (for a unified `set/add/remove/move/swap/save/create` handler using public APIs only):
  - **(a)** facade wiring only — the leaf API exists;
  - **(b)** small leaf API addition — an adapter, wrapper, or missing setter over an existing engine;
  - **(c)** new leaf engine needed.
- "Existing" = editing an opened file; "Create" = from-scratch builder output.

## Cross-cutting findings

1. **The facade is read-only for every format.** No `edit`/`commit`/`save`/`set_*` exists in
   `crates/litchi/src` (grep), consistent with design §16. Facade holders today:
   `DocumentImpl::{Doc(doc::Document, Metadata), Docx(Box<docx::Package>,..), DocxSource(docx::source_backed::Package,..), Pages(litchi_pages::Package), Rtf(litchi_rtf::RtfDocument<'static>), Odt(litchi_odt::Document)}`
   (`F/document/types.rs:32-57`, plus `OdtSource` at `:59`);
   `PresentationImpl::{Ppt(ppt::Presentation), Pptx(Box<pptx::Package>), PptxSource(SourceBackedPresentation), Keynote(litchi_keynote::Package), Odp(..), OdpSource(..)}`
   (`F/presentation/types.rs:35-58`);
   `WorkbookImpl::{Numbers(litchi_numbers::Package), Xlsx(adapters::Workbook), Xlsb(xlsb::Workbook), XlsbSource(..), XlsFile(..), XlsMem(..), Ods(RefCell<Spreadsheet>), OdsSource(..)}`
   (`F/sheet/workbook_types.rs:16-45`).
2. **Editable owners are dropped or absent on several primary open paths**, so M0a must change what
   the facade retains before any write handler works:
   - doc: the `Package` and its bytes are dropped; only the eager read `doc::Document` is kept
     (`F/document/doc.rs:960-975`, `F/document/types.rs:35`).
   - ppt: only `ppt::Presentation` (read model) is kept (`F/presentation/prs.rs:785-796,890-909`).
   - rtf: the raw `RtfDocument` is kept, not the editable `litchi_rtf::Document` snapshot (`F/document/types.rs:54`).
   - xls: `xls::Workbook<R>` / `SourceBackedWorkbook` are read-only; the edit families open bytes.
   - xlsx/xlsb/pptx/docx path-opens hold **source-backed** owners whose write paths are stream-only or
     absent (xlsx `SourceBackedWorkbook` has no `edit()`; xlsb `SourceBackedWorkbook` has no edit/save).
3. **No ADR 0006 `Locked<T>` / `Signed<T>` consumption types exist anywhere** in the workspace
   (repo-wide grep for `Locked<`, `Signed<`, `Encrypted<`, `Protected<`, `Sensitive<` finds only
   unrelated numeric structs). Security state is modelled per crate as error variants, bools, or
   policy enums. The unified `File` enum's `Locked`/`Signed` states are therefore a new facade-level
   construct (class (b) per leaf for a uniform "consume" API).
4. **Silent signature stripping on changed commits** (contradicts ADR 0006 "never silently strip"):
   - pptx `opened::Transaction::commit` calls `working.unsign()` (`crates/litchi-pptx/src/opened/transaction.rs:1255-1270`);
   - xlsb cell-values / workbook-structure / several families call `unsign()` (e.g. `crates/litchi-xlsb/src/cell_values/workbook.rs:167`);
   - odf-formula commit → `rebuild_package` → `copy_source_files_from_except` skips signature members (`crates/litchi-odf-common/src/core/writer.rs:1618-1622`);
   - ods legacy `MutableSpreadsheet` mutators and odp in-place annotation/RDF mutators (*code-path inference*).
   Formats that **refuse** instead: docx (until `unsign()`), xlsx (`Error::Signed`), doc/xls/ppt
   (incremental editors refuse signed/encrypted CFB), odt/odm/odc/odi (refuse). The unified layer
   must expose a single explicit policy; the stripping leaves are (b) fixes.
5. **Commit shape varies.** Only docx main-document, doc `body_text::source`, xls `cell_values`,
   ods `document::Edit`, odf-formula and rtf return a diagnostics member; xlsx/pptx/ppt/xlsb/odt do
   not. Several families publish by `&mut self` (xlsb `apply_*`, odt `MutableDocument`, ods legacy
   mutators, odf-formula `set_root`) rather than returning a new snapshot.
6. **Atomic path save is not uniform.** Atomic sibling-temp save exists on: docx `Package::save`,
   xlsx `Workbook::save`, pptx `Package::save`, xlsb `Package::save`, iWork `Package::save`, ppt/xls/doc
   `Writer::save` (create path only), and the narrow source-backed commits (doc, xls, ppt text_edit).
   It does **not** exist on: doc `body_text::Snapshot`, xls `cell_values::Snapshot`, ppt `slide_order`
   (bytes only); docx/pptx/xlsx source-backed publications (stream only); rtf (`write_to` only, by
   design); all ODF packaged saves (`fs::write` or bytes only). A shared facade atomic writer over
   bytes/stream covers all of these (class (a)).
7. **Three formats' create builders silently drop metadata**: ppt `Writer::set_property`
   (`crates/litchi-ppt/src/writer/core/model/semantic.rs:1764`, never serialized —
   `writer/core/package.rs:129-160`), doc `Writer::set_property`
   (`crates/litchi-doc/src/writer/core/model/semantic.rs:1046`, no SummaryInformation stream written),
   xls `Writer` writes no property streams (`crates/litchi-xls/src/writer/core/package.rs:93-129`).
   Workaround: chain the property-set transaction after the write (b).

## Summary matrix

Legend: ✅ typed public API; 🟡 partial/bounded; ❌ none. "txn" = ADR 0003 snapshot/edit/commit.

| Format | Read (facade today) | Edit txn (existing file) | Commit shape | Save | Create | Text set/add/remove/move | Props / formatting | Structure | Gap class |
|---|---|---|---|---|---|---|---|---|---|
| docx | ✅ source-backed + eager | ✅ `edit_document()` + ~20 scoped families; legacy `document_mut()` | (snapshot, reversible Patch, Diagnostics) | ✅ atomic (owning); stream-only (source-backed) | ✅ `Package::new()` | ✅ set / 🟡 add plain / 🟡 remove (not in `Edit`) / ❌ move | 🟡 core/custom props (no patch); ❌ formatting on existing runs | 🟡 sections (source-backed only); ❌ styles | a + b; c for formatting/table structure/styles |
| xlsx | ✅ source-backed + eager | ✅ `Workbook::edit()` (sheets, cells, layout, merges, hyperlinks) + lazy per-family editors | (Workbook, reversible Patch); no Diagnostics | ✅ atomic | ✅ `Workbook::create()` | ✅ cell set/clear/remove; 🟡 move (≤65,536 cells) | 🟡 apply existing Style only; custom props ✅, core via raw | ✅ sheet add/remove/move/rename/hide; ❌ row/col shift | a + b; c for row/col shift, pictures/charts/tables |
| xlsb | ✅ source-backed + eager | 🟡 ~15 per-family txns (cells, workbook-structure, theme, …) | (Snapshot, Patch) or `&mut` apply; no Diagnostics | ✅ atomic (`Package`) | ✅ `WorkbookWriter` / `Package::create()` | ✅ cells (same storage family) | 🟡 existing XF index only; ❌ doc props | 🟡 rename only | a + b; c for sheet lifecycle, hyperlinks/comments/tables on opened files |
| pptx | ✅ source-backed + eager | ✅ `opened::Transaction` (slides, shape text, add/remove shapes, comments, notes, masters) | (Snapshot, reversible joinable Patch); no Diagnostics | ✅ atomic (eager); stream (source-backed) | ✅ `Package::new()` + `MutablePresentation` | 🟡 whole-shape text replace; add/remove shapes | ❌ formatting on existing shapes; custom props ✅, core via raw | ✅ move/reorder; 🟡 remove/duplicate; ❌ blank-slide add | a + b; c for run/paragraph formatting |
| doc | 🟡 eager read model only (package dropped) | 🟡 `body_text::Edit` + ~12 scoped owners | (snapshot, reversible Patch); Diagnostics only on source-backed splice | 🟡 bytes (`finish`); atomic only on narrow source commit | ✅ `Writer` (metadata dropped) | 🟡 replace in place; ❌ structural add/remove/move | 🟡 bold/italic/underline; ✅ property sets txn | ❌ | a + b; c for structural body editing |
| xls | 🟡 eager + lazy read only | 🟡 ~10 families (`cell_values`, `sheet_visibility`, `comments`, `property_set`, …) | (Snapshot, reversible Patch, Diagnostics) for `cell_values` | 🟡 bytes; atomic on source-backed commits | ✅ `Writer` (no property streams) | 🟡 cell set (family-preserving), insert/remove; rows/cols insert/delete | 🟡 XF index / raw XF bytes; ✅ property sets txn | 🟡 rename, hide, reorder-all (bytes); ❌ add/remove sheet | a + b; c for sheet add/remove, hyperlinks/merges/names/pictures |
| ppt | 🟡 eager read model only | 🟡 `slide_order::Transaction` root + ~20 single-feature owners | (Snapshot, reversible Patch); no Diagnostics | 🟡 bytes (append-only incremental save) | ✅ `Writer` (metadata dropped) | 🟡 whole-shape text atom replace | ❌ formatting; ✅ property sets txn | ✅ slide move/remove/transfer, hidden, transition | a + b; c for rich text, inserting shapes/pictures/tables/slides |
| rtf | ✅ eager (raw model) | ✅ `Document::edit()` | (snapshot, reversible Patch, Diagnostics) + History | 🟡 `write_to` / `to_bytes` (byte-exact) | ❌ semantic builder (raw `Writer` only) | ✅ set / 🟡 add / ✅ remove/move (plain sources) | 🟡 14 run toggles, alignment/layout; ❌ metadata | ❌ sections/tables | a + b; c for create, structure |
| odt | ✅ eager + lazy | ✅ `transaction::Edit` (+ lossy `MutableDocument`) | (snapshot, reversible Patch, results); no Diagnostics | ❌ atomic (fs::write) | ✅ `Builder` | ✅ plain paragraph set/add/remove; 🟡 move (plain bodies) | 🟡 metadata via `MutableDocument`; ❌ formatting publish | 🟡 sections/master pages (MutableDocument only) | a + b; c for rich inline/table/block edits |
| ods | 🟡 names/count/text only | ✅ `document::Edit` (cells, rows, sheets, styles, drawings, charts) | (snapshot, reversible joinable Patch, Diagnostics) | ❌ atomic (fs::write / bytes) | 🟡 `Builder` (zero sheets) | ✅ cell set/clear, formula, rich text | 🟡 automatic styles; metadata outside `Edit` | 🟡 append/remove/move sheet (refused on typical producer files) | a + b; c for structural edits on producer files |
| odp | ✅ eager + lazy (read only) | 🟡 `edit::Snapshot::transaction()` (+ chart, flat, tail-append, in-place `&mut` paths) | (snapshot, reversible Patch); no Diagnostics | ❌ atomic (fs::write / bytes / stream) | ✅ `Builder` | 🟡 named text boxes ✅; title/body only on slides added in the same txn | 🟡 style-name refs only; ❌ metadata on existing decks | 🟡 slide add/remove/move (restricted); masters/layouts ✅ | a + b; c for retained-slide rewrite |
| odg | ❌ not opened by facade | ✅ `Drawing::edit()` | (snapshot, reversible composable Patch w/ typed changes); no Diagnostics | ❌ (bytes only) | 🟡 raw `content_xml` `Builder` | 🟡 single-span shape text | ✅ style props; ❌ metadata | 🟡 page insert/remove/rename, shape insert/remove; ❌ move | a + b; c for rich text, annotations, tables/charts |
| odc | ❌ not opened by facade | ✅ `Chart::edit()` | (chart, reversible Patch) | 🟡 bytes only | ✅ `Builder` | 🟡 titles via whole replace | ✅ chart style props | ✅ axes/series/points CRUD | a + b |
| odi | ❌ | ✅ `Image::edit()` | (image, reversible Patch) | 🟡 bytes only | ✅ `Builder` | N/A | ✅ frame props, metadata | 🟡 frame props only | a + b |
| odb | ❌ | ✅ `Database::edit()` | (database, reversible Patch) | 🟡 bytes only | 🟡 raw content_xml builder | N/A | ❌ metadata | ✅ tables/columns/queries CRUD | a + b; c for data rows/forms |
| oth | ❌ | ✅ `Template::edit()` | (template, reversible Patch) | 🟡 bytes only | ✅ `Builder` | ✅ set; 🟡 append-only add | 🟡 inline replace; ✅ metadata | 🟡 sections/notes append/remove | a + b |
| odm | ❌ | ✅ `Master::edit()` + title/link edits | (snapshot, reversible Patch) | 🟡 bytes only | ✅ `Builder` | 🟡 body item add/remove | 🟡 named styles; ✅ metadata | ✅ sections add/rename/remove | a + b |
| odf (formula) | ❌ | ✅ `Formula::edit()` | (formula, reversible Patch, Diagnostics) | 🟡 non-atomic `fs::write` | ✅ `create` / `Builder` | ✅ MathML tree insert/remove/replace | ❌ metadata | N/A | a + b (security fix) |
| pages | ✅ (`litchi_pages::Package`) | 🟡 ~25 scoped owners (section text, footnotes, tables, charts, page layout) | per-family `*Commit`, `inverse()` | ✅ atomic `Package::save` | ❌ (section `Builder` only, detached) | ✅ section/body text replace/insert/delete | ❌ text formatting | 🟡 table delete, drawable order; ❌ section add/remove | a; iWork requests for create/structure |
| keynote | ✅ (`litchi_keynote::Package`) | 🟡 ~45 scoped owners (`edit()`, slide text, order, delete, notes, tables, charts, media) | per-family `Commit`, `inverse()` | ✅ atomic `Package::save` | ❌ (detached `Builder`s only) | ✅ slide title/body text replace/insert/delete | ❌ text formatting | ✅ slide order, delete, skip; ❌ slide add | a; iWork requests for create/slide add/shape CRUD |
| numbers | ✅ (`litchi_numbers::Package`) | 🟡 ~30 scoped owners (cells, formats, sheet order, table relocation, comments) | per-family `Commit`, `inverse()` | ✅ atomic `Package::save` | ❌ (detached `Builder`s only) | ✅ cell set/clear (text/number/bool/date/duration/formula) | 🟡 display formats ✅; ❌ text style | 🟡 sheet move, table move; ❌ sheet/table add/remove, row/col insert/delete | a; iWork requests for create/structure |

---

# Tier 1 — OOXML

## DOCX (`crates/litchi-docx`)

Prefixes: `D/` = `crates/litchi-docx/src/`.

**Verdict.** Read strong (lazy source-backed + eager). Write exists in three engines that cannot be
mixed before save; existing body content is text-replace plus plain-paragraph insert only; no
formatting edits on existing content. Gap class mostly (a)+(b); (c) for formatting, table structure,
styles.

### 1. Open
- Owning (eager OPC): `Package::open`, `open_with_limits(path, ReadLimits)` `D/package/codec.rs:277,286`;
  `from_opc_package` `:373`; `from_reader*` `:420,429`; password (feature `encryption`):
  `open_with_password` `:298`, `open_with` `:317`, `open_with_password_and_limits` `:337`,
  `from_reader_with_password*` `:441,460,476`. Limits: `ReadLimits` (re-export `D/lib.rs:140`).
- Source-backed (lazy): `source_backed::Package::open` `D/source_backed.rs:326`, `from_path*` `:336-410`,
  `from_read_at*` `:428-592` (incl. `SourceReadPolicy`, `ExecutionContext`), `from_source_backed_package`
  `:639`; `to_owned_package()` `:653` (refused for budget-managed sources). No password opener.
- Facade: path open → `DocumentImpl::DocxSource` (`F/document/doc.rs:771-775`, `F/document/types.rs:48`);
  bytes → `docx::Package::from_opc_package` → `DocumentImpl::Docx` (`F/document/doc.rs:987-991`).

### 2. Edit / commit / patch
- Main txn: `Package::edit_document()` `D/package/package/document.rs:36`; source-backed `edit_document`
  `D/source_backed.rs:1142`. `Edit` `D/document/transaction.rs:2144`; `commit` `:5059` → `Commit`
  `:5266` {`snapshot()`, `patch()`, `diagnostics()`}; `Diagnostics` `:5245` (operations, changed only).
- `Patch` `:5300`: `inverse()` `:5338`, `apply(&Snapshot)` `:5357` (exact-source). Durable/compose:
  `to_durable` `D/document/transaction/durable.rs:623`; `Snapshot::compose` `:431` → `Composition::join/commit`
  `:154,183`; `plan_three_way` `:444`; `History` `:287`.
- Publication (owning): `apply_document_patch` `D/package/package/document.rs:99`,
  `publish_document_commit` `:168`, `undo_document`/`redo_document` `:208,221`. Source-backed:
  `publish_document_commit_to_stream` `D/source_backed.rs:1500` (consumes self; refuses transfers `:1517-1528`).
- `Edit` operations: `replace_paragraph_text` `:2273` (preserves run boundaries), `replace_body_paragraph_texts`
  `:2950`, `replace_run_text` `:3563`, hyperlink/field/revision/content-control/table-cell text replacements
  `:3446-4354`, `insert_paragraph(Position, text)` `:4382` (plain), `insert_paragraph_transfer` `:4449`.
  `Operation::RemoveParagraph` exists `:1793` but only as the inverse of insert (`:2044-2048`) — not public.
- Scoped families (each Snapshot/Transaction/Commit/Patch): numbering (`restartNumberingAfterBreak` only,
  `D/numbering/transaction.rs:126,163`), content controls, tracked revisions, ink, task panes,
  stylesWithEffects, glossary, mail merge, variables, web, settings extensions, OpenType, run symbols,
  collapsed paragraphs, footnote columns, charts (`D/chart/transaction.rs:110`, raw `edit_opc` only).
  Source-backed only: section layout `edit_section_layout` `D/source_backed.rs:1027` (page size/margins/
  start/columns, `D/section/layout.rs:255-306`), story text across header/footer/notes/comments/glossary
  (`D/source_backed/story_text.rs:177`), plain paragraph copy/remove (`D/source_backed/paragraph_copy.rs:658,984`),
  tail append, SVG attach, hyperlink sanitize.
- Legacy `document_mut()` → `writer::MutableDocument` `D/package/package/document.rs:430`: not a txn,
  no patch; preserved paragraphs/tables return `None` from `paragraph(i)`/`table(i)`
  (`D/writer/doc/package.rs:644-646,670-673`), so it can only add/insert/remove.
- **Engines cannot be mixed before save**: `ensure_opc_current` refuses (`D/package/codec.rs:682-716`).

### 3. Save / create
- `save(path)` `D/package/codec.rs:522` / `save_with_durability` `:542` — atomic sibling temp + rename
  (`:580-586`); `to_stream(W)` `:606`; encrypted `to_encrypted`/`save_encrypted`/`save_reencrypted`
  `:626-670`. Source-backed: `publish_*_to_stream(self, W)` only, no path save.
- Create: `Package::new()` `:58`, `new_template()` `:249` (.dotx); `MutableDocument` builders
  (`add_paragraph`, `add_heading`, `add_table`, footnotes, `add_comment`, section breaks, headers/footers,
  protection, theme, watermark, TOC — `D/writer/doc/model.rs:350-1196`); `StreamingDocumentWriter`
  `D/streaming.rs:496-933` (plain).

### 4. Typed coverage (existing documents)
| Area | Status |
|---|---|
| Text set | ✅ paragraph/run/hyperlink/cell/field (`D/document/transaction.rs:2273-4354`); secondary stories via source-backed story text |
| Text add | 🟡 plain paragraph `insert_paragraph` `:4382`; no run insert |
| Text remove | 🟡 not on `Edit`; `MutableDocument::remove_paragraph` (`D/writer/doc/model.rs:704`, irreversible); source-backed `remove_plain_paragraph` |
| Text move/swap | ❌ |
| Run/paragraph formatting | 🟡 new content only (`D/writer/run.rs:137-193`, `D/writer/paragraph.rs:830-903`) |
| Sections | 🟡 source-backed layout edit (reversible); owning: section breaks via `MutableDocument` |
| Core / custom props | ✅ `props_mut`/`put_props`/`custom_props_mut` (`D/package/model.rs:823-895`) — no Patch/inverse; app.xml ❌ |
| Comments | 🟡 text edit via story selector; add/remove via `MutableDocument::add_comment`/`remove_comment` (`D/writer/doc/model.rs:954,990`). **Suspected data loss** on opened docs: comments start empty (`model.rs:299`) and `update_comments_part` overwrites `/word/comments.xml` (`D/package/package/parts.rs:141-148`) — needs a test |
| Hyperlinks | 🟡 text set ✅; add only on new paragraphs (`D/writer/paragraph.rs:492`); target change ❌ |
| Pictures | 🟡 add on new paragraphs (`D/writer/paragraph.rs:526,564`); replace/remove existing ❌ |
| Tables | 🟡 cell text ✅; table add/remove via `MutableDocument` (irreversible); rows/cols ❌ |
| Charts | 🟡 graph-byte replace via raw `edit_opc`; no add |
| Styles | ❌ read only (`D/document/package/codec.rs:110`) |

### 5. Gaps
- (a) Wire `edit_document`/`publish_document_commit`/`save`/`Package::new`; for path-opened
  source-backed docs either a facade atomic temp-file wrapper over `publish_*_to_stream` or
  `to_owned_package()` + `save`. Pick one engine per session (mixing is refused).
- (b) Public `Edit::remove_paragraph` (op + inverse + durable codec already exist,
  `transaction.rs:1793,2044`, `durable.rs:1944`); move/swap op; core-props Patch/inverse + app.xml CRUD;
  comment add/remove on opened docs with ID allocation over existing comments (fixes the suspected
  overwrite); hyperlink target change/insert in existing paragraphs; owning-package section-layout edit;
  managed chart edit method; cross-family batch commit.
- (c) Source-preserving run/paragraph/cell property edits on existing content; table row/column
  structure; picture replace/remove; styles CRUD.

### 6. Security
- Encryption: feature `encryption` re-exports `litchi_crypto::ooxml` (`D/lib.rs:109-110`); `inspect` →
  `Kind` (`crates/litchi-crypto/src/ooxml/mod.rs:505`). Plain `open` on encrypted CFB: no typed error
  (*unverified*, likely ZIP error). After password open, `save`/`to_stream`/`edit_document`/semantic edits
  refuse (`D/package/codec.rs:683-691,721-731`); `save_plain`/`save_reencrypted`/`save_encrypted` allowed.
  `PackageEncryption` provenance (`crates/litchi-ooxml-common/src/package_encryption.rs:14-66`).
- Signatures: `is_signed`, `signatures`, `sign`, `resign`, `unsign` (`D/package/package/access.rs:34-110`);
  semantic edits on signed packages refused until `unsign()` (`:228-234`). No silent strip.
- Protection: source-backed changed publish refuses when protection/tracked revisions on
  (`D/source_backed/document_policy.rs:105-108`); no equivalent on owning `apply_document_patch` (inconsistent).

## XLSX (`crates/litchi-xlsx`)

Prefixes: `X/` = `crates/litchi-xlsx/src/`.

**Verdict.** Read strong. Write strong for cells, sheets, layout, merges, hyperlinks through one ADR 0003
`Edit` with reversible patches, atomic save and create. Gaps: (a) facade; (b) styles allocation,
comments, properties, lazy-model write path, swap/duplicate; (c) row/col shift, pictures/charts/tables.

### 1. Open
- `Workbook` (`Arc<Inner>`, `X/workbook/model.rs:423`): `open`/`open_with_limits` `:441,447`, `from_bytes` `:477`,
  `from_slice` `:513`, `from_reader` `:548`, `from_package`/`from_opc` `:588,593`, leniency `:600-615`,
  password variants `:454-575`. `Package` (mutable OPC wrapper) `X/package.rs:36`, same matrix `:44-193`.
- `SourceBackedWorkbook` `X/workbook/source.rs:117`, openers `:181-462`; **no `edit()`, no save, no access
  to its package**.
- Facade: `WorkbookModel::{Owned, Source}` (`F/sheet/adapters/xlsx.rs:86-89`); path/bytes/prepared opens
  use `SourceBackedWorkbook::from_source_backed_package` (`F/sheet/workbook.rs:292-357,497-512`); Owned only on
  the detection fallback (`:560-575`). No password path.

### 2. Edit / commit / patch
- `Workbook::edit()` `X/workbook/model.rs:1313` (refuses encrypted, `:1349-1351`); `Edit::commit()`
  `X/workbook/edit/semantic/transaction.rs:1501` (calls `ensure_unsigned` → `Error::Signed`,
  `X/workbook/edit/codec.rs:99-105`) → `Commit` {`workbook()`, `patch()`, `into_parts()`, `record(&mut History)`}
  (`X/workbook/edit/model.rs:2909-2951`). **No Diagnostics.**
- `Patch` `X/workbook/edit/model.rs:2204`: typed `changes()`, `inverse()` `:2641`, `durable()` `:2760`;
  `Workbook::apply(&Patch)` `X/workbook/model.rs:1319`; `DurablePatch`/`SealedPatch` `X/workbook/edit/wire.rs:70-148`;
  `History` `X/workbook/mod.rs:45`. Composition: `Edit::join` (`transaction.rs:1279`), `plan_three_way`
  (`X/workbook/edit/planning.rs:103-143`). No patch compose.
- Staged ops: sheets `add/add_before/add_after` `transaction.rs:544-590`, `move_*` `:632-690`, `remove` `:1206`
  (cannot mix with other edits, `:1213-1220`); `replace/clear_defined_names` (whole catalog) `:517,535`;
  page setup family `:692-990`; `put/remove_hyperlink` `:922,936`; `copy_cells`/`copy_cells_from`/`move_cells`
  `:1015-1078`. `TabEdit`: `rename` (rewrites modeled references), `activate`, `show/hide/very_hide`
  (`X/workbook/edit/semantic/worksheet.rs:52-92`). `WorksheetEdit`: `set/clear/remove`, `style`, `reset_style`,
  `merge/unmerge`, row/column/defaults layout (`X/workbook/edit/semantic/layout.rs:45-376`),
  `put/replace/remove_hyperlink`, `attach_svg/detach_svg`, form controls (`worksheet.rs:95-440`).
- Other owners outside `Edit`: `Package::edit_{data_model, custom_data, surveys, calculation_metadata,
  page_breaks, slicers, timelines, task_panes, smart_tags}` + `put/remove_custom_props` (`X/package.rs:275-725`);
  pivot edits on `Workbook` (`X/workbook/model.rs:948-1090`); lazy source-backed editors (cell_values,
  hyperlinks, data_validation, conditional_formatting, defined_names, sheet_protection, page_*, auto_filter,
  tab_state, merges) each with `Patch::inverse` + `Diagnostics` (e.g. `X/cell_values/patch.rs:53,315`) but
  publishing **only** to a stream (`X/cell_values/source.rs:537,600`); raw-OPC owners: classic comments
  (`X/workbook/comments/transaction.rs:29`), calc chain (`X/chain/mod.rs:15`), chart sheets.

### 3. Save / create
- `Workbook::to_bytes` `X/workbook/model.rs:1156`, `write_to(W)` `:1171`, `save`/`save_with_durability`
  atomic `:1187-1208`; `*_plain`, `to_encrypted`, `save_encrypted`, `save_reencrypted` `:1163-1311`.
  `Package` same set `X/package.rs:601-825`. Lazy editors: stream publication only, one family per output.
- Create: `Workbook::new()/create()` `X/workbook/model.rs:431-438`; `Package::create()` `X/package.rs:883`;
  `StreamingWorkbookWriter` `X/streaming.rs:267,405,534` (one sheet, scalars).

### 4. Typed coverage
| Area | Status |
|---|---|
| Cell set/clear/remove | ✅ `WorksheetEdit::set/clear/remove` (`worksheet.rs:352-436`); `Content` = Value or scalar Formula (`X/cell.rs:194,342-367`); array/data-table formulas refused |
| Cell move | ✅ limited `move_cells` (≤65,536 cells; merges/range formulas refused) |
| Row/col insert/delete with shift | ❌ (`remove` is "without shifting", `worksheet.rs:429`) |
| Rich text | 🟡 plain `Value::Text` only; `intern_rich` (`X/raw/shared_strings.rs:586`) not wired |
| Cell formatting | 🟡 apply existing `Style` only (`worksheet.rs:397`, `Error::ForeignStyle`); no xf allocation |
| Row/col props, merges | ✅ |
| Sheets add/remove/reorder/rename/hide/activate | ✅; duplicate ❌ |
| Sheet views (freeze/zoom) | ❌ read only (`X/workbook/model.rs:1582`) |
| Custom props | ✅ (`X/package.rs:568-597`); core/app only via `into_plain_opc()` + `litchi_ooxml_common::properties::write` (`crates/litchi-ooxml-common/src/properties/write.rs:30`) |
| Defined names | 🟡 whole-catalog replace |
| Comments | 🟡 raw-OPC txn by `PackURI`; new notes get no VML shape (`X/workbook/comments/transaction.rs:66-68`); threaded ❌ |
| Hyperlinks | ✅ |
| Pictures | 🟡 SVG attach/detach on existing pictures; raster add/remove/replace ❌ |
| Tables / charts / shapes | ❌ read + codec only |
| DV / CF / protection / autofilter | 🟡 lazy editors only |

### 5. Gaps
- (a) Wrap `edit/commit/save/create`; facade must open an Owned `Workbook` for writes (Source cannot edit).
- (b) `SourceBackedWorkbook → Workbook` conversion or a combined atomic multi-family publication;
  `Workbook::into_package` (core/custom props, comments, slicers outside `Edit`); selector-based comments
  on `Edit` + `Worksheet::comments()`; style allocation (model/writer exist, `X/style/stylesheet/writer.rs:25`);
  per-name defined-name CRUD; sheet duplicate; sheet-view edits; DV/CF/protection/autofilter on owned `Edit`;
  Diagnostics on `Commit`; swap (sheet swap as two moves *unverified*).
- (c) Row/column insert/delete with reference rewrite; picture/chart/table package lifecycle; VML for new notes.

### 6. Security
- Encryption (feature `encryption`): opened-with-password workbooks refuse `save/to_bytes/write_to` and
  `edit()` with `Error::EncryptionPolicy` (`X/error.rs:45`; `crates/litchi-xlsx/tests/encryption.rs:29-58`);
  explicit `*_plain`/`into_plain_opc` declassify; `save_reencrypted` keeps mode.
- Signatures: `Edit::commit`/`apply` refuse with `Error::Signed` (`X/error.rs:202-203`); lazy owners refuse
  (`SignedSourceRequiresExplicitPolicy`, `crates/litchi-opc/src/package.rs:2539-2565`); strip only via raw
  `OpcPackage::unsign()`. Facade maps signed error to `Unsupported` (`F/sheet/adapters/xlsx.rs:28-33`).
- Sheet protection: edits refused with `Error::ProtectedSheet`-family (`X/error.rs:254-294`).

## XLSB (`crates/litchi-xlsb`)

Prefixes: `B/` = `crates/litchi-xlsb/src/`.

**Verdict.** Read strong. Create strong (`WorkbookWriter`). Editing opened workbooks is narrow and
fragmented (cells, rename, image/drawing append, merges, many inert metadata families); no sheet
add/remove/reorder, hyperlink/comment/table/layout edits or document properties on opened files.

### 1. Open
- Eager `Package` `B/package/mod.rs:110`: `open*` `:226-255`, `from_bytes*` `:267-309`, `from_slice*`
  `:312-354`, `from_reader*` `:357-399`, `from_opc*` `:402,468`; `workbook()`/`into_workbook()` `:504,509`.
  `Workbook::new*` `B/workbook/package.rs:1039-1131`.
- Lazy `SourceBackedWorkbook` `B/workbook/source.rs:268-451` — immutable, no edit/apply/save.
- Facade: `open_xlsb_workbook*` → `Package::open_with_limits(...).into_workbook()` (`F/sheet/functions.rs:195-231`);
  unified open → `SourceBackedWorkbook::from_source_backed_package` → `XlsbSource` (`F/sheet/workbook.rs:312-320`);
  detected path → `Package::from_opc` (`:581-596`).

### 2. Edit / commit / patch (per family; no document-level `Edit`)
- Cells: `Workbook::edit_cell_values[_by_name]` `B/workbook/package.rs:380`; `Edit::commit` → `Commit{snapshot,patch}`
  (`B/cell_values/worksheet.rs:1088,1126-1149`); `Patch::inverse` `:1201`, `compose` `:1224`,
  `compose_disjoint` `:1244`, `merge_three_way` `:1282`, bytes `:1299,1352`, `History` `:1560,1574`.
  Publication `apply_cell_values(Commit)` `B/workbook/package.rs:418` mutates `*self` (Workbook is not Clone,
  `B/workbook/model.rs:28-32`); `Patch::apply` works on worksheet bytes only (`:1190`).
- Workbook structure: `edit_workbook_structure` `B/workbook/package.rs:325` → `rename_sheet`
  (`B/cell_values/root.rs:187`), `transfer_cell` `:210`, `insert_shared_string/rich_string/formula`
  with `AuthoredStyle` `:241-322`, `insert_image` `:324`, `transfer_image` `:345`, `transfer_drawing_object`
  `:389`; `WorkbookPatch::apply/inverse/merge_three_way` `:549,569,587`.
- Others: sparklines (no public inverse), cell watches, theme, custom data, data model, calc chain,
  volatile deps, XML maps, slicers/timelines (only these return a new `Package`, `B/package/mod.rs:166,219,455`),
  external links, connections, threaded comments (`B/comments/threaded/edit.rs:26-285`, **no inverse**),
  shared workbook (no inverse), web-extension bindings. Direct in-place setters (no patch): merged cells
  `B/workbook/codec/package.rs:236-334`, scenarios, ribbon/task panes, VBA, connections.
- Raw escape: `Workbook::edit_opc(|&mut OpcPackage| ..)` (`B/workbook/package.rs:699-707`).
- No commit carries diagnostics.

### 3. Save / create
- `Package::to_bytes` `B/package/mod.rs:519`, `write_to` `:524`, `save`/`save_with_durability` atomic `:529-550`.
  `Workbook::save<W: Write+Seek>` `B/workbook/package.rs:1033` (no path save; `into_opc_package` is `pub(crate)`, `:673`).
- Create: `WorkbookWriter::new` `B/writer/workbook/model.rs:83` (`add_worksheet` `:187`, `add_chart_sheet` `:197`,
  names, links, theme, styles); `save<W: Write+Seek>` `B/writer/workbook/package.rs:44`; `Package::create()`
  `B/package/mod.rs:114`. Document properties are a fixed template (`B/writer/workbook/package.rs:191-221`).
  No `Workbook → WorkbookWriter` conversion.

### 4. Typed coverage (existing / create)
| Area | Existing | Create |
|---|---|---|
| Cell set/insert/remove | ✅ same storage family (`B/cell_values/worksheet.rs:769-1050`) | ✅ |
| Move/swap, row/col insert/delete | ❌ | 🟡 `insert_row/column`, `delete_row/column` |
| Formula | ✅ inert token blob (`:948`) | ✅ |
| Cell formatting | 🟡 existing XF index; new XF only for inserted cells | ✅ |
| Row/col layout | ❌ | ✅ |
| Sheets | 🟡 rename only | add only |
| Merges | ✅ setters | ✅ |
| Doc properties | ❌ (raw `edit_opc` + ooxml-common `properties::write`) | ❌ fixed template |
| Classic comments / hyperlinks | ❌ | ✅ |
| Threaded comments | ✅ (no inverse) | *unverified* |
| Pictures / shapes / charts | 🟡 append/transfer only | ✅ |
| Tables / DV / CF / names / protection | ❌ | ✅ |

### 5. Gaps
- (a) No facade edit/save; unified open returns immutable `XlsbSource` — edit path must reopen through
  `Package`. Map cells/rename/merges/images/threaded comments/create/save.
- (b) `Workbook::save(path)` or `into_package()`; typed `edit_properties` (+ `WorkbookWriter` setters);
  publish a cell-values `Patch` at workbook level (undo); uniform snapshot-returning commit + diagnostics;
  `set_style(AuthoredStyle)` for existing cells; inverses for threaded comments/shared workbook/sparklines;
  single explicit signature policy.
- (c) Sheet lifecycle on opened packages (catalog + XTI/name remap); hyperlink and classic-comment record
  splicing; table/DV/CF/layout/defined-name/protection edits on opened worksheets; image/chart remove/replace.

### 6. Security
- Encryption: not supported (no `litchi-crypto` dependency; `crates/litchi-xlsb/docs/FEATURE_MATRIX.md:102`);
  encrypted input fails at OPC ingress (exact error *unverified*).
- Signatures: `is_signed`/`signatures`/`sign`/`resign`/`unsign` (`B/workbook/package.rs:746-806`). On edit,
  behaviour differs by family: **silently stripped** by `edit_opc`, cell values (`B/cell_values/workbook.rs:167`),
  workbook structure, watches, sparklines, slicers, timelines, connections, threaded comments; **refused**
  with `SignedSourceRequiresExplicitPolicy` by theme (`B/theme.rs:458`), data model, calc chain, custom data.

## PPTX (`crates/litchi-pptx`)

Prefixes: `P/` = `crates/litchi-pptx/src/`.

**Verdict.** Read strong (eager + lazy). Write strong as transactional infrastructure (reversible,
joinable, durable patches) but semantically narrow: whole-shape text, add/remove whole objects, slide
order, comments. No formatting edits on opened decks; no hyperlink write.

### 1. Open
- Eager `Package`: `open`/`open_with_limits` `P/package/codec.rs:180,189`, `from_vec*` `:225,234`,
  `from_bytes*` `:246,255`, `from_reader*` `:201,210`, `from_opc_package` `:264`; password openers
  `P/package/encryption.rs:34-135`. Read model `Package::presentation()` `P/package/model.rs:176`; shape scene
  enum `P/shape/model.rs:428-440` (text is a flattened string, no paragraph/run model).
- Lazy `SourceBackedPresentation` `P/presentation/source.rs:408`, openers `:879-1145`; no shape enumeration,
  no password opener.
- Facade: path/bytes → `SourceBackedPresentation` (`F/presentation/prs.rs:690-691,826-828,879`); eager
  `Package::from_opc_package` only on detection fallback (`:928`).

### 2. Edit / commit / patch
- Main: `Package::opened_presentation()` `P/package/model.rs:253` / `opened_presentation_transaction()` `:296`
  (refused for freshly authored packages until save+reopen, `:271-276`); `Transaction::commit()`
  `P/opened/transaction.rs:1241` → `Commit{snapshot, patch}` (`:2025-2072`), **no Diagnostics**;
  publish `apply_opened_presentation_commit/patch` `P/package/model.rs:550,568` (refuse encrypted, `:584`).
- Patch: `inverse` `P/opened/patch.rs:380`, `conflicts_with` `:394`, `join` `:412`, `three_way` `:436`,
  `to_bytes/from_bytes` `:492,523`, `History` `:646-742`.
- Ops (`P/opened/transaction.rs`): `move_slide` `:109`, `reorder_slides` `:138`, `remove_slide` `:181`,
  `set_shape_text(s)` `:210,263`, `add_text_box/rectangle/ellipse/picture` `:293-354`, `remove_shape` `:432`,
  `transfer_shape` `:486`, `set_notes_text` `:650` (existing notes only), table styles `:697,706`,
  `add_table` `:719`, `add_chart` `:751`, `store_media` `:782`, masters/layouts `:797-823`, legacy comments
  `:833-866`, modern comments/replies `:891-1165`, `transfer_relationship_closure` `:1197`.
- Plans: slide copy (`P/opened/copy_plan.rs:188`, no inverse), slide removal (`P/opened/remove_plan.rs:282`,
  inverse `:74`), cross-slide copy.
- Source-backed editor `P/presentation/source.rs:672`: `edit_slide` `:1684` (transition, shape text; one op per
  edit `:2291-2295`), batch ≤32 slides `:1713`, slide order (`P/presentation/order.rs:118,132`), SVG replace/attach.
- Other families: notes (`P/package/model.rs:703-755`), change tracking (only family with Diagnostics),
  readonly-recommended, placeholder ext, task panes, ink, table-style catalog. Sections/custom shows,
  color map, guides, legacy comment update are **free functions on `&mut OpcPackage`** — reachable only
  via `Package::edit_opc` (`P/package/model.rs:1867`).

### 3. Save / create
- `Package::save` (atomic, `P/package/codec.rs:326-353`), `to_bytes` `:360`; encrypted variants
  `P/package/encryption.rs:157-297`. No `Write`-sink save on `Package`. Source-backed: stream publications only.
- Create: `Package::new()` `P/package/codec.rs:38-174`; `presentation_mut()` → `MutablePresentation`
  (add/insert/delete/duplicate/move slide, `P/writer/presentation.rs:58-217`), `MutableSlide`
  (title, notes, transition, background, text boxes, shapes), `MutableShape` formatting
  (`P/writer/shape.rs:200-272`). No picture/table/chart/hyperlink adder in the writer.

### 4. Typed coverage (opened decks)
| Area | Status |
|---|---|
| Text | 🟡 whole-shape replace (first `a:t` gets text, `:200-205`); add = new text box; remove = whole shape; no paragraph/run addressing; no table-cell text set |
| Formatting | ❌ on opened decks; ✅ creation writer only |
| Slides | ✅ move/reorder; 🟡 remove/duplicate (dependency-free only, `P/opened/remove_plan.rs:270-276`); ❌ blank/layout-based add; ❌ hide/rename |
| Sections / custom shows | ✅ typed with inverse but only via `edit_opc` |
| Masters/layouts | ✅ add master/layout, remove unused layout |
| Metadata | ✅ custom props (`P/package/model.rs:74-100`); core only via `edit_opc` |
| Comments | ✅ legacy + modern CRUD |
| Hyperlinks | ❌ write (read only, `P/presentation/model.rs:198`) — FEATURE_MATRIX ✅ is not backed by code |
| Pictures | ✅ add/remove; raster replace ❌ |
| Tables | ✅ add/remove; cell edit ❌ |
| Charts | ✅ add/remove; data edit ❌ (`litchi_drawingml::chart::edit` exists, `crates/litchi-drawingml/src/chart/edit/transaction.rs:24-219`, not wired) |

### 5. Gaps
- (a) Facade holds only `SourceBackedPresentation`; writes need the eager `Package` path
  (`opened_presentation → edit → commit → apply → save`); patch/History map directly.
- (a)→(b) Sections, custom shows, core props, legacy comment update via typed `Package` wrappers instead of `edit_opc`.
- (b) Hyperlink set/remove; slide hide/rename; shape bounds + z-order; table cell text; raster picture
  replace; create notes for slides without notes (`P/notes/package.rs:1216`); `put_core_props`; wire
  drawingml chart edit; atomic path save for source-backed publication; Diagnostics on `opened::Commit`;
  inverse for `apply_slide_copy_plan`; blank/layout slide add (crate-internal `insert_slide_binding`
  suggests basis exists — b/c).
- (c) Paragraph/run-addressed text + formatting edits on existing `txBody` (no DrawingML text-body
  parser/splicer, `crates/litchi-drawingml/src/text/body/mod.rs:8-10`); closure-aware removal of slides with
  notes/charts/media/comments.

### 6. Security
- Encryption: password openers take `&str` (not the redacted credential type ADR 0006 asks for);
  decrypted packages refuse typed mutation, `save`, `to_bytes` with `Error::EncryptionPolicy`
  (`P/package/model.rs:1966-1972`; `P/package/codec.rs:342-348`).
- Signatures: only at OPC level via `Package::opc()`/`edit_opc`. **`opened::Transaction::commit` silently
  unsigns** changed signed packages (`P/opened/transaction.rs:1255-1270`); `edit_opc`, comments, tags,
  events also unsign; copy/remove plans, readonly-recommended, collaboration and source-backed publisher
  refuse. Inconsistent — (b) to unify behind an explicit policy.

---

# Tier 2 — Legacy OLE

## DOC (`crates/litchi-doc`)

Prefixes: `W/` = `crates/litchi-doc/src/`.

**Verdict.** Read complete. Write: fresh documents via `Writer`, and bounded in-place content
replacement plus ~12 scoped owners on existing files. No structural editing; no facade wiring;
facade drops the package at open.

### 1. Open
- Eager `Package` (`W/package/model.rs:559-564`): `open`, `open_with(PackageOpenOptions)`, `open_with_limits`
  `W/package/codec.rs:28,37,43`; `from_reader*` `:65,70`; `from_ole_file*` `:134,139`; `document*()` `:158-184`
  builds a read-only `Document` (no OLE handle or source bytes kept, `W/document/model/state.rs:67-88`).
  Options: `Limits`, `PackageOpenOptions`, `OpenOptions{password, leniency}`, zeroizing `Password`
  (`W/package/model.rs:13-425`).
- Edit-oriented: `body_text::Snapshot::open/open_bounded/open_bounded_with_policy/parse/from_bytes`
  (`W/body_text.rs:642-824`); lazy `body_text::source::SourceSnapshot::open(Arc<dyn ReadAt>)`
  (`W/body_text/source.rs:511-554`).
- Facade: `doc::Package::from_ole_file(ole)` → `document()`; package dropped (`F/document/doc.rs:960-975`,
  `F/document/types.rs:36`); password reaches the reader via `open_with_doc_options` (`F/document/doc.rs:861,881,1067-1083`).
  No source-backed DOC variant.

### 2. Edit / commit / patch
- Root: `Snapshot::edit()` `W/body_text.rs:930` → `Edit`: `replace_paragraph` `:1654`, `replace_text(TextTarget)`
  `:1669` (story paragraph, simple table cell, simple field result), `set_paragraph_bold` `:1747`,
  `set_character_property(Bold|Italic|Underline)` `:1757`, non-destructive `dispose_revision` `:1997-2033`,
  embedded object add/remove `:1913,1956`, donor transfers `:1768-1789`, saved selection / DOFR, `rollback` `:2074`.
- `commit()` `:2086` → `Commit{snapshot, patch}` (`:2296-2326`) — re-opened and validated; **no Diagnostics**
  (opt-in profiled commit behind `performance-diagnostics`, `:2143-2171`).
- `Patch` (`:2329-2490`): exact-source `apply` `:2412`, `inverse` `:2439`, `to_durable` `:2457` (refuses
  auxiliary-table changes); `compose` `:945` → `Composition::join/commit` `:2981,3015`; `History` `:48,955`;
  `plan_three_way` `:1310`.
- Typed refusals (`:315-375`): `StoryLengthChange` (length change outside main story), `StructuralContent`,
  `TrackedText`, `FormattingDependency`, `ComplexField`, `ComplexTableCell`, `SignedSource`, `Protected`, limits.
- Source-backed splice: `SourceSnapshot::edit(Position)` same-UTF-16-length only; `Commit` has
  `diagnostics()` (`W/body_text/source.rs:875-1125`).
- Scoped families (own Snapshot/Commit/Patch with inverse/revert): tracked revisions (add, add_text,
  replace, remove, accept, reject — `W/tracked_revision/transaction/edit.rs:67-149`), embedded objects,
  **property sets** (`edit_summary_information`, `edit_document_summary_information`,
  `edit_user_defined_properties` — `W/package/property_set.rs:300-449`; fails if the stream is absent, `:309`),
  annotation bookmarks, smart tags, subdocuments, captions, envelope, command bars, route slip, DOFR,
  saved selection, OLE controls, protection. **No shared cross-family transaction.** No existing-doc edit
  families for numbering, styles, sections, comments add/remove, bookmarks, hyperlink fields, tables.

### 3. Save / create
- `Writer::new()` `W/writer/core/model/semantic.rs:22`; `save`/`save_with_durability` atomic via CFB writer
  (`W/writer/core/package/package.rs:112-138`, `crates/litchi-cfb/src/writer/core.rs:2034-2057`); `write_to<W: Write+Seek>` `:1178`.
  No `Writer::from(Document)`.
- Edits: `Snapshot::finish() -> Vec<u8>` (`W/body_text.rs:1322`) — no path save; `RevisionEditor::finish`
  (`W/tracked_revision/package.rs:2631-2643`); source-backed `Commit::write_to/save/save_with_durability`
  atomic (`W/body_text/source.rs:1133-1160`).
- Writer encryption: `set_password(pw, EncryptionProfile)` XOR / Office RC4 / CryptoAPI RC4 (`semantic.rs:66`).
- **`Writer::set_property` is a silent no-op** (`semantic.rs:1046`; no SummaryInformation written).

### 4. Typed coverage
| Area | Existing | Create |
|---|---|---|
| Text set | 🟡 whole paragraph / simple cell / simple field; length change main story only | ✅ |
| Text add / remove | 🟡 via tracked insertion/deletion + accept (*unverified end to end*) | append only |
| Move / swap | ❌ | ❌ |
| Run formatting | 🟡 bold/italic/underline on uniform targets | ✅ ~25 fields |
| Paragraph formatting | ❌ | ✅ |
| Sections | ❌ | 🟡 one section, no page size/margins |
| Metadata | ✅ property-set txn (existing streams only) | ❌ silently dropped |
| Comments | 🟡 equal-length text overwrite | ✅ `add_comment` |
| Hyperlinks | 🟡 display text of simple field | ✅ |
| Pictures | 🟡 donor transfer only | ✅ |
| Embedded OLE | ✅ add/remove/reorder/replace | *unverified* |
| Tables | 🟡 simple cell text | ✅ (emitted after all paragraphs, cannot interleave, `package.rs:302,436`) |
| Charts | ❌ | ❌ |

### 5. Gaps
- (a) Retain source bytes (or add `DocSource(body_text::Snapshot)`), map `set` text → `replace_text`,
  b/i/u → `set_character_property`, save → atomic write of `finish()` bytes, `create` → `Writer`,
  metadata → property-set txn. Verify paragraph-index agreement between `body_text::Snapshot::paragraphs(Projection)`
  and `Document::paragraphs()` (*unverified*).
- (b) Atomic path save for `finish()` owners; make `Writer::set_property` emit property streams; let the
  property-set txn create a missing stream; non-main-story length changes; cross-family join adapter;
  untracked "insert text at CP" (wrap add_text+accept); `Writer` page size/margins.
- (c) Structural body editing (paragraph insert/delete/move/swap, table/row add/remove, sections —
  needs PAPX FKP insertion and PLCF shifting); arbitrary-range run formatting; existing-doc comment/
  bookmark/hyperlink-field CRUD; numbering/style edits; picture insert/delete; interleaved tables in
  `Writer`; charts; `Document → Writer` round trip.

### 6. Security
- Encryption detection via FIB flags after open (`W/parts/fib.rs:209`); `EncryptionKind`
  (`W/package/model.rs:431-444`). Open without password → `Error::PasswordRequired`; wrong →
  `InvalidPassword`; unsupported → `UnsupportedEncryption(kind)` (`W/encryption/codec.rs:233-301`).
  Edit engines refuse encrypted DOCs (`W/tracked_revision/package.rs:1009-1011`). No decrypt-to-bytes/
  re-encrypt for existing docs (`encryption` module is private, `W/lib.rs:147`).
- Protection: changed publication needs `ProtectionPolicy::AllowProtected(ProtectionAuthorization::audited(..))`
  (`W/parts/protection/policy.rs:103-176`).
- Signatures: verify only (`W/package/codec.rs:302-311`, `litchi_sign::cfb::verify`); changed commits refused
  (`Refusal::SignedSource`, `W/body_text.rs:1479-1480`); CFB containers with signature/DRM/encryption storages
  rejected at admission (`crates/litchi-ole-common/src/protection.rs:12-29`). Never stripped.

## XLS (`crates/litchi-xls`)

Prefixes: `L/` = `crates/litchi-xls/src/`.

**Verdict.** Read strong (eager + lazy). `Writer` is create-only and broad. Edits of opened files go
through ~10 separate byte-level families; no workbook-wide edit model; facade wires none.

### 1. Open
- Eager `Workbook<R: Read+Seek>`: `new` `L/workbook/package.rs:76`, `new_with_options(OpenOptions)` `:84`,
  `from_ole_file*` `:103,111`; decrypts via `prepare_workbook_stream(.., password)` `:131`. `OpenOptions`
  (`L/workbook/model.rs:80-190`: password, leniency, compatibility profile).
- Lazy `SourceBackedWorkbook` (read-only, `L/workbook/source.rs:1`): `from_read_at*`/`from_path*`
  `:869-920`; `materialize_eager()` `:1058`; no password.
- Per-family byte opens: `cell_values::Snapshot::from_bytes*` (`L/cell_values/mod.rs:568-584`),
  `sheet_visibility` (`L/sheet_visibility.rs:80`), `comments` (`L/comments/transaction.rs:154`),
  `property_set` (`L/workbook/package/property_set.rs:36`), `ole_object` (`L/ole_object/edit/model.rs:30`),
  `chart::Editor::open` (`L/chart/package.rs:76`), `PivotViewEditor::new` (`L/pivot_editor.rs:39`).
- Facade: `open_xls_workbook*` use `Workbook::new` default options — **no password path**
  (`F/sheet/functions.rs:147-187`); detection reuses OLE via `from_ole_file` (`F/sheet/workbook.rs:555`);
  path source via `SourceBackedWorkbook` (`F/sheet/workbook.rs:117-135`).

### 2. Edit / commit / patch
- `cell_values`: `Snapshot::edit()` `L/cell_values/mod.rs:795`; ops `set_number` `:1284`, `set_numeric` `:1332`,
  `set_value` `:1665` (only RK↔LabelSst family change, `:1651-1663`), copy from other snapshots `:1368-1625`,
  `clear_cell` (Blank only) `:1736`, `remove_cell` `:1927`, `insert_cell`/`insert_rich_text_cell`/`insert_formula`
  (empty cells only) `:1763-1886`, `set_style` `:1950`, `duplicate_style` `:1979`, `author_style` (raw 20-byte XF)
  `:1999`, `insert/delete_rows/columns` `:2337-2366`, `rename_sheet` `:2375`, `join` `:1089`.
  `commit()` `:2662` → `Commit{snapshot, patch, diagnostics}` (`:3192-3260`). `Patch::apply/inverse`
  `:5181,5195`; `SemanticPatch` inverse/apply/three-way/transfer/JSON (`:3694-3905`); `History` `:555,812`.
  Source-backed commits `:2614-2655` (numeric, family-preserving).
- `sheet_visibility` (hide/very_hide/unhide, `L/sheet_visibility.rs:297-354`), `comments` (replace author/text
  only; remove/visibility refuse, `L/comments/transaction.rs:334-489`), `property_set`
  (`L/workbook/package/property_set.rs:140-235`, creates missing sections), `ole_object` (full CRUD), all
  reversible. `chart::Editor` (formula/cache/chart area, chart-sheet remove, **reorder all tabs**
  `L/chart/package.rs:868-905`) and `PivotViewEditor` return `Vec<u8>` only (no patch). Authoring of
  charts always refuses (`L/chart/package.rs:296-497,742`).

### 3. Save / create
- `Writer::new` `L/writer/core/codec/lifecycle.rs:16`; `save`/`save_with_durability` atomic, `write_to<W: Write+Seek>`
  (`L/writer/core/package.rs:31-90`). No property streams written (`:93-129`). Broad builders: cells,
  formats, merges, widths, panes, page setup, CF, DV, autofilter, list objects, hyperlinks, names,
  comments, shapes, pivots, protection, password. Sheets: `add_worksheet` only. No pictures/charts.
- Edit path: `Snapshot::bytes()` `L/cell_values/mod.rs:744` (no save); source-backed commits have atomic
  `save`/`write_to` (`:3421-3570`); visibility/comments same.

### 4. Typed coverage (existing / create)
| Area | Existing | Create |
|---|---|---|
| Cell set | 🟡 family-preserving; formula cached result only | ✅ |
| Cell add / remove | ✅ empty cells / 🟡 refuses MulRk interior & formula groups | ✅ / ❌ |
| Rows/cols insert/delete | ✅ | n/a |
| Formatting | 🟡 XF index or raw XF bytes | ✅ |
| Sheets | rename ✅, hide ✅, reorder-all (bytes, no patch), remove chart sheets only; add/remove worksheet ❌ | add only |
| Metadata | ✅ property-set txn | ❌ |
| Comments | 🟡 author/text replace | ✅ |
| Hyperlinks / merges / names / DV / CF | ❌ | ✅ |
| Pictures | ❌ (not even read) | ❌ |
| Charts | 🟡 series formula/cache, chart area | ❌ |
| Pivots | ✅ (bytes) | ✅ |

### 5. Gaps
- (a) Facade edit/save/create wiring; password passthrough; metadata on new files by chaining
  `Writer::write_to` → `property_set` txn.
- (b) `replace_cell` (any-type set); save/write_to on ordinary `cell_values` Commit; Patch for
  `chart::Editor`/`PivotViewEditor`; typed XF builder; `Writer` error/rich-text cells and sheet
  rename/remove/reorder; move sheet reorder into the main txn; CFB splice wrappers for record-level families.
- (c) Workbook-wide edit txn; worksheet add/delete in opened files; comment add/remove/visibility (OfficeArt);
  hyperlink/merge/name/DV/CF CRUD in opened files; pictures; chart authoring; `Workbook → Writer` round trip.

### 6. Security
- FILEPASS: eager reader decrypts XOR/BinaryRc4/CryptoAPI RC4 with password; else `PasswordRequired`/
  `InvalidPassword`/`UnsupportedEncryption` (`L/encryption.rs:556-581`); lazy reader `EncryptedUnsupported`
  (`L/workbook/source.rs:337`). Edit families return `PasswordRequired` on FILEPASS
  (`L/cell_values/mod.rs:5802`) and reject protected containers (`crates/litchi-ole-common/src/object/codec.rs:1327`).
  Property-set FILEPASS check *unverified*.
- Signatures: verify only (`L/workbook/package.rs:318-330`); edits refuse signed inputs. Writer encryption
  RC4/CryptoAPI; XOR needs explicit `WeakEncryptionPolicy` (`L/encryption.rs:39-48`).

## PPT (`crates/litchi-ppt`)

Prefixes: `T/` = `crates/litchi-ppt/src/`.

**Verdict.** Read strong (eager + lazy). Create strong via `Writer` (metadata dropped). Existing-file edits
via an append-only `slide_order` root (move/remove/transfer slides, hidden/advance/transition, whole-shape
text, anchors, media) plus ~20 single-feature owners; no rich text, no inserting shapes/pictures/tables/slides.

### 1. Open
- Eager `Package<R>` `T/package/model.rs:284`: `open*` `T/package/codec.rs:32,41`, `from_reader*` `:68,77`,
  `from_ole_file*` `:122,131`, `presentation*` `:155-190`; property-set readers `:242-304`.
- Lazy `SourceBackedPackage` `T/package/source.rs:27` (`from_read_at*`, `from_path`, `from_shared_if_powerpoint`).
- `RecordLimits`, `OpenOptions{password, leniency}` (`T/package/model.rs:12-147`).
- Facade: `SourceBackedPackage` → `.presentation()` with no options (`F/presentation/prs.rs:785-796`); OLE path
  → `Package::from_ole_file` (`:890-909`). **No password path.** Slide catalog unsupported (`:1141`).

### 2. Edit / commit / patch
- Root `slide_order` (`T/slide_order.rs`): `Snapshot::from_bytes*` `:596,606`, `edit()` `:859` → `Transaction`
  `:1511`: `move_slide` `:1599`, `remove_slide` `:2028` (payload remains as unreachable history),
  `insert_transfer` `:2056` (donor copy), `set_slide_hidden/advance/transition_visual` `:1634-1710`,
  `set_shape_text(s)` `:1752,1811`, `set_shape_anchor` `:1903`, external media `:1947,1986`; `commit` `:2121`,
  `commit_profiled` `:2250`, `rollback` `:2402`. `Commit{snapshot, patch}` (`:2420-2440`), no Diagnostics
  value. `Patch` `:2447`: `apply` `:2589`, `inverse` `:2601`, `to_durable` `:2784`; `History`; only moves
  compose (`:890-920`); `plan_three_way` `:976`. Refusals: review history, nested CFB, signed/encrypted
  (`:1051-1064`, `:219-248`).
- `text_edit`: length change only for 1 paragraph run + 1 character run (`T/text_edit.rs:1-9`); lazy variant
  same-length only, with atomic `SourceBackedCommit::save` (`:700-736`).
- `property_set` txn (`T/package/property_set.rs:96-388`) — edit SI/DSI/user-defined; missing stream →
  `StreamNotFound` (`:188`).
- Single-feature owners: fonts (remove/reorder always refuse), picture bullets, document comparison,
  document structure (record-level), external media, broadcast, client anchor/data, diagrams, masters,
  slide sync, animation (`finish() -> Vec<u8>`, no patch), OLE `Editor` (no patch), chart `PackageEditor`
  (standalone Graph package; bridged via `Chart::replace_package`, `T/chart/model.rs:248`). Record-level
  owners require `embedded::object::Editor::replace_persisted_record` + `finish()`.
- Parsed shape setters return `MutationError::SourceBound` (`T/shapes/picture.rs:281-291`) — dead end.

### 3. Save / create
- `Writer::new/new_widescreen/with_dimensions` (`T/writer/core/model/semantic.rs:42-59`); `save`/
  `save_with_durability` atomic (`T/writer/core/package.rs:305,323`); `write_to<W: Write+Seek>` `:759`.
  Cannot load an existing file. **`set_property` silently dropped** (`semantic.rs:1764`; SummaryInformation
  fixed, `package.rs:129-160`).
- Existing-file owners return bytes (append UserEdit + persist directory); no atomic path helper except `text_edit`.

### 4. Typed coverage (existing / create)
| Area | Existing | Create |
|---|---|---|
| Text set | 🟡 whole text atom; length rule | ✅ rich text boxes, notes |
| Text add/remove/move | ❌ | add only |
| Formatting | ❌ | ✅ |
| Shape geometry | 🟡 anchor only | ✅ |
| Shape add/remove | ❌ | add only |
| Slides | ✅ move/remove/transfer, hidden, advance, transition; ❌ blank insert | add/delete/move |
| Metadata | ✅ property-set txn | ❌ dropped |
| Comments / hyperlinks / pictures / tables | ❌ (read only) | ✅ |
| Charts | 🟡 replace standalone Graph package | 🟡 Bar/Line/Pie |

### 5. Gaps
- (a) Facade: bytes → `slide_order::Snapshot`; verb dispatch; atomic path save for `bytes()`; `create` →
  `Writer`; password passthrough.
- (b) Serialize `Writer::properties` (or chain property-set pass); compose single-feature owners into
  one batch root (today chain bytes → reopen; patches don't join); package-level adapters producing
  patches for record-level owners; patches/inverses for animation/OLE/chart editors; swap via two moves.
- (c) Rich text editing of existing shapes (StyleTextProp remapping); inserting shapes/pictures/tables/
  hyperlinks/comments/blank slides into existing files (OfficeArt/BStore/ExObjList allocation); notes
  editing; shape remove/reorder; font remove/reorder.

### 6. Security
- Detection: `CurrentUser::is_encrypted` (`T/current_user.rs:254`); `EncryptionKind{CryptoApi, Unknown}`.
  Open with feature `encryption` → `PasswordRequired`/`InvalidPassword`/`UnsupportedEncryption`
  (`T/presentation/package.rs:69-88`); facade surfaces missing password as `InvalidFormat`
  (`T/package/model.rs:316-318`). Writer can encrypt (`set_password`, CryptoAPI RC4 40–128-bit).
- Signatures: verify `Package::signatures` (`T/package/codec.rs:260,270`); signing via `litchi_sign::cfb::Editor`
  (`crates/litchi-sign/src/cfb.rs:403-527`). Incremental editors refuse signed/encrypted packages
  (`T/embedded/object/editor/lifecycle/open.rs:158-191`). Never stripped.

---

# Tier 3 — delayed formats (summary only)

## RTF (`crates/litchi-rtf`)

Prefixes: `R/` = `crates/litchi-rtf/src/`.

- **Open**: eager `litchi_rtf::Document::{parse, from_bytes, from_owned_bytes, open}` (+limits/options)
  `R/api/facade.rs:57-175`; `ParseLimits`/`ParseOptions`/`Leniency` (`R/codec/limits.rs:18-216`). Facade keeps the
  raw `RtfDocument<'static>` (`F/document/types.rs:54`, `F/document/doc.rs:977-984`), not the editable snapshot.
- **Edit/commit**: ✅ `Document::edit()` `R/api/facade.rs:482-490` → `Edit::commit()` `R/edit.rs:3397` →
  `Commit{snapshot, patch, diagnostics}` (`R/edit.rs:8949-9000`); `Patch::apply/inverse/to_durable`
  (`R/edit.rs:9913-9960`), composition/merge (`R/edit/composition.rs:103-816`), `History` (`R/edit.rs:47-116`).
  Ops: body/paragraph/span text replace, 14 run property toggles, paragraph insert/remove/move/split/merge
  (plain sources), alignment/layout, table-cell/header-footer/annotation/note/shape text, same-length
  picture payload replace/remove, transfer, tail append. Protected sources refuse changed publication
  (`R/edit.rs:4719-4725`).
- **Save**: `write_to(W)`/`to_bytes()` (`R/api/facade.rs:434-448`), byte-exact for retained source; no path
  save (deliberately left to facade, `:428-430`).
- **Create**: ❌ semantic builder; only low-level `write::Writer` / `streaming::Writer` (plain runs).
- **Gap class**: (a) owner + atomic save + limits wiring, selector→`TextSpan` mapping; (b) swap, font face/
  colour/style setters, non-ASCII/compressed edits; (c) create builder, section/table/comment/hyperlink/
  metadata structure CRUD.
- **Security**: no encryption/signatures in RTF; `\password` protection inert metadata; changed edits refused.

## ODF shared layer (`crates/litchi-odf-common`)

- Package read (`OwnedPackage`, `SourceBackedPackage`), manifest decryption, `PackageWriter` with
  `set_encryption` and `set_document_signer` (`core/writer.rs:772,804`), content splice publication
  (`package/edit.rs:41-837`), streaming publication (`core/source_publication.rs`), `MetaXmlPatch` (7 simple
  fields, `core/metadata/codec.rs:218-226`), signature verify. Rebuild **always drops signatures**
  (`core/writer.rs:1618-1622`) and refuses encrypted sources unless the writer encrypts (`:1605-1608`).
  No public atomic path writer (flat atomic save is `pub(crate)` in odt). (b): public atomic writer, typed
  locked/signed envelope, wider `MetaXmlPatch`.
- **Facade reachability**: only odt/ods/odp are opened by the facade; odg/odc/odi/odb/oth/odm/odf are
  detected (`crates/litchi-odf-common/src/detect.rs:241-246`) but dispatch falls to `_ => None`
  (`F/detection_smart/detected.rs:1753-1766,1915-1929`), and `crates/litchi/Cargo.toml` has no dependency on
  those crates. Flat ODF is refused for all families (`F/document/doc.rs:1027-1033`).

## ODT (`crates/litchi-odt`)

- **Open**: `Document::open/from_bytes/*_with_password/*_with_limits` (`crates/litchi-odt/src/document/package.rs:82-188`);
  lazy `SourceBackedDocument` (`document/source.rs:51-196`). Facade: `OdtSource` via `from_read_at`; bytes via
  `from_prepared_package` (`F/document/doc.rs:1036-1042`); no password path.
- **Edit/commit**: ✅ `Document::edit()` (`document/package.rs:274`) → `transaction::Edit` → `Commit{snapshot, patch,
  results}` (`transaction.rs:2314-2361`), no diagnostics; `Patch` = exact before/after swap with `inverse`,
  `durable`, `JoinedEdit`, `MergePlan`, `History`. Ops: plain paragraph insert/replace/remove/move, append
  run/hyperlink/line break, notes, ruby, fields, tracked changes, RDF, protection, forms, embedded charts,
  images/objects, scripts. Second, **lossy** engine `MutableDocument` (sections, master pages, headers/footers,
  tables, metadata; structural edits regenerate `content.xml` from a projection, `mutable/package.rs:452-460`).
- **Save**: `Document::save` / `MutableDocument::save` / `Builder::save` use `std::fs::write` — **not atomic**
  (`document/package.rs:316-319`, `mutable/package.rs:396-399`); atomic only for flat `.fodt` (`flat/atomic.rs:21-31`).
- **Create**: ✅ `Builder` (`builder/model.rs:113`, paragraphs/headings/lists/tables/sections/notes/forms; no images/charts/comments).
- **Gap class**: (a) facade; (b) atomic save, metadata/comments as `Edit` ops, hyperlink update/remove,
  sections/master pages in `Edit`, commit diagnostics; (c) direct formatting on existing runs, rich
  inline/block/table edits without loss.
- **Security**: `EnvelopeKind{Plain, Signed, Encrypted}`; changed commits on signed/encrypted refused
  (`transaction.rs:5850-5861`); `MutableDocument` rebuild drops signatures and refuses encrypted.

## ODS (`crates/litchi-ods`)

- **Open**: `Spreadsheet::open/from_bytes/*_with_password` (`crates/litchi-ods/src/facade/mod.rs:126-196`); lazy
  `SourceBackedSpreadsheet` (`facade/source.rs:61-157`); `document::Snapshot::from_bytes[_with]` (`document.rs:190,199`).
  Facade exposes names/count/text/tolerance/metadata only.
- **Edit/commit**: ✅ `Snapshot::edit()` (`document.rs:496`) → `commit()`/`commit_with_security` (`:2042,2052`) →
  `Commit{snapshot, patch, diagnostics}` (`:2594-2714`); `Patch` inverse/join/three_way/JSON, `History`. Ops: 9
  scoped closures (worksheets, definitions, annotations, validations, rdf, protection, data pilot, tracked
  changes, charts) + cells (rich text, formula, style), rows/columns append/remove, `edit_logical_rows`,
  sheets append/remove/move/copy, styles, CF, sparklines, forms, drawings, chart objects, resources.
  Metadata, settings, table templates, DDE, scenarios, sheet metadata, database ranges live only on legacy
  `MutableSpreadsheet`/`Spreadsheet` `&mut` mutators (metadata has no Patch).
- **Save**: `Snapshot::as_bytes()`; `MutableSpreadsheet::save` = `fs::write` (not atomic, `authoring/mutable.rs:811-813`);
  `SourceCellCommit::write_to(W)`; `encrypt_to_bytes`.
- **Create**: 🟡 `Builder::new()` → empty spreadsheet, zero sheets (`authoring/builder.rs:36,364`).
- **Gap class**: (a) facade get/set/batch/save/create; (b) fold legacy owners into `Edit`, swap/move-cell,
  shape/chart removal, structured diagnostics, signing; (c) column-aware structural edits on producer files
  (`move_sheet`, `edit_logical_rows`, full-table rewrite refuse typical LibreOffice output — *code-path
  inference*), rename with reference rewriting, encrypted round trip, protected-structure unlock.
- **Security**: `SignatureWritePolicy{Refuse, StripInvalidated}`, `EncryptionWritePolicy` (DecryptAndReencrypt
  declared but always refused, `document.rs:600-604,3164-3166`); legacy mutators silently strip signatures
  (*code-path inference*).

## ODP / ODG (`crates/litchi-odp`, `crates/litchi-odg`)

Prefixes: `O/` = `crates/litchi-odp/src/`, `G/` = `crates/litchi-odg/src/`. Note: `crates/litchi-odraw` is MS
OfficeArt (legacy Office), not ODF (`crates/litchi-odraw/src/lib.rs:1-4`); neither crate depends on it.

**ODP.**
- **Open**: eager `Presentation::open/from_bytes/from_prepared_package/*_with_password` (`O/package/presentation.rs:85-186`);
  lazy read-only `SourceBackedPresentation` ("no edit or output methods", `O/package/source.rs:51-181`) and
  `SourceBackedPresentationCatalog` (`O/package/catalog.rs:108-396`); `edit::Snapshot::open/from_bytes` (128 MiB cap,
  `O/authoring/edit.rs:241-254`). Facade: `OdpSource` via `from_read_at` (`F/detection_smart/detected.rs:2915,3137`),
  eager via `from_prepared_package`/`from_bytes` (`F/presentation/prs.rs:762,965`); read only; slide catalog `Unsupported` (`prs.rs:1149-1156`).
- **Edit/commit**: `Snapshot::transaction()` (`O/authoring/edit.rs:397`) → `commit()` `:2314` → `Commit{changed, snapshot, patch}`
  (`:3112-3158`), no Diagnostics. `Patch` = whole before/after: `apply` (exact source), `inverse`, `join`/`three_way`
  planning, durable bytes (`:3168-3368`), `History` (`:3407-3565`); no sequential compose. Families: slides, shapes, media,
  named text boxes/tables/form controls, RDF, charts (incl. series CRUD), layouts/master pages, annotations, layers
  (`:688-2297`). **Restriction**: `replace`/`add_shape`/`remove_shape`/`clear` refuse slides retained from the source
  (`check_slide_rewrite`, `:3069-3075`; `O/authoring/mutable.rs:1511-1538`) — title/body/shape edits only work on slides added
  in the same transaction. Structural edits refused with header/footer declarations (`:3011-3018`); `move_slide` refused
  with declarations/settings (`mutable.rs:420-426`). Separate mechanisms: `charts::Snapshot::edit` (has
  `charts::Diagnostics`, `O/charts/snapshot.rs:111-413`), flat, tail-append, and **in-place `&mut self` annotation/RDF
  methods on `Presentation`** (`O/package/presentation.rs:369-418,629-709`) that bypass ADR 0003.
- **Save/create**: `Presentation::save` = `fs::write` (`O/package/presentation.rs:751-755`, not atomic); `to_bytes`;
  `Builder` (`O/authoring/builder.rs:125-1021`: slides, images, media, layouts, metadata, transitions); streaming to `Write`
  (`O/streaming.rs:549,570`).
- **Gaps**: (a) facade edit/save/create, password open; (b) metadata set via `MetaXmlPatch`, atomic save,
  `Snapshot::save`, Diagnostics, sequential compose, paragraph insert/remove + run setters on `TextBoxModel`, addressing
  unnamed placeholder frames (inventory only covers named `draw:text-box` frames, `O/content.rs:1234-1249`), slide rename,
  move in-place annotation/RDF onto snapshots; (c) lossless rewrite of retained slides (blocks title/body/shape edits,
  picture insert, shape move/swap), structural edits with declarations/settings, general slide duplicate, automatic style
  editing, notes/transitions in the transaction.
- **Security**: `SecurityPolicy::{Editable, EncryptedReadOnly, SignedReadOnly, ..}` (`O/authoring/edit.rs:65-82`);
  transactions refuse signed/encrypted (`:3722-3735`). **In-place annotation/RDF mutators silently strip signatures**
  via `rebuild_package` (`O/annotation/package.rs:5-8`, *code-path inference*).

**ODG.**
- **Open**: eager only, `Drawing::open/from_bytes/open_template/*_with_password` (`G/facade/mod.rs:18-75`,
  `G/package/snapshot.rs:398-463`); flat `FlatDrawing`. **Not reachable from the facade** (no Cargo dependency; only
  `crates/litchi-odf/src/lib.rs:33-34` re-exports it).
- **Edit/commit**: `Snapshot::edit()`/`edit_with_policies` (`G/package/snapshot.rs:714-741`) → `commit()` → `Commit{changed,
  snapshot, patch}` (`:3393-4007`), no Diagnostics. `Patch` has typed `Change` list, `then` (sequential compose, `:4128`),
  `inverse` `:4148`, `durable` `:4168`; `DurablePatch` JSON + fingerprint; `JoinedEdits`; core `History`. Families: shape
  text/name/layer/geometry/transform/points/style/path/control, groups, pages insert/add/remove/rename/style/transition,
  shapes insert/remove/transfer, styles, forms, layers, resources (`:1231-3229`). `set_shape_text` needs one plain span (`:1612-1614`).
- **Save/create**: bytes only (`G/facade/mod.rs:121,305`) — no save. `Builder` takes raw `content_xml` only (no meta/styles);
  `build_encrypted`/`build_signed` exist (`G/authoring/builder.rs:19-110`).
- **Gaps**: (a) facade dependency/feature/variant; (b) atomic save + `write_to`, metadata, Builder meta/styles,
  text-style-name and z-index setters, page/shape move, picture-frame insert (serializer emits no `draw:image`,
  `G/package/snapshot.rs:10464-10560`), typed hyperlinks, Diagnostics; (c) rich multi-paragraph text, annotations, tables/charts.
- **Security**: `SecurityStatus`, `SecurityCapabilities`, `SecurityWritePolicy::{Refuse, RemoveSignatures}` (explicit
  strip), encrypted always refused on commit (`G/package/snapshot.rs:125-302,9475-9483`); signature verification available.

## ODC, ODI, ODB, OTH, ODM, ODF formula

None is opened by the facade (see ODF shared layer). None depends on `litchi-odt`; all use
`litchi-odf-common`. All have ADR 0003-style edit/commit/reversible patch; **none has an atomic path save**.

| Crate | Open | Edit → Commit | Save | Create | Security | Gap |
|---|---|---|---|---|---|---|
| odc | `Chart::open/from_bytes[_with_limits]` (`crates/litchi-odc/src/facade/mod.rs:59-89`); `FlatChart` | `Chart::edit()` `:240` → `Commit{chart, patch}` `:879-905`; axes/series/points/style/cached table/resources CRUD; patch inverse/compose/join/transfer/bytes | bytes only (`:188,254`) | `Builder` (`authoring/builder.rs:26-95`), `Chart::from_definition` | bools; rebuild refuses signed/encrypted (`package/snapshot.rs:207-215`) | a; b: path save, metadata & title setters, move/swap |
| odi | `Image::open/from_bytes[_with_password]` (`crates/litchi-odi/src/facade/mod.rs:24-47`) | `Image::edit()` `:208` → `Commit{image, patch}` `:783-823`; frame props, image map, metadata, resources | bytes only | `Builder` | `ProtectionDisposition`; rebuild refuses signed/encrypted | a; b: path save, patch compose/bytes, limits |
| odb | `Database::open/from_bytes[_with_password]` (`crates/litchi-odb/src/facade/mod.rs:24-55`) | `edit()`/`edit_with_policy` `:373,379` → `Commit{database, patch}`; tables/columns/keys/queries/connection/settings/components | bytes only | raw `content_xml` `Builder` | `EditPolicy{SignaturePolicy, EncryptionPolicy}` (`authoring/policy.rs:6-28`) | a; b: path save, metadata, typed builder; c: data rows, forms |
| oth | `Template::open/from_bytes/from_shared_bytes` (`crates/litchi-oth/src/facade/mod.rs:201-220`) | `edit()` `:402` → `Commit{template, patch}` `:2908-2959`; paragraph/heading text set, inline, append blocks, lists, metadata, styles, forms, resources | bytes only | rich `Builder` | rebuild refuses signed; no password open | a; b: paragraph remove/move/insert-at, table edits, path save |
| odm | `open/from_bytes[_with_password]` (`crates/litchi-odm/src/facade/mod.rs:79-110`) | `edit()` `:225` → `Commit{snapshot, patch}` (`transaction.rs:1038-1069`); title, links, metadata, sections, body items, styles, resources | bytes only | `Builder` + `BodyItemSpec` | `ChangedWriteDisposition`; refuses signed/encrypted | a; b: text set on existing items, move/swap, insert-at, path save |
| odf formula | `Formula::open/from_reader/from_bytes[_with_limits]` (`crates/litchi-odf-formula/src/facade/mod.rs:223-273`) | `edit()` `:328` → `Commit{formula, patch, diagnostics}` `:1066-1143`; MathML tree insert/remove/replace/attrs/text | `save` = non-atomic `fs::write` (`:487-490`) | `Formula::create`, `Builder` | **changed commit on signed package silently strips signatures** (via `rebuild_package`); `set_root(&mut self)` breaks immutability (`:342-347`) | a; b: atomic save, signature refusal, remove `&mut` API |

## Pages / Keynote / Numbers (iWork — do not modify; requests in `IWORK_API_REQUESTS.md`)

All three expose `Package::{open, from_bytes, write_to, save}` with an **atomic** sibling-temp save
(`crates/litchi-pages/src/package.rs:536,558,656,713`; `crates/litchi-keynote/src/package.rs:681,713,941,997`;
`crates/litchi-numbers/src/package.rs:593,624,763,823`) and many scoped `edit_*`/`apply_*` owners, each with
its own `*Commit` and patch `inverse()`. The facade already holds the `Package` for each
(`F/document/types.rs:51`, `F/presentation/types.rs:47`, `F/sheet/workbook_types.rs:18`). Encrypted iWork input
maps to `InvalidFormat` (e.g. `crates/litchi-numbers/src/document.rs:1282`). No create-from-scratch: only
detached model `Builder`s (`crates/litchi-pages/src/section.rs:565`, `crates/litchi-keynote/src/slide.rs:281`,
`crates/litchi-numbers/src/sheet.rs:255`) that do not produce a package.

- **Pages**: section/body text `edit_section_text`/`edit_body_text` with `replace/insert/delete/set/clear`
  (`crates/litchi-pages/src/package/section_text.rs:227-312,650-682`); footnote text, page layout, section
  name/pagination/settings/background, body tables (title, headers, merges, dimensions, lock, sort, hidden
  axes, deletion `body_table_deletion.rs:527`), charts (data/arrangement), image adjustments, drawable order.
  Missing: text formatting, paragraph-level structure, section add/remove, create.
- **Keynote**: `Package::edit()` slide skip/include (`crates/litchi-keynote/src/package.rs:914`,
  `package/edit.rs:99-156`); slide title/body text replace/insert/delete/set (`package/slide_text.rs:337-442,756-866`);
  slide order, `remove_slide` (`package/slide_delete.rs:635`), notes, background, transitions, build order,
  placeholders, tables, charts, comments, media creation/duplication. Missing: slide add/duplicate,
  shape add/remove/geometry, free text-box text, text formatting, create.
- **Numbers**: `edit_table_cells` / `set_table_cell[_a1]` / `clear_table_cell` with
  `Input::{Text, Number, Boolean, Date, Duration, Formula}` (`crates/litchi-numbers/src/package/table_cells.rs:233-277`,
  `table/cells.rs:54`); display formats (number/percentage/currency/scientific/fraction/text/date/duration/
  custom/controls), sheet order `move_sheet` (`package/sheet_order.rs:202`), table relocation/move, names,
  comments/replies, table appearance/dimension/lock/sort/headers/merges, charts. `RowDeletion`/`ColumnDeletion`
  types exist (`table/topology.rs:69`) but no public package edit entry was found (*unverified*). Missing: sheet/
  table add/remove/rename, row/column insert/delete, cell text style, create.
- **Gap class**: (a) facade wiring of existing owners; everything else is an iWork-owner request.

---

# Ranked recommendation: write-handler streams

Effort: S ≈ ≤1 week, M ≈ 1–3 weeks, L ≈ 1–2 months, XL ≈ > 2 months (one engineer, rough).

## Tier 1 — OOXML (incl. xlsb)

| Rank | Stream | Wiring-only part (a) | Leaf work needed | Effort |
|---|---|---|---|---|
| 1 | **xlsx** | Owned `Workbook::edit/commit/apply/save/create` covers cells, sheets lifecycle, rename/hide, layout, merges, hyperlinks | (b) style allocation, `into_package` (props/comments), per-name defined names, sheet duplicate, views, Diagnostics, source→owned conversion; (c) row/col shift, pictures/charts/tables | a: M; b: M; c: L |
| 2 | **docx** | `edit_document` text set/insert, core/custom props, save/create, section layout (source-backed) | (b) public `remove_paragraph`, move op, comments on opened docs (+ verify suspected overwrite), hyperlink target, props patch, app.xml; (c) formatting on existing runs/paragraphs/cells, table rows/cols, styles CRUD, picture replace | a: M; b: M; c: L–XL |
| 3 | **pptx** | `opened::Transaction` slides move/reorder, whole-shape text, add/remove shapes/pictures/tables/charts, comments, notes text, masters, custom props, save/create | (b) typed wrappers for sections/core props, hyperlinks, hide/rename, bounds/z-order, table cell text, picture replace, slide add, chart edit wiring, Diagnostics, explicit signature policy (stop silent unsign); (c) txBody paragraph/run editing + formatting, closure-aware slide removal | a: M; b: M–L; c: L |
| 4 | **xlsb** | Cells (same family), rename, merges, image insert, threaded comments, `Package::save`, `WorkbookWriter` create | (b) workbook-level save/patch publication, typed props, XF authoring for existing cells, uniform commit shape, explicit signature policy; (c) sheet add/remove/reorder, hyperlink/comment/table/layout/name edits on opened files | a: S–M; b: M; c: L |

## Tier 2 — legacy OLE (doc/xls/ppt)

All three need M0a to **retain an editable owner** (bytes or source) instead of dropping the package, a
facade atomic save over `Vec<u8>`, and password passthrough. Each has a strong `Writer` for `create`
(but doc/ppt/xls lose metadata — chain the property-set txn). Edits of opened files are content-level
only; structural verbs (add/remove paragraph, shape, sheet) need new engines.

| Rank | Stream | Wiring-only part (a) | Leaf work needed | Effort |
|---|---|---|---|---|
| 5 | **xls** | `cell_values` (set/insert/remove, rows/cols insert/delete, rename), visibility, comments text, property sets, `Writer` create | (b) `replace_cell`, save on Commit, typed XF builder, patches for chart/pivot editors, Writer metadata + sheet ops; (c) worksheet add/delete, comments/hyperlinks/merges/names in opened files, pictures | a: M; b: M; c: L |
| 6 | **ppt** | `slide_order` (slide move/remove/transfer, hidden, transition, whole-shape text, anchors), property sets, `Writer` create | (b) Writer metadata, unified batch root over single-feature owners, record-owner adapters, patches for animation/OLE/chart; (c) rich text, inserting shapes/pictures/tables/blank slides | a: M; b: M–L; c: XL |
| 7 | **doc** | `body_text` text replace + b/i/u, tracked revisions, embedded objects, property sets, `Writer` create | (b) Writer metadata, path save for `finish()` owners, cross-family join, untracked insert-at-CP; (c) structural body editor (paragraph/table/section insert/delete/move), range formatting, comments/hyperlinks/bookmarks CRUD | a: M; b: M; c: XL |

## Tier 3 — delayed (RTF, ODF, iWork)

| Stream | Summary | Effort |
|---|---|---|
| ods | Strongest ODF write root (`document::Edit` with Diagnostics); (a) facade; (b) fold legacy owners, atomic save; (c) producer-file structural edits | a: M; b: M; c: L |
| odt | (a) facade; (b) atomic save, metadata/comments in `Edit`; (c) rich inline/table/block editing; avoid lossy `MutableDocument` | a: M; b: M; c: L |
| odp | (a) facade write; (b) metadata, atomic save, text-box paragraph ops; (c) retained-slide lossless rewrite (largest ODF blocker) | a: M; b: M; c: L |
| odg | (a) facade must first open it; (b) save, metadata, move, picture insert | a: S–M; b: M; c: L |
| odc/odi/odb/oth/odm/odf | Facade must first open them at all (new `File` variants + Cargo deps); leaf edit roots exist; (b) atomic save each (or shared writer); odf-formula signature-strip fix | a: M (all six); b: S each |
| rtf | (a) keep editable `Document`, atomic save; (b) setters/swap; (c) semantic create builder, structure CRUD | a: S; b: S–M; c: M–L |
| pages / keynote / numbers | (a) wire existing owners (text, cells, slide order/delete, formats); create + structural verbs blocked on iWork-owner APIs (see `IWORK_API_REQUESTS.md`) | a: M each; rest: external |

**Cross-cutting prerequisites (do first, all tiers):** (1) facade atomic path writer over bytes/streams;
(2) one explicit signature/encryption consumption policy surfaced as `Locked`/`Signed` `File` states, with
leaf fixes where commits silently strip (pptx, xlsb, odf-formula, ods legacy); (3) commit-shape adapter
normalizing `(snapshot, patch, diagnostics)` across leaves (synthesize empty diagnostics where absent);
(4) M0a retention of editable owners for doc/ppt/xls/rtf and owned (not source-backed) models for xlsx/xlsb/pptx write sessions.

## What is not claimed

- No behaviour was executed; "silent strip", "suspected comment overwrite" (docx), and "refused on typical
  producer files" (ods) are code-path inferences that need tests before being treated as bugs.
- Effort figures are rough planning estimates, not measurements.
- FEATURE_MATRIX disagreements noted inline (pptx hyperlinks ✅, docx app.xml ✅, doc XML-signature editing,
  xls signature editing, ppt password-change workflows) were not resolved here.
