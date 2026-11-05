# 0345 — DOCX P0 residue, final: add cell, the table `col` element, cell text, hyperlink tooltip, comment replies, resolution and anchors

- Status: implemented on `uo2/docx-p0b`, cut from `feat/unified-ops-api` at
  `926cbde41e`; trunk merged at `2bdfd00598` (record 0344) and `dfaffd5e48`
  (record 0347).
- Date: 2026-10-19
- Input: the second definition-of-done audit
  (`(deleted-coord-artifact)` §2, §3, §12.1 item 1, package A):
  10 DOCX P0 rows of the OfficeCLI capability delta were still `planned`
  (`docx-p0-residue`: the `column` element, `column.width`, `column.text`,
  `cell.text`, `hyperlink.tooltip`; `docx-picture-comment-follow-ups`:
  `comment.anchoredTo`, `runStart`, `range`, `parentId`, `done`), and the
  corpus class `docx|add|cell|unsupported_element` (`docx-structure-writes`).
  Records 0335 (F2, F8) and 0342 (§8).
- Owned: `crates/litchi-docx`; `crates/litchi-ops/src/format/docx*`; the
  DOCX rows of the registry, capability, CRUD-checklist and fidelity tests
  and goldens; the DOCX keys of `tools/officecli_dispositions.json`. Small
  hunks outside them: `capabilities.rs` (the DOCX `tc`/`col` claims), the
  DOCX probe of `tests/corpus_fidelity.rs`, and the `p0345` parser of
  `tools/unified_bench_matrix.py`.
- `performance_claim`: every new write is within the 10% single-op budget
  on a warm handle (§6); one existing path got faster (plain comment add
  1.65 → 1.02).

## 1. Result

OfficeCLI capability delta, P0 `planned` rows: **26 → 16** (DOCX **10 → 0**;
the 16 left are PPTX, package B). DOCX elements supported 33 → 34, gap
11 → 10; property reads supported 307 → 311, gap 4 → 2; property writes
supported 279 → 287, gap 482 → 474; DOCX `planned` rows 43 → 34 (the P1
chart series and axis formatting keys of record 0335 F4).

| element | before | now |
|---|---|---|
| `cell` (`/body/tbl[N]/tr[M]/tc[K]`) | `sgqr`; `text` set only | `asgqr`: `add` under a row with `text` and `width` (OfficeCLI's `AddCell` honours only these two on add); `text` is written by add and set |
| `column` (`/body/tbl[N]/col[M]`, alias `col`) | element gap | `asgr`: `width` (the `w:gridCol` width, twips) read, set and added; `text` seeds the new cells; add and remove. Not a listed table child: table views and child counts stay the rows, and `query` is not claimed (OfficeCLI has no column get or query) |
| `hyperlink` | `tooltip` read | `tooltip` added and set; `null` removes it |
| `comment` | `parentId`, `done` typed gaps, refused on add and set | `anchoredTo` read; `done` (alias `resolved`) read, added and set; `parentId` read and added; `runStart` and `range` added |

The corpus line `word/revisions.sh:196` (`add cell` with
`revision.author`/`revision.date`) moved from `unsupported_element` to
`warned` (the tracked-change keys are deferred-P2 breadth).

## 2. Leaf surface (`litchi-docx`)

Every verb travels in the main-document `Patch` (element replacements, and
graph changes for the comment parts), so replay and the exact inverse
apply. The grid and tooltip verbs are on both owners (`document::Edit<P>`
and the snapshot `Edit`); the comment verbs stay on the eager
package-bound `Edit`.

| verb | what it stages |
|---|---|
| `insert_cell_at(row, at, &NewCell)` | one table replacement: the cell (`w:tcW` when a width is given, one paragraph of text) before cell `at` or after the last; when the row then spans more grid columns than `w:tblGrid` has, the grid gains columns of the grid's average width (2400 twips for an empty grid) and every other row as many empty cells (OfficeCLI BUG-R1-P0-2). A table without a grid gets none |
| `insert_column_at(table, at, width, text)` | one table replacement: a `w:gridCol` (the given width, else the average) before grid column `at` or after the last, and a cell holding `text` in every row at that slot; a slot covered by a horizontal span or a merged cell is `Refusal::ComplexContent`, as OfficeCLI's `GuardColumnSlotAddressable`; a missing grid is created after `w:tblPr` |
| `remove_column_at(table, column)` | the `w:gridCol` and each row's cell at the slot; merged or spanning cells, the last column, a row's last cell and a cell holding one end of an outside range are refused |
| `set_column_width_at(table, column, width)` | the `w:gridCol` width and, so fixed layouts follow, the `w:tcW` of each row's cell occupying the slot alone with a twips (`dxa`) width; an unchanged width is a no-op |
| `set_hyperlink_tooltip_at(hyperlink, Option<&str>)` | `w:tooltip` on the start tag, other attributes as written |
| `add_comment_at` with `NewComment::{with_parent, with_done, with_start_after, with_range}` | a reply or a comment with a state gets a fresh `w14:paraId` on its (last) paragraph (above every identifier of the main document and the comments part); the parent's last paragraph gets one if it has none; `commentsExtended` gains `w15:commentEx` entries (the thread root's when missing, then the reply's `w15:paraIdParent`/`w15:done`), created with its relationship and content type when absent; an existing `commentsIds` part maps every new identifier to the next durable id. `with_start_after(run)` starts a whole-paragraph range after that direct run; `with_range(false)` writes the reference run alone |
| `set_comment` with `CommentChange::with_done` | the comment's entry's `w15:done`, upserted lexically (other entries keep their bytes) |

Reads: `Comment::paragraph_id`, `Document::comments_extended` and
`source_backed::Package::comments_extended` (the parsed entries).
`NewCell` follows `NewTable`. Tests: `tests/comment_threads.rs`, the grid
module's unit tests (merged slots refused, average widths, grid sync) and
the doctests.

## 3. Adapter (`litchi-ops`)

- `format/docx/column.rs`: the column node (`Value::Column`: the table's
  container entry and the grid position), `pick_column`, `add_column`,
  `try_set_column` and `remove_columns`. The grid is read by scanning the
  table's start up to the end of `w:tblGrid`, so resolving a column does
  not parse the table's properties or rows (§6).
- `format/docx/nested.rs`: the `cell` add arm and `new_cell` (widths as
  OfficeCLI's `AddCell` reads them: twips with an optional `dxa`, `N%`,
  `auto`, `nil`, or a length); the hyperlink add takes `tooltip`; the
  comment add takes `parentId` (a missing parent is `invalid_value`, as
  OfficeCLI refuses it), `done`, `runStart` (1-based; 0 keeps the
  paragraph start) and `range`. The new comment's ordinal is the count of
  the comments before the edit, read on the warm handle (the leaf appends).
  The record 0335 `refuse_reply` is deleted.
- `format/docx/props.rs`: `done` on every comment (`false` without an
  entry, as OfficeCLI), `parentId` on replies (the parent comment's id),
  and `anchoredTo` read on demand: one walk of the main story (body,
  tables, content controls) finds every comment's range start, else its
  reference, and the handle caches the map, as it caches
  `commentsExtended` (`parts::Cache`). A `commentsExtended` part the leaf
  refuses as malformed reads as no entries, so the comment's other
  properties stay readable.
- Registry: `COMMENT_REGISTRY` becomes a table with the `resolved` alias;
  `CELL` `width` is add-and-read, `text` set-and-add; `HYPERLINK`
  `tooltip` set-and-add (clearable); `COLUMN` (`width`, `text`). The
  `col` element gap, the two comment property gaps, the `add cell`
  operation gap and the stale `set table/row table-properties` operation
  gap (record 0342 sets both) are deleted.
- `edit` and `batch` reach the column (the formatting run treats it as a
  barrier, as it treats charts).

## 4. Deltas from OfficeCLI

| # | OfficeCLI | Litchi |
|---|---|---|
| 1 | `column` has no get or set | `width` is read and set (litchi-only verbs over the same grid) |
| 2 | `hyperlink.tooltip` is add-only | also set and cleared |
| 3 | a new cell or column paragraph gets a `w14:paraId` | none, as Litchi's other new paragraphs; only a comment paragraph that `commentsExtended` keys gets one |
| 4 | `pointRef` (alias of `range` with the opposite sense), `rangeOpen`, `rangeEnd`, `commentParaId`, `annotationRef` | not taken (dump-replay controls); `range=false` is the point form |
| 5 | a missing grid is prepended to the table | created after `w:tblPr` (schema order) |
| 6 | a removed comment's `commentsExtended` entry is removed, the part stays | the same (record 0251c `strip_modern_entries`); the empty part and its relationship stay |
| 7 | `add cell` with `revision.*` marks the cell inserted | the revision keys are skipped with a warning (deferred-P2, tracked changes) |
| 8 | `commentsExtended` is keyed by the comment's *first* paragraph | by its *last* paragraph, as [MS-DOCX] `CT_CommentEx` specifies and Word writes; the same paragraph for one-paragraph comments (every comment either tool writes) |

## 5. Goldens, checklist, fidelity, corpus

- Capability and registry goldens re-blessed in both feature sets
  (`--all-features`, and every feature but `encryption`): `docx.json` and
  `docx-encryption.json` of each. Every hunk is a §1 or §3 row. The 0330
  consistency test stays green; `tests/capabilities.rs` asserts the new
  flags (25 + 8 operation gaps).
- CRUD checklist: `cell add` n/a → ok; the new `column` element is ok for
  get, find, view, set, add, remove, edit, batch, create and save (query
  n/a; move and swap are format verbs). 0 FAIL, 0 unreached.
- Fidelity: the DOCX structural probe also resolves its comment
  (`done=true`, which writes `commentsExtended`; the removal strips the
  entry). Every row keeps its verdict; the probe's republished sizes grow
  by the part. Open → empty edit → save stays byte-exact.
- Corpus re-recorded; one outcome changed (§1). Delta and corpus `--check`
  are green. Dispositions: `docx-p0-residue` and `docx-structure-writes`
  deleted; `comment.*` removed from `docx-picture-comment-follow-ups`; the
  follow-up maps `docx-cell-insertion` and `docx-table-properties` (record
  0344) deleted with the gaps they mapped.

## 6. Measurements

`examples/docx_p0_final_bench.rs`, release, `taskset -c 13` on a shared
32-core host (load average 5.1–6.1), ABBA, 41 rounds after 5 warm-ups,
medians, three runs; fixture `create()` with 200 paragraphs, a 4×4 table,
a hyperlink and a comment; warm eager handle. The leaf leg stages the same
`document::Edit` verb on the same snapshot and commits; outputs are
compared outside timing. File:
[results/0345/docx-p0-final-bench.txt](results/0345/docx-p0-final-bench.txt)
(binary SHA-256 in its header).

| case | leaf µs | unified µs | unified/leaf |
|---|---:|---:|---:|
| hyperlink `set tooltip` | 191.0–194.2 | 203.1–206.9 | 1.063–1.070 |
| `add cell` | 192.8–197.4 | 202.7–213.7 | 1.047–1.083 |
| `add column` | 196.4–201.6 | 200.6–211.5 | 1.016–1.049 |
| column `set width` | 190.5–194.6 | 193.2–196.9 | 1.012–1.014 |
| comment `add` with `parentId` | 223.0–227.1 | 228.4–233.2 | 1.024–1.030 |
| comment `set done` | 112.4–116.9 | 116.7–121.4 | 1.036–1.039 |
| comment `add` (existing path) | 213.3–220.4 | 219.0–225.6 | 1.024–1.027 |

Attribution legs in the same file: resolving `/body/tbl[1]/col[1]` costs
1.35–1.37 µs against 0.30–0.32 µs for `/body/tbl[1]`; a comment's view
(`get` + `format`) 6.8–6.9 µs against 3.1–3.3 µs for a paragraph's.
`tools/unified_bench_matrix.py` parses the file (`p0345`); none of its
rows is a miss.

- Every new write is within the 10% budget.
- **Fixed on the way.** The first version measured `add column` 1.13 and
  column `set width` 1.13: resolving `col[M]` parsed the table's
  properties (24 µs; now 1.3 µs, against 0.3 µs for the table path). A
  comment add read the comments of the *new* handle to find its ordinal,
  which first builds that handle's document (about 140 µs): the existing
  plain comment add measured **1.65** and a reply 2.23. Both now read the
  comments before the edit on the warm handle (3 µs). A document's *first*
  comment still looks itself up on the new handle, because its part may
  carry a Strict relationship the eager comment reader does not list
  (`strict.docx` keeps its typed refusal in the fidelity report); that
  first add pays the old cost.
- **Existing reads.** A comment's view now also reads `done`, `parentId`
  and `anchoredTo`. Measured with a probe on the public API (`get` +
  `view().format()`, 200 paragraphs, comment at `p[150]`), base
  `926cbde41e` against the branch, alternated: 6.6–6.7 µs before, 6.6–6.8
  µs after ([results/0345/comment-view.txt](results/0345/comment-view.txt));
  a per-view story walk measured 15.9 µs before the per-handle cache. The
  first comment view of a handle pays one walk of the main story.
  Other paths are untouched; existing outputs are byte-identical (the
  fidelity byte-exact rows and the corpus outcomes other than §1 are
  unchanged).

## 7. Gaps and follow-ups

| # | gap | follow-up |
|---|---|---|
| F1 | `column` move and swap (OfficeCLI moves a column within its table) | format verbs over a grid slot; P1 |
| F2 | comment `pointRef`, `rangeOpen`/`rangeEnd` multi-paragraph ranges, `commentParaId`, `annotationRef` (dump-replay controls) | with `add --from` / dump replay (corpus G2, deferred-P1) |
| F3 | `add cell` and `add column` with `revision.*` (`w:cellIns`) | tracked-change workflows (§13 P2) |
| F4 | a reply's `commentsExtensible` entry (Word 2018 dates, reactions) | not written; Word does not require it |
| F5 | an emptied `commentsExtended` part stays after its last comment is removed | a dependency sweep on publication (record 0335 F5) |

## 8. Gates

Run on the merged tree; cargo through the memory-capped shim,
`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`. Commands and exit codes are in
the branch's final report.
