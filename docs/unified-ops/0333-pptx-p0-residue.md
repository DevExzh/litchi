# 0333 — PPTX P0 residue (audit WP3)

- Status: implemented on `uo2/pptx-p0`, cut from `feat/unified-ops-api` at
  `d761892de3`; trunk merged at `84ef2cd65e` (records 0327, 0328, 0329,
  0332) and again at `92d622b38f` (record 0335).
- Date: 2026-10-17
- Input: the definition-of-done audit of 2026-10-17 (§2, WP3): in design §13's
  P0 table, PPTX "Basic CRUD verbs" and "Comments" were partly met
  (paragraph/run/`tr`/`tc` had no add or remove, comment remove was open,
  picture `set` covered the click action only, formatting batches were not
  coalesced or measured); record 0330 follow-ups F1 (`tc fill=accent1`) and
  F2; record 0252 §6.
- Owned: `crates/litchi-pptx`, `crates/litchi-ops/src/format/pptx*`, the PPTX
  rows of the registry, checklist and fidelity tests. Small hunks outside
  them: none in the shared dispatch files (every new verb is reached through
  `format::pptx`); `tools/officecli_dispositions.json` (the PPTX rows);
  the regenerated reports; one Python test.
- `performance_claim`: §6.

## 1. What landed

| P0 item | get | set | add | remove |
|---|---|---|---|---|
| paragraph (shape, placeholder) | ✅ | ✅ | ✅ start, end, position (past the end appends), before/after a sibling of the same body; `text` (a line feed is a line break), `align`, `bold`, `italic`, `size`, `color`, `font`, `underline`, `strike` | ✅; a body keeps one paragraph |
| paragraph of a table cell, `/slide[N]/table[M]/tr[R]/tc[C]/paragraph[K]` | ✅ new (path only) | ✅ new | ✅ new | ✅ new |
| paragraph of the notes, `/slide[N]/notes/paragraph[K]` | ✅ new (path only) | ✅ new | ✅ new | ✅ new |
| run (and line break) of any of those bodies | ✅ | ✅ | ✅ run: `text`, `bold`, `italic`, `size`, `color`, `font`, `font.latin`, `underline`, `strike` | ✅ run and line break |
| table | ✅ | — | ✅ new: `rows`, `cols`, `data` (`a,b;c,d`), `x`, `y`, `width`, `height`, `name` | ✅ |
| `tr` | ✅ `height` (new read) | ✅ `height` | ✅ `height`, `cols` check, `c1`…`c20` seed texts; start, end, position, before/after | ✅ |
| `col` (0321) | ✅ | ✅ | ✅ merged tables now accepted where safe | ✅ merged tables now accepted where safe |
| `tc` | ✅ `fill` (RGB or theme), `valign`, `padding.left/right/top/bottom` (new reads) | ✅ `fill` takes theme colours and `none` (0330 F1); `valign`; `padding` and the four sides; `underline`, `strike`; OfficeCLI's aliases (`background`, `shd`, `shading`, `margin`, `margin.bottom`, `font.*`, `halign`, …) | operation gap (as 0330 decided for DOCX) | operation gap |
| shape, placeholder, paragraph, run | | ✅ `underline` and `strike` (new) | ✅ on text boxes | |
| comment | ✅ | ✅ | ✅ | ✅ new |
| modernComment, reply | ✅ | — | — | ✅ new |
| picture | ✅ | ✅ new: `x`, `y`, `width`, `height`, `alt` (`altText`, `description`), `src` (image replacement), with `link`/`tooltip` | ✅ `alt` new | ✅ |

Formatting batches: an atomic batch or `edit` closure of direct run,
paragraph or shape-body formatting sets is one native
`set_text_property_chunks` call per slide and one leaf commit; §6
measures it.

### Leaf (`litchi-pptx`)

| Name | What it does |
|---|---|
| `opened::TextBody::{Shape, Cell, Notes}`; `opened::Edit::{insert_paragraph, remove_paragraph, insert_run, remove_run}`; the same four on `source::Edit` (slide bodies) | byte splices of one `a:p` or run child; new markup in the body's own prefix carries only the assigned formatting (empty text: an `a:endParaRPr` with it, as `PowerPoint` sizes an empty line); the part is read back as a scene with unchanged shape identities |
| `opened::TextTarget::CellParagraph` | formatting of one paragraph of a cell body |
| `opened::{CellFill, CellMargin}`; `TextProperties::{with_fill(CellFill), with_cell_anchor, with_cell_margin, with_underline, with_strike}` | `a:tcPr` solid RGB, `a:schemeClr` or `a:noFill`; `@anchor`; `@marL/marR/marT/marB`; `a:rPr@u`, `@strike` |
| `table::{CellProperties, Cell::properties, Row::height, Table::row_cell_counts}` | reads of the above; the row cell counts in one pass |
| `opened::Edit::{insert_table_row, remove_table_row, set_table_row_height}` | row edits through the 0321 grid editor, now tracking `gridSpan`/`rowSpan`/`hMerge`/`vMerge` |
| `opened::Edit::set_text_property_chunks` (eager, `source::Edit`, `SourceBackedSlideEdit`) | many 256-edit chunks in one pass |
| `opened::Edit::{remove_comment_at, remove_modern_comment_at, remove_modern_comment_reply_at}` | positional siblings of `set_comment_at` over the identity verbs |
| `opened::Edit::{set_picture_image, set_shape_description, set_shape_name}` | `a:blip@r:embed` retargeted at a new image part (the old relationship and part go when nothing else uses them); `p:cNvPr@descr` and `@name` |

The geometry editor admits `p:pic` frames, so `set_shape_properties` sets a
picture's bounds.

**Text-property staging is single-pass.** `stage_text_properties` read the
whole slide as a scene and copied the part once per edit, so a batch of n
edits on a slide of n shapes was O(n²); a 1,000-run debug test took 192 s. It
now reads the scene once, rewrites each edited element (leaf elements, so
disjoint; a nested target is refused) and splices them once. Outputs are
byte-identical (§5). Every existing caller (unified run, paragraph, shape and
cell formatting, find/replace, template merge, field refresh) gains from it.

### Merge consistency (columns and rows)

Every row of a DrawingML table holds one `a:tc` per grid column, merged ones
included, so a cell's grid column is its position in the row; tables whose
rows do not are refused for structural edits when they have merges.

| Edit | Rule |
|---|---|
| insert a column at `p` | refused when a row's cell at `p` is `hMerge` (the boundary cuts a merge); otherwise new plain cells |
| remove a column | a covered (`hMerge`) cell: its merge's first cell spans one column less; a merge's first cell: the next cell drops `hMerge` and takes `gridSpan − 1`; a cell in a merge spanning rows and columns: refused |
| insert a row at `p` | refused when the row at `p` holds a `vMerge` cell; otherwise one plain cell per grid column |
| remove a row | a covered (`vMerge`) cell: the first cell above spans one row less; a merge's first cell: the cells below drop `vMerge`, the first takes `rowSpan − 1` (and the `gridSpan`); the last row is kept |

The frame's `cx`/`cy` follow the column widths and row heights, as for
0321's columns.

## 2. Decisions and deltas from OfficeCLI

| # | OfficeCLI | Litchi |
|---|---|---|
| 1 | removing a comment leaves its authors; the comments part goes with the last comment | the same; the author list and each author's `lastIdx` are kept (record 0319's rule: no edit of content the caller did not name), so a later comment takes a fresh index |
| 2 | a text body can lose its last paragraph | refused (`unsupported_capability`): `CT_TextBody` requires one |
| 3 | a new run or paragraph gets `lang="en-US"` | no language is invented: the run carries only the assigned formatting |
| 4 | `add paragraph` at an index returns `paragraph[count]` | the paragraph's actual position |
| 5 | header/footer-style story children are listed | cell and notes paragraphs are addressed by path only: views and unscoped queries keep their shape (record 0325's rule for DOCX stories) |
| 6 | `fill=none` reads back as `none` | an explicit no-fill is written (`a:noFill`) and reads as no fill colour, so the `fill` value stays colour-typed |
| 7 | `padding` and `padding.bottom` in one call: the later wins | refused (`invalid_argument`): one call assigns each inset once |
| 8 | picture `src` replacement derives a new `alt` from the file name | `alt` changes only when assigned |
| 9 | a table without `name` is `Table N` by its table count | the same |
| 10 | `crop*` on pictures | declared property gaps `0333-F4`: the picture reader does not model `a:srcRect` |

## 3. Registry, checklist and fidelity

Goldens re-blessed in both feature sets (`--all-features`, and every feature
but `encryption`); record 0330's `element_operations_agree_with_property_flags`
stays green. PPTX element claims: `paragraph`, `run`: `gq`+`s` → `asgqr`;
`linebreak`: `gq` → `gqr`; `tr`: `gq` → `asgqr`; `table`: `gqr` → `agqr`;
`comment`: `asgq` → `asgqr`; `modernComment`, `reply`: `gq` → `gqr`;
`picture` set covers the frame, `alt` and `src`. New property gaps:
`picture.crop`, `cropLeft`, `cropTop`, `cropRight`, `cropBottom`
(0333-F4); the gaps `tc/paragraph`, `notes/paragraph` and `tr height` are
closed.

CRUD checklist (`results/0289/crud-checklist`): PPTX claimed cells
279 → **293**, all ok, 0 FAIL (new: `table` add; `tr` set, add, remove,
edit, batch; `paragraph`, `run` add and remove; `linebreak`, `comment`,
`modernComment`, `reply` remove). The generator found fixtures for every new
cell without a recipe.

Fidelity (`results/0289/fidelity`): every PPTX file stays `exact` on
open → empty edit → save. One set-then-revert row moved:
`tdf169496_hidden_graphic.pptx` now reverts the picture `alt` (newly
settable) value-exact rather than a layout name byte-exact: the source spells
a character of its description raw that the writer escapes.

## 4. OfficeCLI parity

Delta table (`officecli-capability-delta.md` and its `pptx` page), trunk
`84ef2cd65e` → this record:

| pptx | supported | partial | gap | write-only | litchi-only |
|---|---:|---:|---:|---:|---:|
| property reads | 228 → 234 | | | 123 → 121 | 269 → 291 |
| property writes | 112 → 142 | 34 → 23 | 415 → 396 | | |

Element rows are unchanged (29 supported, 3 gaps). Dispositions
(`tools/officecli_dispositions.json`): the stale keys `paragraph.text` and
`run.text` of `pptx-p0-residue`, the whole `pptx-structure-adds` entry and
the `tc` value-vocabulary key are deleted; `pptx-p0-residue` now names this
record's follow-ups F1–F3 (the P0 rows left: comment `direction`, modern
comment writes, placeholder and slide text, slide layout and title); the
picture and `table.style` rows are P1 property breadth (the policy's
`propertyBreadthTier`) and are `deferred-P1` (`pptx-picture-table-breadth`);
table-level `set` is `planned 0333` (F5, `pptx-table-set`).

Corpus (`LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
officecli_corpus`, `tools/officecli_corpus.py`, `--check`):

| family | passed | warned | gap | delta | cascade |
|---|---|---|---|---|---|
| ppt | 129 → 129 | 37 → 37 | 141 → 141 | 2 | 1 |
| ppt/shapes | 102 → 106 | 92 → 94 | 26 → 20 | 0 | 0 |
| ppt/tables | 95 → **358** | 7 → 90 | 65 → 9 | 0 | 290 → **0** |
| ppt/textboxes | 50 → **90** | 17 → 40 | 65 → 2 | 0 | 0 |
| pptx total | 1286 → **1593** | 716 → 825 | 410 → 284 | 146 | 401 → 111 |
| all formats | 3948 → 4255 | 1178 → 1287 | 663 → 537 | 201 | 441 → 151 |

Bugs: 0. The warned calls that grew are the table and run lines that now run
and skip their remaining P1 breadth (table `style`, borders and banding,
run `baseline`/`spacing`), each with a disposition.

**v1-scope subset** (design §14): 3,443 lines; before: 1,877 passed, 941
accounted, **625 open** (432 blocked by WP3); after: 2,184 passed, 1,062
accounted, **197 open** (1 blocked by this record: `tables-nested.sh:86`,
a table-level `set`, F5). PPTX v1 families: tables 354 → 10 open, textboxes
65 → 2, pictures 17 → 3, shapes 8 → 2.

## 5. Byte identity

A probe of the public `File` API (`results/0333/byte-identity-probe.rs.txt`,
built against the base tree `git archive d761892de3` and against this
record's tree; binaries' SHA-256 in `results/0333/binaries.sha256`) runs, on
every one of the 83 PPTX decks of `test-data/` with both owners, a scoped
query fan-out (`run` under `/slide[1]`: bold and size), a shape-body set, a
paragraph alignment and a 16-item best-effort batch of run colours. The 166
result rows (byte digests or error codes) are identical
(`results/0333/byte-identity-diff.txt`); the batch went from per-item sets
to the coalesced run, the fan-out from per-chunk calls to the chunk verb.
Capability and registry goldens changed only as §3 lists.

## 6. Measurements

`crates/litchi-ops/examples/pptx_p0_bench.rs` pairs each unified call with
the leaf call doing the same work on the same owner, ABBA (leaf, unified,
unified, leaf), checking outside the clock that both publish the same bytes
(they do in every scenario). Release build (`cargo build --release -p
litchi-ops --example pptx_p0_bench --features pptx` at `067cd3e8fb`; SHA-256
`22286511…5b28`), `taskset -c 7`, load 6–8, two runs
(`results/0333/pptx-p0-bench-run{1,2}.jsonl`). The batch leaf leg is one edit
staging the N `TextEdit`s in 256-edit chunks through one
`set_text_property_chunks` call and one commit.

| scenario | N | eager leaf / unified µs | ratio (run 1 / 2) | source leaf / unified µs | ratio (run 1 / 2) |
|---|---:|---|---|---|---|
| run format batch | 1 | 282 / 304 | 1.077 / 1.079 | 47 / 51 | 1.084 / 1.091 |
| | 10 | 398 / 448 | 1.126 / 1.121 | 243 / 270 | 1.111 / 1.113 |
| | 100 | 1,480 / 1,761 | 1.190 / 1.183 | 2,126 / 2,399 | 1.128 / 1.134 |
| | **1,000** | 11,961 / 14,580 | **1.219 / 1.220** | 20,963 / 23,695 | **1.130 / 1.139** |
| | 10,000 | 117,627 / 145,135 | 1.234 / 1.230 | 210,397 / 239,447 | 1.138 / 1.154 |
| shape text batch | 1,000 | 9,471 / 9,926 | 1.048 / 1.047 | 19,217 / 19,635 | 1.022 / 1.029 |
| | 10,000 | 93,906 / 99,233 | 1.057 / 1.051 | 192,664 / 197,648 | 1.026 / 1.033 |

The 1k-op batch budget (within 2× of one leaf edit) is met on both owners,
and both legs scale linearly from 1 to 10,000. Before this record the same
unified batch ran one native commit per item (`results/0333/batch-before-after.txt`,
`batch-time.rs.txt`): 100 items 195 ms → 2.0 ms (eager), 300 items
1,519 ms → 5.3 ms, and 1,000 items were refused (`limit_exceeded`, the
snapshot bytes of 1,000 separate steps) → 16.6 ms.

Single operations, eager owner, a one-slide deck of 20 text boxes, a 3 × 3
table, a picture and two comments (budget 10%):

| scenario | leaf µs | unified µs | ratio (run 1 / 2) |
|---|---:|---:|---|
| `paragraph_add` | 504 | 510 | 1.012 / 1.015 |
| `run_add` | 502 | 510 | 1.016 / 1.015 |
| `paragraph_remove` | 501 | 508 | 1.015 / 1.021 |
| `row_add` | 814 | 894 | 1.099 / 1.092 |
| `row_remove` | 745 | 784 | 1.053 / 1.048 |
| `cell_fill` (`accent1`) | 516 | 556 | 1.077 / 1.077 |
| `comment_remove` | 509 | 523 | 1.027 / 1.028 |
| `picture_frame_set` | 666 | 672 | 1.009 / 1.007 |

Fixes made to reach these numbers (D5: fix before explaining):

- the first eager one-item batches were 1.45–1.53×: a transaction's admission
  built and dropped a whole `Package::edit` (a package capture with its
  fingerprint) before the run built the real one. The eager owner now caches
  that admission per handle (it depends on the immutable package only);
- `cell_fill` was 1.20× and `row_remove` 1.12×: resolving `tr[R]/tc[C]`
  scanned the table once per row with a fresh namespace scope each time.
  Table rows and cells now resolve through one cached grid per frame
  (`Table::row_cell_counts`, one pass); `row_add` reads the column grid only
  when a key needs it;
- the 10,000-op batches were superlinear in the leaf (1.9 s eager): 40
  native calls each read the whole slide. `set_text_property_chunks` stages
  every chunk in one pass (118 ms).

The eager run-format batch keeps a 12–23% overhead at 10–10,000 items (about
2.7 µs per item): planning parses each target shape's text body to resolve
the run and judge its change. It is within the 2× budget; a per-handle body
cache would trim it (follow-up F6).

## 7. Gates

Run on the merged tree (`92d622b38f` merged, the record commits on top),
once before the second merge and again, all green, after it;
cargo through the memory-capped shim, `CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`; every exit status checked
(`(deleted-coord-artifact)`).

| gate | exit |
|---|---:|
| `cargo fmt --all --check` | 0 (after a rustfmt of `pptx_table_rows.rs`) |
| `cargo clippy -p litchi-pptx -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --features <all but encryption> --all-targets -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features pptx] --all-targets -- -D warnings` | 0, 0 |
| `cargo test -p litchi-pptx --all-features` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib` (929), `--doc`, 8 `--test` chunks of 20 (executables deleted per chunk) | 0 (the first run failed `spec_tables_agree_with_their_vocabularies`: a `para` paragraph alias the shared path vocabulary does not know; the alias was dropped, the goldens re-blessed and the lib and the affected chunk rerun) |
| `cargo test -p litchi --all-features`; `cargo test -p litchi-cli --all-features` | 0; 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry` (all features; all but `encryption`) | 0, 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pptx -p litchi-ops` | 0 |
| `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify` | 0; 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (40 profiles, 2,313 bans, 253 families) | 0 |
| `tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI` and `--check` | 0, 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus`; `tools/officecli_corpus.py`; `--check` | 0, 0, 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta`; `tools/test_no_3rdparty_references.py`; `tools/unified_bench_matrix.py --check` | 0, 0, 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (286 tests, new `test_record_0333_p0_residue`) | 0 |

## 8. Follow-ups

- F1: comment `direction`; modern (threaded) comment `add` and `set`
  (record 0252 §6); notes text is a set, not an `add`.
- F2: slide `title` and `text` writes, slide `layout` set, placeholder text
  on `add`; run `link`/`tooltip` on `add`; `linebreak` `add`.
- F3: formatting coalescing of cell and notes text (they take the ordinary
  writer today) and of paragraph `text`.
- F4: picture `crop*` (the reader models no `a:srcRect`), `rotation`,
  flips and effects; table `style`.
- F5: table-level `set` (`firstRow`, banding, borders, `style`); cell
  `colspan`/`rowspan`/`merge.*`, borders, `wrap`, `textDirection`,
  `lineSpacing`, `image` fills; row and column insertion inside a merge.
- F6: the eager run-format batch's per-item planning (§6); picture,
  comment, row and table writes still promote a source owner to an eager
  package (record 0252 §6).
