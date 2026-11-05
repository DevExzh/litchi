# 0346 — PPTX P0 residue, final (second-audit package B)

- Status: implemented on `uo2/pptx-p0b`, cut from `feat/unified-ops-api` at
  `926cbde41e`; trunk merged at `d392343e02` (records 0343, 0344, 0345,
  0347).
- Date: 2026-10-19
- Input: the second definition-of-done audit (`(deleted-coord-artifact)`
  §2, §3, §12.1 item 2): 16 PPTX P0 delta rows `planned` (`pptx-p0-residue`,
  record 0333 F1–F2) and 4 P0 modern-comment reads `unlisted`; record 0333
  F3 (cell and notes text coalescing).
- Owned: `crates/litchi-pptx`, `crates/litchi-ops/src/format/pptx*`, the
  PPTX rows of the registry, checklist and fidelity tests, the pptx keys of
  `tools/officecli_dispositions.json`. Small hunks outside them: one line in
  the shared `capabilities.rs` (placeholder insertion follows shape
  insertion), one CRUD recipe, the capability test's PPTX expectations.
- `performance_claim`: §6.

## 1. Result

| Delta (`officecli-capability-delta.md`) | before (`926cbde41e`) | after |
|---|---:|---:|
| P0 gap rows `planned`, all formats | 26 | **0** (DOCX's 10 closed by record 0345) |
| P0 gap rows `planned`, PPTX | 16 | **0** |
| P0 property reads `unlisted`, PPTX | 4 | **0** |
| P0 gap rows `out-of-scope` | 4 | 5 (`modernComment.id`, below) |
| PPTX property writes: supported / partial / gap | 156 / 28 / 377 | 172 / 22 / 367 |
| PPTX property reads: supported / unlisted | 232 / 279 | 234 / 276 |

`pptx-p0-residue` is deleted from the dispositions; the new entry
`pptx-native-comment-ids` (out-of-scope) holds the modern-comment GUID, a
native identity read through raw-admitted queries only (ADR 0034), as the
comment and paragraph identities are. Both `--check` runs are green.

| Row | OfficeCLI | Litchi now |
|---|---|---|
| comment `author`, `initials`, `date` | asg | asg |
| comment `direction` (alias `dir`) | a-- | a-- |
| modernComment `author`, `initials`, `text`, `resolved`, `created` | asg | asg |
| modernComment `parent` | a-g | a-g |
| modernComment `id` | --g | raw read (out-of-scope) |
| notes `text` | asg | asg |
| placeholder `text` | asg | as- (the node's text is read; as for `shape`) |
| slide `layout` | asg | asg |
| slide `title`, `text` | a-- | a-- |

## 2. What landed

### Legacy comments

- **Leaf.** `opened::Edit::set_comment_date_at(slide, Position, date)` and
  `set_comment_author_at(slide, Position, name, initials)`, positional
  siblings of `set_comment_at`. A comment author is a record the comments
  share: when no other comment of the presentation names it, it is renamed
  in place; otherwise the comment moves to the author of that name and
  initials (a new one, with the next free identifier, when absent) and takes
  that author's next index, so the other comments keep their attribution.
- **Unified.** `set /slide[N]/comment[K] {author, initials, date}`;
  `add … comment {direction=rtl}` leads the text with U+200F.

### Modern (threaded) comments

- **Leaf.** `opened::Edit::{add_modern_comment_text,
  add_modern_comment_reply_text, set_modern_comment_at}` with
  `opened::ModernCommentChange { text, author, initials, created, resolved }`;
  `presentation::modern_comments::{Comment, Reply}::author_initials`. A new
  thread is anchored at the slide origin; its body is a plain
  `p188:txBody` in the list's own prefix (one `a:p` per line). Authors match
  by name and initials; a shared author is forked, not renamed (as above).
  New comment, reply and author GUIDs are derived (SHA-256) from the edit's
  inputs and checked unique in the graph: no clock or random source, so
  equal edits of equal packages produce equal bytes.
- **Unified.** `add /slide[N] modernComment {text, author, initials,
  created, resolved, parent}` (`parent=/slide[N]/moderncomment[K]` adds a
  reply and returns `…/reply[R]`); `set` of a thread (`text`, `author`,
  `initials`, `created`, `resolved`) and of a reply (no `resolved`). Reads
  gain `initials`, `resolved` (the thread's `resolved` status), `parent`
  (a reply's thread path) and the raw `id`.

### Notes, placeholders, slides

- **Leaf.** `opened::Edit::add_placeholder(slide, PlaceholderKind, text)`
  binds the new placeholder by OfficeCLI's `AddPlaceholder` rules: a title
  adopts the layout's title-family type (`ctrTitle` on "Title Slide"),
  other types the index of the layout slot of their type, else the smallest
  index the slide does not use; a second title is refused. A bound
  placeholder keeps an empty `p:spPr` (it inherits the slot, as PowerPoint
  writes it); an unbound one gets OfficeCLI's default rectangle.
  `put_placeholder_text(slide, kind, index, text)` fills the slide's
  matching placeholder or adds one. `set_slide_layout(slide, name)`
  retargets the slide's layout relationship (shapes are kept).
- **Unified.** `add /slide[N] placeholder {phType (alias type), text}`
  returns the canonical `/slide[N]/shape[K]`; `add /slide[N] notes {text}`
  creates the notes page when absent and sets its text, returning
  `/slide[N]/notes`; `add / slide {layout, title, text}` fills the title
  and body placeholders Litchi's insertion materializes from the layout,
  adding one only where the layout has none (OfficeCLI's new slides have no
  placeholders and always add `Title`/`Content` shapes); `set /slide[N]
  {layout}`.

### Coalescing (record 0333 F3)

Table cells (`…/tr[R]/tc[C]`: text, character formatting, fill, anchor,
insets), cell paragraphs and runs, notes paragraphs and runs (eager owners),
and direct paragraph `text` join the shape-text run: one native commit per
batch with per-item change judgement. Overlap barriers are per cell,
paragraph and run, kept in O(1) indexes (`Touched`). Tables are read once
per run and memoized per handle (`Table::cell_contents`, one namespace
scope for a table without declarations).

Two leaf passes were quadratic and are now linear (outputs equal the
one-edit-at-a-time result; leaf test
`batched_paragraph_and_cell_edits_equal_one_edit_at_a_time`):
`set_text_property_chunks` reparsed the whole table frame for every cell
edit (1,000 cells: 1.24 s → 7.1 ms) and the whole shape for every
paragraph edit of one body (1,000 notes paragraphs: 411 ms → 6.8 ms).

## 3. Decisions and deltas from OfficeCLI

| # | OfficeCLI | Litchi |
|---|---|---|
| 1 | modern comment `created` defaults to now; author defaults to `OfficeCli` | `created` and `author` are required on add (`missing_property`): the editor reads no clock (ADR 0034) |
| 2 | new GUIDs are random | derived from the edit's inputs and the graph's identities, unique, deterministic |
| 3 | re-attributing a comment to a forked author keeps its `idx` | the comment takes the new author's next index (`(authorId, idx)` stays unique and within `lastIdx`) |
| 4 | `rtl` is an alias of `direction` | only `dir`: `rtl=rtl` is not a meaningful spelling of a token |
| 5 | a new slide has no placeholders; `title`/`text` always add shapes | the layout's placeholders exist and take the text; shapes are added only where missing |
| 6 | a placeholder always gets an explicit `a:xfrm` | only an unbound one; a bound one inherits its slot (empty `p:spPr`) |
| 7 | `lang="en-US"` on new runs | no language invented (record 0333 #3) |
| 8 | modernComment `id` is a plain read | raw-admitted reads only (`query::raw::query("moderncomment[@id=…]")`) |
| 9 | `noGrp` / `size` / `phIndex` on placeholder add | not taken (P1 breadth, `pptx-breadth`) |

## 4. Registry, checklist, corpus, fidelity

Goldens re-blessed in both feature sets (all features; every feature but
`encryption`); `element_operations_agree_with_property_flags` stays green.
Element claims: `modernComment` `gqr` → `asgqr`, `reply` `gqr` → `sgqr`,
`placeholder` gains `add`, `notes` gains `add`.

- CRUD checklist: PPTX claimed cells 299 → **308**, all ok (new recipe:
  `modernComment` add with author, text and creation stamp).
- Corpus: PPTX passed 1,689 → 1,698, warned 823 → 814 (`ppt/animations.sh`,
  `ppt/video.sh`: slide `title`); Bugs: none; v1 subset unchanged and green
  (3,443 lines, 0 open).
- Fidelity (`results/0289/fidelity`): unchanged (every PPTX file exact).

## 5. Byte identity

Existing outputs keep their bytes: no reader or writer of an existing call
changed its output (the new verbs and keys are new calls). The leaf
batching changes rewrite the same spans with the same functions; one-edit
calls take the same path, and many-edit calls are checked equal to the
sequential result (leaf test above; `pptx_p0_final`
`cell_batches_are_one_commit`, `overlapping_cell_plans_keep_sequential_semantics`,
`paragraph_text_batches_are_one_commit`, `notes_batches_*`,
`source_owned_notes_keep_the_ordinary_writer`, each comparing the batch
with item-by-item `set`s on both owners). The corpus-fidelity run (open →
empty edit → save byte-exact, set-then-revert) is unchanged.

## 6. Measurements

`crates/litchi-ops/examples/pptx_p0_final_bench.rs`, release, `taskset -c 7`,
ABBA (leaf, unified, unified, leaf), medians; both legs' packages compared
part by part outside the clock (all equal). Two runs:
`results/0346/pptx-p0-final-bench-run{1,2}.jsonl`, binary SHA-256 in
`results/0346/binaries.sha256`.

Single operations, eager owner (budget 10%):

| scenario | leaf µs | unified µs | ratio (run 1 / 2) |
|---|---:|---:|---|
| `comment_date_set` | 424 | 460 | 1.085 / 1.088 |
| `comment_author_set` | 459 | 497 | 1.082 / 1.086 |
| `modern_comment_add` | 481 | 485 | 1.008 / 1.008 |
| `modern_comment_set` | 414 | 432 | 1.042 / 1.046 |
| `modern_reply_add` | 488 | 510 | 1.045 / 1.039 |
| `placeholder_add` | 408 | 410 | 1.004 / 1.009 |
| `notes_add` | 1,001 | 1,001 | 1.000 / 1.001 |
| `slide_add_title_text` | 2,823 | 2,827 | 1.001 / 1.001 |
| `slide_layout_set` | 486 | 487 | 1.003 / 1.005 |

Batches (budget: 1k within 2× of one leaf edit; linear 1 → 10k). Leaf leg:
one native edit staging the same `TextEdit`s in 256-edit chunks and one
commit. `cold` batches on a fresh handle each time (opened outside the
clock); `warm` reuses one handle, whose table memo then holds.

| scenario | N | leaf µs | unified µs | ratio (run 1 / 2) |
|---|---:|---:|---:|---|
| `cell_format_batch/eager` cold | 1,000 | 7,135 | 13,573 | 1.902 / 1.906 |
| | 10,000 | 67,884 | 136,572 | **2.012 / 1.977** |
| `cell_format_batch/eager` warm | 1,000 | 7,101 | 7,811 | 1.100 / 1.094 |
| `cell_format_batch/source` cold | 1,000 | 10,284 | 16,410 | 1.596 / 1.595 |
| | 10,000 | 101,423 | 167,000 | 1.647 / 1.636 |
| `cell_paragraph_batch/eager` cold | 1,000 | 6,045 | 12,486 | **2.065 / 2.043** |
| | 10,000 | 57,583 | 124,104 | **2.155 / 2.164** |
| `cell_paragraph_batch/eager` warm | 1,000 | 6,072 | 6,699 | 1.103 / 1.101 |
| `cell_paragraph_batch/source` cold | 1,000 | 9,494 | 15,342 | 1.616 / 1.622 |
| `paragraph_text_batch/eager` | 1,000 | 10,607 | 13,141 | 1.239 / 1.239 |
| | 10,000 | 104,006 | 132,289 | 1.272 / 1.271 |
| `paragraph_text_batch/source` | 1,000 | 20,220 | 22,826 | 1.129 / 1.127 |
| `notes_format_batch/eager` | 1,000 | 6,761 | 11,292 | 1.670 / 1.668 |
| | 10,000 | 511,544 | 564,855 | 1.104 / 1.094 |

Every unified leg scales linearly (cell, paragraph-text: ×9.3–10.6 per
×10). Fixes made to reach these numbers (D5):

- cell plans first resolved each path through the navigator, which parsed
  the table per item: O(n²) (100 cells 78 ms against 17 ms). Plans now read
  the table once per run;
- that table read resolved the namespace scope by rescanning the part for
  every cell (`Cell::text_body`) and row (`Row::cells`): `Table::cell_contents`
  resolves it once (1k cells 300 ms → 26 ms);
- the run's overlap test scanned every scope of the frame (all cells of one
  table share it): O(1) indexes;
- the batch's first plan and its run each read the table: the read is
  memoized per handle (1k cold 2.4× → 1.9×).

**Misses.** Cold cell-paragraph batches are 2.04–2.16× at 1k and 10k
(eager), and cold cell-format batches 1.98–2.01× at 10k: a fresh handle
reads every cell's properties and text body once (about 5 µs per cell) to
judge each item's change before staging, which the leaf leg does not do,
plus the first edit's admission on a fresh eager handle (≈ 0.35 ms, also at
N = 1). Warm batches are 1.06–1.15×. Follow-up 0346-P1: per-edit change
flags from `set_text_property_chunks`, so plans need not read the bodies.
The small notes batches (1.5–2.3× at N ≤ 100, ≤ 0.85 ms absolute) pay the
notes and slide scene reads of the first plan and the run; the 10k notes
leaf is superlinear in the leaf itself (40 256-edit chunks, each rewriting
the notes part: follow-up 0346-P2, one pass over all chunks).

## 7. Gates

On the merged tree (`d392343e02` and the record commits on top); cargo
through the memory-capped shim, `CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`;
every exit status checked (`(deleted-coord-artifact)`).

| gate | exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-pptx -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 (first run: `items_after_test_module` in `table/shape.rs`, fixed) |
| `cargo clippy -p litchi-ops --features <all but encryption> --all-targets -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features pptx] --all-targets -- -D warnings` | 0, 0 |
| `cargo test -p litchi-pptx --all-features` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, 9 `--test` chunks of 20 (executables deleted per chunk) | 0 |
| `cargo test -p litchi --all-features`; `cargo test -p litchi-cli --all-features` | 0; 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry` (all features; all but `encryption`) | 0, 0 (no change after the merge) |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus`; `tools/officecli_corpus.py`; `--check` | 0, 0, 0 |
| `tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI`; `--check` | 0, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pptx -p litchi-ops` | 0 |
| `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify` | 0; 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_unified_bench_matrix` | 0 |
| `tools/unified_bench_matrix.py --check` | 0 (first run: record 0346 had evidence but no parser; `p0346` added, the cold misses annotated) |
| `tools/check_feature_matrix.py` | 0 |

## 8. Follow-ups

- 0346-P1: per-edit change flags from the leaf text staging (cold cell
  batches, §6).
- 0346-P2: `set_notes_text_property_chunks` (one notes-part pass for all
  chunks).
- 0346-P3: placeholder add `name`, `noGrp`, `size`, `phIndex`; notes
  `direction` and `lang` (P1 breadth).
- 0346-P4: source owners are promoted to an eager package for comment,
  modern comment, notes, placeholder, slide `title`/`text` and layout
  writes (as record 0252 §6 notes for comments); a source-backed variant
  of the slide relationship and notes-part edits.
- Modern comments added here carry no `p188:commentRel` slide extension
  (neither do OfficeCLI's, nor the leaf's existing `add_modern_comment`).
