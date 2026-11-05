# 0348 — PPTX batches within budget: leaf-reported changes, one-pass indexes, one-pass notes chunks

- Status: implemented on `uo2/pptx-batch`, cut from `feat/unified-ops-api`
  at `8ba0c2260f`; trunk merged at `857d5c9976` (docs and dispositions)
  and at `c065dfafd8` (record 0349: canonical durable JSON, workspace
  hygiene; the bench-matrix Markdown regenerated).
- Date: 2026-10-19
- Input: record 0346 §6 and follow-ups 0346-P1 (cold cell batches over
  the 2× budget: cell paragraphs 2.04–2.16× at 1k and 10k, cell formats
  1.98–2.01× at 10k) and 0346-P2 (the 10k notes leaf leg superlinear); the
  coordinator's scope addition F-0344-TSV (`bench-matrix.tsv` at 88.5% of
  1 MiB).
- Owned: `crates/litchi-pptx` (batched text staging, table and body text
  indexes), `crates/litchi-ops/src/format/pptx*`, the 0346 bench example,
  `tools/unified_bench_matrix.py`. Small hunks outside them:
  `transaction.rs` (a `reported` list: the field, its fill in the PPTX
  flush arm, its truncation in `isolated`, `take_reported`) and
  `batch/execute.rs` (results take the reported changes after the last
  flush).
- Revised after the coordinator review of 2026-10-19 (§3, "Review"): the
  first version let a coalesced item succeed on a body the text-body
  reader refuses; batches now refuse at that item with its single set's
  code, and `tests/pptx_batch_equivalence.rs` checks it over the corpus.
- `performance_claim`: measured (§5). Every PPTX cell, paragraph, run and
  shape batch route at 1k and 10k is within 1.31× of the leaf (cold and
  warm handles, eager and source owners); notes batches are 1.19–1.24× on
  a warm handle and 1.89–1.97× on a fresh one.

## 1. Result

| Route (`pptx_p0_final_bench`, 1k / 10k) | 0346 | 0348 |
|---|---|---|
| `cell_format_batch/eager` cold | 1.90 / **2.01** | 1.26 / 1.25 |
| `cell_paragraph_batch/eager` cold | **2.07** / **2.16** | 1.31 / 1.30 |
| `cell_format_batch/source` cold | 1.60 / 1.65 | 1.19 / 1.18 |
| `cell_paragraph_batch/source` cold | 1.62 / — | 1.20 / 1.20 |
| cell batches, warm (every owner) | 1.06–1.15 | 1.05–1.11 |
| `notes_format_batch/eager` (0346: one handle) | 1.67 / 1.10 | warm 1.19 / 1.22, cold 1.90 / 1.94 |

The 0346 misses are closed; no PPTX batch row of the matrix misses its
budget. The 10k notes leaf leg is linear (511 ms → 22 ms).

## 2. What landed

### Leaf-reported changes (0346-P1)

A cell plan used to read every cell's properties and text body
(`Table::cell_contents`, about 5 µs a cell on a fresh handle) to judge
whether each item changes anything. The leaf already knows: its staging
compares each rewritten element with the bytes before it.

- **Leaf.** `opened::Edit::set_text_property_chunks_per_edit(slide,
  chunks) -> Vec<bool>` and the same on `source::Edit` (and the focused
  `SourceBackedSlideEdit`): the staging of `set_text_property_chunks`,
  returning each edit's change in order, as one `set_text_properties` call
  per edit applied one after the other reports it.
  `set_notes_text_property_chunks_per_edit` does the same for notes.
- **Unified.** Cell plans (and notes plans, below) are `reported`: their
  staging returns `false`, and the run's commit hands each one's leaf
  change back with its wire index. The transaction keeps them
  (`reported`), drops those of an isolated command (whose change is its
  steps'), and `File::batch` writes them into the item results after its
  last flush. Item results therefore equal the item-by-item ones
  (`cell_batch_changes_are_the_items_own`,
  `notes_batch_changes_and_misses_are_the_items_own`).

### One-pass indexes

Planning still resolves each address before staging, so a miss fails with
the item's own error (`cell_batch_misses_fail_as_the_item_alone`): it
needs the cells, paragraph counts and the positions of text runs among a
paragraph's run children, not their properties or text.

- `table::CellTextIndex`, from `Table::cell_text_index` or, in one pass
  over the graphic frame, `shape::owned::Handle::cell_text_index`. When the
  fragment's namespace declarations are on its root only (PowerPoint and
  Litchi tables alike), a tag scanner resolves each prefix once and needs
  no namespace reader. Markup the pass does not vouch for — markup
  compatibility, a nested or second table, a second text body in a cell,
  declarations below the root, a fragment without its owner, processing
  instructions — is indexed cell by cell through `Cell::text_body`.
- **Validation (after the review).** The single `set` of a cell
  paragraph or run reads that cell's body with the shared text-body
  reader, and refuses with the reader's error (a cell `set` reads no
  body). The index therefore validates every cell body with that same
  reader and indexes a refused body as no body: its paragraph and run
  items take the ordinary writer, which refuses them at their own index
  with the single set's code, while every other item still coalesces. The
  reader's verdict depends on a body's markup and attribute values, not
  on the plain character data of its runs, so the bodies are keyed by
  their bytes with each run of plain character data (no references or
  control characters, at most 1 KiB, valid UTF-8) reduced to one byte
  telling whitespace from text, and each distinct key is read once: a
  table of 1,000 cells written by one producer costs a handful of
  readings. Bodies with comments, character data sections or more than
  64 KiB are read on their own. The cell-by-cell fallback indexes a
  refused body as no body too, instead of failing the whole table (which
  the first version, and record 0346, did for every item of that table).
- `shape::BodyTextIndex`: `Common::text_index` reads the body with the
  reader; `Common::read_text_index` indexes a body the reader has already
  read in one pass.
- Equality with the readers is tested on hand-made tables (runs, breaks,
  fields, empty paragraphs, empty and missing bodies, foreign and
  unmodeled children, strict namespaces, comments, CDATA, quoted `>`,
  inner declarations, and 13 bodies of which 5 the reader refuses:
  `cell_text_index_drops_the_bodies_the_reader_refuses`) and on every
  table and text body of the tracked PPTX corpus
  (`text_indexes_of_the_corpus_equal_the_readers`).

### Notes (0346-P2)

- **Leaf.** `opened::Edit::set_notes_text_property_chunks(slide, chunks)`
  stages every chunk in one notes-part read, rewrite and read-back; the
  result equals one `set_notes_text_properties` call per chunk
  (`notes_chunks_equal_one_call_per_chunk`). The coalesced notes run and
  the ordinary notes text-format writer both use it.
- **Unified.** Notes plans resolve their address as the navigator does:
  `Slide::notes` (which validates the notes graph and reads the body
  placeholder's text with the reader), `/slide[N]/notes` only when that
  text is not empty, and the first body placeholder's paragraphs, indexed
  with `read_text_index`. The view is memoized per handle (a batch's first
  plan and its run read it once), and plans take their change from the
  leaf. The first version read an eager notes snapshot instead (0.65 ms
  against 1.5 ms at 1,000 paragraphs), which the corpus equivalence test
  showed diverging on a notes page with an empty body (`tdf131082.pptx`:
  the single set misses `/slide[1]/notes`, the batch item succeeded); the
  record 0346 plans had the same divergence.
- **Admission edit.** A fresh eager handle's transaction admission builds
  a native edit to run the strict checks; the handle now keeps it for its
  first coalesced run instead of building a second one (an edit is a
  function of the immutable package; it is matched to the package by
  identity).

### Bench matrix (F-0344-TSV)

`tools/unified_bench_matrix.py` writes one row table per format,
`results/0289/bench-matrix/<format>.tsv` (largest: `rtf.tsv`, 211 KB),
next to the per-format pages; `bench-matrix.tsv` is deleted, and
`--check` reports it while it is still on a tree. The 1 MiB rule still
covers every output. No reader of the TSV existed besides the record 0289
prose, which now names the new layout. A `p0348` parser collects this
record's runs under the 0333 and 0346 series names, so those rows are
superseded.

## 3. Decisions and deltas

| # | Before | Now |
|---|---|---|
| 1 | a cell or notes item's `changed` is the reader's judgement of its properties | the leaf's byte comparison, as the ordinary writer's commit reports it; the two agree on every test |
| 2 | planning reads the table's every cell (properties and body): a refused body anywhere fails every item of the table | a refused body fails exactly the items whose single set reads it (that cell's paragraphs and runs), with the single set's code at their index |
| 3 | the 0346 notes route ran on one reused handle without a memo | it now has `cold` (fresh handle per batch) and `warm` legs; the old series is superseded by the warm leg's rows only in name (`notes_format_batch/eager/<N>/{cold,warm}`) |

### Review

The coordinator's review of the first version (decision #2: "a value the
reader refuses no longer fails a coalesced batch item") is resolved by the
validation above; that decision is withdrawn, and row 2 records the
replacement. Probe of the single `set` on a table whose cell (1,1) holds
`algn="bogus"`, `sz="abc"` or two `a:rPr`, both owners: `tc[1]` succeeds,
`tc[1]/paragraph[1]` refuses (`corrupt_file`), other cells succeed; the
batch now matches each case.

`tests/pptx_batch_equivalence.rs` (as record 0314's iWork test): every
PPTX fixture under `test-data/` on both owners, three seeded sequences of
12 sets each (cells, cell paragraphs and runs, notes paragraphs and runs,
shape paragraphs; repeated and missing targets; bold, italic, align,
size, text, fill), run as one `File::edit` closure, one atomic
`File::batch` and a chain of single sets: the same bytes, or the same code
at the same item (a refusal before the first item, a signed package's
admission, counts as item 0). The crafted decks put `algn="bogus"`,
`marL="wide"` or two `a:pPr` into a cell and a notes body.

## 4. Byte identity

No writer's output changed: the leaf staging is the same code with a flag
vector; `set_notes_text_property_chunks` equals the per-chunk calls. Every
bench scenario compares the leaf and unified packages part by part (all
equal, 3 × 89 rows); the batch-equivalence tests compare each batch with
item-by-item sets on both owners, over the corpus too.

## 5. Measurements

`crates/litchi-ops/examples/pptx_p0_final_bench.rs` (record 0346's routes;
notes gained `cold`/`warm`) and the batch scenarios of `pptx_p0_bench.rs`
(record 0333's), release, `taskset -c 7`, ABBA, medians, three runs each:
`results/0348/pptx-p0-final-bench-run{1,2,3}.jsonl`,
`results/0348/pptx-p0-bench-run{1,2,3}.jsonl`; binaries built at
`34685b9c25`, SHA-256 in `results/0348/binaries.sha256`. Ratios of the
three runs:

| scenario | N | leaf µs | unified µs | ratios |
|---|---:|---:|---:|---|
| `cell_format_batch/eager` cold | 1,000 | 7,228 | 9,090 | 1.261 / 1.258 / 1.267 |
| | 10,000 | 68,552 | 85,981 | 1.253 / 1.258 / 1.254 |
| `cell_format_batch/eager` warm | 1,000 | 7,269 | 7,835 | 1.083 / 1.078 / 1.083 |
| | 10,000 | 68,384 | 74,275 | 1.093 / 1.088 / 1.086 |
| `cell_paragraph_batch/eager` cold | 1,000 | 6,140 | 8,049 | 1.307 / 1.311 / 1.303 |
| | 10,000 | 57,678 | 74,971 | 1.300 / 1.311 / 1.295 |
| `cell_paragraph_batch/eager` warm | 1,000 | 6,150 | 6,790 | 1.099 / 1.104 / 1.095 |
| | 10,000 | 57,655 | 63,472 | 1.112 / 1.102 / 1.097 |
| `cell_format_batch/source` cold | 1,000 | 10,359 | 12,246 | 1.191 / 1.187 / 1.181 |
| | 10,000 | 101,815 | 121,251 | 1.193 / 1.184 / 1.179 |
| `cell_format_batch/source` warm | 1,000 | 10,365 | 10,932 | 1.067 / 1.056 / 1.050 |
| | 10,000 | 102,297 | 107,742 | 1.064 / 1.053 / 1.050 |
| `cell_paragraph_batch/source` cold | 1,000 | 9,536 | 11,508 | 1.209 / 1.201 / 1.197 |
| | 10,000 | 93,149 | 111,678 | 1.213 / 1.199 / 1.187 |
| `cell_paragraph_batch/source` warm | 1,000 | 9,471 | 10,175 | 1.074 / 1.061 / 1.056 |
| | 10,000 | 93,625 | 99,987 | 1.075 / 1.061 / 1.060 |
| `notes_format_batch/eager` cold | 1,000 | 2,641 | 5,012 | 1.920 / 1.898 / 1.891 |
| | 10,000 | 22,071 | 42,601 | 1.939 / 1.922 / 1.966 |
| `notes_format_batch/eager` warm | 1,000 | 2,651 | 3,147 | 1.200 / 1.187 / 1.193 |
| | 10,000 | 22,200 | 26,985 | 1.217 / 1.216 / 1.236 |
| `paragraph_text_batch/eager` | 1,000 | 10,926 | 13,532 | 1.247 / 1.231 / 1.236 |
| | 10,000 | 107,969 | 136,396 | 1.278 / 1.262 / 1.271 |
| `paragraph_text_batch/source` | 10,000 | 208,180 | 235,389 | 1.156 / 1.131 / 1.131 |
| `run_format_batch/eager` (0333) | 1,000 | 12,349 | 15,563 | 1.230 / 1.234 / 1.260 |
| | 10,000 | 122,719 | 154,603 | 1.247 / 1.251 / 1.260 |
| `run_format_batch/source` (0333) | 10,000 | 216,735 | 248,880 | 1.144 / 1.148 / 1.148 |
| `shape_text_batch/eager` (0333) | 10,000 | 96,983 | 104,125 | 1.080 / 1.070 / 1.093 |
| `shape_text_batch/source` (0333) | 10,000 | 195,243 | 202,580 | 1.038 / 1.039 / 1.052 |

The first version (binaries at `0e5ecb8a20`, without body validation and
with the notes snapshot) measured cold cells 1.33–1.39× (eager) and cold
notes 1.61–1.72×; the validation memo costs about 1% on these tables, the
kept admission edit gains about 5% on every cold route, and the exact
notes view costs cold notes 0.3–0.4×.

The remaining rows (N = 1–100, the single operations) are in the JSONL
files and the matrix. Every unified leg scales linearly: ×8.5–10.2 from
1k to 10k (below ×10 where the fresh handle's fixed cost is amortized).

Profile (callgrind, `cell_paragraph_batch/eager/1000`) of the path to
these numbers, per cold batch, against about 145 M instructions for the
leaf leg:

- the cell read the plans made (`cell_contents`): 185 M → a namespace
  reader index 25 M → the same with root-only declarations 20 M → the
  one-pass frame index (no `Handle::table`, whose own frame scan was 13 M)
  about 1 M;
- remaining cold-only costs: the fresh eager handle's first native edit
  (it opens the presentation; the leaf leg reuses one package) and the
  slide scene read.

Notes, per cold batch at 1,000 paragraphs: two `Slide::notes` reads and
body decodes (the first plan's and the run's, about 40 M each) → one
`Slide::notes` read (the navigator's, about 31 M) and one structure pass.
That read is as large as the leaf's own staging of the batch, which is
why cold notes stay near 1.9×.

**Small batches.** Cold batches of 1–10 items are 1.5–2.1× (≈ 0.3–0.45 ms
absolute): the fresh handle's first edit admission and scene reads, with
no ratio budget below 1k (scaling rows).

## 6. Gates

On the merged tree `c065dfafd8` (after the review's changes and trunk
record 0349), rerun in full; cargo through the memory-capped shim,
`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`; every exit status checked. The
litchi-ops test chunks include `tests/pptx_batch_equivalence.rs`.

Logs: `(deleted-coord-artifact)` (`summary.txt`).

| gate | exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-pptx -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --features <all but encryption> --all-targets -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features pptx] --all-targets -- -D warnings` | 0, 0 |
| `cargo test -p litchi-pptx --all-features` | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, 9 `--test` chunks of 20 (executables deleted per chunk) | 0 |
| `cargo test -p litchi --all-features`; `cargo test -p litchi-cli --all-features` | 0; 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pptx -p litchi-ops` | 0 |
| `tools/check_crate_boundaries.py`; `tools/non_iwork_gate.py verify` | 0; 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_unified_bench_matrix` | 0 |
| `tools/officecli_capability_delta.py --check`; `tools/officecli_corpus.py --check` | 0; 0 |
| `tools/unified_bench_matrix.py --check` (1 MiB rule included) | 0 |
| `tools/check_feature_matrix.py` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry` (all features; all but `encryption`), then a clean tree | 0, 0, 0 (no golden changed) |

The capability delta and the OfficeCLI corpus are unchanged (no claim or
corpus-visible behaviour changed), so neither was regenerated.

## 7. Follow-ups

- 0348-P1: cold notes batches (1.89–1.97×) pay one `Slide::notes` read
  per fresh handle, as large as the leaf staging; a leaf notes view that
  validates and indexes the body in one pass (the reader's verdict and
  structure together), shared with the native notes editor's own read,
  would bring them towards the cell routes.
- 0348-P2: the kept admission edit stays in a handle until a coalesced
  run takes it; a handle whose first transaction has no coalesced PPTX
  run keeps it (the edit shares the package's parts) until it is dropped.
- 0346-P3 and 0346-P4 are unchanged.
