# 0348 — PPTX batches within budget: leaf-reported changes, one-pass indexes, one-pass notes chunks

- Status: implemented on `uo2/pptx-batch`, cut from `feat/unified-ops-api`
  at `8ba0c2260f`; trunk merged at `857d5c9976` (docs and dispositions
  only; no crate changed).
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
- `performance_claim`: measured (§5). Every PPTX batch route at 1k and
  10k is within 1.39× of the leaf (cold and warm handles, eager and source
  owners) except cold-handle notes batches, 1.61–1.72×.

## 1. Result

| Route (`pptx_p0_final_bench`, 1k / 10k) | 0346 | 0348 |
|---|---|---|
| `cell_format_batch/eager` cold | 1.90 / **2.01** | 1.33 / 1.30 |
| `cell_paragraph_batch/eager` cold | **2.07** / **2.16** | 1.38 / 1.35 |
| `cell_format_batch/source` cold | 1.60 / 1.65 | 1.18 / 1.19 |
| `cell_paragraph_batch/source` cold | 1.62 / — | 1.19 / 1.18 |
| cell batches, warm (every owner) | 1.06–1.15 | 1.06–1.12 |
| `notes_format_batch/eager` (0346: one handle) | 1.67 / 1.10 | warm 1.19 / 1.22, cold 1.70 / 1.61 |

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
  no namespace reader (1,000 cells: about 20 µs of the former 5 ms). Markup
  the pass does not vouch for — markup compatibility, a nested or second
  table, a second text body in a cell, a fragment without its owner,
  processing instructions — is indexed from `cell_contents`, with its
  errors.
- `shape::BodyTextIndex` from `Common::text_index`, for a shape's body
  (the notes page's), with the same fallback to `text_body`.
- Equality with the readers is tested on hand-made tables (runs, breaks,
  fields, empty paragraphs, empty and missing bodies, foreign and
  unmodeled children, strict namespaces, comments, CDATA, quoted `>`,
  inner declarations) and on every table and text body of the tracked
  PPTX corpus (`text_indexes_of_the_corpus_equal_the_readers`).

### Notes (0346-P2)

- **Leaf.** `opened::Edit::set_notes_text_property_chunks(slide, chunks)`
  stages every chunk in one notes-part read, rewrite and read-back; the
  result equals one `set_notes_text_properties` call per chunk
  (`notes_chunks_equal_one_call_per_chunk`). The coalesced notes run and
  the ordinary notes text-format writer both use it.
- **Unified.** Notes plans read the body placeholder from the eager notes
  snapshot the native notes editor stages on (0.65 ms against 1.5 ms for
  `Slide::notes` at 1,000 paragraphs; a notes page holding markup
  compatibility keeps `Slide::notes`), index it with `BodyTextIndex`,
  memoize it per handle (a batch's first plan and its run read it once),
  and take their change from the leaf.

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
| 2 | planning a cell item reads the cell's properties and body, so a value the reader refuses (an out-of-domain attribute) fails the item | the index reads structure only; such a value no longer fails a coalesced batch item (the ordinary writer, which reads the body, still refuses it). Malformed XML, markup compatibility and the other cases above keep the reader and its errors |
| 3 | the 0346 notes route ran on one reused handle without a memo | it now has `cold` (fresh handle per batch) and `warm` legs; the old series is superseded by the warm leg's rows only in name (`notes_format_batch/eager/<N>/{cold,warm}`) |

## 4. Byte identity

No writer's output changed: the leaf staging is the same code with a flag
vector; `set_notes_text_property_chunks` equals the per-chunk calls. Every
bench scenario compares the leaf and unified packages part by part (all
equal, 3 × 89 rows); the batch-equivalence tests compare each batch with
item-by-item sets on both owners.

## 5. Measurements

`crates/litchi-ops/examples/pptx_p0_final_bench.rs` (record 0346's routes;
notes gained `cold`/`warm`) and the batch scenarios of `pptx_p0_bench.rs`
(record 0333's), release, `taskset -c 7`, ABBA, medians, three runs each:
`results/0348/pptx-p0-final-bench-run{1,2,3}.jsonl`,
`results/0348/pptx-p0-bench-run{1,2,3}.jsonl`; binaries built at
`0e5ecb8a20`, SHA-256 in `results/0348/binaries.sha256`. Ratios of the
three runs:

| scenario | N | leaf µs | unified µs | ratios |
|---|---:|---:|---:|---|
| `cell_format_batch/eager` cold | 1,000 | 7,277 | 9,683 | 1.338 / 1.331 / 1.330 |
| | 10,000 | 68,809 | 90,521 | 1.305 / 1.283 / 1.316 |
| `cell_format_batch/eager` warm | 1,000 | 7,239 | 7,915 | 1.106 / 1.085 / 1.095 |
| | 10,000 | 69,503 | 75,939 | 1.089 / 1.100 / 1.093 |
| `cell_paragraph_batch/eager` cold | 1,000 | 6,109 | 8,405 | 1.390 / 1.372 / 1.376 |
| | 10,000 | 58,171 | 78,285 | 1.357 / 1.349 / 1.346 |
| `cell_paragraph_batch/eager` warm | 1,000 | 6,072 | 6,702 | 1.106 / 1.101 / 1.098 |
| | 10,000 | 57,780 | 64,173 | 1.111 / 1.116 / 1.111 |
| `cell_format_batch/source` cold | 1,000 | 10,304 | 12,203 | 1.180 / 1.184 / 1.178 |
| | 10,000 | 103,540 | 122,711 | 1.172 / 1.188 / 1.185 |
| `cell_format_batch/source` warm | 10,000 | 103,959 | 111,386 | 1.062 / 1.072 / 1.079 |
| `cell_paragraph_batch/source` cold | 1,000 | 9,481 | 11,283 | 1.194 / 1.189 / 1.184 |
| | 10,000 | 93,971 | 111,228 | 1.179 / 1.179 / 1.184 |
| `cell_paragraph_batch/source` warm | 10,000 | 93,823 | 99,971 | 1.066 / 1.068 / 1.066 |
| `notes_format_batch/eager` cold | 1,000 | 2,663 | 4,582 | 1.680 / 1.721 / 1.700 |
| | 10,000 | 22,297 | 35,929 | 1.612 / 1.634 / 1.611 |
| `notes_format_batch/eager` warm | 1,000 | 2,668 | 3,197 | 1.191 / 1.205 / 1.184 |
| | 10,000 | 22,132 | 27,086 | 1.219 / 1.235 / 1.223 |
| `paragraph_text_batch/eager` | 10,000 | 107,625 | 137,208 | 1.268 / 1.275 / 1.276 |
| `paragraph_text_batch/source` | 10,000 | 210,803 | 240,046 | 1.136 / 1.157 / 1.143 |
| `run_format_batch/eager` (0333) | 1,000 | 12,020 | 14,973 | 1.234 / 1.253 / 1.250 |
| | 10,000 | 118,880 | 150,165 | 1.263 / 1.258 / 1.274 |
| `run_format_batch/source` (0333) | 10,000 | 211,040 | 243,147 | 1.143 / 1.156 / 1.150 |
| `shape_text_batch/eager` (0333) | 10,000 | 95,345 | 103,233 | 1.077 / 1.083 / 1.063 |
| `shape_text_batch/source` (0333) | 10,000 | 191,699 | 199,858 | 1.043 / 1.049 / 1.032 |

The remaining rows (N = 1–100, the single operations) are in the JSONL
files and the matrix. Every unified leg scales linearly: ×7.8–10.3 from
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
snapshot read and one index.

**Small batches.** Cold batches of 1–10 items stay 1.5–2.3× (≈ 0.35–0.5 ms
absolute): the fresh handle's first edit admission and scene reads, with
no ratio budget below 1k (scaling rows).

## 6. Gates

On the merged tree (`857d5c9976` and the record commit); cargo through the
memory-capped shim, `CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`; every exit
status checked.

Logs: `.coord/pptx-batch/gate-logs/` (`summary.txt`).

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

- 0348-P1: cold notes batches (1.61–1.72×) still read the notes snapshot
  and slide scene once per fresh handle; a leaf notes index taken from the
  native edit's own snapshot would remove the second read.
- 0348-P2: a fresh eager handle pays its first native edit twice (the
  transaction's admission and the run's edit), about 12 M instructions;
  handing the admission edit to the run would remove it, at the cost of
  retaining an edit in the handle.
- 0346-P3 and 0346-P4 are unchanged.
