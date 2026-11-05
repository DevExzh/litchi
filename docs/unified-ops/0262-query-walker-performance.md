# 0262 — Query walker performance, eager PPTX scenes, source PPTX metadata

- Status: implemented on `uo2/p8-perf`, cut from `feat/unified-ops-api` at
  `949224c319` and merged with trunk at `1ab00788e0` (records 0258–0260,
  0263). Plan step 8 of record [0248](0248-coordinator-restart-plan.md) §6
  Phase B, plus the misses of records 0252 (§5) and 0253 (§7) and the
  shared-code findings of record 0259.
- Date: 2026-10-13
- Scope: `crates/litchi-ops/src/{query,value,node,path}`, `format/pptx`
  (nodes, scenes, metadata), small hunks in `format/{xlsx,xls,xlsb}`, an
  additive leaf hunk in `litchi-xlsx` (`RecordIndex`), the harness
  `tools/perf-baseline/unified-query`, a first-set route in
  `examples/ooxml_metadata_bench.rs`.
- `performance_claim`: paired ABBA measurements below, release build,
  `taskset -c 28` on the shared 32-core host. Evidence:
  [results/0262/](results/0262/).

Reference: the stale branch `uo/m1-r4-perf` (`012fd509ff`, WIP) was read and
re-derived, not cherry-picked. Its value-layer patch (typed comparison and
its oracle tests) applied with one conflict; its walker, node and leaf ideas
were re-implemented against trunk's drifted adapters.

## 1. What changed

Query engine and resolver (shared, every format):
- **Fence scopes** (`path/fence.rs`). A source-backed handle must report a
  changed source on every read; trunk checked before and after every
  adapter call, so a query paid two checks per visited child and property
  (on a file source each is a mutex and an `fstat`). The resolver and the
  walker now run each `get`/`find` and each `Matches::next` step inside a
  per-thread scope keyed by the handle, and check once after the step.
  Adapter reads inside the scope skip their own checks (XLSX
  `XlsxOwner::finish`, PPTX `check_owner`, XLSB `nav::finish`; XLS already
  relied on the walker's check, record 0259). A source that changed before
  or during a step fails that check and its error wins, as before; the
  source-change tests (`retained_nodes_and_lazy_reads_refuse_changed_sources`,
  `cached_nodes_and_children_iterators_fence_source_revision_changes`,
  `lazy_file_query_and_cached_nodes_refuse_changed_source`) pass unchanged.
- **Walker shortcuts** (`query/plan.rs` `Shortcut`): a one-compound chain
  that tests only the kind, only a filter, or only content skips the
  structural and ancestor machinery. A condition that is its compound's
  only property reader reads one value (`Inspect::prop`) instead of the
  whole property set.
- **Typed comparison** (`value/compare.rs`, `length.rs`, `number.rs`):
  integers, finite numbers and lengths are compared from their typed value;
  the EMU value of a text is parsed only when a comparison reaches it. The
  old write-then-read evaluation is kept as a test oracle
  (`reference_eval`) and compared over generated values.
- **Small nodes and on-demand paths**: `NodePath` carries up to four steps
  by value, so nodes stop storing theirs. XLSX `Node` 248 → 88 bytes
  (asserted ≤ 96), PPTX 400 → 136 bytes, XLS drops its two stored steps
  (112 bytes). Labels and cell formats are read lazily / boxed.
- **Direct iteration** (XLSX): a worksheet's cells are its records in
  order (no range scan), and a listed cell keeps its record position, read
  again in O(1) (`litchi_xlsx::RecordIndex`, `Cells::indexed`,
  `SourceGrid::record`, `Worksheet::record`). Column/range filters read the
  coordinate directly (`Inspect::cell_at`) instead of re-parsing an A1
  label.
- **PPTX**: a query from `/` no longer re-reads the presentation part for
  the master count (cached per handle; the first match on a 500-slide deck
  went from 186 µs to 1 µs); a property filter reads only its key's inputs
  (§4, the one error-path delta).

Writes:
- **0252 miss, eager run formatting** (`format/pptx/scenes.rs`): the eager
  owner caches strict slide scenes keyed by the slide part's payload
  identity (`Arc<Vec<u8>>`, at most four, most recent first), shared by
  every handle a write derives. A commit shares untouched part payloads, so
  unchanged slides stay cached and a changed slide is read again.
- **0253 miss, source-backed metadata set** (`format/pptx/metadata.rs`): the
  eager package a source handle commits metadata on is built once per
  handle and kept, as the XLSX owner keeps its promotion. Measured
  breakdown on the generated deck: source-graph adoption 98 µs + eager
  package construction 27 µs, against 73 µs for the leaf edit itself.

## 2. Differential evidence

`tests/query_corpus_differential.rs`: 98 digests (37 XLSX, 25 DOCX, 33 PPTX
selectors plus file counts) over every `.xlsx`, `.docx` and `.pptx` under
`test-data` (181/60/81 files opened), each hashing every match's kind, path,
positional path and text and every error. The golden was printed by the
trunk build before any change (`examples/query_corpus_digest.rs`) and the
final build reproduces it exactly, with the `xlsx docx pptx` feature set and
with all features (encrypted files are skipped so digests do not depend on
the crypto feature). The selectors cover every comparison route (numbers,
lengths in cm/pt/in/EMU, booleans, tokens, colours, text), header rows,
positions, unions, `:has`, `:not` and combinators. Debug run: about 140 s.

## 3. Measurements

### Queries and gets (`unified-query`, routes `*_v1`)

Two before binaries (trunk `949224c319`) and two after binaries (merged
branch), alternated B A A B; each run is itself an in-process ABBA of nine
samples. Ratio = unified / leaf loop, one value per run.

| route | matches | before unified µs | after unified µs | leaf µs | ratio before | ratio after |
|---|---:|---:|---:|---:|---:|---:|
| xlsx_1000_bytes_get_cell | 1 | 0.51 | 0.42 | 0.16 | 3.09/3.09 | 2.60/2.61 |
| xlsx_1000_bytes_first_cell_value_gt | 1 | 1.90 | 0.86 | 0.14 | 13.01/13.48 | 5.86/5.91 |
| xlsx_1000_bytes_limit10_cell_value_gt | 10 | 14.28 | 5.37 | 0.44 | 33.48/32.93 | 12.29/11.53 |
| xlsx_1000_bytes_all_cell_value_gt | 500 | 649.51 | 239.33 | 16.83 | 39.64/39.97 | 13.73/13.69 |
| xlsx_1000_bytes_all_cell | 3003 | 342.61 | 154.79 | 9.09 | 37.47/43.82 | 15.36/17.25 |
| xlsx_1000_file_get_cell | 1 | 2.11 | 1.18 | 0.65 | 3.18/3.24 | 1.82/1.82 |
| xlsx_1000_file_first_cell_value_gt | 1 | 9.40 | 2.09 | 0.64 | 14.74/15.06 | 3.25/3.28 |
| xlsx_1000_file_limit10_cell_value_gt | 10 | 79.11 | 8.16 | 0.93 | 86.27/86.34 | 8.82/8.36 |
| xlsx_1000_file_all_cell_value_gt | 500 | 3966.67 | 320.33 | 17.47 | 215.60/229.52 | 17.74/18.10 |
| xlsx_1000_file_all_cell | 3003 | 1373.52 | 665.17 | 9.69 | 137.25/161.95 | 68.99/68.34 |
| xlsx_50000_bytes_get_cell | 1 | 0.53 | 0.44 | 0.18 | 2.92/2.90 | 2.43/2.47 |
| xlsx_50000_bytes_first_cell_value_gt | 1 | 2.01 | 0.87 | 0.16 | 12.76/13.12 | 5.56/5.58 |
| xlsx_50000_bytes_limit10_cell_value_gt | 10 | 15.45 | 5.39 | 0.45 | 35.30/34.95 | 12.13/11.57 |
| xlsx_50000_bytes_all_cell_value_gt | 25000 | 37287.52 | 11834.80 | 1085.21 | 30.15/34.82 | 10.78/11.12 |
| xlsx_50000_bytes_all_cell | 150003 | 17980.76 | 7715.00 | 714.34 | 5.63/25.63 | 10.50/11.11 |
| xlsx_50000_bytes_all_row_header_gt | 25000 | 23415.13 | 10729.99 | 4679.19 | 2.93/4.63 | 2.36/2.29 |
| xlsx_50000_file_get_cell | 1 | 2.16 | 1.19 | 0.67 | 3.11/3.12 | 1.78/1.77 |
| xlsx_50000_file_first_cell_value_gt | 1 | 9.66 | 2.09 | 0.64 | 14.74/14.82 | 3.25/3.28 |
| xlsx_50000_file_limit10_cell_value_gt | 10 | 80.27 | 8.19 | 0.94 | 86.11/86.02 | 8.75/8.51 |
| xlsx_50000_file_all_cell_value_gt | 25000 | 188971.75 | 16021.98 | 1104.70 | 164.24/169.09 | 14.56/15.02 |
| xlsx_50000_file_all_cell | 150003 | 69179.15 | 33119.02 | 1231.88 | 41.45/39.17 | 41.29/47.15 |
| xlsx_50000_file_all_row_header_gt | 25000 | 84684.10 | 40692.48 | 6180.24 | 12.57/11.38 | 7.68/7.37 |
| docx_drawing_bytes_get_paragraph | 1 | 1.01 | 0.95 | 0.47 | 2.13/2.01 | 2.01/2.05 |
| docx_drawing_bytes_first_run_bold | 1 | 8.16 | 7.84 | 7.08 | 1.14/1.13 | 1.12/1.12 |
| docx_drawing_bytes_all_run_bold | 177 | 4355.24 | 4248.60 | 3998.43 | 1.09/1.07 | 1.07/1.07 |
| docx_drawing_file_get_paragraph | 1 | 2.08 | 1.84 | 1.20 | 1.73/1.73 | 1.56/1.53 |
| docx_drawing_file_first_run_bold | 1 | 10.22 | 9.78 | 7.78 | 1.31/1.31 | 1.26/1.27 |
| docx_drawing_file_all_run_bold | 177 | 5131.57 | 5071.77 | 4003.48 | 1.28/1.28 | 1.27/1.27 |
| docx_saut_bytes_get_paragraph | 1 | 1.00 | 0.96 | 0.47 | 2.12/2.03 | 2.05/2.04 |
| docx_saut_bytes_first_run_bold | 0 | 20.27 | 20.19 | 17.12 | 1.19/1.20 | 1.16/1.18 |
| docx_saut_bytes_all_run_bold | 0 | 20.35 | 20.24 | 17.23 | 1.17/1.20 | 1.17/1.18 |
| docx_saut_file_get_paragraph | 1 | 2.04 | 1.84 | 1.19 | 1.73/1.71 | 1.55/1.51 |
| docx_saut_file_first_run_bold | 0 | 28.61 | 28.33 | 17.79 | 1.61/1.61 | 1.56/1.60 |
| docx_saut_file_all_run_bold | 0 | 28.81 | 28.24 | 17.75 | 1.63/1.62 | 1.56/1.61 |
| pptx_small_bytes_get_shape | 1 | 0.62 | 0.59 | 0.10 | 6.30/6.34 | 5.96/6.13 |
| pptx_small_bytes_first_shape_contains | 1 | 13.10 | 0.63 | 0.09 | 147.90/151.73 | 7.27/7.11 |
| pptx_small_bytes_limit10_shape_width | 10 | 289.25 | 12.53 | 1.58 | 185.95/181.07 | 7.78/8.13 |
| pptx_small_bytes_all_shape_width | 10 | 363.59 | 63.65 | 1.66 | 221.74/217.10 | 38.35/38.29 |
| pptx_small_bytes_all_shape_contains | 10 | 74.81 | 59.08 | 1.73 | 43.50/43.17 | 34.13/34.23 |
| pptx_small_bytes_all_shape | 90 | 77.84 | 61.84 | 1.67 | 46.52/46.32 | 37.01/36.78 |
| pptx_small_file_get_shape | 1 | 2.66 | 1.17 | 0.43 | 6.29/6.21 | 2.71/2.72 |
| pptx_small_file_first_shape_contains | 1 | 20.35 | 1.15 | 0.42 | 48.71/48.88 | 2.73/2.74 |
| pptx_small_file_limit10_shape_width | 10 | 371.60 | 17.87 | 4.91 | 76.33/75.94 | 3.57/3.66 |
| pptx_small_file_all_shape_width | 10 | 466.21 | 76.43 | 5.00 | 93.84/93.25 | 15.07/15.30 |
| pptx_small_file_all_shape_contains | 10 | 169.95 | 72.25 | 5.07 | 33.70/33.48 | 14.07/14.29 |
| pptx_small_file_all_shape | 90 | 144.98 | 88.45 | 5.00 | 29.10/29.08 | 17.58/17.72 |
| pptx_large_bytes_get_shape | 1 | 1.00 | 0.99 | 0.10 | 10.09/10.10 | 9.98/10.06 |
| pptx_large_bytes_first_shape_contains | 1 | 185.63 | 0.99 | 0.09 | 2131.39/2100.00 | 11.41/11.30 |
| pptx_large_bytes_limit10_shape_width | 10 | 1488.85 | 47.27 | 4.87 | 303.34/309.22 | 9.55/9.91 |
| pptx_large_bytes_all_shape_width | 500 | 72034.69 | 3198.98 | 292.55 | 251.32/195.24 | 11.75/10.49 |
| pptx_large_bytes_all_shape_contains | 500 | 3170.39 | 2176.53 | 311.61 | 10.92/9.33 | 7.38/6.74 |
| pptx_large_bytes_all_shape | 20500 | 3941.09 | 2945.80 | 285.10 | 14.20/12.98 | 10.87/10.08 |
| pptx_large_file_get_shape | 1 | 3.07 | 1.56 | 0.43 | 7.05/7.09 | 3.64/3.65 |
| pptx_large_file_first_shape_contains | 1 | 196.40 | 1.52 | 0.42 | 467.50/465.42 | 3.63/3.61 |
| pptx_large_file_limit10_shape_width | 10 | 1768.07 | 52.50 | 8.22 | 215.00/216.87 | 6.29/6.41 |
| pptx_large_file_all_shape_width | 500 | 87847.54 | 3469.50 | 471.27 | 178.45/178.91 | 7.84/7.64 |
| pptx_large_file_all_shape_contains | 500 | 17576.94 | 2437.95 | 471.85 | 37.33/37.84 | 5.16/5.15 |
| pptx_large_file_all_shape | 20500 | 12001.13 | 6540.31 | 448.57 | 26.27/27.54 | 14.65/14.36 |

Reading the table against the targets:
- **Warm `get` does not regress** on any route; it improves on every XLSX
  and file-backed route (e.g. XLSX 50k file 3.1× → 1.8×).
- **Early termination** is O(1) in document size: `first()` on 1k and 50k
  rows costs the same (0.86 µs), and on PPTX 10 or 500 slides 0.6–1.0 µs
  (was 13–186 µs).
- **Against record 0248/R4b** (cells 33×, rows 1.8–2.7×; PPTX 4.5–15.6×):
  XLSX full cell queries now 10.5–13.7× on memory sources (were 25–44×; one
  before run's leaf leg was noisy at 5.6×),
  rows 2.3× (within R4b); PPTX full queries 6.7–11.8× on large decks (were
  9–251×). **Still a miss against a +10% trigger**: what remains is the
  generic per-node cost (one `Arc` grid clone per cell, node moves through
  the walker, ~50 ns per XLSX cell against ~5 ns for the leaf scan) and,
  on file sources, the one per-match check (`fstat`), which the
  source-change contract requires (XLSX 50k file `cell`: 41×).
- DOCX is unchanged within noise (1.07–1.6×; its adapter was not touched,
  agent p7a owns it).

### XLS reads (`examples/xls_read_bench.rs`, record 0259's bench)

Two before (trunk with 0259) and two after runs, nine samples each,
ratio unified/native:

| Cells | Owner | get before | get after | scan before | scan after |
|---:|---|---:|---:|---:|---:|
| 1 | eager | 3.46–3.56 | 3.12–3.19 | 6.10–6.16 | 4.84–4.93 |
| 100 | eager | 4.60 | 3.99–4.10 | 2.81–2.93 | 2.13–2.23 |
| 1,000 | eager | 4.13–4.17 | 3.58–3.70 | 2.76–2.78 | 2.03–2.09 |
| 10,000 | eager | 3.52 | 3.07–3.17 | 2.57–2.62 | 1.96–2.03 |
| 1 | positional | 1.61–1.62 | 1.49 | 2.20–2.23 | 1.88–1.89 |
| 100 | positional | 1.85–1.87 | 1.70–1.71 | 6.03–6.33 | 4.49–4.59 |
| 1,000 | positional | 1.82–1.84 | 1.68–1.69 | 6.84–7.00 | 5.00–5.07 |
| 10,000 | positional | 1.62–1.63 | 1.50–1.51 | 6.78–7.14 | 5.06–5.17 |

Still a miss. Remaining: the XLS node (it no longer stores its two path
steps, 112 bytes, but keeps a `Label` and a `CellSource` with an inline
`CellData`),
number formatting in the view, and the per-match file check. Record 0259's
item (c), early termination on a cold worksheet, is not fixed: every
worksheet adapter (XLSX, XLSB, XLS) materializes the sorted grid through its
leaf before the first match. Warm walks stop at the first match (above), but
a cold first match still parses the sheet; that needs a streaming cell
cursor in each leaf (follow-up F2).

### PPTX run formatting (0252 harness, routes `pptx_*_run_format_v1`)

B A A B, five ABBA rounds each, unified/native:

| shapes | eager before | eager after | source before | source after |
|---:|---:|---:|---:|---:|
| 1 | 1.035 / 1.064 | 1.029 / 1.018 | 1.051 / 1.069 | 1.061 / 1.085 |
| 10 | 1.069 / 1.076 | 1.009 / 1.021 | 1.015 / 0.990 | 1.005 / 1.030 |
| 100 | 1.153 / 1.151 | 1.007 / 1.011 | 0.997 / 0.999 | 1.006 / 0.996 |
| 1,000 | 1.177 / 1.183 | 1.007 / 1.005 | 0.995 / 0.994 | 0.999 / 0.992 |
| 10,000 | 1.170 / 1.183 | 1.008 / 1.004 | 0.993 / 0.997 | 1.005 / 0.998 |

**Met** (≤ 10%) at every size for both owners. The harness sets on the same
prepared handle each sample. A chain of sets on one slide (each on the
handle the previous set returned) still reads the changed slide once per
set, because its payload is new; that pattern was not measured.

### PPTX metadata (`ooxml_metadata_bench`)

| Scenario | before | after |
|---|---|---|
| PPTX set, source open (repeated on one handle) | 2.42–2.45× | 1.006–1.013× |
| PPTX set, eager open | 1.011–1.014× | 1.006–1.015× |
| PPTX open + first set (fresh source handle vs leaf eager open + edit) | — | 1.31× (490 vs 375 µs) |

The repeated-set miss is fixed. The first set on a fresh source handle
still pays the one-time materialization (≈125 µs). A source-backed metadata
edit in `litchi-pptx` that returns a source revision was examined and not
built: the properties engine reads source-backed (`Snapshot::read_source_backed`,
20 µs; a full source edit computes its patch in 41 µs), but publishing it
needs a part-overlay or topology revision in `source::Package` for the
`docProps` parts and the package relationships, which only structural slide
edits have today. Follow-up F3.

Binaries (not committed), SHA-256:
- `unified-query` before `1155a0fa…4461`, after `262652da…82d3`
  (`CARGO_TARGET_DIR=target/bench cargo build --release --manifest-path
  tools/perf-baseline/unified-query/Cargo.toml`);
- `unified-pptx-text-format` before `0f2b0311…3780`, after `0a32b2bf…2492`;
- `ooxml_metadata_bench` before `10d2f22b…d5e1`, after `603195e8…c908`
  (`cargo build --release -p litchi-ops --all-features --example
  ooxml_metadata_bench`);
- `xls_read_bench` before `af86a661…e484`, after `2209da18…4999`
  (`cargo build --release -p litchi-ops --features xls --example
  xls_read_bench`).

## 4. Semantic deltas

- Result sets: none (§2).
- **Error path of PPTX property filters.** A filter on one shape-tree key
  now reads only that key's inputs. A shape whose text body (or transform,
  or `cNvPr`) is malformed no longer fails a query that filters on a key not
  using that input (`shape[width>=10cm]` reads the scene bounds only);
  `format` on that node, and filters on the affected keys, still report the
  error. Before, the whole property set was decoded per candidate, which
  made a geometry filter cost 3.5 µs per shape. This follows the walker's
  documented contract (one property read per candidate) and the owner's
  2026-09-16 decision 8 (the timing of an error return may move); flagged
  for the coordinator.
- Timing of freshness checks: one check after each step instead of before
  and after each nested read; the observable results are the same.

## 5. Gaps and follow-ups

- F1 — per-node cost on full scans (XLSX grid `Arc` per cell, walker node
  moves), and the per-match `fstat` on file sources: needs a cheaper
  source-version probe in `litchi-core`/`litchi-opc`, or an owner decision
  to fence per batch.
- F2 — streaming first match on a cold worksheet (XLSX, XLSB, XLS leaves).
- F3 — a `source::Package` metadata revision in `litchi-pptx`, so a fresh
  source handle's first metadata set needs no materialization.
- F4 — the XLSB node still stores its path steps (the XLSB agent owns the
  adapter; the same `NodePath::short` change applies), and the DOCX node is
  now the largest generic node (176 bytes).
- F5 — chained same-slide eager sets re-read the changed slide.
- The `litchi-xlsx` test `form_control_owner_read` needs a file excluded by
  the worktree's sparse checkout; clippy and tests of `litchi-xlsx` were run
  on `--lib` here.
