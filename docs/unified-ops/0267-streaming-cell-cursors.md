# 0267 — Streaming worksheet cell cursors, small XLSB nodes, fence analysis

- Status: implemented on `uo2/stream-cursor`, cut from `feat/unified-ops-api`
  at `4d1afe92bc`. Closes follow-ups F2 and F4 of record
  [0262](0262-query-walker-performance.md) and item (c) of record
  [0259](0259-xls-unified-p0.md) §"Read overhead"; analyses F1.
- Date: 2026-10-13
- Scope: new cursor modules `litchi-xlsx/src/workbook/source/cursor.rs`,
  `litchi-xlsb/src/workbook/source/cursor.rs`,
  `litchi-xls/src/workbook/source/cursor.rs`, with small hooks in their
  owners (listed in §1); `litchi-ops/src/query/{inspect,walk}.rs`; the
  XLSX, XLSB and XLS `nav`/`node`/`children` adapters; new routes in
  `tools/perf-baseline/unified-query`; tests.
- `performance_claim`: paired measurements in §4 (release, `taskset -c 28`,
  B A A B across processes, each run an in-process ABBA). Evidence:
  [results/0267/](results/0267/).

## 1. What changed

### Leaf cursors

Each leaf gains `cell_cursor()` on its source-backed worksheet handle: a
lazy iterator over the stored cells in row-major order that decodes only as
many cells as the caller reads. All three return exactly the cells of the
worksheet's sorted grid, in its order; §2 lists the one remaining
difference, which concerns refusals only.

- **XLSX** (`SourceWorksheet::cell_cursor -> CellCursor`, items
  `(Address, RecordIndex)`, `grid()`, `is_complete()`). The cursor parses
  *prefix projections* of the worksheet part: the bytes through the
  `sheetData` start tag, the first N complete rows, and the two end tags
  that close the document, with the parser, shared strings and style
  repair the grid uses (the same technique as the `blocks.rs` projection).
  The first prefix holds 64 rows, each later one 8× more; past an eighth of
  the part the cursor reads the complete (cached) grid and continues after
  the last coordinate it returned. Before the first prefix one byte scan
  checks that rows ascend (`rows_ascend`); a cut that is not between two
  rows makes the prefix parse fail; a prefix whose parse records a repair
  is not used; in every such case the cursor reads the complete grid. The
  largest prefix is offered to the workbook's weighted value cache
  (`Axis::Prefix`, a longer prefix replaces a shorter one), so a later
  cursor on the same workbook starts from it. Hooks: `values.rs` (the
  prefix axis), `source.rs` / `blocks.rs` (module and match arms).
- **XLSB** (`SourceBackedWorksheet::cell_cursor -> CellCursor`, re-exported
  as `litchi_xlsb::sheet::CellCursor`, items `(u32, u32)`, `grid()`,
  `is_complete()`). Before the first cell the cursor raises every refusal
  the grid's construction raises before it decodes cells (the
  materialization inputs, now shared with `parse_materialized` through
  `materialization_inputs`), parses the scenarios and comments as the grid
  does, and walks the record headers to check that the cell records
  ascend. It then decodes batches of 64, 512 and 4,096 cells with the
  grid's `CellsReader` (each batch re-reads the stream and skips the cells
  already returned) and afterwards reads the complete grid. The first
  batch is kept on the sheet's metadata (`first_batch`, a few dozen cells),
  so a later cursor reads no part.
- **XLS** (`SourceBackedWorksheet::cell_cursor -> Result<Option<CellCursor>>`,
  items `SourceBackedCell`, `needs_formula_context()`). A validating scan
  first runs the whole-sheet scan on the measure path (every check the
  parse runs, no value decoded), noting whether positions strictly ascend
  and whether a formula renders no text without the name tables; a sheet
  whose positions do not ascend gets no cursor (`None`). The cursor then
  runs the same scan in batches of 64 cells that *resume* where the last
  one stopped (`ScanResume`: stream position plus the byte and record
  counters, so the worksheet limits count from the sheet start as for one
  scan). The verdict is kept per sheet (`cursor_checks`). Hooks in
  `source.rs`: `CellSink::{stop, wants_formula}`, `scan_worksheet_from`,
  `WorksheetScan::resume`.

### Query walker and adapters

- `Inspect::early_children` (default: `children`): the walker calls it
  when the walk is likely to stop early — `limit(n)`/`first()`, or every
  top-level alternative selects its `[N]`th match.
- XLSX, XLSB and XLS implement it for a source-backed worksheet whose cells
  alone (and, for XLSX/XLSB, tables) are wanted. Cells streamed from an
  XLSX prefix or an XLSB batch keep that grid for their own record and
  read merges, rows and formats from the complete grid on first use
  (`GridInner::Prefix`, `GridInner::complete`). XLS cells carry their data
  (`CellSource::Direct`); a sheet whose formulas need the name tables loads
  the eager worksheet before the first cell, as the grid does. Any cursor
  creation refusal falls back to the ordinary walk, which owns the refusal.
- The walker's `next` moves the match out once (no `Hit` wrapper).

### Small XLSB nodes (F4)

The XLSB node writes its path from its item (as record 0262 did for XLSX),
keeps labels lazy, shares its grid and tab handle behind one `Arc` (a
listed cell no longer clones the tab handle and its name), boxes the sheet
item and direct targets, and reads `Inspect::cell_at` directly:
**360 → 88 bytes**, asserted ≤ 96.

## 2. Differential evidence and semantic deltas

`tests/query_corpus_differential.rs` now also digests the XLS (126 files,
119 opened) and XLSB (18) corpora with the XLSX selectors, and adds, for
the three workbook formats, a cold `limit(3)` digest per selector on a
freshly opened file handle and a cold `first()` digest on a freshly opened
in-memory handle (396 digests). That golden was printed by trunk
`4d1afe92bc` before any change (commit `b6f6405d5b`) and the final build
reproduces it byte for byte. `tests/query_early_stop.rs` holds every
early-stopping query (13 selectors × limits 1, 3, 200, 5,000 × memory and
file × cold and warm, plus `first()`) to the full query's leading matches,
with each match's properties, on 3,000-row generated workbooks that stream
in all three formats. Leaf tests compare each cursor with its grid
(XLSX: every stage, fallbacks for comments, out-of-order and repeated rows,
prefixed markup; XLSB and XLS: the whole corpus).

Ordering is exact in all three formats (the gates above). **The one
difference**, observable only by a reader that stops early, is the timing
of a refusal located *after* the cells read: the grid refuses the
worksheet before its first cell, the cursor returns the leading cells and
reports the refusal when a later stage reaches it. This covers, per
format: XLSX — any parse refusal or whole-part limit (e.g. the stream
event limit) after the prefix; XLSB — a cell record the reader refuses
after the batch; XLS — a shared-string index or `STRING` result that only
value decoding refuses (framing, parse and XF checks are validated up
front). This is the class the owner accepted for the XLSX read gate on
2026-09-16 (decision 8, "the timing of the error return moves") and is
flagged for the coordinator; no corpus file shows it. **Owner decision
(2026-10-14): accepted** — the delayed damage report stands.

## 3. The per-match source check on file sources (F1)

The walker checks the source once per pulled match (record 0262 fence
scopes). On a file source that is one `FileSource::version()` per match: a
mutex and a `statx` (one per match, measured with `strace`: 3,010 calls
for a 3,003-match query). Batching the check — reading K matches, checking
once, then yielding them — keeps one half of the guarantee (no yielded
match was read from a changed source) but not the other: a pull after a
change must fail (`lazy_file_query_and_cached_nodes_refuse_changed_source`,
`retained_nodes_and_lazy_reads_refuse_changed_sources` pull one match,
change the source and require the next pull to fail). A batch of K ≥ 2
answers that pull from its buffer, so the check frequency cannot change
without an owner decision; it is unchanged. On memory sources the check is
not a cost: removing it moves `xlsx_50000_bytes_all_cell` by less than the
noise (§4). The remaining file cost is the probe itself: `perf` puts about
18% of that route in `FileSource::version`'s own code (mutex, `Metadata`
construction) and about 30% in the `statx` call; a lock-free fingerprint
compare in `litchi-core` would remove the first part (follow-up F1').

## 4. Measurements

Harness `tools/perf-baseline/unified-query`: the before binary is trunk
`4d1afe92bc` with the new routes (`c1ffa128…58a8`), the after binary the
final branch (`e86ac810…68e4f`), both `CARGO_TARGET_DIR=target/bench cargo
build --release --manifest-path tools/perf-baseline/unified-query/Cargo.toml`.
Four processes, B A A B, each `time <family> 9` (nine in-process ABBA
samples per route) on `taskset -c 28` of the shared 32-core host. Unified
and leaf columns: the lower median of the two runs; ratio = unified / leaf
per run. The 50,000-row warm XLSB routes were not run: on that size the
XLSB worksheet cache does not retain the grid (its weight), so every warm
read re-materializes it on both binaries (≈40 ms per call). Raw lines:
`results/0267/query-*.jsonl`.

### 4.1 Cold first match (new routes: a fresh handle per call)

The leaf leg opens a fresh leaf workbook and stops at the same match with
the best trunk leaf read: XLSX `visit_cells_until` (its selected-stream
route validates the whole part first), XLSB the materialized grid, XLS from
a file the positional visitor stopped by its first match (no validation of
the rest), from memory the eager reader.

| route | matches | before unified µs | after unified µs | leaf µs (after) | ratio before | ratio after |
|---|---:|---:|---:|---:|---:|---:|
| xlsx_1000_bytes_cold_first_cell_value_gt | 1 | 976.83 | 157.45 | 6578.03 | 0.15/0.15 | 0.02/0.02 |
| xlsx_1000_bytes_cold_limit10_cell_value_gt | 10 | 989.45 | 161.72 | 6564.66 | 0.15/0.15 | 0.02/0.02 |
| xlsx_1000_file_cold_first_cell_value_gt | 1 | 993.05 | 170.38 | 6593.07 | 0.15/0.15 | 0.03/0.03 |
| xlsx_1000_file_cold_limit10_cell_value_gt | 10 | 1001.28 | 176.48 | 6590.53 | 0.15/0.15 | 0.03/0.03 |
| xlsx_50000_bytes_cold_first_cell_value_gt | 1 | 48626.64 | 2316.44 | 326075.96 | 0.15/0.15 | 0.01/0.01 |
| xlsx_50000_bytes_cold_limit10_cell_value_gt | 10 | 48820.25 | 2335.61 | 326108.11 | 0.15/0.15 | 0.01/0.01 |
| xlsx_50000_file_cold_first_cell_value_gt | 1 | 48820.59 | 2358.53 | 326165.49 | 0.15/0.15 | 0.01/0.01 |
| xlsx_50000_file_cold_limit10_cell_value_gt | 10 | 48857.69 | 2384.15 | 326603.41 | 0.15/0.15 | 0.01/0.01 |
| xls_1000_bytes_cold_first_cell_value_gt | 1 | 727.98 | 725.15 | 722.63 | 1.00/1.01 | 1.00/1.00 |
| xls_1000_bytes_cold_limit10_cell_value_gt | 10 | 732.43 | 728.76 | 721.28 | 1.01/1.02 | 1.00/1.01 |
| xls_1000_file_cold_first_cell_value_gt | 1 | 356.46 | 70.78 | 14.90 | 24.70/24.28 | 4.74/4.75 |
| xls_1000_file_cold_limit10_cell_value_gt | 10 | 365.25 | 78.91 | 22.31 | 16.87/16.70 | 3.39/3.54 |
| xls_50000_bytes_cold_first_cell_value_gt | 1 | 47200.75 | 46658.51 | 46831.11 | 1.00/1.01 | 1.00/1.00 |
| xls_50000_bytes_cold_limit10_cell_value_gt | 10 | 47431.56 | 46791.51 | 46470.70 | 1.00/1.00 | 0.99/1.00 |
| xls_50000_file_cold_first_cell_value_gt | 1 | 16621.05 | 2092.85 | 32.49 | 416.41/422.69 | 64.42/64.88 |
| xls_50000_file_cold_limit10_cell_value_gt | 10 | 16646.44 | 2108.37 | 40.14 | 350.32/348.54 | 51.16/52.90 |
| xlsb_1000_bytes_cold_first_cell_value_gt | 1 | 508.06 | 122.72 | 480.73 | 1.07/1.05 | 0.25/0.26 |
| xlsb_1000_bytes_cold_limit10_cell_value_gt | 10 | 513.41 | 130.18 | 484.27 | 1.08/1.08 | 0.27/0.27 |
| xlsb_1000_file_cold_first_cell_value_gt | 1 | 539.00 | 150.63 | 499.61 | 1.10/1.08 | 0.30/0.30 |
| xlsb_1000_file_cold_limit10_cell_value_gt | 10 | 587.17 | 199.55 | 497.71 | 1.20/1.19 | 0.40/0.40 |
| xlsb_50000_bytes_cold_first_cell_value_gt | 1 | 38697.38 | 2188.03 | 37785.37 | 1.02/1.01 | 0.06/0.06 |
| xlsb_50000_bytes_cold_limit10_cell_value_gt | 10 | 38929.71 | 2204.14 | 37756.99 | 1.02/1.00 | 0.06/0.06 |
| xlsb_50000_file_cold_first_cell_value_gt | 1 | 38598.74 | 2224.98 | 38639.59 | 1.02/1.02 | 0.06/0.06 |
| xlsb_50000_file_cold_limit10_cell_value_gt | 10 | 38526.15 | 2305.83 | 38249.89 | 1.02/1.02 | 0.06/0.06 |

- XLSX: **21×** faster on 50,000 rows (48.6 → 2.3 ms), 6× on 1,000; the
  remainder is mostly whole-part inflate and CRC (G1) and the row-order
  scan (G3).
- XLSB: **17.7×** on 50,000 rows (38.7 → 2.2 ms), 4× on 1,000.
- XLS from a file: **7.9×** on 50,000 rows (16.6 → 2.1 ms), 5× on 1,000;
  the validating scan is most of it. The leaf leg stops without validating
  the rest of the sheet, so the 64× ratio compares unlike work. XLS from
  memory is an eager owner whose open parses the workbook; nothing to
  stream (G4).

### 4.2 Warm queries and gets (record 0262 routes, plus warm XLSB)

| route | matches | before unified µs | after unified µs | leaf µs (after) | ratio before | ratio after |
|---|---:|---:|---:|---:|---:|---:|
| xlsx_1000_bytes_get_cell | 1 | 0.43 | 0.41 | 0.16 | 2.64/2.68 | 2.56/2.53 |
| xlsx_1000_bytes_first_cell_value_gt | 1 | 0.87 | 0.79 | 0.14 | 6.04/6.05 | 5.51/5.51 |
| xlsx_1000_bytes_limit10_cell_value_gt | 10 | 5.51 | 5.49 | 0.44 | 12.67/12.76 | 12.56/12.87 |
| xlsx_1000_bytes_all_cell_value_gt | 500 | 247.67 | 241.19 | 16.75 | 15.39/14.79 | 13.81/15.02 |
| xlsx_1000_bytes_all_cell | 3003 | 159.72 | 149.47 | 8.19 | 24.57/19.40 | 16.45/20.52 |
| xlsx_1000_file_get_cell | 1 | 1.17 | 1.13 | 0.65 | 1.80/1.81 | 1.71/1.75 |
| xlsx_1000_file_first_cell_value_gt | 1 | 2.09 | 1.29 | 0.63 | 3.32/3.33 | 2.04/2.05 |
| xlsx_1000_file_limit10_cell_value_gt | 10 | 8.20 | 7.56 | 0.93 | 8.90/8.91 | 8.09/8.18 |
| xlsx_1000_file_all_cell_value_gt | 500 | 335.22 | 324.49 | 17.20 | 19.51/19.49 | 18.64/19.56 |
| xlsx_1000_file_all_cell | 3003 | 687.72 | 647.32 | 8.80 | 80.16/75.04 | 73.56/65.10 |
| xlsx_50000_bytes_get_cell | 1 | 0.45 | 0.44 | 0.18 | 2.48/2.49 | 2.41/2.39 |
| xlsx_50000_bytes_first_cell_value_gt | 1 | 0.87 | 0.79 | 0.15 | 5.64/5.72 | 5.15/5.24 |
| xlsx_50000_bytes_limit10_cell_value_gt | 10 | 5.50 | 5.49 | 0.45 | 12.39/12.24 | 12.22/12.52 |
| xlsx_50000_bytes_all_cell_value_gt | 25000 | 12518.05 | 12068.11 | 1056.86 | 12.16/11.56 | 11.34/12.05 |
| xlsx_50000_bytes_all_cell | 150003 | 7981.04 | 7449.60 | 728.95 | 15.92/6.78 | 10.23/11.48 |
| xlsx_50000_bytes_all_row_header_gt | 25000 | 12044.16 | 10143.41 | 4795.81 | 2.00/2.12 | 1.95/2.17 |
| xlsx_50000_file_get_cell | 1 | 1.19 | 1.16 | 0.67 | 1.77/1.80 | 1.70/1.72 |
| xlsx_50000_file_first_cell_value_gt | 1 | 2.10 | 1.32 | 0.64 | 3.32/3.29 | 2.06/2.05 |
| xlsx_50000_file_limit10_cell_value_gt | 10 | 8.20 | 7.59 | 0.94 | 8.89/8.95 | 8.10/8.12 |
| xlsx_50000_file_all_cell_value_gt | 25000 | 16439.13 | 16320.76 | 1065.85 | 15.44/15.93 | 15.05/15.75 |
| xlsx_50000_file_all_cell | 150003 | 33175.48 | 33624.84 | 860.22 | 36.97/41.51 | 39.19/37.42 |
| xlsx_50000_file_all_row_header_gt | 25000 | 41376.81 | 40066.53 | 5819.18 | 8.02/6.73 | 6.71/6.96 |
| docx_drawing_bytes_get_paragraph | 1 | 0.97 | 0.96 | 0.48 | 2.06/2.04 | 2.00/2.01 |
| docx_drawing_bytes_first_run_bold | 1 | 7.79 | 7.95 | 7.09 | 1.12/1.12 | 1.12/1.12 |
| docx_drawing_bytes_all_run_bold | 177 | 4326.62 | 4247.79 | 3959.99 | 1.07/1.07 | 1.07/1.07 |
| docx_drawing_file_get_paragraph | 1 | 1.83 | 1.84 | 1.19 | 1.54/1.53 | 1.55/1.55 |
| docx_drawing_file_first_run_bold | 1 | 9.72 | 9.58 | 7.62 | 1.25/1.26 | 1.26/1.26 |
| docx_drawing_file_all_run_bold | 177 | 5070.78 | 4992.27 | 3951.61 | 1.26/1.26 | 1.27/1.26 |
| docx_saut_bytes_get_paragraph | 1 | 0.97 | 0.94 | 0.47 | 2.12/2.09 | 1.99/2.03 |
| docx_saut_bytes_first_run_bold | 0 | 20.26 | 20.20 | 17.06 | 1.19/1.18 | 1.18/1.19 |
| docx_saut_bytes_all_run_bold | 0 | 20.19 | 20.36 | 16.93 | 1.17/1.20 | 1.19/1.21 |
| docx_saut_file_get_paragraph | 1 | 1.83 | 1.83 | 1.19 | 1.56/1.54 | 1.55/1.56 |
| docx_saut_file_first_run_bold | 0 | 28.09 | 27.66 | 17.33 | 1.56/1.61 | 1.60/1.57 |
| docx_saut_file_all_run_bold | 0 | 28.31 | 27.53 | 17.29 | 1.59/1.61 | 1.58/1.58 |
| pptx_small_bytes_get_shape | 1 | 0.59 | 0.56 | 0.10 | 6.02/6.00 | 5.72/5.68 |
| pptx_small_bytes_first_shape_contains | 1 | 0.59 | 0.58 | 0.09 | 6.78/6.85 | 7.00/6.65 |
| pptx_small_bytes_limit10_shape_width | 10 | 12.84 | 12.79 | 1.57 | 8.17/8.21 | 8.20/8.14 |
| pptx_small_bytes_all_shape_width | 10 | 62.44 | 61.81 | 1.65 | 37.76/37.91 | 38.19/37.34 |
| pptx_small_bytes_all_shape_contains | 10 | 57.41 | 57.39 | 1.72 | 33.37/33.59 | 33.56/33.33 |
| pptx_small_bytes_all_shape | 90 | 60.52 | 58.07 | 1.67 | 36.23/36.61 | 35.48/34.81 |
| pptx_small_file_get_shape | 1 | 1.15 | 1.12 | 0.43 | 2.70/2.68 | 2.61/2.60 |
| pptx_small_file_first_shape_contains | 1 | 1.12 | 1.11 | 0.41 | 2.67/2.70 | 2.73/2.69 |
| pptx_small_file_limit10_shape_width | 10 | 17.84 | 17.85 | 4.85 | 3.67/3.69 | 3.66/3.69 |
| pptx_small_file_all_shape_width | 10 | 74.91 | 74.74 | 4.96 | 14.88/15.19 | 15.14/15.07 |
| pptx_small_file_all_shape_contains | 10 | 71.01 | 70.26 | 5.06 | 14.04/14.33 | 13.90/13.90 |
| pptx_small_file_all_shape | 90 | 87.33 | 84.53 | 4.97 | 17.59/17.76 | 17.16/16.91 |
| pptx_large_bytes_get_shape | 1 | 0.97 | 0.92 | 0.10 | 9.83/9.92 | 9.40/9.39 |
| pptx_large_bytes_first_shape_contains | 1 | 0.95 | 0.93 | 0.09 | 10.96/10.95 | 11.19/10.76 |
| pptx_large_bytes_limit10_shape_width | 10 | 49.03 | 48.15 | 4.86 | 10.10/10.12 | 10.02/9.91 |
| pptx_large_bytes_all_shape_width | 500 | 3303.33 | 3258.53 | 273.98 | 12.42/12.00 | 12.49/11.88 |
| pptx_large_bytes_all_shape_contains | 500 | 2171.26 | 2163.34 | 291.44 | 7.31/7.37 | 7.42/7.35 |
| pptx_large_bytes_all_shape | 20500 | 3015.88 | 2558.84 | 275.85 | 10.93/10.94 | 9.66/9.27 |
| pptx_large_file_get_shape | 1 | 1.53 | 1.49 | 0.43 | 3.56/3.57 | 3.48/3.49 |
| pptx_large_file_first_shape_contains | 1 | 1.47 | 1.46 | 0.41 | 3.50/3.59 | 3.56/3.55 |
| pptx_large_file_limit10_shape_width | 10 | 53.97 | 53.80 | 8.15 | 6.60/6.70 | 6.65/6.58 |
| pptx_large_file_all_shape_width | 500 | 3582.41 | 3545.44 | 441.66 | 8.25/8.12 | 8.33/8.04 |
| pptx_large_file_all_shape_contains | 500 | 2443.28 | 2439.43 | 461.69 | 5.19/5.29 | 5.28/5.30 |
| pptx_large_file_all_shape | 20500 | 6537.72 | 6131.92 | 443.54 | 14.69/14.78 | 14.08/13.83 |
| xlsb_1000_bytes_get_cell | 1 | 0.67 | 0.56 | 0.09 | 7.16/7.17 | 6.24/6.15 |
| xlsb_1000_bytes_first_cell_value_gt | 1 | 1.63 | 1.09 | 0.13 | 12.19/12.23 | 8.13/8.11 |
| xlsb_1000_bytes_all_cell_value_gt | 500 | 587.62 | 436.72 | 10.69 | 58.93/55.74 | 40.56/40.81 |
| xlsb_1000_bytes_all_cell | 3003 | 461.83 | 173.05 | 11.56 | 43.03/39.87 | 15.30/14.92 |
| xlsb_1000_file_get_cell | 1 | 3.17 | 3.07 | 0.93 | 3.39/3.43 | 3.33/3.30 |
| xlsb_1000_file_first_cell_value_gt | 1 | 7.04 | 6.95 | 0.96 | 7.21/7.26 | 7.24/7.23 |
| xlsb_1000_file_all_cell_value_gt | 500 | 2730.73 | 2578.15 | 12.13 | 236.73/238.61 | 212.59/213.16 |
| xlsb_1000_file_all_cell | 3003 | 1433.92 | 1170.12 | 12.49 | 140.79/135.12 | 93.71/97.21 |

Warm `get` does not regress on any route (XLSX 0.41–1.16 µs vs
0.43–1.19 µs; DOCX, PPTX within noise; XLSB 0.67 → 0.56 µs from the small
node). Warm early-stopping queries on file-backed XLSX handles improve
(`first` 2.09 → 1.29–1.32 µs, measured; the walk now reaches the resident
grid through the cursor rather than the sheet node's pin), on XLSB from memory too (1.63 → 1.09 µs). XLSB full
scans gain from the small node (`all_cell` from memory 462 → 173 µs,
43× → 15×; from a file 1,434 → 1,170 µs).

### 4.3 Full scans

XLSX, DOCX and PPTX full-scan rows are unchanged within noise (the
`*_bytes_all_cell` leaf leg itself varies 0.67–0.86 ms between runs on the
loaded host). A profile of `xlsx_50000_bytes_all_cell` splits the
unified cost between the adapter's step (node construction, one `Arc`
increment per cell), node drop, and moves of the 184-byte unified node,
with no single dominant term; on files the per-match check dominates (§3).
Removing the per-match check altogether on the memory route moved it by
less than the run-to-run noise (7.6–7.9 ms either way, ABBA), so on memory
sources the check is not the cost. PPTX full queries spend most of their
time outside the walker: about a fifth of the profile re-reads slide scenes
in the PPTX adapter (`strict_scene`, the shape scanner, XML limit
validation, `quick_xml`), outside this record's files.

### 4.4 XLS reads (`examples/xls_read_bench.rs`, record 0259's bench)

Before `9c742e30…d277d` (trunk XLS code), after `c97e5a7d…2a76`, B A A B,
nine samples each; `cargo build --release -p litchi-ops --features xls
--example xls_read_bench`. Raw: `results/0267/xls-*.json`.

| Cells | Owner | get before | get after | scan before | scan after | unified get ns before→after | unified scan ns before→after |
|---:|---|---:|---:|---:|---:|---:|---:|
| 1 | eager | 3.07–3.08 | 3.00–3.01 | 4.44–4.58 | 5.00–5.12 | 218→213 | 275→280 |
| 100 | eager | 4.04–4.07 | 3.90–4.06 | 2.12 | 2.06–2.16 | 23,845→23,162 | 10,533→10,574 |
| 1,000 | eager | 3.58–3.59 | 3.43–3.56 | 2.03–2.05 | 1.98–2.05 | 242,531→237,596 | 106,747→107,410 |
| 10,000 | eager | 3.05–3.06 | 2.97–3.06 | 1.97–2.00 | 1.93–2.01 | 2,484,962→2,456,501 | 1,120,465→1,122,925 |
| 1 | positional | 1.45–1.46 | 1.46–1.47 | 1.90 | 1.91–1.96 | 1,824→1,800 | 2,167→2,165 |
| 100 | positional | 1.67 | 1.67 | 4.44–4.47 | 4.49 | 186,110→184,122 | 52,050→51,760 |
| 1,000 | positional | 1.65 | 1.65 | 4.95–4.97 | 5.05–5.08 | 1,892,256→1,874,093 | 504,146→503,733 |
| 10,000 | positional | 1.48–1.49 | 1.48 | 5.05–5.16 | 5.01–5.10 | 23,406,308→23,013,416 | 5,190,374→5,165,304 |

Unchanged: a full `cell` scan does not stop early, so it does not use the
cursor (the 1-cell eager scan's 4.5 → 5.0 is a 5 ns move on a 60 ns
denominator). The XLS early-stopping gain is in 4.1.

## 5. Gaps and follow-ups

- **G1 — whole-part decode (XLSX, XLSB).** The cursors read the part's
  verified decoded payload, as the grid does; inflate and CRC are now most
  of a cold first match on the 50,000-row XLSX (about 75% of the profile).
  A verified *prefix* decode in `litchi-opc` (decode on demand, verify the
  CRC when the stream is drained or report it with the remainder) would
  remove it; it changes the OPC reader contract, so it is not done here.
- **G2 — refusals after the cells read** (§2): the one semantic delta,
  flagged for the coordinator.
- **G3 — the XLSX row-order scan** reads the whole `sheetData` once per
  cold cursor (≈0.7 ms on 5.5 MB); it is skipped once a prefix is cached.
- **G4 — owners without cursors.** Eager owners (an XLSX or XLSB handle
  after a write, an XLS handle opened from memory, whose open already
  parses the workbook) and scoped ranges (`/Data/A1:C10` as a query scope)
  still read the complete grid. A range *filter* streams under a limit but
  does not stop the walk past its last row.
- **F1' — the per-match file check** (§3): a lock-free fingerprint
  compare in `litchi-core::FileSource::version`, or an owner decision to
  check per batch.
- **Full-scan overhead** (§4.3) is unchanged apart from the XLSB node; the
  remaining XLSX cost per cell is spread over node construction (an `Arc`
  increment), drop and moves of the 184-byte unified node, and the walk's
  bookkeeping, with no single dominant term; the PPTX full queries spend
  most of their time re-reading slide scenes in the PPTX adapter
  (`strict_scene`, XML limit validation), which this record does not own.
- The worktree's sparse checkout was widened to
  `docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/`
  (20 KB) so the `litchi-xlsx` integration tests compile for clippy.
