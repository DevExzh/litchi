# 0287 — Remaining query performance misses

- Status: implemented on `uo2/perf-query2`, cut from `feat/unified-ops-api`
  at `9fe45629ab`. Follows up record
  [0280](0280-shared-query-transaction-performance.md) gaps G1 (leaf probe
  detail), G2, G3 and G5, and its ODT figure.
- Date: 2026-10-14
- Scope: `litchi-ops` `query/{walk,matches,arg}`, `node/path`,
  `format/{doc,rtf,odt,ods}` nodes, navigators and write paths; small leaf
  hunks in `litchi-cfb` (`shared.rs`) and `litchi-xls`
  (`workbook/source.rs` and its index-cache test); paired examples
  (`ods_unified_bench`, `rtf_destination_bench`) and the `unified-query`
  harness (`all-formats` feature).
- `performance_claim`: paired measurements in §3 (release, `taskset -c 31`
  on the shared 32-core host at load 9–25; processes in B A A B order, each
  an in-process ABBA). Deterministic counts are callgrind (`perf` is
  unavailable, `perf_event_paranoid=4`). Evidence:
  [results/0287/](results/0287/), with the driver script
  `results/0287/measure.sh`.

## 1. What changed

1. **Small DOC, RTF and ODT nodes (G2).** The DOC and RTF nodes stored their
   positional paths as six and four inline `PathStep`s (56 bytes each). That
   made the unified `Node` 408 bytes in any build with every format, such as
   the `litchi` facade with all features. They now keep kinds and positions,
   and write the steps only when the path is read. A DOC path deeper than
   four steps is spelled once into a lazily built list, as ODT already did.
   RTF works out its paragraph and run positions from the steps instead of
   storing them, with the same results as before. ODT keeps its deep-path
   steps and an unescaped style name in a single lazy slot. Sizes on 64-bit
   targets, each guarded by a compile-time assertion:

   | node | before | after |
   |---|---:|---:|
   | DOC | 408 | 144 |
   | RTF | 384 | 160 |
   | ODT | 192 | 176 |
   | unified `Node`, every non-iWork format | 408 | 184 |

   The largest nodes now are DOCX (176) and PPTX (168). Neither adapter is
   owned here.
2. **Each match moves once (shared walker).** The walker used to move every
   match through `Result<Option<Found>>` at each layer: `next_found`,
   `advance_checked`, `advance`, `Iterator::next`, and then into the unified
   enum. The match now waits in a slot in the walk state. The step returns
   only small metadata, and `Walk::next_map` wraps the node as it leaves the
   slot. If the source check after the step fails, the slot is cleared.
   `NodePath::inline_with` writes a compact path's steps in place, without a
   staging array.
3. **ODS adapter (G3).**
   - Cell, range, row, column and merge labels are now written on first use
     (`OnceLock<Label>`, as XLSX does). Before, `Label::push_number` alone
     cost about 76 instructions per walked cell.
   - A sheet walk that wants only cells now lists the content cells
     directly. It no longer steps through the column, row, comment and merge
     phases, which return nothing in that case.
   - Coalesced batch planning (0277 P1) used to resolve each coordinate
     through `Sheet::cell`, which scans the row runs from the top. On a
     1,000-row sheet that was about 5.6 K instructions per item. It now goes
     through a navigator that binary-searches cached row-run ends. These are
     built once per sheet and snapshot, and only by that planner. A single
     direct lookup still scans.
4. **RTF single-target fixed costs (G5).**
   - `&QueryArg` now implements `IntoQuery`. The RTF, ODT and ODS mutation
     resolvers compile the selector from a borrow instead of making a deep
     copy first.
   - The RTF coalescing run still admits its positional source index when
     it opens, with the same charges, order and fallback. It now builds the
     index on the first plan that reads it, so writes to cells, headers and
     comments never build it.
   - The RTF child iterator keeps a non-body parent by value instead of
     boxing it.
5. **XLS leaf probes inside one read (G1 detail).** The owner's decision
   stands: the unified layer keeps its per-match check, and none of the
   unified layer's checks were removed. Only leaf probes that a later probe
   of the same leaf call already covers were removed. A warm indexed
   source-backed cell read used to observe the source five times. It now
   observes it twice:
   - The worksheet lookup in `cell_by_index` reads only retained metadata,
     and the cell read observes the source afterwards on every path. A
     missing worksheet still checks before it reports the error.
   - The leading `ensure_current` in `query_cell` is dropped on the indexed
     path. That path publishes nothing, and every one of its reads and its
     trailing check observes the source. The scan path keeps its leading
     check, because it updates the index cache's miss state before it reads.
   - The frame header is read with the new
     `SharedOleStreamCursor::read_exact_deferred`. The payload read that
     follows fences it. Every error between the two reads goes through
     `fence_error`, so a `SourceChanged` error still takes precedence.

No public `litchi-ops` API was added or renamed. `litchi-cfb` gains
`SharedOleStreamCursor::{read_exact_deferred, check_source, fence_error}`.
The query-corpus differential and every golden are unchanged.

## 2. Breakdown before optimizing (callgrind)

- **ODS warm full `cell` query, 1,000 × 10 cells: about 800 instructions per
  cell.** The walker's own step took about 270, the adapter's child iterator
  118, `Label::push_number` 76, `Node::new_cell` 61, the per-match source
  check 38 (two calls), and the unified-enum wrapping 99. After the change it
  is about 610 instructions per cell.
- **ODT full `p` query, 10,000 paragraphs: about 1,900 instructions per
  paragraph.**
  - About 860 of these are the leaf's `Node::text`, which the leaf leg also
    pays.
  - The walker's own share is about 360, the adapter's child step about 130,
    and moves of the 176-byte node through result layers about 280.
  - With the slot and the in-place path, the total drops to about 1,840.
- **XLS positional `get` plus text: thirteen source observations per call.**
  Four are ops checks, in the resolver and `View::text`. Five are in the leaf
  `query_cell` and frame reads, and one is the leaf metadata lookup. The rest
  come from formatting reads.
- **RTF single-target writes, per call (instructions):**

  | route | before | after | over native |
  |---|---:|---:|---:|
  | native | 76.1 K | | |
  | edit-closure set | 81.0 K | | about +4.9 K |
  | query fan-out | 90.9 K | 89.0 K | |

  Of the edit-closure overhead, 2.9 K is coalescing (`Run::new` building the
  index, about 1.25 K). The fan-out overhead is the internal selector copy
  (about 1.9 K), compiling (3.6 K) and the walk (4.2 K). Compiling once into
  a `Plan` removes the compile cost (route F in §3.5).

## 3. Measurements

Binaries are not committed. SHA-256 prefixes and suffixes:

| harness | base | after |
|---|---|---|
| `unified-query` | `a0e165f8…4cad57` | `8ffbf4cf…fa9f5c` |
| `unified-query --features all-formats` | `e1e03e18…0af7dd` | `69c5daf7…37ef8b` |
| `ods_unified_bench` | `9f3f504d…fdb83a` | `a431ffb1…580b44` |
| `odt_paired_bench` | `f6aa7ac1…239637` | `39cd72ea…407179` |
| `rtf_destination_bench` | `bf6a3705…5ee08f` | `ee71d575…77fd50` |
| `xls_read_bench` | `dbb95812…aa1b30` | `61854017…cc5ca4` |

How they were built:

- `unified-query`: `cargo build --release --offline --manifest-path
  tools/perf-baseline/unified-query/Cargo.toml [--features all-formats]`.
- The examples: `cargo build --release -p litchi-ops --no-default-features
  --features F --example …`.
- Base binaries are the base commit (the `all-formats` base uses this
  record's harness `Cargo.toml`). After binaries are `f3101e8602`.

The base ODS and RTF examples have no `LITCHI_ODS_BENCH_ONLY` or plan
route; the common columns are compared. In the tables, columns are the lower
of two runs' medians, ratios are per run, and µs are per call.

### 3.1 Queries (`unified-query`, harness features)

| route | base µs | after µs | leaf µs | ratio base | ratio after |
|---|---:|---:|---:|---:|---:|
| xlsx_1000_bytes_all_cell | 163.9 | 151.1 | 9.2 | 16.8/16.9 | 16.5/14.1 |
| xlsx_1000_bytes_all_cell_owned | 163.8 | 151.0 | 33.8 | 4.85/4.87 | 4.45/4.45 |
| xlsx_1000_bytes_all_cell_fenced | 164.1 | 150.8 | 55.0 | 2.98/3.01 | 2.73/2.74 |
| xlsx_1000_bytes_all_cell_value_gt | 241.5 | 205.5 | 16.9 | 14.1/14.5 | 11.3/11.6 |
| xlsx_1000_file_all_cell_fenced | 543.0 | 528.5 | 432.3 | 1.24/1.23 | 1.24/1.21 |
| pptx_small_bytes_all_shape | 12.2 | 9.8 | 1.7 | 7.82/7.30 | 5.88/6.18 |
| pptx_large_bytes_all_shape | 2252.7 | 1842.0 | 281.0 | 7.85/7.29 | 6.62/5.14 |
| pptx_large_bytes_all_shape_owned | 2256.4 | 1875.6 | 801.5 | 2.66/2.43 | 2.35/1.98 |
| pptx_large_file_all_shape | 5201.2 | 4865.3 | 422.0 | 10.9/11.3 | 11.5/7.07 |
| pptx_large_file_all_shape_fenced | 5272.9 | 4899.6 | 3613.3 | 1.41/1.41 | 1.36/1.07 |
| docx_drawing_file_all_run_bold | 4930.1 | 4905.5 | 3999.8 | 1.23/1.19 | 1.23/1.24 |
| xlsb_1000_file_all_cell | 966.0 | 952.9 | 10.6 | 81.1/92.7 | 77.2/88.0 |
| xls_1000_file_cold_first_cell_value_gt | 68.1 | 70.0 | 14.3 | 4.76/4.76 | 4.83/4.83 |

- Callgrind, 20 queries per route:
  - XLSX 1,000-row full cell query: 50.2 M → 46.0 M instructions.
  - `…value_gt`: 67.5 M → 64.3 M.
  - PPTX 500-slide full shape query: 569.6 M → 478.9 M.
- `get`s, early stops and cold first matches are unchanged within ±3%. The
  full table is `results/0287/query-*.jsonl`.
- The one route that reads slower in both runs is XLS
  `file_cold_first_cell_value_gt`, at 68.1 → 70.0 µs (+2.8%). Its ratio
  moves from 4.76 to 4.83. It is a cold scan whose probes were not changed
  here. It is not explained and is listed in §5.

### 3.2 G2: XLSX in a build with every format (`all-formats`)

Callgrind, 20 queries of `xlsx_1000_bytes_all_cell`:

| build | before | after |
|---|---:|---:|
| `all-formats` | 53.6 M | 48.9 M |
| harness features | 50.2 M | 46.0 M |

The node size accounts for about 7% of the instructions in the
`all-formats` build. In cycles it accounted for more, as wall-clock times
measured during the work show. The figures for full XLSX queries in the
`all-formats` build disagree between the two sessions:

| session | base µs | after µs | owned ratio before | owned ratio after |
|---|---:|---:|---:|---:|
| first (B G G B, during the work) | 225.6 / 227.0 | 153.1 / 155.1 | 6.69 / 6.62 | 4.55 / 4.60 |
| final (B A A B) | 158.3 | 146.6 | 4.70 / 4.68 | 4.64 / 4.30 |

The final session's base did not reproduce the first session's 225 µs. The
callgrind count, 53.6 M → 48.9 M, is the stable figure. In the final
session, `…all_cell_value_gt` read 219 → 239 µs, although callgrind counts
1.3% fewer instructions. This is treated as layout or noise and listed in
§5. PPTX large in the same build: 2759 → 2259 µs, owned ratio
3.19/3.29 → 2.84/2.74.

**0280's 6% XLSX-from-memory regression.** The 0280 base `80e24f8745` was
rebuilt with its own harness and compared with this record's base.

- Callgrind counts 123.8 M instructions for 0280's base against 125.4 M for
  this record's base (+1.3%, 50 queries).
- Simulated I1 misses are 3.4 K vs 6.8 K and branch mispredictions 1.6 K vs
  1.7 K. Spread over 150 K matches, both are negligible.
- Three interleaved wall-clock runs at load 15 spread ±20%, and did not
  reproduce the 6%.

The work difference is therefore about 1%, and the rest is layout or noise.
After this record the route counts 2.30 M instructions per query, compared
with 2.48 M at `80e24f8745`.

### 3.3 ODS (`ods_unified_bench`, 1,000 × 10)

| scenario | base µs | after µs | leaf µs | ratio base | ratio after |
|---|---:|---:|---:|---:|---:|
| query_warm | 530.0 | 330.8 | 14.6 | 36.3/36.7 | 23.1/21.2 |
| query_cold | 576.1 | 433.7 | 14.2 | 22.0/40.4 | 16.6/13.8 |
| batch_10000 | 52841.8 | 51199.9 | 38517.1 | 1.39/1.37 | 1.30/1.29 |
| set, batch_1…1000 | | | | 1.00–1.02 | 1.00–1.02 |

- Every write output is byte-identical to the leaf's.
- Callgrind of the warm query: 8.0 M → 6.1 M instructions.
- The remaining gap is the generic walker step. The leaf leg is a loop over
  borrowed runs costing about 1.4 ns per cell.

### 3.4 ODT (`odt_paired_bench`)

| scenario | paragraphs | base unified µs | after unified µs | ratio base | ratio after |
|---|---:|---:|---:|---:|---:|
| query_all | 1,000 | 123.2 / 129.8 | 100.0 / 100.2 | 4.03/4.32 | 3.41/3.57 |
| query_all | 10,000 | 1275.5 / 1322.7 | 992.7 / 1046.9 | 3.80/4.10 | 3.58/3.31 |
| query_all_owned | 10,000 | 1269.5 / 1322.5 | 993.3 / 1048.6 | 3.70/3.95 | 3.34/3.26 |

- Writes and open are unchanged.
- About 45% of the remaining unified cost is the leaf's `Node::text`, which
  the leaf leg also pays.

### 3.5 RTF (`rtf_destination_bench`, ratios to native)

| family | targets | batch before | batch after | fan-out before | fan-out after | fan-out with `Plan` (new route F) | direct before | direct after |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| table_cell_text | 1 | 1.10 / 1.09 | 1.08 / 1.06 | 1.31 / 1.28 | 1.22 / 1.24 | 1.13 / 1.13 | 1.03 / 1.02 | 1.02 / 1.03 |
| footnote_text | 1 | 1.09 / 1.10 | 1.11 / 1.09 | 1.28 / 1.28 | 1.22 / 1.25 | 1.10 / 1.10 | 1.04 / 1.02 | 1.03 / 1.03 |
| comment_text | 1 | 1.09 / 1.10 | 1.08 / 1.07 | 1.27 / 1.28 | 1.22 / 1.22 | 1.12 / 1.12 | 1.03 / 1.02 | 1.03 / 1.02 |
| picture_removal | 1 | 1.06 / 1.06 | 1.07 / 1.08 | 1.48 / 1.46 | 1.37 / 1.36 | 1.18 / 1.18 | 1.06 / 1.07 | 1.05 / 1.05 |
| picture_removal | 64 | 1.55 / 1.55 | 1.54 / 1.55 | 1.42 / 1.43 | 1.35 / 1.37 | 1.33 / 1.34 | — | — |

- The edit-closure set is unchanged within noise at 1.06–1.11×. In
  instructions it fell by about 1.2 K (the index is no longer built).
- Query fan-out improved by 5–8 points.
- A selector compiled once into a `Plan` (`edit.set(&plan, …)`) comes within
  1.10–1.18× of native.
- Larger batches are within noise of 0280 (`results/0287/rtf-*.json`).
- Every sample is byte-identical to the native oracle.

### 3.6 XLS positional get (`xls_read_bench`, µs per cell)

| cells | native before | native after | unified before | unified after | ratio before | ratio after |
|---:|---:|---:|---:|---:|---:|---:|
| 1 | 1.18 | 0.86 | 1.66 | 1.23 | 1.41/1.41 | 1.43/1.44 |
| 1,000 | 1.10 | 0.81 | 1.71 | 1.32 | 1.55/1.57 | 1.62/1.62 |
| 10,000 | 1.54 | 1.24 | 2.18 | 1.75 | 1.41/1.42 | 1.43/1.41 |

- Both legs save the same three `fstat` calls, about 0.3–0.4 µs.
- The ratio rises slightly because the leaf leg got faster too.
- Scans are unchanged.

## 4. Differential and parity evidence

- `tests/query_corpus_differential.rs` passes with `--all-features`, and its
  golden was not touched. `tests/query_early_stop_parity.rs`,
  `node_view_allocations` and `path_allocations` also pass.
- New `format::ods::tests::cell_only_walks_match_the_phased_sheet_walk`: on
  six ODS fixtures, `query("cell")` gives the same kinds, paths, previews and
  texts as the cells of `query("cell, row, col, mergeCell")`.
- New `format::ods::tests::row_ends_find_the_runs_sheet_cell_finds`: the
  binary search returns the very run (by pointer) that `Sheet::cell`
  returns, at every row and cell run boundary of every fixture sheet, padding
  runs included.
- New `litchi-cfb`
  `deferred_reads_leave_the_success_fence_to_the_next_fenced_read`: a
  deferred read makes no observation, and the next fenced read or
  `check_source` reports a change. A failed deferred read still observes
  the source, and `SourceChanged` takes precedence over the read's own
  error.
- New `litchi-xls` `an_indexed_hit_fences_every_observation_of_its_frame`:
  an indexed hit fails with `SourceChanged` when the source changes before
  the call, or before either observation the hit makes.
  `an_indexed_missing_target_still_takes_the_trailing_freshness_fence` now
  injects at observation 1, the hit's only one, because it no longer has a
  leading fence.
- The ODS batch outputs, RTF outputs, and the officecli corpus
  (`litchi-cli --test officecli_corpus`, not re-blessed) are unchanged.

## 5. Gaps and follow-ups

- **Unified checks in the XLS adapter.** `format/xls` (not owned here)
  observes the source four times per `get` plus `view().text()`: in the
  resolver, in two nested `finish` calls under `View::text`, and once in
  formatting. Only one check per public read is required, so the nested
  `finish` calls can share the outer one.
- **Generic walker step.** It costs about 190 instructions per match on XLSX
  and about 350 on ODT. Most of that is spread over the counter search,
  alternative test, descend test and visit charge, with no single hotspot.
  A fast path for single, kind-only, positionless plans is the next step.
  This is what keeps ODS at 21–23× its borrowed leaf loop and XLSX at
  4.3–4.6× its owned leaf.
- **ODT.** The query is now 3.3–3.6× the fair leaf. The remaining adapter
  cost is the child step (about 130 instructions) and address copies.
- **RTF plan compile.** It costs about 0.5 µs per call when a selector is
  passed as text or as a `Query`. Callers that repeat a selector can compile
  it once into a `Plan`. A per-handle plan cache keyed by selector text was
  not added: it would be hidden mutable state, and the measured route passes
  a `Query` anyway.
- **Other mutation resolvers** (XLSX, PPTX, ODP, `odf_tree`, `xlsx/find`)
  still copy their `QueryArg` with `query.clone()`. They can pass `query`
  instead, now that `&QueryArg` implements `IntoQuery`. Those adapters are
  owned elsewhere.
- **Node sizes.** DOCX (176) and PPTX (168) now set the unified `Node`
  size, and they belong to the OOXML agent.
- **Unexplained wall-clock readings:**
  - XLS `file_cold_first_cell_value_gt`: +2.8% in both runs.
  - `all-formats` `xlsx_1000_bytes_all_cell_value_gt`: +9% in wall-clock
    time against −1.3% in instructions.
  - Both appeared at load 17–25. Re-measure at a quieter time.
- **Unchanged here:** `header_paragraph_text` batches (1.24×) and ODT
  `batch_bold_1000` (1.35×).

## 6. Gates

Each exit status was checked, with builds pinned to CPUs 0–29 and the build
environment set per the rules file.

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-cfb -p litchi-xls -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, docx, xlsx, pptx, xlsb, xls, doc, ppt, rtf, odt, ods, odp, odg | 0 each |
| `cargo test -p litchi-cfb`, `-p litchi-xls` | 0, 0 |
| `cargo test -p litchi-ops --all-features` (includes the corpus differential) | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --all-features --test capabilities --test registry` (not re-blessed) | 0 |
| the same with `--no-default-features --features <every feature but encryption>` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-cfb -p litchi-xls -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir …` (33 profiles, 1,410 bans, 46 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo check --locked --offline --manifest-path tools/perf-baseline/unified-query/Cargo.toml --all-targets --features all-formats` | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` | 0 |

The `unified-query` harness `Cargo.lock` gains the entries for the leaf
crates that `all-formats` pulls in. No golden or delta table changed, so
`officecli_capability_delta.py` was not re-run.
