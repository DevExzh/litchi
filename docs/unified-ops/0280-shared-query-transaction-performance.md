# 0280 — Shared query and transaction performance

- Status: implemented on `uo2/perf-query`, cut from `feat/unified-ops-api`
  at `80e24f8745`. Follows up the open PERFORMANCE BAR misses of records
  [0262](0262-query-walker-performance.md), [0267](0267-streaming-cell-cursors.md),
  [0275](0275-rtf-parity.md), [0276](0276-odt-unified-adapter.md) and
  [0277](0277-ods-unified-adapter.md).
- Date: 2026-10-14
- Scope: `litchi-core` `FileSource::version`; `litchi-ops`
  `query/{walk,inspect,compile,plan}`, `path/arg`, `transaction`; small
  hunks in `format/pptx/{owner,metadata}.rs` and `format/pptx.rs` (cache
  fields), `format/odt/nav.rs` (one `Inspect` method), `format/rtf/write.rs`
  (module line) plus the new `format/rtf/write/removal.rs`; new harness
  routes; tests.
- `performance_claim`: paired measurements in §4 (release, `taskset -c 31`
  on the shared 32-core host, which ran other agents' builds at load 13–17;
  processes in B A A B order, each an in-process ABBA). Evidence:
  [results/0280/](results/0280/). `perf` is unavailable on the host after
  its reboot (`perf_event_paranoid=4`), so the profiles after the first
  session are callgrind.

## 1. Breakdown before optimizing

Profiles of the trunk build (`80e24f8745`):

- **XLSX full `cell` query from memory** (callgrind, 1,000 rows, about 770
  instructions per match; `perf` with IBS cycles on 50,000 rows): the
  walker's own step about 190 instructions, the adapter's child step about
  150, moving each match through `Option<Result<Node>>` layers about 130,
  node drop about 45, the per-match source check about 70. In cycles the
  `Arc` grid increment per listed cell (23%) and its decrement at drop
  (20%) dominate; on this AMD EPYC 9R45 VM an `Arc` clone and drop costs
  about 9 ns. Skipping the per-match check on memory sources changes
  nothing measurable (157–197 → 150–157 µs, inside the noise), as record
  0267 found.
- **The same query on a file source**: checking only every 1,024th pull
  takes the 1,000-row route from 655 to 153 µs, so the check is 75% of it,
  about 170 ns per match. A microbenchmark splits that into the `statx`
  call through `File::metadata` (156 ns; a raw `fstat` is 125 ns) and the
  mutex (9 ns). The system call itself is the floor.
- **PPTX full `shape` query** (500 slides × 40 boxes): about 30% of the
  steady-state instructions re-read masters and layouts on every query
  (`layouts_in_catalog`, `layout_scene`, `master_scene`: each re-parses the
  presentation part and loads the related part through the source-backed
  leaf); `perf` also showed 30% `memmove` and 15% `Handles::next` (one `Arc`
  clone per shape in `strict_scene`'s check).
- **ODT full `p` query**: the walker entered every matched paragraph (a
  paragraph can hold frames with paragraphs), so each match cost a node
  clone, a frame push with two `memcpy`s, a child iterator
  (`Children::of`) and a pop — about a quarter of the query.
- **ODS full `cell` query**: about 670 instructions per cell; the walker's
  own share is about 180, the rest is the ODS adapter (eager A1 label per
  cell, `Label::push_number`, and its child iterator).
- **RTF single-target writes**: a one-target query fan-out costs about
  15 K instructions more than the direct `set` (compiling the plan about
  7 K, of which 0.8 K is the per-compile transitive closure of the spec's
  child lists; the walk about 6 K; the harness's `Query` clone about 2.5 K).
  The edit closure around one `set` adds about 3 K. 64 direct picture
  `remove`s publish 64 native commits.

## 2. What changed

1. **Lock-free, cheaper file-source version check** (`litchi-core`
   `FileSource::version`). The last published fingerprint and revision sit
   in a sequence-locked set of atomics; a check reads the file's metadata
   with one `fstat` (new unix-only `libc` dependency, one `unsafe` call),
   compares it without the mutex, and returns the published revision. Any
   difference, or a concurrent update, takes the old locked path, which
   reads the metadata again under the mutex, advances the revision and
   publishes it. The guarantee is unchanged: every call reads the metadata,
   so the next pull after a source change fails. Tests: a 4-reader
   concurrent test (every change observed, revisions monotonic per thread)
   and the `fstat` fingerprint equal to the `Metadata` one.
2. **`Inspect::childless`** (shared walker): an adapter can tell the walker
   that a node has no children of the kinds the walk still wants; the
   walker then yields it without a clone, frame or child read. Default
   `false`. ODT implements it for body elements (a paragraph of runs when
   the walk looks for paragraphs).
3. **Per-handle master/layout cache** (PPTX, small hunks): layout counts and
   the strict scenes of masters and layouts are kept per handle, as the
   master count already was (record 0262). A handle's content never changes
   (a changed source fails its fence; a write returns a new handle with an
   empty cache); failed reads are not kept.
4. **Coalesced direct RTF picture removals** (`transaction.rs` +
   `format/rtf/write/removal.rs`): consecutive `Edit::remove` of exact
   `/body/p[N]/picture[M]` paths stage into one run and publish one native
   removal, the commit a query `remove` makes. Each later path resolves as
   it would after the earlier removals (removing a picture moves no
   paragraph and no other anchor). Any other target, a read, a set, or the
   native cap of 64 pictures closes the run; a path that names no remaining
   picture closes it and takes the ordinary path, so its error and item
   index are unchanged.
5. **Plan compilation** keeps one transitive-closure table per spec table
   for the process (`Plan::reach` is now `&'static [u128]`).
6. **Path admission** skips writing the canonical spelling when a
   conservative bound of its length (quoting at most doubles a name, key or
   value) is under the input limit; long paths keep the exact check. Not
   separately timed; the 0275 profile charged about 1.3–1.6 K instructions
   per direct write to that write.

No public API changed; no golden changed.

## 3. Fair leaf baselines

The earlier leaf legs iterate borrowed cells or shapes and do nothing per
match. A unified query yields owned, retainable nodes and checks the source
on every pull. New routes (new names; the frozen `_v1` routes are kept):

- `xlsx_*_all_cell_owned_v1`: the leaf keeps each stored cell as an owned
  handle (sheet position, a shared `Grid` clone, the address), which is
  what a unified node holds; `*_fenced_v1` also calls the leaf's
  `source_version()` per match.
- `pptx_*_all_shape_owned_v1` / `*_fenced_v1`: the query from the root
  visits every slide, master and layout, so the leaf leg reads them all
  through the leaf catalog and keeps each match's owned `Handle`.
- ODT `query_all_owned`: the leaf keeps each paragraph and its positional
  address (parent, ordinal).
- ODS `query_warm` (both legs warm) and `query_cold` (both build their
  index on a freshly opened handle inside the clock); record 0277's `query`
  compared a cold unified handle with a warm leaf projection.

## 4. Measurements

Binaries (not committed), SHA-256: `unified-query` base `8346bf9e…1691e`,
after `32d45814…be02b4` (`CARGO_TARGET_DIR=target/bench cargo build
--release --manifest-path tools/perf-baseline/unified-query/Cargo.toml`;
the base build is trunk with this record's harness source);
`odt_paired_bench` base `090f5555…39ca7e`, after `1264d412…6118`;
`ods_unified_bench` base `34ae5e02…d73a6`, after `fa3a8947…2e25c`
(`cargo build --release -p litchi-ops --no-default-features --features
odt|ods --example …`); `unified-rtf-destination-baseline` before
`1d3cef04…dfa41`, after `d1bcddeb…aa47` (root `Cargo.lock` copied next to
the manifest, not committed). Run: `measure.sh` order B A A B, `time
<family> 9`; the 50,000-row warm XLSB routes are skipped as in record 0267.
Columns: lower of the two runs' medians; ratios per run.

### 4.1 Queries (`unified-query`)

| route | matches | base µs | after µs | leaf µs | ratio base | ratio after |
|---|---:|---:|---:|---:|---:|---:|
| xlsx_1000_bytes_all_cell | 3003 | 151.6 | 161.1 | 9.8 | 18.4/16.2 | 16.7/16.5 |
| xlsx_1000_bytes_all_cell_owned | 3003 | 151.7 | 161.0 | 33.8 | 4.53/4.48 | 4.84/4.76 |
| xlsx_1000_bytes_all_cell_fenced | 3003 | 151.4 | 161.1 | 55.0 | 2.79/2.75 | 2.99/2.93 |
| xlsx_50000_bytes_all_cell_owned | 150003 | 7728 | 7759 | 1699 | 4.53/4.45 | 4.42/4.73 |
| xlsx_1000_file_all_cell | 3003 | 653.8 | 547.0 | 9.0 | 73.4/73.7 | 60.4/56.9 |
| xlsx_1000_file_all_cell_fenced | 3003 | 654.6 | 537.4 | 453.8 | 1.11/1.18 | 1.17/1.20 |
| xlsx_50000_file_all_cell | 150003 | 32706 | 26647 | 1127 | 23.8/28.7 | 12.5/24.0 |
| xlsx_50000_file_all_cell_fenced | 150003 | 32500 | 27082 | 22763 | 1.10/1.26 | 1.26/1.19 |
| xlsx_50000_file_all_row_header_gt | 25000 | 42618 | 33012 | 5323 | 6.02/6.83 | 4.12/6.21 |
| xlsx_1000_file_get_cell | 1 | 1.13 | 1.06 | 0.55 | 1.73/1.78 | 1.92/1.93 |
| xlsb_1000_file_all_cell | 3003 | 1184 | 958 | 11.3 | 95.7/95.0 | 77.7/84.6 |
| xlsb_1000_file_get_cell | 1 | 3.13 | 2.58 | 0.73 | 3.38/3.33 | 3.55/3.55 |
| pptx_small_bytes_all_shape | 90 | 57.3 | 9.6 | 1.7 | 34.2/34.4 | 5.70/6.33 |
| pptx_small_bytes_all_shape_owned | 90 | 57.0 | 9.6 | 32.0 | 1.78/1.77 | 0.30/0.33 |
| pptx_small_bytes_all_shape_width | 10 | 60.4 | 13.2 | 1.7 | 36.8/36.4 | 7.99/7.95 |
| pptx_large_bytes_all_shape | 20500 | 2354 | 1744 | 277 | 8.48/8.68 | 6.31/7.21 |
| pptx_large_bytes_all_shape_owned | 20500 | 2361 | 1744 | 793 | 2.94/3.03 | 2.20/2.48 |
| pptx_large_bytes_all_shape_width | 500 | 3102 | 2524 | 275 | 11.4/11.3 | 9.21/9.16 |
| pptx_large_bytes_all_shape_contains | 500 | 2040 | 1452 | 296 | 6.95/6.87 | 4.91/4.91 |
| pptx_large_file_all_shape | 20500 | 5955 | 4560 | 420 | 13.4/13.4 | 10.8/11.3 |
| pptx_large_file_all_shape_fenced | 20500 | 5924 | 4543 | 3492 | 1.39/1.39 | 1.30/1.36 |
| docx_drawing_file_all_run_bold | 177 | 5209 | 4843 | 3954 | 1.26/1.26 | 1.24/1.22 |

The full table, with every get, early-stop and cold route, is
`results/0280/query-*.jsonl`. Reading it:

- **File sources**: every file-backed full query gains 16–23% (the check
  is now about 125 ns, the `fstat` floor); file-backed `get`s gain 6–18%
  (XLSB 3.13 → 2.58 µs). Against the fenced leaf (the same guarantee) XLSX
  and PPTX full queries on files are **1.17–1.36×**.
- **PPTX**: full queries 26–83% faster; the small deck 34× → 5.7× the bare
  leaf loop, and against the fair leaf (which must read the masters and
  layouts it searches) 0.30× on 10 slides and 2.2–2.5× on 500.
- **XLSX memory full scans**: unchanged in work. The 1,000-row route reads
  6% slower in both runs; callgrind counts the same instructions on both
  binaries (146.09 M base, 145.79 M after, 50 queries), so this is code
  layout, not work. Against the fair leaf it stays **4.4–4.8×**: the
  remaining cost is the owned node itself (an `Arc` per match, moves of the
  184-byte unified node) and the walker's step.
- Early termination, cold first matches, `get`s from memory and DOCX are
  unchanged within noise (±3%; XLSX `get` from memory 0.43 → 0.45 µs).

### 4.2 ODT (`odt_paired_bench`, µs)

| scenario | paragraphs | base unified | after unified | leaf | ratio base | ratio after |
|---|---:|---:|---:|---:|---:|---:|
| query_all | 100 | 17.5 | 12.6 | 3.0 | 6.17/5.99 | 4.25/4.20 |
| query_all | 1,000 | 181.8 | 123.6 | 31.6 | 6.86/6.14 | 3.99/3.87 |
| query_all | 10,000 | 1728 | 1219 | 316 | 5.54/6.02 | 3.90/3.86 |
| query_all_owned | 10,000 | 1739 | 1228 | 335 | 5.28/5.92 | 3.77/3.67 |

Writes and open are unchanged (`results/0280/odt-*.txt`). The full query is
30% faster (≈6× → **3.7–3.9×** the fair leaf); callgrind: 26.8 M → 19.5 M
instructions per 10,000-paragraph query.

### 4.3 ODS (`ods_unified_bench`, µs, 1,000 × 10)

| scenario | base unified | after unified | leaf | ratio base | ratio after |
|---|---:|---:|---:|---:|---:|
| query (0277 pairing) | 571.7 | 568.0 | 22.2 | 13.3/11.8 | 12.0/26.3 |
| query_warm | 522.4 | 524.1 | 14.2 | 34.3/34.7 | 36.9/37.4 |
| query_cold | 571.6 | 574.0 | 14.1 | 38.9/42.6 | 37.4/41.1 |

Unchanged, and the fair pairings show a larger gap than 0277's ≈11×:
**35–41×**. About 70% of it is in the ODS adapter (eager A1 labels and its
child iterator), outside this record's files (follow-up G3).

### 4.4 RTF (`unified-rtf-destination`, ratios to native)

| family | targets | batch before | batch after | fan-out before | fan-out after | direct before | direct after |
|---|---:|---:|---:|---:|---:|---:|---:|
| picture_removal | 64 | 55.7 / 52.8 | **1.59 / 1.58** | 1.33 / 1.39 | 1.45 / 1.43 | — | — |
| picture_removal | 1 | 1.09 / 1.09 | 1.07 / 1.09 | 1.53 / 1.50 | 1.49 / 1.50 | 1.07 / 1.05 | 1.07 / 1.06 |
| table_cell_text | 1 | 1.11 / 1.12 | 1.13 / 1.14 | 1.36 / 1.37 | 1.35 / 1.34 | 1.04 / 1.04 | 1.04 / 1.03 |
| footnote_text | 1 | 1.12 / 1.13 | 1.13 / 1.12 | 1.34 / 1.35 | 1.32 / 1.32 | 1.03 / 1.05 | 1.01 / 1.03 |
| comment_text | 1 | 1.12 / 1.09 | 1.14 / 1.13 | 1.33 / 1.33 | 1.33 / 1.32 | 1.04 / 1.03 | 1.03 / 1.02 |

64 direct removals inside one edit: **52.8–55.7× → 1.58–1.59×** (budget 2×),
byte-identical to the native oracle in every sample. The single-target
edit-closure set (1.11–1.14×) and fan-out (1.32–1.50×) are unchanged:
compiling a plan, the walk and the transaction frame cost a fixed
≈0.3–1.4 µs that one ~5 µs native edit does not amortize. Larger batches
are within noise of record 0275 (`results/0280/rtf-*.json`).

## 5. Differential and parity evidence

- `tests/query_corpus_differential.rs` passes unchanged with
  `--all-features` (golden not touched).
- New `query::tests::odt_queries_match_an_unpruned_walk`: on every ODT and
  flat-ODT fixture in `test-data`, a query for each element kind returns
  exactly the nodes an unpruned walk over `Navigate::children` finds, in
  order, and `limit`/`first()` their leading matches. A deliberately wrong
  `childless` fails it.
- New `tests/query_early_stop_parity.rs`: for DOCX, PPTX, ODT, ODP, ODS,
  RTF, DOC and PPT fixtures (up to 25 per format) and 18 selectors,
  `limit(1, 2, 3, 7)` and `first()` return the full query's leading matches
  (kind, path, text) and a warm second query the same matches.
- RTF: `direct_picture_removals_coalesce_into_one_commit`,
  `picture_removal_runs_close_at_barriers_and_attribute_errors` and
  `picture_removal_runs_split_at_the_native_cap` compare against the same
  removals made one call at a time (bytes, commit counts, inverse patch,
  error item index).

## 6. Gaps and follow-ups

- **G1 — per-match file check** (owner decision 2026-10-14: keep the per-match check; a once-per-walk check is rejected). The probe now costs the `fstat` floor
  (≈125 ns). Removing it per pull needs an owner decision: checking once
  per walk (record 0281's ask) breaks "the next pull after a source change
  fails" (`lazy_file_query_and_cached_nodes_refuse_changed_source`). On
  positional XLS owners most probes are leaf-side: per `get` plus a text
  read, callgrind counts eight `FileSource::version` calls, of which the
  shared resolver and view make two; `litchi-cfb`
  `SharedOleStreamCursor::read_exact` makes three and `litchi-xls`
  `query_cell`/`SheetHandle::cell` three. Those leaves could skip their
  probes inside the ops fence scope (not owned here).
- **G2 — owned-node cost on memory scans** (XLSX 4.4–4.8×, PPTX 2.2–2.5× the
  fair leaf): one `Arc` per node and moves of the unified node (184 bytes
  with the harness features, 408 with all features because of the DOC and
  RTF nodes' inline steps). Shrinking DOC/RTF nodes (on-demand paths, as
  0262 did for XLSX) belongs to those adapters.
- **G3 — ODS query** (35–41× fair): eager A1 labels and the child iterator
  in `format/ods`.
- **G4 — PPTX**: `strict_scene` still iterates every shape with an `Arc`
  clone per handle on each slide read; a borrowed iterator in `litchi-pptx`
  would remove it.
- **G5 — RTF single-target fixed costs**: plan compilation (≈5 K
  instructions after the cached closure) and the transaction frame; a
  compiled-plan cache keyed by the selector text, and a lighter one-command
  edit path, are the next steps.
- **G6 — DOCX single set (0281's 1.19×)**: in this branch's callgrind of a
  direct `File::set` on `drawing.docx`, the shared resolver is about 0.2 K
  instructions per set, and the commit wrapping lives in
  `format/docx/write.rs` (`Docx::edit`, `WorkEdit::commit`,
  `nav::load_body` ≈118 K per set), not in `transaction.rs`; it is the DOCX
  adapter's.

## 7. Gates

Each exit status checked (build env per the rules file):

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-core -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, docx, xlsx, pptx, xlsb, xls, doc, ppt, rtf, odt, ods, odp | 0 each |
| `cargo test -p litchi-core` | 0 |
| `cargo test -p litchi-ops --all-features` (includes the corpus differential) | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `cargo test -p litchi-ops --no-default-features --features <every feature but encryption> --test capabilities --test registry` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-core -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir …` (26 profiles, 1,201 bans, 39 families) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

The new unix `libc` edge of `litchi-core` adds one line to each tracked
tool `Cargo.lock` that was `--locked`-consistent at the base (16 files);
the others were already stale there and are left alone. No golden was
re-blessed; the Python surface is unchanged, so its suite was not run.
