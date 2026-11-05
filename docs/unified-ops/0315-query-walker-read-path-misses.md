# 0315 — Query walker and read-path misses (record 0310 groups G4–G9)

- Status: implemented on `uo2/query-walker`, cut from `feat/unified-ops-api`
  at `30eed270b5`; trunk `db96990ff6` (records 0314, 0316–0321) merged at
  `bfdd60d3c1` before the final measurement and gates; trunk `fde847fe90`
  (record 0322, XLSB features) merged at `78d782b49c` after them, followed by
  `cargo fmt --all --check`, the all-features clippy of `litchi-ops`, its
  `--lib` tests (926), the query-corpus differential and the matrix
  `--check` (exit 0 each).
- Date: 2026-10-17
- Input: record [0310](0310-miss-sweep.md) §2 groups G4–G9 and §7
  follow-ups; the 98 open misses of `tools/unified_bench_matrix.py`.
- Scope: `litchi-ops` `query/{walk,plan,compile,matches,tests}`,
  `format/docx` (`Children::next`, `DocxOwner`), `format/xls`
  (`XlsState`, the sheet-step cache), one hunk in `read.rs` (the XLS `get`
  arm); new routes in the `unified-query` harness and in the
  `ods_unified_bench`, `rtf_batch_set_bench` and `rtf_layout_bench`
  examples; `tools/unified_bench_matrix.py` and its curated file.
- `performance_claim`: paired measurements in §4. Release (`unified-query`,
  both builds) and `profiling` (examples) builds, one pinned CPU
  (`taskset -c 31`) of the shared 32-core host at load 4–11, processes in
  B A A B order, each harness an in-process ABBA. Instruction counts are
  callgrind (`perf` is blocked, `perf_event_paranoid=4`). Evidence:
  [results/0315/](results/0315/): driver `measure.sh`, `binaries.sha256`,
  `callgrind-summary.txt`, `pptx-trunk-attribution.txt` and the run set
  `abba/` (base = `30eed270b5` library with this record's new routes;
  after = the merged tree).

## 1. Summary

| | open misses |
|---|---:|
| matrix before this record (`30eed270b5`) | **98** |
| closed: DOCX `drawing` file `r[bold=true]` (G7, 1.228 → 1.07–1.08) | −1 |
| closed, single-target RTF sets now at the bar (no RTF code changed; read 1.100–1.127 in 0310, 1.065–1.100 here): `run_text` ×2, `run_compound_format` ×2, `comment_text` fan-out with `&Query` | −5 |
| new rows of this record's new routes: XLSB fair legs (bytes owned, bytes fenced, file owned), ODS `query_warm_owned`, two RTF typed-props routes | +6 |
| **matrix after this record** (merged tree) | **98** |

Every group was improved where this record owns the code, except G9, whose
typed-props route showed that typed lengths are coerced through their
text (§3.4). The ratios of G4, G5, G6 and G8 fell but stay above their
1.10 trigger; the remaining cost is in adapters this record does not own
(node construction, per-match checks by owner decision) or in the
comparison itself (G4, G6 borrowed leaf loops).

| group | before (0310 / base) | after | what changed |
|---|---|---|---|
| G4 full queries, borrowed leaf | XLSX bytes `cell` 15.6–17.6, `cell[value>1000]` 12.6–13.1 | 13.1–15.9, 12.0–12.4 | walker fast step (§3.1) |
| G5 full queries, fair leaves | XLSX bytes owned 4.31–5.19 / fenced 2.65–3.14; 50,000 rows owned 4.27–4.29 | 3.89–4.39 / 2.39–2.71; 3.85–3.89 | walker fast step |
| G4/G5 new XLSB fair legs | — | bytes owned 4.78–4.83, fenced 2.40–2.42; file fenced 0.64–0.65 (meets) | new routes (§3.5) |
| G6 ODS `query_warm`; ODT `query_all` 10,000 ¶ | 21.9–22.1; 2.99–3.10 | 17.5–18.1; 2.43–2.67 | walker fast step |
| G6 ODS fair leg | — | `query_warm_owned` 2.42–2.51 | new route |
| G7 DOCX `saut` file / `drawing` file | 1.49–1.51 / 1.23–1.24 | 1.15–1.19 / 1.07–1.08 (meets) | DOCX step check (§3.2) |
| G8 XLS eager `get` (1–10,000 cells); positional | 2.58–3.54; 1.40–1.67 | 2.06–2.84; 1.33–1.57 | sheet-step cache (§3.3) |
| G9 RTF single sets | 1.08–1.31 | 1.07–1.30; typed routes 1.17–1.30 | typed-props routes, measured (§3.4) |

## 2. Groups and causes

- **G4/G5 (query walker).** Callgrind of `xlsx_1000_bytes_all_cell`
  (5 queries, 15,015 matches): 671 instructions per match. The walker's
  own step was ≈ 140 (`next_found` self), `Matches::next` and the slot
  ≈ 88, moves of the 184-byte node ≈ 30, the drop the harness pays ≈ 30;
  the adapter's child step ≈ 150 and the per-match source check ≈ 63.
  The kind-only shortcut of record 0291 had already removed most of the
  per-alternative tests, so the walker's cost was mostly moves through the
  match slot and function frames.
- **G6.** ODS and ODT run the same walker; the ODS warm query cost
  ≈ 610 instructions per cell.
- **G7.** `format/docx` `Children::next` checked the source before and
  after every child, and `DocxOwner::finish` checked on every node read,
  inside fenced steps too: `Nav` had no fence key, so the walker's scope
  never applied to DOCX.
- **G8.** Callgrind of 1,001 eager `get` + `view().text()`: 7.14 K
  instructions per call, of which the sheet and cell steps' vocabulary
  lookups (`Vocabulary::rule`, `alias`, `Kind::from_name`) were ≈ 2.0 K
  and path admission 0.35 K.
- **G9.** Record 0310 §7: ≈ 0.9 K instructions per length string.

## 3. Changes

### 3.1 Walker fast step (G4–G6)

- `Plan` carries a `FastPlan` (computed once by `compile`) when the plan
  has one alternative, one compound, no position, no header rows, no lone
  `*`, and a kind-only, single-filter or content shortcut (`cell`,
  `cell[value>1000]`, `shape:contains(Revenue)`, `p:has(r)`). Its interest
  cannot change, so the matching kinds, the kinds worth entering and the
  kinds to ask the adapter for are fixed masks.
- `Matches` decides once, at construction, whether its walk takes the
  fast step, and then calls `Walk::next_fast_map`: the same fence bracket,
  limit, fusing and error handling as `advance`, the same reads in the same
  order as the general step (kind, shortcut test, then the child test of a
  node it may enter), and the match is handed to the caller's wrapper
  without passing through the walk's slot. The `Hit` route (paths) and
  every other plan keep the general step.
- The general route's code is deliberately untouched. Earlier versions
  that dispatched inside `next_map`, kept the fast state in the walk, or
  gave `next_map` a second call site left the general step's instruction
  count unchanged (callgrind 16.45 M against 16.50 M) but read 9–15 %
  slower in wall-clock time on `xlsx_1000_bytes_all_cell_value_gt`, which
  then took the general step: a code-layout effect. They were not kept.
- No public type changed: `Matches` still yields owned `Node`s. Deferring
  the node until the caller consumes it would need a borrowed (lending)
  item type, a new public shape the glossary does not name; it is not
  proposed here (§6). The harness routes consume every node, so it would
  not change their ratios.

### 3.2 DOCX step check (G7)

`DocxOwner` gains `fence_key` (its address in the handle's shared state)
and `ensure_current` (the unconditional check). `Nav::fence_key` returns
it, so the resolver's and walker's scopes cover DOCX reads; inside such a
scope `DocxOwner::finish`/`check_source` and `Children::next` pass the read
through, and the bracket's closing check, which `Nav::check_source` runs
unconditionally, fences it. The walker still checks once per match (owner
decision 2026-10-14, record 0280 G1). Outside a step every read checks as
before.

### 3.3 XLS sheet-step cache (G8)

`Xls` now holds `XlsState { package, sheet_step }`. A `get` of an
absolute two-step cell path (`/<sheet step>/<cell>`, typed or text, not
raw) whose sheet step equals the cached one (name and predicates) runs
path admission, then the root, the worksheet by its cached catalog
position and the cell — the resolver's reads, inside the resolver's
fence — and the resolver's closing check. Anything else (no cache, a busy
cache, a miss, an error, a changed source) takes the general resolver,
which then gives its own result and error; a successful resolver `get` of
such a path fills the cache. Every write publishes a new state
(`Xls::with_package`), so a write invalidates the cache by construction.

### 3.4 RTF typed-props routes (G9)

New routes pass lengths as typed `Length` values:
`paragraph_layout_alignment_typed` and `run_compound_format_typed` in
`rtf_batch_set_bench` (`RTF_BATCH_SET_COUNT=1` here) and
`last_path_typed` in `rtf_layout_bench`; outputs are byte-identical to the
native oracle. They read **slower** than the text routes — layout path
1.17–1.18 against 1.12, paragraph layout 1.29–1.30 against 1.23 — because
`PropValue::coerce` renders a typed `Length` to text and parses it again
through the same `Length::parse_in`. A typed coercion path must keep each
length context's validation (font sizes, extents, line widths), which the
rendered text currently enforces; `props.rs`/`value` are not owned here,
so this is a follow-up (§6). No RTF code changed.

### 3.5 Fair leaf legs (G4–G6)

- `unified-query`: `xlsb_*_all_cell_owned_v1` and `_fenced_v1`, as record
  0280's XLSX/PPTX legs: each stored cell kept as (worksheet position,
  shared grid, coordinate), and in the fenced leg one leaf
  `source_version` per match.
- `ods_unified_bench`: `query_warm_owned`, the warm query against a leaf
  leg that expands repeated rows and cells as the walk does and keeps each
  content cell as (shared workbook, coordinates) while reading its text;
  its text total is checked against the unified leg outside the clock.

## 4. Measurements

### 4.1 Instructions (callgrind, `callgrind-summary.txt`)

| route | base | after | Δ |
|---|---:|---:|---:|
| `xlsx_1000_bytes_all_cell` ×5 | 10.08 M | 9.22 M | −8.5 % |
| `xlsx_1000_bytes_all_cell_value_gt` ×5 | 16.50 M | 14.26 M | −13.6 % |
| `xlsb_1000_bytes_all_cell` ×5 | 12.29 M | 11.69 M | −4.9 % |
| `xlsb_1000_bytes_all_cell_value_gt` ×5 | 27.09 M | 25.45 M | −6.1 % |
| `pptx_small_bytes_all_shape_width` / `_contains` ×5 | 0.839 M / 0.680 M | 0.814 M / 0.660 M | −2.9 % / −2.9 % |
| `pptx_small_bytes_all_shape` ×5 | 0.606 M | 0.636 M | +5.0 % (trunk, §4.2) |
| XLS 1,001 eager `get` + text | 7.14 M | 5.24 M | −26.7 % |
| XLS 1,001 positional `get` + text | 15.81 M | 13.93 M | −11.9 % |
| XLS scan, eager / positional | 3.69 M / 7.07 M | 3.55 M / 6.91 M | −3.7 % / −2.2 % |
| ODS warm `query("cell")` (6 calls) | 19.74 M | 17.90 M | −9.3 % |
| DOCX `drawing` / `saut` bytes `r[bold=true]` | 401.2 M / 1.852 M | 400.0 M / 1.838 M | −0.3 % / −0.7 % |

The DOCX file routes gain in time, not instructions: the removed checks
were `fstat` calls.

### 4.2 Queries (`abba/query-*.jsonl`, unified µs, lower of two medians)

| route | base | after | ratio base → after |
|---|---:|---:|---|
| `xlsx_1000_bytes_all_cell` | 145.6 | 131.1 | 17.6/15.6 → 13.1/15.9 |
| `xlsx_1000_bytes_all_cell_owned` / `_fenced` | 145.7 | 131.1 | 5.19/4.31 → 3.89/4.39; 3.14/2.65 → 2.39/2.71 |
| `xlsx_1000_bytes_all_cell_value_gt` | 217.0 | 201.9 | 13.1/12.6 → 12.4/12.0 |
| `xlsx_50000_bytes_all_cell` | 7,248 | 6,538 | 9.8/10.0 → 4.7/4.9 |
| `xlsx_50000_bytes_all_cell_owned` / `_fenced` | 7,212 | 6,507 | 4.27/4.29 → 3.85/3.89; 2.63/2.61 → 2.37/2.40 |
| `xlsx_1000_file_all_cell_fenced` | 522.1 | 515.4 | 1.20/1.20 → 1.18/1.19 |
| `xlsb_1000_bytes_all_cell` | 151.8 | 149.3 | 15.4/13.8 → 12.9/12.3 |
| `xlsb_1000_bytes_all_cell_value_gt` | 339.2 | 352.4 | 31.9/31.6 → 33.3/33.6 |
| `docx_saut_file_all_run_bold` | 26.58 | 21.15 | 1.51/1.49 → 1.19/1.15 |
| `docx_drawing_file_all_run_bold` | 4,849 | 4,325 | 1.24/1.23 → 1.08/1.07 |
| `pptx_large_bytes_all_shape` | 1,636 | 1,922 | 6.29/5.80 → 6.57/6.96 |
| `pptx_large_bytes_all_shape_contains` | 1,481 | 1,494 | 5.25/4.96 → 4.91/5.06 |
| all-formats `xlsx_1000_bytes_all_cell` | 152.0 | 138.8 | 15.3/15.3 → 14.9/17.9 |

- **PPTX kind-only routes** read 9–23 % slower after; the after binaries
  include trunk `db96990ff6`, whose records 0319–0321 add PPTX element
  kinds. A harness built from trunk alone (`pptx-trunk-attribution.txt`,
  three binaries interleaved) reads `pptx_large_bytes_all_shape`
  1,986–2,420 µs against base 1,719–1,806 and after 1,959–2,261: the
  slowdown is trunk's, and after is within noise of trunk (owned leg:
  trunk 1,775–2,193, after 1,800–2,219).
- **XLSB `cell[value>1000]`** reads 3.9–5.1 % slower although callgrind
  counts 6.1 % fewer instructions; record 0315 found the same kind of
  layout sensitivity on XLSX filter queries during the work (§3.1). Not
  explained beyond that; listed in §6.
- Early stops (`first`, `limit(10)`) and `get`s are within ±3 % except
  PPTX `first_shape_contains` (+3–8 %, trunk's PPTX changes or noise).

### 4.3 Examples (`abba/`, ratios per run, base → after)

| harness, case | base | after |
|---|---|---|
| ODS `query_warm` (318 → 252 µs) | 21.9/22.1 | 17.5/18.1 |
| ODS `query_warm_owned` (new) | 3.08/3.06 | 2.42/2.51 |
| ODS `query` / `query_cold` | 21.5/24.6; 25.9/26.1 | 13.9/20.6; 21.3/21.5 |
| ODT `query_all` 100 / 1,000 / 10,000 ¶ | 3.14–3.27 / 3.03–3.30 / 2.99–3.10 | 2.83–2.88 / 2.56–2.73 / 2.43–2.67 |
| XLS eager `get` 1 / 100 / 1,000 / 10,000 cells | 2.66/2.58, 3.54/3.46, 3.22/2.95, 2.72/2.66 | 2.06/2.07, 2.82/2.84, 2.56/2.58, 2.22/2.26 |
| XLS positional `get` 1 / 100 / 1,000 / 10,000 | 1.40–1.42, 1.64–1.67, 1.61–1.64, 1.42–1.43 | 1.33–1.35, 1.54–1.57, 1.52–1.53, 1.35–1.37 |
| XLS positional scan 1,000 / 10,000 | 4.65–4.66 / 4.74–4.93 | 4.48–4.51 / 4.61–4.65 |
| RTF layout path ×1, text / typed | 1.107–1.126 / 1.159–1.175 | 1.119–1.120 / 1.170–1.181 |
| RTF `paragraph_layout_alignment` ×1, text / typed | 1.231–1.236 / 1.309–1.316 | 1.228–1.232 / 1.291–1.298 |
| RTF `run_compound_format` ×1 (paragraphs), text / typed | 1.078–1.085 / 1.087–1.100 | 1.086–1.089 / 1.093–1.094 |
| RTF fan-out `table_cell_text` ×1 | 1.121–1.127 | 1.134–1.139 |

The ODT leaf leg reads slower in the after runs (`408 µs` against `342`
at 10,000 ¶); the unified leg moved 1,023–1,072 → 1,077–1,090 µs, so
part of the ODT ratio change is the leaf leg's (trunk changes or noise).
The RTF fan-out rows move +1–3 % with no RTF change; within noise.

## 5. Correctness evidence

- `query::tests::fast_path_matches_the_general_step`: on the fake Word,
  presentation and workbook documents, 20 selectors give the same nodes
  (by identity), visit counts and errors on the fast step as on the
  general step, also under `limit(1)` and `limit(2)`, and the `Hit`
  route's nodes are the iterator's; 10 more selectors are checked to take
  the general step.
- `format::docx::tests::query_children_share_the_step_check`: on
  generated bodies of 10 and 20 paragraphs, ten more matching paragraphs
  or runs cost exactly ten more source checks and ten more non-matching
  runs cost none (before: two checks per child); a source changed between
  two pulls fails the next pull with `IoError`, and outside a step a read
  still checks.
- `format::xls::tests::cached_sheet_steps_resolve_as_the_resolver_does`:
  on `Simple.xls` and `HyperlinksOnManySheets.xls`, both owners, every
  sheet-step form (`sheet`, `sheet[last()]`, `sheet[N]`, quoted names,
  upper-cased names, a missing ordinal and name) with seven cells each
  (including `IW1`, `A65537` and a lower-case cell) gives the resolver's
  node, text or error byte for byte, uncached, cached and after the cache
  moved to another step.
- `format::xls::tests::cached_sheet_steps_keep_the_resolver_checks`: a
  cached `get` makes exactly the resolver's source observations; a source
  changed between two cached `get`s fails with `SourceChanged`; a renamed
  sheet's old name no longer resolves on the written handle while the old
  handle keeps its snapshot.
- The query-corpus differential, corpus fidelity, the checklist goldens,
  `query_early_stop(_parity)`, `path_allocations`, `node_view_allocations`,
  `xlsx_query_allocations` and the capability and registry goldens pass
  unchanged; nothing was re-blessed (§7).

## 6. Gaps and follow-ups

- **Path-text parser (coordinator scope addition, record 0314).** Not
  started: a fast parser for absolute `/kind[N]/…` paths with a
  differential test against the current parser over every test and golden
  path, measured with record 0314's `batch_*` and `*_parsed` legs. Left
  for a follow-up record.
- **Typed length coercion (G9).** `PropValue::coerce` should accept a
  typed `Length` for a length key without rendering it, applying each
  `LengthContext`'s checks to the EMU value (`props.rs`, `value/length`);
  the typed routes of §3.4 then measure it. Making text coercion cheaper
  (a unit-suffix fast path in `value/length`) is the other half.
- **Adapter node cost (G4/G5, owned elsewhere).** XLSX and PPTX build a
  168–184-byte node per match and move it through the adapter's iterator;
  smaller nodes or an in-place child API in those adapters is what remains
  on the fair-leaf ratios. A lending (borrowed) match type for `Matches`
  would let callers skip node construction where they only test or count;
  it needs a glossary-conforming public name first.
- **Layout-sensitive wall-clock readings.** XLSB `cell[value>1000]` +4–5 %
  against −6 % instructions; record 0287 saw the same on XLSX. Re-measure
  on a quieter host; a pinned code-alignment build would separate layout
  from work.
- **PPTX after trunk.** `pptx_*_all_shape` routes read 15–35 % slower on
  trunk `db96990ff6` than at `30eed270b5` (§4.2), unrelated to this
  record; for the PPTX owner (records 0319–0321).
- **XLS.** The remaining `get` cost is the cell parse, path admission,
  `View::text`'s number formatting and, on the positional owner, the
  per-call checks (owner decision).

## 7. Gates

Merged tree (`bfdd60d3c1` + this record), build environment per the rules
(`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`, the memory-capped cargo
shim, one cargo command at a time). Every exit status was checked.

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, docx, xls, xlsb, ods, rtf, odt, xlsx, pptx | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops` | 0 |
| `cargo test -p litchi-ops --all-features` in chunks: `--lib` (926), `--doc` (208 + 2), the 145 test targets in five groups (1,291) — includes the query-corpus differential, corpus fidelity, checklist goldens, `path_allocations` | 0 each |
| `cargo test -p litchi-ops --all-features --test capabilities --test registry` (not re-blessed) | 0 |
| the same with `--no-default-features --features` every litchi-ops feature but `encryption` (not re-blessed) | 0 |
| `cargo test -p litchi --all-features` (230) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,313 bans, 238 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (281 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo check --locked --offline --manifest-path tools/perf-baseline/unified-query/Cargo.toml --all-targets [--features all-formats]` | 0 |
| `python3 tools/unified_bench_matrix.py --check` | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so the
delta table and the OfficeCLI corpus were not regenerated. No public name
was added or changed; no glossary row changes. No dependency was added.

## 8. Not claimed

- That any G4–G9 miss other than the six in §1 meets its budget.
- That typed props make RTF single sets cheaper today (§3.4: they do not).
- Any wall-clock gain on XLSB filter queries (§4.2).
