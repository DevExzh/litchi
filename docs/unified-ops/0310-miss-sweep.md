# 0310 — Open benchmark misses: root causes, re-measurement, cheap fixes

- Status: implemented on `uo2/miss-sweep`, cut from `feat/unified-ops-api`
  at `0f3d8158cb`; trunk `8f5569d0d8` (records 0305, 0308, 0309) merged at
  `b53ed53006`, then trunk `1100472b75` (records 0311, 0312) at
  `6dd73bb204`, then trunk `09a57ed610` (record 0313) at `fde3041d13`.
- Date: 2026-10-16
- Input: the 102 open misses of `tools/unified_bench_matrix.py`
  (`results/0289/bench-matrix.{tsv,md}` at `0f3d8158cb`), each citing a
  record's explanation or follow-up.
- Scope: `litchi-ops` `format/{xlsx,xlsb,xls}` (adapters), two example
  harnesses (case filters only), `tools/unified_bench_matrix.py` and its
  curated file, the `unified-query` harness lockfile. Out of scope (another
  agent): DOCX single-set and large-document set/add costs, iWork lazy
  trees (record 0305); they are re-measured and cited, not changed.
- `performance_claim`: paired measurements in §4. Release (`unified-query`,
  `unified-rtf-*`) and `profiling` (litchi-ops examples, stripped before
  timing) builds, one pinned CPU (`taskset -c 31`) of the shared 32-core
  host at load 4–7, processes in B A A B order, each harness an in-process
  ABBA. Instruction counts are callgrind with cache simulation where noted
  (`perf` is blocked, `perf_event_paranoid=4`). Evidence:
  [results/0310/](results/0310/): driver `measure.sh`, `binaries.sha256`,
  `callgrind-summary.txt`, and three run sets (§4.1).

## 1. Summary

| | open misses |
|---|---:|
| matrix at `0f3d8158cb` | **102** |
| fixed by this record: XLS positional axis writes (G3, §3.3) | −3 |
| fixed by records merged during this work: DOCX single paragraph `set` (0305; 1.049–1.050 in `merged/`), Keynote open + `get` and Numbers open + `get` (0305, then 0312: 1.034–1.042 and 1.047) | −3 |
| no longer reproduce on trunk: DOCX table-cell `align` and `text` (1.08), RTF dense-run add of 10,000 groups (1.01), record 0304's ODT text-set reading (1.000–1.004 against 0304's own before runs) | −4 |
| rows the new 0305, 0309, 0311 and 0312 parsers add: DOCX 1,000-paragraph set and add (0305 §4.5, fixed by 0311: 0.54 and 0.40–0.47), Keynote and Pages table-cell writes (0309 §6, fixed by 0312: 1.014–1.019), the DOC body-text comparison (0311 §4.2, 1.12–1.27) | +1 |
| 0305's new RTF route G (fan-out with `&Query`), three single-target families | +3 |
| RTF single sets at the bar: `add run (paragraphs, middle)` 1.103 and `run_text (separate_paragraphs)` 1.100, which read 1.05–1.09 and 1.09–1.12 in this record's earlier runs | +2 |
| **matrix after this record** (merged tree) | **98** |

G1 (XLSB) and G2 (XLSX header rows) are fixed in what they cost — 1.05–5.1×
less time, one source check per match as the owner's rule requires — but
stay open misses: what remains is the G4/G5 residual against a borrowed or
typed leaf leg. Every open miss cites a 0310 measurement and its root-cause
group (§2), or the measurement and explanation of the record that owns it
(0311 §4.2 for its DOC comparison). The two before/after-build rows 0310 cannot rebuild
(0263, 0290) carry 0310's note.

## 2. The open misses by root cause

Counts are the matrix's open misses at `0f3d8158cb`; "now" is the merged
tree (§4.3).

| group | misses | root cause (measured) | outcome |
|---|---:|---|---|
| **G1** XLSB full queries | 4 | Nested freshness checks: `XlsbOwner::finish` checked twice per node property or text read even inside a fenced query step, and every check read the source version twice. On a file: 2 `fstat`s per match, 4.3 per `value>` candidate (callgrind) | **fixed** (§3.1); what remains is G4 |
| **G2** XLSX header-row queries (`row[Sales>1000]`, 50,000 rows) | 2 | `Sheet::grid` re-entered the leaf's fenced `grid()` on every cached read: 7 source observations per probed row on a file | **fixed** (§3.2): file 32.8 → 11.2 ms, bytes 9.8 → 7.9 ms; now 1.12–2.15 and 0.90–1.35 against a leaf leg that itself ranges 5.2–5.9 ms |
| **G3** XLS positional `row[N]`/`col[X]` layout writes | 3 | Resolving the target built the worksheet grid: every cell of the sheet read, ≈130 K instructions per call at 100 cells | **fixed** (§3.3): 1.11–1.13 → 1.03–1.04 (100 cells), 1.07–1.11 → 1.00 (10,000) |
| **G4** full queries against a borrowed leaf loop (XLSX, PPTX, and XLSB after G1) | 25 | Comparison: the leaf leg returns nothing retainable and checks nothing; the unified query yields an owned node per match (671 instructions per XLSX match, 1,068 per PPTX shape) and, on a file, checks the source per match (owner decision 2026-10-14) | inherent to the comparison; the fair pairs are G5 |
| **G5** full queries against the fair leaves (`_owned`, `_fenced`) | 16 | The generic walker step and the owned node: XLSX bytes 4.3–5.0× owned and 2.7–3.0× fenced, file fenced 1.19–1.21×; PPTX bytes 2.0–2.8×, file fenced 1.26–1.42× | open; follow-up: the walker fast path for single kind-only plans (0287 §5) |
| **G6** ODS and ODT queries | 9 | Comparison as G4: ODS's leaf loop costs ≈1.5 ns per cell; ODT's leaf `Node::text` is ≈45 % of the unified cost (0287) | inherent to the comparison; follow-up: an owned ODS leaf leg |
| **G7** DOCX queries | 3 | `format/docx` `Children::next` checks the source per yielded child inside fenced steps (491 version reads over 11 queries with no match); the 176-byte DOCX node | follow-up for the OOXML owner |
| **G8** XLS `get` and scans | 16 | Per-call text-path resolution (3.57 K instructions per eager `get` against 1.18 K native: vocabulary 0.75 K, child steps 0.73 K, validation and cell parsing 0.34 K) and, on the positional owner, the per-match and `View::text` checks | inherent at this design; follow-up: cache the parsed sheet step |
| **G9** RTF single-target writes | 14 | Fixed costs one target cannot amortize: 90.6 K instructions per edit-closure set against the native 80.5 K, the excess in property coercion (≈0.9 K per length string), planning, the source index and the transaction | open; one case no longer reproduces; follow-up: a typed-props route |
| **G10** DOCX single-op sets | 3 | 0302 §4.1 (the leaf's own commit, L1 misses) | the two table-cell sets meet on trunk (1.06–1.08); 0305 fixed the paragraph set (1.049–1.050 merged); 0305's 1,000-paragraph set and add (its §4.5) were fixed by 0311 |
| **G11** iWork opens | 4 | Eager trees (record 0305) and Keynote's bytes copy (`Iwork::detect_bytes` copies the owned `Vec`, 500 KB) | 0305 and 0312 fixed both open + get rows; Keynote bytes open still misses (follow-up for the iWork owner) |
| **G12** after/before readings | 3 | Wall-clock readings of before/after builds with no added work: 0304 ODT text set (callgrind +0.22 %), 0290 ODP metadata (+7 ns, the requested string), 0263 XLS get (+30 ns on a 10 ns timer step) | the 0304 reading no longer reproduces (1.000–1.004 against 0304's before runs); the other two cannot be rebuilt and carry 0310's note |

Counts by group after this record (98): G1 4, G2 2, G3 0, G4 25, G5 16,
G6 9, G7 3, G8 16, G9 18 (0305's route G added three), G10 0, G11 2
(Keynote bytes open), G12 2, and 0311's DOC body-text comparison (its
§4.2).

## 3. Fixes

### 3.1 XLSB: nested checks inside fenced steps (G1)

- `XlsbOwner::finish` (`format/xlsb/owner.rs`) runs its read alone when a
  fence scope is active for its handle, as `XlsxOwner::finish` and the
  XLSB navigator's own `finish` already did (record 0262). Node property
  and text reads (`cell_state`, `props`, `text`, `preview`, `sheet_name`,
  `child_count`) go through it.
- `SourceWorkbook::ensure_current` (`format/xlsb/source.rs`) makes the
  leaf's checks in the leaf's order — execution context, source version,
  execution context — with one source observation. The leaf's
  `SourceBackedWorkbook::source_version` reads the version in a pre- and a
  post-flight with nothing between them; on a file that was a second
  `fstat` per check. The errors are the same conversions as before
  (`PackageError::from` of the package error).
- The per-match check and the check of every public read stay (owner
  decision 2026-10-14).

### 3.2 XLSX: header-row probes inside fenced steps (G2)

`Sheet::grid` (`format/xlsx/node.rs`) re-enters the leaf's fenced
`SourceWorksheet::grid()` on a cached read, so a cached grid still observes
the source. Inside a resolver or query step for the sheet's own handle it
no longer does: the step's closing check observes the same source after
the step. The sheet now records its handle's fence key when it is built;
a sheet lives only inside a node that borrows that handle.

### 3.3 XLS: axis writes resolve coordinates only (G3)

A layout write to a path target (`format/xls/layout.rs`) resolves with
`Nav::coordinates`, the coordinates-only mode the cell writer already uses.
In that mode `row[N]` and `col[X]` get a grid with no cells
(`Grid::coordinates_only`). The layout writer reads only the node's kind,
number and sheet. Query targets and `last()` selectors still read the grid.

### 3.4 Tried and not kept

A plain-integer fast path in `value::number::parse_f64` (exact for up to
15 digits) saved 220 instructions per RTF single set (0.24 %) and nothing
measurable in time; it was reverted.

## 4. Measurements

### 4.1 Run sets

| set | binaries | what |
|---|---|---|
| `trunk/` | base `0f3d8158cb` | every harness behind an open miss, two runs (three for `text_set_bench`), pinned, before any change: the re-measurement of §2 |
| `abba/` | base `0f3d8158cb`, after this record's code before the merge | `unified-query` (both builds) and `xls_p0_write_bench`, B A A B |
| `merged/` | the merged tree (`b53ed53006` + record) | every harness again, two runs: the rows the matrix parses |

`measure.sh trunk BASE - OUT` and `measure.sh abba BASE AFTER OUT` are the
commands (`trunk/` was taken with the same commands before the driver
gathered them, in the order its `log.txt` shows). A stray callgrind process (the native leg of an unfiltered RTF
profile) shared the host during the first three minutes of `abba/`, the
base-1 legs of the two 1,000-row XLSX families; B A A B keeps those
routes' after/base comparison balanced, and their ratios match base-2.

### 4.2 The fixed routes, base against after (`abba/`)

| route | base µs (1 / 2) | after µs (1 / 2) | ratio base → after |
|---|---:|---:|---:|
| `xlsb_1000_file_all_cell_value_gt` | 2,116.5 | 414.1 | 184.1 / 183.4 → 35.3 / 35.1 |
| `xlsb_1000_file_all_cell` | 926.0 | 532.6 | 85.8 / 78.2 → 43.5 / 50.3 |
| `xlsb_1000_bytes_all_cell_value_gt` | 438.2 | 340.5 | 40.2 / 40.6 → 31.9 / 31.0 |
| `xlsb_1000_bytes_all_cell` | 161.6 | 154.1 | 14.6 / 15.1 → 12.6 / 13.4 |
| `xlsb_1000_file_get_cell` (no budget) | 2.50 | 1.83 | 3.38 / 3.44 → 2.49 / 2.52 |
| `xlsb_1000_file_first_cell_value_gt` (no budget) | 5.69 | 2.56 | 7.36 / 7.39 → 3.34 / 3.36 |
| `xlsx_50000_file_all_row_header_gt` | 32,793.7 | 11,697.2 | 5.96 / 6.64 → 2.42 / 2.43 |
| `xlsx_50000_bytes_all_row_header_gt` | 9,779.4 | 7,944.2 | 2.00 / 1.58 → 1.01 / 1.44 |

µs are the lower of the two runs' medians. `xls_p0_write_bench` (ratios
per run, base-1 / after-1 / after-2 / base-2; every output byte-identical):

| case | base | after |
|---|---:|---:|
| `row_height`, positional, 100 cells | 1.120 / 1.111 | 1.030 / 1.027 |
| `column_width`, positional, 100 cells | 1.125 / 1.118 | 1.038 / 1.032 |
| `row_height`, positional, 10,000 cells | 1.094 / 1.070 | 1.002 / 1.000 |
| `column_width`, positional, 10,000 cells | 1.100 / 1.110 | 1.001 / 1.001 |
| `row_height`, `column_width`, eager, 10,000 cells | 1.024–1.046 | 0.999–1.002 |

Every other `unified-query` route stays within ±5 % between the builds,
except the PPTX large routes, which read +3.3 to +8 % in the after legs of
this session and −3 % in an earlier session of an intermediate build. No
PPTX code changed; callgrind counts 2.2 % *fewer* instructions after
(118.6 M → 116.0 M per five queries of `pptx_large_file_all_shape`), so the
reading is layout or host noise, as 0287 §5 found for similar readings.

### 4.3 The merged tree (`merged/`)

Two runs of every harness on the merged tree, the rows the matrix parses
(ratio ranges over the runs). They were taken at `b53ed53006`, before
records 0311 and 0312 merged. Those two change no query or XLS/RTF route:
0311 changes DOCX and DOC publication (`litchi-opc` overlay audits; the
DOCX rows here already meet) and 0312 the iWork tables, whose own runs
supersede the 0310 iWork rows in the matrix.

| group | routes | merged |
|---|---|---:|
| G1 | XLSB file `cell` / `cell[value>1000]`; bytes the same | 43.5–44.3 / 35.4 (531 / 412 µs); 13.1–13.3 / 30.5–30.8 |
| G2 | XLSX 50,000-row `row[Sales>1000]`, file / bytes | 1.12–2.15 (11.2 ms) / 0.90–1.35 (7.9 ms) |
| G3 | XLS positional `row_height` / `column_width`, 100 cells; 10,000 cells | 1.029 / 1.029–1.032; 1.000–1.014 / 1.001 |
| G4 | XLSX bytes `cell` (borrowed), PPTX large bytes `shape` | 15.8–16.7, 6.30–6.33 |
| G5 | XLSX bytes owned / fenced, file fenced; PPTX bytes owned / fenced, file fenced | 4.24–4.82 / 2.61–2.96, 1.19–1.22; 2.03–2.20 / 2.20–2.30, 1.28–1.29 |
| G6 | ODS `query_warm`, ODT `query_all` 10,000 ¶ | 23.0–23.7, 3.36–3.42 |
| G7 | DOCX `saut` file, bytes; `drawing` file | 1.49–1.51, 1.18–1.20; 1.23 |
| G8 | XLS eager `get`, positional `get`, positional scan 10,000 cells | 2.64–3.83, 1.39–1.68, 4.8–4.9 |
| G9 | RTF edit-closure single sets; layout path ×1; fan-out ×1 (owned clone / `&Query` / `Plan`) | 1.08–1.29; 1.14–1.15; 1.13–1.16 / 1.10–1.13 / 1.09–1.12 |
| G10 | DOCX single paragraph `set`; table-cell sets | 1.049–1.050; 1.08 (met) |
| G11 | Keynote bytes open; open + `get`; table cell; Numbers open + `get` | 1.058–1.061; 1.033–1.038 (met); 1.99–2.01; 1.047–1.058 (met) |
| G12 | ODT plain text set against 0304's before runs; ODP metadata read | 0.999–1.005 (met); 36–37 ns (no before build) |

## 5. Bench matrix

- New parsers: **0305** (its keynote, numbers, docx, rtf, pages and
  `docx_large_bench` after legs; `RTF_ROUTES` gains its route G, the fan-out
  passing `&Query`), **0309** (its element writes, against its stated 5 %
  or 20 µs bar), **0310** (`merged/`), **0311** (`docx_large_bench` with
  its committed-handle legs, the `office_gaps_bench` DOC scenarios) and
  **0312** (its Keynote, Pages and Numbers runs) and **0313**
  (`docx_large_bench`; every row meets). Record 0308 has no paired
  measurement and is listed in the notes. The record range runs to 0313.
- New curated annotations attach §2's group and measured cause to every
  open 0310 row, and 0305 §4.5, 0309 §6 and 0311 §4.2's explanations to
  their records' misses;
  the 0290 and 0263 before/after rows gain 0310's note.
- `python3 tools/unified_bench_matrix.py --check` exits 0, and no miss lacks
  an explanation.

## 6. Correctness evidence

- `format::xlsb::tests::query_steps_share_one_check_with_their_node_reads`:
  on generated sheets of 10 and 20 numbers, ten more cells cost exactly ten
  more source observations for `cell` and `cell[value>0]` and none for
  `cell[value>1000]` (before: twenty for `cell`); a source changed between
  two pulls fails the next pull with `IoError`.
- `format::xlsx::unified_tests::policy::header_row_probes_share_the_step_check`:
  ten more matching rows of a header-row query cost exactly ten more
  observations (before: forty); the kitchen-sink query keeps its paths and a
  change between pulls fails the next pull.
- `tests/xls_unified_layout.rs::layout_path_writes_publish_the_native_edit_bytes`:
  `set /Data/row[1]` and `set /Data/col[C]` publish exactly the native
  `Edit::set_row_height` / `set_column_width` bytes on both owner routes,
  and `row[last()]` still resolves through the grid.
- The query-corpus differential, the corpus fidelity test, the checklist
  goldens, `path_allocations`, the capability and registry goldens (both
  feature sets) and the OfficeCLI corpus pass unchanged; nothing was
  re-blessed (§8).

## 7. Gaps and follow-ups

By owner. Nothing here is claimed fixed.

- **Shared query walker (G4, G5, G6; 50 misses).** The residual against
  the fair leaves is the walker step and the owned node, with no single
  hotspot (0287 §5, 0291 §5). Next step as 0287 proposed: a fast path for
  a single kind-only positionless plan that skips the generic visit and
  descend tests, and smaller DOCX (176 bytes) and PPTX (168 bytes) nodes.
- **Harness comparisons (G4, G6).** XLSB and ODS have no fair leaf leg; add
  `_owned`/`_fenced` routes under new names so their misses compare like
  for like, as 0280 did for XLSX and PPTX.
- **DOCX adapter, OOXML owner (G7, G10).** `Children::next` should skip its
  per-child source check inside a fenced step, as XLSX, XLSB and PPTX do
  (record 0262); on `saut_page` that is about 7 µs of a 27 µs query. The
  large-document set and add are 0305 §4.5's (the overlay's second audit of
  the replaced main part).
- **XLS reads (G8).** Each text-path `get` re-resolves the sheet step
  (vocabulary lookups, the sheet child, path validation: ≈1.7 K of the
  2.4 K instructions over native). A per-handle cache of the parsed sheet
  step, or a cheaper vocabulary miss for sheet names, is the cheapest next
  step. The positional owner's per-match and `View::text` checks stay by
  owner decision.
- **RTF single targets (G9).** Coercing string props is about 0.9 K
  instructions per length (`Length::parse_in` → `parse_emu`: the unit table
  scan and the .NET-style float parse) and 3.6 K for a paragraph's nine
  layout props. A typed-props route (`prop::*` keys with typed values)
  would separate coercion from the write path in the harness; making the
  coercion itself cheaper means a unit-suffix fast path in `value/length`.
  Planning, the walk and the transaction are the rest.
- **iWork owner (G11).** Keynote bytes open: `Iwork::detect_bytes` copies the
  owned `Vec` into an `Arc<[u8]>` (500 KB for `basic.key`); adopting the
  `Vec` without the copy needs an ingress that accepts it. Keynote and Pages
  table cells: 0309 §6's lazy table parts.
- **Leaf, `litchi-xlsb`.** `SourceBackedWorkbook::source_version` reads the
  source version twice in a row (pre- and post-flight with nothing between
  them). The adapter now avoids it (§3.1); the leaf's own callers still pay
  it.
- **G12.** No follow-up: the 0290 and 0263 readings are before/after builds
  with no added work beyond the requested output, and cannot be rebuilt
  here.

## 8. Gates

On the final merged tree (`6dd73bb204` + this record's documents and the
`fmt` fix below), build environment per the
rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`, one cargo command at a
time in the memory-capped scopes, test executables pruned between chunks).
Every exit status was checked (`gates.sh`, its summary in the scratch
directory). The same list also passed on `b53ed53006`, before the second
merge. Record 0313, merged last, changes only `xml-minifier` and
`litchi-opc`; after it `cargo fmt --all --check`, the all-features clippy
of `litchi-ops`, `litchi` and `litchi-py`, `cargo test -p litchi-ops
--all-features --lib` and the matrix `--check` were run again (§8, last
rows).

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` (after `404f635d27`: trunk `1100472b75`'s `tests/iwork_unified.rs` from record 0312 was not formatted) | 0 |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, xlsx, xlsb, xls, rtf | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops` | 0 |
| `cargo test -p litchi-ops --all-features` in chunks: `--lib` (922), `--doc` (210), the 138 test targets in five groups (1,253) — includes the query-corpus differential, corpus fidelity, checklist goldens, `path_allocations` | 0 each |
| `cargo test -p litchi-ops --all-features --test capabilities --test registry` (not re-blessed) | 0 |
| the same with `--no-default-features --features` every litchi-ops feature but `encryption` (not re-blessed) | 0 |
| `cargo test -p litchi --all-features` (230) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,306 bans, 228 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (279 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `python3 tools/unified_bench_matrix.py --check` | 0 |
| after the 0313 merge (`fde3041d13`): `cargo fmt --all --check`; `cargo clippy -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`; `cargo test -p litchi-ops --all-features --lib` (922); `python3 tools/unified_bench_matrix.py --check` | 0 each |

No capability, registry or OfficeCLI-visible behaviour changed, so the
delta table and the OfficeCLI corpus were not regenerated. No public name
was added or changed; no glossary row changes. No dependency was added
(the `unified-query` lockfile gains the existing `memchr` entry its
`--locked` build needed on trunk).

## 9. Not claimed

- That any G4–G9 miss meets its budget, or that the G4 and G6 comparisons
  can (their leaf legs do not produce what a unified query returns).
- That G1 and G2 meet 1.10: they are now G4- and G5-shaped residuals.
- Any change of a frozen route name: two examples gained case filters, and
  their default runs are unchanged.
