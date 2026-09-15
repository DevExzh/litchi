# Survey: litchi-xls (BIFF8 over CFB) — remaining opportunities at HEAD 2fc5fc657

Sources: current tree (`crates/litchi-xls/src`, line numbers below), the retained 0584 packet
(`docs/performance/results/change-0584/PROFILE.txt`, `analysis/`), records 0136-0138, 0168, 0172,
0300, 0568, 0574, 0576, 0584, 0585. One fresh measurement: `rustc -Zprint-type-sizes` over
`litchi-xls --release --locked` in a scratch target (285 MB, deleted); summary retained at
`scratchpad/agents/xls/type-sizes-summary.txt`. No timing was run. Every Ir figure is 0584's
callgrind isolation-pair number (instructions rank work, not latency).

## 1. Path map

**Open** (`SourceBackedWorkbook::from_shared_ole_file_with_limits`, `workbook/source.rs:1013`):
one source observation (`ensure_current_parts` 1354, `source.version()` 1366),
`select_workbook_stream` (1489), `parse_globals` (1685). The globals pass frames every record once
into `GlobalsBuffer` (1532; fills 512 B doubling to 64 KiB, `ensure` 1604, `resize(finish, 0)`
1647, `read_stream_range_hinted` 1649), retains **all globals bytes**, frames them a **second**
time (`BiffRecords::with_limits` 1830) and interprets twelve kinds: `FilePass`, `CodePage`,
`BoundSheet8`, `SST`+`Continue` (1889-1972) and, in `Formatting::parse_globals`
(`number_format/codec.rs:57`, arms 71-118), `Date1904`, `Format`, `XF`, `XFCRC`, `XFExt`, `DXF`
(plus `BOF`/`EOF`). Everything else is read, framed twice and dropped. The SST is measure-walked
(`records.rs:1132`, `walk_one_shared_string` 1064), retaining `segments` (24 B each) and `entries`
(16 B each; `records.rs:989`), no text. Retained: `SourceInner` (378) with `Box<[SheetEntry]>`
(72 B each), names, `Arc<SharedStringSstScan>`, `Arc<Formatting>`. Cost (0584): flagship open
2,379,097 Ir, 53 reads/565,201 B, memcpy 28.6% + memset 23.4%; `54016.xls` open 7,239,535 Ir of
which the five SST symbols are 3,581,589 Ir (49.5%) for 7,893 strings = **454 Ir per string**;
`SstCursor::read_exact` issues 7,900 `copy_from_slice` calls of 2-3 bytes (`records.rs:706`).

**List** (`worksheet_names` 1070): metadata clone under one `ensure_current`; nothing scanned.

**One cell** (`query_cell` 2735): `WorksheetScan::new` (2135) builds a cursor with a cold chain
walk (`stream_cursor_at` → `stream_cursor_at_hinted`, `litchi-cfb/src/shared.rs:729/762`; 599
links on 54016, up to 2,477 on the flagship's last sheet), then `loop { next_frame()? }` (2766)
**to EOF**: `read_payload` for the seven cell kinds, `skip_payload` (2384) otherwise, `process_cell`
(3164) parses every cell (`CellRecord::parse`, `validate_cell_xf` `number_format/codec.rs:27`) and
resolves only the target. Counts on 54016 (0584): 37,929 frames, 32,358 payload reads, 5,570
skips, 29,608 cell parses, 1 resolve; 24,525,331 Ir. Per frame: `next_frame` 112 Ir, `ensure`
2 calls at 23 Ir, `read_payload` 46, `query_cell` self 69, `drop_in_place<SourceBackedError>`
2×13.6 (opp. 1). Retained: nothing; the window (≤ 64 KiB; `drain(..framed)` 2221 memmoves the
tail, `resize(…,0)` 2238 zero-fills each fill) dies with the query. Fences: 2 observations per
query + 3 per resolved string (3018, 3070, 3109/3113).

**All cells**: no whole-sheet iterator exists (0568 limitations); a "full cell scan" is N
`query_cell` calls, each re-scanning to EOF.

**Full text** (`write_text_to_impl` 845 → `scan_text_sheet` 2436 → `write_text_sheet` 2594): one
`SharedStringResolver` (2969) across sheets (0585); per sheet the same scan but every cell is
decoded (`collect_source_cell` 3145) into `SourceTextSheet` (641): `HashMap<(u16,u16), CellValue>`
with **three hash operations per insert** (`get` 664, `contains_key` 688, `insert` 716) and two
eagerly built error values (675, 694). Each `LabelSst` resolve (2987) linearly scans `segments`
(3022-3034; 27 on 54016), allocates `chunks: Vec<Vec<u8>>`, one zero-filled `Vec` per chunk
(3079-3086), a `slices` Vec (3094) and the `String`, and observes the source **three times**.
Output walks the declared rectangle `(max_row+1)×(max_col+1)` seeded from `DIMENSIONS`
(2532-2541), one hash lookup per position and one observation per row (2602 → 2408). On a
`from_path` workbook every observation is an `fstat` (`FileSource::version`,
`litchi-core/src/source/file.rs:146-160`); the harness never sees this because every XLS selector
wraps bytes (`InstrumentedSource::new_xls`, `tools/perf-baseline/src/lib.rs:7477`).

**Eager path**: `litchi::sheet::Workbook::open(path)` routes XLS to the source-backed owner
(`crates/litchi/src/sheet/workbook.rs:350`); `from_bytes` (505-514) and every
`litchi_xls::Workbook::new/from_ole_file` (`workbook/package.rs:76-111`) read the whole `Workbook`
stream and parse every record of every sheet. Selectors: `xls_eager_open_*`, `xls_semantic_*`
(`lib.rs:1868-1878`).

**Edit + save** (`cell_values/mod.rs`): `Snapshot::from_bytes` (516) = `PackageEditor::open` over
the whole file + `parse_workbook_stream` inventory + a **complete eager `Workbook::new`** (540) +
per-string property clones. Generic `Transaction::commit` (2552) copies the workbook stream (2593)
and the file bytes into a second `PackageEditor` (2645), rebuilds the `Snapshot`
(`from_package_editor` 2656, or `from_fixed_numeric_package_editor` 577 which still runs
`Workbook::new` at 591), then `verify_readback` (2660). The plan-only fixed-width path
(`commit_source_backed_numeric_plan` 5126; 0137/0138/0168/0172): splice plan → composed target →
CFB reopen → `Workbook::new(ComposedPositionalReader)` **over the complete target** (5230) →
changed-cell readback (6463) → fingerprints. The non-plan `SourceBackedCommit` retains
`snapshot: Snapshot`, a complete target (3250).

## 2. Remaining opportunities, ranked

### 1. Lean worksheet frame loop, including 0584 candidate 2 — GOAL step 1 (and 3)
**Mechanism.** Two `ok_or({ SourceBackedError::ResourceLimit{..} })` calls in `next_frame`
(`source.rs:2307`, 2321) build a 48-byte error on **every** frame and drop it on the `Some` path;
one more in `parse_globals` (1754), two in `SourceTextSheet::insert` (675, 694). 0584's 75,858
drops per 54016 query = exactly 2×37,929 frames. Each frame also pays two non-inlined `ensure`
calls (2211), one `consume_payload` (2343), each returning a 48-byte `Result` through memory, and
two `check_execution` (2401). Fix: `ok_or_else`, an inlined "already resident" fast path,
cold-path error construction. Same checks, same order; no refusal moves.
**Size (measured, PROFILE.txt 54016/one-cell):** `drop_in_place<SourceBackedError>` 1,030,627 Ir
= 4.20%; total framing overhead (`next_frame` 17.32% + `ensure` 7.17% + `read_payload` 6.07% +
`query_cell` self 10.67% + drops 4.20% + `skip_payload` 0.66%) = **46.1%** of the query versus
18.0% of cell semantics; also 0.61% of the 54016 open (site 1754). Measured sizes:
`SourceBackedError` 48 B, `Result<WorksheetFrame, SourceBackedError>` 48 B, `WorksheetFrame` 16 B.
**Records:** 0584 cand 2 proposed, **not landed** (git log on `source.rs`: 57b25a820 is 0585's
hint; 0585 never mentions it). Nothing covers the wider loop cost.
**Scenarios:** one-cell, all-cells, full text (`xls_source_backed_open_one_cell`,
`xls_owned_source_open_one_cell`). **Risk:** low; no design record. Price in cycles (0579).
**Falsified if:** paired callgrind + `perf stat` on 54016 one-cell shows <2% Ir and no cycle change.

### 2. Lazy SST indexing (0584 candidate 3) — step 1
**Mechanism.** Defer the per-string measure walk (`records.rs:1207-1214`) to first use. The
frozen design must decide: (a) `Continue` boundaries are **already recorded at open for free** —
`segments` is built from framed records (1170-1183) before any string is walked, so only
`entries` is deferrable; (b) the deferred index is a prefix index (entry k needs the walk through
k-1) held in the snapshot under a lock/`OnceLock` — ADR 0005 permits "semantic payloads load
lazily into thread-safe weighted caches", `max_sst_entries` still bounds it; (c) **when** a
malformed SST is refused moves from open to the first resolve past the defect, and open/list-only
callers never see it (0576's "error-identity trap"); (d) header checks (`total >= unique`,
count-vs-bytes, 1190-1205) can stay at open, so only per-string refusals move; (e) 0300's
invalid-index `CellValue::Error` texts stay.
**Size (measured):** 3,581,589 Ir = **49.5% of the 54016 open**; 6.6% of the flagship open. It is
a per-open saving only when few strings are resolved; a full text of 54016 resolves all 16,055
`LabelSst` and pays the walk anyway.
**Records:** 0576 defers it; 0584 cand 3 (high risk); 0300 fixed the retained shape.
**Requires its own frozen design record.**
**Cheaper sibling 2b (step 3, low risk, no design record):** keep the eager walk but cut its
454 Ir/string: a fast path when header and characters lie within one segment (direct indexing
instead of `read_exact` → `copy_from_slice` for 2-byte reads, `records.rs:694-718`) and a measure
instantiation of `read_formatting_runs` (824) without the per-string `Vec` (830). Modelled ≤ half
of the 49.5%; 0576's differential harness proves error identity. **Falsified if** the cost is
dominated by segment-boundary logic.

### 3. Snapshot-scoped retained sheet index and a whole-sheet iterator — step 4
**Mechanism.** All-cells = N full re-scans today. A per-(snapshot, sheet) index built by the first
validated scan — per cell record its stream offset, `(row, col)`, kind, XF (≈16-24 B per cell;
29,608 cells → ≤0.7 MB on 54016), or a row-start index like XLSX 0005 — plus a "validated to EOF"
mark so later queries seek within the validated range. ADR 0005: the snapshot is immutable and
version-fenced, so an index keyed by `expected_version` is a clean-value cache; it must be
weighted, bounded by `max_worksheet_scan_records` and evictable — litchi-xls has no such cache, so
this needs a frozen design record (retention accounting, lock discipline, what a second query
re-validates). Even without the index, a public `SourceBackedWorksheet::cells()` iterator turns N
scans into one and is the missing all-cells selector.
**Size (measured per query):** 54016 scan delta 17,285,796 Ir (70.5% of the query); flagship
237,761 (9.1%); N-fold across repeated queries. The index itself is modelled.
**Records:** 0568 limitations flag the re-scan; 0584 names the retained scan "the tractable
form"; 0574 opp 6 (early exit) is rejected and this is not that — the first scan still validates
to EOF. Risk medium.
**Falsified if:** second-query cost is not ≥5x below a re-scan on 54016, or retained weight
breaks the bounded-memory contract.

### 4. Full text: per-string/per-row freshness fences and the rectangle walk — steps 2 and 1
**Mechanism.** (a) Three source observations per string (3018, 3070 per chunk, 3109/3113) and
one per row (2602); on `from_path` each is an `fstat`. ADR 0005 requires only that "mutation
during a read returns `SourceChanged`"; fences at scan start and end detect the same mutations
but change *when* the refusal surfaces, so a short design note is needed. (b) The output loop
walks the declared rectangle with a hash lookup per position (2600-2637); iterating retained
cells in sorted order emits identical bytes. (c) `insert` can be one `entry` operation.
**Size:** modelled from code, **unmeasured**: ≥48,165 `fstat` (3×16,055) plus one per row for one
54016 text extraction on a file source; zero in the harness (in-memory `version()` is a field
read, `litchi-core/src/source.rs:182`). Rectangle cost unknown (no selector).
**Records:** none found (grepped "fence", "fstat", "rectangle"; 0563/a8ada83c1 cut observations
per check, not checks). Scenario: full text (facade `text()`, `workbook.rs:828`). Risk
low-medium. **Falsified if** `strace -c` on a `from_path` extraction puts fences under 5% of wall.

### 5. Measure-only validation of non-target cells in `query_cell` — step 2
`CellRecord::parse` materializes `Label { value: String }` (`records.rs:2124`) and
`Formula { formula: Vec<u8> }` (2150; `formula_metadata/codec.rs:65` `to_vec()`) for every record;
`process_cell` drops all but the target. A `MeasuredText`-style instantiation (0576 pattern)
validates without allocating. `CellRecord` is 88 B; `drop_in_place<CellRecord>` is 503,336 Ir
(2.05%) on 54016 one-cell. **Size:** modelled; small on 54016 (mostly `LabelSst`/RK), larger on
formula-heavy sheets — the census (`analysis/xls-labelsst-census.txt`) counted only `LabelSst`,
so the mix is **unknown**. Records: none found. Risk low-medium (needs 0576's differential
proof). **Falsified if** a census shows <5% Formula/Label records on every primary fixture.

### 6. Skip never-interpreted globals payloads + frame once (0574 opps 2 and 5) — step 2
**Status:** open. Consumed kinds confirmed in the current tree (§1). The 75% figure is 0574's
static model over 126 fixtures (3,002,998 of 3,999,694 globals bytes); flagship 530,960 of
551,377 B, 95.2% one `MsoDrawingGroup`+`Continue` chain. Mechanism: for an unconsumed kind whose
payload lies past the filled window, `skip_forward` (as `skip_payload` 2384) instead of filling.
**Gate:** the `dense_frames` mean-bytes-per-record hysteresis (2180, 1 KiB). Unlike the
worksheet gate (untestable, 0568) it can fire on a real fixture: the flagship's drawing chain is
524,839 B over ≈261 records ≈ 2 KiB/record (modelled from 0574 + census; verify). Costs reads
(+46 requests modelled) — a range-source regression. Double framing (1.78% flagship, 3.21% 54016)
goes inside the same change. Size ~30% of the flagship open (0574, Ir upper bound); opp. 7 shrinks
with it. Risk medium-high; frozen design record needed. **Falsified if** dense-globals fixtures
(103 of 126) regress in cycles more than sparse ones gain.

### 7. Zero-fill of fill buffers (0574 opp 4 / 0584 cand 5) — step 2
Blocked by GOAL rule 10; **no append-style read exists** (`ReadAt::read_at(&mut [u8])`, cursor
`read_exact(&mut [u8])`; no `spare_capacity`/`MaybeUninit` in litchi-cfb or litchi-core).
Measured upper bounds: 548,308 Ir memset from `GlobalsBuffer::ensure` = 23.0% of the flagship
open; 612,727 + 303,790 Ir = 3.7% of 54016 one-cell. A safe partial form (a `ReadAt` default
method `read_append_at(offset, len, &mut Vec)` overridden by `OwnedSource` with
`extend_from_slice`) helps only in-memory sources — the harness, not `from_path` — so it must not
land on harness evidence. Defer behind opp. 6.

### 8. Per-sheet cursor construction hint (0585 limitation) — step 4
28,143 links on the 16-sheet flagship, 91% resumable (`analysis/sheet-cursor-model.txt`); 0 on
single-sheet 54016. Needs a second `StreamChainHint` with a lifetime disjoint from the resolver's
(thrashing, 0585). Benefits full text/multi-sheet only; cycles unknown; low risk, no ADR issue.
**Blocker:** the flagship's text extraction is refused (§4); measure on
`HyperlinksOnManySheets.xls` (3 sheets, 57% modelled) or `WithCustomViews.xls`.

### 9. XLS edit + save: the candidate readback is a complete eager open — step 1
Whole-workbook-proportional steps that remain (§1): two eager parses at open, one eager parse of
the composed candidate (5230; 591 on the eager path), ≥2 fingerprint scans (0168 removed two of
four), CFB reopen chain validation, and — on the non-plan path — a retained complete target (3250).
HOTSPOTS rank 13's "second complete target artifact" is gone only for plan-only Number/RK
families (0138: "zero complete target-artifact bytes"), not for string, formula, style or
structural edits. Opportunity: a source-backed candidate readback (globals + scan of the edited
sheets) replacing `Workbook::new`, if `require_public_worksheet_coverage`,
`require_unprotected_workbook`, `require_macro_free_workbook` (5231-5233) are proven equivalent
on the source-backed owner. **Size unknown**: 0138 p50 Number 105.3 ms plan-only vs 145.4 ms
source-backed, RK 1.23 vs 1.63 ms, and no attribution profile of the edit path exists. ADR 0006
keeps the readback mandatory; changing its owner needs a proposed ADR or frozen record. Risk
medium-high. **Falsified if** a callgrind of `xls_numeric_plan_only_*` shows `Workbook::new`
under 15% of the commit.

### 10. `resolve_shared_string` per-resolve allocations — step 2
Per resolve: `chunks` Vec, one zero-filled `Vec` per chunk, `slices` Vec, linear segment scan.
`SharedStringResolver` (0585) is the natural home for a reused scratch buffer and a binary search.
16,055 resolves per 54016 text; unmeasured (no selector). Low risk.

## 3. Looks like an opportunity but is not
- **Early exit from `query_cell`**, or seeking via `INDEX`/`DBCELL` row blocks
  (`src/row_block_index/`, an eager-owner model): rejected, 0574 opp 6 / 0584 (ADR 0005 mandatory
  validation, GOAL rule 12). Item 3 is the sanctioned form.
- **`directory_name_data` handoff** (7.8% on small fixtures): rejected, 0554.
- **`collect_exact` / collector variants**: rejected 0524/0548/0549/0533/0536; 0533 forbids reviving 0524.
- **Pending-role accounting**: rejected 0555. **DOC-style resolve hint**: rejected 0586.
- **One chain hint shared by resolver and worksheet cursor**: thrashes (0585).
- **Windows wider than 64 KiB**: 0568 chose the cap; 0572 found 4→64 KiB indistinguishable.
- **Touched-closure CFB validation / cross-run coalescing**: 0574 opps 6-7, ADR 0005/0006.
- **Zero-fill via `MaybeUninit`**: GOAL rule 10. **A decoded `Vec<String>` SST cache**: 0300
  deliberately retains offsets only.
- **Parallel per-sheet scans on a global pool**: no hidden global Rayon pool; small tasks regress.

## 4. Measurement blockers and defects noticed
- **No source-backed XLS full-text or all-cells selector**; `xls_source_attribution` supports
  only `open`/`list`/`one-cell` (`bin/xls_source_attribution.rs:52-54`). Items 3, 4, 8, 10 cannot
  be sized today.
- **All XLS harness sources are in-memory** (`InstrumentedSource::new_xls`); no `from_path`
  selector, so `FileSource::version` fstat costs are invisible (item 4).
- **Flagship text extraction is refused**: `Invalid record 0x0006: shared Formula metadata
  requires a leading PtgExp token` (0568, 0585). Whether the eager path refuses the same file is
  unrecorded; triage as a correctness/leniency question before any text measurement on it.
- **No attribution of the XLS edit/save path** (0584 profiled reads; DOC/PPT had a throwaway
  driver, XLS edits none). Item 9 is unsized.
- **Density gate**: untestable on real worksheets (0568); the globals variant may fire on the
  flagship (modelled only, item 6).
- **Ir vs cycles**: items 1, 3, 5 are branch/pointer-heavy; price in cycles (GOAL_AUDIT).
- **Formula/Label record mix unknown** across the corpus (census counted `LabelSst` only).
- Oddities (report only): `SourceTextSheet::insert` does three hash operations and two eager
  error constructions per cell (`source.rs:664-716`); `write_text_sheet` emits
  `DIMENSIONS`-declared empty trailing columns as tabs (format choice, rectangle-proportional);
  a stale `size.rs` from a previous survey wave sat in this agent's scratch dir and was removed.

Scratch retained: `scratchpad/agents/xls/type-sizes-summary.txt`; the 285 MB scratch build
target and raw type-size dumps were deleted. No tracked file was modified; nothing committed.
